"""
=====================================================================
  Program Name  : testWorkload.py
  Author        : Manuel Figallo
  Purpose       : Measure how an environment scales as a workload grows.
                  Builds a synthetic dataset in increasing steps up to a
                  stated maximum, timing the build and a function run at
                  every step, and reports whether the cost per unit of
                  data stayed flat or got worse.
  Version       : 1.0beta
  Created       : 2026-10-13
  Last Modified : 2026-10-13

  Dependencies:
    STANDARD LIBRARY ONLY. pandas and numpy are used WHEN PRESENT, for a
    faster build and a real memory figure, and the module works without
    them. Nothing here reads or writes a file, so it runs unchanged inside
    the Python in Excel sandbox.

  Description:
    WHAT IT ANSWERS. Not "how fast is this machine" -- a single timing tells
    you almost nothing. It answers "does ten times the data cost ten times
    the time, or more?", which is the question behind every capacity
    decision. The headline column is ScalingIndex: 1.0 means linear, 2.0
    means a row costs twice as much to process at full size as it did at
    the first step, which is where capacity planning goes wrong.

    RUNNING IT IN EXCEL -- read this before you trust the numbers.
    Python in Excel executes in a Microsoft-managed container in Azure, NOT
    on your laptop and NOT on a CGS server. So the figures describe
    Microsoft's container, not your environment. That is useful for
    comparing one algorithm against another, and useless for sizing a CGS
    server. To measure a CGS environment, run this module there -- from a
    command line, or from the SAS/PowerShell side -- and compare.

    See the EXCEL INSTRUCTIONS block at the bottom of this file for the
    exact steps and a pasteable snippet.

  Input Parameters (required first):
    RowCount    (REQUIRED, int) - rows at 100%, the maximum size tested.
    ColumnCount (REQUIRED, int) - columns in the synthetic dataset.
    Increment   (optional, int, default 10) - percentage step. 10 gives
                  10%, 20% ... 100%. 25 gives 25%, 50%, 75%, 100%.
    Work        (optional, callable) - the function to time at each step.
                  It is called with the dataset and its result is ignored.
                  Default: a group-by-and-aggregate, which is the shape of
                  most reporting work.
    ScaleBy     (optional, str, default "rows") - "rows", "columns" or
                  "both". "both" grows the grid in two directions at once,
                  so cells rise with the SQUARE of the step; the default is
                  the one whose table reads straight.
    Repeats     (optional, int, default 1) - time Work this many times per
                  step and keep the median. Raise it to 3 or 5 when the
                  timings are small and noisy.
    Seed        (optional, int, default 42) - keeps the data reproducible,
                  so two runs are comparable.
    MeasureCsv  (optional, bool, default False) - also measure the size the
                  data would be as CSV. Honest but slow; it is timed
                  separately so it never contaminates CreateSeconds.
    AsDataFrame (optional, bool, default True) - return a pandas DataFrame
                  when pandas is available, else a list of dicts.
    Label       (optional, str) - a name for this run, carried on every row
                  so two environments' tables can be stacked and compared.

  Returns:
    A table (DataFrame or list of dicts), one row per step. See COLUMNS.

  Exit codes (CLI):
    0 = success, 2 = config error.
=====================================================================
"""

from __future__ import annotations

import argparse
import io
import platform
import random
import statistics
import string
import sys
import time
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Sequence

_HERE = Path(__file__).resolve()
if str(_HERE.parent.parent.parent) not in sys.path:
    sys.path.insert(0, str(_HERE.parent.parent.parent))

try:                                              # pragma: no cover - trivial
    from src.utils.logger import logInfo, logWarn
except ImportError:                               # pasted into Excel, alone
    def logInfo(message: str) -> None:            # type: ignore[misc]
        """Fallback logger used when the module is pasted in on its own."""
        print(f"INFO  {message}", file=sys.stderr)

    def logWarn(message: str) -> None:            # type: ignore[misc]
        """Fallback logger used when the module is pasted in on its own."""
        print(f"WARN  {message}", file=sys.stderr)

__version__ = "1.0beta"

EXIT_OK, EXIT_CONFIG_ERROR = 0, 2

#: Output columns, in order.
COLUMNS = [
    "Label", "StepPct", "Rows", "Columns", "Cells",
    "CreateSeconds", "MemoryMB", "CsvMB",
    "WorkSeconds", "CreateSecPerMCell", "WorkSecPerMCell",
    "ScalingIndex", "CellsPerSecond", "Note",
]

#: Below this, a timing is mostly scheduler noise and the ScalingIndex built
#: from it means nothing. Carried as a NOTE ON THE ROW rather than a log
#: line, because in Excel stderr is invisible: a caveat that is not in the
#: table is a caveat the reader never sees.
NOISE_FLOOR_SECONDS = 0.05

#: The repeating column shape of the synthetic data. Mixed types on purpose:
#: a grid of identical integers is far cheaper to build, hold and process
#: than real data, and would flatter the environment under test.
COLUMN_KINDS = ("id", "code", "amount", "date", "flag", "text")

VALID_SCALE_BY = ("rows", "columns", "both")


def columnPlan(ColumnCount: int) -> List[Dict[str, str]]:
    """Decide the name and type of each column.

    Parameters: ColumnCount (int) - how many columns.
    Returns: list of {"name", "kind"} in grid order.
    """
    plan = []
    for index in range(ColumnCount):
        kind = COLUMN_KINDS[index % len(COLUMN_KINDS)]
        plan.append({"name": f"{kind}_{index + 1}", "kind": kind})
    return plan


def buildDataset(Rows: int, Plan: Sequence[Dict[str, str]],
                 Seed: int = 42) -> Dict[str, List[Any]]:
    """Build a synthetic dataset as a dict of column lists.

    Parameters:
        Rows (int)       - how many rows.
        Plan (sequence)  - from columnPlan().
        Seed (int)       - seeded so two runs are comparable.
    Returns: dict mapping column name -> list of values.

    A dict of lists rather than a list of dicts: it is the shape pandas
    takes directly, and it does not pay a per-row dictionary overhead that
    would be measured as if it were the environment's fault.
    """
    generator = random.Random(Seed)
    codes = ["".join(generator.choices(string.ascii_uppercase, k=3))
             for _ in range(50)]
    data: Dict[str, List[Any]] = {}
    for column in Plan:
        kind, name = column["kind"], column["name"]
        if kind == "id":
            data[name] = list(range(1, Rows + 1))
        elif kind == "code":
            data[name] = [generator.choice(codes) for _ in range(Rows)]
        elif kind == "amount":
            data[name] = [round(generator.uniform(0, 10000), 2)
                          for _ in range(Rows)]
        elif kind == "date":
            data[name] = [f"2026-{generator.randint(1, 12):02d}-"
                          f"{generator.randint(1, 28):02d}" for _ in range(Rows)]
        elif kind == "flag":
            data[name] = [generator.choice([True, False]) for _ in range(Rows)]
        else:
            data[name] = [f"row {index} note" for index in range(Rows)]
    return data


def toFrame(data: Dict[str, List[Any]]) -> Any:
    """Return a pandas DataFrame when pandas is installed, else the dict.

    Parameters: data (dict) - from buildDataset().
    Returns: DataFrame or the dict unchanged.
    """
    try:
        import pandas
    except ImportError:
        return data
    return pandas.DataFrame(data)


def measureMemoryMb(dataset: Any) -> float:
    """Measure how much memory the dataset occupies, in MB.

    Parameters: dataset - a DataFrame or a dict of lists.
    Returns: float MB, rounded to 3 places.

    deep=True on a DataFrame so Python string objects are counted; without
    it a column of text reports as 8 bytes a row, which is the pointer, not
    the string, and understates a text-heavy table by an order of magnitude.
    """
    try:
        import pandas
        if isinstance(dataset, pandas.DataFrame):
            return round(dataset.memory_usage(deep=True).sum() / 1_048_576, 3)
    except ImportError:
        pass
    total = sys.getsizeof(dataset)
    for values in dataset.values():
        total += sys.getsizeof(values)
        # Sample rather than walk millions of objects: the whole point is to
        # measure the workload, not to be the workload.
        sample = values[:1000]
        if sample:
            total += int(sum(sys.getsizeof(v) for v in sample)
                         / len(sample) * len(values))
    return round(total / 1_048_576, 3)


def measureCsvMb(dataset: Any) -> float:
    """Measure the size the dataset would be as CSV, in MB.

    Parameters: dataset - a DataFrame or a dict of lists.
    Returns: float MB, rounded to 3 places.

    Written to an in-memory buffer, never to disk: Python in Excel has no
    filesystem, and the figure is the same either way.
    """
    buffer = io.StringIO()
    try:
        import pandas
        if isinstance(dataset, pandas.DataFrame):
            dataset.to_csv(buffer, index=False)
            return round(len(buffer.getvalue().encode("utf-8")) / 1_048_576, 3)
    except ImportError:
        pass
    import csv
    names = list(dataset.keys())
    writer = csv.writer(buffer)
    writer.writerow(names)
    for row in zip(*[dataset[name] for name in names]):
        writer.writerow(row)
    return round(len(buffer.getvalue().encode("utf-8")) / 1_048_576, 3)


def defaultWorkload(dataset: Any) -> Any:
    """Group by the first code column and total the first amount column.

    Parameters: dataset - a DataFrame or a dict of lists.
    Returns: the aggregate, which the caller ignores.

    This shape -- group, aggregate, sort -- is what most reporting actually
    does, and it exercises hashing, comparison and memory together rather
    than one of them in isolation.
    """
    try:
        import pandas
        if isinstance(dataset, pandas.DataFrame):
            codeColumns = [c for c in dataset.columns if c.startswith("code_")]
            amountColumns = [c for c in dataset.columns
                             if c.startswith("amount_")]
            if codeColumns and amountColumns:
                return (dataset.groupby(codeColumns[0])[amountColumns[0]]
                        .sum().sort_values(ascending=False))
            return dataset.describe()
    except ImportError:
        pass
    codeName = next((n for n in dataset if n.startswith("code_")), None)
    amountName = next((n for n in dataset if n.startswith("amount_")), None)
    if codeName is None or amountName is None:
        return len(next(iter(dataset.values()), []))
    totals: Dict[Any, float] = {}
    for code, amount in zip(dataset[codeName], dataset[amountName]):
        totals[code] = totals.get(code, 0.0) + amount
    return sorted(totals.items(), key=lambda item: item[1], reverse=True)


def stepSizes(RowCount: int, ColumnCount: int, Increment: int,
              ScaleBy: str) -> List[Dict[str, int]]:
    """Work out the (rows, columns) to test at each step.

    Parameters:
        RowCount, ColumnCount (int) - the 100% dimensions.
        Increment (int)             - percentage step.
        ScaleBy (str)               - rows | columns | both.
    Returns: list of {"pct", "rows", "columns"}, always ending at 100%.
    Raises: ValueError - Increment is not between 1 and 100.
    """
    if not 1 <= Increment <= 100:
        raise ValueError(f"Increment must be between 1 and 100, got {Increment}")
    percentages = list(range(Increment, 101, Increment))
    if percentages[-1] != 100:
        percentages.append(100)          # always finish at full size
    steps = []
    for pct in percentages:
        rows = RowCount if ScaleBy == "columns" else max(1, round(RowCount * pct / 100))
        columns = (ColumnCount if ScaleBy == "rows"
                   else max(1, round(ColumnCount * pct / 100)))
        steps.append({"pct": pct, "rows": rows, "columns": columns})
    return steps


def describeEnvironment() -> Dict[str, Any]:
    """Describe the machine the measurements came from.

    Parameters: none.
    Returns: dict - python version, platform, processor, and whether pandas
             was available (which changes the numbers materially).

    Always report this next to the timings. A table of seconds with no note
    of what produced them cannot be compared with anything later.
    """
    try:
        import pandas
        pandasVersion = pandas.__version__
    except ImportError:
        pandasVersion = "not installed"
    return {
        "Python": platform.python_version(),
        "Platform": platform.platform(),
        "Processor": platform.processor() or "unknown",
        "Pandas": pandasVersion,
    }


def testWorkload(RowCount: int, ColumnCount: int, Increment: int = 10,
                 Work: Optional[Callable[[Any], Any]] = None,
                 ScaleBy: str = "rows", Repeats: int = 1, Seed: int = 42,
                 MeasureCsv: bool = False, AsDataFrame: bool = True,
                 Label: str = "") -> Any:
    """Measure how build time and work time grow as a dataset grows.

    Parameters:
        RowCount (int)     - REQUIRED rows at 100%.
        ColumnCount (int)  - REQUIRED columns.
        Increment (int)    - percentage step; default 10.
        Work (callable)    - the function to time; default defaultWorkload.
        ScaleBy (str)      - rows | columns | both; default rows.
        Repeats (int)      - timings per step, median kept; default 1.
        Seed (int)         - reproducible data; default 42.
        MeasureCsv (bool)  - also measure CSV size; default False.
        AsDataFrame (bool) - return a DataFrame when pandas is present.
        Label (str)        - a name carried on every row.
    Returns:
        One row per step with the COLUMNS above. ScalingIndex is the one to
        read: 1.0 is linear scaling, above 1.0 means each unit of data got
        MORE expensive as the dataset grew.
    Raises:
        ValueError - a parameter is out of range.

    Use in claims processing:
        Before promising a monthly run will cope with a year of claims,
        measure it at 10% and 100% of that year and look at ScalingIndex. A
        job that is linear will finish; one that is not will quietly stop
        finishing somewhere in between.
    """
    if RowCount < 1:
        raise ValueError(f"RowCount must be at least 1, got {RowCount}")
    if ColumnCount < 1:
        raise ValueError(f"ColumnCount must be at least 1, got {ColumnCount}")
    if Repeats < 1:
        raise ValueError(f"Repeats must be at least 1, got {Repeats}")
    if ScaleBy not in VALID_SCALE_BY:
        raise ValueError(f"ScaleBy must be one of {', '.join(VALID_SCALE_BY)}, "
                         f"got {ScaleBy!r}")

    work = Work or defaultWorkload
    steps = stepSizes(RowCount, ColumnCount, Increment, ScaleBy)
    logInfo(f"testWorkload {__version__}: {len(steps)} step(s) up to "
            f"{RowCount:,} x {ColumnCount}, scaling by {ScaleBy}")

    rows: List[Dict[str, Any]] = []
    baseline: Optional[float] = None
    for step in steps:
        plan = columnPlan(step["columns"])

        startedBuild = time.perf_counter()
        dataset = toFrame(buildDataset(step["rows"], plan, Seed))
        createSeconds = time.perf_counter() - startedBuild

        memoryMb = measureMemoryMb(dataset)
        # Measured AFTER the build timer stops, so it never inflates it.
        csvMb = measureCsvMb(dataset) if MeasureCsv else None

        timings = []
        for _ in range(Repeats):
            startedWork = time.perf_counter()
            work(dataset)
            timings.append(time.perf_counter() - startedWork)
        workSeconds = statistics.median(timings)

        cells = step["rows"] * step["columns"]
        millionCells = cells / 1_000_000 or 1e-9
        workPerMCell = workSeconds / millionCells
        if baseline is None:
            baseline = workPerMCell or 1e-9

        rows.append({
            "Label": Label,
            "StepPct": step["pct"],
            "Rows": step["rows"],
            "Columns": step["columns"],
            "Cells": cells,
            "CreateSeconds": round(createSeconds, 4),
            "MemoryMB": memoryMb,
            "CsvMB": csvMb,
            "WorkSeconds": round(workSeconds, 4),
            "CreateSecPerMCell": round(createSeconds / millionCells, 4),
            "WorkSecPerMCell": round(workPerMCell, 4),
            "ScalingIndex": round(workPerMCell / baseline, 3),
            "CellsPerSecond": int(cells / workSeconds) if workSeconds else 0,
            "Note": ("timing under %.2fs - too small to trust, raise RowCount "
                     "or Repeats" % NOISE_FLOOR_SECONDS)
                    if workSeconds < NOISE_FLOOR_SECONDS else "",
        })
        logInfo(f"  {step['pct']:>3}%  {step['rows']:>10,} x "
                f"{step['columns']:<3}  build {createSeconds:7.3f}s  "
                f"work {workSeconds:7.3f}s  index {rows[-1]['ScalingIndex']}")
        del dataset                       # free before the next, larger step

    noisy = [row for row in rows if row["Note"]]
    if noisy:
        logWarn(f"{len(noisy)} of {len(rows)} step(s) ran in under "
                f"{NOISE_FLOOR_SECONDS}s. Those ScalingIndex values are "
                f"measuring scheduler noise, not scaling. Raise RowCount, or "
                f"Repeats, until every step takes a measurable time.")
    final = rows[-1]
    if not final["Note"] and final["ScalingIndex"] >= 1.5:
        logWarn(f"scaling is NOT linear: at full size each million cells cost "
                f"{final['ScalingIndex']}x what they did at the first step. "
                f"Expect this to get worse beyond the sizes tested.")
    elif not final["Note"]:
        logInfo(f"scaling held at full size: ScalingIndex "
                f"{final['ScalingIndex']} (1.0 is linear)")

    if AsDataFrame:
        try:
            import pandas
            frame = pandas.DataFrame(rows, columns=COLUMNS)
            return frame.drop(columns=["CsvMB"]) if not MeasureCsv else frame
        except ImportError:
            pass
    if not MeasureCsv:
        for row in rows:
            row.pop("CsvMB", None)
    return rows


# =====================================================================
# EXCEL INSTRUCTIONS -- Python in Excel (the PY function)
# =====================================================================
#
# WHAT YOU ARE MEASURING. Python in Excel does NOT run on your computer.
# Microsoft executes it in a container in Azure. The timings therefore
# describe Microsoft's hardware, not your laptop and not a CGS server. Use
# Excel to compare one approach against another; use a command line on the
# target machine to size that machine.
#
# WHAT THE SANDBOX FORBIDS. No files, no network, no pip install. That is
# why this module imports nothing outside the standard library and pandas,
# and why size is measured in memory rather than by writing a file. It also
# means you cannot "import cgs_ai" in Excel -- paste the code instead.
#
# STEP 1. Put your inputs in cells, so they can be changed without editing
#         any code:
#             A1: 1000000     (max rows)
#             B1: 12          (columns)
#             C1: 10          (increment %)
#
# STEP 2. Click an empty cell -- say A3 -- and type  =PY(  then paste the
#         WHOLE of this file's code, from "import io" down to the end of
#         testWorkload. Use Alt+Enter for new lines if you type it by hand.
#         (Simpler: keep the code in one PY cell and call it from another.)
#
# STEP 3. In the NEXT cell, call it, reading the inputs from the grid:
#
#             =PY( testWorkload(RowCount=int(xl("A1")),
#                               ColumnCount=int(xl("B1")),
#                               Increment=int(xl("C1"))) )
#
# STEP 4. The cell will show a DataFrame object. To spill the table across
#         the grid, set the cell's output to "Excel Value": click the
#         PY icon at the left of the formula bar and choose Excel Value, or
#         press Ctrl+Alt+Shift+M.
#
# STEP 5. Read the ScalingIndex column. 1.0 means ten times the data costs
#         ten times the time. 2.0 means a row at full size costs twice what
#         it cost at the first step -- the environment is NOT keeping up,
#         and the next doubling will hurt more than this one did.
#
# START SMALL. Try RowCount=50000 first and see how long it takes. Python
# in Excel has a per-session time limit, and a run that exceeds it is
# cancelled with nothing to show for it. Raise the number once you know
# roughly what it costs.
#
# =====================================================================


def buildArgParser() -> argparse.ArgumentParser:
    """Build the command-line parser. Parameters: none."""
    parser = argparse.ArgumentParser(
        prog="testWorkload.py",
        description="Measure how build and run time grow with dataset size.")
    parser.add_argument("--row-count", type=int, required=True)
    parser.add_argument("--column-count", type=int, required=True)
    parser.add_argument("--increment", type=int, default=10)
    parser.add_argument("--scale-by", default="rows", choices=VALID_SCALE_BY)
    parser.add_argument("--repeats", type=int, default=1)
    parser.add_argument("--seed", type=int, default=42)
    parser.add_argument("--measure-csv", action="store_true")
    parser.add_argument("--label", default="")
    parser.add_argument("--version", action="version",
                        version=f"testWorkload {__version__}")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    """CLI entry point. Parameters: argv (sequence[str]). Returns: exit code."""
    args = buildArgParser().parse_args(argv)
    try:
        rows = testWorkload(
            RowCount=args.row_count, ColumnCount=args.column_count,
            Increment=args.increment, ScaleBy=args.scale_by,
            Repeats=args.repeats, Seed=args.seed,
            MeasureCsv=args.measure_csv, AsDataFrame=False, Label=args.label)
    except ValueError as exc:
        logWarn(str(exc))
        return EXIT_CONFIG_ERROR

    environment = describeEnvironment()
    print("environment: " + "; ".join(f"{k}={v}" for k, v in environment.items()))
    headers = [c for c in COLUMNS if c != "CsvMB" or args.measure_csv]
    widths = {h: max(len(h), *(len(f"{row.get(h, '')}") for row in rows))
              for h in headers}
    print("  ".join(h.rjust(widths[h]) for h in headers))
    print("  ".join("-" * widths[h] for h in headers))
    for row in rows:
        print("  ".join(f"{row.get(h, '')}".rjust(widths[h]) for h in headers))
    return EXIT_OK


if __name__ == "__main__":
    sys.exit(main())
