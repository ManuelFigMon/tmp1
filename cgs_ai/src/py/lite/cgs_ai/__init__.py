"""
=====================================================================
  Program Name  : __init__.py  (cgs_ai LITE)
  Author        : Manuel Figallo
  Purpose       : A single-file, drop-in build of cgs_ai carrying only the
                  two functions an end user needs to turn a CSV into a
                  formatted workbook and tell somebody it is ready.
  Version       : 1.1beta-lite
  Created       : 2026-08-28
  Last Modified : 2026-10-28

  Dependencies:
    STANDARD LIBRARY ONLY to load. openpyxl is imported lazily by formatCSV
    and is the single third-party requirement; sendEmail needs nothing but
    smtplib and email, both standard library.

  Description:
    The full package is a folder tree that has to be on sys.path. This file
    is deliberately ONE file with NO internal imports, so it can be executed
    straight off a UNC share:

        import runpy
        runpy.run_path(r"\\\\server\\...\\src\\py\\lite\\cgs_ai\\__init__.py",
                       run_name="__init__")
        from cgs_ai import formatCSV, sendEmail

    The second line works because running this file registers a module named
    "cgs_ai" in sys.modules (see REGISTRATION at the bottom). If the full
    cgs_ai package is already imported, this step is skipped and the full
    package wins -- the lite build never shadows it.

    An equally valid pattern, which does not rely on that registration:

        ns = runpy.run_path(PATH, run_name="cgs_ai_lite")
        formatCSV, sendEmail = ns["formatCSV"], ns["sendEmail"]

  Function Index:
    doBasicHello        - fixed greeting; proves the import worked
    doPersonalizedHello - greeting addressed to a name you pass in
    doDetailedHello     - greeting returned as a dict, in three styles
    formatData - CSV to a styled Excel workbook (SAS ODS look and feel);
                 FormatType corporate | corporatev2 | plain | minimal.
                 The full package also reads .xlsx; this build does not --
                 see the note in the docstring.
    formatCSV  - deprecated alias for formatData
    sendEmail  - SMTP notification to one or many recipients
    surveySelect - reproducible random sample of n rows from a DataFrame
    testWorkload - measure how this environment scales as a dataset grows

  Change Log:
    v1.1beta-lite - ADDED doDetailedHello, surveySelect and testWorkload.
                    Nothing else in
                    this file changed: every function that was here
                    before is byte-for-byte what it was.
    v1.0beta-lite - Renamed formatCSV to formatData, keeping formatCSV as
                    an alias.
                  - First lite build. formatCSV and sendEmail only, with
                    every src.utils helper inlined so the file stands alone.
                    Adds the corporatev2 FormatType and saves the workbook
                    through BytesIO for Snowflake workspace compatibility.
=====================================================================
"""

from __future__ import annotations

import csv
import io
import platform
import random
import smtplib
import statistics
import string
import sys
import time
import types
from email.message import EmailMessage
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Sequence

__version__ = "1.1beta-lite"
__all__ = ["doBasicHello", "doPersonalizedHello", "doDetailedHello",
           "formatData",
           "formatCSV", "sendEmail", "surveySelect", "testWorkload",
           "__version__"]

# --------------------------------------------------------------------- #
# Inlined helpers. The full package imports these from src.utils; the lite
# build carries its own copies so the file has no internal dependencies.
# --------------------------------------------------------------------- #

#: Palette per format type: banner, header, stripe and font colours.
FORMAT_STYLES: Dict[str, Dict[str, str]] = {
    "corporate":   {"banner": "1F3864", "header": "2E75B6", "stripe": "DCE6F1",
                    "headerFont": "FFFFFF", "bannerFont": "FFFFFF"},
    "corporatev2": {"banner": "1F3864", "header": "2E75B6", "stripe": "DCE6F1",
                    "headerFont": "FFFFFF", "bannerFont": "FFFFFF"},
    "plain":       {"banner": "FFFFFF", "header": "D9D9D9", "stripe": "FFFFFF",
                    "headerFont": "000000", "bannerFont": "000000"},
    "minimal":     {"banner": "FFFFFF", "header": "FFFFFF", "stripe": "FFFFFF",
                    "headerFont": "000000", "bannerFont": "000000"},
}
#: Format types that get zebra striping on the data rows.
STRIPED_FORMATS = frozenset({"corporate", "corporatev2"})
DEFAULT_FORMAT_TYPE = "corporate"
MAX_COLUMN_WIDTH = 60
DEFAULT_SMTP_SERVER = "smtp.example.com"
DEFAULT_SMTP_PORT = 25


def logInfo(message: str) -> None:
    """Write an informational line to stderr. Parameters: message (str)."""
    print(f"[cgs_ai-lite] INFO  {message}", file=sys.stderr)


def logError(message: str) -> None:
    """Write an error line to stderr. Parameters: message (str)."""
    print(f"[cgs_ai-lite] ERROR {message}", file=sys.stderr)


def asList(value: Any) -> List[str]:
    """Normalise a value into a list of non-empty strings.

    Parameters:
        value - a list/tuple, a ';'-delimited string, a scalar, or None.
    Returns: list of stripped, non-empty strings.
    """
    if value is None:
        return []
    if isinstance(value, (list, tuple, set)):
        items = [str(item) for item in value]
    else:
        items = str(value).split(";")
    return [item.strip() for item in items if str(item).strip()]


def ensureParent(path: str) -> Path:
    """Create the parent directory of `path` if needed. Returns: the Path."""
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    return target


def readCsv(inputPath: str) -> List[Dict[str, str]]:
    """Read a CSV into a list of dicts.

    Parameters: inputPath (str) - the CSV to read.
    Returns: list of row dicts keyed by column name.
    Raises: OSError if the file cannot be read.
    """
    with open(inputPath, encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


# --------------------------------------------------------------------- #
# Greetings -- the smoke test a newcomer runs first
# --------------------------------------------------------------------- #

def doBasicHello() -> str:
    """Return a fixed greeting, to prove the import worked.

    Parameters: none.
    Returns: str "Hello, World!".

    The first line of cgs_ai anybody runs. It touches nothing outside the
    package, so if this prints, the share is reachable, sys.path is right
    and the import succeeded -- which is most of what goes wrong on a new
    machine.
    """
    return "Hello, World!"


def doPersonalizedHello(name: str) -> str:
    """Return a greeting addressed to `name`.

    Parameters: name (str) - who to greet; blank falls back to "World".
    Returns: str e.g. "Hello, Alice!".

    Shows that a function can take an ARGUMENT and give a different answer
    back, which is the whole idea behind every other function in cgs_ai.
    """
    cleaned = str(name).strip() if name else "World"
    return f"Hello, {cleaned}!"


def doDetailedHello(style: str = "friendly") -> Dict[str, str]:
    """Return a structured greeting rather than a bare string.

    Parameters:
        style (str) - friendly (default) | formal | pirate. An unknown style
                      falls back to friendly rather than raising: a greeting
                      is not worth stopping a script over.
    Returns:
        dict with "message" (the greeting) and "style_used" (the style as it
        was asked for, NOT the one that was applied -- so a typo is visible
        in the output instead of being silently swallowed).

    The third greeting, and the one that makes the point: a function can hand
    back a STRUCTURE, not just text. Every real cgs_ai function returns a
    dict like this, so the caller can read one field without parsing a
    sentence -- result["message"] rather than hunting for a substring.
    """
    styles = {
        "friendly": "Hello, World! Wonderful day, isn't it?",
        "formal": "Greetings, World. It is a pleasure to connect.",
        "pirate": "Ahoy, World! Avast ye scallywags!",
    }
    greeting = styles.get(str(style).lower(), styles["friendly"])
    return {"message": greeting, "style_used": style}


# --------------------------------------------------------------------- #
# formatData  (formatCSV is kept as an alias)
# --------------------------------------------------------------------- #

def formatData(InputPath: str = "", OutputExcelPath: str = "",
               FormatType: str = DEFAULT_FORMAT_TYPE,
               SheetName: str = "Report",
               Title: str = "",
               InputCsvPath: str = "") -> Dict[str, Any]:
    """Render a CSV as a styled Excel workbook.

    Parameters:
        InputPath (str)       - REQUIRED source CSV.
        OutputExcelPath (str) - REQUIRED destination .xlsx.
        FormatType (str)      - corporate (default) | corporatev2 |
                                plain | minimal. "corporate" and
                                "corporatev2" are a navy banner, blue header
                                and zebra striping; they render identically,
                                with corporatev2 reserved as the versioned
                                name for that look.
        SheetName (str)       - worksheet name; default "Report".
        Title (str)           - banner text; defaults to the input filename.
        InputCsvPath (str)    - DEPRECATED alias for InputPath.
    Returns:
        dict with OutputExcelPath, RowCount, ColumnCount, FormatType and
        InputPath.
    Raises:
        ValueError  - a required parameter is missing, FormatType is unknown,
                      or an .xlsx was passed (see below).
        ImportError - openpyxl is not installed (message says how to fix).
        OSError     - the CSV cannot be read.

    LITE BUILD LIMITATION. The full package's formatData also READS .xlsx.
    This build deliberately does not: it is one self-contained teaching file,
    and an .xlsx reader is a hundred lines of zip and XML parsing that would
    bury the part a trainee is meant to read. Passing one raises a message
    saying where to get the full version rather than failing obscurely.

    Use in claims processing:
        Turn a claims extract into a report an analyst or manager can open
        directly, without hand-formatting it in Excel every cycle.
    """
    InputPath = str(InputPath or InputCsvPath or "").strip()
    if not InputPath:
        raise ValueError("required parameter 'InputPath' is missing or empty")
    if Path(InputPath).suffix.lower() in (".xlsx", ".xlsm"):
        raise ValueError(
            f"the lite build of formatData reads CSV only, and {InputPath} is "
            f"a workbook. Use the full package (src/py/formatData.py), which "
            f"reads .xlsx, or export the sheet to CSV first.")
    if not OutputExcelPath or not str(OutputExcelPath).strip():
        raise ValueError("required parameter 'OutputExcelPath' is missing or empty")
    if FormatType not in FORMAT_STYLES:
        raise ValueError(f"unknown FormatType {FormatType!r}; expected one of: "
                         f"{', '.join(FORMAT_STYLES)}")
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter
    except ImportError:
        raise ImportError(
            "formatData requires openpyxl for Excel styling. Install it with "
            "'pip install openpyxl'.")

    rows = readCsv(InputPath)
    columns = list(rows[0].keys()) if rows else []
    style = FORMAT_STYLES[FormatType]
    banner = Title or Path(InputPath).stem

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = SheetName

    # Row 1: title banner spanning the table width.
    cell = sheet.cell(row=1, column=1, value=banner)
    cell.font = Font(bold=True, size=14, color=style["bannerFont"])
    cell.fill = PatternFill("solid", fgColor=style["banner"])
    cell.alignment = Alignment(vertical="center")
    sheet.row_dimensions[1].height = 26
    if columns:
        sheet.merge_cells(start_row=1, start_column=1,
                          end_row=1, end_column=len(columns))

    # Row 2: column headers.
    for index, column in enumerate(columns, start=1):
        header = sheet.cell(row=2, column=index, value=column)
        header.font = Font(bold=True, color=style["headerFont"])
        header.fill = PatternFill("solid", fgColor=style["header"])
        header.alignment = Alignment(vertical="center", wrap_text=True)

    # Rows 3+: data with zebra striping.
    stripe = PatternFill("solid", fgColor=style["stripe"])
    for rowIndex, row in enumerate(rows, start=3):
        for colIndex, column in enumerate(columns, start=1):
            dataCell = sheet.cell(row=rowIndex, column=colIndex,
                                  value=row.get(column, ""))
            if FormatType in STRIPED_FORMATS and rowIndex % 2 == 1:
                dataCell.fill = stripe

    if columns:
        sheet.freeze_panes = "A3"
        sheet.auto_filter.ref = (f"A2:{get_column_letter(len(columns))}"
                                 f"{max(2, len(rows) + 2)}")
        for index, column in enumerate(columns, start=1):
            widest = max([len(str(column))] +
                         [len(str(r.get(column, ""))) for r in rows[:200]] or [10])
            sheet.column_dimensions[get_column_letter(index)].width = \
                min(MAX_COLUMN_WIDTH, max(10, widest + 2))

    ensureParent(OutputExcelPath)
    # Save via BytesIO: the Snowflake workspace filesystem does not support
    # seek() on writable files, which openpyxl's ZIP writer requires. Writing
    # to an in-memory buffer and flushing it bypasses that limit, and behaves
    # identically everywhere else.
    buffer = io.BytesIO()
    workbook.save(buffer)
    with open(OutputExcelPath, "wb") as handle:
        handle.write(buffer.getvalue())
    logInfo(f"formatted {len(rows)} row(s) x {len(columns)} column(s) "
            f"[{FormatType}] -> {OutputExcelPath}")
    return {"OutputExcelPath": OutputExcelPath, "RowCount": len(rows),
            "ColumnCount": len(columns), "FormatType": FormatType,
            "InputPath": InputPath}


def formatCSV(InputCsvPath: str = "", OutputExcelPath: str = "",
              FormatType: str = DEFAULT_FORMAT_TYPE,
              SheetName: str = "Report", Title: str = "",
              **kwargs: Any) -> Dict[str, Any]:
    """DEPRECATED alias for formatData, kept so existing callers keep working.

    Parameters: as formatData, with the original parameter name InputCsvPath.
    Returns: whatever formatData returns.

    Not removed and not warned about on every call: this name is in the
    training pipelines and the handouts that go with them.
    """
    return formatData(InputPath=InputCsvPath, OutputExcelPath=OutputExcelPath,
                      FormatType=FormatType, SheetName=SheetName, Title=Title,
                      **kwargs)


# --------------------------------------------------------------------- #
# sendEmail
# --------------------------------------------------------------------- #

def sendEmail(To: Any, From: str, Subject: str, Body: str,
              SmtpServer: Optional[str] = None, Port: Optional[int] = None,
              Html: bool = False, Attachments: Any = ()) -> Dict[str, Any]:
    """Send an email alert over SMTP.

    Parameters:
        To (list|str)     - REQUIRED recipient(s); ';'-delimited string allowed.
        From (str)        - REQUIRED sender address.
        Subject (str)     - REQUIRED subject line.
        Body (str)        - REQUIRED message body.
        SmtpServer (str)  - SMTP host; defaults to smtp.example.com.
        Port (int)        - SMTP port; defaults to 25.
        Html (bool)       - send the body as HTML instead of plain text.
        Attachments       - optional file path(s) to attach.
    Returns:
        dict with To (list), Subject, SmtpServer, Port and Sent (bool).
    Raises:
        ValueError            - a required parameter is missing.
        smtplib.SMTPException - the server rejected the message.
        OSError               - the server is unreachable.

    Use in claims processing:
        Notify the operations mailbox when a claims report is ready,
        including the row count and the output location, so nobody has to
        watch the job.

    NOTE: the lite build has no .env loader, so SmtpServer and Port are NOT
    read from configuration -- pass them explicitly, or you get the
    smtp.example.com default, which will not deliver anything.
    """
    recipients: List[str] = asList(To)
    if not recipients:
        raise ValueError("required parameter 'To' is missing or empty")
    for name, value in (("From", From), ("Subject", Subject), ("Body", Body)):
        if value is None or str(value).strip() == "":
            raise ValueError(f"required parameter '{name}' is missing or empty")

    server = SmtpServer or DEFAULT_SMTP_SERVER
    port = int(Port or DEFAULT_SMTP_PORT)

    message = EmailMessage()
    message["To"] = ", ".join(recipients)
    message["From"] = From
    message["Subject"] = Subject
    message.set_content(Body, subtype="html" if Html else "plain")

    for attachment in asList(Attachments):
        path = Path(attachment)
        if not path.is_file():
            logError(f"attachment not found, skipping: {path}")
            continue
        message.add_attachment(path.read_bytes(), maintype="application",
                               subtype="octet-stream", filename=path.name)

    logInfo(f"sending mail to {len(recipients)} recipient(s) via {server}:{port}")
    with smtplib.SMTP(server, port, timeout=30) as smtp:
        smtp.send_message(message)
    logInfo(f"sent: {Subject}")
    return {"To": recipients, "Subject": Subject, "SmtpServer": server,
            "Port": port, "Sent": True}


# --------------------------------------------------------------------- #
# surveySelect
# --------------------------------------------------------------------- #

#: The default seed. A fixed, published number rather than a random one: the
#: point of the function is that two people running it get the same rows.
DEFAULT_SURVEY_SEED = 8111305759926929283


def surveySelect(data: Any, n: int, out: str = "",
                 seed: int = DEFAULT_SURVEY_SEED) -> Any:
    """Draw a reproducible random sample of n rows.

    Parameters:
        data  - the rows to sample: a pandas DataFrame, or anything
                DataFrame() accepts (a dict of lists, a list of dicts). From
                Excel, pass xl("Claims_Data[#All]", headers=True).
        n (int)    - how many sample units to draw.
        out (str)  - optional .csv or .xlsx path to write the sample to. The
                     sample is RETURNED either way, so in Excel -- which has
                     no filesystem -- leave it out entirely.
        seed (int) - the number that makes the draw repeatable. Change it and
                     you get a different sample; keep it and you get the same
                     one every time, on any machine. Pass 0 or None for a
                     genuinely random draw.
    Returns:
        A DataFrame of n rows, indexed 1..n the way SAS numbers a sample.
    Raises:
        ValueError  - n is larger than the data, or not at least 1.
        ImportError - pandas or numpy is missing (message says how to fix).

    PCG64DXSM is a bit generator from the PCG family. It is used instead of
    the default because it is explicitly seeded and its stream is
    reproducible across platforms and numpy versions -- which is what an
    auditable sample needs. A sample nobody else can reproduce is an
    anecdote, not evidence.

    Use in claims processing:
        Pull the sample a medical reviewer or an auditor will work through,
        and be able to produce exactly the same sample again months later
        when somebody asks how those claims were chosen.
    """
    try:
        import pandas as pd
        from numpy.random import PCG64DXSM
    except ImportError:
        raise ImportError(
            "surveySelect requires pandas and numpy. Install them with "
            "'pip install pandas numpy'. Both are already present inside "
            "Python in Excel and inside a Snowflake notebook.")

    n = int(n)
    if n < 1:
        raise ValueError(f"n must be at least 1, got {n}")
    seed = seed if seed else random.randint(0, sys.maxsize)
    frame = data if isinstance(data, pd.DataFrame) else pd.DataFrame(data)
    if n > len(frame):
        raise ValueError(f"asked for {n} sample units but the data has only "
                         f"{len(frame)} rows")

    generator = PCG64DXSM(seed)
    sample = frame.sample(n=n, random_state=generator).reset_index(drop=True)
    sample.index = sample.index + 1          # 1-based, as SAS numbers them

    if out:
        if str(out).lower().endswith((".xlsx", ".xlsm")):
            sample.to_excel(out, index=True)
        else:
            sample.to_csv(out, index=True)
        logInfo(f"sampled {n} of {len(frame)} row(s) [seed {seed}] -> {out}")
    else:
        logInfo(f"sampled {n} of {len(frame)} row(s) [seed {seed}]")
    return sample


# --------------------------------------------------------------------- #
# testWorkload
#
# WHAT IT ANSWERS. Not "how fast is this machine" -- a single timing tells
# you almost nothing. It answers "does ten times the data cost ten times the
# time, or more?", which is the question behind every capacity decision. The
# headline column is ScalingIndex: 1.0 means linear, 2.0 means a row costs
# twice as much to process at full size as it did at the first step.
#
# RUNNING IT IN EXCEL. Python in Excel executes in a Microsoft-managed
# container in Azure, NOT on your laptop and NOT on a CGS server. The figures
# describe Microsoft's container. That is useful for comparing one algorithm
# against another and useless for sizing a CGS server -- for that, run this
# on the CGS machine from a command line and compare the two tables.
#
# START SMALL. RowCount=50000 first. Python in Excel has a per-session time
# limit, and a run that exceeds it is cancelled with nothing to show.
# --------------------------------------------------------------------- #

#: testWorkload output columns, in order.
WORKLOAD_COLUMNS = [
    "Label", "StepPct", "Rows", "Columns", "Cells",
    "CreateSeconds", "MemoryMB", "WorkSeconds",
    "CreateSecPerMCell", "WorkSecPerMCell",
    "ScalingIndex", "CellsPerSecond", "Note",
]

#: Below this, a timing is mostly scheduler noise and a ScalingIndex built
#: from it means nothing. Carried as a NOTE ON THE ROW rather than a log
#: line, because inside Excel stderr is invisible: a caveat that is not in
#: the table is a caveat the reader never sees.
NOISE_FLOOR_SECONDS = 0.05

#: The repeating column shape of the synthetic data. Mixed types on purpose:
#: a grid of identical integers is far cheaper to build, hold and process
#: than real data, and would flatter the environment under test.
COLUMN_KINDS = ("id", "code", "amount", "date", "flag", "text")

VALID_SCALE_BY = ("rows", "columns", "both")


def columnPlan(ColumnCount: int) -> List[Dict[str, str]]:
    """Decide the name and type of each synthetic column.

    Parameters: ColumnCount (int) - how many columns.
    Returns: list of {"name", "kind"} in grid order.
    """
    return [{"name": f"{COLUMN_KINDS[i % len(COLUMN_KINDS)]}_{i + 1}",
             "kind": COLUMN_KINDS[i % len(COLUMN_KINDS)]}
            for i in range(ColumnCount)]


def buildDataset(Rows: int, Plan: Sequence[Dict[str, str]],
                 Seed: int = 42) -> Dict[str, List[Any]]:
    """Build a synthetic dataset as a dict of column lists.

    Parameters:
        Rows (int)      - how many rows.
        Plan (sequence) - from columnPlan().
        Seed (int)      - seeded, so two runs are comparable.
    Returns: dict mapping column name -> list of values.

    A dict of lists rather than a list of dicts: it is the shape pandas takes
    directly, and it does not pay a per-row dictionary overhead that would be
    measured as if it were the environment's fault.
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
                          f"{generator.randint(1, 28):02d}"
                          for _ in range(Rows)]
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

    deep=True on a DataFrame so Python string objects are counted; without it
    a column of text reports as 8 bytes a row, which is the pointer, not the
    string, and understates a text-heavy table by an order of magnitude.
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
        # Sample rather than walk millions of objects: the point is to
        # measure the workload, not to be the workload.
        sample = values[:1000]
        if sample:
            total += int(sum(sys.getsizeof(v) for v in sample)
                         / len(sample) * len(values))
    return round(total / 1_048_576, 3)


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
        percentages.append(100)              # always finish at full size
    steps = []
    for pct in percentages:
        rows = (RowCount if ScaleBy == "columns"
                else max(1, round(RowCount * pct / 100)))
        columns = (ColumnCount if ScaleBy == "rows"
                   else max(1, round(ColumnCount * pct / 100)))
        steps.append({"pct": pct, "rows": rows, "columns": columns})
    return steps


def describeEnvironment() -> Dict[str, Any]:
    """Describe the machine the measurements came from.

    Parameters: none.
    Returns: dict - python version, platform, processor, and the pandas
             version, which changes the numbers materially.

    Always report this next to the timings. A table of seconds with no note
    of what produced them cannot be compared with anything later.
    """
    try:
        import pandas
        pandasVersion = pandas.__version__
    except ImportError:
        pandasVersion = "not installed"
    return {"Python": platform.python_version(),
            "Platform": platform.platform(),
            "Processor": platform.processor() or "unknown",
            "Pandas": pandasVersion}


def testWorkload(RowCount: int, ColumnCount: int, Increment: int = 10,
                 Work: Optional[Callable[[Any], Any]] = None,
                 ScaleBy: str = "rows", Repeats: int = 1, Seed: int = 42,
                 AsDataFrame: bool = True, Label: str = "") -> Any:
    """Measure how build time and work time grow as a dataset grows.

    Parameters:
        RowCount (int)     - REQUIRED rows at 100%, the maximum size tested.
        ColumnCount (int)  - REQUIRED columns in the synthetic dataset.
        Increment (int)    - percentage step; default 10, giving 10%, 20% ...
                             100%. 25 gives 25%, 50%, 75%, 100%.
        Work (callable)    - the function to time at each step; it is called
                             with the dataset and its result is ignored.
                             Default: a group-by-and-aggregate.
        ScaleBy (str)      - rows | columns | both; default rows. "both"
                             grows the grid in two directions, so cells rise
                             with the SQUARE of the step.
        Repeats (int)      - timings per step, median kept; default 1. Raise
                             it to 3 or 5 when timings are small and noisy.
        Seed (int)         - reproducible data; default 42.
        AsDataFrame (bool) - return a DataFrame when pandas is present, else
                             a list of dicts.
        Label (str)        - a name carried on every row, so two
                             environments' tables can be stacked.
    Returns:
        One row per step with the WORKLOAD_COLUMNS above. ScalingIndex is the
        one to read: 1.0 is linear scaling, above 1.0 means each unit of data
        got MORE expensive as the dataset grew.
    Raises:
        ValueError - a parameter is out of range.

    Use in claims processing:
        Before promising a monthly run will cope with a year of claims,
        measure it at 10% and 100% of that year and look at ScalingIndex. A
        job that is linear will finish; one that is not will quietly stop
        finishing somewhere in between.

    LITE BUILD NOTE. The full module (src/py/testWorkload.py) also offers
    MeasureCsv and a command line. Those are left out here: this build is one
    self-contained file, and in Excel there is no filesystem to size a CSV
    against anyway.
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
    logInfo(f"testWorkload: {len(steps)} step(s) up to {RowCount:,} x "
            f"{ColumnCount}, scaling by {ScaleBy}")

    rows: List[Dict[str, Any]] = []
    baseline: Optional[float] = None
    for step in steps:
        plan = columnPlan(step["columns"])

        startedBuild = time.perf_counter()
        dataset = toFrame(buildDataset(step["rows"], plan, Seed))
        createSeconds = time.perf_counter() - startedBuild

        # Measured AFTER the build timer stops, so it never inflates it.
        memoryMb = measureMemoryMb(dataset)

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

    final = rows[-1]
    if not final["Note"] and final["ScalingIndex"] >= 1.5:
        logError(f"scaling is NOT linear: at full size each million cells "
                 f"cost {final['ScalingIndex']}x what they did at the first "
                 f"step. Expect this to get worse beyond the sizes tested.")

    if AsDataFrame:
        try:
            import pandas
            return pandas.DataFrame(rows, columns=WORKLOAD_COLUMNS)
        except ImportError:
            pass
    return rows


# --------------------------------------------------------------------- #
# REGISTRATION
#
# runpy.run_path() executes this file but does not create an importable
# module, so `from cgs_ai import formatCSV` would fail on the next line.
# Publishing a module object into sys.modules makes that import work.
#
# The guard matters: if the FULL cgs_ai package is already imported, it is
# left alone. The lite build never shadows the real one.
# --------------------------------------------------------------------- #

def _register(moduleName: str = "cgs_ai") -> None:
    """Publish this file's functions as an importable module.

    Parameters:
        moduleName (str) - the name to register under; default "cgs_ai".
    Returns: None. Does nothing if that name is already imported.
    """
    if moduleName in sys.modules:
        return
    module = types.ModuleType(moduleName)
    module.__doc__ = (f"cgs_ai lite build {__version__} "
                      "(3 greetings, formatData, sendEmail, surveySelect, "
                      "testWorkload)")
    module.__version__ = __version__
    module.__all__ = list(__all__)
    module.doBasicHello = doBasicHello
    module.doPersonalizedHello = doPersonalizedHello
    module.doDetailedHello = doDetailedHello
    module.formatData = formatData
    module.formatCSV = formatCSV
    module.sendEmail = sendEmail
    module.surveySelect = surveySelect
    module.testWorkload = testWorkload
    # The helpers testWorkload is built from, so a trainee can call one on
    # its own -- buildDataset to see the synthetic data, stepSizes to see
    # which sizes a given Increment will test, describeEnvironment to record
    # what produced a set of timings.
    module.buildDataset = buildDataset
    module.columnPlan = columnPlan
    module.defaultWorkload = defaultWorkload
    module.describeEnvironment = describeEnvironment
    module.stepSizes = stepSizes
    sys.modules[moduleName] = module


_register()
