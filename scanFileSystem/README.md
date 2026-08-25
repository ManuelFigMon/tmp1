# scanFileSystem

**v1.3.3** · A general-purpose file-system scanner and text-extraction utility.

`scanFileSystem.py` crawls one or more directory roots, captures filesystem
metadata for every matching file, extracts keywords with surrounding context,
optionally filters by a date range, and — via an **opt-in** metric profile —
parses structured performance metrics out of log files. The flagship profile
(`sas_log`) pulls per-step SAS `real time` / `cpu time`, but the same engine
generalizes to any keyword sweep or log-metric use case.

Built to run unattended under Windows Task Scheduler (or from a SAS `SYSTASK`
wrapper). It never prompts, logs everything to stderr, and exits non-zero on
failure so the scheduler can detect it.

**Dependencies are minimal by design (v1.3.3):** CSV output uses only the
Python standard library. XLSX output needs `openpyxl` (or `xlsxwriter`); with
neither installed the scan falls back to CSV with a warning. **pandas and numpy
are not used** — they were the source of an `Unable to import required
dependency numpy` failure and bought nothing over the stdlib `csv` module.

---

## Install

```bash
python -m venv .venv
.venv\Scripts\activate           # Windows;  source .venv/bin/activate on Linux/macOS
pip install -r requirements.txt
```

Run the self-checks:

```bash
pytest -q                        # 42 passed
```

## Running it in VS Code

Open the **`scanFileSystem` folder itself** as the workspace (not the repo
root) — the bundled `.vscode/` configs use `${workspaceFolder}` paths.

**One-time setup**

1. Install the Microsoft **Python** extension.
2. Create the venv and install deps (VS Code terminal, ``Ctrl+` ``):
   ```bash
   python -m venv .venv
   .venv\Scripts\activate            # Linux/macOS: source .venv/bin/activate
   pip install -r requirements.txt
   ```
3. `Ctrl+Shift+P` → **Python: Select Interpreter** → pick `./.venv`.

**Run it** — `F5`, or the Run and Debug panel (`Ctrl+Shift+D`), then pick:

| Configuration | What it does |
|---|---|
| `scan: fixtures -> xlsx (sas_log)` | Safe first run against the bundled fixtures → `out/scan.xlsx` |
| `scan: fixtures -> auto-named CSV` | Omits `--output-file-path` → `out/scan_YYYYMMDD_HHMMSS.csv` |
| `scan: .accdb/.mdb keyword sweep` | Pure keyword run, no metric profile |
| `scan: prompt for root + output` | Prompts for root/output/profile — use this for real UNC roots |
| `Python: current file` | Debugs the focused file |

Set breakpoints in the gutter and `F5` stops on them. Log lines go to stderr in
the integrated terminal.

> The prompting in `scan: prompt for root + output` is **VS Code's**, not the
> scanner's — `scanFileSystem.py` never prompts, so it stays safe to run
> unattended under Task Scheduler.

**Edit the arguments.** To change flags permanently, edit the `args` array in
`.vscode/launch.json`. Each flag and its value are separate strings:

```jsonc
"args": [
  "--input-folder-root", "\\\\A70admed.com\\r1\\CGS\\...\\HHH", "\\\\A70admed.com\\r1\\CGS\\...\\DME",
  "--output-file-path", "C:\\Logs\\scan.xlsx",
  "--metric-profile", "sas_log"
]
```

Backslashes must be doubled inside JSON, so a UNC root `\\server\share`
is written `"\\\\server\\share"`.

**Tests** — the Test Explorer (beaker icon) discovers the pytest suite
automatically; click ▶ to run all 42, or the ▶ beside a single test. Or
`Ctrl+Shift+P` → **Tasks: Run Task** → `run tests` / `install deps` /
`rebuild fixtures`.

**Just want a terminal?** ``Ctrl+` `` and run any command from the
[Examples](#examples) section directly.

## Running it from SAS

The optional SYSTASK wrapper is split in two:

| File | Contents |
|---|---|
| `sas/Run_scanFileSystem_v1.sas` | The `%scanFileSystem()` **macro definition** only, plus the `PYTHON_EXE` / `SCRIPT_PATH` configuration. |
| `sas/Examples_scanFileSystem_v1.sas` | **Example calls** that `%INCLUDE` the macro. All commented out by default. |
| `sas/Find_python_exe.sas` | `%findPython()` — locates `python.exe` and sets `PYTHON_EXE` for you. |

```sas
%include "C:\code\python\cgs_ai\scanFileSystem\sas\Run_scanFileSystem_v1.sas";
```

**List parameters** (`input_folder_root`, `folder_exclusion_list`,
`extract_keyword`, `file_exclusion_list`) take **semicolon-delimited** strings
wrapped in `%str()`. Python splits them back into a list — verified to produce
byte-identical output to passing the values as separate CLI arguments.

This command line:

```bat
python scanFileSystem.py ^
  --input-folder-root "\\A70admed.com\r1\...\UNIT\HHH\Old_Programs\Old_logs" "\\A70admed.com\r1\...\UNIT\DME\Logs" ^
  --output-file-path "C:\code\python\cgs_ai\tests\scanFileSystem\scan.xlsx" ^
  --metric-profile sas_log ^
  --extract-keyword "real time" "cpu time"
```

becomes this macro call (Example A in the examples file) — two roots joined by
`;`, two keywords joined by `;`:

```sas
%scanFileSystem(
  input_folder_root=%str(\\A70admed.com\r1\...\UNIT\HHH\Old_Programs\Old_logs;\\A70admed.com\r1\...\UNIT\DME\Logs),
  output_file_path=C:\code\python\cgs_ai\tests\scanFileSystem\scan.xlsx,
  metric_profile=sas_log,
  extract_keyword=%str(real time;cpu time)
);
```

The macro mirrors the Python validation (a missing `input_folder_root` aborts
before launching), and `%abort`s with the Python exit code when the scan fails
(`2` = config error, `3` = I/O error), so a failed scan fails the SAS job.

### Finding `python.exe`

`PYTHON_EXE` must point at the interpreter that has the dependencies — if you
installed into `.venv`, that is `<project>\.venv\Scripts\python.exe`, **not**
the system Python that `where python` finds first. The system one lacks
`openpyxl`, so `.xlsx` output would quietly fall back to CSV.

The authoritative answer comes from Python itself, run in the terminal where
your command already works:

```bat
python -c "import sys; print(sys.executable)"
```

From SAS, `%findPython()` probes the venv first, then `PATH`, then the `py`
launcher, then standard install locations, and validates that `openpyxl`
imports:

```sas
%include "C:\code\python\cgs_ai\scanFileSystem\sas\Find_python_exe.sas";
%findPython(project_root=C:\code\python\cgs_ai\scanFileSystem);
%put &=PYTHON_EXE;
```

`FILENAME PIPE` needs the `XCMD` option; check with
`%put %sysfunc(getoption(xcmd));`. Under `NOXCMD` the pipe probes are skipped
and only the filesystem probes run.

## Design — three separable concerns

This is **not** a SAS-only tool. The architecture keeps three concerns apart so
new use cases never touch the core:

| Layer | Responsibility |
|---|---|
| **Crawl / filter / metadata** | Walk roots, filter by extension, apply folder exclusions and the date range, capture metadata. Format-agnostic. |
| **Keyword + context extraction** | Works on any text file and any keyword list; returns the matched line, a ±3-line window, and a count. |
| **Pluggable metric profiles** | A profile parses structured metrics into StepDetail rows. Regex/config-driven, so a new profile is a registry entry — no changes to crawl or output code. |

Adding a profile means adding one entry to `METRIC_PROFILES`:

```python
"my_profile": RegexStepProfile(
    name="my_profile",
    step_pattern=r"^BEGIN\s+(?P<label>.+)$",
    metric_patterns={"real_time_sec": r"elapsed\s+(?P<value>[0-9:.]+)"},
    counter_patterns={"error_count": r"^\s*FATAL"},
)
```

## Parameters

Every parameter has a module-level default in the **CONFIG** block at the top of
`scanFileSystem.py`, and a matching kebab-case CLI flag that overrides it.

| Name (Python) | Type | Required | Default | Description |
|---|---|---|---|---|
| `input_folder_root` | list[str] | **Yes** | UNC default root | One or more root paths. Accepts a single string or an array. Empty → ERROR + non-zero exit. |
| `output_file_path` | str | No | *(auto)* | `.csv` → CSV (no index); `.xlsx` → Excel, by extension. A directory auto-names `scan_YYYYMMDD_HHMMSS.csv` inside it. **Omit entirely and the scan writes `scan_YYYYMMDD_HHMMSS.csv` to the current directory.** |
| `file_extensions` | list[str] | No | `["log","txt","sas"]` | Extensions to include. Case-insensitive, with or without a leading dot. |
| `include_subdirectories` | bool | No | `True` | `True` recurses; `False` is top level only. |
| `folder_exclusion_list` | list[str] | No | `[]` | Folder names/tokens to exclude, e.g. `["Old","Test"]`. **Empty default — nothing is excluded unless set.** Matches an ancestor directory segment (case-insensitive) or a full-path prefix. |
| `file_exclusion_list` | list[str] | No | `[]` | Prefixes/tokens stripped from the filename to derive `program_name`. |
| `extract_keyword` | list[str] | No | `[]` | Keywords to extract; each yields a matched line, a ±3-line context window, and a count. |
| `date_from` | str | No | `None` | Inclusive lower bound (`YYYY-MM-DD` or ISO datetime). |
| `date_to` | str | No | `None` | Inclusive upper bound. A bare date extends through end-of-day. |
| `date_field` | str | No | `"modified"` | Which timestamp the range filters on: `created` / `modified` / `accessed`. |
| `metric_profile` | str | No | `"none"` | `"none"` = off (no StepDetail at all); `"sas_log"` = per-step real/cpu time. |

**Validation.** Before any crawling, the scanner checks that `input_folder_root`
is present and non-empty, and that `metric_profile` / `date_field` are known
values. Any failure logs a specific `ERROR` naming the parameter and exits
non-zero — it never prompts.

**Output naming (v1.3.2).** `output_file_path` is optional. Omit it and the scan
writes `scan_YYYYMMDD_HHMMSS.csv` to the current directory; pass a directory and
the same name is generated inside it; pass a `.csv`/`.xlsx` path to control it
exactly. With a metric profile active the companion follows the same stem, e.g.
`scan_20260825_014044_StepDetail.csv`.

## Output

### Files grain — always, one row per file

`program_name`, `log_file_name`, `full_path`, `directory`, `extension`,
`file_size_bytes`, `created_time`, `modified_time`, `accessed_time` (ISO 8601),
`step_count`, `total_real_time_sec`, `total_cpu_time_sec`,
`max_step_real_time_sec`, `max_step_label`, `error_count`, `warning_count`,
`kw_<K>_line` / `kw_<K>_context` / `kw_<K>_count` per keyword, `parse_status`,
`scanned_at`.

Metric columns are zero/blank when `metric_profile="none"`.

### StepDetail grain — only when a profile is active

`full_path`, `program_name`, `step_index`, `step_label`, `real_time_sec`,
`cpu_time_sec`.

| Target | Behavior |
|---|---|
| `.xlsx` | Sheet `Files`; sheet `StepDetail` added when a profile is active. |
| `.csv` | Main file (Files); companion `<stem>_StepDetail.csv` when a profile is active (an INFO line names it). |
| `.xlsx`, no Excel engine | Falls back to CSV for the active grain(s) with a warning. |
| Engine order | `openpyxl` → `xlsxwriter` → CSV fallback. Written directly, no pandas. |
| `metric_profile="none"` | No StepDetail sheet or companion at all. |

`program_name` = filename minus extension, minus any prefix/token in
`file_exclusion_list` (case-insensitive).

## Examples

### A. Two roots (HHH + DME), SAS timing profile, Excel out

```bat
python scanFileSystem.py ^
  --input-folder-root "\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT\HHH" "\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT\DME" ^
  --output-file-path "C:\Logs\scan.xlsx" ^
  --metric-profile sas_log ^
  --extract-keyword "real time" "cpu time"
```

```sas
%scanFileSystem(
  input_folder_root=%str(\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT\HHH;\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT\DME),
  output_file_path=C:\Logs\scan.xlsx,
  metric_profile=sas_log,
  extract_keyword=%str(real time;cpu time)
);
```

### B. Access-database reference sweep (`.accdb` / `.mdb`), CSV out

A pure keyword use case — `metric_profile` stays `none`.

```bat
python scanFileSystem.py ^
  --output-file-path "C:\Logs\accdb_scan.csv" ^
  --extract-keyword ".accdb" ".mdb"
```

```sas
%scanFileSystem(
  output_file_path=C:\Logs\accdb_scan.csv,
  extract_keyword=%str(.accdb;.mdb)
);
```

### C. Date-range filter (files modified in H1 2026)

```bat
python scanFileSystem.py ^
  --output-file-path "C:\Logs\recent.csv" ^
  --date-from 2026-01-01 --date-to 2026-06-30 --date-field modified
```

### D. Defaults only

Uses the default root; `folder_exclusion_list` is empty so nothing is excluded;
`metric_profile=none`.

```bat
python scanFileSystem.py --output-file-path "C:\Logs\scan.csv"
```

### E. No output path at all — auto-named CSV

Writes `scan_YYYYMMDD_HHMMSS.csv` (plus a matching `_StepDetail.csv` when a
profile is active) into the current directory.

```bat
python scanFileSystem.py --input-folder-root "\\srv\logs"
```

```
INFO  output_file_path not supplied; writing scan_20260825_014044.csv
INFO  wrote 13 Files row(s) to scan_20260825_014044.csv
INFO  wrote 9 StepDetail row(s) to companion scan_20260825_014044_StepDetail.csv
```

> The UNC roots above are not reachable from a dev box. To try the examples,
> point `--input-folder-root` at `tests/fixtures/logs`.

## Sample output

**Files** (from the fixture tree, `--metric-profile sas_log`):

| program_name | log_file_name | ext | size | modified_time | step_count | total_real | total_cpu | max_step_label | err | warn | parse_status |
|---|---|---|---|---|---|---|---|---|---|---|---|
| jobA | jobA.log | log | 690 | 2026-03-15T00:00:00 | 3 | 3.25 | 2.43 | PROCEDURE SORT | 1 | 2 | OK |
| jobB | jobB.log | log | 325 | 2026-05-20T00:00:00 | 2 | 63.15 | 60.05 | PROCEDURE SQL | 0 | 0 | OK |
| broken_link | broken_link.log | log | 0 | | 0 | 0.00 | 0.00 | | 0 | 0 | `stat error: FileNotFoundError: ...` |

**StepDetail** (jobA):

| full_path | program_name | step_index | step_label | real_time_sec | cpu_time_sec |
|---|---|---|---|---|---|
| .../jobA.log | jobA | 1 | DATA statement | 0.05 | 0.03 |
| .../jobA.log | jobA | 2 | PROCEDURE MEANS | 1.20 | 0.90 |
| .../jobA.log | jobA | 3 | PROCEDURE SORT | 2.00 | 1.50 |

**Keyword extraction** (`--extract-keyword ".accdb"`, file `notes.txt`):

```
kw_accdb_count = 2
kw_accdb_line  = The legacy tracker lives in S:/shared/claims.accdb today.
kw_accdb_context =
    Migration notes
    The legacy tracker lives in S:/shared/claims.accdb today.
    A second extract still reads Archive.mdb every Monday.
    Both should move to Snowflake in Q3.
    Contact the DME team before touching claims.accdb again.
```

## Robustness & exit codes

- `pathlib` throughout; case-insensitive, trailing-slash-safe path comparisons.
- Files are read UTF-8, falling back to latin-1, then UTF-8 with replacement.
- Per-file `try/except`: a failure still emits the Files row with the error in
  `parse_status` and the scan continues.
- Missing roots and permission errors are logged and skipped, not fatal.
- A file whose `stat()` fails is emitted with a non-OK `parse_status` rather
  than silently dropped — including under a date filter, since its timestamp is
  unknowable.

| Code | Meaning |
|---|---|
| `0` | Success |
| `2` | Config error — missing `input_folder_root`, unknown profile/date_field, bad or inverted date |
| `3` | I/O error — no reachable root, unwritable output |

The output path is resolved (and its parent created) **before** crawling, so a
bad output path fails in seconds rather than after a long network scan.

### Troubleshooting

| Symptom | Cause / fix |
|---|---|
| `Unable to import required dependency numpy` | A broken pandas install in an older version. v1.3.3 removed pandas entirely — pull the latest and re-run. |
| `no Excel engine ... falling back to CSV` | `pip install openpyxl` if you want `.xlsx`; otherwise the CSV output is complete. |
| `none of the supplied input_folder_root path(s) are reachable` | UNC path typo, or the share isn't mounted for the account running the scan. |

## Layout

```
scanFileSystem/
├── scanFileSystem.py            # primary deliverable
├── requirements.txt
├── README.md
├── CLAUDE.md
├── .vscode/                        # launch/tasks/settings for VS Code
├── sas/
│   ├── Run_scanFileSystem_v1.sas      # macro definition (SYSTASK wrapper)
│   └── Examples_scanFileSystem_v1.sas # example calls (%INCLUDEs the macro)
└── tests/
    ├── make_fixtures.py             # generates the synthetic tree
    ├── fixtures/logs/...            # .log/.txt/.sas + Old/ Test/ Older/ + dated + malformed
    └── test_scanFileSystem.py       # 42 self-checks
```
