# scanFileSystem

**v1.3.1** · A general-purpose file-system scanner and text-extraction utility.

`scanFileSystem.py` crawls one or more directory roots, captures filesystem
metadata for every matching file, extracts keywords with surrounding context,
optionally filters by a date range, and — via an **opt-in** metric profile —
parses structured performance metrics out of log files. The flagship profile
(`sas_log`) pulls per-step SAS `real time` / `cpu time`, but the same engine
generalizes to any keyword sweep or log-metric use case.

Built to run unattended under Windows Task Scheduler (or from a SAS `SYSTASK`
wrapper). It never prompts, logs everything to stderr, and exits non-zero on
failure so the scheduler can detect it.

---

## Install

```bash
python -m venv .venv
.venv\Scripts\activate           # Windows;  source .venv/bin/activate on Linux/macOS
pip install -r requirements.txt
```

Run the self-checks:

```bash
pytest -q                        # 40 passed
```

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
| `output_file_path` | str | **Yes** | *(none)* | `.csv` → CSV (no index); `.xlsx` → Excel, by extension. A directory auto-names `scanFileSystem_YYYYMMDD_HHMMSS.csv` inside it. Empty → ERROR + non-zero exit. |
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
and `output_file_path` are present and non-empty, and that `metric_profile` /
`date_field` are known values. Any failure logs a specific `ERROR` naming the
parameter and exits non-zero — it never prompts. (A directory given as
`output_file_path` is valid; only empty/None is an error.)

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
| `2` | Config error — missing required parameter, unknown profile/date_field, bad or inverted date |
| `3` | I/O error — no reachable root, unwritable output |

## Layout

```
scanFileSystem/
├── scanFileSystem.py            # primary deliverable
├── requirements.txt
├── README.md
├── CLAUDE.md
├── sas/Run_scanFileSystem_v1.sas   # optional SYSTASK wrapper
└── tests/
    ├── make_fixtures.py             # generates the synthetic tree
    ├── fixtures/logs/...            # .log/.txt/.sas + Old/ Test/ Older/ + dated + malformed
    └── test_scanFileSystem.py       # 40 self-checks
```
