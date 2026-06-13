# Claude Code Prompt — `Invoke-DataAnalysis.ps1`

> This is a build prompt that reproduces the current `Invoke-DataAnalysis.ps1`.
> Edit any section below and re-run it through Claude Code to regenerate a
> modified version. It is written to be self-contained: it does not assume the
> existing script is present, but it does note the lineage (the script began as
> a generalization of `Invoke-HiglasAnalysis.ps1`).

---

Write a production-quality PowerShell script named `Invoke-DataAnalysis.ps1` that
performs automated exploratory data analysis (EDA) on a CSV dataset and generates
a formatted report for business analysts. The objective is to produce an initial
analysis of the data as a discussion starting point for business analysts. The
script must be **generic and reusable across projects** — all project identity
comes from the `PROJ_NAME` parameter; no project-specific naming anywhere in the
code, console output, or report boilerplate.

It runs in a locked-down corporate Windows environment (no Python, no internet
access, no ability to install modules), so it must rely only on built-in
**Windows PowerShell 5.1** capabilities, **.NET Framework** classes, and
**Microsoft Office COM automation**. Target PowerShell 5.1 — avoid PowerShell
7-only syntax.

## Header / Attribution

Include a comment-based help block at the top with:

* `.SYNOPSIS` and `.DESCRIPTION` explaining that the script produces an initial
  automated analysis of a dataset to serve as a discussion starting point for
  business analysts, reusable across projects. The `.DESCRIPTION` should also
  summarize the two-pass memory design and the four analysis tiers.
* `.NOTES` stating: **Prepared by Manuel Figallo**. Also note it is based on
  `Invoke-HiglasAnalysis.ps1` (generalized and re-engineered for bounded-memory
  streaming to prevent `OutOfMemoryException`). List target platform
  (PowerShell 5.1 / .NET 4.x), dependencies (none beyond the OS; Word output
  needs Microsoft Word), and exit codes.
* `.PARAMETER` documentation for **every** parameter.
* `.EXAMPLE` blocks showing usage (include the sample call at the bottom of this
  prompt, plus one `v3` HTML example using `-MaxSampleRows`).

## Parameters

Fully parameter-driven via a `param()` block with validation:

1. `PROJ_NAME` (string, mandatory, `ValidateNotNullOrEmpty`) — project name used
   in the report title, output file names, and headers (e.g. `"HIGLAS"`).
2. `ANALYSIS_VERSION` (string, mandatory) — `ValidateSet` of `v0`, `v1`, `v2`,
   `v3`, `lite`. Treat input case-insensitively and map `lite` to `v1`.
3. `PATH_TO_DATA` (string, mandatory) — full path (UNC paths must be supported)
   to the input CSV. Validate existence with `ValidateScript` and fail with a
   clear, actionable error message (mention UNC support and read access) if not.
4. `FINAL_OUTPUT` (string, mandatory) — `ValidateSet` of `WORD` or `HTML`.
5. `OutputFolder` (string, optional) — defaults to the script's directory
   (`$PSScriptRoot`); where the report and chart images are written. Create it
   if missing.
6. `MaxSampleRows` (int, optional, default `100000`, `ValidateRange(1000,
   10000000)`) — the maximum number of rows held in memory for statistics,
   plots, and clustering. Lets users tune the memory/fidelity trade-off per
   machine. Document clearly that full-file figures are unaffected by it.

After the `param()` block: `Set-StrictMode -Off`, `$ErrorActionPreference =
'Stop'`, and a block of script-scoped constants/state including: OOM-retry floor
(`MinSampleRows = 10000`), per-scatter point cap (`MaxPlotPoints = 5000`), caps
on categorical columns in frequency analysis (`20`), bar charts (`5`), histograms
(`5`), distinct values tracked per column (`MaxDistinctTracked = 100000`), row
hashes tracked for duplicate detection (`MaxDupHashes = 1000000`), a charting-
available flag, holders for the top correlated pairs and the cluster result,
`List[string]` accumulators for executive-summary headlines and memory notes,
and cached `InvariantCulture` / `NumberStyles.Float` objects.

## Memory Management (OutOfMemoryException Prevention) — CRITICAL

The naive approach of loading the whole file with `Import-Csv` into an array of
`PSCustomObject`s causes `OutOfMemoryException` on large files. **Never load the
full file into PSObjects, and never call `Import-Csv` or `Get-Content` on the
data file.** Use a streaming, two-pass design:

* **Shared file-open helper** — a function that opens the CSV via a
  `System.IO.FileStream` with a **1 MB buffer** and `FileOptions::SequentialScan`
  (minimizes SMB round-trips on UNC paths), wraps it in a `StreamReader` (BOM
  detection on), and constructs a `Microsoft.VisualBasic.FileIO.TextFieldParser`
  for proper quoted-field handling. **Cast the parser constructor argument
  explicitly to `[System.IO.TextReader]`** so PowerShell 5.1 cannot bind a wrong
  overload. Configure delimited mode, the detected delimiter, quotes enabled,
  whitespace not trimmed. Return the stream/reader/parser together.

* **Pass 1 — streaming aggregation over the full file** (exact, bounded memory):
  row count, per-column null/blank counts, distinct-value counts (capped
  `HashSet`/dictionary; record `">100,000"` past the cap rather than growing),
  capped frequency counts per categorical column, **running min/max/mean/variance
  via Welford's online algorithm** for numeric columns, date min/max, per-column
  numeric/date parse tallies for **type inference by sampling** (numeric if
  ≥80% of non-null values parse as numbers, date if ≥80% parse as dates, else
  categorical; flag **mixed-type** when the leading ratio is 20–80%), an
  approximate **duplicate-row count** (hash of joined fields, capped at
  `MaxDupHashes`), and a **10-row preview**. Adaptively stop attempting numeric/
  date parses on columns that are clearly neither (after enough non-null samples)
  for speed. This yields exact full-file metadata/frequency/distribution figures
  without retaining rows.

* **Pass 2 — bounded reservoir sample** (seeded for reproducibility): build a
  reservoir sample of at most `MaxSampleRows` rows, **stored as plain
  `string[]` arrays, never PSObjects**, for analyses that genuinely need
  row-level data: percentiles/IQR outliers, correlation, all charts, pair plots,
  scatter plots, and k-means.

* **File-size handling > 2 GB** — file lengths exceed `Int32.MaxValue`, so
  **cast `FileStream.Length` to `[double]` before any arithmetic** (e.g. for the
  progress-percent denominator). Do **not** use `[math]::Max(1, $fs.Length)` —
  PS 5.1 binds the `Max(int,int)` overload and overflows on large files.

* **Transient network resilience** — UNC shares can drop mid-read. Provide a
  retry wrapper that retries an action on `IOException` (including a wrapped
  inner `IOException`) up to 3 attempts with linear backoff (5s, 10s), but
  **rethrows non-I/O exceptions immediately** (so `OutOfMemoryException` is never
  swallowed). Run **both passes** through this wrapper from the main flow.

* **OOM auto-retry on the sampling pass** — wrap Pass 2 so that if it throws
  `OutOfMemoryException`, automatically retry with `MaxSampleRows` halved (down to
  the `MinSampleRows = 10000` floor), logging the degradation into a memory-notes
  list that surfaces in the report's executive summary. Below the floor, fail.

* **Defensive measures** — at startup check `[Environment]::Is64BitProcess` and
  warn (with guidance to launch 64-bit PowerShell from System32, not SysWOW64)
  if 32-bit, recording a note for the report; build HTML with `StringBuilder`
  (not concatenation); dispose readers/streams/parsers in `finally`; dispose
  chart and bitmap objects right after each `SaveImage`; release large
  intermediates (`$var = $null`) and call `[GC]::Collect()` between major stages;
  report progress by **bytes read vs. file size** via `Write-Progress` during the
  streaming passes.

## File Signature Detection

Before Pass 1, sniff the file: read the first bytes to detect a BOM (UTF-8 /
UTF-16 LE / UTF-16 BE, else "No BOM (ANSI/UTF-8 assumed)"), and detect the
delimiter by counting candidates (`,` `;` tab `|`) in the header line. Report
both in the File Metadata section and use the detected delimiter for both passes.

## Analysis Versions

Cumulative (V1 ⊇ V0, V2 ⊇ V1, V3 ⊇ V2). Each analysis section is its own
function; a version-level switch (0/1/2/3) determines which run. Wrap **each
section** in try/catch so one failure logs into that section of the report
instead of killing the run. Every report **opens with an executive summary
page** listing the version run, scope/sampling notes, memory notes, and a table
of headline findings accumulated by the sections. **V0 produces no charts** and
must run fast; chart-only sections, when charting is unavailable, are replaced by
a stub section noting the charts were skipped.

### V0 — Data Profile (basic)
* **File metadata**: file path, size, last modified date, data row count, column
  count, detected delimiter, encoding/BOM, malformed-lines-skipped count.
* **Metadata analysis (column inventory)**: per column — name, inferred type
  (numeric/date/categorical), null/blank count and %, distinct value count.
* **Simple frequency analysis**: top-5 value counts per categorical column.
* **Simple distribution analysis**: min, max, mean, median per numeric column.
* **Data quality flags**: fully empty columns, constant-value columns,
  mixed-type columns, columns with ≥20% nulls, and an approximate duplicate-row
  count.
* **Sample preview** of the first 10 rows (cap displayed columns for readability).
* **Data readiness summary**: a **pass/warn/fail verdict per column** with a
  reason, so analysts can quickly decide if the data is fit for deeper analysis
  (FAIL = empty / ≥50% null / mixed-type; WARN = ≥20% null / constant /
  identifier-like; PASS otherwise), plus a PASS/WARN/FAIL tally.

### V1 (lite) — Descriptive Analysis (all of V0, plus)
* **Frequency analysis (full)**: top-15 value tables with counts and percentages.
* **Distribution analysis (full)**: min, max, mean (exact from Pass 1), plus
  median, 25th/75th percentiles, and 1.5×IQR outlier counts (from the sample),
  standard deviation (exact from Pass 1), and a **skew indicator** comparing
  mean vs. median.
* **Metadata analysis (full)**: extend the column inventory with **mode**, an
  in-memory **memory-footprint estimate**, and **date range** (min/max) for date
  columns.
* **Correlation matrix**: Pearson across all numeric columns as a formatted
  table; flag pairs with **|r| > 0.7**; remember the top-5 by |r| for V2.
* **Line graph**: if a date column exists, a line chart of record counts over
  time (auto day/month bucketing) and, if a numeric column exists, the sum of the
  first numeric column over time on a secondary axis.
* **Bar graphs**: for up to 5 categorical columns, top-10 value frequencies.
* **Histograms**: for up to 5 leading numeric columns (20 bins, from sample).
* **Missing-data bar chart**: null percentage by column.

### V2 — Relationship / Intermediate Analysis (all of V1, plus)
* **Scatter plots**: for the top-5 most correlated numeric pairs (by |r|), from
  the sample, point-capped.
* **Pair plots**: a single composed image — grid of pairwise scatter plots
  across up to the first 5 numeric columns, with histograms on the diagonal.
* **Grouped comparisons**: pick a suitable categorical grouping column (2–50
  distinct values); for the first ~3 numeric columns, show descriptive stats
  (count/min/P25/median/mean/P75/max — a box-plot-style quartile summary) broken
  down by the top categories, plus grouped bar charts of the per-category mean.
* **Correlation heatmap**: render the correlation matrix as a colored heatmap
  image (blue negative / white ~0 / red positive) using `System.Drawing`,
  in addition to the V1 table.
* **Month-over-month trends**: if a date column exists, a monthly table of record
  counts (and the first numeric column's sum) with **percentage change** vs. the
  prior month.

### V3 — Pattern Discovery / Advanced Analysis (all of V2, plus)
* **K-means clustering from scratch** (no external libraries) on the
  **standardized (z-score)** numeric columns, operating on the typed-array
  sample. Use k-means++ seeding and Lloyd iterations. Choose k via a simple
  **elbow heuristic over k = 2..6** (report WCSS per k; pick the largest k while
  each step still cuts WCSS by ≥15%). Report cluster sizes, per-cluster means for
  each numeric variable (in original units), and a scatter plot of the first two
  numeric variables colored by cluster. Cap clustering to the first 8 numeric
  columns for tractability.
  * **Important PS 5.1 gotcha**: PowerShell variables are case-insensitive, so do
    **not** name the cluster-count parameter `$K` while using `$k` as a loop
    variable — they collide. Use a distinct name (e.g. `$ClusterCount`).
* **Outlier / anomaly spotlight**: list the ~10 sampled records furthest (in
  standardized distance) from their assigned cluster centroid — an investigation
  shortlist (useful in a program-integrity context).
* **Cluster narrative table**: a plain-language description per cluster based on
  its distinguishing variables vs. the overall mean (e.g. "high amount, low
  frequency").
* **Z-score multivariate outlier count**: how many sampled records exceed 3 SD
  on at least one numeric variable.
* **Cluster composition over time**: if a date column exists, a monthly table of
  each cluster's share of rows.

## Charting

Render all charts with the .NET
`System.Windows.Forms.DataVisualization.Charting` assembly (load with `Add-Type`;
also load `System.Drawing`), save as PNG in the output folder, then embed into
the report. **Dispose the chart (and any bitmap) immediately after `SaveImage`**
to avoid GDI/memory buildup. If the charting assembly is unavailable, degrade
gracefully: skip charts, note it in the affected sections, continue with tables.
V0 produces no charts; only initialize charting for version ≥ 1.

## Report Output

* **WORD**: build a `.docx` via Word COM automation (`New-Object -ComObject
  Word.Application`): a **title page** (project name, analysis version, data file
  path, run timestamp, "Prepared by Manuel Figallo"), a **table of contents**,
  one heading per section, formatted bordered tables (shaded header row),
  embedded chart images (scaled to page width). Set `Visible=$false`,
  `DisplayAlerts=0`. **Quit and release all COM objects in a `finally` block**
  (`ReleaseComObject` on doc and app, then `GC::Collect`/`WaitForPendingFinalizers`).
  If Word is not installed/available, **fall back to HTML with a warning**.
* **HTML**: a self-contained, styled report assembled with `StringBuilder`, same
  structure (banner, TOC, sections, tables, images), embedding chart PNGs as
  **base64** so the file is portable. HTML-encode all dynamic text.

Output file name: `{PROJ_NAME}_Analysis_{ANALYSIS_VERSION}_{yyyyMMdd_HHmmss}.docx`
(or `.html`). Sanitize `PROJ_NAME` into a safe filename token.

## Robustness Requirements

* Detect numeric and date columns by sampling values during Pass 1 — never assume
  types.
* Wrap each analysis section in try/catch; log the error into the report section.
* Repair blank or duplicate header names (e.g. `Column3`, `Name_2`).
* Parse numbers tolerantly: invariant culture first, then a cleaned fallback
  ($, commas, %, and accounting-style `(123)` negatives), then current culture.
* Write progress to the console with `Write-Host`/`Write-Progress` for each
  stage (timestamped stage messages).
* Exit codes: `0` success; `1` unexpected fatal error (print message + stack
  trace); `2` data could not be loaded (file unreadable, empty, or no parseable
  columns).

## Deliverables

1. `Invoke-DataAnalysis.ps1` — the complete script.
2. An updated `README.md` covering parameters, the four version tiers and their
   differences, prerequisites, the memory-management design and `MaxSampleRows`
   tuning guidance, troubleshooting (including persistent memory pressure, UNC
   "network path was not found" / I/O errors, Word fallback, missing charting
   assembly, and >2 GB files), and a short "what changed from
   `Invoke-HiglasAnalysis.ps1`" section.
3. Include this sample usage call both in the README and in the script's
   `.EXAMPLE` help block:

```powershell
.\Invoke-DataAnalysis.ps1 `
    -PROJ_NAME "HIGLAS" `
    -ANALYSIS_VERSION "v0" `
    -PATH_TO_DATA "\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\HIGLAS\HIGLAS_tbl_HIGLASRBDReport.csv" `
    -FINAL_OUTPUT "WORD"
```

After writing the script, review it for PowerShell 5.1 syntax errors; verify all
COM objects, readers, streams, and chart/bitmap objects are disposed/released;
confirm **no code path loads the entire CSV into memory at once** (no
`Import-Csv`/`Get-Content` on the data file); and confirm large-file arithmetic
uses `[double]` (no `Int32` overflow on files > 2 GB).
