# Invoke-HiglasAnalysis.ps1

Automated exploratory data analysis (EDA) for CSV datasets, producing a formatted
Word or HTML report intended as a discussion starting point for business analysts.

Designed for locked-down corporate Windows environments: **no Python, no internet
access, no module installation required**. The script uses only built-in Windows
PowerShell 5.1 capabilities, .NET Framework classes, and (optionally) Microsoft
Office COM automation.

*Prepared by Manuel Figallo*

## Quick start

```powershell
.\Invoke-HiglasAnalysis.ps1 `
    -PROJ_NAME "HIGLAS" `
    -ANALYSIS_VERSION "v0" `
    -PATH_TO_DATA "\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\HIGLAS\HIGLAS_tbl_HIGLASRBDReport.csv" `
    -FINAL_OUTPUT "WORD"
```

## Parameters

| Parameter          | Required | Description |
|--------------------|----------|-------------|
| `PROJ_NAME`        | Yes      | Project name used in the report title, output file names, and headers (e.g. `HIGLAS`). |
| `ANALYSIS_VERSION` | Yes      | `v0`, `v1` (alias `lite`), `v2`, or `v3`. Case-insensitive; `lite` maps to `v1`. Versions are cumulative (see below). |
| `PATH_TO_DATA`     | Yes      | Full path to the input CSV file. UNC paths (`\\server\share\...`) are supported. The script fails with a clear error if the file does not exist. |
| `FINAL_OUTPUT`     | Yes      | `WORD` (builds a `.docx` via Word COM automation) or `HTML` (self-contained `.html`). If Word is not installed, WORD requests fall back to HTML with a warning. |
| `OutputFolder`     | No       | Folder where the report and chart PNGs are written. Defaults to the script's directory. Created if missing. |

The output file is named `{PROJ_NAME}_Analysis_{version}_{yyyyMMdd_HHmmss}.docx`
(or `.html`). Chart images are saved alongside the report as PNG files; the HTML
report additionally embeds them as base64 so the `.html` file is fully portable.

## Analysis versions

Each version includes everything from the previous one.

### v0 — basic
- **Dataset size** — file size, row count, and column count.
- **Metadata analysis** — per-column inferred type (numeric, date,
  categorical/string), null/blank counts and percentages, distinct value counts.
- **Frequency analysis** — top 15 values (count + percent) for each categorical column.
- **Distribution analysis** — min, max, mean, median, standard deviation,
  25th/75th percentiles, and potential outlier counts (beyond 1.5 × IQR) per numeric column.
- **Automated observations & recommendations** — flags high-null columns,
  constant columns, identifier-like columns, and flag-like numeric columns as
  discussion points for analysts.

### v1 (alias `lite`) — adds
- **Correlation matrix** — Pearson correlations across all numeric columns;
  pairs with |r| > 0.7 are flagged.
- **Line graph** — record counts over time (plus the sum of the first numeric
  column on a secondary axis) when a date column is detected.
- **Bar graphs** — top-10 value frequencies for up to 5 categorical columns.

### v2 — adds
- **Scatter plots** for the top 5 numeric pairs ranked by |Pearson r|.
- **Pair plot grid** — pairwise scatter plots across up to the first 5 numeric
  columns, composed as a single image (histograms on the diagonal).

### v3 — adds
- **K-means clustering** implemented from scratch (no external libraries) on
  standardized (z-score) numeric columns. k is chosen via an elbow heuristic over
  k = 2..6 (WCSS reported per k). The report includes cluster sizes, per-cluster
  means in original units, and a scatter plot of the first two numeric variables
  colored by cluster.

## Prerequisites

- Windows with **Windows PowerShell 5.1** (pre-installed on Windows 10 / Server 2016+).
- .NET Framework 4.x (standard on the above).
- **Charts**: the `System.Windows.Forms.DataVisualization` assembly (included with
  .NET Framework). If unavailable, charts are skipped gracefully and the report
  notes it — all tabular analysis still runs.
- **WORD output**: a local Microsoft Word installation (COM automation). Without
  Word, the script automatically falls back to HTML.
- Read access to the input CSV (local or UNC path) and write access to the output folder.

No internet access, Python, or PowerShell module installation is required.

## Behavior on large files

The CSV is read with a **streaming parser** (.NET `TextFieldParser`) in a single
pass with bounded memory, so very large files do not cause
`OutOfMemoryException` the way `Import-Csv` does. If the dataset exceeds
100,000 rows, a random 100,000-row reservoir sample (fixed seed, so runs are
reproducible) is used for descriptive statistics, correlations, charts, and
clustering. Metadata and frequency counts always cover the full dataset. The
sampling is noted in the report header and the affected sections.

## Exit codes

| Code | Meaning |
|------|---------|
| 0    | Success. |
| 1    | Unexpected fatal error (details written to the error stream). |
| 2    | Data could not be loaded: file unreadable, empty, or no parseable columns. |

Each analysis section is wrapped in its own try/catch — a failure in one section
is logged into that section of the report and the remaining sections still run.

## Troubleshooting

- **"Input data file not found"** — verify the path and that your account has read
  access to the share; test with `Test-Path "\\server\share\file.csv"`.
- **"Running scripts is disabled on this system"** — your execution policy blocks
  scripts. Run with:
  `powershell.exe -ExecutionPolicy Bypass -File .\Invoke-HiglasAnalysis.ps1 ...`
- **Word report fails / falls back to HTML** — Word is not installed, not licensed,
  or COM automation is blocked. Use `-FINAL_OUTPUT HTML`, which has no Office dependency.
- **Charts missing from the report** — the charting assembly could not be loaded
  (noted in the report). Tables are unaffected. This can occur on Server Core
  installations without the .NET chart components.
- **`System.OutOfMemoryException` while loading** — should no longer occur: the
  script streams the CSV instead of loading it whole. If you still see memory
  errors, make sure you are running **64-bit** PowerShell (`[Environment]::Is64BitProcess`
  should return `True`; launch from `C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe`,
  not `SysWOW64`).
- **Slow runs on very large CSVs** — profiling is row-by-row in PowerShell;
  expect several minutes for files in the hundreds of MB. Memory stays bounded,
  but CPU time grows with row count.
- **Wrong type inference** (e.g. zero-padded codes detected as numeric) — types are
  inferred by sampling values; ID-like numeric codes may be classified as numeric.
  This affects which sections a column appears in but not the underlying data.

## Output artifacts

- `{PROJ_NAME}_Analysis_{version}_{timestamp}.docx` or `.html` — the report.
- `{PROJ_NAME}_*.png` — chart images (kept in the output folder; the HTML report
  embeds copies, so it remains portable on its own).
