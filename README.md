# Invoke-DataAnalysis.ps1

Automated exploratory data analysis (EDA) for CSV datasets, producing a formatted
Word or HTML report intended as a discussion starting point for business analysts.
Generic and reusable across projects — the project identity comes entirely from
the `PROJ_NAME` parameter.

Designed for locked-down corporate Windows environments: **no Python, no internet
access, no module installation required**. The script uses only built-in Windows
PowerShell 5.1 capabilities, .NET Framework classes, and (optionally) Microsoft
Office COM automation.

*Prepared by Manuel Figallo. Based on `Invoke-HiglasAnalysis.ps1` (kept in this
repo unchanged), re-engineered for bounded-memory streaming.*

## Quick start

```powershell
.\Invoke-DataAnalysis.ps1 `
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
| `PATH_TO_DATA`     | Yes      | Full path to the input CSV file. UNC paths (`\\server\share\...`) are supported. Fails with a clear error if the file does not exist. |
| `FINAL_OUTPUT`     | Yes      | `WORD` (builds a `.docx` via Word COM automation) or `HTML` (self-contained `.html`). If Word is not installed, WORD requests fall back to HTML with a warning. |
| `OutputFolder`     | No       | Folder where the report and chart PNGs are written. Defaults to the script's directory. Created if missing. |
| `MaxSampleRows`    | No       | Maximum rows held in memory for sample-based analyses (default `100000`, range 1,000–10,000,000). See *Memory management* below. |

The output file is named `{PROJ_NAME}_Analysis_{version}_{yyyyMMdd_HHmmss}.docx`
(or `.html`). Chart images are saved alongside as PNG files; the HTML report
embeds them as base64 so the `.html` file is fully portable.

## The four version tiers

Each tier includes everything from the previous one, and every report opens with
an **executive summary** of headline findings.

### v0 — Data Profile (basic; fast, no charts)
- File metadata: size, last modified date, row/column count, delimiter and encoding (BOM) check
- Column inventory: inferred type (numeric/date/categorical), null counts and %, distinct values
- Simple frequency analysis: top-5 values per categorical column
- Simple distribution analysis: min, max, mean, median per numeric column
- Data quality flags: fully empty columns, constant columns, mixed-type columns, duplicate row count
- Sample preview of the first 10 rows
- Data readiness summary: a pass/warn/fail verdict per column

### v1 (alias `lite`) — Descriptive Analysis
- Full frequency tables (top 15, with counts and percentages)
- Full descriptive statistics: std dev, 25th/75th percentiles, 1.5×IQR outlier counts, skew indicators
- Extended column inventory: mode, memory footprint estimate, date ranges
- Pearson correlation matrix (|r| > 0.7 pairs flagged)
- Line graph of records over time (if a date column exists)
- Bar graphs for top categorical columns, histograms for leading numeric columns,
  and a missing-data chart (null % by column)

### v2 — Relationship / Intermediate Analysis
- Scatter plots for the top 5 most correlated pairs
- Pair plot grid across the first 5 numeric columns
- Grouped comparisons: numeric statistics by the top categorical variable, with
  quartile (box-plot-style) tables and grouped bar charts
- Correlation heatmap image
- Month-over-month trend table with percentage change (if a date column exists)

### v3 — Pattern Discovery / Advanced Analysis
- K-means clustering implemented from scratch (k-means++ seeding, elbow heuristic
  over k = 2..6 with WCSS per k), cluster sizes, per-cluster means, cluster scatter plot
- Anomaly spotlight: the 10 records furthest from any cluster centroid (an
  investigation shortlist, e.g. for program-integrity work)
- Cluster narrative table: plain-language descriptions ("high amount, low frequency")
- Z-score multivariate outlier counts (records beyond 3 standard deviations)
- Cluster composition over time (if a date column exists)

## Memory management (OutOfMemoryException prevention)

This is the key engineering difference from the original script. The file is
**never loaded whole into PSObjects**:

1. **Pass 1 — streaming aggregation** over the full file with a `StreamReader` +
   `TextFieldParser` (proper quoted-field handling): row count, null counts,
   capped distinct/frequency dictionaries, running min/max/mean/variance via
   **Welford's online algorithm**, and date ranges. These full-file figures are
   *exact* regardless of file size, with bounded memory.
2. **Pass 2 — reservoir sample** of at most `MaxSampleRows` rows, stored as plain
   string arrays (never PSObjects), used only where row-level data is genuinely
   needed: percentiles/IQR outliers, correlation, charts, pair plots, and
   clustering. Affected report sections say so explicitly.
3. **Defensive measures**: a 32-bit-process warning at startup with guidance to
   launch 64-bit PowerShell; `StringBuilder` for HTML assembly; readers disposed
   in `finally`; charts/bitmaps disposed after each save; large intermediates
   released and `[GC]::Collect()` called between stages; progress reported by
   bytes read vs. file size; and if an `OutOfMemoryException` still occurs during
   sampling, the pass **automatically retries with the sample size halved**
   (floor 10,000) and notes the degradation in the report.

### Tuning `MaxSampleRows`

| Situation | Suggested value |
|-----------|-----------------|
| Default / typical desktop | `100000` (default) |
| 32-bit PowerShell or < 4 GB RAM | `25000`–`50000` |
| Memory pressure persists | `10000` (the retry floor) |
| Large-memory server, maximum fidelity | `250000`+ |

Lowering `MaxSampleRows` does **not** affect full-file figures (row counts,
frequencies, mean/min/max/std) — only percentile precision, chart density, and
clustering granularity.

## Prerequisites

- Windows with **Windows PowerShell 5.1** (pre-installed on Windows 10 / Server 2016+).
- .NET Framework 4.x (standard on the above).
- **Charts**: the `System.Windows.Forms.DataVisualization` assembly (included with
  .NET Framework). If unavailable, charts are skipped gracefully — all tabular
  analysis still runs. v0 produces no charts by design and runs fast.
- **WORD output**: a local Microsoft Word installation (COM automation). Without
  Word, the script automatically falls back to HTML.
- Read access to the input CSV (local or UNC) and write access to the output folder.

## Exit codes

| Code | Meaning |
|------|---------|
| 0    | Success. |
| 1    | Unexpected fatal error (details on the error stream). |
| 2    | Data could not be loaded: file unreadable, empty, or no parseable columns. |

Each analysis section runs in its own try/catch — a failure in one section is
logged into that section of the report and the rest still runs.

## Troubleshooting

- **"Input data file not found"** — verify the path and share permissions; test
  with `Test-Path "\\server\share\file.csv"`.
- **"Running scripts is disabled on this system"** — run with
  `powershell.exe -ExecutionPolicy Bypass -File .\Invoke-DataAnalysis.ps1 ...`
- **Memory pressure persists despite the streaming design** — (1) confirm 64-bit
  PowerShell: `[Environment]::Is64BitProcess` must be `True` (launch from
  `C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe`, not `SysWOW64`);
  (2) lower `-MaxSampleRows` (e.g. `25000`); (3) close other large processes;
  (4) note the script already auto-halves the sample on OOM down to 10,000 rows —
  if it still fails there, the machine genuinely lacks memory for any analysis
  and the file should be split or processed on a bigger machine.
- **Word report fails / falls back to HTML** — Word is missing, unlicensed, or COM
  is blocked. Use `-FINAL_OUTPUT HTML`, which has no Office dependency.
- **Charts missing** — the charting assembly could not be loaded (noted in the
  report). Tables are unaffected.
- **Slow runs on very large CSVs** — the two streaming passes read the file twice
  with row-by-row processing; expect several minutes for files in the hundreds of
  MB. Memory stays bounded; only CPU time grows with row count.
- **Wrong type inference** (e.g. zero-padded codes detected as numeric) — types
  are inferred by sampling values; ID-like codes may classify as numeric. This
  affects which sections a column appears in, not the underlying data.

## What changed from Invoke-HiglasAnalysis.ps1

`Invoke-HiglasAnalysis.ps1` remains in this repo unchanged. The new script:

1. **Generic**: all naming, help text, and report boilerplate are project-neutral;
   HIGLAS appears only as the example project.
2. **Memory-safe by design**: replaced the single-pass loader with a two-pass
   streaming design (Welford full-file statistics + bounded reservoir sample),
   added the `MaxSampleRows` tuning parameter, an OOM auto-retry that halves the
   sample, a 32-bit process warning, byte-based progress, and explicit GC/dispose
   between stages. Full-file mean/min/max/std are now *exact* even on huge files
   (the original computed them from the sample).
3. **New v0 tier** is richer: file metadata with delimiter/encoding sniffing,
   duplicate-row estimate, mixed-type detection, 10-row preview, and a
   pass/warn/fail data readiness summary.
4. **New analyses**: executive summary page, histograms, missing-data chart,
   skew indicators, mode/memory/date-range column details, grouped comparisons
   with quartile tables, correlation heatmap, month-over-month trends, anomaly
   spotlight, cluster narratives, z-score outlier counts, and cluster composition
   over time.

## Output artifacts

- `{PROJ_NAME}_Analysis_{version}_{timestamp}.docx` or `.html` — the report.
- `{PROJ_NAME}_*.png` — chart images (kept in the output folder; the HTML report
  embeds copies, so it remains portable on its own).
