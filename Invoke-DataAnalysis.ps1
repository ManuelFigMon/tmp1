<#
.SYNOPSIS
    Performs an automated exploratory data analysis (EDA) of a CSV dataset and
    generates a formatted Word or HTML report for business analysts. Reusable
    across projects: the project identity comes entirely from -PROJ_NAME.

.DESCRIPTION
    Invoke-DataAnalysis.ps1 produces an initial, automated analysis of a dataset
    intended to serve as a discussion starting point for business analysts. It is
    designed for locked-down corporate Windows environments and relies only on
    built-in Windows PowerShell 5.1 capabilities, .NET Framework classes, and
    (optionally) Microsoft Office COM automation.

    MEMORY DESIGN (OutOfMemoryException prevention)
    The file is never loaded whole into PSObjects. A streaming, two-pass design
    is used instead:
      Pass 1 - streaming aggregation over the full file: row count, per-column
               null counts, capped distinct/frequency counts, running min/max/
               mean/variance via Welford's online algorithm, and date min/max.
               This yields exact full-file metadata, frequency, and core
               distribution figures with bounded memory.
      Pass 2 - a reservoir sample of at most MaxSampleRows rows, stored as
               lightweight typed arrays, used only for analyses that need
               row-level data (percentiles, IQR outliers, correlation, charts,
               clustering). If an OutOfMemoryException still occurs during
               sampling, the pass automatically retries with the sample size
               halved (floor 10,000) and the degradation is noted in the report.

    Four cumulative analysis versions are supported:

      v0 - Data Profile (basic; fast, no charts)
        - File metadata (size, last modified, row/column count, delimiter and
          encoding check)
        - Column inventory (inferred type, nulls, distinct values)
        - Simple frequency analysis (top 5 values per categorical column)
        - Simple distribution analysis (min, max, mean, median per numeric column)
        - Data quality flags (empty/constant/mixed-type columns, duplicate rows)
        - Sample preview of the first 10 rows
        - Data readiness summary with a pass/warn/fail verdict per column

      v1 (alias: lite) - Descriptive Analysis = v0 plus
        - Full frequency tables (top 15 values with counts and percentages)
        - Full descriptive statistics (std dev, 25th/75th percentiles,
          1.5 x IQR outlier counts, skewness indicators)
        - Extended column inventory (mode, memory footprint estimate, date ranges)
        - Pearson correlation matrix (pairs with |r| > 0.7 flagged)
        - Line graph of records over time when a date column is detected
        - Bar graphs for top categorical columns, histograms for leading numeric
          columns, and a missing-data chart of null percentage by column

      v2 - Relationship / Intermediate Analysis = v1 plus
        - Scatter plots for the top 5 most correlated numeric pairs
        - Pair plot grid across up to the first 5 numeric columns
        - Grouped comparisons: numeric statistics broken down by the top
          categorical variable, with quartile tables and grouped bar charts
        - Correlation heatmap image
        - Month-over-month trend table with percentage change (if a date exists)

      v3 - Pattern Discovery / Advanced Analysis = v2 plus
        - K-means clustering implemented from scratch on the standardized
          numeric columns (typed-array sample), with k chosen by an elbow
          heuristic over k = 2..6 (WCSS reported per k), cluster sizes,
          per-cluster means, and a cluster-colored scatter plot
        - Outlier/anomaly spotlight of records furthest from any centroid
        - Cluster narrative table (plain-language cluster descriptions)
        - Z-score based multivariate outlier counts
        - Cluster composition over time (if a date column exists)

    Every report opens with an executive summary listing the version run and
    headline findings. Charts are rendered with the .NET
    System.Windows.Forms.DataVisualization.Charting assembly and saved as PNG
    files; if the assembly is unavailable, charts are skipped gracefully and the
    report notes it.

.PARAMETER PROJ_NAME
    Project name used in the report title, output file names, and headers
    (for example: "HIGLAS"). This is the only place project identity comes from.

.PARAMETER ANALYSIS_VERSION
    The analysis depth to run: v0, v1 (alias: lite), v2, or v3. Input is treated
    case-insensitively and "lite" is mapped to v1. Versions are cumulative.

.PARAMETER PATH_TO_DATA
    Full path (UNC paths supported) to the input CSV file. The file must exist;
    the script fails with a clear error message otherwise.

.PARAMETER FINAL_OUTPUT
    Output report format: WORD (Microsoft Word .docx via COM automation) or HTML
    (self-contained .html with embedded charts). If WORD is requested but
    Microsoft Word is not installed, the script falls back to HTML with a warning.

.PARAMETER OutputFolder
    Optional folder where the report and chart images are written.
    Defaults to the directory containing this script.

.PARAMETER MaxSampleRows
    Maximum number of rows held in memory for statistics, plots, and clustering
    (default 100,000). Lower it on memory-constrained machines (e.g. 25000);
    raise it for higher fidelity on machines with plenty of RAM. Full-file
    metadata, frequency, and mean/min/max/std figures are unaffected by this
    setting - they are always computed by streaming over the whole file.

.EXAMPLE
    .\Invoke-DataAnalysis.ps1 `
        -PROJ_NAME "HIGLAS" `
        -ANALYSIS_VERSION "v0" `
        -PATH_TO_DATA "\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\HIGLAS\HIGLAS_tbl_HIGLASRBDReport.csv" `
        -FINAL_OUTPUT "WORD"

    Runs the basic (v0) data profile against a CSV on a UNC share and produces a
    Word report in the script's folder.

.EXAMPLE
    .\Invoke-DataAnalysis.ps1 -PROJ_NAME "CLAIMS" -ANALYSIS_VERSION "v3" `
        -PATH_TO_DATA "C:\Data\claims.csv" -FINAL_OUTPUT "HTML" `
        -OutputFolder "C:\Reports" -MaxSampleRows 50000

    Runs the full advanced analysis with a reduced 50,000-row sample budget and
    writes a self-contained HTML report to C:\Reports.

.NOTES
    Prepared by Manuel Figallo
    Based on    : Invoke-HiglasAnalysis.ps1 (generalized and re-engineered for
                  bounded-memory streaming to prevent OutOfMemoryException)
    Target      : Windows PowerShell 5.1 / .NET Framework 4.x
    Dependencies: None beyond the OS. Word output requires Microsoft Word.
    Exit codes  : 0 = success, 1 = unexpected fatal error,
                  2 = data could not be loaded / no parseable columns.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = 'Project name used in titles and file names, e.g. HIGLAS')]
    [ValidateNotNullOrEmpty()]
    [string]$PROJ_NAME,

    [Parameter(Mandatory = $true, HelpMessage = 'Analysis version: v0, v1 (alias lite), v2 or v3')]
    [ValidateSet('v0', 'v1', 'v2', 'v3', 'lite')]
    [string]$ANALYSIS_VERSION,

    [Parameter(Mandatory = $true, HelpMessage = 'Full path (UNC supported) to the input CSV file')]
    [ValidateScript({
        if (Test-Path -LiteralPath $_ -PathType Leaf) { $true }
        else { throw "Input data file not found or not a file: '$_'. Verify the path (UNC paths are supported) and that you have read access." }
    })]
    [string]$PATH_TO_DATA,

    [Parameter(Mandatory = $true, HelpMessage = 'Report format: WORD or HTML')]
    [ValidateSet('WORD', 'HTML')]
    [string]$FINAL_OUTPUT,

    [Parameter(Mandatory = $false)]
    [string]$OutputFolder = $PSScriptRoot,

    [Parameter(Mandatory = $false, HelpMessage = 'Maximum rows held in memory for sample-based analyses')]
    [ValidateRange(1000, 10000000)]
    [int]$MaxSampleRows = 100000
)

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Script-wide constants and state
# ---------------------------------------------------------------------------
$script:MinSampleRows      = 10000    # OOM-retry floor for the sampling pass
$script:MaxPlotPoints      = 5000     # point cap per scatter chart
$script:MaxFreqColumns     = 20       # cap on categorical columns in frequency analysis
$script:MaxBarCharts       = 5
$script:MaxHistograms      = 5
$script:MaxDistinctTracked = 100000   # cap on distinct values tracked per column
$script:MaxDupHashes       = 1000000  # cap on row hashes tracked for duplicate detection
$script:ChartingAvailable  = $false
$script:TopCorrelatedPairs = @()      # set by the correlation section, used by v2 scatter plots
$script:ClusterResult      = $null    # set by the k-means section, used by cluster insights
$script:Headlines          = New-Object System.Collections.Generic.List[string]
$script:MemoryNotes        = New-Object System.Collections.Generic.List[string]
$script:InvariantCulture   = [System.Globalization.CultureInfo]::InvariantCulture
$script:NumStyle           = [System.Globalization.NumberStyles]::Float

# ---------------------------------------------------------------------------
# Console helpers
# ---------------------------------------------------------------------------
function Write-Stage {
    param([string]$Message, [string]$Color = 'Cyan')
    Write-Host ("[{0}] {1}" -f (Get-Date -Format 'HH:mm:ss'), $Message) -ForegroundColor $Color
}

# ---------------------------------------------------------------------------
# Value parsing helpers (types are inferred by sampling, never assumed)
# ---------------------------------------------------------------------------
function ConvertTo-NullableDouble {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $d = 0.0
    if ([double]::TryParse($Value, $script:NumStyle, $script:InvariantCulture, [ref]$d)) { return $d }
    $clean = $Value.Trim().TrimStart('$').Replace(',', '').TrimEnd('%')
    if ($clean.StartsWith('(') -and $clean.EndsWith(')')) {
        $clean = '-' + $clean.Substring(1, $clean.Length - 2)   # accounting-style negative
    }
    if ([double]::TryParse($clean, $script:NumStyle, $script:InvariantCulture, [ref]$d)) { return $d }
    if ([double]::TryParse($clean, [ref]$d)) { return $d }      # current-culture fallback
    return $null
}

function ConvertTo-NullableDate {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $dt = [datetime]::MinValue
    if ([datetime]::TryParse($Value.Trim(), [ref]$dt)) { return $dt }
    return $null
}

function Format-Num {
    param($Value, [int]$Decimals = 4)
    if ($null -eq $Value) { return 'n/a' }
    return [math]::Round([double]$Value, $Decimals).ToString($script:InvariantCulture)
}

function Format-Bytes {
    param([double]$Bytes)
    if ($Bytes -ge 1GB) { return ('{0:N1} GB' -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ('{0:N1} MB' -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ('{0:N1} KB' -f ($Bytes / 1KB)) }
    return ('{0:N0} bytes' -f $Bytes)
}

# ---------------------------------------------------------------------------
# Statistics helpers
# ---------------------------------------------------------------------------
function Get-Percentile {
    # Linear-interpolation percentile over an already-sorted array.
    param([double[]]$SortedValues, [double]$Percentile)
    $n = $SortedValues.Length
    if ($n -eq 0) { return $null }
    if ($n -eq 1) { return $SortedValues[0] }
    $rank = ($Percentile / 100.0) * ($n - 1)
    $lo   = [int][math]::Floor($rank)
    $hi   = [int][math]::Ceiling($rank)
    $frac = $rank - $lo
    return $SortedValues[$lo] + $frac * ($SortedValues[$hi] - $SortedValues[$lo])
}

function Get-PearsonCorrelation {
    # $X and $Y are row-aligned arrays whose elements are [double] or $null.
    param($X, $Y)
    $xs = New-Object System.Collections.Generic.List[double]
    $ys = New-Object System.Collections.Generic.List[double]
    for ($i = 0; $i -lt $X.Count; $i++) {
        if ($null -ne $X[$i] -and $null -ne $Y[$i]) {
            $xs.Add([double]$X[$i])
            $ys.Add([double]$Y[$i])
        }
    }
    $n = $xs.Count
    if ($n -lt 3) { return $null }
    $mx = 0.0; $my = 0.0
    for ($i = 0; $i -lt $n; $i++) { $mx += $xs[$i]; $my += $ys[$i] }
    $mx /= $n; $my /= $n
    $sxy = 0.0; $sxx = 0.0; $syy = 0.0
    for ($i = 0; $i -lt $n; $i++) {
        $dx = $xs[$i] - $mx
        $dy = $ys[$i] - $my
        $sxy += $dx * $dy
        $sxx += $dx * $dx
        $syy += $dy * $dy
    }
    if ($sxx -eq 0 -or $syy -eq 0) { return $null }
    return $sxy / [math]::Sqrt($sxx * $syy)
}

# ---------------------------------------------------------------------------
# Report section model
# ---------------------------------------------------------------------------
function New-Section {
    param([string]$Title)
    return @{
        Title      = $Title
        Paragraphs = New-Object System.Collections.Generic.List[string]
        Tables     = New-Object System.Collections.Generic.List[object]   # @{ Caption; Rows = PSCustomObject[] (string-valued) }
        Images     = New-Object System.Collections.Generic.List[object]   # @{ Caption; Path }
        ErrorText  = $null
    }
}

function Add-SectionTable {
    param($Section, [string]$Caption, $Rows)
    # enumerate manually rather than @($Rows): works identically for arrays,
    # generic Lists, and single objects across PowerShell hosts
    $rowList = New-Object System.Collections.Generic.List[object]
    if ($null -ne $Rows) {
        foreach ($r in $Rows) {
            if ($null -ne $r) { $rowList.Add($r) }
        }
    }
    if ($rowList.Count -gt 0) {
        $Section.Tables.Add(@{ Caption = $Caption; Rows = $rowList.ToArray() })
    }
}

function Add-SectionImage {
    param($Section, [string]$Caption, [string]$Path)
    if ($Path -and (Test-Path -LiteralPath $Path)) {
        $Section.Images.Add(@{ Caption = $Caption; Path = $Path })
    }
}

function Add-Headline {
    param([string]$Text)
    $script:Headlines.Add($Text)
}

function Invoke-SafeSection {
    # Runs one analysis section; a failure is logged into the report instead of
    # killing the run.
    param([string]$Title, [scriptblock]$Action)
    try {
        $result = & $Action
        if ($null -eq $result) {
            $result = New-Section $Title
            $result.Paragraphs.Add('No applicable data was found for this section.')
        }
        return $result
    }
    catch {
        Write-Warning ("Section '{0}' failed: {1}" -f $Title, $_.Exception.Message)
        $s = New-Section $Title
        $s.ErrorText = ('This section could not be completed. Error: {0}' -f $_.Exception.Message)
        return $s
    }
}

function New-ChartsUnavailableSection {
    # Stand-in used when the charting assembly could not be loaded, so chart-only
    # section functions are never invoked at all in that state.
    param([string]$Title)
    $s = New-Section $Title
    $s.Paragraphs.Add('The charting assembly (System.Windows.Forms.DataVisualization) is unavailable in this environment, so the charts for this section were skipped. Tabular results elsewhere in the report are unaffected.')
    return $s
}

# ---------------------------------------------------------------------------
# Charting helpers (System.Windows.Forms.DataVisualization)
# ---------------------------------------------------------------------------
function Initialize-Charting {
    try {
        Add-Type -AssemblyName System.Windows.Forms.DataVisualization -ErrorAction Stop
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
        $script:ChartingAvailable = $true
        Write-Stage 'Charting assembly loaded (System.Windows.Forms.DataVisualization).'
    }
    catch {
        $script:ChartingAvailable = $false
        Write-Warning ('Charting assembly is unavailable; charts will be skipped and noted in the report. ({0})' -f $_.Exception.Message)
    }
}

function New-AnalysisChart {
    param([int]$Width = 900, [int]$Height = 500, [string]$Title = '')
    $chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
    $chart.Width  = $Width
    $chart.Height = $Height
    $chart.BackColor = [System.Drawing.Color]::White
    $chart.AntiAliasing = [System.Windows.Forms.DataVisualization.Charting.AntiAliasingStyles]::All
    if ($Title) {
        $t = $chart.Titles.Add($Title)
        $t.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
    }
    return $chart
}

function New-DefaultChartArea {
    param($Chart, [string]$Name = 'main')
    $area = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea($Name)
    $area.AxisX.MajorGrid.LineColor = [System.Drawing.Color]::Gainsboro
    $area.AxisY.MajorGrid.LineColor = [System.Drawing.Color]::Gainsboro
    $area.AxisX.LabelStyle.Font = New-Object System.Drawing.Font('Segoe UI', 8)
    $area.AxisY.LabelStyle.Font = New-Object System.Drawing.Font('Segoe UI', 8)
    $Chart.ChartAreas.Add($area)
    return $area
}

function Save-AnalysisChart {
    # Saves the PNG and disposes the chart immediately to avoid GDI/memory buildup.
    param($Chart, [string]$FileName)
    $path = Join-Path $script:ResolvedOutputFolder $FileName
    try {
        $Chart.SaveImage($path, [System.Windows.Forms.DataVisualization.Charting.ChartImageFormat]::Png)
    }
    finally {
        $Chart.Dispose()
    }
    return $path
}

function Get-SafeFileToken {
    param([string]$Text)
    $token = ($Text -replace '[^\w\-]', '_')
    if ($token.Length -gt 40) { $token = $token.Substring(0, 40) }
    return $token
}

# ---------------------------------------------------------------------------
# File signature: delimiter sniffing + encoding (BOM) check
# ---------------------------------------------------------------------------
function Get-FileSignature {
    param([string]$Path)
    $encoding = 'No BOM (ANSI/UTF-8 assumed)'
    $delimiter = ','
    $fs = $null; $sr = $null
    try {
        $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $bom = New-Object byte[] 4
        $read = $fs.Read($bom, 0, 4)
        if ($read -ge 3 -and $bom[0] -eq 0xEF -and $bom[1] -eq 0xBB -and $bom[2] -eq 0xBF) { $encoding = 'UTF-8 with BOM' }
        elseif ($read -ge 2 -and $bom[0] -eq 0xFF -and $bom[1] -eq 0xFE) { $encoding = 'UTF-16 LE (BOM)' }
        elseif ($read -ge 2 -and $bom[0] -eq 0xFE -and $bom[1] -eq 0xFF) { $encoding = 'UTF-16 BE (BOM)' }
        $fs.Position = 0
        $sr = New-Object System.IO.StreamReader($fs, $true)
        $firstLine = $sr.ReadLine()
        if ($null -ne $firstLine) {
            $best = ','; $bestCount = -1
            foreach ($cand in @(',', ';', "`t", '|')) {
                $count = ($firstLine.ToCharArray() | Where-Object { $_ -eq $cand[0] }).Count
                if ($count -gt $bestCount) { $bestCount = $count; $best = $cand }
            }
            if ($bestCount -gt 0) { $delimiter = $best }
        }
    }
    finally {
        if ($null -ne $sr) { $sr.Dispose() } elseif ($null -ne $fs) { $fs.Dispose() }
    }
    $delimName = switch ($delimiter) {
        ','  { 'comma (,)' }
        ';'  { 'semicolon (;)' }
        "`t" { 'tab' }
        '|'  { 'pipe (|)' }
        default { $delimiter }
    }
    return @{ Encoding = $encoding; Delimiter = $delimiter; DelimiterName = $delimName }
}

# ---------------------------------------------------------------------------
# Resilient file access helpers (UNC shares can drop mid-read)
# ---------------------------------------------------------------------------
function Open-CsvParser {
    # Opens the CSV with a 1 MB sequential-scan buffer (fewer SMB round-trips on
    # UNC paths) and returns the FileStream/StreamReader/TextFieldParser trio.
    # The parser constructor argument is cast to TextReader explicitly so
    # PowerShell 5.1 cannot bind any other overload.
    param([string]$Path, [string]$Delimiter)
    $fs = New-Object System.IO.FileStream($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read,
        [System.IO.FileShare]::ReadWrite, 1048576, [System.IO.FileOptions]::SequentialScan)
    $sr = New-Object System.IO.StreamReader($fs, $true)
    $parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser([System.IO.TextReader]$sr)
    $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
    $parser.SetDelimiters(@($Delimiter))
    $parser.HasFieldsEnclosedInQuotes = $true
    $parser.TrimWhiteSpace = $false
    return @{ Stream = $fs; Reader = $sr; Parser = $parser }
}

function Invoke-WithNetworkRetry {
    # Retries an action when it fails with an I/O error (e.g. "The network path
    # was not found" on a flaky SMB share). Non-I/O errors are rethrown at once.
    param([string]$Description, [scriptblock]$Action, [int]$MaxAttempts = 3)
    $attempt = 1
    while ($true) {
        try {
            return (& $Action)
        }
        catch {
            $ex = $_.Exception
            $isIO = ($ex -is [System.IO.IOException]) -or ($ex.InnerException -is [System.IO.IOException])
            if (-not $isIO -or $attempt -ge $MaxAttempts) { throw }
            $delay = 5 * $attempt
            Write-Warning ("{0} failed with an I/O error: {1} Retrying in {2}s (attempt {3} of {4})..." -f `
                $Description, $ex.Message.Trim(), $delay, $attempt, $MaxAttempts)
            Start-Sleep -Seconds $delay
            $attempt++
        }
    }
}

# ---------------------------------------------------------------------------
# PASS 1: streaming aggregation over the FULL file (bounded memory).
# Computes exact row count, null counts, capped distinct/frequency counts,
# Welford running mean/variance, min/max, date ranges, type-inference tallies,
# duplicate-row estimate, and a 10-row preview - without retaining rows.
# ---------------------------------------------------------------------------
function Invoke-StreamingProfile {
    param([string]$Path, [string]$Delimiter)

    Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop

    $fs = $null; $sr = $null; $parser = $null
    $columns = @()
    $profiles = @{}
    $preview = New-Object System.Collections.Generic.List[object]
    $rowCount = 0
    $malformed = 0
    $dupCount = 0
    $dupCapped = $false
    $rowHashes = New-Object 'System.Collections.Generic.HashSet[int]'
    try {
        $handles = Open-CsvParser -Path $Path -Delimiter $Delimiter
        $fs = $handles.Stream; $sr = $handles.Reader; $parser = $handles.Parser
        # cast to double BEFORE any math: file lengths > 2 GB overflow the
        # [math]::Max(int, int) overload that PowerShell 5.1 binds to
        $fileLen = [double]$fs.Length
        if ($fileLen -lt 1) { $fileLen = 1.0 }

        if ($parser.EndOfData) {
            return @{ Columns = @(); RowCount = 0; Profiles = @{}; Preview = $preview; MalformedRows = 0; DuplicateCount = 0; DuplicateCapped = $false }
        }

        # ----- headers (blank or duplicate names are repaired) -----
        $rawHeaders = $parser.ReadFields()
        $seen = @{}
        $colList = New-Object System.Collections.Generic.List[string]
        for ($c = 0; $c -lt $rawHeaders.Length; $c++) {
            $h = ([string]$rawHeaders[$c]).Trim()
            if ([string]::IsNullOrWhiteSpace($h)) { $h = 'Column{0}' -f ($c + 1) }
            $base = $h; $n = 2
            while ($seen.ContainsKey($h)) { $h = '{0}_{1}' -f $base, $n; $n++ }
            $seen[$h] = $true
            $colList.Add($h)
        }
        $columns = $colList.ToArray()
        $colCount = $columns.Length

        # per-column accumulators kept in parallel arrays for loop speed
        $nullCounts  = New-Object long[]   $colCount
        $nonNull     = New-Object long[]   $colCount
        $valueCounts = New-Object object[] $colCount
        $overflow    = New-Object bool[]   $colCount
        $numHits     = New-Object long[]   $colCount
        $dateHits    = New-Object long[]   $colCount
        $wN          = New-Object long[]   $colCount
        $wMean       = New-Object double[] $colCount
        $wM2         = New-Object double[] $colCount
        $mins        = New-Object double[] $colCount
        $maxs        = New-Object double[] $colCount
        $dateMin     = New-Object object[] $colCount
        $dateMax     = New-Object object[] $colCount
        $sumLen      = New-Object long[]   $colCount
        $checkNum    = New-Object bool[]   $colCount
        $checkDate   = New-Object bool[]   $colCount
        for ($c = 0; $c -lt $colCount; $c++) {
            $valueCounts[$c] = @{}
            $mins[$c] = [double]::MaxValue
            $maxs[$c] = [double]::MinValue
            $checkNum[$c]  = $true
            $checkDate[$c] = $true
        }

        $inv = $script:InvariantCulture
        $numStyle = $script:NumStyle
        $maxDistinct = $script:MaxDistinctTracked
        $maxDup = $script:MaxDupHashes

        while (-not $parser.EndOfData) {
            $fields = $null
            try { $fields = $parser.ReadFields() }
            catch [Microsoft.VisualBasic.FileIO.MalformedLineException] { $malformed++; continue }

            for ($c = 0; $c -lt $colCount; $c++) {
                $v = $null
                if ($c -lt $fields.Length) { $v = $fields[$c] }
                if ([string]::IsNullOrWhiteSpace($v)) {
                    $nullCounts[$c]++
                    continue
                }
                $nonNull[$c]++
                $sumLen[$c] += $v.Length

                # capped distinct/frequency tracking
                $vc = $valueCounts[$c]
                if ($vc.ContainsKey($v)) { $vc[$v]++ }
                elseif ($vc.Count -lt $maxDistinct) { $vc[$v] = 1 }
                else { $overflow[$c] = $true }

                # numeric attempt (fast path first, cleaned fallback) + Welford
                if ($checkNum[$c]) {
                    $d = 0.0
                    $isNum = [double]::TryParse($v, $numStyle, $inv, [ref]$d)
                    if (-not $isNum -and $v.IndexOfAny([char[]]@('$', ',', '%', '(')) -ge 0) {
                        $clean = $v.Trim().TrimStart('$').Replace(',', '').TrimEnd('%')
                        if ($clean.StartsWith('(') -and $clean.EndsWith(')')) { $clean = '-' + $clean.Substring(1, $clean.Length - 2) }
                        $isNum = [double]::TryParse($clean, $numStyle, $inv, [ref]$d)
                    }
                    if ($isNum) {
                        $numHits[$c]++
                        $wN[$c]++
                        $delta = $d - $wMean[$c]
                        $wMean[$c] += $delta / $wN[$c]
                        $wM2[$c] += $delta * ($d - $wMean[$c])
                        if ($d -lt $mins[$c]) { $mins[$c] = $d }
                        if ($d -gt $maxs[$c]) { $maxs[$c] = $d }
                        continue
                    }
                }

                # date attempt (only while plausible)
                if ($checkDate[$c]) {
                    $dtv = [datetime]::MinValue
                    if ([datetime]::TryParse($v, [ref]$dtv)) {
                        $dateHits[$c]++
                        if ($null -eq $dateMin[$c] -or $dtv -lt $dateMin[$c]) { $dateMin[$c] = $dtv }
                        if ($null -eq $dateMax[$c] -or $dtv -gt $dateMax[$c]) { $dateMax[$c] = $dtv }
                    }
                }
            }

            # duplicate-row estimate via hash of the joined fields (approximate)
            if ($rowHashes.Count -lt $maxDup) {
                $joined = [string]::Join([string][char]31, $fields)
                if (-not $rowHashes.Add($joined.GetHashCode())) { $dupCount++ }
            }
            else { $dupCapped = $true }

            if ($rowCount -lt 10) { $preview.Add($fields) }
            $rowCount++

            if (($rowCount % 2000) -eq 0) {
                # adaptively stop numeric/date parse attempts on hopeless columns
                for ($c = 0; $c -lt $colCount; $c++) {
                    if ($nonNull[$c] -gt 500) {
                        if ($checkNum[$c]  -and ($numHits[$c]  / $nonNull[$c]) -lt 0.05) { $checkNum[$c]  = $false }
                        if ($checkDate[$c] -and ($dateHits[$c] / $nonNull[$c]) -lt 0.05) { $checkDate[$c] = $false }
                    }
                }
            }
            if (($rowCount % 20000) -eq 0) {
                $pct = [math]::Min(100, 100.0 * $fs.Position / $fileLen)
                Write-Progress -Activity 'Pass 1: streaming profile (full file)' `
                    -Status ("{0:N0} rows read ({1:N0}% of file)" -f $rowCount, $pct) -PercentComplete $pct
            }
        }
        Write-Progress -Activity 'Pass 1: streaming profile (full file)' -Completed

        # ----- assemble per-column profiles + type inference -----
        for ($c = 0; $c -lt $colCount; $c++) {
            $type = 'Categorical'
            $mixed = $false
            if ($nonNull[$c] -gt 0) {
                $numRatio  = $numHits[$c]  / $nonNull[$c]
                $dateRatio = $dateHits[$c] / $nonNull[$c]
                if ($numRatio -ge 0.8) { $type = 'Numeric' }
                elseif ($dateRatio -ge 0.8) { $type = 'Date' }
                else {
                    $top = [math]::Max($numRatio, $dateRatio)
                    if ($top -ge 0.2) { $mixed = $true }
                }
            }
            $profiles[$columns[$c]] = @{
                Name             = $columns[$c]
                Type             = $type
                MixedType        = $mixed
                NullCount        = $nullCounts[$c]
                NonNull          = $nonNull[$c]
                ValueCounts      = $valueCounts[$c]
                DistinctOverflow = $overflow[$c]
                NumericCount     = $wN[$c]
                Mean             = $(if ($wN[$c] -gt 0) { $wMean[$c] } else { $null })
                StdDev           = $(if ($wN[$c] -gt 1) { [math]::Sqrt($wM2[$c] / ($wN[$c] - 1)) } else { $null })
                Min              = $(if ($wN[$c] -gt 0) { $mins[$c] } else { $null })
                Max              = $(if ($wN[$c] -gt 0) { $maxs[$c] } else { $null })
                DateMin          = $dateMin[$c]
                DateMax          = $dateMax[$c]
                AvgLength        = $(if ($nonNull[$c] -gt 0) { $sumLen[$c] / $nonNull[$c] } else { 0 })
            }
        }
    }
    finally {
        if ($null -ne $parser) { $parser.Close(); $parser.Dispose() }
        elseif ($null -ne $sr) { $sr.Dispose() }
        elseif ($null -ne $fs) { $fs.Dispose() }
    }

    return @{
        Columns         = $columns
        RowCount        = $rowCount
        Profiles        = $profiles
        Preview         = $preview
        MalformedRows   = $malformed
        DuplicateCount  = $dupCount
        DuplicateCapped = $dupCapped
    }
}

# ---------------------------------------------------------------------------
# PASS 2: bounded reservoir sample (seeded, reproducible). Rows are stored as
# plain string arrays - never PSObjects - so memory stays proportional to
# SampleSize regardless of file size.
# ---------------------------------------------------------------------------
function Get-ReservoirSample {
    param([string]$Path, [string]$Delimiter, [int]$SampleSize, [int]$Seed = 42)

    $fs = $null; $sr = $null; $parser = $null
    $sample = New-Object System.Collections.Generic.List[object]
    try {
        $handles = Open-CsvParser -Path $Path -Delimiter $Delimiter
        $fs = $handles.Stream; $sr = $handles.Reader; $parser = $handles.Parser
        # cast to double BEFORE any math: file lengths > 2 GB overflow the
        # [math]::Max(int, int) overload that PowerShell 5.1 binds to
        $fileLen = [double]$fs.Length
        if ($fileLen -lt 1) { $fileLen = 1.0 }

        if (-not $parser.EndOfData) { $parser.ReadFields() | Out-Null }   # skip header

        $rand = New-Object System.Random($Seed)
        $rowIdx = 0
        while (-not $parser.EndOfData) {
            $fields = $null
            try { $fields = $parser.ReadFields() }
            catch [Microsoft.VisualBasic.FileIO.MalformedLineException] { continue }

            if ($rowIdx -lt $SampleSize) {
                $sample.Add($fields)
            }
            else {
                $j = $rand.Next($rowIdx + 1)
                if ($j -lt $SampleSize) { $sample[$j] = $fields }
            }
            $rowIdx++
            if (($rowIdx % 20000) -eq 0) {
                $pct = [math]::Min(100, 100.0 * $fs.Position / $fileLen)
                Write-Progress -Activity 'Pass 2: reservoir sampling' `
                    -Status ("{0:N0} rows scanned ({1:N0}% of file)" -f $rowIdx, $pct) -PercentComplete $pct
            }
        }
        Write-Progress -Activity 'Pass 2: reservoir sampling' -Completed
    }
    finally {
        if ($null -ne $parser) { $parser.Close(); $parser.Dispose() }
        elseif ($null -ne $sr) { $sr.Dispose() }
        elseif ($null -ne $fs) { $fs.Dispose() }
    }
    return $sample
}

function Get-NumericColumnArray {
    # Returns a row-aligned array of [double] or $null for one column of the
    # sampled rows (each row is a string[] produced by the streaming reader).
    param($Rows, [int]$Index)
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($row in $Rows) {
        $v = $null
        if ($Index -lt $row.Length) { $v = $row[$Index] }
        $list.Add((ConvertTo-NullableDouble ([string]$v)))
    }
    return $list
}

# ===========================================================================
# V0 ANALYSIS SECTIONS - DATA PROFILE
# ===========================================================================
function Get-FileMetadataSection {
    param([string]$DataPath, $Signature, [long]$RowCount, [int]$ColumnCount, [long]$MalformedRows)
    $section = New-Section 'File Metadata'
    $item = Get-Item -LiteralPath $DataPath
    $rows = @(
        [PSCustomObject]@{ 'Property' = 'File path';          'Value' = $DataPath }
        [PSCustomObject]@{ 'Property' = 'File size';          'Value' = (Format-Bytes $item.Length) }
        [PSCustomObject]@{ 'Property' = 'Last modified';      'Value' = $item.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss') }
        [PSCustomObject]@{ 'Property' = 'Row count (data)';   'Value' = ('{0:N0}' -f $RowCount) }
        [PSCustomObject]@{ 'Property' = 'Column count';       'Value' = ('{0:N0}' -f $ColumnCount) }
        [PSCustomObject]@{ 'Property' = 'Delimiter detected'; 'Value' = $Signature.DelimiterName }
        [PSCustomObject]@{ 'Property' = 'Encoding (BOM)';     'Value' = $Signature.Encoding }
        [PSCustomObject]@{ 'Property' = 'Malformed lines skipped'; 'Value' = ('{0:N0}' -f $MalformedRows) }
    )
    Add-SectionTable $section 'File facts' $rows
    Add-Headline ("Dataset: {0:N0} rows x {1} columns ({2})." -f $RowCount, $ColumnCount, (Format-Bytes $item.Length))
    return $section
}

function Get-MetadataSection {
    param([long]$RowCount, [string[]]$Columns, $Profiles, [int]$Level)
    $section = New-Section 'Metadata Analysis (Column Inventory)'
    $section.Paragraphs.Add(("The dataset contains {0:N0} rows and {1:N0} columns. " -f $RowCount, $Columns.Count) +
        'Column types are inferred by sampling values during the streaming pass (numeric, date, or categorical/string).')

    $rows = foreach ($col in $Columns) {
        $prof = $Profiles[$col]
        $distinct = $prof.ValueCounts.Count
        $distinctText = '{0:N0}' -f $distinct
        if ($prof.DistinctOverflow) { $distinctText = '> {0:N0}' -f $script:MaxDistinctTracked }
        $nullPct = 0.0
        if ($RowCount -gt 0) { $nullPct = 100.0 * $prof.NullCount / $RowCount }
        $o = [ordered]@{
            'Column'           = $col
            'Inferred Type'    = $prof.Type
            'Null/Blank Count' = ('{0:N0}' -f $prof.NullCount)
            'Null %'           = ('{0:N2}%' -f $nullPct)
            'Distinct Values'  = $distinctText
        }
        if ($Level -ge 1) {
            # mode (most frequent value)
            $mode = 'n/a'
            if ($prof.ValueCounts.Count -gt 0 -and -not $prof.DistinctOverflow) {
                $top = $null
                foreach ($e in $prof.ValueCounts.GetEnumerator()) {
                    if ($null -eq $top -or $e.Value -gt $top.Value) { $top = $e }
                }
                if ($null -ne $top) {
                    $mode = [string]$top.Key
                    if ($mode.Length -gt 30) { $mode = $mode.Substring(0, 27) + '...' }
                }
            }
            $o['Mode'] = $mode
            # memory footprint estimate (UTF-16 in-memory string size)
            $o['Est. Memory'] = (Format-Bytes ($prof.AvgLength * 2 * $RowCount))
            # date range for date columns
            $range = ''
            if ($prof.Type -eq 'Date' -and $null -ne $prof.DateMin) {
                $range = '{0:yyyy-MM-dd} to {1:yyyy-MM-dd}' -f $prof.DateMin, $prof.DateMax
            }
            $o['Date Range'] = $range
        }
        [PSCustomObject]$o
    }
    Add-SectionTable $section 'Column metadata' $rows
    return $section
}

function Get-FrequencySection {
    param([long]$RowCount, [string[]]$Columns, $Profiles, [int]$TopN)
    $section = New-Section 'Frequency Analysis (Categorical Columns)'
    $catCols = @($Columns | Where-Object { $Profiles[$_].Type -eq 'Categorical' })
    if ($catCols.Count -eq 0) {
        $section.Paragraphs.Add('No categorical columns were detected in this dataset.')
        return $section
    }
    $section.Paragraphs.Add(("For each categorical column, the top {0} values by frequency are shown with counts and percentages of all rows (computed over the full file)." -f $TopN))
    if ($catCols.Count -gt $script:MaxFreqColumns) {
        $section.Paragraphs.Add(("Note: {0} categorical columns were detected; only the first {1} are tabulated to keep the report readable." -f $catCols.Count, $script:MaxFreqColumns))
        $catCols = @($catCols | Select-Object -First $script:MaxFreqColumns)
    }
    foreach ($col in $catCols) {
        $prof = $Profiles[$col]
        $top = $prof.ValueCounts.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First $TopN
        $rows = foreach ($entry in $top) {
            $pct = 0.0
            if ($RowCount -gt 0) { $pct = 100.0 * $entry.Value / $RowCount }
            [PSCustomObject]@{
                'Value'   = [string]$entry.Key
                'Count'   = ('{0:N0}' -f $entry.Value)
                'Percent' = ('{0:N2}%' -f $pct)
            }
        }
        Add-SectionTable $section ("Top {0} values: {1}" -f $TopN, $col) $rows
    }
    return $section
}

function Get-DistributionSection {
    param([string[]]$NumericColumns, $Profiles, $NumericArrays, [int]$Level, [bool]$Sampled)
    $section = New-Section 'Distribution Analysis (Numeric Columns)'
    if ($NumericColumns.Count -eq 0) {
        $section.Paragraphs.Add('No numeric columns were detected in this dataset.')
        return $section
    }
    if ($Level -ge 1) {
        $section.Paragraphs.Add('Descriptive statistics per numeric column. Count, min, max, mean, and standard deviation are exact (streamed over the full file); median, percentiles, and 1.5 x IQR outlier counts are computed from the in-memory sample. The skew indicator compares mean to median.')
    }
    else {
        $section.Paragraphs.Add('Basic statistics per numeric column. Count, min, max, and mean are exact (streamed over the full file); the median is computed from the in-memory sample.')
    }
    if ($Sampled) {
        $section.Paragraphs.Add(("Sample-based figures use the random {0:N0}-row reservoir sample (the dataset exceeds that size)." -f $script:EffectiveSampleRows))
    }

    $rows = foreach ($col in $NumericColumns) {
        $prof = $Profiles[$col]
        # sample-based order statistics
        $vals = New-Object System.Collections.Generic.List[double]
        if ($NumericArrays.ContainsKey($col)) {
            foreach ($v in $NumericArrays[$col]) { if ($null -ne $v) { $vals.Add([double]$v) } }
        }
        $median = $null; $p25 = $null; $p75 = $null; $outliers = $null
        if ($vals.Count -gt 0) {
            $arr = $vals.ToArray()
            [array]::Sort($arr)
            $median = Get-Percentile -SortedValues $arr -Percentile 50
            if ($Level -ge 1) {
                $p25 = Get-Percentile -SortedValues $arr -Percentile 25
                $p75 = Get-Percentile -SortedValues $arr -Percentile 75
                $iqr = $p75 - $p25
                $lowFence  = $p25 - 1.5 * $iqr
                $highFence = $p75 + 1.5 * $iqr
                $outliers = 0
                foreach ($v in $arr) { if ($v -lt $lowFence -or $v -gt $highFence) { $outliers++ } }
            }
        }
        $o = [ordered]@{
            'Column' = $col
            'Count'  = ('{0:N0}' -f $prof.NumericCount)
            'Min'    = (Format-Num $prof.Min)
            'Max'    = (Format-Num $prof.Max)
            'Mean'   = (Format-Num $prof.Mean)
            'Median' = (Format-Num $median)
        }
        if ($Level -ge 1) {
            $o['Std Dev']  = (Format-Num $prof.StdDev)
            $o['P25']      = (Format-Num $p25)
            $o['P75']      = (Format-Num $p75)
            $o['Outliers'] = $(if ($null -ne $outliers) { '{0:N0}' -f $outliers } else { 'n/a' })
            $skew = 'n/a'
            if ($null -ne $prof.Mean -and $null -ne $median -and $null -ne $prof.StdDev -and $prof.StdDev -gt 0) {
                $s = ($prof.Mean - $median) / $prof.StdDev
                if ($s -gt 0.1) { $skew = 'Right-skewed' }
                elseif ($s -lt -0.1) { $skew = 'Left-skewed' }
                else { $skew = 'Symmetric' }
            }
            $o['Skew'] = $skew
        }
        [PSCustomObject]$o
    }
    Add-SectionTable $section 'Numeric column statistics' $rows
    return $section
}

function Get-DataQualitySection {
    param([long]$RowCount, [string[]]$Columns, $Profiles, [long]$DuplicateCount, [bool]$DuplicateCapped)
    $section = New-Section 'Data Quality Flags'
    $section.Paragraphs.Add('Automated checks over the full file intended to flag structural issues before deeper analysis.')

    $emptyCols    = New-Object System.Collections.Generic.List[string]
    $constantCols = New-Object System.Collections.Generic.List[string]
    $mixedCols    = New-Object System.Collections.Generic.List[string]
    $highNullCols = New-Object System.Collections.Generic.List[string]
    foreach ($col in $Columns) {
        $prof = $Profiles[$col]
        if ($prof.NonNull -eq 0) { $emptyCols.Add($col); continue }
        if ($prof.ValueCounts.Count -eq 1 -and -not $prof.DistinctOverflow) { $constantCols.Add($col) }
        if ($prof.MixedType) { $mixedCols.Add($col) }
        if ($RowCount -gt 0 -and (100.0 * $prof.NullCount / $RowCount) -ge 20) { $highNullCols.Add($col) }
    }

    $dupText = '{0:N0}' -f $DuplicateCount
    if ($DuplicateCapped) { $dupText += (' (approximate; checked on the first {0:N0} rows)' -f $script:MaxDupHashes) }
    else { $dupText += ' (approximate, hash-based)' }

    function Local-JoinOrNone($list) {
        if ($list.Count -eq 0) { return 'None' }
        return ($list -join ', ')
    }

    $rows = @(
        [PSCustomObject]@{ 'Check' = 'Fully empty columns';        'Result' = ('{0:N0}' -f $emptyCols.Count);    'Details' = (Local-JoinOrNone $emptyCols) }
        [PSCustomObject]@{ 'Check' = 'Constant-value columns';     'Result' = ('{0:N0}' -f $constantCols.Count); 'Details' = (Local-JoinOrNone $constantCols) }
        [PSCustomObject]@{ 'Check' = 'Mixed-type columns';         'Result' = ('{0:N0}' -f $mixedCols.Count);    'Details' = (Local-JoinOrNone $mixedCols) }
        [PSCustomObject]@{ 'Check' = 'Columns with >= 20% nulls';  'Result' = ('{0:N0}' -f $highNullCols.Count); 'Details' = (Local-JoinOrNone $highNullCols) }
        [PSCustomObject]@{ 'Check' = 'Duplicate rows';             'Result' = $dupText;                          'Details' = 'Rows identical across every column' }
    )
    Add-SectionTable $section 'Quality checks' $rows

    $issueCount = $emptyCols.Count + $constantCols.Count + $mixedCols.Count + $highNullCols.Count
    if ($DuplicateCount -gt 0) { $issueCount++ }
    if ($issueCount -eq 0) { Add-Headline 'Data quality: no structural issues flagged.' }
    else { Add-Headline ("Data quality: {0} potential issue type(s) flagged (see Data Quality Flags)." -f $issueCount) }
    return $section
}

function Get-PreviewSection {
    param($Preview, [string[]]$Columns)
    $section = New-Section 'Sample Preview (First 10 Rows)'
    if ($Preview.Count -eq 0) {
        $section.Paragraphs.Add('No data rows were available for preview.')
        return $section
    }
    $maxCols = 12
    $shownCols = $Columns
    if ($Columns.Count -gt $maxCols) {
        $shownCols = $Columns[0..($maxCols - 1)]
        $section.Paragraphs.Add(("Only the first {0} of {1} columns are shown to keep the table readable." -f $maxCols, $Columns.Count))
    }
    $rows = foreach ($fields in $Preview) {
        $o = [ordered]@{}
        for ($c = 0; $c -lt $shownCols.Count; $c++) {
            $v = ''
            if ($c -lt $fields.Length) { $v = [string]$fields[$c] }
            if ($v.Length -gt 40) { $v = $v.Substring(0, 37) + '...' }
            $o[$shownCols[$c]] = $v
        }
        [PSCustomObject]$o
    }
    Add-SectionTable $section 'First rows of the file' $rows
    return $section
}

function Get-ReadinessSection {
    param([long]$RowCount, [string[]]$Columns, $Profiles)
    $section = New-Section 'Data Readiness Summary'
    $section.Paragraphs.Add('A pass/warn/fail verdict per column so analysts can quickly decide whether the data is fit for deeper analysis. FAIL = unusable or unreliable as-is; WARN = usable with caution; PASS = no issues detected.')

    $passN = 0; $warnN = 0; $failN = 0
    $rows = foreach ($col in $Columns) {
        $prof = $Profiles[$col]
        $nullPct = 0.0
        if ($RowCount -gt 0) { $nullPct = 100.0 * $prof.NullCount / $RowCount }
        $verdict = 'PASS'; $reason = 'No issues detected'
        if ($prof.NonNull -eq 0) {
            $verdict = 'FAIL'; $reason = 'Column is fully empty'
        }
        elseif ($nullPct -ge 50) {
            $verdict = 'FAIL'; $reason = ('{0:N1}% of values are null/blank' -f $nullPct)
        }
        elseif ($prof.MixedType) {
            $verdict = 'FAIL'; $reason = 'Mixed types: values are partly numeric/date, partly text'
        }
        elseif ($nullPct -ge 20) {
            $verdict = 'WARN'; $reason = ('{0:N1}% of values are null/blank' -f $nullPct)
        }
        elseif ($prof.ValueCounts.Count -eq 1 -and -not $prof.DistinctOverflow) {
            $verdict = 'WARN'; $reason = 'Constant column (single value) - no analytical signal'
        }
        elseif ($prof.DistinctOverflow -or ($prof.NonNull -gt 1 -and $prof.ValueCounts.Count -eq $prof.NonNull)) {
            $verdict = 'WARN'; $reason = 'Identifier-like (all or very many distinct values)'
        }
        switch ($verdict) {
            'PASS' { $passN++ }
            'WARN' { $warnN++ }
            'FAIL' { $failN++ }
        }
        $distinctText = '{0:N0}' -f $prof.ValueCounts.Count
        if ($prof.DistinctOverflow) { $distinctText = '> {0:N0}' -f $script:MaxDistinctTracked }
        [PSCustomObject]@{
            'Column'   = $col
            'Type'     = $prof.Type
            'Null %'   = ('{0:N1}%' -f $nullPct)
            'Distinct' = $distinctText
            'Verdict'  = $verdict
            'Reason'   = $reason
        }
    }
    Add-SectionTable $section 'Per-column readiness verdicts' $rows
    $section.Paragraphs.Add(("Summary: {0} PASS, {1} WARN, {2} FAIL out of {3} columns." -f $passN, $warnN, $failN, $Columns.Count))
    Add-Headline ("Data readiness: {0} pass / {1} warn / {2} fail across {3} columns." -f $passN, $warnN, $failN, $Columns.Count)
    return $section
}

# ===========================================================================
# V1 ANALYSIS SECTIONS - DESCRIPTIVE ANALYSIS
# ===========================================================================
function Get-CorrelationSection {
    param([string[]]$NumericColumns, $NumericArrays, [bool]$Sampled)
    $section = New-Section 'Correlation Matrix (Pearson)'
    if ($NumericColumns.Count -lt 2) {
        $section.Paragraphs.Add('Fewer than two numeric columns were detected; a correlation matrix is not applicable.')
        return $section
    }
    $section.Paragraphs.Add('Pearson correlation coefficients computed pairwise over rows where both values are present. Pairs with |r| > 0.7 are flagged below the matrix.')
    if ($Sampled) {
        $section.Paragraphs.Add(("Correlations are computed on the random {0:N0}-row sample (the dataset exceeds that size)." -f $script:EffectiveSampleRows))
    }

    $nCols = $NumericColumns.Count
    $corr = @{}
    $pairs = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $nCols; $i++) {
        for ($j = $i; $j -lt $nCols; $j++) {
            $a = $NumericColumns[$i]; $b = $NumericColumns[$j]
            if ($i -eq $j) {
                $corr["$a|$b"] = 1.0
            }
            else {
                $r = Get-PearsonCorrelation -X $NumericArrays[$a] -Y $NumericArrays[$b]
                $corr["$a|$b"] = $r
                $corr["$b|$a"] = $r
                if ($null -ne $r) {
                    $pairs.Add(@{ A = $a; B = $b; R = $r })
                }
            }
        }
    }
    $script:CorrelationLookup = $corr

    $rows = for ($i = 0; $i -lt $nCols; $i++) {
        $a = $NumericColumns[$i]
        $o = [ordered]@{ 'Variable' = $a }
        for ($j = 0; $j -lt $nCols; $j++) {
            $b = $NumericColumns[$j]
            $key = "$a|$b"
            if ($i -eq $j) { $o[$b] = '1.000' }
            elseif ($corr.ContainsKey($key) -and $null -ne $corr[$key]) { $o[$b] = ('{0:N3}' -f $corr[$key]) }
            else { $o[$b] = 'n/a' }
        }
        [PSCustomObject]$o
    }
    Add-SectionTable $section 'Correlation matrix' $rows

    $sorted = @($pairs | Sort-Object -Property @{ Expression = { [math]::Abs($_.R) }; Descending = $true })
    $script:TopCorrelatedPairs = @($sorted | Select-Object -First 5)
    $strong = @($sorted | Where-Object { [math]::Abs($_.R) -gt 0.7 })
    if ($strong.Count -gt 0) {
        $flagRows = foreach ($p in $strong) {
            [PSCustomObject]@{
                'Variable 1' = $p.A
                'Variable 2' = $p.B
                'r'          = ('{0:N3}' -f $p.R)
                'Strength'   = $(if ([math]::Abs($p.R) -gt 0.9) { 'Very strong' } else { 'Strong' })
            }
        }
        Add-SectionTable $section 'Strongly correlated pairs (|r| > 0.7)' $flagRows
        $first = $strong[0]
        Add-Headline ("Strongest correlation: {0} ~ {1} (r = {2:N3}); {3} pair(s) exceed |r| > 0.7." -f $first.A, $first.B, $first.R, $strong.Count)
    }
    else {
        $section.Paragraphs.Add('No pairs exceeded the |r| > 0.7 threshold.')
        Add-Headline 'No strongly correlated numeric pairs (|r| > 0.7) were found.'
    }
    return $section
}

function Get-TimeSeriesSection {
    param($PlotData, $ColIndex, [string[]]$DateColumns, [string[]]$NumericColumns)
    $section = New-Section 'Trend Over Time (Line Graph)'
    if ($DateColumns.Count -eq 0) {
        $section.Paragraphs.Add('No date column was detected, so no time trend chart was produced.')
        return $section
    }
    $dateCol = $DateColumns[0]
    $dateIdx = [int]$ColIndex[$dateCol]
    $numCol  = $null
    $numIdx  = -1
    if ($NumericColumns.Count -gt 0) {
        $numCol = $NumericColumns[0]
        $numIdx = [int]$ColIndex[$numCol]
    }

    # aggregate record counts (and optional numeric sum) per period
    $buckets = @{}
    $minDate = [datetime]::MaxValue
    $maxDate = [datetime]::MinValue
    foreach ($row in $PlotData) {
        if ($dateIdx -ge $row.Length) { continue }
        $dt = ConvertTo-NullableDate ([string]$row[$dateIdx])
        if ($null -eq $dt) { continue }
        if ($dt -lt $minDate) { $minDate = $dt }
        if ($dt -gt $maxDate) { $maxDate = $dt }
    }
    if ($minDate -gt $maxDate) {
        $section.Paragraphs.Add(("Column '{0}' did not yield any parseable dates; the chart was skipped." -f $dateCol))
        return $section
    }
    $byMonth = (($maxDate - $minDate).TotalDays -gt 120)
    foreach ($row in $PlotData) {
        if ($dateIdx -ge $row.Length) { continue }
        $dt = ConvertTo-NullableDate ([string]$row[$dateIdx])
        if ($null -eq $dt) { continue }
        if ($byMonth) { $key = New-Object datetime($dt.Year, $dt.Month, 1) }
        else { $key = $dt.Date }
        if (-not $buckets.ContainsKey($key)) { $buckets[$key] = @{ Count = 0; Sum = 0.0 } }
        $buckets[$key].Count++
        if ($numIdx -ge 0 -and $numIdx -lt $row.Length) {
            $nv = ConvertTo-NullableDouble ([string]$row[$numIdx])
            if ($null -ne $nv) { $buckets[$key].Sum += [double]$nv }
        }
    }

    $granularity = 'day'
    if ($byMonth) { $granularity = 'month' }
    $section.Paragraphs.Add(("Record counts over time based on column '{0}' (aggregated by {1})." -f $dateCol, $granularity))

    $chart = New-AnalysisChart -Width 950 -Height 480 -Title ("Records over time by '{0}'" -f $dateCol)
    $area = New-DefaultChartArea -Chart $chart
    $area.AxisX.LabelStyle.Format = 'yyyy-MM-dd'
    if ($byMonth) { $area.AxisX.LabelStyle.Format = 'yyyy-MM' }
    $area.AxisX.LabelStyle.Angle = -45

    $sCount = New-Object System.Windows.Forms.DataVisualization.Charting.Series('Record count')
    $sCount.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
    $sCount.XValueType = [System.Windows.Forms.DataVisualization.Charting.ChartValueType]::DateTime
    $sCount.BorderWidth = 2
    $sCount.Color = [System.Drawing.Color]::SteelBlue
    $chart.Series.Add($sCount) | Out-Null

    $sSum = $null
    if ($numCol) {
        $sSum = New-Object System.Windows.Forms.DataVisualization.Charting.Series(("Sum of {0}" -f $numCol))
        $sSum.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
        $sSum.XValueType = [System.Windows.Forms.DataVisualization.Charting.ChartValueType]::DateTime
        $sSum.YAxisType = [System.Windows.Forms.DataVisualization.Charting.AxisType]::Secondary
        $sSum.BorderWidth = 2
        $sSum.Color = [System.Drawing.Color]::DarkOrange
        $chart.Series.Add($sSum) | Out-Null
        $area.AxisY2.MajorGrid.Enabled = $false
        $legend = New-Object System.Windows.Forms.DataVisualization.Charting.Legend('legend')
        $chart.Legends.Add($legend) | Out-Null
    }

    foreach ($key in ($buckets.Keys | Sort-Object)) {
        $sCount.Points.AddXY($key, $buckets[$key].Count) | Out-Null
        if ($sSum) { $sSum.Points.AddXY($key, $buckets[$key].Sum) | Out-Null }
    }

    $file = '{0}_timeseries.png' -f $script:ProjToken
    $path = Save-AnalysisChart -Chart $chart -FileName $file
    Add-SectionImage $section ("Records over time by '{0}'" -f $dateCol) $path
    return $section
}

function Get-BarChartSection {
    param([string[]]$Columns, $Profiles)
    $section = New-Section 'Categorical Bar Graphs'
    $catCols = @($Columns | Where-Object {
            $Profiles[$_].Type -eq 'Categorical' -and
            $Profiles[$_].ValueCounts.Count -ge 2
        } | Select-Object -First $script:MaxBarCharts)
    if ($catCols.Count -eq 0) {
        $section.Paragraphs.Add('No suitable categorical columns were found for bar charts.')
        return $section
    }
    $section.Paragraphs.Add(("Top-10 value frequencies (full-file counts) for up to {0} categorical columns." -f $script:MaxBarCharts))

    foreach ($col in $catCols) {
        $top = @($Profiles[$col].ValueCounts.GetEnumerator() |
                Sort-Object -Property Value -Descending | Select-Object -First 10)
        $chart = New-AnalysisChart -Width 850 -Height 420 -Title ("Top values: {0}" -f $col)
        $area = New-DefaultChartArea -Chart $chart
        $area.AxisX.Interval = 1
        $series = New-Object System.Windows.Forms.DataVisualization.Charting.Series('freq')
        $series.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Bar
        $series.Color = [System.Drawing.Color]::SteelBlue
        $chart.Series.Add($series) | Out-Null
        [array]::Reverse($top)   # most frequent value renders at the top
        foreach ($entry in $top) {
            $label = [string]$entry.Key
            if ($label.Length -gt 35) { $label = $label.Substring(0, 32) + '...' }
            $series.Points.AddXY($label, $entry.Value) | Out-Null
        }
        $file = '{0}_bar_{1}.png' -f $script:ProjToken, (Get-SafeFileToken $col)
        $path = Save-AnalysisChart -Chart $chart -FileName $file
        Add-SectionImage $section ("Top-10 values: {0}" -f $col) $path
    }
    return $section
}

function Get-HistogramSection {
    param($NumericArrays, [string[]]$NumericColumns)
    $section = New-Section 'Histograms (Leading Numeric Columns)'
    $cols = @($NumericColumns | Select-Object -First $script:MaxHistograms)
    if ($cols.Count -eq 0) {
        $section.Paragraphs.Add('No numeric columns were detected; no histograms were produced.')
        return $section
    }
    $section.Paragraphs.Add(("Value distributions (20 bins) for the first {0} numeric column(s), based on the in-memory sample." -f $cols.Count))

    foreach ($col in $cols) {
        $vals = New-Object System.Collections.Generic.List[double]
        foreach ($v in $NumericArrays[$col]) { if ($null -ne $v) { $vals.Add([double]$v) } }
        if ($vals.Count -lt 2) { continue }
        $arr = $vals.ToArray()
        [array]::Sort($arr)
        $vMin = $arr[0]; $vMax = $arr[$arr.Length - 1]
        $range = $vMax - $vMin
        if ($range -le 0) { continue }
        $binN = 20
        $bins = New-Object int[] $binN
        foreach ($v in $arr) {
            $b = [int][math]::Floor($binN * ($v - $vMin) / $range)
            if ($b -ge $binN) { $b = $binN - 1 }
            $bins[$b]++
        }
        $chart = New-AnalysisChart -Width 700 -Height 380 -Title ("Histogram: {0}" -f $col)
        $area = New-DefaultChartArea -Chart $chart
        $area.AxisY.Title = 'Frequency'
        $series = New-Object System.Windows.Forms.DataVisualization.Charting.Series('hist')
        $series.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Column
        $series.Color = [System.Drawing.Color]::SteelBlue
        $series['PointWidth'] = '1'
        $chart.Series.Add($series) | Out-Null
        for ($b = 0; $b -lt $binN; $b++) {
            $center = $vMin + ($b + 0.5) * $range / $binN
            $series.Points.AddXY($center, $bins[$b]) | Out-Null
        }
        $file = '{0}_hist_{1}.png' -f $script:ProjToken, (Get-SafeFileToken $col)
        $path = Save-AnalysisChart -Chart $chart -FileName $file
        Add-SectionImage $section ("Histogram: {0}" -f $col) $path
    }
    return $section
}

function Get-MissingDataChartSection {
    param([long]$RowCount, [string[]]$Columns, $Profiles)
    $section = New-Section 'Missing Data by Column'
    if ($RowCount -le 0) {
        $section.Paragraphs.Add('No rows were available.')
        return $section
    }
    $section.Paragraphs.Add('Null/blank percentage per column over the full file. Columns with high missing rates may need upstream fixes before analysis.')

    $shown = $Columns
    if ($Columns.Count -gt 25) {
        $shown = $Columns[0..24]
        $section.Paragraphs.Add(("Only the first 25 of {0} columns are charted." -f $Columns.Count))
    }
    $chart = New-AnalysisChart -Width 850 -Height ([math]::Max(380, 24 * $shown.Count + 120)) -Title 'Null/blank percentage by column'
    $area = New-DefaultChartArea -Chart $chart
    $area.AxisX.Interval = 1
    $area.AxisY.Maximum = 100
    $area.AxisY.Title = 'Null %'
    $series = New-Object System.Windows.Forms.DataVisualization.Charting.Series('nulls')
    $series.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Bar
    $series.Color = [System.Drawing.Color]::IndianRed
    $chart.Series.Add($series) | Out-Null
    for ($c = $shown.Count - 1; $c -ge 0; $c--) {
        $col = $shown[$c]
        $pct = 100.0 * $Profiles[$col].NullCount / $RowCount
        $label = $col
        if ($label.Length -gt 30) { $label = $label.Substring(0, 27) + '...' }
        $series.Points.AddXY($label, [math]::Round($pct, 2)) | Out-Null
    }
    $file = '{0}_missing.png' -f $script:ProjToken
    $path = Save-AnalysisChart -Chart $chart -FileName $file
    Add-SectionImage $section 'Null/blank percentage by column' $path
    return $section
}

# ===========================================================================
# V2 ANALYSIS SECTIONS - RELATIONSHIP / INTERMEDIATE ANALYSIS
# ===========================================================================
function Get-ScatterPlotSection {
    param($PlotData, $ColIndex, [string[]]$NumericColumns)
    $section = New-Section 'Scatter Plots (Most Correlated Pairs)'
    if ($NumericColumns.Count -lt 2) {
        $section.Paragraphs.Add('Fewer than two numeric columns were detected; scatter plots are not applicable.')
        return $section
    }

    $pairs = @($script:TopCorrelatedPairs)
    if ($pairs.Count -eq 0) {
        $pairs = @()
        for ($i = 0; $i -lt ([math]::Min(5, $NumericColumns.Count - 1)); $i++) {
            $pairs += @{ A = $NumericColumns[$i]; B = $NumericColumns[$i + 1]; R = $null }
        }
        $section.Paragraphs.Add('Correlation results were unavailable; scatter plots show adjacent numeric column pairs instead.')
    }
    else {
        $section.Paragraphs.Add('Scatter plots for the top numeric column pairs ranked by |Pearson r|, drawn from the in-memory sample.')
    }

    foreach ($pair in $pairs) {
        $a = $pair.A; $b = $pair.B
        $titleSuffix = ''
        if ($null -ne $pair.R) { $titleSuffix = ('  (r = {0:N3})' -f $pair.R) }
        $chart = New-AnalysisChart -Width 700 -Height 520 -Title ("{0} vs {1}{2}" -f $a, $b, $titleSuffix)
        $area = New-DefaultChartArea -Chart $chart
        $area.AxisX.Title = $a
        $area.AxisY.Title = $b
        $series = New-Object System.Windows.Forms.DataVisualization.Charting.Series('points')
        $series.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Point
        $series.MarkerStyle = [System.Windows.Forms.DataVisualization.Charting.MarkerStyle]::Circle
        $series.MarkerSize = 5
        $series.Color = [System.Drawing.Color]::FromArgb(140, 70, 130, 180)
        $chart.Series.Add($series) | Out-Null

        $ai = [int]$ColIndex[$a]
        $bi = [int]$ColIndex[$b]
        $plotted = 0
        foreach ($row in $PlotData) {
            if ($ai -ge $row.Length -or $bi -ge $row.Length) { continue }
            $xv = ConvertTo-NullableDouble ([string]$row[$ai])
            $yv = ConvertTo-NullableDouble ([string]$row[$bi])
            if ($null -ne $xv -and $null -ne $yv) {
                $series.Points.AddXY([double]$xv, [double]$yv) | Out-Null
                $plotted++
                if ($plotted -ge $script:MaxPlotPoints) { break }
            }
        }
        $file = '{0}_scatter_{1}_vs_{2}.png' -f $script:ProjToken, (Get-SafeFileToken $a), (Get-SafeFileToken $b)
        $path = Save-AnalysisChart -Chart $chart -FileName $file
        $caption = ("Scatter: {0} vs {1}" -f $a, $b)
        if ($null -ne $pair.R) { $caption += (' (r = {0:N3})' -f $pair.R) }
        Add-SectionImage $section $caption $path
    }
    return $section
}

function Get-PairPlotSection {
    param($PlotData, $ColIndex, [string[]]$NumericColumns)
    $section = New-Section 'Pair Plot Grid'
    $cols = @($NumericColumns | Select-Object -First 5)
    if ($cols.Count -lt 2) {
        $section.Paragraphs.Add('Fewer than two numeric columns were detected; the pair plot grid is not applicable.')
        return $section
    }
    $section.Paragraphs.Add(("Pairwise scatter plots across the first {0} numeric columns (diagonal cells show value histograms), drawn from the in-memory sample." -f $cols.Count))

    $colValues = @{}
    foreach ($c in $cols) { $colValues[$c] = New-Object System.Collections.Generic.List[object] }
    $taken = 0
    foreach ($row in $PlotData) {
        foreach ($c in $cols) {
            $idx = [int]$ColIndex[$c]
            $v = $null
            if ($idx -lt $row.Length) { $v = $row[$idx] }
            $colValues[$c].Add((ConvertTo-NullableDouble ([string]$v)))
        }
        $taken++
        if ($taken -ge $script:MaxPlotPoints) { break }
    }

    $n = $cols.Count
    $cell = 100.0 / $n
    $chart = New-AnalysisChart -Width 1000 -Height 1000 -Title ('Pair plot: first {0} numeric columns' -f $n)

    for ($i = 0; $i -lt $n; $i++) {
        for ($j = 0; $j -lt $n; $j++) {
            $areaName = "pp_${i}_${j}"
            $area = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea($areaName)
            $area.Position = New-Object System.Windows.Forms.DataVisualization.Charting.ElementPosition(
                ($j * $cell), (4.0 + $i * $cell * 0.96), $cell, ($cell * 0.96))
            $area.AxisX.LabelStyle.Enabled = $false
            $area.AxisY.LabelStyle.Enabled = $false
            $area.AxisX.MajorGrid.Enabled = $false
            $area.AxisY.MajorGrid.Enabled = $false
            $area.AxisX.MajorTickMark.Enabled = $false
            $area.AxisY.MajorTickMark.Enabled = $false
            $chart.ChartAreas.Add($area)

            $series = New-Object System.Windows.Forms.DataVisualization.Charting.Series("s_${i}_${j}")
            $series.ChartArea = $areaName

            if ($i -eq $j) {
                $series.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Column
                $series.Color = [System.Drawing.Color]::LightSteelBlue
                $vals = New-Object System.Collections.Generic.List[double]
                foreach ($v in $colValues[$cols[$i]]) { if ($null -ne $v) { $vals.Add([double]$v) } }
                if ($vals.Count -gt 1) {
                    $arr = $vals.ToArray()
                    [array]::Sort($arr)
                    $vMin = $arr[0]; $vMax = $arr[$arr.Length - 1]
                    $range = $vMax - $vMin
                    if ($range -le 0) { $range = 1.0 }
                    $bins = New-Object int[] 10
                    foreach ($v in $arr) {
                        $b = [int][math]::Floor(10 * ($v - $vMin) / $range)
                        if ($b -gt 9) { $b = 9 }
                        $bins[$b]++
                    }
                    for ($b = 0; $b -lt 10; $b++) {
                        $center = $vMin + ($b + 0.5) * $range / 10
                        $series.Points.AddXY($center, $bins[$b]) | Out-Null
                    }
                }
                $title = $chart.Titles.Add([string]$cols[$i])
                $title.DockedToChartArea = $areaName
                $title.IsDockedInsideChartArea = $true
                $title.Font = New-Object System.Drawing.Font('Segoe UI', 7, [System.Drawing.FontStyle]::Bold)
            }
            else {
                $series.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Point
                $series.MarkerStyle = [System.Windows.Forms.DataVisualization.Charting.MarkerStyle]::Circle
                $series.MarkerSize = 2
                $series.Color = [System.Drawing.Color]::FromArgb(120, 70, 130, 180)
                $xList = $colValues[$cols[$j]]
                $yList = $colValues[$cols[$i]]
                for ($r = 0; $r -lt $xList.Count; $r++) {
                    if ($null -ne $xList[$r] -and $null -ne $yList[$r]) {
                        $series.Points.AddXY([double]$xList[$r], [double]$yList[$r]) | Out-Null
                    }
                }
            }
            $chart.Series.Add($series) | Out-Null
        }
    }

    $file = '{0}_pairplot.png' -f $script:ProjToken
    $path = Save-AnalysisChart -Chart $chart -FileName $file
    Add-SectionImage $section ('Pair plot grid ({0} x {0})' -f $n) $path
    return $section
}

function Get-GroupedComparisonSection {
    param($PlotData, $ColIndex, [string[]]$Columns, $Profiles, [string[]]$NumericColumns, [bool]$Sampled)
    $section = New-Section 'Grouped Comparisons (Numeric by Category)'
    $groupCol = $null
    foreach ($col in $Columns) {
        $prof = $Profiles[$col]
        if ($prof.Type -eq 'Categorical' -and -not $prof.DistinctOverflow -and
            $prof.ValueCounts.Count -ge 2 -and $prof.ValueCounts.Count -le 50) {
            $groupCol = $col
            break
        }
    }
    if ($null -eq $groupCol -or $NumericColumns.Count -eq 0) {
        $section.Paragraphs.Add('No suitable categorical grouping column (2-50 distinct values) together with a numeric column was found; grouped comparisons are not applicable.')
        return $section
    }
    $numCols = @($NumericColumns | Select-Object -First 3)
    $topCats = @($Profiles[$groupCol].ValueCounts.GetEnumerator() |
        Sort-Object -Property Value -Descending | Select-Object -First 8 | ForEach-Object { [string]$_.Key })
    $section.Paragraphs.Add(("Descriptive statistics of {0} broken down by '{1}' (top {2} categories by frequency). Quartile columns give a box-plot-style summary per category." -f ($numCols -join ', '), $groupCol, $topCats.Count))
    if ($Sampled) {
        $section.Paragraphs.Add(("Figures are computed on the random {0:N0}-row sample." -f $script:EffectiveSampleRows))
    }

    $gIdx = [int]$ColIndex[$groupCol]
    foreach ($numCol in $numCols) {
        $nIdx = [int]$ColIndex[$numCol]
        $byCat = @{}
        foreach ($cat in $topCats) { $byCat[$cat] = New-Object System.Collections.Generic.List[double] }
        foreach ($row in $PlotData) {
            if ($gIdx -ge $row.Length -or $nIdx -ge $row.Length) { continue }
            $cat = [string]$row[$gIdx]
            if (-not $byCat.ContainsKey($cat)) { continue }
            $v = ConvertTo-NullableDouble ([string]$row[$nIdx])
            if ($null -ne $v) { $byCat[$cat].Add([double]$v) }
        }

        $rows = foreach ($cat in $topCats) {
            $vals = $byCat[$cat]
            if ($vals.Count -eq 0) { continue }
            $arr = $vals.ToArray()
            [array]::Sort($arr)
            $sum = 0.0
            foreach ($v in $arr) { $sum += $v }
            $catLabel = $cat
            if ($catLabel.Length -gt 35) { $catLabel = $catLabel.Substring(0, 32) + '...' }
            [PSCustomObject]@{
                'Category' = $catLabel
                'Count'    = ('{0:N0}' -f $arr.Length)
                'Min'      = (Format-Num $arr[0] 2)
                'P25'      = (Format-Num (Get-Percentile -SortedValues $arr -Percentile 25) 2)
                'Median'   = (Format-Num (Get-Percentile -SortedValues $arr -Percentile 50) 2)
                'Mean'     = (Format-Num ($sum / $arr.Length) 2)
                'P75'      = (Format-Num (Get-Percentile -SortedValues $arr -Percentile 75) 2)
                'Max'      = (Format-Num $arr[$arr.Length - 1] 2)
            }
        }
        Add-SectionTable $section ("{0} by {1}" -f $numCol, $groupCol) $rows

        if ($script:ChartingAvailable) {
            $chart = New-AnalysisChart -Width 850 -Height 420 -Title ("Mean {0} by {1}" -f $numCol, $groupCol)
            $area = New-DefaultChartArea -Chart $chart
            $area.AxisX.Interval = 1
            $area.AxisX.LabelStyle.Angle = -45
            $series = New-Object System.Windows.Forms.DataVisualization.Charting.Series('means')
            $series.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Column
            $series.Color = [System.Drawing.Color]::SeaGreen
            $chart.Series.Add($series) | Out-Null
            foreach ($cat in $topCats) {
                $vals = $byCat[$cat]
                if ($vals.Count -eq 0) { continue }
                $sum = 0.0
                foreach ($v in $vals) { $sum += $v }
                $label = $cat
                if ($label.Length -gt 25) { $label = $label.Substring(0, 22) + '...' }
                $series.Points.AddXY($label, [math]::Round($sum / $vals.Count, 2)) | Out-Null
            }
            $file = '{0}_grouped_{1}.png' -f $script:ProjToken, (Get-SafeFileToken $numCol)
            $path = Save-AnalysisChart -Chart $chart -FileName $file
            Add-SectionImage $section ("Mean {0} by {1}" -f $numCol, $groupCol) $path
        }
    }
    if (-not $script:ChartingAvailable) {
        $section.Paragraphs.Add('The charting assembly is unavailable; grouped bar charts were skipped (tables above are unaffected).')
    }
    return $section
}

function Get-CorrelationHeatmapSection {
    param([string[]]$NumericColumns)
    $section = New-Section 'Correlation Heatmap'
    $corr = $script:CorrelationLookup
    if ($null -eq $corr -or $NumericColumns.Count -lt 2) {
        $section.Paragraphs.Add('Correlation results were unavailable; the heatmap was skipped.')
        return $section
    }
    $section.Paragraphs.Add('Visual rendering of the Pearson correlation matrix: blue = negative, white = near zero, red = positive.')

    $n = $NumericColumns.Count
    $cellPx = [math]::Max(28, [math]::Min(60, [int](560 / $n)))
    $margin = 150
    $w = $margin + $n * $cellPx + 20
    $h = $margin + $n * $cellPx + 20
    $bmp = New-Object System.Drawing.Bitmap($w, $h)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    try {
        $g.Clear([System.Drawing.Color]::White)
        $font = New-Object System.Drawing.Font('Segoe UI', 7)
        $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Black)
        $cellFont = New-Object System.Drawing.Font('Segoe UI', 6.5)
        for ($i = 0; $i -lt $n; $i++) {
            for ($j = 0; $j -lt $n; $j++) {
                $r = 0.0
                if ($i -eq $j) { $r = 1.0 }
                else {
                    $key = "{0}|{1}" -f $NumericColumns[$i], $NumericColumns[$j]
                    if ($corr.ContainsKey($key) -and $null -ne $corr[$key]) { $r = [double]$corr[$key] } else { $r = [double]::NaN }
                }
                $x = $margin + $j * $cellPx
                $y = $margin + $i * $cellPx
                if ([double]::IsNaN($r)) {
                    $color = [System.Drawing.Color]::Gainsboro
                }
                elseif ($r -ge 0) {
                    $shade = [int](255 - 195 * $r)
                    $color = [System.Drawing.Color]::FromArgb(255, $shade, $shade)
                }
                else {
                    $shade = [int](255 + 195 * $r)
                    $color = [System.Drawing.Color]::FromArgb($shade, $shade, 255)
                }
                $cellBrush = New-Object System.Drawing.SolidBrush($color)
                $g.FillRectangle($cellBrush, $x, $y, $cellPx - 1, $cellPx - 1)
                $cellBrush.Dispose()
                if (-not [double]::IsNaN($r)) {
                    $g.DrawString(('{0:N2}' -f $r), $cellFont, $brush, ($x + 2), ($y + [int]($cellPx / 2) - 6))
                }
            }
            # row label
            $label = $NumericColumns[$i]
            if ($label.Length -gt 24) { $label = $label.Substring(0, 21) + '...' }
            $g.DrawString($label, $font, $brush, 4, ($margin + $i * $cellPx + [int]($cellPx / 2) - 6))
            # column label (rotated would be nicer; angled placement keeps it simple)
            $sf = $g.Save()
            $g.TranslateTransform(($margin + $i * $cellPx + [int]($cellPx / 2)), ($margin - 6))
            $g.RotateTransform(-60)
            $g.DrawString($label, $font, $brush, 0, 0)
            $g.Restore($sf)
        }
        $font.Dispose(); $cellFont.Dispose(); $brush.Dispose()
        $file = '{0}_heatmap.png' -f $script:ProjToken
        $path = Join-Path $script:ResolvedOutputFolder $file
        $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
        Add-SectionImage $section 'Pearson correlation heatmap' $path
    }
    finally {
        $g.Dispose()
        $bmp.Dispose()
    }
    return $section
}

function Get-MonthlyTrendSection {
    param($PlotData, $ColIndex, [string[]]$DateColumns, [string[]]$NumericColumns, [bool]$Sampled)
    $section = New-Section 'Month-over-Month Trends'
    if ($DateColumns.Count -eq 0) {
        $section.Paragraphs.Add('No date column was detected; month-over-month trends are not applicable.')
        return $section
    }
    $dateCol = $DateColumns[0]
    $dateIdx = [int]$ColIndex[$dateCol]
    $numCol = $null; $numIdx = -1
    if ($NumericColumns.Count -gt 0) {
        $numCol = $NumericColumns[0]
        $numIdx = [int]$ColIndex[$numCol]
    }

    $buckets = @{}
    foreach ($row in $PlotData) {
        if ($dateIdx -ge $row.Length) { continue }
        $dt = ConvertTo-NullableDate ([string]$row[$dateIdx])
        if ($null -eq $dt) { continue }
        $key = New-Object datetime($dt.Year, $dt.Month, 1)
        if (-not $buckets.ContainsKey($key)) { $buckets[$key] = @{ Count = 0; Sum = 0.0 } }
        $buckets[$key].Count++
        if ($numIdx -ge 0 -and $numIdx -lt $row.Length) {
            $nv = ConvertTo-NullableDouble ([string]$row[$numIdx])
            if ($null -ne $nv) { $buckets[$key].Sum += [double]$nv }
        }
    }
    if ($buckets.Count -eq 0) {
        $section.Paragraphs.Add(("Column '{0}' did not yield any parseable dates." -f $dateCol))
        return $section
    }
    $months = @($buckets.Keys | Sort-Object)
    if ($months.Count -gt 24) {
        $months = $months[($months.Count - 24)..($months.Count - 1)]
        $section.Paragraphs.Add('Only the most recent 24 months are shown.')
    }
    $note = ("Record counts per month based on '{0}'" -f $dateCol)
    if ($numCol) { $note += (" with the monthly sum of '{0}'" -f $numCol) }
    $note += '. Percentage change is month-over-month.'
    $section.Paragraphs.Add($note)
    if ($Sampled) {
        $section.Paragraphs.Add(("Figures are computed on the random {0:N0}-row sample." -f $script:EffectiveSampleRows))
    }

    $prevCount = $null; $prevSum = $null
    $rows = foreach ($m in $months) {
        $cnt = $buckets[$m].Count
        $sum = $buckets[$m].Sum
        $cntChg = 'n/a'
        if ($null -ne $prevCount -and $prevCount -gt 0) { $cntChg = ('{0:N1}%' -f (100.0 * ($cnt - $prevCount) / $prevCount)) }
        $o = [ordered]@{
            'Month'      = $m.ToString('yyyy-MM')
            'Records'    = ('{0:N0}' -f $cnt)
            'Records %Chg' = $cntChg
        }
        if ($numCol) {
            $sumChg = 'n/a'
            if ($null -ne $prevSum -and [math]::Abs($prevSum) -gt 1e-9) { $sumChg = ('{0:N1}%' -f (100.0 * ($sum - $prevSum) / [math]::Abs($prevSum))) }
            $o[('Sum {0}' -f $numCol)] = ('{0:N2}' -f $sum)
            $o[('Sum %Chg')] = $sumChg
        }
        $prevCount = $cnt; $prevSum = $sum
        [PSCustomObject]$o
    }
    Add-SectionTable $section 'Monthly trend' $rows
    return $section
}

# ===========================================================================
# V3 ANALYSIS SECTIONS - PATTERN DISCOVERY (k-means from scratch)
# ===========================================================================
function Invoke-KMeans {
    # Plain k-means with k-means++ style seeding. $Points is an array of double[].
    # NOTE: PowerShell variables are case-INsensitive, so the cluster-count
    # parameter must not be named $K ($k loop variables would collide with it).
    param($Points, [int]$ClusterCount, [int]$Seed = 42, [int]$MaxIterations = 100)
    $n   = $Points.Count
    $dim = $Points[0].Length
    $rand = New-Object System.Random($Seed)

    # --- k-means++ initialization ---
    $centroids = New-Object System.Collections.Generic.List[object]
    $centroids.Add(($Points[$rand.Next($n)].Clone()))
    while ($centroids.Count -lt $ClusterCount) {
        $d2 = New-Object double[] $n
        $totalD2 = 0.0
        for ($i = 0; $i -lt $n; $i++) {
            $best = [double]::MaxValue
            foreach ($c in $centroids) {
                $dist = 0.0
                for ($d = 0; $d -lt $dim; $d++) {
                    $diff = $Points[$i][$d] - $c[$d]
                    $dist += $diff * $diff
                }
                if ($dist -lt $best) { $best = $dist }
            }
            $d2[$i] = $best
            $totalD2 += $best
        }
        if ($totalD2 -le 0) {
            $centroids.Add(($Points[$rand.Next($n)].Clone()))
            continue
        }
        $target = $rand.NextDouble() * $totalD2
        $cum = 0.0
        $chosen = $n - 1
        for ($i = 0; $i -lt $n; $i++) {
            $cum += $d2[$i]
            if ($cum -ge $target) { $chosen = $i; break }
        }
        $centroids.Add(($Points[$chosen].Clone()))
    }

    # --- Lloyd iterations ---
    $assign = New-Object int[] $n
    for ($i = 0; $i -lt $n; $i++) { $assign[$i] = -1 }
    for ($iter = 0; $iter -lt $MaxIterations; $iter++) {
        $changed = $false
        for ($i = 0; $i -lt $n; $i++) {
            $bestK = 0
            $bestDist = [double]::MaxValue
            for ($k = 0; $k -lt $ClusterCount; $k++) {
                $dist = 0.0
                for ($d = 0; $d -lt $dim; $d++) {
                    $diff = $Points[$i][$d] - $centroids[$k][$d]
                    $dist += $diff * $diff
                }
                if ($dist -lt $bestDist) { $bestDist = $dist; $bestK = $k }
            }
            if ($assign[$i] -ne $bestK) { $assign[$i] = $bestK; $changed = $true }
        }
        if (-not $changed) { break }
        # recompute centroids
        $sums   = New-Object 'object[]' $ClusterCount
        $counts = New-Object int[] $ClusterCount
        for ($k = 0; $k -lt $ClusterCount; $k++) { $sums[$k] = New-Object double[] $dim }
        for ($i = 0; $i -lt $n; $i++) {
            $k = $assign[$i]
            $counts[$k]++
            for ($d = 0; $d -lt $dim; $d++) { $sums[$k][$d] += $Points[$i][$d] }
        }
        for ($k = 0; $k -lt $ClusterCount; $k++) {
            if ($counts[$k] -gt 0) {
                for ($d = 0; $d -lt $dim; $d++) { $centroids[$k][$d] = $sums[$k][$d] / $counts[$k] }
            }
            else {
                $centroids[$k] = $Points[$rand.Next($n)].Clone()   # re-seed empty cluster
            }
        }
    }

    # --- WCSS + per-point distance to its own centroid ---
    $wcss = 0.0
    $dists = New-Object double[] $n
    for ($i = 0; $i -lt $n; $i++) {
        $k = $assign[$i]
        $dist = 0.0
        for ($d = 0; $d -lt $dim; $d++) {
            $diff = $Points[$i][$d] - $centroids[$k][$d]
            $dist += $diff * $diff
        }
        $wcss += $dist
        $dists[$i] = [math]::Sqrt($dist)
    }
    return @{ Assignments = $assign; Centroids = $centroids; WCSS = $wcss; Distances = $dists }
}

function Get-KMeansSection {
    param($PlotData, $ColIndex, [string[]]$NumericColumns)
    $section = New-Section 'K-Means Clustering'
    if ($NumericColumns.Count -lt 2) {
        $section.Paragraphs.Add('Fewer than two numeric columns were detected; k-means clustering is not applicable.')
        return $section
    }

    Write-Stage 'Running k-means clustering (this can take a moment)...'
    $cols = @($NumericColumns | Select-Object -First 8)
    $colIdxs = New-Object int[] $cols.Count
    for ($c = 0; $c -lt $cols.Count; $c++) { $colIdxs[$c] = [int]$ColIndex[$cols[$c]] }
    $rawRows = New-Object System.Collections.Generic.List[object]
    $rowRefs = New-Object System.Collections.Generic.List[object]   # original string[] rows for the complete cases
    foreach ($row in $PlotData) {
        $vec = New-Object double[] $cols.Count
        $ok = $true
        for ($c = 0; $c -lt $cols.Count; $c++) {
            $v = $null
            if ($colIdxs[$c] -lt $row.Length) { $v = ConvertTo-NullableDouble ([string]$row[$colIdxs[$c]]) }
            if ($null -eq $v) { $ok = $false; break }
            $vec[$c] = [double]$v
        }
        if ($ok) { $rawRows.Add($vec); $rowRefs.Add($row) }
    }
    if ($rawRows.Count -lt 20) {
        $section.Paragraphs.Add(("Only {0} complete rows were available across the numeric columns; at least 20 are required for clustering." -f $rawRows.Count))
        return $section
    }

    $note = ("Clustering uses {0:N0} complete rows across {1} standardized (z-score) numeric columns: {2}." -f `
        $rawRows.Count, $cols.Count, ($cols -join ', '))
    if ($NumericColumns.Count -gt 8) {
        $note += ' (Numeric columns beyond the first 8 were excluded to keep clustering tractable.)'
    }
    $section.Paragraphs.Add($note)

    # ----- standardize (z-score) -----
    $dim = $cols.Count
    $nPts = $rawRows.Count
    $means = New-Object double[] $dim
    $sds   = New-Object double[] $dim
    for ($d = 0; $d -lt $dim; $d++) {
        $sum = 0.0
        foreach ($vec in $rawRows) { $sum += $vec[$d] }
        $means[$d] = $sum / $nPts
        $ssd = 0.0
        foreach ($vec in $rawRows) { $diff = $vec[$d] - $means[$d]; $ssd += $diff * $diff }
        $sds[$d] = [math]::Sqrt($ssd / [math]::Max(1, ($nPts - 1)))
        if ($sds[$d] -le 0) { $sds[$d] = 1.0 }
    }
    $points = New-Object 'object[]' $nPts
    for ($i = 0; $i -lt $nPts; $i++) {
        $z = New-Object double[] $dim
        for ($d = 0; $d -lt $dim; $d++) { $z[$d] = ($rawRows[$i][$d] - $means[$d]) / $sds[$d] }
        $points[$i] = $z
    }

    # ----- elbow heuristic over k = 2..6 -----
    $maxK = [math]::Min(6, $nPts - 1)
    $results = @{}
    $wcssRows = New-Object System.Collections.Generic.List[object]
    for ($k = 2; $k -le $maxK; $k++) {
        Write-Progress -Activity 'K-means clustering' -Status ("Evaluating k = {0}" -f $k) `
            -PercentComplete ((($k - 2) / [math]::Max(1, ($maxK - 1))) * 100)
        $results[$k] = Invoke-KMeans -Points $points -ClusterCount $k -Seed 42
        $wcssRows.Add([PSCustomObject]@{
            'k'    = [string]$k
            'WCSS' = (Format-Num $results[$k].WCSS 2)
        })
    }
    Write-Progress -Activity 'K-means clustering' -Completed

    $bestK = 2
    for ($k = 3; $k -le $maxK; $k++) {
        $prev = $results[$k - 1].WCSS
        if ($prev -le 0) { break }
        $improvement = ($prev - $results[$k].WCSS) / $prev
        if ($improvement -ge 0.15) { $bestK = $k } else { break }
    }
    Add-SectionTable $section 'Within-cluster sum of squares (WCSS) per k' $wcssRows
    $section.Paragraphs.Add(("Elbow heuristic selected k = {0} (k was increased while each additional cluster reduced WCSS by at least 15%)." -f $bestK))

    $final = $results[$bestK]

    # ----- cluster sizes -----
    $sizes = New-Object int[] $bestK
    foreach ($a in $final.Assignments) { $sizes[$a]++ }
    $sizeRows = for ($k = 0; $k -lt $bestK; $k++) {
        [PSCustomObject]@{
            'Cluster'    = ('Cluster {0}' -f ($k + 1))
            'Rows'       = ('{0:N0}' -f $sizes[$k])
            '% of Total' = ('{0:N1}%' -f (100.0 * $sizes[$k] / $nPts))
        }
    }
    Add-SectionTable $section 'Cluster sizes' $sizeRows

    # ----- per-cluster means in ORIGINAL units -----
    $clusterSums = New-Object 'object[]' $bestK
    for ($k = 0; $k -lt $bestK; $k++) { $clusterSums[$k] = New-Object double[] $dim }
    for ($i = 0; $i -lt $nPts; $i++) {
        $k = $final.Assignments[$i]
        for ($d = 0; $d -lt $dim; $d++) { $clusterSums[$k][$d] += $rawRows[$i][$d] }
    }
    $clusterMeans = New-Object 'object[]' $bestK
    $meanRows = for ($k = 0; $k -lt $bestK; $k++) {
        $cm = New-Object double[] $dim
        $o = [ordered]@{ 'Cluster' = ('Cluster {0}' -f ($k + 1)) }
        for ($d = 0; $d -lt $dim; $d++) {
            $m = 0.0
            if ($sizes[$k] -gt 0) { $m = $clusterSums[$k][$d] / $sizes[$k] }
            $cm[$d] = $m
            $o[$cols[$d]] = (Format-Num $m 2)
        }
        $clusterMeans[$k] = $cm
        [PSCustomObject]$o
    }
    Add-SectionTable $section 'Per-cluster means (original units)' $meanRows
    Add-Headline ("K-means found {0} clusters (largest: {1:N0} rows, {2:N1}% of the sample)." -f `
        $bestK, ($sizes | Measure-Object -Maximum).Maximum, (100.0 * ($sizes | Measure-Object -Maximum).Maximum / $nPts))

    # stash everything the cluster-insights section needs
    $script:ClusterResult = @{
        Cols = $cols; ColIdxs = $colIdxs; RawRows = $rawRows; RowRefs = $rowRefs
        Points = $points; Means = $means; Sds = $sds
        BestK = $bestK; Sizes = $sizes; ClusterMeans = $clusterMeans
        Assignments = $final.Assignments; Distances = $final.Distances
    }

    # ----- cluster scatter plot on the first two numeric variables -----
    if ($script:ChartingAvailable) {
        $palette = @(
            [System.Drawing.Color]::SteelBlue, [System.Drawing.Color]::DarkOrange,
            [System.Drawing.Color]::SeaGreen,  [System.Drawing.Color]::IndianRed,
            [System.Drawing.Color]::MediumPurple, [System.Drawing.Color]::Goldenrod
        )
        $chart = New-AnalysisChart -Width 760 -Height 560 -Title ("Clusters: {0} vs {1} (k = {2})" -f $cols[0], $cols[1], $bestK)
        $area = New-DefaultChartArea -Chart $chart
        $area.AxisX.Title = $cols[0]
        $area.AxisY.Title = $cols[1]
        $legend = New-Object System.Windows.Forms.DataVisualization.Charting.Legend('legend')
        $chart.Legends.Add($legend) | Out-Null
        $seriesList = @()
        for ($k = 0; $k -lt $bestK; $k++) {
            $s = New-Object System.Windows.Forms.DataVisualization.Charting.Series(('Cluster {0}' -f ($k + 1)))
            $s.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Point
            $s.MarkerStyle = [System.Windows.Forms.DataVisualization.Charting.MarkerStyle]::Circle
            $s.MarkerSize = 5
            $s.Color = $palette[$k % $palette.Count]
            $chart.Series.Add($s) | Out-Null
            $seriesList += $s
        }
        $plotted = 0
        for ($i = 0; $i -lt $nPts; $i++) {
            $seriesList[$final.Assignments[$i]].Points.AddXY($rawRows[$i][0], $rawRows[$i][1]) | Out-Null
            $plotted++
            if ($plotted -ge $script:MaxPlotPoints) { break }
        }
        $file = '{0}_clusters.png' -f $script:ProjToken
        $path = Save-AnalysisChart -Chart $chart -FileName $file
        Add-SectionImage $section ("Cluster assignments on {0} vs {1}" -f $cols[0], $cols[1]) $path
    }
    else {
        $section.Paragraphs.Add('The charting assembly is unavailable; the cluster scatter plot was skipped.')
    }
    return $section
}

function Get-ClusterInsightsSection {
    param($ColIndex, [string[]]$DateColumns)
    $section = New-Section 'Cluster Insights (Anomalies & Narratives)'
    $cr = $script:ClusterResult
    if ($null -eq $cr) {
        $section.Paragraphs.Add('Clustering results were unavailable, so anomaly and narrative analyses were skipped.')
        return $section
    }
    $cols = $cr.Cols
    $dim  = $cols.Count
    $nPts = $cr.RawRows.Count

    # ----- outlier/anomaly spotlight: records furthest from their centroid -----
    $section.Paragraphs.Add('The anomaly spotlight lists the sampled records furthest (in standardized distance) from their assigned cluster centroid - a practical shortlist for closer investigation, e.g. in a program-integrity context.')
    $order = New-Object 'System.Collections.Generic.List[object]'
    for ($i = 0; $i -lt $nPts; $i++) { $order.Add(@{ Index = $i; Dist = $cr.Distances[$i] }) }
    $top = @($order | Sort-Object -Property @{ Expression = { $_.Dist }; Descending = $true } | Select-Object -First 10)
    $showCols = @($cols | Select-Object -First 5)
    $spotRows = foreach ($entry in $top) {
        $i = [int]$entry.Index
        $o = [ordered]@{
            'Rank'     = ''
            'Cluster'  = ('Cluster {0}' -f ($cr.Assignments[$i] + 1))
            'Distance' = ('{0:N2}' -f $entry.Dist)
        }
        for ($d = 0; $d -lt $showCols.Count; $d++) {
            $o[$showCols[$d]] = (Format-Num $cr.RawRows[$i][$d] 2)
        }
        [PSCustomObject]$o
    }
    $rank = 1
    foreach ($r in $spotRows) { $r.Rank = [string]$rank; $rank++ }
    Add-SectionTable $section 'Anomaly spotlight: top 10 records furthest from any centroid' $spotRows

    # ----- z-score based multivariate outlier counts -----
    $z3 = 0
    for ($i = 0; $i -lt $nPts; $i++) {
        $p = $cr.Points[$i]
        $maxAbs = 0.0
        for ($d = 0; $d -lt $dim; $d++) {
            $a = [math]::Abs($p[$d])
            if ($a -gt $maxAbs) { $maxAbs = $a }
        }
        if ($maxAbs -gt 3) { $z3++ }
    }
    $section.Paragraphs.Add(("Z-score screening: {0:N0} of {1:N0} sampled records ({2:N2}%) have at least one variable more than 3 standard deviations from its mean." -f $z3, $nPts, (100.0 * $z3 / $nPts)))
    Add-Headline ("Anomalies: {0:N0} record(s) ({1:N2}%) exceed 3 standard deviations on at least one numeric variable." -f $z3, (100.0 * $z3 / $nPts))

    # ----- cluster narrative table -----
    $narrRows = for ($k = 0; $k -lt $cr.BestK; $k++) {
        $traits = New-Object 'System.Collections.Generic.List[object]'
        for ($d = 0; $d -lt $dim; $d++) {
            $z = ($cr.ClusterMeans[$k][$d] - $cr.Means[$d]) / $cr.Sds[$d]
            if ([math]::Abs($z) -ge 0.5) {
                $traits.Add(@{ Col = $cols[$d]; Z = $z })
            }
        }
        $desc = 'Near the overall average on all variables'
        if ($traits.Count -gt 0) {
            $topTraits = @($traits | Sort-Object -Property @{ Expression = { [math]::Abs($_.Z) }; Descending = $true } | Select-Object -First 3)
            $parts = foreach ($t in $topTraits) {
                $word = 'high'
                if ($t.Z -lt 0) { $word = 'low' }
                if ([math]::Abs($t.Z) -ge 1.5) { $word = 'very ' + $word }
                ('{0} {1}' -f $word, $t.Col)
            }
            $desc = ($parts -join ', ')
        }
        [PSCustomObject]@{
            'Cluster'     = ('Cluster {0}' -f ($k + 1))
            'Rows'        = ('{0:N0}' -f $cr.Sizes[$k])
            'Description' = $desc
        }
    }
    Add-SectionTable $section 'Cluster narratives (vs. overall averages)' $narrRows

    # ----- cluster composition over time -----
    if ($DateColumns.Count -gt 0) {
        $dateIdx = [int]$ColIndex[$DateColumns[0]]
        $byMonth = @{}
        for ($i = 0; $i -lt $nPts; $i++) {
            $row = $cr.RowRefs[$i]
            if ($dateIdx -ge $row.Length) { continue }
            $dt = ConvertTo-NullableDate ([string]$row[$dateIdx])
            if ($null -eq $dt) { continue }
            $m = New-Object datetime($dt.Year, $dt.Month, 1)
            if (-not $byMonth.ContainsKey($m)) { $byMonth[$m] = New-Object int[] $cr.BestK }
            ($byMonth[$m])[$cr.Assignments[$i]]++
        }
        if ($byMonth.Count -gt 0) {
            $months = @($byMonth.Keys | Sort-Object)
            if ($months.Count -gt 24) { $months = $months[($months.Count - 24)..($months.Count - 1)] }
            $compRows = foreach ($m in $months) {
                $counts = $byMonth[$m]
                $tot = 0
                foreach ($c in $counts) { $tot += $c }
                $o = [ordered]@{ 'Month' = $m.ToString('yyyy-MM'); 'Rows' = ('{0:N0}' -f $tot) }
                for ($k = 0; $k -lt $cr.BestK; $k++) {
                    $pct = 0.0
                    if ($tot -gt 0) { $pct = 100.0 * $counts[$k] / $tot }
                    $o[('Cluster {0}' -f ($k + 1))] = ('{0:N1}%' -f $pct)
                }
                [PSCustomObject]$o
            }
            Add-SectionTable $section ("Cluster composition by month (based on '{0}')" -f $DateColumns[0]) $compRows
        }
    }
    else {
        $section.Paragraphs.Add('No date column was detected, so cluster composition over time was not computed.')
    }
    return $section
}

# ===========================================================================
# EXECUTIVE SUMMARY
# ===========================================================================
function Get-ExecutiveSummarySection {
    param($Meta)
    $section = New-Section 'Executive Summary'
    $section.Paragraphs.Add(("This report presents an automated initial analysis of the {0} dataset (analysis version {1}), generated {2}. It is intended as a discussion starting point for business analysts, not a finished study." -f $Meta.ProjName, $Meta.Version, $Meta.Timestamp))
    $section.Paragraphs.Add($Meta.ScopeNote)
    foreach ($note in $script:MemoryNotes) { $section.Paragraphs.Add($note) }
    if ($script:Headlines.Count -gt 0) {
        $rows = New-Object System.Collections.Generic.List[object]
        $i = 1
        foreach ($hl in $script:Headlines) {
            $rows.Add([PSCustomObject]@{ '#' = [string]$i; 'Headline finding' = $hl })
            $i++
        }
        Add-SectionTable $section 'Headline findings' $rows
    }
    else {
        $section.Paragraphs.Add('No headline findings were recorded.')
    }
    return $section
}

# ===========================================================================
# REPORT RENDERERS
# ===========================================================================
function Export-HtmlReport {
    param($Sections, $Meta, [string]$OutputPath)
    Write-Stage ("Building HTML report: {0}" -f $OutputPath)

    function HtmlEnc([string]$s) {
        if ($null -eq $s) { return '' }
        return [System.Net.WebUtility]::HtmlEncode($s)
    }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<!DOCTYPE html>')
    [void]$sb.AppendLine('<html lang="en"><head><meta charset="utf-8">')
    [void]$sb.AppendLine(('<title>{0} Automated Data Analysis ({1})</title>' -f (HtmlEnc $Meta.ProjName), (HtmlEnc $Meta.Version)))
    [void]$sb.AppendLine(@'
<style>
  body { font-family: "Segoe UI", Calibri, Arial, sans-serif; margin: 0; background: #f4f6f8; color: #222; }
  .banner { background: #1f4e79; color: #fff; padding: 28px 40px; }
  .banner h1 { margin: 0 0 6px 0; font-size: 26px; }
  .banner p { margin: 2px 0; color: #cfe0f0; font-size: 13px; }
  .content { max-width: 1080px; margin: 24px auto; padding: 0 24px 48px 24px; }
  .card { background: #fff; border: 1px solid #dde3ea; border-radius: 6px; padding: 20px 26px; margin-bottom: 22px; box-shadow: 0 1px 2px rgba(0,0,0,.05); }
  h2 { color: #1f4e79; border-bottom: 2px solid #1f4e79; padding-bottom: 6px; font-size: 19px; margin-top: 0; }
  h3 { color: #2e6da4; font-size: 15px; margin-bottom: 6px; }
  table { border-collapse: collapse; margin: 8px 0 18px 0; font-size: 12.5px; }
  th { background: #1f4e79; color: #fff; padding: 6px 10px; text-align: left; white-space: nowrap; }
  td { border: 1px solid #d4dae1; padding: 5px 10px; }
  tr:nth-child(even) td { background: #f0f4f8; }
  img.chart { max-width: 100%; border: 1px solid #dde3ea; border-radius: 4px; margin: 6px 0 16px 0; }
  .err { background: #fdecea; border: 1px solid #e6a8a1; color: #92322a; padding: 10px 14px; border-radius: 4px; }
  .toc a { color: #2e6da4; text-decoration: none; }
  .toc li { margin: 3px 0; }
  .caption { font-style: italic; color: #555; font-size: 12px; margin: 0 0 4px 0; }
  .footer { color: #888; font-size: 11px; text-align: center; padding-bottom: 24px; }
</style>
'@)
    [void]$sb.AppendLine('</head><body>')
    [void]$sb.AppendLine('<div class="banner">')
    [void]$sb.AppendLine(('<h1>{0} &mdash; Automated Data Analysis ({1})</h1>' -f (HtmlEnc $Meta.ProjName), (HtmlEnc $Meta.Version)))
    [void]$sb.AppendLine(('<p>Data file: {0}</p>' -f (HtmlEnc $Meta.DataPath)))
    [void]$sb.AppendLine(('<p>Generated: {0} &nbsp;|&nbsp; Prepared by Manuel Figallo</p>' -f (HtmlEnc $Meta.Timestamp)))
    [void]$sb.AppendLine(('<p>{0}</p>' -f (HtmlEnc $Meta.ScopeNote)))
    [void]$sb.AppendLine('</div><div class="content">')

    [void]$sb.AppendLine('<div class="card toc"><h2>Table of Contents</h2><ul>')
    for ($i = 0; $i -lt $Sections.Count; $i++) {
        [void]$sb.AppendLine(('<li><a href="#sec{0}">{1}</a></li>' -f $i, (HtmlEnc $Sections[$i].Title)))
    }
    [void]$sb.AppendLine('</ul></div>')

    for ($i = 0; $i -lt $Sections.Count; $i++) {
        $sec = $Sections[$i]
        [void]$sb.AppendLine(('<div class="card" id="sec{0}">' -f $i))
        [void]$sb.AppendLine(('<h2>{0}</h2>' -f (HtmlEnc $sec.Title)))
        if ($sec.ErrorText) {
            [void]$sb.AppendLine(('<div class="err">{0}</div>' -f (HtmlEnc $sec.ErrorText)))
        }
        foreach ($p in $sec.Paragraphs) {
            [void]$sb.AppendLine(('<p>{0}</p>' -f (HtmlEnc $p)))
        }
        foreach ($tbl in $sec.Tables) {
            [void]$sb.AppendLine(('<h3>{0}</h3>' -f (HtmlEnc $tbl.Caption)))
            $rows = $tbl.Rows
            $props = @($rows[0].PSObject.Properties.Name)
            [void]$sb.Append('<table><thead><tr>')
            foreach ($p in $props) { [void]$sb.Append(('<th>{0}</th>' -f (HtmlEnc $p))) }
            [void]$sb.AppendLine('</tr></thead><tbody>')
            foreach ($r in $rows) {
                [void]$sb.Append('<tr>')
                foreach ($p in $props) { [void]$sb.Append(('<td>{0}</td>' -f (HtmlEnc ([string]$r.$p)))) }
                [void]$sb.AppendLine('</tr>')
            }
            [void]$sb.AppendLine('</tbody></table>')
        }
        foreach ($img in $sec.Images) {
            try {
                $b64 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($img.Path))
                [void]$sb.AppendLine(('<p class="caption">{0}</p>' -f (HtmlEnc $img.Caption)))
                [void]$sb.AppendLine(('<img class="chart" alt="{0}" src="data:image/png;base64,{1}">' -f (HtmlEnc $img.Caption), $b64))
            }
            catch {
                [void]$sb.AppendLine(('<p class="err">Chart image could not be embedded: {0}</p>' -f (HtmlEnc $img.Path)))
            }
        }
        [void]$sb.AppendLine('</div>')
    }

    [void]$sb.AppendLine(('<p class="footer">{0} automated data analysis &middot; generated {1} &middot; Prepared by Manuel Figallo</p>' -f (HtmlEnc $Meta.ProjName), (HtmlEnc $Meta.Timestamp)))
    [void]$sb.AppendLine('</div></body></html>')

    [System.IO.File]::WriteAllText($OutputPath, $sb.ToString(), [System.Text.Encoding]::UTF8)
    return $true
}

function Export-WordReport {
    param($Sections, $Meta, [string]$OutputPath)
    Write-Stage ("Building Word report: {0}" -f $OutputPath)

    # Word constants
    $wdStyleTitle    = -63
    $wdStyleSubtitle = -75
    $wdStyleNormal   = -1
    $wdStyleHeading1 = -2
    $wdStory         = 6
    $wdFormatDocx    = 16
    $headerShade     = 14277081  # light gray

    $word = $null
    $doc  = $null
    try {
        try {
            $word = New-Object -ComObject Word.Application
        }
        catch {
            Write-Warning ('Microsoft Word is not available via COM automation ({0}).' -f $_.Exception.Message)
            return $false
        }

        $word.Visible = $false
        $word.DisplayAlerts = 0
        $doc = $word.Documents.Add()
        $sel = $word.Selection

        # ----- Title page -----
        $sel.Style = $doc.Styles.Item($wdStyleTitle)
        $sel.TypeText(("{0} - Automated Data Analysis" -f $Meta.ProjName))
        $sel.TypeParagraph()
        $sel.Style = $doc.Styles.Item($wdStyleSubtitle)
        $sel.TypeText(("Analysis version: {0}" -f $Meta.Version))
        $sel.TypeParagraph()
        $sel.Style = $doc.Styles.Item($wdStyleNormal)
        $sel.TypeParagraph()
        $sel.TypeText(("Data file: {0}" -f $Meta.DataPath));  $sel.TypeParagraph()
        $sel.TypeText(("Run timestamp: {0}" -f $Meta.Timestamp)); $sel.TypeParagraph()
        $sel.TypeText($Meta.ScopeNote); $sel.TypeParagraph()
        $sel.TypeParagraph()
        $sel.Font.Bold = $true
        $sel.TypeText('Prepared by Manuel Figallo')
        $sel.Font.Bold = $false
        $sel.TypeParagraph()
        $sel.InsertNewPage()

        # ----- Table of contents -----
        $sel.Style = $doc.Styles.Item($wdStyleHeading1)
        $sel.TypeText('Table of Contents')
        $sel.TypeParagraph()
        $sel.Style = $doc.Styles.Item($wdStyleNormal)
        $tocRange = $sel.Range
        $doc.TablesOfContents.Add($tocRange, $true, 1, 2) | Out-Null
        $sel.EndKey($wdStory) | Out-Null
        $sel.InsertNewPage()

        # ----- Sections -----
        foreach ($sec in $Sections) {
            $sel.Style = $doc.Styles.Item($wdStyleHeading1)
            $sel.TypeText($sec.Title)
            $sel.TypeParagraph()
            $sel.Style = $doc.Styles.Item($wdStyleNormal)

            if ($sec.ErrorText) {
                $sel.Font.Color = 255  # red (BGR)
                $sel.TypeText($sec.ErrorText)
                $sel.Font.Color = 0
                $sel.TypeParagraph()
            }
            foreach ($p in $sec.Paragraphs) {
                $sel.TypeText($p)
                $sel.TypeParagraph()
            }
            foreach ($tbl in $sec.Tables) {
                $sel.Font.Italic = $true
                $sel.TypeText($tbl.Caption)
                $sel.Font.Italic = $false
                $sel.TypeParagraph()

                $rows  = $tbl.Rows
                $props = @($rows[0].PSObject.Properties.Name)
                $wordTable = $doc.Tables.Add($sel.Range, ($rows.Count + 1), $props.Count)
                $wordTable.Borders.InsideLineStyle  = 1
                $wordTable.Borders.OutsideLineStyle = 1
                $wordTable.Range.Font.Size = 8
                try { $wordTable.AutoFitBehavior(2) | Out-Null } catch { }  # wdAutoFitWindow

                for ($c = 0; $c -lt $props.Count; $c++) {
                    $wordTable.Cell(1, $c + 1).Range.Text = [string]$props[$c]
                }
                $headerRow = $wordTable.Rows.Item(1)
                $headerRow.Range.Font.Bold = $true
                $headerRow.Shading.BackgroundPatternColor = $headerShade

                for ($r = 0; $r -lt $rows.Count; $r++) {
                    for ($c = 0; $c -lt $props.Count; $c++) {
                        $wordTable.Cell($r + 2, $c + 1).Range.Text = [string]($rows[$r].($props[$c]))
                    }
                }
                $sel.EndKey($wdStory) | Out-Null
                $sel.TypeParagraph()
            }
            foreach ($img in $sec.Images) {
                $sel.Font.Italic = $true
                $sel.TypeText($img.Caption)
                $sel.Font.Italic = $false
                $sel.TypeParagraph()
                $shape = $sel.InlineShapes.AddPicture($img.Path, $false, $true)
                if ($shape.Width -gt 460) {
                    $ratio = 460.0 / $shape.Width
                    $shape.Height = $shape.Height * $ratio
                    $shape.Width  = 460
                }
                $sel.EndKey($wdStory) | Out-Null
                $sel.TypeParagraph()
            }
            $sel.InsertNewPage()
        }

        # ----- Finalize -----
        foreach ($toc in $doc.TablesOfContents) { $toc.Update() | Out-Null }
        $doc.SaveAs2([string]$OutputPath, $wdFormatDocx)
        $doc.Close($false)
        $doc = $null
        Write-Stage ("Word report saved: {0}" -f $OutputPath) 'Green'
        return $true
    }
    catch {
        Write-Warning ('Word report generation failed: {0}' -f $_.Exception.Message)
        return $false
    }
    finally {
        if ($null -ne $doc) {
            try { $doc.Close($false) } catch { }
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) } catch { }
        }
        if ($null -ne $word) {
            try { $word.Quit() } catch { }
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch { }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

# ===========================================================================
# MAIN
# ===========================================================================
try {
    # ----- Normalize parameters -----
    $version = $ANALYSIS_VERSION.ToLowerInvariant()
    if ($version -eq 'lite') { $version = 'v1' }
    $versionLevel = [int]($version.Substring(1))
    $outputFormat = $FINAL_OUTPUT.ToUpperInvariant()

    if ([string]::IsNullOrWhiteSpace($OutputFolder)) {
        $OutputFolder = Split-Path -Parent $MyInvocation.MyCommand.Path
        if ([string]::IsNullOrWhiteSpace($OutputFolder)) { $OutputFolder = (Get-Location).Path }
    }
    if (-not (Test-Path -LiteralPath $OutputFolder)) {
        New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
    }
    $script:ResolvedOutputFolder = (Resolve-Path -LiteralPath $OutputFolder).Path
    $script:ProjToken = Get-SafeFileToken $PROJ_NAME
    $script:EffectiveSampleRows = $MaxSampleRows

    Write-Host ''
    Write-Host ('=' * 70) -ForegroundColor DarkCyan
    Write-Host (" {0} - Automated Data Analysis ({1})" -f $PROJ_NAME, $version) -ForegroundColor White
    Write-Host ' Prepared by Manuel Figallo' -ForegroundColor Gray
    Write-Host ('=' * 70) -ForegroundColor DarkCyan

    # ----- Defensive: warn early when running 32-bit PowerShell -----
    if (-not [Environment]::Is64BitProcess) {
        Write-Warning ('This session is 32-bit PowerShell, which hits memory limits far sooner. ' +
            'Launch 64-bit PowerShell instead (C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe, not SysWOW64). ' +
            'The streaming design will still keep memory bounded, but consider lowering -MaxSampleRows if problems persist.')
        $script:MemoryNotes.Add('Note: this analysis ran in a 32-bit PowerShell session; 64-bit PowerShell is recommended for large files.')
    }

    # ----- File signature (delimiter + encoding) -----
    $sig = Get-FileSignature -Path $PATH_TO_DATA
    Write-Stage ("Detected delimiter: {0}; encoding: {1}" -f $sig.DelimiterName, $sig.Encoding)

    # ----- PASS 1: streaming profile over the full file -----
    $fileSizeMB = (Get-Item -LiteralPath $PATH_TO_DATA).Length / 1MB
    Write-Stage ("Pass 1: streaming profile ({0:N1} MB): {1}" -f $fileSizeMB, $PATH_TO_DATA)
    try {
        $profileResult = Invoke-WithNetworkRetry -Description 'Pass 1 (streaming profile)' -Action {
            Invoke-StreamingProfile -Path $PATH_TO_DATA -Delimiter $sig.Delimiter
        }
    }
    catch {
        Write-Error ("Failed to read CSV file '{0}': {1}" -f $PATH_TO_DATA, $_.Exception.Message) -ErrorAction Continue
        exit 2
    }
    $Columns  = $profileResult.Columns
    $RowCount = [long]$profileResult.RowCount
    $Profiles = $profileResult.Profiles
    if ($RowCount -eq 0) {
        Write-Error ("The CSV file '{0}' contains no data rows." -f $PATH_TO_DATA) -ErrorAction Continue
        exit 2
    }
    if ($Columns.Count -eq 0) {
        Write-Error ("No parseable columns were found in '{0}'." -f $PATH_TO_DATA) -ErrorAction Continue
        exit 2
    }
    if ($profileResult.MalformedRows -gt 0) {
        Write-Warning ("{0:N0} malformed CSV line(s) were skipped during loading." -f $profileResult.MalformedRows)
    }
    $ColIndex = @{}
    for ($i = 0; $i -lt $Columns.Count; $i++) { $ColIndex[$Columns[$i]] = $i }
    $NumericColumns = @($Columns | Where-Object { $Profiles[$_].Type -eq 'Numeric' })
    $DateColumns    = @($Columns | Where-Object { $Profiles[$_].Type -eq 'Date' })
    Write-Stage ("Pass 1 complete: {0:N0} rows x {1} columns ({2} numeric, {3} date, {4} categorical)." -f `
        $RowCount, $Columns.Count, $NumericColumns.Count, $DateColumns.Count, `
        ($Columns.Count - $NumericColumns.Count - $DateColumns.Count)) 'Green'
    [GC]::Collect()

    # ----- PASS 2: reservoir sample with OutOfMemoryException auto-retry -----
    $currentSample = $MaxSampleRows
    $PlotData = $null
    Write-Stage ("Pass 2: reservoir sampling (budget {0:N0} rows)..." -f $currentSample)
    while ($null -eq $PlotData) {
        try {
            $PlotData = Invoke-WithNetworkRetry -Description 'Pass 2 (reservoir sampling)' -Action {
                Get-ReservoirSample -Path $PATH_TO_DATA -Delimiter $sig.Delimiter -SampleSize $currentSample
            }
        }
        catch [System.OutOfMemoryException] {
            if ($currentSample -le $script:MinSampleRows) { throw }
            $newSize = [math]::Max($script:MinSampleRows, [int]($currentSample / 2))
            Write-Warning ("OutOfMemoryException during sampling at {0:N0} rows; retrying with {1:N0} rows." -f $currentSample, $newSize)
            $script:MemoryNotes.Add(("Memory pressure: the sampling pass hit an OutOfMemoryException at {0:N0} rows and automatically retried at {1:N0} rows; sample-based figures use the reduced sample." -f $currentSample, $newSize))
            $currentSample = $newSize
            $PlotData = $null
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }
    $script:EffectiveSampleRows = $currentSample
    $sampled = ($RowCount -gt $currentSample)
    if ($sampled) {
        Write-Stage ("Dataset exceeds {0:N0} rows; percentiles, correlations, charts, and clustering use a random {0:N0}-row reservoir sample (metadata, frequencies, and mean/min/max/std cover the full file)." -f $currentSample) 'Yellow'
    }
    [GC]::Collect()

    # ----- Parse numeric columns once from the sample -----
    $NumericArrays = @{}
    if ($NumericColumns.Count -gt 0) {
        Write-Stage 'Parsing numeric columns from the sample...'
        $ci = 0
        foreach ($col in $NumericColumns) {
            $ci++
            Write-Progress -Activity 'Parsing numeric columns' -Status $col -PercentComplete (100 * $ci / $NumericColumns.Count)
            $NumericArrays[$col] = Get-NumericColumnArray -Rows $PlotData -Index ([int]$ColIndex[$col])
        }
        Write-Progress -Activity 'Parsing numeric columns' -Completed
    }

    # ----- Charting (v0 produces no charts and must run fast) -----
    if ($versionLevel -ge 1) { Initialize-Charting }

    # ----- Report metadata -----
    $scopeNote = ("Dataset: {0:N0} rows, {1} columns, {2:N1} MB. Full-file figures were computed by streaming; no full in-memory load was performed." -f $RowCount, $Columns.Count, $fileSizeMB)
    if ($sampled) {
        $scopeNote += (" Because the dataset exceeds {0:N0} rows, percentiles, correlations, charts, and clustering use a random {0:N0}-row sample." -f $currentSample)
    }
    if ($MaxSampleRows -ne 100000) {
        $scopeNote += (" (MaxSampleRows was set to {0:N0}.)" -f $MaxSampleRows)
    }
    $Meta = @{
        ProjName  = $PROJ_NAME
        Version   = $version
        DataPath  = $PATH_TO_DATA
        Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        ScopeNote = $scopeNote
    }

    # ----- Run analysis sections per version (cumulative; v0 = data profile) -----
    $Sections = New-Object System.Collections.Generic.List[object]

    Write-Stage 'Running V0 analysis: file metadata...'
    $Sections.Add((Invoke-SafeSection 'File Metadata' {
        Get-FileMetadataSection -DataPath $PATH_TO_DATA -Signature $sig -RowCount $RowCount -ColumnCount $Columns.Count -MalformedRows $profileResult.MalformedRows }))

    Write-Stage 'Running V0 analysis: column inventory...'
    $Sections.Add((Invoke-SafeSection 'Metadata Analysis (Column Inventory)' {
        Get-MetadataSection -RowCount $RowCount -Columns $Columns -Profiles $Profiles -Level $versionLevel }))

    Write-Stage 'Running V0 analysis: frequency tables...'
    $freqTopN = 5
    if ($versionLevel -ge 1) { $freqTopN = 15 }
    $Sections.Add((Invoke-SafeSection 'Frequency Analysis (Categorical Columns)' {
        Get-FrequencySection -RowCount $RowCount -Columns $Columns -Profiles $Profiles -TopN $freqTopN }))

    Write-Stage 'Running V0 analysis: distributions...'
    $Sections.Add((Invoke-SafeSection 'Distribution Analysis (Numeric Columns)' {
        Get-DistributionSection -NumericColumns $NumericColumns -Profiles $Profiles -NumericArrays $NumericArrays -Level $versionLevel -Sampled $sampled }))

    Write-Stage 'Running V0 analysis: data quality flags...'
    $Sections.Add((Invoke-SafeSection 'Data Quality Flags' {
        Get-DataQualitySection -RowCount $RowCount -Columns $Columns -Profiles $Profiles -DuplicateCount $profileResult.DuplicateCount -DuplicateCapped $profileResult.DuplicateCapped }))

    Write-Stage 'Running V0 analysis: sample preview...'
    $Sections.Add((Invoke-SafeSection 'Sample Preview (First 10 Rows)' {
        Get-PreviewSection -Preview $profileResult.Preview -Columns $Columns }))

    Write-Stage 'Running V0 analysis: data readiness summary...'
    $Sections.Add((Invoke-SafeSection 'Data Readiness Summary' {
        Get-ReadinessSection -RowCount $RowCount -Columns $Columns -Profiles $Profiles }))

    if ($versionLevel -ge 1) {
        Write-Stage 'Running V1 analysis: correlation matrix...'
        $Sections.Add((Invoke-SafeSection 'Correlation Matrix (Pearson)' {
            Get-CorrelationSection -NumericColumns $NumericColumns -NumericArrays $NumericArrays -Sampled $sampled }))

        Write-Stage 'Running V1 analysis: time trend chart...'
        if ($script:ChartingAvailable) {
            $Sections.Add((Invoke-SafeSection 'Trend Over Time (Line Graph)' {
                Get-TimeSeriesSection -PlotData $PlotData -ColIndex $ColIndex -DateColumns $DateColumns -NumericColumns $NumericColumns }))
        }
        else { $Sections.Add((New-ChartsUnavailableSection 'Trend Over Time (Line Graph)')) }

        Write-Stage 'Running V1 analysis: categorical bar charts...'
        if ($script:ChartingAvailable) {
            $Sections.Add((Invoke-SafeSection 'Categorical Bar Graphs' {
                Get-BarChartSection -Columns $Columns -Profiles $Profiles }))
        }
        else { $Sections.Add((New-ChartsUnavailableSection 'Categorical Bar Graphs')) }

        Write-Stage 'Running V1 analysis: histograms...'
        if ($script:ChartingAvailable) {
            $Sections.Add((Invoke-SafeSection 'Histograms (Leading Numeric Columns)' {
                Get-HistogramSection -NumericArrays $NumericArrays -NumericColumns $NumericColumns }))
        }
        else { $Sections.Add((New-ChartsUnavailableSection 'Histograms (Leading Numeric Columns)')) }

        Write-Stage 'Running V1 analysis: missing-data chart...'
        if ($script:ChartingAvailable) {
            $Sections.Add((Invoke-SafeSection 'Missing Data by Column' {
                Get-MissingDataChartSection -RowCount $RowCount -Columns $Columns -Profiles $Profiles }))
        }
        else { $Sections.Add((New-ChartsUnavailableSection 'Missing Data by Column')) }
    }

    if ($versionLevel -ge 2) {
        Write-Stage 'Running V2 analysis: scatter plots...'
        if ($script:ChartingAvailable) {
            $Sections.Add((Invoke-SafeSection 'Scatter Plots (Most Correlated Pairs)' {
                Get-ScatterPlotSection -PlotData $PlotData -ColIndex $ColIndex -NumericColumns $NumericColumns }))
        }
        else { $Sections.Add((New-ChartsUnavailableSection 'Scatter Plots (Most Correlated Pairs)')) }

        Write-Stage 'Running V2 analysis: pair plot grid...'
        if ($script:ChartingAvailable) {
            $Sections.Add((Invoke-SafeSection 'Pair Plot Grid' {
                Get-PairPlotSection -PlotData $PlotData -ColIndex $ColIndex -NumericColumns $NumericColumns }))
        }
        else { $Sections.Add((New-ChartsUnavailableSection 'Pair Plot Grid')) }

        Write-Stage 'Running V2 analysis: grouped comparisons...'
        $Sections.Add((Invoke-SafeSection 'Grouped Comparisons (Numeric by Category)' {
            Get-GroupedComparisonSection -PlotData $PlotData -ColIndex $ColIndex -Columns $Columns -Profiles $Profiles -NumericColumns $NumericColumns -Sampled $sampled }))

        Write-Stage 'Running V2 analysis: correlation heatmap...'
        if ($script:ChartingAvailable) {
            $Sections.Add((Invoke-SafeSection 'Correlation Heatmap' {
                Get-CorrelationHeatmapSection -NumericColumns $NumericColumns }))
        }
        else { $Sections.Add((New-ChartsUnavailableSection 'Correlation Heatmap')) }

        Write-Stage 'Running V2 analysis: month-over-month trends...'
        $Sections.Add((Invoke-SafeSection 'Month-over-Month Trends' {
            Get-MonthlyTrendSection -PlotData $PlotData -ColIndex $ColIndex -DateColumns $DateColumns -NumericColumns $NumericColumns -Sampled $sampled }))
    }

    if ($versionLevel -ge 3) {
        Write-Stage 'Running V3 analysis: k-means clustering...'
        $Sections.Add((Invoke-SafeSection 'K-Means Clustering' {
            Get-KMeansSection -PlotData $PlotData -ColIndex $ColIndex -NumericColumns $NumericColumns }))

        Write-Stage 'Running V3 analysis: cluster insights...'
        $Sections.Add((Invoke-SafeSection 'Cluster Insights (Anomalies & Narratives)' {
            Get-ClusterInsightsSection -ColIndex $ColIndex -DateColumns $DateColumns }))
    }

    # ----- Executive summary goes first -----
    $Sections.Insert(0, (Invoke-SafeSection 'Executive Summary' { Get-ExecutiveSummarySection -Meta $Meta }))

    # release the big intermediates before report rendering
    $PlotData = $null
    $NumericArrays = $null
    $script:ClusterResult = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    # ----- Render report -----
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $baseName = '{0}_Analysis_{1}_{2}' -f $script:ProjToken, $version, $stamp

    $reportPath = $null
    if ($outputFormat -eq 'WORD') {
        $docxPath = Join-Path $script:ResolvedOutputFolder ($baseName + '.docx')
        $ok = Export-WordReport -Sections $Sections -Meta $Meta -OutputPath $docxPath
        if ($ok) {
            $reportPath = $docxPath
        }
        else {
            Write-Warning 'Falling back to HTML output because Word automation is unavailable or failed.'
            $outputFormat = 'HTML'
        }
    }
    if ($outputFormat -eq 'HTML') {
        $htmlPath = Join-Path $script:ResolvedOutputFolder ($baseName + '.html')
        Export-HtmlReport -Sections $Sections -Meta $Meta -OutputPath $htmlPath | Out-Null
        $reportPath = $htmlPath
        Write-Stage ("HTML report saved: {0}" -f $htmlPath) 'Green'
    }

    Write-Host ''
    Write-Host ('=' * 70) -ForegroundColor DarkCyan
    Write-Stage ("Analysis complete. Report: {0}" -f $reportPath) 'Green'
    Write-Host ('=' * 70) -ForegroundColor DarkCyan
    exit 0
}
catch {
    Write-Error ("Fatal error: {0}" -f $_.Exception.Message) -ErrorAction Continue
    Write-Error ($_.ScriptStackTrace) -ErrorAction Continue
    exit 1
}
