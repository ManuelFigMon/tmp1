<#
.SYNOPSIS
    Performs an automated exploratory data analysis (EDA) of a CSV dataset and
    generates a formatted Word or HTML report for business analysts.

.DESCRIPTION
    Invoke-HiglasAnalysis.ps1 produces an initial, automated analysis of a dataset
    intended to serve as a discussion starting point for business analysts. It is
    designed for locked-down corporate Windows environments and relies only on
    built-in Windows PowerShell 5.1 capabilities, .NET Framework classes, and
    (optionally) Microsoft Office COM automation.

    Four cumulative analysis versions are supported:

      v0 (basic)
        - Dataset size (file size, row count, column count)
        - Metadata analysis (inferred types, nulls, distinct values)
        - Simple frequency analysis (top 15 values per categorical column)
        - Simple distribution analysis (descriptive statistics + outlier counts
          for numeric columns)
        - Automated observations & recommendations (high-null columns, constant
          columns, identifier-like columns, flag-like numeric columns)

      v1 (alias: lite) = v0 plus
        - Pearson correlation matrix across numeric columns (pairs with |r| > 0.7 flagged)
        - Line graph of record counts (and first numeric column sum) over time, when a
          date column is detected
        - Bar graphs of top-10 value frequencies for up to 5 categorical columns

      v2 = v1 plus
        - Scatter plots for the top 5 most strongly correlated numeric pairs
        - Pair plot grid across up to the first 5 numeric columns

      v3 = v2 plus
        - K-means clustering (implemented from scratch) on standardized numeric columns,
          with k chosen via an elbow heuristic over k = 2..6, WCSS per k, cluster sizes,
          per-cluster means, and a cluster-colored scatter plot

    Charts are rendered with the .NET System.Windows.Forms.DataVisualization.Charting
    assembly and saved as PNG files in the output folder. If the charting assembly is
    unavailable, charts are skipped gracefully and the report notes their absence.

    The CSV is read with a streaming parser (one pass, bounded memory) so very
    large files do not cause OutOfMemoryException. For datasets larger than
    100,000 rows, a seeded random 100,000-row reservoir sample is used for
    descriptive statistics, correlations, plots, and clustering (the full data
    is still used for metadata and frequency counts), and the sampling is noted
    in the report.

.PARAMETER PROJ_NAME
    Project name used in the report title, output file names, and section headers
    (for example: "HIGLAS").

.PARAMETER ANALYSIS_VERSION
    The analysis depth to run: v0, v1 (alias: lite), v2, or v3. Input is treated
    case-insensitively and "lite" is mapped to v1. Versions are cumulative.

.PARAMETER PATH_TO_DATA
    Full path (UNC paths supported) to the input CSV file. The file must exist;
    the script fails with a clear error message otherwise.

.PARAMETER FINAL_OUTPUT
    Output report format: WORD (Microsoft Word .docx via COM automation) or HTML
    (self-contained .html with embedded charts). If WORD is requested but Microsoft
    Word is not installed, the script falls back to HTML output with a warning.

.PARAMETER OutputFolder
    Optional folder where the report and chart images are written.
    Defaults to the directory containing this script.

.EXAMPLE
    .\Invoke-HiglasAnalysis.ps1 `
        -PROJ_NAME "HIGLAS" `
        -ANALYSIS_VERSION "v0" `
        -PATH_TO_DATA "\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\HIGLAS\HIGLAS_tbl_HIGLASRBDReport.csv" `
        -FINAL_OUTPUT "WORD"

    Runs the basic (v0) analysis against a CSV on a UNC share and produces a Word
    report in the script's folder.

.EXAMPLE
    .\Invoke-HiglasAnalysis.ps1 -PROJ_NAME "HIGLAS" -ANALYSIS_VERSION "v3" `
        -PATH_TO_DATA "C:\Data\claims.csv" -FINAL_OUTPUT "HTML" -OutputFolder "C:\Reports"

    Runs the full analysis (including scatter/pair plots and k-means clustering) and
    writes a self-contained HTML report to C:\Reports.

.NOTES
    Prepared by Manuel Figallo
    Target platform : Windows PowerShell 5.1 / .NET Framework 4.x
    Dependencies    : None beyond the OS. Word output requires Microsoft Word.
    Exit codes      : 0 = success, 1 = unexpected fatal error,
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
    [string]$OutputFolder = $PSScriptRoot
)

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Script-wide constants and state
# ---------------------------------------------------------------------------
$script:MaxSampleRows      = 100000   # row cap for plots and clustering
$script:MaxPlotPoints      = 5000     # point cap per scatter chart
$script:MaxFreqColumns     = 20       # cap on categorical columns in frequency analysis
$script:MaxBarCharts       = 5
$script:MaxDistinctTracked = 100000   # cap on distinct values tracked per column
$script:ChartingAvailable  = $false
$script:TopCorrelatedPairs = @()      # populated by the correlation section, used by v2 scatter plots

# ---------------------------------------------------------------------------
# Console helpers
# ---------------------------------------------------------------------------
function Write-Stage {
    param([string]$Message, [string]$Color = 'Cyan')
    Write-Host ("[{0}] {1}" -f (Get-Date -Format 'HH:mm:ss'), $Message) -ForegroundColor $Color
}

# ---------------------------------------------------------------------------
# Value parsing helpers (type inference never assumes types)
# ---------------------------------------------------------------------------
function ConvertTo-NullableDouble {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $clean = $Value.Trim().TrimStart('$').Replace(',', '').TrimEnd('%')
    if ($clean.StartsWith('(') -and $clean.EndsWith(')')) {
        # accounting-style negative, e.g. (1,234.56)
        $clean = '-' + $clean.Substring(1, $clean.Length - 2)
    }
    $d = 0.0
    if ([double]::TryParse($clean, [System.Globalization.NumberStyles]::Float,
            [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d)) {
        return $d
    }
    # fall back to the current culture (e.g. regional decimal separators)
    if ([double]::TryParse($clean, [ref]$d)) { return $d }
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
    return [math]::Round([double]$Value, $Decimals).ToString([System.Globalization.CultureInfo]::InvariantCulture)
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
    # Stand-in section used when the charting assembly could not be loaded, so
    # chart-only section functions are never invoked at all in that state.
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
    param($Chart, [string]$FileName)
    $path = Join-Path $script:ResolvedOutputFolder $FileName
    $Chart.SaveImage($path, [System.Windows.Forms.DataVisualization.Charting.ChartImageFormat]::Png)
    $Chart.Dispose()
    return $path
}

function Get-SafeFileToken {
    param([string]$Text)
    $token = ($Text -replace '[^\w\-]', '_')
    if ($token.Length -gt 40) { $token = $token.Substring(0, 40) }
    return $token
}

# ---------------------------------------------------------------------------
# Streaming CSV load, column profiling, and type inference
# ---------------------------------------------------------------------------
function Read-CsvData {
    # Streams the CSV with the .NET TextFieldParser instead of Import-Csv so
    # very large files cannot exhaust memory (Import-Csv materializes every row
    # as a PSCustomObject, which throws OutOfMemoryException on big datasets).
    # A single pass computes the row count and per-column null counts and value
    # frequencies over the FULL data, and draws a seeded reservoir sample of
    # rows (as lightweight string arrays) for statistics, plots, and clustering.
    param([string]$Path, [int]$SampleSize)

    Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop

    $parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($Path)
    $sample = New-Object System.Collections.Generic.List[object]
    $columns = @()
    $rowCount = 0
    $malformed = 0
    $profiles = @{}
    try {
        $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
        $parser.SetDelimiters(',')
        $parser.HasFieldsEnclosedInQuotes = $true
        $parser.TrimWhiteSpace = $false

        if ($parser.EndOfData) {
            return @{ Columns = @(); RowCount = 0; Profiles = @{}; SampleRows = $sample; MalformedRows = 0 }
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

        $nullCounts  = New-Object long[] $colCount
        $valueCounts = New-Object object[] $colCount
        $overflow    = New-Object bool[] $colCount
        for ($c = 0; $c -lt $colCount; $c++) { $valueCounts[$c] = @{} }

        $rand = New-Object System.Random(42)

        while (-not $parser.EndOfData) {
            $fields = $null
            try { $fields = $parser.ReadFields() }
            catch [Microsoft.VisualBasic.FileIO.MalformedLineException] { $malformed++; continue }

            for ($c = 0; $c -lt $colCount; $c++) {
                $v = $null
                if ($c -lt $fields.Length) { $v = $fields[$c] }
                if ([string]::IsNullOrWhiteSpace($v)) {
                    $nullCounts[$c]++
                }
                else {
                    $vc = $valueCounts[$c]
                    if ($vc.ContainsKey($v)) { $vc[$v]++ }
                    elseif ($vc.Count -lt $script:MaxDistinctTracked) { $vc[$v] = 1 }
                    else { $overflow[$c] = $true }
                }
            }

            # seeded reservoir sample keeps memory bounded regardless of file size
            if ($rowCount -lt $SampleSize) {
                $sample.Add($fields)
            }
            else {
                $j = $rand.Next($rowCount + 1)
                if ($j -lt $SampleSize) { $sample[$j] = $fields }
            }
            $rowCount++
            if (($rowCount % 25000) -eq 0) {
                Write-Progress -Activity 'Reading and profiling CSV (streaming)' -Status ("{0:N0} rows read" -f $rowCount)
            }
        }
        Write-Progress -Activity 'Reading and profiling CSV (streaming)' -Completed

        for ($c = 0; $c -lt $colCount; $c++) {
            $profiles[$columns[$c]] = @{
                Name             = $columns[$c]
                NullCount        = $nullCounts[$c]
                ValueCounts      = $valueCounts[$c]
                DistinctOverflow = $overflow[$c]
            }
        }
    }
    finally {
        $parser.Close()
        $parser.Dispose()
    }

    return @{
        Columns       = $columns
        RowCount      = $rowCount
        Profiles      = $profiles
        SampleRows    = $sample
        MalformedRows = $malformed
    }
}

function Get-InferredColumnType {
    # Samples up to 500 distinct values; >= 80% parse success decides the type.
    param($Profile)
    $keys = @($Profile.ValueCounts.Keys)
    if ($keys.Count -eq 0) { return 'Categorical' }
    $sample = $keys | Select-Object -First 500
    $numericHits = 0
    $dateHits    = 0
    $tested      = 0
    foreach ($v in $sample) {
        $tested++
        if ($null -ne (ConvertTo-NullableDouble $v)) { $numericHits++ }
        elseif ($null -ne (ConvertTo-NullableDate $v)) { $dateHits++ }
    }
    if ($tested -eq 0) { return 'Categorical' }
    if (($numericHits / $tested) -ge 0.8) { return 'Numeric' }
    if (($dateHits / $tested) -ge 0.8) { return 'Date' }
    return 'Categorical'
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
# V1 ANALYSIS SECTIONS
# ===========================================================================
function Get-MetadataSection {
    param([long]$RowCount, [string[]]$Columns, $Profiles, $ColumnTypes, [string]$FileSizeText)
    $section = New-Section 'Metadata Analysis'
    $rowCount = $RowCount
    $section.Paragraphs.Add(("The dataset contains {0:N0} rows and {1:N0} columns (file size: {2}). " -f $rowCount, $Columns.Count, $FileSizeText) +
        'Column types are inferred by sampling values (numeric, date, or categorical/string).')

    $rows = foreach ($col in $Columns) {
        $prof = $Profiles[$col]
        $distinct = $prof.ValueCounts.Count
        $distinctText = '{0:N0}' -f $distinct
        if ($prof.DistinctOverflow) { $distinctText = '> {0:N0}' -f $script:MaxDistinctTracked }
        $nullPct = 0
        if ($rowCount -gt 0) { $nullPct = 100.0 * $prof.NullCount / $rowCount }
        [PSCustomObject]@{
            'Column'          = $col
            'Inferred Type'   = $ColumnTypes[$col]
            'Null/Blank Count'= ('{0:N0}' -f $prof.NullCount)
            'Null %'          = ('{0:N2}%' -f $nullPct)
            'Distinct Values' = $distinctText
        }
    }
    Add-SectionTable $section 'Column metadata' $rows
    return $section
}

function Get-FrequencySection {
    param([long]$RowCount, [string[]]$Columns, $Profiles, $ColumnTypes)
    $section = New-Section 'Frequency Analysis (Categorical Columns)'
    $rowCount = $RowCount
    $catCols = @($Columns | Where-Object { $ColumnTypes[$_] -eq 'Categorical' })
    if ($catCols.Count -eq 0) {
        $section.Paragraphs.Add('No categorical columns were detected in this dataset.')
        return $section
    }
    $section.Paragraphs.Add('For each categorical column, the top 15 values by frequency are shown with counts and percentages of all rows.')
    if ($catCols.Count -gt $script:MaxFreqColumns) {
        $section.Paragraphs.Add(("Note: {0} categorical columns were detected; only the first {1} are tabulated to keep the report readable." -f $catCols.Count, $script:MaxFreqColumns))
        $catCols = @($catCols | Select-Object -First $script:MaxFreqColumns)
    }
    foreach ($col in $catCols) {
        $prof = $Profiles[$col]
        $top = $prof.ValueCounts.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First 15
        $rows = foreach ($entry in $top) {
            $pct = 0
            if ($rowCount -gt 0) { $pct = 100.0 * $entry.Value / $rowCount }
            [PSCustomObject]@{
                'Value'   = [string]$entry.Key
                'Count'   = ('{0:N0}' -f $entry.Value)
                'Percent' = ('{0:N2}%' -f $pct)
            }
        }
        Add-SectionTable $section ("Top 15 values: {0}" -f $col) $rows
    }
    return $section
}

function Get-ObservationsSection {
    # V0+ extra: automated data-quality observations to seed analyst discussion.
    param([long]$RowCount, [string[]]$Columns, $Profiles, $ColumnTypes)
    $section = New-Section 'Automated Observations & Recommendations'
    $section.Paragraphs.Add('Automatically generated data-quality observations intended as discussion points for business analysts.')

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($col in $Columns) {
        $prof = $Profiles[$col]
        $distinct = $prof.ValueCounts.Count
        $nullPct = 0.0
        if ($RowCount -gt 0) { $nullPct = 100.0 * $prof.NullCount / $RowCount }
        $nonNull = $RowCount - $prof.NullCount

        $notes = New-Object System.Collections.Generic.List[string]
        if ($nullPct -ge 90) {
            $notes.Add(('Almost entirely null/blank ({0:N1}%) - consider dropping the column or fixing data capture upstream.' -f $nullPct))
        }
        elseif ($nullPct -ge 20) {
            $notes.Add(('High null rate ({0:N1}%) - investigate whether blanks are expected.' -f $nullPct))
        }
        if ($distinct -eq 1 -and -not $prof.DistinctOverflow -and $nonNull -gt 0) {
            $notes.Add('Constant column (a single value) - adds no analytical signal.')
        }
        if ($prof.DistinctOverflow) {
            $notes.Add(('Very high cardinality (more than {0:N0} distinct values) - likely an identifier or free text.' -f $script:MaxDistinctTracked))
        }
        elseif ($nonNull -gt 1 -and $distinct -eq $nonNull) {
            $notes.Add('Every value is unique - likely an identifier/key rather than an analysis variable.')
        }
        if ($ColumnTypes[$col] -eq 'Numeric' -and $distinct -gt 0 -and $distinct -le 2 -and -not $prof.DistinctOverflow) {
            $notes.Add('Numeric with at most 2 distinct values - may be a flag/indicator; consider treating it as categorical.')
        }
        foreach ($note in $notes) {
            $rows.Add([PSCustomObject]@{ 'Column' = $col; 'Observation' = $note })
        }
    }
    if ($rows.Count -gt 0) {
        Add-SectionTable $section 'Data quality observations' $rows
    }
    else {
        $section.Paragraphs.Add('No notable data quality issues were detected (nulls, constant columns, or identifier-like columns).')
    }
    $section.Paragraphs.Add('Suggested next steps: confirm column meanings with the data owners, validate null-handling rules, and review the frequency and distribution sections for unexpected categories or extreme values.')
    return $section
}

function Get-DistributionSection {
    param([string[]]$NumericColumns, $NumericArrays, [bool]$Sampled = $false)
    $section = New-Section 'Distribution Analysis (Numeric Columns)'
    if ($NumericColumns.Count -eq 0) {
        $section.Paragraphs.Add('No numeric columns were detected in this dataset.')
        return $section
    }
    $section.Paragraphs.Add('Descriptive statistics per numeric column. Potential outliers are values beyond 1.5 x IQR from the 25th/75th percentiles.')
    if ($Sampled) {
        $section.Paragraphs.Add(("Statistics are computed on the random {0:N0}-row sample (the dataset exceeds that size)." -f $script:MaxSampleRows))
    }

    $rows = foreach ($col in $NumericColumns) {
        $vals = New-Object System.Collections.Generic.List[double]
        foreach ($v in $NumericArrays[$col]) { if ($null -ne $v) { $vals.Add([double]$v) } }
        if ($vals.Count -eq 0) { continue }
        $arr = $vals.ToArray()
        [array]::Sort($arr)
        $n      = $arr.Length
        $sum    = 0.0
        foreach ($v in $arr) { $sum += $v }
        $mean   = $sum / $n
        $ssd    = 0.0
        foreach ($v in $arr) { $ssd += ($v - $mean) * ($v - $mean) }
        $stdev  = 0.0
        if ($n -gt 1) { $stdev = [math]::Sqrt($ssd / ($n - 1)) }
        $p25    = Get-Percentile -SortedValues $arr -Percentile 25
        $median = Get-Percentile -SortedValues $arr -Percentile 50
        $p75    = Get-Percentile -SortedValues $arr -Percentile 75
        $iqr    = $p75 - $p25
        $lowFence  = $p25 - 1.5 * $iqr
        $highFence = $p75 + 1.5 * $iqr
        $outliers = 0
        foreach ($v in $arr) { if ($v -lt $lowFence -or $v -gt $highFence) { $outliers++ } }
        [PSCustomObject]@{
            'Column'   = $col
            'Count'    = ('{0:N0}' -f $n)
            'Min'      = (Format-Num $arr[0])
            'P25'      = (Format-Num $p25)
            'Median'   = (Format-Num $median)
            'Mean'     = (Format-Num $mean)
            'P75'      = (Format-Num $p75)
            'Max'      = (Format-Num $arr[$n - 1])
            'Std Dev'  = (Format-Num $stdev)
            'Outliers' = ('{0:N0}' -f $outliers)
        }
    }
    Add-SectionTable $section 'Descriptive statistics' $rows
    return $section
}

function Get-CorrelationSection {
    param([string[]]$NumericColumns, $NumericArrays, [bool]$Sampled = $false)
    $section = New-Section 'Correlation Matrix (Pearson)'
    if ($NumericColumns.Count -lt 2) {
        $section.Paragraphs.Add('Fewer than two numeric columns were detected; a correlation matrix is not applicable.')
        return $section
    }
    $section.Paragraphs.Add('Pearson correlation coefficients computed pairwise over rows where both values are present. Pairs with |r| > 0.7 are flagged below the matrix.')
    if ($Sampled) {
        $section.Paragraphs.Add(("Correlations are computed on the random {0:N0}-row sample (the dataset exceeds that size)." -f $script:MaxSampleRows))
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

    # Matrix table
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

    # Flag strong pairs and remember the strongest for v2 scatter plots
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
    }
    else {
        $section.Paragraphs.Add('No pairs exceeded the |r| > 0.7 threshold.')
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

    if (-not $script:ChartingAvailable) {
        $section.Paragraphs.Add("The charting assembly is unavailable in this environment; the time trend chart for column '$dateCol' was skipped.")
        return $section
    }

    # Aggregate record counts (and optional numeric sum) per period
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
        $section.Paragraphs.Add("Column '$dateCol' did not yield any parseable dates; the chart was skipped.")
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
    param([string[]]$Columns, $Profiles, $ColumnTypes)
    $section = New-Section 'Categorical Bar Graphs'
    $catCols = @($Columns | Where-Object {
            $ColumnTypes[$_] -eq 'Categorical' -and
            $Profiles[$_].ValueCounts.Count -ge 2
        } | Select-Object -First $script:MaxBarCharts)
    if ($catCols.Count -eq 0) {
        $section.Paragraphs.Add('No suitable categorical columns were found for bar charts.')
        return $section
    }
    if (-not $script:ChartingAvailable) {
        $section.Paragraphs.Add('The charting assembly is unavailable in this environment; bar charts were skipped. Refer to the Frequency Analysis tables instead.')
        return $section
    }
    $section.Paragraphs.Add(("Top-10 value frequencies for up to {0} categorical columns." -f $script:MaxBarCharts))

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
        # reverse so the most frequent value renders at the top of the bar chart
        [array]::Reverse($top)
        foreach ($entry in $top) {
            $label = [string]$entry.Key
            if ($label.Length -gt 35) { $label = $label.Substring(0, 32) + '...' }
            $idx = $series.Points.AddXY($label, $entry.Value)
        }
        $file = '{0}_bar_{1}.png' -f $script:ProjToken, (Get-SafeFileToken $col)
        $path = Save-AnalysisChart -Chart $chart -FileName $file
        Add-SectionImage $section ("Top-10 values: {0}" -f $col) $path
    }
    return $section
}

# ===========================================================================
# V2 ANALYSIS SECTIONS
# ===========================================================================
function Get-ScatterPlotSection {
    param($PlotData, $ColIndex, [string[]]$NumericColumns)
    $section = New-Section 'Scatter Plots (Most Correlated Pairs)'
    if ($NumericColumns.Count -lt 2) {
        $section.Paragraphs.Add('Fewer than two numeric columns were detected; scatter plots are not applicable.')
        return $section
    }
    if (-not $script:ChartingAvailable) {
        $section.Paragraphs.Add('The charting assembly is unavailable in this environment; scatter plots were skipped.')
        return $section
    }

    $pairs = @($script:TopCorrelatedPairs)
    if ($pairs.Count -eq 0) {
        # correlation section unavailable - fall back to adjacent column pairs
        $pairs = @()
        for ($i = 0; $i -lt ([math]::Min(5, $NumericColumns.Count - 1)); $i++) {
            $pairs += @{ A = $NumericColumns[$i]; B = $NumericColumns[$i + 1]; R = $null }
        }
        $section.Paragraphs.Add('Correlation results were unavailable; scatter plots show adjacent numeric column pairs instead.')
    }
    else {
        $section.Paragraphs.Add('Scatter plots for the top numeric column pairs ranked by |Pearson r|.')
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
    if (-not $script:ChartingAvailable) {
        $section.Paragraphs.Add('The charting assembly is unavailable in this environment; the pair plot grid was skipped.')
        return $section
    }
    $section.Paragraphs.Add(("Pairwise scatter plots across the first {0} numeric columns (diagonal cells show value histograms)." -f $cols.Count))

    # Pre-extract values per column (capped) keeping row alignment
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
                # diagonal: histogram of the column
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

# ===========================================================================
# V3 ANALYSIS SECTION - K-MEANS (implemented from scratch)
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
                # re-seed an empty cluster
                $centroids[$k] = $Points[$rand.Next($n)].Clone()
            }
        }
    }

    # --- WCSS ---
    $wcss = 0.0
    for ($i = 0; $i -lt $n; $i++) {
        $k = $assign[$i]
        for ($d = 0; $d -lt $dim; $d++) {
            $diff = $Points[$i][$d] - $centroids[$k][$d]
            $wcss += $diff * $diff
        }
    }
    return @{ Assignments = $assign; Centroids = $centroids; WCSS = $wcss }
}

function Get-KMeansSection {
    param($PlotData, $ColIndex, [string[]]$NumericColumns)
    $section = New-Section 'K-Means Clustering'
    if ($NumericColumns.Count -lt 2) {
        $section.Paragraphs.Add('Fewer than two numeric columns were detected; k-means clustering is not applicable.')
        return $section
    }

    Write-Stage 'Running k-means clustering (this can take a moment)...'
    # Build complete-case matrix of raw values (cap dimensions for tractability)
    $cols = @($NumericColumns | Select-Object -First 8)
    $colIdxs = New-Object int[] $cols.Count
    for ($c = 0; $c -lt $cols.Count; $c++) { $colIdxs[$c] = [int]$ColIndex[$cols[$c]] }
    $rawRows = New-Object System.Collections.Generic.List[object]
    foreach ($row in $PlotData) {
        $vec = New-Object double[] $cols.Count
        $ok = $true
        for ($c = 0; $c -lt $cols.Count; $c++) {
            $v = $null
            if ($colIdxs[$c] -lt $row.Length) { $v = ConvertTo-NullableDouble ([string]$row[$colIdxs[$c]]) }
            if ($null -eq $v) { $ok = $false; break }
            $vec[$c] = [double]$v
        }
        if ($ok) { $rawRows.Add($vec) }
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

    # Standardize (z-score)
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

    # Elbow heuristic over k = 2..6
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

    # Elbow: keep increasing k while the relative WCSS improvement stays >= 15%
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

    # Cluster sizes
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

    # Per-cluster means in ORIGINAL units
    $clusterSums = New-Object 'object[]' $bestK
    for ($k = 0; $k -lt $bestK; $k++) { $clusterSums[$k] = New-Object double[] $dim }
    for ($i = 0; $i -lt $nPts; $i++) {
        $k = $final.Assignments[$i]
        for ($d = 0; $d -lt $dim; $d++) { $clusterSums[$k][$d] += $rawRows[$i][$d] }
    }
    $meanRows = for ($k = 0; $k -lt $bestK; $k++) {
        $o = [ordered]@{ 'Cluster' = ('Cluster {0}' -f ($k + 1)) }
        for ($d = 0; $d -lt $dim; $d++) {
            $m = 0.0
            if ($sizes[$k] -gt 0) { $m = $clusterSums[$k][$d] / $sizes[$k] }
            $o[$cols[$d]] = (Format-Num $m 2)
        }
        [PSCustomObject]$o
    }
    Add-SectionTable $section 'Per-cluster means (original units)' $meanRows

    # Cluster scatter plot on the first two numeric variables
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
    [void]$sb.AppendLine(('<title>{0} Automated Analysis ({1})</title>' -f (HtmlEnc $Meta.ProjName), (HtmlEnc $Meta.Version)))
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

    # Table of contents
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
            $rows = @($tbl.Rows)
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

    [void]$sb.AppendLine(('<p class="footer">{0} automated analysis &middot; generated {1} &middot; Prepared by Manuel Figallo</p>' -f (HtmlEnc $Meta.ProjName), (HtmlEnc $Meta.Timestamp)))
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

                $rows  = @($tbl.Rows)
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

    Write-Host ''
    Write-Host ('=' * 70) -ForegroundColor DarkCyan
    Write-Host (" {0} - Automated Data Analysis ({1})" -f $PROJ_NAME, $version) -ForegroundColor White
    Write-Host ' Prepared by Manuel Figallo' -ForegroundColor Gray
    Write-Host ('=' * 70) -ForegroundColor DarkCyan

    # ----- Load + profile data (streamed; Import-Csv loads every row into
    #       memory as an object and throws OutOfMemoryException on large files) -----
    $fileSizeMB = (Get-Item -LiteralPath $PATH_TO_DATA).Length / 1MB
    Write-Stage ("Loading CSV (streaming, {0:N1} MB): {1}" -f $fileSizeMB, $PATH_TO_DATA)
    try {
        $csv = Read-CsvData -Path $PATH_TO_DATA -SampleSize $script:MaxSampleRows
    }
    catch {
        Write-Error ("Failed to load CSV file '{0}': {1}" -f $PATH_TO_DATA, $_.Exception.Message) -ErrorAction Continue
        exit 2
    }
    $Columns  = @($csv.Columns)
    $RowCount = [long]$csv.RowCount
    $Profiles = $csv.Profiles
    if ($RowCount -eq 0) {
        Write-Error ("The CSV file '{0}' contains no data rows." -f $PATH_TO_DATA) -ErrorAction Continue
        exit 2
    }
    if ($Columns.Count -eq 0) {
        Write-Error ("No parseable columns were found in '{0}'." -f $PATH_TO_DATA) -ErrorAction Continue
        exit 2
    }
    if ($csv.MalformedRows -gt 0) {
        Write-Warning ("{0:N0} malformed CSV line(s) were skipped during loading." -f $csv.MalformedRows)
    }
    $ColIndex = @{}
    for ($i = 0; $i -lt $Columns.Count; $i++) { $ColIndex[$Columns[$i]] = $i }
    Write-Stage ("Loaded {0:N0} rows x {1} columns ({2:N1} MB)." -f $RowCount, $Columns.Count, $fileSizeMB) 'Green'

    Write-Stage 'Inferring column types by sampling values...'
    $ColumnTypes = @{}
    foreach ($col in $Columns) { $ColumnTypes[$col] = Get-InferredColumnType -Profile $Profiles[$col] }
    $NumericColumns = @($Columns | Where-Object { $ColumnTypes[$_] -eq 'Numeric' })
    $DateColumns    = @($Columns | Where-Object { $ColumnTypes[$_] -eq 'Date' })
    Write-Stage ("Detected {0} numeric, {1} date, {2} categorical column(s)." -f `
        $NumericColumns.Count, $DateColumns.Count, ($Columns.Count - $NumericColumns.Count - $DateColumns.Count))

    # ----- Sampled rows (reservoir drawn during streaming) -----
    $sampled = ($RowCount -gt $script:MaxSampleRows)
    $PlotData = $csv.SampleRows
    if ($sampled) {
        Write-Stage ("Dataset exceeds {0:N0} rows; statistics, charts, and clustering use a random {0:N0}-row reservoir sample (metadata and frequency counts cover the full data)." -f $script:MaxSampleRows) 'Yellow'
    }

    # ----- Parse numeric columns once (sampled rows; feeds stats + correlation) -----
    $NumericArrays = @{}
    if ($NumericColumns.Count -gt 0) {
        Write-Stage 'Parsing numeric columns...'
        $ci = 0
        foreach ($col in $NumericColumns) {
            $ci++
            Write-Progress -Activity 'Parsing numeric columns' -Status $col -PercentComplete (100 * $ci / $NumericColumns.Count)
            $NumericArrays[$col] = Get-NumericColumnArray -Rows $PlotData -Index ([int]$ColIndex[$col])
        }
        Write-Progress -Activity 'Parsing numeric columns' -Completed
    }

    # ----- Charting -----
    Initialize-Charting

    # ----- Report metadata -----
    $scopeNote = ("Dataset: {0:N0} rows, {1} columns, {2:N1} MB." -f $RowCount, $Columns.Count, $fileSizeMB)
    if ($sampled) {
        $scopeNote += (" Note: because the dataset exceeds {0:N0} rows, descriptive statistics, correlations, charts, and clustering use a random {0:N0}-row sample; metadata and frequency counts cover the full dataset." -f $script:MaxSampleRows)
    }
    $Meta = @{
        ProjName  = $PROJ_NAME
        Version   = $version
        DataPath  = $PATH_TO_DATA
        Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        ScopeNote = $scopeNote
    }

    # ----- Run analysis sections per version (cumulative; v0 is the baseline) -----
    $Sections = New-Object System.Collections.Generic.List[object]

    Write-Stage 'Running V0 analysis: metadata...'
    $Sections.Add((Invoke-SafeSection 'Metadata Analysis' {
        Get-MetadataSection -RowCount $RowCount -Columns $Columns -Profiles $Profiles -ColumnTypes $ColumnTypes -FileSizeText ('{0:N1} MB' -f $fileSizeMB) }))

    Write-Stage 'Running V0 analysis: frequency tables...'
    $Sections.Add((Invoke-SafeSection 'Frequency Analysis (Categorical Columns)' {
        Get-FrequencySection -RowCount $RowCount -Columns $Columns -Profiles $Profiles -ColumnTypes $ColumnTypes }))

    Write-Stage 'Running V0 analysis: distributions...'
    $Sections.Add((Invoke-SafeSection 'Distribution Analysis (Numeric Columns)' {
        Get-DistributionSection -NumericColumns $NumericColumns -NumericArrays $NumericArrays -Sampled $sampled }))

    Write-Stage 'Running V0 analysis: automated observations...'
    $Sections.Add((Invoke-SafeSection 'Automated Observations & Recommendations' {
        Get-ObservationsSection -RowCount $RowCount -Columns $Columns -Profiles $Profiles -ColumnTypes $ColumnTypes }))

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
                Get-BarChartSection -Columns $Columns -Profiles $Profiles -ColumnTypes $ColumnTypes }))
        }
        else { $Sections.Add((New-ChartsUnavailableSection 'Categorical Bar Graphs')) }
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
    }

    if ($versionLevel -ge 3) {
        Write-Stage 'Running V3 analysis: k-means clustering...'
        $Sections.Add((Invoke-SafeSection 'K-Means Clustering' {
            Get-KMeansSection -PlotData $PlotData -ColIndex $ColIndex -NumericColumns $NumericColumns }))
    }

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
