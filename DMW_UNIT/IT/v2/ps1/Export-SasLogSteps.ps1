<#
.SYNOPSIS
  Script Name   : Export-SasLogSteps.ps1
  Purpose       : Run the cgs_ai scanFileSystem.ps1 with -metric_profile sas_log
                  and flatten its Metrics sheet to CSV for the DMW
                  "Visualize Performance" tab.
  Version       : 1.0.0

  -----------------------------------------------------------------------------
  WHY A WRAPPER RATHER THAN A SECOND PARSER
  -----------------------------------------------------------------------------
  cgs_ai\src\ps\scanFileSystem.ps1 already parses SAS logs into per-step
  timings, and it has a Python twin with matching parameters and columns. It is
  the single source of truth for that parse, so nothing here re-implements it.

  Two things make a wrapper necessary rather than optional:

    1. The cgs_ai script writes its metrics to an XLSX SHEET, not to CSV. Its
       own code is explicit about this: when the output path does not end in
       .xlsx it writes ONLY the match rows, and the metric rows are dropped.
       So asking it for a .csv silently produces a file with no timings in it.

    2. A browser cannot read .xlsx. The DMW page reads CSV.

  So this script asks for .xlsx, then extracts the "Metrics" sheet to CSV.

  -----------------------------------------------------------------------------
  RELATIONSHIP TO Export-SasLogRollup.ps1
  -----------------------------------------------------------------------------
  Two sources, two grains, both feeding the same tab:

      THIS (cgs_ai)                 one row per SAS STEP
                                    "where does the time go inside a program"
      Export-SasLogRollup.ps1       one row per LOG FILE
                                    "which programs cost most, which fail"

  Neither derives the other.

  -----------------------------------------------------------------------------
  OUTPUT -> data\sas_log_steps.csv
      FullPath, ProgramName, StepIndex, StepLabel, RealTimeSec, CpuTimeSec
  These are the cgs_ai column names, unchanged, so the CSV stays recognisable
  to anyone who knows that tool.
  -----------------------------------------------------------------------------

  .PARAMETER ScanScript
    Path to cgs_ai\src\ps\scanFileSystem.ps1. REQUIRED unless the default
    below is correct for your checkout. Note the cgs_ai script dot-sources
    cgsUtils.ps1 from its own folder, so it must be run where it lives --
    copying it alone to a share will fail.

  .PARAMETER Root
    Log directories to scan. Passed through as -input_folder_root.

  .PARAMETER Keyword
    The cgs_ai script requires -extract_keyword even when you only want
    metrics. Metric rows are collected per FILE regardless of whether any
    keyword matched, so this does not filter the timings -- but it does drive
    the (discarded) Matches sheet, so keep it cheap.

  .PARAMETER KeepWorkbook
    Keep the intermediate .xlsx instead of deleting it. Useful when you also
    want the Matches sheet.

  .EXAMPLE
    .\Export-SasLogSteps.ps1

  .EXAMPLE
    .\Export-SasLogSteps.ps1 -Root "\\srv\sas\UNIT\logs" -KeepWorkbook

  DEPENDENCIES
    None beyond Windows PowerShell 5.1. If the ImportExcel module is present
    it is used to read the workbook; if not, the workbook is unzipped and its
    sheet XML parsed with the built-in .NET classes. Nothing is installed.
#>

#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]   $ScanScript   = "",
    [string[]] $Root         = @(),
    [string[]] $Keyword      = @("NOTE:"),
    [string]   $OutputPath   = "",
    [string]   $WorkbookPath = "",
    [switch]   $KeepWorkbook,
    [switch]   $TestMode
)

###############################################################################
#region CONFIGURATION  -- the only section an operator should need to edit
###############################################################################

# Where the cgs_ai scanner lives. Adjust to your checkout.
$DefaultScanScript = "$PSScriptRoot\..\..\..\cgs_ai\src\ps\scanFileSystem.ps1"

$DefaultRoots = @(
    "\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\Web_Stored_Processes\Logs"
    "\\A70admed.com\r1\CGS\APPS\SAS\PROD\SAS_G\Web_Stored_Processes\Logs"
)

$DataDir      = "$PSScriptRoot\..\data"
$DataFileName = "sas_log_steps.csv"

$LogDir   = "$PSScriptRoot\..\log"
$LogLevel = "INFO"

# Column order written to the CSV. These are the cgs_ai names, unchanged.
$StepColumns = @('FullPath','ProgramName','StepIndex','StepLabel','RealTimeSec','CpuTimeSec')

# The sheet the cgs_ai script writes its metrics to.
$MetricSheetName = 'Metrics'

#endregion CONFIGURATION

###############################################################################
#region LOGGING
###############################################################################
$script:LogLevelMap = @{ DEBUG=0; INFO=1; WARN=2; ERROR=3 }
function Write-RunLog {
    param([string]$Message, [string]$Level="INFO")
    try {
        if ($script:LogLevelMap[$Level] -lt $script:LogLevelMap[$LogLevel]) { return }
        if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
        $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        "[$ts] [$Level] $Message" | Out-File -FilePath (Join-Path $LogDir ("Export-SasLogSteps_{0}.log" -f (Get-Date -Format "yyyyMMdd"))) -Append -Encoding UTF8
        if ($Level -eq "ERROR") { Write-Host "[$ts] [$Level] $Message" -ForegroundColor Red }
        elseif ($Level -eq "WARN") { Write-Host "[$ts] [$Level] $Message" -ForegroundColor Yellow }
    } catch { }
}
#endregion LOGGING

###############################################################################
#region XLSX READING  (no module required)
###############################################################################

function Read-XlsxSheet {
    <#
      .SYNOPSIS Read one worksheet from an .xlsx into PSCustomObjects.
      .NOTES    An .xlsx is a zip of XML. The first row is treated as the
                header. Values may be inline, numeric, or an index into
                sharedStrings.xml, and all three are handled.

                Used only when ImportExcel is unavailable, so the dashboard
                can be fed on a locked-down box with nothing installed.
    #>
    param([string]$Path, [string]$SheetName)

    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        function Get-EntryXml {
            param($Archive, [string]$EntryName)
            $e = $Archive.Entries | Where-Object { $_.FullName -eq $EntryName }
            if (-not $e) { return $null }
            $sr = New-Object System.IO.StreamReader($e.Open())
            try { return [xml]$sr.ReadToEnd() } finally { $sr.Dispose() }
        }

        # Shared strings table (optional).
        $shared = @()
        $ssXml = Get-EntryXml -Archive $zip -EntryName 'xl/sharedStrings.xml'
        if ($ssXml) {
            foreach ($si in $ssXml.sst.si) {
                if ($si.t -is [string])      { $shared += $si.t }
                elseif ($si.t.'#text')       { $shared += $si.t.'#text' }
                elseif ($si.r)               { $shared += (($si.r | ForEach-Object { $_.t.'#text'; $_.t }) -join '') }
                else                         { $shared += '' }
            }
        }

        # Find the sheet's target file via the workbook relationships.
        $wb  = Get-EntryXml -Archive $zip -EntryName 'xl/workbook.xml'
        if (-not $wb) { throw "xl/workbook.xml not found in $Path" }
        $sheet = $wb.workbook.sheets.sheet | Where-Object { $_.name -eq $SheetName }
        if (-not $sheet) {
            throw ("sheet '{0}' not found. Sheets present: {1}" -f $SheetName,
                   (($wb.workbook.sheets.sheet | ForEach-Object { $_.name }) -join ', '))
        }
        $rid  = $sheet.id       # r:id attribute
        $rels = Get-EntryXml -Archive $zip -EntryName 'xl/_rels/workbook.xml.rels'
        $target = ($rels.Relationships.Relationship | Where-Object { $_.Id -eq $rid }).Target
        if (-not $target) { throw "relationship $rid not resolved" }
        $entryName = 'xl/' + ($target -replace '^/xl/', '' -replace '^/', '')

        $sx = Get-EntryXml -Archive $zip -EntryName $entryName
        if (-not $sx) { throw "sheet part $entryName not found" }

        $cellValue = {
            param($c)
            if ($null -eq $c) { return '' }
            $t = $c.t
            if ($t -eq 's')      { $i = [int]$c.v; if ($i -lt $shared.Count) { return $shared[$i] } else { return '' } }
            elseif ($t -eq 'inlineStr') { return [string]$c.is.t }
            elseif ($t -eq 'str') { return [string]$c.v }
            else { return [string]$c.v }
        }
        $colLetters = { param([string]$ref) ($ref -replace '[0-9]','') }

        $rows = @($sx.worksheet.sheetData.row)
        if (-not $rows -or $rows.Count -lt 1) { return @() }

        # Header row -> column letter => name
        $headerMap = @{}
        foreach ($c in @($rows[0].c)) {
            $letter = & $colLetters $c.r
            $headerMap[$letter] = (& $cellValue $c)
        }

        $out = New-Object System.Collections.Generic.List[object]
        foreach ($r in $rows[1..($rows.Count-1)]) {
            $o = [ordered]@{}
            foreach ($name in $headerMap.Values) { if ($name) { $o[$name] = '' } }
            foreach ($c in @($r.c)) {
                $letter = & $colLetters $c.r
                $name = $headerMap[$letter]
                if ($name) { $o[$name] = (& $cellValue $c) }
            }
            if (($o.Values | Where-Object { "$_" -ne '' }).Count -gt 0) { $out.Add([PSCustomObject]$o) }
        }
        return , $out.ToArray()
    }
    finally { $zip.Dispose() }
}

#endregion XLSX READING

###############################################################################
#region MAIN
###############################################################################

$scan  = if ($ScanScript) { $ScanScript } else { $DefaultScanScript }
$roots = if ($Root)       { $Root }       else { $DefaultRoots }
$out   = if ($OutputPath) { $OutputPath } else { Join-Path $DataDir $DataFileName }
$book  = if ($WorkbookPath) { $WorkbookPath } else { Join-Path $env:TEMP ("dmw_saslog_{0}.xlsx" -f (Get-Date -Format 'yyyyMMdd_HHmmss')) }

Write-RunLog "Export-SasLogSteps start; scanner=$scan"

if (-not (Test-Path $scan)) {
    Write-RunLog "cgs_ai scanFileSystem.ps1 not found at: $scan" -Level "ERROR"
    Write-Host "Pass -ScanScript with the path to cgs_ai\src\ps\scanFileSystem.ps1." -ForegroundColor Yellow
    exit 2
}
# The cgs_ai script dot-sources cgsUtils.ps1 from its own directory.
$utils = Join-Path (Split-Path $scan -Parent) 'cgsUtils.ps1'
if (-not (Test-Path $utils)) {
    Write-RunLog "cgsUtils.ps1 not found beside the scanner. It must be run where it lives." -Level "ERROR"
    exit 2
}

$reachable = @($roots | Where-Object { Test-Path $_ })
if (-not $reachable) {
    Write-RunLog ("none of the roots are reachable: {0}" -f ($roots -join '; ')) -Level "ERROR"
    exit 3
}
if ($reachable.Count -lt $roots.Count) {
    Write-RunLog ("{0} of {1} roots unreachable; continuing with the rest" -f ($roots.Count - $reachable.Count), $roots.Count) -Level "WARN"
}

# --- run the cgs_ai scanner ----------------------------------------------
# .xlsx is REQUIRED: with any other extension it writes the Matches sheet
# only and drops every metric row.
Write-RunLog ("invoking cgs_ai scanner -> {0}" -f $book)
& $scan -input_folder_root $reachable `
        -extract_keyword   $Keyword `
        -output_file_path  $book `
        -file_extensions   @('log') `
        -metric_profile    'sas_log'
$rc = $LASTEXITCODE
Write-RunLog "cgs_ai scanner exit code $rc"
if ($rc -ne 0) { Write-RunLog "scanner reported a failure; see its own output" -Level "WARN" }
if (-not (Test-Path $book)) {
    Write-RunLog "scanner produced no workbook at $book" -Level "ERROR"
    exit 3
}

# --- read the Metrics sheet ----------------------------------------------
$rows = @()
try {
    if (Get-Module -ListAvailable -Name ImportExcel) {
        Import-Module ImportExcel -ErrorAction Stop
        Write-RunLog "reading '$MetricSheetName' with ImportExcel"
        $rows = @(Import-Excel -Path $book -WorksheetName $MetricSheetName -ErrorAction Stop)
    } else {
        Write-RunLog "ImportExcel not available; reading the workbook directly"
        $rows = @(Read-XlsxSheet -Path $book -SheetName $MetricSheetName)
    }
} catch {
    Write-RunLog ("could not read the '{0}' sheet: {1}" -f $MetricSheetName, $_.Exception.Message) -Level "ERROR"
    Write-Host "If the scan found no SAS step timings the sheet is not written at all." -ForegroundColor Yellow
    if (-not $KeepWorkbook) { Remove-Item $book -Force -ErrorAction SilentlyContinue }
    exit 3
}

if (-not $rows -or $rows.Count -eq 0) {
    Write-RunLog "the Metrics sheet is empty; no step timings were found" -Level "WARN"
}

# Normalise to the documented column order; anything extra is dropped so the
# dashboard contract stays fixed.
$normalised = foreach ($r in $rows) {
    $o = [ordered]@{}
    foreach ($c in $StepColumns) { $o[$c] = $(if ($null -ne $r.$c) { $r.$c } else { '' }) }
    [PSCustomObject]$o
}

if ($TestMode) {
    $normalised | Select-Object -First 20 | Format-Table -AutoSize
    Write-RunLog "TestMode: CSV not written."
    if (-not $KeepWorkbook) { Remove-Item $book -Force -ErrorAction SilentlyContinue }
    return
}

try {
    $dir = Split-Path $out -Parent
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $normalised | Export-Csv -Path $out -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
    Write-RunLog ("wrote {0} step row(s) to {1}" -f @($normalised).Count, $out)
} catch {
    Write-RunLog "CSV write failed: $($_.Exception.Message)" -Level "ERROR"
    exit 3
}

if ($KeepWorkbook) { Write-RunLog "workbook kept at $book" }
else { Remove-Item $book -Force -ErrorAction SilentlyContinue }

exit 0

#endregion MAIN
