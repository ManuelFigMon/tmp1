<#
.SYNOPSIS
  Script Name   : Export-SasLogRollup.ps1
  Purpose       : Roll SAS .log files up to ONE ROW PER FILE for the DMW
                  "Visualize Performance" tab.
  Version       : 1.1.0

  -----------------------------------------------------------------------------
  RELATIONSHIP TO cgs_ai\src\ps\scanFileSystem.ps1  -- READ THIS FIRST
  -----------------------------------------------------------------------------
  This script was previously, and wrongly, named scanFileSystem.ps1. That name
  belongs to the cgs_ai tool, which is a different program with a different
  output contract and a Python twin. The name has been corrected; the two now
  coexist and feed DMW as two sources:

      cgs_ai scanFileSystem.ps1 -metric_profile sas_log
          GRAIN  one row per SAS STEP
          GIVES  FullPath, ProgramName, StepIndex, StepLabel,
                 RealTimeSec, CpuTimeSec
          USE    "where does the time go inside a program"
          -> DMW reads it via Export-SasLogSteps.ps1, which runs the cgs_ai
             script and flattens its Metrics sheet to CSV.

      THIS SCRIPT
          GRAIN  one row per LOG FILE
          GIVES  session totals plus memory, step counts, error and warning
                 counts, top PROC and librefs
          USE    "which programs cost the most, and which are failing"

  They are complementary, not duplicates. Neither derives the other: the
  cgs_ai script does not capture memory, errors or session totals, and this
  one does not capture per-step timings.

  -----------------------------------------------------------------------------
  WHY THIS ONE STILL EXISTS
  -----------------------------------------------------------------------------
  The DMW dashboard needs per-program cost and failure counts. Those are not in
  the cgs_ai output, so they have to be parsed separately. Two further
  differences are deliberate:

    * ENCODING. SAS on Windows writes the ANSI code page. This script reads
      Windows-1252; the cgs_ai script falls back to ISO-8859-1, which differs
      in 0x80-0x9F and turns a smart quote or em dash into a control character.

    * PROGRAM NAME. This script strips TRAILING schedule and version suffixes
      (_1200PM, _V1_1, _20261110) so repeated runs of one program roll up to a
      single name. The cgs_ai script strips a LEADING keyword token instead,
      which would scatter those runs across separate rows.

  -----------------------------------------------------------------------------
  OUTPUT -> data\sas_log_rollup.csv   (read by the DMW Visualize Performance tab)
  -----------------------------------------------------------------------------

  .PARAMETER metric_profile
    sas_log | file_inventory. Named to mirror the cgs_ai parameter.

  .PARAMETER Root
    One or more directories to scan. Defaults to $DefaultRoots below.

  .PARAMETER OutputPath
    Target CSV. Defaults to $DataDir\$DataFileName.

  .PARAMETER MaxFiles
    Safety cap. 0 = unlimited.

  .EXAMPLE
    .\Export-SasLogRollup.ps1 -metric_profile sas_log

  .EXAMPLE
    .\Export-SasLogRollup.ps1 -metric_profile sas_log -Root "\\srv\sas\logs" -MaxFiles 500
#>

#Requires -Version 5.1
[CmdletBinding()]
param(
    [ValidateSet("sas_log","file_inventory")]
    [string]$metric_profile = "sas_log",

    [string[]]$Root,
    [string]$OutputPath,
    [int]$MaxFiles = 0,
    [switch]$TestMode
)

###############################################################################
#region CONFIGURATION
###############################################################################

$DefaultRoots = @(
    "\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\Web_Stored_Processes\Logs"
    "\\A70admed.com\r1\CGS\APPS\SAS\PROD\SAS_G\Web_Stored_Processes\Logs"
)

$Extensions  = @("log")        # sas_log profile
$MaxAgeDays  = 180             # ignore logs older than this (0 = no limit)

$DataDir      = "$PSScriptRoot\..\data"
$DataFileName = "sas_log_rollup.csv"

$LogDir   = "$PSScriptRoot\..\log"
$LogLevel = "INFO"

# Strip scheduling and version suffixes so repeated runs of the same program
# roll up to one ProgramName on the dashboard.
#   DME_Top_Denials_V1_1_1200PM.log -> DME_Top_Denials
$SuffixPatterns = @(
    '_\d{3,5}(AM|PM)$'      # _1200PM
    '_[Vv]\d+(_\d+)*$'      # _V1_1
    '_\d{8}(_\d{6})?$'      # _20261110_031500
)

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
        "[$ts] [$Level] $Message" | Out-File -FilePath (Join-Path $LogDir ("Export-SasLogRollup_{0}.log" -f (Get-Date -Format "yyyyMMdd"))) -Append -Encoding UTF8
        if ($Level -eq "ERROR") { Write-Host "[$ts] [$Level] $Message" -ForegroundColor Red }
    } catch { }
}
#endregion LOGGING

###############################################################################
#region PARSING
###############################################################################

function ConvertTo-Seconds {
    <#
      Normalises the three SAS duration shapes to decimal seconds:
        "12.34"        -> 12.34
        "2:45.12"      -> 165.12
        "1:02:45.12"   -> 3765.12
      Returns $null when the text does not match any of them.
    #>
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $t = $Text.Trim()
    $parts = $t -split ':'
    try {
        switch ($parts.Count) {
            1 { return [double]$parts[0] }
            2 { return ([double]$parts[0] * 60) + [double]$parts[1] }
            3 { return ([double]$parts[0] * 3600) + ([double]$parts[1] * 60) + [double]$parts[2] }
            default { return $null }
        }
    } catch { return $null }
}

function Get-ProgramName {
    param([string]$BaseName)
    $n = $BaseName
    foreach ($p in $SuffixPatterns) { $n = $n -replace $p, '' }
    return ($n.TrimEnd('_','.','-'))
}

function Read-SasLog {
    param([System.IO.FileInfo]$File)

    try {
        $enc  = [System.Text.Encoding]::GetEncoding(1252)   # SAS writes ANSI, not UTF-8
        $text = [System.IO.File]::ReadAllText($File.FullName, $enc)
    } catch {
        Write-RunLog "Unreadable: $($File.FullName) - $($_.Exception.Message)" -Level "WARN"
        return [PSCustomObject]@{ ParseStatus = "PARSE_ERROR"; ErrorText = $_.Exception.Message }
    }

    $lines = $text -split "`r?`n"

    # Session totals: the LAST "NOTE: The SAS System used:" block wins.
    $realSec = $null; $cpuSec = $null; $memKb = $null; $osMemKb = $null
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match 'NOTE:\s+The SAS System used:') {
            for ($j = $i + 1; $j -lt [Math]::Min($i + 8, $lines.Count); $j++) {
                if ($lines[$j] -match 'real time\s+([\d:.]+)')      { $realSec = ConvertTo-Seconds $Matches[1] }
                elseif ($lines[$j] -match 'cpu time\s+([\d:.]+)')   { $cpuSec  = ConvertTo-Seconds $Matches[1] }
                elseif ($lines[$j] -match 'memory\s+([\d.]+)k')     { $memKb   = [double]$Matches[1] }
                elseif ($lines[$j] -match 'OS Memory\s+([\d.]+)k')  { $osMemKb = [double]$Matches[1] }
            }
        }
    }

    # Step counts and per-step CPU, for the "CPU by program" view.
    $procCount = ([regex]::Matches($text, 'NOTE:\s+PROCEDURE\s+(\w+)\s+used')).Count
    $dataCount = ([regex]::Matches($text, 'NOTE:\s+DATA statement used')).Count
    $errCount  = ([regex]::Matches($text, '(?m)^ERROR')).Count
    $warnCount = ([regex]::Matches($text, '(?m)^WARNING')).Count

    # Most frequent procedure, a cheap signal of what the program actually does.
    $procs = [regex]::Matches($text, 'NOTE:\s+PROCEDURE\s+(\w+)\s+used') |
             ForEach-Object { $_.Groups[1].Value } |
             Group-Object | Sort-Object Count -Descending
    $topProc = if ($procs) { $procs[0].Name } else { "" }

    # LIBREF / library references, used to attribute a program to a data source.
    $librefs = [regex]::Matches($text, "(?m)Libref\s+(\w+)\s+was successfully assigned") |
               ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique
    $librefCount = @($librefs).Count

    $status = if ($null -eq $realSec -and $procCount -eq 0 -and $dataCount -eq 0) { "NO_STEPS_FOUND" } else { "OK" }

    [PSCustomObject]@{
        ParseStatus     = $status
        RealTimeSec     = $realSec
        CpuTimeSec      = $cpuSec
        MemoryKb        = $memKb
        OsMemoryKb      = $osMemKb
        ProcStepCount   = $procCount
        DataStepCount   = $dataCount
        TotalStepCount  = $procCount + $dataCount
        ErrorCount      = $errCount
        WarningCount    = $warnCount
        TopProcedure    = $topProc
        LibrefCount     = $librefCount
        Librefs         = ($librefs -join ';')
        ErrorText       = ""
    }
}

#endregion PARSING

###############################################################################
#region MAIN
###############################################################################

$roots = if ($Root) { $Root } else { $DefaultRoots }
$out   = if ($OutputPath) { $OutputPath } else { Join-Path $DataDir $DataFileName }

Write-RunLog "Export-SasLogRollup start: profile=$metric_profile roots=$($roots -join '; ')"

$files = @()
foreach ($r in $roots) {
    if (-not (Test-Path $r)) { Write-RunLog "Root not reachable, skipping: $r" -Level "WARN"; continue }
    try {
        $found = Get-ChildItem -Path $r -Recurse -File -ErrorAction SilentlyContinue |
                 Where-Object { $Extensions -contains $_.Extension.TrimStart('.').ToLower() }
        if ($MaxAgeDays -gt 0) {
            $cut = (Get-Date).AddDays(-$MaxAgeDays)
            $found = $found | Where-Object { $_.LastWriteTime -ge $cut }
        }
        $files += $found
        Write-RunLog "$r -> $(@($found).Count) files"
    } catch { Write-RunLog "Scan failed for $r : $($_.Exception.Message)" -Level "ERROR" }
}

if ($MaxFiles -gt 0 -and $files.Count -gt $MaxFiles) {
    $files = $files | Sort-Object LastWriteTime -Descending | Select-Object -First $MaxFiles
    Write-RunLog "Capped at $MaxFiles most recent files"
}

if (-not $files -or $files.Count -eq 0) {
    Write-RunLog "No files matched. Nothing written." -Level "WARN"
    exit 0
}

$rows = foreach ($f in $files) {

    $base = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
    $envName = if ($f.FullName -match '\\UNIT\\') { "UNIT" } elseif ($f.FullName -match '\\PROD\\') { "PROD" } else { "UNKNOWN" }

    if ($metric_profile -eq "file_inventory") {
        [PSCustomObject]@{
            ScanTimestamp = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
            MetricProfile = $metric_profile
            Environment   = $envName
            ProgramName   = Get-ProgramName $base
            FileName      = $f.Name
            FullPath      = $f.FullName
            SizeBytes     = $f.Length
            LastModified  = $f.LastWriteTime.ToString("yyyy-MM-ddTHH:mm:ss")
            AgeDays       = [int]((Get-Date) - $f.LastWriteTime).TotalDays
        }
        continue
    }

    $m = Read-SasLog -File $f
    [PSCustomObject]@{
        ScanTimestamp   = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
        MetricProfile   = $metric_profile
        Environment     = $envName
        ProgramName     = Get-ProgramName $base
        FileName        = $f.Name
        FullPath        = $f.FullName
        LastModified    = $f.LastWriteTime.ToString("yyyy-MM-ddTHH:mm:ss")
        SizeBytes       = $f.Length
        ParseStatus     = $m.ParseStatus
        RealTimeSec     = $m.RealTimeSec
        CpuTimeSec      = $m.CpuTimeSec
        MemoryKb        = $m.MemoryKb
        OsMemoryKb      = $m.OsMemoryKb
        ProcStepCount   = $m.ProcStepCount
        DataStepCount   = $m.DataStepCount
        TotalStepCount  = $m.TotalStepCount
        ErrorCount      = $m.ErrorCount
        WarningCount    = $m.WarningCount
        TopProcedure    = $m.TopProcedure
        LibrefCount     = $m.LibrefCount
        Librefs         = $m.Librefs
        ErrorText       = ($m.ErrorText -replace '[\r\n,]',' ')
    }
}

if ($TestMode) {
    $rows | Select-Object -First 20 | Format-Table Environment, ProgramName, CpuTimeSec, RealTimeSec, ErrorCount -AutoSize
    Write-RunLog "TestMode: CSV not written."
    return
}

try {
    $dir = Split-Path $out -Parent
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $rows | Export-Csv -Path $out -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
    Write-RunLog "Wrote $(@($rows).Count) rows to $out"
} catch {
    Write-RunLog "CSV write failed: $($_.Exception.Message)" -Level "ERROR"
    exit 2
}

$bad = @($rows | Where-Object { $_.ParseStatus -eq "PARSE_ERROR" }).Count
if ($bad -gt 0) { Write-RunLog "$bad files could not be parsed." -Level "WARN"; exit 1 }
exit 0

#endregion MAIN
