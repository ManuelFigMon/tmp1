<#
.SYNOPSIS
  Script Name   : Export-ServerLatencyDaily.ps1
  Purpose       : Roll the per-check server status file up to ONE ROW PER
                  ENDPOINT PER DAY, with latency percentiles, for the DMW
                  Charts tab.
  Version       : 1.0.0

  -----------------------------------------------------------------------------
  WHY THIS EXISTS
  -----------------------------------------------------------------------------
  Check-SASServers_v2.0.ps1 appends one row per endpoint per sweep. At one
  sweep every five minutes that is 288 rows per endpoint per day, so a year of
  four endpoints is roughly 420,000 rows. The browser was downloading and
  aggregating all of it on every visit to the Charts tab.

  This script does that arithmetic once, on the server, and leaves 4 rows per
  day behind. The page reads those instead: about 1,500 rows for a year rather
  than 420,000.

  The raw file is still what you want for the scatter -- one dot per check is
  the point of that chart -- so this does not replace it. Keep both.

  -----------------------------------------------------------------------------
  PERCENTILE DEFINITION -- MUST MATCH THE PAGE
  -----------------------------------------------------------------------------
  Nearest-rank on the ascending sample:

      index = min(n - 1, floor(n * p))        0-based

  The DMW page uses exactly this when it computes percentiles from the raw
  file. Two different definitions would make the daily file and the raw file
  disagree by a millisecond or two on the same data, and someone would
  eventually spend an afternoon on it. If you change it here, change it there.

  -----------------------------------------------------------------------------
  WHICH CHECKS COUNT
  -----------------------------------------------------------------------------
  A DOWN check records roughly the timeout ceiling, not a measurement. Left in
  the percentiles it drags them toward the timeout value and hides real
  movement in healthy latency. So:

      Latency statistics  -- computed from NON-DOWN checks only.
      Availability        -- computed from ALL checks.
      Timeouts            -- counted separately, so nothing is hidden.

  A day on which every check failed therefore has Checks > 0, UpChecks = 0,
  AvailabilityPct = 0 and EMPTY percentile columns. Empty means "no healthy
  check to measure", which is not the same as zero milliseconds, and the page
  skips those points rather than plotting them on the floor.

  -----------------------------------------------------------------------------
  OUTPUT -> data\sas_server_latency_daily.csv
      Date, Environment, ServiceName, Checks, UpChecks, Timeouts, Degraded,
      AvailabilityPct, MinMs, P50Ms, P90Ms, P95Ms, P99Ms, MaxMs, MeanMs
  -----------------------------------------------------------------------------

  .PARAMETER InputPath
    The per-check CSV. Defaults to data\sas_server_status.csv.

  .PARAMETER OutputPath
    Defaults to data\sas_server_latency_daily.csv.

  .PARAMETER Days
    Only roll up the last N days. 0 = everything in the file.

  .PARAMETER TestMode
    Print the result, write nothing.

  .EXAMPLE
    .\Export-ServerLatencyDaily.ps1

  .EXAMPLE
    .\Export-ServerLatencyDaily.ps1 -Days 400 -TestMode

  SCHEDULING
    Run it once a day, after midnight, so the previous day is complete:

      Program : powershell.exe
      Args    : -NonInteractive -ExecutionPolicy Bypass -File "<path>\Export-ServerLatencyDaily.ps1"
      Trigger : Daily at 00:20

    It is idempotent -- it rewrites the whole output file from the input every
    run, so a missed day is repaired by the next one rather than leaving a
    hole. You can also add it as a second action on the Check-SASServers task,
    but once a day is enough: re-aggregating every five minutes spends work to
    produce the same answer.
#>

#Requires -Version 5.1
[CmdletBinding()]
param(
    [string] $InputPath  = "",
    [string] $OutputPath = "",
    [int]    $Days       = 0,
    [switch] $TestMode
)

###############################################################################
#region CONFIGURATION
###############################################################################

$DataDir        = "$PSScriptRoot\..\data"
$DefaultInput   = Join-Path $DataDir "sas_server_status.csv"
$DefaultOutput  = Join-Path $DataDir "sas_server_latency_daily.csv"

$LogDir   = "$PSScriptRoot\..\log"
$LogLevel = "INFO"

# Status values that mean "no measurement was taken".
$DownStatuses = @('DOWN')

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
        "[$ts] [$Level] $Message" | Out-File -FilePath (Join-Path $LogDir ("Export-ServerLatencyDaily_{0}.log" -f (Get-Date -Format "yyyyMMdd"))) -Append -Encoding UTF8
        if ($Level -eq "ERROR") { Write-Host "[$ts] [$Level] $Message" -ForegroundColor Red }
        elseif ($Level -eq "WARN") { Write-Host "[$ts] [$Level] $Message" -ForegroundColor Yellow }
    } catch { }
}
#endregion LOGGING

###############################################################################
#region STATS
###############################################################################

function Get-Percentile {
    <# .SYNOPSIS Nearest-rank percentile. See the header: this definition is
                 shared with the DMW page and must not drift from it.
       .PARAMETER Sorted  Ascending [double[]].
       .PARAMETER P       0..1 #>
    param([double[]] $Sorted, [double] $P)
    if (-not $Sorted -or $Sorted.Count -eq 0) { return $null }
    $i = [Math]::Floor($Sorted.Count * $P)
    if ($i -gt $Sorted.Count - 1) { $i = $Sorted.Count - 1 }
    if ($i -lt 0) { $i = 0 }
    return $Sorted[[int]$i]
}

#endregion STATS

###############################################################################
#region MAIN
###############################################################################

$in  = if ($InputPath)  { $InputPath }  else { $DefaultInput }
$out = if ($OutputPath) { $OutputPath } else { $DefaultOutput }

Write-RunLog "start; input=$in"

if (-not (Test-Path $in)) {
    Write-RunLog "input not found: $in" -Level "ERROR"
    Write-Host "Run Check-SASServers_v2.0.ps1 first, or pass -InputPath." -ForegroundColor Yellow
    exit 3
}

try { $rows = @(Import-Csv -Path $in -ErrorAction Stop) }
catch {
    Write-RunLog "could not read ${in}: $($_.Exception.Message)" -Level "ERROR"
    exit 3
}
if ($rows.Count -eq 0) { Write-RunLog "input is empty; nothing written" -Level "WARN"; exit 0 }

$cutoff = if ($Days -gt 0) { (Get-Date).AddDays(-$Days) } else { [datetime]::MinValue }

# Bucket by Date + Environment + ServiceName in one pass.
$buckets = @{}
$skipped = 0
foreach ($r in $rows) {
    $ts = $null
    if (-not [datetime]::TryParse($r.Timestamp, [ref]$ts)) { $skipped++; continue }
    if ($ts -lt $cutoff) { continue }

    $key = '{0:yyyy-MM-dd}|{1}|{2}' -f $ts, $r.Environment, $r.ServiceName
    if (-not $buckets.ContainsKey($key)) {
        $buckets[$key] = [PSCustomObject]@{
            Date = $ts.ToString('yyyy-MM-dd'); Environment = $r.Environment
            ServiceName = $r.ServiceName; Checks = 0; Timeouts = 0; Degraded = 0
            Latencies = (New-Object System.Collections.Generic.List[double])
        }
    }
    $b = $buckets[$key]
    $b.Checks++

    if ($DownStatuses -contains $r.Status) {
        $b.Timeouts++
        continue                      # no measurement to record
    }
    if ($r.Status -eq 'DEGRADED') { $b.Degraded++ }

    $ms = 0.0
    if ([double]::TryParse($r.ResponseTimeMs, [ref]$ms)) { $b.Latencies.Add($ms) }
}
if ($skipped -gt 0) { Write-RunLog "$skipped row(s) had an unparseable Timestamp and were ignored" -Level "WARN" }

$outRows = foreach ($key in ($buckets.Keys | Sort-Object)) {
    $b = $buckets[$key]
    $sorted = [double[]]($b.Latencies | Sort-Object)
    $n = $sorted.Count

    [PSCustomObject][ordered]@{
        Date            = $b.Date
        Environment     = $b.Environment
        ServiceName     = $b.ServiceName
        Checks          = $b.Checks
        UpChecks        = $n
        Timeouts        = $b.Timeouts
        Degraded        = $b.Degraded
        AvailabilityPct = if ($b.Checks -gt 0) { [Math]::Round((($b.Checks - $b.Timeouts) / $b.Checks) * 100, 3) } else { '' }
        # Empty, not zero, when no healthy check was taken that day.
        MinMs  = if ($n) { [Math]::Round($sorted[0], 1) } else { '' }
        P50Ms  = if ($n) { [Math]::Round((Get-Percentile $sorted 0.50), 1) } else { '' }
        P90Ms  = if ($n) { [Math]::Round((Get-Percentile $sorted 0.90), 1) } else { '' }
        P95Ms  = if ($n) { [Math]::Round((Get-Percentile $sorted 0.95), 1) } else { '' }
        P99Ms  = if ($n) { [Math]::Round((Get-Percentile $sorted 0.99), 1) } else { '' }
        MaxMs  = if ($n) { [Math]::Round($sorted[$n - 1], 1) } else { '' }
        MeanMs = if ($n) { [Math]::Round((($sorted | Measure-Object -Average).Average), 1) } else { '' }
    }
}

$outRows = @($outRows)
Write-RunLog ("rolled {0} check row(s) up to {1} daily row(s)" -f $rows.Count, $outRows.Count)

if ($TestMode) {
    $outRows | Select-Object -Last 12 | Format-Table Date, Environment, ServiceName, Checks, Timeouts, P50Ms, P95Ms, MaxMs -AutoSize
    Write-RunLog "TestMode: CSV not written."
    return
}

try {
    $dir = Split-Path $out -Parent
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    # Rewritten whole every run, so a missed day repairs itself next time.
    $outRows | Export-Csv -Path $out -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
    Write-RunLog "wrote $out"
} catch {
    Write-RunLog "CSV write failed: $($_.Exception.Message)" -Level "ERROR"
    exit 3
}

exit 0

#endregion MAIN
