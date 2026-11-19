<#
.SYNOPSIS
  Script Name   : Check-SASServers_v2.0.ps1
  Purpose       : Poll the SAS HTTP endpoints in UNIT and PROD and write one
                  CSV row per sweep for the DMW Home tab.
  Version       : 2.0.0
  Derived from  : SAS_Monitor_Lite_v2.ps1 (URL monitoring, logging, recovery)

  WHAT CHANGED FROM SAS_Monitor_Lite_v2
    The Lite monitor polls ONE url in a long-running loop with email alerting.
    DMW needs the opposite shape: poll SEVERAL endpoints across TWO environments,
    once, quickly, and leave behind a flat CSV the dashboard can read. So this
    script keeps the logging, the recovery/outage tracking and the CSV append,
    and drops the polling loop, the SAS performance tests and the SMTP stack.

    Run it from Task Scheduler every N minutes instead of leaving it resident.

  .DESCRIPTION
    Checks each endpoint in $Targets, records status, HTTP code and response
    time, and appends one row per endpoint per sweep to a monthly CSV.
    Writes a rolling text log beside it.

    OUTPUT -> data\sas_server_status.csv   (read by the DMW Home tab)

  .PARAMETER CheckOnce
    Run one sweep and exit. This is the default and the only supported mode;
    the switch exists for symmetry with the Lite monitor.

  .PARAMETER TestMode
    Print results to the console without writing the CSV.

  EXECUTION POLICY (UNC paths)
    PowerShell blocks unsigned scripts from network shares. Either:
      Unblock-File -Path "\\server\share\Check-SASServers_v2.0.ps1"
    or invoke as:
      powershell.exe -ExecutionPolicy Bypass -File "<path>"

  TASK SCHEDULER
    Program : powershell.exe
    Args    : -NonInteractive -ExecutionPolicy Bypass -File "<path>\Check-SASServers_v2.0.ps1"
    Trigger : Every 5 minutes
    Run As  : an account with write access to $DataDir and $LogDir
#>

#Requires -Version 5.1
[CmdletBinding()]
param(
    [switch]$CheckOnce,
    [switch]$TestMode
)

###############################################################################
#region CONFIGURATION  -- the only section an operator should need to edit
###############################################################################

# --- WHAT TO CHECK -------------------------------------------------------
# Add or remove rows freely. Name and Environment appear verbatim on the
# DMW Home tab, so keep them short and stable.
$Targets = @(
    [PSCustomObject]@{ Name = "SAS WIP";       Environment = "UNIT"; Url = "https://cgssaswebu.a70admed.com/SASWIP/" }
    [PSCustomObject]@{ Name = "SAS Scheduler"; Environment = "UNIT"; Url = "https://cgssaswebu.a70admed.com/SASStoredProcess/" }
    [PSCustomObject]@{ Name = "SAS WIP";       Environment = "PROD"; Url = "https://cgssasweb.a70admed.com/SASWIP/" }
    [PSCustomObject]@{ Name = "SAS Scheduler"; Environment = "PROD"; Url = "https://cgssasweb.a70admed.com/SASStoredProcess/" }
)

# --- TIMING --------------------------------------------------------------
$TimeoutSec = 30     # max seconds to wait for each HTTP response
$RetryCount = 1      # extra attempts before declaring DOWN (0 = single attempt)
$RetryWaitSec = 3

# --- OUTPUT --------------------------------------------------------------
# The dashboard reads $DataDir\$DataFileName. Keep it beside the HTML page,
# or point the page's CONFIG.SERVER_STATUS_CSV at wherever you put it.
$DataDir      = "$PSScriptRoot\..\data"
$DataFileName = "sas_server_status.csv"
$KeepDays     = 90   # rows older than this are pruned on each sweep (0 = keep all)

$LogDir           = "$PSScriptRoot\..\log"
$LogRetentionDays = 90
$LogLevel         = "INFO"   # DEBUG | INFO | WARN | ERROR

# --- STATUS THRESHOLDS ---------------------------------------------------
# A reachable endpoint that is slow is reported as DEGRADED rather than UP,
# so the dashboard can show amber before anything actually breaks.
$DegradedMs = 3000

#endregion CONFIGURATION

###############################################################################
#region LOGGING
###############################################################################

$script:LogLevelMap = @{ DEBUG = 0; INFO = 1; WARN = 2; ERROR = 3 }

function Write-RunLog {
    param([string]$Message, [string]$Level = "INFO")
    try {
        if ($script:LogLevelMap[$Level] -lt $script:LogLevelMap[$LogLevel]) { return }
        if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
        $ts   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $file = Join-Path $LogDir ("Check-SASServers_{0}.log" -f (Get-Date -Format "yyyyMMdd"))
        "[$ts] [$Level] $Message" | Out-File -FilePath $file -Append -Encoding UTF8
        if ($Level -eq "ERROR") { Write-Host "[$ts] [$Level] $Message" -ForegroundColor Red }
        elseif ($Level -eq "WARN") { Write-Host "[$ts] [$Level] $Message" -ForegroundColor Yellow }
        else { Write-Verbose "[$ts] [$Level] $Message" }
    } catch {
        # Logging must never break the sweep.
        Write-Host "Logging failed: $($_.Exception.Message)" -ForegroundColor DarkYellow
    }
}

function Remove-OldLogs {
    if ($LogRetentionDays -le 0 -or -not (Test-Path $LogDir)) { return }
    try {
        Get-ChildItem -Path $LogDir -Filter "Check-SASServers_*.log" -ErrorAction Stop |
            Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$LogRetentionDays) } |
            Remove-Item -Force -ErrorAction SilentlyContinue
    } catch { Write-RunLog "Log purge failed: $($_.Exception.Message)" -Level "WARN" }
}

#endregion LOGGING

###############################################################################
#region CHECK
###############################################################################

function Test-Endpoint {
    <#
      Returns one PSCustomObject per endpoint. Status is one of:
        UP        reachable, HTTP 2xx/3xx, inside the latency threshold
        DEGRADED  reachable but slower than $DegradedMs
        DOWN      unreachable, timed out, or returned 4xx/5xx
      A SAS logon redirect counts as UP: the server answered, which is what
      this script is testing. Authentication is a separate concern.
    #>
    param([PSCustomObject]$Target)

    $attempt = 0
    $maxAttempts = $RetryCount + 1
    $status = "DOWN"; $code = 0; $ms = 0; $err = ""

    while ($attempt -lt $maxAttempts) {
        $attempt++
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        try {
            $resp = Invoke-WebRequest -Uri $Target.Url `
                                      -TimeoutSec $TimeoutSec `
                                      -UseBasicParsing `
                                      -MaximumRedirection 5 `
                                      -ErrorAction Stop
            $sw.Stop()
            $ms   = [int]$sw.Elapsed.TotalMilliseconds
            $code = [int]$resp.StatusCode
            if ($code -ge 200 -and $code -lt 400) {
                $status = if ($ms -gt $DegradedMs) { "DEGRADED" } else { "UP" }
                $err = ""
                break
            } else {
                $status = "DOWN"; $err = "HTTP $code"
            }
        } catch {
            $sw.Stop()
            $ms = [int]$sw.Elapsed.TotalMilliseconds
            $status = "DOWN"
            $err = $_.Exception.Message
            # An HTTP error response still proves the server answered; record
            # the code so a 503 is distinguishable from a dead host.
            if ($_.Exception.Response) {
                try { $code = [int]$_.Exception.Response.StatusCode } catch { $code = 0 }
            }
        }
        if ($attempt -lt $maxAttempts) { Start-Sleep -Seconds $RetryWaitSec }
    }

    [PSCustomObject]@{
        Timestamp      = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
        Environment    = $Target.Environment
        ServiceName    = $Target.Name
        ServiceType    = "HTTP"
        Url            = $Target.Url
        Status         = $status
        HttpCode       = $code
        ResponseTimeMs = $ms
        Attempts       = $attempt
        ErrorText      = ($err -replace '[\r\n,]', ' ')
        CheckedBy      = $env:COMPUTERNAME
    }
}

#endregion CHECK

###############################################################################
#region MAIN
###############################################################################

Write-RunLog "Sweep started on $env:COMPUTERNAME ($($Targets.Count) endpoints)"
Remove-OldLogs

$results = foreach ($t in $Targets) {
    $r = Test-Endpoint -Target $t
    $lvl = switch ($r.Status) { "DOWN" { "ERROR" } "DEGRADED" { "WARN" } default { "INFO" } }
    Write-RunLog ("{0}/{1}: {2} (HTTP {3}, {4} ms)" -f $r.Environment, $r.ServiceName, $r.Status, $r.HttpCode, $r.ResponseTimeMs) -Level $lvl
    $r
}

if ($TestMode) {
    $results | Format-Table Environment, ServiceName, Status, HttpCode, ResponseTimeMs -AutoSize
    Write-RunLog "TestMode: CSV not written."
    return
}

# --- append to the monthly CSV ------------------------------------------
try {
    if (-not (Test-Path $DataDir)) { New-Item -ItemType Directory -Path $DataDir -Force | Out-Null }
    $csvPath = Join-Path $DataDir $DataFileName

    if (Test-Path $csvPath) {
        $results | Export-Csv -Path $csvPath -Append -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
    } else {
        $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
    }
    Write-RunLog "Wrote $($results.Count) rows to $csvPath"

    # Prune old rows so the dashboard's fetch stays small.
    if ($KeepDays -gt 0) {
        $cutoff = (Get-Date).AddDays(-$KeepDays)
        $all = Import-Csv -Path $csvPath
        $kept = $all | Where-Object {
            $d = $null
            [datetime]::TryParse($_.Timestamp, [ref]$d) -and $d -ge $cutoff
        }
        if ($kept.Count -lt $all.Count) {
            $kept | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            Write-RunLog "Pruned $($all.Count - $kept.Count) rows older than $KeepDays days"
        }
    }
} catch {
    Write-RunLog "CSV write failed: $($_.Exception.Message)" -Level "ERROR"
    exit 2
}

$down = @($results | Where-Object { $_.Status -eq "DOWN" }).Count
Write-RunLog "Sweep complete. DOWN=$down of $($results.Count)."
if ($down -gt 0) { exit 1 }
exit 0

#endregion MAIN
