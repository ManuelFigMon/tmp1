<#
.SYNOPSIS
  Script Name   : SAS_Monitor_Lite_v1.ps1
  Author        : Manuel Figallo
  Purpose       : Continuous URL monitoring, structured logging, and
                  alert emails for the SAS Stored Process Server
  Version       : 1.0.0-Lite
  Created       : <date placeholder>
  Last Modified : <date placeholder>

  Description:
    Polls the SAS Stored Process web endpoint on a configurable interval.
    All script activity is written to timestamped log files in $LogDir.
    Check records are appended to monthly CSVs in $DataDir.
    Sends daily summary, Monday-morning weekly digest, and immediate
    URGENT / recovery alert emails. No automated service restart is
    performed — this script is monitoring and reporting only.
    All paths, addresses, and behavioral flags are defined in the
    CONFIGURATION region.

  Task Scheduler setup:
    Program : powershell.exe
    Args    : -NonInteractive -ExecutionPolicy Bypass -File "C:\scripts\SAS_Monitor_Lite_v1.ps1"
    Trigger : At system startup; also repeat daily (for daily report timing)
    Run As  : Any domain account with read/write access to $DataDir and $LogDir

  Parameters:
    -TestMode   : One check, one test email to $UrgentGroup, then exit
    -CheckOnce  : One check, write CSV/log, print to console, then exit

  Change Log:
    v1.0.0-Lite — Lite edition: monitoring and reporting only; no service restart
#>

#Requires -Version 5.1
[CmdletBinding()]
param(
    [switch]$TestMode,
    [switch]$CheckOnce
)

###############################################################################
#region CONFIGURATION
# This is the ONLY section an operator should need to edit.
###############################################################################

# --- MONITORING ---
$MonitorURL         = "https://cgssaswebu.a70admed.com/SASStoredProcess/"  # URL to poll
$PollingIntervalSec = 60    # seconds between each URL check
$TimeoutSec         = 30    # max seconds to wait for HTTP response
$FailureThreshold   = 3     # consecutive DOWN checks before triggering URGENT alert

# --- DATA STORAGE (CSV records — one file per month, append-only) ---
$DataDir         = "\\a70tucgssasr006\custom\projects\sas_monitor\data"  # UNC path for CSV files
$DataFilePrefix  = "sas_monitor"    # filename stem; e.g. sas_monitor_2026-05.csv
$ReportSharePath = "\\a70tucgssasr006\custom\projects\sas_monitor\data"  # share shown in email footers

# --- LOG STORAGE (text log files — separate from CSV data) ---
$LogDir           = "\\a70tucgssasr006\custom\projects\sas_monitor\log"  # UNC path for run/error logs
$LogRetentionDays = 90       # auto-purge log files older than N days
$LogLevel         = "INFO"   # minimum level to write: DEBUG | INFO | WARN | ERROR
$EnableEmailAuditLog = $true # append every email send attempt to a monthly audit log

# --- EMAIL / SMTP ---
$SMTPServer     = "smtp.yourdomain.com"         # SMTP relay hostname or IP
$SMTPPort       = 25                            # SMTP port (25 = relay, 587 = submission)
$FromAddress    = "sas-monitor@yourdomain.com"  # envelope From address
$UseSSL         = $false                        # $true for STARTTLS / implicit TLS
$SMTPCredential = $null                         # set to (Get-Credential) if SMTP auth is required

# --- DISTRIBUTION LISTS ---
$DailyGroup  = @("user1@domain.com", "user2@domain.com")        # daily summary recipients
$WeeklyGroup = @("manager1@domain.com", "manager2@domain.com")  # weekly digest recipients
$UrgentGroup = @("oncall1@domain.com", "oncall2@domain.com")    # outage/recovery alerts

# --- REPORT SCHEDULE ---
$DailyReportHour  = 7           # 24-hour clock hour to send the daily report
$WeeklyReportDay  = "Monday"    # day of week for weekly digest
$WeeklyReportHour = 7           # 24-hour clock hour to send the weekly digest

#endregion CONFIGURATION

###############################################################################
#region LOGGING SUBSYSTEM
###############################################################################

# Maps level strings to numeric priority for threshold comparisons
$script:LogLevelMap = @{ DEBUG = 0; INFO = 1; WARN = 2; ERROR = 3 }

# Tracks whether the daily log-purge has already run today
$script:LastPurgeDate = [datetime]::MinValue

function Write-RunLog {
    <#
    .SYNOPSIS Writes a timestamped entry to the daily run log and console.
    #>
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    try {
        # Honour the configured minimum log level
        $configuredPriority = $script:LogLevelMap[$LogLevel]
        $entryPriority      = $script:LogLevelMap[$Level]
        if ($entryPriority -lt $configuredPriority) { return }

        $ts   = (Get-Date).ToString("o")
        $line = "[$ts] [$Level] $Message"

        # Write to console so operators watching the window see live activity
        Write-Host $line

        # Emit to the -Verbose stream for operators who run with -Verbose
        Write-Verbose $line

        # Build today's log file path and ensure the directory exists
        $logFile = "$LogDir\sas_monitor_run_$(Get-Date -Format 'yyyy-MM-dd').log"
        $line | Out-File -FilePath $logFile -Append -Encoding UTF8

        # Once per day, purge log files older than $LogRetentionDays
        $today = (Get-Date).Date
        if ($script:LastPurgeDate -ne $today) {
            $script:LastPurgeDate = $today
            $cutoff = $today.AddDays(-$LogRetentionDays)
            Get-ChildItem -Path $LogDir -Filter "*.log" -ErrorAction SilentlyContinue |
                Where-Object { $_.LastWriteTime -lt $cutoff } |
                ForEach-Object {
                    try {
                        Remove-Item $_.FullName -Force -ErrorAction Stop
                        $purgeTs = (Get-Date).ToString("o")
                        "[$purgeTs] [DEBUG] Purged old log: $($_.FullName)" |
                            Out-File -FilePath $logFile -Append -Encoding UTF8
                    } catch { <# swallow purge errors — never crash #> }
                }
        }
    } catch {
        Write-Warning "Write-RunLog failed: $_"
    }
}

function Write-ErrorLog {
    <#
    .SYNOPSIS Writes full exception detail to the daily error log, then calls Write-RunLog.
    #>
    param(
        [string]$Message,
        [System.Exception]$Exception = $null
    )
    try {
        $ts      = (Get-Date).ToString("o")
        $errFile = "$LogDir\sas_monitor_errors_$(Get-Date -Format 'yyyy-MM-dd').log"

        "[$ts] [ERROR] $Message" | Out-File -FilePath $errFile -Append -Encoding UTF8

        if ($null -ne $Exception) {
            "[$ts] [ERROR]   Exception Type : $($Exception.GetType().FullName)" |
                Out-File -FilePath $errFile -Append -Encoding UTF8
            "[$ts] [ERROR]   Message        : $($Exception.Message)" |
                Out-File -FilePath $errFile -Append -Encoding UTF8

            # Limit stack trace to first 3 lines to keep logs readable
            $stackLines = ($Exception.StackTrace -split "`n") | Select-Object -First 3
            foreach ($sl in $stackLines) {
                "[$ts] [ERROR]   StackTrace     : $($sl.Trim())" |
                    Out-File -FilePath $errFile -Append -Encoding UTF8
            }
        }

        # Mirror to run log so a single tail shows all activity
        Write-RunLog -Message $Message -Level "ERROR"
    } catch {
        Write-Warning "Write-ErrorLog itself failed: $_"
    }
}

function Write-EmailAuditLog {
    <#
    .SYNOPSIS Appends one record per email send attempt to the monthly audit log.
    #>
    param(
        [string]$Group,
        [string]$Subject,
        [string]$Result
    )
    if (-not $EnableEmailAuditLog) { return }
    try {
        $ts        = (Get-Date).ToString("o")
        $auditFile = "$LogDir\sas_monitor_email_$(Get-Date -Format 'yyyy-MM').log"
        "[$ts] | TO: $Group | SUBJECT: $Subject | RESULT: $Result" |
            Out-File -FilePath $auditFile -Append -Encoding UTF8
    } catch {
        Write-Warning "Write-EmailAuditLog failed: $_"
    }
}

#endregion LOGGING SUBSYSTEM

###############################################################################
#region CORE HELPER FUNCTIONS
###############################################################################

function Get-MonthlyCSVPath {
    <#
    .SYNOPSIS Returns the path of the current month's append-only CSV and ensures $DataDir exists.
    #>
    if (-not (Test-Path $DataDir)) {
        New-Item -ItemType Directory -Path $DataDir -Force | Out-Null
        Write-RunLog "Created DataDir: $DataDir"
    }
    return "$DataDir\$($DataFilePrefix)_$(Get-Date -Format 'yyyy-MM').csv"
}

function Get-EnsureDir {
    <#
    .SYNOPSIS Creates a directory if it does not already exist, logging the creation.
    #>
    param([string]$Path)
    try {
        if (-not (Test-Path $Path)) {
            New-Item -ItemType Directory -Path $Path -Force | Out-Null
            Write-RunLog "Created directory: $Path"
        }
    } catch {
        Write-ErrorLog "Failed to create directory: $Path" -Exception $_.Exception
    }
}

function Get-UptimePct24h {
    <#
    .SYNOPSIS Calculates rolling 24-hour uptime percentage from the monthly CSV.
    #>
    try {
        $csvPath = Get-MonthlyCSVPath
        if (-not (Test-Path $csvPath)) { return 100.0 }
        $cutoff  = (Get-Date).AddHours(-24)
        $rows    = Import-Csv $csvPath |
            Where-Object { [datetime]$_.Timestamp_ISO8601 -gt $cutoff }
        if ($rows.Count -eq 0) { return 100.0 }
        $upCount = ($rows | Where-Object { $_.ServerStatus -eq "UP" }).Count
        return [math]::Round(($upCount / $rows.Count) * 100, 2)
    } catch {
        return -1
    }
}

function Invoke-URLCheck {
    <#
    .SYNOPSIS Performs a single HTTP check against $MonitorURL and returns a structured result object.
    #>
    $ts        = (Get-Date).ToString("o")
    $dateLoc   = (Get-Date).ToString("MM/dd/yyyy")
    $timeLoc   = (Get-Date).ToString("HH:mm:ss.fff")
    $tz        = [System.TimeZoneInfo]::Local.DisplayName
    $utcOffset = [System.TimeZoneInfo]::Local.GetUtcOffset((Get-Date)).ToString("hh\:mm")
    if ([System.TimeZoneInfo]::Local.GetUtcOffset((Get-Date)).TotalHours -lt 0) {
        $utcOffset = "-$utcOffset"
    } else {
        $utcOffset = "+$utcOffset"
    }

    # Default values — overwritten on success or classified on failure
    $serverStatus   = "DOWN"
    $detailedStatus = "Unknown"
    $serverError    = "Unknown"
    $httpCode       = 0
    $responseMs     = 0
    $contentLen     = -1
    $redirectCount  = 0
    $finalURL       = $MonitorURL
    $sslExpiry      = "N/A"
    $sslDays        = -1
    $dnsMs          = 0

    # --- 1. DNS resolution timing ---
    try {
        $dnsSw = [System.Diagnostics.Stopwatch]::StartNew()
        $uri   = [System.Uri]$MonitorURL
        [System.Net.Dns]::GetHostEntry($uri.Host) | Out-Null
        $dnsSw.Stop()
        $dnsMs = $dnsSw.ElapsedMilliseconds
        Write-Verbose "DNS resolved '$($uri.Host)' in ${dnsMs}ms"
    } catch {
        $serverError    = "DNS"
        $detailedStatus = "DNS Resolution Failed"
        Write-ErrorLog "DNS resolution failed for $MonitorURL" -Exception $_.Exception
        # Return early — no point attempting HTTP if DNS is broken
        return [PSCustomObject]@{
            Timestamp_ISO8601   = $ts
            Date_Local          = $dateLoc
            Time_Local          = $timeLoc
            TimeZone            = $tz
            UTCOffset           = $utcOffset
            ServerStatus        = $serverStatus
            DetailedStatus      = $detailedStatus
            ServerError         = $serverError
            HTTPStatusCode      = $httpCode
            ResponseTime_ms     = $responseMs
            ContentLength_bytes = $contentLen
            RedirectCount       = $redirectCount
            FinalURL            = $finalURL
            SSLCertExpiryDate   = $sslExpiry
            SSLDaysUntilExpiry  = $sslDays
            DNSResolution_ms    = $dnsMs
            CheckedFromHost     = $env:COMPUTERNAME
            CheckedByUser       = $env:USERNAME
            ConsecutiveFailures = $script:ConsecutiveFailures
            UptimePct_24h       = (Get-UptimePct24h)
            OutageEventID       = $script:OutageEventID
        }
    }

    # --- 2. HTTP request with stopwatch ---
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $resp = Invoke-WebRequest -Uri $MonitorURL `
            -TimeoutSec $TimeoutSec `
            -UseBasicParsing `
            -MaximumRedirection 10 `
            -ErrorAction Stop

        $sw.Stop()
        $responseMs     = $sw.ElapsedMilliseconds
        $httpCode       = [int]$resp.StatusCode
        $detailedStatus = "HTTP $httpCode $($resp.StatusDescription)"
        $contentLen     = if ($resp.Headers["Content-Length"]) { [long]$resp.Headers["Content-Length"] } else { -1 }
        $finalURL       = $resp.BaseResponse.ResponseUri.AbsoluteUri
        # Count redirects as difference between original and final URI hops (approximate)
        $redirectCount  = ($resp.BaseResponse.ResponseUri.AbsoluteUri -ne $MonitorURL) ? 1 : 0

        if ($httpCode -eq 200) {
            $serverStatus = "UP"
            $serverError  = "None"
        } elseif ($httpCode -ge 400 -and $httpCode -lt 500) {
            $serverError  = "HTTP_4xx"
        } elseif ($httpCode -ge 500) {
            $serverError  = "HTTP_5xx"
        }
        Write-Verbose "HTTP $httpCode from $MonitorURL in ${responseMs}ms"

    } catch [System.Net.WebException] {
        $sw.Stop()
        $responseMs = $sw.ElapsedMilliseconds
        $ex = $_.Exception

        # Classify the failure type from exception characteristics
        if ($ex.Status -eq [System.Net.WebExceptionStatus]::Timeout) {
            $serverError    = "Timeout"
            $detailedStatus = "Connection Timed Out"
        } elseif ($ex.Status -eq [System.Net.WebExceptionStatus]::NameResolutionFailure) {
            $serverError    = "DNS"
            $detailedStatus = "DNS Resolution Failed"
        } elseif ($ex.Status -eq [System.Net.WebExceptionStatus]::ConnectFailure) {
            $serverError    = "ConnectionRefused"
            $detailedStatus = "Connection Refused"
        } elseif ($ex.Message -match "SSL|TLS|certificate") {
            $serverError    = "SSL"
            $detailedStatus = "SSL/TLS Error"
        } elseif ($null -ne $ex.Response) {
            $httpCode       = [int]$ex.Response.StatusCode
            $detailedStatus = "HTTP $httpCode"
            $serverError    = if ($httpCode -ge 500) { "HTTP_5xx" } elseif ($httpCode -ge 400) { "HTTP_4xx" } else { "Unknown" }
        } else {
            $serverError    = "Unknown"
            $detailedStatus = $ex.Message
        }
        Write-ErrorLog "URL check failed: $detailedStatus" -Exception $ex
    } catch {
        $sw.Stop()
        $responseMs     = $sw.ElapsedMilliseconds
        $serverError    = "Unknown"
        $detailedStatus = $_.Exception.Message
        Write-ErrorLog "URL check unexpected error" -Exception $_.Exception
    }

    # --- 3. SSL certificate expiry via raw TLS socket ---
    try {
        $uri = [System.Uri]$MonitorURL
        if ($uri.Scheme -eq "https") {
            $tcpClient = [System.Net.Sockets.TcpClient]::new($uri.Host, $uri.Port)
            $sslStream = [System.Net.Security.SslStream]::new(
                $tcpClient.GetStream(), $false,
                { $true }   # accept-all callback — we just want the cert
            )
            $sslStream.AuthenticateAsClient($uri.Host)
            $cert      = $sslStream.RemoteCertificate
            $sslExpiry = ([datetime]::Parse($cert.GetExpirationDateString())).ToString("o")
            $sslDays   = ([datetime]::Parse($cert.GetExpirationDateString()) - (Get-Date)).Days
            $sslStream.Dispose()
            $tcpClient.Dispose()
        }
    } catch {
        # Non-fatal — SSL check failure does not change UP/DOWN status
        Write-Verbose "SSL cert check failed: $($_.Exception.Message)"
    }

    return [PSCustomObject]@{
        Timestamp_ISO8601   = $ts
        Date_Local          = $dateLoc
        Time_Local          = $timeLoc
        TimeZone            = $tz
        UTCOffset           = $utcOffset
        ServerStatus        = $serverStatus
        DetailedStatus      = $detailedStatus
        ServerError         = $serverError
        HTTPStatusCode      = $httpCode
        ResponseTime_ms     = $responseMs
        ContentLength_bytes = $contentLen
        RedirectCount       = $redirectCount
        FinalURL            = $finalURL
        SSLCertExpiryDate   = $sslExpiry
        SSLDaysUntilExpiry  = $sslDays
        DNSResolution_ms    = $dnsMs
        CheckedFromHost     = $env:COMPUTERNAME
        CheckedByUser       = $env:USERNAME
        ConsecutiveFailures = $script:ConsecutiveFailures
        UptimePct_24h       = (Get-UptimePct24h)
        OutageEventID       = $script:OutageEventID
    }
}

#endregion CORE HELPER FUNCTIONS

###############################################################################
#region EMAIL FUNCTIONS
###############################################################################

function Get-EmailFooter {
    <#
    .SYNOPSIS Returns the standard HTML footer appended to every outgoing email.
    #>
    $month = Get-Date -Format 'yyyy-MM'
    return @"
<hr style="border:none;border-top:1px solid #ddd;margin:20px 0">
<table style="font-size:12px;color:#555;width:100%">
  <tr>
    <td><strong>Monitoring data:</strong></td>
    <td><code>$DataDir\$($DataFilePrefix)_$month.csv</code></td>
  </tr>
  <tr>
    <td><strong>Log files:</strong></td>
    <td><code>$LogDir\</code></td>
  </tr>
</table>
"@
}

function Send-MonitorEmail {
    <#
    .SYNOPSIS Sends an HTML email via System.Net.Mail and audits the result.
    #>
    param(
        [string[]]$To,
        [string]$Subject,
        [string]$HTMLBody
    )
    $toStr = $To -join ";"
    try {
        $mail            = [System.Net.Mail.MailMessage]::new()
        $mail.From       = $FromAddress
        $mail.Subject    = $Subject
        $mail.Body       = $HTMLBody + (Get-EmailFooter)
        $mail.IsBodyHtml = $true
        foreach ($addr in $To) { $mail.To.Add($addr) }

        $smtp            = [System.Net.Mail.SmtpClient]::new($SMTPServer, $SMTPPort)
        $smtp.EnableSsl  = $UseSSL
        if ($null -ne $SMTPCredential) {
            $smtp.Credentials = $SMTPCredential.GetNetworkCredential()
        }
        $smtp.Send($mail)
        $smtp.Dispose()
        $mail.Dispose()

        Write-RunLog "Email sent: '$Subject' → $toStr"
        Write-EmailAuditLog -Group $toStr -Subject $Subject -Result "OK"
    } catch {
        Write-ErrorLog "Email send failed: '$Subject' → $toStr" -Exception $_.Exception
        Write-EmailAuditLog -Group $toStr -Subject $Subject -Result "FAILED: $($_.Exception.Message)"
    }
}

function Send-UrgentOutageEmail {
    param($Check)
    $subj = "*** URGENT *** SAS Server DOWN — $(Get-Date -Format 'o')"
    $body = @"
<div style="background:#c0392b;color:#fff;padding:16px;font-size:20px;font-weight:bold;text-align:center">
  &#9888; SERVER IS DOWN
</div>
<table style="width:100%;margin-top:16px;font-size:14px;border-collapse:collapse">
  <tr><td style="padding:4px;font-weight:bold">URL</td><td>$MonitorURL</td></tr>
  <tr><td style="padding:4px;font-weight:bold">First Failure</td><td>$($script:OutageStartTime.ToString('o'))</td></tr>
  <tr><td style="padding:4px;font-weight:bold">Timezone</td><td>$($Check.TimeZone)</td></tr>
  <tr><td style="padding:4px;font-weight:bold">Consecutive Failures</td><td>$($script:ConsecutiveFailures)</td></tr>
  <tr><td style="padding:4px;font-weight:bold">Error Type</td><td>$($Check.ServerError)</td></tr>
  <tr><td style="padding:4px;font-weight:bold">Detailed Status</td><td>$($Check.DetailedStatus)</td></tr>
  <tr><td style="padding:4px;font-weight:bold">Outage Event ID</td><td>$($script:OutageEventID)</td></tr>
  <tr><td style="padding:4px;font-weight:bold">Action Required</td><td>Manual intervention required — no automated restart in Lite mode</td></tr>
  <tr><td style="padding:4px;font-weight:bold">Escalation</td><td>[ESCALATION CONTACT PLACEHOLDER]</td></tr>
</table>
"@
    Send-MonitorEmail -To $UrgentGroup -Subject $subj -HTMLBody $body
}

function Send-RecoveryEmail {
    param($Check, [timespan]$Duration)
    $subj = "[RESOLVED] SAS Server RESTORED — $(Get-Date -Format 'o')"
    $body = @"
<div style="background:#27ae60;color:#fff;padding:16px;font-size:20px;font-weight:bold;text-align:center">
  &#10003; SERVER RESTORED
</div>
<table style="width:100%;margin-top:16px;font-size:14px;border-collapse:collapse">
  <tr><td style="padding:4px;font-weight:bold">Outage Event ID</td><td>$($script:OutageEventID)</td></tr>
  <tr><td style="padding:4px;font-weight:bold">Outage Start</td><td>$($script:OutageStartTime.ToString('o'))</td></tr>
  <tr><td style="padding:4px;font-weight:bold">Outage End</td><td>$(Get-Date -Format 'o')</td></tr>
  <tr><td style="padding:4px;font-weight:bold">Duration</td><td>$($Duration.ToString('hh\:mm\:ss'))</td></tr>
  <tr><td style="padding:4px;font-weight:bold">First Successful HTTP Status</td><td>$($Check.DetailedStatus)</td></tr>
  <tr><td style="padding:4px;font-weight:bold">Response Time</td><td>$($Check.ResponseTime_ms)ms</td></tr>
</table>
"@
    Send-MonitorEmail -To $UrgentGroup -Subject $subj -HTMLBody $body
}

function Send-DailyReport {
    $subj  = "[SAS Monitor] Daily Status Report — $(Get-Date -Format 'MM/dd/yyyy')"
    $today = (Get-Date).Date
    try {
        $csvPath = Get-MonthlyCSVPath
        $rows    = @()
        if (Test-Path $csvPath) {
            $rows = Import-Csv $csvPath |
                Where-Object { ([datetime]$_.Timestamp_ISO8601).Date -eq $today }
        }
        $totalChecks = $rows.Count
        $upChecks    = ($rows | Where-Object { $_.ServerStatus -eq "UP" }).Count
        $downRows    = $rows | Where-Object { $_.ServerStatus -eq "DOWN" }
        $uptimePct   = if ($totalChecks -gt 0) { [math]::Round(($upChecks / $totalChecks) * 100, 2) } else { 100 }
        $downtimeMin = [math]::Round(($totalChecks - $upChecks) * ($PollingIntervalSec / 60), 1)
        $incidents   = ($rows | Where-Object { $_.OutageEventID -ne "" } |
                        Select-Object -ExpandProperty OutageEventID -Unique).Count

        $upTimes = $rows | Where-Object { $_.ServerStatus -eq "UP" } |
                   ForEach-Object { [double]$_.ResponseTime_ms }
        $avgMs = if ($upTimes.Count -gt 0) { [math]::Round(($upTimes | Measure-Object -Average).Average, 1) } else { 0 }
        $minMs = if ($upTimes.Count -gt 0) { ($upTimes | Measure-Object -Minimum).Minimum } else { 0 }
        $maxMs = if ($upTimes.Count -gt 0) { ($upTimes | Measure-Object -Maximum).Maximum } else { 0 }

        $downTableRows = ""
        foreach ($r in $downRows) {
            $downTableRows += "<tr><td>$($r.Timestamp_ISO8601)</td><td>$($r.HTTPStatusCode)</td>" +
                              "<td>$($r.ResponseTime_ms)</td><td>$($r.ServerError)</td>" +
                              "<td>$($r.DetailedStatus)</td></tr>"
        }
        if ($downTableRows -eq "") {
            $downTableRows = "<tr><td colspan='5' style='text-align:center'>No DOWN events today</td></tr>"
        }

        $sslWarning = ""
        if ($rows.Count -gt 0) {
            $lastSSL = $rows | Select-Object -Last 1
            if ([int]$lastSSL.SSLDaysUntilExpiry -lt 30 -and [int]$lastSSL.SSLDaysUntilExpiry -ge 0) {
                $sslWarning = "<p style='color:#e67e22'><strong>&#9888; SSL Certificate expiring in $($lastSSL.SSLDaysUntilExpiry) days ($($lastSSL.SSLCertExpiryDate))</strong></p>"
            }
        }

        $body = @"
<h2>SAS Monitor — Daily Status Report</h2>
<p><strong>URL:</strong> $MonitorURL<br>
<strong>Generated:</strong> $(Get-Date -Format 'o') ($([System.TimeZoneInfo]::Local.DisplayName))</p>
<table border="1" style="border-collapse:collapse;font-size:14px">
  <tr style="background:#f0f0f0">
    <th>Total Checks</th><th>Uptime %</th><th>Downtime (min)</th><th>Incidents</th>
  </tr>
  <tr>
    <td style="text-align:center">$totalChecks</td>
    <td style="text-align:center">$uptimePct%</td>
    <td style="text-align:center">$downtimeMin</td>
    <td style="text-align:center">$incidents</td>
  </tr>
</table>
<h3>DOWN Events</h3>
<table border="1" style="border-collapse:collapse;width:100%;font-size:12px">
  <tr style="background:#f0f0f0">
    <th>Timestamp</th><th>HTTP Code</th><th>Response (ms)</th>
    <th>Error Type</th><th>Detailed Status</th>
  </tr>
  $downTableRows
</table>
<h3>Response Time (UP checks)</h3>
<p>Avg: ${avgMs}ms &nbsp;|&nbsp; Min: ${minMs}ms &nbsp;|&nbsp; Max: ${maxMs}ms</p>
$sslWarning
"@
        Send-MonitorEmail -To $DailyGroup -Subject $subj -HTMLBody $body
    } catch {
        Write-ErrorLog "Failed to build daily report" -Exception $_.Exception
    }
}

function Send-WeeklyReport {
    $weekStart = (Get-Date).Date.AddDays(-6)
    $subj      = "[SAS Monitor] Weekly Status Report — Week of $($weekStart.ToString('MM/dd/yyyy'))"
    try {
        $csvPath = Get-MonthlyCSVPath
        $allRows = @()
        if (Test-Path $csvPath) {
            $allRows = Import-Csv $csvPath |
                Where-Object { ([datetime]$_.Timestamp_ISO8601) -ge $weekStart }
        }

        $totalUp   = ($allRows | Where-Object { $_.ServerStatus -eq "UP" }).Count
        $totalAll  = $allRows.Count
        $weekUpPct = if ($totalAll -gt 0) { [math]::Round(($totalUp / $totalAll) * 100, 2) } else { 100 }

        # Day-by-day breakdown table
        $dayRows = ""
        for ($i = 6; $i -ge 0; $i--) {
            $day     = (Get-Date).Date.AddDays(-$i)
            $dayData = $allRows | Where-Object { ([datetime]$_.Timestamp_ISO8601).Date -eq $day }
            $dayUp   = ($dayData | Where-Object { $_.ServerStatus -eq "UP" }).Count
            $dayAll  = $dayData.Count
            $dayPct  = if ($dayAll -gt 0) { [math]::Round(($dayUp / $dayAll) * 100, 2) } else { 100 }
            $dayInc  = ($dayData | Where-Object { $_.OutageEventID -ne "" } |
                        Select-Object -ExpandProperty OutageEventID -Unique).Count
            $dayRows += "<tr><td>$($day.ToString('ddd MM/dd'))</td><td>$dayAll</td>" +
                        "<td>$dayPct%</td><td>$dayInc</td></tr>"
        }

        # Longest outage
        $longestOutage = "No outages this week"
        $outageIDs = $allRows | Where-Object { $_.OutageEventID -ne "" } |
                     Select-Object -ExpandProperty OutageEventID -Unique
        if ($outageIDs.Count -gt 0) {
            $longestDur    = [timespan]::Zero
            $longestDetail = ""
            foreach ($oid in $outageIDs) {
                $outRows = $allRows | Where-Object { $_.OutageEventID -eq $oid }
                $dur     = [timespan]::FromSeconds($outRows.Count * $PollingIntervalSec)
                if ($dur -gt $longestDur) {
                    $longestDur    = $dur
                    $start         = ($outRows | Select-Object -First 1).Timestamp_ISO8601
                    $errType       = ($outRows | Select-Object -First 1).ServerError
                    $longestDetail = "Start: $start | Duration: $($dur.ToString('hh\:mm\:ss')) | Error: $errType"
                }
            }
            $longestOutage = $longestDetail
        }

        # Response time trend
        $trendRows = ""
        for ($i = 6; $i -ge 0; $i--) {
            $day   = (Get-Date).Date.AddDays(-$i)
            $dayUp = $allRows | Where-Object {
                ([datetime]$_.Timestamp_ISO8601).Date -eq $day -and $_.ServerStatus -eq "UP"
            }
            $times = $dayUp | ForEach-Object { [double]$_.ResponseTime_ms }
            $avg   = if ($times.Count -gt 0) { [math]::Round(($times | Measure-Object -Average).Average, 1) } else { "N/A" }
            $min   = if ($times.Count -gt 0) { ($times | Measure-Object -Minimum).Minimum } else { "N/A" }
            $max   = if ($times.Count -gt 0) { ($times | Measure-Object -Maximum).Maximum } else { "N/A" }
            $trendRows += "<tr><td>$($day.ToString('ddd MM/dd'))</td><td>$avg</td><td>$min</td><td>$max</td></tr>"
        }

        $sslWarning = ""
        if ($allRows.Count -gt 0) {
            $lastSSL = $allRows | Select-Object -Last 1
            if ([int]$lastSSL.SSLDaysUntilExpiry -lt 30 -and [int]$lastSSL.SSLDaysUntilExpiry -ge 0) {
                $sslWarning = "<p style='color:#e67e22'><strong>&#9888; SSL Certificate expiring in $($lastSSL.SSLDaysUntilExpiry) days</strong></p>"
            }
        }

        $body = @"
<h2>SAS Monitor — Weekly Status Report</h2>
<p><strong>Week of:</strong> $($weekStart.ToString('MM/dd/yyyy')) — $(Get-Date -Format 'MM/dd/yyyy')</p>
<p><strong>7-Day Uptime:</strong> $weekUpPct%</p>
<h3>Day-by-Day Breakdown</h3>
<table border="1" style="border-collapse:collapse;width:100%;font-size:12px">
  <tr style="background:#f0f0f0">
    <th>Day</th><th>Checks</th><th>Uptime%</th><th>Incidents</th>
  </tr>
  $dayRows
</table>
<h3>Longest Outage</h3>
<p>$longestOutage</p>
<h3>Response Time Trend (ms)</h3>
<table border="1" style="border-collapse:collapse;width:100%;font-size:12px">
  <tr style="background:#f0f0f0"><th>Day</th><th>Avg</th><th>Min</th><th>Max</th></tr>
  $trendRows
</table>
$sslWarning
"@
        Send-MonitorEmail -To $WeeklyGroup -Subject $subj -HTMLBody $body
    } catch {
        Write-ErrorLog "Failed to build weekly report" -Exception $_.Exception
    }
}

#endregion EMAIL FUNCTIONS

###############################################################################
#region MAIN — State Variables and Monitoring Loop
###############################################################################

# Per-outage state (reset on recovery)
$script:ConsecutiveFailures       = 0
$script:OutageEventID             = ""
$script:AlertSentForCurrentOutage = $false
$script:OutageStartTime           = $null
$script:LastDailyReportDate       = [datetime]::MinValue
$script:LastWeeklyReportDate      = [datetime]::MinValue

# Ensure all storage directories exist before writing anything
Get-EnsureDir $DataDir
Get-EnsureDir $LogDir

Write-RunLog "SAS Monitor Lite started. Polling: $MonitorURL | Interval: ${PollingIntervalSec}s"
Write-RunLog "DataDir: $DataDir"
Write-RunLog "LogDir:  $LogDir"

# Log a clean shutdown message if PowerShell exits for any reason
Register-EngineEvent PowerShell.Exiting -Action {
    Write-RunLog "SAS Monitor Lite terminated." -Level "WARN"
}

# ---- MONITORING LOOP --------------------------------------------------------
do {
    try {
        $check = Invoke-URLCheck

        # ---------- Evaluate result and update outage state ------------------
        if ($check.ServerStatus -eq "DOWN") {
            $script:ConsecutiveFailures++
            Write-RunLog "CHECK DOWN ($($script:ConsecutiveFailures)x): $($check.DetailedStatus) [$($check.ServerError)]" -Level "WARN"

            # Open a new outage event on the first failure
            if ($script:OutageEventID -eq "") {
                $script:OutageEventID   = (New-Guid).Guid
                $script:OutageStartTime = Get-Date
                Write-RunLog "New outage event: $script:OutageEventID"
            }

            # Send URGENT alert once per outage when threshold is crossed
            if ($script:ConsecutiveFailures -ge $FailureThreshold -and
                -not $script:AlertSentForCurrentOutage) {
                Send-UrgentOutageEmail -Check $check
                $script:AlertSentForCurrentOutage = $true
            }

        } elseif ($check.ServerStatus -eq "UP") {
            if ($script:ConsecutiveFailures -gt 0) {
                # Server came back — close the outage event
                $duration = (Get-Date) - $script:OutageStartTime
                Write-RunLog "RECOVERY: Server UP after $($duration.ToString('hh\:mm\:ss')) outage"
                Send-RecoveryEmail -Check $check -Duration $duration

                $script:ConsecutiveFailures       = 0
                $script:OutageEventID             = ""
                $script:AlertSentForCurrentOutage = $false
                $script:OutageStartTime           = $null
            } else {
                # Normal healthy check — only log at DEBUG to avoid log spam
                Write-RunLog "CHECK OK: $($check.DetailedStatus) | $($check.ResponseTime_ms)ms" -Level "DEBUG"
                $script:ConsecutiveFailures = 0
            }
        }

        # ---------- Append check record to monthly CSV -----------------------
        try {
            $check | Export-Csv -Path (Get-MonthlyCSVPath) -Append -NoTypeInformation -ErrorAction Stop
        } catch {
            Write-ErrorLog "Failed to write CSV record" -Exception $_.Exception
        }

        # ---------- Scheduled email timing -----------------------------------
        $now = Get-Date
        if ($now.Hour -eq $DailyReportHour -and
            $script:LastDailyReportDate.Date -ne $now.Date) {
            Send-DailyReport
            $script:LastDailyReportDate = $now
        }
        if ($now.DayOfWeek.ToString() -eq $WeeklyReportDay -and
            $now.Hour -eq $WeeklyReportHour -and
            $script:LastWeeklyReportDate.Date -ne $now.Date) {
            Send-WeeklyReport
            $script:LastWeeklyReportDate = $now
        }

    } catch {
        # Never let an unhandled exception kill the monitoring loop
        Write-ErrorLog "Unhandled exception in main loop" -Exception $_.Exception
    }

    if ($TestMode -or $CheckOnce) { break }
    Start-Sleep -Seconds $PollingIntervalSec

} while ($true)

# ---------- Post-loop: TestMode output ---------------------------------------
if ($TestMode) {
    $check | Format-List
    Send-MonitorEmail -To $UrgentGroup `
        -Subject "[TEST MODE] SAS Monitor Lite Config Test — $(Get-Date -Format 'o')" `
        -HTMLBody "<h2>Test mode</h2><p>SMTP is configured correctly if you received this.</p>"
    Write-RunLog "Test mode complete."
}

#endregion MAIN
