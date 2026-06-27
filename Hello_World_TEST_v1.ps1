<#
  Script Name   : Hello_World_TEST_v1.ps1
  Author        : Manuel Figallo
  Purpose       : Validation/test script confirming that logging and SMTP
                  email delivery work correctly in this environment before
                  building larger SAS monitoring automation on top of it.
  Version       : 1.0.0
  Created       : <date placeholder>
  Last Modified : <date placeholder>

  Description:
    On each execution, this script:
      1. Writes a single log line "Hello world on MM_DD_YYYY at HH:MM:SS"
         to a timestamped log file in $LogDir.
      2. Sends a test email to every address in $DailyGroup via $SMTPServer.
      3. Records the email send attempt in a monthly email audit log if
         $EnableEmailAuditLog is $true.
    This script is designed to be invoked repeatedly by an external
    scheduler (SAS systask facility, Windows Task Scheduler, or a SAS
    program using the X statement / PIPE engine). It performs ONE run per
    invocation; it does not loop internally. Scheduling frequency is
    controlled entirely by how often the external scheduler calls it —
    see the $PollingIntervalSec note in the CONFIGURATION region for how
    this value maps to scheduler frequency when this script's pattern is
    later extended into a persistent monitor.

  Usage:
    Interactive : .\Hello_World_TEST_v1.ps1
    Verbose     : .\Hello_World_TEST_v1.ps1 -Verbose
    Via systask : see embedded comment block below and companion guide

  Change Log:
    v1.0.0 — Initial release
#>

#Requires -Version 5.1
[CmdletBinding()]
param()

###############################################################################
#region CONFIGURATION
###############################################################################

# --- LOG STORAGE (text log files — separate from any future data store) ---
$LogDir              = "\\a70tucgssasr006\custom\projects\sas_monitor\tests"
                        # UNC path for run/error logs. Must be writable by
                        # the account that runs this script.
$LogRetentionDays    = 90      # auto-purge log files older than N days
$LogLevel            = "INFO"  # minimum level to write: DEBUG | INFO | WARN | ERROR
$EnableEmailAuditLog = $true   # append every email send attempt to a monthly audit log

# --- EMAIL / SMTP ---
$SMTPServer     = "smtprelay.TESTTEST.com"        # SMTP relay hostname or IP
$SMTPPort       = 25                              # SMTP port (25 = relay, 587 = submission)
$FromAddress    = "Manuel.Figallo@cgsadmin.com"   # envelope From address
$UseSSL         = $false                          # $true for STARTTLS / implicit TLS
$SMTPCredential = $null                           # set to (Get-Credential) if SMTP auth required

# --- DISTRIBUTION LISTS ---
$DailyGroup = @("Manuel.Figallo@cgsadmin.com", "mfigallo4cgs@gmail.com")
                        # recipients of the test email on every run

# --- SCHEDULING REFERENCE VALUES ---
# These three variables are NOT used by this one-shot script directly —
# they document the intended polling cadence for when this pattern is
# extended into a persistent monitoring loop (e.g. SAS_Monitor.ps1).
# Keep them here so the configuration block is forward-compatible.
$PollingIntervalSec = 3660   # seconds between checks (~hourly); use 21660 for
                             # ~6 hours, 86400 for daily, in a future loop version
$TimeoutSec         = 30    # max seconds to wait for any future HTTP response
$FailureThreshold   = 1     # consecutive failures before an URGENT alert,
                             # in a future monitoring version

# --- SCRIPT IDENTITY (used in log lines and email subject) ---
$ScriptName = "Hello_World_TEST_v1"

#endregion CONFIGURATION

###############################################################################
#region LOGGING FUNCTIONS
###############################################################################

function Get-EnsureDir {
    param([string]$Path)
    # Creates the directory if it does not exist. This script cannot proceed
    # without a writable log directory, so a failure here is fatal (re-throw).
    try {
        if (-not (Test-Path -Path $Path)) {
            New-Item -ItemType Directory -Path $Path -Force -ErrorAction Stop | Out-Null
            Write-Verbose "Created directory: $Path"
        }
    } catch {
        Write-Warning "Failed to create or access directory '$Path': $($_.Exception.Message)"
        throw   # fatal — re-throw so the top-level catch sets exit code 1
    }
}

function Write-TestLog {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    # Level hierarchy: DEBUG(0) < INFO(1) < WARN(2) < ERROR(3).
    # Only write the entry if $Level is at or above the configured $LogLevel.
    $levelRank = @{ "DEBUG" = 0; "INFO" = 1; "WARN" = 2; "ERROR" = 3 }

    # Resolve numeric ranks, defaulting unknown values to INFO(1).
    $thisRank = if ($levelRank.ContainsKey($Level))    { $levelRank[$Level]    } else { 1 }
    $minRank  = if ($levelRank.ContainsKey($LogLevel)) { $levelRank[$LogLevel] } else { 1 }
    if ($thisRank -lt $minRank) { return }   # below threshold — skip entirely

    # Build the structured log line and the per-day log file path.
    $line    = "[$(Get-Date -Format 'o')] [$Level] $Message"
    $logFile = "$LogDir\$ScriptName`_run_$(Get-Date -Format 'yyyy-MM-dd').log"

    # Always echo to the console, color-coded by severity for easy reading.
    switch ($Level) {
        "ERROR" { Write-Host $line -ForegroundColor Red }
        "WARN"  { Write-Host $line -ForegroundColor Yellow }
        default { Write-Host $line }
    }

    # Persist to file — a logging failure must never crash the script.
    try {
        Add-Content -Path $logFile -Value $line -ErrorAction Stop
    } catch {
        Write-Warning "Could not write to log file '$logFile': $($_.Exception.Message)"
    }
}

function Write-EmailAuditLog {
    param(
        [string]$Recipients,
        [string]$Subject,
        [string]$Result
    )
    # Email auditing is optional — bail out immediately when disabled.
    if (-not $EnableEmailAuditLog) { return }

    # Monthly audit file (one file per calendar month).
    $auditFile = "$LogDir\$ScriptName`_email_audit_$(Get-Date -Format 'yyyy-MM').log"
    $auditLine = "[$(Get-Date -Format 'o')] | TO: $Recipients | SUBJECT: $Subject | RESULT: $Result"

    try {
        Add-Content -Path $auditFile -Value $auditLine -ErrorAction Stop
    } catch {
        Write-Warning "Could not write email audit log '$auditFile': $($_.Exception.Message)"
    }

    # Mirror the outcome into the structured run log at the appropriate level.
    if ($Result -like "OK*") {
        Write-TestLog "Email audit: TO=$Recipients RESULT=$Result" -Level INFO
    } else {
        Write-TestLog "Email audit: TO=$Recipients RESULT=$Result" -Level ERROR
    }
}

function Remove-OldLogs {
    # Auto-purge *.log files in $LogDir older than $LogRetentionDays.
    # A purge failure should warn, not crash the script.
    try {
        $cutoff = (Get-Date).AddDays(-$LogRetentionDays)
        Get-ChildItem -Path $LogDir -Filter "*.log" -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -lt $cutoff } |
            ForEach-Object {
                try {
                    Remove-Item -Path $_.FullName -Force -ErrorAction Stop
                    Write-TestLog "Purged old log: $($_.Name)" -Level DEBUG
                } catch {
                    Write-Warning "Could not purge '$($_.Name)': $($_.Exception.Message)"
                }
            }
    } catch {
        Write-Warning "Log purge scan failed: $($_.Exception.Message)"
    }
}

#endregion LOGGING FUNCTIONS

###############################################################################
#region EMAIL FUNCTION
###############################################################################

function Send-TestEmail {
    param(
        [string[]]$To,
        [string]$Subject,
        [string]$HTMLBody
    )
    $toStr = $To -join ";"
    try {
        # Construct the SMTP client and the message object.
        $smtp           = [System.Net.Mail.SmtpClient]::new($SMTPServer, $SMTPPort)
        $smtp.EnableSsl = $UseSSL
        if ($null -ne $SMTPCredential) {
            # Supply network credentials only when SMTP auth is configured.
            $smtp.Credentials = $SMTPCredential.GetNetworkCredential()
        }

        $mail            = [System.Net.Mail.MailMessage]::new()
        $mail.From       = $FromAddress
        $mail.Subject    = $Subject
        $mail.Body       = $HTMLBody
        $mail.IsBodyHtml = $true
        foreach ($addr in $To) { $mail.To.Add($addr) }   # add each recipient

        $smtp.Send($mail)        # blocking send
        $mail.Dispose()
        $smtp.Dispose()

        Write-Verbose "Email sent to $toStr via $SMTPServer`:$SMTPPort"
        Write-EmailAuditLog -Recipients $toStr -Subject $Subject -Result "OK"
    } catch {
        # Record the failure in the audit log AND the structured run log.
        Write-EmailAuditLog -Recipients $toStr -Subject $Subject -Result "FAILED: $($_.Exception.Message)"
        Write-TestLog "Email send failed: $($_.Exception.Message)" -Level ERROR
    }
}

#endregion EMAIL FUNCTION

###############################################################################
#region MAIN SCRIPT BODY
###############################################################################

# Wrap the entire body so any unhandled exception sets exit code 1, allowing
# systask / SAS to detect failure via the return code.
try {

    # 1. Ensure the log directory exists (fatal if it cannot be created).
    Get-EnsureDir $LogDir

    # 2. Purge logs older than the retention window.
    Remove-OldLogs

    # 3-4. Startup banner + execution context.
    Write-TestLog "=== $ScriptName started ===" -Level INFO
    Write-TestLog "Running as: $env:USERNAME on $env:COMPUTERNAME" -Level INFO
    Write-Verbose "PowerShell version: $($PSVersionTable.PSVersion)"

    # 5. REQUIRED LOG STATEMENT — "Hello world on MM_DD_YYYY at HH:MM:SS".
    $dateStamp    = Get-Date -Format "MM_dd_yyyy"
    $timeStamp    = Get-Date -Format "HH:mm:ss"
    $helloMessage = "Hello world on $dateStamp at $timeStamp"
    Write-TestLog $helloMessage -Level INFO

    # Also write the hello line to its own simple log file, satisfying the
    # literal "produce a log in same folder with the statement" requirement
    # in the simplest possible form.
    $simpleLogPath = "$LogDir\$ScriptName`_hello_$(Get-Date -Format 'yyyy-MM-dd').log"
    try {
        Add-Content -Path $simpleLogPath -Value $helloMessage -ErrorAction Stop
        Write-Verbose "Wrote hello line to simple log: $simpleLogPath"
    } catch {
        Write-TestLog "Failed to write simple hello log '$simpleLogPath': $($_.Exception.Message)" -Level ERROR
    }

    # 6. SEND EMAIL — build the structured run-log path for inclusion in the body.
    $runLogPath   = "$LogDir\$ScriptName`_run_$(Get-Date -Format 'yyyy-MM-dd').log"
    $emailSubject = "[$ScriptName] Test Run — $helloMessage"
    $emailBody    = @"
<h2>Hello World Test &mdash; Execution Confirmation</h2>
<p><strong>Message:</strong> $helloMessage</p>
<table style="font-size:13px;border-collapse:collapse">
  <tr><td><strong>Host:</strong></td><td>$env:COMPUTERNAME</td></tr>
  <tr><td><strong>User:</strong></td><td>$env:USERNAME</td></tr>
  <tr><td><strong>Log file:</strong></td><td><code>$runLogPath</code></td></tr>
</table>
<p style="font-size:12px;color:#555;margin-top:12px">
  This is an automated test message confirming that PowerShell logging and
  SMTP email delivery are working correctly from this environment.
</p>
"@

    Send-TestEmail -To $DailyGroup -Subject $emailSubject -HTMLBody $emailBody

    # 7. Completion banner.
    Write-TestLog "=== $ScriptName completed successfully ===" -Level INFO

    # 8a. Explicit success exit code.
    exit 0

} catch {
    # 8b. Any unhandled exception lands here: log it and exit non-zero so the
    #     external scheduler (systask / SAS) can detect the failure.
    Write-TestLog "UNHANDLED EXCEPTION: $($_.Exception.Message)" -Level ERROR
    exit 1
}

#endregion MAIN SCRIPT BODY

<#
=====================================================================
SCHEDULER INTEGRATION REFERENCE
=====================================================================
This script is designed to be called repeatedly by an external
scheduler rather than running its own internal loop. Two supported
integration patterns:

1. SAS systask facility (see companion deployment guide for full detail):
   systask command
     "powershell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass -File ""\\a70tucgssasr006\custom\projects\sas_monitor\Hello_World_TEST_v1.ps1"""
     taskname = "Hello_World_TEST_v1"
     status   = rc
     nowait;

2. SAS program using the X statement or PIPE engine, scheduled in your
   organization's existing SAS batch scheduler (see companion guide).

Exit code 0 = success, 1 = failure. Check $SYSRC / the systask "status=rc"
variable, or your scheduler's return-code mechanism, to detect failures.
=====================================================================
#>
