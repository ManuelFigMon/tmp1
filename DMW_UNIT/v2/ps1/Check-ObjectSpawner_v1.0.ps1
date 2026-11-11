<#
.SYNOPSIS
  Script Name   : Check-ObjectSpawner_v1.0.ps1
  Purpose       : Check the SAS Windows services in UNIT and PROD and emit an
                  OverallStatus the DMW Home tab reads.
  Version       : 1.0.0

  .DESCRIPTION
    Queries the named SAS services on each host and writes one CSV row per
    service per sweep, plus a single OverallStatus row per environment.

    OverallStatus is what the DMW Home tab displays. It is the worst state
    found among the services in that environment:

        HEALTHY   every monitored service is Running
        DEGRADED  at least one non-critical service is not Running
        CRITICAL  at least one service marked Critical is not Running
        UNKNOWN   the host could not be reached at all

    The distinction matters operationally: the Object Spawner being down stops
    workspace servers from launching and is therefore Critical, whereas a
    stopped Connect Spawner may be expected in some configurations.

    OUTPUT -> data\objectspawner_status.csv   (read by the DMW Home tab)

  .PARAMETER TestMode
    Print to console, write nothing.

  .PARAMETER Environment
    Limit the sweep to one environment name, e.g. -Environment UNIT.

  PREREQUISITES
    Remote service queries use Get-Service -ComputerName, which needs the
    Remote Registry service running on the target and the executing account
    to hold at least read rights to the service control manager. If that is
    not available in your environment, set $UseWinRM = $true below to query
    via Invoke-Command instead, which requires WinRM to be configured.

  EXECUTION POLICY (UNC paths)
      Unblock-File -Path "\\server\share\Check-ObjectSpawner_v1.0.ps1"
    or
      powershell.exe -ExecutionPolicy Bypass -File "<path>"

  TASK SCHEDULER
    Program : powershell.exe
    Args    : -NonInteractive -ExecutionPolicy Bypass -File "<path>\Check-ObjectSpawner_v1.0.ps1"
    Trigger : Every 5 minutes
#>

#Requires -Version 5.1
[CmdletBinding()]
param(
    [switch]$TestMode,
    [string]$Environment = ""
)

###############################################################################
#region CONFIGURATION  -- the only section an operator should need to edit
###############################################################################

# --- WHAT TO CHECK -------------------------------------------------------
# ServiceName must be the Windows SERVICE NAME, not the display name. Find it
# with:  Get-Service -ComputerName <host> | Where-Object DisplayName -like '*SAS*'
#
# Critical = $true means a stopped service makes the whole environment CRITICAL
# rather than DEGRADED.
$Services = @(
    [PSCustomObject]@{ Environment="UNIT"; ComputerName="a70tucgssasr006"; ServiceName="SASObjectSpawner";       Label="Object Spawner";       Critical=$true  }
    [PSCustomObject]@{ Environment="UNIT"; ComputerName="a70tucgssasr006"; ServiceName="SASMetadataServer";      Label="Metadata Server";      Critical=$true  }
    [PSCustomObject]@{ Environment="UNIT"; ComputerName="a70tucgssasr006"; ServiceName="SASConnectSpawner";      Label="Connect Spawner";      Critical=$false }
    [PSCustomObject]@{ Environment="PROD"; ComputerName="a70pcgssasr006";  ServiceName="SASObjectSpawner";       Label="Object Spawner";       Critical=$true  }
    [PSCustomObject]@{ Environment="PROD"; ComputerName="a70pcgssasr006";  ServiceName="SASMetadataServer";      Label="Metadata Server";      Critical=$true  }
    [PSCustomObject]@{ Environment="PROD"; ComputerName="a70pcgssasr006";  ServiceName="SASConnectSpawner";      Label="Connect Spawner";      Critical=$false }
)

# --- QUERY METHOD --------------------------------------------------------
$UseWinRM  = $false   # $true -> Invoke-Command (needs WinRM); $false -> Get-Service -ComputerName
$QueryTimeoutSec = 20

# --- OUTPUT --------------------------------------------------------------
$DataDir      = "$PSScriptRoot\..\data"
$DataFileName = "objectspawner_status.csv"
$KeepDays     = 90

$LogDir           = "$PSScriptRoot\..\log"
$LogRetentionDays = 90
$LogLevel         = "INFO"

# --- RECOVERY ------------------------------------------------------------
# Left OFF by design. Restarting a SAS service is not a safe unattended action:
# it drops every live workspace session on that host. Turn this on only after
# agreeing the behaviour with the platform owner.
$EnableAutoRestart   = $false
$RestartOnlyCritical = $true

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
        $file = Join-Path $LogDir ("Check-ObjectSpawner_{0}.log" -f (Get-Date -Format "yyyyMMdd"))
        "[$ts] [$Level] $Message" | Out-File -FilePath $file -Append -Encoding UTF8
        if ($Level -eq "ERROR") { Write-Host "[$ts] [$Level] $Message" -ForegroundColor Red }
        elseif ($Level -eq "WARN") { Write-Host "[$ts] [$Level] $Message" -ForegroundColor Yellow }
    } catch { Write-Host "Logging failed: $($_.Exception.Message)" -ForegroundColor DarkYellow }
}

#endregion LOGGING

###############################################################################
#region CHECK
###############################################################################

function Get-ServiceState {
    <#
      Returns the service State as a string, or "Unreachable" when the host
      could not be queried at all, or "NotFound" when the host answered but
      has no such service. Those two are different problems and the dashboard
      shows them differently.
    #>
    param([string]$ComputerName, [string]$ServiceName)

    try {
        if ($UseWinRM) {
            $job = Invoke-Command -ComputerName $ComputerName -AsJob -ScriptBlock {
                param($svc)
                $s = Get-Service -Name $svc -ErrorAction SilentlyContinue
                if ($null -eq $s) { "NotFound" } else { $s.Status.ToString() }
            } -ArgumentList $ServiceName
            if (Wait-Job $job -Timeout $QueryTimeoutSec) {
                $v = Receive-Job $job
                Remove-Job $job -Force
                return [string]$v
            }
            Stop-Job $job -ErrorAction SilentlyContinue
            Remove-Job $job -Force -ErrorAction SilentlyContinue
            return "Unreachable"
        }
        else {
            $s = Get-Service -ComputerName $ComputerName -Name $ServiceName -ErrorAction Stop
            return $s.Status.ToString()
        }
    } catch {
        if ($_.Exception.Message -match "Cannot find any service") { return "NotFound" }
        return "Unreachable"
    }
}

#endregion CHECK

###############################################################################
#region MAIN
###############################################################################

$set = if ($Environment) { $Services | Where-Object { $_.Environment -eq $Environment } } else { $Services }
if (-not $set) { Write-RunLog "No services matched Environment='$Environment'." -Level "ERROR"; exit 2 }

Write-RunLog "Sweep started ($($set.Count) services)"
$stamp = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")

$rows = foreach ($svc in $set) {
    $state = Get-ServiceState -ComputerName $svc.ComputerName -ServiceName $svc.ServiceName
    $isUp  = ($state -eq "Running")

    $lvl = if ($isUp) { "INFO" } elseif ($svc.Critical) { "ERROR" } else { "WARN" }
    Write-RunLog ("{0}/{1} on {2}: {3}" -f $svc.Environment, $svc.Label, $svc.ComputerName, $state) -Level $lvl

    if (-not $isUp -and $EnableAutoRestart -and ($svc.Critical -or -not $RestartOnlyCritical) -and $state -ne "Unreachable") {
        try {
            Write-RunLog "Attempting restart of $($svc.ServiceName) on $($svc.ComputerName)" -Level "WARN"
            (Get-Service -ComputerName $svc.ComputerName -Name $svc.ServiceName).Start()
            Start-Sleep -Seconds 10
            $state = Get-ServiceState -ComputerName $svc.ComputerName -ServiceName $svc.ServiceName
            Write-RunLog "Post-restart state: $state"
        } catch { Write-RunLog "Restart failed: $($_.Exception.Message)" -Level "ERROR" }
    }

    [PSCustomObject]@{
        Timestamp     = $stamp
        Environment   = $svc.Environment
        ComputerName  = $svc.ComputerName
        ServiceName   = $svc.ServiceName
        Label         = $svc.Label
        Critical      = $svc.Critical
        State         = $state
        IsRunning     = if ($isUp) { 1 } else { 0 }
        OverallStatus = ""     # filled in below, per environment
        CheckedBy     = $env:COMPUTERNAME
    }
}

# --- roll up to OverallStatus per environment ---------------------------
foreach ($envName in ($rows | Select-Object -ExpandProperty Environment -Unique)) {
    $envRows = $rows | Where-Object { $_.Environment -eq $envName }

    $overall =
        if ($envRows | Where-Object { $_.State -eq "Unreachable" }) { "UNKNOWN"  }
        elseif ($envRows | Where-Object { $_.Critical -and $_.IsRunning -eq 0 }) { "CRITICAL" }
        elseif ($envRows | Where-Object { $_.IsRunning -eq 0 }) { "DEGRADED" }
        else { "HEALTHY" }

    foreach ($r in $envRows) { $r.OverallStatus = $overall }
    Write-RunLog "$envName OverallStatus = $overall"
}

if ($TestMode) {
    $rows | Format-Table Environment, Label, State, Critical, OverallStatus -AutoSize
    Write-RunLog "TestMode: CSV not written."
    return
}

try {
    if (-not (Test-Path $DataDir)) { New-Item -ItemType Directory -Path $DataDir -Force | Out-Null }
    $csvPath = Join-Path $DataDir $DataFileName

    if (Test-Path $csvPath) { $rows | Export-Csv -Path $csvPath -Append -NoTypeInformation -Encoding UTF8 -ErrorAction Stop }
    else                    { $rows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop }
    Write-RunLog "Wrote $($rows.Count) rows to $csvPath"

    if ($KeepDays -gt 0) {
        $cutoff = (Get-Date).AddDays(-$KeepDays)
        $all  = Import-Csv -Path $csvPath
        $kept = $all | Where-Object { $d=$null; [datetime]::TryParse($_.Timestamp,[ref]$d) -and $d -ge $cutoff }
        if ($kept.Count -lt $all.Count) {
            $kept | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            Write-RunLog "Pruned $($all.Count - $kept.Count) rows older than $KeepDays days"
        }
    }
} catch {
    Write-RunLog "CSV write failed: $($_.Exception.Message)" -Level "ERROR"
    exit 2
}

$bad = @($rows | Where-Object { $_.OverallStatus -in @("CRITICAL","UNKNOWN") }).Count
Write-RunLog "Sweep complete."
if ($bad -gt 0) { exit 1 }
exit 0

#endregion MAIN
