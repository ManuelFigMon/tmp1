<#
=====================================================================
  Program Name  : collectSystemMetrics.ps1
  Author        : Manuel Figallo
  Purpose       : Collect as many host metrics as the platform allows and
                  append or overwrite them in a CSV time series.
  Version       : 1.0beta
  Created       : 2026-08-26
  Last Modified : 2026-08-26

  Dependencies: none (Get-Counter / CIM are built in on Windows).

  Description:
    PowerShell twin of src/py/collectSystemMetrics.py. FAILS GRACEFULLY BY
    DESIGN: every probe is wrapped independently, so a counter that does not
    exist on this server is recorded blank and the run still succeeds. Names
    of failed probes land in the row's Errors column.

  Input Parameters (required first):
    -OutputCsvPath (REQUIRED)  -ServerName (defaults to $env:COMPUTERNAME)
    -WriteMode (default append; append|overwrite)
  Exit codes: 0 = success, 2 = config error, 3 = cannot write the CSV.
=====================================================================
#>
[CmdletBinding()]
param(
    [string] $OutputCsvPath = '',
    [string] $ServerName    = '',
    [string] $WriteMode     = 'append'
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\cgsUtils.ps1"

$script:MetricColumns = @(
    'Timestamp','ServerName','OSVersion','CPUName','CPUCount','CPUUsagePercent',
    'TotalPhysicalMemoryGB','MemoryAvailableMB','MemoryUsedPercent',
    'DiskTotalGB','DiskFreeGB','DiskFreePercent','PowerShellVersion','Errors'
)

function Invoke-SafeProbe {
    <# .SYNOPSIS Run one probe, recording rather than raising on failure.
       .PARAMETER Probe   Scriptblock returning the metric.
       .PARAMETER Label   Metric name used in the error note.
       .PARAMETER Errors  [ref] list accumulating failure labels.
       .OUTPUTS  The value, or '' when the probe failed. #>
    param([scriptblock] $Probe, [string] $Label, [ref] $Errors)
    try {
        $value = & $Probe
        if ($null -eq $value) { return '' }
        return $value
    } catch {
        $Errors.Value += ("{0}:{1}" -f $Label, $_.Exception.GetType().Name)
        return ''
    }
}

function Invoke-Main {
    <# .SYNOPSIS Gather metrics and write the CSV. .OUTPUTS [int] exit code. #>
    if (-not $OutputCsvPath) { Write-CgsError "required parameter 'OutputCsvPath' is missing or empty"; return 2 }
    if (@('append','overwrite') -notcontains $WriteMode) {
        Write-CgsError "unknown WriteMode '$WriteMode'; expected one of: append, overwrite"; return 2
    }
    $errors = @()
    $host_ = if ($ServerName) { $ServerName } else { $env:COMPUTERNAME }
    if (-not $host_) { $host_ = 'unknown' }

    $cpuUsage = Invoke-SafeProbe -Label 'cpuUsage' -Errors ([ref]$errors) -Probe {
        [math]::Round((Get-Counter '\Processor(_Total)\% Processor Time' -ErrorAction Stop).CounterSamples.CookedValue, 2) }
    $memAvail = Invoke-SafeProbe -Label 'memAvailable' -Errors ([ref]$errors) -Probe {
        [math]::Round((Get-Counter '\Memory\Available MBytes' -ErrorAction Stop).CounterSamples.CookedValue, 2) }
    $os = Invoke-SafeProbe -Label 'os' -Errors ([ref]$errors) -Probe {
        (Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop).Version }
    $cpuName = Invoke-SafeProbe -Label 'cpuName' -Errors ([ref]$errors) -Probe {
        (Get-CimInstance -ClassName Win32_Processor -ErrorAction Stop | Select-Object -First 1).Name }
    $totalMem = Invoke-SafeProbe -Label 'totalMemory' -Errors ([ref]$errors) -Probe {
        [math]::Round((Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop).TotalPhysicalMemory / 1GB, 2) }
    $disk = Invoke-SafeProbe -Label 'disk' -Errors ([ref]$errors) -Probe {
        Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction Stop }

    $memUsedPct = ''
    if ($totalMem -and $memAvail) {
        $memUsedPct = [math]::Round((1 - ($memAvail / 1024) / [double]$totalMem) * 100, 2)
    }
    $diskTotal = ''; $diskFree = ''; $diskFreePct = ''
    if ($disk) {
        $diskTotal   = [math]::Round($disk.Size / 1GB, 2)
        $diskFree    = [math]::Round($disk.FreeSpace / 1GB, 2)
        if ($disk.Size) { $diskFreePct = [math]::Round($disk.FreeSpace / $disk.Size * 100, 2) }
    }

    $row = [PSCustomObject][ordered]@{
        Timestamp             = (Get-Date -Format "MM/dd/yyyy HH:mm:ss")
        ServerName            = $host_
        OSVersion             = $os
        CPUName               = $cpuName
        CPUCount              = (Invoke-SafeProbe -Label 'cpuCount' -Errors ([ref]$errors) -Probe { $env:NUMBER_OF_PROCESSORS })
        CPUUsagePercent       = $cpuUsage
        TotalPhysicalMemoryGB = $totalMem
        MemoryAvailableMB     = $memAvail
        MemoryUsedPercent     = $memUsedPct
        DiskTotalGB           = $diskTotal
        DiskFreeGB            = $diskFree
        DiskFreePercent       = $diskFreePct
        PowerShellVersion     = $PSVersionTable.PSVersion.ToString()
        Errors                = ($errors -join ';')
    }

    $all = @()
    if ($WriteMode -eq 'append' -and (Test-Path -LiteralPath $OutputCsvPath)) {
        try { $all = @(Import-Csv -LiteralPath $OutputCsvPath) }
        catch { Write-CgsWarn "cannot read existing $OutputCsvPath; overwriting" }
    }
    $all += $row
    [void](Write-CgsCsv -Rows $all -Columns $script:MetricColumns -Target $OutputCsvPath)

    Write-CgsInfo "collected metrics for $host_ ($WriteMode) -> $OutputCsvPath"
    if ($errors.Count -gt 0) {
        Write-CgsWarn ("{0} metric(s) unavailable on this host: {1}" -f $errors.Count, ($errors -join ';'))
    }
    return 0
}
try { exit (Invoke-Main) }
catch { Write-CgsError ("unhandled error: {0}" -f $_.Exception.Message); exit 3 }
