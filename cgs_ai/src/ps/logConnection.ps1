<#
=====================================================================
  Program Name  : logConnection.ps1
  Author        : Manuel Figallo
  Purpose       : Log the network connections to the SAS Object Spawner
                  port (default 8591) once per interval into a CSV time
                  series, to find WHEN an Enterprise Guide connection
                  disappears (idle timeout, network event or outage).
  Version       : 1.0beta
  Created       : 2026-11-05
  Last Modified : 2026-11-05

  Dependencies: Get-NetTCPConnection (NetTCPIP module, built in on
  Windows 8 / Server 2012 and later). Windows PowerShell 5.1 or 7+.

  Description:
    PowerShell twin of src/sas/logConnection.sas, with one addition. It
    runs in either of two places:

      SERVER mode (default) - run it on the SAS compute server. Logs every
        client connected to the local port, optionally only -client_ip.
        This is what the SAS macro does.

      CLIENT mode (-server_host given) - run it on your own PC. Logs the
        PC's connections to that server's port AND whether the server
        port answers (ServerReachable). The reachability column is the
        evidence that separates a network outage from a server outage.

    Run both at once and line up the Time columns: a connection that
    disappears while ServerReachable stays True points to an idle timeout
    on the path; ServerReachable False points to the network or server.

    A sample with no connection is written as State=none, so a gap in the
    file always means the script was not running.

  Input Parameters (required first):
    -output_file_path (REQUIRED)  .csv, appended if it exists
    -port             default 8591
    -interval_seconds default 60
    -duration_minutes default 60
    -client_ip        SERVER mode only; log just this client
    -server_host      CLIENT mode; the SAS server name or IP

  Output columns: Time, State, RemoteAddress, RemotePort, ServerReachable
    Time is local time, ISO 8601. ServerReachable is blank in SERVER mode.
  Exit codes: 0 = success, 2 = config error, 3 = cannot write the CSV.
=====================================================================
#>
[CmdletBinding()]
param(
    [string] $output_file_path = '',
    [int]    $port             = 8591,
    [int]    $interval_seconds = 60,
    [int]    $duration_minutes = 60,
    [string] $client_ip        = '',
    [string] $server_host      = ''
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\cgsUtils.ps1"

$script:Header = 'Time,State,RemoteAddress,RemotePort,ServerReachable'

function Test-TcpPort {
    <# .SYNOPSIS Can this machine open a TCP connection to host:port?
       .PARAMETER Hostname  Server name or IP.
       .PARAMETER Port      TCP port.
       .PARAMETER TimeoutMs Give up after this long (default 5000).
       .OUTPUTS  [bool]. Faster than Test-NetConnection when the port is closed. #>
    param([string] $Hostname, [int] $Port, [int] $TimeoutMs = 5000)
    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $task = $client.ConnectAsync($Hostname, $Port)
        return ($task.Wait($TimeoutMs) -and $client.Connected)
    } catch { return $false }
    finally { $client.Close() }
}

function Get-PortConnections {
    <# .SYNOPSIS Connections on the port, in SERVER or CLIENT mode.
       .OUTPUTS  Objects with State, RemoteAddress, RemotePort; may be empty. #>
    if ($server_host) {
        $targets = @([System.Net.Dns]::GetHostAddresses($server_host) | ForEach-Object { $_.IPAddressToString })
        $found = @(Get-NetTCPConnection -RemotePort $port -ErrorAction SilentlyContinue |
                   Where-Object { $targets -contains $_.RemoteAddress })
    } elseif ($client_ip) {
        $found = @(Get-NetTCPConnection -LocalPort $port -RemoteAddress $client_ip -ErrorAction SilentlyContinue)
    } else {
        $found = @(Get-NetTCPConnection -LocalPort $port -ErrorAction SilentlyContinue)
    }
    return $found
}

function Invoke-Main {
    <# .SYNOPSIS Sample the port until the duration ends. .OUTPUTS [int] exit code. #>
    if (-not $output_file_path) { Write-CgsError "required parameter 'output_file_path' is missing or empty"; return 2 }
    if ($interval_seconds -lt 1 -or $duration_minutes -lt 1) {
        Write-CgsError 'interval_seconds and duration_minutes must be 1 or more'; return 2
    }
    if (-not (Get-Command Get-NetTCPConnection -ErrorAction SilentlyContinue)) {
        Write-CgsError 'Get-NetTCPConnection is not available on this machine'; return 2
    }
    $mode = if ($server_host) { 'CLIENT' } else { 'SERVER' }
    $samples = [math]::Floor($duration_minutes * 60 / $interval_seconds)

    try {
        $parent = Split-Path -Parent $output_file_path
        if ($parent -and -not (Test-Path -LiteralPath $parent)) { [void](New-Item -ItemType Directory -Path $parent -Force) }
        if (-not (Test-Path -LiteralPath $output_file_path)) { Set-Content -LiteralPath $output_file_path -Value $script:Header }
    } catch {
        Write-CgsError ("cannot write {0}: {1}" -f $output_file_path, $_.Exception.Message); return 3
    }

    Write-CgsInfo ("logConnection {0} mode, port {1}, {2} samples every {3}s -> {4}" -f $mode, $port, $samples, $interval_seconds, $output_file_path)

    for ($i = 1; $i -le $samples; $i++) {
        $stamp = Get-Date -Format 's'
        $reach = ''
        if ($server_host) { $reach = [string](Test-TcpPort -Hostname $server_host -Port $port) }
        $conns = @(Get-PortConnections)
        $lines = @()
        if ($conns.Count -eq 0) {
            $lines += ('{0},none,,,{1}' -f $stamp, $reach)
        } else {
            foreach ($c in $conns) {
                $lines += ('{0},{1},{2},{3},{4}' -f $stamp, $c.State, $c.RemoteAddress, $c.RemotePort, $reach)
            }
        }
        try { Add-Content -LiteralPath $output_file_path -Value $lines }
        catch { Write-CgsError ("cannot write {0}: {1}" -f $output_file_path, $_.Exception.Message); return 3 }
        if ($i -lt $samples) { Start-Sleep -Seconds $interval_seconds }
    }
    Write-CgsInfo ("logConnection finished. Samples written: {0}. Output: {1}" -f $samples, $output_file_path)
    return 0
}
try { exit (Invoke-Main) }
catch { Write-CgsError ("unhandled error: {0}" -f $_.Exception.Message); exit 3 }
