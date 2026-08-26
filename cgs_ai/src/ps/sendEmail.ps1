<#
=====================================================================
  Program Name  : sendEmail.ps1
  Author        : Manuel Figallo
  Purpose       : Send email alerts over SMTP.
  Version       : 1.0beta
  Created       : 2026-08-26
  Last Modified : 2026-08-26

  Dependencies: none (Send-MailMessage is built in).

  Description:
    PowerShell twin of src/py/sendEmail.py -- same function name and same
    parameter names. Multiple recipients allowed in To (array or
    ';'-delimited). SmtpServer/Port default to smtp.example.com:25, or
    SMTP_SERVER/SMTP_PORT from .env. All other parameters are required.

  Input Parameters (required first):
    -To (REQUIRED) -From (REQUIRED) -Subject (REQUIRED) -Body (REQUIRED)
    -SmtpServer (default smtp.example.com)  -Port (default 25)
  Exit codes: 0 = success, 2 = config error, 3 = send failure.
=====================================================================
#>
[CmdletBinding()]
param(
    [string[]] $To         = @(),
    [string]   $From       = '',
    [string]   $Subject    = '',
    [string]   $Body       = '',
    [string]   $SmtpServer = '',
    [int]      $Port       = 0,
    [string]   $Html       = 'false'
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\cgsUtils.ps1"

function Invoke-Main {
    <# .SYNOPSIS Validate and send. .OUTPUTS [int] exit code. #>
    $recipients = ConvertTo-CgsList $To
    if ($recipients.Count -eq 0) { Write-CgsError "required parameter 'To' is missing or empty"; return 2 }
    foreach ($pair in @(@('From',$From), @('Subject',$Subject), @('Body',$Body))) {
        if (-not $pair[1]) { Write-CgsError "required parameter '$($pair[0])' is missing or empty"; return 2 }
    }
    $server = if ($SmtpServer) { $SmtpServer } else { Get-CgsConfig -Key 'SMTP_SERVER' -Default 'smtp.example.com' }
    $portNo = if ($Port -gt 0) { $Port } else { [int](Get-CgsConfig -Key 'SMTP_PORT' -Default '25') }

    Write-CgsInfo "sending mail to $($recipients.Count) recipient(s) via ${server}:${portNo}"
    $mailMessage = @{
        To = $recipients; From = $From; Subject = $Subject
        Body = $Body; SmtpServer = $server; Port = $portNo
    }
    if ((ConvertTo-CgsBool $Html)) { $mailMessage['BodyAsHtml'] = $true }
    Send-MailMessage @mailMessage
    Write-CgsInfo "sent: $Subject"
    return 0
}
try { exit (Invoke-Main) }
catch { Write-CgsError ("send failed: {0}" -f $_.Exception.Message); exit 3 }
