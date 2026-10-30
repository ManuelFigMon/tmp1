<#
=====================================================================
  Program Name  : sendEmail.ps1
  Author        : Manuel Figallo
  Purpose       : Send email alerts over SMTP.
  Version       : 1.0beta
  Created       : 2026-08-26
  Last Modified : 2026-09-06

  Dependencies: none (System.Net.Mail is part of .NET).

  Description:
    PowerShell twin of src/py/sendEmail.py -- same function name and same
    parameter names. Multiple recipients allowed in To (array or
    ';'-delimited). SmtpServer/Port default to smtp.example.com:25, or
    SMTP_SERVER/SMTP_PORT from .env. All other parameters are required.

    Built on System.Net.Mail rather than Send-MailMessage: Send-MailMessage
    cannot add a custom header, and the Outlook red flag IS a custom header
    (X-Message-Flag). Send-MailMessage is also obsolete.

  HOW TO FORMAT AN EMAIL
    1. ADDRESSES. Give every address a real name, so the recipient sees
       "Manuel Figallo" in their inbox instead of a raw mailbox:
           -From "Manuel Figallo <manuel.figallo@cgsadmin.com>"
           -To   "Al Cordoba <al.cordoba@cgsadmin.com>"
       The form is  Display Name <address>. Separate several recipients
       with a SEMICOLON, never a comma -- a comma is legal INSIDE a display
       name ("Cordoba, Al"), so splitting on it would tear the address in
       half.

    2. PLAIN TEXT OR HTML. Pass -Html true to send HTML. Alignment, bold
       and spacing DO NOT EXIST in a plain text mail, so anything centered,
       indented or coloured needs -Html true.

    3. LAYOUT. One idea per paragraph, and the deliverable's location on
       its own line -- an operator reading on a phone should find the path
       without scrolling sideways.

    4. URGENCY. -Urgent true marks the message high priority AND sets the
       Outlook message flag, so it arrives with a red flag. Use it for
       something that needs action tonight, not for every run: a mailbox
       where everything is urgent has nothing urgent in it.

    5. SUBJECT. Say the outcome, not the mechanism -- "Issue Log of DB
       Tables - scan complete" beats "SAS job finished".

  Input Parameters (required first):
    -To (REQUIRED) -From (REQUIRED) -Subject (REQUIRED) -Body (REQUIRED)
    -SmtpServer (default smtp.example.com)  -Port (default 25)
    -Html (default false)   -Urgent (default false)
    -UrgentFlag (default 'Follow up'; the text Outlook shows on the flag)
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
    # [string] not [bool]: -File mode passes every argument as a string.
    [string]   $Html       = 'false',
    [string]   $Urgent     = 'false',
    [string]   $UrgentFlag = 'Follow up'
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
    $isHtml   = [bool](ConvertTo-CgsBool $Html)
    $isUrgent = [bool](ConvertTo-CgsBool $Urgent)

    $message = New-Object System.Net.Mail.MailMessage
    try {
        # MailAddress parses "Display Name <address>", so the recipient sees
        # the name in their inbox rather than a bare mailbox.
        $message.From = New-Object System.Net.Mail.MailAddress($From)
        foreach ($recipient in $recipients) { $message.To.Add($recipient) }
        $message.Subject    = $Subject
        $message.Body       = $Body
        $message.IsBodyHtml = $isHtml

        if ($isUrgent) {
            # Priority sets Importance / X-Priority: the red exclamation mark
            # and the sort to the top. The red FLAG is a separate header --
            # both are set so the message stands out in either view.
            $message.Priority = [System.Net.Mail.MailPriority]::High
            $message.Headers.Add('X-Message-Flag', $UrgentFlag)
        }

        $urgentLabel = if ($isUrgent) { 'URGENT ' } else { '' }
        Write-CgsInfo "sending ${urgentLabel}mail to $($recipients.Count) recipient(s) via ${server}:${portNo}"
        $client = New-Object System.Net.Mail.SmtpClient($server, $portNo)
        try { $client.Send($message) } finally { $client.Dispose() }
    } finally { $message.Dispose() }

    Write-CgsInfo "sent: $Subject"
    return 0
}
try { exit (Invoke-Main) }
catch { Write-CgsError ("send failed: {0}" -f $_.Exception.Message); exit 3 }
