<#
=====================================================================
  Program Name  : downloadBulkFiles.ps1
  Author        : Manuel Figallo
  Purpose       : Download every attachment referenced by a CSV column of
                  HTTP links (e.g. regulations.gov comment attachments).
  Version       : 1.0beta
  Created       : 2026-08-26
  Last Modified : 2026-08-26

  Dependencies: none (Invoke-WebRequest is built in).

  Description:
    PowerShell twin of src/py/downloadBulkFiles.py. A cell may be BLANK, a
    single URL, or several joined by '|'. Blank cells are skipped, not
    errors. Saved names are prefixed with the row id so identical
    attachment names from different comments cannot collide. A failed
    download is logged and the run continues.

  Input Parameters (required first):
    -InputCsvPath (REQUIRED)  -OutputFolder (REQUIRED)
    -LinkColumn (default attachmentLinks)  -IdColumn (default commentId)
    -Overwrite (default false)
  Exit codes: 0 = success, 2 = config error, 3 = I/O error.
=====================================================================
#>
[CmdletBinding()]
param(
    [string] $InputCsvPath   = '',
    [string] $OutputFolder   = '',
    [string] $LinkColumn     = 'attachmentLinks',
    [string] $IdColumn       = 'commentId',
    [string] $Overwrite      = 'false',
    [int]    $TimeoutSeconds = 60
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\cgsUtils.ps1"

function Invoke-Main {
    <# .SYNOPSIS Download every link in the CSV column. .OUTPUTS [int] code. #>
    if (-not $InputCsvPath) { Write-CgsError "required parameter 'InputCsvPath' is missing or empty"; return 2 }
    if (-not $OutputFolder) { Write-CgsError "required parameter 'OutputFolder' is missing or empty"; return 2 }
    if (-not (Test-Path -LiteralPath $InputCsvPath -PathType Leaf)) {
        Write-CgsError "InputCsvPath not found: $InputCsvPath"; return 3
    }
    $overwriteFlag = ConvertTo-CgsBool $Overwrite
    if ($null -eq $overwriteFlag) { Write-CgsError "unknown Overwrite '$Overwrite'; expected true/false"; return 2 }

    $rows = @(Import-Csv -LiteralPath $InputCsvPath)
    if ($rows.Count -gt 0 -and -not $rows[0].PSObject.Properties[$LinkColumn]) {
        Write-CgsError ("column '$LinkColumn' not found in $InputCsvPath. Available columns: " +
                        (($rows[0].PSObject.Properties | ForEach-Object { $_.Name }) -join ', '))
        return 2
    }
    if (-not (Test-Path -LiteralPath $OutputFolder)) { [void](New-Item -ItemType Directory -Path $OutputFolder -Force) }

    $downloaded = 0; $skipped = 0; $failed = 0
    foreach ($row in $rows) {
        $cell = [string]$row.$LinkColumn
        if ([string]::IsNullOrWhiteSpace($cell)) { $skipped++; continue }   # BLANK is normal
        $rowId = if ($IdColumn -and $row.PSObject.Properties[$IdColumn]) { ([string]$row.$IdColumn).Trim() } else { '' }
        foreach ($url in ($cell -split '\|' | Where-Object { $_.Trim() })) {
            $url = $url.Trim()
            $name = [System.IO.Path]::GetFileName(([uri]$url).LocalPath)
            if (-not $name) { Write-CgsWarn "cannot derive a filename from $url; skipped"; $failed++; continue }
            $target = Join-Path $OutputFolder $(if ($rowId) { "${rowId}_${name}" } else { $name })
            if ((Test-Path -LiteralPath $target) -and -not $overwriteFlag) {
                Write-CgsInfo "exists, skipping: $(Split-Path -Leaf $target)"; $skipped++; continue
            }
            try {
                Invoke-WebRequest -Uri $url -OutFile $target -TimeoutSec $TimeoutSeconds -UseBasicParsing
                $downloaded++; Write-CgsInfo "downloaded $(Split-Path -Leaf $target)"
            } catch {
                $failed++; Write-CgsError "failed ${url}: $($_.Exception.Message)"
            }
        }
    }
    Write-CgsInfo "done; downloaded=$downloaded skipped=$skipped failed=$failed"
    return 0
}
try { exit (Invoke-Main) }
catch { Write-CgsError ("unhandled error: {0}" -f $_.Exception.Message); exit 3 }
