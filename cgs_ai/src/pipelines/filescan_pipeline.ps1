<#
=====================================================================
  Program Name  : filescan_pipeline.ps1
  Author        : Manuel Figallo
  Purpose       : End-to-end pipeline: scan log folders for keywords,
                  render a styled Excel report, and email a completion
                  notice.
  Version       : 1.0beta
  Created       : 2026-08-26
  Last Modified : 2026-08-26

  Dependencies:
    scanFileSystem.ps1 (none), formatCSV.ps1 (ImportExcel), sendEmail.ps1 (none).

  Description:
    PowerShell twin of src/pipelines/filescan_pipeline.py. Orchestration
    only -- all real work lives in the three functions it calls. Email
    failure is reported but does NOT fail the pipeline, because the scan
    output is already on disk and is the deliverable.

  Input Parameters (required first):
    -input_folder_root -extract_keyword -output_file_path -excel_output_path
    -metric_profile (default sas_log) -email_to -email_from -email_subject
  Exit codes: 0 = success, non-zero from the failing stage.
=====================================================================
#>
[CmdletBinding()]
param(
    [string[]] $input_folder_root = @(),
    [string[]] $extract_keyword   = @(),
    [string]   $output_file_path  = '',
    [string]   $excel_output_path = '',
    [string]   $metric_profile    = 'sas_log',
    [string[]] $email_to          = @(),
    [string]   $email_from        = '',
    [string]   $email_subject     = '',
    [string]   $format_type       = 'corporate'
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\..\ps\cgsUtils.ps1"

$script:PsFolder = Join-Path $PSScriptRoot '..\ps'
# The production roots this pipeline was built for.
$script:DefaultRoots = @(
    '\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT\HHH\Old_Programs\Old_logs',
    '\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT\DME\Logs'
)
$script:DefaultKeywords = @('real time','cpu time')

function Invoke-Main {
    <# .SYNOPSIS Run scan -> format -> notify. .OUTPUTS [int] exit code. #>
    $roots    = ConvertTo-CgsList $input_folder_root; if ($roots.Count -eq 0)    { $roots = $script:DefaultRoots }
    $keywords = ConvertTo-CgsList $extract_keyword;   if ($keywords.Count -eq 0) { $keywords = $script:DefaultKeywords }
    $steps = @()
    $stamp = Get-CgsTimestampSuffix
    $dataRoot = Get-CgsConfig -Key 'ROOT_DATA' -Default '.'
    $scanTarget = if ($output_file_path) { $output_file_path } else { Join-Path $dataRoot "scan_$stamp.csv" }

    Write-CgsInfo "filescan_pipeline 1.0beta starting; $($roots.Count) root(s), $($keywords.Count) keyword(s)"

    # --- 1. scan ---------------------------------------------------------
    & (Join-Path $script:PsFolder 'scanFileSystem.ps1') `
        -input_folder_root ($roots -join ';') -extract_keyword ($keywords -join ';') `
        -output_file_path $scanTarget -metric_profile $metric_profile
    if ($LASTEXITCODE -ne 0) { Write-CgsError "scanFileSystem failed ($LASTEXITCODE)"; return $LASTEXITCODE }
    $steps += 'scanFileSystem'

    # metric_profile active means the scan wrote .xlsx instead of .csv.
    $scanOutput = $scanTarget
    if ($metric_profile -ne 'none' -and -not $scanTarget.ToLowerInvariant().EndsWith('.xlsx')) {
        $scanOutput = [System.IO.Path]::ChangeExtension($scanTarget, '.xlsx')
    }
    Write-CgsInfo "scan output: $scanOutput"

    # --- 2. format -------------------------------------------------------
    $reportTarget = if ($excel_output_path) { $excel_output_path } else {
        Join-Path ([System.IO.Path]::GetDirectoryName($scanOutput)) `
            ("{0}_report.xlsx" -f [System.IO.Path]::GetFileNameWithoutExtension($scanOutput)) }
    if ($scanOutput.ToLowerInvariant().EndsWith('.csv')) {
        & (Join-Path $script:PsFolder 'formatCSV.ps1') `
            -InputCsvPath $scanOutput -OutputExcelPath $reportTarget `
            -FormatType $format_type -Title "File Scan $stamp"
        if ($LASTEXITCODE -eq 0) { $steps += 'formatCSV' }
        else { Write-CgsWarn "styled report skipped (formatCSV exit $LASTEXITCODE)" }
    } else {
        # Re-styling would drop the Metrics sheet, so keep the scan workbook.
        Write-CgsInfo 'scan already produced Excel (metric profile active); keeping it as the report'
        $reportTarget = $scanOutput
    }

    # --- 3. notify -------------------------------------------------------
    $recipients = ConvertTo-CgsList $email_to
    if ($recipients.Count -gt 0) {
        $subject = if ($email_subject) { $email_subject } else { "cgs_ai file scan complete - $stamp" }
        $body = @"
The cgs_ai file scan has completed.

Roots scanned : $($roots.Count)
Keywords      : $($keywords -join ', ')
Scan output   : $scanOutput
Report        : $reportTarget
"@
        & (Join-Path $script:PsFolder 'sendEmail.ps1') `
            -To ($recipients -join ';') -From $email_from -Subject $subject -Body $body
        if ($LASTEXITCODE -eq 0) { $steps += 'sendEmail' }
        else { Write-CgsError "notification email failed (scan output is still valid)" }
    } else {
        Write-CgsInfo 'no email_to supplied; skipping notification'
    }

    Write-CgsInfo ("pipeline complete; steps: {0}" -f ($steps -join ', '))
    return 0
}
try { exit (Invoke-Main) }
catch { Write-CgsError ("unhandled error: {0}" -f $_.Exception.Message); exit 3 }
