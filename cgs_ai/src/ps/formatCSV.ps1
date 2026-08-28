<#
=====================================================================
  Program Name  : formatCSV.ps1
  Author        : Manuel Figallo
  Purpose       : Convert a CSV into a styled Excel workbook with a SAS
                  ODS-style look and feel.
  Version       : 1.0beta
  Created       : 2026-08-26
  Last Modified : 2026-08-26

  Dependencies:
    ImportExcel module (Install-Module ImportExcel -Scope CurrentUser).
    Checked at runtime; a clear message is given when it is absent.

  Description:
    PowerShell twin of src/py/formatCSV.py. The default "corporate" style is
    a navy title banner, blue header row and zebra striping.

  Input Parameters (required first):
    -InputCsvPath (REQUIRED)  -OutputExcelPath (REQUIRED)
    -FormatType (default corporate; corporate|corporatev2|plain|minimal)
    -SheetName (default Report)  -Title (defaults to the CSV name)
  Exit codes: 0 = success, 2 = config error, 3 = I/O error.
=====================================================================
#>
[CmdletBinding()]
param(
    [string] $InputCsvPath    = '',
    [string] $OutputExcelPath = '',
    [string] $FormatType      = 'corporate',
    [string] $SheetName       = 'Report',
    [string] $Title           = ''
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\cgsUtils.ps1"

# Palette per format type -- matches FORMAT_STYLES in the Python twin.
$script:FormatStyles = @{
    'corporate'   = @{ Banner='#1F3864'; Header='#2E75B6'; Stripe='#DCE6F1'; HeaderFont='White'; BannerFont='White' }
    'corporatev2' = @{ Banner='#1F3864'; Header='#2E75B6'; Stripe='#DCE6F1'; HeaderFont='White'; BannerFont='White' }
    'plain'       = @{ Banner='#FFFFFF'; Header='#D9D9D9'; Stripe='#FFFFFF'; HeaderFont='Black'; BannerFont='Black' }
    'minimal'     = @{ Banner='#FFFFFF'; Header='#FFFFFF'; Stripe='#FFFFFF'; HeaderFont='Black'; BannerFont='Black' }
}

# Format types that get zebra striping -- matches STRIPED_FORMATS in Python.
$script:StripedFormats = @('corporate', 'corporatev2')

function Invoke-Main {
    <# .SYNOPSIS Style the CSV into a workbook. .OUTPUTS [int] exit code. #>
    if (-not $InputCsvPath)    { Write-CgsError "required parameter 'InputCsvPath' is missing or empty"; return 2 }
    if (-not $OutputExcelPath) { Write-CgsError "required parameter 'OutputExcelPath' is missing or empty"; return 2 }
    if (-not $script:FormatStyles.Contains($FormatType)) {
        Write-CgsError ("unknown FormatType '$FormatType'; expected one of: " + (($script:FormatStyles.Keys | Sort-Object) -join ', '))
        return 2
    }
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-CgsError "formatCSV requires the ImportExcel module. Install it with: Install-Module ImportExcel -Scope CurrentUser"
        return 3
    }
    Import-Module ImportExcel -ErrorAction Stop

    $rows = @(Import-Csv -LiteralPath $InputCsvPath)
    $style = $script:FormatStyles[$FormatType]
    $banner = if ($Title) { $Title } else { [System.IO.Path]::GetFileNameWithoutExtension($InputCsvPath) }
    if (Test-Path -LiteralPath $OutputExcelPath) { Remove-Item -LiteralPath $OutputExcelPath -Force }

    # StartRow 2 leaves row 1 for the banner, matching the Python twin -- so
    # copyExcelSheet2CSV must be called with -HeaderRow 2 to round-trip.
    $excel = $rows | Export-Excel -Path $OutputExcelPath -WorksheetName $SheetName `
                -AutoSize -FreezeTopRowFirstColumn:$false -StartRow 2 -PassThru
    $sheet = $excel.Workbook.Worksheets[$SheetName]
    $columnCount = if ($rows.Count) { @($rows[0].PSObject.Properties).Count } else { 1 }

    Set-ExcelRange -Worksheet $sheet -Range "A1:$([char](64+$columnCount))1" -Merge `
        -BackgroundColor $style.Banner -FontColor $style.BannerFont -Bold -Value $banner
    Set-ExcelRange -Worksheet $sheet -Range "A2:$([char](64+$columnCount))2" `
        -BackgroundColor $style.Header -FontColor $style.HeaderFont -Bold
    if ($script:StripedFormats -contains $FormatType) {
        for ($r = 3; $r -lt (3 + $rows.Count); $r++) {
            if ($r % 2 -eq 1) {
                Set-ExcelRange -Worksheet $sheet -Range "A${r}:$([char](64+$columnCount))${r}" -BackgroundColor $style.Stripe
            }
        }
    }
    $sheet.View.FreezePanes(3, 1)
    Close-ExcelPackage $excel

    Write-CgsInfo ("formatted {0} row(s) x {1} column(s) [{2}] -> {3}" -f $rows.Count, $columnCount, $FormatType, $OutputExcelPath)
    return 0
}
try { exit (Invoke-Main) }
catch { Write-CgsError ("unhandled error: {0}" -f $_.Exception.Message); exit 3 }
