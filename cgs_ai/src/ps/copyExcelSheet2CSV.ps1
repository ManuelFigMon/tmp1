<#
=====================================================================
  Program Name  : copyExcelSheet2CSV.ps1
  Author        : Manuel Figallo
  Purpose       : Export one Excel worksheet to CSV, refusing to proceed
                  when the sheet is not shaped for flat output.
  Version       : 1.0beta
  Created       : 2026-08-26
  Last Modified : 2026-08-26

  Dependencies: ImportExcel module.

  Description:
    PowerShell twin of src/py/copyExcelSheet2CSV.py. Validates BEFORE
    writing and stops on the first problem -- sheet missing, sheet empty,
    blank or duplicate header names -- rather than emitting a malformed CSV
    that fails downstream.

  Input Parameters (required first):
    -InputExcelPath (REQUIRED)  -SheetName (REQUIRED)
    -OutputCsvPath (REQUIRED)   -HeaderRow (default 1; use 2 for formatCSV output)
  Exit codes: 0 = success, 2 = config/validation error, 3 = I/O error.
=====================================================================
#>
[CmdletBinding()]
param(
    [string] $InputExcelPath = '',
    [string] $SheetName      = '',
    [string] $OutputCsvPath  = '',
    [int]    $HeaderRow      = 1
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\cgsUtils.ps1"

function Invoke-Main {
    <# .SYNOPSIS Validate then export. .OUTPUTS [int] exit code. #>
    foreach ($pair in @(@('InputExcelPath',$InputExcelPath), @('SheetName',$SheetName), @('OutputCsvPath',$OutputCsvPath))) {
        if (-not $pair[1]) { Write-CgsError "required parameter '$($pair[0])' is missing or empty"; return 2 }
    }
    if ($HeaderRow -lt 1) { Write-CgsError 'HeaderRow must be 1 or greater'; return 2 }
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-CgsError "copyExcelSheet2CSV requires the ImportExcel module. Install it with: Install-Module ImportExcel -Scope CurrentUser"
        return 3
    }
    Import-Module ImportExcel -ErrorAction Stop

    $sheets = @(Get-ExcelSheetInfo -Path $InputExcelPath | ForEach-Object { $_.Name })
    if ($sheets -notcontains $SheetName) {
        Write-CgsError ("worksheet '$SheetName' not found in $InputExcelPath. Available sheets: " + ($sheets -join ', '))
        return 2
    }
    $rows = @(Import-Excel -Path $InputExcelPath -WorksheetName $SheetName -StartRow $HeaderRow)
    if ($rows.Count -eq 0) {
        Write-CgsError "worksheet '$SheetName' is empty; nothing to export."
        return 2
    }
    $header = @($rows[0].PSObject.Properties | ForEach-Object { $_.Name })
    $blank = @($header | Where-Object { -not $_ -or $_.Trim() -eq '' -or $_ -match '^[A-Z]+\d*$' -and $false })
    if (@($header | Where-Object { -not $_ -or $_.Trim() -eq '' }).Count -gt 0) {
        Write-CgsError "worksheet '$SheetName' has blank header name(s). Every column needs a name for CSV output."
        return 2
    }
    $duplicates = @($header | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })
    if ($duplicates.Count -gt 0) {
        Write-CgsError ("worksheet '$SheetName' has duplicate header name(s): " + ($duplicates -join ', ') + ". Column names must be unique.")
        return 2
    }

    $parent = Split-Path -Parent $OutputCsvPath
    if ($parent -and -not (Test-Path -LiteralPath $parent)) { [void](New-Item -ItemType Directory -Path $parent -Force) }
    [void](Write-CgsCsv -Rows $rows -Columns $header -Target $OutputCsvPath)
    Write-CgsInfo ("exported '{0}': {1} row(s) x {2} column(s) -> {3}" -f $SheetName, $rows.Count, $header.Count, $OutputCsvPath)
    return 0
}
try { exit (Invoke-Main) }
catch { Write-CgsError ("unhandled error: {0}" -f $_.Exception.Message); exit 3 }
