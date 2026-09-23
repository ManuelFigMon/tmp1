<#
=====================================================================
  Program Name  : formatCSV.ps1
  Author        : Manuel Figallo
  Purpose       : Backward-compatible shim. formatCSV was renamed to
                  formatData when it learned to read .xlsx as well as
                  .csv; this forwards every call to the new script.
  Version       : 1.2beta
  Created       : 2026-08-26
  Last Modified : 2026-09-06

  Dependencies:
    None of its own -- it invokes src/ps/formatData.ps1.

  Description:
    The .ps1 files are copied to a share individually, so deleting this one
    would break any caller copied before the rename with "the term
    formatCSV.ps1 is not recognized" rather than anything useful. Both names
    run the same code.

  Input Parameters (required first):
    As formatData.ps1. -InputCsvPath is the original name for -InputPath and
    is still accepted here and there.
  Exit codes: passed through from formatData.ps1.
=====================================================================
#>
[CmdletBinding()]
param(
    [string] $InputPath       = '',
    [string] $OutputExcelPath = '',
    [string] $FormatType      = 'corporate',
    [string] $SheetName       = 'Report',
    [string] $Title           = '',
    [string] $InputSheet      = '',
    [int]    $HeaderRow       = 0,
    [string] $Writer          = 'auto',
    [string] $InputCsvPath    = ''
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

& "$PSScriptRoot\formatData.ps1" `
    -InputPath $InputPath -OutputExcelPath $OutputExcelPath `
    -FormatType $FormatType -SheetName $SheetName -Title $Title `
    -InputSheet $InputSheet -HeaderRow $HeaderRow -Writer $Writer `
    -InputCsvPath $InputCsvPath
exit $LASTEXITCODE
