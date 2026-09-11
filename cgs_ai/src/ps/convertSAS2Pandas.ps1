<#
=====================================================================
  Program Name  : convertSAS2Pandas.ps1
  Author        : Manuel Figallo
  Purpose       : Convert a sas7bdat data set to a pandas-readable file.
  Version       : 1.0beta
  Created       : 2026-08-26
  Last Modified : 2026-08-26

  Dependencies:
    Python with pandas. PowerShell has no native sas7bdat reader, so this
    function DELEGATES to the Python twin (src/py/convertSAS2Pandas.py)
    rather than reimplementing a proprietary binary format badly.
    This is the one function where the two languages are not independent
    implementations; the parameter names and behaviour remain identical.

  Input Parameters (required first):
    -InputSas7bdatPath (REQUIRED)  -OutputPath (REQUIRED)
    -PythonExe (defaults to 'python' on PATH)
  Exit codes: 0 = success, 2 = config error, 3 = conversion failure.
=====================================================================
#>
[CmdletBinding()]
param(
    [string] $InputSas7bdatPath = '',
    [string] $OutputPath        = '',
    [string] $PythonExe         = 'python',
    [string] $Encoding          = 'latin-1'
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\cgsUtils.ps1"

function Invoke-Main {
    <# .SYNOPSIS Delegate to the Python twin. .OUTPUTS [int] exit code. #>
    if (-not $InputSas7bdatPath) { Write-CgsError "required parameter 'InputSas7bdatPath' is missing or empty"; return 2 }
    if (-not $OutputPath)        { Write-CgsError "required parameter 'OutputPath' is missing or empty"; return 2 }
    if (-not (Test-Path -LiteralPath $InputSas7bdatPath -PathType Leaf)) {
        Write-CgsError "sas7bdat file not found: $InputSas7bdatPath"; return 3
    }
    $projectRoot = Get-ProjectRoot
    Write-CgsInfo "delegating to the Python twin (PowerShell has no native sas7bdat reader)"

    $snippet = @"
import sys
sys.path.insert(0, r'$projectRoot')
from src.py.convertSAS2Pandas import convertSAS2Pandas
convertSAS2Pandas(InputSas7bdatPath=r'$InputSas7bdatPath',
                  OutputPath=r'$OutputPath', Encoding='$Encoding')
"@
    $snippet | & $PythonExe -
    if ($LASTEXITCODE -ne 0) {
        Write-CgsError "the Python twin failed with exit code $LASTEXITCODE (is pandas installed?)"
        return 3
    }
    Write-CgsInfo "converted -> $OutputPath"
    return 0
}
try { exit (Invoke-Main) }
catch { Write-CgsError ("unhandled error: {0}" -f $_.Exception.Message); exit 3 }
