<#
=====================================================================
  Program Name  : cgsUtils.ps1
  Author        : Manuel Figallo
  Purpose       : Shared helpers for every cgs_ai PowerShell function --
                  .env loading, logging, list parsing and CSV writing.
  Version       : 1.0beta
  Created       : 2026-08-26
  Last Modified : 2026-08-26

  Dependencies:
    None. Windows PowerShell 5.1 or PowerShell 7+.

  Description:
    Dot-source this from any cgs_ai .ps1:  . "$PSScriptRoot\cgsUtils.ps1"
    Mirrors src/utils/{config,logger,helpers}.py so the two languages
    behave identically.
=====================================================================
#>

$script:CgsVersion = '1.0beta'

function Get-ProjectRoot {
    <# .SYNOPSIS Return the cgs_ai project root (folder holding __init__.py).
       .OUTPUTS  [string] full path. #>
    return (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
}

function Import-DotEnv {
    <# .SYNOPSIS Read the project .env into a hashtable.
       .PARAMETER EnvPath  Optional path; defaults to <projectRoot>\.env.
       .OUTPUTS  [hashtable] KEY -> VALUE. Real environment variables win. #>
    param([string] $EnvPath = '')
    if (-not $EnvPath) { $EnvPath = Join-Path (Get-ProjectRoot) '.env' }
    $values = @{}
    if (Test-Path -LiteralPath $EnvPath) {
        foreach ($line in (Get-Content -LiteralPath $EnvPath)) {
            $trimmed = $line.Trim()
            if (-not $trimmed -or $trimmed.StartsWith('#') -or ($trimmed -notmatch '=')) { continue }
            $key, $value = $trimmed -split '=', 2
            $key = $key.Trim(); $value = $value.Trim().Trim('"').Trim("'")
            if ($key) { $values[$key] = $value }
        }
    }
    foreach ($key in @($values.Keys)) {
        $fromEnv = [Environment]::GetEnvironmentVariable($key)
        if ($fromEnv) { $values[$key] = $fromEnv }
    }
    return $values
}

function Get-CgsConfig {
    <# .SYNOPSIS Fetch one configuration value from .env / environment.
       .PARAMETER Key       Variable name, e.g. ROOT_DATA.
       .PARAMETER Default   Returned when the key is absent.
       .PARAMETER Required  Throw instead of returning the default.
       .OUTPUTS  [string] #>
    param([string] $Key, [string] $Default = '', [switch] $Required)
    $config = Import-DotEnv
    $value = if ($config.ContainsKey($Key)) { $config[$Key] } else {
        [Environment]::GetEnvironmentVariable($Key) }
    if (-not $value) { $value = $Default }
    if ($Required -and -not $value) {
        throw "Required configuration '$Key' is not set. Add it to .env (see .env.example)."
    }
    return $value
}

function Write-CgsLog {
    <# .SYNOPSIS Write a timestamped line to stderr, matching the Python format.
       .PARAMETER Level   INFO, WARNING or ERROR.
       .PARAMETER Message Text to log. #>
    param([string] $Level, [string] $Message)
    $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    [Console]::Error.WriteLine(('{0} {1,-7} {2}' -f $stamp, $Level, $Message))
}
function Write-CgsInfo  { param([string] $Message) Write-CgsLog 'INFO'    $Message }
function Write-CgsWarn  { param([string] $Message) Write-CgsLog 'WARNING' $Message }
function Write-CgsError { param([string] $Message) Write-CgsLog 'ERROR'   $Message }

function ConvertTo-CgsList {
    <# .SYNOPSIS Normalize a list parameter to a string array.
       .PARAMETER Value  Null, a string ("a;b"), or an array.
       .OUTPUTS  [string[]] -- semicolons split, because that is how the SAS
                 wrappers pass lists on a command line. #>
    param([object] $Value)
    $out = New-Object System.Collections.Generic.List[string]
    if ($null -eq $Value) { return , $out.ToArray() }
    foreach ($item in @($Value)) {
        if ($null -eq $item) { continue }
        foreach ($part in ([string]$item -split ';')) {
            $trimmed = $part.Trim()
            if ($trimmed) { $out.Add($trimmed) }
        }
    }
    return , $out.ToArray()
}

function ConvertTo-CgsBool {
    <# .SYNOPSIS Parse a boolean-ish string.
       .DESCRIPTION In "powershell.exe -File" mode every argument arrives as a
                    string, so a [bool] parameter CANNOT be set. Script
                    parameters are [string] and pass through here.
       .OUTPUTS  [bool] or $null when unrecognized. #>
    param([string] $Value)
    if ($null -eq $Value) { return $null }
    switch ($Value.Trim().ToLowerInvariant()) {
        { $_ -in @('true','1','yes','y','$true')  } { return $true }
        { $_ -in @('false','0','no','n','$false') } { return $false }
        default { return $null }
    }
}

function Format-CgsCell {
    <# .SYNOPSIS Render a value the way Python's csv writer does.
       .DESCRIPTION Doubles keep a decimal place (2 -> "2.0") so CSVs produced
                    by the two languages match byte for byte. #>
    param([object] $Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [double] -or $Value -is [single] -or $Value -is [decimal]) {
        return ([double]$Value).ToString('0.0###############',
            [System.Globalization.CultureInfo]::InvariantCulture)
    }
    return [string]$Value
}

function ConvertTo-CgsCsvField {
    <# .SYNOPSIS Minimal CSV quoting matching Python's csv module defaults. #>
    param([object] $Value)
    if ($null -eq $Value) { return '""' }
    $text = Format-CgsCell $Value
    if ($text -match '[",\r\n]') { return '"' + $text.Replace('"','""') + '"' }
    return $text
}

function Write-CgsCsv {
    <# .SYNOPSIS Write rows to UTF-8 CSV with no BOM (matches Python output).
       .PARAMETER Rows     Array of PSCustomObject / hashtable.
       .PARAMETER Columns  Column order.
       .PARAMETER Target   Destination path. #>
    param([object[]] $Rows, [string[]] $Columns, [string] $Target)
    $parent = Split-Path -Parent $Target
    if ($parent -and -not (Test-Path -LiteralPath $parent)) {
        [void](New-Item -ItemType Directory -Path $parent -Force)
    }
    $encoding = New-Object System.Text.UTF8Encoding($false)
    $writer = New-Object System.IO.StreamWriter($Target, $false, $encoding)
    try {
        $writer.NewLine = "`r`n"
        $writer.WriteLine((($Columns | ForEach-Object { ConvertTo-CgsCsvField $_ }) -join ','))
        foreach ($row in $Rows) {
            $fields = foreach ($column in $Columns) {
                $value = if ($row -is [hashtable]) { $row[$column] }
                         elseif ($row.PSObject.Properties[$column]) { $row.$column }
                         else { '' }
                ConvertTo-CgsCsvField $value
            }
            $writer.WriteLine(($fields -join ','))
        }
    } finally { $writer.Dispose() }
    return $Target
}

function Get-CgsTimestampSuffix {
    <# .SYNOPSIS Current time as yyyyMMdd_HHmmss for generated filenames. #>
    return (Get-Date).ToString('yyyyMMdd_HHmmss')
}
