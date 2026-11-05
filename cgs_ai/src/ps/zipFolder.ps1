<#
=====================================================================
  Program Name  : zipFolder.ps1
  Author        : Manuel Figallo
  Purpose       : Zip a folder plus a list of accompanying files.
  Version       : 1.0beta
  Created       : 2026-08-26
  Last Modified : 2026-08-26

  Dependencies: none (System.IO.Compression is built in).

  Description:
    PowerShell twin of src/py/zipFolder.py. The folder is stored under its
    own name inside the archive; accompanying files land at the archive
    root beside it. __pycache__/.git/.venv are skipped.

  Input Parameters (required first):
    -FolderToZip (REQUIRED)  -OutputZipPath (REQUIRED)  -AccompanyFiles
  Exit codes: 0 = success, 2 = config error, 3 = I/O error.
=====================================================================
#>
[CmdletBinding()]
param(
    [string]   $FolderToZip    = '',
    [string]   $OutputZipPath  = '',
    [string[]] $AccompanyFiles = @()
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\cgsUtils.ps1"
Add-Type -AssemblyName System.IO.Compression.FileSystem

$script:SkipDirs = @('__pycache__', '.git', '.pytest_cache', '.venv')

function Invoke-Main {
    <# .SYNOPSIS Build the archive. .OUTPUTS [int] exit code. #>
    if (-not $FolderToZip)   { Write-CgsError "required parameter 'FolderToZip' is missing or empty"; return 2 }
    if (-not $OutputZipPath) { Write-CgsError "required parameter 'OutputZipPath' is missing or empty"; return 2 }
    if (-not (Test-Path -LiteralPath $FolderToZip -PathType Container)) {
        Write-CgsError "FolderToZip is not a directory: $FolderToZip"; return 3
    }
    $parent = Split-Path -Parent $OutputZipPath
    if ($parent -and -not (Test-Path -LiteralPath $parent)) { [void](New-Item -ItemType Directory -Path $parent -Force) }
    if (Test-Path -LiteralPath $OutputZipPath) { Remove-Item -LiteralPath $OutputZipPath -Force }

    $source   = (Resolve-Path -LiteralPath $FolderToZip).Path
    $rootName = Split-Path -Leaf $source
    $count    = 0
    $archive  = [System.IO.Compression.ZipFile]::Open($OutputZipPath, 'Create')
    try {
        foreach ($file in (Get-ChildItem -LiteralPath $source -Recurse -File -Force)) {
            $skip = $false
            foreach ($bad in $script:SkipDirs) {
                if ($file.FullName -split '[\\/]' -contains $bad) { $skip = $true; break }
            }
            if ($skip) { continue }
            $relative = $file.FullName.Substring($source.Length).TrimStart('\','/')
            $arcName  = Join-Path $rootName $relative
            [void][System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, $file.FullName, ($arcName -replace '\\','/'))
            $count++
        }
        foreach ($extra in (ConvertTo-CgsList $AccompanyFiles)) {
            if (-not (Test-Path -LiteralPath $extra -PathType Leaf)) {
                Write-CgsWarn "accompanying file not found, skipped: $extra"; continue
            }
            [void][System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, (Resolve-Path -LiteralPath $extra).Path, (Split-Path -Leaf $extra))
            $count++
        }
    } finally { $archive.Dispose() }

    Write-CgsInfo "wrote $count file(s) to $OutputZipPath"
    return 0
}
try { exit (Invoke-Main) }
catch { Write-CgsError ("unhandled error: {0}" -f $_.Exception.Message); exit 3 }
