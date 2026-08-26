<#
=====================================================================
  Program Name  : scanFileSystem.ps1
  Author        : Manuel Figallo
  Purpose       : Scan one or more directory roots for keyword matches in
                  text files and emit one row per MATCH, with surrounding
                  context lines and token extraction. Optionally parses
                  structured performance metrics from log files via an
                  opt-in metric profile, which adds an Excel sheet.
  Version       : 1.0beta
  Created       : 2026-08-20
  Last Modified : 2026-08-26

  Dependencies:
    CSV output needs nothing beyond Windows PowerShell 5.1. XLSX output
    needs the ImportExcel module (Install-Module ImportExcel -Scope
    CurrentUser); without it the scan falls back to CSV with a warning.

  Description:
    PowerShell twin of src/py/scanFileSystem.py. Same function name, same
    parameter names, same output columns, same exit codes.

    GRAIN: one row per KEYWORD MATCH, not one row per file.

    Never prompts: no parameter is Mandatory, because a Mandatory parameter
    would make PowerShell prompt and hang an unattended run.

  Input Parameters (required first):
    -input_folder_root      (REQUIRED) root path(s); array or ';'-string
    -extract_keyword        (REQUIRED) keyword(s); array or ';'-string
    -output_file_path       (optional) .csv/.xlsx, a directory, or omitted
    -file_extensions        (default log,txt,sas)
    -include_subdirectories (default true; accepts true/false/1/0/yes/no)
    -folder_exclusion_list  (default empty -- nothing excluded)
    -file_exclusion_list    (default empty)
    -lines_above            (default 5)
    -lines_below            (default 5)
    -nth_token_after        (default 1)
    -nth_token_before       (default 1)
    -numeric_token_after    (default 1)
    -date_from / -date_to   (optional YYYY-MM-DD, inclusive)
    -date_field             (default modified; created|modified|accessed)
    -metric_profile         (default none; none|sas_log) -- when active,
                              EXCEL output is produced and announced.

  Exit codes: 0 = success, 2 = config error, 3 = I/O error.

  Change Log:
    v1.0beta - Regrained to one row per match; configurable context window
               and token positions; metric_profile announces Excel output.
=====================================================================
#>

[CmdletBinding()]
param(
    [string[]] $input_folder_root      = @(),
    [string[]] $extract_keyword        = @(),
    [string]   $output_file_path       = '',
    [string[]] $file_extensions        = @('log','txt','sas'),
    # [string] not [bool]: -File mode passes every argument as a string.
    [string]   $include_subdirectories = 'true',
    [string[]] $folder_exclusion_list  = @(),
    [string[]] $file_exclusion_list    = @(),
    [int]      $lines_above            = 5,
    [int]      $lines_below            = 5,
    [int]      $nth_token_after        = 1,
    [int]      $nth_token_before       = 1,
    [int]      $numeric_token_after    = 1,
    [string]   $date_from              = '',
    [string]   $date_to                = '',
    [string]   $date_field             = 'modified',
    [string]   $metric_profile         = 'none'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\cgsUtils.ps1"

$script:Version         = '1.0beta'
$script:ValidDateFields = @('created','modified','accessed')
$script:ExitOk = 0; $script:ExitConfigError = 2; $script:ExitIoError = 3
$script:MatchSheet = 'Matches'; $script:MetricSheet = 'Metrics'

# The default output columns, in order. One row per keyword match.
$script:MatchColumns = @(
    'SourceDir','FileName','Line','LinesAbove','LinesBelow','FullPath',
    'LineNumber','Keyword','ExtractedString','NthTokenAfter','NthTokenBefore',
    'NumericTokenAfter','LastToken','FirstToken','FileTimestamp','extension',
    'file_size_bytes','created_time','modified_time','accessed_time','scanned_at'
)
$script:MetricColumns = @('FullPath','ProgramName','StepIndex','StepLabel',
                          'RealTimeSec','CpuTimeSec')

# Metric profile registry -- add an entry to support a new log format.
$script:MetricProfiles = @{
    'none'    = @{ Active = $false }
    'sas_log' = @{
        Active      = $true
        StepPattern = '^NOTE:\s+(?<label>.+?)\s+used\s+\(Total process time\)'
        Metrics     = [ordered]@{
            RealTimeSec = '^\s*real time\s+(?<value>[0-9:.]+)'
            # 'user cpu time' matched too, so FULLSTIMER logs are not missed.
            CpuTimeSec  = '^\s*(?:user\s+)?cpu time\s+(?<value>[0-9:.]+)'
        }
        Lookahead   = 10
    }
}

function Get-DurationSeconds {
    <# .SYNOPSIS Parse "0.05", "1.20 seconds", "1:03.05" or "1:00:30.00".
       .OUTPUTS  [double] seconds, or $null when unparseable. #>
    param([string] $Raw)
    if ([string]::IsNullOrWhiteSpace($Raw)) { return $null }
    $text = $Raw.ToLowerInvariant().Replace('seconds','').Replace('second','').Trim()
    if (-not $text) { return $null }
    $ic = [System.Globalization.CultureInfo]::InvariantCulture
    $m = [regex]::Match($text, '^(\d+):(\d{1,2}):(\d{1,2}(?:\.\d+)?)$')
    if ($m.Success) { return [double]$m.Groups[1].Value*3600 + [double]$m.Groups[2].Value*60 + [double]::Parse($m.Groups[3].Value,$ic) }
    $m = [regex]::Match($text, '^(\d+):(\d{1,2}(?:\.\d+)?)$')
    if ($m.Success) { return [double]$m.Groups[1].Value*60 + [double]::Parse($m.Groups[2].Value,$ic) }
    $m = [regex]::Match($text, '^(\d+(?:\.\d+)?)$')
    if ($m.Success) { return [double]::Parse($m.Groups[1].Value,$ic) }
    return $null
}

function Test-NumericToken {
    <# .SYNOPSIS True when a token looks numeric (allows $ , . % and a sign). #>
    param([string] $Token)
    return [bool]([regex]::IsMatch($Token, '^[+-]?\$?\d[\d,]*\.?\d*%?$'))
}

function Get-TokenExtract {
    <# .SYNOPSIS Pull the requested tokens out of a matched line.
       .PARAMETER Line          The full matched line.
       .PARAMETER Keyword       Keyword that matched (case-insensitive).
       .PARAMETER NthAfter      Which token AFTER the keyword (1-based).
       .PARAMETER NthBefore     Which token BEFORE, counting backwards.
       .PARAMETER NumericAfter  Which NUMERIC token after the keyword.
       .OUTPUTS  [hashtable] with the five token fields; '' when absent. #>
    param([string] $Line, [string] $Keyword, [int] $NthAfter,
          [int] $NthBefore, [int] $NumericAfter)

    $tokens = @($Line -split '\s+' | Where-Object { $_ })
    $needleWords = @($Keyword.ToLowerInvariant() -split '\s+' | Where-Object { $_ })
    $firstWord = if ($needleWords.Count) { $needleWords[0] } else { $Keyword.ToLowerInvariant() }

    $anchor = -1
    for ($i = 0; $i -lt $tokens.Count; $i++) {
        if ($tokens[$i].ToLowerInvariant().Contains($firstWord)) { $anchor = $i; break }
    }

    $after = @(); $before = @()
    if ($anchor -ge 0) {
        $start = $anchor + $needleWords.Count
        if ($start -lt $tokens.Count) { $after = @($tokens[$start..($tokens.Count-1)]) }
        if ($anchor -gt 0) { $before = @($tokens[0..($anchor-1)]) }
    }
    $numerics = @($after | Where-Object { Test-NumericToken $_ })

    $pick = { param($seq, $n) if ($n -gt 0 -and $n -le $seq.Count) { $seq[$n-1] } else { '' } }
    return @{
        NthTokenAfter     = (& $pick $after $NthAfter)
        NthTokenBefore    = $(if ($NthBefore -gt 0 -and $NthBefore -le $before.Count) { $before[$before.Count - $NthBefore] } else { '' })
        NumericTokenAfter = (& $pick $numerics $NumericAfter)
        FirstToken        = $(if ($tokens.Count) { $tokens[0] } else { '' })
        LastToken         = $(if ($tokens.Count) { $tokens[$tokens.Count-1] } else { '' })
    }
}

function Test-FolderExcluded {
    <# .SYNOPSIS True when an ancestor segment or full-path prefix matches.
       .DESCRIPTION Segment matching is exact, so excluding "Old" keeps "Older". #>
    param([string] $FullPath, [string[]] $Exclusions)
    if (-not $Exclusions -or $Exclusions.Count -eq 0) { return $false }
    $parent = [System.IO.Path]::GetDirectoryName($FullPath)
    if (-not $parent) { return $false }
    $segments = @($parent -split '[\\/]' | Where-Object { $_ } | ForEach-Object { $_.ToLowerInvariant() })
    $normalized = ($FullPath -replace '\\','/').TrimEnd('/').ToLowerInvariant()
    foreach ($raw in $Exclusions) {
        $token = $raw.Trim(); if (-not $token) { continue }
        if ($segments -contains (($token -replace '\\','/').TrimEnd('/').ToLowerInvariant())) { return $true }
        $prefix = ($token -replace '\\','/').TrimEnd('/').ToLowerInvariant()
        if ($prefix -and ($normalized -eq $prefix -or $normalized.StartsWith($prefix + '/'))) { return $true }
    }
    return $false
}

function Read-TextLines {
    <# .SYNOPSIS Read a file as lines: UTF-8, then latin-1, then replacement. #>
    param([string] $Path)
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $strict = New-Object System.Text.UTF8Encoding($false, $true)
    $text = $null
    try { $text = $strict.GetString($bytes) }
    catch {
        try   { $text = [System.Text.Encoding]::GetEncoding('iso-8859-1').GetString($bytes) }
        catch { $text = [System.Text.Encoding]::UTF8.GetString($bytes) }
    }
    if ($null -eq $text) { return @() }
    $lines = $text -split "`r`n|`n|`r"
    # Match Python's splitlines(): a trailing newline yields no final empty item.
    if ($lines.Count -gt 0 -and $lines[-1] -eq '') { $lines = $lines[0..($lines.Count-2)] }
    return , $lines
}

function Get-MetricRows {
    <# .SYNOPSIS Extract structured metric rows from a file's lines.
       .OUTPUTS  Array of PSCustomObject shaped by $script:MetricColumns. #>
    param([hashtable] $ProfileDef, [string[]] $Lines,
          [string] $FullPath, [string] $ProgramName)
    $rows = New-Object System.Collections.Generic.List[object]
    if (-not $ProfileDef.Active) { return , $rows.ToArray() }
    for ($i = 0; $i -lt $Lines.Count; $i++) {
        $header = [regex]::Match($Lines[$i], $ProfileDef.StepPattern,
                  [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if (-not $header.Success) { continue }
        $row = [ordered]@{
            FullPath    = $FullPath
            ProgramName = $ProgramName
            StepIndex   = $rows.Count + 1
            StepLabel   = $header.Groups['label'].Value.Trim()
        }
        $end = [Math]::Min($Lines.Count - 1, $i + $ProfileDef.Lookahead)
        foreach ($metric in $ProfileDef.Metrics.Keys) {
            $row[$metric] = $null
            if ($i + 1 -le $end) {
                foreach ($candidate in $Lines[($i+1)..$end]) {
                    $hit = [regex]::Match($candidate, $ProfileDef.Metrics[$metric],
                           [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                    if ($hit.Success) { $row[$metric] = Get-DurationSeconds $hit.Groups['value'].Value; break }
                }
            }
        }
        $rows.Add([PSCustomObject]$row)
    }
    return , $rows.ToArray()
}

function Write-MatchExcel {
    <# .SYNOPSIS Write matches (and metrics) to an .xlsx via ImportExcel.
       .OUTPUTS  [string] path written, or $null when ImportExcel is absent. #>
    param([object[]] $MatchRows, [object[]] $MetricRows, [string] $Target)
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) { return $null }
    Import-Module ImportExcel -ErrorAction Stop
    if (Test-Path -LiteralPath $Target) { Remove-Item -LiteralPath $Target -Force }
    $MatchRows | Select-Object $script:MatchColumns |
        Export-Excel -Path $Target -WorksheetName $script:MatchSheet -AutoSize
    if ($MetricRows.Count -gt 0) {
        $MetricRows | Select-Object $script:MetricColumns |
            Export-Excel -Path $Target -WorksheetName $script:MetricSheet -AutoSize
    }
    return $Target
}

function Invoke-Main {
    <# .SYNOPSIS Validate, crawl, extract and write. .OUTPUTS [int] exit code. #>
    $roots      = ConvertTo-CgsList $input_folder_root
    $keywords   = ConvertTo-CgsList $extract_keyword
    $extensions = @(ConvertTo-CgsList $file_extensions | ForEach-Object { $_.TrimStart('.').ToLowerInvariant() })
    $folderEx   = ConvertTo-CgsList $folder_exclusion_list
    $fileEx     = ConvertTo-CgsList $file_exclusion_list

    if ($roots.Count -eq 0) {
        Write-CgsError "required parameter 'input_folder_root' is missing or empty; pass -input_folder_root"
        return $script:ExitConfigError
    }
    if ($keywords.Count -eq 0) {
        Write-CgsError "required parameter 'extract_keyword' is missing or empty; pass -extract_keyword (no keywords means no matches)"
        return $script:ExitConfigError
    }
    if (-not $script:MetricProfiles.Contains($metric_profile)) {
        Write-CgsError ("unknown metric_profile '$metric_profile'; expected one of: " + (($script:MetricProfiles.Keys | Sort-Object) -join ', '))
        return $script:ExitConfigError
    }
    if ($script:ValidDateFields -notcontains $date_field) {
        Write-CgsError ("unknown date_field '$date_field'; expected one of: " + ($script:ValidDateFields -join ', '))
        return $script:ExitConfigError
    }
    $recurse = ConvertTo-CgsBool $include_subdirectories
    if ($null -eq $recurse) {
        Write-CgsError "unknown include_subdirectories '$include_subdirectories'; expected true/false (or 1/0, yes/no)"
        return $script:ExitConfigError
    }

    $dateLow = $null; $dateHigh = $null
    $ic = [System.Globalization.CultureInfo]::InvariantCulture
    $parsed = [datetime]::MinValue
    if ($date_from) {
        if (-not [datetime]::TryParse($date_from, $ic, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
            Write-CgsError "unparseable date '$date_from'; expected YYYY-MM-DD or ISO datetime"
            return $script:ExitConfigError
        }
        $dateLow = $parsed
    }
    if ($date_to) {
        if (-not [datetime]::TryParse($date_to, $ic, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
            Write-CgsError "unparseable date '$date_to'; expected YYYY-MM-DD or ISO datetime"
            return $script:ExitConfigError
        }
        $dateHigh = if ($date_to.Trim().Length -eq 10) { $parsed.Date.AddDays(1).AddTicks(-1) } else { $parsed }
    }
    if ($dateLow -and $dateHigh -and $dateLow -gt $dateHigh) {
        Write-CgsError "date_from ($date_from) is after date_to ($date_to)"
        return $script:ExitConfigError
    }

    $profileDef = $script:MetricProfiles[$metric_profile]
    Write-CgsInfo ("scanFileSystem {0} (PowerShell) starting; profile={1}; roots={2}; keywords={3}" -f `
                   $script:Version, $metric_profile, $roots.Count, $keywords.Count)

    # Resolve the output target BEFORE crawling so a bad path fails fast.
    $target = $output_file_path.Trim()
    if (-not $target) {
        $target = "scan_$(Get-CgsTimestampSuffix).csv"
        Write-CgsInfo "output_file_path not supplied; writing $target"
    } else {
        $suffix = [System.IO.Path]::GetExtension($target).ToLowerInvariant()
        $isDir = (Test-Path -LiteralPath $target -PathType Container) -or
                 $target.EndsWith('/') -or $target.EndsWith('\') -or
                 ($suffix -ne '.csv' -and $suffix -ne '.xlsx')
        if ($isDir) {
            if (-not (Test-Path -LiteralPath $target)) { [void](New-Item -ItemType Directory -Path $target -Force) }
            $target = Join-Path $target "scan_$(Get-CgsTimestampSuffix).csv"
            Write-CgsInfo "output path is a directory; writing $target"
        }
    }

    # An active metric profile REQUIRES Excel (the metrics go on a second
    # sheet). Tell the user plainly rather than switching silently.
    if ($profileDef.Active -and -not $target.ToLowerInvariant().EndsWith('.xlsx')) {
        $requested = [System.IO.Path]::GetExtension($target)
        $target = [System.IO.Path]::ChangeExtension($target, '.xlsx')
        Write-CgsInfo "metric_profile='$metric_profile' is active, so EXCEL output is produced: $target"
        Write-CgsInfo "(requested $requested; the extra '$($script:MetricSheet)' sheet cannot be written to CSV)"
    }

    $scannedAt = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
    $matchRows  = New-Object System.Collections.Generic.List[object]
    $metricRows = New-Object System.Collections.Generic.List[object]
    $reachable  = 0
    $candidates = New-Object System.Collections.Generic.List[string]

    foreach ($rawRoot in $roots) {
        $root = $rawRoot.Trim()
        if (-not (Test-Path -LiteralPath $root -PathType Container)) {
            Write-CgsWarn "root not found or not a directory, skipping: $root"; continue
        }
        $reachable++
        $items = Get-ChildItem -LiteralPath $root -File -Recurse:$recurse -Force -ErrorAction SilentlyContinue
        foreach ($item in $items) { $candidates.Add($item.FullName) }
    }
    Write-CgsInfo ("discovered {0} file(s) across {1} reachable root(s)" -f $candidates.Count, $reachable)

    [string[]] $ordered = $candidates.ToArray()
    [Array]::Sort($ordered, [StringComparer]::Ordinal)

    foreach ($path in $ordered) {
        $extension = [System.IO.Path]::GetExtension($path).TrimStart('.').ToLowerInvariant()
        if ($extensions.Count -gt 0 -and $extensions -notcontains $extension) { continue }
        if (Test-FolderExcluded -FullPath $path -Exclusions $folderEx) { continue }

        try { $info = Get-Item -LiteralPath $path -Force -ErrorAction Stop }
        catch { Write-CgsWarn "cannot stat ${path}: $($_.Exception.Message)"; continue }

        $times = @{ created = $info.CreationTime; modified = $info.LastWriteTime; accessed = $info.LastAccessTime }
        if ($dateLow  -and $times[$date_field] -lt $dateLow)  { continue }
        if ($dateHigh -and $times[$date_field] -gt $dateHigh) { continue }

        try { $lines = Read-TextLines -Path $path }
        catch { Write-CgsWarn "cannot read ${path}: $($_.Exception.Message)"; continue }

        $programName = [System.IO.Path]::GetFileNameWithoutExtension($path)
        foreach ($raw in $fileEx) {
            $token = $raw.Trim()
            if ($token -and $programName.ToLowerInvariant().StartsWith($token.ToLowerInvariant())) {
                $programName = $programName.Substring($token.Length)
            }
        }

        if ($profileDef.Active) {
            foreach ($metricRow in (Get-MetricRows -ProfileDef $profileDef -Lines $lines -FullPath $path -ProgramName $programName)) {
                $metricRows.Add($metricRow)
            }
        }

        for ($i = 0; $i -lt $lines.Count; $i++) {
            $line = $lines[$i]
            $lowered = $line.ToLowerInvariant()
            foreach ($keyword in $keywords) {
                if (-not $lowered.Contains($keyword.ToLowerInvariant())) { continue }
                $lo = [Math]::Max(0, $i - $lines_above)
                $above = if ($i -gt 0 -and $lo -le ($i-1)) { ($lines[$lo..($i-1)] -join "`n") } else { '' }
                $hi = [Math]::Min($lines.Count - 1, $i + $lines_below)
                $below = if (($i+1) -le $hi) { ($lines[($i+1)..$hi] -join "`n") } else { '' }
                $tokens = Get-TokenExtract -Line $line -Keyword $keyword `
                            -NthAfter $nth_token_after -NthBefore $nth_token_before `
                            -NumericAfter $numeric_token_after

                $matchRows.Add([PSCustomObject][ordered]@{
                    SourceDir         = [System.IO.Path]::GetDirectoryName($path)
                    FileName          = [System.IO.Path]::GetFileName($path)
                    Line              = $line
                    LinesAbove        = $above
                    LinesBelow        = $below
                    FullPath          = $path
                    LineNumber        = $i + 1
                    Keyword           = $keyword
                    ExtractedString   = $line
                    NthTokenAfter     = $tokens.NthTokenAfter
                    NthTokenBefore    = $tokens.NthTokenBefore
                    NumericTokenAfter = $tokens.NumericTokenAfter
                    LastToken         = $tokens.LastToken
                    FirstToken        = $tokens.FirstToken
                    FileTimestamp     = $info.LastWriteTime.ToString('yyyy-MM-ddTHH:mm:ss')
                    extension         = $extension
                    file_size_bytes   = $info.Length
                    created_time      = $info.CreationTime.ToString('yyyy-MM-ddTHH:mm:ss')
                    modified_time     = $info.LastWriteTime.ToString('yyyy-MM-ddTHH:mm:ss')
                    accessed_time     = $info.LastAccessTime.ToString('yyyy-MM-ddTHH:mm:ss')
                    scanned_at        = $scannedAt
                })
            }
        }
    }

    if ($reachable -eq 0) {
        Write-CgsError 'none of the supplied input_folder_root path(s) are reachable'
        return $script:ExitIoError
    }

    if ($target.ToLowerInvariant().EndsWith('.xlsx')) {
        $written = Write-MatchExcel -MatchRows $matchRows.ToArray() -MetricRows $metricRows.ToArray() -Target $target
        if ($null -eq $written) {
            Write-CgsWarn 'no Excel engine (ImportExcel module) available; falling back to CSV'
            $target = [System.IO.Path]::ChangeExtension($target, '.csv')
            [void](Write-CgsCsv -Rows $matchRows.ToArray() -Columns $script:MatchColumns -Target $target)
            Write-CgsInfo ("wrote {0} match row(s) to {1}" -f $matchRows.Count, $target)
            if ($metricRows.Count -gt 0) {
                $companion = Join-Path ([System.IO.Path]::GetDirectoryName($target)) `
                    ("{0}_Metrics.csv" -f [System.IO.Path]::GetFileNameWithoutExtension($target))
                [void](Write-CgsCsv -Rows $metricRows.ToArray() -Columns $script:MetricColumns -Target $companion)
                Write-CgsInfo ("wrote {0} metric row(s) to companion {1}" -f $metricRows.Count, $companion)
            }
        } else {
            Write-CgsInfo ("wrote {0} match row(s) to sheet '{1}' in {2}" -f $matchRows.Count, $script:MatchSheet, $written)
            if ($metricRows.Count -gt 0) {
                Write-CgsInfo ("wrote {0} metric row(s) to sheet '{1}'" -f $metricRows.Count, $script:MetricSheet)
            }
        }
    } else {
        [void](Write-CgsCsv -Rows $matchRows.ToArray() -Columns $script:MatchColumns -Target $target)
        Write-CgsInfo ("wrote {0} match row(s) to {1}" -f $matchRows.Count, $target)
    }

    Write-CgsInfo ("done; {0} match row(s), {1} metric row(s)" -f $matchRows.Count, $metricRows.Count)
    return $script:ExitOk
}

# A trailing exception would exit 1, indistinguishable from "PowerShell could
# not start the script". Catch everything and report a controlled code.
try {
    exit (Invoke-Main)
} catch {
    Write-CgsError ("unhandled error: {0}" -f $_.Exception.Message)
    if ($_.ScriptStackTrace) { Write-CgsError ("at: {0}" -f ($_.ScriptStackTrace -split "`n")[0]) }
    exit $script:ExitIoError
}
