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
    -metric_profile         (default none; none|sas_log|access_db) -- when
                              active, EXCEL output is produced and announced,
                              with a second sheet shaped by the profile:
                                sas_log   - step timings from a SAS log.
                                access_db - one row per LIBNAME pointing at a
                                  Microsoft Access file (.accdb/.mdb), with
                                  the line it is defined on and the line
                                  numbers where the libref is used. Scan .sas
                                  files for this: in a .log the numbers are
                                  log lines, not program lines.
                              KNOWN LIMITATION (access_db):
                                libname db pcfiles path="&mdbPath";
                              builds the path from a macro variable, so no
                              'accdb'/'mdb' is visible in the statement and it
                              cannot be reported. Those are counted and a
                              WARNING names the count per file.

  Exit codes: 0 = success, 2 = config error, 3 = I/O error.

  Change Log:
    v1.0beta - Regrained to one row per match; configurable context window
               and token positions; metric_profile announces Excel output.
             - Profiles generalized: each entry names an extractor and carries
               its own columns, so a profile need not be a timing parser.
               sas_log is unchanged. Added access_db.
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
# Columns of the second sheet, per profile. Each profile carries its own set
# because the profiles answer different questions -- timings, or usage.
$script:SasLogColumns   = @('FullPath','ProgramName','StepIndex','StepLabel',
                            'RealTimeSec','CpuTimeSec')
$script:AccessDbColumns = @('FullPath','ProgramName','Libref','DatabaseFile',
                            'Keyword','DefinitionLine','UsageCount','UsageLines')
# Backward-compatible alias for anything still expecting the sas_log shape.
$script:MetricColumns   = $script:SasLogColumns

# Profile registry. Each entry names an EXTRACTOR and carries its own Columns,
# so a new profile is a registry entry plus one function -- the crawl, the
# writer and the parameter list need no change.
$script:MetricProfiles = @{
    'none'    = @{ Active = $false; Columns = @() }
    'sas_log' = @{
        Active      = $true
        Extractor   = 'sas_log'
        Columns     = $script:SasLogColumns
        StepPattern = '^NOTE:\s+(?<label>.+?)\s+used\s+\(Total process time\)'
        Metrics     = [ordered]@{
            RealTimeSec = '^\s*real time\s+(?<value>[0-9:.]+)'
            # 'user cpu time' matched too, so FULLSTIMER logs are not missed.
            CpuTimeSec  = '^\s*(?:user\s+)?cpu time\s+(?<value>[0-9:.]+)'
        }
        Lookahead   = 10
    }
    'access_db' = @{
        Active            = $true
        Extractor         = 'access_db'
        Columns           = $script:AccessDbColumns
        Keywords          = @('accdb','mdb')
        # A LIBNAME may wrap; stop looking for the ';' after this many lines so
        # a file with an unterminated statement cannot swallow the whole file.
        MaxStatementLines = 20
    }
}

# LIBNAME statement parsing, used by the access_db profile.
$script:LibnameStart    = '\blibname\s+(?<libref>[A-Za-z_]\w*)'
$script:QuotedPath      = '(?<quote>[''"])(?<path>.*?)\k<quote>'
$script:MacroReference  = '[&%]\w+'
# Engines that reach a Microsoft Access file. Used only to decide whether a
# macro-built path is worth warning about, so an ordinary SAS library whose
# path contains a macro variable does not raise a false alarm.
$script:AccessEngine    = '\b(?:access|pcfiles)\b'

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
    <# .SYNOPSIS Dispatch to the extractor named by a profile.
       .OUTPUTS  Array of PSCustomObject shaped by the profile's own Columns;
                 empty when the profile is inactive. #>
    param([hashtable] $ProfileDef, [string[]] $Lines,
          [string] $FullPath, [string] $ProgramName)
    if (-not $ProfileDef.Active) { return , @() }
    switch ($ProfileDef.Extractor) {
        'sas_log'   { return , (Get-SasLogMetricRows -ProfileDef $ProfileDef -Lines $Lines -FullPath $FullPath -ProgramName $ProgramName) }
        'access_db' { return , (Get-AccessDbRows     -ProfileDef $ProfileDef -Lines $Lines -FullPath $FullPath -ProgramName $ProgramName) }
    }
    Write-CgsWarn ("unknown extractor '{0}'; no rows produced" -f $ProfileDef.Extractor)
    return , @()
}

function Get-SasLogMetricRows {
    <# .SYNOPSIS Extract step timings from a SAS log.
       .OUTPUTS  Array of PSCustomObject shaped by $script:SasLogColumns.
       .NOTES    Finds each "NOTE: <step> used (Total process time):" header
                 and reads the real/cpu times from the lines that follow. #>
    param([hashtable] $ProfileDef, [string[]] $Lines,
          [string] $FullPath, [string] $ProgramName)
    $rows = New-Object System.Collections.Generic.List[object]
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

function Join-LibnameStatement {
    <# .SYNOPSIS Join a LIBNAME statement from its first line to its ';'.
       .PARAMETER Lines     The file's lines.
       .PARAMETER Start     Index of the line the LIBNAME keyword is on.
       .PARAMETER Offset    Column the LIBNAME keyword starts at, so a ';'
                            belonging to an earlier statement on the same line
                            ("run; libname x ...") does not end this one.
       .PARAMETER MaxLines  Give up after this many lines.
       .OUTPUTS  [hashtable] Text (statement without the ';') and EndIndex. #>
    param([string[]] $Lines, [int] $Start, [int] $Offset, [int] $MaxLines)
    $pieces = New-Object System.Collections.Generic.List[string]
    $last = [Math]::Min($Lines.Count - 1, $Start + $MaxLines - 1)
    for ($i = $Start; $i -le $last; $i++) {
        $text = if ($i -eq $Start) { $Lines[$i].Substring($Offset) } else { $Lines[$i] }
        $cut = $text.IndexOf(';')
        if ($cut -ge 0) {
            $pieces.Add($text.Substring(0, $cut))
            return @{ Text = ($pieces -join ' '); EndIndex = $i }
        }
        $pieces.Add($text)
    }
    return @{ Text = ($pieces -join ' '); EndIndex = $last }
}

function Get-AccessDbRows {
    <# .SYNOPSIS Inventory LIBNAMEs pointing at Microsoft Access files.
       .OUTPUTS  Array of PSCustomObject shaped by $script:AccessDbColumns,
                 one per (file, libref), ordered by definition line.
       .NOTES    A libref is USED where it appears followed by a dot --
                 issuelog.claims -- which is how a dataset in that library is
                 named. Matching the bare word would also count a macro
                 variable or a column of the same name.

                 The lines of the LIBNAME statements themselves are never
                 counted as usage. That is not tidiness: in
                 "libname issuelog access path='...\issuelog.mdb'" the file
                 name matches the libref-dot pattern, so counting the
                 definition line would report a use that is not one. #>
    param([hashtable] $ProfileDef, [string[]] $Lines,
          [string] $FullPath, [string] $ProgramName)
    $ic = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    $keywords = @($ProfileDef.Keywords)
    $maxLines = [int]$ProfileDef.MaxStatementLines

    $definitions    = New-Object System.Collections.Specialized.OrderedDictionary
    $statementLines = @{}
    $skippedMacroPaths = 0

    for ($i = 0; $i -lt $Lines.Count; $i++) {
        $found = [regex]::Match($Lines[$i], $script:LibnameStart, $ic)
        if (-not $found.Success) { continue }
        $statement = Join-LibnameStatement -Lines $Lines -Start $i `
                        -Offset $found.Index -MaxLines $maxLines
        $libref = $found.Groups['libref'].Value
        $key = $libref.ToLowerInvariant()
        # Every line of every LIBNAME statement for this libref -- the
        # definition, a redefinition, and "libname x clear;" -- is excluded
        # from the usage count below.
        if (-not $statementLines.ContainsKey($key)) { $statementLines[$key] = @{} }
        for ($j = $i; $j -le $statement.EndIndex; $j++) { $statementLines[$key][$j] = $true }

        # Match the keyword as a FILE EXTENSION on a macro-free statement. A
        # plain substring search reports "path=&mdbPath" as a real hit -- the
        # macro variable's own name contains "mdb" -- which is precisely the
        # case that must be counted as skipped instead.
        $visible = [regex]::Replace($statement.Text, $script:MacroReference, ' ')
        $keyword = $null
        foreach ($candidate in $keywords) {
            if ([regex]::IsMatch($visible, '\.' + [regex]::Escape($candidate) + '\b', $ic)) {
                $keyword = $candidate; break
            }
        }
        if ($null -eq $keyword) {
            if ([regex]::IsMatch($statement.Text, $script:MacroReference) -and
                [regex]::IsMatch($statement.Text, $script:AccessEngine, $ic)) {
                $skippedMacroPaths++
            }
            continue
        }
        if ($definitions.Contains($key)) { continue }   # first definition wins

        $quoted = [regex]::Match($statement.Text, $script:QuotedPath)
        $databaseFile = if ($quoted.Success) { $quoted.Groups['path'].Value.Trim() } else { 'NA' }
        if (-not $databaseFile) { $databaseFile = 'NA' }
        # The path's own extension is the better answer when both keywords
        # appear somewhere in the statement (a comment, another path).
        foreach ($candidate in $keywords) {
            if ($databaseFile.ToLowerInvariant().EndsWith('.' + $candidate)) {
                $keyword = $candidate; break
            }
        }

        $definitions[$key] = [ordered]@{
            FullPath       = $FullPath
            ProgramName    = $ProgramName
            Libref         = $libref
            DatabaseFile   = $databaseFile
            Keyword        = $keyword
            DefinitionLine = $i + 1
        }
    }

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($key in $definitions.Keys) {
        $row = $definitions[$key]
        $usePattern = '\b' + [regex]::Escape($row.Libref) + '\.'
        $ignore = $statementLines[$key]
        $hits = New-Object System.Collections.Generic.List[int]
        for ($i = 0; $i -lt $Lines.Count; $i++) {
            if ($ignore.ContainsKey($i)) { continue }
            if ([regex]::IsMatch($Lines[$i], $usePattern, $ic)) { $hits.Add($i + 1) }
        }
        $row.UsageCount = $hits.Count
        $row.UsageLines = if ($hits.Count -gt 0) { $hits -join ',' } else { 'NA' }
        $rows.Add([PSCustomObject]$row)
    }

    if ($skippedMacroPaths -gt 0) {
        Write-CgsWarn ("access_db: {0} LIBNAME statement(s) in {1} build the path from a macro variable, so no 'accdb'/'mdb' is visible and they are not reported" -f $skippedMacroPaths, $FullPath)
    }
    return , $rows.ToArray()
}

function Write-MatchExcel {
    <# .SYNOPSIS Write matches (and metrics) to an .xlsx via ImportExcel.
       .OUTPUTS  [string] path written, or $null when ImportExcel is absent. #>
    param([object[]] $MatchRows, [object[]] $MetricRows, [string] $Target,
          [string[]] $MetricColumns = $script:MetricColumns)
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) { return $null }
    Import-Module ImportExcel -ErrorAction Stop
    if (Test-Path -LiteralPath $Target) { Remove-Item -LiteralPath $Target -Force }
    $MatchRows | Select-Object $script:MatchColumns |
        Export-Excel -Path $Target -WorksheetName $script:MatchSheet -AutoSize
    if ($MetricRows.Count -gt 0) {
        $MetricRows | Select-Object $MetricColumns |
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
    # Which cgsUtils.ps1 actually loaded. Copying the .ps1 files to a share one
    # at a time is how versions drift, and this line is what makes it obvious.
    Write-CgsInfo (Get-CgsUtilsBanner)

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

    # Settle the target BEFORE the crawl: on a big share the scan takes
    # minutes, and a locked file must not cost us that work.
    $target = Resolve-CgsWritableTarget -Target $target

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
        $metricColumns = if ($profileDef.Columns) { @($profileDef.Columns) } else { $script:MetricColumns }
        $written = Write-MatchExcel -MatchRows $matchRows.ToArray() -MetricRows $metricRows.ToArray() -Target $target -MetricColumns $metricColumns
        if ($null -eq $written) {
            Write-CgsWarn 'no Excel engine (ImportExcel module) available; falling back to CSV'
            $target = [System.IO.Path]::ChangeExtension($target, '.csv')
            [void](Write-CgsCsv -Rows $matchRows.ToArray() -Columns $script:MatchColumns -Target $target)
            Write-CgsInfo ("wrote {0} match row(s) to {1}" -f $matchRows.Count, $target)
            if ($metricRows.Count -gt 0) {
                $companion = Join-Path ([System.IO.Path]::GetDirectoryName($target)) `
                    ("{0}_Metrics.csv" -f [System.IO.Path]::GetFileNameWithoutExtension($target))
                [void](Write-CgsCsv -Rows $metricRows.ToArray() -Columns $metricColumns -Target $companion)
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
