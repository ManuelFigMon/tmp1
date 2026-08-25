<#
=====================================================================
  Program Name  : scanFileSystem.ps1
  Author        : Manuel Figallo
  Purpose       : General-purpose file-system scanner and text-extraction
                  utility. It crawls one or more directory roots, captures
                  file-system metadata for every matching file, extracts
                  caller-specified keywords together with surrounding
                  context, optionally filters by a date range, and (via an
                  opt-in metric profile) parses structured performance
                  metrics from log files. The flagship profile extracts SAS
                  per-step "real time" and "cpu time", but the same engine
                  generalizes to any keyword sweep or log-metric use case.
  Version       : 1.3.3
  Created       : 2026-08-25
  Last Modified : 2026-08-25

  This is the PowerShell port of scanFileSystem.py. Same parameters, same
  output columns, same exit codes. See ../README.md and ../scanFileSystem.py
  for the authoritative documentation.

  Dependencies:
    CSV output requires nothing beyond Windows PowerShell 5.1 (or PowerShell
    7+). XLSX output requires the ImportExcel module; without it the scan
    falls back to CSV with a warning, exactly like the Python version falls
    back when openpyxl is missing.

  Description:
    Runs unattended on a schedule (Windows Task Scheduler) or from a SAS
    SYSTASK wrapper. Fully parameterized with CLI overrides; NEVER prompts
    interactively -- note that no parameter is declared Mandatory, because a
    Mandatory parameter would make PowerShell prompt and hang an unattended
    run. Missing required values are validated manually and exit non-zero.

  Input Parameters (required first):
    -InputFolderRoot      (REQUIRED, string[]) - root path(s) to search.
                            Accepts an array or a semicolon-delimited string
                            (how the SAS wrapper passes it). Empty -> exit 2.
    -OutputFilePath       (optional, string) - .csv or .xlsx by extension; a
                            directory auto-names a timestamped .csv inside it.
                            When omitted, writes scan_YYYYMMDD_HHMMSS.csv to
                            the current directory.
    -FileExtensions       (string[], default log,txt,sas) - case-insensitive.
    -IncludeSubdirectories (string, default 'true') - recurse when true.
                            Accepts true/false/1/0/yes/no ($true/$false work
                            interactively). A [bool] cannot be used because
                            -File mode passes every argument as a string.
    -FolderExclusionList  (string[], default empty) - folder names/tokens to
                            exclude; empty means nothing is excluded.
    -FileExclusionList    (string[], default empty) - prefixes/tokens stripped
                            from the filename to derive program_name.
    -ExtractKeyword       (string[], default empty) - keywords to extract with
                            a matched line, a +/-3 line context window, and a
                            match count.
    -DateFrom             (string, default none) - inclusive lower bound.
    -DateTo               (string, default none) - inclusive upper bound.
    -DateField            (string, default modified) - created/modified/accessed.
    -MetricProfile        (string, default none) - none | sas_log.

  Output:
    Files grain (one row per file) and, when a metric profile is active, a
    StepDetail grain -- as a second "StepDetail" sheet (.xlsx) or a companion
    "<stem>_StepDetail.csv" (.csv). Columns match the Python version exactly.

  Exit codes:
    0 = success, 2 = config error, 3 = I/O error.

  Usage:
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File scanFileSystem.ps1 `
        -InputFolderRoot "\\srv\logs" -OutputFilePath "C:\Logs\scan.xlsx" `
        -MetricProfile sas_log -ExtractKeyword "real time","cpu time"

  Change Log:
    v1.3.3 - Initial PowerShell port, feature-matched to scanFileSystem.py
             v1.3.3 (optional output path, empty folder exclusions by
             default, opt-in metric profiles, no third-party CSV deps).
=====================================================================
#>

[CmdletBinding()]
param(
    # NOTE: deliberately NOT Mandatory -- a Mandatory parameter prompts, which
    # would hang an unattended Task Scheduler / SYSTASK run.
    [string[]] $InputFolderRoot       = @(),
    [string]   $OutputFilePath        = '',
    [string[]] $FileExtensions        = @('log', 'txt', 'sas'),
    # NOTE: typed [string], not [bool], on purpose. In "powershell.exe -File"
    # mode every argument arrives as a string, so a [bool] parameter cannot be
    # set at all ("Cannot convert value System.String to type System.Boolean").
    # Accepts true/false/1/0/yes/no; $true / $false also work interactively.
    [string]   $IncludeSubdirectories = 'true',
    [string[]] $FolderExclusionList   = @(),
    [string[]] $FileExclusionList     = @(),
    [string[]] $ExtractKeyword        = @(),
    [string]   $DateFrom              = '',
    [string]   $DateTo                = '',
    [string]   $DateField             = 'modified',
    [string]   $MetricProfile         = 'none'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# =====================================================================
# Constants (mirror scanFileSystem.py)
# =====================================================================

$script:Version           = '1.3.3'
$script:ContextLines      = 3
$script:ValidDateFields   = @('created', 'modified', 'accessed')
$script:DefaultOutPrefix  = 'scan'
$script:TimestampFormat   = 'yyyyMMdd_HHmmss'

$script:ExitOk            = 0
$script:ExitConfigError   = 2
$script:ExitIoError       = 3

$script:FilesSheet        = 'Files'
$script:StepDetailSheet   = 'StepDetail'

$script:FilesBaseColumns = @(
    'program_name', 'log_file_name', 'full_path', 'directory', 'extension',
    'file_size_bytes', 'created_time', 'modified_time', 'accessed_time',
    'step_count', 'total_real_time_sec', 'total_cpu_time_sec',
    'max_step_real_time_sec', 'max_step_label', 'error_count', 'warning_count'
)
$script:FilesTailColumns  = @('parse_status', 'scanned_at')
$script:StepDetailColumns = @(
    'full_path', 'program_name', 'step_index', 'step_label',
    'real_time_sec', 'cpu_time_sec'
)

# Metric-profile registry. Add an entry here to support a new log format --
# nothing in the crawl or output code changes.
$script:MetricProfiles = @{
    'none' = @{ Active = $false }
    'sas_log' = @{
        Active      = $true
        StepPattern = '^NOTE:\s+(?<label>.+?)\s+used\s+\(Total process time\)'
        Metrics     = [ordered]@{
            real_time_sec = '^\s*real time\s+(?<value>[0-9:.]+)'
            cpu_time_sec  = '^\s*cpu time\s+(?<value>[0-9:.]+)'
        }
        Counters    = [ordered]@{
            error_count   = '^\s*ERROR[: ]'
            warning_count = '^\s*WARNING[: ]'
        }
        Lookahead   = 6
    }
}

# =====================================================================
# Logging -- stderr only, never interactive
# =====================================================================

function Write-Log {
    param([string] $Level, [string] $Message)
    $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    [Console]::Error.WriteLine(('{0} {1,-7} {2}' -f $stamp, $Level, $Message))
}
function Write-InfoLog  { param([string] $Message) Write-Log 'INFO'    $Message }
function Write-WarnLog  { param([string] $Message) Write-Log 'WARNING' $Message }
function Write-ErrorLog { param([string] $Message) Write-Log 'ERROR'   $Message }

# =====================================================================
# Helpers
# =====================================================================

function ConvertTo-List {
    <#  Accept an array, a single string, or a semicolon-delimited string
        (how the SAS wrapper passes list parameters). #>
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

function ConvertTo-Bool {
    <#  Parse a boolean-ish string. Returns $null when unrecognized so the
        caller can report a config error instead of guessing. #>
    param([string] $Value)
    if ($null -eq $Value) { return $null }
    switch ($Value.Trim().ToLowerInvariant()) {
        { $_ -in @('true',  '1', 'yes', 'y', '$true')  } { return $true }
        { $_ -in @('false', '0', 'no',  'n', '$false') } { return $false }
        default { return $null }
    }
}

function Get-DurationSeconds {
    <#  "0.05 seconds" -> 0.05 ; "1:03.05" -> 63.05 ; "1:00:30.00" -> 3630 #>
    param([string] $Raw)
    if ([string]::IsNullOrWhiteSpace($Raw)) { return $null }
    $text = $Raw.ToLowerInvariant().Replace('seconds', '').Replace('second', '').Trim()
    if (-not $text) { return $null }
    $ic = [System.Globalization.CultureInfo]::InvariantCulture

    $m = [regex]::Match($text, '^(\d+):(\d{1,2}):(\d{1,2}(?:\.\d+)?)$')   # hh:mm:ss
    if ($m.Success) {
        return [double]$m.Groups[1].Value * 3600 +
               [double]$m.Groups[2].Value * 60 +
               [double]::Parse($m.Groups[3].Value, $ic)
    }
    $m = [regex]::Match($text, '^(\d+):(\d{1,2}(?:\.\d+)?)$')             # mm:ss
    if ($m.Success) {
        return [double]$m.Groups[1].Value * 60 + [double]::Parse($m.Groups[2].Value, $ic)
    }
    $m = [regex]::Match($text, '^(\d+(?:\.\d+)?)$')                       # seconds
    if ($m.Success) { return [double]::Parse($m.Groups[1].Value, $ic) }
    return $null
}

function Get-KeywordSlug {
    param([string] $Keyword)
    $slug = ([regex]::Replace($Keyword, '[^0-9a-zA-Z]+', '_')).Trim('_').ToLowerInvariant()
    if (-not $slug) { return 'kw' }
    return $slug
}

function Get-UniqueSlugs {
    param([string[]] $Keywords)
    $seen = @{}
    $slugs = New-Object System.Collections.Generic.List[string]
    foreach ($keyword in $Keywords) {
        $base = Get-KeywordSlug $keyword
        if ($seen.ContainsKey($base)) { $seen[$base] += 1 } else { $seen[$base] = 1 }
        if ($seen[$base] -eq 1) { $slugs.Add($base) } else { $slugs.Add("$base`_$($seen[$base])") }
    }
    return , $slugs.ToArray()
}

function Get-KeywordColumns {
    param([string[]] $Keywords)
    $columns = New-Object System.Collections.Generic.List[string]
    [string[]] $slugs = Get-UniqueSlugs $Keywords
    foreach ($slug in $slugs) {
        $columns.Add("kw_${slug}_line")
        $columns.Add("kw_${slug}_context")
        $columns.Add("kw_${slug}_count")
    }
    return , $columns.ToArray()
}

function Get-KeywordExtract {
    <#  First matched line, a +/-ContextLines window, and a match count. #>
    param([string[]] $Lines, [string[]] $Keywords)
    $result = [ordered]@{}
    [string[]] $slugs = Get-UniqueSlugs $Keywords
    for ($k = 0; $k -lt $Keywords.Count; $k++) {
        $needle = $Keywords[$k].ToLowerInvariant()
        $slug = $slugs[$k]
        $firstLine = ''; $context = ''; $count = 0
        for ($i = 0; $i -lt $Lines.Count; $i++) {
            if ($Lines[$i].ToLowerInvariant().Contains($needle)) {
                $count++
                if (-not $firstLine) {
                    $firstLine = $Lines[$i].Trim()
                    $lo = [Math]::Max(0, $i - $script:ContextLines)
                    $hi = [Math]::Min($Lines.Count - 1, $i + $script:ContextLines)
                    $context = ($Lines[$lo..$hi] -join "`n")
                }
            }
        }
        $result["kw_${slug}_line"]    = $firstLine
        $result["kw_${slug}_context"] = $context
        $result["kw_${slug}_count"]   = $count
    }
    return $result
}

function ConvertTo-NormalizedPath {
    param([string] $Text)
    return ($Text -replace '\\', '/').TrimEnd('/').ToLowerInvariant()
}

function Test-FolderExcluded {
    <#  True when an ancestor directory segment matches (exact, case-
        insensitive) or a full-path prefix matches. Excluding "Old" must NOT
        exclude a sibling "Older". #>
    param([string] $FullPath, [string[]] $Exclusions)
    if (-not $Exclusions -or $Exclusions.Count -eq 0) { return $false }

    $parent = [System.IO.Path]::GetDirectoryName($FullPath)
    if (-not $parent) { return $false }
    $segments = @($parent -split '[\\/]' | Where-Object { $_ } |
                  ForEach-Object { $_.ToLowerInvariant() })
    $normalizedPath = ConvertTo-NormalizedPath $FullPath

    foreach ($raw in $Exclusions) {
        $token = $raw.Trim()
        if (-not $token) { continue }
        $tokenSegment = ($token -replace '\\', '/').TrimEnd('/').ToLowerInvariant()
        if ($segments -contains $tokenSegment) { return $true }
        $prefix = ConvertTo-NormalizedPath $token
        if ($prefix -and ($normalizedPath -eq $prefix -or
                          $normalizedPath.StartsWith($prefix + '/'))) { return $true }
    }
    return $false
}

function Get-NormalizedExtensions {
    param([string[]] $Extensions)
    $set = New-Object System.Collections.Generic.HashSet[string]
    foreach ($extension in $Extensions) {
        $clean = $extension.Trim().TrimStart('.').ToLowerInvariant()
        if ($clean) { [void]$set.Add($clean) }
    }
    return $set
}

function Get-DateBoundary {
    <#  Parse YYYY-MM-DD or an ISO datetime. A bare date used as an upper
        bound extends through end-of-day so the bound is truly inclusive.
        Throws on a malformed value. #>
    param([string] $Raw, [bool] $EndOfDay)
    if ([string]::IsNullOrWhiteSpace($Raw)) { return $null }
    $text = $Raw.Trim()
    $parsed = [datetime]::MinValue
    $styles = [System.Globalization.DateTimeStyles]::None
    $ic = [System.Globalization.CultureInfo]::InvariantCulture
    if (-not [datetime]::TryParse($text, $ic, $styles, [ref]$parsed)) {
        throw "unparseable date '$Raw'; expected YYYY-MM-DD or ISO datetime"
    }
    if ($EndOfDay -and $text.Length -eq 10) {
        $parsed = $parsed.Date.AddDays(1).AddTicks(-1)
    }
    return $parsed
}

function Get-ProgramName {
    <#  Filename minus extension, minus any configured prefix/token. #>
    param([string] $FileName, [string[]] $Exclusions)
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    foreach ($raw in $Exclusions) {
        $token = $raw.Trim()
        if (-not $token) { continue }
        if ($stem.ToLowerInvariant().StartsWith($token.ToLowerInvariant())) {
            $stem = $stem.Substring($token.Length)
        } else {
            $stem = [regex]::Replace($stem, [regex]::Escape($token), '',
                                     [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        }
    }
    $stem = $stem.Trim(' ', '_', '-', '.')
    if (-not $stem) { return [System.IO.Path]::GetFileNameWithoutExtension($FileName) }
    return $stem
}

function Read-TextLines {
    <#  UTF-8 (strict), falling back to latin-1, then UTF-8 with replacement. #>
    param([string] $Path)
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $strictUtf8 = New-Object System.Text.UTF8Encoding($false, $true)
    $text = $null
    try { $text = $strictUtf8.GetString($bytes) }
    catch {
        try   { $text = [System.Text.Encoding]::GetEncoding('iso-8859-1').GetString($bytes) }
        catch { $text = [System.Text.Encoding]::UTF8.GetString($bytes) }
    }
    if ($null -eq $text) { return @() }
    $lines = $text -split "`r`n|`n|`r"
    # Match Python's str.splitlines(): a trailing newline does NOT produce a
    # final empty element (this would otherwise show up in keyword context).
    if ($lines.Count -gt 0 -and $lines[-1] -eq '') {
        $lines = $lines[0..($lines.Count - 2)]
    }
    return , $lines
}

# =====================================================================
# Metric profile parsing
# =====================================================================

function Invoke-MetricProfile {
    <#  Return @{ Steps = @(...); Counters = @{...} } for the given profile. #>
    param([hashtable] $ProfileDef, [string[]] $Lines)

    $steps = New-Object System.Collections.Generic.List[object]
    $counters = [ordered]@{ error_count = 0; warning_count = 0 }
    if (-not $ProfileDef.Active) { return @{ Steps = $steps.ToArray(); Counters = $counters } }

    for ($i = 0; $i -lt $Lines.Count; $i++) {
        $header = [regex]::Match($Lines[$i], $ProfileDef.StepPattern,
                  [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if (-not $header.Success) { continue }

        $step = [ordered]@{
            step_index = $steps.Count + 1
            step_label = $header.Groups['label'].Value.Trim()
        }
        $windowEnd = [Math]::Min($Lines.Count - 1, $i + $ProfileDef.Lookahead)
        foreach ($metric in $ProfileDef.Metrics.Keys) {
            $step[$metric] = $null
            if ($i + 1 -gt $windowEnd) { continue }
            foreach ($candidate in $Lines[($i + 1)..$windowEnd]) {
                $hit = [regex]::Match($candidate, $ProfileDef.Metrics[$metric],
                       [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                if ($hit.Success) {
                    $step[$metric] = Get-DurationSeconds $hit.Groups['value'].Value
                    break
                }
            }
        }
        $steps.Add($step)
    }

    foreach ($line in $Lines) {
        foreach ($counter in $ProfileDef.Counters.Keys) {
            if ([regex]::IsMatch($line, $ProfileDef.Counters[$counter],
                [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)) {
                $counters[$counter] = $counters[$counter] + 1
            }
        }
    }
    return @{ Steps = $steps.ToArray(); Counters = $counters }
}

function Get-StepAggregate {
    <#  Roll StepDetail rows up to the Files grain. #>
    param([object[]] $Steps)
    $totalReal = 0.0; $totalCpu = 0.0; $maxReal = 0.0; $maxLabel = ''
    foreach ($step in $Steps) {
        if ($null -ne $step.real_time_sec) {
            $totalReal += $step.real_time_sec
            if ($step.real_time_sec -gt $maxReal) {
                $maxReal = $step.real_time_sec; $maxLabel = $step.step_label
            }
        }
        if ($null -ne $step.cpu_time_sec) { $totalCpu += $step.cpu_time_sec }
    }
    return [ordered]@{
        step_count             = $Steps.Count
        total_real_time_sec    = [Math]::Round($totalReal, 6)
        total_cpu_time_sec     = [Math]::Round($totalCpu, 6)
        max_step_real_time_sec = $maxReal
        max_step_label         = $maxLabel
    }
}

# =====================================================================
# Crawl
# =====================================================================

function Get-CandidateFile {
    <#  Every file under the roots. Missing roots are logged and skipped.
        Returns @{ Files = @(paths); Reachable = n }. #>
    param([string[]] $Roots, [bool] $Recurse)
    $found = New-Object System.Collections.Generic.List[string]
    $reachable = 0

    foreach ($rawRoot in $Roots) {
        $root = $rawRoot.Trim()
        if (-not (Test-Path -LiteralPath $root)) {
            Write-WarnLog "root not found, skipping: $root"; continue
        }
        if (-not (Test-Path -LiteralPath $root -PathType Container)) {
            Write-WarnLog "root is not a directory, skipping: $root"; continue
        }
        $reachable++
        try {
            $items = Get-ChildItem -LiteralPath $root -File -Recurse:$Recurse `
                                   -Force -ErrorAction SilentlyContinue -ErrorVariable walkErrors
            foreach ($item in $items) { $found.Add($item.FullName) }
            foreach ($walkError in $walkErrors) {
                Write-WarnLog "cannot descend into $($walkError.TargetObject): $($walkError.Exception.Message)"
            }
        } catch {
            Write-WarnLog "permission denied while walking ${root}: $($_.Exception.Message)"
        }
    }
    return @{ Files = $found.ToArray(); Reachable = $reachable }
}

function Invoke-Scan {
    param(
        [string[]] $Roots,
        [string[]] $Extensions,
        [bool]     $Recurse,
        [string[]] $FolderExclusions,
        [string[]] $FileExclusions,
        [string[]] $Keywords,
        [hashtable] $ProfileDef,
        [object]   $DateLow,
        [object]   $DateHigh,
        [string]   $WhichDate
    )

    $wanted    = Get-NormalizedExtensions $Extensions
    $scannedAt = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
    $filesRows = New-Object System.Collections.Generic.List[object]
    $stepRows  = New-Object System.Collections.Generic.List[object]
    [string[]] $keywordColumns = Get-KeywordColumns $Keywords

    $discovered = Get-CandidateFile -Roots $Roots -Recurse $Recurse
    Write-InfoLog ("discovered {0} file(s) across {1} reachable root(s)" -f `
                   $discovered.Files.Count, $discovered.Reachable)

    # Ordinal sort so row order matches the Python port (which sorts by
    # code point: uppercase folder names come before lowercase file names).
    [string[]] $orderedPaths = $discovered.Files
    [Array]::Sort($orderedPaths, [StringComparer]::Ordinal)
    foreach ($path in $orderedPaths) {
        $extension = ([System.IO.Path]::GetExtension($path)).TrimStart('.').ToLowerInvariant()
        if ($wanted.Count -gt 0 -and -not $wanted.Contains($extension)) { continue }
        if (Test-FolderExcluded -FullPath $path -Exclusions $FolderExclusions) { continue }

        $fileName    = [System.IO.Path]::GetFileName($path)
        $programName = Get-ProgramName -FileName $fileName -Exclusions $FileExclusions

        $row = [ordered]@{
            program_name           = $programName
            log_file_name          = $fileName
            full_path              = $path
            directory              = [System.IO.Path]::GetDirectoryName($path)
            extension              = $extension
            file_size_bytes        = 0
            created_time           = ''
            modified_time          = ''
            accessed_time          = ''
            step_count             = 0
            total_real_time_sec    = 0.0
            total_cpu_time_sec     = 0.0
            max_step_real_time_sec = 0.0
            max_step_label         = ''
            error_count            = 0
            warning_count          = 0
        }
        foreach ($column in $keywordColumns) {
            $row[$column] = if ($column.EndsWith('_count')) { 0 } else { '' }
        }
        $row['parse_status'] = 'OK'
        $row['scanned_at']   = $scannedAt

        # --- metadata + date filter (a stat failure still emits a row) ---
        try {
            $info = Get-Item -LiteralPath $path -Force -ErrorAction Stop
            if (-not $info.PSObject.Properties['Length']) { throw 'not a file' }
            $times = @{
                created  = $info.CreationTime
                modified = $info.LastWriteTime
                accessed = $info.LastAccessTime
            }
            if ($null -ne $DateLow  -and $times[$WhichDate] -lt $DateLow)  { continue }
            if ($null -ne $DateHigh -and $times[$WhichDate] -gt $DateHigh) { continue }

            $row['file_size_bytes'] = $info.Length
            $row['created_time']    = $times['created'].ToString('yyyy-MM-ddTHH:mm:ss')
            $row['modified_time']   = $times['modified'].ToString('yyyy-MM-ddTHH:mm:ss')
            $row['accessed_time']   = $times['accessed'].ToString('yyyy-MM-ddTHH:mm:ss')
        } catch {
            $row['parse_status'] = "stat error: $($_.Exception.GetType().Name): $($_.Exception.Message)"
            Write-WarnLog "cannot stat ${path}: $($_.Exception.Message)"
            $filesRows.Add([PSCustomObject]$row)
            continue
        }

        # --- content: keywords + metric profile ---
        try {
            $lines = Read-TextLines -Path $path
            if ($Keywords.Count -gt 0) {
                $extracted = Get-KeywordExtract -Lines $lines -Keywords $Keywords
                foreach ($key in $extracted.Keys) { $row[$key] = $extracted[$key] }
            }
            $parsed = Invoke-MetricProfile -ProfileDef $ProfileDef -Lines $lines
            $row['error_count']   = $parsed.Counters['error_count']
            $row['warning_count'] = $parsed.Counters['warning_count']

            if ($ProfileDef.Active) {
                $aggregate = Get-StepAggregate -Steps $parsed.Steps
                foreach ($key in $aggregate.Keys) { $row[$key] = $aggregate[$key] }
                foreach ($step in $parsed.Steps) {
                    $stepRows.Add([PSCustomObject][ordered]@{
                        full_path     = $path
                        program_name  = $programName
                        step_index    = $step.step_index
                        step_label    = $step.step_label
                        real_time_sec = $step.real_time_sec
                        cpu_time_sec  = $step.cpu_time_sec
                    })
                }
            }
        } catch {
            $row['parse_status'] = "read error: $($_.Exception.GetType().Name): $($_.Exception.Message)"
            Write-WarnLog "cannot read ${path}: $($_.Exception.Message)"
        }

        $filesRows.Add([PSCustomObject]$row)
    }

    return @{
        Files     = $filesRows.ToArray()
        Steps     = $stepRows.ToArray()
        Reachable = $discovered.Reachable
    }
}

# =====================================================================
# Output
# =====================================================================

function Get-DefaultOutputName {
    param([string] $Directory = '')
    $stamp = (Get-Date).ToString($script:TimestampFormat)
    $name = "$($script:DefaultOutPrefix)_$stamp.csv"
    if ($Directory) { return (Join-Path $Directory $name) }
    return $name
}

function Resolve-OutputPath {
    <#  A directory target auto-names a timestamped .csv inside it. #>
    param([string] $Raw)
    $path = $Raw.Trim()
    $suffix = ([System.IO.Path]::GetExtension($path)).ToLowerInvariant()
    $looksLikeDir = (Test-Path -LiteralPath $path -PathType Container) -or
                    $path.EndsWith('/') -or $path.EndsWith('\') -or
                    ($suffix -ne '.csv' -and $suffix -ne '.xlsx')
    if ($looksLikeDir) {
        if (-not (Test-Path -LiteralPath $path)) {
            [void](New-Item -ItemType Directory -Path $path -Force)
        }
        $generated = Get-DefaultOutputName -Directory $path
        Write-InfoLog "output path is a directory; writing $generated"
        return $generated
    }
    $parent = [System.IO.Path]::GetDirectoryName($path)
    if ($parent -and -not (Test-Path -LiteralPath $parent)) {
        [void](New-Item -ItemType Directory -Path $parent -Force)
    }
    return $path
}

function Write-CsvRows {
    <#  Write rows as UTF-8 CSV with no BOM, matching the Python output byte
        for byte. Export-Csv is avoided because Windows PowerShell 5.1 emits
        a BOM and quotes every field. #>
    param([object[]] $Rows, [string[]] $Columns, [string] $Target)

    $encoding = New-Object System.Text.UTF8Encoding($false)
    $writer = New-Object System.IO.StreamWriter($Target, $false, $encoding)
    try {
        $writer.NewLine = "`r`n"
        $writer.WriteLine(($Columns | ForEach-Object { ConvertTo-CsvField $_ }) -join ',')
        foreach ($row in $Rows) {
            $fields = foreach ($column in $Columns) {
                $value = if ($row.PSObject.Properties[$column]) { $row.$column } else { '' }
                ConvertTo-CsvField $value
            }
            $writer.WriteLine($fields -join ',')
        }
    } finally { $writer.Dispose() }
}

function Format-CellValue {
    <#  Render a value the way Python's csv writer does, so the two ports
        produce identical files: doubles always keep a decimal place
        (2 -> "2.0"), everything else is plain ToString(). #>
    param([object] $Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [double] -or $Value -is [single] -or $Value -is [decimal]) {
        $ic = [System.Globalization.CultureInfo]::InvariantCulture
        return ([double]$Value).ToString('0.0###############', $ic)
    }
    return [string]$Value
}

function ConvertTo-CsvField {
    <#  Minimal CSV quoting, matching Python's csv module defaults. #>
    param([object] $Value)
    if ($null -eq $Value) { return '""' }
    $text = Format-CellValue $Value
    if ($text -match '[",\r\n]') { return '"' + $text.Replace('"', '""') + '"' }
    return $text
}

function Test-ExcelModule {
    return [bool](Get-Module -ListAvailable -Name ImportExcel)
}

function Write-ScanOutput {
    param(
        [object[]] $FilesRows, [object[]] $StepRows, [string] $Target,
        [string[]] $Keywords, [bool] $ProfileActive
    )
    [string[]] $keywordColumns = Get-KeywordColumns $Keywords
    [string[]] $columns = @($script:FilesBaseColumns) + $keywordColumns +
                          @($script:FilesTailColumns)
    $written = New-Object System.Collections.Generic.List[string]
    $target = $Target

    if (([System.IO.Path]::GetExtension($target)).ToLowerInvariant() -eq '.xlsx') {
        if (-not (Test-ExcelModule)) {
            Write-WarnLog 'no Excel engine (ImportExcel module) available; falling back to CSV'
            $target = [System.IO.Path]::ChangeExtension($target, '.csv')
        } else {
            Import-Module ImportExcel -ErrorAction Stop
            if (Test-Path -LiteralPath $target) { Remove-Item -LiteralPath $target -Force }
            $FilesRows | Select-Object $columns |
                Export-Excel -Path $target -WorksheetName $script:FilesSheet -AutoSize
            Write-InfoLog ("wrote {0} Files row(s) to {1} [ImportExcel]" -f $FilesRows.Count, $target)
            if ($ProfileActive) {
                $StepRows | Select-Object $script:StepDetailColumns |
                    Export-Excel -Path $target -WorksheetName $script:StepDetailSheet -AutoSize
                Write-InfoLog ("wrote {0} StepDetail row(s) to sheet '{1}'" -f `
                               $StepRows.Count, $script:StepDetailSheet)
            }
            $written.Add($target)
            return , $written.ToArray()
        }
    }

    Write-CsvRows -Rows $FilesRows -Columns $columns -Target $target
    $written.Add($target)
    Write-InfoLog ("wrote {0} Files row(s) to {1}" -f $FilesRows.Count, $target)

    if ($ProfileActive) {
        $stem = [System.IO.Path]::GetFileNameWithoutExtension($target)
        $dir  = [System.IO.Path]::GetDirectoryName($target)
        $companion = if ($dir) { Join-Path $dir "${stem}_StepDetail.csv" } else { "${stem}_StepDetail.csv" }
        Write-CsvRows -Rows $StepRows -Columns $script:StepDetailColumns -Target $companion
        $written.Add($companion)
        Write-InfoLog ("wrote {0} StepDetail row(s) to companion {1}" -f $StepRows.Count, $companion)
    }
    return , $written.ToArray()
}

# =====================================================================
# Main
# =====================================================================

function Invoke-Main {
    $roots            = ConvertTo-List $InputFolderRoot
    $extensions       = ConvertTo-List $FileExtensions
    $folderExclusions = ConvertTo-List $FolderExclusionList
    $fileExclusions   = ConvertTo-List $FileExclusionList
    $keywords         = ConvertTo-List $ExtractKeyword

    # ---- validation (before any crawling; never prompts) ----
    if ($roots.Count -eq 0) {
        Write-ErrorLog ("required parameter 'input_folder_root' is missing or empty; " +
                        'pass -InputFolderRoot')
        return $script:ExitConfigError
    }
    if (-not $script:MetricProfiles.Contains($MetricProfile)) {
        Write-ErrorLog ("unknown metric_profile '$MetricProfile'; expected one of: " +
                        (($script:MetricProfiles.Keys | Sort-Object) -join ', '))
        return $script:ExitConfigError
    }
    if ($script:ValidDateFields -notcontains $DateField) {
        Write-ErrorLog ("unknown date_field '$DateField'; expected one of: " +
                        ($script:ValidDateFields -join ', '))
        return $script:ExitConfigError
    }

    try {
        $dateLow  = Get-DateBoundary -Raw $DateFrom -EndOfDay $false
        $dateHigh = Get-DateBoundary -Raw $DateTo   -EndOfDay $true
    } catch {
        Write-ErrorLog $_.Exception.Message
        return $script:ExitConfigError
    }
    if ($null -ne $dateLow -and $null -ne $dateHigh -and $dateLow -gt $dateHigh) {
        Write-ErrorLog "date_from ($DateFrom) is after date_to ($DateTo)"
        return $script:ExitConfigError
    }

    $recurse = ConvertTo-Bool $IncludeSubdirectories
    if ($null -eq $recurse) {
        Write-ErrorLog ("unknown include_subdirectories '$IncludeSubdirectories'; " +
                        'expected true/false (or 1/0, yes/no)')
        return $script:ExitConfigError
    }

    $profileDef = $script:MetricProfiles[$MetricProfile]
    Write-InfoLog ("scanFileSystem {0} (PowerShell) starting; profile={1}; roots={2}" -f `
                   $script:Version, $MetricProfile, $roots.Count)

    # ---- resolve output BEFORE crawling so a bad path fails fast ----
    try {
        if ($OutputFilePath -and $OutputFilePath.Trim()) {
            $target = Resolve-OutputPath $OutputFilePath
        } else {
            $target = Get-DefaultOutputName
            Write-InfoLog "output_file_path not supplied; writing $target"
        }
    } catch {
        Write-ErrorLog "cannot prepare output path: $($_.Exception.Message)"
        return $script:ExitIoError
    }

    try {
        $scan = Invoke-Scan -Roots $roots -Extensions $extensions `
                    -Recurse $recurse -FolderExclusions $folderExclusions `
                    -FileExclusions $fileExclusions -Keywords $keywords -ProfileDef $profileDef `
                    -DateLow $dateLow -DateHigh $dateHigh -WhichDate $DateField
    } catch {
        Write-ErrorLog "fatal I/O error while scanning: $($_.Exception.Message)"
        return $script:ExitIoError
    }

    if ($scan.Reachable -eq 0) {
        Write-ErrorLog 'none of the supplied input_folder_root path(s) are reachable'
        return $script:ExitIoError
    }

    try {
        $written = Write-ScanOutput -FilesRows $scan.Files -StepRows $scan.Steps `
                       -Target $target -Keywords $keywords -ProfileActive $profileDef.Active
    } catch {
        Write-ErrorLog "cannot write output: $($_.Exception.Message)"
        return $script:ExitIoError
    }

    Write-InfoLog ("done; {0} file row(s), {1} step row(s), {2} output file(s)" -f `
                   $scan.Files.Count, $scan.Steps.Count, $written.Count)
    return $script:ExitOk
}

# A trailing exception would exit 1, which is indistinguishable from
# "powershell.exe could not start the script at all". Catch everything and
# report a controlled I/O error code with a real message instead.
try {
    exit (Invoke-Main)
} catch {
    Write-ErrorLog ("unhandled error: {0}" -f $_.Exception.Message)
    if ($_.ScriptStackTrace) {
        Write-ErrorLog ("at: {0}" -f ($_.ScriptStackTrace -split "`n")[0])
    }
    exit $script:ExitIoError
}
