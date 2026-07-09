# =============================================================================
#  Get-SasLogPerformance_v1.ps1
# =============================================================================
#  Version : 1.0
#  Author  : <AUTHOR NAME HERE>
#  Date    : 2026-07-09
#
#  PURPOSE
#  -------
#  Crawls one or more directories of SAS log files and builds a performance
#  inventory: one output row per SAS step (PROC / DATA step / initialization /
#  session total) with real time and CPU time normalized to decimal seconds,
#  plus file metadata and the SAS program name derived from the log filename.
#  Built to run unattended on a Windows Task Scheduler schedule - it never
#  prompts, survives bad log files, and emails a run summary when finished.
#
#  PARAMETERS (all overridable from the command line)
#  --------------------------------------------------
#  LogRoot                   string   \\a70admed.com\R1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT
#                                     Fallback root crawled when the input
#                                     directories CSV is absent.
#  InputDirectoriesCsv       string   <script dir>\input_directories.csv
#                                     CSV of root dirs (column FULL_PATH_DIR);
#                                     the primary way to specify log locations.
#  LogFileExtensions         string[] @('log','txt','sas')  Extensions included
#                                     (case-insensitive).
#  IncludeSubdirectories     bool     $true   Recurse into subfolders.
#  OutputPath                string   <script dir>\sas_log_inventory.csv
#                                     .csv -> CSV; .xlsx -> Excel (ImportExcel,
#                                     else Excel COM, else CSV fallback).
#  ExcelSheetName            string   'SAS_Log_Inventory'  Sheet name for .xlsx.
#  FolderExclusionCsv        string   <script dir>\folder_exclusion_list.csv
#                                     CSV of folders to skip (EXCLUDE_FOLDER,NOTES).
#  EnableRobocopy            bool     $false  Stage logs with robocopy first.
#  RobocopyConfigCsv         string   <script dir>\robocopy_config.csv
#                                     Robocopy job CSV (Copy-FilesWithRobocopy style).
#  EmailEnabled              bool     $true   Send completion/failure email.
#  EmailTo                   string[] @('dl-sas-admins@cgsadmin.com')
#  EmailFrom                 string   'sas-log-crawler@cgsadmin.com'
#  SmtpServer                string   'smtprelay.bcbssc.com'
#  SmtpPort                  int      25
#  ProgramNameSuffixPattern  string   '_\d{3,5}(AM|PM)$'  Trims schedule suffix.
#  ProgramNameVersionPattern string   '_V\d+(_\d+)*$'     Trims version suffix.
#  ProgramNameTokens         string[] @()  Literal trailing tokens to strip.
#
#  EXAMPLE input_directories.csv
#  -----------------------------
#      FULL_PATH_DIR,NOTES
#      E:\SAS\Logs\DME_Monthly,Monthly DME reporting logs
#      E:\SAS\Logs\DMEB_Quarterly,Quarterly POE and SSR logs
#      \\server01\sas_share\batch_logs,Network share - batch jobs
#
#  EXAMPLE folder_exclusion_list.csv
#  ---------------------------------
#      EXCLUDE_FOLDER,NOTES
#      E:\SAS\Logs\DME_Monthly\archive,Old logs - do not inventory
#      \\server01\sas_share\batch_logs\test,Test runs only
#
#  EXAMPLE robocopy_config.csv (used only when -EnableRobocopy $true)
#  ------------------------------------------------------------------
#      Source,Target,FileExtension,UpdateIfNewer,IncludeSubfolders,Notes
#      \\server01\sas_share\batch_logs,E:\SAS\LogStage\batch,log,Y,Y,Stage batch logs locally
#
#  EXAMPLE COMMAND LINES
#  ---------------------
#      powershell.exe -NoProfile -ExecutionPolicy Bypass -File "E:\Scripts\Get-SasLogPerformance_v1.ps1"
#      powershell.exe -NoProfile -ExecutionPolicy Bypass -File "E:\Scripts\Get-SasLogPerformance_v1.ps1" -OutputPath "E:\Reports\sas_log_metrics.xlsx" -ExcelSheetName "LogMetrics"
#      powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "& 'E:\Scripts\Get-SasLogPerformance_v1.ps1' -EnableRobocopy $true"
#
#  EXIT CODES:  0 = success   1 = completed with parse errors   2 = fatal
# =============================================================================

<#
.SYNOPSIS
    Crawls directories of SAS 9.4 log files, extracts per-step performance
    metrics (real time / CPU time), captures file-system metadata, derives the
    SAS program name from each log filename, and writes a single CSV or Excel
    inventory file.

.DESCRIPTION
    Get-SasLogPerformance_v1.ps1 is designed to run unattended under Windows
    Task Scheduler on PowerShell 5.1. It never prompts, logs everything to the
    console and to a rolling text log file, wraps every file parse in its own
    try/catch (one bad log cannot kill the run), and sends a completion or
    failure email through the corporate SMTP relay.

    Root directories to crawl come from an input_directories.csv file
    (column FULL_PATH_DIR); if that CSV is not found, the single -LogRoot
    path is crawled instead. Folders listed in folder_exclusion_list.csv are
    skipped. Optionally, logs can first be staged locally with robocopy using
    a CSV config in the Copy-FilesWithRobocopy_v1_3.ps1 style.

    Output format is chosen by the -OutputPath extension: .csv writes via
    Export-Csv -NoTypeInformation; .xlsx uses the ImportExcel module when
    available, falls back to Excel COM automation, and finally falls back to
    a CSV written alongside (with a warning) so the run never fails on output.

.PARAMETER LogRoot
    [string] Fallback root directory of SAS logs, crawled when
    -InputDirectoriesCsv is not supplied or the file does not exist.
    Default: \\a70admed.com\R1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT

.PARAMETER InputDirectoriesCsv
    [string] Path to input_directories.csv (column FULL_PATH_DIR, one row per
    root directory to traverse). This is the primary way to specify log
    locations. Default: input_directories.csv next to this script.

.PARAMETER LogFileExtensions
    [string[]] Extensions to include, matched case-insensitively and with or
    without a leading dot. Default: @('log','txt','sas').

.PARAMETER IncludeSubdirectories
    [bool] Recurse into subdirectories of each root. Default: $true.

.PARAMETER OutputPath
    [string] Full path of the output file. Extension decides the format:
    a .csv path writes CSV; a .xlsx path writes Excel (ImportExcel module,
    then Excel COM, then CSV fallback with warning).
    Default: sas_log_inventory.csv next to this script.

.PARAMETER ExcelSheetName
    [string] Worksheet name used when OutputPath is .xlsx.
    Default: 'SAS_Log_Inventory'.

.PARAMETER FolderExclusionCsv
    [string] Path to folder_exclusion_list.csv (columns EXCLUDE_FOLDER, NOTES).
    Any file whose full path starts with an excluded folder (case-insensitive,
    trailing-slash-safe) is skipped.
    Default: folder_exclusion_list.csv next to this script (optional).

.PARAMETER EnableRobocopy
    [bool] When $true, stage log files with robocopy before crawling, driven
    by -RobocopyConfigCsv. Robocopy exit codes 0-7 are treated as success.
    Default: $false.

.PARAMETER RobocopyConfigCsv
    [string] Path to the robocopy config CSV (columns: Source, Target,
    FileExtension, UpdateIfNewer, IncludeSubfolders, Notes) in the same style
    as Copy-FilesWithRobocopy_v1_3.ps1.
    Default: robocopy_config.csv next to this script.

.PARAMETER EmailEnabled
    [bool] Send a summary email on success and a failure email on fatal error.
    Default: $true.

.PARAMETER EmailTo
    [string[]] Recipient list (supports distribution lists),
    e.g. @('dl-sas-admins@cgsadmin.com').

.PARAMETER EmailFrom
    [string] From address for notifications.
    Default: 'sas-log-crawler@cgsadmin.com'.

.PARAMETER SmtpServer
    [string] SMTP relay host. Default: 'smtprelay.bcbssc.com'.

.PARAMETER SmtpPort
    [int] SMTP port. Default: 25.

.PARAMETER ProgramNameSuffixPattern
    [string] Regex that trims the scheduling-time suffix off a log filename to
    recover the SAS program name. Default: '_\d{3,5}(AM|PM)$'
    (e.g. DMEB_POE_Qtrly_Report_0915AM -> DMEB_POE_Qtrly_Report).

.PARAMETER ProgramNameVersionPattern
    [string] Regex applied after the schedule suffix to trim trailing program
    version tokens. Default: '_V\d+(_\d+)*$'
    (e.g. DME_Prepay_Top_Denials_V1_1 -> DME_Prepay_Top_Denials).
    Set to '' to keep version tokens in SAS_PROGRAM_NAME.

.PARAMETER ProgramNameTokens
    [string[]] Optional literal tokens (Strip-Tokens_v3 style) trimmed from the
    end of the derived program name as a secondary mechanism,
    e.g. @('_FINAL','_PROD'). Default: @() (none).

.EXAMPLE
    .\Get-SasLogPerformance_v1.ps1
    Crawl the directories in .\input_directories.csv (or -LogRoot if the CSV is
    absent) and write .\sas_log_inventory.csv.

.EXAMPLE
    .\Get-SasLogPerformance_v1.ps1 -OutputPath 'E:\Reports\sas_log_metrics.xlsx' -ExcelSheetName 'LogMetrics'
    Write an Excel workbook with worksheet 'LogMetrics'.

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "E:\Scripts\Get-SasLogPerformance_v1.ps1"
    Typical Windows Task Scheduler action (string/int parameters may be
    appended after the -File path).

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "& 'E:\Scripts\Get-SasLogPerformance_v1.ps1' -EnableRobocopy $true -EmailEnabled $false"
    Use -Command style when overriding [bool] parameters from a scheduled task
    (-File passes booleans as literal text, which will not bind).

.NOTES
    Exit codes: 0 = success, 1 = completed with parse errors, 2 = fatal error.
#>
#region ===================== CONFIG / PARAMETERS =============================
[CmdletBinding()]
param(
    [string]   $LogRoot                   = '\\a70admed.com\R1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT',
    [string]   $InputDirectoriesCsv       = '',
    [string[]] $LogFileExtensions         = @('log', 'txt', 'sas'),
    [bool]     $IncludeSubdirectories     = $true,
    [string]   $OutputPath                = '',
    [string]   $ExcelSheetName            = 'SAS_Log_Inventory',
    [string]   $FolderExclusionCsv        = '',
    [bool]     $EnableRobocopy            = $false,
    [string]   $RobocopyConfigCsv         = '',
    [bool]     $EmailEnabled              = $true,
    [string[]] $EmailTo                   = @('dl-sas-admins@cgsadmin.com'),
    [string]   $EmailFrom                 = 'sas-log-crawler@cgsadmin.com',
    [string]   $SmtpServer                = 'smtprelay.bcbssc.com',
    [int]      $SmtpPort                  = 25,
    [string]   $ProgramNameSuffixPattern  = '_\d{3,5}(AM|PM)$',
    [string]   $ProgramNameVersionPattern = '_V\d+(_\d+)*$',
    [string[]] $ProgramNameTokens         = @()
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Continue'
$ProgressPreference    = 'SilentlyContinue'   # never draw progress bars unattended
$ConfirmPreference     = 'None'               # never prompt

# Resolve script-relative defaults (kept out of param() so $PSScriptRoot is safe
# under every host).
$script:ScriptDir = $PSScriptRoot
if (-not $script:ScriptDir) { $script:ScriptDir = (Get-Location).Path }

if (-not $InputDirectoriesCsv) { $InputDirectoriesCsv = Join-Path $script:ScriptDir 'input_directories.csv' }
if (-not $FolderExclusionCsv)  { $FolderExclusionCsv  = Join-Path $script:ScriptDir 'folder_exclusion_list.csv' }
if (-not $RobocopyConfigCsv)   { $RobocopyConfigCsv   = Join-Path $script:ScriptDir 'robocopy_config.csv' }
if (-not $OutputPath)          { $OutputPath          = Join-Path $script:ScriptDir 'sas_log_inventory.csv' }

$script:RunTimestamp  = Get-Date
$script:RunLogFile    = $null      # set once OutputPath's folder is known
$script:ExitCode      = 0
#endregion

#region ===================== LOGGING HELPERS =================================
function Write-Log {
    <# Timestamped console + rolling-file logger. Levels: INFO / WARN / ERROR. #>
    param(
        [Parameter(Position = 0)][AllowEmptyString()][string] $Message = '',
        [Parameter(Position = 1)][ValidateSet('INFO', 'WARN', 'ERROR')][string] $Level = 'INFO'
    )
    $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $line  = "[$stamp] [$Level] $Message"
    switch ($Level) {
        'ERROR' { Write-Host $line -ForegroundColor Red }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        default { Write-Host $line }
    }
    if ($script:RunLogFile) {
        try { Add-Content -LiteralPath $script:RunLogFile -Value $line -Encoding Default } catch { }
    }
}

function Initialize-RunLog {
    <# Creates/rolls Get-SasLogPerformance.log next to the output file (5 MB roll). #>
    param([Parameter(Mandatory = $true)][string] $NextToPath)
    try {
        $dir = Split-Path -Parent $NextToPath
        if (-not $dir) { $dir = $script:ScriptDir }
        if (-not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
        $logPath = Join-Path $dir 'Get-SasLogPerformance.log'
        if (Test-Path -LiteralPath $logPath) {
            $size = (Get-Item -LiteralPath $logPath).Length
            if ($size -gt 5MB) {
                $rolled = "$logPath.1"
                if (Test-Path -LiteralPath $rolled) { Remove-Item -LiteralPath $rolled -Force }
                Move-Item -LiteralPath $logPath -Destination $rolled -Force
            }
        }
        $script:RunLogFile = $logPath
    } catch {
        # Console-only logging if the file cannot be created; never fatal.
        $script:RunLogFile = $null
    }
}
#endregion

#region ===================== INPUT CONFIG READERS ============================
function Read-InputDirectories {
    <# Returns the list of root directories to crawl.
       Primary source: $InputDirectoriesCsv (column FULL_PATH_DIR);
       fallback: $LogRoot. #>
    param(
        [string] $CsvPath,
        [string] $FallbackRoot
    )
    $dirs = @()
    if ($CsvPath -and (Test-Path -LiteralPath $CsvPath)) {
        try {
            $rows = @(Import-Csv -LiteralPath $CsvPath)
            foreach ($row in $rows) {
                $p = $null
                if ($row.PSObject.Properties['FULL_PATH_DIR']) { $p = [string]$row.FULL_PATH_DIR }
                if ($p) { $p = $p.Trim() }
                if ($p) { $dirs += $p }
            }
            Write-Log "Read $($dirs.Count) root director$(if ($dirs.Count -eq 1) {'y'} else {'ies'}) from '$CsvPath'."
        } catch {
            Write-Log "Failed to read input directories CSV '$CsvPath': $($_.Exception.Message)" 'WARN'
        }
    } else {
        Write-Log "Input directories CSV not found at '$CsvPath'." 'WARN'
    }
    if ($dirs.Count -eq 0 -and $FallbackRoot) {
        Write-Log "Falling back to -LogRoot '$FallbackRoot'."
        $dirs = @($FallbackRoot)
    }
    return , $dirs
}

function Get-ExclusionList {
    <# Returns normalized excluded-folder prefixes (lower case, trailing
       separator appended) from folder_exclusion_list.csv. #>
    param([string] $CsvPath)
    $excl = @()
    if ($CsvPath -and (Test-Path -LiteralPath $CsvPath)) {
        try {
            $rows = @(Import-Csv -LiteralPath $CsvPath)
            foreach ($row in $rows) {
                $p = $null
                if ($row.PSObject.Properties['EXCLUDE_FOLDER']) { $p = [string]$row.EXCLUDE_FOLDER }
                if ($p) { $p = $p.Trim() }
                if ($p) {
                    $p = $p.TrimEnd('\', '/')
                    $excl += ($p + [System.IO.Path]::DirectorySeparatorChar).ToLowerInvariant()
                }
            }
            Write-Log "Loaded $($excl.Count) folder exclusion(s) from '$CsvPath'."
        } catch {
            Write-Log "Failed to read folder exclusion CSV '$CsvPath': $($_.Exception.Message)" 'WARN'
        }
    } else {
        Write-Log "No folder exclusion CSV found at '$CsvPath' (none applied)."
    }
    return , $excl
}

function Test-ExcludedPath {
    <# True when a file path sits under any excluded folder
       (case-insensitive, trailing-slash-safe). #>
    param(
        [Parameter(Mandatory = $true)][string]   $Path,
        [string[]] $Exclusions = @()
    )
    if (-not $Exclusions -or $Exclusions.Count -eq 0) { return $false }
    $sep  = [System.IO.Path]::DirectorySeparatorChar
    $norm = ($Path.TrimEnd('\', '/') + $sep).ToLowerInvariant()
    foreach ($prefix in $Exclusions) {
        if ($norm.StartsWith($prefix)) { return $true }
    }
    return $false
}
#endregion

#region ===================== ROBOCOPY STAGING ================================
function Invoke-LogCopy {
    <# CSV-driven robocopy staging (Copy-FilesWithRobocopy_v1_3 style).
       Columns: Source, Target, FileExtension, UpdateIfNewer, IncludeSubfolders, Notes.
       Args are built as an array - no Invoke-Expression. Exit codes 0-7 = success. #>
    param([Parameter(Mandatory = $true)][string] $ConfigCsv)

    if (-not (Test-Path -LiteralPath $ConfigCsv)) {
        Write-Log "Robocopy enabled but config CSV '$ConfigCsv' not found - skipping staging." 'WARN'
        return
    }
    $robo = Get-Command robocopy.exe -ErrorAction SilentlyContinue
    if (-not $robo) { $robo = Get-Command robocopy -ErrorAction SilentlyContinue }
    if (-not $robo) {
        Write-Log 'robocopy not found on this system - skipping staging.' 'WARN'
        return
    }

    $jobs = @(Import-Csv -LiteralPath $ConfigCsv)
    Write-Log "Robocopy staging: $($jobs.Count) job(s) from '$ConfigCsv'."
    foreach ($job in $jobs) {
        $src = [string]$job.Source
        $tgt = [string]$job.Target
        if (-not $src -or -not $tgt) {
            Write-Log 'Robocopy job skipped: Source or Target missing.' 'WARN'
            continue
        }
        $spec = ([string]$job.FileExtension).Trim()
        if (-not $spec) { $spec = '*.*' }
        elseif ($spec.IndexOf('*') -lt 0) { $spec = '*.' + $spec.TrimStart('.') }

        $roboArgs = @($src, $tgt, $spec, '/NP', '/NDL', '/R:2', '/W:5')
        if (('Y', 'YES', 'TRUE', '1') -contains ([string]$job.IncludeSubfolders).Trim().ToUpperInvariant()) {
            $roboArgs += '/E'
        }
        if (('Y', 'YES', 'TRUE', '1') -contains ([string]$job.UpdateIfNewer).Trim().ToUpperInvariant()) {
            $roboArgs += '/XO'   # copy only when source is newer than target
        }
        Write-Log "robocopy $($roboArgs -join ' ')"
        & $robo.Source @roboArgs | Out-Null
        $rc = $LASTEXITCODE
        if ($rc -le 7) {
            Write-Log "Robocopy job OK (exit $rc): '$src' -> '$tgt'  [$([string]$job.Notes)]"
        } else {
            Write-Log "Robocopy job FAILED (exit $rc): '$src' -> '$tgt'" 'ERROR'
        }
    }
}
#endregion

#region ===================== FILE DISCOVERY ==================================
function Get-LogFiles {
    <# Enumerates candidate log files under each root, applying the extension
       filter (case-insensitive) and folder exclusions. #>
    param(
        [string[]] $Roots,
        [string[]] $Extensions,
        [bool]     $Recurse,
        [string[]] $Exclusions
    )
    $wanted = @{}
    foreach ($ext in $Extensions) {
        if ($ext) { $wanted['.' + $ext.TrimStart('.').ToLowerInvariant()] = $true }
    }
    $files = New-Object System.Collections.Generic.List[object]
    foreach ($root in $Roots) {
        if (-not (Test-Path -LiteralPath $root)) {
            Write-Log "Root directory not found (skipped): '$root'" 'WARN'
            continue
        }
        Write-Log "Scanning '$root' (recurse=$Recurse)..."
        try {
            $gciParams = @{ LiteralPath = $root; File = $true; Force = $true; ErrorAction = 'SilentlyContinue' }
            if ($Recurse) { $gciParams['Recurse'] = $true }
            $found = @(Get-ChildItem @gciParams)
            foreach ($f in $found) {
                if (-not $wanted.ContainsKey($f.Extension.ToLowerInvariant())) { continue }
                if (Test-ExcludedPath -Path $f.FullName -Exclusions $Exclusions) {
                    Write-Log "Excluded by folder list: $($f.FullName)"
                    continue
                }
                $files.Add($f)
            }
        } catch {
            Write-Log "Error scanning '$root': $($_.Exception.Message)" 'ERROR'
        }
    }
    Write-Log "Discovered $($files.Count) log file(s) to parse."
    return , $files
}
#endregion

#region ===================== PARSING HELPERS =================================
function ConvertTo-Seconds {
    <# Normalizes a SAS timing string to decimal seconds.
       Handles: '0.06 seconds' | 'mm:ss.ff' (35:17.06) | 'h:mm:ss.ff'.
       Returns $null when unparseable. #>
    param([string] $TimeText)
    if (-not $TimeText) { return $null }
    $t   = $TimeText.Trim() -replace '\s*seconds?\s*$', ''
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    try {
        if ($t -match '^(\d+):([0-5]?\d):([0-5]?\d(\.\d+)?)$') {
            return [math]::Round(
                ([double]$Matches[1] * 3600) + ([double]$Matches[2] * 60) +
                [double]::Parse($Matches[3], $inv), 4)
        }
        if ($t -match '^(\d+):([0-5]?\d(\.\d+)?)$') {
            return [math]::Round(
                ([double]$Matches[1] * 60) + [double]::Parse($Matches[2], $inv), 4)
        }
        if ($t -match '^\d+(\.\d+)?$') {
            return [double]::Parse($t, $inv)
        }
    } catch { }
    return $null
}

function ConvertTo-LogDateTime {
    <# Best-effort parse of timestamps found in SAS logs ('Fri Jun 05 12:00:01 2026',
       '06/05/2026 12:35:18 PM', '05JUN2026:12:00:01', ...). Returns $null on failure. #>
    param([string] $Text)
    if (-not $Text) { return $null }
    $t = $Text.Trim()
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    $exact = @(
        'ddd MMM dd HH:mm:ss yyyy',
        'ddd MMM  d HH:mm:ss yyyy',
        'MM/dd/yyyy hh:mm:ss tt',
        'M/d/yyyy h:mm:ss tt',
        'ddMMMyyyy:HH:mm:ss',
        'ddMMMyyyy:HH:mm:ss.ff'
    )
    foreach ($fmt in $exact) {
        try { return [datetime]::ParseExact($t, $fmt, $inv) } catch { }
    }
    try { return [datetime]::Parse($t, $inv) } catch { }
    try { return [datetime]::Parse($t) } catch { }
    return $null
}

function Format-Ts {
    <# Uniform timestamp text for the output file (blank for $null). #>
    param($Value)
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd HH:mm:ss') }
    return ''
}

function Get-SasProgramName {
    <# Derives the SAS program name from a log filename:
       1) trim schedule suffix via $SuffixPattern  (…_0915AM -> …)
       2) trim version suffix via $VersionPattern  (…_V1_1   -> …)
       3) trim literal trailing tokens (Strip-Tokens_v3 style)
       4) trim trailing '_' and '.'
       Returns @{ Name = <string>; Source = 'SUFFIX_TRIMMED'|'FILENAME_UNTRIMMED' } #>
    param(
        [Parameter(Mandatory = $true)][string] $FileName,
        [string]   $SuffixPattern,
        [string]   $VersionPattern,
        [string[]] $Tokens = @()
    )
    $base    = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $name    = $base
    $trimmed = $false

    if ($SuffixPattern) {
        $candidate = $name -replace $SuffixPattern, ''
        if ($candidate -ne $name -and $candidate) { $name = $candidate; $trimmed = $true }
    }
    if ($VersionPattern) {
        $candidate = $name -replace $VersionPattern, ''
        if ($candidate -ne $name -and $candidate) { $name = $candidate; $trimmed = $true }
    }
    if ($Tokens) {
        foreach ($token in $Tokens) {
            if (-not $token) { continue }
            while ($name.Length -gt $token.Length -and
                   $name.EndsWith($token, [System.StringComparison]::OrdinalIgnoreCase)) {
                $name = $name.Substring(0, $name.Length - $token.Length)
                $trimmed = $true
            }
        }
    }
    $cleaned = $name.TrimEnd('_', '.')
    if ($cleaned) {
        if ($cleaned -ne $name) { $trimmed = $true }
        $name = $cleaned
    }
    if (-not $name) { $name = $base; $trimmed = $false }

    $source = 'FILENAME_UNTRIMMED'
    if ($trimmed) { $source = 'SUFFIX_TRIMMED' }
    return @{ Name = $name; Source = $source }
}

function New-InventoryRow {
    <# One fully-schema'd output row; all fields default to blank strings so the
       CSV/Excel column set is identical on every row. #>
    param([hashtable] $Values = @{})
    $row = [ordered]@{
        FULL_PATH               = ''
        FILENAME                = ''
        FOLDER                  = ''
        PARENT_FOLDER           = ''
        GRAND_PARENT_FOLDER     = ''
        FILE_EXTENSION          = ''
        FILE_SIZE_KB            = ''
        CREATED_TS              = ''
        LAST_UPDATED_TS         = ''
        LAST_ACCESSED_TS        = ''
        FILE_OWNER              = ''
        SAS_PROGRAM_NAME        = ''
        SAS_PROGRAM_NAME_SOURCE = ''
        LAST_RUN                = ''
        LAST_RUN_SOURCE         = ''
        STEP_SEQ                = ''
        STEP_TYPE               = ''
        STEP_NAME               = ''
        REAL_TIME_SEC           = ''
        CPU_TIME_SEC            = ''
        REAL_TIME_RAW           = ''
        CPU_TIME_RAW            = ''
        USER_CPU_TIME_SEC       = ''
        SYSTEM_CPU_TIME_SEC     = ''
        MEMORY_KB               = ''
        OS_MEMORY_KB            = ''
        STEP_COUNT              = ''
        SWITCH_COUNT            = ''
        LOG_OPENED_TS           = ''
        LOG_CLOSED_TS           = ''
        ERROR_COUNT             = ''
        WARNING_COUNT           = ''
        PARSE_STATUS            = ''
        RUN_TIMESTAMP           = $script:RunTimestamp.ToString('yyyy-MM-dd HH:mm:ss')
    }
    foreach ($key in $Values.Keys) { $row[$key] = $Values[$key] }
    return [pscustomobject]$row
}

function Get-FileMetadataValues {
    <# File-system metadata + derived program name for one log file. #>
    param([Parameter(Mandatory = $true)][System.IO.FileInfo] $File)

    $owner = ''
    try { $owner = (Get-Acl -LiteralPath $File.FullName -ErrorAction Stop).Owner } catch { $owner = '' }

    $folder     = ''
    $parentName = ''
    $grandName  = ''
    try {
        $dirInfo = $File.Directory
        if ($dirInfo) {
            $folder     = $dirInfo.FullName
            $parentName = $dirInfo.Name
            if ($dirInfo.Parent) { $grandName = $dirInfo.Parent.Name }
        }
    } catch { }

    $prog = Get-SasProgramName -FileName $File.Name `
                               -SuffixPattern  $ProgramNameSuffixPattern `
                               -VersionPattern $ProgramNameVersionPattern `
                               -Tokens         $ProgramNameTokens
    return @{
        FULL_PATH               = $File.FullName
        FILENAME                = $File.Name
        FOLDER                  = $folder
        PARENT_FOLDER           = $parentName
        GRAND_PARENT_FOLDER     = $grandName
        FILE_EXTENSION          = $File.Extension.TrimStart('.')
        FILE_SIZE_KB            = [math]::Round($File.Length / 1KB, 2)
        CREATED_TS              = Format-Ts $File.CreationTime
        LAST_UPDATED_TS         = Format-Ts $File.LastWriteTime
        LAST_ACCESSED_TS        = Format-Ts $File.LastAccessTime
        FILE_OWNER              = $owner
        SAS_PROGRAM_NAME        = $prog.Name
        SAS_PROGRAM_NAME_SOURCE = $prog.Source
    }
}
#endregion

#region ===================== SAS LOG PARSER ==================================
function Parse-SasLog {
    <# Parses one SAS log and returns the list of inventory rows for it
       (one per step; exactly one NO_STEPS_FOUND row for stub logs).
       Throws on unreadable files - the caller catches and emits PARSE_ERROR. #>
    param([Parameter(Mandatory = $true)][System.IO.FileInfo] $File)

    # --- read with Windows-1252 so ISO-8859-1 bytes (e.g. the (c) sign) never crash us
    $encoding = $null
    try { $encoding = [System.Text.Encoding]::GetEncoding(1252) } catch { }
    if (-not $encoding) { $encoding = [System.Text.Encoding]::Default }
    $reader = New-Object System.IO.StreamReader($File.FullName, $encoding, $true)
    $lines  = New-Object System.Collections.Generic.List[string]
    try {
        while ($null -ne ($line = $reader.ReadLine())) { $lines.Add($line) }
    } finally {
        $reader.Close()
    }

    # --- regexes for step markers and indented metric lines
    $rxProc     = '^NOTE: PROCEDURE (\S+) used \(Total process time\):'
    $rxData     = '^NOTE: DATA statement used \(Total process time\):'
    $rxInit     = '^NOTE: SAS initialization used:'
    $rxSession  = '^NOTE: The SAS System used:'
    $rxOpened   = '^NOTE: Log file opened at (.+?)\s*$'
    $rxClosed   = '^NOTE: Log file closing at (.+?)\s*$'
    $rxReal     = '^\s+real time\s+(\S.*?)\s*$'
    $rxUserCpu  = '^\s+user cpu time\s+(\S.*?)\s*$'
    $rxSysCpu   = '^\s+system cpu time\s+(\S.*?)\s*$'
    $rxCpu      = '^\s+cpu time\s+(\S.*?)\s*$'
    $rxMemory   = '^\s+memory\s+([\d.]+)\s*k\b'
    $rxOsMemory = '^\s+OS Memory\s+([\d.]+)\s*k\b'
    $rxTimestamp= '^\s+Timestamp\s+(\S.*?)\s*$'
    $rxCounts   = '^\s+Step Count\s+(\d+)\s+Switch Count\s+(\d+)'

    $steps        = New-Object System.Collections.Generic.List[object]
    $current      = $null
    $errorCount   = 0
    $warningCount = 0
    $openedRaw    = ''
    $closedRaw    = ''

    foreach ($line in $lines) {
        if ($line -match '^ERROR:')   { $errorCount++ }
        if ($line -match '^WARNING:') { $warningCount++ }

        if ($line -match $rxOpened) { if (-not $openedRaw) { $openedRaw = $Matches[1] }; $current = $null; continue }
        if ($line -match $rxClosed) { $closedRaw = $Matches[1]; $current = $null; continue }

        $newStep = $null
        if     ($line -match $rxProc)    { $newStep = @{ Type = 'PROCEDURE';     Name = $Matches[1].ToUpperInvariant() } }
        elseif ($line -match $rxData)    { $newStep = @{ Type = 'DATA_STEP';     Name = '' } }
        elseif ($line -match $rxInit)    { $newStep = @{ Type = 'INITIALIZATION';Name = '' } }
        elseif ($line -match $rxSession) { $newStep = @{ Type = 'SESSION_TOTAL'; Name = '' } }

        if ($newStep) {
            $current = @{
                Type = $newStep.Type; Name = $newStep.Name
                RealRaw = ''; CpuRaw = ''; UserCpuRaw = ''; SysCpuRaw = ''
                MemoryKb = ''; OsMemoryKb = ''; TimestampRaw = ''
                StepCount = ''; SwitchCount = ''
            }
            $steps.Add($current)
            continue
        }

        if ($null -ne $current) {
            # order matters: 'user cpu time'/'system cpu time' also match the bare 'cpu time' pattern
            if     ($line -match $rxReal)      { $current.RealRaw      = $Matches[1] }
            elseif ($line -match $rxUserCpu)   { $current.UserCpuRaw   = $Matches[1] }
            elseif ($line -match $rxSysCpu)    { $current.SysCpuRaw    = $Matches[1] }
            elseif ($line -match $rxCpu)       { $current.CpuRaw       = $Matches[1] }
            elseif ($line -match $rxMemory)    { $current.MemoryKb     = $Matches[1] }
            elseif ($line -match $rxOsMemory)  { $current.OsMemoryKb   = $Matches[1] }
            elseif ($line -match $rxTimestamp) { $current.TimestampRaw = $Matches[1] }
            elseif ($line -match $rxCounts)    { $current.StepCount    = $Matches[1]; $current.SwitchCount = $Matches[2] }
            elseif ($line.Trim() -ne '')       { $current = $null }   # block ended
        }
    }

    # --- log-level values shared by every row of this file
    $openedTs = ConvertTo-LogDateTime $openedRaw
    $closedTs = ConvertTo-LogDateTime $closedRaw
    $meta     = Get-FileMetadataValues -File $File

    $lastRun       = Format-Ts $File.LastWriteTime
    $lastRunSource = 'FILE_LAST_WRITE'
    if ($openedTs) { $lastRun = Format-Ts $openedTs; $lastRunSource = 'LOG_OPENED_NOTE' }
    elseif ($openedRaw) { $lastRun = $openedRaw; $lastRunSource = 'LOG_OPENED_NOTE' }

    $shared = @{
        LAST_RUN        = $lastRun
        LAST_RUN_SOURCE = $lastRunSource
        LOG_OPENED_TS   = $(if ($openedTs) { Format-Ts $openedTs } else { $openedRaw })
        LOG_CLOSED_TS   = $(if ($closedTs) { Format-Ts $closedTs } else { $closedRaw })
        ERROR_COUNT     = $errorCount
        WARNING_COUNT   = $warningCount
    }

    $rows = New-Object System.Collections.Generic.List[object]

    if ($steps.Count -eq 0) {
        # Stub log: one row, metadata populated, step/metric fields blank.
        $values = @{}
        foreach ($k in $meta.Keys)   { $values[$k] = $meta[$k] }
        foreach ($k in $shared.Keys) { $values[$k] = $shared[$k] }
        $values['PARSE_STATUS'] = 'NO_STEPS_FOUND'
        $rows.Add((New-InventoryRow -Values $values))
        return , $rows
    }

    $seq = 0
    foreach ($step in $steps) {
        $seq++
        $values = @{}
        foreach ($k in $meta.Keys)   { $values[$k] = $meta[$k] }
        foreach ($k in $shared.Keys) { $values[$k] = $shared[$k] }

        $values['PARSE_STATUS'] = 'OK'
        $values['STEP_SEQ']     = $seq
        $values['STEP_TYPE']    = $step.Type
        $values['STEP_NAME']    = $step.Name

        $realSec = ConvertTo-Seconds $step.RealRaw
        $values['REAL_TIME_RAW'] = $step.RealRaw
        if ($null -ne $realSec) { $values['REAL_TIME_SEC'] = $realSec }

        if ($step.Type -eq 'SESSION_TOTAL') {
            $userSec = ConvertTo-Seconds $step.UserCpuRaw
            $sysSec  = ConvertTo-Seconds $step.SysCpuRaw
            if ($null -ne $userSec) { $values['USER_CPU_TIME_SEC']   = $userSec }
            if ($null -ne $sysSec)  { $values['SYSTEM_CPU_TIME_SEC'] = $sysSec }
            $cpuParts = @()
            if ($null -ne $userSec) { $cpuParts += $userSec }
            if ($null -ne $sysSec)  { $cpuParts += $sysSec }
            if ($cpuParts.Count -gt 0) {
                $values['CPU_TIME_SEC'] = [math]::Round(($cpuParts | Measure-Object -Sum).Sum, 4)
            }
            $rawParts = @()
            if ($step.UserCpuRaw) { $rawParts += "user: $($step.UserCpuRaw)" }
            if ($step.SysCpuRaw)  { $rawParts += "system: $($step.SysCpuRaw)" }
            $values['CPU_TIME_RAW'] = ($rawParts -join '; ')
            $values['MEMORY_KB']    = $step.MemoryKb
            $values['OS_MEMORY_KB'] = $step.OsMemoryKb
            $values['STEP_COUNT']   = $step.StepCount
            $values['SWITCH_COUNT'] = $step.SwitchCount
            if ($step.TimestampRaw -and -not $closedRaw) {
                $sessionTs = ConvertTo-LogDateTime $step.TimestampRaw
                if ($sessionTs) { $values['LOG_CLOSED_TS'] = Format-Ts $sessionTs }
            }
        } else {
            $cpuSec = ConvertTo-Seconds $step.CpuRaw
            $values['CPU_TIME_RAW'] = $step.CpuRaw
            if ($null -ne $cpuSec) { $values['CPU_TIME_SEC'] = $cpuSec }
        }
        $rows.Add((New-InventoryRow -Values $values))
    }
    return , $rows
}
#endregion

#region ===================== OUTPUT WRITERS ==================================
function Export-InventoryExcelCom {
    <# Excel COM automation fallback (no ImportExcel module). Throws on failure. #>
    param(
        [Parameter(Mandatory = $true)][object[]] $Rows,
        [Parameter(Mandatory = $true)][string]   $Path,
        [Parameter(Mandatory = $true)][string]   $SheetName
    )
    $excel = $null
    $workbook = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Add()
        $sheet = $workbook.Worksheets.Item(1)
        $sheet.Name = $SheetName

        $columns = @($Rows[0].PSObject.Properties | ForEach-Object { $_.Name })
        $rowCount = $Rows.Count
        $colCount = $columns.Count

        # Bulk write through a 2-D array - orders of magnitude faster than per cell.
        $data = New-Object 'object[,]' ($rowCount + 1), $colCount
        for ($c = 0; $c -lt $colCount; $c++) { $data[0, $c] = $columns[$c] }
        for ($r = 0; $r -lt $rowCount; $r++) {
            for ($c = 0; $c -lt $colCount; $c++) {
                $data[($r + 1), $c] = [string]$Rows[$r].($columns[$c])
            }
        }
        $range = $sheet.Range($sheet.Cells.Item(1, 1), $sheet.Cells.Item($rowCount + 1, $colCount))
        $range.Value2 = $data
        $sheet.Rows.Item(1).Font.Bold = $true

        if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
        $workbook.SaveAs($Path, 51)   # 51 = xlOpenXMLWorkbook (.xlsx)
    } finally {
        if ($workbook) { $workbook.Close($false) | Out-Null }
        if ($excel)    { $excel.Quit() }
        foreach ($com in @($workbook, $excel)) {
            if ($com) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($com) }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Export-Inventory {
    <# Writes the inventory rows to CSV or Excel based on the OutputPath
       extension. Returns the path actually written. #>
    param(
        [Parameter(Mandatory = $true)][object[]] $Rows,
        [Parameter(Mandatory = $true)][string]   $Path,
        [string] $SheetName = 'SAS_Log_Inventory'
    )
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    $ext = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()

    if ($ext -eq '.xlsx') {
        # 1st choice: ImportExcel module
        $importExcel = Get-Module -ListAvailable -Name ImportExcel | Select-Object -First 1
        if ($importExcel) {
            try {
                Import-Module ImportExcel -ErrorAction Stop
                if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
                $Rows | Export-Excel -Path $Path -WorksheetName $SheetName -AutoSize -FreezeTopRow -BoldTopRow
                Write-Log "Wrote $($Rows.Count) row(s) to Excel '$Path' (sheet '$SheetName') via ImportExcel."
                return $Path
            } catch {
                Write-Log "ImportExcel export failed: $($_.Exception.Message)" 'WARN'
            }
        } else {
            Write-Log 'ImportExcel module not available; trying Excel COM automation.' 'WARN'
        }
        # 2nd choice: Excel COM automation
        try {
            Export-InventoryExcelCom -Rows $Rows -Path $Path -SheetName $SheetName
            Write-Log "Wrote $($Rows.Count) row(s) to Excel '$Path' (sheet '$SheetName') via Excel COM."
            return $Path
        } catch {
            Write-Log "Excel COM export failed: $($_.Exception.Message)" 'WARN'
        }
        # last resort: CSV alongside, with a warning
        $Path = [System.IO.Path]::ChangeExtension($Path, '.csv')
        Write-Log "Falling back to CSV output: '$Path'" 'WARN'
    } elseif ($ext -ne '.csv') {
        Write-Log "Unrecognized output extension '$ext' - writing CSV format to '$Path'." 'WARN'
    }

    $Rows | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
    Write-Log "Wrote $($Rows.Count) row(s) to CSV '$Path'."
    return $Path
}
#endregion

#region ===================== EMAIL NOTIFICATION ==============================
function Send-EmailAlert {
    <# Send-MailMessage via the corporate SMTP relay (SendEmailAlert.ps1 pattern).
       Failures are logged, never fatal. #>
    param(
        [Parameter(Mandatory = $true)][string]   $Subject,
        [Parameter(Mandatory = $true)][string]   $Body,
        [Parameter(Mandatory = $true)][string[]] $To,
        [Parameter(Mandatory = $true)][string]   $From,
        [Parameter(Mandatory = $true)][string]   $Server,
        [int] $Port = 25
    )
    try {
        Send-MailMessage -To $To -From $From -Subject $Subject -Body $Body `
                         -SmtpServer $Server -Port $Port -ErrorAction Stop
        Write-Log "Email sent to $($To -join ', '): $Subject"
        return $true
    } catch {
        Write-Log "Email send failed: $($_.Exception.Message)" 'WARN'
        return $false
    }
}

function Send-CompletionEmail {
    <# Builds and sends the success or failure summary email. #>
    param(
        [Parameter(Mandatory = $true)][bool] $Success,
        [string[]] $DirectoriesCrawled = @(),
        [int]      $FilesParsed  = 0,
        [int]      $RowsWritten  = 0,
        [string]   $FinalOutputPath = '',
        [string[]] $FailedFiles  = @(),
        [string]   $FatalMessage = ''
    )
    if (-not $EmailEnabled) {
        Write-Log 'Email notifications disabled (-EmailEnabled $false); skipping.'
        return
    }
    $host_ = $env:COMPUTERNAME
    if (-not $host_) { $host_ = [System.Environment]::MachineName }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('SAS Log Performance Crawler - run summary')
    [void]$sb.AppendLine('==========================================')
    [void]$sb.AppendLine("Run timestamp   : $($script:RunTimestamp.ToString('yyyy-MM-dd HH:mm:ss'))")
    [void]$sb.AppendLine("Host            : $host_")
    [void]$sb.AppendLine("Directories     : $($DirectoriesCrawled.Count)")
    foreach ($d in $DirectoriesCrawled) { [void]$sb.AppendLine("    $d") }
    [void]$sb.AppendLine("Log files parsed: $FilesParsed")
    [void]$sb.AppendLine("Rows written    : $RowsWritten")
    [void]$sb.AppendLine("Output path     : $FinalOutputPath")
    if ($FailedFiles.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("Files that FAILED to parse ($($FailedFiles.Count)):")
        foreach ($f in $FailedFiles) { [void]$sb.AppendLine("    $f") }
    }
    if (-not $Success -and $FatalMessage) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("FATAL ERROR: $FatalMessage")
    }

    if ($Success) {
        $subject = "SAS Log Performance Crawler - SUCCESS ($RowsWritten rows) - $($script:RunTimestamp.ToString('yyyy-MM-dd'))"
        if ($FailedFiles.Count -gt 0) {
            $subject = "SAS Log Performance Crawler - COMPLETED WITH $($FailedFiles.Count) PARSE ERROR(S) - $($script:RunTimestamp.ToString('yyyy-MM-dd'))"
        }
    } else {
        $subject = "SAS Log Performance Crawler - FAILED - $($script:RunTimestamp.ToString('yyyy-MM-dd'))"
    }
    [void](Send-EmailAlert -Subject $subject -Body $sb.ToString() -To $EmailTo `
                           -From $EmailFrom -Server $SmtpServer -Port $SmtpPort)
}
#endregion

#region ===================== MAIN ============================================
$directories = @()
$failedFiles = @()
$filesParsed = 0
$rowsWritten = 0
$finalOutput = $OutputPath

try {
    Initialize-RunLog -NextToPath $OutputPath
    Write-Log '================ SAS Log Performance Crawler starting ================'
    Write-Log "Output target: '$OutputPath' (sheet '$ExcelSheetName' if Excel)"

    # --- optional robocopy staging ------------------------------------------
    if ($EnableRobocopy) {
        Invoke-LogCopy -ConfigCsv $RobocopyConfigCsv
    }

    # --- resolve roots, exclusions, files ------------------------------------
    $directories = Read-InputDirectories -CsvPath $InputDirectoriesCsv -FallbackRoot $LogRoot
    if ($directories.Count -eq 0) {
        throw 'No input directories resolved (CSV empty/missing and no -LogRoot).'
    }
    $exclusions = Get-ExclusionList -CsvPath $FolderExclusionCsv
    $logFiles   = Get-LogFiles -Roots $directories -Extensions $LogFileExtensions `
                               -Recurse $IncludeSubdirectories -Exclusions $exclusions

    # --- parse every file; one bad log never kills the run -------------------
    $allRows = New-Object System.Collections.Generic.List[object]
    foreach ($file in $logFiles) {
        try {
            $rows = Parse-SasLog -File $file
            foreach ($r in $rows) { $allRows.Add($r) }
            $filesParsed++
            $status = 'OK'
            if ($rows.Count -eq 1 -and $rows[0].PARSE_STATUS -eq 'NO_STEPS_FOUND') { $status = 'NO_STEPS_FOUND' }
            Write-Log "Parsed '$($file.Name)': $($rows.Count) row(s) [$status]"
        } catch {
            $failedFiles += $file.FullName
            Write-Log "PARSE_ERROR '$($file.FullName)': $($_.Exception.Message)" 'ERROR'
            # Still emit a row so no file is ever silently skipped.
            $values = @{}
            try {
                $meta = Get-FileMetadataValues -File $file
                foreach ($k in $meta.Keys) { $values[$k] = $meta[$k] }
                $values['LAST_RUN']        = Format-Ts $file.LastWriteTime
                $values['LAST_RUN_SOURCE'] = 'FILE_LAST_WRITE'
            } catch {
                $values['FULL_PATH'] = $file.FullName
                $values['FILENAME']  = $file.Name
            }
            $values['PARSE_STATUS'] = 'PARSE_ERROR'
            $allRows.Add((New-InventoryRow -Values $values))
        }
    }

    # --- export ---------------------------------------------------------------
    if ($allRows.Count -eq 0) {
        Write-Log 'No rows produced (no matching log files found).' 'WARN'
    } else {
        $finalOutput = Export-Inventory -Rows $allRows.ToArray() -Path $OutputPath -SheetName $ExcelSheetName
    }
    $rowsWritten = $allRows.Count

    # --- summary + email --------------------------------------------------------
    if ($failedFiles.Count -gt 0) { $script:ExitCode = 1 }
    Write-Log "Run complete: $filesParsed file(s) parsed, $($failedFiles.Count) parse failure(s), $rowsWritten row(s) written to '$finalOutput'."
    Send-CompletionEmail -Success $true -DirectoriesCrawled $directories `
                         -FilesParsed $filesParsed -RowsWritten $rowsWritten `
                         -FinalOutputPath $finalOutput -FailedFiles $failedFiles
    Write-Log "Exiting with code $($script:ExitCode)."
} catch {
    $script:ExitCode = 2
    $msg = $_.Exception.Message
    Write-Log "FATAL: $msg" 'ERROR'
    Write-Log ($_.ScriptStackTrace | Out-String).Trim() 'ERROR'
    Send-CompletionEmail -Success $false -DirectoriesCrawled $directories `
                         -FilesParsed $filesParsed -RowsWritten $rowsWritten `
                         -FinalOutputPath $finalOutput -FailedFiles $failedFiles `
                         -FatalMessage $msg
    Write-Log 'Exiting with code 2.'
}
exit $script:ExitCode
#endregion
