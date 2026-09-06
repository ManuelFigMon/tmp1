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

# Bump this whenever a function is added, renamed or removed. The callers log
# it at startup, so a share holding a stale copy of this file is visible in the
# log instead of surfacing later as "the term X is not recognized".
#   1  original helpers
#   2  Assert-CgsWritable split into Test-CgsWritable + Resolve-CgsWritableTarget
#   3  Write-CgsXlsx added -- native multi-sheet .xlsx, no ImportExcel needed
#   4  Read-CgsXlsxSheet added -- native .xlsx reading, for formatData
$script:CgsUtilsApi = 4

function Get-CgsUtilsBanner {
    <# .SYNOPSIS One line describing the cgsUtils.ps1 that actually loaded.
       .OUTPUTS  [string] version, API level and the file's timestamp, so two
                 machines running different copies can be told apart. #>
    $path = Join-Path $PSScriptRoot 'cgsUtils.ps1'
    $stamp = if (Test-Path -LiteralPath $path) {
        (Get-Item -LiteralPath $path).LastWriteTime.ToString('yyyy-MM-dd HH:mm')
    } else { 'unknown' }
    return ('cgsUtils {0} api={1} ({2}, modified {3})' -f
            $script:CgsVersion, $script:CgsUtilsApi, $path, $stamp)
}

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

function Test-CgsWritable {
    <# .SYNOPSIS Probe whether a path can be opened for writing.
       .PARAMETER Target  The candidate output file.
       .OUTPUTS [string] the reason it cannot be written, or $null when it can.
       .NOTES  FileShare Read matches what StreamWriter itself opens with, so
               the probe never rejects a file the write would have accepted. #>
    param([string] $Target)
    $parent = Split-Path -Parent $Target
    if ($parent -and -not (Test-Path -LiteralPath $parent)) {
        [void](New-Item -ItemType Directory -Path $parent -Force)
    }
    $existed = Test-Path -LiteralPath $Target
    try {
        $probe = [System.IO.File]::Open(
            $Target, [System.IO.FileMode]::OpenOrCreate,
            [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read)
        $probe.Close()
    }
    catch {
        return $_.Exception.Message
    }
    # Do not leave a stray empty file behind when the probe created it.
    if (-not $existed) {
        Remove-Item -LiteralPath $Target -Force -ErrorAction SilentlyContinue
    }
    return $null
}

function Resolve-CgsWritableTarget {
    <# .SYNOPSIS Return a path the caller can actually write to.
       .PARAMETER Target  The requested output file.
       .OUTPUTS [string] the requested path, or a timestamped sibling when it
                is locked. Throws only when the fallback fails too.
       .NOTES  A crawl of a large share takes minutes. Throwing that away
               because somebody left the workbook open in Excel is the wrong
               trade, so a locked target is downgraded to a timestamped name
               and announced loudly rather than ending the run. #>
    param([string] $Target)
    $reason = Test-CgsWritable -Target $Target
    if ($null -eq $reason) { return $Target }

    $folder = [System.IO.Path]::GetDirectoryName($Target)
    $stem   = [System.IO.Path]::GetFileNameWithoutExtension($Target)
    $ext    = [System.IO.Path]::GetExtension($Target)
    $fallback = Join-Path $folder ("{0}_{1}{2}" -f $stem, (Get-CgsTimestampSuffix), $ext)

    Write-CgsWarn ("cannot write '{0}': {1}" -f $Target, $reason)
    Write-CgsWarn ("the file is most likely open in Excel; writing '{0}' instead" -f $fallback)

    $second = Test-CgsWritable -Target $fallback
    if ($null -ne $second) {
        $template = "cannot write '{0}' ({1}) and the fallback '{2}' is not " +
                    "writable either ({3}). Check the folder exists and you " +
                    "have permission to write to it."
        throw ($template -f $Target, $reason, $fallback, $second)
    }
    return $fallback
}

function Assert-CgsWritable {
    <# .SYNOPSIS Backward-compatible shim for callers written before the split.
       .PARAMETER Target  The requested output file.
       .DESCRIPTION This function was replaced by Test-CgsWritable and
                    Resolve-CgsWritableTarget. It is kept because the .ps1
                    files are copied to a share individually: a scanner copied
                    before the rename would otherwise die with "the term
                    'Assert-CgsWritable' is not recognized", which says nothing
                    about the real problem. Old behaviour is preserved -- it
                    returns nothing and throws when the target is locked --
                    because a stale caller ignores return values, so it must
                    not be handed a fallback path it would silently discard.
       .NOTES  DEPRECATED. New code calls Resolve-CgsWritableTarget, which
               falls back to a timestamped sibling instead of throwing. #>
    param([string] $Target)
    $reason = Test-CgsWritable -Target $Target
    if ($null -eq $reason) { return }
    $template = "cannot write '{0}': {1} The file is most likely open in " +
                "Excel. Close it, or re-copy every file in src\ps to the " +
                "share together -- this script predates {2}."
    throw ($template -f $Target, $reason, (Get-CgsUtilsBanner))
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
    try {
        $writer = New-Object System.IO.StreamWriter($Target, $false, $encoding)
    }
    catch {
        $template = "cannot write '{0}': {1} If the file is open in Excel " +
                    "or another program, close it and run again."
        throw ($template -f $Target, $_.Exception.Message)
    }
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

function ConvertTo-CgsXmlText {
    <# .SYNOPSIS Escape a value for XML text content.
       .OUTPUTS [string] with & < > escaped and control characters dropped.
       .NOTES  Control characters are illegal in XML 1.0 and Excel rejects the
               whole workbook rather than the one bad cell. #>
    param($Value)
    if ($null -eq $Value) { return '' }
    $text = [regex]::Replace([string]$Value, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')
    return $text.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
}

function Get-CgsColumnRef {
    <# .SYNOPSIS Convert a 1-based column number to letters (1 -> A, 27 -> AA).
       .OUTPUTS [string] the column reference. #>
    param([int] $Index)
    $ref = ''
    while ($Index -gt 0) {
        $remainder = ($Index - 1) % 26
        $ref = [char](65 + $remainder) + $ref
        $Index = [int](($Index - $remainder - 1) / 26)
    }
    return $ref
}

function Get-CgsSafeSheetName {
    <# .SYNOPSIS Make a worksheet name Excel accepts, unique in the workbook.
       .PARAMETER Name  The requested name.
       .PARAMETER Used  Names already taken.
       .OUTPUTS [string] at most 31 characters, none of []:*?/\, never empty. #>
    param([string] $Name, [string[]] $Used = @())
    $cleaned = [regex]::Replace([string]$Name, '[\[\]:*?/\\]', '_').Trim("'")
    if ($cleaned.Length -gt 31) { $cleaned = $cleaned.Substring(0, 31) }
    if (-not $cleaned) { $cleaned = 'Sheet' }
    $candidate = $cleaned; $suffix = 2
    while ($Used -contains $candidate) {
        $tail = "_$suffix"
        $head = if ($cleaned.Length -gt (31 - $tail.Length)) {
            $cleaned.Substring(0, 31 - $tail.Length) } else { $cleaned }
        $candidate = $head + $tail
        $suffix++
    }
    return $candidate
}

function New-CgsXlsxSheetXml {
    <# .SYNOPSIS Build one xl/worksheets/sheetN.xml.
       .PARAMETER Columns  Column order; also the header row.
       .PARAMETER Rows     Objects or hashtables keyed by column name.
       .OUTPUTS [string] the worksheet part.
       .NOTES  The child order is fixed by the schema -- dimension, sheetViews,
               sheetFormatPr, cols, sheetData, autoFilter. Excel refuses the
               file when they are out of order, with no useful message. #>
    param([string[]] $Columns, [object[]] $Rows)
    if (-not $Columns -or $Columns.Count -eq 0) { $Columns = @('(no columns)') }
    $lastCol = Get-CgsColumnRef -Index $Columns.Count
    $lastRow = $Rows.Count + 1

    $sb = New-Object System.Text.StringBuilder
    [void] $sb.Append('<row r="1">')
    for ($i = 1; $i -le $Columns.Count; $i++) {
        [void] $sb.Append('<c r="' + (Get-CgsColumnRef -Index $i) + '1" s="1" t="inlineStr"><is><t xml:space="preserve">' + (ConvertTo-CgsXmlText $Columns[$i - 1]) + '</t></is></c>')
    }
    [void] $sb.Append('</row>')

    $rowNumber = 2
    foreach ($row in $Rows) {
        [void] $sb.Append('<row r="' + $rowNumber + '">')
        for ($i = 1; $i -le $Columns.Count; $i++) {
            $value = ConvertTo-CgsXmlText $row.($Columns[$i - 1])
            [void] $sb.Append('<c r="' + (Get-CgsColumnRef -Index $i) + $rowNumber + '" s="0" t="inlineStr"><is><t xml:space="preserve">' + $value + '</t></is></c>')
        }
        [void] $sb.Append('</row>')
        $rowNumber++
    }

    # Column widths, sampled from the first 200 rows for speed.
    $sample = if ($Rows.Count -gt 200) { $Rows[0..199] } else { $Rows }
    $cols = New-Object System.Text.StringBuilder
    for ($i = 1; $i -le $Columns.Count; $i++) {
        $name = $Columns[$i - 1]
        $widest = $name.Length
        foreach ($row in $sample) {
            $length = ([string] $row.$name).Length
            if ($length -gt $widest) { $widest = $length }
        }
        $width = [Math]::Min(60, [Math]::Max(10, $widest + 2))
        [void] $cols.Append('<col min="' + $i + '" max="' + $i + '" width="' + $width + '" customWidth="1"/>')
    }

    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<dimension ref="A1:' + $lastCol + $lastRow + '"/>' +
      '<sheetViews><sheetView workbookViewId="0">' +
      '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' +
      '</sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>' +
      '<cols>' + $cols.ToString() + '</cols>' +
      '<sheetData>' + $sb.ToString() + '</sheetData>' +
      '<autoFilter ref="A1:' + $lastCol + $lastRow + '"/>' +
      '</worksheet>'
}

function Write-CgsXlsx {
    <# .SYNOPSIS Write a multi-sheet .xlsx with no third-party module.
       .DESCRIPTION An .xlsx is a zip of XML parts, written here directly with
                    System.IO.Compression so nothing beyond .NET is needed.
                    Values are inline strings, so a claim number keeps its
                    leading zeros instead of being turned into a number.
       .PARAMETER Sheets  Array of hashtables: Name, Columns, Rows.
       .PARAMETER Path    Destination .xlsx.
       .OUTPUTS [string] the path written. Throws on I/O failure.
       .NOTES  Mirrors writeWorkbook in src/utils/xlsx.py; the two produce the
               same XML, and a test compares their structure. #>
    param([object[]] $Sheets, [string] $Path)
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    # Excel rejects a workbook with no sheets, so always write at least one.
    if (-not $Sheets -or $Sheets.Count -eq 0) {
        $Sheets = @(@{ Name = 'Sheet1'; Columns = @(); Rows = @() })
    }
    $prepared = New-Object System.Collections.Generic.List[object]
    foreach ($sheet in $Sheets) {
        $used = @($prepared | ForEach-Object { $_.Name })
        $prepared.Add(@{
            Name    = (Get-CgsSafeSheetName -Name $sheet.Name -Used $used)
            Columns = @($sheet.Columns)
            Rows    = @($sheet.Rows)
        })
    }

    $sheetTags = ''; $relTags = ''; $overrides = ''
    for ($i = 1; $i -le $prepared.Count; $i++) {
        $sheetTags += '<sheet name="' + (ConvertTo-CgsXmlText $prepared[$i-1].Name) + '" sheetId="' + $i + '" r:id="rId' + $i + '"/>'
        $relTags   += '<Relationship Id="rId' + $i + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + $i + '.xml"/>'
        $overrides += '<Override PartName="/xl/worksheets/sheet' + $i + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    }
    $stylesId = $prepared.Count + 1

    $stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>' +
      '<font><b/><sz val="11"/><name val="Calibri"/></font></fonts>' +
      # Fill 0 MUST be none and fill 1 MUST be gray125 -- Excel rejects the
      # workbook otherwise, however unused those two entries are.
      '<fills count="3"><fill><patternFill patternType="none"/></fill>' +
      '<fill><patternFill patternType="gray125"/></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/><bgColor indexed="64"/></patternFill></fill></fills>' +
      '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="2">' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
      '</cellXfs>' +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '</styleSheet>'

    $parts = [ordered] @{
        '[Content_Types].xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
            '<Default Extension="xml" ContentType="application/xml"/>' +
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
            $overrides +
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
        '_rels/.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        'xl/workbook.xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
            '<sheets>' + $sheetTags + '</sheets></workbook>'
        'xl/_rels/workbook.xml.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            $relTags +
            '<Relationship Id="rId' + $stylesId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
        'xl/styles.xml' = $stylesXml
    }
    for ($i = 1; $i -le $prepared.Count; $i++) {
        $parts['xl/worksheets/sheet' + $i + '.xml'] = New-CgsXlsxSheetXml `
            -Columns $prepared[$i-1].Columns -Rows $prepared[$i-1].Rows
    }

    $parent = Split-Path -Parent $Path
    if ($parent -and -not (Test-Path -LiteralPath $parent)) {
        [void](New-Item -ItemType Directory -Path $parent -Force)
    }
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }

    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::CreateNew)
    try {
        $archive = New-Object System.IO.Compression.ZipArchive(
            $stream, [System.IO.Compression.ZipArchiveMode]::Create, $true)
        try {
            # UTF-8 with NO byte-order mark: Excel rejects a BOM inside parts.
            $encoding = New-Object System.Text.UTF8Encoding($false)
            foreach ($name in $parts.Keys) {
                $entry = $archive.CreateEntry($name,
                    [System.IO.Compression.CompressionLevel]::Optimal)
                $entryStream = $entry.Open()
                try {
                    $bytes = $encoding.GetBytes($parts[$name])
                    $entryStream.Write($bytes, 0, $bytes.Length)
                } finally { $entryStream.Dispose() }
            }
        } finally { $archive.Dispose() }
    } finally { $stream.Dispose() }
    return $Path
}

$script:XlsxMainNs = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$script:XlsxRelNs  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$script:XlsxPkgNs  = 'http://schemas.openxmlformats.org/package/2006/relationships'

function Get-CgsColumnIndex {
    <# .SYNOPSIS Convert a cell reference to a 0-based column index (AA3 -> 26).
       .OUTPUTS [int] 0-based index; 0 when the reference is unusable. #>
    param([string] $CellRef)
    $letters = [regex]::Match([string]$CellRef.ToUpperInvariant(), '^([A-Z]+)')
    if (-not $letters.Success) { return 0 }
    $index = 0
    foreach ($character in $letters.Groups[1].Value.ToCharArray()) {
        $index = $index * 26 + ([int][char]$character - 64)
    }
    return $index - 1
}

function Read-CgsXlsxGrid {
    <# .SYNOPSIS Read one worksheet as a grid of strings, with no add-in.
       .PARAMETER Path   The .xlsx to read.
       .PARAMETER Sheet  Worksheet name; default the first sheet.
       .OUTPUTS [object[]] array of string arrays, all the same length.
       .NOTES  Values come back as STRINGS on purpose: a claim number stored
               as text keeps its leading zeros, and this feeds a formatter,
               not a calculation. Mirrors readSheetRows in src/utils/xlsx.py. #>
    param([string] $Path, [string] $Sheet = '')
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        function Read-Entry([string] $name) {
            $entry = $archive.GetEntry($name)
            if ($null -eq $entry) { return $null }
            $reader = New-Object System.IO.StreamReader($entry.Open())
            try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
        }

        # Shared strings. A string can be split across several runs, so every
        # <t> under an <si> is joined -- taking the first would truncate any
        # cell Excel happened to style mid-word.
        $strings = New-Object System.Collections.Generic.List[string]
        $sharedXml = Read-Entry 'xl/sharedStrings.xml'
        if ($sharedXml) {
            $shared = New-Object System.Xml.XmlDocument
            $shared.LoadXml($sharedXml)
            foreach ($si in $shared.GetElementsByTagName('si', $script:XlsxMainNs)) {
                $text = ''
                foreach ($t in $si.GetElementsByTagName('t', $script:XlsxMainNs)) { $text += $t.InnerText }
                $strings.Add($text)
            }
        }

        # Worksheet name -> part, in workbook order.
        $relations = @{}
        $relsXml = Read-Entry 'xl/_rels/workbook.xml.rels'
        if ($relsXml) {
            $rels = New-Object System.Xml.XmlDocument
            $rels.LoadXml($relsXml)
            foreach ($node in $rels.GetElementsByTagName('Relationship', $script:XlsxPkgNs)) {
                $target = $node.GetAttribute('Target')
                if ($target.StartsWith('/')) { $relations[$node.GetAttribute('Id')] = $target.Substring(1) }
                else { $relations[$node.GetAttribute('Id')] = "xl/$target" }
            }
        }
        $workbook = New-Object System.Xml.XmlDocument
        $workbook.LoadXml((Read-Entry 'xl/workbook.xml'))
        $sheets = New-Object System.Collections.Generic.List[object]
        $ordinal = 1
        foreach ($node in $workbook.GetElementsByTagName('sheet', $script:XlsxMainNs)) {
            $id = $node.GetAttribute('id', $script:XlsxRelNs)
            $part = if ($relations.ContainsKey($id)) { $relations[$id] } else { "xl/worksheets/sheet$ordinal.xml" }
            $sheets.Add(@{ Name = $node.GetAttribute('name'); Path = $part })
            $ordinal++
        }
        if ($sheets.Count -eq 0) { return , @() }

        if (-not $Sheet) { $chosen = $sheets[0] }
        else {
            $chosen = $sheets | Where-Object { $_.Name -eq $Sheet } | Select-Object -First 1
            if ($null -eq $chosen) {
                throw ("worksheet '{0}' not found; this workbook has: {1}" -f
                       $Sheet, (($sheets | ForEach-Object { $_.Name }) -join ', '))
            }
        }
        $sheetDoc = New-Object System.Xml.XmlDocument
        $sheetDoc.LoadXml((Read-Entry $chosen.Path))
    } finally { $archive.Dispose() }

    $grid = New-Object System.Collections.Generic.List[object]
    foreach ($rowNode in $sheetDoc.GetElementsByTagName('row', $script:XlsxMainNs)) {
        $cells = New-Object System.Collections.Generic.List[string]
        foreach ($cellNode in $rowNode.GetElementsByTagName('c', $script:XlsxMainNs)) {
            $index = Get-CgsColumnIndex -CellRef $cellNode.GetAttribute('r')
            while ($cells.Count -lt $index) { $cells.Add('') }   # sparse rows
            $kind = $cellNode.GetAttribute('t')
            $text = ''
            if ($kind -eq 's') {
                $valueNode = $cellNode.GetElementsByTagName('v', $script:XlsxMainNs)
                if ($valueNode.Count -gt 0) {
                    $position = [int] $valueNode[0].InnerText
                    if ($position -ge 0 -and $position -lt $strings.Count) { $text = $strings[$position] }
                }
            }
            elseif ($kind -eq 'inlineStr') {
                foreach ($t in $cellNode.GetElementsByTagName('t', $script:XlsxMainNs)) { $text += $t.InnerText }
            }
            else {
                $valueNode = $cellNode.GetElementsByTagName('v', $script:XlsxMainNs)
                if ($valueNode.Count -gt 0) { $text = $valueNode[0].InnerText }
            }
            $cells.Add($text)
        }
        $grid.Add($cells.ToArray())
    }

    $width = 0
    foreach ($row in $grid) { if ($row.Count -gt $width) { $width = $row.Count } }
    $padded = New-Object System.Collections.Generic.List[object]
    foreach ($row in $grid) {
        $list = New-Object System.Collections.Generic.List[string]
        $list.AddRange([string[]]$row)
        while ($list.Count -lt $width) { $list.Add('') }
        $padded.Add($list.ToArray())
    }
    return , $padded.ToArray()
}

function Get-CgsHeaderRow {
    <# .SYNOPSIS Work out which row of a grid holds the column headers.
       .OUTPUTS [int] 1-based header row.
       .NOTES  A workbook formatData produced has a merged TITLE BANNER on
               row 1 and the headers on row 2 -- a first row with a single
               filled cell above a much wider second row. Detecting that lets
               a formatted workbook be fed straight back in. #>
    param([object[]] $Grid)
    if ($Grid.Count -lt 2) { return 1 }
    $filledFirst  = @($Grid[0] | Where-Object { "$_".Trim() }).Count
    $filledSecond = @($Grid[1] | Where-Object { "$_".Trim() }).Count
    if ($filledFirst -le 1 -and $filledSecond -gt 1) { return 2 }
    return 1
}

function Read-CgsXlsxSheet {
    <# .SYNOPSIS Read one worksheet as objects keyed by column header.
       .PARAMETER Path       The .xlsx to read.
       .PARAMETER Sheet      Worksheet name; default the first.
       .PARAMETER HeaderRow  1-based header row; 0 (default) detects it.
       .OUTPUTS [object[]] PSCustomObjects, shaped like Import-Csv output so
                either input can feed the same formatter. #>
    param([string] $Path, [string] $Sheet = '', [int] $HeaderRow = 0)
    $grid = @(Read-CgsXlsxGrid -Path $Path -Sheet $Sheet)
    if ($grid.Count -eq 0) { return , @() }
    $header = if ($HeaderRow -gt 0) { $HeaderRow } else { Get-CgsHeaderRow -Grid $grid }
    if ($header -lt 1) { $header = 1 }
    if ($header -gt $grid.Count) { $header = $grid.Count }

    $columns = New-Object System.Collections.Generic.List[string]
    $position = 1
    foreach ($name in $grid[$header - 1]) {
        $clean = "$name".Trim()
        # An unnamed trailing column would collide on the empty-string key.
        if (-not $clean) { $clean = "Column$position" }
        $columns.Add($clean)
        $position++
    }

    $rows = New-Object System.Collections.Generic.List[object]
    for ($i = $header; $i -lt $grid.Count; $i++) {
        $record = [ordered] @{}
        for ($c = 0; $c -lt $columns.Count; $c++) {
            $record[$columns[$c]] = if ($c -lt $grid[$i].Count) { $grid[$i][$c] } else { '' }
        }
        $rows.Add([PSCustomObject] $record)
    }
    return , $rows.ToArray()
}

function Get-CgsTimestampSuffix {
    <# .SYNOPSIS Current time as yyyyMMdd_HHmmss for generated filenames. #>
    return (Get-Date).ToString('yyyyMMdd_HHmmss')
}
