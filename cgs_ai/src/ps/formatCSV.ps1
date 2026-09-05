<#
=====================================================================
  Program Name  : formatCSV.ps1
  Author        : Manuel Figallo
  Purpose       : Convert a CSV into a styled Excel workbook with a SAS
                  ODS-style look and feel.
  Version       : 1.1beta
  Created       : 2026-08-26
  Last Modified : 2026-08-28

  Dependencies:
    NONE that must be installed. ImportExcel is USED WHEN PRESENT because
    it is well tested; when it is absent the script writes the .xlsx
    itself using only .NET types that ship with Windows PowerShell, so a
    locked-down server with no PSGallery access still produces a workbook.

  Description:
    PowerShell twin of src/py/formatCSV.py. The default "corporate" style is
    a navy title banner, blue header row and zebra striping.

  Input Parameters (required first):
    -InputCsvPath (REQUIRED)  -OutputExcelPath (REQUIRED)
    -FormatType (default corporate; corporate|corporatev2|plain|minimal)
    -SheetName (default Report)  -Title (defaults to the CSV name)
    -Writer (auto|native|module) -- force a writer; auto prefers ImportExcel.
  Exit codes: 0 = success, 2 = config error, 3 = I/O error.

  Change Log:
    v1.1beta - Added the native OOXML writer so a missing ImportExcel module
               is no longer a fatal error, and a -Writer switch to choose.
=====================================================================
#>
[CmdletBinding()]
param(
    [string] $InputCsvPath    = '',
    [string] $OutputExcelPath = '',
    [string] $FormatType      = 'corporate',
    [string] $SheetName       = 'Report',
    [string] $Title           = '',
    [string] $Writer          = 'auto'
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\cgsUtils.ps1"

# Palette per format type -- matches FORMAT_STYLES in the Python twin.
# Hex WITHOUT the leading '#'; the module writer adds it, the native writer
# prefixes the ARGB alpha byte instead.
$script:FormatStyles = @{
    'corporate'   = @{ Banner='1F3864'; Header='2E75B6'; Stripe='DCE6F1'; HeaderFont='FFFFFF'; BannerFont='FFFFFF' }
    'corporatev2' = @{ Banner='1F3864'; Header='2E75B6'; Stripe='DCE6F1'; HeaderFont='FFFFFF'; BannerFont='FFFFFF' }
    'plain'       = @{ Banner='FFFFFF'; Header='D9D9D9'; Stripe='FFFFFF'; HeaderFont='000000'; BannerFont='000000' }
    'minimal'     = @{ Banner='FFFFFF'; Header='FFFFFF'; Stripe='FFFFFF'; HeaderFont='000000'; BannerFont='000000' }
}

# Format types that get zebra striping -- matches STRIPED_FORMATS in Python.
$script:StripedFormats = @('corporate', 'corporatev2')

$script:MaxColumnWidth = 60

function ConvertTo-XmlText {
    <#
    .SYNOPSIS Escape a value for XML text content.
    .PARAMETER Value The value to escape; $null becomes an empty string.
    .OUTPUTS [string] with &, < and > escaped and control characters dropped.
    #>
    param([Parameter(ValueFromPipeline = $true)] $Value)
    if ($null -eq $Value) { return '' }
    $text = [string] $Value
    # Control characters are illegal in XML 1.0 and Excel refuses the file.
    $text = [regex]::Replace($text, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')
    return $text.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
}

function Get-ColumnRef {
    <#
    .SYNOPSIS Convert a 1-based column number to its letters (1 -> A, 27 -> AA).
    .PARAMETER Index 1-based column number.
    .OUTPUTS [string] the column reference.
    #>
    param([int] $Index)
    $ref = ''
    while ($Index -gt 0) {
        $remainder = ($Index - 1) % 26
        $ref = [char](65 + $remainder) + $ref
        $Index = [int](($Index - $remainder - 1) / 26)
    }
    return $ref
}

function New-XlsxStylesXml {
    <#
    .SYNOPSIS Build xl/styles.xml for the chosen palette.
    .PARAMETER Style Hashtable from $script:FormatStyles.
    .OUTPUTS [string] the styles part.
    .NOTES cellXfs indexes used by the sheet: 1 banner, 2 header, 3 striped
           data, 4 plain data. Fill 0 MUST be none and fill 1 MUST be
           gray125 -- Excel rejects the workbook otherwise.
    #>
    param([hashtable] $Style)
    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="3"><font><sz val="11"/><name val="Calibri"/></font>
<font><b/><sz val="14"/><color rgb="FF$($Style.BannerFont)"/><name val="Calibri"/></font>
<font><b/><sz val="11"/><color rgb="FF$($Style.HeaderFont)"/><name val="Calibri"/></font></fonts>
<fills count="5"><fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF$($Style.Banner)"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF$($Style.Header)"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF$($Style.Stripe)"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="5">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center"/></xf>
<xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
"@
}

function Write-XlsxNative {
    <#
    .SYNOPSIS Write a styled .xlsx with no third-party module.
    .DESCRIPTION An .xlsx is a zip of XML parts. This writes those parts
                 directly using System.IO.Compression, so it needs nothing
                 beyond .NET. Values are written as inline strings, which
                 avoids a shared-string table and keeps every cell exactly
                 as it appeared in the CSV.
    .PARAMETER Rows Objects from Import-Csv.
    .PARAMETER Columns Ordered column names.
    .PARAMETER Style Hashtable from $script:FormatStyles.
    .PARAMETER Striped Whether data rows alternate colour.
    .PARAMETER Banner Row 1 text.
    .PARAMETER Path Destination .xlsx.
    .PARAMETER Sheet Worksheet name.
    .OUTPUTS None. Throws on I/O failure.
    #>
    param(
        [object[]] $Rows, [string[]] $Columns, [hashtable] $Style,
        [bool] $Striped, [string] $Banner, [string] $Path, [string] $Sheet
    )
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $columnCount = [Math]::Max(1, $Columns.Count)
    $lastCol = Get-ColumnRef -Index $columnCount
    $lastRow = $Rows.Count + 2

    $sb = New-Object System.Text.StringBuilder

    # Row 1: the banner. Every cell in the merge range carries the banner
    # style so the fill spans the whole width in every reader.
    [void] $sb.Append('<row r="1" ht="20" customHeight="1">')
    for ($i = 1; $i -le $columnCount; $i++) {
        $text = if ($i -eq 1) { ConvertTo-XmlText $Banner } else { '' }
        [void] $sb.Append('<c r="' + (Get-ColumnRef -Index $i) + '1" s="1" t="inlineStr"><is><t xml:space="preserve">' + $text + '</t></is></c>')
    }
    [void] $sb.Append('</row>')

    # Row 2: the column headers.
    [void] $sb.Append('<row r="2">')
    for ($i = 1; $i -le $Columns.Count; $i++) {
        [void] $sb.Append('<c r="' + (Get-ColumnRef -Index $i) + '2" s="2" t="inlineStr"><is><t xml:space="preserve">' + (ConvertTo-XmlText $Columns[$i - 1]) + '</t></is></c>')
    }
    [void] $sb.Append('</row>')

    # Rows 3+: the data, striped on odd rows to match the Python twin.
    $rowNumber = 3
    foreach ($row in $Rows) {
        $cellStyle = if ($Striped -and ($rowNumber % 2 -eq 1)) { '3' } else { '4' }
        [void] $sb.Append('<row r="' + $rowNumber + '">')
        for ($i = 1; $i -le $Columns.Count; $i++) {
            $value = ConvertTo-XmlText $row.($Columns[$i - 1])
            [void] $sb.Append('<c r="' + (Get-ColumnRef -Index $i) + $rowNumber + '" s="' + $cellStyle + '" t="inlineStr"><is><t xml:space="preserve">' + $value + '</t></is></c>')
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
        $width = [Math]::Min($script:MaxColumnWidth, [Math]::Max(10, $widest + 2))
        [void] $cols.Append('<col min="' + $i + '" max="' + $i + '" width="' + $width + '" customWidth="1"/>')
    }

    # NOTE: the child order below is fixed by the schema -- dimension,
    # sheetViews, sheetFormatPr, cols, sheetData, autoFilter, mergeCells.
    # Excel refuses the file if autoFilter comes after mergeCells.
    $sheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<dimension ref="A1:' + $lastCol + $lastRow + '"/>' +
      '<sheetViews><sheetView workbookViewId="0">' +
      '<pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/>' +
      '</sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>' +
      '<cols>' + $cols.ToString() + '</cols>' +
      '<sheetData>' + $sb.ToString() + '</sheetData>' +
      '<autoFilter ref="A2:' + $lastCol + $lastRow + '"/>' +
      '<mergeCells count="1"><mergeCell ref="A1:' + $lastCol + '1"/></mergeCells>' +
      '</worksheet>'

    $workbookXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="' + (ConvertTo-XmlText $Sheet) + '" sheetId="1" r:id="rId1"/></sheets></workbook>'

    $contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'

    $rootRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'

    $workbookRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'

    $parts = [ordered] @{
        '[Content_Types].xml'          = $contentTypes
        '_rels/.rels'                  = $rootRels
        'xl/workbook.xml'              = $workbookXml
        'xl/_rels/workbook.xml.rels'   = $workbookRels
        'xl/styles.xml'                = (New-XlsxStylesXml -Style $Style)
        'xl/worksheets/sheet1.xml'     = $sheetXml
    }

    $parent = Split-Path -Parent $Path
    if ($parent -and -not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
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
}

function Write-XlsxWithModule {
    <#
    .SYNOPSIS Write the workbook using the ImportExcel module.
    .PARAMETER Rows / Columns / Style / Striped / Banner / Path / Sheet
               As for Write-XlsxNative.
    .OUTPUTS None.
    #>
    param(
        [object[]] $Rows, [string[]] $Columns, [hashtable] $Style,
        [bool] $Striped, [string] $Banner, [string] $Path, [string] $Sheet
    )
    Import-Module ImportExcel -ErrorAction Stop
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    $columnCount = [Math]::Max(1, $Columns.Count)
    $lastCol = Get-ColumnRef -Index $columnCount

    # StartRow 2 leaves row 1 for the banner, matching the Python twin -- so
    # copyExcelSheet2CSV must be called with -HeaderRow 2 to round-trip.
    $excel = $Rows | Export-Excel -Path $Path -WorksheetName $Sheet `
                -AutoSize -FreezeTopRowFirstColumn:$false -StartRow 2 -PassThru
    $sheetObject = $excel.Workbook.Worksheets[$Sheet]

    Set-ExcelRange -Worksheet $sheetObject -Range "A1:${lastCol}1" -Merge `
        -BackgroundColor "#$($Style.Banner)" -FontColor "#$($Style.BannerFont)" -Bold -Value $Banner
    Set-ExcelRange -Worksheet $sheetObject -Range "A2:${lastCol}2" `
        -BackgroundColor "#$($Style.Header)" -FontColor "#$($Style.HeaderFont)" -Bold
    if ($Striped) {
        for ($r = 3; $r -lt (3 + $Rows.Count); $r++) {
            if ($r % 2 -eq 1) {
                Set-ExcelRange -Worksheet $sheetObject -Range "A${r}:${lastCol}${r}" `
                    -BackgroundColor "#$($Style.Stripe)"
            }
        }
    }
    $sheetObject.View.FreezePanes(3, 1)
    Close-ExcelPackage $excel
}

function Invoke-Main {
    <# .SYNOPSIS Style the CSV into a workbook. .OUTPUTS [int] exit code. #>
    if (-not $InputCsvPath)    { Write-CgsError "required parameter 'InputCsvPath' is missing or empty"; return 2 }
    if (-not $OutputExcelPath) { Write-CgsError "required parameter 'OutputExcelPath' is missing or empty"; return 2 }
    if (-not $script:FormatStyles.Contains($FormatType)) {
        Write-CgsError ("unknown FormatType '$FormatType'; expected one of: " + (($script:FormatStyles.Keys | Sort-Object) -join ', '))
        return 2
    }
    if ($Writer -notin @('auto', 'native', 'module')) {
        Write-CgsError "unknown Writer '$Writer'; expected one of: auto, native, module"
        return 2
    }
    if (-not (Test-Path -LiteralPath $InputCsvPath)) {
        Write-CgsError "input CSV not found: $InputCsvPath"
        return 3
    }

    $rows = @(Import-Csv -LiteralPath $InputCsvPath)
    $columns = @()
    if ($rows.Count -gt 0) {
        $columns = @($rows[0].PSObject.Properties | ForEach-Object { $_.Name })
    }
    $style   = $script:FormatStyles[$FormatType]
    $striped = $script:StripedFormats -contains $FormatType
    $banner  = if ($Title) { $Title } else { [System.IO.Path]::GetFileNameWithoutExtension($InputCsvPath) }

    # Choose a writer. ImportExcel is preferred when present because it is
    # well travelled; the native writer is the fallback so a server without
    # the module -- and without PSGallery access to install it -- still works.
    $haveModule = [bool] (Get-Module -ListAvailable -Name ImportExcel)
    $useModule = switch ($Writer) {
        'module' { $true }
        'native' { $false }
        default  { $haveModule }
    }
    if ($useModule -and -not $haveModule) {
        Write-CgsError "-Writer module was requested but the ImportExcel module is not installed"
        return 3
    }

    if ($useModule) {
        Write-CgsInfo "using the ImportExcel module"
        Write-XlsxWithModule -Rows $rows -Columns $columns -Style $style `
            -Striped $striped -Banner $banner -Path $OutputExcelPath -Sheet $SheetName
    }
    else {
        Write-CgsInfo "ImportExcel not available; writing the workbook natively"
        Write-XlsxNative -Rows $rows -Columns $columns -Style $style `
            -Striped $striped -Banner $banner -Path $OutputExcelPath -Sheet $SheetName
    }

    Write-CgsInfo ("formatted {0} row(s) x {1} column(s) [{2}] -> {3}" -f $rows.Count, $columns.Count, $FormatType, $OutputExcelPath)
    return 0
}
try { exit (Invoke-Main) }
catch { Write-CgsError ("unhandled error: {0}" -f $_.Exception.Message); exit 3 }
