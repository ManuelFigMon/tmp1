<#
=====================================================================
  Program Name  : formatExcel.ps1
  Author        : Manuel Figallo
  Purpose       : Turn one worksheet of an .xlsx into a formatted status
                  report for management -- Word, PDF or HTML -- picking
                  the renderer from the OUTPUT file's extension.
  Version       : 1.0beta
  Created       : 2026-10-27
  Last Modified : 2026-10-27

  Dependencies:
    NONE that must be installed, for .docx and .html. The workbook is read
    with Read-CgsXlsxSheet from cgsUtils.ps1, which needs no add-in, and the
    Word document is written as OOXML directly.

    .pdf IS THE EXCEPTION. PowerShell cannot write a PDF on its own, so the
    script drives a converter: Microsoft Word through COM when it is there,
    otherwise LibreOffice in headless mode, which needs no licence and is
    often present where Word is not. Where NEITHER is installed the script
    does NOT quietly hand you a different file -- it stops and names the
    alternatives. Pass -OnWordMissing html or -OnWordMissing docx to opt in
    to a fallback deliberately.

  Description:
    Reads the sheet, keeps only the columns asked for and in that order,
    then renders a titled, banded report. The look is chosen with
    -format_name, using the same palette as formatData so a Word report and
    an Excel workbook from this toolkit match.

    Columns are selected by NAME, not position, so a Planner or Snowflake
    export that gains a column upstream does not silently shift the report.
    A name that is not in the sheet is reported with the names that ARE.

  Input Parameters (required first):
    -ExcelPath   (REQUIRED) source .xlsx.
    -SheetName   (REQUIRED) worksheet to format, e.g. "Consolidated Data".
    -OutputPath  (REQUIRED) destination. Its EXTENSION picks the renderer:
                   .docx -> Word          (native, no Office needed)
                   .html -> web page      (native, prints to PDF anywhere)
                   .pdf  -> PDF           (needs Word or LibreOffice)
    -Columns     (optional) names to keep, in order. Empty = every column.
    -format_name (optional, default "corporate")
                   corporate | corporatev2 | plain | minimal
    -Title       (optional) banner text; defaults to the workbook name.
    -Subtitle    (optional) second line; defaults to sheet and row count.
    -ExcludeWhere (optional) drop rows matching "Column=Value". Several
                 values with | and several clauses with ;, e.g.
                 "Bucket=BACKLOG|TODO (Up next)".
    -IncludeWhere (optional) the inverse: keep ONLY matching rows. Applied
                 first when both are given.
    -SortBy      (optional) "Column", "Column ASC" or "Column DESC";
                 several keys with ;. Blanks sort last in both directions.
    -HeaderRow   (optional, default 0 = detect) 1-based header row.
    -OnWordMissing (optional, default "fail") fail | html | docx -- what to
                   do when a .pdf is asked for and NO converter (Word or
                   LibreOffice) is installed.
    -MaxCellLength (optional, default 0 = never shorten) cut long free-text
                   cells to this many characters, marked with an ellipsis.
                   Useful for a Notes column that would otherwise run to
                   half a page per row.
  Exit codes: 0 = success, 2 = config error, 3 = I/O or renderer failure.

  Change Log:
    v1.0beta - First release. docx/html native, pdf via Word COM.
    v1.1beta - PDF also via LibreOffice when Word is absent, so a server
               without Office can still produce the PDF itself.
=====================================================================
#>
[CmdletBinding()]
param(
    [string]   $ExcelPath     = '',
    [string]   $SheetName     = '',
    [string]   $OutputPath    = '',
    [object]   $Columns      = @(),
    [string]   $format_name   = 'corporate',
    [string]   $Title         = '',
    [string]   $Subtitle      = '',
    [int]      $HeaderRow     = 0,
    [string]   $OnWordMissing = 'fail',
    [int]      $MaxCellLength = 0,
    [string]   $ExcludeWhere  = '',
    [string]   $IncludeWhere  = '',
    [string]   $SortBy        = ''
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\cgsUtils.ps1"

# Palette per format type -- matches FORMAT_STYLES in formatData, so a Word
# report and an Excel workbook from this toolkit look like one family.
# Hex WITHOUT the leading '#'.
$script:FormatStyles = @{
    'corporate'   = @{ Banner='1F3864'; Header='2E75B6'; Stripe='DCE6F1'; HeaderFont='FFFFFF'; BannerFont='FFFFFF' }
    'corporatev2' = @{ Banner='1F3864'; Header='2E75B6'; Stripe='DCE6F1'; HeaderFont='FFFFFF'; BannerFont='FFFFFF' }
    'plain'       = @{ Banner='FFFFFF'; Header='D9D9D9'; Stripe='FFFFFF'; HeaderFont='000000'; BannerFont='000000' }
    'minimal'     = @{ Banner='FFFFFF'; Header='FFFFFF'; Stripe='FFFFFF'; HeaderFont='000000'; BannerFont='000000' }
}
$script:StripedFormats = @('corporate', 'corporatev2')
$script:Version = '1.1beta'

# Landscape Letter in twentieths of a point. Seven columns of task text do
# not fit portrait, and a report nobody can read is not a report.
$script:PageWidth  = 15840
$script:PageHeight = 12240
$script:PageMargin = 720
$script:TableWidth = $script:PageWidth - (2 * $script:PageMargin)

function ConvertTo-HtmlText {
    <# .SYNOPSIS Escape a value for HTML text content.
       .OUTPUTS [string] with & < > and quotes escaped. #>
    param($Value)
    if ($null -eq $Value) { return '' }
    return ([string]$Value).Replace('&','&amp;').Replace('<','&lt;').
        Replace('>','&gt;').Replace('"','&quot;')
}

function Resolve-ReportColumn {
    <# .SYNOPSIS Match a column name against the sheet's, forgiving case.
       .OUTPUTS [string] the sheet's own spelling of the column.
       .NOTES  Exact match first, then case-insensitive. Values are already
               matched case-insensitively, so having "bucket=BACKLOG" fail
               on the column while "Bucket=backlog" works would be a
               distinction nobody can predict. #>
    param([string] $Name, [string[]] $Available)
    $wanted = $Name.Trim()
    if ($Available -ccontains $wanted) { return $wanted }
    $match = @($Available | Where-Object { $_ -ieq $wanted })
    if ($match.Count -gt 0) { return $match[0] }
    throw ("column '{0}' is not in this sheet`r`n  the sheet has: {1}" -f
           $Name, ($Available -join ', '))
}

function ConvertTo-FilterClauses {
    <# .SYNOPSIS Parse "Column=Value|Value;Column=Value" into clauses.
       .OUTPUTS [object[]] of @{ Column; Values }.
       .NOTES  Split on the FIRST '=' only. A value may legitimately
               contain one, and "Notes=a=b" has to mean the column Notes
               and the value "a=b", not a parse error. #>
    param([object] $Spec)
    $clauses = @()
    foreach ($clause in (ConvertTo-CgsList $Spec)) {
        $at = $clause.IndexOf('=')
        if ($at -lt 0) {
            # The parentheses are load-bearing: -f binds TIGHTER than +,
            # so without them the format would apply to the last fragment
            # alone and '{0}' in the first would reach the user literally.
            throw (("filter clause '{0}' has no '='. Write it as " +
                    "Column=Value, several values as Column=One|Two, and " +
                    "several clauses as Column=One;Other=Two.") -f $clause)
        }
        # An empty value is kept on purpose: "Bucket=" means the rows where
        # Bucket is blank, which is a real thing to ask for.
        $clauses += @{
            Column = $clause.Substring(0, $at).Trim()
            Values = @($clause.Substring($at + 1).Split('|') |
                       ForEach-Object { $_.Trim() })
        }
    }
    return , $clauses
}

function Test-RowMatchesClauses {
    <# .SYNOPSIS Does this row match ANY clause?
       .OUTPUTS [bool]
       .NOTES  ANY, not ALL: "filter out BACKLOG and DONE" means drop both,
               which is what anyone saying it out loud intends.

               Matching is trimmed and case-insensitive, and EXACT rather
               than substring. Substring would make Status=DONE quietly
               drop "NOT DONE" as well, and a status report that removes
               the wrong rows is worse than one that fails. #>
    param($Row, [object[]] $Clauses)
    foreach ($clause in $Clauses) {
        $cell = ''
        if ($null -ne $Row.($clause.Column)) {
            $cell = ([string] $Row.($clause.Column)).Trim()
        }
        foreach ($value in $clause.Values) {
            if ($cell -ieq $value) { return $true }
        }
    }
    return $false
}

function Select-ReportRows {
    <# .SYNOPSIS Keep or drop rows by column value.
       .OUTPUTS [hashtable] @{ Rows; Note }.
       .NOTES  THIS RUNS BEFORE COLUMN SELECTION, deliberately. A
               management report may well show Task Name and Due date but
               not Bucket -- and still need to filter on Bucket. Filtering
               after column selection would fail on the most ordinary
               request there is.

               A missing column THROWS rather than matching nothing. A typo
               that silently matches nothing ships an unfiltered report
               that its author believes is filtered, which is the failure
               mode worth being loud about. #>
    param([object[]] $Rows, [object] $ExcludeWhere, [object] $IncludeWhere)
    $exclude = ConvertTo-FilterClauses $ExcludeWhere
    $include = ConvertTo-FilterClauses $IncludeWhere
    if ($exclude.Count -eq 0 -and $include.Count -eq 0) {
        return @{ Rows = $Rows; Note = '' }
    }
    if ($Rows.Count -eq 0) { return @{ Rows = $Rows; Note = '' } }

    $available = @($Rows[0].PSObject.Properties | ForEach-Object { $_.Name })
    foreach ($clause in ($exclude + $include)) {
        $clause.Column = Resolve-ReportColumn -Name $clause.Column -Available $available
    }

    $kept = $Rows
    if ($include.Count -gt 0) {
        $kept = @($kept | Where-Object { Test-RowMatchesClauses -Row $_ -Clauses $include })
    }
    if ($exclude.Count -gt 0) {
        $kept = @($kept | Where-Object { -not (Test-RowMatchesClauses -Row $_ -Clauses $exclude) })
    }

    if ($kept.Count -eq 0) {
        throw (("the filter removed every row, so there is nothing to report.`r`n" +
                "  rows before filtering: {0}`r`n" +
                "  check the spelling and the case of the values you asked for.") -f $Rows.Count)
    }

    $parts = @()
    foreach ($pair in @(@('Excluded', $exclude), @('Included', $include))) {
        foreach ($clause in $pair[1]) {
            $parts += ("{0}: {1} = {2}" -f $pair[0], $clause.Column,
                       ($clause.Values -join ' or '))
        }
    }
    $note = ("Showing {0} of {1} rows.  {2}" -f $kept.Count, $Rows.Count,
             ($parts -join '  '))
    return @{ Rows = $kept; Note = $note }
}

function ConvertTo-SortKeys {
    <# .SYNOPSIS Parse "Column DESC; Other Column" into sort keys.
       .OUTPUTS [object[]] of @{ Column; Descending }.
       .NOTES  Space-separated rather than "Column=DESC", because '='
               already means "equals this value" in -ExcludeWhere and
               reusing it for a direction would make two parameters that
               look alike behave differently.

               The split is on the LAST whitespace, and only when that last
               word is ASC or DESC. That is what lets "Due date" stay one
               column name while "Due date DESC" becomes a column and a
               direction. #>
    param([object] $Spec)
    $keys = @()
    foreach ($clause in (ConvertTo-CgsList $Spec)) {
        $text = $clause.Trim()
        $descending = $false
        $at = $text.LastIndexOfAny(@(' ', "`t"))
        if ($at -gt 0) {
            $tail = $text.Substring($at + 1).ToUpperInvariant()
            if ($tail -eq 'ASC' -or $tail -eq 'DESC') {
                $descending = ($tail -eq 'DESC')
                $text = $text.Substring(0, $at).Trim()
            }
        }
        if (-not $text) {
            throw (("sort clause '{0}' names a direction but no column. " +
                    "Write it as ""Column"", ""Column ASC"" or ""Column DESC"".") -f $clause)
        }
        $keys += @{ Column = $text; Descending = $descending }
    }
    return , $keys
}

function Get-SortRank {
    <# .SYNOPSIS Build a sort key that survives a column of mixed types.
       .OUTPUTS [hashtable] @{ Rank; Number; Text }.
       .NOTES  A worksheet column happily holds 3, "3", "" and "Urgent" at
               once, and comparing a number to a string is not reliable.
               Ranking first keeps the comparison inside one type: numbers
               (0), then text (1), then blanks.

               Blanks rank LAST in both directions, which is why the rank
               depends on $Descending: the whole list is reversed for a
               DESC sort, so a blank has to start in front to finish at the
               back. A report whose DESC column opens with twenty empty
               cells is not sorted in any useful sense. #>
    param($Value, [bool] $Descending)
    $text = ''
    if ($null -ne $Value) { $text = ([string] $Value).Trim() }
    if ($text -eq '') {
        return @{ Rank = $(if ($Descending) { -1 } else { 2 }); Number = [double] 0; Text = '' }
    }
    [double] $number = 0
    if ([double]::TryParse($text, [ref] $number)) {
        return @{ Rank = 0; Number = $number; Text = '' }
    }
    return @{ Rank = 1; Number = [double] 0; Text = $text.ToLowerInvariant() }
}

function Sort-ReportRows {
    <# .SYNOPSIS Sort rows by one or more columns.
       .OUTPUTS [hashtable] @{ Rows; Note }.
       .NOTES  Keys are applied LAST FIRST, and each pass is stable, so
               rows that tie on the first key stay ordered by the second --
               which is what "sort by Bucket then Due date" means, and what
               one pass over a composite key would get wrong.

               NOT Sort-Object. Sort-Object compares strings the way the
               CURRENT CULTURE does, which treats punctuation as lower
               weight: it puts "CS_CGS01" before "CS0483966", while Python
               -- and the .py twin -- compare by code point and put
               "CS0483966" first. Two engines that are meant to produce
               byte-identical documents cannot disagree about row order, so
               this sorts with an explicit ORDINAL comparison instead.

               The comparison also carries the original position and falls
               back to it on a tie, which makes the sort stable whatever
               algorithm Array.Sort picks -- introsort is not stable on its
               own. Ties keep their original order in BOTH directions, the
               same as Python's reverse=True.

               Like the filter, this runs before column selection, so you
               can sort on a column the report does not display. #>
    param([object[]] $Rows, [object] $SortBy)
    $keys = ConvertTo-SortKeys $SortBy
    if ($keys.Count -eq 0 -or $Rows.Count -eq 0) {
        return @{ Rows = $Rows; Note = '' }
    }

    $available = @($Rows[0].PSObject.Properties | ForEach-Object { $_.Name })
    foreach ($key in $keys) {
        $key.Column = Resolve-ReportColumn -Name $key.Column -Available $available
    }

    $ordered = $Rows
    for ($index = $keys.Count - 1; $index -ge 0; $index--) {
        $key = $keys[$index]
        $descending = [bool] $key.Descending

        $decorated = New-Object 'System.Collections.Generic.List[object]'
        for ($position = 0; $position -lt $ordered.Count; $position++) {
            $rank = Get-SortRank -Value $ordered[$position].($key.Column) -Descending $descending
            [void] $decorated.Add([pscustomobject] @{
                Rank = [int] $rank.Rank; Number = [double] $rank.Number
                Text = [string] $rank.Text; Position = $position
                Row  = $ordered[$position] })
        }

        $comparison = [System.Comparison[object]] {
            param($a, $b)
            $result = 0
            if ($a.Rank -ne $b.Rank)           { $result = $a.Rank.CompareTo($b.Rank) }
            elseif ($a.Number -ne $b.Number)   { $result = $a.Number.CompareTo($b.Number) }
            else { $result = [string]::CompareOrdinal($a.Text, $b.Text) }
            if ($result -ne 0) {
                if ($descending) { return -$result }
                return $result
            }
            return $a.Position.CompareTo($b.Position)
        }.GetNewClosure()

        $decorated.Sort($comparison)
        $ordered = @($decorated | ForEach-Object { $_.Row })
    }

    $labels = @($keys | ForEach-Object {
        "{0} {1}" -f $_.Column, $(if ($_.Descending) { 'DESC' } else { 'ASC' }) })
    return @{ Rows = $ordered; Note = ("Sorted by: {0}" -f ($labels -join ', ')) }
}

function Select-ReportColumns {
    <# .SYNOPSIS Keep the requested columns, in the requested order.
       .PARAMETER Rows     Objects from Read-CgsXlsxSheet.
       .PARAMETER Wanted   Column names; empty keeps everything.
       .OUTPUTS [string[]] the column names to render.
       .NOTES  Selection is BY NAME so an upstream export that gains a
               column cannot silently shift the report sideways. A name
               that is absent is an error naming what is present, which is
               far more useful than a column of blanks.

               THE LIST ARRIVES IN THREE SHAPES. Called as PowerShell,
               -Columns 'a','b' is a real array. Launched with -File -- how
               SAS and a command line call it -- EVERY argument is a string,
               so the same text arrives as the single string "a,b". Both
               are accepted: the project's ';' convention first, then ','
               but only when the whole string is not itself a column name,
               so a column legitimately containing a comma still works. #>
    param([object[]] $Rows, [object] $Wanted)
    if ($Rows.Count -eq 0) { return , @() }
    $available = @($Rows[0].PSObject.Properties | ForEach-Object { $_.Name })

    $requested = ConvertTo-CgsList $Wanted
    if ($requested.Count -eq 1 -and $requested[0].Contains(',') -and
        $available -notcontains $requested[0]) {
        $requested = @($requested[0].Split(',') |
                       ForEach-Object { $_.Trim() } |
                       Where-Object { $_ })
    }
    if ($requested.Count -eq 0) { return , $available }

    $missing = @($requested | Where-Object { $available -notcontains $_ })
    if ($missing.Count -gt 0) {
        throw ("column(s) not found in this sheet: {0}`r`n  the sheet has: {1}" -f
               ($missing -join ', '), ($available -join ', '))
    }
    return , @($requested)
}

function Get-ColumnWidths {
    <# .SYNOPSIS Share the page width out between columns by content length.
       .OUTPUTS [int[]] widths in twips, summing to the table width.
       .NOTES  Sampled from the first 200 rows and clamped, so one long
               Notes field cannot squeeze every other column to nothing.

               THEN EVERY COLUMN IS GIVEN A FLOOR WIDE ENOUGH FOR ITS OWN
               HEADING. Proportional sharing alone put "Priority" in 0.47in
               and broke the word across two lines, while Notes took a
               third of the page. A heading that wraps mid-word makes the
               report look broken, so narrow columns take their floor first
               and the long ones share what is left. #>
    param([object[]] $Rows, [string[]] $Names)
    $sample = if ($Rows.Count -gt 200) { $Rows[0..199] } else { $Rows }
    $weights = @(); $floors = @()
    foreach ($name in $Names) {
        $widest = $name.Length
        foreach ($row in $sample) {
            $length = ([string] $row.$name).Length
            if ($length -gt $widest) { $widest = $length }
        }
        # Clamp before sharing out: a 400-character note is not worth 40x
        # the width of a date, however long it is.
        $weights += [Math]::Max(8, [Math]::Min(45, $widest))
        # ~115 twips per character of 9pt bold Calibri, plus cell padding.
        $floors  += [int](($name.Length * 115) + 260)
    }

    $total = ($weights | Measure-Object -Sum).Sum
    $widths = @()
    for ($i = 0; $i -lt $Names.Count; $i++) {
        $widths += [int]($script:TableWidth * $weights[$i] / $total)
    }

    # Raise anything under its floor, and recover the difference from the
    # columns that have room to give.
    $deficit = 0
    for ($i = 0; $i -lt $widths.Count; $i++) {
        if ($widths[$i] -lt $floors[$i]) {
            $deficit += $floors[$i] - $widths[$i]
            $widths[$i] = $floors[$i]
        }
    }
    if ($deficit -gt 0) {
        $donors = @(0..($widths.Count - 1) | Where-Object { $widths[$_] -gt $floors[$_] })
        $spare = 0
        foreach ($index in $donors) { $spare += $widths[$index] - $floors[$index] }
        if ($spare -gt 0) {
            foreach ($index in $donors) {
                $share = [int]($deficit * ($widths[$index] - $floors[$index]) / $spare)
                $widths[$index] -= $share
            }
        }
    }

    # Give any rounding remainder to the widest column so the row is exact.
    $widestIndex = 0
    for ($i = 1; $i -lt $widths.Count; $i++) {
        if ($widths[$i] -gt $widths[$widestIndex]) { $widestIndex = $i }
    }
    $widths[$widestIndex] += $script:TableWidth - (($widths | Measure-Object -Sum).Sum)
    return , $widths
}

function Limit-CellText {
    <# .SYNOPSIS Shorten a value for a report, visibly.
       .PARAMETER Value      The cell value.
       .PARAMETER MaxLength  0 means never shorten.
       .OUTPUTS [string] the value, with a trailing ellipsis when cut.
       .NOTES  Off by default. Truncating a report silently would hide data
               from the people relying on it, so the cut is always marked
               and the caller has to ask for it. #>
    param($Value, [int] $MaxLength)
    $text = [string] $Value
    if ($MaxLength -le 0 -or $text.Length -le $MaxLength) { return $text }
    return $text.Substring(0, [Math]::Max(1, $MaxLength - 1)) + [char]0x2026
}

function New-DocxBytes {
    <# .SYNOPSIS Build a .docx as OOXML, with no Word and no add-in.
       .PARAMETER Rows / Names / Widths / Style / Striped / Title / Subtitle
       .PARAMETER Path  Destination .docx.
       .OUTPUTS [string] the path written.
       .NOTES  A .docx is a zip of XML parts, written here directly. The
               child order inside sectPr and tblPr is fixed by the schema;
               Word refuses the file when they are out of order. #>
    param([object[]] $Rows, [string[]] $Names, [int[]] $Widths,
          [hashtable] $Style, [bool] $Striped, [string] $Title,
          [string] $Subtitle, [string] $Path, [int] $MaxLength = 0,
          [string] $Note = '')
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $sb = New-Object System.Text.StringBuilder

    # Title banner: a one-cell shaded table, which prints as a solid band
    # where a shaded paragraph leaves white gutters at the margins.
    [void] $sb.Append('<w:tbl><w:tblPr><w:tblW w:w="' + $script:TableWidth + '" w:type="dxa"/>')
    [void] $sb.Append('<w:tblLayout w:type="fixed"/></w:tblPr>')
    [void] $sb.Append('<w:tblGrid><w:gridCol w:w="' + $script:TableWidth + '"/></w:tblGrid>')
    [void] $sb.Append('<w:tr><w:tc><w:tcPr><w:tcW w:w="' + $script:TableWidth + '" w:type="dxa"/>')
    [void] $sb.Append('<w:shd w:val="clear" w:fill="' + $Style.Banner + '"/>')
    [void] $sb.Append('<w:tcMar><w:top w:w="140" w:type="dxa"/><w:left w:w="180" w:type="dxa"/><w:bottom w:w="140" w:type="dxa"/></w:tcMar></w:tcPr>')
    [void] $sb.Append('<w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="32"/><w:color w:val="' + $Style.BannerFont + '"/></w:rPr>')
    [void] $sb.Append('<w:t xml:space="preserve">' + (ConvertTo-CgsXmlText $Title) + '</w:t></w:r></w:p>')
    if ($Subtitle) {
        [void] $sb.Append('<w:p><w:pPr><w:spacing w:before="40" w:after="0"/></w:pPr><w:r><w:rPr><w:sz w:val="19"/><w:color w:val="' + $Style.BannerFont + '"/></w:rPr>')
        [void] $sb.Append('<w:t xml:space="preserve">' + (ConvertTo-CgsXmlText $Subtitle) + '</w:t></w:r></w:p>')
    }
    # The filter line is its OWN paragraph, not a newline glued onto the
    # subtitle: a newline inside a <w:t> is collapsed to a space by Word,
    # so the note would silently run on and read as one sentence.
    if ($Note) {
        [void] $sb.Append('<w:p><w:pPr><w:spacing w:before="60" w:after="0"/></w:pPr>')
        [void] $sb.Append('<w:r><w:rPr><w:i/><w:sz w:val="17"/><w:color w:val="' + $Style.BannerFont + '"/></w:rPr>')
        [void] $sb.Append('<w:t xml:space="preserve">' + (ConvertTo-CgsXmlText $Note) + '</w:t></w:r></w:p>')
    }
    [void] $sb.Append('</w:tc></w:tr></w:tbl>')
    [void] $sb.Append('<w:p><w:pPr><w:spacing w:after="160"/></w:pPr></w:p>')

    # The data table.
    $borders = '<w:tblBorders>'
    foreach ($edge in 'top','left','bottom','right','insideH','insideV') {
        $borders += '<w:' + $edge + ' w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
    }
    $borders += '</w:tblBorders>'
    [void] $sb.Append('<w:tbl><w:tblPr><w:tblW w:w="' + $script:TableWidth + '" w:type="dxa"/>')
    [void] $sb.Append($borders + '<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>')
    foreach ($width in $Widths) { [void] $sb.Append('<w:gridCol w:w="' + $width + '"/>') }
    [void] $sb.Append('</w:tblGrid>')

    # Header row. tblHeader repeats it at the top of every printed page --
    # without it a three-page report has unlabelled columns after page one.
    [void] $sb.Append('<w:tr><w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>')
    for ($i = 0; $i -lt $Names.Count; $i++) {
        [void] $sb.Append('<w:tc><w:tcPr><w:tcW w:w="' + $Widths[$i] + '" w:type="dxa"/>')
        [void] $sb.Append('<w:shd w:val="clear" w:fill="' + $Style.Header + '"/></w:tcPr>')
        [void] $sb.Append('<w:p><w:pPr><w:spacing w:after="40"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="18"/><w:color w:val="' + $Style.HeaderFont + '"/></w:rPr>')
        [void] $sb.Append('<w:t xml:space="preserve">' + (ConvertTo-CgsXmlText $Names[$i]) + '</w:t></w:r></w:p></w:tc>')
    }
    [void] $sb.Append('</w:tr>')

    $rowNumber = 0
    foreach ($row in $Rows) {
        $rowNumber++
        $fill = if ($Striped -and ($rowNumber % 2 -eq 1)) { $Style.Stripe } else { '' }
        [void] $sb.Append('<w:tr>')
        for ($i = 0; $i -lt $Names.Count; $i++) {
            [void] $sb.Append('<w:tc><w:tcPr><w:tcW w:w="' + $Widths[$i] + '" w:type="dxa"/>')
            if ($fill) { [void] $sb.Append('<w:shd w:val="clear" w:fill="' + $fill + '"/>') }
            [void] $sb.Append('</w:tcPr><w:p><w:pPr><w:spacing w:after="30"/></w:pPr><w:r><w:rPr><w:sz w:val="17"/></w:rPr>')
            $value = Limit-CellText -Value $row.($Names[$i]) -MaxLength $MaxLength
            [void] $sb.Append('<w:t xml:space="preserve">' + (ConvertTo-CgsXmlText $value) + '</w:t></w:r></w:p></w:tc>')
        }
        [void] $sb.Append('</w:tr>')
    }
    [void] $sb.Append('</w:tbl>')

    $document = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:body>' + $sb.ToString() +
      '<w:sectPr><w:pgSz w:w="' + $script:PageWidth + '" w:h="' + $script:PageHeight + '" w:orient="landscape"/>' +
      '<w:pgMar w:top="' + $script:PageMargin + '" w:right="' + $script:PageMargin +
      '" w:bottom="' + $script:PageMargin + '" w:left="' + $script:PageMargin + '"/>' +
      '</w:sectPr></w:body></w:document>'

    $styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>' +
      '<w:sz w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/></w:style>' +
      '</w:styles>'

    $parts = [ordered] @{
        '[Content_Types].xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
            '<Default Extension="xml" ContentType="application/xml"/>' +
            '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
            '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'
        '_rels/.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
        'word/_rels/document.xml.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
        'word/styles.xml'   = $styles
        'word/document.xml' = $document
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
            # UTF-8 with NO byte-order mark: Word rejects a BOM inside parts.
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

function New-HtmlReport {
    <# .SYNOPSIS Write the report as a self-contained web page.
       .OUTPUTS [string] the path written.
       .NOTES  Every style is inline in one file -- no CSS to lose when the
               page is emailed. Any browser prints it to PDF, which is the
               fallback when Word is not installed. #>
    param([object[]] $Rows, [string[]] $Names, [hashtable] $Style,
          [bool] $Striped, [string] $Title, [string] $Subtitle, [string] $Path,
          [int] $MaxLength = 0, [string] $Note = '')
    $stripe = if ($Striped) { $Style.Stripe } else { 'FFFFFF' }
    $html = New-Object System.Text.StringBuilder
    [void] $html.Append('<!doctype html><html><head><meta charset="utf-8">')
    [void] $html.Append('<title>' + (ConvertTo-HtmlText $Title) + '</title><style>')
    [void] $html.Append('body{font-family:Calibri,''Segoe UI'',Arial,sans-serif;margin:24px;color:#1F2937}')
    [void] $html.Append('.banner{background:#' + $Style.Banner + ';color:#' + $Style.BannerFont + ';padding:14px 18px;border-radius:3px}')
    [void] $html.Append('.banner h1{margin:0;font-size:20pt}.banner p{margin:4px 0 0;font-size:10pt;opacity:.9}.banner p.filter{font-style:italic;font-size:9pt;opacity:.95}')
    [void] $html.Append('table{border-collapse:collapse;width:100%;margin-top:16px;font-size:9.5pt}')
    [void] $html.Append('th{background:#' + $Style.Header + ';color:#' + $Style.HeaderFont + ';text-align:left;padding:7px 8px;border:1px solid #BFBFBF}')
    [void] $html.Append('td{padding:6px 8px;border:1px solid #BFBFBF;vertical-align:top}')
    [void] $html.Append('tbody tr:nth-child(odd){background:#' + $stripe + '}')
    [void] $html.Append('@media print{@page{size:landscape;margin:12mm}thead{display:table-header-group}}')
    [void] $html.Append('</style></head><body>')
    [void] $html.Append('<div class="banner"><h1>' + (ConvertTo-HtmlText $Title) + '</h1>')
    if ($Subtitle) { [void] $html.Append('<p>' + (ConvertTo-HtmlText $Subtitle) + '</p>') }
    if ($Note) { [void] $html.Append('<p class="filter">' + (ConvertTo-HtmlText $Note) + '</p>') }
    [void] $html.Append('</div><table><thead><tr>')
    foreach ($name in $Names) { [void] $html.Append('<th>' + (ConvertTo-HtmlText $name) + '</th>') }
    [void] $html.Append('</tr></thead><tbody>')
    foreach ($row in $Rows) {
        [void] $html.Append('<tr>')
        foreach ($name in $Names) {
            $value = Limit-CellText -Value $row.$name -MaxLength $MaxLength
            [void] $html.Append('<td>' + (ConvertTo-HtmlText $value) + '</td>')
        }
        [void] $html.Append('</tr>')
    }
    [void] $html.Append('</tbody></table></body></html>')

    $parent = Split-Path -Parent $Path
    if ($parent -and -not (Test-Path -LiteralPath $parent)) {
        [void](New-Item -ItemType Directory -Path $parent -Force)
    }
    $encoding = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $html.ToString(), $encoding)
    return $Path
}

function Test-WordAvailable {
    <# .SYNOPSIS Is Microsoft Word installed and usable from COM?
       .OUTPUTS [bool] #>
    if (-not $IsWindows -and $null -ne (Get-Variable -Name IsWindows -ErrorAction SilentlyContinue)) {
        return $false        # COM automation is Windows-only
    }
    try {
        $word = New-Object -ComObject Word.Application -ErrorAction Stop
        $word.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
        return $true
    } catch { return $false }
}

function Find-LibreOffice {
    <# .SYNOPSIS Locate the LibreOffice command line, if it is installed.
       .OUTPUTS [string] the executable path, or $null.
       .NOTES  LibreOffice converts a .docx to PDF with no Word and no
               licence. It is worth looking for: on a locked-down server
               Word is often absent while LibreOffice is not. #>
    foreach ($name in 'soffice', 'libreoffice') {
        $found = Get-Command $name -ErrorAction SilentlyContinue
        if ($found) { return $found.Source }
    }
    foreach ($candidate in
             "$env:ProgramFiles\LibreOffice\program\soffice.exe",
             "${env:ProgramFiles(x86)}\LibreOffice\program\soffice.exe") {
        if ($candidate -and (Test-Path -LiteralPath $candidate)) { return $candidate }
    }
    return $null
}

function Convert-DocxToPdfWithLibreOffice {
    <# .SYNOPSIS Save a .docx as .pdf using LibreOffice in headless mode.
       .OUTPUTS [string] the .pdf path.
       .NOTES  LibreOffice always names its output after the INPUT file and
               drops it in --outdir, so when the caller asked for a
               different name the result is renamed afterwards. #>
    param([string] $Executable, [string] $DocxPath, [string] $PdfPath)
    $outputDirectory = Split-Path -Parent $PdfPath
    if (-not $outputDirectory) { $outputDirectory = '.' }
    $arguments = @('--headless', '--convert-to', 'pdf', '--outdir',
                   $outputDirectory, $DocxPath)
    & $Executable @arguments | Out-Null
    $produced = Join-Path $outputDirectory (
        [System.IO.Path]::GetFileNameWithoutExtension($DocxPath) + '.pdf')
    if (-not (Test-Path -LiteralPath $produced)) {
        throw "LibreOffice did not produce a PDF from $DocxPath"
    }
    if ($produced -ne $PdfPath) {
        if (Test-Path -LiteralPath $PdfPath) { Remove-Item -LiteralPath $PdfPath -Force }
        Move-Item -LiteralPath $produced -Destination $PdfPath
    }
    return $PdfPath
}

function Convert-DocxToPdf {
    <# .SYNOPSIS Save a .docx as .pdf by driving Word through COM.
       .OUTPUTS [string] the .pdf path.
       .NOTES  wdFormatPDF = 17. Word is quit and released in a finally
               block: a COM instance left running holds the file open and
               accumulates invisible WINWORD.EXE processes on the server. #>
    param([string] $DocxPath, [string] $PdfPath)
    $word = $null; $document = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $document = $word.Documents.Open($DocxPath, $false, $true)   # read-only
        $document.SaveAs([ref] $PdfPath, [ref] 17)
    } finally {
        if ($document) { $document.Close([ref] $false); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) }
        if ($word) { $word.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
    }
    return $PdfPath
}

function Invoke-Main {
    <# .SYNOPSIS Read, select, render. .OUTPUTS [int] exit code. #>
    if (-not $ExcelPath)  { Write-CgsError "required parameter 'ExcelPath' is missing or empty";  return 2 }
    if (-not $SheetName)  { Write-CgsError "required parameter 'SheetName' is missing or empty";  return 2 }
    if (-not $OutputPath) { Write-CgsError "required parameter 'OutputPath' is missing or empty"; return 2 }
    if (-not $script:FormatStyles.Contains($format_name)) {
        Write-CgsError ("unknown format_name '$format_name'; expected one of: " +
                        (($script:FormatStyles.Keys | Sort-Object) -join ', '))
        return 2
    }
    if ($OnWordMissing -notin @('fail','html','docx')) {
        Write-CgsError "unknown OnWordMissing '$OnWordMissing'; expected fail, html or docx"
        return 2
    }
    if (-not (Test-Path -LiteralPath $ExcelPath)) {
        Write-CgsError "input workbook not found: $ExcelPath"
        return 3
    }

    Write-CgsInfo "formatExcel $script:Version starting"
    Write-CgsInfo (Get-CgsUtilsBanner)

    # NO @() around this call: Read-CgsXlsxSheet returns ",$array", which
    # arrives as one object that @() would re-wrap into a 1-element array.
    $rows = Read-CgsXlsxSheet -Path $ExcelPath -Sheet $SheetName -HeaderRow $HeaderRow
    if ($null -eq $rows) { $rows = @() }
    Write-CgsInfo ("read {0} row(s) from sheet '{1}'" -f $rows.Count, $SheetName)

    # Filter and sort BEFORE choosing columns, so either can name a column
    # the report does not display.
    $before   = $rows.Count
    $filtered = Select-ReportRows -Rows $rows -ExcludeWhere $ExcludeWhere -IncludeWhere $IncludeWhere
    $rows     = $filtered.Rows
    $note     = $filtered.Note
    if ($note) { Write-CgsInfo $note }

    $sorted   = Sort-ReportRows -Rows $rows -SortBy $SortBy
    $rows     = $sorted.Rows
    if ($sorted.Note) {
        Write-CgsInfo $sorted.Note
        $note = if ($note) { "$note  " + $sorted.Note } else { $sorted.Note }
    }

    $names = Select-ReportColumns -Rows $rows -Wanted $Columns
    Write-CgsInfo ("keeping {0} column(s): {1}" -f $names.Count, ($names -join ', '))

    $style   = $script:FormatStyles[$format_name]
    $striped = $script:StripedFormats -contains $format_name
    $banner  = if ($Title) { $Title } else {
        [System.IO.Path]::GetFileNameWithoutExtension($ExcelPath) }
    $second  = if ($Subtitle) { $Subtitle } else {
        "{0}  |  {1} row(s)  |  generated {2}" -f $SheetName, $rows.Count,
            (Get-Date).ToString('dd MMM yyyy HH:mm') }

    $extension = [System.IO.Path]::GetExtension($OutputPath).ToLowerInvariant()
    if ($extension -notin @('.docx','.pdf','.html','.htm')) {
        Write-CgsError ("unsupported output extension '$extension'. " +
                        "Use .docx, .pdf, .html or .htm -- the extension " +
                        "chooses the renderer.")
        return 2
    }

    $widths = Get-ColumnWidths -Rows $rows -Names $names

    switch ($extension) {
        { $_ -in '.html', '.htm' } {
            $written = New-HtmlReport -Rows $rows -Names $names -Style $style `
                -Striped $striped -Title $banner -Subtitle $second -Path $OutputPath `
                -MaxLength $MaxCellLength -Note $note
        }
        '.docx' {
            $written = New-DocxBytes -Rows $rows -Names $names -Widths $widths `
                -Style $style -Striped $striped -Title $banner -Subtitle $second `
                -Path $OutputPath -MaxLength $MaxCellLength -Note $note
        }
        '.pdf' {
            $libreOffice = $null
            $haveWord = Test-WordAvailable
            if (-not $haveWord) { $libreOffice = Find-LibreOffice }
            if (-not $haveWord -and -not $libreOffice) {
                $alternative = [System.IO.Path]::ChangeExtension($OutputPath,
                    $(if ($OnWordMissing -eq 'docx') { '.docx' } else { '.html' }))
                if ($OnWordMissing -eq 'fail') {
                    Write-CgsError "cannot produce a PDF: neither Microsoft Word nor LibreOffice is available on this machine."
                    Write-CgsError "  WHY: PowerShell cannot write a PDF by itself. This script makes one by driving Word or LibreOffice, and neither is installed here."
                    Write-CgsError "  FIX 1: run this on a machine with Word installed."
                    Write-CgsError "  FIX 2: install LibreOffice, which converts the same file with no licence needed."
                    Write-CgsError "  FIX 3: ask for .html instead and print it to PDF from any browser -- no software needed."
                    Write-CgsError "  FIX 4: re-run with -OnWordMissing html (or docx) to accept that file instead of a PDF."
                    return 3
                }
                Write-CgsWarn "no PDF converter found; -OnWordMissing $OnWordMissing was set, so writing $alternative instead of a PDF"
                if ($OnWordMissing -eq 'docx') {
                    $written = New-DocxBytes -Rows $rows -Names $names -Widths $widths `
                        -Style $style -Striped $striped -Title $banner -Subtitle $second `
                        -Path $alternative -MaxLength $MaxCellLength -Note $note
                } else {
                    $written = New-HtmlReport -Rows $rows -Names $names -Style $style `
                        -Striped $striped -Title $banner -Subtitle $second -Path $alternative `
                        -MaxLength $MaxCellLength -Note $note
                }
            } else {
                # Both converters read a .docx, so build one first.
                $staging = [System.IO.Path]::ChangeExtension($OutputPath, '.docx')
                [void](New-DocxBytes -Rows $rows -Names $names -Widths $widths `
                    -Style $style -Striped $striped -Title $banner -Subtitle $second `
                    -Path $staging -MaxLength $MaxCellLength -Note $note)
                if ($haveWord) {
                    Write-CgsInfo "converting to PDF with Microsoft Word"
                    $written = Convert-DocxToPdf -DocxPath $staging -PdfPath $OutputPath
                } else {
                    Write-CgsInfo "Word not found; converting to PDF with LibreOffice ($libreOffice)"
                    $written = Convert-DocxToPdfWithLibreOffice -Executable $libreOffice `
                        -DocxPath $staging -PdfPath $OutputPath
                }
                Remove-Item -LiteralPath $staging -Force -ErrorAction SilentlyContinue
            }
        }
    }

    $size = (Get-Item -LiteralPath $written).Length
    Write-CgsInfo ("wrote {0} row(s) x {1} column(s) [{2}] -> {3} ({4:N0} bytes)" -f
                   $rows.Count, $names.Count, $format_name, $written, $size)
    return 0
}

try { exit (Invoke-Main) }
catch { Write-CgsError ("formatExcel failed: {0}" -f $_.Exception.Message); exit 3 }
