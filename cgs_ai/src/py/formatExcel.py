"""
=====================================================================
  Program Name  : formatExcel.py
  Author        : Manuel Figallo
  Purpose       : Turn one worksheet of an .xlsx into a formatted status
                  report for management -- Word, PDF or HTML -- picking
                  the renderer from the OUTPUT file's extension.
  Version       : 1.1beta
  Created       : 2026-10-28
  Last Modified : 2026-10-28

  Dependencies:
    STANDARD LIBRARY ONLY for .docx and .html. The workbook is read with
    src/utils/xlsx.py, which needs no openpyxl, and the Word document is
    written as OOXML directly.

    .pdf IS THE EXCEPTION. Python cannot write a PDF from the standard
    library, so the module drives a converter: Microsoft Word through COM
    (pywin32, Windows only) when it is there, otherwise LibreOffice in
    headless mode, which needs no licence and is often present where Word
    is not. Where NEITHER is installed it does NOT quietly hand back a
    different file -- it raises, naming the alternatives. Pass
    OnWordMissing="html" or "docx" to opt in to a fallback deliberately.

  Description:
    PYTHON TWIN of src/ps/formatExcel.ps1 -- same function name, same
    parameter names, same output. The two write byte-identical HTML and an
    identical word/document.xml, and a test diffs them, so a change to one
    engine that is not made to the other is caught rather than discovered
    by a reader comparing two reports.

    Reads the sheet, keeps only the columns asked for and in that order,
    then renders a titled, banded report. Columns are selected BY NAME, so
    an export that gains a column upstream cannot silently shift the
    report sideways; a name that is not present is an error listing the
    names that are.

  Input Parameters (required first):
    ExcelPath     (REQUIRED, str) source .xlsx.
    SheetName     (REQUIRED, str) worksheet, e.g. "Consolidated Data".
    OutputPath    (REQUIRED, str) destination. Its EXTENSION picks the
                    renderer:
                      .docx -> Word     (native, no Office needed)
                      .html -> web page (native, prints to PDF anywhere)
                      .pdf  -> PDF      (needs Word or LibreOffice)
    Columns       (optional) names to keep, in order. A list, or a ';' or
                    ','-delimited string. Empty = every column.
    format_name   (optional, str, default "corporate")
                    corporate | corporatev2 | plain | minimal
    Title         (optional, str) banner text; defaults to the file name.
    Subtitle      (optional, str) second line; defaults to sheet and count.
    HeaderRow     (optional, int, default 0 = detect) 1-based header row.
    OnWordMissing (optional, str, default "fail") fail | html | docx.
    MaxCellLength (optional, int, default 0 = never shorten) cut long
                    free-text cells, marked with an ellipsis.

  Returns:
    dict with OutputPath, RowCount, ColumnCount, Columns, FormatType and
    Renderer.

  Exit codes (CLI):
    0 = success, 2 = config error, 3 = I/O or renderer failure.
=====================================================================
"""

from __future__ import annotations

import argparse
import os
import platform
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

_HERE = Path(__file__).resolve()
if str(_HERE.parent.parent.parent) not in sys.path:
    sys.path.insert(0, str(_HERE.parent.parent.parent))

from src.utils.logger import logError, logInfo, logWarn  # noqa: E402
from src.utils.xlsx import readSheet                     # noqa: E402

__version__ = "1.1beta"

EXIT_OK, EXIT_CONFIG_ERROR, EXIT_IO_ERROR = 0, 2, 3

#: Palette per format type -- matches $script:FormatStyles in the .ps1 twin
#: and FORMAT_STYLES in formatData. Hex WITHOUT the leading '#'.
FORMAT_STYLES: Dict[str, Dict[str, str]] = {
    "corporate":   {"Banner": "1F3864", "Header": "2E75B6", "Stripe": "DCE6F1",
                    "HeaderFont": "FFFFFF", "BannerFont": "FFFFFF"},
    "corporatev2": {"Banner": "1F3864", "Header": "2E75B6", "Stripe": "DCE6F1",
                    "HeaderFont": "FFFFFF", "BannerFont": "FFFFFF"},
    "plain":       {"Banner": "FFFFFF", "Header": "D9D9D9", "Stripe": "FFFFFF",
                    "HeaderFont": "000000", "BannerFont": "000000"},
    "minimal":     {"Banner": "FFFFFF", "Header": "FFFFFF", "Stripe": "FFFFFF",
                    "HeaderFont": "000000", "BannerFont": "000000"},
}
STRIPED_FORMATS = ("corporate", "corporatev2")
VALID_ON_WORD_MISSING = ("fail", "html", "docx")
SUPPORTED_EXTENSIONS = (".docx", ".pdf", ".html", ".htm")

# Landscape Letter in twentieths of a point. Seven columns of task text do
# not fit portrait, and a report nobody can read is not a report.
PAGE_WIDTH = 15840
PAGE_HEIGHT = 12240
PAGE_MARGIN = 720
TABLE_WIDTH = PAGE_WIDTH - (2 * PAGE_MARGIN)

#: Control characters are illegal in XML 1.0; Word rejects the whole file.
ILLEGAL_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def psInt(value: float) -> int:
    """Round the way PowerShell's [int] cast does.

    Parameters: value (float).
    Returns: int.

    NOT int(), which truncates. PowerShell's [int] is a .NET convert, which
    rounds half to even -- and Python's round() does the same. Getting this
    wrong makes every column width differ by one twip from the .ps1 twin,
    which is invisible on the page and fails the byte-identical test that
    keeps the two engines honest.
    """
    return int(round(value))


def escapeXml(value: Any) -> str:
    """Escape a value for XML text content. Mirrors ConvertTo-CgsXmlText.

    Parameters: value (any) - None becomes an empty string.
    Returns: str with control characters dropped and & < > escaped.
    """
    if value is None:
        return ""
    text = ILLEGAL_XML.sub("", str(value))
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def escapeHtml(value: Any) -> str:
    """Escape a value for HTML text content. Mirrors ConvertTo-HtmlText.

    Parameters: value (any) - None becomes an empty string.
    Returns: str with & < > and double quotes escaped.
    """
    if value is None:
        return ""
    return (str(value).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def asColumnList(value: Any) -> List[str]:
    """Normalise the Columns parameter, whichever shape it arrives in.

    Parameters: value - a list, a ';'-delimited string, or None.
    Returns: list[str] of stripped, non-empty names.

    Mirrors ConvertTo-CgsList: a real list passes through, a string splits
    on ';'. The ',' case is handled in selectReportColumns, where the real
    column names are known.
    """
    if value is None:
        return []
    items = list(value) if isinstance(value, (list, tuple, set)) else \
        str(value).split(";")
    return [str(item).strip() for item in items if str(item).strip()]


def selectReportColumns(rows: Sequence[Dict[str, Any]],
                        wanted: Any) -> List[str]:
    """Keep the requested columns, in the requested order.

    Parameters:
        rows (sequence) - from readSheet.
        wanted          - names to keep; empty keeps everything.
    Returns: list[str] the column names to render.
    Raises: KeyError - a requested name is not in the sheet; the message
            lists the names that ARE, which is far more useful than a
            column of blanks.

    The ',' fallback matches the .ps1 twin: a command line passes
    "a,b,c" as one string, and splitting it is safe only when that whole
    string is not itself a column name.
    """
    if not rows:
        return []
    available = list(rows[0].keys())

    requested = asColumnList(wanted)
    if (len(requested) == 1 and "," in requested[0]
            and requested[0] not in available):
        requested = [part.strip() for part in requested[0].split(",")
                     if part.strip()]
    if not requested:
        return available

    missing = [name for name in requested if name not in available]
    if missing:
        raise KeyError(
            f"column(s) not found in this sheet: {', '.join(missing)}\n"
            f"  the sheet has: {', '.join(available)}")
    return requested


def columnWidths(rows: Sequence[Dict[str, Any]],
                 names: Sequence[str]) -> List[int]:
    """Share the page width out between columns by content length.

    Parameters: rows (sequence); names (sequence) - the chosen columns.
    Returns: list[int] widths in twips, summing to TABLE_WIDTH.

    Sampled from the first 200 rows and clamped, so one long Notes field
    cannot squeeze every other column to nothing. THEN every column is
    given a floor wide enough for its own heading: proportional sharing
    alone put "Priority" in 0.47in and broke the word across two lines,
    and a heading that wraps mid-word makes the report look broken.
    """
    sample = rows[:200]
    weights, floors = [], []
    for name in names:
        widest = len(name)
        for row in sample:
            length = len(str(row.get(name, "")))
            if length > widest:
                widest = length
        # Clamp before sharing out: a 400-character note is not worth 40x
        # the width of a date, however long it is.
        weights.append(max(8, min(45, widest)))
        # ~115 twips per character of 9pt bold Calibri, plus cell padding.
        floors.append(int((len(name) * 115) + 260))

    total = sum(weights)
    widths = [psInt(TABLE_WIDTH * weight / total) for weight in weights]

    # Raise anything under its floor, and recover the difference from the
    # columns that have room to give.
    deficit = 0
    for index in range(len(widths)):
        if widths[index] < floors[index]:
            deficit += floors[index] - widths[index]
            widths[index] = floors[index]
    if deficit > 0:
        donors = [i for i in range(len(widths)) if widths[i] > floors[i]]
        spare = sum(widths[i] - floors[i] for i in donors)
        if spare > 0:
            for index in donors:
                widths[index] -= psInt(
                    deficit * (widths[index] - floors[index]) / spare)

    # Give any rounding remainder to the widest column so the row is exact.
    widestIndex = widths.index(max(widths))
    widths[widestIndex] += TABLE_WIDTH - sum(widths)
    return widths


def limitCellText(value: Any, maxLength: int) -> str:
    """Shorten a value for a report, visibly.

    Parameters: value (any); maxLength (int) - 0 never shortens.
    Returns: str, with a trailing ellipsis when cut.

    Off by default. Truncating a report silently would hide data from the
    people relying on it, so the cut is always marked and the caller has
    to ask for it.
    """
    text = "" if value is None else str(value)
    if maxLength <= 0 or len(text) <= maxLength:
        return text
    return text[:max(1, maxLength - 1)] + "\u2026"


def buildDocumentXml(rows: Sequence[Dict[str, Any]], names: Sequence[str],
                     widths: Sequence[int], style: Dict[str, str],
                     striped: bool, title: str, subtitle: str,
                     maxLength: int = 0) -> str:
    """Build word/document.xml.

    Parameters: as newDocx.
    Returns: str the document part.

    Kept separate from the zipping so a test can diff it against the
    .ps1 twin's output without unpacking two archives.
    """
    parts: List[str] = []

    # Title banner: a one-cell shaded table, which prints as a solid band
    # where a shaded paragraph leaves white gutters at the margins.
    parts.append(f'<w:tbl><w:tblPr><w:tblW w:w="{TABLE_WIDTH}" w:type="dxa"/>')
    parts.append('<w:tblLayout w:type="fixed"/></w:tblPr>')
    parts.append(f'<w:tblGrid><w:gridCol w:w="{TABLE_WIDTH}"/></w:tblGrid>')
    parts.append(f'<w:tr><w:tc><w:tcPr><w:tcW w:w="{TABLE_WIDTH}" w:type="dxa"/>')
    parts.append(f'<w:shd w:val="clear" w:fill="{style["Banner"]}"/>')
    parts.append('<w:tcMar><w:top w:w="140" w:type="dxa"/><w:left w:w="180" '
                 'w:type="dxa"/><w:bottom w:w="140" w:type="dxa"/></w:tcMar></w:tcPr>')
    parts.append('<w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:rPr><w:b/>'
                 f'<w:sz w:val="32"/><w:color w:val="{style["BannerFont"]}"/></w:rPr>')
    parts.append(f'<w:t xml:space="preserve">{escapeXml(title)}</w:t></w:r></w:p>')
    if subtitle:
        parts.append('<w:p><w:pPr><w:spacing w:before="40" w:after="0"/></w:pPr>'
                     f'<w:r><w:rPr><w:sz w:val="19"/>'
                     f'<w:color w:val="{style["BannerFont"]}"/></w:rPr>')
        parts.append(f'<w:t xml:space="preserve">{escapeXml(subtitle)}'
                     '</w:t></w:r></w:p>')
    parts.append('</w:tc></w:tr></w:tbl>')
    parts.append('<w:p><w:pPr><w:spacing w:after="160"/></w:pPr></w:p>')

    borders = "<w:tblBorders>"
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        borders += (f'<w:{edge} w:val="single" w:sz="4" w:space="0" '
                    f'w:color="BFBFBF"/>')
    borders += "</w:tblBorders>"
    parts.append(f'<w:tbl><w:tblPr><w:tblW w:w="{TABLE_WIDTH}" w:type="dxa"/>')
    parts.append(borders + '<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>')
    for width in widths:
        parts.append(f'<w:gridCol w:w="{width}"/>')
    parts.append("</w:tblGrid>")

    # tblHeader repeats the header row at the top of every printed page --
    # without it a three-page report has unlabelled columns after page one.
    parts.append('<w:tr><w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>')
    for index, name in enumerate(names):
        parts.append(f'<w:tc><w:tcPr><w:tcW w:w="{widths[index]}" w:type="dxa"/>')
        parts.append(f'<w:shd w:val="clear" w:fill="{style["Header"]}"/></w:tcPr>')
        parts.append('<w:p><w:pPr><w:spacing w:after="40"/></w:pPr><w:r><w:rPr>'
                     f'<w:b/><w:sz w:val="18"/>'
                     f'<w:color w:val="{style["HeaderFont"]}"/></w:rPr>')
        parts.append(f'<w:t xml:space="preserve">{escapeXml(name)}'
                     '</w:t></w:r></w:p></w:tc>')
    parts.append("</w:tr>")

    for rowNumber, row in enumerate(rows, start=1):
        fill = style["Stripe"] if striped and rowNumber % 2 == 1 else ""
        parts.append("<w:tr>")
        for index, name in enumerate(names):
            parts.append(f'<w:tc><w:tcPr><w:tcW w:w="{widths[index]}" '
                         'w:type="dxa"/>')
            if fill:
                parts.append(f'<w:shd w:val="clear" w:fill="{fill}"/>')
            parts.append('</w:tcPr><w:p><w:pPr><w:spacing w:after="30"/></w:pPr>'
                         '<w:r><w:rPr><w:sz w:val="17"/></w:rPr>')
            value = limitCellText(row.get(name, ""), maxLength)
            parts.append(f'<w:t xml:space="preserve">{escapeXml(value)}'
                         '</w:t></w:r></w:p></w:tc>')
        parts.append("</w:tr>")
    parts.append("</w:tbl>")

    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/'
            'wordprocessingml/2006/main">'
            "<w:body>" + "".join(parts) +
            f'<w:sectPr><w:pgSz w:w="{PAGE_WIDTH}" w:h="{PAGE_HEIGHT}" '
            'w:orient="landscape"/>'
            f'<w:pgMar w:top="{PAGE_MARGIN}" w:right="{PAGE_MARGIN}" '
            f'w:bottom="{PAGE_MARGIN}" w:left="{PAGE_MARGIN}"/>'
            "</w:sectPr></w:body></w:document>")


STYLES_XML = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<w:styles xmlns:w="http://schemas.openxmlformats.org/'
              'wordprocessingml/2006/main">'
              '<w:docDefaults><w:rPrDefault><w:rPr>'
              '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>'
              '<w:sz w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:styleId="Normal" w:default="1">'
              '<w:name w:val="Normal"/></w:style>'
              "</w:styles>")


def newDocx(rows: Sequence[Dict[str, Any]], names: Sequence[str],
            widths: Sequence[int], style: Dict[str, str], striped: bool,
            title: str, subtitle: str, path: str,
            maxLength: int = 0) -> str:
    """Build a .docx as OOXML, with no Word and no add-in.

    Parameters:
        rows, names, widths, style, striped, title, subtitle - the content.
        path (str)      - destination .docx.
        maxLength (int) - cut long cells; 0 never does.
    Returns: str the path written.
    Raises: OSError - the destination cannot be written.
    """
    parts = {
        "[Content_Types].xml":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
        "_rels/.rels":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
        "word/_rels/document.xml.rels":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
        "word/styles.xml": STYLES_XML,
        "word/document.xml": buildDocumentXml(rows, names, widths, style,
                                              striped, title, subtitle,
                                              maxLength),
    }
    destination = Path(path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in parts.items():
            # UTF-8 with NO byte-order mark: Word rejects a BOM in a part.
            archive.writestr(name, content.encode("utf-8"))
    return str(destination)


def newHtmlReport(rows: Sequence[Dict[str, Any]], names: Sequence[str],
                  style: Dict[str, str], striped: bool, title: str,
                  subtitle: str, path: str, maxLength: int = 0) -> str:
    """Write the report as a self-contained web page.

    Parameters: as newDocx, without widths (the browser lays it out).
    Returns: str the path written.

    Every style is inline in one file -- no CSS to lose when the page is
    emailed. Any browser prints it to PDF, which is the fallback when no
    converter is installed.
    """
    stripe = style["Stripe"] if striped else "FFFFFF"
    html = ['<!doctype html><html><head><meta charset="utf-8">',
            f"<title>{escapeHtml(title)}</title><style>",
            "body{font-family:Calibri,'Segoe UI',Arial,sans-serif;"
            "margin:24px;color:#1F2937}",
            f'.banner{{background:#{style["Banner"]};'
            f'color:#{style["BannerFont"]};padding:14px 18px;border-radius:3px}}',
            ".banner h1{margin:0;font-size:20pt}"
            ".banner p{margin:4px 0 0;font-size:10pt;opacity:.9}",
            "table{border-collapse:collapse;width:100%;margin-top:16px;"
            "font-size:9.5pt}",
            f'th{{background:#{style["Header"]};color:#{style["HeaderFont"]};'
            "text-align:left;padding:7px 8px;border:1px solid #BFBFBF}",
            "td{padding:6px 8px;border:1px solid #BFBFBF;vertical-align:top}",
            f"tbody tr:nth-child(odd){{background:#{stripe}}}",
            "@media print{@page{size:landscape;margin:12mm}"
            "thead{display:table-header-group}}",
            "</style></head><body>",
            f'<div class="banner"><h1>{escapeHtml(title)}</h1>']
    if subtitle:
        html.append(f"<p>{escapeHtml(subtitle)}</p>")
    html.append("</div><table><thead><tr>")
    for name in names:
        html.append(f"<th>{escapeHtml(name)}</th>")
    html.append("</tr></thead><tbody>")
    for row in rows:
        html.append("<tr>")
        for name in names:
            value = limitCellText(row.get(name, ""), maxLength)
            html.append(f"<td>{escapeHtml(value)}</td>")
        html.append("</tr>")
    html.append("</tbody></table></body></html>")

    destination = Path(path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text("".join(html), encoding="utf-8")
    return str(destination)


def wordAvailable() -> bool:
    """Is Microsoft Word installed and reachable through COM?

    Parameters: none. Returns: bool.

    COM automation is Windows-only and needs pywin32, so this is False
    everywhere else without trying.
    """
    if platform.system() != "Windows":
        return False
    try:
        import win32com.client
        word = win32com.client.Dispatch("Word.Application")
        word.Quit()
        return True
    except Exception:                       # noqa: BLE001 - absence, not error
        return False


def findLibreOffice() -> Optional[str]:
    """Locate the LibreOffice command line, if it is installed.

    Parameters: none.
    Returns: str the executable path, or None.

    Worth looking for: on a locked-down server Word is often absent while
    LibreOffice, which needs no licence, is not.
    """
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    for candidate in (
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if os.path.isfile(candidate):
            return candidate
    return None


def convertDocxToPdfWithWord(docxPath: str, pdfPath: str) -> str:
    """Save a .docx as .pdf by driving Word through COM.

    Parameters: docxPath (str); pdfPath (str).
    Returns: str the .pdf path.
    Raises: RuntimeError - the conversion failed.

    wdFormatPDF = 17. Word is quit in a finally block: an instance left
    running holds the file open and accumulates invisible WINWORD.EXE
    processes on a server.
    """
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    document = None
    try:
        document = word.Documents.Open(os.path.abspath(docxPath),
                                       False, True)   # read-only
        document.SaveAs(os.path.abspath(pdfPath), 17)
    finally:
        if document is not None:
            document.Close(False)
        word.Quit()
    return pdfPath


def convertDocxToPdfWithLibreOffice(executable: str, docxPath: str,
                                    pdfPath: str) -> str:
    """Save a .docx as .pdf using LibreOffice in headless mode.

    Parameters: executable (str); docxPath (str); pdfPath (str).
    Returns: str the .pdf path.
    Raises: RuntimeError - LibreOffice produced nothing.

    LibreOffice always names its output after the INPUT file and drops it
    in --outdir, so when the caller asked for a different name the result
    is renamed afterwards.
    """
    outputDirectory = os.path.dirname(os.path.abspath(pdfPath)) or "."
    subprocess.run([executable, "--headless", "--convert-to", "pdf",
                    "--outdir", outputDirectory, os.path.abspath(docxPath)],
                   capture_output=True, text=True, timeout=600, check=False)
    produced = os.path.join(outputDirectory,
                            Path(docxPath).stem + ".pdf")
    if not os.path.isfile(produced):
        raise RuntimeError(f"LibreOffice did not produce a PDF from {docxPath}")
    if os.path.abspath(produced) != os.path.abspath(pdfPath):
        shutil.move(produced, pdfPath)
    return pdfPath


def formatExcel(ExcelPath: str, SheetName: str, OutputPath: str,
                Columns: Any = (), format_name: str = "corporate",
                Title: str = "", Subtitle: str = "", HeaderRow: int = 0,
                OnWordMissing: str = "fail",
                MaxCellLength: int = 0) -> Dict[str, Any]:
    """Turn one worksheet into a formatted status report.

    Parameters:
        ExcelPath (str)     - REQUIRED source .xlsx.
        SheetName (str)     - REQUIRED worksheet to format.
        OutputPath (str)    - REQUIRED destination; its extension picks the
                              renderer (.docx, .html, .pdf).
        Columns             - names to keep, in order; empty keeps all.
        format_name (str)   - corporate | corporatev2 | plain | minimal.
        Title (str)         - banner text; defaults to the file name.
        Subtitle (str)      - second line; defaults to sheet and row count.
        HeaderRow (int)     - 1-based header row; 0 detects it.
        OnWordMissing (str) - fail | html | docx, when no PDF converter is
                              installed.
        MaxCellLength (int) - shorten long cells; 0 never does.
    Returns:
        dict with OutputPath, RowCount, ColumnCount, Columns, FormatType
        and Renderer.
    Raises:
        ValueError   - a required parameter is missing or an option unknown.
        FileNotFoundError - the workbook is not there.
        KeyError     - a requested column or the sheet is absent.
        RuntimeError - a .pdf was asked for and nothing can produce one.

    Use in claims processing:
        Publish a weekly governance or issue-log status pack straight from
        the tracking workbook, with no manual formatting each cycle.
    """
    if not str(ExcelPath or "").strip():
        raise ValueError("required parameter 'ExcelPath' is missing or empty")
    if not str(SheetName or "").strip():
        raise ValueError("required parameter 'SheetName' is missing or empty")
    if not str(OutputPath or "").strip():
        raise ValueError("required parameter 'OutputPath' is missing or empty")
    if format_name not in FORMAT_STYLES:
        raise ValueError(f"unknown format_name '{format_name}'; expected one "
                         f"of: {', '.join(sorted(FORMAT_STYLES))}")
    if OnWordMissing not in VALID_ON_WORD_MISSING:
        raise ValueError(f"unknown OnWordMissing '{OnWordMissing}'; expected "
                         f"fail, html or docx")
    if not os.path.isfile(ExcelPath):
        raise FileNotFoundError(f"input workbook not found: {ExcelPath}")

    extension = Path(OutputPath).suffix.lower()
    if extension not in SUPPORTED_EXTENSIONS:
        raise ValueError(f"unsupported output extension '{extension}'. Use "
                         f".docx, .pdf, .html or .htm -- the extension "
                         f"chooses the renderer.")

    logInfo(f"formatExcel {__version__} starting")
    rows = readSheet(ExcelPath, sheet=SheetName, headerRow=HeaderRow)
    logInfo(f"read {len(rows)} row(s) from sheet '{SheetName}'")

    names = selectReportColumns(rows, Columns)
    logInfo(f"keeping {len(names)} column(s): {', '.join(names)}")

    style = FORMAT_STYLES[format_name]
    striped = format_name in STRIPED_FORMATS
    banner = Title or Path(ExcelPath).stem
    second = Subtitle or (f"{SheetName}  |  {len(rows)} row(s)  |  generated "
                          f"{__import__('datetime').datetime.now():%d %b %Y %H:%M}")
    widths = columnWidths(rows, names)

    if extension in (".html", ".htm"):
        written = newHtmlReport(rows, names, style, striped, banner, second,
                                OutputPath, MaxCellLength)
        renderer = "html"
    elif extension == ".docx":
        written = newDocx(rows, names, widths, style, striped, banner, second,
                          OutputPath, MaxCellLength)
        renderer = "docx"
    else:
        haveWord = wordAvailable()
        libreOffice = None if haveWord else findLibreOffice()
        if not haveWord and not libreOffice:
            alternative = str(Path(OutputPath).with_suffix(
                ".docx" if OnWordMissing == "docx" else ".html"))
            if OnWordMissing == "fail":
                raise RuntimeError(
                    "cannot produce a PDF: neither Microsoft Word nor "
                    "LibreOffice is available on this machine.\n"
                    "  WHY: Python cannot write a PDF from the standard "
                    "library. This module makes one by driving Word or "
                    "LibreOffice, and neither is installed here.\n"
                    "  FIX 1: run this on a machine with Word installed.\n"
                    "  FIX 2: install LibreOffice, which converts the same "
                    "file with no licence needed.\n"
                    "  FIX 3: ask for .html instead and print it to PDF from "
                    "any browser -- no software needed.\n"
                    "  FIX 4: pass OnWordMissing='html' (or 'docx') to accept "
                    "that file instead of a PDF.")
            logWarn(f"no PDF converter found; OnWordMissing={OnWordMissing} "
                    f"was set, so writing {alternative} instead of a PDF")
            if OnWordMissing == "docx":
                written = newDocx(rows, names, widths, style, striped, banner,
                                  second, alternative, MaxCellLength)
                renderer = "docx-fallback"
            else:
                written = newHtmlReport(rows, names, style, striped, banner,
                                        second, alternative, MaxCellLength)
                renderer = "html-fallback"
        else:
            # Both converters read a .docx, so build one first.
            staging = str(Path(OutputPath).with_suffix(".docx"))
            newDocx(rows, names, widths, style, striped, banner, second,
                    staging, MaxCellLength)
            try:
                if haveWord:
                    logInfo("converting to PDF with Microsoft Word")
                    written = convertDocxToPdfWithWord(staging, OutputPath)
                    renderer = "pdf-word"
                else:
                    logInfo(f"Word not found; converting to PDF with "
                            f"LibreOffice ({libreOffice})")
                    written = convertDocxToPdfWithLibreOffice(
                        libreOffice, staging, OutputPath)
                    renderer = "pdf-libreoffice"
            finally:
                # The staging file is an implementation detail; a PDF
                # request must leave a PDF and nothing else beside it.
                if os.path.isfile(staging) and os.path.abspath(staging) != \
                        os.path.abspath(OutputPath):
                    os.remove(staging)

    size = os.path.getsize(written)
    logInfo(f"wrote {len(rows)} row(s) x {len(names)} column(s) "
            f"[{format_name}] -> {written} ({size:,} bytes)")
    return {"OutputPath": written, "RowCount": len(rows),
            "ColumnCount": len(names), "Columns": names,
            "FormatType": format_name, "Renderer": renderer}


def buildArgParser() -> argparse.ArgumentParser:
    """Build the command-line parser. Parameters: none."""
    parser = argparse.ArgumentParser(
        prog="formatExcel.py",
        description="Turn one worksheet into a formatted status report.")
    parser.add_argument("--excel-path", required=True)
    parser.add_argument("--sheet-name", required=True)
    parser.add_argument("--output-path", required=True)
    parser.add_argument("--columns", default="")
    parser.add_argument("--format-name", default="corporate")
    parser.add_argument("--title", default="")
    parser.add_argument("--subtitle", default="")
    parser.add_argument("--header-row", type=int, default=0)
    parser.add_argument("--on-word-missing", default="fail",
                        choices=list(VALID_ON_WORD_MISSING))
    parser.add_argument("--max-cell-length", type=int, default=0)
    parser.add_argument("--version", action="version",
                        version=f"formatExcel {__version__}")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    """CLI entry point. Parameters: argv (sequence[str]). Returns: exit code."""
    args = buildArgParser().parse_args(argv)
    try:
        formatExcel(ExcelPath=args.excel_path, SheetName=args.sheet_name,
                    OutputPath=args.output_path, Columns=args.columns,
                    format_name=args.format_name, Title=args.title,
                    Subtitle=args.subtitle, HeaderRow=args.header_row,
                    OnWordMissing=args.on_word_missing,
                    MaxCellLength=args.max_cell_length)
    except ValueError as exc:
        logError(str(exc))
        return EXIT_CONFIG_ERROR
    except (OSError, KeyError, RuntimeError) as exc:
        # KeyError's str() adds quotes around a multi-line message, which
        # makes the column list unreadable; use the argument itself.
        message = exc.args[0] if isinstance(exc, KeyError) else str(exc)
        for line in str(message).splitlines():
            logError(line)
        return EXIT_IO_ERROR
    return EXIT_OK


if __name__ == "__main__":
    sys.exit(main())
