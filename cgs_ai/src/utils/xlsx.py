"""
=====================================================================
  Program Name  : xlsx.py
  Author        : Manuel Figallo
  Purpose       : Write a multi-sheet .xlsx with the standard library
                  alone, so a workbook can be produced on a server that
                  has neither openpyxl nor the ImportExcel module.
  Version       : 1.0beta
  Created       : 2026-09-06
  Last Modified : 2026-09-06

  Dependencies:
    STANDARD LIBRARY ONLY (zipfile, xml escaping done here).

  Description:
    An .xlsx is a zip of XML parts. This writes those parts directly.
    Values go in as INLINE strings, which avoids a shared-string table and
    keeps every cell exactly as it was handed over -- leading zeros in a
    claim number survive, where a numeric conversion would eat them.

    Mirrored by Write-CgsXlsx in src/ps/cgsUtils.ps1. The two produce the
    same XML; a test compares the two implementations' structure.

  Input Parameters (required first):
    sheets     (REQUIRED, list[dict]) - each {name, columns, rows}.
    outputPath (REQUIRED, str)        - destination .xlsx.

  Usage:
      from src.utils.xlsx import writeWorkbook
      writeWorkbook([{"name": "Matches", "columns": cols, "rows": rows}],
                    "out.xlsx")
=====================================================================
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Sequence

__version__ = "1.0beta"

#: Excel refuses these in a sheet name, and caps the name at 31 characters.
ILLEGAL_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")
#: Control characters are illegal in XML 1.0; Excel rejects the whole file.
ILLEGAL_XML_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

MAX_COLUMN_WIDTH = 60
WIDTH_SAMPLE_ROWS = 200


def escapeXml(value: Any) -> str:
    """Escape a value for XML text content.

    Parameters: value (any) - None becomes an empty string.
    Returns: str with control characters dropped and & < > escaped.
    """
    if value is None:
        return ""
    text = ILLEGAL_XML_CHARS.sub("", str(value))
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def columnRef(index: int) -> str:
    """Convert a 1-based column number to letters (1 -> A, 27 -> AA).

    Parameters: index (int) - 1-based column number.
    Returns: str the column reference.
    """
    ref = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        ref = chr(65 + remainder) + ref
    return ref


def safeSheetName(name: str, used: Sequence[str] = ()) -> str:
    """Make a worksheet name Excel will accept, and unique in the workbook.

    Parameters:
        name (str)      - the requested name.
        used (sequence) - names already taken.
    Returns: str at most 31 characters, none of []:*?/\\ , not empty.
    """
    cleaned = ILLEGAL_SHEET_CHARS.sub("_", str(name or "Sheet")).strip("'")[:31]
    if not cleaned:
        cleaned = "Sheet"
    candidate, suffix = cleaned, 2
    while candidate in used:
        tail = f"_{suffix}"
        candidate = cleaned[:31 - len(tail)] + tail
        suffix += 1
    return candidate


def stylesXml() -> str:
    """Build xl/styles.xml: 0 = plain, 1 = bold header on a grey fill.

    Parameters: none. Returns: str the styles part.

    Fill 0 MUST be none and fill 1 MUST be gray125 -- Excel rejects the
    workbook otherwise, however unused those two entries are.
    """
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>'
        '<font><b/><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="3"><fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid">'
        '<fgColor rgb="FFD9E1F2"/><bgColor indexed="64"/></patternFill></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="2">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" '
        'applyFont="1" applyFill="1"/>'
        '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '</styleSheet>')


def sheetXml(columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> str:
    """Build one xl/worksheets/sheetN.xml.

    Parameters:
        columns (sequence) - column order; also the header row.
        rows (sequence)    - dicts keyed by column name.
    Returns: str the worksheet part.

    The child order below is fixed by the schema -- dimension, sheetViews,
    sheetFormatPr, cols, sheetData, autoFilter. Excel refuses the file when
    they are out of order, with no useful message.
    """
    columns = list(columns) or ["(no columns)"]
    lastCol = columnRef(len(columns))
    lastRow = len(rows) + 1

    body = ['<row r="1">']
    for index, name in enumerate(columns, start=1):
        body.append(f'<c r="{columnRef(index)}1" s="1" t="inlineStr">'
                    f'<is><t xml:space="preserve">{escapeXml(name)}</t></is></c>')
    body.append("</row>")

    for number, row in enumerate(rows, start=2):
        body.append(f'<row r="{number}">')
        for index, name in enumerate(columns, start=1):
            body.append(
                f'<c r="{columnRef(index)}{number}" s="0" t="inlineStr">'
                f'<is><t xml:space="preserve">{escapeXml(row.get(name, ""))}'
                f'</t></is></c>')
        body.append("</row>")

    sample = rows[:WIDTH_SAMPLE_ROWS]
    widths = []
    for index, name in enumerate(columns, start=1):
        widest = max([len(str(name))]
                     + [len(str(row.get(name, ""))) for row in sample] or [0])
        width = min(MAX_COLUMN_WIDTH, max(10, widest + 2))
        widths.append(f'<col min="{index}" max="{index}" width="{width}" '
                      f'customWidth="1"/>')

    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f'<dimension ref="A1:{lastCol}{lastRow}"/>'
            '<sheetViews><sheetView workbookViewId="0">'
            '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
            '</sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>'
            f'<cols>{"".join(widths)}</cols>'
            f'<sheetData>{"".join(body)}</sheetData>'
            f'<autoFilter ref="A1:{lastCol}{lastRow}"/>'
            '</worksheet>')


def writeWorkbook(sheets: Sequence[Dict[str, Any]], outputPath: str) -> str:
    """Write a multi-sheet .xlsx using the standard library alone.

    Parameters:
        sheets (sequence) - each {name (str), columns (seq), rows (seq[dict])}.
                            An empty list still produces a valid one-sheet
                            workbook, because Excel rejects a sheetless file.
        outputPath (str)  - destination .xlsx (parent folders are created).
    Returns: str the path written.
    Raises: OSError when the destination cannot be written.

    Use in claims processing:
        Emit a scan's matches and its metrics as two worksheets of one
        workbook on a locked-down server, with no Excel and no add-ins.
    """
    prepared: List[Dict[str, Any]] = []
    for sheet in (sheets or [{"name": "Sheet1", "columns": [], "rows": []}]):
        prepared.append({
            "name": safeSheetName(sheet.get("name", "Sheet"),
                                  [s["name"] for s in prepared]),
            "columns": list(sheet.get("columns") or []),
            "rows": list(sheet.get("rows") or []),
        })

    sheetTags, relTags, overrides = [], [], []
    for index, sheet in enumerate(prepared, start=1):
        sheetTags.append(f'<sheet name="{escapeXml(sheet["name"])}" '
                         f'sheetId="{index}" r:id="rId{index}"/>')
        relTags.append(
            f'<Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org'
            f'/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{index}.xml"/>')
        overrides.append(
            f'<Override PartName="/xl/worksheets/sheet{index}.xml" '
            f'ContentType="application/vnd.openxmlformats-officedocument'
            f'.spreadsheetml.worksheet+xml"/>')
    stylesId = len(prepared) + 1

    parts = {
        "[Content_Types].xml":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            + "".join(overrides) +
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '</Types>',
        "_rels/.rels":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>',
        "xl/workbook.xml":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'<sheets>{"".join(sheetTags)}</sheets></workbook>',
        "xl/_rels/workbook.xml.rels":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(relTags) +
            f'<Relationship Id="rId{stylesId}" Type="http://schemas.openxmlformats.org'
            f'/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '</Relationships>',
        "xl/styles.xml": stylesXml(),
    }
    for index, sheet in enumerate(prepared, start=1):
        parts[f"xl/worksheets/sheet{index}.xml"] = sheetXml(
            sheet["columns"], sheet["rows"])

    destination = Path(outputPath)
    if destination.parent and not destination.parent.exists():
        destination.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in parts.items():
            # UTF-8 with NO byte-order mark: Excel rejects a BOM inside parts.
            archive.writestr(name, content.encode("utf-8"))
    return str(destination)
