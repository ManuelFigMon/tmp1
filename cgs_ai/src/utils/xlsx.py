"""
=====================================================================
  Program Name  : xlsx.py
  Author        : Manuel Figallo
  Purpose       : Read and write .xlsx with the standard library alone, so
                  a workbook can be produced -- and consumed -- on a server
                  that has neither openpyxl nor the ImportExcel module.
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


# =====================================================================
# Reading
# =====================================================================

MAIN_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PKG_REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"

CELL_REF = re.compile(r"^([A-Z]+)")


def columnIndex(cellRef: str) -> int:
    """Convert a cell reference to a 0-based column index (A1 -> 0, AA3 -> 26).

    Parameters: cellRef (str) - e.g. "AA12".
    Returns: int 0-based column index; 0 when the reference is unusable.
    """
    letters = CELL_REF.match(str(cellRef or "").upper())
    if not letters:
        return 0
    index = 0
    for character in letters.group(1):
        index = index * 26 + (ord(character) - 64)
    return index - 1


def _sharedStrings(archive: zipfile.ZipFile) -> List[str]:
    """Read xl/sharedStrings.xml. Returns: list[str], empty when absent.

    A shared string can be split across several runs (<r><t>..</t></r>), so
    every <t> under an <si> is concatenated -- taking only the first would
    silently truncate any cell Excel happened to style mid-word.
    """
    import xml.etree.ElementTree as ET
    try:
        raw = archive.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    return ["".join(node.text or "" for node in item.iter(f"{MAIN_NS}t"))
            for item in ET.fromstring(raw).findall(f"{MAIN_NS}si")]


def _sheetPaths(archive: zipfile.ZipFile) -> List[Dict[str, str]]:
    """Map worksheet names to their parts, in workbook order.

    Parameters: archive (ZipFile) - the open .xlsx.
    Returns: list of {"name", "path"}.
    """
    import xml.etree.ElementTree as ET
    relations = {}
    try:
        relsXml = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        for node in relsXml.findall(f"{PKG_REL_NS}Relationship"):
            target = node.get("Target", "")
            relations[node.get("Id", "")] = (
                target[1:] if target.startswith("/") else f"xl/{target}")
    except KeyError:
        pass

    found: List[Dict[str, str]] = []
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    for index, node in enumerate(workbook.iter(f"{MAIN_NS}sheet"), start=1):
        path = relations.get(node.get(f"{REL_NS}id", ""))
        if path is None or path not in archive.namelist():
            path = f"xl/worksheets/sheet{index}.xml"
        found.append({"name": node.get("name", f"Sheet{index}"), "path": path})
    return found


def readSheetRows(path: str, sheet: Any = None) -> List[List[str]]:
    """Read one worksheet as a grid of strings, with the standard library.

    Parameters:
        path (str)      - the .xlsx to read.
        sheet (str|int) - worksheet name, or 0-based index; default the first.
    Returns:
        list of rows, each a list of cell values as strings. Short rows are
        padded so every row has the same length.
    Raises:
        OSError / KeyError - the file is not a readable .xlsx.

    Values come back as STRINGS on purpose: a claim number stored as text
    keeps its leading zeros, and this reader feeds a formatter, not a
    calculation.
    """
    import xml.etree.ElementTree as ET
    with zipfile.ZipFile(path) as archive:
        strings = _sharedStrings(archive)
        sheets = _sheetPaths(archive)
        if not sheets:
            return []
        if sheet is None or sheet == "":
            chosen = sheets[0]
        elif isinstance(sheet, int):
            chosen = sheets[sheet]
        else:
            matches = [s for s in sheets if s["name"] == sheet]
            if not matches:
                raise KeyError(
                    f"worksheet {sheet!r} not found; this workbook has: "
                    f"{', '.join(s['name'] for s in sheets)}")
            chosen = matches[0]
        tree = ET.fromstring(archive.read(chosen["path"]))

    grid: List[List[str]] = []
    for rowNode in tree.iter(f"{MAIN_NS}row"):
        cells: List[str] = []
        for cellNode in rowNode.findall(f"{MAIN_NS}c"):
            index = columnIndex(cellNode.get("r", ""))
            while len(cells) < index:      # sparse rows omit empty cells
                cells.append("")
            kind = cellNode.get("t", "n")
            if kind == "s":
                value = cellNode.find(f"{MAIN_NS}v")
                position = int(value.text) if value is not None and value.text else -1
                text = strings[position] if 0 <= position < len(strings) else ""
            elif kind == "inlineStr":
                node = cellNode.find(f"{MAIN_NS}is")
                text = "".join(t.text or "" for t in node.iter(f"{MAIN_NS}t")) \
                    if node is not None else ""
            else:
                value = cellNode.find(f"{MAIN_NS}v")
                text = value.text if value is not None and value.text else ""
            cells.append(text)
        grid.append(cells)

    width = max((len(row) for row in grid), default=0)
    for row in grid:
        row.extend([""] * (width - len(row)))
    return grid


def detectHeaderRow(grid: Sequence[Sequence[str]]) -> int:
    """Work out which row holds the column headers.

    Parameters: grid (sequence) - rows from readSheetRows.
    Returns: int the 1-based header row number.

    A workbook this package produced has a merged TITLE BANNER on row 1 and
    the headers on row 2. That shows up as a first row with a single filled
    cell above a much wider second row, which is what is detected here, so a
    formatted workbook can be fed straight back in without a parameter.
    """
    if len(grid) < 2:
        return 1
    filledFirst = sum(1 for cell in grid[0] if str(cell).strip())
    filledSecond = sum(1 for cell in grid[1] if str(cell).strip())
    return 2 if filledFirst <= 1 < filledSecond else 1


def readSheet(path: str, sheet: Any = None,
              headerRow: int = 0) -> List[Dict[str, str]]:
    """Read one worksheet as a list of dicts keyed by column header.

    Parameters:
        path (str)      - the .xlsx to read.
        sheet (str|int) - worksheet name or 0-based index; default the first.
        headerRow (int) - 1-based header row; 0 (default) detects it.
    Returns:
        list[dict] in sheet order, shaped like readCsv's output so either
        input can feed the same formatter.

    Use in claims processing:
        Re-format a workbook a colleague sent, or re-style one this package
        produced last cycle, without converting it back to CSV first.
    """
    grid = readSheetRows(path, sheet)
    if not grid:
        return []
    header = headerRow if headerRow > 0 else detectHeaderRow(grid)
    header = max(1, min(header, len(grid)))
    columns = [str(cell).strip() for cell in grid[header - 1]]
    # Unnamed trailing columns would collide on the empty-string key.
    columns = [name or f"Column{index}"
               for index, name in enumerate(columns, start=1)]
    return [dict(zip(columns, row)) for row in grid[header:]]
