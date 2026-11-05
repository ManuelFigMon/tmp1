"""
=====================================================================
  Program Name  : build_tests_workbook.py
  Author        : Manuel Figallo
  Purpose       : Build python_for_medicare_claims_pipelines_TESTS1.xlsx --
                  a workbook that loads the cgs_ai package from a worksheet
                  and then calls it the way you would in VS Code.
  Version       : 1.0beta
  Created       : 2026-10-29
  Last Modified : 2026-10-29

  Dependencies:
    STANDARD LIBRARY ONLY. Scaffolding parts (styles, theme, and the Python
    environment definition) are copied out of an existing workbook so the
    =PY() plumbing is known-good rather than hand-rolled.

  Description:
    THE PROBLEM THIS TESTS. Python in Excel runs in a Microsoft-managed
    container with no filesystem and no network, so "import cgs_ai" from a
    OneDrive path can never resolve there. But `import` checks sys.modules
    BEFORE it goes near a disk, and the lite build's last line, _register(),
    puts a module object into sys.modules under the name "cgs_ai". So if one
    cell runs that source first, every later cell can do an ordinary import.

    WHY A WHOLE SHEET. The lite __init__.py is 37,732 characters and Excel's
    cell limit is 32,767 -- it does not fit in one cell, and splitting a
    Python file across cells at arbitrary points does not work. One LINE PER
    ROW sidesteps the limit completely: the longest line in the file is 84
    characters, so no function can ever outgrow a cell no matter how much
    the package gains.

    CALCULATION ORDER IS THE WHOLE DESIGN. Excel evaluates =PY() cells in one
    Python session, in workbook order: row-major within a sheet, then sheet
    by sheet LEFT TO RIGHT across the tabs. The namespace dependency is
    invisible to Excel -- it is not a cell reference -- so nothing reorders
    it for you. That is why "__init__" is the second tab and the TEST sheets
    follow it. Drag "__init__" to the right of a TEST sheet and that test
    breaks.

  Input Parameters (required first):
    template (REQUIRED) - a workbook to copy styles/theme/env from
    target   (REQUIRED) - the .xlsx to write

  Usage:
    python build_tests_workbook.py <template.xlsx> <target.xlsx>

  Exit codes:
    0 = success, 2 = usage error.
=====================================================================
"""

from __future__ import annotations

import math
import random
import sys
import zipfile
from pathlib import Path
from typing import Dict, List, Sequence, Tuple

__version__ = "1.0beta"

EXIT_OK, EXIT_USAGE = 0, 2

#: Where the package source is read from, with room to grow. The sheet uses
#: 869 rows today; the loader trims the empty tail itself, so this never has
#: to be edited when the package gains a function.
SOURCE_RANGE = "'__init__'!$A$1:$A$1200"
SOURCE_ROWS = 1200

#: The claims data TEST2 samples.
CLAIMS_RANGE = "Claims_Data!$A$1:$N$51"

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_PY = "http://schemas.microsoft.com/office/spreadsheetml/2023/python"

# Style ids inherited from the template workbook's styles.xml.
S_TITLE, S_HEADING, S_BODY, S_CODE, S_NOTE = "5", "6", "7", "8", "9"
S_TABLE_HEADER = "1"

HT_HEADING = "20.100000000000001"
HT_LINE = 15.0


# =====================================================================
# The loader
# =====================================================================

LOADER_CODE = '''# Load the cgs_ai package from the "__init__" worksheet.
#
# Python in Excel has no filesystem and no network, so it cannot reach the
# OneDrive folder that holds cgs_ai. It can, however, read a worksheet --
# and `import` checks sys.modules before it ever looks at a disk. The last
# line of __init__.py calls _register(), which puts a module object into
# sys.modules under the name "cgs_ai". Run the source once and every cell
# to the right can do an ordinary import.
import sys
import pandas as pd

rows = xl(%P2%)
column = rows.iloc[:, 0] if getattr(rows, "ndim", 1) == 2 else rows

# dropna() is used for ONE thing: finding the last row that still has source
# on it. Everything past that is empty grid, not code, and trimming it keeps
# the range above from needing an edit every time the package grows.
#
# Interior blank lines are then kept as "". Dropping those runs perfectly
# well and silently flattens every docstring in the package -- paragraph
# breaks inside a """...""" are content, not whitespace.
filled = column.dropna()
if len(filled) == 0:
    raise ValueError("the __init__ sheet is empty - paste the package into "
                     "column A, one line of the file per row")
lastLine = column.index.get_loc(filled.index[-1])
source = "\\n".join("" if pd.isna(value) else str(value)
                    for value in column.iloc[:lastLine + 1])

exec(compile(source, "__init__ sheet", "exec"), {})

module = sys.modules["cgs_ai"]
f"cgs_ai {module.__version__} loaded - {len(source):,} characters, " \\
f"{lastLine + 1} rows, {len(module.__all__) - 1} functions"'''


# =====================================================================
# The three tests
# =====================================================================

TEST1_CODE = '''import cgs_ai

cgs_ai.doBasicHello()'''

TEST2_TEXT = '''import cgs_ai

claims = xl("Claims_Data!$A$1:$N$51", headers=True)
cgs_ai.surveySelect(data=claims, n=5)'''

TEST2_CODE = '''import cgs_ai

claims = xl(%P2%, headers=True)
cgs_ai.surveySelect(data=claims, n=5)'''

TEST3_CODE = '''import cgs_ai

RowCount    = 50000     # rows at 100%
ColumnCount = 12        # columns in the synthetic dataset
Increment   = 25        # 25% -> four steps: 25, 50, 75, 100
cgs_ai.testWorkload(RowCount, ColumnCount, Increment, Repeats=3, Label="small")'''


# =====================================================================
# Sheet text
# =====================================================================

README_TITLE = "TESTS1 — loading cgs_ai from a worksheet"

README_INTRO = (
    "This workbook tests one idea: that a custom Python package can be "
    "imported inside Excel without a filesystem, a network or a pip "
    "install. If the three TEST tabs return values, it works, and the "
    "pattern can go into the real workbook.")

README_SECTIONS: List[Tuple[str, List[str]]] = [
    ("Why the package lives on a worksheet", [
        "Python in Excel does not run on your computer. Microsoft executes "
        "it in a managed container in Azure with no access to your disk, "
        "OneDrive, SharePoint or the network, and no pip. So "
        "sys.path.insert(0, r\"...OneDrive...\") can never find anything, "
        "and import cgs_ai fails there no matter how correct the path is.",
        "What the container CAN read is a worksheet. And Python's import "
        "checks sys.modules — the interpreter's cache of already-imported "
        "modules — BEFORE it goes looking on disk. The last line of the lite "
        "__init__.py calls _register(), which builds a module object and "
        "files it in that cache under the name \"cgs_ai\". Run the source "
        "once and every later cell gets a normal import for free.",
    ]),
    ("Why one line per row, and not one big cell", [
        "The lite __init__.py is 37,732 characters. Excel's limit for the "
        "text in a single cell is 32,767, so it does not fit — and a Python "
        "file cannot be split across cells at arbitrary points.",
        "One line of the file per row removes the ceiling entirely. The "
        "longest line in the package is 84 characters, so no function can "
        "outgrow a cell however much the package gains. It also means you "
        "can read, search and diff the code in the grid, and see at a glance "
        "which version is loaded.",
    ]),
    ("Tab order is load-bearing — do not reorder these", [
        "Excel runs =PY() cells in ONE Python session, in workbook order: "
        "row-major within a sheet, then sheet by sheet LEFT TO RIGHT across "
        "the tabs. The dependency here is a shared namespace, not a cell "
        "reference, so Excel cannot see it and will not reorder anything for "
        "you.",
        "That is why __init__ is the second tab and TEST1–TEST3 follow it. "
        "Drag __init__ to the right of a TEST sheet and that test breaks "
        "with ModuleNotFoundError. Add new sheets that call cgs_ai to the "
        "RIGHT of __init__, never the left.",
    ]),
    ("How to hide the __init__ sheet", [
        "Right-click the __init__ tab and choose Hide. It still calculates "
        "while hidden, so the loader keeps working and the TEST sheets are "
        "unaffected. Right-click any tab and choose Unhide to bring it back.",
        "To hide it so it does not even appear in the Unhide list, open the "
        "VBA editor with Alt+F11, select the __init__ sheet in the Project "
        "pane, and set its Visible property to 2 - xlSheetVeryHidden. Only "
        "VBA can undo that, which is the point: it stops somebody deleting "
        "the package by accident. Leave it visible while you are testing.",
    ]),
    ("What to do if a test shows a stale value", [
        "The cached values saved in this file are placeholders. Press "
        "Ctrl+Alt+F9 to force a full recalculation, or edit any =PY() cell "
        "and press Ctrl+Enter. Remember that Enter adds a newline inside a "
        "Python cell — Ctrl+Enter is what commits it.",
        "If a result shows as a Python object rather than spilling across "
        "the grid, select the cell, click the PY icon at the left of the "
        "formula bar and choose Excel Value (or press Ctrl+Alt+Shift+M).",
    ]),
    ("The honest limitation", [
        "This is a COPY of the package, not a link to it. Update cgs_ai on "
        "SharePoint and this workbook keeps running the pasted version until "
        "somebody re-pastes it. Outside Excel — VS Code, Anaconda, Jupyter, "
        "a Snowflake notebook — the import is live and always current. The "
        "call syntax is identical in both places, which is the whole point "
        "of the design.",
    ]),
]

INIT_TITLE = "__init__ — the cgs_ai package source"

INIT_INTRO = (
    "Column A holds src/py/lite/cgs_ai/__init__.py, one line of the file per "
    "row. Do not sort it, do not delete blank rows, and do not let anything "
    "else into column A — this is executable code and the row order IS the "
    "file. The loader in cell C3 runs it.")

INIT_WARNING = (
    "If you ever retype a line by hand: a cell that starts with = or + is a "
    "formula to Excel, and several lines of this file start with = (the "
    "comment banners). Paste, or prefix with an apostrophe, rather than "
    "typing. The cells in this file were written as text and are safe as "
    "they stand.")

CLAIMS_TITLE = "Claims_Data — 50 synthetic claims for TEST2"

#: Every variable in the Metadata data dictionary. Deliberately NOT an Excel
#: table: a table's column names must be unique ignoring case, and this list
#: carries both Allowed_Amount and allowed_amount.
BIG_COLUMNS = [
    "Claim_ID", "Bene_ID", "Claim_Type", "From_Date", "Thru_Date",
    "ICD10_Diagnosis_Code", "HCPCS_Code", "Allowed_Amount",
    "Medicare_Payment_Amount", "allowed_amount", "medicare_payment",
    "patient_responsibility", "length_of_stay", "claim_type",
]
NUMERIC_COLUMNS = frozenset({"Allowed_Amount", "Medicare_Payment_Amount",
                             "allowed_amount", "medicare_payment",
                             "patient_responsibility", "length_of_stay"})

TESTS: List[Dict[str, str]] = [
    {
        "name": "TEST1",
        "title": "TEST1 — import cgs_ai and call a function",
        "intro": "The smoke test. If this returns a greeting, the package "
                 "loaded from the __init__ sheet and the import resolved "
                 "from sys.modules. Everything else depends on this one "
                 "working.",
        "text": TEST1_CODE,
        "code": TEST1_CODE,
        "expect": "Expected result:  Hello, World!",
        "argument": "",
        "cached": "Hello, World!",
    },
    {
        "name": "TEST2",
        "title": "TEST2 — cgs_ai.surveySelect on real claims",
        "intro": "A reproducible 5-row sample of the Claims_Data tab. The "
                 "seed is a published constant inside the function, so this "
                 "returns the SAME five claims every time it recalculates — "
                 "which is what makes a review sample defensible. Pass a "
                 "different seed= to get a different five.",
        "text": TEST2_TEXT,
        "code": TEST2_CODE,
        "expect": "Expected result:  a 5-row table, numbered 1 to 5, with "
                  "all 14 columns of Claims_Data.",
        "argument": CLAIMS_RANGE,
        "cached": "",
    },
    {
        "name": "TEST3",
        "title": "TEST3 — cgs_ai.testWorkload, unchanged",
        "intro": "The call from the Test_Scaling tab, exactly as written "
                 "there, with no copy of the module in the cell. Read the "
                 "ScalingIndex column: 1.0 means ten times the data costs "
                 "ten times the time; above 1.0 means each row got more "
                 "expensive as the data grew. Note these timings describe "
                 "Microsoft's container, not a CGS server.",
        "text": TEST3_CODE,
        "code": TEST3_CODE,
        "expect": "Expected result:  four rows (25%, 50%, 75%, 100%) with "
                  "timings and a ScalingIndex column. Give it a few seconds.",
        "argument": "",
        "cached": "",
    },
]


# =====================================================================
# XML helpers
# =====================================================================

def escapeXml(text: str) -> str:
    """Escape text for XML content. Parameters: text (str). Returns: str."""
    return (str(text).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;"))


def columnLetter(index: int) -> str:
    """1 -> A, 27 -> AA. Parameters: index (int), 1-based. Returns: str."""
    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def wrappedHeight(text: str, charsPerLine: int, minimum: float = 33.95) -> str:
    """Estimate the row height a wrapped cell needs, in points.

    Parameters: text (str); charsPerLine (int); minimum (float).
    Returns: str, for ht="...".

    Excel does not auto-fit a wrapped cell on open, so a row written without
    a height shows one line and hides the rest.
    """
    lines = sum(max(1, math.ceil(len(part) / charsPerLine))
                for part in str(text).split("\n"))
    return str(max(minimum, round(lines * HT_LINE + 4, 2)))


def inlineCell(reference: str, style: str, value: str) -> str:
    """One inline-string cell. Parameters: reference, style, value.

    Inline strings rather than shared strings: this workbook is generated
    whole, nothing is shared between sheets, and an inline string keeps the
    text next to the cell that holds it where it can be read.
    """
    return (f'<c r="{reference}" s="{style}" t="inlineStr"><is>'
            f'<t xml:space="preserve">{escapeXml(value)}</t></is></c>')


def numberCell(reference: str, value: object) -> str:
    """One numeric cell. Parameters: reference (str); value."""
    return f'<c r="{reference}"><v>{value}</v></c>'


def pythonCell(reference: str, scriptIndex: int, style: str = "",
               cached: str = "", argument: str = "") -> str:
    """Build a Python-in-Excel cell.

    Parameters:
        reference (str)   - the anchor cell.
        scriptIndex (int) - ZERO-BASED index into python.xml. PY(14,0) in a
                            real workbook runs the fifteenth script; this was
                            read off a working file, not guessed.
        style (str)       - style id, or "" for the sheet default.
        cached (str)      - the value shown until Excel recalculates.
        argument (str)    - an Excel reference the script reads as %P2%.
    Returns: str, the <c> XML.

    cm="1" points at cellMetadata entry 1, which marks the cell a dynamic
    array. Without it Excel treats _xlfn._xlws.PY as an unknown function and
    shows #NAME?.
    """
    attributes = f'<c r="{reference}"' + (f' s="{style}"' if style else "")
    extra = f",{argument}" if argument else ""
    value = f"<v>{escapeXml(cached)}</v>" if cached else "<v/>"
    return (f'{attributes} t="str" cm="1"><f t="array" ref="{reference}">'
            f"_xlfn._xlws.PY({scriptIndex},0{extra})</f>{value}</c>")


def row(number: int, cells: str, height: str = "", spans: str = "1:3") -> str:
    """Wrap cell XML in a <row>. Parameters: number, cells, height, spans."""
    attributes = f'<row r="{number}" spans="{spans}"'
    if height:
        attributes += f' ht="{height}" customHeight="1"'
    return f'{attributes} x14ac:dyDescent="0.25">{cells}</row>'


def sheetXml(dimension: str, cols: str, rows: str,
             freeze: str = "") -> str:
    """Assemble a worksheet part.

    Parameters:
        dimension (str) - e.g. "A1:C40".
        cols (str)      - the <cols> block, or "".
        rows (str)      - the <row> elements.
        freeze (str)    - a topLeftCell to freeze above, or "".
    Returns: str, the whole worksheet XML.

    Child order is not advisory: dimension, sheetViews, sheetFormatPr, cols,
    sheetData, then pageMargins. Out of order, Excel repairs the sheet by
    discarding it.
    """
    pane = ""
    if freeze:
        pane = (f'<pane ySplit="{int(freeze[1:]) - 1}" topLeftCell="{freeze}" '
                'activePane="bottomLeft" state="frozen"/>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<worksheet xmlns="{NS_MAIN}" xmlns:r="{NS_REL}" '
            'xmlns:mc="http://schemas.openxmlformats.org/'
            'markup-compatibility/2006" mc:Ignorable="x14ac" '
            'xmlns:x14ac="http://schemas.microsoft.com/office/'
            'spreadsheetml/2009/9/ac">'
            f'<dimension ref="{dimension}"/>'
            f'<sheetViews><sheetView workbookViewId="0">{pane}'
            "</sheetView></sheetViews>"
            '<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>'
            f"{cols}<sheetData>{rows}</sheetData>"
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" '
            'header="0.3" footer="0.3"/></worksheet>')


# =====================================================================
# The sheets
# =====================================================================

def buildReadmeSheet() -> str:
    """The README tab. Parameters: none. Returns: the worksheet XML."""
    body = [row(1, inlineCell("A1", S_TITLE, README_TITLE), "24", "1:1"),
            row(2, inlineCell("A2", S_NOTE, README_INTRO),
                wrappedHeight(README_INTRO, 110), "1:1")]
    number = 4
    for heading, paragraphs in README_SECTIONS:
        body.append(row(number, inlineCell(f"A{number}", S_HEADING, heading),
                        HT_HEADING, "1:1"))
        number += 1
        for paragraph in paragraphs:
            body.append(row(number, inlineCell(f"A{number}", S_BODY, paragraph),
                            wrappedHeight(paragraph, 110), "1:1"))
            number += 1
        number += 1
    return sheetXml(f"A1:A{number}",
                    '<cols><col min="1" max="1" width="110" '
                    'customWidth="1"/></cols>', "".join(body))


def buildInitSheet(sourceLines: Sequence[str], loaderIndex: int) -> str:
    """The __init__ tab: the package source, plus the loader cell.

    Parameters:
        sourceLines (sequence[str]) - the file, one entry per line.
        loaderIndex (int)           - 0-based index of the loader script.
    Returns: the worksheet XML.

    The source starts at row 1 so the row number in the grid IS the line
    number in the file -- row 412 is line 412. That makes a traceback from
    the loaded module point straight at a row.
    """
    # The column C annotations wrap; Excel will not auto-fit them on open,
    # so those few rows carry a stated height.
    heights = {1: "24", 2: wrappedHeight(INIT_INTRO, 60),
               3: "20.100000000000001", 5: wrappedHeight(INIT_WARNING, 60)}
    body = []
    for number, line in enumerate(sourceLines, start=1):
        cells = ""
        if line:
            cells = inlineCell(f"A{number}", S_CODE, line)
        if number == 1:
            cells += inlineCell("C1", S_TITLE, INIT_TITLE)
        elif number == 2:
            cells += inlineCell("C2", S_BODY, INIT_INTRO)
        elif number == 3:
            cells += pythonCell("C3", loaderIndex, style=S_CODE,
                                argument=SOURCE_RANGE)
        elif number == 5:
            cells += inlineCell("C5", S_NOTE, INIT_WARNING)
        if cells:
            body.append(row(number, cells, heights.get(number, ""),
                            spans="1:3"))
    cols = ('<cols><col min="1" max="1" width="92" customWidth="1"/>'
            '<col min="2" max="2" width="2.42578125" customWidth="1"/>'
            '<col min="3" max="3" width="62" customWidth="1"/></cols>')
    return sheetXml(f"A1:C{len(sourceLines)}", cols, "".join(body))


def buildClaimsRows(count: int = 50,
                    seed: int = 8111305759926929283) -> List[Dict[str, object]]:
    """Build the synthetic claims TEST2 samples.

    Parameters: count (int); seed (int).
    Returns: list of dicts keyed by BIG_COLUMNS.

    Seeded, so rebuilding this workbook gives the same 50 claims and TEST2's
    expected output does not move underneath it.
    """
    generator = random.Random(seed)
    types = ["Inpatient", "Outpatient", "Carrier/Physician", "DME"]
    analysis = {"Inpatient": "Inpatient", "Outpatient": "Outpatient",
                "Carrier/Physician": "Carrier", "DME": "DME"}
    diagnoses = ["I50.9", "M17.11", "E11.9", "Z99.81", "J18.9", "N18.3",
                 "I10", "E78.5", "M54.5", "J44.9"]
    procedures = ["99213", "99204", "E1390", "93000", "71046", "80053",
                  "97110", "99285", "36415"]
    rows: List[Dict[str, object]] = []
    for index in range(count):
        claimType = generator.choice(types)
        startDay = generator.randint(0, 300)
        stay = (generator.randint(2, 14) if claimType == "Inpatient"
                else generator.randint(0, 1))
        allowed = round(generator.uniform(85, 24000), 2)
        paid = round(allowed * generator.uniform(0.65, 0.92), 2)
        rows.append({
            "Claim_ID": f"CLM{100300 + index}",
            "Bene_ID": f"BENE{generator.randint(1000, 9999)}",
            "Claim_Type": claimType,
            "From_Date": f"2026-{startDay // 28 + 1:02d}-{startDay % 28 + 1:02d}",
            "Thru_Date": f"2026-{(startDay + stay) // 28 + 1:02d}-"
                         f"{(startDay + stay) % 28 + 1:02d}",
            "ICD10_Diagnosis_Code": generator.choice(diagnoses),
            "HCPCS_Code": ("" if claimType == "Inpatient" and stay > 7
                           else generator.choice(procedures)),
            "Allowed_Amount": allowed,
            "Medicare_Payment_Amount": paid,
            "allowed_amount": allowed,
            "medicare_payment": paid,
            "patient_responsibility": round(allowed - paid, 2),
            "length_of_stay": stay,
            "claim_type": analysis[claimType],
        })
    return rows


def buildClaimsSheet(claims: Sequence[Dict[str, object]]) -> str:
    """The Claims_Data tab. Parameters: claims. Returns: the worksheet XML."""
    span = f"1:{len(BIG_COLUMNS)}"
    header = "".join(inlineCell(f"{columnLetter(i)}1", S_TABLE_HEADER, name)
                     for i, name in enumerate(BIG_COLUMNS, start=1))
    body = [row(1, header, spans=span)]
    for number, claim in enumerate(claims, start=2):
        cells = []
        for i, name in enumerate(BIG_COLUMNS, start=1):
            reference = f"{columnLetter(i)}{number}"
            value = claim[name]
            if name in NUMERIC_COLUMNS:
                cells.append(numberCell(reference, value))
            elif value != "":
                cells.append(f'<c r="{reference}" t="inlineStr"><is>'
                             f'<t xml:space="preserve">{escapeXml(value)}'
                             "</t></is></c>")
        body.append(row(number, "".join(cells), spans=span))
    cols = "<cols>" + "".join(
        f'<col min="{i}" max="{i}" width="{max(12, min(26, len(n) + 4))}" '
        f'customWidth="1"/>' for i, n in enumerate(BIG_COLUMNS, start=1)
    ) + "</cols>"
    return sheetXml(f"A1:{columnLetter(len(BIG_COLUMNS))}{len(claims) + 1}",
                    cols, "".join(body), freeze="A2")


def buildTestSheet(test: Dict[str, str], scriptIndex: int) -> str:
    """One TEST tab. Parameters: test (dict); scriptIndex (int)."""
    body = [
        row(1, inlineCell("A1", S_TITLE, test["title"]), "24", "1:1"),
        row(2, inlineCell("A2", S_NOTE, test["intro"]),
            wrappedHeight(test["intro"], 100), "1:1"),
        row(4, inlineCell("A4", S_HEADING,
                          "The code — exactly what you would write in VS Code"),
            HT_HEADING, "1:1"),
        row(5, inlineCell("A5", S_CODE, test["text"]),
            wrappedHeight(test["text"], 90), "1:1"),
        row(7, inlineCell("A7", S_BODY, test["expect"]),
            wrappedHeight(test["expect"], 100), "1:1"),
        row(9, inlineCell("A9", S_HEADING, "Live result — spills from C10"),
            HT_HEADING, "1:1"),
        # The result spills down and right from C10. Anchoring it in column A
        # would give the table a 100-character-wide first column.
        row(10, inlineCell("A10", S_BODY,
                           "Select C10 to see the =PY() formula.")
            + pythonCell("C10", scriptIndex, cached=test["cached"],
                         argument=test["argument"]), spans="1:18"),
    ]
    cols = ('<cols><col min="1" max="1" width="100" customWidth="1"/>'
            '<col min="2" max="2" width="2.42578125" customWidth="1"/>'
            '<col min="3" max="18" width="17" customWidth="1"/></cols>')
    return sheetXml("A1:R30", cols, "".join(body))


# =====================================================================
# Package parts
# =====================================================================

def buildPythonXml(environment: str, scripts: Sequence[str]) -> str:
    """Build xl/python.xml. Parameters: environment (str); scripts."""
    body = "".join(f"<pythonScript><code>{escapeXml(code)}</code>"
                   f"</pythonScript>" for code in scripts)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<python xmlns="{NS_PY}">{environment}'
            f"<pythonScripts>{body}</pythonScripts></python>")


def buildMetadataXml() -> str:
    """Build xl/metadata.xml.

    Parameters: none. Returns: str.

    The minimum a =PY() cell needs: one XLDAPR metadata type, the dynamic
    array property, and one cellMetadata entry -- which is what cm="1" on a
    cell points at (1-based). The rich-value types are left out on purpose:
    they index into richData parts, and a valueMetadata entry pointing at a
    part that is not in the package is a repair prompt.
    """
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<metadata xmlns="{NS_MAIN}" xmlns:xda="http://schemas.microsoft'
            '.com/office/spreadsheetml/2017/dynamicarray">'
            '<metadataTypes count="1"><metadataType name="XLDAPR" '
            'minSupportedVersion="120000" copy="1" pasteAll="1" '
            'pasteValues="1" merge="1" splitFirst="1" rowColShift="1" '
            'clearFormats="1" clearComments="1" assign="1" coerce="1" '
            'cellMeta="1"/></metadataTypes>'
            '<futureMetadata name="XLDAPR" count="1"><bk><extLst>'
            '<ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">'
            '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/>'
            "</ext></extLst></bk></futureMetadata>"
            '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk>'
            "</cellMetadata></metadata>")


def buildWorkbookXml(sheets: Sequence[Tuple[str, str]]) -> str:
    """Build xl/workbook.xml.

    Parameters: sheets - (name, relationship id) in TAB ORDER.
    Returns: str.

    Tab order is the calculation order for =PY() cells, so this list is the
    contract the whole workbook rests on. fullCalcOnLoad because every cached
    value in the file is a placeholder.
    """
    entries = "".join(
        f'<sheet name="{escapeXml(name)}" sheetId="{index}" r:id="{rid}"/>'
        for index, (name, rid) in enumerate(sheets, start=1))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<workbook xmlns="{NS_MAIN}" xmlns:r="{NS_REL}">'
            '<workbookPr defaultThemeVersion="124226"/>'
            "<bookViews><workbookView xWindow=\"0\" yWindow=\"0\" "
            'windowWidth="28800" windowHeight="15600"/></bookViews>'
            f"<sheets>{entries}</sheets>"
            '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')


# =====================================================================
# Main
# =====================================================================

def build(template: Path, target: Path) -> Dict[str, object]:
    """Produce the test workbook.

    Parameters: template (Path) - a workbook to copy scaffolding from;
                target (Path).
    Returns: dict - a summary.
    Raises: ValueError - the template is missing a part this needs.
    """
    with zipfile.ZipFile(template) as archive:
        names = set(archive.namelist())
        for required in ("xl/styles.xml", "xl/theme/theme1.xml",
                         "xl/python.xml"):
            if required not in names:
                raise ValueError(f"{template.name} has no {required}")
        styles = archive.read("xl/styles.xml")
        theme = archive.read("xl/theme/theme1.xml")
        templatePython = archive.read("xl/python.xml").decode("utf-8")

    start = templatePython.index("<environmentDefinition")
    end = templatePython.index("</environmentDefinition>") + len(
        "</environmentDefinition>")
    environment = templatePython[start:end]

    sourcePath = (Path(__file__).resolve().parent.parent / "py" / "lite" /
                  "cgs_ai" / "__init__.py")
    source = sourcePath.read_text(encoding="utf-8")
    sourceLines = source.splitlines()
    if len(sourceLines) > SOURCE_ROWS:
        raise ValueError(f"the package is {len(sourceLines)} lines, past the "
                         f"{SOURCE_ROWS}-row range the loader reads")

    scripts = [LOADER_CODE] + [test["code"] for test in TESTS]
    claims = buildClaimsRows()

    # Tab order IS calculation order. __init__ must precede every TEST.
    sheetOrder = [("README", "rId1"), ("__init__", "rId2"),
                  ("Claims_Data", "rId3"), ("TEST1", "rId4"),
                  ("TEST2", "rId5"), ("TEST3", "rId6")]

    parts: Dict[str, bytes] = {
        "xl/styles.xml": styles,
        "xl/theme/theme1.xml": theme,
        "xl/python.xml": buildPythonXml(environment, scripts).encode("utf-8"),
        "xl/metadata.xml": buildMetadataXml().encode("utf-8"),
        "xl/workbook.xml": buildWorkbookXml(sheetOrder).encode("utf-8"),
        "xl/worksheets/sheet1.xml": buildReadmeSheet().encode("utf-8"),
        "xl/worksheets/sheet2.xml": buildInitSheet(sourceLines, 0).encode("utf-8"),
        "xl/worksheets/sheet3.xml": buildClaimsSheet(claims).encode("utf-8"),
    }
    for index, test in enumerate(TESTS):
        parts[f"xl/worksheets/sheet{index + 4}.xml"] = buildTestSheet(
            test, index + 1).encode("utf-8")

    parts["_rels/.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<Relationships xmlns="{NS_PKG}">'
        f'<Relationship Id="rId1" Type="{NS_REL}/officeDocument" '
        'Target="xl/workbook.xml"/></Relationships>').encode("utf-8")

    workbookRels = [
        f'<Relationship Id="{rid}" Type="{NS_REL}/worksheet" '
        f'Target="worksheets/sheet{index + 1}.xml"/>'
        for index, (_, rid) in enumerate(sheetOrder)]
    workbookRels += [
        f'<Relationship Id="rId10" Type="{NS_REL}/styles" '
        'Target="styles.xml"/>',
        f'<Relationship Id="rId11" Type="{NS_REL}/theme" '
        'Target="theme/theme1.xml"/>',
        f'<Relationship Id="rId12" Type="{NS_REL}/sheetMetadata" '
        'Target="metadata.xml"/>',
        '<Relationship Id="rId13" Type="http://schemas.microsoft.com/office/'
        '2023/09/relationships/Python" Target="python.xml"/>']
    parts["xl/_rels/workbook.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<Relationships xmlns="{NS_PKG}">{"".join(workbookRels)}'
        "</Relationships>").encode("utf-8")

    overrides = "".join(
        f'<Override PartName="/xl/worksheets/sheet{index + 1}.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.'
        'spreadsheetml.worksheet+xml"/>' for index in range(len(sheetOrder)))
    parts["[Content_Types].xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/'
        'content-types">'
        '<Default Extension="rels" ContentType="application/vnd.'
        'openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.'
        'openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        f"{overrides}"
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.'
        'openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '<Override PartName="/xl/theme/theme1.xml" ContentType="application/'
        'vnd.openxmlformats-officedocument.theme+xml"/>'
        '<Override PartName="/xl/metadata.xml" ContentType="application/vnd.'
        'openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>'
        '<Override PartName="/xl/python.xml" ContentType="application/vnd.'
        'ms-excel.python+xml"/>'
        "</Types>").encode("utf-8")

    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)

    return {"Target": str(target), "Sheets": [n for n, _ in sheetOrder],
            "SourceLines": len(sourceLines),
            "SourceCharacters": len(source),
            "LongestLine": max(len(line) for line in sourceLines),
            "ExcelCellLimit": 32767,
            "Scripts": len(scripts), "ClaimRows": len(claims)}


def main(argv: Sequence[str]) -> int:
    """CLI entry point. Parameters: argv. Returns: an exit code."""
    if len(argv) != 2:
        print("usage: build_tests_workbook.py <template.xlsx> <target.xlsx>",
              file=sys.stderr)
        return EXIT_USAGE
    for key, value in build(Path(argv[0]), Path(argv[1])).items():
        print(f"{key}: {value}")
    return EXIT_OK


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
