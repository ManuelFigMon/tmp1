"""
=====================================================================
  Program Name  : build_claims_pipelines_v15.py
  Author        : Manuel Figallo
  Purpose       : Turn v1.4 of python_for_medicare_claims_pipelines.xlsx
                  into v1.5, and permanently repair the table that Excel
                  reported as unreadable.
  Version       : 1.0beta
  Created       : 2026-10-29
  Last Modified : 2026-10-29

  Dependencies:
    STANDARD LIBRARY ONLY.

  Description:
    RAW XML SURGERY, NOT openpyxl, for the same reason as the v1.4 builder:
    a round-trip through openpyxl drops every image, xl/python.xml and with
    it every =PY() cell.

    THE REPAIR. Excel reported "Repaired Records: Table from
    /xl/tables/table6.xml part (Table)" and then quietly renamed two
    columns -- allowed_amount became allowed_amount2 and claim_type became
    claim_type3. The cause is not corruption: EXCEL REQUIRES TABLE COLUMN
    NAMES TO BE UNIQUE IGNORING CASE, and the Metadata data dictionary
    legitimately contains both Allowed_Amount and allowed_amount, and both
    Claim_Type and claim_type. They are two different variables under two
    naming conventions, and renaming either one to satisfy Excel would make
    the sheet disagree with the dictionary it is supposed to mirror.

    So the fix is to stop making Claims_Data_BIG a table. It becomes an
    ordinary range with a frozen header and an autofilter, which looks and
    behaves the same, has no uniqueness rule, and cannot be repaired again.
    The =PY() cells that read it move from Claims_Data_BIG[#All] to the
    range Claims_Data_BIG!$A$1:$N$51, and the two header cells Excel
    renamed are put back.

    THE WORKSHEET CHANGES.
      CGI_AI     - a new "How to import CGS_AI" section after Source, with
                   the import in column C, and every column C cell calling
                   a cgs_ai function rather than a copy pasted into the
                   cell.
      Test_Scaling - "Run it live" now imports cgs_ai in A12 and calls
                   cgs_ai.testWorkload() below it. The full module is gone
                   from this sheet; it lives in the package.
      README     - a new "The Analytics Lifecycle" section before Version
                   history, with a six-step cycle drawn as DrawingML
                   shapes: data -> describe data -> correlation matrix ->
                   scatter plot -> pair plot -> linear regression -> data.

    WHAT THE EXCEL SANDBOX WILL DO WITH THE IMPORT. Python in Excel runs
    in a Microsoft-managed container with no filesystem and no network, so
    "import cgs_ai" from a SharePoint path raises ModuleNotFoundError
    there. The import pattern is correct and is what runs in VS Code,
    Anaconda, Jupyter and a Snowflake notebook; a note row on each sheet
    says so, so the reader is not left guessing at an error message.

  Input Parameters (required first):
    source (REQUIRED) - the v1.4 workbook
    target (REQUIRED) - the .xlsx to write

  Usage:
    python build_claims_pipelines_v15.py <source.xlsx> <target.xlsx>

  Exit codes:
    0 = success, 2 = usage error.
=====================================================================
"""

from __future__ import annotations

import math
import re
import sys
import zipfile
from pathlib import Path
from typing import Dict, List, Sequence, Tuple

__version__ = "1.0beta"

EXIT_OK, EXIT_USAGE = 0, 2

# Style ids as they stand in the v1.4 workbook after Excel re-saved it.
S_TITLE, S_HEADING, S_BODY, S_CODE, S_NOTE = "5", "6", "7", "8", "9"
# Test_Scaling keeps its own palette.
T_TITLE, T_NOTE, T_LABEL, T_CODE, T_HEADING, T_BODY = "19", "15", "13", "16", "17", "18"
T_FILL = "14"            # the blank cells that sit under a merged A:J row

HT_HEADING = "20.100000000000001"
HT_LABEL = "18.95"
HT_LINE = 15.0

#: Width of the merged A:J band on Test_Scaling, in characters.
SCALE_BAND = 6 + 8 + 9 + 11 + 10 + 12 + 14 + 11 + 13 + 18
README_BAND = 100        # column A is 96 characters wide
CGI_BAND = 94            # column A on CGI_AI

#: Where Claims_Data_BIG lives, as a plain range. Absolute, so a reader who
#: copies the cell somewhere else still points at the data.
BIG_RANGE = "Claims_Data_BIG!$A$1:$N$51"

NS_SS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


# =====================================================================
# The import snippet, verbatim
# =====================================================================

IMPORT_SNIPPET = '''# The folder that CONTAINS the cgs_ai package folder.
SHARE = r"https://bcbsscgov-my.sharepoint.us/my?id=%2Fpersonal%2Fmanuel%5Ffigallo%5Fcgsadmin%5Fcom%2FDocuments%2Fcode%2Fpython%2Fcgs%5Fai&ga=1"
CGS_AI_HOME = SHARE + r"\\src\\py\\lite"

# STEP 1. Import the python package
# Telling Python where to look, then importing cgs_ai exactly like pandas.
import sys
sys.path.insert(0, CGS_AI_HOME)
import cgs_ai'''

SANDBOX_NOTE = (
    "One caveat, so an error message does not come as a surprise: Python in "
    "Excel runs in a Microsoft-managed container with no filesystem and no "
    "network, so it cannot reach SharePoint and the import above will report "
    "that cgs_ai cannot be found. The pattern is correct and is exactly what "
    "runs in VS Code, Anaconda, Jupyter or a Snowflake notebook — where the "
    "cells below return real results. Use Excel to read the code; use one of "
    "those to run it against the live package.")


# =====================================================================
# CGI_AI content
# =====================================================================

CGI_IMPORT_HEADING = "How to import CGS_AI"

CGI_IMPORT_INTRO = (
    "cgs_ai is our own Python package — code the team wrote, reviewed and "
    "shares, sitting in one folder on OneDrive/SharePoint rather than in "
    "somebody's private notebook. You do not install it and you do not copy "
    "it: you tell Python where the folder is, and then you import it exactly "
    "the way you import pandas. Two lines, and every function in the package "
    "is available as cgs_ai.something(). That is the whole point of a "
    "package — write a function once, review it once, and let everyone call "
    "it from anywhere.")

CGI_IMPORT_LABEL = ("The import — it is in column C, and every live cell "
                    "below depends on it having run first:")

CGI_LIVE_HEADING = "Live preview — calling cgs_ai functions"

CGI_LIVE_INTRO = (
    "Each cell in column C now CALLS a cgs_ai function instead of carrying "
    "its own copy of the code. That is the difference worth noticing: the "
    "function lives in one file, __init__.py, and the sheet just calls it. "
    "Fix a bug there and every sheet, notebook and script that calls it is "
    "fixed with it.")

CGI_BASIC_TEXT = '''# Example 1 — the smoke test. If this returns, the import worked.
result = cgs_ai.doBasicHello()
print(result)        # -> Hello, World!'''

CGI_PERSONAL_TEXT = '''# Example 2 — a function that takes an argument.
result = cgs_ai.doPersonalizedHello('Alice')
print(result)        # -> Hello, Alice!'''

CGI_DETAILED_TEXT = '''# Example 3 — a function that returns a STRUCTURE, not a sentence.
result = cgs_ai.doDetailedHello('pirate')
print(result['message'])   # -> Ahoy, World! Avast ye scallywags!'''

CGI_SURVEY_HEADING = "Live preview — cgs_ai.surveySelect on Claims_Data_BIG"

CGI_SURVEY_INTRO = (
    "surveySelect draws a random sample you can reproduce. Pass it the rows "
    "and how many units you want, and it hands back that many, numbered from "
    "1. The seed is what makes it repeatable: the same seed returns the same "
    "rows on any machine, months later — which is what lets you defend a "
    "Medicare review sample after the fact. The two cells below are the same "
    "call with two different seeds, which is the fastest way to see what the "
    "seed actually does.")

CGI_SURVEY_TEXT = '''# A reproducible 10-row sample of Claims_Data_BIG.
claims = xl("Claims_Data_BIG!$A$1:$N$51", headers=True)
cgs_ai.surveySelect(data=claims, n=10)'''

CGI_SURVEY_SEED_TEXT = '''# The same call with a different seed — a different ten rows.
# Change the seed for a fresh sample; keep it to get last month's back.
claims = xl("Claims_Data_BIG!$A$1:$N$51", headers=True)
cgs_ai.surveySelect(data=claims, n=10, seed=20261028)'''

CGI_FUNCTIONS_LINE = (
    "Functions in the package:  doBasicHello()  ->  returns \"Hello, "
    "World!\" ;   doPersonalizedHello(name)  ->  returns \"Hello, <name>!\" ;   "
    "doDetailedHello(style)  ->  returns a dictionary, in three styles ;   "
    "formatData(InputPath, OutputExcelPath)  ->  a styled Excel report ;   "
    "sendEmail(To, From, Subject, Body)  ->  an SMTP notification ;   "
    "surveySelect(data, n, out)  ->  a reproducible random sample ;   "
    "testWorkload(RowCount, ColumnCount, Increment)  ->  a scaling benchmark "
    "(see the Test_Scaling tab).")

CGI_OUTSIDE_HEADING = "Run it outside Excel — where the import really works"

CGI_OUTSIDE_INTRO = (
    "Paste this into VS Code, Anaconda, Jupyter or a Snowflake notebook. Edit "
    "SHARE so it points at the folder that CONTAINS the cgs_ai folder, and "
    "run. This is the same import as column C, followed by the same calls.")

CGI_OUTSIDE_CODE = IMPORT_SNIPPET + '''

# STEP 2. Run examples
result = cgs_ai.doBasicHello()
print(result)

result = cgs_ai.doPersonalizedHello('Alice')
print(result)

result = cgs_ai.doDetailedHello('formal')
print(result['message'])

# A reproducible sample of any DataFrame.
import pandas as pd
claims = pd.read_csv('claims.csv')
print(cgs_ai.surveySelect(data=claims, n=30))

# Does ten times the data cost ten times the time? Read ScalingIndex.
print(cgs_ai.testWorkload(RowCount=200000, ColumnCount=12, Increment=25))'''

CGI_CLOSING = (
    "Claims_Data_BIG is synthetic and contains no PHI or PII, and "
    "surveySelect writes no files unless you give it an output path.")


# =====================================================================
# Test_Scaling content
# =====================================================================

SCALE_RUNLIVE_INTRO = (
    "testWorkload is a function in cgs_ai — our own Python package, kept in "
    "one shared folder on OneDrive/SharePoint so the team calls the same code "
    "instead of each keeping a copy. You do not install it and you do not "
    "paste it into the sheet: you point Python at the folder and import it "
    "like pandas, then call cgs_ai.testWorkload(). The full function is "
    "defined and documented once, in __init__.py.")

SCALE_STEP1 = "STEP 1 — import cgs_ai. The code is in cell A12:"

SCALE_STEP2 = ("STEP 2 — call cgs_ai.testWorkload() and read the result. The "
               "code is in cell A16; edit the parameters there and it "
               "recomputes:")

SCALE_CALL = '''RowCount    = 50000     # rows at 100%
ColumnCount = 12        # columns in the synthetic dataset
Increment   = 25        # 25% -> four steps: 25, 50, 75, 100
cgs_ai.testWorkload(RowCount, ColumnCount, Increment, Repeats=3, Label="small")'''

SCALE_IMPORTANT = (
    "IMPORTANT: Python in Excel runs in a Microsoft-managed Azure container, "
    "not on your laptop and not on a CGS server — so any timing produced here "
    "describes Microsoft's hardware. To size a CGS environment, run the same "
    "call there and compare. The container also has no network, so the import "
    "in A12 will not find cgs_ai inside Excel; run it from VS Code, Anaconda, "
    "Jupyter or a Snowflake notebook to see live numbers.")

SCALE_CLOSING = (
    "New in v1.5: this tab no longer carries a copy of the module. "
    "testWorkload is defined and documented once in the cgs_ai package, and "
    "this sheet imports it and calls it — which is the same thing the CGI_AI "
    "tab does, and the habit worth taking away from both. The synthetic data "
    "contains no PHI/PII and the function reads and writes no files.")


# =====================================================================
# README content
# =====================================================================

LIFECYCLE_HEADING = "The Analytics Lifecycle"

LIFECYCLE_INTRO = (
    "Every tab in this workbook is one step of the same loop, and the loop is "
    "what a data pipeline actually is. You start with the data, describe it "
    "to see its shape, find which variables move together, look at those "
    "relationships one pair at a time and then all at once, and fit a model "
    "to the strongest one. Then the arrow comes back to the data — because "
    "the model tells you what you were missing, and the next pass is better "
    "than the last. Nobody gets it right on the first lap; the value is in "
    "going round again.")

LIFECYCLE_STEPS = [
    "Data", "Describe Data", "Correlation Matrix",
    "Scatter Plot", "Pair Plot", "Linear Regression",
]

VERSION_LINE_15 = (
    "v1.5 (2026-10-29) — Added this Analytics Lifecycle section and its "
    "diagram. CGI_AI gained a How to import CGS_AI section, and every live "
    "cell in column C now calls a cgs_ai function instead of carrying a copy "
    "of it. Test_Scaling imports cgs_ai and calls cgs_ai.testWorkload() "
    "rather than embedding the module. Claims_Data_BIG is now a plain range "
    "rather than a table, which is the permanent fix for the repair prompt "
    "Excel raised: a table may not hold two column names that differ only by "
    "case, and the data dictionary has both Allowed_Amount and "
    "allowed_amount.")


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
    Returns: str, suitable for ht="...".

    Excel does not auto-fit a wrapped or merged cell when a file is opened,
    so a row written without a height shows one line and hides the rest.
    """
    lines = sum(max(1, math.ceil(len(part) / charsPerLine))
                for part in str(text).split("\n"))
    return str(max(minimum, round(lines * HT_LINE + 4, 2)))


def sharedStringList(sharedXml: str) -> List[str]:
    """Read sharedStrings.xml into a list. Parameters: sharedXml (str)."""
    import xml.etree.ElementTree as ET
    return ["".join(t.text or "" for t in si.iter(NS_SS + "t"))
            for si in ET.fromstring(sharedXml)]


def addSharedStrings(sharedXml: str,
                     values: Sequence[str]) -> Tuple[str, List[int]]:
    """Append strings to sharedStrings.xml.

    Parameters: sharedXml (str); values (sequence[str]).
    Returns: (new xml, list of 0-based indices).

    Appending never renumbers what is already there, so no existing cell has
    to be touched.
    """
    count = int(re.search(r'\bcount="(\d+)"', sharedXml).group(1))
    unique = int(re.search(r'uniqueCount="(\d+)"', sharedXml).group(1))
    indices = list(range(unique, unique + len(values)))
    addition = "".join(f'<si><t xml:space="preserve">{escapeXml(v)}</t></si>'
                       for v in values)
    sharedXml = sharedXml.replace("</sst>", addition + "</sst>", 1)
    sharedXml = sharedXml.replace(
        f'count="{count}" uniqueCount="{unique}"',
        f'count="{count + len(values)}" uniqueCount="{unique + len(values)}"', 1)
    return sharedXml, indices


def addPythonScripts(pythonXml: str,
                     codes: Sequence[str]) -> Tuple[str, List[int]]:
    """Append code blocks to xl/python.xml.

    Parameters: pythonXml (str); codes (sequence[str]).
    Returns: (new xml, list of ZERO-BASED indices for =PY(n,0)).
    """
    existing = len(re.findall(r"<pythonScript>", pythonXml))
    indices = list(range(existing, existing + len(codes)))
    addition = "".join(f"<pythonScript><code>{escapeXml(code)}</code>"
                       f"</pythonScript>" for code in codes)
    return (pythonXml.replace("</pythonScripts>", addition + "</pythonScripts>",
                              1), indices)


def shiftRows(sheetXml: str, fromRow: int, by: int) -> str:
    """Move every row at or below `fromRow` down by `by` rows.

    Parameters: sheetXml (str); fromRow (int) 1-based; by (int).
    Returns: str, the rewritten sheet.

    The <row r> attribute, every <c r> inside it and every array-formula ref
    move together; a cell still saying A7 inside a row numbered 10 is the
    kind of mismatch Excel repairs by throwing the sheet away.
    """
    def row(match: "re.Match") -> str:
        number = int(match.group(1))
        return (match.group(0) if number < fromRow else
                match.group(0).replace(f'r="{number}"', f'r="{number + by}"', 1))

    def cell(match: "re.Match") -> str:
        column, number = match.group(1), int(match.group(2))
        return (match.group(0) if number < fromRow
                else f'<c r="{column}{number + by}"')

    def ref(match: "re.Match") -> str:
        column, number = match.group(1), int(match.group(2))
        return (match.group(0) if number < fromRow
                else f'ref="{column}{number + by}"')

    sheetXml = re.sub(r'<row r="(\d+)"', row, sheetXml)
    sheetXml = re.sub(r'<c r="([A-Z]+)(\d+)"', cell, sheetXml)
    return re.sub(r'ref="([A-Z]+)(\d+)"', ref, sheetXml)


def insertAfterRow(sheetXml: str, afterRow: int, rowsXml: str) -> str:
    """Insert raw <row> XML immediately after a row number.

    Parameters: sheetXml (str); afterRow (int); rowsXml (str).
    Returns: str. Raises: ValueError if that row is not in the sheet.
    """
    match = re.search(r'<row r="%d"[^>]*>.*?</row>' % afterRow, sheetXml, re.S)
    if not match:
        raise ValueError(f"row {afterRow} is not in this sheet")
    return sheetXml[:match.end()] + rowsXml + sheetXml[match.end():]


def setDimension(sheetXml: str, ref: str) -> str:
    """Replace the sheet's declared dimension. Parameters: sheetXml, ref."""
    return re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="{ref}"/>',
                  sheetXml, count=1)


def replaceSheetData(sheetXml: str, rowsXml: str) -> str:
    """Swap a sheet's entire <sheetData> for new rows.

    Parameters: sheetXml (str); rowsXml (str) - the new <row> elements.
    Returns: str, the rewritten sheet.
    """
    return re.sub(r"<sheetData>.*?</sheetData>",
                  lambda _: f"<sheetData>{rowsXml}</sheetData>",
                  sheetXml, count=1, flags=re.S)


def pythonCell(reference: str, scriptIndex: int, style: str = "",
               cached: str = "", argument: str = "") -> str:
    """Build a Python-in-Excel cell.

    Parameters:
        reference (str)   - the anchor cell, e.g. "C15".
        scriptIndex (int) - ZERO-BASED index into python.xml.
        style (str)       - style id, or "" for the sheet default.
        cached (str)      - the value shown until Excel recalculates.
        argument (str)    - an Excel reference the script reads as %P2%.
    Returns: str, the <c> XML.

    cm="1" points at cellMetadata entry 1, which marks the cell a dynamic
    array; without it Excel shows #NAME? for _xlfn._xlws.PY.
    """
    attributes = f'<c r="{reference}"' + (f' s="{style}"' if style else "")
    extra = f",{argument}" if argument else ""
    value = f"<v>{escapeXml(cached)}</v>" if cached else "<v/>"
    return (f'{attributes} t="str" cm="1"><f t="array" ref="{reference}">'
            f"_xlfn._xlws.PY({scriptIndex},0{extra})</f>{value}</c>")


def stringCell(reference: str, style: str, index: int) -> str:
    """One shared-string cell. Parameters: reference, style, index."""
    return f'<c r="{reference}" s="{style}" t="s"><v>{index}</v></c>'


def simpleRow(number: int, cells: str, height: str, spans: str) -> str:
    """Wrap cell XML in a <row>. Parameters: number, cells, height, spans."""
    return (f'<row r="{number}" spans="{spans}" ht="{height}" '
            f'customHeight="1" x14ac:dyDescent="0.25">{cells}</row>')


# =====================================================================
# The Analytics Lifecycle diagram
# =====================================================================

#: Overall size of the diagram, in EMU (914400 EMU to the inch). Column A on
#: the README is 96 characters, roughly 6.9 inches, so 6.5 inches fits.
DIAGRAM_WIDTH = 5943600          # 6.5"
DIAGRAM_HEIGHT = 3657600         # 4.0"
NODE_WIDTH = 1600200             # 1.75"
NODE_HEIGHT = 685800             # 0.75"

COLOR_NODE = "2E75B6"            # the workbook's header blue
COLOR_START = "1F3864"           # the workbook's banner navy
COLOR_LINE = "1F3864"


def nodeCentres() -> List[Tuple[int, int]]:
    """Place six nodes clockwise around an ellipse, starting at the top.

    Parameters: none.
    Returns: list of (x, y) centres in EMU, in flow order.

    An ellipse rather than a circle because the space available is wider
    than it is tall; the arrows still read as a cycle, which is the point.
    """
    centreX, centreY = DIAGRAM_WIDTH // 2, DIAGRAM_HEIGHT // 2
    radiusX = (DIAGRAM_WIDTH - NODE_WIDTH) // 2
    radiusY = (DIAGRAM_HEIGHT - NODE_HEIGHT) // 2
    points = []
    for step in range(6):
        angle = math.radians(-90 + 60 * step)
        points.append((int(centreX + radiusX * math.cos(angle)),
                       int(centreY + radiusY * math.sin(angle))))
    return points


def shapeXml(shapeId: int, name: str, label: str, centre: Tuple[int, int],
             fill: str) -> str:
    """One labelled ellipse. Parameters: shapeId, name, label, centre, fill."""
    x = centre[0] - NODE_WIDTH // 2
    y = centre[1] - NODE_HEIGHT // 2
    return (
        '<xdr:sp macro="" textlink="">'
        f'<xdr:nvSpPr><xdr:cNvPr id="{shapeId}" name="{escapeXml(name)}"/>'
        "<xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr>"
        f'<a:xfrm><a:off x="{x}" y="{y}"/>'
        f'<a:ext cx="{NODE_WIDTH}" cy="{NODE_HEIGHT}"/></a:xfrm>'
        '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
        f'<a:ln w="19050"><a:solidFill><a:srgbClr val="{COLOR_LINE}"/>'
        "</a:solidFill></a:ln></xdr:spPr>"
        '<xdr:txBody><a:bodyPr rtlCol="0" anchor="ctr" wrap="square" '
        'lIns="36000" rIns="36000" tIns="18000" bIns="18000"/><a:lstStyle/>'
        '<a:p><a:pPr algn="ctr"/><a:r>'
        '<a:rPr lang="en-US" sz="1000" b="1">'
        '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
        '<a:latin typeface="Calibri"/></a:rPr>'
        f"<a:t>{escapeXml(label)}</a:t></a:r></a:p></xdr:txBody></xdr:sp>")


def arrowXml(shapeId: int, start: Tuple[float, float],
             end: Tuple[float, float]) -> str:
    """One arrow from `start` to `end`.

    Parameters: shapeId (int); start, end - (x, y) in EMU.
    Returns: str, the <xdr:cxnSp> XML.

    A connector is drawn left-to-right, top-to-bottom from its offset; flipH
    and flipV are what put the arrowhead on the end that the flow is going
    to, rather than the one it came from.
    """
    x0, y0 = int(min(start[0], end[0])), int(min(start[1], end[1]))
    flips = ""
    if end[0] < start[0]:
        flips += ' flipH="1"'
    if end[1] < start[1]:
        flips += ' flipV="1"'
    return (
        '<xdr:cxnSp macro="">'
        f'<xdr:nvCxnSpPr><xdr:cNvPr id="{shapeId}" name="Flow {shapeId}"/>'
        "<xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr>"
        f'<a:xfrm{flips}>'
        f'<a:off x="{x0}" y="{y0}"/>'
        f'<a:ext cx="{int(abs(end[0] - start[0]))}" '
        f'cy="{int(abs(end[1] - start[1]))}"/></a:xfrm>'
        '<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>'
        f'<a:ln w="28575"><a:solidFill><a:srgbClr val="{COLOR_LINE}"/>'
        '</a:solidFill><a:tailEnd type="triangle" w="med" len="med"/>'
        "</a:ln></xdr:spPr></xdr:cxnSp>")


def centreLabelXml(shapeId: int) -> str:
    """The caption inside the ring. Parameters: shapeId (int)."""
    width, height = 2286000, 457200
    x = DIAGRAM_WIDTH // 2 - width // 2
    y = DIAGRAM_HEIGHT // 2 - height // 2
    return (
        '<xdr:sp macro="" textlink="">'
        f'<xdr:nvSpPr><xdr:cNvPr id="{shapeId}" name="Cycle caption"/>'
        "<xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr>"
        f'<a:xfrm><a:off x="{x}" y="{y}"/>'
        f'<a:ext cx="{width}" cy="{height}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        "<a:noFill/><a:ln><a:noFill/></a:ln></xdr:spPr>"
        '<xdr:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/>'
        '<a:p><a:pPr algn="ctr"/><a:r>'
        '<a:rPr lang="en-US" sz="1200" b="1">'
        f'<a:solidFill><a:srgbClr val="{COLOR_START}"/></a:solidFill>'
        '<a:latin typeface="Calibri"/></a:rPr>'
        "<a:t>The Analytics Lifecycle</a:t></a:r></a:p>"
        '<a:p><a:pPr algn="ctr"/><a:r>'
        '<a:rPr lang="en-US" sz="900" i="1">'
        '<a:solidFill><a:srgbClr val="595959"/></a:solidFill>'
        '<a:latin typeface="Calibri"/></a:rPr>'
        "<a:t>each lap asks a better question</a:t></a:r></a:p>"
        "</xdr:txBody></xdr:sp>")


def lifecycleAnchorXml(row: int) -> str:
    """Build the whole diagram, anchored at a row.

    Parameters: row (int) - the 0-based row to anchor at.
    Returns: str, the <xdr:oneCellAnchor> XML.

    One group, so the shapes keep their positions relative to each other no
    matter what happens to the row heights around them. oneCellAnchor with
    an explicit ext, the way the screenshot on this sheet is anchored, so
    the diagram keeps its size instead of stretching to fit two cells.
    """
    centres = nodeCentres()
    shapes = [centreLabelXml(99)]
    for index, (label, centre) in enumerate(zip(LIFECYCLE_STEPS, centres)):
        fill = COLOR_START if index == 0 else COLOR_NODE
        shapes.append(shapeXml(100 + index, label, label, centre, fill))

    # Arrows run between the middle 30% of each pair of centres, which
    # leaves the ellipses clear at both ends without measuring their edges.
    for index in range(6):
        here, nxt = centres[index], centres[(index + 1) % 6]
        deltaX, deltaY = nxt[0] - here[0], nxt[1] - here[1]
        start = (here[0] + deltaX * 0.36, here[1] + deltaY * 0.36)
        end = (here[0] + deltaX * 0.64, here[1] + deltaY * 0.64)
        shapes.append(arrowXml(110 + index, start, end))

    return (
        "<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col>"
        "<xdr:colOff>152400</xdr:colOff>"
        f"<xdr:row>{row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
        f'<xdr:ext cx="{DIAGRAM_WIDTH}" cy="{DIAGRAM_HEIGHT}"/>'
        '<xdr:grpSp><xdr:nvGrpSpPr>'
        '<xdr:cNvPr id="98" name="Analytics Lifecycle"/>'
        "<xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr>"
        f'<a:xfrm><a:off x="0" y="0"/>'
        f'<a:ext cx="{DIAGRAM_WIDTH}" cy="{DIAGRAM_HEIGHT}"/>'
        f'<a:chOff x="0" y="0"/>'
        f'<a:chExt cx="{DIAGRAM_WIDTH}" cy="{DIAGRAM_HEIGHT}"/></a:xfrm>'
        "</xdr:grpSpPr>" + "".join(shapes) +
        "</xdr:grpSp><xdr:clientData/></xdr:oneCellAnchor>")


# =====================================================================
# Sheet rewrites
# =====================================================================

def rewriteCgiAi(sheetXml: str, strings: Dict[str, int],
                 scripts: Dict[str, int], textOf: Dict[str, str]) -> str:
    """Rebuild the CGI_AI sheet from scratch.

    Parameters: sheetXml (str); strings (dict); scripts (dict).
    Returns: str, the rewritten sheet.

    Rebuilt rather than patched: the live cells moved, the sections are in a
    new order, and the spilled results of the old =PY() cells had to go.
    Patching around all of that leaves stale cached values on the sheet that
    look like fresh ones.
    """
    def body(number: int, key: str, style: str = S_BODY) -> str:
        return simpleRow(number, stringCell(f"A{number}", style,
                                            strings[key]),
                         wrappedHeight(textOf[key], CGI_BAND), "1:3")

    def heading(number: int, key: str) -> str:
        return simpleRow(number, stringCell(f"A{number}", S_HEADING,
                                            strings[key]), HT_HEADING, "1:3")

    def codeAndLive(number: int, key: str, script: str,
                    cached: str = "", argument: str = "") -> str:
        cells = (stringCell(f"A{number}", S_CODE, strings[key])
                 + pythonCell(f"C{number}", scripts[script], style=S_CODE,
                              cached=cached, argument=argument))
        return simpleRow(number, cells, wrappedHeight(textOf[key], CGI_BAND),
                         "1:3")

    rows = "".join([
        simpleRow(1, stringCell("A1", S_TITLE, strings["cgiTitle"]),
                  "24", "1:3"),
        body(2, "cgiIntro", S_NOTE),

        heading(4, "sourceHeading"),
        body(5, "cgiSource"),
        body(6, "cgiFunctions"),

        heading(8, "importHeading"),
        body(9, "importIntro"),
        body(10, "sandboxNote", S_NOTE),
        body(11, "importLabel"),
        codeAndLive(12, "importCode", "import"),

        heading(14, "liveHeading"),
        body(15, "liveIntro"),
        codeAndLive(16, "basicText", "basic", cached="Hello, World!"),
        codeAndLive(17, "personalText", "personal", cached="Hello, Alice!"),
        codeAndLive(18, "detailedText", "detailed",
                    cached="Ahoy, World! Avast ye scallywags!"),

        heading(20, "surveyHeading"),
        body(21, "surveyIntro"),
        # Each sample spills eleven rows across C:Q, so the next code cell
        # starts twelve rows further down rather than underneath it.
        codeAndLive(22, "surveyText", "survey", argument=BIG_RANGE),
        codeAndLive(34, "surveySeedText", "surveySeed", argument=BIG_RANGE),

        heading(46, "outsideHeading"),
        body(47, "outsideIntro"),
        body(48, "outsideCode", S_CODE),

        body(50, "cgiClosing", S_NOTE),
    ])
    return setDimension(replaceSheetData(sheetXml, rows), "A1:Q50")


def rewriteTestScaling(sheetXml: str, strings: Dict[str, int],
                       scripts: Dict[str, int],
                       textOf: Dict[str, str]) -> str:
    """Rebuild the Test_Scaling sheet from scratch.

    Parameters: sheetXml (str); strings (dict); scripts (dict); textOf
                (dict) - each key's actual text, for the row heights.
    Returns: str, the rewritten sheet.

    Column A is six characters wide here, so every text row is merged across
    A:J. The merges are rebuilt to match the new row numbers -- a mergeCell
    pointing at a row that has moved is a merge over the wrong text.
    """
    merged: List[int] = []

    def band(number: int, style: str, key: str, height: str = "") -> str:
        merged.append(number)
        fillers = "".join(f'<c r="{columnLetter(i)}{number}" s="{T_FILL}"/>'
                          for i in range(2, 11))
        return simpleRow(number,
                         stringCell(f"A{number}", style, strings[key]) + fillers,
                         height or wrappedHeight(textOf[key], SCALE_BAND),
                         "1:10")

    def live(number: int, script: str) -> str:
        return (f'<row r="{number}" spans="1:14" x14ac:dyDescent="0.25">'
                + pythonCell(f"A{number}", scripts[script]) + "</row>")

    rows = "".join([
        band(1, T_TITLE, "scaleTitle", "24"),
        band(2, T_NOTE, "scaleIntro"),

        band(4, T_HEADING, "whatHeading", HT_HEADING),
        band(5, T_BODY, "whatBody"),

        band(7, T_HEADING, "runHeading", HT_HEADING),
        band(8, T_BODY, "runIntro"),
        band(9, T_NOTE, "important"),
        band(10, T_LABEL, "step1", HT_LABEL),
        live(12, "scaleImport"),
        band(14, T_LABEL, "step2"),
        live(16, "scaleCall"),

        band(23, T_HEADING, "readHeading", HT_HEADING),
        band(24, T_BODY, "readIntro", HT_LABEL),
        band(25, T_BODY, "readCols"),
        band(26, T_BODY, "readTimes"),
        band(27, T_BODY, "readPerCell"),
        band(28, T_BODY, "readIndex"),
        band(29, T_BODY, "readNote"),

        band(31, T_HEADING, "resultHeading", HT_HEADING),
        band(32, T_BODY, "resultBody"),

        band(34, T_HEADING, "paramsHeading", HT_HEADING),
        band(35, T_BODY, "paramsBody"),

        band(37, T_NOTE, "scaleClosing"),
    ])
    sheetXml = replaceSheetData(sheetXml, rows)

    mergeXml = "".join(f'<mergeCell ref="A{n}:J{n}"/>' for n in merged)
    sheetXml = re.sub(
        r"<mergeCells count=\"\d+\">.*?</mergeCells>",
        lambda _: f'<mergeCells count="{len(merged)}">{mergeXml}</mergeCells>',
        sheetXml, count=1, flags=re.S)
    return setDimension(sheetXml, "A1:N37")


def editReadme(sheetXml: str, strings: Dict[str, int]) -> Tuple[str, int]:
    """Add the Analytics Lifecycle section and the v1.5 version line.

    Parameters: sheetXml (str); strings (dict).
    Returns: (rewritten sheet, the 0-based row the diagram anchors at).

    Bottom-up again: the version line goes in first, so the row numbers read
    off the file are still valid when the lifecycle block is inserted above
    it.
    """
    # v1.5 under v1.4, which is row 60.
    sheetXml = shiftRows(sheetXml, 61, 1)
    sheetXml = insertAfterRow(sheetXml, 60, simpleRow(
        61, stringCell("A61", S_BODY, strings["version15"]),
        wrappedHeight(VERSION_LINE_15, README_BAND), "1:1"))

    # The lifecycle section goes above "Version history" at row 55: a
    # heading, the paragraph, then blank rows the diagram sits over.
    diagramRows = 24
    sheetXml = shiftRows(sheetXml, 55, diagramRows + 3)
    sheetXml = insertAfterRow(sheetXml, 53, (
        simpleRow(55, stringCell("A55", S_HEADING, strings["lifeHeading"]),
                  HT_HEADING, "1:1")
        + simpleRow(56, stringCell("A56", S_BODY, strings["lifeIntro"]),
                    wrappedHeight(LIFECYCLE_INTRO, README_BAND), "1:1")))
    return setDimension(sheetXml, "A1:A90"), 56      # 0-based row 56 = row 57


# =====================================================================
# Main
# =====================================================================

def build(source: Path, target: Path) -> Dict[str, object]:
    """Produce the v1.5 workbook.

    Parameters: source (Path); target (Path).
    Returns: dict - a summary of what changed.
    Raises: ValueError - a part the edit depends on is missing or has moved.
    """
    with zipfile.ZipFile(source) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
        order = list(archive.namelist())

    def text(name: str) -> str:
        if name not in parts:
            raise ValueError(f"{source.name} has no {name}")
        return parts[name].decode("utf-8")

    existing = sharedStringList(text("xl/sharedStrings.xml"))

    def find(value: str) -> int:
        """Index of an existing shared string. Raises if it is not there."""
        try:
            return existing.index(value)
        except ValueError:
            raise ValueError(f"shared string not found: {value!r}")

    # Text already in the workbook that the rebuilt sheets keep.
    keep = {
        "cgiTitle": "CGI_AI — run cgs_ai functions from your OneDrive / SharePoint",
        "sourceHeading": "Source",
        "whatHeading": "What it does",
        "runHeading": "Run it live (Python in Excel)",
        "readHeading": "How to read the result (in detail)",
        "resultHeading": "Reading the result",
        "paramsHeading": "Key parameters",
    }
    kept = {key: find(value) for key, value in keep.items()}
    # These are long; match them on a prefix rather than retyping them.
    def findPrefix(prefix: str) -> int:
        for index, value in enumerate(existing):
            if value.startswith(prefix):
                return index
        raise ValueError(f"no shared string starts with {prefix!r}")

    kept.update({
        "cgiIntro": findPrefix("cgs_ai is a reusable Python package."),
        "cgiSource": findPrefix("Package source (GitHub):"),
        "scaleTitle": findPrefix("Test_Scaling — benchmark"),
        "scaleIntro": findPrefix("This tab runs the cgs_ai testWorkload()"),
        "whatBody": findPrefix("testWorkload grows a synthetic dataset"),
        "readIntro": findPrefix("The function returns one row per size step"),
        "readCols": findPrefix("StepPct — this step's size"),
        "readTimes": findPrefix("CreateSeconds — time to build"),
        "readPerCell": findPrefix("CreateSecPerMCell / WorkSecPerMCell"),
        "readIndex": findPrefix("ScalingIndex — this step's WorkSecPerMCell"),
        "readNote": findPrefix("Note — flags any step"),
        "resultBody": findPrefix("If ScalingIndex stays near 1.0"),
        "paramsBody": findPrefix("RowCount, ColumnCount — the full (100%)"),
    })

    # New text.
    values = {
        "cgiFunctions": CGI_FUNCTIONS_LINE,
        "importHeading": CGI_IMPORT_HEADING,
        "importIntro": CGI_IMPORT_INTRO,
        "sandboxNote": SANDBOX_NOTE,
        "importLabel": CGI_IMPORT_LABEL,
        "importCode": IMPORT_SNIPPET,
        "liveHeading": CGI_LIVE_HEADING,
        "liveIntro": CGI_LIVE_INTRO,
        "basicText": CGI_BASIC_TEXT,
        "personalText": CGI_PERSONAL_TEXT,
        "detailedText": CGI_DETAILED_TEXT,
        "surveyHeading": CGI_SURVEY_HEADING,
        "surveyIntro": CGI_SURVEY_INTRO,
        "surveyText": CGI_SURVEY_TEXT,
        "surveySeedText": CGI_SURVEY_SEED_TEXT,
        "outsideHeading": CGI_OUTSIDE_HEADING,
        "outsideIntro": CGI_OUTSIDE_INTRO,
        "outsideCode": CGI_OUTSIDE_CODE,
        "cgiClosing": CGI_CLOSING,
        "runIntro": SCALE_RUNLIVE_INTRO,
        "important": SCALE_IMPORTANT,
        "step1": SCALE_STEP1,
        "step2": SCALE_STEP2,
        "scaleClosing": SCALE_CLOSING,
        "lifeHeading": LIFECYCLE_HEADING,
        "lifeIntro": LIFECYCLE_INTRO,
        "version15": VERSION_LINE_15,
    }
    keys = list(values)
    sharedXml, indices = addSharedStrings(text("xl/sharedStrings.xml"),
                                          [values[k] for k in keys])
    strings = dict(zip(keys, indices))
    strings.update(kept)
    # The two header cells Excel renamed already exist as strings; reuse them
    # rather than adding duplicates.
    strings["allowedAmount"] = find("allowed_amount")
    strings["claimType"] = find("claim_type")

    # --- python scripts ---
    scriptSource = {
        "import": IMPORT_SNIPPET,
        "basic": "cgs_ai.doBasicHello()",
        "personal": "cgs_ai.doPersonalizedHello('Alice')",
        "detailed": "cgs_ai.doDetailedHello('pirate')['message']",
        "survey": ("claims = xl(%P2%, headers=True)\n"
                   "cgs_ai.surveySelect(data=claims, n=10)"),
        "surveySeed": ("claims = xl(%P2%, headers=True)\n"
                       "cgs_ai.surveySelect(data=claims, n=10, seed=20261028)"),
        "scaleImport": IMPORT_SNIPPET,
        "scaleCall": SCALE_CALL,
    }
    scriptKeys = list(scriptSource)
    pythonXml, scriptIndices = addPythonScripts(
        text("xl/python.xml"), [scriptSource[k] for k in scriptKeys])
    scripts = dict(zip(scriptKeys, scriptIndices))

    parts["xl/sharedStrings.xml"] = sharedXml.encode("utf-8")
    parts["xl/python.xml"] = pythonXml.encode("utf-8")

    # --- the permanent table repair ---
    big = text("xl/worksheets/sheet3.xml")
    if "<tableParts" not in big:
        raise ValueError("Claims_Data_BIG no longer has a table to remove")
    big = re.sub(r"<tableParts count=\"\d+\">.*?</tableParts>", "", big,
                 count=1, flags=re.S)
    # Put back the two headers Excel renamed to dodge its own uniqueness rule.
    big = re.sub(r'(<c r="J1"[^>]*t="s"[^>]*>)<v>\d+</v>',
                 lambda m: m.group(1) + f'<v>{strings["allowedAmount"]}</v>',
                 big, count=1)
    big = re.sub(r'(<c r="N1"[^>]*t="s"[^>]*>)<v>\d+</v>',
                 lambda m: m.group(1) + f'<v>{strings["claimType"]}</v>',
                 big, count=1)
    # An autofilter on the range keeps the sort-and-filter arrows the table
    # provided. It goes after </sheetData> and before pageMargins.
    if "<autoFilter" not in big:
        big = big.replace("<pageMargins",
                          '<autoFilter ref="A1:N51"/><pageMargins', 1)
    parts["xl/worksheets/sheet3.xml"] = big.encode("utf-8")
    parts.pop("xl/tables/table1.xml", None)
    parts.pop("xl/worksheets/_rels/sheet3.xml.rels", None)

    contentTypes = text("[Content_Types].xml").replace(
        '<Override PartName="/xl/tables/table1.xml" ContentType="application/'
        'vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>', "", 1)
    parts["[Content_Types].xml"] = contentTypes.encode("utf-8")

    # --- the rewritten sheets ---
    # Every key's actual text, so each row can be given a height that fits
    # what is in it -- Excel will not work that out on its own.
    textOf = dict(values)
    textOf.update({key: existing[index] for key, index in kept.items()})

    parts["xl/worksheets/sheet11.xml"] = rewriteCgiAi(
        text("xl/worksheets/sheet11.xml"), strings, scripts,
        textOf).encode("utf-8")
    parts["xl/worksheets/sheet13.xml"] = rewriteTestScaling(
        text("xl/worksheets/sheet13.xml"), strings, scripts,
        textOf).encode("utf-8")

    readme, diagramRow = editReadme(text("xl/worksheets/sheet1.xml"), strings)
    parts["xl/worksheets/sheet1.xml"] = readme.encode("utf-8")

    # --- the diagram ---
    drawing = text("xl/drawings/drawing1.xml")
    anchor = re.search(r"<xdr:row>(\d+)</xdr:row>", drawing)
    if not anchor:
        raise ValueError("the README drawing has no row anchor")
    parts["xl/drawings/drawing1.xml"] = drawing.replace(
        "</xdr:wsDr>", lifecycleAnchorXml(diagramRow) + "</xdr:wsDr>",
        1).encode("utf-8")

    # --- recalculate on open, and drop the stale chain ---
    workbook = text("xl/workbook.xml")
    workbook = re.sub(r"<calcPr([^>]*?)/>",
                      lambda m: '<calcPr' + m.group(1)
                      + ' fullCalcOnLoad="1"/>',
                      workbook, count=1)
    parts["xl/workbook.xml"] = workbook.encode("utf-8")

    if "xl/calcChain.xml" in parts:
        parts.pop("xl/calcChain.xml")
        parts["[Content_Types].xml"] = re.sub(
            r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', "",
            parts["[Content_Types].xml"].decode("utf-8")).encode("utf-8")
        rels = re.sub(r'<Relationship [^>]*calcChain[^>]*/>', "",
                      text("xl/_rels/workbook.xml.rels"))
        parts["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")

    # --- write, keeping the original part order ---
    order = [n for n in order if n in parts]
    order += [n for n in parts if n not in order]
    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        for name in order:
            archive.writestr(name, parts[name])

    return {"Target": str(target), "NewStrings": len(keys),
            "NewScripts": len(scriptKeys), "ScriptIndices": scripts,
            "DiagramAnchorRow0Based": diagramRow,
            "TableRemoved": "Claims_Data_BIG (now a plain range)"}


def main(argv: Sequence[str]) -> int:
    """CLI entry point. Parameters: argv. Returns: an exit code."""
    if len(argv) != 2:
        print("usage: build_claims_pipelines_v15.py <source.xlsx> "
              "<target.xlsx>", file=sys.stderr)
        return EXIT_USAGE
    for key, value in build(Path(argv[0]), Path(argv[1])).items():
        print(f"{key}: {value}")
    return EXIT_OK


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
