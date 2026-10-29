"""
=====================================================================
  Program Name  : build_claims_pipelines_workbook.py
  Author        : Manuel Figallo
  Purpose       : Turn python_for_medicare_claims_synthetic.xlsx into
                  python_for_medicare_claims_pipelines.xlsx.
  Version       : 1.0beta
  Created       : 2026-10-28
  Last Modified : 2026-10-28

  Dependencies:
    STANDARD LIBRARY ONLY.

  Description:
    RAW XML SURGERY, NOT openpyxl. A round-trip through openpyxl drops 31
    parts from this workbook -- all seven images, xl/python.xml and with it
    every =PY() cell. So the archive is unzipped, the handful of parts that
    change are edited as text, and everything else is copied through byte
    for byte.

    A Python-in-Excel cell is three things: an entry in xl/python.xml, a
    cell holding =_xlfn._xlws.PY(n,0) as an array formula, and cm="1" to
    mark it a dynamic array. The index n is ZERO-BASED into python.xml's
    <pythonScript> list -- PY(14,0) in this workbook is the fifteenth
    script, which is how it was confirmed rather than assumed. A reference
    the code needs is NOT written into the code: the code says
    xl(%P2%, headers=True) and the cell's formula carries the real range as
    its third argument, PY(n,0,Claims_Data_BIG[#All]).

    The five changes:
      1. README gains a Background section after Objective, a new title and
         subtitle, an updated tab list and a v1.4 version-history line.
      2. A new Claims_Data_BIG sheet follows Claims_Data: 50 rows carrying
         every column in the Metadata data dictionary, as a real Excel
         table so xl("Claims_Data_BIG[#All]") resolves.
      3. CGI_AI shows the updated cgs_ai __init__.py, surveySelect source
         and all, with three new live =PY() previews.
      4. "Test Scaling" becomes "Test_Scaling" and gains the cgs_ai import
         pattern plus two live calls, keeping every section it had.
      5. calcChain.xml is dropped and fullCalcOnLoad is set, so Excel
         recomputes the Python cells instead of trusting a stale chain.

  Input Parameters (required first):
    source (REQUIRED) - python_for_medicare_claims_synthetic.xlsx
    target (REQUIRED) - the .xlsx to write

  Usage:
    python build_claims_pipelines_workbook.py <source.xlsx> <target.xlsx>

  Exit codes:
    0 = success, 2 = usage error.
=====================================================================
"""

from __future__ import annotations

import math
import random
import re
import sys
import zipfile
from pathlib import Path
from typing import Dict, List, Sequence, Tuple

__version__ = "1.0beta"

EXIT_OK, EXIT_USAGE = 0, 2

# Style ids already defined in the workbook's styles.xml. Reusing them is
# what keeps the new rows looking like the old ones; adding a style would
# mean renumbering cellXfs, and every existing cell points into that list.
S_TITLE = "5"        # 16pt bold, the sheet title
S_HEADING = "6"      # 12pt bold, a section heading
S_BODY = "7"         # wrapped body text
S_CODE = "8"         # Consolas on a grey fill, wrapped -- a code block
S_NOTE = "9"         # grey italic, a footnote
S_LABEL = "10"       # bold, unwrapped -- a one-line label or table header
S_CLAIMS_HEADER = "1"  # centred and bordered: the Claims_Data header look

# Test Scaling has its own palette. Same look, different ids.
T_HEADING = "16"
T_BODY = "17"
T_NOTE = "15"
T_LABEL = "18"

#: Row heights, in points. Excel does not auto-fit a merged or wrapped cell
#: on open, so every row this script writes states its own height.
HT_HEADING = "20.100000000000001"
HT_LABEL = "18.95"
HT_LINE = 15.0

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


# =====================================================================
# Content
# =====================================================================

NEW_TITLE = "Python for Medicare Claims — Building Analytics Pipelines"

NEW_SUBTITLE = ("A hands-on Excel + Python (Python in Excel) workbook for "
                "building claims analytics pipelines with pandas on a "
                "synthetic Medicare dataset — from correlated variables "
                "through to a linear regression model.")

BACKGROUND = (
    "Python training and shared practice are what turn a one-off analysis "
    "into a pipeline somebody else can run. When every analyst solves the "
    "same problem privately, each solution carries its own assumptions, its "
    "own bugs and its own silent failures — and none of it survives the "
    "person who wrote it. Agreeing on how we read a claims extract, how we "
    "name things, how we check our work and how we hand it on is what makes "
    "a result reproducible and reviewable, which is the standard claims "
    "analytics has to meet. That is also why the functions in this workbook "
    "come out of a shared package rather than a private notebook: knowledge "
    "that is written down as code other people can call is the only kind "
    "that compounds. The Python covered here walks the whole analytics "
    "lifecycle in order — understand the data types and structures, "
    "describe the distributions, find the correlated variables, see those "
    "relationships in a scatter and a pair plot, and finish by fitting a "
    "linear regression model. Each tab is one step of that lifecycle, and "
    "together they are the shape of every claims pipeline you will build: "
    "iterate through them, and the same pattern carries from five rows to "
    "five million.")

TABS_LINE = ("README · Claims_Data · Claims_Data_BIG · Metadata · Data Types "
             "and Structures · Describe Data · Correlation Matrix · Scatter "
             "Plot · Pair Plot · Linear Regression · CGI_AI · Random Sample · "
             "Test_Scaling · Streamlit · Contact.")

HOW_TO_LINE = (
    "1) Open in Excel with Python in Excel enabled.  2) Read Claims_Data to "
    "see the example claim records, and Claims_Data_BIG for a 50-row version "
    "carrying every field in the data dictionary.  3) On Data Types and "
    "Structures, read each commented code block and select a cell in column C "
    "to run it live with =PY().  4) Try CGI_AI, Random Sample, Test_Scaling "
    "and Streamlit.  5) Use Metadata to look up fields.")

VERSION_LINE = (
    "v1.4 (2026-10-28) — Renamed to python_for_medicare_claims_pipelines. "
    "Added this Background section; added Claims_Data_BIG (50 rows, every "
    "column in the data dictionary) as a real Excel table; CGI_AI now shows "
    "the updated cgs_ai __init__.py carrying surveySelect, with live "
    "previews; Test Scaling renamed Test_Scaling and now shows the cgs_ai "
    "import pattern and calls testWorkload live.")

# --- CGI_AI ----------------------------------------------------------------

FUNCTIONS_LINE = (
    "Functions shown here:  doBasicHello()  ->  returns \"Hello, World!\" ;   "
    "doPersonalizedHello(name)  ->  returns \"Hello, <name>!\" (falls back to "
    "\"World\" if name is blank) ;   doDetailedHello(style)  ->  returns a "
    "dictionary, in three styles ;   surveySelect(data, n, out)  ->  returns "
    "a reproducible random sample ;   testWorkload(RowCount, ColumnCount, "
    "Increment)  ->  returns a scaling benchmark (see the Test_Scaling tab).")

SHARE = (r"https://bcbsscgov-my.sharepoint.us/my?id=%2Fpersonal%2Fmanuel%5F"
         r"figallo%5Fcgsadmin%5Fcom%2FDocuments%2Fcode%2Fpython%2Fcgs%5Fai&ga=1")

SHARE_SNIPPET = '''# The folder that CONTAINS the cgs_ai package folder.
SHARE = r"''' + SHARE + r'''"
CGS_AI_HOME = SHARE + r"\src\py\lite"

# STEP 1. Import the python package
# Telling Python where to look, then importing cgs_ai exactly like pandas.
import sys
sys.path.insert(0, CGS_AI_HOME)
import cgs_ai

# STEP 2. Run examples
# Example 1 -- the smoke test: if this prints, the import worked.
result = cgs_ai.doBasicHello()
print(result)

# Example 2 -- a function that takes an argument and gives a different answer.
result = cgs_ai.doPersonalizedHello('Alice')
print(result)

# Example 3 -- a function that returns a STRUCTURE, not a sentence.
result = cgs_ai.doDetailedHello('formal')
print(result['message'])

# Example 4 -- surveySelect: a reproducible random sample. Same seed, same
# rows, on any machine -- which is what makes a sample auditable.
import pandas as pd
claims = pd.read_csv('claims.csv')
sample = cgs_ai.surveySelect(data=claims, n=30)
print(sample)

# Example 5 -- testWorkload: does ten times the data cost ten times the
# time, or more? Read the ScalingIndex column.
table = cgs_ai.testWorkload(RowCount=200000, ColumnCount=12, Increment=25)
print(table)'''

LIVE_PREVIEW_LINE = (
    "The cells in column C run self-contained copies of these functions so "
    "you can see their output right here in Excel. The code block above is "
    "the real pattern that imports cgs_ai from your OneDrive/SharePoint "
    "folder.")

DETAILED_CODE_TEXT = '''# Example 3 — a greeting that returns a structure
result = cgs_ai.doDetailedHello('pirate')
print(result['message'])   # -> Ahoy, World! Avast ye scallywags!'''

DETAILED_SCRIPT = '''def doDetailedHello(style="friendly"):
    styles = {
        "friendly": "Hello, World! Wonderful day, isn't it?",
        "formal": "Greetings, World. It is a pleasure to connect.",
        "pirate": "Ahoy, World! Avast ye scallywags!",
    }
    greeting = styles.get(str(style).lower(), styles["friendly"])
    return {"message": greeting, "style_used": style}
str(doDetailedHello('pirate')['message'])'''

SURVEY_SELECT_SOURCE = '''def surveySelect(data, n, out="",
                 seed=8111305759926929283):
    """Draw a reproducible random sample of n rows.

    Parameters:
        data  - the rows to sample: a pandas DataFrame, or anything
                DataFrame() accepts. From Excel, pass
                xl("Claims_Data_BIG[#All]", headers=True).
        n     - how many sample units to draw.
        out   - optional .csv or .xlsx path to write the sample to. The
                sample is RETURNED either way, so in Excel -- which has no
                filesystem -- leave it out entirely.
        seed  - the number that makes the draw repeatable. Change it and you
                get a different sample; keep it and you get the same one
                every time, on any machine. Pass 0 for a random draw.
    Returns:
        A DataFrame of n rows, indexed 1..n the way SAS numbers a sample.

    PCG64DXSM is a bit generator from the PCG family. It is used instead of
    the default because it is explicitly seeded and its stream is
    reproducible across platforms and numpy versions -- which is what an
    auditable sample needs. A sample nobody else can reproduce is an
    anecdote, not evidence.
    """
    import random
    import sys
    import pandas as pd
    from numpy.random import PCG64DXSM

    n = int(n)
    if n < 1:
        raise ValueError(f"n must be at least 1, got {n}")
    seed = seed if seed else random.randint(0, sys.maxsize)
    frame = data if isinstance(data, pd.DataFrame) else pd.DataFrame(data)
    if n > len(frame):
        raise ValueError(f"asked for {n} sample units but the data has only "
                         f"{len(frame)} rows")

    generator = PCG64DXSM(seed)
    sample = frame.sample(n=n, random_state=generator).reset_index(drop=True)
    sample.index = sample.index + 1          # 1-based, as SAS numbers them

    if out:
        if str(out).lower().endswith((".xlsx", ".xlsm")):
            sample.to_excel(out, index=True)
        else:
            sample.to_csv(out, index=True)
    return sample'''

SURVEY_INTRO = (
    "surveySelect is new in this version of cgs_ai. It is the Random Sample "
    "tab's code turned into a function anybody can call: pass it a DataFrame "
    "and a number of sample units, and it hands back that many rows, "
    "numbered from 1. The whole point is the seed — run it again next month "
    "with the same seed and you get the same rows, which is what lets you "
    "defend a Medicare review sample after the fact. This is the function as "
    "it appears in __init__.py; the complete file is at the GitHub link in "
    "the Source section above.")

SURVEY_SIBLING = (
    "testWorkload is in the same file, added at the same time. It measures "
    "how an environment copes as a dataset grows — the Test_Scaling tab runs "
    "it live and explains how to read the result.")

SURVEY_LIVE_INTRO = (
    "These two cells sample the Claims_Data_BIG sheet. They are the same "
    "call with two different seeds, which is the fastest way to see what the "
    "seed actually does: the first pair of rows is identical every time you "
    "recalculate, and changing the seed changes which rows come out. Select "
    "cell C25 or C36 to see the =PY() formula.")

SURVEY_DEMO_TEXT = '''# surveySelect — a reproducible 10-row sample of Claims_Data_BIG.
# The seed is what makes it repeatable: same seed, same ten rows, always.
claims = xl("Claims_Data_BIG[#All]", headers=True)
surveySelect(data=claims, n=10)'''

SURVEY_DEMO_SCRIPT = '''# surveySelect, inlined: Python in Excel runs in a Microsoft container with
# no filesystem and no network, so it cannot reach the share that holds
# cgs_ai. Everywhere else, you import it -- see the code block above.
import random
import sys
import pandas as pd
from numpy.random import PCG64DXSM

def surveySelect(data, n, out="", seed=8111305759926929283):
    n = int(n)
    if n < 1:
        raise ValueError(f"n must be at least 1, got {n}")
    seed = seed if seed else random.randint(0, sys.maxsize)
    frame = data if isinstance(data, pd.DataFrame) else pd.DataFrame(data)
    if n > len(frame):
        raise ValueError(f"asked for {n} sample units but the data has only "
                         f"{len(frame)} rows")
    generator = PCG64DXSM(seed)
    sample = frame.sample(n=n, random_state=generator).reset_index(drop=True)
    sample.index = sample.index + 1
    return sample

claims = xl(%P2%, headers=True)
surveySelect(data=claims, n=10)'''

SURVEY_SEED_TEXT = '''# The same call with a different seed — a different ten rows.
# Change the seed when you want a fresh sample; keep it when you need the
# sample from last month's report back.
claims = xl("Claims_Data_BIG[#All]", headers=True)
surveySelect(data=claims, n=10, seed=20261028)'''

SURVEY_SEED_SCRIPT = '''# surveySelect is already defined by the cell above: Python in Excel
# evaluates a workbook's =PY() cells in one shared namespace, in calculation
# order, so a function defined once is in scope further down.
claims = xl(%P2%, headers=True)
surveySelect(data=claims, n=10, seed=20261028)'''

CGI_CLOSING = ("Claims_Data_BIG is synthetic and contains no PHI or PII, and "
               "surveySelect writes no files when `out` is left empty — which "
               "is why both run unchanged inside the Python in Excel sandbox.")

# --- Test_Scaling ----------------------------------------------------------

SCALE_TITLE = "Test_Scaling — benchmark how a workload scales (runs live)"

SCALE_HEADING = "Run it from cgs_ai — the package, not a pasted copy"

SCALE_INTRO = (
    "testWorkload now ships inside the cgs_ai package, alongside "
    "surveySelect. Outside Excel you no longer paste the module — you import "
    "it, exactly as you would pandas, and call cgs_ai.testWorkload(). The two "
    "code blocks below are the real pattern; copy either into Anaconda, "
    "Jupyter, VS Code or a Snowflake notebook.")

SCALE_WHY = (
    "Why the live cell in A12 above still carries a copy of the module: "
    "Python in Excel runs in a Microsoft-managed container with no filesystem "
    "and no network, so it cannot reach the OneDrive/SharePoint share that "
    "holds cgs_ai. In Excel you embed the code; everywhere else you import "
    "it. It is the same function, and the same numbers, either way.")

SCALE_LABEL_1 = "Example 1 — import cgs_ai and measure a small claims dataset:"
SCALE_LABEL_2 = ("Example 2 — import cgs_ai and ask whether a year of claims "
                 "will still finish:")

SCALE_DEMO_SMALL = '''# testWorkload from cgs_ai — a first look at how an environment copes as a
# claims dataset grows. Start small and see what it costs before going big.
import sys
sys.path.insert(0, CGS_AI_HOME)   # the folder that CONTAINS the cgs_ai folder
import cgs_ai

RowCount    = 50000     # rows at 100%
ColumnCount = 12        # columns in the synthetic dataset
Increment   = 25        # 25% -> four steps: 25, 50, 75, 100
cgs_ai.testWorkload(RowCount, ColumnCount, Increment, Repeats=3, Label="small")'''

SCALE_DEMO_CLAIMS = '''# The question that matters for a pipeline: will the monthly run still
# finish when it is a year of claims? Measure it, do not guess.
# Read ScalingIndex: 1.0 is linear; above 1.0 means each row got MORE
# expensive as the data grew, and that pipeline will stop finishing.
import sys
sys.path.insert(0, CGS_AI_HOME)
import cgs_ai

cgs_ai.testWorkload(RowCount=400000, ColumnCount=14, Increment=25,
                    Repeats=3, Label="year-of-claims")'''

SCALE_LIVE_HEADING = "The same two examples, live in Excel"

SCALE_LIVE_INTRO = (
    "These two =PY() cells call testWorkload directly, with no copy of the "
    "module in them. They can, because Python in Excel evaluates a workbook's "
    "=PY() cells in one shared namespace in calculation order — so the module "
    "loaded by cell A12 above is still in scope down here. Select cell A48 or "
    "A56 to see the call, and compare the two ScalingIndex columns.")

SCALE_LIVE_LABEL_1 = "Example 1 live — a small run (50,000 rows at full size):"
SCALE_LIVE_LABEL_2 = ("Example 2 live — eight times the rows. This one takes "
                      "longer; if it is cancelled, lower RowCount.")

SCALE_LIVE_SMALL = '''# Example 1, live. testWorkload comes from cell A12 above.
RowCount    = 50000     # rows at 100%
ColumnCount = 12        # columns in the synthetic dataset
Increment   = 25        # 25% -> four steps: 25, 50, 75, 100
testWorkload(RowCount, ColumnCount, Increment, Repeats=3, Label="small")'''

SCALE_LIVE_CLAIMS = '''# Example 2, live: a year of claims. Compare its ScalingIndex column with
# Example 1's -- if the index climbs as the data grows, the work is getting
# more expensive per row and the pipeline will not scale.
testWorkload(RowCount=400000, ColumnCount=14, Increment=25,
             Repeats=3, Label="year-of-claims")'''

SCALE_CLOSING = (
    "New in v1.4: testWorkload and surveySelect both live in the cgs_ai "
    "package now, so the code on this tab and the code on CGI_AI are the same "
    "code. The synthetic data still contains no PHI/PII and the function "
    "still reads and writes no files, which is why it runs unchanged inside "
    "the Python in Excel sandbox.")

# --- Claims_Data_BIG -------------------------------------------------------

#: Every variable in the Metadata data dictionary, in dictionary order. The
#: first nine are the Claims_Data columns; the last five are the analysis
#: columns the Describe / Correlation / Regression tabs use, which is why
#: allowed_amount and Allowed_Amount both appear -- they are two different
#: variables in the dictionary, one per naming convention.
BIG_COLUMNS = [
    "Claim_ID", "Bene_ID", "Claim_Type", "From_Date", "Thru_Date",
    "ICD10_Diagnosis_Code", "HCPCS_Code", "Allowed_Amount",
    "Medicare_Payment_Amount", "allowed_amount", "medicare_payment",
    "patient_responsibility", "length_of_stay", "claim_type",
]
NUMERIC_COLUMNS = frozenset({"Allowed_Amount", "Medicare_Payment_Amount",
                             "allowed_amount", "medicare_payment",
                             "patient_responsibility", "length_of_stay"})
BIG_ROW_COUNT = 50
BIG_SEED = 8111305759926929283


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

    Parameters:
        text (str)          - the cell contents.
        charsPerLine (int)  - how many characters fit across the column.
        minimum (float)     - never return less than this.
    Returns: str, a height suitable for ht="...".

    Excel does NOT auto-fit a wrapped or merged cell when the file is opened,
    so a row written without a height shows one line and hides the rest. An
    estimate that is slightly too tall is harmless; one that is too short
    silently truncates the text the reader came for.
    """
    lines = 0
    for paragraph in str(text).split("\n"):
        lines += max(1, math.ceil(len(paragraph) / charsPerLine))
    return str(max(minimum, round(lines * HT_LINE + 4, 2)))


def addSharedStrings(sharedXml: str, values: Sequence[str]) -> Tuple[str, List[int]]:
    """Append strings to sharedStrings.xml.

    Parameters: sharedXml (str); values (sequence[str]) - strings to add.
    Returns: (new xml, list of 0-based indices for the added strings).

    Appending never renumbers what is already there, so no existing cell has
    to be touched.
    """
    count = int(re.search(r'\bcount="(\d+)"', sharedXml).group(1))
    unique = int(re.search(r'uniqueCount="(\d+)"', sharedXml).group(1))
    indices = list(range(unique, unique + len(values)))
    addition = "".join(f'<si><t xml:space="preserve">{escapeXml(v)}</t></si>'
                       for v in values)
    sharedXml = sharedXml.replace("</sst>", addition + "</sst>")
    sharedXml = sharedXml.replace(
        f'count="{count}" uniqueCount="{unique}"',
        f'count="{count + len(values)}" uniqueCount="{unique + len(values)}"', 1)
    return sharedXml, indices


def addPythonScripts(pythonXml: str, codes: Sequence[str]) -> Tuple[str, List[int]]:
    """Append code blocks to xl/python.xml.

    Parameters: pythonXml (str); codes (sequence[str]).
    Returns: (new xml, list of ZERO-BASED indices for =PY(n,0)).

    Zero-based because that is what this workbook's existing cells do:
    PY(14,0) on CGI_AI runs the fifteenth <pythonScript>. Guessing 1-based
    would have silently run the wrong script in every new cell.
    """
    existing = len(re.findall(r"<pythonScript>", pythonXml))
    indices = list(range(existing, existing + len(codes)))
    addition = "".join(f"<pythonScript><code>{escapeXml(code)}</code>"
                       f"</pythonScript>" for code in codes)
    return pythonXml.replace("</pythonScripts>",
                             addition + "</pythonScripts>", 1), indices


def shiftRows(sheetXml: str, fromRow: int, by: int) -> str:
    """Move every row at or below `fromRow` down by `by` rows.

    Parameters: sheetXml (str); fromRow (int) 1-based; by (int).
    Returns: str, the rewritten sheet.

    The <row r="..."> attribute, every cell reference inside it and every
    array-formula ref all have to move together; a cell still saying A7
    inside a row now numbered A10 is the kind of mismatch Excel repairs by
    throwing the sheet away.
    """
    def row(match: "re.Match") -> str:
        number = int(match.group(1))
        if number < fromRow:
            return match.group(0)
        return match.group(0).replace(f'r="{number}"',
                                      f'r="{number + by}"', 1)

    def cell(match: "re.Match") -> str:
        column, number = match.group(1), int(match.group(2))
        if number < fromRow:
            return match.group(0)
        return f'<c r="{column}{number + by}"'

    def ref(match: "re.Match") -> str:
        column, number = match.group(1), int(match.group(2))
        if number < fromRow:
            return match.group(0)
        return f'ref="{column}{number + by}"'

    sheetXml = re.sub(r'<row r="(\d+)"', row, sheetXml)
    sheetXml = re.sub(r'<c r="([A-Z]+)(\d+)"', cell, sheetXml)
    return re.sub(r'ref="([A-Z]+)(\d+)"', ref, sheetXml)


def insertAfterRow(sheetXml: str, afterRow: int, rowsXml: str) -> str:
    """Insert raw <row> XML immediately after a given row number.

    Parameters: sheetXml (str); afterRow (int); rowsXml (str).
    Returns: str, the rewritten sheet.
    Raises: ValueError - that row is not in the sheet.
    """
    match = re.search(r'<row r="%d"[^>]*>.*?</row>' % afterRow, sheetXml, re.S)
    if not match:
        raise ValueError(f"row {afterRow} is not in this sheet")
    return sheetXml[:match.end()] + rowsXml + sheetXml[match.end():]


def appendRows(sheetXml: str, rowsXml: str) -> str:
    """Append raw <row> XML at the end of <sheetData>."""
    return sheetXml.replace("</sheetData>", rowsXml + "</sheetData>", 1)


def setDimension(sheetXml: str, ref: str) -> str:
    """Replace the sheet's declared dimension. Parameters: sheetXml, ref."""
    return re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="{ref}"/>',
                  sheetXml, count=1)


def replaceStringCell(sheetXml: str, reference: str, index: int) -> str:
    """Point an existing shared-string cell at a different string.

    Parameters: sheetXml (str); reference (str) e.g. "A1"; index (int),
                0-based into sharedStrings.
    Returns: str, the rewritten sheet.
    Raises: ValueError - the cell is missing or is not a shared string.
    """
    pattern = re.compile(r'(<c r="%s"[^>]*\bt="s"[^>]*>)<v>\d+</v>' % reference)
    if not pattern.search(sheetXml):
        raise ValueError(f"{reference} is not a shared-string cell")
    return pattern.sub(lambda m: m.group(1) + f"<v>{index}</v>", sheetXml,
                       count=1)


def textRow(number: int, cells: Sequence[Tuple[str, str, int]],
            height: str, spans: str = "1:1") -> str:
    """Build one row of shared-string cells.

    Parameters:
        number (int) - the row number.
        cells        - (columnLetter, styleId, sharedStringIndex) each.
        height (str) - ht="..." value.
        spans (str)  - the row's spans attribute.
    Returns: str, the <row> XML.
    """
    body = "".join(f'<c r="{column}{number}" s="{style}" t="s">'
                   f"<v>{index}</v></c>" for column, style, index in cells)
    return (f'<row r="{number}" spans="{spans}" ht="{height}" '
            f'customHeight="1" x14ac:dyDescent="0.25">{body}</row>')


def pythonCell(reference: str, scriptIndex: int, style: str = "",
               cached: str = "", spill: str = "", argument: str = "") -> str:
    """Build a Python-in-Excel cell.

    Parameters:
        reference (str)   - the anchor cell, e.g. "C25".
        scriptIndex (int) - ZERO-BASED index into python.xml.
        style (str)       - style id, or "" for the sheet default.
        cached (str)      - the value shown until Excel recalculates. Leave
                            empty for anything that spills.
        spill (str)       - the range the result spills over, e.g. "C25:P35".
                            Defaults to the anchor cell alone.
        argument (str)    - an Excel reference passed to the script, which
                            reads it as %P2%.
    Returns: str, the <c> XML.

    cm="1" points at cellMetadata entry 1, which marks the cell a dynamic
    array. Without it Excel treats _xlfn._xlws.PY as an unknown function and
    shows #NAME?.
    """
    attributes = f'<c r="{reference}"'
    if style:
        attributes += f' s="{style}"'
    attributes += ' t="str" cm="1"'
    extra = f",{argument}" if argument else ""
    value = f"<v>{escapeXml(cached)}</v>" if cached else "<v/>"
    return (f'{attributes}><f t="array" ref="{spill or reference}">'
            f"_xlfn._xlws.PY({scriptIndex},0{extra})</f>{value}</c>")


# =====================================================================
# Claims_Data_BIG
# =====================================================================

def buildBigRows(count: int = BIG_ROW_COUNT,
                 seed: int = BIG_SEED) -> List[Dict[str, object]]:
    """Build the Claims_Data_BIG rows.

    Parameters: count (int) - how many rows; seed (int) - reproducibility.
    Returns: list of dicts keyed by BIG_COLUMNS.

    Seeded, so rebuilding the workbook produces the same 50 rows and a diff
    shows only what actually changed. The values follow the Metadata data
    dictionary: Medicare pays roughly 65-92% of the allowed amount, the
    residual is the beneficiary's, only Inpatient claims have a real length
    of stay, and an Inpatient claim may have no line-item HCPCS code.
    """
    generator = random.Random(seed)
    types = ["Inpatient", "Outpatient", "Carrier/Physician", "DME"]
    analysisTypes = {"Inpatient": "Inpatient", "Outpatient": "Outpatient",
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
            "claim_type": analysisTypes[claimType],
        })
    return rows


def buildBigSheetXml(rows: Sequence[Dict[str, object]]) -> str:
    """Build the whole Claims_Data_BIG worksheet part.

    Parameters: rows (sequence[dict]).
    Returns: str, the sheet XML.

    Inline strings rather than shared strings: 700 new <si> entries would
    renumber nothing, but they would make the sharedStrings diff unreadable,
    and inline costs nothing at this size. autoFilter MUST come before
    tableParts, and tableParts last -- the schema's child order is not
    advisory, and Excel repairs a sheet that gets it wrong.
    """
    lastColumn = columnLetter(len(BIG_COLUMNS))
    lastRow = len(rows) + 1
    span = f"1:{len(BIG_COLUMNS)}"

    header = "".join(
        f'<c r="{columnLetter(i)}1" s="{S_CLAIMS_HEADER}" t="inlineStr">'
        f'<is><t xml:space="preserve">{escapeXml(name)}</t></is></c>'
        for i, name in enumerate(BIG_COLUMNS, start=1))
    body = [f'<row r="1" spans="{span}" x14ac:dyDescent="0.25">{header}</row>']

    for number, row in enumerate(rows, start=2):
        cells = []
        for i, name in enumerate(BIG_COLUMNS, start=1):
            reference = f"{columnLetter(i)}{number}"
            value = row[name]
            if name in NUMERIC_COLUMNS:
                cells.append(f'<c r="{reference}"><v>{value}</v></c>')
            elif value == "":
                continue                      # a blank cell is simply absent
            else:
                cells.append(f'<c r="{reference}" t="inlineStr">'
                             f'<is><t xml:space="preserve">'
                             f"{escapeXml(value)}</t></is></c>")
        body.append(f'<row r="{number}" spans="{span}" '
                    f'x14ac:dyDescent="0.25">{"".join(cells)}</row>')

    widths = "".join(
        f'<col min="{i}" max="{i}" width="{max(12, min(26, len(name) + 4))}" '
        f'customWidth="1"/>' for i, name in enumerate(BIG_COLUMNS, start=1))

    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<worksheet xmlns="{NS_MAIN}" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships" xmlns:mc="http://schemas.openxmlformats.org/'
            'markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" '
            'xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/'
            '2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/'
            'spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft'
            '.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://'
            'schemas.microsoft.com/office/spreadsheetml/2016/revision3" '
            'xr:uid="{A1B2C3D4-0000-4000-8000-0000000B1600}">'
            f'<dimension ref="A1:{lastColumn}{lastRow}"/>'
            '<sheetViews><sheetView workbookViewId="0">'
            '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" '
            'state="frozen"/></sheetView></sheetViews>'
            '<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>'
            f"<cols>{widths}</cols>"
            f'<sheetData>{"".join(body)}</sheetData>'
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" '
            'header="0.3" footer="0.3"/>'
            '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>'
            "</worksheet>")


def buildBigTableXml(rowCount: int) -> str:
    """Build the Claims_Data_BIG table part.

    Parameters: rowCount (int) - data rows, excluding the header.
    Returns: str, the table XML.

    A real table, not just a range, so xl("Claims_Data_BIG[#All]") resolves
    the way the live =PY() cells on CGI_AI are written. Table id 7: ids 1, 2
    and 4-6 are taken, and two tables sharing an id is a repair prompt.
    """
    reference = f"A1:{columnLetter(len(BIG_COLUMNS))}{rowCount + 1}"
    columns = "".join(
        f'<tableColumn id="{i}" name="{escapeXml(name)}"/>'
        for i, name in enumerate(BIG_COLUMNS, start=1))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<table xmlns="{NS_MAIN}" id="7" name="Claims_Data_BIG" '
            f'displayName="Claims_Data_BIG" ref="{reference}" '
            'totalsRowShown="0">'
            f'<autoFilter ref="{reference}"/>'
            f'<tableColumns count="{len(BIG_COLUMNS)}">{columns}'
            "</tableColumns>"
            '<tableStyleInfo name="TableStyleLight1" showFirstColumn="0" '
            'showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>'
            "</table>")


# =====================================================================
# The five edits
# =====================================================================

def editReadme(sheetXml: str, strings: Dict[str, int]) -> str:
    """Rewrite the README sheet.

    Parameters: sheetXml (str); strings (dict) - shared-string indices.
    Returns: str, the rewritten sheet.

    Worked bottom-up: inserting the version line first, then the Background
    block, means the row numbers in the second edit are still the ones read
    off the original file. Top-down, every number after the first insert
    would have had to be adjusted by hand.
    """
    # Title, subtitle, tab list and the how-to line, all in place.
    sheetXml = replaceStringCell(sheetXml, "A1", strings["title"])
    sheetXml = replaceStringCell(sheetXml, "A2", strings["subtitle"])
    sheetXml = replaceStringCell(sheetXml, "A58", strings["howTo"])
    sheetXml = replaceStringCell(sheetXml, "A61", strings["tabs"])
    for number, line in ((58, HOW_TO_LINE), (61, TABS_LINE)):
        sheetXml = re.sub(
            r'(<row r="%d"[^>]*?)ht="[\d.]+"' % number,
            lambda m, t=line: m.group(1) + 'ht="%s"' % wrappedHeight(t, 105),
            sheetXml, count=1)

    # The new version-history line, under v1.3.
    sheetXml = shiftRows(sheetXml, 68, 1)
    sheetXml = insertAfterRow(sheetXml, 67, textRow(
        68, [("A", S_BODY, strings["version"])],
        wrappedHeight(VERSION_LINE, 105)))

    # Background: a heading, a body paragraph and a blank row, after
    # Objective and before Acknowledgments.
    sheetXml = shiftRows(sheetXml, 7, 3)
    sheetXml = insertAfterRow(sheetXml, 5, (
        textRow(7, [("A", S_HEADING, strings["backgroundHeading"])],
                HT_HEADING)
        + textRow(8, [("A", S_BODY, strings["background"])],
                  wrappedHeight(BACKGROUND, 105))))
    return setDimension(sheetXml, "A1:A73")


def editCgiAi(sheetXml: str, strings: Dict[str, int],
              scripts: Dict[str, int]) -> str:
    """Rewrite the CGI_AI sheet. Parameters: sheetXml, strings, scripts."""
    sheetXml = replaceStringCell(sheetXml, "A6", strings["functions"])
    sheetXml = replaceStringCell(sheetXml, "A10", strings["share"])
    sheetXml = replaceStringCell(sheetXml, "A13", strings["livePreview"])
    # The code block grew from 19 lines to 37; its row height has to grow too
    # or the reader sees the first half of it.
    for row, text in ((6, FUNCTIONS_LINE), (10, SHARE_SNIPPET),
                      (13, LIVE_PREVIEW_LINE)):
        sheetXml = re.sub(
            r'(<row r="%d"[^>]*?)ht="[\d.]+"' % row,
            lambda m, t=text: m.group(1) + 'ht="%s"' % wrappedHeight(t, 100),
            sheetXml, count=1)

    rows = [
        # Third greeting, beside the two that are already there.
        f'<row r="16" spans="1:3" ht="{wrappedHeight(DETAILED_CODE_TEXT, 92)}"'
        f' customHeight="1" x14ac:dyDescent="0.25">'
        f'<c r="A16" s="{S_CODE}" t="s"><v>{strings["detailedCode"]}</v></c>'
        + pythonCell("C16", scripts["detailed"], style=S_CODE,
                     cached="Ahoy, World! Avast ye scallywags!")
        + "</row>",

        textRow(18, [("A", S_HEADING, strings["surveyHeading"])], HT_HEADING,
                spans="1:3"),
        textRow(19, [("A", S_BODY, strings["surveyIntro"])],
                wrappedHeight(SURVEY_INTRO, 100), spans="1:3"),
        textRow(20, [("A", S_CODE, strings["surveySource"])],
                wrappedHeight(SURVEY_SELECT_SOURCE, 92), spans="1:3"),
        textRow(21, [("A", S_BODY, strings["surveySibling"])],
                wrappedHeight(SURVEY_SIBLING, 100), spans="1:3"),

        textRow(23, [("A", S_HEADING, strings["surveyLiveHeading"])],
                HT_HEADING, spans="1:3"),
        textRow(24, [("A", S_BODY, strings["surveyLiveIntro"])],
                wrappedHeight(SURVEY_LIVE_INTRO, 100), spans="1:3"),

        f'<row r="25" spans="1:3" ht="{wrappedHeight(SURVEY_DEMO_TEXT, 92)}"'
        f' customHeight="1" x14ac:dyDescent="0.25">'
        f'<c r="A25" s="{S_CODE}" t="s"><v>{strings["surveyDemoText"]}</v></c>'
        + pythonCell("C25", scripts["surveyDemo"], style=S_CODE,
                     argument="Claims_Data_BIG[#All]")
        + "</row>",

        f'<row r="36" spans="1:3" ht="{wrappedHeight(SURVEY_SEED_TEXT, 92)}"'
        f' customHeight="1" x14ac:dyDescent="0.25">'
        f'<c r="A36" s="{S_CODE}" t="s"><v>{strings["surveySeedText"]}</v></c>'
        + pythonCell("C36", scripts["surveySeed"], style=S_CODE,
                     argument="Claims_Data_BIG[#All]")
        + "</row>",

        textRow(48, [("A", S_NOTE, strings["cgiClosing"])],
                wrappedHeight(CGI_CLOSING, 100), spans="1:3"),
    ]
    return setDimension(appendRows(sheetXml, "".join(rows)), "A1:P48")


def editTestScaling(sheetXml: str, strings: Dict[str, int],
                    scripts: Dict[str, int]) -> str:
    """Rewrite the Test Scaling sheet. Parameters: sheetXml, strings, scripts.

    Column A is six characters wide here, so every text row on this sheet is
    merged across A:J to have somewhere to go. New rows have to be merged the
    same way, and the mergeCells count attribute kept honest -- Excel reads
    the count, not the number of children.
    """
    # Width of the merged A:J band, in characters, from the <cols> widths.
    band = 6 + 8 + 9 + 11 + 10 + 12 + 14 + 11 + 13 + 18

    def row(number: int, styleId: str, index: int, text: str,
            height: str = "") -> str:
        """One merged A:J row holding a shared string."""
        blanks = "".join(f'<c r="{columnLetter(i)}{number}" s="14"/>'
                         for i in range(2, 11))
        return (f'<row r="{number}" spans="1:10" '
                f'ht="{height or wrappedHeight(text, band)}" '
                f'customHeight="1" x14ac:dyDescent="0.25">'
                f'<c r="A{number}" s="{styleId}" t="s"><v>{index}</v></c>'
                f"{blanks}</row>")

    def liveRow(number: int, scriptIndex: int) -> str:
        """One row holding a =PY() cell that spills across the sheet."""
        return (f'<row r="{number}" spans="1:14" x14ac:dyDescent="0.25">'
                + pythonCell(f"A{number}", scriptIndex) + "</row>")

    rows = "".join([
        row(35, T_HEADING, strings["scaleHeading"], SCALE_HEADING, HT_HEADING),
        row(36, T_BODY, strings["scaleIntro"], SCALE_INTRO),
        row(37, T_NOTE, strings["scaleWhy"], SCALE_WHY),
        row(39, T_LABEL, strings["scaleLabel1"], SCALE_LABEL_1, HT_LABEL),
        row(40, S_CODE, strings["scaleSmall"], SCALE_DEMO_SMALL),
        row(42, T_LABEL, strings["scaleLabel2"], SCALE_LABEL_2, HT_LABEL),
        row(43, S_CODE, strings["scaleClaims"], SCALE_DEMO_CLAIMS),
        row(45, T_HEADING, strings["scaleLiveHeading"], SCALE_LIVE_HEADING,
            HT_HEADING),
        row(46, T_BODY, strings["scaleLiveIntro"], SCALE_LIVE_INTRO),
        row(47, T_LABEL, strings["scaleLiveLabel1"], SCALE_LIVE_LABEL_1,
            HT_LABEL),
        liveRow(48, scripts["scaleLiveSmall"]),
        row(55, T_LABEL, strings["scaleLiveLabel2"], SCALE_LIVE_LABEL_2,
            HT_LABEL),
        liveRow(56, scripts["scaleLiveClaims"]),
        row(63, T_NOTE, strings["scaleClosing"], SCALE_CLOSING),
    ])
    sheetXml = appendRows(sheetXml, rows)
    # The tab is now Test_Scaling; the title on it said Test Scaling.
    sheetXml = replaceStringCell(sheetXml, "A1", strings["scaleTitle"])

    merged = [35, 36, 37, 39, 40, 42, 43, 45, 46, 47, 55, 63]
    addition = "".join(f'<mergeCell ref="A{n}:J{n}"/>' for n in merged)
    count = int(re.search(r'<mergeCells count="(\d+)">', sheetXml).group(1))
    sheetXml = sheetXml.replace(f'<mergeCells count="{count}">',
                                f'<mergeCells count="{count + len(merged)}">', 1)
    sheetXml = sheetXml.replace("</mergeCells>", addition + "</mergeCells>", 1)
    return setDimension(sheetXml, "A1:N63")


def editWorkbook(workbookXml: str) -> str:
    """Add Claims_Data_BIG, rename Test Scaling, force a full recalculation.

    Parameters: workbookXml (str). Returns: str.

    fullCalcOnLoad matters more than it looks: the cached values in the new
    =PY() cells are placeholders, and without this Excel would show them as
    final until something else forced a recalculation.
    """
    workbookXml = workbookXml.replace(
        '<sheet name="Metadata"',
        '<sheet name="Claims_Data_BIG" sheetId="17" r:id="rId29"/>'
        '<sheet name="Metadata"', 1)
    workbookXml = workbookXml.replace('<sheet name="Test Scaling"',
                                      '<sheet name="Test_Scaling"', 1)
    return workbookXml.replace('<calcPr calcId="191029"/>',
                               '<calcPr calcId="191029" fullCalcOnLoad="1"/>', 1)


def editAppProperties(appXml: str) -> str:
    """Add the new sheet to docProps/app.xml. Parameters: appXml (str)."""
    appXml = appXml.replace("<vt:i4>14</vt:i4>", "<vt:i4>15</vt:i4>", 1)
    appXml = appXml.replace('<vt:vector size="14" baseType="lpstr">',
                            '<vt:vector size="15" baseType="lpstr">', 1)
    appXml = appXml.replace("<vt:lpstr>Metadata</vt:lpstr>",
                            "<vt:lpstr>Claims_Data_BIG</vt:lpstr>"
                            "<vt:lpstr>Metadata</vt:lpstr>", 1)
    return appXml.replace("<vt:lpstr>Test Scaling</vt:lpstr>",
                          "<vt:lpstr>Test_Scaling</vt:lpstr>", 1)


# =====================================================================
# Main
# =====================================================================

def build(source: Path, target: Path) -> Dict[str, object]:
    """Produce the new workbook.

    Parameters: source (Path) - the .xlsx to read; target (Path) - to write.
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

    bigRows = buildBigRows()

    # --- shared strings, all in one append ---
    values = {
        "title": NEW_TITLE, "subtitle": NEW_SUBTITLE,
        "backgroundHeading": "Background", "background": BACKGROUND,
        "tabs": TABS_LINE, "howTo": HOW_TO_LINE, "version": VERSION_LINE,
        "functions": FUNCTIONS_LINE, "share": SHARE_SNIPPET,
        "livePreview": LIVE_PREVIEW_LINE, "detailedCode": DETAILED_CODE_TEXT,
        "surveyHeading": "The updated __init__.py — surveySelect",
        "surveyIntro": SURVEY_INTRO, "surveySource": SURVEY_SELECT_SOURCE,
        "surveySibling": SURVEY_SIBLING,
        "surveyLiveHeading": "Live preview — surveySelect on Claims_Data_BIG",
        "surveyLiveIntro": SURVEY_LIVE_INTRO,
        "surveyDemoText": SURVEY_DEMO_TEXT,
        "surveySeedText": SURVEY_SEED_TEXT, "cgiClosing": CGI_CLOSING,
        "scaleTitle": SCALE_TITLE,
        "scaleHeading": SCALE_HEADING, "scaleIntro": SCALE_INTRO,
        "scaleWhy": SCALE_WHY, "scaleLabel1": SCALE_LABEL_1,
        "scaleSmall": SCALE_DEMO_SMALL, "scaleLabel2": SCALE_LABEL_2,
        "scaleClaims": SCALE_DEMO_CLAIMS,
        "scaleLiveHeading": SCALE_LIVE_HEADING,
        "scaleLiveIntro": SCALE_LIVE_INTRO,
        "scaleLiveLabel1": SCALE_LIVE_LABEL_1,
        "scaleLiveLabel2": SCALE_LIVE_LABEL_2,
        "scaleClosing": SCALE_CLOSING,
    }
    keys = list(values)
    sharedXml, indices = addSharedStrings(text("xl/sharedStrings.xml"),
                                          [values[k] for k in keys])
    strings = dict(zip(keys, indices))

    # --- python scripts, all in one append ---
    scriptSource = {
        "detailed": DETAILED_SCRIPT,
        "surveyDemo": SURVEY_DEMO_SCRIPT,
        "surveySeed": SURVEY_SEED_SCRIPT,
        "scaleLiveSmall": SCALE_LIVE_SMALL,
        "scaleLiveClaims": SCALE_LIVE_CLAIMS,
    }
    scriptKeys = list(scriptSource)
    pythonXml, scriptIndices = addPythonScripts(
        text("xl/python.xml"), [scriptSource[k] for k in scriptKeys])
    scripts = dict(zip(scriptKeys, scriptIndices))

    # --- the parts that change ---
    parts["xl/sharedStrings.xml"] = sharedXml.encode("utf-8")
    parts["xl/python.xml"] = pythonXml.encode("utf-8")
    parts["xl/worksheets/sheet1.xml"] = editReadme(
        text("xl/worksheets/sheet1.xml"), strings).encode("utf-8")
    parts["xl/worksheets/sheet10.xml"] = editCgiAi(
        text("xl/worksheets/sheet10.xml"), strings, scripts).encode("utf-8")
    parts["xl/worksheets/sheet12.xml"] = editTestScaling(
        text("xl/worksheets/sheet12.xml"), strings, scripts).encode("utf-8")
    parts["xl/workbook.xml"] = editWorkbook(
        text("xl/workbook.xml")).encode("utf-8")
    parts["docProps/app.xml"] = editAppProperties(
        text("docProps/app.xml")).encode("utf-8")

    # The screenshot is anchored to a row, and Background pushed that row
    # down by three. An anchor left alone would park the image on top of the
    # text it is meant to illustrate.
    drawing = text("xl/drawings/drawing1.xml")
    if "<xdr:row>11</xdr:row>" not in drawing:
        raise ValueError("the README drawing is not anchored where expected")
    parts["xl/drawings/drawing1.xml"] = drawing.replace(
        "<xdr:row>11</xdr:row>", "<xdr:row>14</xdr:row>", 1).encode("utf-8")

    # --- the new parts ---
    parts["xl/worksheets/sheet15.xml"] = buildBigSheetXml(
        bigRows).encode("utf-8")
    parts["xl/tables/table6.xml"] = buildBigTableXml(
        len(bigRows)).encode("utf-8")
    parts["xl/worksheets/_rels/sheet15.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
        '2006/relationships"><Relationship Id="rId1" Type="http://schemas.'
        'openxmlformats.org/officeDocument/2006/relationships/table" '
        'Target="../tables/table6.xml"/></Relationships>').encode("utf-8")

    rels = text("xl/_rels/workbook.xml.rels")
    rels = rels.replace("</Relationships>", (
        '<Relationship Id="rId29" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet15.xml"/></Relationships>'), 1)
    # calcChain is an optimisation, not data. Dropping it is cheaper and far
    # safer than hand-editing a dependency chain that no longer matches the
    # formulas, and Excel rebuilds it on the first recalculation.
    rels = re.sub(r'<Relationship Id="rId27"[^>]*calcChain[^>]*/>', "", rels)
    parts["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")

    types = text("[Content_Types].xml")
    types = types.replace("</Types>", (
        '<Override PartName="/xl/worksheets/sheet15.xml" ContentType='
        '"application/vnd.openxmlformats-officedocument.spreadsheetml.'
        'worksheet+xml"/>'
        '<Override PartName="/xl/tables/table6.xml" ContentType='
        '"application/vnd.openxmlformats-officedocument.spreadsheetml.'
        'table+xml"/></Types>'), 1)
    types = re.sub(r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', "", types)
    parts["[Content_Types].xml"] = types.encode("utf-8")

    parts.pop("xl/calcChain.xml", None)

    # --- write, keeping the original part order ---
    order = [name for name in order if name != "xl/calcChain.xml"]
    order += [name for name in parts if name not in order]
    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        for name in order:
            archive.writestr(name, parts[name])

    return {"Target": str(target), "Sheets": 15, "BigRows": len(bigRows),
            "BigColumns": len(BIG_COLUMNS),
            "NewStrings": len(keys), "NewScripts": len(scriptKeys),
            "ScriptIndices": scripts}


def main(argv: Sequence[str]) -> int:
    """CLI entry point. Parameters: argv. Returns: an exit code."""
    if len(argv) != 2:
        print("usage: build_claims_pipelines_workbook.py <source.xlsx> "
              "<target.xlsx>", file=sys.stderr)
        return EXIT_USAGE
    summary = build(Path(argv[0]), Path(argv[1]))
    for key, value in summary.items():
        print(f"{key}: {value}")
    return EXIT_OK


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
