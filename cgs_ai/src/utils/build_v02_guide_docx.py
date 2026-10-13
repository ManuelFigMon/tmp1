"""
=====================================================================
  Program Name  : build_v02_guide_docx.py
  Author        : Manuel Figallo
  Purpose       : Generate the one-page training handout for
                  claims_report_pipeline_v0_2.py.
  Version       : 1.0beta
  Created       : 2026-08-28
  Last Modified : 2026-08-28

  Dependencies:
    STANDARD LIBRARY ONLY. Reuses the OOXML helpers from
    build_readme_docx.py, so no python-docx dependency is introduced.

  Description:
    Audience is novice Python users. Large type, one page, five numbered
    steps: make a folder, create the file, type the code, run it, check
    the output. Opens with why cgs_ai is worth their time.

  Usage:
    python src/utils/build_v02_guide_docx.py [output_path]
=====================================================================
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from build_readme_docx import (BLUE, NAVY, bullet, para,  # noqa: E402
                               writeDocx)

__version__ = "1.0beta"

SHARE = r"\\a70admed.com\R1\CGS\APPS\SAS\UNIT\SAS_G\GSIT_Prod\MANUAL\cgs_ai"

BODY_SIZE = 12          # points -- deliberately large for a handout
CODE_SIZE = 9
NOTE_SIZE = 10

#: The code the trainee types. Kept in step with the file on the share.
PIPELINE_CODE = r'''SHARE = r"\\a70admed.com\R1\CGS\...\MANUAL\cgs_ai"
CGS_AI_HOME = SHARE + r"\src\py\lite"
# STEP 1. Import our own python package
import sys
sys.path.insert(0, CGS_AI_HOME)
import cgs_ai
print('cgs_ai version', cgs_ai.__version__)

# STEP 2. Load the csv data
import csv
with open(SHARE + r"\data\synthetic_medicare_claims.csv",
          newline='', encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f))
print(len(rows), 'rows x', len(rows[0]), 'columns')

# STEP 3. Turn the csv into a formatted Excel file
result = cgs_ai.formatCSV(
    InputCsvPath=SHARE + r"\data\synthetic_medicare_claims.csv",
    OutputExcelPath=SHARE + r"\data\claims_report.xlsx",
    FormatType='corporatev2',
    SheetName='CMS Synthetic Data',
    Title='CMS Reporting and Analysis')
print(result)

# STEP 4. Send an email notification
cgs_ai.sendEmail(
    To='manuel.figallo@cgsadmin.com',
    From='manuel.figallo@cgsadmin.com',
    Subject='CMS Reporting and Analysis - workbook ready',
    Body=f"Rows: {result['RowCount']}",
    SmtpServer='smtprelay.bcbssc.com', Port=25)'''


def buildBody() -> str:
    """Assemble the handout body XML. Returns: the <w:body> content."""
    parts = []

    parts.append(para("Build Your First cgs_ai Pipeline",
                      style="Title", size=21, spaceAfter=20))
    parts.append(para("claims_report_pipeline_v0_2.py  ·  five steps  ·  "
                      "about ten minutes",
                      style="Subtitle", size=12, spaceAfter=80))

    # ---- Why this matters --------------------------------------------- #
    parts.append(para("Why cgs_ai", style="Heading1", spaceAfter=50))
    parts.append(para(
        "cgs_ai is our own Python package. It holds the work we all repeat "
        "— formatting an extract, emailing it out — as functions anyone calls.",
        size=BODY_SIZE, spaceAfter=60))
    parts.append(bullet("You write four lines instead of four hundred."))
    parts.append(bullet("Every report comes out looking the same."))
    parts.append(bullet("What you run by hand today, a scheduler can run "
                        "at 2 a.m. tomorrow."))

    # ---- The steps ----------------------------------------------------- #
    parts.append(para("Do this", style="Heading1", spaceAfter=50))

    parts.append(para("1.  Make a folder for your work, for example "
                      "C:\\code\\python.", size=BODY_SIZE, spaceAfter=50))
    parts.append(para("2.  Open Visual Studio Code, then File > New File.",
                      size=BODY_SIZE, spaceAfter=50))
    parts.append(para("3.  Save it into that folder as "
                      "claims_report_pipeline_v0_2.py — the .py matters.",
                      size=BODY_SIZE, spaceAfter=50))
    parts.append(para("4.  Type this in, then save:",
                      size=BODY_SIZE, spaceAfter=50))
    parts.append(para(PIPELINE_CODE, style="Code", size=CODE_SIZE, mono=True,
                      spaceAfter=40))
    parts.append(para("5.  Press the Run button, or press F5.",
                      size=BODY_SIZE, spaceAfter=50))

    # ---- What success looks like --------------------------------------- #
    parts.append(para("You should see", style="Heading3", spaceAfter=40))
    parts.append(para(
        "cgs_ai version 1.0beta-lite\n"
        "100 rows x 10 columns\n"
        "{'OutputExcelPath': '...claims_report.xlsx', 'RowCount': 100 ...}",
        style="Code", size=CODE_SIZE, mono=True, spaceAfter=60))
    parts.append(para(
        "Open claims_report.xlsx: navy banner, blue headers, striped rows "
        "— plus an email in your inbox.",
        size=NOTE_SIZE, color=BLUE, spaceAfter=40))

    parts.append(para(
        "Now try it: change FormatType to 'plain' and run it again.",
        size=NOTE_SIZE, color=NAVY, spaceAfter=0))

    return "".join(parts)


if __name__ == "__main__":
    output = sys.argv[1] if len(sys.argv) > 1 else str(
        Path(__file__).resolve().parent.parent.parent
        / "docs" / "cgs_ai_First_Pipeline_Guide.docx")
    print("wrote", writeDocx(buildBody(), output, margin=720))
