"""
=====================================================================
  Program Name  : build_onepager_docx.py
  Author        : Manuel Figallo
  Purpose       : Generate the end-user one-pager for the cgs_ai lite
                  claims reporting pipeline.
  Version       : 1.0beta
  Created       : 2026-08-28
  Last Modified : 2026-08-28

  Dependencies:
    STANDARD LIBRARY ONLY. Reuses the OOXML helpers from
    build_readme_docx.py, so no python-docx dependency is introduced.

  Description:
    One page: what cgs_ai is and why it is worth using at the top, then the
    four steps an end user runs, then troubleshooting.

  Usage:
    python src/utils/build_onepager_docx.py [output_path]
=====================================================================
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from build_readme_docx import (BLUE, NAVY, bullet, para,  # noqa: E402
                               table, writeDocx)

__version__ = "1.0beta"

SHARE = r"\\a70admed.com\R1\CGS\APPS\SAS\UNIT\SAS_G\GSIT_Prod\MANUAL\cgs_ai"


def buildBody() -> str:
    """Assemble the one-pager body XML. Returns: the <w:body> content."""
    parts = []

    parts.append(para("cgs_ai — Claims Reporting in Four Steps",
                      style="Title", size=20, spaceAfter=20))
    parts.append(para("Lite build v1.0beta  ·  formatCSV + sendEmail  ·  "
                      "no installation required",
                      style="Subtitle", spaceAfter=90))

    # ---- What it is and why it matters -------------------------------- #
    parts.append(para("What cgs_ai is", style="Heading1", spaceAfter=50))
    parts.append(para(
        "cgs_ai is a shared library of reusable functions for routine data work "
        "— formatting extracts into reports, sending notifications, scanning "
        "file shares. Each does one job and is driven by parameters, not by "
        "editing code, the same way a SAS PROC is. The lite build carries just "
        "the two functions claims reporting needs.", spaceAfter=70))

    parts.append(para("Why it is worth using", style="Heading2", spaceAfter=50))
    parts.append(bullet("Write once, use everywhere — one implementation, "
                        "not many near-identical scripts."))
    parts.append(bullet("Change a parameter, not the code."))
    parts.append(bullet("Consistent output — no hand-formatting in Excel."))
    parts.append(bullet("Ready to be scheduled; nothing prompts for input."))

    # ---- The four steps ----------------------------------------------- #
    parts.append(para("The pipeline", style="Heading1", spaceAfter=50))
    parts.append(para(
        "Needs Python 3.9 or later, openpyxl (pip install openpyxl), and "
        "access to the share. Nothing else.", spaceAfter=80))

    parts.append(para("STEP 1.  Import the package",
                      style="Heading3", spaceAfter=30))
    parts.append(para(
        "import runpy\n"
        f'SHARE = r"{SHARE}"\n'
        'runpy.run_path(SHARE + r"\\src\\py\\lite\\cgs_ai\\__init__.py", '
        'run_name="__init__")\n'
        "from cgs_ai import formatCSV, sendEmail",
        style="Code", size=8, mono=True, spaceAfter=50))
    parts.append(para(
        "run_path executes the file; the file registers itself as cgs_ai, "
        "which is what makes the last line work.",
        size=8, color=BLUE, spaceAfter=60))

    parts.append(para("STEP 2.  Load the claims data",
                      style="Heading3", spaceAfter=30))
    parts.append(para(
        "import pandas as pd\n"
        'CSV = SHARE + r"\\data\\synthetic_medicare_claims.csv"\n'
        "df = pd.read_csv(CSV)\n"
        "print(df.shape)",
        style="Code", size=8, mono=True, spaceAfter=80))

    parts.append(para("STEP 3.  Publish a formatted Excel workbook",
                      style="Heading3", spaceAfter=30))
    parts.append(para(
        "result = formatCSV(\n"
        "    InputCsvPath=CSV,\n"
        '    OutputExcelPath=SHARE + r"\\data\\'
        'synthetic_medicare_claims_corporate.xlsx",\n'
        '    FormatType="corporate", SheetName="CMS Synthetic Data",\n'
        '    Title="CMS Reporting and Analysis")\n'
        "print(result)",
        style="Code", size=8, mono=True, spaceAfter=50))
    parts.append(para(
        "Navy banner, blue header, zebra striping, frozen auto-filtered header "
        "— the SAS ODS look. Returns OutputExcelPath, RowCount, ColumnCount.",
        size=8, color=BLUE, spaceAfter=60))

    parts.append(para("STEP 4.  Send the notification",
                      style="Heading3", spaceAfter=30))
    parts.append(para(
        'sendEmail(To="manuel.figallo@cgsadmin.com", '
        'From="manuel.figallo@cgsadmin.com",\n'
        '    Subject="CMS Reporting and Analysis - workbook ready",\n'
        '    Body=f"Rows: {result[\'RowCount\']}  |  '
        '{result[\'OutputExcelPath\']}",\n'
        '    SmtpServer="smtprelay.bcbssc.com", Port=25)',
        style="Code", size=8, mono=True, spaceAfter=80))

    # ---- Run it all at once ------------------------------------------- #
    parts.append(para("Or run all four steps at once",
                      style="Heading3", spaceAfter=30))
    parts.append(para(
        "python %SHARE%\\src\\pipelines\\claims_report_pipeline.py"
        "            (add --no-email to stop after step 3)",
        style="Code", size=8, mono=True, spaceAfter=50))
    parts.append(para(
        "Every path and address is a constant at the top of that file — to "
        "report on a different extract, change CLAIMS_CSV.",
        size=8, color=BLUE, spaceAfter=60))

    # ---- Troubleshooting ---------------------------------------------- #
    parts.append(para("If something goes wrong", style="Heading3",
                      spaceAfter=30))
    parts.append(para(
        '"cgs_ai lite package not found" — the share is not reachable; map '
        "the drive and re-run.\n"
        '"formatCSV requires openpyxl" — run pip install openpyxl.\n'
        '"Connection refused" on port 25 — the relay is unreachable from this '
        "machine; re-run with --no-email and check with IT.",
        size=8, spaceAfter=70))

    parts.append(para(
        "Questions, or a report you would like added: Manuel Figallo.   "
        "cgs_ai is in beta — feedback on names and parameters is welcome now.",
        size=8, color=NAVY, spaceAfter=0))

    return "".join(parts)


if __name__ == "__main__":
    output = sys.argv[1] if len(sys.argv) > 1 else str(
        Path(__file__).resolve().parent.parent.parent
        / "docs" / "cgs_ai_Lite_Pipeline_OnePager.docx")
    print("wrote", writeDocx(buildBody(), output, margin=720))
