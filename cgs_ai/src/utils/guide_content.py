"""
=====================================================================
  Program Name  : guide_content.py
  Author        : Manuel Figallo
  Purpose       : Single source of the training content for
                  claims_report_pipeline_v0_2.py -- the code a trainee
                  types, and the prose that explains it.
  Version       : 1.0beta
  Created       : 2026-08-28
  Last Modified : 2026-08-28

  Dependencies:
    STANDARD LIBRARY ONLY.

  Description:
    The handouts and the slide deck all import from here, so the code on
    a slide is the same code on the page. A check executes ASSEMBLED_CODE
    verbatim, so none of these documents can describe something that does
    not run.

  Usage:
      from guide_content import CODE_STEPS, ASSEMBLED_CODE, CONTACT
=====================================================================
"""

from __future__ import annotations

from typing import Dict, List

__version__ = "1.0beta"

SHARE_FULL = r"\\a70admed.com\R1\CGS\APPS\SAS\UNIT\SAS_G\GSIT_Prod\MANUAL\cgs_ai"
SHARE_SHORT = r"\\a70admed.com\R1\CGS\...\MANUAL\cgs_ai"

CONTACT = "Manuel Figallo"
FILENAME = "claims_report_pipeline_v0_2.py"

#: The two lines every step depends on.
CODE_HEADER = (f'SHARE = r"{SHARE_SHORT}"\n'
               'CGS_AI_HOME = SHARE + r"\\src\\py\\lite"')

#: One entry per step: the code, and why it is there.
CODE_STEPS: List[Dict[str, str]] = [
    {
        "number": "1",
        "title": "Import our own python package",
        "code": ("import sys\n"
                 "sys.path.insert(0, CGS_AI_HOME)\n"
                 "import cgs_ai\n"
                 "print('cgs_ai version', cgs_ai.__version__)"),
        "why": ("Line one tells Python where our package lives; line two "
                "imports it, exactly the way you would import pandas."),
    },
    {
        "number": "2",
        "title": "Load the csv data",
        "code": ('import csv\n'
                 'with open(SHARE + r"\\data\\synthetic_medicare_claims.csv",\n'
                 "          newline='', encoding='utf-8-sig') as f:\n"
                 "    rows = list(csv.DictReader(f))\n"
                 "print(len(rows), 'rows x', len(rows[0]), 'columns')"),
        "why": ("csv is built into Python — nothing to install. Printing "
                "the size confirms you are pointed at the right file."),
    },
    {
        "number": "3",
        "title": "Turn the csv into a formatted Excel file",
        "code": ("result = cgs_ai.formatCSV(\n"
                 '    InputCsvPath=SHARE + r"\\data\\synthetic_medicare_claims.csv",\n'
                 '    OutputExcelPath=SHARE + r"\\data\\claims_report.xlsx",\n'
                 "    FormatType='corporatev2',\n"
                 "    SheetName='CMS Synthetic Data',\n"
                 "    Title='CMS Reporting and Analysis')\n"
                 "print(result)"),
        "why": ("One call replaces all the formatting you would otherwise "
                "do by hand, and hands back a result you can check."),
    },
    {
        "number": "4",
        "title": "Send an email notification",
        "code": ("cgs_ai.sendEmail(\n"
                 "    To='manuel.figallo@cgsadmin.com',\n"
                 "    From='manuel.figallo@cgsadmin.com',\n"
                 "    Subject='CMS Reporting and Analysis - workbook ready',\n"
                 '    Body=f"Rows: {result[\'RowCount\']}",\n'
                 "    SmtpServer='smtprelay.bcbssc.com', Port=25)"),
        "why": ("The row count comes from step 3's result, so the message "
                "always reports what was actually written."),
    },
]

#: The whole program, exactly as a trainee would type it.
ASSEMBLED_CODE = CODE_HEADER + "\n" + "\n\n".join(
    f"# STEP {step['number']}. {step['title']}\n{step['code']}"
    for step in CODE_STEPS)

#: What the console prints when it works.
EXPECTED_OUTPUT = ("cgs_ai version 1.0beta-lite\n"
                   "100 rows x 10 columns\n"
                   "{'OutputExcelPath': '...claims_report.xlsx', "
                   "'RowCount': 100 ...}")

#: Why the package is worth their time.
VALUE_POINTS = [
    ("You write four lines instead of four hundred.",
     "The formatting, the styling and the mail all live inside cgs_ai. "
     "You call them; you do not rewrite them."),
    ("Every report comes out looking the same.",
     "One implementation means one look, whoever runs it and whenever."),
    ("Change a parameter, not the code.",
     "A different extract, sheet name or title is an argument you pass, "
     "not a program you edit."),
    ("What you run by hand today, a scheduler runs at 2 a.m. tomorrow.",
     "Nothing in these functions prompts for input, so the same call works "
     "unattended."),
]

#: Setup, before any code is typed.
SETUP_STEPS = [
    ("Make a folder", "Somewhere easy to find, for example C:\\code\\python."),
    ("Open Visual Studio Code", "Then choose File > New File."),
    (f"Save it as {FILENAME}",
     "Into the folder you just made. The .py on the end is what makes it a "
     "Python program."),
]

#: Common first-run failures.
TROUBLESHOOTING = [
    ("ModuleNotFoundError: No module named 'cgs_ai'",
     "CGS_AI_HOME is pointing at the wrong folder, or the share is not "
     "mapped on this machine."),
    ("formatCSV requires openpyxl",
     "Run pip install openpyxl once, then run the program again."),
    ("Connection refused on port 25",
     "The mail relay is not reachable from your machine. Comment out step 4 "
     "and check with IT."),
]
