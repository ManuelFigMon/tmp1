"""
claims_report_pipeline_v0_1.py -- the four steps, nothing else.

Load a Medicare claims CSV, publish it as a formatted Excel workbook, and
email the notification. Needs openpyxl and access to the share.
"""

SHARE = r"\\a70admed.com\R1\CGS\APPS\SAS\UNIT\SAS_G\GSIT_Prod\MANUAL\cgs_ai"


# STEP 1. Import the python package
import runpy
runpy.run_path(SHARE + r"\src\py\lite\cgs_ai\__init__.py", run_name="__init__")
from cgs_ai import formatCSV, sendEmail


# STEP 2. Load the csv data into python
# csv is built into Python, so there is nothing to install.
import csv
with open(SHARE + r"\data\synthetic_medicare_claims.csv", newline='',
          encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f))
print(len(rows), 'rows x', len(rows[0]), 'columns')


# STEP 3. Update the csv into an Excel file with ODS like formatting
result = formatCSV(
    InputCsvPath=SHARE + r"\data\synthetic_medicare_claims.csv",
    OutputExcelPath=SHARE + r"\data\synthetic_medicare_claims_corporate.xlsx",
    FormatType='corporatev2',
    SheetName='CMS Synthetic Data',
    Title='CMS Reporting and Analysis'
)
print(result)


# STEP 4. Send an email notification
sendEmail(
    To='manuel.figallo@cgsadmin.com',
    From='manuel.figallo@cgsadmin.com',
    Subject='CMS Reporting and Analysis - workbook ready',
    Body=f"Rows: {result['RowCount']}\nLocation: {result['OutputExcelPath']}",
    SmtpServer='smtprelay.bcbssc.com',
    Port=25
)
