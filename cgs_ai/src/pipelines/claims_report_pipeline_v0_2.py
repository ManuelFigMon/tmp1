"""
claims_report_pipeline_v0_2.py -- the four steps, using our own package.

Same as v0_1, except cgs_ai is imported the way any Python package is
imported, instead of being loaded from a file. Needs openpyxl.
"""

# The folder that CONTAINS the cgs_ai package folder.
CGS_AI_HOME = r"\\a70admed.com\R1\CGS\APPS\SAS\UNIT\SAS_G\GSIT_Prod\MANUAL"
SHARE = CGS_AI_HOME + r"\cgs_ai"


# STEP 1. Import the python package
# Telling Python where to look, then importing cgs_ai exactly like pandas.
import sys
sys.path.insert(0, CGS_AI_HOME)
import cgs_ai

functions = [name for name in cgs_ai.__all__ if callable(getattr(cgs_ai, name))]
print('cgs_ai version', cgs_ai.__version__)
print('loaded from   ', cgs_ai.__file__)
print('functions     ', len(functions))
print(functions)


# STEP 2. Load the csv data into python
# csv is built into Python, so there is nothing to install.
import csv
with open(SHARE + r"\data\synthetic_medicare_claims.csv", newline='',
          encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f))
print(len(rows), 'rows x', len(rows[0]), 'columns')


# STEP 3. Update the csv into an Excel file with ODS like formatting
result = cgs_ai.formatCSV(
    InputCsvPath=SHARE + r"\data\synthetic_medicare_claims.csv",
    OutputExcelPath=SHARE + r"\data\synthetic_medicare_claims_corporate.xlsx",
    FormatType='corporatev2',
    SheetName='CMS Synthetic Data',
    Title='CMS Reporting and Analysis'
)
print(result)


# STEP 4. Send an email notification
cgs_ai.sendEmail(
    To='manuel.figallo@cgsadmin.com',
    From='manuel.figallo@cgsadmin.com',
    Subject='CMS Reporting and Analysis - workbook ready',
    Body=f"Rows: {result['RowCount']}\nLocation: {result['OutputExcelPath']}",
    SmtpServer='smtprelay.bcbssc.com',
    Port=25
)
