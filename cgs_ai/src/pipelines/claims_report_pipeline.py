"""
=====================================================================
  Program Name  : claims_report_pipeline.py
  Author        : Manuel Figallo
  Purpose       : End-user pipeline. Load a Medicare claims CSV, publish it
                  as a formatted Excel workbook, and email the notification.
  Version       : 1.0beta
  Created       : 2026-08-28
  Last Modified : 2026-08-28

  Dependencies:
    cgs_ai LITE (src/py/lite/__init__.py) -- loaded from the share at run
    time, no installation. openpyxl is required by formatCSV. pandas is
    used for the DataFrame if present; if it is not, the pipeline falls
    back to the standard-library csv module and still runs.

  Description:
    Four steps, in the order an end user thinks about them:
      STEP 1  import the cgs_ai lite package from the network share
      STEP 2  load the claims CSV into a DataFrame
      STEP 3  write a corporate-styled Excel workbook
      STEP 4  email the notification

    Every path and address is a constant at the top of this file. To point
    the pipeline at a different extract, change CLAIMS_CSV and nothing else.

  Usage:
      python claims_report_pipeline.py
      python claims_report_pipeline.py --no-email      (steps 1-3 only)
=====================================================================
"""

from __future__ import annotations

import runpy
import sys
from pathlib import Path
from typing import Any, Dict

__version__ = "1.0beta"

# --------------------------------------------------------------------- #
# Configuration -- everything an end user needs to change lives here.
# --------------------------------------------------------------------- #

SHARE = r"\\a70admed.com\R1\CGS\APPS\SAS\UNIT\SAS_G\GSIT_Prod\MANUAL\cgs_ai"

LITE_PACKAGE = rf"{SHARE}\src\py\lite\__init__.py"
CLAIMS_CSV = rf"{SHARE}\data\synthetic_medicare_claims.csv"
OUTPUT_XLSX = rf"{SHARE}\data\synthetic_medicare_claims_corporate.xlsx"

SHEET_NAME = "CMS Synthetic Data"
REPORT_TITLE = "CMS Reporting and Analysis"

EMAIL_TO = "manuel.figallo@cgsadmin.com"
EMAIL_FROM = "manuel.figallo@cgsadmin.com"
SMTP_SERVER = "smtprelay.bcbssc.com"   # SMTP relay hostname or IP
SMTP_PORT = 25                         # 25 = relay, 587 = submission


# --------------------------------------------------------------------- #
# STEP 1. Import the python package
# --------------------------------------------------------------------- #

def loadPackage(packagePath: str = LITE_PACKAGE) -> Any:
    """Load the cgs_ai lite package straight off the network share.

    Parameters:
        packagePath (str) - full path to src/py/lite/__init__.py.
    Returns: the imported cgs_ai module, exposing formatCSV and sendEmail.
    Raises: FileNotFoundError if the share is unreachable or the path is wrong.

    NOTE: runpy.run_path() executes the file but does not by itself create an
    importable module. The lite __init__.py registers itself in sys.modules
    as "cgs_ai" as its last action, which is what makes the import below work.
    """
    if not Path(packagePath).is_file():
        raise FileNotFoundError(
            f"cgs_ai lite package not found at {packagePath}. Check that the "
            f"share is mapped and reachable from this machine.")

    runpy.run_path(packagePath, run_name="__init__")

    import cgs_ai  # noqa: E402  -- registered by the line above
    print(f"STEP 1  cgs_ai loaded, version {cgs_ai.__version__}")
    return cgs_ai


# --------------------------------------------------------------------- #
# STEP 2. Load the CSV into a dataframe
# --------------------------------------------------------------------- #

def loadClaims(csvPath: str = CLAIMS_CSV) -> Any:
    """Read the claims extract into a DataFrame.

    Parameters:
        csvPath (str) - the claims CSV on the share.
    Returns:
        a pandas DataFrame when pandas is available, otherwise a list of row
        dicts. Either way the row and column counts are printed.
    Raises: FileNotFoundError if the CSV is not on the share.
    """
    if not Path(csvPath).is_file():
        raise FileNotFoundError(f"claims CSV not found at {csvPath}")

    try:
        import pandas as pd
        frame = pd.read_csv(csvPath)
        rowCount, columnCount = frame.shape
    except ImportError:
        # pandas is not installed on every desktop; the pipeline still runs.
        import csv
        with open(csvPath, encoding="utf-8-sig", newline="") as handle:
            frame = list(csv.DictReader(handle))
        rowCount = len(frame)
        columnCount = len(frame[0]) if frame else 0
        print("STEP 2  pandas not installed, using the csv module instead")

    print(f"STEP 2  loaded {rowCount} row(s) x {columnCount} column(s) "
          f"from {Path(csvPath).name}")
    return frame


# --------------------------------------------------------------------- #
# STEP 3. Write the styled Excel workbook
# --------------------------------------------------------------------- #

def publishWorkbook(cgs_ai: Any, csvPath: str = CLAIMS_CSV,
                    outputPath: str = OUTPUT_XLSX) -> Dict[str, Any]:
    """Convert the claims CSV into an ODS-styled Excel workbook.

    Parameters:
        cgs_ai (module)   - the package returned by loadPackage().
        csvPath (str)     - source CSV.
        outputPath (str)  - destination .xlsx.
    Returns: formatCSV's result dict -- OutputExcelPath, RowCount,
             ColumnCount, FormatType.
    Raises: ImportError if openpyxl is missing; OSError if the share is
            read-only.
    """
    result = cgs_ai.formatCSV(
        InputCsvPath=csvPath,
        OutputExcelPath=outputPath,
        FormatType="corporate",
        SheetName=SHEET_NAME,
        Title=REPORT_TITLE,
    )
    print(f"STEP 3  {result}")
    return result


# --------------------------------------------------------------------- #
# STEP 4. Send the notification
# --------------------------------------------------------------------- #

def notify(cgs_ai: Any, result: Dict[str, Any]) -> Dict[str, Any]:
    """Email the notification that the workbook is ready.

    Parameters:
        cgs_ai (module)  - the package returned by loadPackage().
        result (dict)    - formatCSV's result, quoted in the body.
    Returns: sendEmail's result dict.
    Raises: OSError or smtplib.SMTPException if the relay is unreachable or
            rejects the message.
    """
    body = (
        f"{REPORT_TITLE}\n\n"
        f"The claims workbook is ready.\n\n"
        f"  Rows      : {result['RowCount']}\n"
        f"  Columns   : {result['ColumnCount']}\n"
        f"  Format    : {result['FormatType']}\n"
        f"  Worksheet : {SHEET_NAME}\n"
        f"  Location  : {result['OutputExcelPath']}\n\n"
        f"Generated by cgs_ai.\n")

    sent = cgs_ai.sendEmail(
        To=EMAIL_TO,
        From=EMAIL_FROM,
        Subject=f"{REPORT_TITLE} - {result['RowCount']} rows ready",
        Body=body,
        SmtpServer=SMTP_SERVER,
        Port=SMTP_PORT,
    )
    print(f"STEP 4  {sent}")
    return sent


# --------------------------------------------------------------------- #
# Orchestration
# --------------------------------------------------------------------- #

def runClaimsReportPipeline(sendNotification: bool = True) -> Dict[str, Any]:
    """Run all four steps end to end.

    Parameters:
        sendNotification (bool) - False stops after the workbook is written,
                                  which is what you want when testing off
                                  the network.
    Returns: dict with RowCount, ColumnCount, OutputExcelPath and Notified.
    """
    cgs_ai = loadPackage()                       # STEP 1
    loadClaims()                                 # STEP 2
    result = publishWorkbook(cgs_ai)             # STEP 3
    notified = False
    if sendNotification:
        notify(cgs_ai, result)                   # STEP 4
        notified = True
    else:
        print("STEP 4  skipped (--no-email)")

    print(f"\nDone. Workbook: {result['OutputExcelPath']}")
    return {"RowCount": result["RowCount"],
            "ColumnCount": result["ColumnCount"],
            "OutputExcelPath": result["OutputExcelPath"],
            "Notified": notified}


if __name__ == "__main__":
    runClaimsReportPipeline(sendNotification="--no-email" not in sys.argv)
