"""
=====================================================================
  Program Name  : filescan_pipeline.py
  Author        : Manuel Figallo
  Purpose       : End-to-end pipeline: scan log folders for keywords,
                  render the result as a styled Excel report, and email a
                  completion notice.
  Version       : 1.0beta
  Created       : 2026-08-26
  Last Modified : 2026-08-26

  Dependencies:
    scanFileSystem (stdlib), formatCSV (openpyxl, lazy), sendEmail (stdlib).

  Description:
    Orchestration only -- all real work lives in the three functions it
    calls, per the "pipelines coordinate, they do not transform" principle.
    Email failure is reported but does NOT fail the pipeline, because the
    scan output is already on disk and is the deliverable.

  Input Parameters (required first):
    input_folder_root (REQUIRED, list[str]) - roots to scan
    extract_keyword   (REQUIRED, list[str]) - keywords to find
    output_file_path  (optional, str)       - scan output (CSV or XLSX)
    excel_output_path (optional, str)       - styled report destination
    metric_profile    (optional, str, default "sas_log")
    email_to / email_from / email_subject   - optional; email is skipped
                                              when email_to is empty
=====================================================================
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

_HERE = Path(__file__).resolve()
if str(_HERE.parent.parent.parent) not in sys.path:
    sys.path.insert(0, str(_HERE.parent.parent.parent))

from src.py.formatCSV import formatCSV                    # noqa: E402
from src.py.scanFileSystem import scanFileSystem          # noqa: E402
from src.py.sendEmail import sendEmail                    # noqa: E402
from src.utils.config import getConfig                    # noqa: E402
from src.utils.helpers import asList, timestampSuffix     # noqa: E402
from src.utils.logger import logError, logInfo, logWarn   # noqa: E402

__version__ = "1.0beta"

#: The production roots this pipeline was built for.
DEFAULT_ROOTS = [
    r"\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT\HHH\Old_Programs\Old_logs",
    r"\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT\DME\Logs",
]
DEFAULT_KEYWORDS = ["real time", "cpu time"]


def runFilescanPipeline(
    input_folder_root: Any = None,
    extract_keyword: Any = None,
    output_file_path: Optional[str] = None,
    excel_output_path: Optional[str] = None,
    metric_profile: str = "sas_log",
    email_to: Any = (),
    email_from: str = "",
    email_subject: str = "",
    format_type: str = "corporate",
) -> Dict[str, Any]:
    """Run scan -> format -> notify as one job.

    Parameters:
        input_folder_root  - roots to scan; defaults to DEFAULT_ROOTS.
        extract_keyword    - keywords; defaults to DEFAULT_KEYWORDS.
        output_file_path   - scan output; defaults under ROOT_DATA.
        excel_output_path  - styled report; defaults beside the scan output.
        metric_profile     - passed through; "sas_log" makes the scan emit
                             Excel with an extra Metrics sheet.
        email_to           - recipient(s); when empty, email is skipped.
        email_from         - sender address (required to send).
        email_subject      - subject; a sensible default is generated.
        format_type        - formatCSV style; default "corporate".
    Returns:
        dict with Scan, Report, Email and Steps (list of completed stages).
    Raises:
        ValueError - invalid scan parameters propagate from scanFileSystem.
        Email problems are logged as warnings, not raised: the scan output
        is the deliverable and it already exists by then.

    Use in claims processing:
        Nightly sweep of SAS job logs for timing and error keywords,
        delivered to the operations mailbox as a formatted workbook.
    """
    roots = asList(input_folder_root) or DEFAULT_ROOTS
    keywords = asList(extract_keyword) or DEFAULT_KEYWORDS
    steps: List[str] = []

    dataRoot = getConfig("ROOT_DATA", ".") or "."
    stamp = timestampSuffix()
    scanTarget = output_file_path or str(Path(dataRoot) / f"scan_{stamp}.csv")

    logInfo(f"filescan_pipeline {__version__} starting; {len(roots)} root(s), "
            f"{len(keywords)} keyword(s)")

    # --- 1. scan -------------------------------------------------------------
    scan = scanFileSystem(
        input_folder_root=roots,
        extract_keyword=keywords,
        output_file_path=scanTarget,
        metric_profile=metric_profile,
    )
    steps.append("scanFileSystem")
    logInfo(f"scan produced {len(scan['matches'])} match row(s) -> {scan['output']}")

    # --- 2. format -----------------------------------------------------------
    report: Optional[Dict[str, Any]] = None
    scanOutput = scan["output"]
    reportTarget = excel_output_path or str(
        Path(scanOutput).with_name(Path(scanOutput).stem + "_report.xlsx"))
    if scanOutput.lower().endswith(".csv"):
        try:
            report = formatCSV(InputCsvPath=scanOutput,
                               OutputExcelPath=reportTarget,
                               FormatType=format_type,
                               Title=f"File Scan {stamp}")
            steps.append("formatCSV")
        except ImportError as exc:
            logWarn(f"styled report skipped: {exc}")
    else:
        # metric_profile already produced Excel; styling it again would drop
        # the Metrics sheet, so keep the workbook the scan wrote.
        logInfo("scan already produced Excel (metric profile active); "
                "keeping it as the report")
        reportTarget = scanOutput

    # --- 3. notify -----------------------------------------------------------
    email: Optional[Dict[str, Any]] = None
    recipients = asList(email_to)
    if recipients:
        subject = email_subject or f"cgs_ai file scan complete - {stamp}"
        body = (f"The cgs_ai file scan has completed.\n\n"
                f"Roots scanned : {len(roots)}\n"
                f"Keywords      : {', '.join(keywords)}\n"
                f"Match rows    : {len(scan['matches'])}\n"
                f"Metric rows   : {len(scan['metrics'])}\n"
                f"Scan output   : {scanOutput}\n"
                f"Report        : {reportTarget}\n")
        try:
            email = sendEmail(To=recipients, From=email_from,
                              Subject=subject, Body=body)
            steps.append("sendEmail")
        except Exception as exc:                 # never fail the pipeline on mail
            logError(f"notification email failed (scan output is still valid): {exc}")
    else:
        logInfo("no email_to supplied; skipping notification")

    logInfo(f"pipeline complete; steps: {', '.join(steps)}")
    return {"Scan": scan, "Report": report, "Email": email, "Steps": steps}


if __name__ == "__main__":
    runFilescanPipeline()
