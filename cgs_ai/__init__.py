"""
=====================================================================
  Program Name  : __init__.py  (cgs_ai package root)
  Author        : Manuel Figallo
  Purpose       : Single import surface for every cgs_ai Python function.
                  Import this package and get the whole toolkit.
  Version       : 1.0beta
  Created       : 2026-08-26
  Last Modified : 2026-08-26

  Dependencies:
    Standard library only to import. Individual functions import their own
    optional third-party packages LAZILY (openpyxl for Excel, pandas for
    SAS conversion, pyodbc for SQL Server), so importing cgs_ai never fails
    because an optional package is missing.

  Description:
    Every public function follows one convention across Python, PowerShell
    and SAS: mixedCase functionNames, identical parameter names, a header
    comment block, and a docstring stating Parameters, Returns, Raises and
    the claims-processing use.

  Usage:
      %run cgs_ai_setup            # notebook one-liner
      import cgs_ai
      cgs_ai.scanFileSystem(input_folder_root=..., extract_keyword=...)

  Function Index:
    basic_hello              - standard greeting (smoke test)
    personalized_hello       - greeting with input protection
    detailed_hello           - structured greeting dictionary
    scanFileSystem           - keyword scan of directory roots; one row/match
    get_comments             - retrieve public comments from Regulations.gov
    runSQLServerQuery        - SQL Server query via Integrated Security
    formatCSV                - CSV to styled Excel (SAS ODS look and feel)
    downloadBulkFiles        - bulk HTTP download from a CSV link column
    sendEmail                - SMTP alert to one or many recipients
    convertSAS2Pandas        - sas7bdat to pandas DataFrame
    copyExcelSheet2CSV       - Excel worksheet to CSV, with validation
    collectSystemMetrics     - host metrics to a CSV time series
    zipFolder                - folder + extra files into one archive
    runFilescanPipeline      - scan -> format -> email, end to end

  Change Log:
    v1.0beta - First consolidated package release.
=====================================================================
"""

from __future__ import annotations

import sys as _sys
from pathlib import Path as _Path

__version__ = "1.0beta"

# Make `src.*` importable whether cgs_ai is imported as a package or the
# project root is the working directory.
_ROOT = _Path(__file__).resolve().parent
if str(_ROOT) not in _sys.path:
    _sys.path.insert(0, str(_ROOT))


# =====================================================================
# Greeting functions -- smoke tests that prove the import worked
# =====================================================================

def basic_hello() -> str:
    """Version 1: Standard greeting."""
    return "Hello, World!"


def personalized_hello(name: str) -> str:
    """Version 2: Personalized greeting with input protection."""
    clean_name = str(name).strip() if name else "World"
    return f"Hello, {clean_name}!"


def detailed_hello(style: str = "friendly") -> dict:
    """Version 3: Returns a structured JSON-like dictionary."""
    styles = {
        "friendly": "Hello, World! Wonderful day, isn't it?",
        "formal": "Greetings, World. It is a pleasure to connect.",
        "pirate": "Ahoy, World! Avast ye scallywags!",
    }
    greeting = styles.get(style.lower(), styles["friendly"])
    return {"message": greeting, "style_used": style}


# =====================================================================
# Toolkit functions -- re-exported from src/py so callers need one import
# =====================================================================

from src.py.collectSystemMetrics import collectSystemMetrics    # noqa: E402
from src.py.convertSAS2Pandas import convertSAS2Pandas          # noqa: E402
from src.py.copyExcelSheet2CSV import copyExcelSheet2CSV        # noqa: E402
from src.py.downloadBulkFiles import downloadBulkFiles          # noqa: E402
from src.py.formatCSV import formatCSV                          # noqa: E402
from src.py.runSQLServerQuery import runSQLServerQuery          # noqa: E402
from src.py.scanFileSystem import scanFileSystem                # noqa: E402
from src.py.sendEmail import sendEmail                          # noqa: E402
from src.py.zipFolder import zipFolder                          # noqa: E402
# Regulations.gov retrieval (the package's first module, kept importable).
from src.py import regulations                                  # noqa: E402
from src.py.regulations import (build_metadata, get_comments,   # noqa: E402
                                write_metadata, write_output)
from src.pipelines.filescan_pipeline import runFilescanPipeline  # noqa: E402

# Shared utilities, exposed for callers who need them directly.
from src.utils.config import getConfig, loadConfig               # noqa: E402
from src.utils.logger import logError, logInfo, logWarn          # noqa: E402

__all__ = [
    "__version__",
    # greetings
    "basic_hello", "personalized_hello", "detailed_hello",
    # toolkit
    "scanFileSystem", "runSQLServerQuery", "formatCSV", "downloadBulkFiles",
    "sendEmail", "convertSAS2Pandas", "copyExcelSheet2CSV",
    "collectSystemMetrics", "zipFolder", "runFilescanPipeline",
    # regulations.gov
    "regulations", "get_comments", "write_output", "build_metadata",
    "write_metadata",
    # utilities
    "getConfig", "loadConfig", "logInfo", "logWarn", "logError",
]
