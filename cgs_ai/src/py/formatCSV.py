"""
=====================================================================
  Program Name  : formatCSV.py
  Author        : Manuel Figallo
  Purpose       : Backward-compatible shim. formatCSV was renamed to
                  formatData when it learned to read .xlsx as well as
                  .csv; this keeps the old import path working.
  Version       : 1.0beta
  Created       : 2026-08-26
  Last Modified : 2026-09-06

  Dependencies:
    None of its own -- everything comes from src/py/formatData.py.

  Description:
    The files in src/py and src/ps are copied to a share individually, so
    deleting this one would break any caller copied before the rename with
    "No module named formatCSV" rather than anything useful. Both names
    resolve to the same implementation.

  Input Parameters (required first):
    As formatData. InputCsvPath is the original name for InputPath.

  Usage:
      from src.py.formatCSV import formatCSV    # old, still works
      from src.py.formatData import formatData  # preferred
=====================================================================
"""

from __future__ import annotations

import sys
from pathlib import Path

_HERE = Path(__file__).resolve()
if str(_HERE.parent.parent.parent) not in sys.path:
    sys.path.insert(0, str(_HERE.parent.parent.parent))

from src.py.formatData import (DEFAULT_FORMAT_TYPE, FORMAT_STYLES,  # noqa: E402,F401
                               MAX_COLUMN_WIDTH, STRIPED_FORMATS,
                               formatCSV, formatData, readTable)

__version__ = "1.0beta"

__all__ = ["formatCSV", "formatData", "readTable", "FORMAT_STYLES",
           "STRIPED_FORMATS", "DEFAULT_FORMAT_TYPE", "MAX_COLUMN_WIDTH"]
