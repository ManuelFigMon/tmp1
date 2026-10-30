"""
=====================================================================
  Program Name  : logger.py
  Author        : Manuel Figallo
  Purpose       : One consistent stderr logger for every cgs_ai function,
                  matching the scanFileSystem log format so Python,
                  PowerShell and SAS output all read alike.
  Version       : 1.0beta
  Created       : 2026-08-26
  Last Modified : 2026-08-26

  Description:
    Writes "YYYY-MM-DD HH:MM:SS LEVEL message" to stderr. stdout is left
    clean so a function's real output can be piped. Never prompts.
=====================================================================
"""

from __future__ import annotations

import datetime as _dt
import sys

__version__ = "1.0beta"


def writeLog(level: str, message: str) -> None:
    """Write one timestamped log line to stderr.

    Parameters:
        level (str)   - INFO, WARNING or ERROR.
        message (str) - text to log.
    Returns: None.
    """
    stamp = _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"{stamp} {level:<7} {message}", file=sys.stderr, flush=True)


def logInfo(message: str) -> None:
    """Log an INFO line. Parameters: message (str). Returns: None."""
    writeLog("INFO", message)


def logWarn(message: str) -> None:
    """Log a WARNING line. Parameters: message (str). Returns: None."""
    writeLog("WARNING", message)


def logError(message: str) -> None:
    """Log an ERROR line. Parameters: message (str). Returns: None."""
    writeLog("ERROR", message)
