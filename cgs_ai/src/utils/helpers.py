"""
=====================================================================
  Program Name  : helpers.py
  Author        : Manuel Figallo
  Purpose       : Small shared helpers used by more than one cgs_ai
                  function -- list parsing, CSV I/O and timestamps.
  Version       : 1.0beta
  Created       : 2026-08-26
  Last Modified : 2026-08-26

  Description:
    Standard library only. The SAS wrappers pass list parameters as ONE
    semicolon-delimited string, so asList() accepts that form as well as a
    real Python list, keeping the three languages interchangeable.
=====================================================================
"""

from __future__ import annotations

import csv
import datetime as _dt
import os
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

__version__ = "1.0beta"

TIMESTAMP_FORMAT = "%Y%m%d_%H%M%S"


def asList(value: Any) -> List[str]:
    """Normalize a list parameter to a list of strings.

    Parameters:
        value - None, a string ("a;b" or "a"), or an iterable of strings.
    Returns:
        list[str] with blanks removed. Semicolons split because that is how
        the SAS wrappers pass lists on a command line.
    """
    if value is None:
        return []
    if isinstance(value, str):
        return [part.strip() for part in value.split(";") if part.strip()]
    out: List[str] = []
    for item in value:
        out.extend(asList(item) if isinstance(item, str) else [str(item)])
    return out


def isoNow() -> str:
    """Current local time as ISO-8601 to the second. Parameters: none."""
    return _dt.datetime.now().isoformat(timespec="seconds")


def timestampSuffix() -> str:
    """Current time as YYYYMMDD_HHMMSS, for auto-generated filenames."""
    return _dt.datetime.now().strftime(TIMESTAMP_FORMAT)


def ensureParent(path: str) -> Path:
    """Create the parent directory of `path` if needed.

    Parameters: path (str) - a file path.
    Returns: Path object for the file.
    """
    target = Path(path)
    if target.parent and not target.parent.exists():
        target.parent.mkdir(parents=True, exist_ok=True)
    return target


def writeCsv(rows: Sequence[Dict[str, Any]], columns: Sequence[str], outputPath: str) -> str:
    """Write rows to CSV using only the standard library.

    Parameters:
        rows (sequence of dict) - the data.
        columns (sequence of str) - column order; missing keys become ''.
        outputPath (str) - destination file.
    Returns: the path written.
    """
    target = ensureParent(outputPath)
    with open(target, "w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(columns), extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow({column: row.get(column, "") for column in columns})
    return str(target)


def readCsv(inputPath: str) -> List[Dict[str, str]]:
    """Read a CSV into a list of dicts. Parameters: inputPath (str)."""
    with open(inputPath, encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def fileTimes(path: str) -> Dict[str, str]:
    """Return created/modified/accessed times as ISO-8601 strings.

    Parameters: path (str) - file to stat.
    Returns: dict with created_time, modified_time, accessed_time and
             file_size_bytes. Raises OSError if the file cannot be stat'ed.
    """
    info = os.stat(path)
    fmt = lambda ts: _dt.datetime.fromtimestamp(ts).isoformat(timespec="seconds")
    return {
        "created_time": fmt(getattr(info, "st_ctime", 0)),
        "modified_time": fmt(info.st_mtime),
        "accessed_time": fmt(info.st_atime),
        "file_size_bytes": info.st_size,
    }
