"""
=====================================================================
  Program Name  : downloadBulkFiles.py
  Author        : Manuel Figallo
  Purpose       : Download every attachment referenced by a CSV column of
                  HTTP links, such as regulations.gov comment attachments.
  Version       : 1.0beta
  Created       : 2026-08-26
  Last Modified : 2026-08-26

  Dependencies:
    Standard library only (urllib, csv).

  Description:
    Each cell in the link column may be BLANK, a single URL, or several URLs
    separated by '|'. Blank cells are skipped, not treated as errors. Files
    are named from the URL and prefixed with the row's id column when one is
    present, so attachment_1.pdf from two comments cannot collide. A failed
    download is logged and the run continues.

  Input Parameters (required first):
    InputCsvPath   (REQUIRED, str) - CSV containing the link column.
    OutputFolder   (REQUIRED, str) - destination folder; created if needed.
    LinkColumn     (optional, str, default "attachmentLinks")
    IdColumn       (optional, str, default "commentId") - filename prefix.
    Overwrite      (optional, bool, default False) - re-download existing files.
=====================================================================
"""

from __future__ import annotations

import sys
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any, Dict, List

_HERE = Path(__file__).resolve()
if str(_HERE.parent.parent.parent) not in sys.path:
    sys.path.insert(0, str(_HERE.parent.parent.parent))

from src.utils.helpers import readCsv                    # noqa: E402
from src.utils.logger import logError, logInfo, logWarn   # noqa: E402

__version__ = "1.0beta"

DEFAULT_LINK_COLUMN = "attachmentLinks"
DEFAULT_ID_COLUMN = "commentId"
LINK_SEPARATOR = "|"


def downloadBulkFiles(InputCsvPath: str, OutputFolder: str,
                      LinkColumn: str = DEFAULT_LINK_COLUMN,
                      IdColumn: str = DEFAULT_ID_COLUMN,
                      Overwrite: bool = False,
                      TimeoutSeconds: int = 60) -> Dict[str, Any]:
    """Download every file listed in a CSV link column.

    Parameters:
        InputCsvPath (str)   - REQUIRED CSV holding the links.
        OutputFolder (str)   - REQUIRED destination folder.
        LinkColumn (str)     - column with the URLs; default "attachmentLinks".
                               Cells may be blank, one URL, or several joined
                               by '|'.
        IdColumn (str)       - column used to prefix saved filenames so that
                               identically named attachments do not collide;
                               default "commentId". Ignored when absent.
        Overwrite (bool)     - re-download a file that already exists.
        TimeoutSeconds (int) - per-request timeout.
    Returns:
        dict with Downloaded, Skipped, Failed (counts) and Files (list of
        saved paths).
    Raises:
        ValueError - a required parameter is missing, or LinkColumn is not
                     present in the CSV header.
        OSError    - the CSV cannot be read or the folder cannot be created.

    Use in claims processing:
        Pull every public-comment attachment for a CMS docket into one folder
        so the documents can be text-extracted and reviewed in bulk.
    """
    if not InputCsvPath or not str(InputCsvPath).strip():
        raise ValueError("required parameter 'InputCsvPath' is missing or empty")
    if not OutputFolder or not str(OutputFolder).strip():
        raise ValueError("required parameter 'OutputFolder' is missing or empty")

    rows = readCsv(InputCsvPath)
    if rows and LinkColumn not in rows[0]:
        raise ValueError(
            f"column '{LinkColumn}' not found in {InputCsvPath}. "
            f"Available columns: {', '.join(rows[0].keys())}")

    destination = Path(OutputFolder)
    destination.mkdir(parents=True, exist_ok=True)

    downloaded, skipped, failed = 0, 0, 0
    saved: List[str] = []

    for row in rows:
        cell = (row.get(LinkColumn) or "").strip()
        if not cell:                      # BLANK is normal, not an error
            skipped += 1
            continue
        rowId = (row.get(IdColumn) or "").strip() if IdColumn else ""
        for url in [u.strip() for u in cell.split(LINK_SEPARATOR) if u.strip()]:
            name = Path(urllib.parse.unquote(urllib.parse.urlparse(url).path)).name
            if not name:
                logWarn(f"cannot derive a filename from {url}; skipped")
                failed += 1
                continue
            target = destination / (f"{rowId}_{name}" if rowId else name)
            if target.exists() and not Overwrite:
                logInfo(f"exists, skipping: {target.name}")
                skipped += 1
                continue
            try:
                request = urllib.request.Request(url, headers={"User-Agent": "cgs_ai/1.0"})
                with urllib.request.urlopen(request, timeout=TimeoutSeconds) as response:
                    target.write_bytes(response.read())
                downloaded += 1
                saved.append(str(target))
                logInfo(f"downloaded {target.name}")
            except (urllib.error.URLError, OSError) as exc:
                failed += 1
                logError(f"failed {url}: {exc}")

    logInfo(f"done; downloaded={downloaded} skipped={skipped} failed={failed}")
    return {"Downloaded": downloaded, "Skipped": skipped, "Failed": failed,
            "Files": saved}
