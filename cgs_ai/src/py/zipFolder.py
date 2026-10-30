"""
=====================================================================
  Program Name  : zipFolder.py
  Author        : Manuel Figallo
  Purpose       : Zip a folder together with a list of accompanying files
                  into a single archive, for packaging or distribution.
  Version       : 1.0beta
  Created       : 2026-08-26
  Last Modified : 2026-08-26

  Dependencies:
    Standard library only (zipfile, os).

  Description:
    The folder is stored under its own name inside the archive; the extra
    files land at the archive root, beside it. __pycache__ is always
    skipped. Deflate compression.

  Input Parameters (required first):
    FolderToZip     (REQUIRED, str)       - folder whose contents are archived.
    AccompanyFiles  (optional, list[str]) - extra files placed at the root.
    OutputZipPath   (REQUIRED, str)       - full path of the .zip to create.
=====================================================================
"""

from __future__ import annotations

import os
import sys
import zipfile
from pathlib import Path
from typing import Any, Dict, List

_HERE = Path(__file__).resolve()
if str(_HERE.parent.parent.parent) not in sys.path:
    sys.path.insert(0, str(_HERE.parent.parent.parent))

from src.utils.helpers import asList, ensureParent  # noqa: E402
from src.utils.logger import logInfo, logWarn       # noqa: E402

__version__ = "1.0beta"

SKIP_DIRS = {"__pycache__", ".git", ".pytest_cache", ".venv"}


def zipFolder(FolderToZip: str, OutputZipPath: str,
              AccompanyFiles: Any = ()) -> Dict[str, Any]:
    """Archive a folder plus extra files into one .zip.

    Parameters:
        FolderToZip (str)    - REQUIRED folder to archive; stored under its
                               own basename inside the zip.
        OutputZipPath (str)  - REQUIRED full path of the .zip to write.
        AccompanyFiles       - extra file path(s), list or ';'-string; each is
                               stored at the archive root next to the folder.
    Returns:
        dict with OutputZipPath, FileCount and Names (sorted archive members).
    Raises:
        ValueError - a required parameter is missing.
        OSError    - the folder does not exist or the zip cannot be written.

    Use in claims processing:
        Bundle a month of scan output, the formatted Excel report and the
        run log into one archive for records retention or hand-off.
    """
    if not FolderToZip or not str(FolderToZip).strip():
        raise ValueError("required parameter 'FolderToZip' is missing or empty")
    if not OutputZipPath or not str(OutputZipPath).strip():
        raise ValueError("required parameter 'OutputZipPath' is missing or empty")

    source = Path(FolderToZip)
    if not source.is_dir():
        raise OSError(f"FolderToZip is not a directory: {source}")

    target = ensureParent(OutputZipPath)
    extras: List[str] = asList(AccompanyFiles)
    names: List[str] = []

    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        for dirpath, dirnames, filenames in os.walk(source):
            dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS]
            for filename in filenames:
                full = Path(dirpath) / filename
                arcname = str(Path(source.name) / full.relative_to(source))
                archive.write(full, arcname)
                names.append(arcname)
        for extra in extras:
            path = Path(extra)
            if not path.is_file():
                logWarn(f"accompanying file not found, skipped: {path}")
                continue
            archive.write(path, path.name)
            names.append(path.name)

    logInfo(f"wrote {len(names)} file(s) to {target}")
    return {"OutputZipPath": str(target), "FileCount": len(names),
            "Names": sorted(names)}
