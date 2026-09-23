"""
=====================================================================
  Program Name  : config.py
  Author        : Manuel Figallo
  Purpose       : Load key=value configuration from the project .env file
                  and expose it to every cgs_ai function. Standard library
                  only -- no python-dotenv dependency.
  Version       : 1.0beta
  Created       : 2026-08-26
  Last Modified : 2026-08-26

  Description:
    Reads the project .env (KEY=VALUE, one per line, '#' comments) once and
    caches it. Real OS environment variables always win over .env values, so
    a scheduled task or CI job can override configuration without editing
    files. Never logs or prints values, since .env may hold credentials.

  Configuration:
    .env lives at the project root (next to __init__.py). Copy .env.example
    to .env and fill it in. .env is git-ignored; .env.example is committed.
=====================================================================
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Dict, Optional

__version__ = "1.0beta"

#: Variables documented in .env.example.
KNOWN_KEYS = (
    "PS_FOLDER_PATH", "SAS_FOLDER_PATH", "PYTHON_FOLDER_PATH",
    "ROOT_DATA", "ROOT_SRC",
    "SMTP_SERVER", "SMTP_PORT",
    "SQL_DATA_SOURCE", "SQL_LOB_CATALOG",
)

_CACHE: Optional[Dict[str, str]] = None


def projectRoot() -> Path:
    """Return the cgs_ai project root (the folder holding __init__.py).

    Parameters: none.
    Returns: Path to the project root.
    """
    return Path(__file__).resolve().parent.parent.parent


def parseEnvFile(envPath: Path) -> Dict[str, str]:
    """Parse a .env file into a dict.

    Parameters:
        envPath (Path) - file to read; a missing file yields {}.
    Returns:
        dict of KEY -> VALUE. Blank lines and '#' comments are skipped;
        surrounding quotes on the value are stripped; whitespace around '='
        is tolerated even though the convention is to omit it.
    """
    values: Dict[str, str] = {}
    if not envPath.is_file():
        return values
    for raw in envPath.read_text(encoding="utf-8", errors="replace").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        key = key.strip()
        value = value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in "\"'":
            value = value[1:-1]
        if key:
            values[key] = value
    return values


def loadConfig(refresh: bool = False) -> Dict[str, str]:
    """Load and cache the merged configuration.

    Parameters:
        refresh (bool) - re-read the .env file instead of using the cache.
    Returns:
        dict of configuration values. Precedence: OS environment > .env.
    """
    global _CACHE
    if _CACHE is not None and not refresh:
        return _CACHE
    merged = parseEnvFile(projectRoot() / ".env")
    for key in list(merged) + list(KNOWN_KEYS):
        if os.environ.get(key):
            merged[key] = os.environ[key]
    _CACHE = merged
    return merged


def getConfig(key: str, default: Optional[str] = None, required: bool = False) -> Optional[str]:
    """Fetch one configuration value.

    Parameters:
        key (str)       - variable name, e.g. "ROOT_DATA".
        default (str)   - returned when the key is absent and not required.
        required (bool) - raise instead of returning the default when absent.
    Returns:
        The value, or `default`.
    Raises:
        KeyError - when required and absent. The message names the key only,
                   never a value, so credentials cannot leak into a log.
    """
    value = loadConfig().get(key, default)
    if required and (value is None or value == ""):
        raise KeyError(
            f"Required configuration '{key}' is not set. Add it to the project "
            f".env (see .env.example) or export it as an environment variable."
        )
    return value
