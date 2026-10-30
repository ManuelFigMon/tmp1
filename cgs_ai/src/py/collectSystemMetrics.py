"""
=====================================================================
  Program Name  : collectSystemMetrics.py
  Author        : Manuel Figallo
  Purpose       : Collect as many host metrics as the platform allows and
                  append or overwrite them in a CSV time series.
  Version       : 1.0beta
  Created       : 2026-08-26
  Last Modified : 2026-08-26

  Dependencies:
    Standard library only. psutil is used opportunistically when installed
    (richer CPU/memory/disk figures) but is never required.

  Description:
    FAILS GRACEFULLY BY DESIGN: every probe is independent and wrapped, so a
    metric that is unavailable on this host is recorded as blank and the run
    still succeeds. This mirrors the PowerShell Get-Counter approach, where
    some counters exist only on some servers.

  Input Parameters (required first):
    ServerName    (optional, str)  - defaults to this host's name.
    OutputCsvPath (REQUIRED, str)  - CSV to write.
    WriteMode     (optional, str, default "append") - append | overwrite.
=====================================================================
"""

from __future__ import annotations

import datetime as dt
import os
import platform
import shutil
import socket
import sys
from pathlib import Path
from typing import Any, Callable, Dict, Optional

_HERE = Path(__file__).resolve()
if str(_HERE.parent.parent.parent) not in sys.path:
    sys.path.insert(0, str(_HERE.parent.parent.parent))

from src.utils.helpers import ensureParent, isoNow, readCsv, writeCsv  # noqa: E402
from src.utils.logger import logInfo, logWarn                          # noqa: E402

__version__ = "1.0beta"

METRIC_COLUMNS = [
    "Timestamp", "ServerName", "OSVersion", "CPUName", "CPUCount",
    "CPUUsagePercent", "TotalPhysicalMemoryGB", "MemoryAvailableMB",
    "MemoryUsedPercent", "DiskTotalGB", "DiskFreeGB", "DiskFreePercent",
    "PythonVersion", "Errors",
]
VALID_WRITE_MODES = ("append", "overwrite")


def _safe(probe: Callable[[], Any], label: str, errors: list) -> Any:
    """Run one probe, recording rather than raising on failure.

    Parameters:
        probe (callable) - zero-arg function returning the metric.
        label (str)      - metric name used in the error note.
        errors (list)    - accumulator for failure labels.
    Returns: the metric value, or "" when the probe failed.
    """
    try:
        value = probe()
        return "" if value is None else value
    except Exception as exc:                      # deliberately broad: fail soft
        errors.append(f"{label}:{type(exc).__name__}")
        return ""


def collectSystemMetrics(OutputCsvPath: str, ServerName: Optional[str] = None,
                         WriteMode: str = "append") -> Dict[str, Any]:
    """Gather host metrics and record them in a CSV.

    Parameters:
        OutputCsvPath (str) - REQUIRED CSV to write.
        ServerName (str)    - host label; defaults to the machine name.
        WriteMode (str)     - "append" (default) adds a row, preserving
                              history; "overwrite" replaces the file.
    Returns:
        dict with Metrics (the row just collected), OutputCsvPath, WriteMode
        and RowCount (rows in the file after writing).
    Raises:
        ValueError - OutputCsvPath missing, or WriteMode not append/overwrite.
        OSError    - the CSV cannot be written.
        Individual metric failures NEVER raise; they are blank and are listed
        in the row's Errors column.

    Use in claims processing:
        Sample the SAS/ETL server during a nightly claims run to correlate
        slow steps (see scanFileSystem's metric profile) with CPU, memory or
        disk pressure on the host.
    """
    if not OutputCsvPath or not str(OutputCsvPath).strip():
        raise ValueError("required parameter 'OutputCsvPath' is missing or empty")
    if WriteMode not in VALID_WRITE_MODES:
        raise ValueError(f"unknown WriteMode {WriteMode!r}; expected one of: "
                         f"{', '.join(VALID_WRITE_MODES)}")

    errors: list = []
    host = ServerName or _safe(socket.gethostname, "hostname", errors) or "unknown"

    psutil = None
    try:
        import psutil as _psutil
        psutil = _psutil
    except ImportError:
        logWarn("psutil not installed; collecting the standard-library subset")

    def cpuPercent():
        return round(psutil.cpu_percent(interval=1.0), 2) if psutil else ""

    def memoryAvailableMb():
        return round(psutil.virtual_memory().available / (1024 ** 2), 2) if psutil else ""

    def memoryUsedPercent():
        return round(psutil.virtual_memory().percent, 2) if psutil else ""

    def totalMemoryGb():
        if psutil:
            return round(psutil.virtual_memory().total / (1024 ** 3), 2)
        pages = os.sysconf("SC_PHYS_PAGES") if hasattr(os, "sysconf") else None
        size = os.sysconf("SC_PAGE_SIZE") if hasattr(os, "sysconf") else None
        return round(pages * size / (1024 ** 3), 2) if pages and size else ""

    usage = _safe(lambda: shutil.disk_usage(os.path.abspath(os.sep)),
                  "disk", errors)

    row: Dict[str, Any] = {
        "Timestamp": dt.datetime.now().strftime("%m/%d/%Y %H:%M:%S"),
        "ServerName": host,
        "OSVersion": _safe(lambda: f"{platform.system()} {platform.release()}",
                           "os", errors),
        "CPUName": _safe(lambda: platform.processor() or platform.machine(),
                         "cpuName", errors),
        "CPUCount": _safe(lambda: os.cpu_count(), "cpuCount", errors),
        "CPUUsagePercent": _safe(cpuPercent, "cpuUsage", errors),
        "TotalPhysicalMemoryGB": _safe(totalMemoryGb, "totalMemory", errors),
        "MemoryAvailableMB": _safe(memoryAvailableMb, "memAvailable", errors),
        "MemoryUsedPercent": _safe(memoryUsedPercent, "memUsed", errors),
        "DiskTotalGB": round(usage.total / (1024 ** 3), 2) if usage else "",
        "DiskFreeGB": round(usage.free / (1024 ** 3), 2) if usage else "",
        "DiskFreePercent": (round(usage.free / usage.total * 100, 2)
                            if usage and usage.total else ""),
        "PythonVersion": platform.python_version(),
        "Errors": ";".join(errors),
    }

    existing = []
    target = Path(OutputCsvPath)
    if WriteMode == "append" and target.is_file():
        try:
            existing = readCsv(OutputCsvPath)
        except OSError as exc:
            logWarn(f"cannot read existing {OutputCsvPath} ({exc}); overwriting")
    ensureParent(OutputCsvPath)
    allRows = existing + [row]
    written = writeCsv(allRows, METRIC_COLUMNS, OutputCsvPath)

    logInfo(f"collected metrics for {host} ({WriteMode}) -> {written}")
    if errors:
        logWarn(f"{len(errors)} metric(s) unavailable on this host: {';'.join(errors)}")
    return {"Metrics": row, "OutputCsvPath": written, "WriteMode": WriteMode,
            "RowCount": len(allRows)}
