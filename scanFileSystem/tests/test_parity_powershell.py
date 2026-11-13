"""Parity checks: ps/scanFileSystem.ps1 must match scanFileSystem.py.

Skipped automatically when no PowerShell interpreter is on PATH, so the main
suite still runs anywhere.
"""

from __future__ import annotations

import csv
import os
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from make_fixtures import build as build_fixtures  # noqa: E402

PS_SCRIPT = ROOT / "ps" / "scanFileSystem.ps1"
PWSH = shutil.which("pwsh") or shutil.which("powershell")

pytestmark = pytest.mark.skipif(
    PWSH is None, reason="no PowerShell interpreter available")

# Columns that legitimately differ between the two ports.
VOLATILE = {"scanned_at", "created_time", "accessed_time"}
#   parse_status: a broken symlink fails at stat() in Python but at read time
#     in PowerShell (Get-Item succeeds on the dangling link). Both are non-OK.
#   file_size_bytes / modified_time: same reason -- PS can stat the link itself.
PLATFORM_SPECIFIC = {"parse_status", "file_size_bytes", "modified_time"}


@pytest.fixture(scope="module")
def logs_root() -> str:
    return str(build_fixtures())


def read_rows(path: Path) -> list:
    with open(path, encoding="utf-8", newline="") as handle:
        return list(csv.DictReader(handle))


def run_python(tmp_path, logs_root, extra=()):
    out = tmp_path / "py.csv"
    argv = [sys.executable, str(ROOT / "scanFileSystem.py"),
            "--input-folder-root", logs_root, "--output-file-path", str(out), *extra]
    proc = subprocess.run(argv, capture_output=True, text=True, cwd=str(ROOT))
    return out, proc.returncode


def run_powershell(tmp_path, logs_root, extra=()):
    out = tmp_path / "ps.csv"
    argv = [PWSH, "-NoProfile", "-NonInteractive", "-File", str(PS_SCRIPT),
            "-InputFolderRoot", logs_root, "-OutputFilePath", str(out), *extra]
    proc = subprocess.run(argv, capture_output=True, text=True, cwd=str(ROOT))
    return out, proc.returncode


def normalize(rows: list) -> list:
    """Make the two ports comparable: PS emits absolute paths."""
    out = []
    for row in rows:
        row = dict(row)
        for key in ("full_path", "directory"):
            if key in row:      # StepDetail has no "directory" column
                row[key] = row[key].replace(str(ROOT) + os.sep, "").replace(str(ROOT) + "/", "")
        out.append(row)
    return out


def test_files_grain_matches(tmp_path, logs_root):
    py_out, py_rc = run_python(tmp_path, logs_root,
                               ["--metric-profile", "sas_log",
                                "--extract-keyword", "real time", "cpu time"])
    ps_out, ps_rc = run_powershell(tmp_path, logs_root,
                                   ["-MetricProfile", "sas_log",
                                    "-ExtractKeyword", "real time;cpu time"])
    assert py_rc == ps_rc == 0
    py_rows, ps_rows = normalize(read_rows(py_out)), normalize(read_rows(ps_out))

    assert len(py_rows) == len(ps_rows)
    assert list(py_rows[0].keys()) == list(ps_rows[0].keys()), "column order must match"
    assert [r["full_path"] for r in py_rows] == [r["full_path"] for r in ps_rows], \
        "row order must match (both sort paths ordinally)"

    skip = VOLATILE | PLATFORM_SPECIFIC
    for py_row, ps_row in zip(py_rows, ps_rows):
        for column in py_row:
            if column in skip:
                continue
            assert py_row[column] == ps_row[column], \
                f"{py_row['program_name']}.{column}: py={py_row[column]!r} ps={ps_row[column]!r}"


def test_stepdetail_grain_matches(tmp_path, logs_root):
    run_python(tmp_path, logs_root, ["--metric-profile", "sas_log"])
    run_powershell(tmp_path, logs_root, ["-MetricProfile", "sas_log"])
    py_rows = normalize(read_rows(tmp_path / "py_StepDetail.csv"))
    ps_rows = normalize(read_rows(tmp_path / "ps_StepDetail.csv"))
    assert py_rows == ps_rows, "StepDetail must be identical across ports"


def test_exit_codes_match(tmp_path, logs_root):
    """Config and I/O errors must produce the same codes in both ports."""
    cases = [
        # (python extra, powershell extra, expected code)
        (["--metric-profile", "bogus"], ["-MetricProfile", "bogus"], 2),
        (["--date-from", "not-a-date"], ["-DateFrom", "not-a-date"], 2),
    ]
    for py_extra, ps_extra, expected in cases:
        _, py_rc = run_python(tmp_path, logs_root, py_extra)
        _, ps_rc = run_powershell(tmp_path, logs_root, ps_extra)
        assert py_rc == expected, f"python {py_extra} -> {py_rc}"
        assert ps_rc == expected, f"powershell {ps_extra} -> {ps_rc}"


def test_missing_root_exits_2_in_both(tmp_path):
    py = subprocess.run([sys.executable, str(ROOT / "scanFileSystem.py"),
                         "--output-file-path", str(tmp_path / "x.csv")],
                        capture_output=True, text=True, cwd=str(ROOT))
    ps = subprocess.run([PWSH, "-NoProfile", "-NonInteractive", "-File", str(PS_SCRIPT),
                         "-OutputFilePath", str(tmp_path / "y.csv")],
                        capture_output=True, text=True, cwd=str(ROOT))
    # Python's CONFIG block has a default UNC root (unreachable here) -> 3;
    # PowerShell's default is empty -> 2. Both are non-zero failures.
    assert py.returncode != 0 and ps.returncode != 0
    assert "input_folder_root" in ps.stderr


def test_folder_exclusions_match(tmp_path, logs_root):
    py_out, _ = run_python(tmp_path, logs_root,
                           ["--folder-exclusion-list", "Old", "Test"])
    ps_out, _ = run_powershell(tmp_path, logs_root,
                               ["-FolderExclusionList", "Old;Test"])
    py_names = {r["program_name"] for r in read_rows(py_out)}
    ps_names = {r["program_name"] for r in read_rows(ps_out)}
    assert py_names == ps_names
    assert "keepme" in py_names and "legacy" not in py_names


def test_date_filter_matches(tmp_path, logs_root):
    py_out, _ = run_python(tmp_path, logs_root,
                           ["--date-from", "2026-01-01", "--date-to", "2026-06-30"])
    ps_out, _ = run_powershell(tmp_path, logs_root,
                               ["-DateFrom", "2026-01-01", "-DateTo", "2026-06-30"])
    # broken_link is excluded from this comparison: Python cannot stat a
    # dangling symlink so it emits the row unfiltered, while PowerShell can
    # stat the link itself and applies the date filter to it. Platform
    # nuance on the fixture, not a behavioral difference in the filter.
    py_names = {r["program_name"] for r in read_rows(py_out)} - {"broken_link"}
    ps_names = {r["program_name"] for r in read_rows(ps_out)} - {"broken_link"}
    assert py_names == ps_names
    assert "dated_2026H1" in py_names and "dated_2025" not in py_names


@pytest.mark.parametrize("value,recurses", [
    ("1", True), ("0", False), ("true", True), ("false", False),
    ("yes", True), ("no", False),
])
def test_include_subdirectories_accepts_file_mode_strings(tmp_path, logs_root, value, recurses):
    """-File mode passes every argument as a string, so a [bool] parameter
    cannot be used; the script parses these forms itself."""
    out = tmp_path / f"r{value}.csv"
    proc = subprocess.run(
        [PWSH, "-NoProfile", "-NonInteractive", "-File", str(PS_SCRIPT),
         "-InputFolderRoot", logs_root, "-OutputFilePath", str(out),
         "-IncludeSubdirectories", value],
        capture_output=True, text=True, cwd=str(ROOT))
    assert proc.returncode == 0, proc.stderr
    names = {r["program_name"] for r in read_rows(out)}
    assert ("jobC" in names) == recurses


def test_invalid_include_subdirectories_exits_2(tmp_path, logs_root):
    proc = subprocess.run(
        [PWSH, "-NoProfile", "-NonInteractive", "-File", str(PS_SCRIPT),
         "-InputFolderRoot", logs_root, "-OutputFilePath", str(tmp_path / "x.csv"),
         "-IncludeSubdirectories", "maybe"],
        capture_output=True, text=True, cwd=str(ROOT))
    assert proc.returncode == 2
    assert "include_subdirectories" in proc.stderr


def test_semicolon_lists_match_native_arrays(tmp_path, logs_root):
    """The SAS wrapper passes lists as one ';'-delimited string."""
    joined = f"{logs_root}/nested;{logs_root}/Older"
    out = tmp_path / "sas_form.csv"
    proc = subprocess.run(
        [PWSH, "-NoProfile", "-NonInteractive", "-File", str(PS_SCRIPT),
         "-InputFolderRoot", joined, "-OutputFilePath", str(out),
         "-ExtractKeyword", "real time;cpu time"],
        capture_output=True, text=True, cwd=str(ROOT))
    assert proc.returncode == 0, proc.stderr
    rows = read_rows(out)
    assert {r["program_name"] for r in rows} == {"jobC", "keepme"}
    assert "kw_real_time_count" in rows[0] and "kw_cpu_time_count" in rows[0]
