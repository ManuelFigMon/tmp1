"""Tests for the scanFileSystem match-grain output and token extraction."""
from __future__ import annotations
import csv, sys
from pathlib import Path
import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from src.py.scanFileSystem import (MATCH_COLUMNS, extractTokens,  # noqa: E402
                                   isNumericToken, parseDuration, scanFileSystem)

SAMPLE = """1    data work.a; set sashelp.class; run;
NOTE: DATA statement used (Total process time):
      real time           0.05 seconds
      cpu time            0.03 seconds
WARNING: Variable height is uninitialized.
NOTE: The claims extract wrote 45231 records to CLAIMS.OUT
ERROR: File WORK.MISSING.DATA does not exist.
"""


@pytest.fixture
def logsRoot(tmp_path):
    (tmp_path / "jobA.log").write_text(SAMPLE, encoding="utf-8")
    return str(tmp_path)


def readRows(path):
    with open(path, encoding="utf-8", newline="") as handle:
        return list(csv.DictReader(handle))


def test_output_has_the_specified_columns_in_order(logsRoot, tmp_path):
    out = tmp_path / "o.csv"
    scanFileSystem(input_folder_root=logsRoot, extract_keyword=["real time"],
                   output_file_path=str(out))
    assert list(readRows(out)[0].keys()) == MATCH_COLUMNS


def test_grain_is_one_row_per_match_not_per_file(logsRoot, tmp_path):
    """A file with three matching lines must yield three rows."""
    out = tmp_path / "o.csv"
    result = scanFileSystem(input_folder_root=logsRoot,
                            extract_keyword=["NOTE", "ERROR"],
                            output_file_path=str(out))
    rows = readRows(out)
    # SAMPLE has 2 lines containing "NOTE" and 1 containing "ERROR".
    assert len(rows) == 3
    assert len({r["FullPath"] for r in rows}) == 1   # all from one file
    assert [r["Keyword"] for r in rows] == ["NOTE", "NOTE", "ERROR"]


def test_no_keyword_means_no_rows(logsRoot, tmp_path):
    out = tmp_path / "o.csv"
    scanFileSystem(input_folder_root=logsRoot, extract_keyword=["zzz-absent"],
                   output_file_path=str(out))
    assert readRows(out) == []


def test_context_window_defaults_to_five_and_is_configurable(logsRoot, tmp_path):
    out = tmp_path / "o.csv"
    scanFileSystem(input_folder_root=logsRoot, extract_keyword=["ERROR"],
                   output_file_path=str(out), lines_above=2, lines_below=0)
    row = readRows(out)[0]
    assert len(row["LinesAbove"].splitlines()) == 2
    assert row["LinesBelow"] == ""

    out2 = tmp_path / "o2.csv"
    scanFileSystem(input_folder_root=logsRoot, extract_keyword=["ERROR"],
                   output_file_path=str(out2))          # defaults
    assert len(readRows(out2)[0]["LinesAbove"].splitlines()) == 5


def test_line_number_is_one_based(logsRoot, tmp_path):
    out = tmp_path / "o.csv"
    scanFileSystem(input_folder_root=logsRoot, extract_keyword=["ERROR"],
                   output_file_path=str(out))
    assert readRows(out)[0]["LineNumber"] == "7"


@pytest.mark.parametrize("nth,expected", [(1, "0.05"), (2, "seconds")])
def test_nth_token_after_is_configurable(nth, expected):
    tokens = extractTokens("      real time           0.05 seconds",
                           "real time", nth, 1, 1)
    assert tokens["NthTokenAfter"] == expected


def test_token_extraction_fields():
    line = "NOTE: The claims extract wrote 45231 records to CLAIMS.OUT"
    tokens = extractTokens(line, "wrote", 1, 1, 1)
    assert tokens["NthTokenAfter"] == "45231"
    assert tokens["NthTokenBefore"] == "extract"
    assert tokens["NumericTokenAfter"] == "45231"
    assert tokens["FirstToken"] == "NOTE:"
    assert tokens["LastToken"] == "CLAIMS.OUT"


@pytest.mark.parametrize("token,expected", [
    ("45231", True), ("1,204", True), ("$99.50", True), ("12%", True),
    ("-3.5", True), ("records", False), ("CLAIMS.OUT", False)])
def test_numeric_token_detection(token, expected):
    assert isNumericToken(token) is expected


@pytest.mark.parametrize("raw,expected", [
    ("0.05 seconds", 0.05), ("1.20", 1.20), ("1:03.05", 63.05),
    ("1:00:30.00", 3630.0), ("garbage", None)])
def test_duration_parser(raw, expected):
    result = parseDuration(raw)
    assert result == pytest.approx(expected) if expected else result is None
