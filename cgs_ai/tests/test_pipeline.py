"""Tests for the filescan pipeline, package exports and cross-language parity."""
from __future__ import annotations
import csv, re, subprocess, sys
from pathlib import Path
import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT.parent))          # so `import cgs_ai` works
sys.path.insert(0, str(ROOT))

import cgs_ai                                            # noqa: E402
from src.pipelines.filescan_pipeline import runFilescanPipeline  # noqa: E402

SAMPLE = """NOTE: DATA statement used (Total process time):
      real time           0.05 seconds
      cpu time            0.03 seconds
NOTE: PROCEDURE MEANS used (Total process time):
      real time           1.20 seconds
      cpu time            0.90 seconds
"""


@pytest.fixture
def logsRoot(tmp_path):
    (tmp_path / "jobA.log").write_text(SAMPLE, encoding="utf-8")
    return str(tmp_path)


# --- package surface ---------------------------------------------------------

def test_version_is_declared():
    assert cgs_ai.__version__ == "1.0beta"


@pytest.mark.parametrize("name", [
    "scanFileSystem", "runSQLServerQuery", "formatCSV", "downloadBulkFiles",
    "sendEmail", "convertSAS2Pandas", "copyExcelSheet2CSV",
    "collectSystemMetrics", "zipFolder", "runFilescanPipeline",
    "basic_hello", "personalized_hello", "detailed_hello"])
def test_every_function_is_exported(name):
    assert hasattr(cgs_ai, name), f"cgs_ai.{name} is missing"
    assert name in cgs_ai.__all__


def test_greeting_functions():
    assert cgs_ai.basic_hello() == "Hello, World!"
    assert cgs_ai.personalized_hello("  Manuel  ") == "Hello, Manuel!"
    assert cgs_ai.personalized_hello("") == "Hello, World!"
    assert cgs_ai.detailed_hello("pirate")["style_used"] == "pirate"
    assert "Ahoy" in cgs_ai.detailed_hello("pirate")["message"]
    assert cgs_ai.detailed_hello("unknown")["message"] == \
        cgs_ai.detailed_hello("friendly")["message"]


def test_importing_cgs_ai_needs_no_third_party_package():
    """The package must import with the standard library alone."""
    code = ("import sys;"
            "blocked=('pandas','numpy','openpyxl','pyodbc','psutil');"
            "[sys.modules.__setitem__(m, None) for m in blocked];"
            f"sys.path.insert(0, r'{ROOT.parent}');"
            "import cgs_ai; print(cgs_ai.__version__)")
    result = subprocess.run([sys.executable, "-c", code],
                            capture_output=True, text=True)
    assert result.returncode == 0, result.stderr
    assert "1.0beta" in result.stdout


# --- pipeline ----------------------------------------------------------------

def test_pipeline_runs_scan_and_reports_steps(logsRoot, tmp_path):
    result = runFilescanPipeline(
        input_folder_root=logsRoot, extract_keyword=["real time"],
        output_file_path=str(tmp_path / "scan.csv"), metric_profile="none")
    assert "scanFileSystem" in result["Steps"]
    assert len(result["Scan"]["matches"]) == 2


def test_pipeline_skips_email_when_no_recipient(logsRoot, tmp_path):
    result = runFilescanPipeline(
        input_folder_root=logsRoot, extract_keyword=["real time"],
        output_file_path=str(tmp_path / "scan.csv"), metric_profile="none")
    assert result["Email"] is None
    assert "sendEmail" not in result["Steps"]


def test_metric_profile_switches_output_to_excel(logsRoot, tmp_path):
    """A .csv request becomes .xlsx, because metrics need a second sheet."""
    pytest.importorskip("openpyxl")
    result = cgs_ai.scanFileSystem(
        input_folder_root=logsRoot, extract_keyword=["real time"],
        output_file_path=str(tmp_path / "scan.csv"), metric_profile="sas_log")
    assert result["output"].endswith(".xlsx")
    assert len(result["metrics"]) == 2
    from openpyxl import load_workbook
    assert load_workbook(result["output"]).sheetnames == ["Matches", "Metrics"]


def test_fullstimer_cpu_time_is_captured(tmp_path):
    """FULLSTIMER logs say 'user cpu time', not 'cpu time'."""
    (tmp_path / "fs.log").write_text(
        "NOTE: PROCEDURE SORT used (Total process time):\n"
        "      real time           12.34 seconds\n"
        "      user cpu time       0.80 seconds\n"
        "      system cpu time     0.20 seconds\n", encoding="utf-8")
    result = cgs_ai.scanFileSystem(
        input_folder_root=str(tmp_path), extract_keyword=["real time"],
        output_file_path=str(tmp_path / "o.csv"), metric_profile="none")
    from src.py.scanFileSystem import METRIC_PROFILES, parseMetricProfile
    lines = (tmp_path / "fs.log").read_text().splitlines()
    rows = parseMetricProfile(METRIC_PROFILES["sas_log"], lines, "p", "fs")
    assert rows[0]["CpuTimeSec"] == pytest.approx(0.80), \
        "FULLSTIMER 'user cpu time' must be captured, not silently dropped"


# --- cross-language parity ---------------------------------------------------

PS_DIR = ROOT / "src" / "ps"
PY_DIR = ROOT / "src" / "py"
FUNCTIONS = ["scanFileSystem", "runSQLServerQuery", "formatCSV",
             "downloadBulkFiles", "sendEmail", "convertSAS2Pandas",
             "copyExcelSheet2CSV", "collectSystemMetrics", "zipFolder"]


@pytest.mark.parametrize("name", FUNCTIONS)
def test_every_function_exists_in_both_languages(name):
    assert (PY_DIR / f"{name}.py").is_file(), f"missing src/py/{name}.py"
    assert (PS_DIR / f"{name}.ps1").is_file(), f"missing src/ps/{name}.ps1"


@pytest.mark.parametrize("name", FUNCTIONS)
def test_every_function_has_the_standard_header(name):
    """Every file carries the scanFileSystem header block."""
    for path in (PY_DIR / f"{name}.py", PS_DIR / f"{name}.ps1"):
        text = path.read_text(encoding="utf-8")[:3000]
        for field in ("Program Name", "Author", "Purpose", "Version",
                      "Dependencies", "Input Parameters"):
            assert field in text, f"{path.name} header is missing '{field}'"


@pytest.mark.parametrize("name", FUNCTIONS)
def test_sas_wrapper_exists_for_every_function(name):
    sasText = (ROOT / "src" / "sas" / "cgsFunctions.sas").read_text()
    assert re.search(rf"%macro\s+{name}\s*\(", sasText), \
        f"cgsFunctions.sas has no %{name} wrapper"


def test_scanfilesystem_parameter_names_match_across_languages():
    """Parameter names must be identical in Python, PowerShell and SAS."""
    expected = {
        "input_folder_root", "extract_keyword", "output_file_path",
        "file_extensions", "include_subdirectories", "folder_exclusion_list",
        "file_exclusion_list", "lines_above", "lines_below",
        "nth_token_after", "nth_token_before", "numeric_token_after",
        "date_from", "date_to", "date_field", "metric_profile"}

    import inspect
    pySig = set(inspect.signature(cgs_ai.scanFileSystem).parameters) - {"self"}
    assert expected <= pySig, f"Python is missing {expected - pySig}"

    psText = (PS_DIR / "scanFileSystem.ps1").read_text()
    psParams = set(re.findall(r"\$(\w+)\s*=", psText.split("Set-StrictMode")[0]))
    assert expected <= psParams, f"PowerShell is missing {expected - psParams}"

    sasText = (ROOT / "src" / "sas" / "cgsFunctions.sas").read_text()
    block = re.search(r"%macro\s+scanFileSystem\s*\((.*?)\);", sasText, re.DOTALL).group(1)
    block = re.sub(r"/\*.*?\*/", "", block, flags=re.DOTALL)
    sasParams = set(re.findall(r"(\w+)\s*=", block))
    assert expected <= sasParams, f"SAS is missing {expected - sasParams}"


# --- formatCSV: the SAS ODS1 renderer and the PowerShell native writer -------

SAS_FUNCTIONS = (ROOT / "src" / "sas" / "cgsFunctions.sas").read_text()
PS_FORMATCSV = (ROOT / "src" / "ps" / "formatCSV.ps1").read_text()


def test_sas_macro_and_mend_are_balanced():
    """Every %macro needs its %mend, or everything after it is swallowed."""
    opens = re.findall(r"^%macro\s+(\w+)", SAS_FUNCTIONS, re.MULTILINE)
    closes = re.findall(r"^%mend\s+(\w+)", SAS_FUNCTIONS, re.MULTILINE)
    assert opens == closes, f"unbalanced: {set(opens) ^ set(closes)}"


def test_sas_formatcsv_routes_ods1_without_leaving_sas():
    block = re.search(r"%macro formatCSV\b.*?%mend formatCSV;",
                      SAS_FUNCTIONS, re.DOTALL).group(0)
    # ODS1 must be intercepted BEFORE %cgsRun, or it would be handed to
    # PowerShell, which has no such FormatType.
    assert "%upcase(&FormatType) = ODS1" in block
    assert block.index("ODS1") < block.index("%cgsRun")
    assert "%cgsFormatCsvOds" in block
    # The parameter comment must list ODS1, or callers never learn it exists.
    header = block[:block.index(");")]
    assert "ODS1" in header, "the FormatType comment does not mention ODS1"


def test_sas_ods1_renderer_is_self_contained():
    block = re.search(r"%macro cgsFormatCsvOds\b.*?%mend cgsFormatCsvOds;",
                      SAS_FUNCTIONS, re.DOTALL).group(0)
    for needed in ("proc import", "ods excel", "proc report", "ods excel close",
                   "cx1F3864", "cx2E75B6", "cxDCE6F1", "frozen_headers",
                   "autofilter"):
        assert needed in block, f"ODS1 renderer is missing {needed!r}"
    # It must not shell out; that is the whole point of ODS1.
    assert "%cgsRun" not in block


def test_sas_put_statements_have_no_embedded_semicolons():
    """A ';' inside %put text ends the statement early -- ERROR 180-322."""
    for line in SAS_FUNCTIONS.splitlines():
        stripped = line.strip()
        if stripped.startswith("%put"):
            body = stripped[len("%put"):].rstrip()
            assert body.count(";") == 1 and body.endswith(";"), \
                f"%put with an embedded semicolon: {stripped}"


def test_powershell_formatcsv_no_longer_fails_without_importexcel():
    """The missing module must be a fallback, not a fatal error."""
    assert "formatCSV requires the ImportExcel module" not in PS_FORMATCSV
    assert "function Write-XlsxNative" in PS_FORMATCSV
    assert "writing the workbook natively" in PS_FORMATCSV


def test_powershell_native_writer_emits_valid_ooxml_shape():
    # autoFilter must precede mergeCells; Excel refuses the other order.
    assert PS_FORMATCSV.index("'<autoFilter ref=") < PS_FORMATCSV.index("'<mergeCells count=")
    # Fill 0 must be none and fill 1 gray125, or Excel rejects the workbook.
    fills = re.search(r'<fills count="5">(.*?)</fills>', PS_FORMATCSV, re.DOTALL).group(1)
    assert fills.index('patternType="none"') < fills.index('patternType="gray125"')
    # A BOM inside a part makes the file unreadable.
    assert "UTF8Encoding($false)" in PS_FORMATCSV
    assert "<cellStyles count=" in PS_FORMATCSV


def test_powershell_writer_switch_is_documented_and_validated():
    assert "[string] $Writer" in PS_FORMATCSV
    assert "'auto', 'native', 'module'" in PS_FORMATCSV


# --- SAS block comments must not nest ----------------------------------------

def _scanBlockComments(text):
    """Return (nested_line_numbers, unterminated).

    SAS block comments do NOT nest: the first close marker ends the comment,
    so a nested open means everything after that close is parsed as live code.
    """
    depth, index, line, nested = 0, 0, 1, []
    while index < len(text) - 1:
        pair = text[index:index + 2]
        if pair == "/*":
            if depth:
                nested.append(line)
            else:
                depth = 1
            index += 2
            continue
        if pair == "*/" and depth:
            depth = 0
            index += 2
            continue
        if text[index] == "\n":
            line += 1
        index += 1
    return nested, bool(depth)


@pytest.mark.parametrize(
    "sasPath",
    sorted((ROOT / "src" / "sas").glob("*.sas")),
    ids=lambda p: p.name)
def test_sas_block_comments_do_not_nest(sasPath):
    nested, unterminated = _scanBlockComments(sasPath.read_text())
    assert not nested, (
        f"{sasPath.name}: block comment opened again at line(s) {nested} while "
        f"one was already open -- the outer comment ends early and the rest "
        f"of the file is parsed as code")
    assert not unterminated, f"{sasPath.name}: a block comment is never closed"
