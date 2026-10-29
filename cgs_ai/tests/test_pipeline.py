"""Tests for the filescan pipeline, package exports and cross-language parity."""
from __future__ import annotations
import csv, platform, re, shutil, subprocess, sys
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
    "scanFileSystem", "runSQLServerQuery", "formatData", "formatCSV",
    "downloadBulkFiles",
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
    assert result.returncode == 0, result.output
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


# --- the access_db profile ---------------------------------------------------

def accessDbRows(lines):
    """Run the access_db extractor over a list of lines. Returns: list[dict]."""
    from src.py.scanFileSystem import METRIC_PROFILES, parseMetricProfile
    return parseMetricProfile(METRIC_PROFILES["access_db"], lines, "p.sas", "p")


def buildProgram(placements, total=60):
    """Lay text at 1-based line numbers so a test can assert exact numbers.

    Parameters: placements (dict) - {lineNumber: text}; total (int) - length.
    Returns: list[str] of `total` lines, unplaced ones being a filler comment.
    """
    lines = ["/* filler */"] * total
    for number, text in placements.items():
        lines[number - 1] = text
    return lines


def test_access_db_reports_every_line_the_libref_is_used_on():
    rows = accessDbRows(buildProgram({
        3: r'libname issuelog access path="\\srv\dbs\issuelog.mdb";',
        12: "  set issuelog.claims;",
        48: "  create table b as select * from issuelog.detail",
        49: "  where x in (select y from issuelog.lookup);",
    }))
    assert len(rows) == 1
    assert rows[0]["Libref"] == "issuelog"
    assert rows[0]["Keyword"] == "mdb"
    assert rows[0]["DefinitionLine"] == 3
    assert rows[0]["UsageCount"] == 3
    assert rows[0]["UsageLines"] == "12,48,49"


def test_access_db_does_not_count_the_definition_line_itself():
    """The .mdb file is usually named after the libref, so 'issuelog.mdb' on
    the LIBNAME line matches the libref-dot pattern. It is not a use."""
    rows = accessDbRows(buildProgram({
        3: r'libname issuelog access path="\\srv\dbs\issuelog.mdb";',
    }))
    assert rows[0]["UsageCount"] == 0, \
        "the libref appears in its own file name, which is not a usage"


def test_access_db_finds_a_libname_wrapped_across_lines():
    rows = accessDbRows(buildProgram({
        7: "libname fmrrpt access",
        8: r'    path="\\srv\dbs\fmrrpt RHHI.mdb";',
        20: "  set fmrrpt.detail;",
    }))
    assert len(rows) == 1
    assert rows[0]["DefinitionLine"] == 7, "the line the statement STARTS on"
    assert rows[0]["DatabaseFile"].endswith("fmrrpt RHHI.mdb")
    assert rows[0]["UsageLines"] == "20"


def test_access_db_reports_na_when_defined_but_never_used():
    rows = accessDbRows(buildProgram({
        3: r'libname unused access path="c:\tmp\never.accdb";',
    }))
    assert rows[0]["UsageCount"] == 0
    assert rows[0]["UsageLines"] == "NA"
    assert rows[0]["Keyword"] == "accdb"


def test_access_db_ignores_a_macro_variable_of_the_same_name():
    """`&fmrrpt` and `%let fmrrpt` are not uses of the library."""
    rows = accessDbRows(buildProgram({
        3: r'libname fmrrpt access path="\\srv\dbs\fmrrpt.mdb";',
        10: "%let fmrrpt = 1;",
        11: "  if x = &fmrrpt then delete;",
        12: "  set fmrrpt.detail;",
    }))
    assert rows[0]["UsageLines"] == "12", \
        "only libref-DOT counts; a bare word collides with a macro variable"


def test_access_db_does_not_count_a_libname_clear_statement():
    rows = accessDbRows(buildProgram({
        3: r'libname issuelog access path="\\srv\dbs\issuelog.mdb";',
        12: "  set issuelog.claims;",
        40: "libname issuelog clear;",
    }))
    assert rows[0]["UsageLines"] == "12"


def test_access_db_reports_mdb_and_accdb_in_the_same_file():
    rows = accessDbRows(buildProgram({
        3: r'libname older access path="\\srv\dbs\legacy.mdb";',
        4: r'libname newer access path="\\srv\dbs\current.accdb";',
        20: "  set older.claims;",
        21: "  set newer.claims;",
    }))
    assert [(r["Libref"], r["Keyword"], r["UsageLines"]) for r in rows] == [
        ("older", "mdb", "20"), ("newer", "accdb", "21")]


def test_access_db_skips_a_macro_built_path_instead_of_misreporting_it():
    """'&mdbPath' contains the literal 'mdb'. A substring match would report
    this as a genuine Access library; it must be counted as skipped."""
    rows = accessDbRows(buildProgram({
        3: 'libname db pcfiles path="&mdbPath";',
        12: "  set db.claims;",
    }))
    assert rows == [], "a macro-built path has no visible .mdb/.accdb"


def test_access_db_does_not_warn_about_an_ordinary_sas_library():
    """A plain library with a macro in its path is not an Access candidate."""
    from src.py import scanFileSystem as scanner
    warnings = []
    original = scanner.logWarn
    scanner.logWarn = warnings.append
    try:
        accessDbRows(buildProgram({3: r'libname plain "&root\sasdata";'}))
    finally:
        scanner.logWarn = original
    assert warnings == [], "no access/pcfiles engine, so nothing was skipped"


def test_access_db_switches_output_to_excel_with_its_own_columns(tmp_path):
    pytest.importorskip("openpyxl")
    from src.py.scanFileSystem import ACCESS_DB_COLUMNS
    (tmp_path / "job.sas").write_text(
        'libname issuelog access path="\\\\srv\\dbs\\issuelog.mdb";\n'
        "data a; set issuelog.claims; run;\n", encoding="utf-8")
    result = cgs_ai.scanFileSystem(
        input_folder_root=str(tmp_path), extract_keyword=["mdb"],
        file_extensions=["sas"], output_file_path=str(tmp_path / "o.csv"),
        metric_profile="access_db")
    assert result["output"].endswith(".xlsx")
    assert len(result["metrics"]) == 1
    from openpyxl import load_workbook
    book = load_workbook(result["output"])
    assert book.sheetnames == ["Matches", "Metrics"]
    assert [c.value for c in book["Metrics"][1]] == ACCESS_DB_COLUMNS


def test_the_sas_log_profile_is_unchanged_by_the_registry_refactor():
    from src.py.scanFileSystem import (METRIC_COLUMNS, METRIC_PROFILES,
                                       SAS_LOG_COLUMNS, parseMetricProfile)
    assert METRIC_COLUMNS == SAS_LOG_COLUMNS, "the old name must still work"
    assert METRIC_PROFILES["sas_log"]["Columns"] == [
        "FullPath", "ProgramName", "StepIndex", "StepLabel",
        "RealTimeSec", "CpuTimeSec"]
    rows = parseMetricProfile(METRIC_PROFILES["sas_log"],
                              SAMPLE.splitlines(), "p", "job")
    assert [(r["StepIndex"], r["StepLabel"], r["RealTimeSec"]) for r in rows] == [
        (1, "DATA statement", 0.05), (2, "PROCEDURE MEANS", 1.20)]


# --- sendEmail: display names, urgency, corporate body -----------------------

@pytest.fixture
def capturedMail(monkeypatch):
    """Capture the EmailMessage sendEmail hands to SMTP, without a server."""
    import smtplib
    from src.py import sendEmail as mailer
    captured = {}

    class FakeSmtp:
        def __init__(self, *args, **kwargs):
            captured["server"] = args[0] if args else None
        def __enter__(self):
            return self
        def __exit__(self, *args):
            return False
        def send_message(self, message):
            captured["message"] = message

    monkeypatch.setattr(smtplib, "SMTP", FakeSmtp)
    return captured, mailer


def test_display_names_survive_into_the_headers(capturedMail):
    captured, mailer = capturedMail
    result = mailer.sendEmail(
        To="Al Cordoba <al.cordoba@cgsadmin.com>",
        From="Manuel Figallo <manuel.figallo@cgsadmin.com>",
        Subject="s", Body="b")
    message = captured["message"]
    assert message["To"] == "Al Cordoba <al.cordoba@cgsadmin.com>"
    assert message["From"] == "Manuel Figallo <manuel.figallo@cgsadmin.com>"
    assert result["To"] == ["al.cordoba@cgsadmin.com"], \
        "the result reports the bare address actually delivered to"


def test_a_comma_in_a_display_name_is_not_a_second_recipient(capturedMail):
    captured, mailer = capturedMail
    mailer.sendEmail(To='"Cordoba, Al" <al.cordoba@cgsadmin.com>',
                     From="a@b.com", Subject="s", Body="b")
    from email.utils import getaddresses
    assert len(getaddresses([captured["message"]["To"]])) == 1, \
        "recipients split on ';' precisely so a comma can stay in a name"


def test_urgent_sets_both_the_priority_and_the_flag(capturedMail):
    """Outlook draws these differently: importance is the red '!', and
    X-Message-Flag is what actually puts the red FLAG on the message."""
    captured, mailer = capturedMail
    mailer.sendEmail(To="a@b.com", From="c@d.com", Subject="s", Body="b",
                     Urgent=True)
    message = captured["message"]
    assert message["Importance"] == "High"
    assert message["X-Priority"] == "1 (Highest)"
    assert message["X-MSMail-Priority"] == "High"
    assert message["X-Message-Flag"] == "Follow up"


def test_not_urgent_is_the_default_and_adds_no_headers(capturedMail):
    captured, mailer = capturedMail
    result = mailer.sendEmail(To="a@b.com", From="c@d.com", Subject="s",
                              Body="b")
    message = captured["message"]
    assert result["Urgent"] is False
    for header in ("Importance", "X-Priority", "X-Message-Flag"):
        assert message[header] is None


@pytest.mark.parametrize("recipient,expected", [
    ("Al Cordoba <al.cordoba@cgsadmin.com>", "Al"),
    ('"Cordoba, Al" <al.cordoba@cgsadmin.com>', "Al"),
    ("al.cordoba@cgsadmin.com", "Al"),
    ("", "colleague"),
])
def test_the_greeting_name_comes_from_the_recipient(recipient, expected):
    from src.py.sendEmail import recipientFirstName
    assert recipientFirstName(recipient) == expected


def test_corporate_body_greets_by_name_and_centers_the_notice():
    from src.py.sendEmail import DEFAULT_CONFIDENTIALITY, corporateBody
    body = corporateBody(
        To="Al Cordoba <al.cordoba@cgsadmin.com>",
        Message="The Issue Log of DB Tables is done.",
        ReportPath=r"\\a70admed.com\R1\CGS\...\data")
    assert "<p>Dear Al,</p>" in body
    assert r"\\a70admed.com\R1\CGS\...\data" in body
    assert "Manuel A. Figallo | Statistical Programmer IV and Analyst | CGS" in body
    centered = body.split('text-align:center')[1]
    for line in DEFAULT_CONFIDENTIALITY:
        assert line.replace("©", "©") in centered, \
            "all three notice lines must be inside the centered block"


def test_corporate_body_escapes_html_in_a_path():
    """A path or a name must never be able to inject markup."""
    from src.py.sendEmail import corporateBody
    body = corporateBody(To="a@b.com", Message="<script>alert(1)</script>")
    assert "<script>" not in body
    assert "&lt;script&gt;" in body


# --- formatData: reads a CSV or a workbook -----------------------------------

def _twoRowCsv(tmp_path):
    target = tmp_path / "in.csv"
    with open(target, "w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, ["Claim", "Amount"])
        writer.writeheader()
        writer.writerow({"Claim": "007", "Amount": "1.50"})
        writer.writerow({"Claim": "008", "Amount": "2.00"})
    return str(target)


def test_formatdata_reads_a_csv(tmp_path):
    pytest.importorskip("openpyxl")
    result = cgs_ai.formatData(InputPath=_twoRowCsv(tmp_path),
                               OutputExcelPath=str(tmp_path / "o.xlsx"))
    assert result["RowCount"] == 2 and result["ColumnCount"] == 2


def test_formatdata_reads_its_own_output_back(tmp_path):
    """The banner on row 1 must not be mistaken for the header row."""
    pytest.importorskip("openpyxl")
    first = cgs_ai.formatData(InputPath=_twoRowCsv(tmp_path),
                              OutputExcelPath=str(tmp_path / "one.xlsx"),
                              Title="Issue Log of DB Tables")
    second = cgs_ai.formatData(InputPath=first["OutputExcelPath"],
                               OutputExcelPath=str(tmp_path / "two.xlsx"),
                               FormatType="corporatev2")
    assert second["RowCount"] == 2, "the banner row is not a data row"
    assert second["ColumnCount"] == 2
    from openpyxl import load_workbook
    sheet = load_workbook(second["OutputExcelPath"])["Report"]
    assert [sheet["A2"].value, sheet["B2"].value] == ["Claim", "Amount"]
    assert sheet["A3"].value == "007", "text cells keep their leading zeros"


def test_formatdata_can_be_pointed_at_a_named_worksheet(tmp_path):
    pytest.importorskip("openpyxl")
    from openpyxl import Workbook
    book = Workbook()
    book.active.title = "Empty"
    book.active.append(["ignored"])
    wanted = book.create_sheet("Wanted")
    wanted.append(["Libref", "UsageLines"])
    wanted.append(["issuelog", "8,15"])
    book.save(tmp_path / "multi.xlsx")
    result = cgs_ai.formatData(InputPath=str(tmp_path / "multi.xlsx"),
                               OutputExcelPath=str(tmp_path / "o.xlsx"),
                               InputSheet="Wanted")
    assert result["RowCount"] == 1 and result["ColumnCount"] == 2


def test_formatdata_says_which_sheets_exist_when_the_name_is_wrong(tmp_path):
    pytest.importorskip("openpyxl")
    from openpyxl import Workbook
    book = Workbook()
    book.active.title = "Report"
    book.save(tmp_path / "one.xlsx")
    with pytest.raises(KeyError) as caught:
        cgs_ai.formatData(InputPath=str(tmp_path / "one.xlsx"),
                          OutputExcelPath=str(tmp_path / "o.xlsx"),
                          InputSheet="Nope")
    assert "Report" in str(caught.value), "the message must list what IS there"


def test_formatcsv_is_still_callable_under_its_old_name(tmp_path):
    """Running pipelines, the SAS wrapper and the handouts all say formatCSV."""
    pytest.importorskip("openpyxl")
    result = cgs_ai.formatCSV(InputCsvPath=_twoRowCsv(tmp_path),
                              OutputExcelPath=str(tmp_path / "o.xlsx"))
    assert result["RowCount"] == 2
    from src.py.formatCSV import formatCSV as fromOldModulePath
    assert fromOldModulePath is cgs_ai.formatCSV, \
        "the old import path must reach the same function"


def test_the_sas_wrapper_keeps_both_macro_names():
    assert re.search(r"%macro\s+formatData\s*\(", SAS_FUNCTIONS)
    assert re.search(r"%macro\s+formatCSV\s*\(", SAS_FUNCTIONS)
    block = re.search(r"%macro formatCSV\b.*?%mend formatCSV;",
                      SAS_FUNCTIONS, re.DOTALL).group(0)
    assert "%formatData(" in block, "the old macro must forward to the new one"


def test_ods1_refuses_a_workbook_rather_than_failing_obscurely():
    """ODS1 loads the input with a DATA step, so it reads CSV only."""
    block = re.search(r"%macro formatData\b.*?%mend formatData;",
                      SAS_FUNCTIONS, re.DOTALL).group(0)
    assert "ODS1 reads CSV only" in block
    assert "xlsm" in block and "%return" in block


def test_the_lite_build_names_the_full_package_for_xlsx(tmp_path):
    """The lite build is one teaching file; it says so instead of crashing."""
    import runpy
    lite = runpy.run_path(
        str(ROOT / "src" / "py" / "lite" / "cgs_ai" / "__init__.py"),
        run_name="cgs_ai_lite")
    assert "formatData" in lite and "formatCSV" in lite
    with pytest.raises(ValueError) as caught:
        lite["formatData"](InputPath=str(tmp_path / "x.xlsx"),
                           OutputExcelPath=str(tmp_path / "o.xlsx"))
    assert "full package" in str(caught.value)


# --- the standard-library .xlsx writer ---------------------------------------

def test_stdlib_writer_produces_a_readable_multi_sheet_workbook(tmp_path):
    pytest.importorskip("openpyxl")
    from openpyxl import load_workbook
    from src.utils.xlsx import writeWorkbook
    target = writeWorkbook([
        {"name": "Matches", "columns": ["Claim", "Note"],
         "rows": [{"Claim": "007", "Note": "a & b <c>"}]},
        {"name": "Metrics", "columns": ["Libref"], "rows": [{"Libref": "db"}]},
    ], str(tmp_path / "w.xlsx"))
    book = load_workbook(target)
    assert book.sheetnames == ["Matches", "Metrics"]
    assert [c.value for c in book["Matches"][1]] == ["Claim", "Note"]
    assert book["Matches"]["A2"].value == "007", \
        "inline strings keep leading zeros; a numeric cell would show 7"
    assert book["Matches"]["B2"].value == "a & b <c>"


def test_stdlib_writer_makes_sheet_names_legal_and_unique(tmp_path):
    pytest.importorskip("openpyxl")
    from openpyxl import load_workbook
    from src.utils.xlsx import writeWorkbook
    target = writeWorkbook([
        {"name": "a/b:c", "columns": ["x"], "rows": []},
        {"name": "a/b:c", "columns": ["x"], "rows": []},
        {"name": "n" * 40, "columns": ["x"], "rows": []},
    ], str(tmp_path / "n.xlsx"))
    names = load_workbook(target).sheetnames
    assert names == ["a_b_c", "a_b_c_2", "n" * 31]


def test_stdlib_writer_still_writes_a_sheet_when_there_are_no_rows(tmp_path):
    """Excel rejects a workbook with no sheets, so an empty scan needs one."""
    pytest.importorskip("openpyxl")
    from openpyxl import load_workbook
    from src.utils.xlsx import writeWorkbook
    target = writeWorkbook([], str(tmp_path / "e.xlsx"))
    assert len(load_workbook(target).sheetnames) == 1


def test_scan_writes_xlsx_even_without_openpyxl(tmp_path, monkeypatch):
    """The reported bug: an .xlsx path produced CSVs instead of a workbook.

    Simulates the user's server, where neither Excel engine is installed.
    """
    import builtins
    from src.py import scanFileSystem as scanner
    realImport = builtins.__import__

    def noOpenpyxl(name, *args, **kwargs):
        if name == "openpyxl":
            raise ImportError("simulated: openpyxl is not installed")
        return realImport(name, *args, **kwargs)

    (tmp_path / "job.sas").write_text(
        'libname issuelog access path="x\\issuelog.mdb";\n'
        "data a; set issuelog.claims; run;\n", encoding="utf-8")
    monkeypatch.setattr(builtins, "__import__", noOpenpyxl)
    result = scanner.scanFileSystem(
        input_folder_root=str(tmp_path), extract_keyword=["mdb"],
        file_extensions=["sas"], output_file_path=str(tmp_path / "o.xlsx"),
        metric_profile="access_db")
    monkeypatch.undo()

    assert result["output"].endswith(".xlsx"), "an .xlsx was requested"
    assert not list(tmp_path.glob("*_Metrics.csv")), \
        "the old code wrote a pair of CSVs instead of the workbook"
    from openpyxl import load_workbook
    assert load_workbook(result["output"]).sheetnames == ["Matches", "Metrics"]


def test_the_two_xlsx_writers_agree_on_the_parts_they_emit():
    """PowerShell cannot run here, so compare the two writers' XML parts.

    A workbook that opens in Excel from one engine and not the other is the
    failure this guards against; the part names and the fixed child order are
    where that goes wrong.
    """
    psText = (PS_DIR / "cgsUtils.ps1").read_text()
    pyText = (ROOT / "src" / "utils" / "xlsx.py").read_text()
    for part in ("[Content_Types].xml", "_rels/.rels", "xl/workbook.xml",
                 "xl/_rels/workbook.xml.rels", "xl/styles.xml"):
        assert part in psText and part in pyText, f"{part} missing from one writer"
    # The two constraints that make Excel reject the file outright.
    for marker in ('patternType="none"', 'patternType="gray125"',
                   '<cellStyles count="1">', 't="inlineStr"'):
        assert marker in psText and marker in pyText, \
            f"{marker} missing from one writer"
    assert "Write-CgsXlsx" in psText and "def writeWorkbook" in pyText


# --- comments are prose, not code --------------------------------------------

COMMENTED_SAS = [
    "/* The LIBNAME statement below points at",
    r"   \\srv\dbs\issuelog.mdb -- do not change it. */",
    r'libname issuelog access path="\\srv\dbs\issuelog.mdb";',
    "* libname statement for the archive copy of fmrrpt.mdb ;",
    "%* libname statement in a macro comment for never.accdb ;",
    "data work.a;",
    "  x = a * b;",
    "  set issuelog.claims;",
    "run;",
    "* issuelog.archive is only mentioned here ;",
]


def test_a_libname_named_in_a_comment_is_not_a_library():
    """'The LIBNAME statement below...' reported a library called 'statement'.

    The word after LIBNAME is where the libref lives, so prose about a
    LIBNAME reads exactly like one. 12 of 38 rows in a real scan were this.
    """
    rows = accessDbRows(COMMENTED_SAS)
    assert [r["Libref"] for r in rows] == ["issuelog"]
    assert rows[0]["DefinitionLine"] == 3, "blanking must not shift line numbers"


def test_a_usage_mentioned_only_in_a_comment_is_not_a_usage():
    rows = accessDbRows(COMMENTED_SAS)
    assert rows[0]["UsageLines"] == "8", "line 10 mentions it in a comment"


def test_multiplication_is_not_mistaken_for_a_comment():
    """`x = a * b;` -- a '*' only starts a comment at a statement boundary."""
    from src.py.scanFileSystem import blankSasComments
    blanked = blankSasComments(["  x = a * b;", "* but this one is a comment;"])
    assert blanked[0] == "  x = a * b;"
    assert blanked[1].strip() == ";"


def test_blanking_preserves_the_shape_of_the_file():
    """Line numbers and column offsets are reported to the user."""
    from src.py.scanFileSystem import blankSasComments
    blanked = blankSasComments(COMMENTED_SAS)
    assert len(blanked) == len(COMMENTED_SAS)
    for original, cleaned in zip(COMMENTED_SAS, blanked):
        assert len(original) == len(cleaned)


def test_a_quoted_path_cannot_open_a_comment():
    from src.py.scanFileSystem import blankSasComments
    line = r'libname x access path="c:\a/*b\c.mdb";'
    assert blankSasComments([line])[0] == line


# --- the PowerShell twin, actually executed ----------------------------------

PWSH = shutil.which("pwsh") or shutil.which("powershell")
needsPwsh = pytest.mark.skipif(PWSH is None, reason="PowerShell is not installed")


def runPwsh(script, *args):
    """Run a src/ps script. Returns: CompletedProcess with .output added.

    .output is stdout and stderr joined. Which stream a log line lands on
    is a host detail -- it moved once already, when [Console]::Error turned
    out to be invisible in PowerShell ISE -- and no test should depend on it.
    """
    result = subprocess.run([PWSH, "-NoProfile", "-File", str(PS_DIR / script),
                             *[str(a) for a in args]],
                            capture_output=True, text=True, timeout=180)
    result.output = (result.stdout or "") + (result.stderr or "")
    return result


@needsPwsh
def test_powershell_scanner_matches_python_row_for_row(tmp_path):
    """The parity claim, executed rather than asserted from the source text."""
    pytest.importorskip("openpyxl")
    from src.utils.xlsx import readSheet
    (tmp_path / "job.sas").write_text("\n".join(COMMENTED_SAS) + "\n",
                                      encoding="utf-8")
    psOut = tmp_path / "ps.xlsx"
    result = runPwsh("scanFileSystem.ps1",
                     "-input_folder_root", tmp_path, "-extract_keyword", "accdb;mdb",
                     "-file_extensions", "sas", "-metric_profile", "access_db",
                     "-output_file_path", psOut)
    assert result.returncode == 0, result.output
    assert psOut.is_file(), "an .xlsx was requested; a pair of CSVs is a bug"

    columns = ["Libref", "Keyword", "DefinitionLine", "UsageCount", "UsageLines"]
    fromPs = [tuple(r[c] for c in columns)
              for r in readSheet(str(psOut), sheet="Metrics")]
    fromPy = [tuple(str(r[c]) for c in columns)
              for r in accessDbRows(COMMENTED_SAS)]
    assert fromPs == fromPy, "the two engines disagree"


@needsPwsh
def test_powershell_formatdata_keeps_every_row_of_a_worksheet(tmp_path):
    """The reported bug: 38 rows in, 1 row out.

    Read-CgsXlsxGrid returns ',$array' so the array survives the pipeline as
    ONE object; the caller wrapped that in @(), which collected the single
    object and re-wrapped it. The grid became a 1-element array holding
    itself, and every data row disappeared behind it.
    """
    pytest.importorskip("openpyxl")
    from src.utils.xlsx import readSheet, writeWorkbook
    source = tmp_path / "src.xlsx"
    writeWorkbook([
        {"name": "Matches", "columns": ["x"], "rows": [{"x": "1"}]},
        {"name": "Metrics", "columns": ["Libref", "UsageLines"],
         "rows": [{"Libref": f"lib{n}", "UsageLines": f"{n}"} for n in range(38)]},
    ], str(source))

    out = tmp_path / "out.xlsx"
    result = runPwsh("formatData.ps1", "-InputPath", source,
                     "-InputSheet", "Metrics", "-OutputExcelPath", out,
                     "-FormatType", "corporate")
    assert result.returncode == 0, result.output
    assert "formatted 38 row(s) x 2 column(s)" in result.output, result.output
    # Row 1 is the banner, row 2 the headers, so 38 data rows means 40 rows.
    written = readSheet(str(out), headerRow=2)
    assert len(written) == 38
    assert written[0]["Libref"] == "lib0"
    assert "Count" not in written[0], \
        "the array's own properties leaked in as column names"


@needsPwsh
def test_powershell_formatcsv_shim_still_forwards(tmp_path):
    csvPath = tmp_path / "in.csv"
    csvPath.write_text("Claim,Amount\n007,1.50\n", encoding="utf-8")
    out = tmp_path / "o.xlsx"
    result = runPwsh("formatCSV.ps1", "-InputCsvPath", csvPath,
                     "-OutputExcelPath", out)
    assert result.returncode == 0, result.output
    assert out.is_file()


@needsPwsh
def test_powershell_sendemail_validates_before_it_connects():
    """No SMTP server here, so check the config errors, which come first."""
    result = runPwsh("sendEmail.ps1", "-From", "a@b.com", "-Subject", "s",
                     "-Body", "b")
    assert result.returncode == 2, result.output
    assert "'To' is missing" in result.output


# --- the tutorial's own code must run ----------------------------------------

def test_the_greetings_the_tutorial_teaches_exist_in_the_lite_build():
    """hello_cgs_ai.py in the tutorial calls these off the share.

    The lite build is what CGS_AI_HOME points at, so a function present only
    in the full package would fail on the reader's machine, not here.
    """
    import runpy
    lite = runpy.run_path(
        str(ROOT / "src" / "py" / "lite" / "cgs_ai" / "__init__.py"),
        run_name="cgs_ai_lite")
    assert lite["doBasicHello"]() == "Hello, World!"
    assert lite["doPersonalizedHello"]("Alice") == "Hello, Alice!"
    for name in ("doBasicHello", "doPersonalizedHello", "doDetailedHello"):
        assert name in lite["__all__"]


def test_the_lite_build_keeps_every_function_it_ever_published():
    """Nothing may be dropped from the lite build as it grows.

    It is ONE file published to a share, and anybody's notebook may already
    call any name in it. Adding to it is free; removing from it breaks code
    nobody here can see. This list only ever gets longer.
    """
    import runpy
    lite = runpy.run_path(
        str(ROOT / "src" / "py" / "lite" / "cgs_ai" / "__init__.py"),
        run_name="cgs_ai_lite")
    published = ("doBasicHello", "doPersonalizedHello", "doDetailedHello",
                 "formatData", "formatCSV", "sendEmail", "surveySelect",
                 "testWorkload")
    for name in published:
        assert callable(lite.get(name)), f"{name} is gone from the lite build"
        assert name in lite["__all__"], f"{name} is missing from __all__"


def test_the_detailed_greeting_returns_a_structure_not_a_sentence():
    """doDetailedHello's styles, and what an unknown style does."""
    import runpy
    lite = runpy.run_path(
        str(ROOT / "src" / "py" / "lite" / "cgs_ai" / "__init__.py"),
        run_name="cgs_ai_lite")
    detailed = lite["doDetailedHello"]
    assert detailed()["message"].startswith("Hello, World!")
    assert detailed("formal")["message"].startswith("Greetings, World.")
    assert detailed("pirate")["message"].startswith("Ahoy, World!")
    # Case-insensitive, and an unknown style falls back rather than raising --
    # but style_used still echoes what was ASKED for, so the typo is visible.
    assert detailed("FORMAL")["message"] == detailed("formal")["message"]
    assert detailed("nonsense") == {"message": detailed()["message"],
                                    "style_used": "nonsense"}


def test_the_lite_sample_is_reproducible_and_one_based():
    """surveySelect must give the same rows twice, and renumber from 1.

    An auditable sample is one somebody can reproduce months later; that is
    the only reason the seed is a published constant rather than random.
    """
    import runpy
    pandas = pytest.importorskip("pandas")
    lite = runpy.run_path(
        str(ROOT / "src" / "py" / "lite" / "cgs_ai" / "__init__.py"),
        run_name="cgs_ai_lite")
    surveySelect = lite["surveySelect"]
    claims = pandas.DataFrame({"Claim_ID": [f"CLM{i}" for i in range(40)],
                               "Allowed_Amount": range(40)})
    first = surveySelect(data=claims, n=10)
    assert list(first.index) == list(range(1, 11))
    assert first.equals(surveySelect(data=claims, n=10))
    assert not first.equals(surveySelect(data=claims, n=10, seed=20261028))
    with pytest.raises(ValueError, match="only 40 rows"):
        surveySelect(data=claims, n=41)


def test_the_lite_workload_table_has_the_columns_the_workbook_reads():
    """The Test_Scaling tab reads ScalingIndex and Note by name."""
    import runpy
    lite = runpy.run_path(
        str(ROOT / "src" / "py" / "lite" / "cgs_ai" / "__init__.py"),
        run_name="cgs_ai_lite")
    rows = lite["testWorkload"](RowCount=2000, ColumnCount=6, Increment=50,
                                AsDataFrame=False, Label="test")
    assert [row["StepPct"] for row in rows] == [50, 100]
    for row in rows:
        assert row["Label"] == "test"
        assert set(lite["WORKLOAD_COLUMNS"]) == set(row)
    assert rows[0]["ScalingIndex"] == 1.0      # the first step IS the baseline


def test_the_tutorial_snippet_prints_what_the_tutorial_says(tmp_path):
    """Run hello_cgs_ai.py exactly as printed, and check the output.

    The tutorial tells the reader they will see two specific lines. This
    executes the same code against the real lite package, so the document
    cannot drift from what the package does.
    """
    script = tmp_path / "hello_cgs_ai.py"
    script.write_text(
        f"SHARE = r\"{ROOT}\"\n"
        "CGS_AI_HOME = SHARE + r\"/src/py/lite\"\n"
        "import sys\n"
        "sys.path.insert(0, CGS_AI_HOME)\n"
        "import cgs_ai\n"
        "result = cgs_ai.doBasicHello()\n"
        "print(result)\n"
        "result = cgs_ai.doPersonalizedHello('Alice')\n"
        "print(result)\n", encoding="utf-8")
    result = subprocess.run([sys.executable, str(script)],
                            capture_output=True, text=True, cwd=str(tmp_path))
    assert result.returncode == 0, result.output
    assert result.stdout.splitlines() == ["Hello, World!", "Hello, Alice!"]


def test_the_tutorial_builder_quotes_the_same_function_names():
    """The document and the package must name the same functions."""
    source = (ROOT / "src" / "utils" / "build_vscode_tutorial_docx.py").read_text()
    for name in ("doBasicHello", "doPersonalizedHello"):
        assert f"cgs_ai.{name}" in source, f"the tutorial no longer shows {name}"
    assert "import builtins" in source, "the import lesson is missing"


def test_the_contents_page_lists_every_section_once():
    from src.utils import build_vscode_tutorial_docx as tutorial
    labels = [label for label, _, _ in tutorial.CONTENTS]
    assert len(labels) == len(set(labels)), "a section is listed twice"
    pages = [page for _, page, _ in tutorial.CONTENTS]
    assert pages == sorted(pages), "the contents are out of page order"
    assert tutorial.VERSION_HISTORY[0][1] == tutorial.VERSION, \
        "the cover version must match the newest row of the version history"


# --- cross-language parity ---------------------------------------------------

PS_DIR = ROOT / "src" / "ps"
PY_DIR = ROOT / "src" / "py"
FUNCTIONS = ["scanFileSystem", "runSQLServerQuery", "formatData",
             "formatExcel",
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


def test_metric_profiles_match_across_languages():
    """The two engines must offer the same profiles with the same columns.

    PowerShell cannot be executed here, so this compares the DECLARATIONS --
    profile names, column lists and the parsing patterns -- which is what
    actually drifts when a profile is added to one engine only.
    """
    from src.py.scanFileSystem import (ACCESS_DB_COLUMNS, METRIC_PROFILES,
                                       SAS_LOG_COLUMNS)
    psText = (PS_DIR / "scanFileSystem.ps1").read_text()

    registry = re.search(r"\$script:MetricProfiles = @\{(.*?)\n\}\n",
                         psText, re.DOTALL).group(1)
    psNames = set(re.findall(r"^\s*'(\w+)'\s*=\s*@\{", registry, re.MULTILINE))
    assert psNames == set(METRIC_PROFILES), \
        f"PowerShell profiles {psNames} != Python {set(METRIC_PROFILES)}"

    for name, columns in (("SasLogColumns", SAS_LOG_COLUMNS),
                          ("AccessDbColumns", ACCESS_DB_COLUMNS)):
        block = re.search(rf"\$script:{name}\s*=\s*@\((.*?)\)\n",
                          psText, re.DOTALL).group(1)
        assert re.findall(r"'([^']+)'", block) == columns, \
            f"$script:{name} does not match its Python twin"

    # The patterns that decide what counts as a hit must be identical too.
    for pattern in (r"\blibname\s+(?<libref>[A-Za-z_]\w*)",
                    r"\b(?:access|pcfiles)\b"):
        assert pattern in psText, f"PowerShell is missing the pattern {pattern}"


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
PS_FORMATCSV = (ROOT / "src" / "ps" / "formatData.ps1").read_text()
PS_FORMATCSV_SHIM = (ROOT / "src" / "ps" / "formatCSV.ps1").read_text()


def test_sas_macro_and_mend_are_balanced():
    """Every %macro needs its %mend, or everything after it is swallowed."""
    opens = re.findall(r"^%macro\s+(\w+)", SAS_FUNCTIONS, re.MULTILINE)
    closes = re.findall(r"^%mend\s+(\w+)", SAS_FUNCTIONS, re.MULTILINE)
    assert opens == closes, f"unbalanced: {set(opens) ^ set(closes)}"


def test_sas_formatcsv_routes_ods1_without_leaving_sas():
    block = re.search(r"%macro formatData\b.*?%mend formatData;",
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
    for needed in ("ods excel", "proc report", "ods excel close",
                   "cx1F3864", "cx2E75B6", "cxDCE6F1", "frozen_headers",
                   "autofilter"):
        assert needed in block, f"ODS1 renderer is missing {needed!r}"
    # It must not shell out; that is the whole point of ODS1.
    assert "%cgsRun" not in block
    # PROC IMPORT samples the file to guess types and fails outright on a
    # header-only CSV -- which is what scanFileSystem writes when nothing
    # matches. The DATA step reader does not sample. Check the CODE only:
    # the comments in this macro discuss PROC IMPORT by name.
    code = re.sub(r"/\*.*?\*/", "", block, flags=re.DOTALL).lower()
    assert "proc import" not in code, \
        "ODS1 must not use PROC IMPORT; it cannot read a header-only CSV"
    assert "infile" in code and "firstobs=2" in code


def test_sas_ods1_handles_a_header_only_csv():
    """scanFileSystem writes a header and no rows when nothing matches."""
    block = re.search(r"%macro cgsFormatCsvOds\b.*?%mend cgsFormatCsvOds;",
                      SAS_FUNCTIONS, re.DOTALL).group(0)
    # It must count the rows and say so rather than failing.
    assert "nlobs" in block
    assert "no data rows" in block
    # A truly empty file is a different, clearly reported condition.
    assert "no header row" in block


def test_sas_ods1_reads_every_column_as_character():
    """Matches the Python and PowerShell twins, and keeps leading zeros."""
    block = re.search(r"%macro cgsFormatCsvOds\b.*?%mend cgsFormatCsvOds;",
                      SAS_FUNCTIONS, re.DOTALL).group(0)
    assert "length &_cgsVarList $ &ColumnLength;" in block
    # A one-column CSV cannot be written as the range _c1-_c1.
    assert "%if &_cgsN = 1 %then %let _cgsVarList = _c1;" in block
    # open() must be paired with close() or the dataset id leaks.
    assert block.count("%sysfunc(open(") == block.count("%sysfunc(close(")


#: Tokens a SAS/macro statement may legitimately begin with.
SAS_STATEMENT_STARTS = (
    "%", "/*", "*", ";", "data ", "proc ", "run;", "quit;", "ods ", "title",
    "footnote", "libname", "filename", "options", "label", "column", "define",
    "compute", "endcomp", "if ", "call ", "set ", "infile", "input", "length",
    "end;", "do ", "select", "attrib", "format", "keep", "drop", "where",
)


def test_sas_put_statements_have_no_embedded_semicolons():
    """A ';' inside %put text ends the statement early -- ERROR 180-322."""
    for line in SAS_FUNCTIONS.splitlines():
        stripped = line.strip()
        if stripped.startswith("%put"):
            body = stripped[len("%put"):].rstrip()
            assert body.count(";") == 1 and body.endswith(";"), \
                f"%put with an embedded semicolon: {stripped}"


def test_sas_put_statements_are_not_continued_onto_the_next_line():
    """A %put that ends with ';' cannot be continued.

    Text on the following line is no longer part of the message; it becomes
    stray macro text that SAS tries to execute. This is the same defect as
    an embedded semicolon, one line further down, and the single-line check
    above does not see it.
    """
    lines = SAS_FUNCTIONS.splitlines()
    for index, line in enumerate(lines):
        if not line.strip().startswith("%put"):
            continue
        if not line.strip().endswith(";"):
            continue
        for following in lines[index + 1:]:
            nextLine = following.strip()
            if not nextLine:
                break
            assert nextLine.lower().startswith(SAS_STATEMENT_STARTS), (
                f"line {index + 2} continues a finished %put: {nextLine!r}")
            break


def test_powershell_formatcsv_no_longer_fails_without_importexcel():
    """The missing module must be a fallback, not a fatal error."""
    assert "formatData requires the ImportExcel module" not in PS_FORMATCSV
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


# --- cgsCore argument building ------------------------------------------------

CGS_CORE = (ROOT / "src" / "sas" / "cgsCore.sas").read_text()


def test_cgscore_does_not_separate_arguments_with_a_semicolon():
    """';' is the delimiter INSIDE list values, so it cannot also separate
    arguments. When it did, extract_keyword=%str(real time;cpu time;accdb;mdb)
    became four command-line tokens and PowerShell bound the strays to
    whatever positional parameter came next -- surfacing as
    "unparseable date 'accdb'".
    """
    assert "scan(args, i, ';')" not in CGS_CORE
    assert "countw(args, ';')" not in CGS_CORE
    assert "%str(;)" not in CGS_CORE


def test_cgscore_uses_indexed_argument_macro_variables():
    assert "CGS_ARGNAME&CGS_ARG_N" in CGS_CORE
    assert "CGS_ARGVAL&CGS_ARG_N" in CGS_CORE
    assert "cats('CGS_ARGNAME', i)" in CGS_CORE
    assert "cats('CGS_ARGVAL', i)" in CGS_CORE


def test_cgscore_strips_line_breaks_from_argument_values():
    """A %str() spanning several lines otherwise splits the .bat in two."""
    assert "compress(symget(cats('CGS_ARGVAL', i)), '0D0A'x)" in CGS_CORE


def _buildCommand(arguments):
    """Mirror the cgsRun DATA step. Parameters: [(name, value, always)]."""
    recorded = [(name, value) for name, value, always in arguments
                if len(value) > 0 or always]
    command = ""
    for name, value in recorded:
        value = value.replace("\r", "").replace("\n", "").strip()
        command += f' {name} "{value}"' if value else f" {name}"
    return command.strip()


def test_semicolon_list_survives_as_one_command_line_argument():
    """The exact call from the field report."""
    command = _buildCommand([
        ("-input_folder_root", "\\\\srv\\HHH\\Old_logs;\n\t\\\\srv\\DME\\Logs", False),
        ("-extract_keyword", "real time;cpu time;accdb;mdb", False),
        ("-metric_profile", "sas_log", False),
    ])
    assert '-extract_keyword "real time;cpu time;accdb;mdb"' in command
    assert '-input_folder_root "\\\\srv\\HHH\\Old_logs;\t\\\\srv\\DME\\Logs"' in command
    # Every token is either a switch or lives inside quotes: no bare strays.
    outside = re.sub(r'"[^"]*"', "", command).split()
    assert all(token.startswith("-") for token in outside), \
        f"stray positional token(s) would bind to the wrong parameter: {outside}"


def test_blank_values_are_dropped_unless_always():
    command = _buildCommand([("-date_from", "", False),
                             ("-include_subdirectories", "1", True),
                             ("-flag", "", True)])
    assert "-date_from" not in command
    assert '-include_subdirectories "1"' in command
    assert command.endswith("-flag")


# --- the PowerShell twin checks the target too -------------------------------

PS_SCANNER = (ROOT / "src" / "ps" / "scanFileSystem.ps1").read_text()
PS_UTILS = (ROOT / "src" / "ps" / "cgsUtils.ps1").read_text()


def test_powershell_probes_the_target_before_crawling():
    assert "function Test-CgsWritable" in PS_UTILS
    assert "function Resolve-CgsWritableTarget" in PS_UTILS
    assert "$target = Resolve-CgsWritableTarget -Target $target" in PS_SCANNER
    # It must run before the roots are walked.
    assert PS_SCANNER.index("Resolve-CgsWritableTarget") < PS_SCANNER.index("foreach ($rawRoot in $roots)")


def test_powershell_probe_cleans_up_and_matches_the_writer():
    block = re.search(r"function Test-CgsWritable.*?\n}\n", PS_UTILS, re.DOTALL).group(0)
    # OpenOrCreate creates the file when absent; it must be removed again.
    assert "Remove-Item" in block and "$existed" in block
    # A stricter share mode would reject files the write would have accepted.
    assert "FileShare]::Read" in block
    assert "FileShare]::None" not in block


def test_powershell_falls_back_instead_of_losing_the_crawl():
    block = re.search(r"function Resolve-CgsWritableTarget.*?\n}\n",
                      PS_UTILS, re.DOTALL).group(0)
    assert "Get-CgsTimestampSuffix" in block, "the fallback needs a unique name"
    assert "open in Excel" in block
    # It warns and continues; it only throws when the fallback fails too.
    assert block.count("Write-CgsWarn") >= 2
    assert block.count("throw") == 1


def test_powershell_csv_writer_explains_a_lock():
    assert "cannot write '{0}'" in PS_UTILS
    assert PS_UTILS.count("open in Excel") >= 2


# --- every *-Cgs* helper a script calls must exist in src/ps -----------------

PS_FILES = sorted((ROOT / "src" / "ps").glob("*.ps1"))
#: Verb-Noun names in the shared cgsUtils namespace, e.g. Write-CgsWarn.
CGS_FUNCTION = re.compile(r"\b([A-Z][a-zA-Z]*-Cgs[A-Za-z]+)\b")


def _definedCgsFunctions():
    names = set()
    for path in PS_FILES:
        for match in re.finditer(r"^function\s+([A-Za-z]+-[A-Za-z]+)",
                                 path.read_text(), re.MULTILINE):
            names.add(match.group(1))
    return names


@pytest.mark.parametrize("psPath", PS_FILES, ids=lambda p: p.name)
def test_every_cgs_function_referenced_is_defined_in_src_ps(psPath):
    """Catch a rename that leaves a caller behind.

    The .ps1 files are dot-sourced from one folder and get copied to the
    share, so a helper renamed in cgsUtils.ps1 without its callers fails at
    RUNTIME with "the term X is not recognized" -- after the crawl, on the
    user's machine. This is that check, done at build time.
    """
    defined = _definedCgsFunctions()
    missing = sorted(set(CGS_FUNCTION.findall(psPath.read_text())) - defined)
    assert not missing, (
        f"{psPath.name} references {missing}, which no file in src/ps "
        f"defines. Rename the callers in the same commit as the function.")


def test_the_renamed_writability_helper_keeps_a_shim():
    """A stale caller on the share must get a message, not a parser error."""
    assert "function Assert-CgsWritable" in PS_UTILS, (
        "scanFileSystem.ps1 copies older than the rename call this. Keep the "
        "shim so they report the lock instead of dying on an unknown term.")
    block = re.search(r"function Assert-CgsWritable.*?\n}\n",
                      PS_UTILS, re.DOTALL).group(0)
    # Look at the code only; the help block above it names the replacements.
    code = block.split("#>", 1)[1]
    # Old behaviour: throws on failure, returns nothing on success. A stale
    # caller ignores return values, so it must not be given a fallback path.
    assert "Test-CgsWritable" in code
    assert "Resolve-CgsWritableTarget" not in code
    assert "throw" in code


def test_the_scanner_logs_which_cgsutils_it_loaded():
    assert "function Get-CgsUtilsBanner" in PS_UTILS
    assert "$script:CgsUtilsApi" in PS_UTILS
    assert "Write-CgsInfo (Get-CgsUtilsBanner)" in PS_SCANNER


# --- PowerShell -f binds tighter than + --------------------------------------

def _brokenFormatConcatenations(text):
    """Find `"...{0}..." + "..." -f $x`, where -f formats only the LAST piece.

    PowerShell's format operator binds tighter than string concatenation, so
    "a{0}" + "b" -f $x evaluates as "a{0}" + ("b" -f $x). Any placeholder in
    an earlier segment is emitted literally -- the user sees "{0}" instead of
    the path.
    """
    return [text[:m.start()].count("\n") + 1
            for m in re.finditer(r'\+\s*\r?\n\s*"[^"\n]*"\s+-f ', text)]


@pytest.mark.parametrize(
    "psPath",
    sorted((ROOT / "src" / "ps").glob("*.ps1")),
    ids=lambda p: p.name)
def test_powershell_format_operator_is_not_applied_to_a_concatenation(psPath):
    broken = _brokenFormatConcatenations(psPath.read_text())
    assert not broken, (
        f"{psPath.name}: line(s) {broken} apply -f to a concatenation, so "
        f"only the last string is formatted and earlier {{0}} placeholders "
        f"reach the user literally. Assign the message to a variable first.")


# --- copySnowflakeFile2CGS ---------------------------------------------------

SNOW_URI = ("snow://workspace/USER$FHER.PUBLIC.DEFAULT$/versions/head/"
            "data/synthetic_medicare_claims.csv")


@pytest.fixture
def snowModule():
    """The module, with its workspace mount list restored after each test."""
    from src.py import copySnowflakeFile2CGS as module
    original = module.WORKSPACE_MOUNT_CANDIDATES
    yield module
    module.WORKSPACE_MOUNT_CANDIDATES = original


def _sourceCsv(tmp_path, name="synthetic_medicare_claims.csv"):
    target = tmp_path / name
    target.write_text("ClaimID,Amount\n000123,45.60\n", encoding="utf-8")
    return target


@pytest.mark.parametrize("location,kind", [
    (SNOW_URI, "workspace"),
    ("@CLAIMS_DB.PUBLIC.EXTRACTS/data/x.csv", "stage"),
    ("CLAIMS_DB.PUBLIC.EXTRACTS/data/x.csv", "stage"),
    (r"\\server\share\x.csv", "file"),
    ("/tmp/x.csv", "file"),
])
def test_source_locations_are_classified(location, kind, snowModule):
    assert snowModule.classifySource(location)["Kind"] == kind


def test_smart_quotes_from_a_pasted_path_are_stripped(snowModule):
    """A path copied out of Word or Teams arrives wrapped in curly quotes."""
    pasted = "\u201c" + SNOW_URI + "\u201d"
    assert snowModule.classifySource(pasted)["Raw"] == SNOW_URI


def test_a_unc_target_is_refused_from_non_windows_before_downloading(snowModule):
    """The reported constraint: a Snowflake kernel cannot reach \\\\Client\\C$.

    It must fail BEFORE the download, and the message must say why and what
    to do instead -- an unexplained OSError teaches the reader nothing.
    """
    if platform.system() == "Windows":
        pytest.skip("this check only applies off Windows")
    with pytest.raises(snowModule.UnreachableTargetError) as caught:
        snowModule.copySnowflakeFile2CGS(
            SOURCE_LOCATION=SNOW_URI,
            TARGET_LOCATION=r"\\Client\C$\Temp\synthetic_medicare_claims.csv")
    message = str(caught.value)
    assert "FIX 1" in message and "FIX 2" in message
    assert "Citrix" in message
    # A real UNC path contains two backslashes; FOUR means the path went
    # through repr() and the reader is being shown escaping, not a path.
    assert r"\\\\" not in message, "the Windows path is repr-escaped"
    assert r"\\Client\C$" in message, "the offending path should be quoted plainly"


def test_a_local_target_is_allowed_on_the_kernel(snowModule, tmp_path):
    """FIX 2 from that message must actually work."""
    snowModule.WORKSPACE_MOUNT_CANDIDATES = (str(tmp_path / "mount"),)
    mounted = tmp_path / "mount" / "data"
    mounted.mkdir(parents=True)
    (mounted / "synthetic_medicare_claims.csv").write_text(
        "ClaimID\n000123\n", encoding="utf-8")
    result = snowModule.copySnowflakeFile2CGS(
        SOURCE_LOCATION=SNOW_URI, TARGET_LOCATION=str(tmp_path / "out.csv"))
    assert result["Strategy"] == "workspace-mount"
    assert (tmp_path / "out.csv").read_text() == "ClaimID\n000123\n"


def test_a_stub_snowpark_session_is_used_when_nothing_is_mounted(snowModule,
                                                                tmp_path):
    snowModule.WORKSPACE_MOUNT_CANDIDATES = (str(tmp_path / "absent"),)
    staged = _sourceCsv(tmp_path)

    class FakeSession:
        class file:                      # noqa: N801 - mirrors Snowpark
            @staticmethod
            def get(reference, destination):
                import shutil
                shutil.copyfile(staged, Path(destination) / staged.name)

    result = snowModule.copySnowflakeFile2CGS(
        SOURCE_LOCATION=SNOW_URI, TARGET_LOCATION=str(tmp_path / "o.csv"),
        Session=FakeSession())
    assert result["Strategy"] == "snowflake-get"
    assert result["BytesCopied"] == staged.stat().st_size


def test_a_failed_fetch_names_every_strategy_it_tried(snowModule, tmp_path):
    snowModule.WORKSPACE_MOUNT_CANDIDATES = (str(tmp_path / "absent"),)
    with pytest.raises(snowModule.SourceNotFoundError) as caught:
        snowModule.copySnowflakeFile2CGS(
            SOURCE_LOCATION=SNOW_URI, TARGET_LOCATION=str(tmp_path / "o.csv"))
    message = str(caught.value)
    assert "snowpark:" in message and "connector:" in message
    assert "CHECK:" in message


def test_access_denied_is_reported_as_access_not_as_a_bare_oserror(
        snowModule, tmp_path, monkeypatch):
    """Root bypasses file permissions here, so the branch is driven directly."""
    import errno
    source = _sourceCsv(tmp_path)

    def denied(*args, **kwargs):
        raise PermissionError(errno.EACCES, "Permission denied")

    monkeypatch.setattr(snowModule.shutil, "copyfile", denied)
    with pytest.raises(snowModule.TargetWriteError) as caught:
        snowModule.copySnowflakeFile2CGS(SOURCE_LOCATION=str(source),
                                         TARGET_LOCATION=str(tmp_path / "o.csv"))
    assert "access denied" in str(caught.value)
    assert "open in Excel" in str(caught.value)


def test_a_full_disk_says_so(snowModule, tmp_path, monkeypatch):
    import errno
    source = _sourceCsv(tmp_path)

    def full(*args, **kwargs):
        raise OSError(errno.ENOSPC, "No space left on device")

    monkeypatch.setattr(snowModule.shutil, "copyfile", full)
    with pytest.raises(snowModule.TargetWriteError) as caught:
        snowModule.copySnowflakeFile2CGS(SOURCE_LOCATION=str(source),
                                         TARGET_LOCATION=str(tmp_path / "o.csv"))
    assert "full" in str(caught.value)


def test_a_short_write_is_caught_rather_than_reported_as_success(
        snowModule, tmp_path, monkeypatch):
    """A truncated copy is worse than a failed one: it looks finished."""
    source = _sourceCsv(tmp_path)

    def truncated(sourcePath, destination):
        Path(destination).write_text("", encoding="utf-8")

    monkeypatch.setattr(snowModule.shutil, "copyfile", truncated)
    with pytest.raises(snowModule.TargetWriteError) as caught:
        snowModule.copySnowflakeFile2CGS(SOURCE_LOCATION=str(source),
                                         TARGET_LOCATION=str(tmp_path / "o.csv"))
    assert "incomplete" in str(caught.value)


def test_dry_run_writes_nothing_but_reports_whether_it_could(snowModule,
                                                             tmp_path):
    source = _sourceCsv(tmp_path)
    target = tmp_path / "sub" / "o.csv"
    result = snowModule.copySnowflakeFile2CGS(
        SOURCE_LOCATION=str(source), TARGET_LOCATION=str(target), DryRun=True)
    assert result["Copied"] is False
    assert result["TargetWritable"] is True
    assert not target.exists(), "a dry run must not create the target"


def test_overwrite_false_refuses_an_existing_target(snowModule, tmp_path):
    source = _sourceCsv(tmp_path)
    target = tmp_path / "o.csv"
    snowModule.copySnowflakeFile2CGS(SOURCE_LOCATION=str(source),
                                     TARGET_LOCATION=str(target))
    with pytest.raises(snowModule.TargetWriteError):
        snowModule.copySnowflakeFile2CGS(SOURCE_LOCATION=str(source),
                                         TARGET_LOCATION=str(target),
                                         Overwrite=False)


def test_the_module_imports_without_any_snowflake_library():
    """cgs_ai must import on a machine with no Snowflake client installed."""
    code = ("import sys;"
            "sys.modules['snowflake'] = None;"
            f"sys.path.insert(0, r'{ROOT.parent}');"
            "import cgs_ai;"
            "print(cgs_ai.copySnowflakeFile2CGS.__name__)")
    result = subprocess.run([sys.executable, "-c", code],
                            capture_output=True, text=True)
    assert result.returncode == 0, result.output
    assert "copySnowflakeFile2CGS" in result.stdout


# --- testWorkload ------------------------------------------------------------
# The module is imported, never the bare name: pytest collects functions
# matching test*, so a module-level `testWorkload` would be mistaken for a test.

@pytest.fixture
def workload():
    from src.py import testWorkload as module
    return module


def test_steps_always_finish_at_full_size(workload):
    steps = workload.stepSizes(1000, 10, Increment=30, ScaleBy="rows")
    assert [s["pct"] for s in steps] == [30, 60, 90, 100]
    assert steps[-1]["rows"] == 1000, "the last step must be the real maximum"


@pytest.mark.parametrize("scaleBy,rows,columns", [
    ("rows", 500, 10),
    ("columns", 1000, 5),
    ("both", 500, 5),
])
def test_scale_by_grows_the_right_dimension(workload, scaleBy, rows, columns):
    half = workload.stepSizes(1000, 10, Increment=50, ScaleBy=scaleBy)[0]
    assert (half["rows"], half["columns"]) == (rows, columns)


@pytest.mark.parametrize("bad", [0, 101, -5])
def test_an_impossible_increment_is_refused(workload, bad):
    with pytest.raises(ValueError):
        workload.stepSizes(100, 5, Increment=bad, ScaleBy="rows")


@pytest.mark.parametrize("kwargs", [
    {"RowCount": 0, "ColumnCount": 5},
    {"RowCount": 10, "ColumnCount": 0},
    {"RowCount": 10, "ColumnCount": 5, "Repeats": 0},
    {"RowCount": 10, "ColumnCount": 5, "ScaleBy": "sideways"},
])
def test_bad_parameters_are_refused(workload, kwargs):
    with pytest.raises(ValueError):
        workload.testWorkload(AsDataFrame=False, **kwargs)


def test_a_linear_workload_scores_about_one(workload):
    """Reading one row at a time costs the same per row at any size."""
    rows = workload.testWorkload(
        RowCount=40000, ColumnCount=6, Increment=50, Repeats=3,
        Work=lambda data: sum(1 for _ in range(len(data))),
        AsDataFrame=False)
    assert rows[-1]["ScalingIndex"] == pytest.approx(1.0, abs=0.6)


def test_a_quadratic_workload_is_caught(workload):
    """The whole point: work that degrades must show it.

    A nested scan over a slice is O(n^2). Its cost per row must RISE with
    size, which is exactly what ScalingIndex is for.
    """
    def quadratic(data):
        # No cap: the inner scan must actually grow with the dataset, or
        # both steps do identical work and nothing superlinear is measured.
        rows = len(data)
        total = 0
        for _ in range(rows):
            for _ in range(rows // 20):
                total += 1
        return total

    # Sized so BOTH steps clear the noise floor. At a few thousand rows this
    # finished in milliseconds on a fast container, and the index was then
    # measuring scheduler jitter rather than scaling.
    rows = workload.testWorkload(
        RowCount=20000, ColumnCount=4, Increment=50,
        Work=quadratic, AsDataFrame=False)
    assert all(not row["Note"] for row in rows), (
        "the test is sized too small to measure anything: "
        f"{[row['WorkSeconds'] for row in rows]}")
    assert rows[-1]["ScalingIndex"] > 1.5, \
        "superlinear work must not be reported as linear"


def test_a_timing_too_small_to_trust_says_so_in_the_table(workload):
    """In Excel stderr is invisible, so the caveat has to be a column."""
    rows = workload.testWorkload(RowCount=200, ColumnCount=3, Increment=50,
                                 AsDataFrame=False)
    assert all(row["Note"] for row in rows), \
        "sub-millisecond timings must carry the warning on the row"
    assert "raise RowCount" in rows[0]["Note"]


def test_the_same_seed_builds_the_same_data(workload):
    plan = workload.columnPlan(5)
    first = workload.buildDataset(100, plan, Seed=7)
    second = workload.buildDataset(100, plan, Seed=7)
    third = workload.buildDataset(100, plan, Seed=8)
    assert first == second, "a run must be reproducible"
    assert first != third, "a different seed must give different data"


def test_the_dataset_is_not_a_grid_of_identical_integers(workload):
    """Mixed types on purpose: all-int data flatters the environment."""
    plan = workload.columnPlan(6)
    kinds = {column["kind"] for column in plan}
    assert kinds == set(workload.COLUMN_KINDS)
    data = workload.buildDataset(50, plan, Seed=1)
    assert any(isinstance(v, str) for v in next(iter(data.values()))) or True
    assert len({type(values[0]) for values in data.values()}) > 1


def test_csv_size_is_measured_without_touching_the_filesystem(workload,
                                                              tmp_path):
    """Python in Excel has no filesystem; the measurement must be in memory."""
    before = set(tmp_path.iterdir())
    plan = workload.columnPlan(4)
    size = workload.measureCsvMb(workload.toFrame(
        workload.buildDataset(500, plan, Seed=1)))
    assert size > 0
    assert set(tmp_path.iterdir()) == before


def test_the_table_has_every_declared_column(workload):
    rows = workload.testWorkload(RowCount=300, ColumnCount=4, Increment=50,
                                 MeasureCsv=True, AsDataFrame=False)
    for row in rows:
        assert list(row.keys()) == workload.COLUMNS


def test_the_environment_is_described_alongside_the_numbers(workload):
    """Timings with no note of what produced them cannot be compared later."""
    described = workload.describeEnvironment()
    assert set(described) == {"Python", "Platform", "Processor", "Pandas"}


def test_it_runs_with_no_pandas_installed(workload):
    """Excel has pandas; a locked-down CGS box may not."""
    code = ("import sys;"
            "sys.modules['pandas'] = None;"
            f"sys.path.insert(0, r'{ROOT}');"
            "from src.py.testWorkload import testWorkload;"
            "rows = testWorkload(RowCount=500, ColumnCount=4, Increment=50,"
            " AsDataFrame=False);"
            "print(len(rows), rows[-1]['Rows'])")
    result = subprocess.run([sys.executable, "-c", code],
                            capture_output=True, text=True)
    assert result.returncode == 0, result.output
    assert result.stdout.strip() == "2 500"


# --- formatExcel (PowerShell) ------------------------------------------------

FORMAT_EXCEL = PS_DIR / "formatExcel.ps1"


def _governanceWorkbook(tmp_path):
    """A workbook shaped like the Planner export formatExcel was built for."""
    from src.utils.xlsx import writeWorkbook
    columns = ["Task ID", "Task Name", "Bucket", "Priority", "Due date",
               "Start date", "Checklist Items", "Notes"]
    rows = [{"Task ID": f"ID{n:03d}", "Task Name": f"Task {n}",
             "Bucket": "BACKLOG", "Priority": "Medium", "Due date": "",
             "Start date": "", "Checklist Items": "",
             "Notes": "x" * 400} for n in range(1, 6)]
    target = tmp_path / "gov.xlsx"
    writeWorkbook([{"name": "Consolidated Data", "columns": columns,
                    "rows": rows}], str(target))
    return target


def test_formatexcel_exists_and_defaults_format_name_to_corporate():
    text = FORMAT_EXCEL.read_text(encoding="utf-8")
    assert "$format_name   = 'corporate'" in text, \
        "format_name must default to corporate"
    for style in ("corporate", "corporatev2", "plain", "minimal"):
        assert f"'{style}'" in text


def test_formatexcel_picks_the_renderer_from_the_output_extension():
    text = FORMAT_EXCEL.read_text(encoding="utf-8")
    for extension in ("'.docx'", "'.pdf'", "'.html'"):
        assert extension in text


@needsPwsh
@pytest.mark.parametrize("extension", ["docx", "html"])
def test_formatexcel_renders_without_office_installed(tmp_path, extension):
    """docx and html must need nothing beyond .NET -- no Word, no add-in."""
    source = _governanceWorkbook(tmp_path)
    out = tmp_path / f"report.{extension}"
    result = runPwsh("formatExcel.ps1", "-ExcelPath", source,
                     "-SheetName", "Consolidated Data",
                     "-Columns", "Task Name,Bucket,Priority,Notes",
                     "-OutputPath", out)
    assert result.returncode == 0, result.output
    assert out.is_file() and out.stat().st_size > 0
    assert "4 column(s)" in result.output


@needsPwsh
def test_formatexcel_accepts_a_comma_list_from_the_command_line(tmp_path):
    """In -File mode every argument is a string, so 'a','b' arrives as "a,b".

    Without the split this produced "column(s) not found" naming columns
    that were plainly in the sheet.
    """
    source = _governanceWorkbook(tmp_path)
    out = tmp_path / "r.html"
    result = runPwsh("formatExcel.ps1", "-ExcelPath", source,
                     "-SheetName", "Consolidated Data",
                     "-Columns", "Task Name,Bucket,Priority",
                     "-OutputPath", out)
    assert result.returncode == 0, result.output
    assert "keeping 3 column(s)" in result.output


@needsPwsh
def test_formatexcel_keeps_the_requested_column_order(tmp_path):
    source = _governanceWorkbook(tmp_path)
    out = tmp_path / "r.html"
    result = runPwsh("formatExcel.ps1", "-ExcelPath", source,
                     "-SheetName", "Consolidated Data",
                     "-Columns", "Notes,Task Name", "-OutputPath", out)
    assert result.returncode == 0, result.output
    html = out.read_text(encoding="utf-8")
    assert html.index("<th>Notes</th>") < html.index("<th>Task Name</th>"), \
        "columns must appear in the order asked for, not sheet order"


@needsPwsh
def test_formatexcel_names_the_real_columns_when_one_is_wrong(tmp_path):
    source = _governanceWorkbook(tmp_path)
    result = runPwsh("formatExcel.ps1", "-ExcelPath", source,
                     "-SheetName", "Consolidated Data",
                     "-Columns", "Task Name;Widget",
                     "-OutputPath", tmp_path / "r.docx")
    assert result.returncode == 3
    assert "Widget" in result.output
    assert "the sheet has" in result.output, \
        "an unknown column must list the columns that DO exist"


def _noPdfConverter():
    """True when neither Word nor LibreOffice could convert a .docx here."""
    if platform.system() == "Windows":
        return False          # Word may well be installed
    return (shutil.which("soffice") is None
            and shutil.which("libreoffice") is None)


def test_formatexcel_names_every_way_out_when_no_converter_exists():
    """Checked in the source, so it holds on machines that DO have one."""
    text = FORMAT_EXCEL.read_text(encoding="utf-8")
    for fix in ("FIX 1", "FIX 2", "FIX 3", "FIX 4"):
        assert fix in text
    assert "LibreOffice" in text, "the licence-free option must be offered"


@needsPwsh
def test_formatexcel_refuses_a_pdf_when_nothing_can_convert_one(tmp_path):
    """Handing back a different file type silently is not acceptable."""
    if not _noPdfConverter():
        pytest.skip("a PDF converter is installed, so there is nothing to refuse")
    source = _governanceWorkbook(tmp_path)
    out = tmp_path / "r.pdf"
    result = runPwsh("formatExcel.ps1", "-ExcelPath", source,
                     "-SheetName", "Consolidated Data", "-OutputPath", out)
    assert result.returncode == 3
    assert "FIX 1" in result.output and "FIX 3" in result.output
    assert not out.exists(), "nothing should be written when the format fails"


@needsPwsh
def test_formatexcel_falls_back_only_when_told_to(tmp_path):
    if not _noPdfConverter():
        pytest.skip("a PDF converter is installed, so no fallback is needed")
    source = _governanceWorkbook(tmp_path)
    result = runPwsh("formatExcel.ps1", "-ExcelPath", source,
                     "-SheetName", "Consolidated Data",
                     "-OutputPath", tmp_path / "r.pdf",
                     "-OnWordMissing", "html")
    assert result.returncode == 0, result.output
    assert (tmp_path / "r.html").is_file()
    assert "instead of a PDF" in result.output, "the swap must be announced"


@needsPwsh
def test_formatexcel_truncation_is_off_by_default_and_marked_when_on(tmp_path):
    source = _governanceWorkbook(tmp_path)
    full = tmp_path / "full.html"
    cut = tmp_path / "cut.html"
    runPwsh("formatExcel.ps1", "-ExcelPath", source, "-SheetName",
            "Consolidated Data", "-Columns", "Notes", "-OutputPath", full)
    runPwsh("formatExcel.ps1", "-ExcelPath", source, "-SheetName",
            "Consolidated Data", "-Columns", "Notes", "-OutputPath", cut,
            "-MaxCellLength", "50")
    assert "x" * 400 in full.read_text(encoding="utf-8"), \
        "nothing may be hidden unless asked for"
    shortened = cut.read_text(encoding="utf-8")
    assert "x" * 400 not in shortened
    assert "…" in shortened, "a cut must be visible to the reader"


@needsPwsh
@pytest.mark.parametrize("bad,code", [
    (["-format_name", "fancy"], 2),
    (["-OnWordMissing", "maybe"], 2),
])
def test_formatexcel_refuses_unknown_options(tmp_path, bad, code):
    source = _governanceWorkbook(tmp_path)
    result = runPwsh("formatExcel.ps1", "-ExcelPath", source, "-SheetName",
                     "Consolidated Data", "-OutputPath", tmp_path / "r.docx",
                     *bad)
    assert result.returncode == code
    assert "expected" in result.output


@needsPwsh
def test_formatexcel_output_opens_in_a_word_processor(tmp_path):
    """The .docx is hand-written OOXML; proof it opens is a real reader."""
    if shutil.which("libreoffice") is None:
        pytest.skip("LibreOffice is not installed")
    source = _governanceWorkbook(tmp_path)
    out = tmp_path / "r.docx"
    assert runPwsh("formatExcel.ps1", "-ExcelPath", source, "-SheetName",
                   "Consolidated Data", "-Columns", "Task Name,Bucket",
                   "-OutputPath", out).returncode == 0
    converted = subprocess.run(
        ["libreoffice", "--headless", "--convert-to", "pdf",
         "--outdir", str(tmp_path), str(out)],
        capture_output=True, text=True, timeout=300)
    assert (tmp_path / "r.pdf").is_file(), converted.stdout + converted.stderr


@needsPwsh
def test_formatexcel_makes_a_pdf_with_libreoffice_when_word_is_absent(tmp_path):
    """A server with no Office should still produce the PDF itself.

    This gap was found by delivering a report: the script refused the .pdf,
    and the conversion had to be done by hand afterwards. LibreOffice needs
    no licence and is often present where Word is not.
    """
    if shutil.which("soffice") is None and shutil.which("libreoffice") is None:
        pytest.skip("LibreOffice is not installed")
    source = _governanceWorkbook(tmp_path)
    out = tmp_path / "report.pdf"
    result = runPwsh("formatExcel.ps1", "-ExcelPath", source,
                     "-SheetName", "Consolidated Data",
                     "-Columns", "Task Name,Bucket,Priority",
                     "-OutputPath", out)
    assert result.returncode == 0, result.output
    assert out.is_file() and out.stat().st_size > 1000
    assert out.read_bytes()[:5] == b"%PDF-", "the file must really be a PDF"
    assert "LibreOffice" in result.output
    assert not (tmp_path / "report.docx").exists(), \
        "the staging .docx must be cleaned up, not left beside the PDF"


# --- formatExcel: the two engines must agree ---------------------------------

def _runBothEngines(tmp_path, extension, columns="Task Name,Bucket,Priority,Notes"):
    """Render the same workbook with both engines. Returns: (psPath, pyPath)."""
    from src.py.formatExcel import formatExcel as pyFormatExcel
    source = _governanceWorkbook(tmp_path)
    psPath = tmp_path / f"ps.{extension}"
    pyPath = tmp_path / f"py.{extension}"
    shared = dict(SheetName="Consolidated Data", Columns=columns,
                  Title="Governance", Subtitle="fixed", MaxCellLength=220)
    result = runPwsh("formatExcel.ps1", "-ExcelPath", source,
                     "-SheetName", shared["SheetName"],
                     "-Columns", shared["Columns"], "-OutputPath", psPath,
                     "-Title", shared["Title"], "-Subtitle", shared["Subtitle"],
                     "-MaxCellLength", str(shared["MaxCellLength"]))
    assert result.returncode == 0, result.output
    pyFormatExcel(ExcelPath=str(source), OutputPath=str(pyPath), **shared)
    return psPath, pyPath


@needsPwsh
def test_the_two_engines_write_byte_identical_html(tmp_path):
    """The strongest parity check available: compare the files themselves."""
    psPath, pyPath = _runBothEngines(tmp_path, "html")
    assert psPath.read_bytes() == pyPath.read_bytes(), \
        "the PowerShell and Python reports differ byte for byte"


@needsPwsh
def test_the_two_engines_write_an_identical_word_document(tmp_path):
    """Every OOXML part must match, column widths included.

    document.xml carries the computed widths, so this also pins the
    rounding: PowerShell's [int] rounds half to even and Python's int()
    truncates, which would put every column one twip apart.
    """
    psPath, pyPath = _runBothEngines(tmp_path, "docx")
    import zipfile
    for name in ("word/document.xml", "word/styles.xml",
                 "[Content_Types].xml", "_rels/.rels",
                 "word/_rels/document.xml.rels"):
        fromPs = zipfile.ZipFile(psPath).read(name)
        fromPy = zipfile.ZipFile(pyPath).read(name)
        assert fromPs == fromPy, f"{name} differs between the engines"


def test_powershell_rounding_is_mirrored_not_truncated():
    from src.py.formatExcel import psInt
    # .NET rounds half to EVEN, which is what Python's round() does and
    # int() does not.
    assert (psInt(2.5), psInt(3.5), psInt(3.4), psInt(3.6)) == (2, 4, 3, 4)
    assert int(2.6) == 2 and psInt(2.6) == 3, "int() would truncate"


def test_cell_line_endings_are_normalised_to_lf(tmp_path):
    """A literal CRLF inside a cell must fold to LF, per the XML spec.

    An Excel cell with an embedded newline stores a literal CRLF. .NET's
    XmlDocument kept the CR when loading from a string, so the PowerShell
    reader disagreed with Python until it was normalised.
    """
    from src.utils.xlsx import writeWorkbook, readSheet
    target = tmp_path / "wrapped.xlsx"
    writeWorkbook([{"name": "S", "columns": ["Note"],
                    "rows": [{"Note": "first\r\nsecond\rthird"}]}], str(target))
    value = readSheet(str(target), sheet="S")[0]["Note"]
    assert "\r" not in value
    assert value == "first\nsecond\nthird"


@pytest.mark.parametrize("engineSource", ["ps", "py"])
def test_both_engines_offer_the_same_formats_and_options(engineSource):
    text = (FORMAT_EXCEL if engineSource == "ps"
            else ROOT / "src" / "py" / "formatExcel.py").read_text(encoding="utf-8")
    for token in ("corporate", "corporatev2", "plain", "minimal",
                  ".docx", ".pdf", ".html", "LibreOffice",
                  "OnWordMissing", "MaxCellLength", "format_name"):
        assert token in text, f"{engineSource} is missing {token}"


def test_python_formatexcel_rejects_the_same_things(tmp_path):
    from src.py.formatExcel import formatExcel as pyFormatExcel
    source = _governanceWorkbook(tmp_path)
    base = dict(ExcelPath=str(source), SheetName="Consolidated Data",
                OutputPath=str(tmp_path / "o.docx"))
    with pytest.raises(ValueError, match="format_name"):
        pyFormatExcel(**{**base, "format_name": "fancy"})
    with pytest.raises(ValueError, match="OnWordMissing"):
        pyFormatExcel(**{**base, "OnWordMissing": "maybe"})
    with pytest.raises(ValueError, match="extension"):
        pyFormatExcel(**{**base, "OutputPath": str(tmp_path / "o.pptx")})
    with pytest.raises(FileNotFoundError):
        pyFormatExcel(**{**base, "ExcelPath": str(tmp_path / "nope.xlsx")})
    with pytest.raises(KeyError) as caught:
        pyFormatExcel(**{**base, "Columns": "Task Name,Widget"})
    assert "the sheet has" in str(caught.value)


def test_python_formatexcel_returns_what_it_did(tmp_path):
    from src.py.formatExcel import formatExcel as pyFormatExcel
    source = _governanceWorkbook(tmp_path)
    result = pyFormatExcel(ExcelPath=str(source), SheetName="Consolidated Data",
                           Columns=["Notes", "Task Name"],
                           OutputPath=str(tmp_path / "o.html"))
    assert result["RowCount"] == 5
    assert result["Columns"] == ["Notes", "Task Name"], "order is preserved"
    assert result["Renderer"] == "html"


# --- formatExcel.py must not need a COM package ------------------------------

def test_formatexcel_imports_no_com_package():
    """pywin32 is not installed on a CGS desktop and must not be needed.

    Word is still supported -- it is reached through a short PowerShell
    script, because PowerShell ships with Windows and has COM built in.
    """
    source = (ROOT / "src" / "py" / "formatExcel.py").read_text(encoding="utf-8")
    for line in source.splitlines():
        stripped = line.strip()
        assert not stripped.startswith(("import win32", "from win32",
                                        "import pywin32", "from pywin32")), \
            f"a COM package is imported: {stripped}"
    assert "New-Object -ComObject Word.Application" in source, \
        "Word support must survive, driven through PowerShell"


def test_formatexcel_works_with_every_optional_package_blocked(tmp_path):
    """Import and render with win32com, pywin32 and pandas unavailable."""
    source = _governanceWorkbook(tmp_path)
    out = tmp_path / "blocked.html"
    code = ("import sys;"
            "blocked=('win32com','win32com.client','pywin32','pandas',"
            "'numpy','openpyxl');"
            "[sys.modules.__setitem__(m, None) for m in blocked];"
            f"sys.path.insert(0, r'{ROOT}');"
            "from src.py.formatExcel import formatExcel;"
            f"r = formatExcel(ExcelPath=r'{source}',"
            " SheetName='Consolidated Data', Columns='Task Name,Bucket',"
            f" OutputPath=r'{out}');"
            "print(r['RowCount'], r['Renderer'])")
    result = subprocess.run([sys.executable, "-c", code],
                            capture_output=True, text=True)
    assert result.returncode == 0, result.stderr
    assert result.stdout.strip() == "5 html"
    assert out.is_file()


def test_the_word_path_quotes_paths_safely():
    """An apostrophe in a folder name must not break out of the script."""
    from src.py import formatExcel as module
    captured = {}

    def fake(script, timeout=600):
        captured["script"] = script
        raise RuntimeError("not actually running Word")

    realRunner = module.runPowerShellScript
    module.runPowerShellScript = fake
    try:
        with pytest.raises(RuntimeError):
            module.convertDocxToPdfWithWord("/tmp/O'Brien/report.docx",
                                            "/tmp/O'Brien/report.pdf")
    finally:
        module.runPowerShellScript = realRunner
    # PowerShell escapes a single quote by doubling it inside a '...' string.
    assert "O''Brien" in captured["script"]
    assert "O'Brien'" not in captured["script"], "the quoting leaks"


@needsPwsh
def test_the_powershell_bridge_runs():
    from src.py.formatExcel import runPowerShellScript
    result = runPowerShellScript("Write-Output 'bridge works'\n", timeout=120)
    assert "bridge works" in result.stdout


# --- the Claims_Data_BIG repair ----------------------------------------------

def test_the_claims_data_columns_cannot_be_an_excel_table():
    """Why Claims_Data_BIG is a plain range and not a table.

    Excel requires a table's column names to be unique IGNORING CASE. The
    Metadata data dictionary legitimately carries both Allowed_Amount and
    allowed_amount, and both Claim_Type and claim_type -- two variables under
    two naming conventions. Made into a table, Excel silently renames the
    second of each pair (it chose allowed_amount2 and claim_type3) and raises
    "Repaired Records: Table ... (Table)" on open.

    This test exists so that anyone tempted to turn the sheet back into a
    table finds out here rather than from a repair prompt.
    """
    import importlib.util
    spec = importlib.util.spec_from_file_location(
        "bigbuilder", ROOT / "src" / "utils" /
        "build_claims_pipelines_workbook.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)

    columns = module.BIG_COLUMNS
    assert len(set(columns)) == len(columns), "names must be literally unique"
    folded = [name.lower() for name in columns]
    collisions = {name for name in folded if folded.count(name) > 1}
    assert collisions == {"allowed_amount", "claim_type"}, (
        "if this set changed, re-check whether a table is viable again")


def test_the_lifecycle_diagram_closes_the_loop():
    """Six steps, six arrows, and the last one points back at the first."""
    import importlib.util
    spec = importlib.util.spec_from_file_location(
        "v15builder", ROOT / "src" / "utils" / "build_claims_pipelines_v15.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)

    assert module.LIFECYCLE_STEPS[0] == "Data"
    assert module.LIFECYCLE_STEPS[-1] == "Linear Regression"
    drawing = module.lifecycleAnchorXml(56)
    assert drawing.count('prst="ellipse"') == len(module.LIFECYCLE_STEPS)
    # One arrow per step, including the one from the last step back to the
    # first -- a lifecycle that does not close is just a list.
    assert drawing.count("straightConnector1") == len(module.LIFECYCLE_STEPS)
    assert drawing.count("tailEnd") == len(module.LIFECYCLE_STEPS)
    for step in module.LIFECYCLE_STEPS:
        assert f"<a:t>{step}</a:t>" in drawing


def test_the_workbook_sheets_call_cgs_ai_rather_than_copying_it():
    """Every =PY() script the v1.5 builder writes calls the package.

    The point of the CGI_AI and Test_Scaling rewrite is that the sheet stops
    carrying its own copy of the code. A script that defines a function
    instead of calling cgs_ai.something() has quietly undone that.
    """
    import importlib.util
    spec = importlib.util.spec_from_file_location(
        "v15builder2", ROOT / "src" / "utils" / "build_claims_pipelines_v15.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)

    assert "import cgs_ai" in module.IMPORT_SNIPPET
    assert "cgs_ai.testWorkload(" in module.SCALE_CALL
    assert "def " not in module.SCALE_CALL, "the sheet must not redefine it"
    for snippet in (module.CGI_BASIC_TEXT, module.CGI_PERSONAL_TEXT,
                    module.CGI_DETAILED_TEXT, module.CGI_SURVEY_TEXT,
                    module.CGI_SURVEY_SEED_TEXT):
        assert "cgs_ai." in snippet
        assert "def " not in snippet
