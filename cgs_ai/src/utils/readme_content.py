"""
=====================================================================
  Program Name  : readme_content.py
  Author        : Manuel Figallo
  Purpose       : Assemble the README.docx content for cgs_ai.
  Version       : 1.0beta
  Created       : 2026-08-26
  Last Modified : 2026-08-26

  Dependencies: standard library only (uses build_readme_docx.py).

  Description:
    Section order and the icon vocabulary follow the supplied guidance
    documents. Separating content (here) from the OOXML writer
    (build_readme_docx.py) keeps both readable.

  Usage:
    python src/utils/readme_content.py [output_path]
=====================================================================
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_readme_docx import (bullet, pageBreak, para, table,  # noqa: E402
                               writeDocx, NAVY)

__version__ = "1.0beta"
DOC_VERSION = "1.0beta"
DOC_DATE = "2026-08-26"


def buildBody() -> str:
    """Assemble the whole document body. Returns: the <w:body> XML."""
    x = []

    # ---------------- Title page ----------------
    x.append(para("cgs_ai", style="Title"))
    x.append(para("Claims Intelligence Toolkit for SAS, Python and PowerShell",
                  style="Subtitle", spaceAfter=200))
    x.append(para("Project Reference, Function Guide and Setup Tutorial",
                  size=13, color=NAVY, spaceAfter=400))
    x.append(para(f"Version {DOC_VERSION}   |   {DOC_DATE}   |   Author: Manuel Figallo",
                  size=11, spaceAfter=400))

    x.append(para("Version Control", style="Heading2"))
    x.append(table(
        ["Version", "Date", "Author", "Summary of Changes"],
        [["1.0beta", "2026-08-26", "Manuel Figallo",
          "First consolidated release. Data-engineering folder structure "
          "adopted; .env/.gitignore configuration added; scanFileSystem "
          "regrained to one row per keyword match; eight new functions "
          "(runSQLServerQuery, formatCSV, downloadBulkFiles, sendEmail, "
          "convertSAS2Pandas, copyExcelSheet2CSV, collectSystemMetrics, "
          "zipFolder); filescan pipeline; Python/PowerShell/SAS parity."],
         ["0.9", "2026-08-25", "Manuel Figallo",
          "scanFileSystem PowerShell port and SAS SYSTASK wrappers; "
          "pandas/numpy dependency removed from the Python scanner."],
         ["0.1", "2026-08-20", "Manuel Figallo",
          "Initial scanFileSystem prototype (Python, file-grain output)."]],
        [1100, 1300, 1900, 5900]))

    x.append(pageBreak())

    # ---------------- Outline ----------------
    x.append(para("Document Outline", style="Heading1"))
    for index, title in enumerate([
        "Overview",
        "Project Folder Structure",
        "Folder and File Descriptions",
        "Why Folder Structure Is Important for Data Engineering",
        "Recommended Engineering Principle",
        "Quick Architecture Flow",
        "Configuration: .env and .gitignore",
        "Function Reference",
        "Running cgs_ai in Visual Studio Code: A Brief Tutorial",
        "Cross-Language Parity and Conventions",
    ], start=1):
        x.append(bullet(f"{index}.  {title}"))
    x.append(pageBreak())

    # ---------------- 1. Overview ----------------
    x.append(para("1. Overview", style="Heading1"))
    x.append(para(
        "cgs_ai is a claims-intelligence toolkit that runs the same operations "
        "from SAS, Python or PowerShell. It scans file systems and job logs for "
        "keywords, queries the LOB data marts, formats results for distribution, "
        "downloads regulatory attachments in bulk, converts SAS data sets, "
        "collects host metrics and sends alerts."))
    x.append(para(
        "The design goal is a single toolkit that an analyst can call from a SAS "
        "session, a scheduler can call from a batch file, and an engineer can "
        "import into Python, with identical function names, identical parameter "
        "names and identical output. A pipeline written once behaves the same "
        "whichever language invokes it."))
    x.append(para("Dependency policy", style="Heading2"))
    x.append(para(
        "The package imports with nothing beyond the Python standard library. "
        "Optional third-party packages are imported lazily by the single "
        "function that needs them, and the error names the exact install "
        "command. This matters on locked-down servers where installs require "
        "change control."))
    x.append(table(
        ["Package", "Needed by", "If missing"],
        [["openpyxl", "formatCSV, copyExcelSheet2CSV, scanFileSystem "
                      "(metric_profile)", "Scan falls back to CSV with a warning"],
         ["pandas", "convertSAS2Pandas", "Clear ImportError naming the pip command"],
         ["pyodbc", "runSQLServerQuery (Python only)",
          "Use the PowerShell twin, which needs no module"],
         ["psutil", "collectSystemMetrics (richer figures)",
          "Standard-library subset is collected instead"]],
        [1900, 4100, 4200]))

    # ---------------- 2. Folder structure ----------------
    x.append(para("2. Project Folder Structure", style="Heading1"))
    x.append(para(
        "A well-organized data engineering project separates configuration, "
        "source code, data pipelines, tests, data assets, logs and execution "
        "entry points. This structure makes the project easier to understand, "
        "maintain, test, scale, deploy and govern."))
    x.append(para(
        "cgs_ai/\n"
        "├── README.docx\n"
        "├── requirements.txt\n"
        "├── .env\n"
        "├── .gitignore\n"
        "├── docker-compose.yml\n"
        "├── __init__.py\n"
        "├── cgs_ai_setup.py\n"
        "├── Examples.py\n"
        "│\n"
        "├── src/\n"
        "│   ├── ps/        PowerShell implementations\n"
        "│   ├── py/        Python implementations\n"
        "│   ├── sas/       SAS wrapper macros\n"
        "│   ├── ingestion/\n"
        "│   ├── transforms/\n"
        "│   ├── models/\n"
        "│   ├── pipelines/ dags.py, pipeline.py, filescan_pipeline.py\n"
        "│   ├── utils/     helpers.py, config.py, logger.py\n"
        "│   └── api/       routes.py, schemas.py\n"
        "│\n"
        "├── tests/\n"
        "├── data/          sample/, reference/\n"
        "├── logs/          .gitkeep\n"
        "└── main.py",
        style="Code"))
    x.append(table(
        ["Icon", "Folder / File", "Description"],
        [["📄", "README.docx", "Project overview, setup, usage and examples."],
         ["📄", "requirements.txt", "List of Python dependencies (all optional)."],
         ["🔐", ".env", "Environment variables and API keys; never commit to GitHub."],
         ["🔶", ".gitignore", "Tells Git which files and folders to ignore."],
         ["🐳", "docker-compose.yml", "Optional: local SMTP sink and SQL Server for development."],
         ["🗂️", "src/", "Source code for the project."],
         ["🐍", "src/py/", "Python implementations of every function."],
         ["▶", "src/ps/", "PowerShell implementations, feature-matched to Python."],
         ["📋", "src/sas/", "SAS wrapper macros only; they launch ps or py."],
         ["🗄️", "ingestion/", "Source extraction and ingestion logic."],
         ["⚙️", "transforms/", "Data transformation and business logic."],
         ["▦", "models/", "Data models, schemas and curated-layer definitions."],
         ["🔀", "pipelines/", "DAGs, workflows and orchestration logic."],
         ["🛠️", "utils/", "Helpers, configuration, logging and common utilities."],
         ["API", "api/", "API layer used to expose data or trigger pipelines."],
         ["🧪", "tests/", "Unit tests, data-quality tests and integration tests."],
         ["🗄️", "data/", "Sample data, reference data and test datasets."],
         ["📄", "logs/", "Log files for debugging and monitoring."],
         ["▶", "main.py", "Entry point to run the project as a CLI or pipeline runner."]],
        [800, 2600, 6800]))

    # ---------------- 3. Descriptions ----------------
    x.append(para("3. Folder and File Descriptions", style="Heading1"))
    x.append(table(
        ["Icon", "Component", "Why It Exists"],
        [["🗂️", "src/", "Holds the production source code. Keeping application "
                        "code under src/ separates it from tests, data, "
                        "documentation and deployment files."],
         ["🗄️", "ingestion/", "Code that extracts or receives data from source "
                              "systems: APIs, databases, files, queues or cloud storage."],
         ["⚙️", "transforms/", "Cleansing, standardization, enrichment, aggregation "
                               "and business-rule logic."],
         ["▦", "models/", "Schemas, dimensions, facts and curated-layer structures."],
         ["🔀", "pipelines/", "Coordinates workflow execution, dependencies, "
                              "scheduling and orchestration."],
         ["🛠️", "utils/", "Shared helper functions; avoids duplicating configuration, "
                          "logging and utility code."],
         ["API", "api/", "Exposes project capabilities through endpoints, or triggers "
                         "data processes."],
         ["📋", "tests/", "Automated tests for code, transformations, data quality and "
                          "integration behavior."],
         ["🗄️", "data/", "Sample, reference or test data supporting development and "
                         "validation."],
         ["📄", "logs/", "Runtime logs used in debugging, operational monitoring and "
                         "troubleshooting."],
         ["▶", "main.py", "A clear, predictable entry point for running the "
                          "application, service, CLI or pipeline."]],
        [800, 2000, 7400]))

    # ---------------- 4. Why it matters ----------------
    x.append(para("4. Why Folder Structure Is Important for Data Engineering",
                  style="Heading1"))
    for index, text in enumerate([
        "Separation of concerns: ingestion, transformation, modeling, "
        "orchestration, APIs, testing and utilities have different "
        "responsibilities. Separating them reduces coupling and eases maintenance.",
        "Faster onboarding: a predictable layout lets a new engineer or analyst "
        "locate source code, pipeline definitions, tests, configuration and "
        "documentation quickly.",
        "Better maintainability: changes are isolated to the appropriate module "
        "instead of accumulating in a few large scripts.",
        "Improved testability: a dedicated tests/ area encourages automated unit, "
        "integration and data-quality testing, and makes failures easier to diagnose.",
        "Scalability: as sources, transformations, models and pipelines grow, the "
        "structure expands without becoming an unmanageable collection of scripts.",
        "Clear orchestration boundaries: keeping workflow code separate from "
        "business transformations makes it easier to change scheduling or "
        "orchestration tools without rewriting processing logic.",
        "Security and configuration management: .env and .gitignore separate "
        "secrets and environment-specific settings from source code.",
        "Reusability: shared utilities, business logic, models and schemas can be "
        "reused across multiple pipelines and applications.",
        "Better governance and collaboration: a standard layout supports code "
        "review, version control, ownership, documentation, CI/CD and consistent "
        "engineering practices.",
        "Production readiness: a clear entry point, tests, logging, configuration "
        "and deployment artifacts ease the move from prototype to reliable "
        "production system.",
    ], start=1):
        x.append(para(f"{index}.  {text}", spaceAfter=80))

    # ---------------- 5. Principle ----------------
    x.append(para("5. Recommended Engineering Principle", style="Heading1"))
    x.append(para("Organize by responsibility and lifecycle, not simply by file type.",
                  bold=True, size=13, color=NAVY))
    x.append(para(
        "Source extraction belongs in ingestion/, business transformations belong "
        "in transforms/, data structures belong in models/, workflow coordination "
        "belongs in pipelines/, and validation belongs in tests/. This creates a "
        "project that is easier to reason about, modify, test and scale."))
    x.append(para(
        "cgs_ai applies the same rule across languages: src/py and src/ps hold "
        "implementations, src/sas holds wrappers only, and src/pipelines holds "
        "orchestration that calls functions but never reimplements them."))

    # ---------------- 6. Flow ----------------
    x.append(para("6. Quick Architecture Flow", style="Heading1"))
    x.append(para("🗄️ Sources  →  📥 ingestion/  →  ⚙️ transforms/  →  ▦ models/  "
                  "→  🔀 pipelines/  →  API / Reports / Data Consumers",
                  size=12, bold=True, spaceAfter=200))
    x.append(para("Worked example: the filescan pipeline", style="Heading2"))
    x.append(para(
        "🗄️ SAS job logs on a UNC share  →  🔀 filescan_pipeline  →  "
        "scanFileSystem (keyword matches + metrics)  →  formatCSV (styled "
        "workbook)  →  sendEmail (operations mailbox)", spaceAfter=160))
    x.append(pageBreak())

    # ---------------- 7. Configuration ----------------
    x.append(para("7. Configuration: .env and .gitignore", style="Heading1"))
    x.append(para(
        "🔐 .env holds environment variables and paths in plain KEY=VALUE form. "
        "It is never committed. .env.example is committed and documents every "
        "variable with placeholder values, so a new user knows what to provide."))
    x.append(para("Rules for .env", style="Heading2"))
    for rule in [
        "One variable per line.",
        "Use KEY=VALUE.",
        "Do not put spaces around the '='.",
        "Lines beginning with # are comments.",
        "The file is named exactly .env, with no .txt extension.",
        "Never commit .env if it contains passwords, tokens or API keys.",
    ]:
        x.append(bullet(rule))
    x.append(para("Variables", style="Heading2"))
    x.append(table(
        ["Variable", "Purpose"],
        [["PS_FOLDER_PATH", "Folder holding the PowerShell implementations (src\\ps\\)."],
         ["SAS_FOLDER_PATH", "Folder holding the SAS wrapper macros (src\\sas\\)."],
         ["PYTHON_FOLDER_PATH", "Folder holding the Python implementations (src\\py)."],
         ["ROOT_DATA", "Root for data inputs and generated output."],
         ["ROOT_SRC", "Root of the source tree."],
         ["SMTP_SERVER / SMTP_PORT", "Optional defaults for sendEmail."],
         ["SQL_DATA_SOURCE / SQL_LOB_CATALOG", "Optional defaults for runSQLServerQuery."]],
        [3400, 6800]))
    x.append(para("🔶 .gitignore excludes:", style="Heading2"))
    x.append(para(".env\n.env.*\n!.env.example", style="Code"))
    x.append(para(
        "The '!' line is important: it re-includes .env.example so the template "
        "stays in source control while every real .env is excluded."))
    x.append(para(
        "Configuration precedence: a real operating-system environment variable "
        "always overrides the .env value, so a scheduled task or CI job can "
        "override settings without editing files."))

    x.append(pageBreak())

    # ---------------- 8. Function reference ----------------
    x.append(para("8. Function Reference", style="Heading1"))
    x.append(para(
        "Every function exists in Python (src/py), PowerShell (src/ps) and as a "
        "SAS wrapper macro (src/sas). Function names and parameter names are "
        "identical in all three. Each entry below gives the parameters and the "
        "role the function plays in claims processing."))

    functions = [
        ("scanFileSystem",
         "Scan directory roots for keyword matches in text files. Emits ONE ROW "
         "PER MATCH with the matched line, configurable context lines above and "
         "below, and extracted tokens. An optional metric profile parses "
         "structured performance metrics and produces Excel output.",
         [("input_folder_root", "REQUIRED. Root path(s); list or ';'-delimited."),
          ("extract_keyword", "REQUIRED. Keyword(s) to find; no keywords means no rows."),
          ("output_file_path", ".csv or .xlsx, or a directory to auto-name inside. "
                               "Omitted: scan_YYYYMMDD_HHMMSS.csv in the current folder."),
          ("file_extensions", "Extensions to include. Default log, txt, sas."),
          ("include_subdirectories", "Recurse. Default true."),
          ("folder_exclusion_list", "Folder names to skip. Default empty: nothing excluded."),
          ("file_exclusion_list", "Prefixes stripped to derive the program name."),
          ("lines_above / lines_below", "Context lines captured. Default 5 each."),
          ("nth_token_after / nth_token_before", "Which token beside the keyword. Default 1."),
          ("numeric_token_after", "Which numeric token after the keyword. Default 1."),
          ("date_from / date_to / date_field", "Inclusive YYYY-MM-DD bounds on "
                                               "created, modified or accessed."),
          ("metric_profile", "none (default) or sas_log. When set, EXCEL output is "
                             "produced -- announced in the log -- with a second "
                             "Metrics sheet of structured numbers.")],
         "Sweep SAS and ETL job logs for ERROR, a claim number, a denial code or a "
         "file reference. The analyst gets the exact line, the surrounding context "
         "and the adjacent tokens (for example the record count following a "
         "keyword) as a filterable table, instead of opening hundreds of logs."),

        ("runSQLServerQuery",
         "Run a SQL Server query against a LOB catalog using Windows Integrated "
         "Security. Equivalent to the PROC SQL 'connect to oledb' pass-through "
         "block. No password is handled, stored or logged.",
         [("SQL_Statement", "REQUIRED. The query text (the SAS &query equivalent)."),
          ("LOB_Catalog", "REQUIRED. Initial Catalog, e.g. DataMartKYA."),
          ("DataSource", "server,port. Defaults to SQL_DATA_SOURCE from .env."),
          ("OutputCsvPath", "Optional CSV for the result set.")],
         "Pull claims, denial or provider extracts straight from the LOB data mart "
         "into CSV or a DataFrame, using the same catalog and credentials the SAS "
         "pass-through uses. The PowerShell twin needs no module at all, which "
         "matters on servers where installs require change control."),

        ("formatCSV",
         "Render a CSV as a styled Excel workbook with a SAS ODS look and feel: "
         "navy title banner, blue header row, zebra striping, frozen and "
         "auto-filtered header.",
         [("InputCsvPath", "REQUIRED. Source CSV."),
          ("OutputExcelPath", "REQUIRED. Destination .xlsx."),
          ("FormatType", "corporate (default), corporatev2, plain or\n                          minimal. corporate and corporatev2 render the\n                          same navy/blue/zebra look."),
          ("SheetName", "Worksheet name. Default Report."),
          ("Title", "Banner text. Defaults to the CSV filename.")],
         "Turn a raw scan result or claims extract into a report a manager or "
         "auditor can open directly, with no hand-formatting each cycle. "
         "Consistent presentation also makes month-over-month comparison easier."),

        ("downloadBulkFiles",
         "Download every file referenced by a CSV column of HTTP links. A cell may "
         "be blank, a single URL, or several joined by '|'. Blank cells are skipped "
         "rather than treated as errors; a failed download is logged and the run "
         "continues.",
         [("InputCsvPath", "REQUIRED. CSV containing the link column."),
          ("OutputFolder", "REQUIRED. Destination folder; created if needed."),
          ("LinkColumn", "Column holding the URLs. Default attachmentLinks."),
          ("IdColumn", "Column used to prefix saved names so identically named "
                       "attachments cannot collide. Default commentId."),
          ("Overwrite", "Re-download existing files. Default false.")],
         "Pull every public-comment attachment for a CMS docket (for example "
         "CMS-2022-0193, the Interoperability and Prior Authorization rule) into "
         "one folder, so the documents can be text-extracted, classified and "
         "reviewed in bulk rather than one browser click at a time."),

        ("sendEmail",
         "Send an email alert over SMTP to one or many recipients.",
         [("To", "REQUIRED. Recipient(s); list or ';'-delimited string."),
          ("From", "REQUIRED. Sender address."),
          ("Subject", "REQUIRED."), ("Body", "REQUIRED."),
          ("SmtpServer", "Default smtp.example.com, or SMTP_SERVER from .env."),
          ("Port", "Default 25, or SMTP_PORT from .env.")],
         "Notify the claims-operations mailbox when an overnight scan, extract or "
         "bulk download finishes, including row counts and output locations, so "
         "nobody has to watch the job or discover a silent failure the next day."),

        ("convertSAS2Pandas",
         "Read a sas7bdat data set into a pandas DataFrame and persist it as "
         "parquet, CSV or pickle. The one function that cannot be standard-library "
         "only, because sas7bdat is a proprietary binary format.",
         [("InputSas7bdatPath", "REQUIRED. Source .sas7bdat."),
          ("OutputPath", "REQUIRED. Destination; the extension selects the writer.")],
         "Bring a SAS claims extract into the Python and Snowflake side of the "
         "stack without a manual export step, so the same data can feed pandas "
         "analysis, a Snowflake load or an AI model without re-keying."),

        ("copyExcelSheet2CSV",
         "Export one Excel worksheet to CSV, validating the sheet BEFORE writing "
         "and stopping on the first problem: sheet missing, sheet empty, blank or "
         "duplicate header names, or merged cells across the header row.",
         [("InputExcelPath", "REQUIRED. Source workbook."),
          ("SheetName", "REQUIRED. Worksheet to export."),
          ("OutputCsvPath", "REQUIRED. Destination CSV."),
          ("HeaderRow", "1-based header row. Default 1; use 2 for a workbook with "
                        "a title banner, which is what formatCSV produces.")],
         "Convert a hand-maintained Excel reference workbook -- fee schedules, "
         "denial-code mappings, provider lists -- into the CSV a pipeline can "
         "consume, catching formatting mistakes at the source instead of letting "
         "a malformed file corrupt a downstream claims run."),

        ("collectSystemMetrics",
         "Gather host metrics into a CSV time series. Fails gracefully by design: "
         "each probe is independent, so a metric unavailable on this server is "
         "recorded blank and named in an Errors column while the run still succeeds.",
         [("OutputCsvPath", "REQUIRED. CSV to write."),
          ("ServerName", "Host label. Defaults to this machine's name."),
          ("WriteMode", "append (default) preserves history; overwrite replaces.")],
         "Sample the SAS or ETL server during a nightly claims run, then correlate "
         "slow steps found by scanFileSystem's metric profile with CPU, memory or "
         "disk pressure on the host -- the difference between 'the job was slow' "
         "and 'the job was slow because the server was starved'."),

        ("zipFolder",
         "Archive a folder together with a list of accompanying files into one zip.",
         [("FolderToZip", "REQUIRED. Folder to archive, stored under its own name."),
          ("OutputZipPath", "REQUIRED. Full path of the .zip."),
          ("AccompanyFiles", "Extra files placed at the archive root.")],
         "Bundle a month of scan output, the formatted Excel report and the run log "
         "into a single archive for records retention, audit response or hand-off "
         "to another team."),

        ("runFilescanPipeline",
         "End-to-end orchestration: scanFileSystem, then formatCSV, then sendEmail. "
         "Orchestration only -- all work lives in the three functions it calls. "
         "Email failure is reported but does not fail the pipeline, because the "
         "scan output is already on disk and is the deliverable.",
         [("input_folder_root / extract_keyword", "Passed to scanFileSystem."),
          ("output_file_path / excel_output_path", "Scan and report destinations."),
          ("metric_profile", "Default sas_log, so Excel with a Metrics sheet."),
          ("email_to / email_from / email_subject", "Notification; skipped when "
                                                    "email_to is empty.")],
         "The nightly job: sweep every SAS log on the HHH and DME shares for timing "
         "and error keywords, produce a formatted workbook, and deliver it to the "
         "operations mailbox before the morning shift starts."),
    ]

    for name, summary, params, use in functions:
        x.append(para(name, style="Heading2"))
        x.append(para(summary, spaceAfter=100))
        x.append(table(["Parameter", "Description"],
                       [[p, d] for p, d in params], [3000, 7200]))
        x.append(para("Use in claims processing", style="Heading3"))
        x.append(para(use, spaceAfter=200))

    x.append(para("Helper functions (src/utils)", style="Heading2"))
    x.append(table(
        ["Function", "Purpose"],
        [["getConfig / loadConfig", "Read a value from .env with the OS environment "
                                    "taking precedence."],
         ["logInfo / logWarn / logError", "Timestamped stderr logging, identical "
                                          "format in Python and PowerShell."],
         ["asList", "Normalize a list parameter, accepting the ';'-delimited form "
                    "the SAS wrappers pass."],
         ["writeCsv / readCsv", "Standard-library CSV I/O with stable column order."],
         ["basic_hello / personalized_hello / detailed_hello",
          "Smoke tests that confirm the package imported correctly."]],
        [3000, 7200]))

    x.append(pageBreak())

    # ---------------- 9. VS Code tutorial ----------------
    x.append(para("9. Running cgs_ai in Visual Studio Code: A Brief Tutorial",
                  style="Heading1"))
    x.append(para(
        "This tutorial takes a new user from a clean machine to a working scan in "
        "about ten minutes. It uses only the Python implementation; PowerShell and "
        "SAS need no installation beyond what Windows already provides."))

    steps = [
        ("Step 1 - Install the prerequisites",
         "Install Visual Studio Code and Python 3.9 or later. During Python setup, "
         "tick 'Add Python to PATH'. In VS Code, open the Extensions panel "
         "(Ctrl+Shift+X) and install the Microsoft Python extension.", None),
        ("Step 2 - Open the project",
         "File > Open Folder, then select the cgs_ai folder -- the one containing "
         "__init__.py and main.py. Opening the parent folder instead is the most "
         "common setup mistake, because imports are resolved from the folder you "
         "open.", None),
        ("Step 3 - Create a virtual environment",
         "Open the integrated terminal with Ctrl+` and run:",
         "python -m venv .venv\n.venv\\Scripts\\activate\n"
         "pip install -r requirements.txt"),
        ("Step 4 - Select the interpreter",
         "Press Ctrl+Shift+P, type 'Python: Select Interpreter', and choose the "
         "one inside ./.venv. The chosen interpreter appears in the status bar; "
         "if it does not say .venv, the packages you just installed will not be "
         "found.", None),
        ("Step 5 - Create your .env",
         "Copy the template and edit the paths for your environment:",
         "copy .env.example .env"),
        ("Step 6 - Confirm the package imports",
         "Open Examples.py and run it (the play button, or Ctrl+F5). Expected "
         "output:",
         ">>> cgs_ai imported from: ...\\cgs_ai\\__init__.py\n"
         ">>> Hello, World!\n>>> Hello, Manuel!\n>>> cgs_ai version: 1.0beta"),
        ("Step 7 - Run your first scan",
         "In the terminal, scan any folder you can read. Every flag has a default "
         "except the two required ones:",
         "python main.py scanFileSystem ^\n"
         "    --input-folder-root \"C:\\Windows\\Logs\" ^\n"
         "    --extract-keyword \"error\" ^\n"
         "    --output-file-path \"data\\first_scan.csv\""),
        ("Step 8 - Read the output",
         "Open data\\first_scan.csv. There is one row per keyword match, not one "
         "per file. Look at LinesAbove and LinesBelow for context, and at "
         "NthTokenAfter and NumericTokenAfter for the values beside the keyword.",
         None),
        ("Step 9 - Produce a formatted report",
         "Turn that CSV into a styled workbook from a Python terminal:",
         ">>> import cgs_ai\n"
         ">>> cgs_ai.formatCSV(InputCsvPath='data/first_scan.csv',\n"
         "...                  OutputExcelPath='data/first_scan.xlsx')"),
        ("Step 10 - Debug interactively",
         "Set a breakpoint by clicking the gutter beside any line in "
         "src/py/scanFileSystem.py, then press F5. Execution stops there and the "
         "Variables panel shows the row being built. Log lines appear in the "
         "terminal, because all logging goes to stderr and stdout is kept clean.",
         None),
    ]
    for title, text, code in steps:
        x.append(para(title, style="Heading3"))
        x.append(para(text, spaceAfter=80 if code else 140))
        if code:
            x.append(para(code, style="Code"))

    x.append(para("Troubleshooting", style="Heading2"))
    x.append(table(
        ["Symptom", "Cause and fix"],
        [["ModuleNotFoundError: cgs_ai",
          "The wrong folder is open, or the interpreter is not the .venv one. "
          "Re-check Steps 2 and 4."],
         ["Excel output requires openpyxl",
          "Run pip install openpyxl, or choose a .csv output path."],
         ["No rows in the output",
          "extract_keyword is required and matching is literal. Confirm the "
          "keyword really appears in the files, and that file_extensions covers "
          "their extension."],
         ["metric_profile produced .xlsx, not the .csv I asked for",
          "Expected. The metrics need a second sheet, which CSV cannot hold; the "
          "log states this at the start of the run."],
         ["Scan is slow on a UNC share",
          "Narrow file_extensions, set include_subdirectories to false, or use "
          "date_from and date_to to limit the sweep."]],
        [3400, 6800]))

    # ---------------- 10. Parity ----------------
    x.append(para("10. Cross-Language Parity and Conventions", style="Heading1"))
    x.append(para(
        "The three languages are kept interchangeable on purpose. A call written "
        "in one can be translated to another by changing only the syntax around "
        "the arguments."))
    x.append(table(
        ["Convention", "Rule"],
        [["Function names", "Identical across Python, PowerShell and SAS "
                            "(mixedCase, e.g. scanFileSystem)."],
         ["Parameter names", "Identical across all three. PowerShell uses -name, "
                             "Python's CLI uses --kebab-case, the SAS macro uses "
                             "name=."],
         ["List parameters", "SAS passes lists as ONE semicolon-delimited string; "
                             "Python and PowerShell split it back."],
         ["Exit codes", "0 success, 2 configuration error, 3 I/O error -- so a "
                        "scheduler or SAS wrapper can branch on them."],
         ["Logging", "Timestamped lines to stderr in the same format; stdout is "
                     "kept clean for real output."],
         ["Never prompt", "No parameter is mandatory in a way that triggers a "
                          "prompt; unattended runs cannot hang."],
         ["File headers", "Every file carries the scanFileSystem header block: "
                          "Program Name, Author, Purpose, Version, Dependencies, "
                          "Description, Input Parameters, Change Log."]],
        [2600, 7600]))
    x.append(para("Calling the same operation three ways", style="Heading2"))
    x.append(para(
        "Python:\n"
        "    cgs_ai.scanFileSystem(input_folder_root=[r'\\\\srv\\logs'],\n"
        "                          extract_keyword=['real time'],\n"
        "                          metric_profile='sas_log')\n\n"
        "PowerShell:\n"
        "    .\\scanFileSystem.ps1 -input_folder_root '\\\\srv\\logs' `\n"
        "                          -extract_keyword 'real time' `\n"
        "                          -metric_profile 'sas_log'\n\n"
        "SAS:\n"
        "    %scanFileSystem(input_folder_root=%str(\\\\srv\\logs),\n"
        "                    extract_keyword=%str(real time),\n"
        "                    metric_profile=sas_log);",
        style="Code"))
    x.append(para(
        "The SAS macro takes engine=ps or engine=py, so the same macro call can "
        "run either implementation without changing any other argument."))

    return "".join(x)


if __name__ == "__main__":
    output = sys.argv[1] if len(sys.argv) > 1 else str(
        Path(__file__).resolve().parent.parent.parent / "README.docx")
    print("wrote", writeDocx(buildBody(), output))
