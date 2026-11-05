"""
=====================================================================
  Program Name  : build_formatexcel_runbook_docx.py
  Author        : Manuel Figallo
  Purpose       : Produce "Running formatExcel.ps1 - Step by Step.docx" --
                  the Word version of the runbook published at
                  https://claude.ai/artifact/VBWfw7xW1E9ZBAFiXFFQys
  Version       : 1.0beta
  Created       : 2026-11-02
  Last Modified : 2026-11-02

  Dependencies:
    STANDARD LIBRARY ONLY, plus build_readme_docx.py beside this file for
    the OOXML helpers (para, bullet, table, writeDocx).

  Description:
    One source, two outputs. The runbook lives as a Claude doc that the
    team reads and comments on; this writes the same content as a .docx for
    anyone who wants it in Word, attached to a mail, or printed.

    Kept as a SCRIPT rather than a one-off export so the Word copy can be
    regenerated when the runbook changes. An export is a snapshot nobody
    can reproduce; this is a build.

  Input Parameters (required first):
    target (REQUIRED) - the .docx to write.

  Usage:
    python build_formatexcel_runbook_docx.py "Running formatExcel.docx"

  Exit codes:
    0 = success, 2 = usage error.
=====================================================================
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import List, Sequence

sys.path.insert(0, str(Path(__file__).resolve().parent))

from build_readme_docx import (bullet, para, table,  # noqa: E402
                               writeDocx)

__version__ = "1.0beta"

EXIT_OK, EXIT_USAGE = 0, 2

#: Table column widths in twips. The page is 8.5" less two 0.75" margins =
#: 7", and a table that oversteps that is silently clipped by Word rather
#: than wrapped, so the widths are set rather than left to be shared out.
WIDTH_TOTAL = 10080
NARROW = 1500


def code(lines: str) -> List[str]:
    """Render a code block as monospaced paragraphs.

    Parameters: lines (str) - the block, newline separated.
    Returns: list[str] of paragraph XML.

    One paragraph per line rather than one paragraph with line breaks: Word
    keeps a paragraph together when it can, and a 12-line command would be
    pushed whole onto the next page, leaving half a page blank.
    """
    rows = str(lines).split("\n")
    # Space only after the LAST line: putting it on every line would double
    # the block's height, and leaving it off all of them lets the next
    # paragraph sit flush against the code.
    return [para(line or " ", style="Code",
                 spaceAfter=140 if index == len(rows) - 1 else 0)
            for index, line in enumerate(rows)]


def heading(text: str) -> str:
    """A section heading, kept with the paragraph under it."""
    return para(text, style="Heading1", spaceAfter=100, keepNext=True)


def subheading(text: str) -> str:
    """A bold lead-in line, kept with what follows it."""
    return para(text, style="Heading2", spaceAfter=80, keepNext=True)


# =====================================================================
# Content -- the same runbook, section for section
# =====================================================================

def buildBody() -> str:
    """Assemble the document body. Parameters: none. Returns: str."""
    parts: List[str] = []

    parts.append(para("Running formatExcel.ps1 \u2014 Step by Step",
                      style="Title", spaceAfter=60))
    parts.append(para("2 November 2026  ·  Manuel Figallo", style="Normal",
                      color="595959", spaceAfter=240))
    parts.append(para(
        "formatExcel turns one worksheet into a formatted status report -- "
        "PDF, Word or HTML -- with no Excel and no modules to install. Run "
        ".\\formatExcel.ps1 -Preflight first: it reports in one command "
        "whether this machine can produce a PDF.", spaceAfter=240))

    # --- dependencies ----------------------------------------------------
    parts.append(heading("What you need \u2014 dependencies"))
    parts.append(para(
        "Nothing to install for .docx or .html. No modules, no "
        "Install-Module, no Excel. The script reads .xlsx by unzipping it "
        "and parsing the XML with .NET classes PowerShell already has."))
    parts.append(table(
        ["Component", "Needed?", "Notes"],
        [["PowerShell", "Yes",
          "Windows PowerShell 5.1 or PowerShell 7. No PS7-only syntax; "
          "verified on 7.4.6"],
         ["cgsUtils.ps1", "Yes",
          "Must sit in the SAME FOLDER as formatExcel.ps1"],
         ["Microsoft Word", "PDF only", "The usual route on a CGS desktop"],
         ["LibreOffice", "PDF only",
          "Alternative if Word is absent; found automatically, nothing to "
          "configure"],
         ["Microsoft Excel", "No", "Never opened"],
         ["openpyxl, any module", "No", "--"]],
        widths=[2400, 1400, WIDTH_TOTAL - 2400 - 1400]))
    parts.append(para(
        "PDF needs EITHER Word or LibreOffice, not both. -Preflight says "
        "which one it will use.", spaceAfter=240))

    # --- step 1: set up and unblock --------------------------------------
    # Before the preflight on purpose: the check cannot run until the
    # execution policy lets the script load at all, which is the first
    # thing most people hit.

    parts.append(heading("Step 1 \u2014 Set up and unblock"))
    parts.append(para(
        "a) Put both files in one folder. formatExcel.ps1 dot-sources "
        "cgsUtils.ps1 from its own directory, so one without the other does "
        "not run.", spaceAfter=80))
    parts.extend(code("C:\\code\\python\\cgs_ai\\src\\ps\\\n"
                      "    formatExcel.ps1\n"
                      "    cgsUtils.ps1          <-- REQUIRED, same folder"))
    parts.append(para("b) Open PowerShell in that folder.", spaceAfter=80))
    parts.extend(code("cd C:\\code\\python\\cgs_ai\\src\\ps"))
    parts.append(para(
        "c) Unblock the files. DO THIS FIRST -- a file copied from "
        "SharePoint, Teams, email or a browser carries a hidden "
        "\"downloaded from the internet\" mark, and that mark alone is "
        "what most policies block. Once per copy, not per session, and it "
        "needs no admin rights. Re-copying a file re-marks it, so do this "
        "again every time you replace one.", spaceAfter=80))
    parts.extend(code("Get-ChildItem .\\*.ps1 | Unblock-File"))
    parts.append(para(
        "d) Only if it still will not run: allow the script for this window. "
        "-Scope Process lasts only for this session and changes nothing "
        "permanently. On a managed machine this reports that Group Policy "
        "overrides it -- which is fine, and not a failure, when the "
        "effective policy is already RemoteSigned. Step c is the fix there.",
        spaceAfter=80))
    parts.extend(code("Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process"))
    parts.append(para(
        "Missing cgsUtils.ps1 produces an error that blames the folder "
        "rather than the file -- see \"When it goes wrong\".", spaceAfter=160))
    parts.append(subheading("If you see \"is not digitally signed\""))
    parts.append(para(
        "The policy is AllSigned or RemoteSigned, not Restricted, and the "
        "two have different fixes. Find out which is in force:", spaceAfter=80))
    parts.extend(code("Get-ExecutionPolicy -List"))
    parts.append(para(
        "The winning scope is the first one in the list with a value set. "
        "MachinePolicy and UserPolicy come from Group Policy and override "
        "everything below them -- including -Scope Process and "
        "powershell -ExecutionPolicy Bypass.", spaceAfter=80))
    for line in (
        "RemoteSigned wins, from any scope including Group Policy -- local "
        "unsigned scripts are allowed, so the block is the downloaded mark "
        "and nothing else. Unblock-File alone fixes it, and you do not need "
        "Set-ExecutionPolicy at all. Check the mark with: "
        "Get-Item .\\formatExcel.ps1 -Stream Zone.Identifier",
        "AllSigned wins in CurrentUser or LocalMachine -- "
        "Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process fixes it for that window.",
        "AllSigned wins in MachinePolicy or UserPolicy -- you cannot "
        "override it. Ask desktop support for an exception, have the script "
        "signed with an internal certificate, or use the Python twin, which "
        "no execution policy applies to.",
    ):
        parts.append(bullet(line))
    parts.append(para("", spaceAfter=160))

    # --- step 2: check the machine ---------------------------------------
    parts.append(heading("Step 2 \u2014 Check the machine"))
    parts.append(para("Once the script will load, run this before anything "
                      "else. It needs no workbook, writes no file, and takes "
                      "a second."))
    parts.extend(code(".\\formatExcel.ps1 -Preflight"))
    parts.append(para("On a machine that can do everything:", spaceAfter=80))
    parts.extend(code(
        "INFO  .xlsx reading  : OK (no Excel and no module needed)\n"
        "INFO  .docx output   : OK (written directly as OOXML)\n"
        "INFO  .html output   : OK (written directly)\n"
        "INFO  Microsoft Word : FOUND (COM automation answered)\n"
        "INFO  .pdf output    : OK -- will convert with Microsoft Word\n"
        "INFO  VERDICT: this machine can produce .pdf, .docx and .html."))
    parts.append(para(
        "If no converter is there, the verdict reads \".docx and .html "
        "will work; .pdf will not\", followed by four numbered fixes.",
        spaceAfter=80))
    parts.append(para(
        "Exit code 0 means PDF works. Exit code 3 means only .docx and .html "
        "do -- the same 3 the script returns when a real run cannot convert, "
        "so a scheduled job can branch on it.", spaceAfter=80))
    parts.append(para(
        "This runs the script's own Word and LibreOffice detection, not a "
        "separate check. A hand-written check is a second opinion that can "
        "disagree with the first, and the disagreement would only surface "
        "during a real run.", spaceAfter=240))

    # --- 4 ---------------------------------------------------------------
    parts.append(heading("Step 3 \u2014 Run it"))
    parts.append(subheading("Start with HTML"))
    parts.append(para("It needs no converter, so it proves everything except "
                      "the PDF step.", spaceAfter=80))
    parts.extend(code(
        ".\\formatExcel.ps1 `\n"
        "  -ExcelPath  \"C:\\reports\\kanban.xlsx\" `\n"
        "  -SheetName  \"Consolidated Data\" `\n"
        "  -OutputPath \"C:\\reports\\test.html\""))
    parts.append(para(
        "Expect \"read 26 row(s)\", then \"wrote 26 row(s) x 19 column(s) "
        "[corporate]\". Open it in a browser. If that works, only the PDF "
        "converter is left to prove.", spaceAfter=160))

    parts.append(subheading("Then the real report"))
    parts.extend(code(
        ".\\formatExcel.ps1 `\n"
        "  -ExcelPath   \"C:\\reports\\Kanban_10_30_2026.xlsx\" `\n"
        "  -SheetName   \"Consolidated Data\" `\n"
        "  -Columns     \"Task Name;Bucket;Priority;Created Date;Labels;Notes\" `\n"
        "  -OutputPath  \"C:\\reports\\Governance_KANBAN_10_30_2026.pdf\" `\n"
        "  -format_name corporate `\n"
        "  -Title       \"Data Analytics Infrastructure and Governance\" `\n"
        "  -Subtitle    \"KANBAN status report  |  30 October 2026\" `\n"
        "  -ExcludeWhere \"Bucket=BACKLOG\" `\n"
        "  -SortBy      \"Bucket; Task Name\" `\n"
        "  -MaxCellLength 180"))
    parts.append(para(
        "Watch for \"converting to PDF with Microsoft Word\" -- the only step "
        "with an external dependency.", spaceAfter=160))
    parts.append(para(
        "Use semicolons, not commas, in -Columns and -SortBy. Run with -File "
        "(how SAS and Task Scheduler call it) and PowerShell hands the whole "
        "list over as one string, so a comma inside a column name would "
        "split in the wrong place.", spaceAfter=80))
    parts.append(para(
        "The backtick must be the last character on the line. A trailing "
        "space after it breaks the continuation silently. To avoid them "
        "entirely, splat:", spaceAfter=80))
    parts.extend(code(
        "$opts = @{\n"
        "    ExcelPath    = \"C:\\reports\\kanban.xlsx\"\n"
        "    SheetName    = \"Consolidated Data\"\n"
        "    Columns      = \"Task Name;Bucket;Priority\"\n"
        "    OutputPath   = \"C:\\reports\\report.pdf\"\n"
        "    ExcludeWhere = \"Bucket=BACKLOG\"\n"
        "    SortBy       = \"Bucket; Task Name\"\n"
        "}\n"
        ".\\formatExcel.ps1 @opts"))
    parts.append(para(
        "Formats for -format_name: corporate (default), corporatev2, plain, "
        "minimal. The output extension picks the renderer -- .pdf, .docx, "
        ".html.", spaceAfter=240))

    # --- 5 ---------------------------------------------------------------
    parts.append(heading("Step 4 \u2014 Filter and sort"))
    parts.append(table(
        ["Parameter", "Example", "What it does"],
        [["-ExcludeWhere", "\"Bucket=BACKLOG\"", "Drops matching rows"],
         ["-ExcludeWhere", "\"Bucket=BACKLOG|TODO (Up next)\"",
          "Pipe = several values for one column"],
         ["-ExcludeWhere", "\"Bucket=BACKLOG;Priority=Low\"",
          "Semicolon = several columns; drops a row matching EITHER"],
         ["-ExcludeWhere", "\"Bucket=\"",
          "Empty value = drop the rows where Bucket is blank"],
         ["-IncludeWhere", "\"Bucket=DONE|UNDER REVIEW\"",
          "Keeps ONLY these rows"],
         ["-SortBy", "\"Due date DESC\"",
          "One key; ASC is the default and can be left off"],
         ["-SortBy", "\"Bucket; Task Name DESC\"",
          "Bucket ascending, then Task Name descending within each"]],
        widths=[2100, 3300, WIDTH_TOTAL - 2100 - 3300]))
    parts.append(para(
        "It is -ExcludeWhere, not -Filter, on purpose. In Excel \"filter to "
        "BACKLOG\" means KEEP it -- the opposite of filtering it out. A "
        "parameter named -Filter gets read the Excel way eventually, and "
        "produces a report that is the exact inverse of what was wanted, "
        "with no error to warn you.", spaceAfter=160))

    parts.append(subheading("Things that are easy to get wrong"))
    for line in (
        "Matching is EXACT, not substring. Status=DONE does not also drop "
        "NOT DONE.",
        "Case and padding are ignored on both the column name and the "
        "value, so \"bucket = backlog \" works.",
        "Both run BEFORE columns are chosen, so either can name a column "
        "the report does not display.",
        "Blanks sort LAST in both directions. A DESC column opening with "
        "twenty empty cells is not usefully sorted.",
        "A typo in a column name is an error that lists the real columns. "
        "Silently matching nothing would ship an unfiltered report you "
        "believed was filtered.",
        "A filter that removes every row is an error, not an empty report.",
    ):
        parts.append(bullet(line))

    parts.append(subheading("The report says so on its face"))
    parts.append(para("When a filter or sort is active, the banner carries a "
                      "line you did not ask for:", spaceAfter=80))
    parts.extend(code("Showing 21 of 26 rows.  Excluded: Bucket = BACKLOG  "
                      "Sorted by: Bucket ASC, Task Name ASC"))
    parts.append(para(
        "That is automatic and not optional. A status report that silently "
        "omits rows misleads whoever reads it, and nothing else in the PDF "
        "would reveal it.", spaceAfter=240))

    # --- 6 ---------------------------------------------------------------
    parts.append(heading("Step 5 \u2014 When it goes wrong"))
    parts.append(table(
        ["What you see", "What it means", "Fix"],
        [["The module 'C:\\...\\ps' could not be loaded",
          "cgsUtils.ps1 is not beside formatExcel.ps1. The error blames the "
          "folder, not the missing file",
          "Copy both files, not one"],
         ["cannot be loaded because running scripts is disabled",
          "Execution policy is Restricted",
          "Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process"],
         ["is not digitally signed",
          "Policy is AllSigned or RemoteSigned, not Restricted",
          "Get-ExecutionPolicy -List to see which scope wins; Unblock-File "
          "for RemoteSigned, -ExecutionPolicy Bypass -Scope Process for "
          "AllSigned, and "
          "neither if Group Policy sets it"],
         ["cannot produce a PDF: neither Microsoft Word nor LibreOffice",
          "No converter on this machine",
          "Four numbered fixes print with it; quickest is -OnWordMissing "
          "html, then print to PDF from a browser"],
         ["column(s) not found in this sheet",
          "A name in -Columns, -ExcludeWhere or -SortBy does not match",
          "The message lists the real column names -- copy one"],
         ["the filter removed every row", "The filter matched everything",
          "Check the spelling and case of the value"],
         ["filter clause 'Bucket' has no '='", "Missing the =",
          "Write Column=Value"],
         ["Only one column in the report",
          "-Columns was comma-separated and arrived as one string",
          "Use ;"],
         ["Exit code 2", "A required parameter is missing",
          "-ExcelPath, -SheetName and -OutputPath are all required unless "
          "you pass -Preflight"]],
        widths=[3100, 3100, WIDTH_TOTAL - 3100 - 3100]))
    parts.append(para(
        "Exit codes: 0 success, 2 bad parameters, 3 cannot produce the "
        "output (no converter, or the workbook is missing).", spaceAfter=80))
    parts.append(para(
        "Where the examples live: src/ps/formatExcel.Examples.ps1 carries "
        "twelve worked examples, all commented out, so it is safe to "
        "dot-source. Example 12 is the exact command behind the October "
        "KANBAN report.", spaceAfter=80))
    parts.append(para(
        "The Python twin: src/py/formatExcel.py takes the same options as "
        "--excel-path, --sheet-name, --exclude-where, --sort-by and "
        "--preflight, and produces a byte-identical document. Use whichever "
        "suits the machine.", spaceAfter=80))
    parts.append(para(
        "The live version of this runbook, which the team can comment on, "
        "is at https://claude.ai/artifact/VBWfw7xW1E9ZBAFiXFFQys",
        color="595959", spaceAfter=0))

    return "".join(parts)


def main(argv: Sequence[str]) -> int:
    """CLI entry point. Parameters: argv. Returns: an exit code."""
    if len(argv) != 1:
        print("usage: build_formatexcel_runbook_docx.py <target.docx>",
              file=sys.stderr)
        return EXIT_USAGE
    target = Path(argv[0])
    target.parent.mkdir(parents=True, exist_ok=True)
    writeDocx(buildBody(), str(target), margin=1080, pageNumbers=True)
    print(f"wrote {target} ({target.stat().st_size:,} bytes)")
    return EXIT_OK


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
