"""
=====================================================================
  Program Name  : build_onepi_incident_docx.py
  Author        : Manuel Figallo
  Purpose       : Produce a completed One PI Incident Request Form for the
                  Snowflake-to-CGS file export problem.
  Version       : 1.0beta
  Created       : 2026-10-12
  Last Modified : 2026-10-12

  Dependencies:
    STANDARD LIBRARY ONLY. OOXML helpers come from build_readme_docx.py.

  Description:
    Keeps the CMS form's own structure -- the instructions block and
    "Table 1: Ticket Information" -- so the service desk sees the document
    it expects, then expands the description underneath, because the real
    content does not fit legibly in one table cell.

    The description is written to be read by somebody who did not write the
    code: what was attempted, what happened verbatim, why it happens, what
    we need One PI to answer, and which workarounds we already have. The
    asks are numbered so the reply can answer them point by point.

    PLACEHOLDERS are deliberate and marked [LIKE THIS]. The CMS ID, phone
    number, issue time and browser are not mine to guess.

  Usage:
    python src/utils/build_onepi_incident_docx.py [output_path]
=====================================================================
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from build_readme_docx import (BLUE, NAVY, bullet, cell, esc,  # noqa: E402
                               para, table, writeDocx)

__version__ = "1.0beta"

ROOT = Path(__file__).resolve().parent.parent.parent

BODY = 11
ISSUE_DATE = "10/12/2026"
IMPACT = "3"

#: Left column 2900 twips, right column 7180 -- the form's own proportions.
LEFT, RIGHT = 2900, 7180

TICKET_ROWS = [
    ("Impact",
     "3 - Application issue affecting less than five users.\n"
     "(Fewer than five users are affected and a manual workaround exists, "
     "so this is raised as Impact 3 rather than 2.)"),
    ("Contact Name", "Manuel A. Figallo"),
    ("Centers for Medicare & Medicaid Services (CMS) ID", "[INSERT CMS ID]"),
    ("Phone Number", "[INSERT PHONE NUMBER]"),
    ("Email Address", "manuel.figallo@cgsadmin.com"),
    ("Affected Tool(s)",
     "Snowsight\n"
     "Also affected: Snowflake Workspaces and Snowflake Notebooks, which are "
     "not listed separately on this form."),
    ("Do you need System or Data Integration Support?",
     "Yes\n"
     "This request is specifically about moving a file between One PI "
     "Snowflake and an external (CGS) environment, which is an integration "
     "and data-egress question rather than an application fault."),
    ("Issue Date(s)", ISSUE_DATE),
    ("Issue Time(s) (including time zone)", "[INSERT TIME] ET"),
    ("Browser(s) Used", "[INSERT BROWSER AND VERSION]"),
    ("Detailed Problem or Request Description "
     "(including error messages if applicable)",
     "A file produced in a One PI Snowflake workspace cannot be written to a "
     "workstation or organisation file share from inside Snowflake, because "
     "the notebook kernel runs on Snowflake-managed compute rather than on "
     "the user's machine.\n"
     "We are asking One PI to confirm the supported export method.\n"
     "Full detail, the verbatim error, our questions and the workarounds we "
     "have already identified are set out below, under \"Detailed Problem or "
     "Request Description\"."),
]

#: The verbatim console output, quoted so the service desk sees the real text.
ERROR_TEXT = (
    "ERROR  cannot write \\\\Client\\C$\\Temp\\synthetic_medicare_claims.csv "
    "from a Linux host.\n"
    "ERROR    WHY: the kernel is executing on a Linux host, not on your "
    "machine. A path beginning \\\\Client\\ or a mapped drive letter only "
    "exists inside your own Windows or Citrix session, and there is no "
    "network route to it from here.\n"
    "ERROR    FIX 1 (recommended): run this function ON the machine that owns "
    "the target -- a CGS Windows desktop or your Citrix session -- with "
    "snowflake-connector-python installed.\n"
    "ERROR    FIX 2: in the notebook, copy to the kernel's own disk first, "
    "then bring it down with the notebook's download control.")

QUESTIONS = [
    ("What is the supported method to export a file from a One PI Snowflake "
     "workspace or stage to a user's workstation, or to an organisation file "
     "share outside One PI?"),
    ("Are the PUT and GET commands available from Snowsight, or do they "
     "require a local client (SnowSQL, snowflake-connector-python, or the "
     "ODBC/JDBC driver)? If a local client is required, is installing one on "
     "a One PI-accessible workstation permitted, and through what process?"),
    ("Is there an approved external stage -- cloud storage that One PI can "
     "write to and that an external environment may read? If so, what is the "
     "request process and are there size or retention limits?"),
    ("Is there a row or file-size limit on the Snowsight \"Download results\" "
     "option, and is that the sanctioned path for datasets of this size?"),
    ("Does CMS policy permit export of data of this type from One PI to a "
     "contractor environment, and if so under what controls? We will not "
     "move any non-synthetic data until this is confirmed in writing."),
]

WORKAROUNDS = [
    ("A. Run the transfer from the workstation instead of from Snowflake",
     "Status: PREFERRED, BLOCKED.",
     "The export runs on the user's own machine, which can see both One PI "
     "Snowflake and the destination folder. This is the only option that can "
     "be scheduled and run unattended. It is blocked until question 2 is "
     "answered, because it needs a Snowflake client installed locally."),
    ("B. Write to the notebook's own storage, then download from the notebook",
     "Status: PARTLY CONFIRMED, MANUAL.",
     "Writing the file to the kernel's own filesystem is confirmed working. "
     "Retrieving it from there relies on the notebook's download control, "
     "which we have not yet exercised in One PI -- please confirm it is "
     "available. Either way it is a per-run manual step, so it does not meet "
     "the recurring-delivery need."),
    ("C. Snowsight \"Download results\" from a worksheet query",
     "Status: NOT YET TESTED BY US, MANUAL.",
     "Run the query in Snowsight and download the result grid as CSV. This "
     "is a standard Snowsight feature rather than something we have "
     "verified in One PI. Simplest option for a small extract; question 4 "
     "asks what the size limit is, which decides whether it is viable for "
     "our volumes."),
    ("D. Hand off through an approved external stage",
     "Status: BEST FIT FOR AUTOMATION, NEEDS APPROVAL.",
     "One PI writes the file to approved cloud storage; our environment "
     "collects it from there. Nothing has to reach across from Snowflake "
     "compute to a private network, so it avoids the problem entirely. "
     "Depends on the answer to question 3."),
]


def heading(text: str) -> str:
    """Section heading inside the expanded description."""
    return para(text, style="Heading2", size=13, spaceAfter=90)


def labelled(label: str, text: str) -> str:
    """A bold navy lead-in followed by body text, in one paragraph."""
    return (f'<w:p><w:pPr><w:spacing w:after="110"/></w:pPr>'
            f'<w:r><w:rPr><w:b/><w:sz w:val="{BODY * 2}"/>'
            f'<w:color w:val="{NAVY}"/></w:rPr>'
            f'<w:t xml:space="preserve">{esc(label)}  </w:t></w:r>'
            f'<w:r><w:rPr><w:sz w:val="{BODY * 2}"/></w:rPr>'
            f'<w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>')


def numbered(index: int, text: str) -> str:
    """A numbered ask, so the reply can answer point by point."""
    return (f'<w:p><w:pPr><w:spacing w:after="110"/>'
            f'<w:ind w:left="460" w:hanging="460"/></w:pPr>'
            f'<w:r><w:rPr><w:b/><w:sz w:val="{BODY * 2}"/>'
            f'<w:color w:val="{BLUE}"/></w:rPr>'
            f'<w:t xml:space="preserve">{index}.   </w:t></w:r>'
            f'<w:r><w:rPr><w:sz w:val="{BODY * 2}"/></w:rPr>'
            f'<w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>')


def formTable() -> str:
    """Rebuild Table 1 with the answers filled in.

    Returns: the table XML.

    Built by hand rather than with table(): the form has a two-column
    Category/Item layout with no zebra striping, and keeping it looking like
    the original matters more here than house style.
    """
    borders = ("<w:tblBorders>" + "".join(
        f'<w:{edge} w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/>'
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV")
    ) + "</w:tblBorders>")
    xml = (f'<w:tbl><w:tblPr><w:tblW w:w="{LEFT + RIGHT}" w:type="dxa"/>'
           f'{borders}<w:tblLayout w:type="fixed"/></w:tblPr>'
           f'<w:tblGrid><w:gridCol w:w="{LEFT}"/>'
           f'<w:gridCol w:w="{RIGHT}"/></w:tblGrid>')
    xml += ("<w:tr><w:trPr><w:tblHeader/></w:trPr>"
            + cell("Category", LEFT, bold=True, fill=NAVY, color="FFFFFF")
            + cell("Item", RIGHT, bold=True, fill=NAVY, color="FFFFFF")
            + "</w:tr>")
    for label, value in TICKET_ROWS:
        xml += ("<w:tr>" + cell(label, LEFT, bold=True, fill="EEF3FB")
                + cell(value, RIGHT) + "</w:tr>")
    return xml + "</w:tbl>" + para("", spaceAfter=200)


def buildBody() -> str:
    """Assemble the completed form. Returns: <w:body> content."""
    parts = []

    parts.append(para("One PI Incident Request Form", style="Title", size=22,
                      spaceAfter=120))
    parts.append(para(
        f"Save this form as:  One PI Incident Request {ISSUE_DATE[6:]}-"
        f"{ISSUE_DATE[0:2]}-{ISSUE_DATE[3:5]}", size=BODY, spaceAfter=50))
    parts.append(para(
        f"Email it to cms_it_service_desk@cms.hhs.gov with the subject line:  "
        f"One PI - Impact {IMPACT}", size=BODY, spaceAfter=220))

    parts.append(para("Table 1: Ticket Information", style="Heading1",
                      spaceAfter=120))
    parts.append(formTable())

    parts.append(para("Detailed Problem or Request Description",
                      style="Heading1", spaceAfter=140))

    parts.append(heading("1.  What we are trying to do"))
    parts.append(para(
        "We need to take a CSV produced in a One PI Snowflake workspace and "
        "place it on a CGS file location, so that existing downstream "
        "reporting can read it. The intention is for this to run on a "
        "schedule rather than be repeated by hand each cycle.",
        size=BODY, spaceAfter=170))

    parts.append(heading("2.  What we attempted"))
    parts.append(para(
        "A small Python utility was run inside a Snowflake notebook in One "
        "PI, with these two parameters:", size=BODY, spaceAfter=90))
    parts.append(para(
        "SOURCE_LOCATION = snow://workspace/USER$FHER.PUBLIC.DEFAULT$/"
        "versions/head/data/synthetic_medicare_claims.csv\n"
        "TARGET_LOCATION = \\\\Client\\C$\\Temp\\synthetic_medicare_claims.csv",
        style="Code", size=9, mono=True, spaceAfter=110))
    parts.append(labelled(
        "Note on the data:",
        "the file used for this test is SYNTHETIC claims data created for "
        "testing. No real beneficiary or claims data was involved in "
        "reproducing this issue."))
    parts.append(para("", spaceAfter=60))

    parts.append(heading("3.  What happened (verbatim)"))
    parts.append(para(ERROR_TEXT, style="Code", size=9, mono=True,
                      spaceAfter=170))

    parts.append(heading("4.  Why we believe this happens"))
    parts.append(para(
        "A Snowflake notebook kernel executes on Snowflake-managed compute, "
        "not on the user's workstation and not inside the user's Citrix "
        "session. Two consequences follow, and we believe both are expected "
        "behaviour rather than a fault:",
        size=BODY, spaceAfter=110))
    parts.append(bullet(
        "A path beginning \\\\Client\\ is a Citrix client-drive redirection. "
        "It exists only inside the user's own session, so it is not "
        "addressable from anywhere else, including Snowflake compute."))
    parts.append(bullet(
        "Organisation file servers are on a private network with no route "
        "from Snowflake compute, so a UNC path to one is equally "
        "unreachable."))
    parts.append(para(
        "We are therefore not reporting this as a defect in One PI. We are "
        "asking what the supported export route is, so that we build against "
        "it rather than around it.", size=BODY, spaceAfter=170))

    parts.append(heading("5.  What we need from One PI"))
    for index, question in enumerate(QUESTIONS, start=1):
        parts.append(numbered(index, question))
    parts.append(para("", spaceAfter=80))

    parts.append(heading("6.  Workarounds already identified"))
    parts.append(para(
        "We are not blocked on urgent work: option B gets a file out today, "
        "and option C should. Both are manual, which is why the request "
        "above matters for the recurring case. Where we have not verified "
        "something ourselves, it is marked as such below.",
        size=BODY, spaceAfter=130))
    for title, status, detail in WORKAROUNDS:
        parts.append(para(title, size=BODY, bold=True, color=NAVY,
                          spaceAfter=30))
        parts.append(para(status, size=BODY - 1, bold=True, color=BLUE,
                          spaceAfter=30))
        parts.append(para(detail, size=BODY, spaceAfter=150))

    parts.append(heading("7.  Business impact"))
    parts.append(para(
        "Fewer than five users are affected. The immediate need is met by "
        "the manual workaround, so no reporting deadline is currently at "
        "risk. The cost is recurring manual effort each cycle, and the "
        "inability to schedule a delivery that otherwise runs unattended. "
        "That is why this is submitted as Impact 3.",
        size=BODY, spaceAfter=170))

    parts.append(heading("8.  What a good resolution looks like"))
    parts.append(para(
        "A written answer to the five questions in section 5, naming the "
        "supported export route for recurring file delivery out of One PI "
        "Snowflake. If that route is an approved external stage or a locally "
        "installed Snowflake client, we would also like to know the request "
        "process so we can raise it.", size=BODY, spaceAfter=0))

    return "".join(parts)


if __name__ == "__main__":
    output = sys.argv[1] if len(sys.argv) > 1 else str(
        ROOT / "docs" / "One PI Incident Request 2026-10-12.docx")
    print("wrote", writeDocx(buildBody(), output, margin=1080))
