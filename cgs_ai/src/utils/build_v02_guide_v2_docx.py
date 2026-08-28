"""
=====================================================================
  Program Name  : build_v02_guide_v2_docx.py
  Author        : Manuel Figallo
  Purpose       : Generate the TWO-PAGE training handout for
                  claims_report_pipeline_v0_2.py.
  Version       : 1.0beta
  Created       : 2026-08-28
  Last Modified : 2026-08-28

  Dependencies:
    STANDARD LIBRARY ONLY. OOXML helpers come from build_readme_docx.py
    and all content comes from guide_content.py, so the handout, the
    one-pager and the slide deck cannot drift apart.

  Description:
    Two pages, so nothing is cramped. Page 1 is why cgs_ai matters and
    the three setup steps; page 2 is the code, one block per step, with a
    sentence on what each block does, then how to run it, what to expect,
    what to do when it fails, and who to ask.

  Usage:
    python src/utils/build_v02_guide_v2_docx.py [output_path]
=====================================================================
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from build_readme_docx import (BLUE, NAVY, bullet, pageBreak,  # noqa: E402
                               para, table, writeDocx)
from guide_content import (CODE_HEADER, CODE_STEPS, CONTACT,  # noqa: E402
                           EXPECTED_OUTPUT, FILENAME, SETUP_STEPS,
                           TROUBLESHOOTING, VALUE_POINTS)

__version__ = "1.0beta"

BODY = 12          # points -- large, for a printed handout
CODE = 10
NOTE = 10


def buildBody() -> str:
    """Assemble the two-page handout body XML. Returns: <w:body> content."""
    parts = []

    # ================= PAGE 1 ================= #
    parts.append(para("Build Your First cgs_ai Pipeline",
                      style="Title", size=22, spaceAfter=40))
    parts.append(para(f"{FILENAME}   ·   a step-by-step guide",
                      style="Subtitle", size=13, spaceAfter=240))

    parts.append(para("Why cgs_ai", style="Heading1", spaceAfter=100))
    parts.append(para(
        "cgs_ai is our own Python package. It holds the work we all repeat "
        "— turning an extract into a clean report, sending it to the people "
        "who need it — as functions anyone can call by name.",
        size=BODY, spaceAfter=160))

    for lead, rest in VALUE_POINTS:
        parts.append(bullet(f"{lead}  {rest}"))
    parts.append(para("", spaceAfter=120))

    parts.append(para("Before you write any code", style="Heading1",
                      spaceAfter=100))
    for index, (title, detail) in enumerate(SETUP_STEPS, start=1):
        parts.append(para(f"{index}.   {title}", size=BODY, bold=True,
                          color=NAVY, spaceAfter=40))
        parts.append(para(f"      {detail}", size=BODY, spaceAfter=150))

    parts.append(para("Now type the code", style="Heading1", spaceAfter=100))
    parts.append(para(
        "Start with these two lines. Every step below uses them.",
        size=BODY, spaceAfter=80))
    parts.append(para(CODE_HEADER, style="Code", size=CODE, mono=True,
                      spaceAfter=150))

    for index, step in enumerate(CODE_STEPS):
        parts.append(para(f"STEP {step['number']}.   {step['title']}",
                          style="Heading3", spaceAfter=50))
        parts.append(para(step["code"], style="Code", size=CODE, mono=True,
                          spaceAfter=60))
        parts.append(para(step["why"], size=NOTE, color=BLUE, spaceAfter=130))
        if index == 0:
            # ================= PAGE 2 ================= #
            parts.append(pageBreak())

    parts.append(para("Run it", style="Heading1", spaceAfter=100))
    parts.append(para(
        "Save the file, then press Run in Visual Studio Code, or press F5:",
        size=BODY, spaceAfter=80))
    parts.append(para(EXPECTED_OUTPUT, style="Code", size=CODE, mono=True,
                      spaceAfter=80))
    parts.append(para(
        "Then open claims_report.xlsx — and check your inbox.",
        size=BODY, spaceAfter=130))

    parts.append(para("If something goes wrong", style="Heading1",
                      spaceAfter=100))
    for message, action in TROUBLESHOOTING:
        parts.append(para(message, size=NOTE, mono=True, color=NAVY,
                          spaceAfter=20))
        parts.append(para(action, size=NOTE, spaceAfter=90))
    parts.append(para("", spaceAfter=60))

    parts.append(para("Questions", style="Heading3", spaceAfter=50))
    parts.append(para(
        f"Please contact {CONTACT} with any questions, or with a report you "
        "would like added. cgs_ai is in beta — feedback is welcome now.",
        size=BODY, spaceAfter=0))

    return "".join(parts)


if __name__ == "__main__":
    output = sys.argv[1] if len(sys.argv) > 1 else str(
        Path(__file__).resolve().parent.parent.parent
        / "docs" / "cgs_ai_First_Pipeline_Guide_v2.docx")
    print("wrote", writeDocx(buildBody(), output, margin=900))
