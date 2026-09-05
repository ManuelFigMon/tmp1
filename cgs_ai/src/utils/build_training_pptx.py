"""
=====================================================================
  Program Name  : build_training_pptx.py
  Author        : Manuel Figallo
  Purpose       : Build CGS_AI_First_Pipeline_Training.pptx -- the slide
                  version of the claims_report_pipeline_v0_2.py guide,
                  with one slide per step.
  Version       : 1.0beta
  Created       : 2026-08-28
  Last Modified : 2026-08-28

  Dependencies:
    STANDARD LIBRARY ONLY. Design helpers come from build_presentation.py
    and every word of content comes from guide_content.py, so the slides
    and the handout cannot drift apart.

  Usage:
    python src/utils/build_training_pptx.py <source.pptx> <output.pptx>
=====================================================================
"""

from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path
from typing import List, Sequence, Tuple

sys.path.insert(0, str(Path(__file__).resolve().parent))

from build_presentation import (CONTENT_W, LIGHT, MARGIN_X, NAVY,  # noqa: E402
                                NAVY_CIRCLE, NAVY_DEEP, RULE, SLATE,
                                SLIDE_H, SLIDE_W, WHITE, bulletBlock,
                                cardRow, codeBlock, ellipse, paragraph,
                                rect, run, slideShell, textBox)
from guide_content import (CODE_HEADER, CODE_STEPS, CONTACT,  # noqa: E402
                           EXPECTED_OUTPUT, FILENAME, SETUP_STEPS,
                           TROUBLESHOOTING, VALUE_POINTS)

__version__ = "1.0beta"

FOOTER_LEFT = "cgs_ai · Build Your First Pipeline"
FOOTER_CENTER = FILENAME
BODY_TOP = 1430000

#: Filled in once the slide list is known, so page numbers stay honest.
_total = 0


def page(index: int) -> str:
    """Return the page indicator for slide `index`. Parameters: index (int)."""
    return f"{index} / {_total}"


def chrome(eyebrow: str, title: str, subtitle: str, pageNo: str) -> Tuple[str, int]:
    """Build the standard content-slide chrome.

    Parameters: eyebrow, title, subtitle, pageNo (all str).
    Returns: (shape XML, next free shape id).
    """
    shapes = [
        rect(2, "TopBar", 0, 0, SLIDE_W, 128016, NAVY),
        rect(3, "AccentBar", 0, 128016, 2194560, 36576, LIGHT),
        textBox(4, "Eyebrow", MARGIN_X, 181051, 8229600, 256032,
                paragraph(run(eyebrow.upper(), 1100, SLATE, bold=True))),
        textBox(5, "Title", MARGIN_X, 400507, 8229600, 566928,
                paragraph(run(title, 2800, NAVY_DEEP, bold=True))),
        textBox(6, "Subtitle", MARGIN_X, 1004011, CONTENT_W, 310896,
                paragraph(run(subtitle, 1250, SLATE))),
        rect(7, "FooterRule", MARGIN_X, 4828032, CONTENT_W, 10973, RULE),
        textBox(8, "FooterLeft", MARGIN_X, 4855464, 4572000, 219456,
                paragraph(run(FOOTER_LEFT, 900, SLATE))),
        textBox(9, "FooterCenter", 4572000, 4855464, 2880360, 219456,
                paragraph(run(FOOTER_CENTER, 900, SLATE))),
        textBox(10, "PageNo", 7498080, 4855464, 1143000, 219456,
                paragraph(run(pageNo, 900, NAVY, bold=True), align="r")),
    ]
    return "".join(shapes), 11


def note(shapeId: int, lead: str, rest: str, y: int) -> str:
    """Build the bold-lead takeaway line that closes a slide."""
    runs = run(f"{lead}  ", 1150, NAVY_DEEP, bold=True) + run(rest, 1150, SLATE)
    return textBox(shapeId, "Note", MARGIN_X, y, CONTENT_W, 500000,
                   paragraph(runs))


def buildSlides() -> List[str]:
    """Assemble the deck. Returns: list of slide XML strings."""
    slides: List[str] = []

    # ---------- 1. Title ----------
    body = "".join([
        rect(2, "Bg", 0, 0, SLIDE_W, SLIDE_H, NAVY),
        ellipse(3, "BigCircle", 6858000, -457200, 2743200, NAVY_CIRCLE),
        ellipse(4, "SmallCircle", 7772400, 457200, 822960, LIGHT),
        ellipse(5, "CornerCircle", -457200, 3429000, 1828800, NAVY_CIRCLE),
        textBox(6, "Title", 640080, 1188720, 6096000, 1200000,
                paragraph(run("Build Your First cgs_ai Pipeline", 3200,
                              WHITE, bold=True))),
        textBox(7, "Subtitle", 640080, 2560320, 6858000, 400000,
                paragraph(run("Four steps. About ten minutes.", 1600, LIGHT))),
        textBox(8, "Footer", 640080, 3931920, 6858000, 600000,
                paragraph(run(FILENAME, 1100, LIGHT))
                + paragraph(run(f"Presented by {CONTACT}  ·  cgs_ai v1.0beta",
                                1100, LIGHT), spaceBefore=400)),
    ])
    slides.append(slideShell(body))

    # ---------- 2. Why cgs_ai ----------
    body, nid = chrome("Why this matters", "cgs_ai Is Our Own Python Package",
                       "It holds the work we all repeat as functions anyone "
                       "can call by name.", page(2))
    body += bulletBlock(nid, MARGIN_X, BODY_TOP, CONTENT_W, 2700000,
                        [(lead, rest) for lead, rest in VALUE_POINTS],
                        size=1250, gap=700)
    slides.append(slideShell(body))

    # ---------- 3. Before you write any code ----------
    body, nid = chrome("Setup", "Before You Write Any Code",
                       "Three things, then you are ready to type.", page(3))
    body += cardRow(nid, BODY_TOP, [
        {"number": str(index), "title": title, "body": detail}
        for index, (title, detail) in enumerate(SETUP_STEPS, start=1)
    ], height=1600000)
    body += note(nid + 40, "You need one package:",
                 "openpyxl. Install it once with pip install openpyxl.",
                 y=3300000)
    slides.append(slideShell(body))

    # ---------- 4. Start with these two lines ----------
    body, nid = chrome("The code", "Start With These Two Lines",
                       "Every step that follows uses them.", page(4))
    body += codeBlock(nid, MARGIN_X, BODY_TOP, CONTENT_W, 1000000,
                      FILENAME, CODE_HEADER.split("\n"))
    body += note(nid + 2, "SHARE is where cgs_ai lives.",
                 "CGS_AI_HOME is the folder that CONTAINS the package "
                 "folder — one level up from cgs_ai itself.", y=2750000)
    slides.append(slideShell(body))

    # ---------- 5-8. One slide per code step ----------
    for index, step in enumerate(CODE_STEPS):
        lines = step["code"].split("\n")
        # codeBlock = padding + heading + one line per row of code.
        height = 480000 + len(lines) * 190000
        body, nid = chrome(f"Step {step['number']} of {len(CODE_STEPS)}",
                           step["title"], "Type this, then move on.",
                           page(5 + index))
        body += codeBlock(nid, MARGIN_X, BODY_TOP, CONTENT_W, height,
                          f"STEP {step['number']}", lines)
        body += note(nid + 2, "What it does:", step["why"],
                     y=BODY_TOP + height + 240000)
        slides.append(slideShell(body))

    # ---------- 9. Run it ----------
    body, nid = chrome("Run it", "Press Run, or Press F5",
                       "Save the file first. This is what a good run looks "
                       "like.", page(9))
    body += codeBlock(nid, MARGIN_X, BODY_TOP, CONTENT_W, 1250000,
                      "Console output", EXPECTED_OUTPUT.split("\n"))
    body += note(nid + 2, "Then open claims_report.xlsx.",
                 "Navy banner, blue headers, striped rows — and a "
                 "notification waiting in your inbox.", y=3000000)
    slides.append(slideShell(body))

    # ---------- 10. If something goes wrong ----------
    body, nid = chrome("Troubleshooting", "If Something Goes Wrong",
                       "The three failures you are most likely to hit on the "
                       "first run.", page(10))
    body += cardRow(nid, BODY_TOP, [
        {"number": str(index), "title": message, "body": action}
        for index, (message, action) in enumerate(TROUBLESHOOTING, start=1)
    ], height=1700000)
    body += note(nid + 40, "None of these are your code.",
                 "They are the environment. Ask if you hit one.", y=3400000)
    slides.append(slideShell(body))

    # ---------- 11. Questions ----------
    body = "".join([
        rect(2, "Bg", 0, 0, SLIDE_W, SLIDE_H, NAVY),
        ellipse(3, "BigCircle", 6858000, -457200, 2743200, NAVY_CIRCLE),
        ellipse(4, "SmallCircle", 7772400, 457200, 822960, LIGHT),
        ellipse(5, "CornerCircle", -457200, 3429000, 1828800, NAVY_CIRCLE),
        textBox(6, "Title", 640080, 1188720, 6096000, 800000,
                paragraph(run("Questions?", 4000, WHITE, bold=True))),
        textBox(7, "Body", 640080, 2194560, 6400800, 1400000,
                paragraph(run(f"Please contact {CONTACT} with any questions, "
                              f"or with a report you would like added.",
                              1600, LIGHT))
                + paragraph(run("cgs_ai is in beta — feedback on the function "
                                "names and parameters is especially welcome "
                                "now, while they are still easy to change.",
                                1400, LIGHT), spaceBefore=800)),
        textBox(8, "Try", 640080, 4114800, 6858000, 400000,
                paragraph(run("Try it: change FormatType to 'plain' and run "
                              "it again.", 1200, LIGHT))),
    ])
    slides.append(slideShell(body))

    return slides


def buildTrainingDeck(sourcePath: str, outputPath: str) -> str:
    """Write the deck, reusing the source's masters, layouts and theme.

    Parameters:
        sourcePath (str) - a deck to take the design parts from.
        outputPath (str) - destination .pptx.
    Returns: the path written.
    """
    global _total
    source = zipfile.ZipFile(sourcePath)

    # Count the slides first so the page indicators are right.
    _total = 0
    _total = len(buildSlides())
    slides = buildSlides()

    drop = re.compile(r"^ppt/(slides|notesSlides|media)/")
    keep = [n for n in source.namelist() if not drop.match(n)]

    slideRels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Relationships xmlns="http://schemas.openxmlformats.org/'
                 'package/2006/relationships">'
                 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.'
                 'org/officeDocument/2006/relationships/slideLayout" '
                 'Target="../slideLayouts/slideLayout1.xml"/></Relationships>')

    presRels = source.read("ppt/_rels/presentation.xml.rels").decode("utf-8")
    presRels = re.sub(r'<Relationship[^>]*slides/slide\d+\.xml"/>', "", presRels)
    used = [int(m) for m in re.findall(r'Id="rId(\d+)"', presRels)] or [0]
    nextRel = max(used) + 1
    relIds = [f"rId{nextRel + i}" for i in range(len(slides))]
    presRels = presRels.replace("</Relationships>", "".join(
        f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/'
        f'officeDocument/2006/relationships/slide" '
        f'Target="slides/slide{i + 1}.xml"/>'
        for i, rid in enumerate(relIds)) + "</Relationships>")

    presentation = source.read("ppt/presentation.xml").decode("utf-8")
    presentation = re.sub(
        r"<p:sldIdLst>.*?</p:sldIdLst>",
        "<p:sldIdLst>" + "".join(f'<p:sldId id="{256 + i}" r:id="{rid}"/>'
                                 for i, rid in enumerate(relIds))
        + "</p:sldIdLst>", presentation, flags=re.DOTALL)

    contentTypes = source.read("[Content_Types].xml").decode("utf-8")
    contentTypes = re.sub(
        r'<Override PartName="/ppt/(?:slides|notesSlides)/[^"]*"[^>]*/>',
        "", contentTypes)
    contentTypes = contentTypes.replace("</Types>", "".join(
        f'<Override PartName="/ppt/slides/slide{i + 1}.xml" '
        f'ContentType="application/vnd.openxmlformats-officedocument.'
        f'presentationml.slide+xml"/>' for i in range(len(slides))) + "</Types>")

    target = Path(outputPath)
    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as out:
        for name in keep:
            if name == "ppt/presentation.xml":
                out.writestr(name, presentation)
            elif name == "ppt/_rels/presentation.xml.rels":
                out.writestr(name, presRels)
            elif name == "[Content_Types].xml":
                out.writestr(name, contentTypes)
            else:
                out.writestr(name, source.read(name))
        for index, slide in enumerate(slides, start=1):
            out.writestr(f"ppt/slides/slide{index}.xml", slide)
            out.writestr(f"ppt/slides/_rels/slide{index}.xml.rels", slideRels)
    source.close()
    return str(target)


if __name__ == "__main__":
    src = sys.argv[1] if len(sys.argv) > 1 else "CGS_AI_DEMO_v2.pptx"
    dst = sys.argv[2] if len(sys.argv) > 2 else str(
        Path(__file__).resolve().parent.parent.parent
        / "CGS_AI_First_Pipeline_Training.pptx")
    print("wrote", buildTrainingDeck(src, dst))
