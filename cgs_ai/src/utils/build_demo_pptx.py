"""
=====================================================================
  Program Name  : build_demo_pptx.py
  Author        : Manuel Figallo
  Purpose       : Build the CGS_AI demo deck -- the GUI-to-code walkthrough
                  that accompanies CGS_AI_Presentation.pptx.
  Version       : 1.0beta
  Created       : 2026-08-27
  Last Modified : 2026-08-27

  Dependencies:
    STANDARD LIBRARY ONLY, like build_presentation.py. Design helpers are
    imported from that module so both decks stay visually identical.

  Description:
    Takes an earlier version of the demo deck as the source, keeps its
    masters, layouts, theme AND its screenshots (ppt/media/*.png), and
    rewrites every slide. The screenshots are the point of the demo, so
    they are re-placed at a corrected aspect ratio rather than recreated.

    Structure: three acts -- the pipeline in the GUI, the three-step
    export, and what the exported code lets you do next.

  Usage:
    python src/utils/build_demo_pptx.py <source.pptx> <output.pptx>
=====================================================================
"""

from __future__ import annotations

import posixpath
import re
import struct
import sys
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

sys.path.insert(0, str(Path(__file__).resolve().parent))

from build_presentation import (CONTENT_W, LIGHT, MARGIN_X, NAVY,  # noqa: E402
                                NAVY_CIRCLE, NAVY_DEEP, NS, RULE, SLATE,
                                SLIDE_H, SLIDE_W, WHITE, bulletBlock, cardRow,
                                codeBlock, ellipse, paragraph, rect, run,
                                slideShell, tableFrame, textBox)

__version__ = "1.0beta"

TOTAL_SLIDES = 12
FOOTER_LEFT = "cgs_ai · Demo"
FOOTER_CENTER = "Prepared by Manuel Figallo"

# Vertical budget for the body of a content slide. The footer rule sits at
# 4828032. NOTE_Y must leave room for a takeaway that wraps to TWO lines
# (~400000 EMU) -- several of them do, and a one-line budget puts the second
# line straight through the footer.
BODY_TOP = 1430000
IMAGE_BOTTOM = 4200000
NOTE_Y = 4270000


def page(index: int) -> str:
    """Return the page indicator for slide `index`. Parameters: index (int)."""
    return f"{index} / {TOTAL_SLIDES}"


def chrome(eyebrow: str, title: str, subtitle: str, pageNo: str) -> Tuple[str, int]:
    """Build the demo slide chrome.

    A local copy of build_presentation.chrome so the demo can carry its own
    footer text. Parameters and return value are otherwise identical.
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


def picture(shapeId: int, relId: str, x: int, y: int, cx: int, cy: int) -> str:
    """Build a <p:pic> shape with a hairline border.

    Parameters:
        relId (str)          - the slide-relationship id of the image part.
        x, y, cx, cy (int)   - geometry in EMU.
    Returns: the picture XML.
    """
    return (f'<p:pic><p:nvPicPr><p:cNvPr id="{shapeId}" name="Screenshot {shapeId}"/>'
            f'<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/>'
            f'</p:nvPicPr><p:blipFill><a:blip r:embed="{relId}"/>'
            f'<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>'
            f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            f'<a:ln w="9525"><a:solidFill><a:srgbClr val="{RULE}"/></a:solidFill>'
            f'</a:ln></p:spPr></p:pic>')


def fitImage(aspect: float, top: int, bottom: int,
             maxWidth: int = 7000000) -> Tuple[int, int, int, int]:
    """Scale an image to fit a band, preserving aspect and centring it.

    Parameters:
        aspect (float)      - image width / height.
        top, bottom (int)   - vertical band in EMU.
        maxWidth (int)      - widest the image may be drawn.
    Returns: (x, y, cx, cy) in EMU.
    """
    maxHeight = bottom - top
    cx = maxWidth
    cy = int(cx / aspect)
    if cy > maxHeight:
        cy = maxHeight
        cx = int(cy * aspect)
    x = (SLIDE_W - cx) // 2
    y = top + (maxHeight - cy) // 2
    return x, y, cx, cy


def note(shapeId: int, lead: str, rest: str, y: int = NOTE_Y) -> str:
    """Build the one-line takeaway that closes a slide.

    Parameters: lead (bold navy phrase), rest (the sentence), y (EMU).
    """
    runs = run(f"{lead}  ", 1150, NAVY_DEEP, bold=True) + run(rest, 1150, SLATE)
    return textBox(shapeId, "Note", MARGIN_X, y, CONTENT_W, 400000,
                   paragraph(runs))


def imageSlide(eyebrow: str, title: str, subtitle: str, pageNo: int,
               aspect: float, leadIn: str, takeaway: str,
               maxWidth: int = 7000000) -> str:
    """Build a screenshot slide: chrome, centred screenshot, one takeaway.

    Parameters:
        aspect (float)  - the screenshot's width/height.
        leadIn (str)    - bold lead of the takeaway line.
        takeaway (str)  - the rest of the takeaway line.
    Returns: the slide XML. The image always uses relationship rId2.
    """
    body, nid = chrome(eyebrow, title, subtitle, page(pageNo))
    x, y, cx, cy = fitImage(aspect, BODY_TOP, IMAGE_BOTTOM, maxWidth)
    body += picture(nid, "rId2", x, y, cx, cy)
    body += note(nid + 1, leadIn, takeaway)
    return slideShell(body)


# =====================================================================
# Slide content
# =====================================================================

def buildSlides(aspects: Dict[str, float]) -> List[Tuple[str, Optional[str]]]:
    """Assemble the deck.

    Parameters:
        aspects (dict) - media part name -> width/height, so screenshots are
                         placed at their true proportions.
    Returns: a list of (slide XML, image part name or None).
    """
    slides: List[Tuple[str, Optional[str]]] = []

    # ---------- 1. Title ----------
    body = "".join([
        rect(2, "Bg", 0, 0, SLIDE_W, SLIDE_H, NAVY),
        ellipse(3, "BigCircle", 6858000, -457200, 2743200, NAVY_CIRCLE),
        ellipse(4, "SmallCircle", 7772400, 457200, 822960, LIGHT),
        ellipse(5, "CornerCircle", -457200, 3429000, 1828800, NAVY_CIRCLE),
        textBox(6, "Title", 640080, 1188720, 6096000, 1000000,
                paragraph(run("CGS_AI Demo", 4000, WHITE, bold=True))),
        textBox(7, "Subtitle", 640080, 2377440, 6858000, 400000,
                paragraph(run("From a process flow to code you can run anywhere.",
                              1600, LIGHT))),
        textBox(8, "Footer", 640080, 3931920, 6858000, 600000,
                paragraph(run("A claims-analysis pipeline built in "
                              "SAS Enterprise Guide", 1100, LIGHT))
                + paragraph(run("Prepared by Manuel Figallo  ·  cgs_ai v1.0beta",
                                1100, LIGHT), spaceBefore=400)),
    ])
    slides.append((slideShell(body), None))

    # ---------- 2. What you will see ----------
    body, nid = chrome("The demo in three acts", "What You Will See",
                       "One pipeline, followed from the point-and-click "
                       "interface to a scheduled job.", page(2))
    body += cardRow(nid, BODY_TOP, [
        {"number": "1", "title": "Build it in the GUI",
         "body": "Three cgs_ai functions wired together in a SAS Enterprise "
                 "Guide process flow. No script written by hand."},
        {"number": "2", "title": "Export it to code",
         "body": "Three clicks turn the process flow into one .sas program "
                 "containing the very same calls."},
        {"number": "3", "title": "Run it anywhere",
         "body": "That program runs outside Enterprise Guide, in Python or "
                 "PowerShell, and on a schedule."},
    ], height=1550000)
    body += note(nid + 40, "The point:",
                 "you never rewrite the logic. The GUI and the code are the "
                 "same three function calls.", y=3220000)
    slides.append((slideShell(body), None))

    # ---------- 3. The pipeline in the GUI ----------
    slides.append((imageSlide(
        "Act 1  ·  the GUI", "The Pipeline an Analyst Builds",
        "Three cgs_ai functions, dragged into one Enterprise Guide process "
        "flow and run end to end.",
        3, aspects["image1"],
        "Nothing custom here.",
        "Each node is a cgs_ai function called with parameters — not a "
        "one-off script written for this project."), "image1"))

    # ---------- 4. What the three nodes do ----------
    body, nid = chrome("The three functions", "What Each Node Actually Does",
                       "Every function exists in Python, PowerShell and as a "
                       "SAS wrapper, with identical names and parameters.",
                       page(4))
    body += tableFrame(nid, MARGIN_X, BODY_TOP,
                       [1950720, 6187440],
                       ["Function", "What it does in this pipeline"],
                       [["scanFileSystem",
                         "Sweeps the claims log shares for keywords and returns one "
                         "row per match, with the matched line and context lines."],
                        ["formatCSV",
                         "Turns that raw result into a styled Excel workbook: navy "
                         "banner, blue header, zebra striping."],
                        ["sendEmail",
                         "Notifies the operations mailbox that the run finished, "
                         "with the row count and the output path."]],
                       headerHeight = 340000, rowHeight = 620000)
    body += note(nid + 1, "Read it top to bottom:",
                 "scan, then format, then alert. That order is the whole "
                 "pipeline.", y=3880000)
    slides.append((slideShell(body), None))

    # ---------- 5-7. The three export steps ----------
    slides.append((imageSlide(
        "Act 2  ·  from GUI to code  ·  step 1 of 3",
        "Export the Process Flow",
        "In the process flow, choose Share  →  “Export all code in "
        "process flow”.",
        5, aspects["image2"],
        "One menu item.",
        "Enterprise Guide writes out everything the flow just ran — in "
        "order, as a single program.",
        maxWidth=7300000), "image2"))

    slides.append((imageSlide(
        "Act 2  ·  from GUI to code  ·  step 2 of 3",
        "Choose Where the Code Lands",
        "Pick a path, keep the default encoding, and leave the four "
        "“include” boxes ticked.",
        6, aspects["image3"],
        "Keep the includes on.",
        "They add the library and WORK assignments, so the exported program "
        "still runs when Enterprise Guide is not the one running it.",
        maxWidth=4200000), "image3"))

    slides.append((imageSlide(
        "Act 2  ·  from GUI to code  ·  step 3 of 3",
        "Review the Generated Code",
        "The result is one self-contained .sas program — the same work, now "
        "in text you can read, diff and version.",
        7, aspects["image4"],
        "This is the handoff.",
        "What was a diagram is now a file an engineer can review and a "
        "scheduler can run.",
        maxWidth=7300000), "image4"))

    # ---------- 8. What is inside the file ----------
    body, nid = chrome("Act 3  ·  what you now have",
                       "What Is Inside the Exported File",
                       "Two parts: a preamble Enterprise Guide adds, and the "
                       "three calls you actually care about.", page(8))
    body += codeBlock(nid, MARGIN_X, BODY_TOP, CONTENT_W, 1950000,
                      "HHH PALS_v2_5.sas",
                      ["/* Code exported from SAS Enterprise Guide */",
                       "%macro enterpriseguide;  ...  %mend;   /* libraries, WORK path */",
                       "",
                       "%scanFileSystem(input_folder_root=..., extract_keyword=...,",
                       "                output_file_path=..., metric_profile=sas_log);",
                       "%formatCSV(InputCsvPath=..., OutputExcelPath=...);",
                       "%sendEmail(To=..., From=..., Subject=..., Body=...);"])
    body += bulletBlock(nid + 2, MARGIN_X, 3500000, CONTENT_W, 1100000, [
        ("The preamble is the only Enterprise Guide part.",
         "Everything below it is ordinary SAS."),
        ("The three calls are the pipeline.",
         "Same function names, same parameter names as the Python and "
         "PowerShell versions."),
    ], size=1150, gap=500)
    slides.append((slideShell(body), None))

    # ---------- 9. The same calls in three languages ----------
    body, nid = chrome("Act 3  ·  why the export matters",
                       "The Same Calls, Three Languages",
                       "Identical function names and parameter names. Only "
                       "the surrounding syntax changes.", page(9))
    # NOTE: code goes in codeBlock, not cardRow -- a card body is one run, so
    # embedded newlines collapse to spaces and the code runs together.
    width = (CONTENT_W - 320040) // 3
    body += codeBlock(nid, MARGIN_X, BODY_TOP, width, 1450000,
                      "SAS  ·  Enterprise Guide",
                      ["%scanFileSystem(", "  extract_keyword=",
                       "    %str(real time),", "  metric_profile=sas_log",
                       ");"])
    body += codeBlock(nid + 2, MARGIN_X + width + 160020, BODY_TOP, width, 1450000,
                      "Python  ·  VS Code",
                      ["cgs_ai.scanFileSystem(", "  extract_keyword=",
                       "    ['real time'],", "  metric_profile='sas_log')"])
    body += codeBlock(nid + 4, MARGIN_X + 2 * (width + 160020), BODY_TOP,
                      width, 1450000, "PowerShell  ·  Scheduler",
                      [".\\scanFileSystem.ps1 `", "  -extract_keyword",
                       "    'real time' `", "  -metric_profile sas_log"])
    body += note(nid + 40, "Nobody translates anything.",
                 "The analyst hands over a working call; the engineer and the "
                 "scheduler run that same call.", y=3120000)
    slides.append((slideShell(body), None))

    # ---------- 10. Scheduling ----------
    body, nid = chrome("Act 3  ·  the payoff", "Now It Can Be Scheduled",
                       "The step that turns a useful click-through into an "
                       "operational service.", page(10))
    body += tableFrame(nid, MARGIN_X, BODY_TOP,
                       [1950720, 2377440, 3810000],
                       ["Environment", "Scheduled by", "How it runs"],
                       [["Desktop", "Manual / F5",
                         "Development only. Not a place to schedule business work."],
                        ["UNIT Server", "Windows Task Scheduler",
                         "A .bat calling the PowerShell or Python entry point."],
                        ["PROD Server", "Task Scheduler / SAS",
                         "Unattended, service account, exit codes the scheduler reads."],
                        ["Snowflake", "Snowflake TASK",
                         "Runs next to the data. No server to patch, no share to mount."]],
                       headerHeight=340000, rowHeight=470000)
    body += note(nid + 1, "Built to be scheduled.",
                 "No cgs_ai function ever prompts — a prompt would hang an "
                 "unattended run forever.", y=3900000)
    slides.append((slideShell(body), None))

    # ---------- 11. Recap ----------
    body, nid = chrome("Recap", "What This Demo Showed",
                       "One pipeline, three acts, no logic rewritten at any "
                       "point.", page(11))
    body += bulletBlock(nid, MARGIN_X, BODY_TOP, CONTENT_W, 2600000, [
        ("The GUI is a first-class way in.",
         "An analyst built a real pipeline without writing a script."),
        ("The export is three clicks.",
         "Share → Export all code → choose a file. No conversion project."),
        ("The code is the same three calls.",
         "scanFileSystem, formatCSV, sendEmail — with the parameters the "
         "analyst chose."),
        ("From there it moves.",
         "Run it outside Enterprise Guide, in Python, in Snowflake, or on a "
         "schedule — unchanged."),
    ], size=1250, gap=650)
    body += note(nid + 1, "The whole loop:",
                 "build it where you are comfortable, then move it where it "
                 "belongs.", y=3900000)
    slides.append((slideShell(body), None))

    # ---------- 12. Beta and feedback ----------
    body, nid = chrome("Where it stands", "cgs_ai Is in Beta",
                       "Documented and covered by automated tests, but early "
                       "— and shaped by how people actually use it.", page(12))
    body += cardRow(nid, BODY_TOP, [
        {"number": "1", "title": "Try it on real work",
         "body": "Pick a routine you run by hand today and point cgs_ai at "
                 "it. That is the fastest way to find what is missing."},
        {"number": "2", "title": "Tell me what is awkward",
         "body": "Function names and parameter names are still easy to "
                 "change. Feedback on those is especially welcome now."},
        {"number": "3", "title": "Ask for a walkthrough",
         "body": "Happy to sit down with anyone and set it up against your "
                 "own folders and data."},
    ], height=1550000)
    body += note(nid + 40, "Next step:",
                 "pick one manual routine you run every week and schedule it. "
                 "That single move buys performance, reliability and your "
                 "time back.", y=3220000)
    slides.append((slideShell(body), None))

    return slides


# =====================================================================
# Package assembly
# =====================================================================

def readAspects(source: zipfile.ZipFile) -> Dict[str, float]:
    """Read each PNG's aspect ratio straight from its IHDR chunk.

    Parameters: source (ZipFile) - the opened source .pptx.
    Returns: {"image1": 3.41, ...}. Standard library only, so no Pillow.
    """
    aspects = {}
    for name in source.namelist():
        if name.startswith("ppt/media/") and name.endswith(".png"):
            data = source.read(name)
            width, height = struct.unpack(">II", data[16:24])
            aspects[Path(name).stem] = width / height
    return aspects


def buildDemo(sourcePath: str, outputPath: str) -> str:
    """Write the demo deck, reusing the source's design parts and screenshots.

    Parameters:
        sourcePath (str) - the previous demo .pptx (supplies masters and media).
        outputPath (str) - destination .pptx.
    Returns: the path written.
    Raises: OSError if the source cannot be read.
    """
    source = zipfile.ZipFile(sourcePath)
    slides = buildSlides(readAspects(source))

    # Keep everything except the source's own slides, their rels and notes.
    drop = re.compile(r"^ppt/(slides|notesSlides)/")
    keep = [n for n in source.namelist() if not drop.match(n)]

    layoutRel = ('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
                 'officeDocument/2006/relationships/slideLayout" '
                 'Target="../slideLayouts/slideLayout1.xml"/>')

    def relsFor(image: Optional[str]) -> str:
        """Build one slide's .rels, adding rId2 for the image when present."""
        extra = ""
        if image:
            extra = ('<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/'
                     'officeDocument/2006/relationships/image" '
                     f'Target="../media/{image}.png"/>')
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
                '2006/relationships">' + layoutRel + extra + '</Relationships>')

    # presentation.xml.rels: drop the old slide relationships, add the new ones.
    presRels = source.read("ppt/_rels/presentation.xml.rels").decode("utf-8")
    presRels = re.sub(r'<Relationship[^>]*slides/slide\d+\.xml"/>', "", presRels)
    used = [int(m) for m in re.findall(r'Id="rId(\d+)"', presRels)] or [0]
    nextRel = max(used) + 1
    slideRelIds = [f"rId{nextRel + i}" for i in range(len(slides))]
    additions = "".join(
        f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/'
        f'officeDocument/2006/relationships/slide" '
        f'Target="slides/slide{i + 1}.xml"/>'
        for i, rid in enumerate(slideRelIds))
    presRels = presRels.replace("</Relationships>", additions + "</Relationships>")

    presentation = source.read("ppt/presentation.xml").decode("utf-8")
    sldIdLst = "".join(f'<p:sldId id="{256 + i}" r:id="{rid}"/>'
                       for i, rid in enumerate(slideRelIds))
    presentation = re.sub(r"<p:sldIdLst>.*?</p:sldIdLst>",
                          f"<p:sldIdLst>{sldIdLst}</p:sldIdLst>",
                          presentation, flags=re.DOTALL)

    # [Content_Types].xml: one override per slide. The ContentType value
    # contains '/', so the trailing attribute must be matched with [^>]*.
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
        for index, (slide, image) in enumerate(slides, start=1):
            out.writestr(f"ppt/slides/slide{index}.xml", slide)
            out.writestr(f"ppt/slides/_rels/slide{index}.xml.rels", relsFor(image))
    source.close()
    return str(target)


if __name__ == "__main__":
    src = sys.argv[1] if len(sys.argv) > 1 else "CGS_AI_DEMO_v1.pptx"
    dst = sys.argv[2] if len(sys.argv) > 2 else str(
        Path(__file__).resolve().parent.parent.parent / "CGS_AI_DEMO_v2.pptx")
    print("wrote", buildDemo(src, dst))
