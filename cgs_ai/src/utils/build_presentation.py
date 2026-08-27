"""
=====================================================================
  Program Name  : build_presentation.py
  Author        : Manuel Figallo
  Purpose       : Generate the cgs_ai presentation from the corporate
                  PowerPoint template.
  Version       : 1.0beta
  Created       : 2026-08-26
  Last Modified : 2026-08-26

  Dependencies:
    STANDARD LIBRARY ONLY. A .pptx is a zip of OOXML parts, so this reuses
    the template package (theme, masters, layouts, media) and writes new
    slide parts directly, rather than adding python-pptx as a dependency.

  Description:
    Reuses TEMPLATE1_0.pptx: its slide masters, layouts, theme and media are
    kept untouched, so the deck inherits the corporate look exactly. Only
    the slide parts are replaced. Every colour, position and font size below
    was measured from the template's own slides so new slides are
    indistinguishable in style from the originals.

  Input Parameters (required first):
    templatePath (REQUIRED, str) - source .pptx template.
    outputPath   (REQUIRED, str) - destination .pptx.

  Usage:
    python src/utils/build_presentation.py <template.pptx> [output.pptx]
=====================================================================
"""

from __future__ import annotations

import re
import shutil
import sys
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

__version__ = "1.0beta"

# --- geometry (EMU) and palette, measured from the template -----------------
SLIDE_W, SLIDE_H = 9144000, 5143500
MARGIN_X, CONTENT_W = 502920, 8138160

NAVY = "1E2761"        # top bar, page number
NAVY_DEEP = "232A4D"   # slide titles
NAVY_CIRCLE = "2A3678" # title-slide circles
LIGHT = "CADCFC"       # accent bar, title-slide subtitle
SLATE = "5A6685"       # eyebrow, body text
RULE = "D7DFF0"        # hairlines, table borders
WHITE = "FFFFFF"

NS = ('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"')

FOOTER_LEFT = "cgs_ai · From GUI to Code"
FOOTER_CENTER = "Prepared by Manuel Figallo"


def esc(text: str) -> str:
    """XML-escape a string. Parameters: text (str). Returns: str."""
    return (str(text).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;"))


def run(text: str, size: int, color: str, bold: bool = False,
        italic: bool = False, mono: bool = False) -> str:
    """Build one <a:r> text run.

    Parameters:
        text (str)  - run text.
        size (int)  - font size in hundredths of a point (1800 = 18pt).
        color (str) - RRGGBB hex.
        bold, italic, mono - styling flags.
    Returns: the run XML.
    """
    face = "Consolas" if mono else "Calibri"
    weight = ' b="1"' if bold else ""
    slant = ' i="1"' if italic else ""
    return (f'<a:r><a:rPr lang="en-US" sz="{size}"'
            f'{weight}{slant} dirty="0">'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:latin typeface="{face}"/><a:cs typeface="{face}"/>'
            f'</a:rPr><a:t>{esc(text)}</a:t></a:r>')


def paragraph(runs: str, align: str = "l", spaceBefore: int = 0,
              bulletChar: str = "", indent: int = 0) -> str:
    """Wrap runs in an <a:p>.

    Parameters:
        runs (str)        - concatenated run XML.
        align (str)       - l, ctr or r.
        spaceBefore (int) - space before the paragraph, in points*100.
        bulletChar (str)  - a bullet glyph, or '' for none.
        indent (int)      - left indent in EMU.
    Returns: the paragraph XML.
    """
    bullet = (f'<a:buChar char="{esc(bulletChar)}"/>' if bulletChar
              else "<a:buNone/>")
    return (f'<a:p><a:pPr marL="{indent}" indent="{-indent if bulletChar else 0}" '
            f'algn="{align}"><a:spcBef><a:spcPts val="{spaceBefore}"/></a:spcBef>'
            f'{bullet}</a:pPr>{runs}</a:p>')


def textBox(shapeId: int, name: str, x: int, y: int, cx: int, cy: int,
            paragraphs: str, anchor: str = "t", wrap: bool = True) -> str:
    """Build a text-box shape. Parameters: geometry (EMU) and paragraph XML."""
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{shapeId}" name="{name}"/>'
            f'<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>'
            f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            f'<p:txBody><a:bodyPr wrap="{"square" if wrap else "none"}" '
            f'lIns="0" tIns="0" rIns="0" bIns="0" anchor="{anchor}">'
            f'<a:normAutofit/></a:bodyPr><a:lstStyle/>{paragraphs}</p:txBody></p:sp>')


def rect(shapeId: int, name: str, x: int, y: int, cx: int, cy: int,
         fill: str, radius: bool = False) -> str:
    """Build a filled rectangle. Parameters: geometry and fill colour."""
    geom = "roundRect" if radius else "rect"
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{shapeId}" name="{name}"/>'
            f'<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>'
            f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            f'<a:prstGeom prst="{geom}"><a:avLst/></a:prstGeom>'
            f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
            f'<a:ln><a:noFill/></a:ln></p:spPr>'
            f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/>'
            f'</a:p></p:txBody></p:sp>')


def ellipse(shapeId: int, name: str, x: int, y: int, size: int, fill: str) -> str:
    """Build a filled circle (used on the title slide)."""
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{shapeId}" name="{name}"/>'
            f'<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>'
            f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{size}" cy="{size}"/></a:xfrm>'
            f'<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>'
            f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
            f'<a:ln><a:noFill/></a:ln></p:spPr>'
            f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/>'
            f'</a:p></p:txBody></p:sp>')


def tableCell(text: str, size: int, color: str, bold: bool = False,
              fill: Optional[str] = None, mono: bool = False,
              align: str = "l") -> str:
    """Build one table cell with the template's border and margin style."""
    border = ("".join(
        f'<a:ln{edge} w="6350" cap="flat" cmpd="sng" algn="ctr">'
        f'<a:solidFill><a:srgbClr val="{RULE}"/></a:solidFill>'
        f'<a:prstDash val="solid"/><a:round/></a:ln{edge}>'
        for edge in ("L", "R", "T", "B")))
    shade = (f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
             if fill else "<a:noFill/>")
    lines = str(text).split("\n")
    paras = "".join(paragraph(run(line, size, color, bold=bold, mono=mono),
                              align=align) for line in lines)
    return (f'<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>{paras}</a:txBody>'
            f'<a:tcPr marL="76200" marR="76200" marT="38100" marB="38100" '
            f'anchor="ctr">{border}{shade}</a:tcPr></a:tc>')


def tableFrame(shapeId: int, x: int, y: int, widths: Sequence[int],
               header: Sequence[str], rows: Sequence[Sequence[str]],
               headerHeight: int = 300000, rowHeight: int = 250000,
               bodySize: int = 1000, headerSize: int = 1050,
               monoColumns: Sequence[int] = ()) -> str:
    """Build a table graphic frame styled like the template's tables.

    Parameters:
        widths (sequence)      - column widths in EMU.
        header (sequence)      - header labels (navy fill, white bold text).
        rows (sequence)        - data rows; alternate rows are shaded.
        monoColumns (sequence) - column indexes rendered in Consolas.
    Returns: the graphicFrame XML.
    """
    grid = "".join(f'<a:gridCol w="{w}"/>' for w in widths)
    xml = (f'<p:graphicFrame><p:nvGraphicFramePr>'
           f'<p:cNvPr id="{shapeId}" name="Table {shapeId}"/>'
           f'<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/>'
           f'</p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>'
           f'<p:xfrm><a:off x="{x}" y="{y}"/>'
           f'<a:ext cx="{sum(widths)}" cy="{headerHeight + rowHeight * len(rows)}"/>'
           f'</p:xfrm><a:graphic><a:graphicData '
           f'uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
           f'<a:tbl><a:tblPr firstRow="1"/><a:tblGrid>{grid}</a:tblGrid>')
    xml += f'<a:tr h="{headerHeight}">' + "".join(
        tableCell(h, headerSize, WHITE, bold=True, fill=NAVY)
        for h in header) + "</a:tr>"
    for index, row in enumerate(rows):
        fill = RULE if index % 2 else WHITE
        xml += f'<a:tr h="{rowHeight}">' + "".join(
            tableCell(value, bodySize, NAVY_DEEP if i == 0 else SLATE,
                      bold=(i == 0), fill=fill, mono=(i in monoColumns))
            for i, value in enumerate(row)) + "</a:tr>"
    return xml + "</a:tbl></a:graphicData></a:graphic></p:graphicFrame>"


def slideShell(shapes: str) -> str:
    """Wrap shape XML in a complete <p:sld> part."""
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<p:sld {NS}><p:cSld><p:spTree>'
            f'<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/>'
            f'</p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/>'
            f'<a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/>'
            f'</a:xfrm></p:grpSpPr>{shapes}</p:spTree></p:cSld>'
            f'<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>')


def chrome(eyebrow: str, title: str, subtitle: str, page: str) -> Tuple[str, int]:
    """Build the standard content-slide chrome.

    Parameters:
        eyebrow (str)  - small caps line above the title.
        title (str)    - slide title.
        subtitle (str) - one-line description under the title.
        page (str)     - page indicator, e.g. "3 / 14".
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
                paragraph(run(page, 900, NAVY, bold=True), align="r")),
    ]
    return "".join(shapes), 11


def bulletBlock(shapeId: int, x: int, y: int, cx: int, cy: int,
                items: Sequence[Any], size: int = 1250,
                gap: int = 700) -> str:
    """Build a bulleted text block.

    Parameters:
        items (sequence) - each item is a string, or a (lead, rest) tuple
                           where `lead` is rendered bold navy.
    Returns: the shape XML.
    """
    paras = []
    for index, item in enumerate(items):
        space = 0 if index == 0 else gap
        if isinstance(item, tuple):
            lead, rest = item
            runs = (run(f"{lead}  ", size, NAVY_DEEP, bold=True)
                    + run(rest, size, SLATE))
        else:
            runs = run(str(item), size, SLATE)
        paras.append(paragraph(runs, spaceBefore=space,
                               bulletChar="▪", indent=180000))
    return textBox(shapeId, "Bullets", x, y, cx, cy, "".join(paras))


def cardRow(startId: int, y: int, cards: Sequence[Dict[str, str]],
            height: int = 1500000) -> str:
    """Build a row of numbered cards, like the template's phase slide.

    Parameters:
        cards (sequence of dict) - each with number, title and body keys.
    Returns: the shape XML for the whole row.
    """
    count = len(cards)
    gutter = 160020
    width = (CONTENT_W - gutter * (count - 1)) // count
    shapes, shapeId = [], startId
    for index, card in enumerate(cards):
        x = MARGIN_X + index * (width + gutter)
        shapes.append(rect(shapeId, f"Card{index}", x, y, width, height, "F4F7FC"))
        shapeId += 1
        shapes.append(rect(shapeId, f"CardBar{index}", x, y, width, 45720, NAVY))
        shapeId += 1
        shapes.append(ellipse(shapeId, f"CardNum{index}", x + 137160,
                              y + 137160, 320040, NAVY))
        shapeId += 1
        shapes.append(textBox(shapeId, f"CardNumT{index}", x + 137160,
                              y + 201168, 320040, 220000,
                              paragraph(run(card["number"], 1100, WHITE,
                                            bold=True), align="ctr")))
        shapeId += 1
        shapes.append(textBox(shapeId, f"CardTitle{index}", x + 137160,
                              y + 548640, width - 274320, 320040,
                              paragraph(run(card["title"], 1250, NAVY_DEEP,
                                            bold=True))))
        shapeId += 1
        shapes.append(textBox(shapeId, f"CardBody{index}", x + 137160,
                              y + 914400, width - 274320, height - 1005840,
                              paragraph(run(card["body"], 1000, SLATE))))
        shapeId += 1
    return "".join(shapes)


def codeBlock(shapeId: int, x: int, y: int, cx: int, cy: int,
              heading: str, lines: Sequence[str]) -> str:
    """Build a shaded monospace code block with a heading."""
    shapes = [rect(shapeId, "CodeBg", x, y, cx, cy, "F4F7FC")]
    paras = [paragraph(run(heading, 1000, NAVY, bold=True))]
    for line in lines:
        paras.append(paragraph(run(line, 900, NAVY_DEEP, mono=True),
                               spaceBefore=300))
    shapes.append(textBox(shapeId + 1, "CodeTx", x + 137160, y + 137160,
                          cx - 274320, cy - 274320, "".join(paras)))
    return "".join(shapes)


# =====================================================================
# Slide content
# =====================================================================

def buildSlides() -> List[str]:
    """Assemble every slide. Returns: list of slide XML strings."""
    slides: List[str] = []
    total = 14

    def page(n: int) -> str:
        return f"{n} / {total}"

    # ---------- 1. Title ----------
    shapes = [
        rect(2, "Bg", 0, 0, SLIDE_W, SLIDE_H, NAVY),
        ellipse(3, "Circle1", 6949440, -1280160, 4023360, NAVY_CIRCLE),
        ellipse(4, "Circle2", -1280160, 3566160, 3291840, NAVY_CIRCLE),
        ellipse(5, "Dot", 7818120, 502920, 822960, LIGHT),
        textBox(6, "Title", 640080, 1417320, 7863840, 1000000,
                paragraph(run("From GUI to Code using CGS_AI", 4000, WHITE, bold=True))),
        textBox(7, "Subtitle", 640080, 2540000, 7315200, 457200,
                paragraph(run("One toolkit. Three languages. Four environments.",
                              2000, LIGHT))),
        textBox(8, "Foot", 640080, 4069080, 7498080, 640080,
                paragraph(run("A claims-analysis toolkit for SAS, Python and PowerShell",
                              1300, LIGHT))
                + paragraph(run("Prepared by Manuel Figallo  ·  cgs_ai v1.0beta",
                                1300, LIGHT), spaceBefore=400)),
    ]
    slides.append(slideShell("".join(shapes)))

    # ---------- 2. Agenda ----------
    body, nid = chrome("Agenda", "What We Will Cover",
                       "From why functions matter, through the four environments, to a working pipeline.",
                       page(2))
    body += tableFrame(nid, MARGIN_X, 1440000, [640080, 4200000, 3298080],
                       ["#", "Topic", "What you take away"],
                       [["1", "Introduction to functions",
                         "Why reusable functions matter"],
                        ["2", "Functions are the new PROCs",
                         "A familiar idea in a new place"],
                        ["3", "GUI and programming environments",
                         "SAS EG, VS Code, Snowflake"],
                        ["4", "The function catalog",
                         "Eleven functions, one interface"],
                        ["5", "Four environments",
                         "Desktop, UNIT, PROD, Snowflake"],
                        ["6", "Moving up the chain",
                         "Performance, reliability, security, automation"],
                        ["7", "Use case and scheduling",
                         "A pipeline running unattended"],
                        ["8", "Conclusions",
                         "The foundation for AI pipelines"]],
                       headerHeight = 280000, rowHeight = 245000)
    slides.append(slideShell(body))

    # ---------- 3. Why functions matter ----------
    body, nid = chrome("Introduction to functions", "Why It Matters",
                       "The same logic, written once, callable from anywhere, producing identical output.",
                       page(3))
    body += bulletBlock(nid, MARGIN_X, 1470000, CONTENT_W, 3100000, [
        ("Write once, run everywhere.",
         "One implementation serves the analyst in SAS, the engineer in Python and the scheduler on the server."),
        ("No more copy-paste drift.",
         "Ten variations of the same log-scanning code become one function with parameters."),
        ("Parameters replace editing.",
         "Changing a folder, a keyword or a date range is an argument, not a code change requiring review."),
        ("Testable and governed.",
         "95 passing automated tests protect behaviour; a change that breaks the contract fails before it ships."),
        ("A path from prototype to production.",
         "The same call you try on your desktop is the one the scheduler runs at 2 a.m."),
    ], size=1200, gap=900)
    slides.append(slideShell(body))

    # ---------- 4. Functions are the new PROCs ----------
    body, nid = chrome("A familiar idea", "Functions Are the Equivalent of SAS PROCs",
                       "You already think this way. PROC SORT does one job well and takes options; so does every cgs_ai function.",
                       page(4))
    body += tableFrame(nid, MARGIN_X, 1450000, [2560320, 2788920, 2788920],
                       ["Concept", "In SAS you write", "In cgs_ai you call"],
                       [["Do one job well", "PROC SORT DATA=claims;", "scanFileSystem(...)"],
                        ["Options control it", "BY claimId; OUT=sorted;",
                         "extract_keyword=, lines_above="],
                        ["Predictable output", "A sorted data set",
                         "A CSV or Excel with fixed columns"],
                        ["Reusable everywhere", "Any SAS program",
                         "SAS, Python, PowerShell, Snowflake"],
                        ["Documented contract", "SAS documentation",
                         "Header block + README.docx"]],
                       headerHeight=290000, rowHeight=300000, monoColumns=(1, 2))
    body += textBox(90, "Note", MARGIN_X, 3560000, CONTENT_W, 500000,
                    paragraph(run("The difference: ", 1150, NAVY_DEEP, bold=True)
                              + run("a PROC only runs inside SAS. A cgs_ai function runs "
                                    "inside SAS, at a command line, in a Python notebook and "
                                    "in Snowflake — with the same name and the same parameters.",
                                    1150, SLATE)))
    slides.append(slideShell(body))

    # ---------- 5. GUI and programming ----------
    body, nid = chrome("Meant for both worlds", "GUI and Programming Environments",
                       "Nobody has to abandon the tool they already use. The function is the common denominator.",
                       page(5))
    body += cardRow(nid, 1450000, [
        {"number": "1", "title": "GUI — SAS Enterprise Guide",
         "body": "The analyst stays in the interface they know. A %macro call in a "
                 "program node runs the same function, with results returned to the "
                 "project. No Python knowledge required."},
        {"number": "2", "title": "Programming — VS Code",
         "body": "The engineer imports the package, gets autocomplete, sets "
                 "breakpoints and steps through the code. Full debugging on the "
                 "same logic the analyst just called."},
        {"number": "3", "title": "Programming — Snowflake Worksheets",
         "body": "The cloud path. The same functions run next to the data as a "
                 "Snowpark procedure, with no download and no local file share."},
    ], height=1900000)
    body += textBox(95, "Bridge", MARGIN_X, 3560000, CONTENT_W, 500000,
                    paragraph(run("The bridge:  ", 1150, NAVY_DEEP, bold=True)
                              + run("src/sas holds wrapper macros only. They launch the "
                                    "PowerShell or Python implementation, so the GUI user and "
                                    "the programmer are running the very same code.",
                                    1150, SLATE)))
    slides.append(slideShell(body))

    # ---------- 6. Function catalog 1 ----------
    body, nid = chrome("The function catalog  ·  1 of 2",
                       "Scanning, Data Access and Reporting",
                       "Every function exists in Python, PowerShell and as a SAS wrapper, with identical names and parameters.",
                       page(6))
    body += tableFrame(nid, MARGIN_X, 1420000, [2011680, 6126480],
                       ["Function", "Description"],
                       [["scanFileSystem",
                         "Scans directory roots for keywords. Returns ONE ROW PER MATCH with the "
                         "matched line, configurable context lines, and extracted tokens."],
                        ["runSQLServerQuery",
                         "Runs a SQL Server query against a LOB catalog using Windows Integrated "
                         "Security. No password is ever handled, stored or logged."],
                        ["formatCSV",
                         "Renders a CSV as a styled Excel workbook with a SAS ODS look and feel: "
                         "navy banner, blue header, zebra striping."],
                        ["downloadBulkFiles",
                         "Downloads every file listed in a CSV link column. Blank cells are skipped; "
                         "a failed download is logged and the run continues."],
                        ["convertSAS2Pandas",
                         "Reads a sas7bdat data set into a pandas DataFrame and saves it as parquet, "
                         "CSV or pickle."],
                        ["copyExcelSheet2CSV",
                         "Exports one worksheet to CSV, validating first and refusing to write when "
                         "the sheet is not shaped for flat output."]],
                       headerHeight=280000, rowHeight=490000,
                       bodySize=1000, monoColumns=(0,))
    slides.append(slideShell(body))

    # ---------- 7. Function catalog 2 ----------
    body, nid = chrome("The function catalog  ·  2 of 2",
                       "Operations, Packaging and Orchestration",
                       "The operational half: alerting, host metrics, archiving and the end-to-end pipeline.",
                       page(7))
    body += tableFrame(nid, MARGIN_X, 1420000, [2011680, 6126480],
                       ["Function", "Description"],
                       [["sendEmail",
                         "Sends an SMTP alert to one or many recipients. Used to notify operations "
                         "when a job finishes or fails."],
                        ["collectSystemMetrics",
                         "Gathers host metrics into a CSV time series. Fails gracefully: an "
                         "unavailable metric is blank and the run still succeeds."],
                        ["zipFolder",
                         "Archives a folder plus accompanying files into one zip, for records "
                         "retention or hand-off."],
                        ["get_comments",
                         "Retrieves public comments from Regulations.gov for a docket such as "
                         "CMS-2022-0193, with an API key held as a secret."],
                        ["runFilescanPipeline",
                         "End-to-end orchestration: scanFileSystem, then formatCSV, then sendEmail. "
                         "Coordinates; never reimplements."],
                        ["basic_hello / detailed_hello",
                         "Smoke tests that confirm the package imported correctly in a new "
                         "environment."]],
                       headerHeight=280000, rowHeight=490000,
                       bodySize=1000, monoColumns=(0,))
    slides.append(slideShell(body))

    # ---------- 8. Four environments ----------
    body, nid = chrome("Where the functions run", "Four Environments",
                       "The same call, promoted from a laptop to the cloud, with no rewrite at any step.",
                       page(8))
    body += cardRow(nid, 1430000, [
        {"number": "1", "title": "Desktop",
         "body": "Develop and explore. VS Code, a local venv, small folders. "
                 "Fast feedback, full debugging, zero risk to shared systems."},
        {"number": "2", "title": "UNIT Server",
         "body": "Test at real scale against real UNC shares and the LOB data "
                 "marts. Validate before anything touches production."},
        {"number": "3", "title": "PROD Server",
         "body": "Run for the business. Scheduled, unattended, monitored, with "
                 "results delivered to the operations mailbox."},
        {"number": "4", "title": "Snowflake Cloud",
         "body": "Run next to the data. Elastic compute, governed access, no "
                 "file shares to mount and no server to patch."},
    ], height=1850000)
    body += textBox(96, "Note", MARGIN_X, 3520000, CONTENT_W, 560000,
                    paragraph(run("One codebase, four homes.  ", 1150, NAVY_DEEP, bold=True)
                              + run("Configuration lives in .env, so promoting a job between "
                                    "environments changes paths and credentials — never the code.",
                                    1150, SLATE)))
    slides.append(slideShell(body))

    # ---------- 9. Environment comparison ----------
    body, nid = chrome("Environment comparison", "Choosing the Right Environment",
                       "Each tier has a purpose. Work moves left to right as it matures.",
                       page(9))
    body += tableFrame(nid, MARGIN_X, 1420000,
                       [1500000, 1659540, 1659540, 1659540, 1659540],
                       ["Dimension", "Desktop", "UNIT Server", "PROD Server", "Snowflake"],
                       [["Purpose", "Develop, explore", "Test at scale",
                         "Run the business", "Run next to data"],
                        ["Data volume", "Small samples", "Full test sets",
                         "Full production", "Warehouse scale"],
                        ["Runs when", "You press F5", "On demand",
                         "Scheduled, unattended", "Scheduled task"],
                        ["Availability", "Your machine only", "Shared, business hours",
                         "Managed, monitored", "Cloud SLA"],
                        ["Credentials", "Your own login", "Service account",
                         "Service account", "Snowflake secret"],
                        ["Interface", "VS Code", "SAS EG, CLI",
                         "Scheduler, SAS", "Worksheets, Snowpark"]],
                       headerHeight=290000, rowHeight=270000, bodySize=950,
                       headerSize=1000)
    slides.append(slideShell(body))

    # ---------- 10. Moving up the chain ----------
    body, nid = chrome("Why promote work", "The Benefits of Moving Up the Chain",
                       "Every step toward the server and the cloud buys something concrete.",
                       page(10))
    body += cardRow(nid, 1430000, [
        {"number": "1", "title": "Better performance",
         "body": "Server-class CPU, memory and network. A scan that crawls a UNC "
                 "share over a laptop VPN runs beside the data instead. Snowflake "
                 "adds elastic compute that scales with the workload."},
        {"number": "2", "title": "More reliability",
         "body": "No dependency on one person's machine being awake, patched or "
                 "connected. Managed hosts, restart policies and monitored "
                 "storage replace 'it worked on my laptop'."},
        {"number": "3", "title": "More security",
         "body": "Service accounts and Integrated Security instead of personal "
                 "logins. Secrets in .env or a Snowflake secret, never in code. "
                 "Access is auditable and revocable."},
        {"number": "4", "title": "Maintenance and automation",
         "body": "One scheduled job replaces a manual routine. Central logs, "
                 "email alerts on completion or failure, and one deployed copy "
                 "to patch rather than many desktop copies."},
    ], height=2100000)
    slides.append(slideShell(body))

    # ---------- 11. Use case: the pipeline ----------
    # NOTE: keep this title short enough for one line at 28pt -- a wrapped
    # title runs into the subtitle, which sits at a fixed y offset.
    body, nid = chrome("Use case", "Scan, Format, Alert: One Pipeline",
                       "The nightly job: sweep the SAS log shares, produce a formatted workbook, notify operations.",
                       page(11))
    body += cardRow(nid, 1430000, [
        {"number": "1", "title": "scanFileSystem",
         "body": "Sweeps the HHH and DME log shares for 'real time' and 'cpu time'. "
                 "Returns one row per match with context lines. With "
                 "metric_profile=sas_log it also writes a Metrics sheet of "
                 "per-step timings."},
        {"number": "2", "title": "formatCSV",
         "body": "Turns the raw result into a corporate-styled Excel workbook: "
                 "navy banner, blue headers, zebra striping, frozen and filtered "
                 "header row. Ready to send to a manager as-is."},
        {"number": "3", "title": "sendEmail",
         "body": "Notifies the operations mailbox with row counts and output "
                 "paths. If mail fails the pipeline still succeeds, because the "
                 "workbook is already on disk."},
    ], height=1950000)
    body += textBox(97, "Note", MARGIN_X, 3620000, CONTENT_W, 500000,
                    paragraph(run("Orchestration only.  ", 1150, NAVY_DEEP, bold=True)
                              + run("The pipeline calls the three functions and does no work "
                                    "itself — so fixing the scanner fixes it everywhere.",
                                    1150, SLATE)))
    slides.append(slideShell(body))

    # ---------- 12. One pipeline, three languages ----------
    body, nid = chrome("The same call, three ways", "One Pipeline, Three Languages",
                       "Identical function names and identical parameter names. Only the surrounding syntax changes.",
                       page(12))
    width = (CONTENT_W - 320040) // 3
    body += codeBlock(nid, MARGIN_X, 1450000, width, 2000000,
                      "SAS  ·  Enterprise Guide",
                      ["%runFilescanPipeline(", "  extract_keyword=",
                       "    %str(real time),", "  metric_profile=sas_log,",
                       "  email_to=ops@cgs.com", ");"])
    body += codeBlock(nid + 2, MARGIN_X + width + 160020, 1450000, width, 2000000,
                      "Python  ·  VS Code",
                      ["cgs_ai.runFilescanPipeline(", "  extract_keyword=",
                       "    ['real time'],", "  metric_profile='sas_log',",
                       "  email_to='ops@cgs.com')"])
    body += codeBlock(nid + 4, MARGIN_X + 2 * (width + 160020), 1450000, width, 2000000,
                      "PowerShell  ·  Scheduler",
                      [".\\filescan_pipeline.ps1 `", "  -extract_keyword",
                       "    'real time' `", "  -metric_profile sas_log `",
                       "  -email_to ops@cgs.com"])
    body += textBox(99, "Note", MARGIN_X, 3620000, CONTENT_W, 560000,
                    paragraph(run("Why this matters:  ", 1150, NAVY_DEEP, bold=True)
                              + run("an analyst can hand a working call to an engineer, or a "
                                    "scheduler, without translation. The SAS macro even takes "
                                    "engine=ps or engine=py to switch implementation without "
                                    "changing any other argument.", 1150, SLATE)))
    slides.append(slideShell(body))

    # ---------- 13. Scheduling ----------
    body, nid = chrome("Automation", "Scheduling the Pipeline",
                       "The step that turns a useful script into an operational service.",
                       page(13))
    body += tableFrame(nid, MARGIN_X, 1420000, [1700000, 2100000, 4338160],
                       ["Environment", "Scheduled by", "How it runs"],
                       [["Desktop", "Manual / F5",
                         "Development only. Not a place to schedule business work."],
                        ["UNIT Server", "Windows Task Scheduler",
                         "A .bat calling the PowerShell or Python entry point, for "
                         "test runs against real data."],
                        ["PROD Server", "Task Scheduler / SAS",
                         "Unattended, service account, exit codes 0/2/3 so the scheduler "
                         "detects failure. Output and alerts are the deliverable."],
                        ["Snowflake", "Snowflake TASK",
                         "A scheduled TASK calls the stored procedure. No server to "
                         "patch and no share to mount."]],
                       headerHeight=290000, rowHeight=460000, bodySize=1000)
    body += textBox(98, "Note", MARGIN_X, 3700000, CONTENT_W, 620000,
                    paragraph(run("Built to be scheduled.  ", 1150, NAVY_DEEP, bold=True)
                              + run("No cgs_ai function ever prompts — a prompt would hang an "
                                    "unattended run forever. Every function validates its "
                                    "parameters first, logs to stderr, and exits with a code a "
                                    "scheduler can branch on.", 1150, SLATE)))
    slides.append(slideShell(body))

    # ---------- 14. Conclusions ----------
    body, nid = chrome("Conclusions", "The Foundation for AI Pipelines",
                       "Where this leaves us, and what it makes possible next.",
                       page(14))
    body += bulletBlock(nid, MARGIN_X, 1460000, CONTENT_W, 2300000, [
        ("cgs_ai sets the foundation for AI pipelines.",
         "Reliable, repeatable data movement is the prerequisite for any AI work. "
         "Scanning, extraction, formatting and alerting are that foundation."),
        ("A single AI-ready claims-analysis toolkit.",
         "An analyst calls it from a SAS session, a scheduler calls it from a batch "
         "file, and an engineer imports it into Python — with identical function "
         "names, identical parameter names and identical output."),
        ("The path from GUI to code is now incremental.",
         "Nobody has to leap. Start in Enterprise Guide, move to VS Code, promote to "
         "the server, and land in Snowflake — the same call at every step."),
    ], size=1200, gap=1000)
    body += rect(200, "CalloutBg", MARGIN_X, 3760000, CONTENT_W, 700000, "F4F7FC")
    body += textBox(201, "Callout", MARGIN_X + 182880, 3900000,
                    CONTENT_W - 365760, 460000,
                    paragraph(run("Next step:  ", 1250, NAVY, bold=True)
                              + run("pick one manual routine you run every week and schedule it. "
                                    "That single move buys performance, reliability, security and "
                                    "your time back.", 1250, NAVY_DEEP)))
    slides.append(slideShell(body))

    return slides


# =====================================================================
# Package assembly
# =====================================================================

def buildPresentation(templatePath: str, outputPath: str) -> str:
    """Write the deck, reusing the template's masters, layouts and theme.

    Parameters:
        templatePath (str) - source .pptx template.
        outputPath (str)   - destination .pptx.
    Returns: the path written.
    Raises: OSError if the template cannot be read.
    """
    slides = buildSlides()
    source = zipfile.ZipFile(templatePath)
    names = source.namelist()

    # Keep everything except the template's own slides, their rels and notes.
    drop = re.compile(r"^ppt/(slides|notesSlides)/")
    keep = [n for n in names if not drop.match(n)]

    layoutFor = "../slideLayouts/slideLayout1.xml"
    slideRels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
                 f'officeDocument/2006/relationships/slideLayout" Target="{layoutFor}"/>'
                 '</Relationships>')

    # presentation.xml: rebuild the slide id list.
    presentation = source.read("ppt/presentation.xml").decode("utf-8")
    presRels = source.read("ppt/_rels/presentation.xml.rels").decode("utf-8")

    # Strip existing slide relationships, then add one per new slide.
    presRels = re.sub(r'<Relationship[^>]*slides/slide\d+\.xml"/>', "", presRels)
    existing = [int(m) for m in re.findall(r'Id="rId(\d+)"', presRels)] or [0]
    nextRel = max(existing) + 1
    slideRelIds = []
    additions = ""
    for index in range(len(slides)):
        relId = f"rId{nextRel + index}"
        slideRelIds.append(relId)
        additions += (f'<Relationship Id="{relId}" Type="http://schemas.openxmlformats.org/'
                      f'officeDocument/2006/relationships/slide" '
                      f'Target="slides/slide{index + 1}.xml"/>')
    presRels = presRels.replace("</Relationships>", additions + "</Relationships>")

    sldIdLst = "".join(f'<p:sldId id="{256 + i}" r:id="{rid}"/>'
                       for i, rid in enumerate(slideRelIds))
    presentation = re.sub(r"<p:sldIdLst>.*?</p:sldIdLst>",
                          f"<p:sldIdLst>{sldIdLst}</p:sldIdLst>",
                          presentation, flags=re.DOTALL)

    # [Content_Types].xml: one override per slide, notes overrides removed.
    contentTypes = source.read("[Content_Types].xml").decode("utf-8")
    # NOTE: the ContentType value itself contains '/', so the trailing attribute
    # must be matched with [^>]* -- [^/]* silently matches nothing and leaves
    # the template's original overrides behind, producing duplicate PartNames.
    contentTypes = re.sub(r'<Override PartName="/ppt/(?:slides|notesSlides)/[^"]*"[^>]*/>',
                          "", contentTypes)
    slideOverrides = "".join(
        f'<Override PartName="/ppt/slides/slide{i + 1}.xml" '
        f'ContentType="application/vnd.openxmlformats-officedocument.'
        f'presentationml.slide+xml"/>' for i in range(len(slides)))
    contentTypes = contentTypes.replace("</Types>", slideOverrides + "</Types>")

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
    template = sys.argv[1] if len(sys.argv) > 1 else "TEMPLATE1_0.pptx"
    output = sys.argv[2] if len(sys.argv) > 2 else str(
        Path(__file__).resolve().parent.parent.parent / "CGS_AI_Presentation.pptx")
    print("wrote", buildPresentation(template, output))
