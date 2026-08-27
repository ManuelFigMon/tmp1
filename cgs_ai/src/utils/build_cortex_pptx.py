"""
=====================================================================
  Program Name  : build_cortex_pptx.py
  Author        : Manuel Figallo
  Purpose       : Build CGS_AI_CORTEX_PIPELINE_v1.pptx -- the CMS
                  regulations Comments analysis pipeline running on
                  Snowflake Cortex.
  Version       : 1.0beta
  Created       : 2026-08-27
  Last Modified : 2026-08-27

  Dependencies:
    STANDARD LIBRARY ONLY. Design helpers come from build_presentation.py
    and the demo deck supplies the masters, layouts and theme, so this deck
    is visually identical to CGS_AI_DEMO_v2.pptx.

  Description:
    Content is taken from cgs_ai_reference.docx (function signatures and
    descriptions) and cms_workflow.pdf (the processing workflow, regrouped
    into an Analysis Pipeline and an Output Pipeline).

    House rule applied throughout: the phrase "Public Comment" is never
    used -- these are "Comments".

  Usage:
    python src/utils/build_cortex_pptx.py <source.pptx> <output.pptx> [imageDir]

    imageDir is searched for the three workflow screenshots:
        regulations_docket.png
        snowflake_cortex_CMS_Regulations_analytics.png
        Executive_Briefing_OUTPUT1.png
    Any that are missing render as a labelled drop-in frame, so the deck
    builds either way. Re-run with the directory once the files exist.
=====================================================================
"""

from __future__ import annotations

import re
import struct
import sys
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

sys.path.insert(0, str(Path(__file__).resolve().parent))

from build_presentation import (CONTENT_W, LIGHT, MARGIN_X, NAVY,  # noqa: E402
                                NAVY_CIRCLE, NAVY_DEEP, RULE, SLATE,
                                SLIDE_H, SLIDE_W, WHITE, bulletBlock,
                                cardRow, codeBlock, ellipse, esc, paragraph,
                                rect, run, slideShell, tableFrame, textBox)

__version__ = "1.0beta"

TOTAL_SLIDES = 12
FOOTER_LEFT = "cgs_ai · Cortex Pipeline"
FOOTER_CENTER = "CMS-2022-0190 Comments Analysis"

BODY_TOP = 1430000

# Workflow block colours. The user specified green / blue / light blue for
# the three stages, so these sit outside the deck's navy palette on purpose.
GREEN = "2E7D4F"
BLUE = "2E75B6"
LIGHTBLUE = "9DC3E6"
PANEL = "F4F7FC"

# The three high-level workflow screenshots, in order.
WORKFLOW_IMAGES = [
    ("regulations_docket.png", "Regulations.gov"),
    ("snowflake_cortex_CMS_Regulations_analytics.png", "Cortex NLP"),
    ("Executive_Briefing_OUTPUT1.png", "Executive Briefing"),
]


def page(index: int) -> str:
    """Return the page indicator for slide `index`. Parameters: index (int)."""
    return f"{index} / {TOTAL_SLIDES}"


def chrome(eyebrow: str, title: str, subtitle: str, pageNo: str) -> Tuple[str, int]:
    """Build the standard content-slide chrome, with this deck's footers.

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
    return textBox(shapeId, "Note", MARGIN_X, y, CONTENT_W, 400000,
                   paragraph(runs))


def picture(shapeId: int, relId: str, x: int, y: int, cx: int, cy: int) -> str:
    """Build a <p:pic> shape with a hairline border. Geometry in EMU."""
    return (f'<p:pic><p:nvPicPr><p:cNvPr id="{shapeId}" name="Shot {shapeId}"/>'
            f'<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/>'
            f'</p:nvPicPr><p:blipFill><a:blip r:embed="{relId}"/>'
            f'<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>'
            f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            f'<a:ln w="9525"><a:solidFill><a:srgbClr val="{RULE}"/></a:solidFill>'
            f'</a:ln></p:spPr></p:pic>')


def arrow(shapeId: int, x: int, y: int, cx: int, cy: int,
          fill: str = NAVY) -> str:
    """Build a right-pointing arrow. Parameters: geometry (EMU), fill colour."""
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{shapeId}" name="Arrow {shapeId}"/>'
            f'<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>'
            f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            f'<a:prstGeom prst="rightArrow"><a:avLst>'
            f'<a:gd name="adj1" fmla="val 50000"/>'
            f'<a:gd name="adj2" fmla="val 50000"/></a:avLst></a:prstGeom>'
            f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
            f'<a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr/>'
            f'<a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p>'
            f'</p:txBody></p:sp>')


def roundPanel(shapeId: int, x: int, y: int, cx: int, cy: int,
               fill: str) -> str:
    """Build a rounded container panel. Parameters: geometry, fill colour."""
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{shapeId}" name="Panel {shapeId}"/>'
            f'<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>'
            f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            f'<a:prstGeom prst="roundRect"><a:avLst>'
            f'<a:gd name="adj" fmla="val 6000"/></a:avLst></a:prstGeom>'
            f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
            f'<a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr/>'
            f'<a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p>'
            f'</p:txBody></p:sp>')


def stageBlock(shapeId: int, x: int, y: int, cx: int, cy: int, fill: str,
               number: str, title: str, lines: Sequence[str],
               textColor: str = WHITE, mono: bool = False) -> str:
    """Build one workflow stage block: numbered circle, title, detail lines.

    Parameters:
        fill (str)       - block colour.
        number (str)     - step number shown in the circle.
        title (str)      - block heading.
        lines (sequence) - detail lines under the heading.
        textColor (str)  - colour for heading and detail text.
        mono (bool)      - render the detail lines in Consolas.
    Returns: the shape XML for the whole block.
    """
    circleFill = WHITE if textColor == WHITE else NAVY
    circleText = fill if textColor == WHITE else WHITE
    shapes = [
        roundPanel(shapeId, x, y, cx, cy, fill),
        ellipse(shapeId + 1, f"Num{shapeId}", x + 137160, y + 137160,
                300000, circleFill),
        textBox(shapeId + 2, f"NumT{shapeId}", x + 137160, y + 195000,
                300000, 220000,
                paragraph(run(number, 1050, circleText, bold=True), align="ctr")),
        textBox(shapeId + 3, f"Title{shapeId}", x + 137160, y + 530000,
                cx - 274320, 300000,
                paragraph(run(title, 1300, textColor, bold=True))),
    ]
    paras = "".join(
        paragraph(run(line, 950, textColor, mono=mono),
                  spaceBefore=0 if i == 0 else 350)
        for i, line in enumerate(lines))
    shapes.append(textBox(shapeId + 4, f"Body{shapeId}", x + 137160,
                          y + 880000, cx - 274320, cy - 960000, paras))
    return "".join(shapes)


def imageFrame(shapeId: int, relId: Optional[str], aspect: Optional[float],
               fileName: str, x: int, y: int, cx: int, cy: int) -> str:
    """Place a screenshot in a fixed frame, or draw a labelled drop-in frame.

    Parameters:
        relId (str|None)    - slide relationship id, or None when the file
                              was not supplied.
        aspect (float|None) - image width/height, when known.
        fileName (str)      - the expected file name, shown on the placeholder.
        x, y, cx, cy (int)  - the frame, in EMU.
    Returns: the shape XML. The image is centred and scaled to fit the frame.
    """
    if relId and aspect:
        width, height = cx, int(cx / aspect)
        if height > cy:
            height, width = cy, int(cy * aspect)
        return picture(shapeId, relId, x + (cx - width) // 2,
                       y + (cy - height) // 2, width, height)
    return (rect(shapeId, f"Frame{shapeId}", x, y, cx, cy, PANEL)
            + textBox(shapeId + 1, f"FrameT{shapeId}", x + 100000,
                      y + cy // 2 - 250000, cx - 200000, 500000,
                      paragraph(run("drop in", 900, SLATE), align="ctr")
                      + paragraph(run(fileName, 800, NAVY_DEEP, mono=True),
                                  align="ctr", spaceBefore=300)))


# =====================================================================
# Slide content
# =====================================================================

def buildSlides(images: Dict[str, Tuple[str, float]]
                ) -> List[Tuple[str, List[str]]]:
    """Assemble the deck.

    Parameters:
        images (dict) - file name -> (media part name, aspect) for each
                        workflow screenshot that was supplied.
    Returns: a list of (slide XML, list of media part names used).
    """
    slides: List[Tuple[str, List[str]]] = []

    # ---------- 1. Title ----------
    body = "".join([
        rect(2, "Bg", 0, 0, SLIDE_W, SLIDE_H, NAVY),
        ellipse(3, "BigCircle", 6858000, -457200, 2743200, NAVY_CIRCLE),
        ellipse(4, "SmallCircle", 7772400, 457200, 822960, LIGHT),
        ellipse(5, "CornerCircle", -457200, 3429000, 1828800, NAVY_CIRCLE),
        textBox(6, "Title", 640080, 1005840, 6096000, 1400000,
                paragraph(run("CGS_AI Cortex Pipeline", 3400, WHITE, bold=True))),
        textBox(7, "Subtitle", 640080, 2377440, 6858000, 400000,
                paragraph(run("Analyzing CMS regulations Comments at scale.",
                              1600, LIGHT))),
        textBox(8, "Footer", 640080, 3931920, 6858000, 600000,
                paragraph(run("Docket CMS-2022-0190  ·  built on Snowflake Cortex",
                              1100, LIGHT))
                + paragraph(run("Prepared by Manuel Figallo  ·  cgs_ai v1.0beta",
                                1100, LIGHT), spaceBefore=400)),
    ])
    slides.append((slideShell(body), []))

    # ---------- 2. The use case ----------
    body, nid = chrome("A new use case", "898 Comments, One Docket",
                       "Docket CMS-2022-0190 closed with 898 Comments on prior "
                       "authorization and interoperability.", page(2))
    body += cardRow(nid, BODY_TOP, [
        {"number": "1", "title": "The volume is the problem",
         "body": "898 Comments, 666 of them substantive. Reading and coding "
                 "them by hand is weeks of analyst time."},
        {"number": "2", "title": "The questions are always the same",
         "body": "What is the sentiment? Which policy domains do they touch? "
                 "What would we tell leadership?"},
        {"number": "3", "title": "The answer should be repeatable",
         "body": "The next docket needs the same analysis. That makes it a "
                 "pipeline, not a project."},
    ], height=1550000)
    body += note(nid + 40, "So we built one.",
                 "cgs_ai turns a docket of Comments into a sentiment-scored, "
                 "categorized data set and an executive briefing.", y=3220000)
    slides.append((slideShell(body), []))

    # ---------- 3. High-level workflow (three screenshots) ----------
    body, nid = chrome("High-level workflow",
                       "Regulations.gov  →  Cortex NLP  →  Briefing",
                       "Three stages. The Comments never leave Snowflake once "
                       "they land.", page(3))
    gap = 420000
    colWidth = (CONTENT_W - 2 * gap) // 3
    used: List[str] = []
    relIndex = 2
    shapeId = nid
    for index, (fileName, label) in enumerate(WORKFLOW_IMAGES):
        x = MARGIN_X + index * (colWidth + gap)
        body += ellipse(shapeId, f"Step{index}", x, 1420000, 300000, NAVY)
        body += textBox(shapeId + 1, f"StepT{index}", x, 1478000, 300000, 220000,
                        paragraph(run(str(index + 1), 1050, WHITE, bold=True),
                                  align="ctr"))
        body += textBox(shapeId + 2, f"StepL{index}", x + 380000, 1455000,
                        colWidth - 380000, 300000,
                        paragraph(run(label.upper(), 1150, NAVY_DEEP, bold=True)))
        shapeId += 3
        relId, aspect = None, None
        if fileName in images:
            part, aspect = images[fileName]
            relId = f"rId{relIndex}"
            relIndex += 1
            used.append(part)
        body += imageFrame(shapeId, relId, aspect, fileName,
                           x, 1880000, colWidth, 1700000)
        shapeId += 2
        if index < 2:
            body += arrow(shapeId, x + colWidth + 90000, 2620000,
                          240000, 220000)
            shapeId += 1
    captions = ["The docket and its Comments, downloaded as CSV.",
                "cgs_ai running in a Snowflake notebook, next to the data.",
                "A formatted memo with sentiment, categories and a narrative."]
    for index, caption in enumerate(captions):
        x = MARGIN_X + index * (colWidth + gap)
        body += textBox(shapeId, f"Cap{index}", x, 3660000, colWidth, 500000,
                        paragraph(run(caption, 1000, SLATE)))
        shapeId += 1
    body += note(shapeId, "One direction, no round trips.",
                 "Source to briefing without moving the data to a laptop.",
                 y=4270000)
    slides.append((slideShell(body), used))

    # ---------- 4. Why Cortex ----------
    body, nid = chrome("Where it runs", "NLP That Runs Next to the Data",
                       "cgs_ai imports into a Snowflake notebook, so the "
                       "analysis happens inside Cortex.", page(4))
    body += codeBlock(nid, MARGIN_X, BODY_TOP, CONTENT_W, 900000,
                      "Snowflake notebook cell",
                      ["%run cgs_ai_setup",
                       "import cgs_ai",
                       "print('cgs_ai imported from:', cgs_ai.__file__)"])
    body += bulletBlock(nid + 2, MARGIN_X, 2500000, CONTENT_W, 1650000, [
        ("Elastic compute.",
         "summarizeText calls Cortex Summarize() across 10 parallel workers; "
         "the warehouse scales, your laptop does not have to."),
        ("No data movement.",
         "The Comments stay in Snowflake. Nothing is downloaded to a share "
         "or a desktop to be analyzed."),
        ("No external packages to install.",
         "Every function uses pre-installed or standard-library dependencies "
         "— no PyPI access required."),
        ("Platform-guarded.",
         "The analytical functions check they are on Linux / Snowflake before "
         "they run, so they fail clearly rather than oddly."),
    ], size=1150, gap=450)
    slides.append((slideShell(body), []))

    # ---------- 5. cgs_ai is a collection of functions ----------
    body, nid = chrome("What cgs_ai is", "A Collection of Functions",
                       "Import the one you need, pass parameters, get a "
                       "predictable result back.", page(5))
    body += codeBlock(nid, MARGIN_X, BODY_TOP, CONTENT_W, 1700000,
                      "from cgs_ai import formatCSV",
                      ["result = formatCSV(",
                       "    InputCsvPath='data/output/cms_2022_0190_SENT_SUMMARIZE_CAT_v1.csv',",
                       "    OutputExcelPath='data/output/report_corporate.xlsx',",
                       "    FormatType='corporate',",
                       "    SheetName='CMS Comments',",
                       "    Title='CMS-2022-0190 Comments Analysis')",
                       "print(result)"])
    body += bulletBlock(nid + 2, MARGIN_X, 3300000, CONTENT_W, 1300000, [
        ("What it does.",
         "Reads the enriched Comments CSV and writes a styled .xlsx — "
         "FormatType='corporate' gives a navy banner, blue header, zebra "
         "stripes and a frozen, auto-filtered header row."),
        ("What it returns.",
         "A dict — OutputExcelPath, RowCount, ColumnCount, FormatType — so "
         "the next step can check the result instead of guessing."),
    ], size=1100, gap=450)
    slides.append((slideShell(body), []))

    # ---------- 6. The processing workflow ----------
    body, nid = chrome("How it runs",
                       "The AI-Powered Processing Workflow",
                       "Two pipelines: analysis enriches the Comments, output "
                       "turns them into deliverables.", page(6))

    analysisX, analysisW = MARGIN_X, 5180000
    outputX = analysisX + analysisW + 340000
    outputW = CONTENT_W - analysisW - 340000
    panelY, panelH = 1560000, 2280000

    body += textBox(nid, "AnalysisLabel", analysisX, BODY_TOP - 130000,
                    analysisW, 250000,
                    paragraph(run("ANALYSIS PIPELINE", 1000, SLATE, bold=True)))
    body += textBox(nid + 1, "OutputLabel", outputX, BODY_TOP - 130000,
                    outputW, 250000,
                    paragraph(run("OUTPUT PIPELINE", 1000, SLATE, bold=True)))
    body += roundPanel(nid + 2, analysisX, panelY, analysisW, panelH, PANEL)
    body += roundPanel(nid + 3, outputX, panelY, outputW, panelH, PANEL)

    innerY, innerH = panelY + 150000, panelH - 300000
    loadW = 1750000
    nlpX = analysisX + 150000 + loadW + 320000
    nlpW = analysisW - 300000 - loadW - 320000

    body += stageBlock(nid + 10, analysisX + 150000, innerY, loadW, innerH,
                       GREEN, "1", "Load CSV",
                       ["Read the docket", "Comments into a", "DataFrame."])
    body += arrow(nid + 20, analysisX + 150000 + loadW + 60000,
                  innerY + innerH // 2 - 110000, 200000, 220000)
    body += stageBlock(nid + 21, nlpX, innerY, nlpW, innerH, BLUE,
                       "2", "NLP Processing",
                       ["doSentimentAnalysis", "visualizeNLP",
                        "summarizeText", "categorizeText"], mono=True)
    body += arrow(nid + 30, outputX - 260000, innerY + innerH // 2 - 110000,
                  200000, 220000)
    body += stageBlock(nid + 31, outputX + 150000, innerY,
                       outputW - 300000, innerH, LIGHTBLUE,
                       "3", "Generate Output",
                       ["formatCSV", "createExecutiveBriefing"],
                       textColor=NAVY_DEEP, mono=True)

    body += note(nid + 40, "Analysis enriches, output delivers.",
                 "Everything the NLP stage adds — sentiment, summary, "
                 "category — is what the output stage formats.", y=3980000)
    slides.append((slideShell(body), []))

    # ---------- 7. Function catalog: NLP ----------
    body, nid = chrome("The function catalog  ·  1 of 2",
                       "NLP and Analysis Functions",
                       "Every analytical function is platform-guarded and "
                       "needs no external package download.", page(7))
    body += tableFrame(nid, MARGIN_X, BODY_TOP,
                       [1800000, 2500000, 3838160],
                       ["Function", "Signature", "What it does"],
                       [["doSentimentAnalysis",
                         "(csv_path, comment_col='COMMENT',\n result_col=None, "
                         "CREATE_BAR=False)",
                         "Classifies each Comment as Positive, Negative or "
                         "Neutral / Other by keyword intersection. Adds "
                         "char_count and word_count; optional bar chart."],
                        ["visualizeNLP",
                         "(csv_path, comment_col='COMMENT',\n viz_type='WORD_DIST')",
                         "Renders a 1x2 matplotlib figure. WORD_DIST is a word-"
                         "count histogram, WORD_CLOUD a top-30 frequency chart, "
                         "BOTH shows them side by side."],
                        ["summarizeText",
                         "(csv_path, comment_col='COMMENT',\n result_col=None)",
                         "Calls Snowflake Cortex Summarize() on every row using "
                         "10 parallel workers. No temp stage required."],
                        ["categorizeText",
                         "(csv_path, comment_col='COMMENT',\n result_col=None)",
                         "Assigns each Comment to a CMS policy domain by keyword "
                         "scoring — reimbursement, care quality, drug pricing, "
                         "provider burden, health equity."]],
                       headerHeight=300000, rowHeight=640000,
                       bodySize=900, monoColumns=(0, 1))
    slides.append((slideShell(body), []))

    # ---------- 8. Function catalog: output ----------
    body, nid = chrome("The function catalog  ·  2 of 2",
                       "Output and Package Functions",
                       "The deliverables, plus the three greetings that prove "
                       "the import worked.", page(8))
    body += tableFrame(nid, MARGIN_X, BODY_TOP,
                       [1800000, 2500000, 3838160],
                       ["Function", "Signature", "What it does"],
                       [["formatCSV",
                         "(InputCsvPath, OutputExcelPath,\n "
                         "FormatType='corporate',\n SheetName='Report', Title='')",
                         "Converts a CSV to a styled .xlsx — corporate, plain or "
                         "minimal. Saves via BytesIO to bypass the Snowflake "
                         "workspace seek limitation. Returns a result dict."],
                        ["createExecutiveBriefing",
                         "(input_csv, text_col='comment',\n secondary_cols=None,\n "
                         "output_file='...md')",
                         "Aggregates sentiment and category stats, calls Cortex "
                         "Complete for a three-paragraph narrative, and writes a "
                         "memo as .md, .pdf or .docx."],
                        ["basic_hello",
                         "()",
                         "Returns \"Hello, World!\" — the baseline package "
                         "smoke test."],
                        ["personalized_hello",
                         "(name)",
                         "Returns a personalized greeting; strips whitespace and "
                         "falls back to \"World\" on empty input."],
                        ["detailed_hello",
                         "(style='friendly')",
                         "Returns a dict with message and style_used. Styles: "
                         "friendly, formal, pirate."]],
                       headerHeight=300000, rowHeight=530000,
                       bodySize=900, monoColumns=(0, 1))
    slides.append((slideShell(body), []))

    # ---------- 9. End to end ----------
    body, nid = chrome("End to end", "The Whole Pipeline in One Cell",
                       "Six calls take the docket from raw CSV to a briefing "
                       "and a workbook.", page(9))
    body += codeBlock(nid, MARGIN_X, BODY_TOP, CONTENT_W, 2750000,
                      "Snowflake notebook  ·  CMS-2022-0190",
                      ["from cgs_ai import (doSentimentAnalysis, visualizeNLP, "
                       "summarizeText,",
                       "                    categorizeText, formatCSV, "
                       "createExecutiveBriefing)",
                       "",
                       "CSV = 'data/CMS_2022_0190_COMMENTS_v3.csv'",
                       "df = doSentimentAnalysis(CSV, comment_col='comment')   "
                       "# 1  sentiment",
                       "visualizeNLP(CSV, viz_type='BOTH')                     "
                       "# 2  charts",
                       "df = summarizeText(CSV, comment_col='comment')         "
                       "# 3  Cortex",
                       "df = categorizeText(CSV, comment_col='comment')        "
                       "# 4  domains",
                       "formatCSV('data/output/enriched.csv',                  "
                       "# 5  workbook",
                       "          'data/output/enriched_corporate.xlsx', "
                       "FormatType='corporate')",
                       "createExecutiveBriefing(                               "
                       "# 6  briefing",
                       "    input_csv='data/output/enriched.csv',",
                       "    secondary_cols=['overall_sentiment', "
                       "'COMMENT_CATEGORY'],",
                       "    output_file='data/output/briefing.md')"])
    body += note(nid + 2, "Same names everywhere.",
                 "These are the function names in the catalog, called with "
                 "the parameter names in the catalog.", y=4300000)
    slides.append((slideShell(body), []))

    # ---------- 10. What comes out ----------
    body, nid = chrome("The result", "What Comes Out the Other End",
                       "One run of the docket, from the executive briefing it "
                       "produced.", page(10))
    body += cardRow(nid, BODY_TOP, [
        {"number": "1", "title": "898 Comments scored",
         "body": "666 substantive entries of more than five words. Every row "
                 "carries a sentiment, a summary and a CMS policy domain."},
        {"number": "2", "title": "Sentiment distribution",
         "body": "676 Neutral, 186 Positive, 36 Negative — with 49.4% raising "
                 "provider administrative burden."},
        {"number": "3", "title": "Two deliverables",
         "body": "A corporate-styled workbook for the analysts, and a memo "
                 "for the steering committee."},
    ], height=1600000)
    body += note(nid + 40, "The narrative is generated, not written.",
                 "createExecutiveBriefing calls Cortex Complete on the "
                 "aggregated statistics and formats the memo around it.",
                 y=3300000)
    slides.append((slideShell(body), []))

    # ---------- 11. Why it matters ----------
    body, nid = chrome("Why it matters", "A Pipeline, Not a One-Off Analysis",
                       "The docket changes. The code does not.", page(11))
    body += bulletBlock(nid, MARGIN_X, BODY_TOP, CONTENT_W, 2600000, [
        ("Point it at the next docket.",
         "Change one CSV path and the same six calls produce the same "
         "deliverables for a different rule."),
        ("The AI work is a function call.",
         "Cortex Summarize and Cortex Complete sit behind summarizeText and "
         "createExecutiveBriefing — no prompt engineering at the call site."),
        ("It is auditable.",
         "Sentiment and categories come from explicit keyword rules you can "
         "read, not an opaque score."),
        ("It is schedulable.",
         "Nothing prompts for input, so a Snowflake TASK can run the whole "
         "pipeline when a comment period closes."),
    ], size=1250, gap=650)
    body += note(nid + 1, "This is what cgs_ai is for:",
                 "making the analysis you did once into the analysis you can "
                 "run every time.", y=3900000)
    slides.append((slideShell(body), []))

    # ---------- 12. Beta ----------
    body, nid = chrome("Where it stands", "cgs_ai Is in Beta",
                       "The Cortex functions are the newest part of the "
                       "package — and the most worth arguing about.", page(12))
    body += cardRow(nid, BODY_TOP, [
        {"number": "1", "title": "Try it on a docket you know",
         "body": "If you already have a view on a rule, run it and tell me "
                 "where the categories are wrong."},
        {"number": "2", "title": "The keyword rules are yours to shape",
         "body": "Sentiment and CMS domain scoring are explicit lists. They "
                 "should reflect how our analysts actually read Comments."},
        {"number": "3", "title": "Ask for a walkthrough",
         "body": "Happy to set this up in your Snowflake workspace against a "
                 "docket you care about."},
    ], height=1600000)
    body += note(nid + 40, "Next step:",
                 "pick the next comment period we care about and run it "
                 "through before the analysis is due.", y=3300000)
    slides.append((slideShell(body), []))

    return slides


# =====================================================================
# Package assembly
# =====================================================================

def pngAspect(data: bytes) -> float:
    """Return width/height for PNG bytes, read from the IHDR chunk."""
    width, height = struct.unpack(">II", data[16:24])
    return width / height


def collectImages(imageDir: Optional[str]) -> Dict[str, Tuple[str, float, bytes]]:
    """Find the workflow screenshots.

    Parameters:
        imageDir (str|None) - directory to search; None skips the search.
    Returns: {fileName: (media part name, aspect, bytes)} for those found.
             Missing files are simply absent, and render as drop-in frames.
    """
    found: Dict[str, Tuple[str, float, bytes]] = {}
    if not imageDir:
        return found
    root = Path(imageDir)
    for index, (fileName, _label) in enumerate(WORKFLOW_IMAGES, start=1):
        candidate = root / fileName
        if not candidate.is_file():
            matches = [p for p in root.glob("*")
                       if p.name.lower() == fileName.lower()]
            candidate = matches[0] if matches else None
        if candidate and candidate.is_file():
            data = candidate.read_bytes()
            found[fileName] = (f"workflow{index}.png", pngAspect(data), data)
    return found


def buildCortexDeck(sourcePath: str, outputPath: str,
                    imageDir: Optional[str] = None) -> str:
    """Write the deck, reusing the demo deck's masters, layouts and theme.

    Parameters:
        sourcePath (str)    - CGS_AI_DEMO_v2.pptx, for the design parts.
        outputPath (str)    - destination .pptx.
        imageDir (str|None) - directory holding the workflow screenshots.
    Returns: the path written.
    Raises: OSError if the source cannot be read.
    """
    source = zipfile.ZipFile(sourcePath)
    found = collectImages(imageDir)
    slides = buildSlides({name: (part, aspect)
                          for name, (part, aspect, _data) in found.items()})

    # Keep the design parts; drop the source's slides and its screenshots,
    # which belong to the demo deck rather than this one.
    drop = re.compile(r"^ppt/(slides|notesSlides|media)/")
    keep = [n for n in source.namelist() if not drop.match(n)]

    layoutRel = ('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
                 'officeDocument/2006/relationships/slideLayout" '
                 'Target="../slideLayouts/slideLayout1.xml"/>')

    def relsFor(parts: Sequence[str]) -> str:
        """Build one slide's .rels: the layout, then rId2.. for each image."""
        extra = "".join(
            f'<Relationship Id="rId{index + 2}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            f'relationships/image" Target="../media/{part}"/>'
            for index, part in enumerate(parts))
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
                '2006/relationships">' + layoutRel + extra + '</Relationships>')

    presRels = source.read("ppt/_rels/presentation.xml.rels").decode("utf-8")
    presRels = re.sub(r'<Relationship[^>]*slides/slide\d+\.xml"/>', "", presRels)
    usedIds = [int(m) for m in re.findall(r'Id="rId(\d+)"', presRels)] or [0]
    nextRel = max(usedIds) + 1
    slideRelIds = [f"rId{nextRel + i}" for i in range(len(slides))]
    presRels = presRels.replace("</Relationships>", "".join(
        f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/'
        f'officeDocument/2006/relationships/slide" '
        f'Target="slides/slide{i + 1}.xml"/>'
        for i, rid in enumerate(slideRelIds)) + "</Relationships>")

    presentation = source.read("ppt/presentation.xml").decode("utf-8")
    presentation = re.sub(
        r"<p:sldIdLst>.*?</p:sldIdLst>",
        "<p:sldIdLst>" + "".join(f'<p:sldId id="{256 + i}" r:id="{rid}"/>'
                                 for i, rid in enumerate(slideRelIds))
        + "</p:sldIdLst>", presentation, flags=re.DOTALL)

    # The ContentType value contains '/', so match the trailing attribute
    # with [^>]* -- [^/]* matches nothing and leaves duplicate PartNames.
    contentTypes = source.read("[Content_Types].xml").decode("utf-8")
    contentTypes = re.sub(
        r'<Override PartName="/ppt/(?:slides|notesSlides)/[^"]*"[^>]*/>',
        "", contentTypes)
    contentTypes = contentTypes.replace("</Types>", "".join(
        f'<Override PartName="/ppt/slides/slide{i + 1}.xml" '
        f'ContentType="application/vnd.openxmlformats-officedocument.'
        f'presentationml.slide+xml"/>' for i in range(len(slides))) + "</Types>")
    if found and 'Extension="png"' not in contentTypes:
        contentTypes = contentTypes.replace(
            "<Types ", '<Types ', 1).replace(
            "</Types>", '<Default Extension="png" ContentType="image/png"/>'
                        "</Types>")

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
        for part, _aspect, data in ((p, a, d) for p, a, d in found.values()):
            out.writestr(f"ppt/media/{part}", data)
        for index, (slide, parts) in enumerate(slides, start=1):
            out.writestr(f"ppt/slides/slide{index}.xml", slide)
            out.writestr(f"ppt/slides/_rels/slide{index}.xml.rels", relsFor(parts))
    source.close()
    missing = [name for name, _ in WORKFLOW_IMAGES if name not in found]
    if missing:
        print("placeholder frames drawn for:", ", ".join(missing))
    return str(target)


if __name__ == "__main__":
    src = sys.argv[1] if len(sys.argv) > 1 else "CGS_AI_DEMO_v2.pptx"
    dst = sys.argv[2] if len(sys.argv) > 2 else str(
        Path(__file__).resolve().parent.parent.parent
        / "CGS_AI_CORTEX_PIPELINE_v1.pptx")
    imgDir = sys.argv[3] if len(sys.argv) > 3 else None
    print("wrote", buildCortexDeck(src, dst, imgDir))
