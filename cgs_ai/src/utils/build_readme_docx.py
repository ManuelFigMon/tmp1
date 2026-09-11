"""
=====================================================================
  Program Name  : build_readme_docx.py
  Author        : Manuel Figallo
  Purpose       : Generate README.docx for the cgs_ai project.
  Version       : 1.0beta
  Created       : 2026-08-26
  Last Modified : 2026-08-26

  Dependencies:
    STANDARD LIBRARY ONLY. A .docx is a zip of OOXML parts, so this writes
    them directly rather than adding python-docx as a dependency. That keeps
    the project's "minimize dependencies" rule intact even for tooling.

  Description:
    Content and section order follow the supplied guidance documents
    (Data_Engineering_Project_Structure_Editable.docx and the "Project
    Structure and Folder Structure" email), including their icon vocabulary.

  Usage:
    python src/utils/build_readme_docx.py [output_path]
=====================================================================
"""

from __future__ import annotations

import struct
import sys
import zipfile
from pathlib import Path
from typing import List, Optional, Sequence, Tuple

__version__ = "1.0beta"

NAVY = "1F3864"
BLUE = "2E75B6"
STRIPE = "DCE6F1"


def esc(text: str) -> str:
    """XML-escape a string. Parameters: text (str). Returns: str."""
    return (str(text).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def para(text: str = "", style: str = "Normal", bold: bool = False,
         size: Optional[int] = None, color: Optional[str] = None,
         align: str = "", spaceAfter: int = 120, mono: bool = False,
         keepNext: bool = False) -> str:
    """Build one <w:p> paragraph.

    Parameters:
        text (str)      - paragraph text ('\n' splits into line breaks).
        style (str)     - style id: Normal, Title, Heading1..3, Code.
        bold, size, color, align, spaceAfter, mono - direct formatting.
        keepNext (bool) - keep on the same page as the next paragraph.
    Returns: the paragraph XML.
    """
    runProps = []
    if mono:
        runProps.append('<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>')
    if bold:
        runProps.append("<w:b/>")
    if color:
        runProps.append(f'<w:color w:val="{color}"/>')
    if size:
        runProps.append(f'<w:sz w:val="{size * 2}"/>')
    rPr = f"<w:rPr>{''.join(runProps)}</w:rPr>" if runProps else ""

    jc = f'<w:jc w:val="{align}"/>' if align else ""
    keep = "<w:keepNext/>" if keepNext else ""
    pPr = (f'<w:pPr><w:pStyle w:val="{style}"/>{keep}'
           f'<w:spacing w:after="{spaceAfter}"/>{jc}</w:pPr>')

    lines = str(text).split("\n")
    runs = ""
    for index, line in enumerate(lines):
        if index:
            runs += f"<w:r>{rPr}<w:br/></w:r>"
        runs += f'<w:r>{rPr}<w:t xml:space="preserve">{esc(line)}</w:t></w:r>'
    return f"<w:p>{pPr}{runs}</w:p>"


def bullet(text: str, level: int = 0) -> str:
    """Build a bulleted paragraph. Parameters: text, level (int indent)."""
    indent = 360 + level * 360
    return (f'<w:p><w:pPr><w:pStyle w:val="Normal"/>'
            f'<w:spacing w:after="60"/>'
            f'<w:ind w:left="{indent}" w:hanging="360"/></w:pPr>'
            f'<w:r><w:t xml:space="preserve">•  {esc(text)}</w:t></w:r></w:p>')


def cell(text: str, width: int, bold: bool = False, fill: Optional[str] = None,
         color: Optional[str] = None, mono: bool = False) -> str:
    """Build one table cell. Parameters: text, width (twips), styling flags."""
    shade = f'<w:shd w:val="clear" w:fill="{fill}"/>' if fill else ""
    runProps = []
    if mono:
        runProps.append('<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="17"/>')
    if bold:
        runProps.append("<w:b/>")
    if color:
        runProps.append(f'<w:color w:val="{color}"/>')
    rPr = f"<w:rPr>{''.join(runProps)}</w:rPr>" if runProps else ""
    lines = str(text).split("\n")
    runs = ""
    for index, line in enumerate(lines):
        if index:
            runs += f"<w:r>{rPr}<w:br/></w:r>"
        runs += f'<w:r>{rPr}<w:t xml:space="preserve">{esc(line)}</w:t></w:r>'
    return (f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{shade}'
            f'<w:vAlign w:val="center"/></w:tcPr>'
            f'<w:p><w:pPr><w:spacing w:after="40"/></w:pPr>{runs}</w:p></w:tc>')


def table(header: Sequence[str], rows: Sequence[Sequence[str]],
          widths: Sequence[int], mono: bool = False) -> str:
    """Build a styled table with a blue header and zebra striping.

    Parameters:
        header (sequence) - header labels.
        rows (sequence)   - data rows.
        widths (sequence) - column widths in twips.
        mono (bool)       - render body cells in a monospace font.
    Returns: the table XML.
    """
    borders = ("<w:tblBorders>" + "".join(
        f'<w:{edge} w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV")
    ) + "</w:tblBorders>")
    grid = ("<w:tblGrid>" +
            "".join(f'<w:gridCol w:w="{w}"/>' for w in widths) + "</w:tblGrid>")
    xml = (f'<w:tbl><w:tblPr><w:tblW w:w="{sum(widths)}" w:type="dxa"/>'
           f'{borders}<w:tblLayout w:type="fixed"/></w:tblPr>{grid}')
    xml += "<w:tr><w:trPr><w:tblHeader/></w:trPr>" + "".join(
        cell(h, widths[i], bold=True, fill=BLUE, color="FFFFFF")
        for i, h in enumerate(header)) + "</w:tr>"
    for rowIndex, row in enumerate(rows):
        fill = STRIPE if rowIndex % 2 else None
        xml += "<w:tr>" + "".join(
            cell(value, widths[i], fill=fill, mono=mono)
            for i, value in enumerate(row)) + "</w:tr>"
    return xml + "</w:tbl>" + para("", spaceAfter=160)


def pageBreak() -> str:
    """Return a page-break paragraph."""
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


#: EMU per inch. Word measures drawings in English Metric Units.
EMU_PER_INCH = 914400


def pngSize(path: str) -> Tuple[int, int]:
    """Read a PNG's pixel dimensions from its IHDR chunk.

    Parameters: path (str) - the .png to measure.
    Returns: (width, height) in pixels.
    Raises: ValueError - the file is not a PNG.

    Stdlib only: the IHDR chunk is always first, and its width and height
    are two big-endian 4-byte integers at a fixed offset.
    """
    with open(path, "rb") as handle:
        header = handle.read(24)
    if header[:8] != b"\x89PNG\r\n\x1a\n":
        raise ValueError(f"{path} is not a PNG")
    width, height = struct.unpack(">II", header[16:24])
    return width, height


def image(relationshipId: str, path: str, maxWidthInches: float = 6.0,
          index: int = 1, border: bool = True, keepNext: bool = False) -> str:
    """Return a centred paragraph holding one inline picture.

    Parameters:
        relationshipId (str)   - the r:embed id, matching writeDocx's images.
        path (str)             - the .png, read only to get its aspect ratio.
        maxWidthInches (float) - scale down to this width; never scale up, so
                                 a small screenshot is not blown up and blurred.
        index (int)            - unique drawing id within the document.
        border (bool)          - draw a hairline frame, which stops a
                                 screenshot with a pale edge from bleeding
                                 into the page.
        keepNext (bool)        - keep the picture on the same page as the
                                 paragraph that follows it, so a screenshot
                                 is never separated from its caption.
    Returns: the <w:p> XML.
    """
    pixelWidth, pixelHeight = pngSize(path)
    # Screenshots are captured at 96 DPI, so pixels map straight to inches.
    widthInches = min(maxWidthInches, pixelWidth / 96)
    width = int(widthInches * EMU_PER_INCH)
    height = int(width * pixelHeight / pixelWidth)
    frame = ('<a:ln w="9525"><a:solidFill><a:srgbClr val="C9D3E4"/>'
             "</a:solidFill></a:ln>") if border else ""
    return (
        '<w:p><w:pPr>' + ("<w:keepNext/>" if keepNext else "") +
        '<w:jc w:val="center"/><w:spacing w:before="60" w:after="60"/></w:pPr>'
        "<w:r><w:drawing>"
        '<wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{width}" cy="{height}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{index}" name="Screenshot {index}"/>'
        "<wp:cNvGraphicFramePr>"
        '<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>'
        "</wp:cNvGraphicFramePr>"
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:nvPicPr><pic:cNvPr id="{index}" name="screenshot{index}.png"/>'
        "<pic:cNvPicPr/></pic:nvPicPr>"
        f'<pic:blipFill><a:blip r:embed="{relationshipId}"/>'
        "<a:stretch><a:fillRect/></a:stretch></pic:blipFill>"
        '<pic:spPr><a:xfrm><a:off x="0" y="0"/>'
        f'<a:ext cx="{width}" cy="{height}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f"{frame}</pic:spPr>"
        "</pic:pic></a:graphicData></a:graphic></wp:inline>"
        "</w:drawing></w:r></w:p>")


STYLES_XML = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/>
<w:pPr><w:spacing w:after="120"/></w:pPr>
<w:rPr><w:b/><w:sz w:val="56"/><w:color w:val="{NAVY}"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/>
<w:rPr><w:sz w:val="26"/><w:color w:val="{BLUE}"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>
<w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="2" w:color="{BLUE}"/></w:pBdr>
<w:spacing w:before="280" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr>
<w:rPr><w:b/><w:sz w:val="32"/><w:color w:val="{NAVY}"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/>
<w:pPr><w:spacing w:before="220" w:after="100"/><w:outlineLvl w:val="1"/></w:pPr>
<w:rPr><w:b/><w:sz w:val="26"/><w:color w:val="{BLUE}"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/>
<w:pPr><w:spacing w:before="160" w:after="80"/><w:outlineLvl w:val="2"/></w:pPr>
<w:rPr><w:b/><w:sz w:val="23"/><w:color w:val="{NAVY}"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Code"/>
<w:pPr><w:shd w:val="clear" w:fill="F2F2F2"/><w:spacing w:after="120"/><w:ind w:left="200"/></w:pPr>
<w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="18"/></w:rPr></w:style>
</w:styles>"""

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Default Extension="png" ContentType="image/png"/>
</Types>"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""


def writeDocx(bodyXml: str, outputPath: str, margin: int = 1080,
              images: Optional[dict] = None) -> str:
    """Zip the OOXML parts into a .docx.

    Parameters:
        bodyXml (str)    - the <w:body> content.
        outputPath (str) - destination .docx.
        margin (int)     - page margin in twips on all four sides; the
                           default 1080 is 0.75in. Pass 720 (0.5in) to fit
                           more on a single page.
        images (dict)    - {relationshipId: local .png path} for every picture
                           the body refers to. Each is copied into the package
                           and given a relationship; a missing entry shows in
                           Word as a broken-image placeholder.
    Returns: the path written.
    """
    images = dict(images or {})
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
        f"<w:body>{bodyXml}"
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        f'<w:pgMar w:top="{margin}" w:right="{margin}" '
        f'w:bottom="{margin}" w:left="{margin}"/>'
        "</w:sectPr></w:body></w:document>")
    target = Path(outputPath)
    target.parent.mkdir(parents=True, exist_ok=True)
    relationships = [DOC_RELS.rsplit("</Relationships>", 1)[0]]
    for relationshipId, source in images.items():
        relationships.append(
            f'<Relationship Id="{relationshipId}" Type="http://schemas.'
            f'openxmlformats.org/officeDocument/2006/relationships/image" '
            f'Target="media/{Path(source).name}"/>')
    relationships.append("</Relationships>")

    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", RELS)
        archive.writestr("word/_rels/document.xml.rels", "".join(relationships))
        archive.writestr("word/styles.xml", STYLES_XML)
        archive.writestr("word/document.xml", document)
        for source in images.values():
            archive.write(source, f"word/media/{Path(source).name}")
    return str(target)
