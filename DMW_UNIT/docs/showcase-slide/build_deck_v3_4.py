# -*- coding: utf-8 -*-
"""v3_4: rebuild slide 2's two graphics as native PowerPoint shapes.

v3_3 carried them as flat PNGs, so nothing in them could be restyled. Every
piece is now an autoshape or a text box: select it and change the font size,
the colour, the wording, or drag it somewhere else.
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

SRC = '/root/.claude/uploads/2db2c2d8-47eb-55b9-a566-250d777a95ef/d17bac00-OnePI_Showcase_11-18-2026_EDITABLE_MF_v3_3.pptx'
OUT = '/home/user/tmp1/DMW_UNIT/docs/OnePI_Showcase_11-18-2026_EDITABLE_MF_v3_4.pptx'

TEAL  = RGBColor(0x45, 0xA5, 0xC4)
TEAL2 = RGBColor(0x3D, 0x97, 0xB8)
STEEL = RGBColor(0x2E, 0x6F, 0x9E)
STEEL2= RGBColor(0x9D, 0xBF, 0xD8)
STEEL3= RGBColor(0xC3, 0xD7, 0xE6)
INK   = RGBColor(0x1E, 0x29, 0x3B)
BODY  = RGBColor(0x64, 0x74, 0x8B)
LINE  = RGBColor(0xD9, 0xE2, 0xEC)
CHIP  = RGBColor(0xF4, 0xF8, 0xFA)
CHIPB = RGBColor(0xBF, 0xE0, 0xEB)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

prs = Presentation(SRC)
slide = prs.slides[1]
shapes = slide.shapes

# ---- drop the two pictures -------------------------------------------------
for sh in list(shapes):
    if sh.shape_type is not None and sh.name in ('Funnel Graphic', 'Conversation Graphic'):
        sh._element.getparent().remove(sh._element)

# ---- helpers ---------------------------------------------------------------
def box(x, y, w, h, shape=MSO_SHAPE.ROUNDED_RECTANGLE, fill=None, line=None,
        lw=1.0, name=None, adj=None):
    s = shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid(); s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line; s.line.width = Pt(lw)
    s.shadow.inherit = False
    if name:
        s.name = name
    if adj is not None:
        try: s.adjustments[0] = adj
        except Exception: pass
    s.text_frame.word_wrap = True
    return s

def label(shape, lines, size=12, color=INK, bold=False, align=PP_ALIGN.CENTER,
          anchor=MSO_ANCHOR.MIDDLE):
    """Put text on a shape. Each entry is (text, size, bold, color) or a str."""
    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.06)
    tf.margin_top = tf.margin_bottom = Inches(0.03)
    for i, item in enumerate(lines):
        t, sz, bd, cl = (item if isinstance(item, tuple) else (item, size, bold, color))
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run(); r.text = t
        r.font.size = Pt(sz); r.font.bold = bd; r.font.color.rgb = cl
    return shape

def textbox(x, y, w, h, lines, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, name=None):
    tb = shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name: tb.name = name
    label(tb, lines, align=align, anchor=anchor)
    return tb

def connector(x1, y1, x2, y2, color, lw=1.5, name=None):
    from pptx.enum.shapes import MSO_CONNECTOR
    c = shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                             Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = color; c.line.width = Pt(lw)
    if name: c.name = name
    return c

# ============================================================== LEFT : funnel
# The region is the inside of the left card. v3_3's picture was 0.43in wider
# than the card it sat in and overhung the edge; these shapes fit inside it.
LX, LY, LW = 0.44, 2.55, 5.90

CHIPS = ['Narratives', 'Comments', 'Policies', 'Regulations']
cw, gap = 1.38, 0.12
for i, t in enumerate(CHIPS):
    cx = LX + i * (cw + gap)
    s = box(cx, LY, cw, 0.32, fill=CHIP, line=CHIPB, lw=1.25, name='Chip %s' % t, adj=0.5)
    label(s, [(t, 12, True, INK)])

# the four feeds converging on the funnel mouth
MOUTH_L, MOUTH_R = 1.84, 4.94
MOUTH_Y = 3.10
for i in range(4):
    cx = LX + i * (cw + gap) + cw / 2
    tx = MOUTH_L + 0.30 + i * ((MOUTH_R - MOUTH_L - 0.60) / 3)
    connector(cx, LY + 0.34, tx, MOUTH_Y - 0.02, CHIPB, 1.5, 'Feed %d' % (i + 1))

FUNNEL_H = 1.05
funnel = box(MOUTH_L, MOUTH_Y, MOUTH_R - MOUTH_L, FUNNEL_H,
             shape=MSO_SHAPE.FLOWCHART_MERGE, fill=TEAL, name='Funnel')
stem = box((MOUTH_L + MOUTH_R) / 2 - 0.28, MOUTH_Y + FUNNEL_H - 0.02, 0.56, 0.28,
           shape=MSO_SHAPE.RECTANGLE, fill=TEAL2, name='Funnel Stem')

# The labels live in their own box so the text can be resized freely -- and
# anchor to the TOP of it, because the cone narrows fast and text centred in
# the box lands where the shape is too narrow to hold it.
textbox(MOUTH_L + 0.31, MOUTH_Y + 0.04, (MOUTH_R - MOUTH_L) - 0.62, 0.52,
        [('Cortex — NLP', 16, True, WHITE),
         ('reads every document', 10.5, False, WHITE)],
        anchor=MSO_ANCHOR.TOP, name='Funnel Label')

box((MOUTH_L + MOUTH_R) / 2 - 0.09, MOUTH_Y + FUNNEL_H + 0.30, 0.18, 0.24,
    shape=MSO_SHAPE.DOWN_ARROW, fill=TEAL2, name='Funnel Outflow')

card = box(LX + 0.42, 4.72, LW - 0.84, 1.40, fill=WHITE, line=LINE, lw=1.25,
           name='Executive Summary Card', adj=0.06)
box(LX + 0.58, 4.86, 0.05, 0.20, shape=MSO_SHAPE.RECTANGLE, fill=TEAL, name='Summary Accent')
textbox(LX + 0.68, 4.80, LW - 1.30, 0.30, [('Executive summary', 14, True, INK)],
        align=PP_ALIGN.LEFT, name='Summary Heading')
textbox(LX + 0.68, 5.12, LW - 1.30, 0.92,
        [('•  What the volume is saying', 11.5, False, BODY),
         ('•  The themes that dominate, ranked', 11.5, False, BODY),
         ('•  What it could not classify', 11.5, False, BODY)],
        align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, name='Summary Lines')

# ============================================================== RIGHT : Q <-> A
RX, RY, RW = 7.26, 2.55, 5.85
COL = 2.40
QX, AX = RX, RX + RW - COL
MIDX, MIDW = QX + COL, RW - 2 * COL

textbox(QX, RY, COL, 0.86, [('Q', 46, True, STEEL)], name='Big Q')
textbox(AX, RY, COL, 0.86, [('A', 46, True, STEEL)], name='Big A')

qcard = box(QX, 3.46, COL, 1.86, fill=STEEL, line=STEEL, lw=1.0, name='Question Card', adj=0.07)
label(qcard, [('“Which providers\nare highest risk?”', 14, True, WHITE)])

acard = box(AX, 3.46, COL, 1.86, fill=WHITE, line=LINE, lw=1.25, name='Answer Card', adj=0.07)

# the little histogram, five rectangles so each bar can be edited on its own
BARW, BARG = 0.12, 0.07
heights = [0.17, 0.28, 0.46, 0.24, 0.13]
total = len(heights) * BARW + (len(heights) - 1) * BARG
bx = AX + (COL - total) / 2
BASE = 4.32
cols = [STEEL3, STEEL2, STEEL, STEEL2, STEEL3]
for i, hgt in enumerate(heights):
    box(bx + i * (BARW + BARG), BASE - hgt, BARW, hgt,
        shape=MSO_SHAPE.RECTANGLE, fill=cols[i], name='Bar %d' % (i + 1))

textbox(AX + 0.08, 4.38, COL - 0.16, 0.86,
        [('A report, and a chart', 13, True, INK),
         ('ranked, color-coded,\nwith the SQL it wrote', 11, False, BODY)],
        name='Answer Caption')

# question out, answer back, and the loop that says you can keep going
box(MIDX + 0.06, 3.60, MIDW - 0.12, 0.20, shape=MSO_SHAPE.RIGHT_ARROW,
    fill=STEEL, name='Arrow Out')
box(MIDX + 0.06, 3.92, MIDW - 0.12, 0.20, shape=MSO_SHAPE.LEFT_ARROW,
    fill=STEEL, name='Arrow Back')
# A double-headed arc: it reads as a loop and carries a head at each end,
# which the plain circular-arrow preset does not render reliably.
box(MIDX + 0.04, 2.92, MIDW - 0.08, 0.56,
    shape=MSO_SHAPE.LEFT_RIGHT_CIRCULAR_ARROW, fill=STEEL2, name='Loop Arrow')

textbox(RX, 5.46, RW, 0.34,
        [('Ask again — each answer leads to the next question', 12, True, STEEL)],
        name='Loop Caption')

prs.save(OUT)
print('wrote', OUT)
print('slide 2 shapes now:', len(slide.shapes))
for sh in slide.shapes:
    print('   %-16s %s' % (str(sh.shape_type), sh.name))
