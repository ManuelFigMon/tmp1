# -*- coding: utf-8 -*-
"""v4_1: corrections found reviewing v4_0 against the two demos it describes.

Notes: the claims app has no left pane, the run takes four seconds not ten,
the 15-minute arithmetic did not add up, the Help vocabulary did not match
what is on screen, and v2.8's new "unavailable" reply was undocumented.

Slide 2: the loop arrow had been deleted while its caption still referred to
it, a stray "Cortex" label overlapped the arrows and the answer card, and the
two cards were different heights.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

SRC = '/root/.claude/uploads/2db2c2d8-47eb-55b9-a566-250d777a95ef/fd0ae6a6-DMW_DEMO_OnePI_Showcase_11-19-2026_EDITABLE_MF_v4_0.pptx'
OUT = '/home/user/tmp1/DMW_UNIT/docs/DMW_DEMO_OnePI_Showcase_11-19-2026_EDITABLE_MF_v4_1.pptx'

STEEL2 = RGBColor(0x9D, 0xBF, 0xD8)
prs = Presentation(SRC)
s2 = prs.slides[1]
by_name = {}
for sh in s2.shapes:
    by_name.setdefault(sh.name, []).append(sh)

# ---------------------------------------------------------------- slide 2
# 1. the stray "Cortex" label: it shares a name with the big A, is 2.40in wide
#    and lies across both the arrows and the answer card. Narrow it, drop it
#    below the arrows in the gap between the cards, and give it its own name.
stray = [sh for sh in by_name.get('Big A', []) if sh.text_frame.text.strip() == 'Cortex']
assert len(stray) == 1, 'expected exactly one stray Cortex label'
lab = stray[0]
lab.name = 'Cortex Label'
# Level with the Q? and A letters, in the clear gap between the two cards,
# so the slide reads Q? -- Cortex -- A with the loop and arrows beneath. Put
# any lower and it is squeezed between the card edges at the same height as
# the question text, which is what it looked like in v4_0.
lab.left, lab.top, lab.width, lab.height = (
    Inches(9.58), Inches(2.66), Inches(1.21), Inches(0.36))
for par in lab.text_frame.paragraphs:
    for r in par.runs:
        r.font.size = Pt(15)

# 2. restore the loop arrow. Its caption -- "Ask again, each answer leads to
#    the next question" -- was still on the slide with nothing to point at.
assert not any(sh.name == 'Loop Arrow' for sh in s2.shapes), 'loop arrow already present'
loop = s2.shapes.add_shape(MSO_SHAPE.LEFT_RIGHT_CIRCULAR_ARROW,
                           Inches(9.70), Inches(2.92), Inches(0.97), Inches(0.56))
loop.name = 'Loop Arrow'
loop.fill.solid(); loop.fill.fore_color.rgb = STEEL2
loop.line.fill.background()
loop.shadow.inherit = False

# 3. the two cards were 5.09in and 4.47in tall against the same top edge.
left_card = by_name['Shape 2'][0]
for nm in ('Shape 6', 'Shape 7'):
    by_name[nm][0].height = left_card.height
by_name['Shape 3'][0].height = left_card.height

# ---------------------------------------------------------------- notes
def fix(slide_idx, pairs):
    tf = prs.slides[slide_idx].notes_slide.notes_text_frame
    for p in tf.paragraphs:
        for r in p.runs:
            for old, new in pairs:
                if old in r.text:
                    r.text = r.text.replace(old, new)

# -- slide 2: the slide says "a report or a chart"; the note said "and"
fix(1, [('You get a report and a chart', 'You get a report or a chart')])

# -- slide 8, demo 1
fix(7, [
    # the claims app has tabs, not a sidebar; there is no pane to collapse
    ('Open on the Home tab, left pane collapsed. Say:',
     'Open on the Home tab. (This app has tabs across the top — Home, Data '
     'Grid, Executive Briefing, Ask DMW Agent, Console — and no left pane, '
     'so there is nothing to collapse here.) Say:'),
    # measured at 3.9-4.0 seconds; the next sentence already said four
    # These two sit in separate runs in the file, so they are replaced apart.
    ('Under 10 seconds.', 'Under four seconds.'),
    ('weeks of reading, done while we talk.', 'days of reading, done while we talk.'),
    # the briefing opens stamped DRAFT unless Approve is ticked
    ('Click createExecutiveBriefing, then Export:',
     'Click createExecutiveBriefing. Tick Approve first if you do not want the '
     'memo to read DRAFT — NOT APPROVED, then Export:'),
    ('Close on the second row of the table:',
     'Close on the second row of the thematic table inside the briefing:'),
])

# -- slide 8, demo 2
fix(7, [
    # Help lists STATE_COUNT and CLAIM_COUNT; there is no "ownership count" there
    ('a multi-state score for providers operating across state lines, and their '
     'claim and ownership counts',
     'a multi-state score for providers operating across state lines, the number '
     'of states their ownership spans, and their matched claim count'),
    ('Question five is the close',
     'If you mistype a question the app answers “The system is currently '
     'unavailable. Please try again later.” and lists the five it knows — '
     'that is the v2.8 wording, not a failure. Retype and carry on. '
     'Question five is the close'),
])

# -- slide 8: the timing block did not add up and is easier to use at the end
OLD_T = ('Slides 1, 3, 4, 5, 6, 7 and the WHY NOW close: 7:30.  Slide 2: 0:45.  '
         'Demo 1: 3:30.  Demo 2: 3:00.  Transitions: 0:15.  Total 15:00, with nothing spare.')
NEW_T = ('Slides 1, 3, 4, 5, 6, 7 and the WHY NOW close: 6:45 (0:30 + 1:00 + 0:45 + '
         '1:30 + 1:00 + 1:00 + 1:00, as timed in each slide’s own notes).  '
         'Slide 2: 0:45.  Demo 1: 3:30.  Demo 2: 3:00.  Transitions: 0:15.  '
         'Total 14:15 — about 45 seconds of slack inside the 15 minutes.')
fix(7, [(OLD_T, NEW_T)])

prs.save(OUT)
print('wrote', OUT)
