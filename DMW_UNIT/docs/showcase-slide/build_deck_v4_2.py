# -*- coding: utf-8 -*-
"""v4_2: fold the Approve/DRAFT instruction into what is actually said.

v4_1 carried it as a stage direction -- "Tick Approve first if you do not want
the memo to read DRAFT - NOT APPROVED" -- which told the presenter to hide the
status line. The status line is worth more said out loud: it is the governance
argument the rest of the deck makes, demonstrated on screen rather than
asserted. So the briefing is left unapproved and the DRAFT stamp becomes a
beat in the script.
"""
import copy
from pptx import Presentation
from pptx.util import Pt

SRC = '/home/user/tmp1/DMW_UNIT/docs/DMW_DEMO_OnePI_Showcase_11-19-2026_EDITABLE_MF_v4_1.pptx'
OUT = '/home/user/tmp1/DMW_UNIT/docs/DMW_DEMO_OnePI_Showcase_11-19-2026_EDITABLE_MF_v4_2.pptx'

prs = Presentation(SRC)
tf = prs.slides[7].notes_slide.notes_text_frame

idx = [i for i, par in enumerate(tf.paragraphs)
       if 'createExecutiveBriefing' in ''.join(r.text for r in par.runs)]
assert len(idx) == 1, 'expected exactly one briefing paragraph, found %d' % len(idx)
par = tf.paragraphs[idx[0]]

STAGE = ('Click createExecutiveBriefing, then Export. Leave Approve unticked — '
         'the DRAFT stamp is the point you are about to make: ')
SPOKEN = ('"And then it writes the memo. An executive summary of all 898 comments: '
          'sentiment is mostly neutral, but nearly half raise administrative burden '
          'on providers. Reimbursement, patient care quality, health equity — each '
          'quantified, with recommendations. Nobody typed this. The pipeline wrote it. '
          'And look at the top of it — it says DRAFT, NOT APPROVED. It stays a draft '
          'until a person ticks approve. The AI writes it; a human signs it. From a '
          'docket of public comments to a briefing on your desk. That’s Snowflake '
          'Cortex AI."')

# Rebuild the paragraph: one plain run for the direction, one marked run for
# the speech, matching how the rest of these notes distinguish the two.
for r in list(par.runs):
    r._r.getparent().remove(r._r)

r1 = par.add_run(); r1.text = STAGE
r1.font.size = Pt(12)

r2 = par.add_run(); r2.text = SPOKEN
r2.font.size = Pt(12); r2.font.bold = True; r2.font.italic = True; r2.font.underline = True

# The alternative, as its own direction straight after, so the choice is
# visible at the moment it has to be made.
ALT = ('(If you would rather show it approved: tick Approve before Export and the '
       'status line reads APPROVED FOR DISTRIBUTION. You lose the beat above, which '
       'is the stronger one with this audience — it answers the governance question '
       'by showing it rather than claiming it.)')
new_p = copy.deepcopy(par._p)
par._p.addnext(new_p)
from pptx.text.text import _Paragraph
alt_par = _Paragraph(new_p, par._parent)
for r in list(alt_par.runs):
    r._r.getparent().remove(r._r)
ra = alt_par.add_run(); ra.text = ALT
ra.font.size = Pt(12)

prs.save(OUT)
print('wrote', OUT)
