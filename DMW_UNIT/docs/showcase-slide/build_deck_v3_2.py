# -*- coding: utf-8 -*-
"""Build v3_2 of the OnePI showcase deck from the supplied v3_1.

  - slide 2: both graphics replaced (bigger type; the right one redrawn as
    Q <-> A with a loop), and speaker notes added on cost and compliance;
  - slide 8: every word that is spoken aloud is set bold + italic +
    underlined, so the presenter can see at a glance what is script and what
    is stage direction.

Everything else in the package is copied through untouched.
"""
import zipfile, re, io

SRC = '/root/.claude/uploads/2db2c2d8-47eb-55b9-a566-250d777a95ef/1b938113-OnePI_Showcase_11-18-2026_EDITABLE_MF_v3_1.pptx'
OUT = '/home/user/tmp1/DMW_UNIT/docs/OnePI_Showcase_11-18-2026_EDITABLE_MF_v3_2.pptx'
GFX = '/tmp/claude-0/-home-user-tmp1/2db2c2d8-47eb-55b9-a566-250d777a95ef/scratchpad/'

zin = zipfile.ZipFile(SRC)
names = zin.namelist()

# ------------------------------------------------------------- slide 2 notes
def esc(t):
    return t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

SPOKEN = dict(b='1', i='1', u='sng')

def run(t, spoken=False):
    """A run of notes text. Spoken text is bold, italic and underlined."""
    attrs = ' lang="en-US" sz="1200"'
    if spoken:
        attrs += ' b="1" i="1" u="sng"'
    return '<a:r><a:rPr%s dirty="0"/><a:t>%s</a:t></a:r>' % (attrs, esc(t))

def para_plain(t):
    return '<a:p>%s</a:p>' % run(t) if t else BLANK

BLANK = '<a:p><a:endParaRPr lang="en-US" sz="1200" dirty="0"/></a:p>'

class Marker(object):
    """Marks the spoken text in a run of note paragraphs.

    Quoted text is what gets said out loud. A quote can run across several
    paragraphs -- the WHY NOW close is one quotation four paragraphs long --
    so the open/closed state has to carry from one paragraph to the next.
    Splitting each paragraph on its own quotes dropped the opening mark on
    the last paragraph and invented a closing one on the first.

    The text itself is never rebuilt: runs are cut at the quote marks and the
    characters pass through exactly as they were.
    """

    def __init__(self):
        self.inside = False

    def para(self, t):
        if not t:
            return BLANK                      # a blank line does not end a quote
        out, buf = [], []

        def flush():
            if buf:
                out.append(run(''.join(buf), self.inside))
                del buf[:]

        for ch in t:
            if ch == '"':
                if self.inside:               # the closing mark is part of the speech
                    buf.append(ch)
                    flush()
                    self.inside = False
                else:
                    flush()
                    self.inside = True
                    buf.append(ch)
            else:
                buf.append(ch)
        flush()
        return '<a:p>%s</a:p>' % ''.join(out)

NOTES2 = [
 'SHOWCASE 1 — THE WHOLE SLIDE IN ONE LINE (~45 sec):',
 '',
 '"Two kinds of claims data. On the left, everything written in words. On the right, everything already in columns. The same AI handles both — and in both cases you ask in English and get something you can act on."',
 '',
 'LEFT, walking the funnel: "Narratives, comments, policies, regulations go in. Cortex reads every document, and what comes out is an executive summary — what the volume is saying, the themes that dominate, and what it could not classify."',
 '',
 'RIGHT, pointing at the loop: "You ask a question. You get a report and a chart, with the SQL it wrote. Then you ask the next question. That loop is the point — it is a conversation with the data, not a report request that takes three weeks."',
 '',
 'COST — if asked what this costs:',
 'Cortex is already part of our Snowflake footprint, so there is no new licence, no new vendor and no procurement cycle. The compute is the same warehouse we already pay for, and it runs in seconds per question rather than hours per analyst. The cost to try this is a few days of our own people, not a capital request.',
 'Say it as: "The cheapest part of this proposal is the software. We already own it."',
 '',
 'COMPLIANCE — if asked about risk:',
 'The data never leaves Snowflake. There is no external model call, no copy sent to a vendor, nothing crossing our security boundary — so the existing access controls, encryption and audit logging continue to apply unchanged.',
 'Every answer is traceable: the structured side shows the exact SQL it wrote, and the unstructured side cites the document behind each extract. Nothing is a black box that cannot be checked.',
 'A person still decides. The AI finds, organises and summarises the evidence; the clinician or reviewer makes the determination and signs it.',
 'Say it as: "It reads and organises. It never decides, and it never leaves our four walls."',
 '',
 'IF PRESSED ON ACCURACY: the honest answer is that the classification on the left is rule-based and reports what it could not classify rather than guessing — that share is on the screen in the demo. Offer the proof of concept as the way to measure accuracy on our own data.',
]
notes2 = zin.read('ppt/notesSlides/notesSlide2.xml').decode('utf-8')
mk2 = Marker()
body2 = ''.join(mk2.para(t) for t in NOTES2)
assert not mk2.inside, 'slide 2 notes leave a quote open'
old_empty = ('<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US" dirty="0">'
             '<a:latin typeface="Aptos" panose="020B0004020202020204" pitchFamily="34" charset="0"/>'
             '</a:endParaRPr></a:p></p:txBody>')
assert old_empty in notes2, 'notes 2 body placeholder not where expected'
notes2 = notes2.replace(old_empty, '<p:txBody><a:bodyPr/><a:lstStyle/>%s</p:txBody>' % body2, 1)

# ------------------------------------------------------------- slide 8 notes
# Rebuild the body so every quoted sentence is marked as spoken. The text is
# taken from the deck as it stands, so nothing is reworded here.
notes8 = zin.read('ppt/notesSlides/notesSlide8.xml').decode('utf-8')
# The extracted fragment has no namespace declarations of its own, so it is
# parsed by pulling the paragraphs out with a regex rather than by XML.
body_start = notes8.index('<p:ph type="body"')
body_txb = notes8.index('<p:txBody>', body_start)
body_end = notes8.index('</p:txBody>', body_txb) + len('</p:txBody>')
old_body = notes8[body_txb:body_end]
body_paras = []
for m in re.finditer(r'<a:p>(.*?)</a:p>', old_body, re.S):
    body_paras.append(''.join(re.findall(r'<a:t>(.*?)</a:t>', m.group(1), re.S)))
# unescape what the XML held
body_paras = [t.replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>')
              for t in body_paras]
mk8 = Marker()
new_body = '<p:txBody><a:bodyPr/><a:lstStyle/>%s</p:txBody>' % ''.join(
    mk8.para(t) for t in body_paras)
assert not mk8.inside, 'slide 8 notes leave a quote open'
notes8 = notes8[:body_txb] + new_body + notes8[body_end:]

# ------------------------------------------------------------- write
zout = zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED)
for n in names:
    if   n == 'ppt/notesSlides/notesSlide2.xml': zout.writestr(n, notes2)
    elif n == 'ppt/notesSlides/notesSlide8.xml': zout.writestr(n, notes8)
    elif n == 'ppt/media/image9.png':            zout.writestr(n, open(GFX + 's2v2_left.png',  'rb').read())
    elif n == 'ppt/media/image10.png':           zout.writestr(n, open(GFX + 's2v2_right.png', 'rb').read())
    else:                                        zout.writestr(n, zin.read(n))
zout.close()
print('wrote', OUT)
print('slide-2 note paragraphs:', len(NOTES2))
print('slide-8 note paragraphs:', len(body_paras))
