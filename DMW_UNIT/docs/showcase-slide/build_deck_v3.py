# -*- coding: utf-8 -*-
"""Build OnePI_Showcase_..._v3.pptx from the supplied v2 deck.

Slide 2 becomes visual: the two column headers are renamed, the body text
under each header is removed, and a graphic goes in its place. Slide 8's
speaker notes get the two demo scripts the TODOs asked for.
Everything else in the package is copied through untouched.
"""
import zipfile, re, io

SRC = '/root/.claude/uploads/2db2c2d8-47eb-55b9-a566-250d777a95ef/e06af718-OnePI_Showcase_11-17-2026_EDITABLE_MF_v2.pptx'
OUT = '/home/user/tmp1/DMW_UNIT/docs/OnePI_Showcase_11-17-2026_EDITABLE_MF_v3.pptx'
GFX = '/tmp/claude-0/-home-user-tmp1/2db2c2d8-47eb-55b9-a566-250d777a95ef/scratchpad/'
EMU = 914400
def emu(v): return str(int(round(v * EMU)))

zin = zipfile.ZipFile(SRC)
names = zin.namelist()
slide2 = zin.read('ppt/slides/slide2.xml').decode('utf-8')
rels2  = zin.read('ppt/slides/_rels/slide2.xml.rels').decode('utf-8')
notes8 = zin.read('ppt/notesSlides/notesSlide8.xml').decode('utf-8')

# ---------------------------------------------------------------- slide 2
# 1. rename the two column headers, keeping the leading rule glyph
for old, new in (('   |  Cortex — NLP / Unstructured Data', '   |  Unstructured (Text) Data'),
                 ('  |  Cortex Analyst — Structured Data (CoCo)', '  |  Structured (Numeric) Data')):
    assert old in slide2, old
    slide2 = slide2.replace(old, new, 1)

# 2. delete the body text under each header — the graphics carry the message now
def drop_shape(xml, name):
    i = xml.index('name="%s"' % name)
    st = xml.rindex('<p:sp>', 0, i)
    en = xml.index('</p:sp>', i) + len('</p:sp>')
    return xml[:st] + xml[en:]
for nm in ('Text 5', 'Text 9'):
    slide2 = drop_shape(slide2, nm)
assert 'Summarize or classify' not in slide2 and 'Ask IDRC' not in slide2

# 3. the graphics, filling the space the text vacated
PIC = ('<p:pic><p:nvPicPr><p:cNvPr id="{id}" name="{name}" descr="{descr}"/>'
       '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
       '<p:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
       '<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
       '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>')
pics = (PIC.format(id=201, name='Funnel Graphic', rid='rId3',
          descr='Narratives, comments, policies and regulations feed into a funnel labelled '
                'Cortex NLP, which reads every document and produces an executive summary: what '
                'the volume is saying, the themes that dominate, and what it could not classify.',
          x=emu(0.98), y=emu(2.56), cx=emu(5.30), cy=emu(3.12))
      + PIC.format(id=202, name='Conversation Graphic', rid='rId4',
          descr='A question asked in plain language returns a report and a chart; a follow-up '
                'question goes a level deeper against the national picture. A loop marked ASK '
                'AGAIN shows the analysis is a conversation, not a one-off report request.',
          x=emu(7.06), y=emu(2.56), cx=emu(5.10), cy=emu(3.12)))
slide2 = slide2.replace('</p:spTree>', pics + '</p:spTree>', 1)

rels2 = rels2.replace('</Relationships>',
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/zz_funnel.png"/>'
    '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/zz_conversation.png"/>'
    '</Relationships>')

# ---------------------------------------------------------------- slide 8 notes
def para(t):
    t = (t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;'))
    return ('<a:p><a:r><a:rPr lang="en-US" sz="1200" dirty="0"/>'
            '<a:t>%s</a:t></a:r></a:p>' % t)
BLANK = '<a:p><a:endParaRPr lang="en-US" sz="1200" dirty="0"/></a:p>'

DEMO = [
 'DEMO 1 — UNSTRUCTURED (TEXT) DATA  ·  dmw_claims_management_v1_6Beta.html  (3:30)',
 '',
 'Open on the Home tab, left pane collapsed. Say:',
 '"This is 898 public comments on one Medicare rule. A million and a half characters — the better part of a thousand pages. Today that is three weeks of reading."',
 '',
 'Press Run Pipeline — then SAY NOTHING for the four seconds it takes. The silence is the demo; talking over it throws away the only part that cannot be put on a slide.',
 '"Eight stages. Eight hundred and ninety-eight comments. Under four seconds."',
 '',
 'Click doSentimentAnalysis: "186 positive, 676 neutral, 36 negative. Neutral is not an opinion — it is a signal that the real story is somewhere else."',
 '',
 'Click summarizeText: "Every one of them summarized. A million and a half characters down to two hundred and eighty-four thousand — eighty-two percent shorter, and nothing invented: every word came out of the original comment." Click one Show more link so they see a full comment behind the summary.',
 '',
 'Click categorizeText: "Forty-nine point four percent — four hundred and forty-four comments — one theme: Provider Administrative Burden. That is not a chart. That is a staffing decision."',
 '',
 'Click createExecutiveBriefing, then Export: "And then it writes the memo. A Word document, on your desktop, ready to send."',
 '',
 'Close on the second row of the table: "Thirty-seven point six percent it could not classify. It tells you what it does not know — and that is exactly where your experts go on Monday."',
 '',
 'DELIVERY: one click, then hands behind your back. Let the cards light up. The Uncategorized line is your credibility — do not skip it to save time.',
 '',
 'ACCURACY, if asked: this page really runs. Sentiment, summaries and categories are computed in the browser on the real CMS file. The executive narrative inside the briefing is the Snowflake Cortex output from the original run, reproduced word for word. The browser summaries are extractive — Cortex writes new sentences in Snowflake. Say that plainly rather than blurring it.',
 '',
 'DEMO 2 — STRUCTURED (NUMERIC) DATA  ·  PROVIDER_RISK_AGENT_DEMO_V2_7.html  (3:00)',
 '',
 'The start-up splash states that the page is a mockup and that no Snowflake connection is made. Click through it before you start talking, or leave it up and say so — the demo site cannot reach Snowflake, and what is replayed is a real session.',
 '',
 'Click About, then Help (15 sec): "Same workbench, pointed at two hundred and seventy-nine thousand DME providers."  Collapse the left pane.',
 '',
 'Type — do not paste, the typing is the point — and press Send or Enter:',
 '',
 '1. "Show the top 20 highest-risk providers"  →  "Plain English in. Cortex writes the SQL — there it is — and returns the table."',
 '',
 '2. "Show the top 5 highest-risk DME providers in Ohio."  →  "Same question, one state. It mapped the word Ohio to the state code on its own."',
 '',
 '3. "show a heatmap of the top 20 highest-risk providers by state and risk tier"  →  "I asked for a chart in words, and I got one."',
 '',
 '4. "Now show that as a bar chart"  →  "That means the last result. No second query — it redraws what it already has."',
 '',
 '5. "why is IL so high?"  →  "A follow-up, in English. It writes a query comparing Illinois against the national averages by tier. And the answer is the interesting part: Illinois is not high on score — it matches the national average exactly. It is high on volume: three hundred and sixteen providers. That is a different action."',
 '',
 'DELIVERY: let the "Querying provider data..." spinner sit. Question five is the close — it shows the thing answering a question nobody wrote a report for.',
 '',
 'ACCURACY, if asked: this one replays a recorded Cortex Analyst session. The questions, the generated SQL and the rows are genuine output; the page makes no live call. The closing paragraph on Illinois was written from the figures in the table, not by Cortex, and the splash says so.',
 '',
 'THE 15 MINUTES',
 'Slides 1, 3, 4, 5, 6, 7 and the WHY NOW close: 7:30.  Slide 2: 0:45.  Demo 1: 3:30.  Demo 2: 3:00.  Transitions: 0:15.  Total 15:00, with nothing spare.',
 'If you are running long, cut the sentiment panel in demo 1 and the About/Help clicks in demo 2 — that buys a minute without losing an argument.',
]
new_block = ''.join(BLANK if t == '' else para(t) for t in DEMO)

# splice: keep "SHOW TWO DEMOS.", drop the two TODO paragraphs, insert the scripts
# Cut from the start of the first TODO paragraph to the end of the second.
# Both must be located by their enclosing <a:p> boundaries: the second TODO
# paragraph ends with <a:endParaRPr/>, not with a run, so searching for a
# run-terminated paragraph end runs past it and swallows the WHY NOW heading.
first  = notes8.index('TODO: produce a demo script for the unstructured data')
second = notes8.index('TODO: produce a demo script for the structured data')
start  = notes8.rindex('<a:p>', 0, first)
end    = notes8.index('</a:p>', second) + len('</a:p>')
notes8 = notes8[:start] + new_block + notes8[end:]
assert 'TODO:' not in notes8
assert 'WHY NOW' in notes8, 'the WHY NOW heading must survive'

# ---------------------------------------------------------------- write
zout = zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED)
for n in names:
    if   n == 'ppt/slides/slide2.xml':             zout.writestr(n, slide2)
    elif n == 'ppt/slides/_rels/slide2.xml.rels':  zout.writestr(n, rels2)
    elif n == 'ppt/notesSlides/notesSlide8.xml':   zout.writestr(n, notes8)
    else:                                          zout.writestr(n, zin.read(n))
zout.writestr('ppt/media/zz_funnel.png',       open(GFX + 's2_left.png',  'rb').read())
zout.writestr('ppt/media/zz_conversation.png', open(GFX + 's2_right.png', 'rb').read())
zout.close()
print('wrote', OUT)
