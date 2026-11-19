const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle, LevelFormat
} = require('docx');

const OUT  = '/home/user/tmp1/DMW_UNIT/docs/OnePI_Showcase_v4_Review_and_Delivery_Notes.docx';
const NAVY = '1F3864', BLUE = '4A7BA7', GREY = '595959', HDRBG = 'DCE6F1', ALTBG = 'F2F6FA';
const RED  = '9C2B28', AMBER = '8A6100', GREEN = '2E7D32';

const RP = (parts, o = {}) => new Paragraph({
  spacing: { after: o.after === undefined ? 130 : o.after, line: 276 },
  indent: o.indent,
  children: (Array.isArray(parts) ? parts : [[parts]]).map(([t, f = {}]) => new TextRun({
    text: t, bold: f.b, italics: f.i,
    font: f.code ? 'Consolas' : undefined, size: f.code ? 18 : o.size,
    color: f.code ? 'A31515' : (f.color || o.color)
  }))
});
const P = (t, o = {}) => RP([[t, {}]], o);
const H1 = t => new Paragraph({
  spacing: { before: 320, after: 140 }, keepNext: true,
  children: [new TextRun({ text: t, font: 'Arial', size: 30, bold: true, color: NAVY })]
});
const H2 = t => new Paragraph({
  spacing: { before: 250, after: 110 }, keepNext: true,
  children: [new TextRun({ text: t, font: 'Arial', size: 23, bold: true, color: BLUE })]
});
const BULLET = (parts) => new Paragraph({
  numbering: { reference: 'bullets', level: 0 },
  spacing: { after: 70, line: 276 },
  children: (Array.isArray(parts) ? parts : [[parts]]).map(([t, f = {}]) => new TextRun({
    text: t, bold: f.b, italics: f.i,
    font: f.code ? 'Consolas' : undefined, size: f.code ? 18 : undefined,
    color: f.code ? 'A31515' : (f.color || undefined)
  }))
});
const NUM = (parts) => new Paragraph({
  numbering: { reference: 'steps', level: 0 },
  spacing: { after: 70, line: 276 },
  children: (Array.isArray(parts) ? parts : [[parts]]).map(([t, f = {}]) => new TextRun({
    text: t, bold: f.b, italics: f.i,
    font: f.code ? 'Consolas' : undefined, size: f.code ? 18 : undefined,
    color: f.code ? 'A31515' : undefined
  }))
});
function inline(line, head) {
  const out = []; const re = /`([^`]+)`|\*\*([^*]+)\*\*/g; let last = 0, m;
  while ((m = re.exec(line))) {
    if (m.index > last) out.push(new TextRun({ text: line.slice(last, m.index), bold: head }));
    if (m[1]) out.push(new TextRun({ text: m[1], font: 'Consolas', size: 17, color: 'A31515', bold: head }));
    else out.push(new TextRun({ text: m[2], bold: true }));
    last = re.lastIndex;
  }
  if (last < line.length) out.push(new TextRun({ text: line.slice(last), bold: head }));
  if (!out.length) out.push(new TextRun({ text: ' ', bold: head }));
  return out;
}
function TBL(headers, rows, widths) {
  const cell = (txt, { head = false, alt = false, w } = {}) => new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: head ? HDRBG : (alt ? ALTBG : 'FFFFFF'), color: 'auto' },
    margins: { top: 65, bottom: 65, left: 110, right: 110 },
    children: String(txt).split('\n').map((line, i, a) => new Paragraph({
      spacing: { after: i === a.length - 1 ? 0 : 40, line: 250 }, children: inline(line, head)
    }))
  });
  return new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    rows: [
      new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, { head: true, w: widths[i] })) }),
      ...rows.map((r, ri) => new TableRow({ children: r.map((c, i) => cell(c, { alt: ri % 2 === 1, w: widths[i] })) }))
    ]
  });
}
const CALLOUT = (title, parts, color) => new Paragraph({
  spacing: { before: 140, after: 160, line: 276 },
  shading: { type: ShadingType.CLEAR, fill: color === RED ? 'FDF0EF' : (color === GREEN ? 'EEF7EE' : 'FFF8E8'), color: 'auto' },
  border: { left: { style: BorderStyle.SINGLE, size: 18, color: color || 'F0AD4E', space: 8 } },
  indent: { left: 180, right: 120 },
  children: [new TextRun({ text: title + '  ', bold: true, color: color || AMBER })]
    .concat(parts.map(([t, f = {}]) => new TextRun({
      text: t, bold: f.b, italics: f.i,
      font: f.code ? 'Consolas' : undefined, size: f.code ? 18 : undefined,
      color: f.code ? 'A31515' : undefined
    })))
});
const RULE = () => new Paragraph({
  spacing: { before: 200, after: 120 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: BLUE, space: 2 } }, children: []
});

/* ===================================================================== */
const children = [
  new Paragraph({ spacing: { after: 60 },
    children: [new TextRun({ text: 'OnePI Showcase — Review and Delivery Notes',
      font: 'Arial', size: 36, bold: true, color: NAVY })] }),
  new Paragraph({ spacing: { after: 40 },
    children: [new TextRun({ text: 'DMW_DEMO_OnePI_Showcase_11-19-2026_EDITABLE_MF · v4_0 reviewed, v4_2 issued',
      font: 'Arial', size: 24, color: BLUE })] }),
  P('Prepared for a 15-minute VP-level presentation with two live demos.', { color: GREY, size: 19 }),

  CALLOUT('Bottom line.', [['The deck is in good shape and the story holds. Eight claims in the demo scripts were checked against the two applications by running them; all the headline numbers are correct. '], ['Five errors in the demo scripts and three layout defects on slide 2 are fixed in v4_1, and the Approve/DRAFT beat is rewritten in v4_2', { b: true }], [' — one of them would have had you looking for a control that does not exist, in front of the room. Three further items need a decision only you can make, and are listed in section 3.']], GREEN),

  RULE(),
  H1('1. What was verified, and checks out'),
  P('Each of these was confirmed by running the application, not by reading the notes.', { after: 120 }),
  TBL(['Claim in the notes', 'Verified'], [
    ['898 public comments, eight stages', 'Yes — 8 stage cards, 898 rows'],
    ['186 positive / 676 neutral / 36 negative', 'Yes — exactly these figures'],
    ['1.5M characters down to 284K, 82% shorter', 'Yes — 1,578K → 284K, 82.0%'],
    ['49.4%, 444 comments, Provider Administrative Burden', 'Yes — row 1 of the thematic table'],
    ['37.6% Uncategorized is the second row', 'Yes — 338 comments, row 2'],
    ['The briefing exports as a Word document', 'Yes — `cgs_executive_briefing.docx`'],
    ['~279K DME providers; top 20 returns names, NPIs, states, scores, tiers',
     'Yes — 8 columns, 20 rows'],
    ['Illinois matches the national average and is high on volume (316)',
     'Yes — 40.3 vs 40.3 national, 316 providers']
  ], [5600, 4400]),
  P(''),
  P('The two demo files named in the notes are the current ones: dmw_claims_management_v1_6Beta.html and PROVIDER_RISK_AGENT_DEMO_V2_8.html.'),

  RULE(),
  H1('2. What was wrong, and is fixed in v4_1'),

  H2('2.1  The claims demo has no left pane  (would have cost you time on stage)'),
  RP([['The notes opened demo 1 with '], ['“Open on the Home tab, left pane collapsed”', { i: true }], ['. That app has no left pane. It has tabs across the top — Home, Data Grid, Executive Briefing, Ask DMW Agent, Console. The sidebar belongs to the '], ['other', { i: true }], [' demo. You would have hunted for a control that is not there while the room watched.']]),
  P('Fixed: the instruction now names the tabs and says plainly that there is nothing to collapse here.'),

  H2('2.2  Two different numbers for the same four seconds'),
  RP([['One sentence said '], ['“SAY NOTHING for the four seconds it takes”', { i: true }], ['; the next said '], ['“Under 10 seconds”', { i: true }], ['. Timed on a cold run it is 3.9 to 4.0 seconds. Fixed to four seconds — which is the stronger claim anyway.']]),

  H2('2.3  The 15-minute budget did not add up'),
  P('The block said the slides take 7:30. Adding up the timings written in each slide’s own notes gives 6:45 — 0:30 + 1:00 + 0:45 + 1:30 + 1:00 + 1:00 + 1:00. The real total is 14:15, so you have about 45 seconds of slack rather than “nothing spare”. Fixed, with the arithmetic shown so it can be re-checked if a slide changes.'),

  H2('2.4  The Help-panel script named a field that is not on screen'),
  RP([['The script says the tool knows '], ['“claim and ownership counts”', { i: true }], ['. The Help panel actually lists '], ['STATE_COUNT', { code: true }], [' (states the ownership spans) and '], ['CLAIM_COUNT', { code: true }], ['. If a VP reads along while you speak, the mismatch is visible. Reworded to match the panel.']]),

  H2('2.5  v2.8 changed what happens when you mistype — and nobody told the script'),
  RP([['In the current build an unrecognised question answers '], ['“The system is currently unavailable. Please try again later.”', { i: true }], [' If you fumble a question live and see that, it looks like the demo just died. It has not. The notes now say so, and tell you to retype and carry on.']]),
  CALLOUT('This one is worth rehearsing.', [['Deliberately mistype a question once in practice so you have seen the message before an audience does. Your recovery line can be as simple as “and there is the error handling — let me type that properly.”']], AMBER),

  H2('2.6  Slide 2 — three layout defects'),
  BULLET([['The '], ['loop arrow had been deleted', { b: true }], [' but its caption — “Ask again — each answer leads to the next question” — was still on the slide, pointing at nothing. Restored.']]),
  BULLET([['A stray '], ['“Cortex”', { b: true }], [' label 2.4 inches wide lay across both the arrows and the answer card. It is now set level with the Q? and A letters, so the slide reads Q? — Cortex — A with the loop beneath.']]),
  BULLET([['The two cards were '], ['different heights', { b: true }], [' (5.09in and 4.47in) against the same top edge. Squared up.']]),
  P('The shape also carried a duplicate name ("Big A", shared with the letter A), which made the selection pane confusing. Renamed to "Cortex Label".'),

  RULE(),
  H1('3. Three judgement calls \u2014 two still yours'),

  H2('3.1  “We’ve already built it at CMS in OnePI”'),
  RP([['Slide 6’s notes say this. An earlier version said simply “in OnePI”. '], ['Built at CMS', { b: true }], [' is a stronger claim about who owns the work, and I cannot verify it from here. If CGS built it in a CMS-hosted environment, say that precisely — a VP will hear the difference, and so will anyone from CMS in the room. '], ['Left unchanged in v4_1 pending your call.', { b: true }]]),

  H2('3.2  “That’s days of reading, done while we talk”'),
  RP([['v4_0 said '], ['weeks', { i: true }], ['. 898 comments is about 1.58 million characters — roughly a thousand pages, or two to three working days of careful reading, not weeks. I changed it to '], ['days', { b: true }], [' because the overstatement is the kind a VP checks afterwards, and losing that one number costs you the rest. If you have a real measured baseline from a reviewer, use it and say where it came from — that is stronger than either word.']]),

  H2('3.3  The exported briefing says DRAFT — NOT APPROVED  (resolved in v4_2)'),
  RP([['The memo carries a status line that reads DRAFT — NOT APPROVED unless you tick '], ['Approve', { b: true }], [' in the briefing panel first. v4_1 handled this with a stage direction telling you to tick it. That was the wrong way round: it told you to '], ['hide', { i: true }], [' the status line.']]),
  RP([['v4_2 does the opposite and says it out loud. The briefing is left unapproved and the script now reads: '], ['“… Nobody typed this. The pipeline wrote it. And look at the top of it — it says DRAFT, NOT APPROVED. It stays a draft until a person ticks approve. The AI writes it; a human signs it.”', { i: true }]]),
  RP([['That turns a cosmetic blemish into the governance argument the rest of the deck only asserts — demonstrated on screen instead of claimed. The alternative is still in the notes directly beneath, in case you would rather show it approved.']]),

  RULE(),
  H1('4. Risks during the live demo, and how to recover'),
  TBL(['If this happens', 'What to do'], [
    ['**The claims pipeline looks frozen**', 'It finishes in about four seconds. If the page was just loaded, the first run reads the data files — run it once before the audience arrives so it is warm.'],
    ['**A typed question is not recognised** (demo 2)', 'You will see "The system is currently unavailable." Retype it. The five recorded questions are in the notes; the wording is matched loosely, so a trailing full stop or lower-case start is fine.'],
    ['**Someone asks you to try their own question** (demo 2)', 'It will answer "currently unavailable." Say plainly that this build replays a recorded session because the demo site cannot reach Snowflake, and offer to run their question live afterwards. Do not improvise an answer.'],
    ['**The Ask DMW Agent tab is blank** (demo 1)', 'The agent’s host refuses to be framed. Use "Open in a new tab". Better: do not open that tab at all — it is not in the script.'],
    ['**You are running long**', 'Cut the sentiment panel in demo 1 and the About/Help clicks in demo 2. That is about a minute and costs you no argument.'],
    ['**A chart or table looks wrong**', 'Both demos are deterministic — the same docket always produces the same figures. If a number differs from your notes, you are on a different source; check the Source picker.']
  ], [3400, 6600]),

  RULE(),
  H1('5. Delivering this to a VP'),

  H2('5.1  The four things a VP is actually listening for'),
  NUM([['What does it cost, and what do I have to approve?', { b: true }], [' Your answer is strong — already owned, no licence, no procurement. Lead with it the moment cost comes up; do not wait for the Q&A.']]),
  NUM([['What happens if it is wrong?', { b: true }], [' The Uncategorized number is your best asset here. A tool that reports what it could not classify is more credible than one that claims 100%. Say it before you are asked.']]),
  NUM([['Who is accountable?', { b: true }], [' “It reads and organises. It never decides.” Land that sentence; it closes the governance question before it opens.']]),
  NUM([['What do you want from me?', { b: true }], [' Have one specific ask ready — a named proof of concept, a timebox, and who you need for how long. A VP who likes it will ask what the next step is, and “we’d like to explore further” wastes the moment.']]),

  H2('5.2  Pacing'),
  BULLET([['The four seconds of silence while the pipeline runs is the most persuasive moment in the deck. Do not talk over it. It will feel like twenty seconds to you and like a magic trick to them.']]),
  BULLET([['Slow down on slide 5. It is the argument; everything else is evidence for it.']]),
  BULLET([['Type the questions in demo 2 — do not paste. The typing is what proves there is no trick.']]),
  BULLET([['Let the “Querying provider data…” spinner sit. Hiding the wait makes it look faked.']]),

  H2('5.3  Where to put your strongest material'),
  RP([['The Illinois question is your closing argument, not a footnote. It is the only moment that shows the tool answering something '], ['nobody wrote a report for', { b: true }], [' — and the answer is counter-intuitive: Illinois is not high on score, it is high on volume, which is a different management action. If you only get through one demo, get through that question.']]),

  H2('5.4  Things not to say'),
  BULLET([['Do not say “the AI decides” or “automatically approves”. Not true, and it invites a compliance objection you do not need.']]),
  BULLET([['Do not claim an accuracy percentage. The honest line is already in the notes: that is what the proof of concept measures, on our own data.']]),
  BULLET([['Do not describe the provider demo as live. It replays a recorded session; the splash says so. If asked directly, say so plainly — the questions, the SQL and the rows are genuine output from a real session.']]),
  BULLET([['Do not read the slide. Slide 2 is now a picture precisely so you can talk over it.']]),

  RULE(),
  H1('6. Pre-flight checklist'),
  NUM([['Open both demo URLs and run each one once, start to finish, so the data files are warm.']]),
  NUM([['In demo 1, decide now whether you are ticking Approve before Export (section 3.3).']]),
  NUM([['In demo 2, click through the start-up splash once so you know what it says.']]),
  NUM([['Mistype one question in demo 2 on purpose, so the "currently unavailable" message is not a surprise.']]),
  NUM([['Check the browser zoom. The SQL block and the risk-tier colours are the two things people squint at.']]),
  NUM([['Have the one specific ask written on a card.']]),
  P(''),
  CALLOUT('One last thing.', [['Both applications are deterministic and run without a network. Nothing in this demo depends on a connection, a login or a warehouse being awake — which is unusual and worth saying out loud if the room has sat through a demo that failed.']], GREEN)
];

const doc = new Document({
  creator: 'DMW — Data Mart Workbench',
  title: 'OnePI Showcase v4 — Review and Delivery Notes',
  numbering: { config: [
    { reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•',
       alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 460, hanging: 240 } } } }] },
    { reference: 'steps', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.',
       alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 460, hanging: 240 } } } }] }
  ]},
  styles: { default: { document: { run: { font: 'Calibri', size: 21, color: '262626' } } } },
  sections: [{ properties: { page: { margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } } }, children }]
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log('wrote', OUT, b.length, 'bytes'); });
