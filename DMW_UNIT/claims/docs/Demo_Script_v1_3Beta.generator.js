const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle, LevelFormat
} = require('docx');

const OUT  = '/home/user/tmp1/DMW_UNIT/claims/docs/Demo_Script_v1_3Beta.docx';
const NAVY = '1F3864', BLUE = '4A7BA7', GREY = '595959', HDRBG = 'DCE6F1', ALTBG = 'F2F6FA';

const RP = (parts, o = {}) => new Paragraph({
  spacing: { after: o.after === undefined ? 130 : o.after, line: o.line || 276 },
  alignment: o.align,
  indent: o.indent,
  children: (Array.isArray(parts) ? parts : [[parts]]).map(([t, f = {}]) => new TextRun({
    text: t, bold: f.b, italics: f.i, allCaps: f.caps,
    font: f.code ? 'Consolas' : o.font,
    size: f.size || o.size, color: f.color || o.color
  }))
});
const P = (t, o = {}) => RP([[t, {}]], o);

const H1 = t => new Paragraph({
  spacing: { before: 320, after: 140 }, keepNext: true,
  children: [new TextRun({ text: t, font: 'Arial', size: 30, bold: true, color: NAVY })]
});
const H2 = t => new Paragraph({
  spacing: { before: 260, after: 110 }, keepNext: true,
  children: [new TextRun({ text: t, font: 'Arial', size: 24, bold: true, color: BLUE })]
});

/* A spoken line: large, generous leading, the way a teleprompter sets it. */
const SAY = (t, o = {}) => new Paragraph({
  spacing: { after: o.after === undefined ? 160 : o.after, line: 300 },
  indent: { left: 360 },
  children: [new TextRun({ text: t, size: 24, color: '1A1A1A', bold: o.b, italics: o.i })]
});

/* A stage direction: smaller, grey, bracketed, never read aloud. */
const DO = t => new Paragraph({
  spacing: { before: 120, after: 120, line: 260 },
  indent: { left: 360 },
  shading: { type: ShadingType.CLEAR, fill: 'F4F7FA', color: 'auto' },
  border: { left: { style: BorderStyle.SINGLE, size: 12, color: BLUE, space: 6 } },
  children: [new TextRun({ text: t, size: 18, italics: true, color: GREY })]
});

/* The running clock down the left margin. */
const CUE = (time, label) => new Paragraph({
  spacing: { before: 300, after: 80 }, keepNext: true,
  children: [
    new TextRun({ text: time, font: 'Consolas', size: 20, bold: true, color: 'A31515' }),
    new TextRun({ text: '   ' + label, font: 'Arial', size: 20, bold: true, color: NAVY })
  ]
});

const BEAT = t => new Paragraph({
  spacing: { before: 60, after: 160 }, indent: { left: 360 },
  children: [new TextRun({ text: t, size: 20, bold: true, color: 'B07400' })]
});

const BULLET = (parts) => new Paragraph({
  numbering: { reference: 'bullets', level: 0 },
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
  shading: { type: ShadingType.CLEAR, fill: color === 'C0504D' ? 'FDF0EF' : 'FFF8E8', color: 'auto' },
  border: { left: { style: BorderStyle.SINGLE, size: 18, color: color || 'F0AD4E', space: 8 } },
  indent: { left: 180, right: 120 },
  children: [new TextRun({ text: title + '  ', bold: true, color: color === 'C0504D' ? '9C2B28' : '8A6100' })]
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

  new Paragraph({
    spacing: { after: 60 },
    children: [new TextRun({ text: 'DMW Claims Management Solution', font: 'Arial', size: 36, bold: true, color: NAVY })]
  }),
  new Paragraph({
    spacing: { after: 40 },
    children: [new TextRun({ text: 'Three-Minute Demo Script', font: 'Arial', size: 26, color: BLUE })]
  }),
  RP([['dmw_claims_management_v1_3Beta.html', { code: true }], ['   ·   898 CMS public comments → sentiment, summary, categories, executive briefing']],
     { color: GREY, size: 19 }),

  CALLOUT('How to use this.', [['Everything in the large type is spoken. Everything in a '], ['grey box', { b: true }], [' is something you do and never read aloud. The timings are cumulative. Pause where it says to pause — the silence while the pipeline runs is the most persuasive part of the whole demo, and talking over it wastes it.']], 'F0AD4E'),

  RULE(),
  H1('The script'),

  CUE('0:00', 'OPEN — one number on the screen, nothing else'),
  DO('Slide: 898, set as large as the projector will allow. No logo, no title, no bullet points. Let it sit for three full seconds before you say anything.'),
  SAY('Eight hundred and ninety-eight.'),
  SAY('That is how many people wrote to CMS about a single Medicare rule.'),
  SAY('Not a survey. Not a focus group. Eight hundred and ninety-eight people who sat down and typed what this rule would do to them. Providers. Hospitals. Patients.'),
  SAY('One and a half million characters. The better part of a thousand pages.'),
  BEAT('[pause]'),
  SAY('And somewhere in those thousand pages is the one thing you need to know before you decide.'),

  CUE('0:40', 'THE PROBLEM'),
  SAY('How long does that take today?'),
  SAY('Three weeks. Four people, reading. And by the time they finish, the comment window has closed and the decision is already made.'),
  DO('Substitute your own baseline here. Three weeks and four people is a placeholder — use the real number from your last docket, and say where it comes from. A figure the room can check is worth more than a dramatic one it cannot.'),
  SAY('Leadership only ever asks three questions about a docket like this. What do people feel. What did they say. What are they talking about.'),

  BEAT('[pause]'),
  SAY('Feel. Say. About. That is the whole job.'),
  SAY('Watch.'),

  CUE('1:05', 'THE RUN — one click, then silence'),
  DO('Home tab. Click Run Pipeline. Then stop talking. Put your hands behind your back and watch the screen with the audience. The eight cards light up left to right, with the arrows turning green between them. It takes about four seconds. Say nothing for all four.'),
  SAY('Done.'),
  BEAT('[pause]'),
  SAY('Eight stages. Eight hundred and ninety-eight comments. Three point nine seconds.'),
  SAY('Three weeks became four seconds.'),

  CUE('1:25', 'QUESTION ONE — what do they feel'),
  DO('Click the doSentimentAnalysis card. The sentiment chart opens.'),
  SAY('Question one. What do they feel.'),
  SAY('A hundred and eighty-six positive. Six hundred and seventy-six neutral. Thirty-six negative.'),
  BEAT('[pause]'),
  SAY('Most people would put that chart on a slide and move on. Do not. Neutral is not an opinion — it is a signal that the real story is somewhere else.'),
  DO('Close the panel.'),

  CUE('1:50', 'QUESTION TWO — what did they say'),
  DO('Click the summarizeText card. The comment-against-summary table opens. Scroll it once, slowly, so the room sees real comments going by.'),
  SAY('Question two. What did they actually say.'),
  SAY('All eight hundred and ninety-eight, summarized. A million and a half characters down to two hundred and eighty-four thousand.'),
  SAY('Eighty-two percent shorter. And nothing invented — every word came out of the original comment.'),
  DO('Close the panel.'),

  CUE('2:10', 'QUESTION THREE — what is it about'),
  DO('Click the categorizeText card. The domain breakdown and word cloud open.'),
  SAY('Question three. What is it about.'),
  SAY('And this is where it stops being a demo and starts being a decision.'),
  BEAT('[pause]'),
  SAY('Forty-nine point four percent. Four hundred and forty-four comments. One theme: Provider Administrative Burden.'),
  SAY('Half the response to this rule is about one thing. Paperwork.'),
  SAY('That is not a chart. That is a staffing decision — which division lead gets a phone call on Monday morning.'),
  DO('Close the panel.'),

  CUE('2:35', 'THE BRIEFING'),
  DO('Click the createExecutiveBriefing card. Scroll the preview from the top: TO / FROM / SUBJECT, the dashboard, the thematic table, the twelve-week plan, the risk index.'),
  SAY('And then it writes the memo.'),
  SAY('To. From. Subject. The dashboard. The thematic table. A twelve-week plan. The risks, named.'),
  DO('Click Export briefing. The .docx lands in the downloads bar. Open it in Word, on screen, in front of them. This is the moment the demo is actually selling — do not rush it.'),
  SAY('A Word document. On your desktop.'),
  BEAT('[pause]'),
  SAY('Ready to send.'),

  CUE('2:50', 'ONE MORE THING — the honest number'),
  DO('Scroll back to the thematic table and put a finger on row two.'),
  SAY('One more thing.'),
  SAY('Look at the second row. Thirty-seven point six percent. Uncategorized.'),
  BEAT('[pause]'),
  SAY('That is the system telling you what it does not know.'),
  SAY('Another tool would have forced those three hundred and thirty-eight into a bucket and handed you a clean pie chart. This one refuses. And that is exactly where you put your experts on Monday.'),

  CUE('3:00', 'CLOSE'),
  SAY('Eight hundred and ninety-eight voices. Four seconds. One memo.'),
  SAY('Not to replace the judgment. To get it in front of the people who have it — while the decision is still open.'),
  DO('Stop. Do not add a summary slide. Do not say "any questions". Let it land.'),

  RULE(),
  H1('Delivery notes'),

  H2('The one rule'),
  RP([['Do not narrate the pipeline while it runs. Four seconds of silence in front of an audience feels like twenty to '], ['you', { i: true }], [' and like a magic trick to '], ['them', { i: true }], ['. It is the only part of this demo that cannot be reproduced on a slide, and filling it with words throws it away.']]),

  H2('Pace'),
  P('427 spoken words. At 145 words a minute that is two minutes fifty-seven of speech; add the four-second silent run and the five marked pauses and you land at about three minutes ten. That is the honest number — if you have a hard three-minute slot, cut the second and third lines of the sentiment section, which are the only lines in the script that restate something already on the screen. Never cut the Uncategorized section to save time; it is the most credible thirty seconds in the demo.'),

  H2('Rehearse the run'),
  RP([['Open the page and run the pipeline once before the audience arrives. It is a static page with no database and no network calls, so it cannot fail on a bad connection — but the first run of a session reads the data files, and you want that already warm. Serve it over http, never '], ['file://', { code: true }], [', or nothing loads at all.']]),

  H2('If someone asks whether this is really AI'),
  P('Answer honestly, and the answer is strong enough:'),
  BULLET([['The executive narrative in the briefing is genuine generative AI — Snowflake Cortex Complete wrote it during the pipeline run that produced this data, and the page reproduces it verbatim.']]),
  BULLET([['Sentiment and categorisation are keyword rules, not a trained model. They are fast, transparent and auditable, and they are why 37.6% comes back Uncategorized instead of being quietly mislabelled.']]),
  BULLET([['The summaries in the browser are extractive — real sentences lifted from the comment. In Snowflake the same stage calls Cortex Summarize, which writes new ones.']]),
  CALLOUT('Do not overclaim this in the room.', [['The strongest version of this demo is the honest one. "These are keyword rules, and here is the share they could not classify" is a more credible claim than "the AI read everything", and it survives the follow-up question. The slide that oversells is the one that loses the budget at the second meeting.']], 'F0AD4E'),

  RULE(),
  H1('Every number in the script'),
  P('All of these come from a live run of v1.3Beta on the CMS-2022-0190 docket and will match what the audience sees on screen — with one exception, marked below.', { after: 120 }),
  TBL(['Said in the script', 'Where it comes from'], [
    ['898 comments', 'Comments pill, and the Load Comments card'],
    ['1.5 million characters → 284,000', '`summarizeText` panel: 1,578K original, 284K summary'],
    ['82% shorter', '`summarizeText` panel: 82.0% character reduction (83.2% by word)'],
    ['186 / 676 / 36', 'Briefing, Overall Sentiment Distribution'],
    ['3.9 seconds', 'Elapsed pill, 3,885 ms on a cold run'],
    ['49.4%, 444 comments', 'Briefing, thematic table row 1 — Provider Administrative Burden'],
    ['37.6%, 338 comments', 'Briefing, thematic table row 2 — Uncategorized'],
    ['**Three weeks, four people**', '**Not measured. A placeholder.** Replace it with your own baseline before you present, or drop the claim.']
  ], [3600, 6480]),
  P(''),
  CALLOUT('The timings move a little.', [['The elapsed figure varies by a few hundred milliseconds between runs and between machines. Say "under four seconds" rather than a precise number if you would rather not be caught out by a slow laptop. Every other figure above is deterministic — the same docket always produces the same counts.']], 'F0AD4E')
];

const doc = new Document({
  creator: 'DMW — Data Mart Workbench',
  title: 'DMW Claims Management Solution — Three-Minute Demo Script',
  description: 'Presenter script for dmw_claims_management_v1_3Beta.html',
  numbering: {
    config: [{
      reference: 'bullets',
      levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
                 style: { paragraph: { indent: { left: 460, hanging: 240 } } } }]
    }]
  },
  styles: { default: { document: { run: { font: 'Calibri', size: 21, color: '262626' } } } },
  sections: [{
    properties: { page: { margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } } },
    children
  }]
});

Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log('wrote', OUT, b.length, 'bytes'); });
