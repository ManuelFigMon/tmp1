# DMW — Claims Management Solution v1.0Beta

`dmw_claims_management_v1_0Beta.html`. Runs a CMS public-comment analysis
pipeline end to end and shows what each stage produced.

**It connects to no database.** It reads two static files from `data/` and does
every enrichment in the browser. There is no SAS session, no logon and no
stored process, so it runs from any web server.

---

## 1. Contents

```
claims/
├── dmw_claims_management_v1_0Beta.html   the application
├── column_metadata.csv                   13 column definitions — must sit beside the HTML
├── data/
│   ├── cms_2022_0190_SENT_SUMMARIZE_CAT_v1.csv    the supplied enriched export
│   ├── cms_2022_0190_SENT_SUMMARIZE_CAT_v1.json   JSON version, built from it
│   ├── CMS_LCD_SIJ_SynPUF_Comments.json           the supplied SIJ sample
│   └── CMS_LCD_SIJ_SynPUF_Comments.normalized.json  same, in the shared schema
├── snowflake/
│   └── __init__.py                       the cgs_ai package that ran in ONEPI
├── scripts/  styles/                     AG Grid, vendored locally
└── README.md
```

Serve the folder over **http(s)**. A `file://` page cannot fetch its own data
files, so nothing will load.

---

## 2. The pipeline

Eight stage cards, two rows, lighting up left to right.

| Row | Stages |
|---|---|
| **Analysis pipeline** | Load Comments → `doSentimentAnalysis` → `visualizeNLP` → `summarizeText` |
| **Output pipeline** | `categorizeText` → `formatCSV` → `createExecutiveBriefing` → Output |

Each card shows queued (grey), running (spinner), or done — a green check with
the elapsed time and a count such as *"898 comments summarized"*. Click a
finished card to see what it produced:

| Card | What opens |
|---|---|
| `doSentimentAnalysis` | Sentiment distribution bar chart with counts and shares |
| `visualizeNLP` | Word-count histogram and the 30 most frequent terms |
| `summarizeText` | Scrollable comment-against-summary table, with character and word reduction |
| `categorizeText` | Configurable word cloud plus the category breakdown |
| `formatCSV` | The formatted output table, scrollable |
| `createExecutiveBriefing` | Approval, export format (docx default / pdf), and a preview |

`CONFIG.STAGE_DELAY_MS` gives each card a visible moment in the running state.
The work finishes between frames; a workflow that blinks from grey to green
teaches nobody what it did. Set it to `0` to run at full speed.

---

## 3. What is reimplemented, and what is not

The deterministic stages are ported to JavaScript from `snowflake/__init__.py`
— the same keyword sets, the same scoring — so the workflow is demonstrable
without a Snowflake session.

> **`summarizeText` is the exception, and it matters.** Cortex generates
> **abstractive** text: new sentences. The browser builds an **extractive**
> summary by scoring the comment's own sentences on term frequency and keeping
> the best. Every word in a summary shown here appears in the original comment.
> The two are not interchangeable. Where they differ, the Snowflake output is
> the system of record, and the page says so on the stage panel.

### The briefing narrative is AI-generated

It was produced by Snowflake Cortex `Complete` during the run that enriched the
CMS-2022-0190 file, and is carried here **verbatim** from
`cgs_executive_briefing.docx` — reproduced, not regenerated, because no language
model runs in a browser. The heading reads *AI-Generated Executive Analysis
(Snowflake Cortex)*, which is accurate: that text is Cortex output.

Its figures were reconciled against the pipeline's own before it was embedded:

| Narrative claims | Pipeline computes |
|---|---|
| 898 comments | 898 |
| 76% neutral | 75.3% |
| 49.4% administrative burden | 49.4% |
| 6.3% reimbursement | 6.3% |
| 5.0% care quality | 5.0% |
| 1.3% health equity | 1.3% |

All six match, so the paragraph describes this dataset and not an earlier one.

> **If you replace the data file, re-run the Snowflake briefing and replace the
> narrative with it.** A verbatim paragraph does not update itself, and a stale
> one would quietly describe a different docket while every figure around it
> moved. The narrative lives on the source entry in `CONFIG.SOURCES`, next to a
> comment saying exactly this.

The SIJ sample was never run through Snowflake, so it carries
`cortexNarrative: null`. For that source the briefing falls back to a paragraph
assembled from the computed figures and labels it *Computed Executive Analysis*
instead — the label follows the provenance, in both directions.

### Where the two can disagree

Both sample files already carry `overall_sentiment` and `COMMENT_CATEGORY` from
the Snowflake run. The pipeline recomputes both, **keeps the stored value**, and
writes any disagreement count to the Console tab, with the recomputed label
available in `_sent_computed` / `_cat_computed`. Silently overwriting a column
produced by the system of record would hide exactly the drift worth knowing
about.

---

## 4. The Snowflake package

`snowflake/__init__.py` is the attached `cgs_ai` package, with one change.

**It runs in Snowflake ONEPI only**, and now says so properly. v1.0 guarded each
function with `platform.platform().startswith("Linux")`. Nearly every build
agent, container and CI runner is also Linux, so that check passed in all the
places the code cannot actually run, and the first sign of trouble was an import
error deep inside a function.

`_require_snowflake()` now checks three things in increasing order of cost:

1. the platform is Linux — necessary, not sufficient
2. `snowflake.snowpark` imports — the library is present
3. an active session exists — we are genuinely inside one

On failure it prints a NOTE naming the failed check and **returns None**. It does
not raise: an unattended pipeline step should report and stop, not abort the
notebook.

**`formatCSV` is deliberately unguarded.** It reads a CSV and writes an xlsx,
touching no session and no Cortex function, so it works anywhere `openpyxl` is
installed. Guarding it would block a step that genuinely runs outside ONEPI.

---

## 5. Data

| Field | Source |
|---|---|
| `id` | `commentId` in the CSV export, mapped so both sources share one schema |
| `agencyId`, `postedDate`, `title`, `comment`, `docketId`, `attachmentLinks` | as supplied |
| `attachmentCount` | native in the SIJ JSON; **derived** for the CSV by counting `attachmentLinks` entries |
| `character_count`, `word_count` | recomputed from the cleaned text |
| `overall_sentiment`, `COMMENT_CATEGORY` | from the Snowflake run where present, else computed |
| `Comment_summary` | produced in the browser at run time; not stored in either file |

**Two cleaning steps applied when the JSON was built**, both at the source rather
than at display time:

- HTML entities are decoded, so an apostrophe reads as an apostrophe rather than
  `&rsquo;`.
- **HTML tags are stripped.** The export carries `<br/>` and `<p>` inside the
  comment body. Escaping that for display shows the reader `<br/>` as text, and
  the summariser treats it as a word — it was appearing at the front of
  summaries before this was fixed. `character_count` and `word_count` are
  recomputed afterwards so they describe the text actually shown.

---

## 6. Reading the enrichment columns honestly

`overall_sentiment` and `COMMENT_CATEGORY` are **keyword rules, not models**.
They read word presence, not meaning, so *"I do not support this"* scores as
positive. Use them to decide which comments to read, never as measured opinion.
Both stage panels say this on screen.

`Uncategorized` is the **second largest group** in the supplied CSV at roughly
36% of comments. That is a finding about how well the five CMS domains fit this
docket, not a gap to be hidden — a breakdown that drops it overstates coverage.
The category panel calls the share out explicitly.

---

## 7. Exports

| Output | Notes |
|---|---|
| `cms_comments_enriched.csv` | the full enriched table |
| `cms_comments_enriched.json` | the same, as JSON |
| `cgs_executive_briefing.md` | the briefing as Markdown |
| `cgs_executive_briefing.docx` | from the briefing card |

> A browser cannot write a true `.docx` without a library. The export is
> Word-compatible HTML under a `.docx` name, which Word opens and re-saves as a
> real document. The Snowflake `createExecutiveBriefing` writes a native file via
> `python-docx`. The Console says which you have; they are not the same artefact.

The approval tick-box **stamps the exported document and nothing else**. There is
no workflow, no second approver and no audit record behind it. The panel says so,
because an approval control that looks like governance and isn't is worse than
none.

---

## 8. What was tested

Exercised headless over http, no console or page errors:

- All eight stages run in order and reach done; mid-run the cards read
  `done, running, queued…` as expected
- 898 comments load, 13 columns formatted, 5,112 distinct terms, 6 domains
- All six output panels open and populate: 3 sentiment bars, 250-row comparison
  table at 82.0% character reduction, 60 cloud words reconfiguring to 25, a
  12-column formatted table, and the briefing with `docx` default
- The briefing shows the verbatim Cortex narrative under an *AI-Generated*
  heading for CMS-2022-0190, and the computed fallback under a *Computed*
  heading for the SIJ sample
- Approval stamps the briefing; export downloads `cgs_executive_briefing.docx`
- Grid shows 24 columns with 12 live metadata icons, none greyed; the panel
  opens with all six fields
- The second source runs clean: `CMS-2026-0042`, 24 comments, 12 columns

**Not tested:** `snowflake/__init__.py` against a real ONEPI session. There is no
Snowflake in this environment, so the guard's blocked path was verified but its
success path was not.
