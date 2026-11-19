# DMW_UNIT

The DMW (Data Mart Workbench) solutions, grouped by what they manage.

| Folder | Solution | Status |
|---|---|---|
| [`IT/`](IT) | **IT Management** — SAS scheduler monitoring, server SLA and pipeline performance, over SQL Server via a SAS Stored Process. | **Active** |
| [`claims/`](claims) | **Claims Management** — a CMS public-comment pipeline: sentiment, NLP, summarisation, categorisation and an executive briefing. Reads static files, no database. | **Active** |
| [`provider/`](provider) | **Provider Risk Agent** — a browser mockup of the Streamlit-in-Snowflake Cortex Analyst app, for demoing where Snowflake is unreachable. | **Demo only** |
| [`docs/`](docs) | Planning and communication documents. Not tied to a solution. | Current |

The solutions are independent. They share no files and none imports from
another; each folder is self-contained and separately deployable.

---

## IT Management — `IT/`

```
IT/
├── v2/     Current. Ten tabs, latest build dmw_it_management_v1_8Beta.html
└── v1/     The original page plus the Job Notes TEST build layered on it
```

**Start with [`IT/v2/README_V2.md`](IT/v2/README_V2.md)** (or `README_V2.docx`).

v2 is the active build. Home SLA overview, Data Grid with column metadata and an
editable `JobName`, Scheduler Overview, Server Performance, Data Pipeline
Performance, Security, Ask DMW Agent. Each beta is kept beside its predecessor
rather than overwritten, so `v1_8Beta` sits next to `v1_7Beta` and so on.

v1 is superseded but still deployable, and is kept rather than deleted because
it is a working page and the Job Notes work layered on it was never folded into
v2. See [`IT/v1/README.md`](IT/v1/README.md) and
[`IT/v1/README_TEST_NOTES.md`](IT/v1/README_TEST_NOTES.md).

### What carried forward into v2, and what did not

- `dmw_run.sas` is **unchanged** in both versions and still holds the
  `SELECT`-only guard. v2 does not modify it.
- The v1 Job Notes stored processes (`dmw_test_note_*.sas`) are **not** used by
  v2. v2's only write path is `dmw_update.sas`, which edits `JobName`.
- Both connect through the same SAS Stored Process bridge, as the service
  account `A70ADMED\SVC_CGSUSASSRV`.

---

## Claims Management — `claims/`

```
claims/
├── dmw_claims_management_v1_7Beta.html   the application (current)
├── dmw_claims_management_v1_6Beta.html   the previous build, kept beside it
├── dmw_claims_management_v1_5Beta.html   and the ones before it
├── dmw_claims_management_v1_4Beta.html
├── dmw_claims_management_v1_3Beta.html
├── dmw_claims_management_v1_2Beta.html
├── dmw_claims_management_v1_1Beta.html
├── dmw_claims_management_v1_0Beta.html
├── column_metadata.csv                   13 column definitions
├── data/                                 the two comment files it reads
├── snowflake/__init__.py                 the cgs_ai package that ran in ONEPI
├── docs/                                 the README diagram and the demo script
└── README.docx
```

**Start with `claims/README.docx`.** There is deliberately no `README.md` in
`claims/`: v1.1Beta ships documentation as Word only. The cost is that GitHub
cannot render a landing page for that folder — download the `.docx` instead.

A Pipeline Runner Workflow: eight stage cards in two rows that light up left to
right as the pipeline runs, each opening its own output when clicked.

**It connects to no database.** Where the IT solution reaches SQL Server through
a SAS Stored Process, this one reads two static files and does every enrichment
in the browser — so it runs from any web server, with no SAS session and no
Snowflake behind it.

`snowflake/__init__.py` is the `cgs_ai` package that produced the enriched data,
and it **runs in Snowflake ONEPI only**. The browser reimplements the
deterministic stages so the workflow is demonstrable; it does not reimplement
Cortex. The README is explicit about where the two differ.

---

## Provider Risk Agent — `provider/`

```
provider/
├── PROVIDER_RISK_AGENT_DEMO_V2_9.html   the mockup (current)
├── PROVIDER_RISK_AGENT_DEMO_V2_8.html   the previous build, kept beside it
├── PROVIDER_RISK_AGENT_DEMO_V2_7.html   and the ones before it
├── PROVIDER_RISK_AGENT_DEMO_V2_6.html
├── PROVIDER_AGENT_DEMO_V2_5.html
└── FFS_PROVIDER_CHAT_HJS_V2_5.py        the Streamlit app it mocks up
```

A browser stand-in for the Streamlit-in-Snowflake Cortex Analyst app, built
because the demo site cannot reach Snowflake. **It makes no connection and runs
no query** — it replays a recorded session.

What is real: the chrome and CSS are the same markup `build_chat_html()`
produces, the four chart renderers are lifted verbatim from the `CHART_JS`
string in the `.py`, and the questions, interpretations, generated SQL and
result rows are the ones from the recorded Cortex Analyst session.

What is not: the connection splash and the spinner timing are staged, and the
closing narrative on the Illinois question was written from the figures in the
result table — the transcript has no such paragraph.

**Both of those are disclosed on the start-up splash**, which from v2.7 is the
only place they are disclosed. v2.6 and earlier also printed a note under the
Illinois paragraph itself. The splash holds for six seconds, and a click
dismisses it early, so the disclosure is long enough to read rather than a
caption that flashes past.

**v2.9 is the one to present.** It is the app and nothing else: the presenter
clicks About and Help and types each question, and Send or Enter both work.
Five questions are recorded — the top 20 providers, the top 5 in Ohio, the
heatmap by state and risk tier, the same result as a bar chart, and the
Illinois follow-up. Wording is matched loosely, so a trailing full stop or a
lower-case start still lands. Anything else answers
**"The system is currently unavailable. Please try again later."** followed by
the five questions it does know. From v2.8 that is the reply on both paths —
typed and clicked — so an unscripted question in front of an audience reads as
an outage rather than as a seam in the prop. The start-up splash still states
plainly that the page is a mockup, so nothing is concealed; the wording only
moves the disclosure out of the middle of the demo.

v2.5 keeps a demo-script strip across the top that steps through the same
sequence and rings the control each step concerns. It is useful for rehearsing
and wrong for presenting, since an audience expecting a live tool would read
the strip as part of the product.

---

## Serving any of them

All three are static pages and must be served over **http(s)**. A `file://`
page cannot fetch its own CSV, JSON or metadata files, so nothing will load.
Copy a solution folder to a web directory as a unit; nothing reaches outside it.
