# DMW_UNIT

The DMW (Data Mart Workbench) solutions, grouped by what they manage.

| Folder | Solution | Status |
|---|---|---|
| [`IT/`](IT) | **IT Management** — SAS scheduler monitoring, server SLA and pipeline performance, over SQL Server via a SAS Stored Process. | **Active** |
| [`claims/`](claims) | **Claims Management** — a CMS public-comment pipeline: sentiment, NLP, summarisation, categorisation and an executive briefing. Reads static files, no database. | **Active** |
| [`docs/`](docs) | Planning and communication documents. Not tied to a solution. | Current |

The two solutions are independent. They share no files and neither imports from
the other; each folder is self-contained and separately deployable.

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
├── dmw_claims_management_v1_1Beta.html   the application (current)
├── dmw_claims_management_v1_0Beta.html   the previous build, kept beside it
├── column_metadata.csv                   13 column definitions
├── data/                                 the two comment files it reads
├── snowflake/__init__.py                 the cgs_ai package that ran in ONEPI
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

## Serving any of them

All three are static pages and must be served over **http(s)**. A `file://`
page cannot fetch its own CSV, JSON or metadata files, so nothing will load.
Copy a solution folder to a web directory as a unit; nothing reaches outside it.
