# DMW_UNIT

The DMW (Data Mart Workbench) IT Management Solution, in two versions plus the
documents that sit alongside them.

| Folder | What it is | Status |
|---|---|---|
| [`v2/`](v2) | Current. Ten tabs: Home SLA overview, Data Grid with column metadata and an editable `JobName`, Chart, Visualize Performance, Security, Ask DMW. | **Active** |
| [`v1/`](v1) | The original single-page app, plus the Job Notes TEST build layered on it. | Superseded, still deployable |
| [`docs/`](docs) | Planning and communication documents. Not tied to either version. | Current |

Each version folder is **self-contained and independently deployable** — its own
`scripts/`, `styles/` and configuration. Copy the folder to a web server as a
unit; nothing reaches outside it.

## v2 — start here

Read [`v2/README_V2.md`](v2/README_V2.md) (or `README_V2.docx`).

```
v2/
├── dmw_it_management_v2.html   the application
├── column_metadata.csv         34 column definitions — must sit beside the HTML
├── dmw_update.sas              stored process: the only thing that writes
├── data/                       CSVs the Home and Visualize Performance tabs read
├── ps1/                        four PowerShell scanners that produce them
├── scripts/  styles/           AG Grid, vendored locally
└── README_V2.md / .docx
```

## v1 — the original, and the Job Notes test build

Read [`v1/README.md`](v1/README.md) for the application and
[`v1/README_TEST_NOTES.md`](v1/README_TEST_NOTES.md) for the notes feature.

```
v1/
├── dmw_it_management.html          the original six-tab page
├── dmw_run.sas                     read-only middleware, SELECT-only guard
├── dmw_it_management_test.html     adds editable Job Notes
├── dmw_test_*.sas / .sql           the notes table and its stored processes
├── scripts/  styles/               AG Grid, vendored locally
└── README.md / .docx, README_TEST_NOTES.md / .docx
```

`v1` is kept rather than deleted because it is a working, deployed page and the
Job Notes work on top of it was never folded into v2. Nothing in `v2` depends on
anything in `v1`; the two share no files.

## What carried forward into v2, and what did not

- `dmw_run.sas` is **unchanged** in both versions and still holds the
  `SELECT`-only guard. v2 does not modify it.
- The v1 Job Notes stored processes (`dmw_test_note_*.sas`) are **not** used by
  v2. v2's only write path is `dmw_update.sas`, which edits `JobName`.
- Both versions connect through the same SAS Stored Process bridge, as the
  service account `A70ADMED\SVC_CGSUSASSRV`.
