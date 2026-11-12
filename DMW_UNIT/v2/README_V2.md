# DMW — IT Management Solution v2

`dmw_it_management_v2.html`. Adds a Home SLA overview, a report-performance
Visualize Performance tab, a Security diagram, an embedded Ask DMW agent, column
metadata with an "i" icon on every column, and an editable `JobName`.

The Data Grid query now joins `View_Scheduler_PBI_Dashboard` to `Job` on
`JobID`, so all 34 scheduler variables are available in one row.

---

## 1. What is in the folder

Keep this structure. Every path below is relative to the HTML file.

```
v2/
├── dmw_it_management_v2.html     the application
├── column_metadata.csv           34 column definitions — MUST sit beside the HTML
├── dmw_update.sas        stored process: the only thing that writes
├── scripts/ag-grid-community.min.js
├── styles/ag-grid.css, ag-theme-alpine.css
├── data/                         CSVs the Home and Visualize Performance tabs read
│   ├── sas_server_status.csv         <- Check-SASServers_v2.0.ps1
│   ├── objectspawner_status.csv      <- Check-ObjectSpawner_v1.0.ps1
│   ├── sas_log_rollup.csv            <- Export-SasLogRollup.ps1
│   └── sas_log_steps.csv             <- Export-SasLogSteps.ps1 (cgs_ai)
└── ps1/
    ├── Check-SASServers_v2.0.ps1
    ├── Check-ObjectSpawner_v1.0.ps1
    ├── Export-SasLogRollup.ps1       one row per LOG FILE
    └── Export-SasLogSteps.ps1        one row per SAS STEP (wraps cgs_ai)
```

**The three CSVs in `data/` are SAMPLE DATA.** They were generated so the
dashboard has something to render before you run anything. Replace them by
running the scripts. See §4.

---

## 2. Step-by-step installation

### Step 1 — Put the folder on a web server

Copy the whole `v2` folder to a web directory, ideally the same server that
hosts SAS so the session cookie is same-origin.

> The page **must** be served over http(s). From `file://` the browser blocks
> the `fetch` of `column_metadata.csv` and the `data/*.csv` files, so the
> metadata icons grey out and the Home tab shows nothing. This is a browser
> origin rule, not a defect.

To test locally before deploying:

```
cd v2
python -m http.server 8080
```

then open `http://localhost:8080/dmw_it_management_v2.html`.

### Step 2 — Check the configuration block

Everything adjustable is in `CONFIG` at the very top of the HTML, in eleven
commented sections. At minimum confirm:

| Key | Default | Note |
|---|---|---|
| `SAS_BASE_URL` | `https://cgssaswebu.a70admed.com/SASStoredProcess/do?_program=` | must end in `_program=` |
| `SAS_PROGRAM` | `/User Folders/731o/My Folder/dmw_run` | the read stored process |
| `METADATA_CSV_PATH` | `column_metadata.csv` | relative to the HTML |
| `SERVER_STATUS_CSV` | `data/sas_server_status.csv` | |
| `SERVICE_STATUS_CSV` | `data/objectspawner_status.csv` | |
| `SAS_LOG_ROLLUP_CSV` | `data/sas_log_rollup.csv` | per-file metrics |
| `SAS_LOG_STEPS_CSV` | `data/sas_log_steps.csv` | per-step metrics, optional |
| `EDITABLE_COLUMNS.JobName.program` | `/User Folders/731o/My Folder/dmw_update` | the write stored process |
| `ASK_DMW_URL` | the Copilot Studio embed URL | |

After editing, press **Clear Cache & Reload** in the header.

### Step 3 — Register the update stored process

Only needed if you want `JobName` to be editable. Skip it and everything else
still works; the pencil column simply fails on save.

1. Copy `dmw_update.sas` to the stored process source folder.
2. SMC → New Stored Process, name `dmw_update`.
3. **Result type = Streaming.** Required.
4. Input parameters, all Text, none required at prompt:
   `job_id`, `new_name`, `old_name`, `debug_mode`.
5. Ask the DBA for column-level UPDATE:

```sql
GRANT UPDATE ON [dbo].[Job] ([JobName]) TO [A70ADMED\SVC_CGSUSASSRV];
```

Column-level is deliberate: this one field can change and nothing else on the
row, so a defect here cannot alter a schedule.

### Step 4 — Schedule the PowerShell scanners

Each writes a CSV into `data/`. All are read-only against the servers
they inspect.

| Script | Writes | Suggested schedule |
|---|---|---|
| `Check-SASServers_v2.0.ps1` | `sas_server_status.csv` | every 5 minutes |
| `Check-ObjectSpawner_v1.0.ps1` | `objectspawner_status.csv` | every 5 minutes |
| `Export-SasLogRollup.ps1 -metric_profile sas_log` | `sas_log_rollup.csv` | nightly |
| `Export-SasLogSteps.ps1` | `sas_log_steps.csv` | nightly |

Try each by hand first:

```powershell
.\Check-SASServers_v2.0.ps1   -TestMode
.\Check-ObjectSpawner_v1.0.ps1 -TestMode
.\Export-SasLogRollup.ps1 -metric_profile sas_log -TestMode -MaxFiles 50
.\Export-SasLogSteps.ps1   -TestMode
```

`-TestMode` prints to the console and writes nothing, so you can confirm the
hostnames and paths before any CSV is touched.

Then as a scheduled task:

```
Program : powershell.exe
Args    : -NonInteractive -ExecutionPolicy Bypass -File "<path>\Check-SASServers_v2.0.ps1"
Trigger : Repeat every 5 minutes
Run As  : an account with write access to data\ and log\
```

> **UNC paths.** PowerShell blocks unsigned scripts loaded from a network
> share. Either `Unblock-File -Path "<unc path>"` once, or always invoke with
> `-ExecutionPolicy Bypass`, or copy to a local drive first.

Edit the `CONFIGURATION` region at the top of each script — it is the only
part an operator should need to touch. In particular check the hostnames in
`Check-ObjectSpawner_v1.0.ps1`; they are a guess at your naming convention.

### Step 5 — Maintain the metadata CSV

`column_metadata.csv` is an external file, deliberately. Data stewards edit it
without touching HTML. Columns, matching
`query_builder_with_metadata_csv_v1_21.html` exactly:

```
LOB, PHYSICAL_COLUMN_NAME, LOGICAL_COLUMN_NAME, TABLE_COLUMN_NAME,
DESCRIPTION, LONG_DESCRIPTION, LOAD_FREQUENCY
```

Rows match grid columns on `PHYSICAL_COLUMN_NAME`, case-insensitively. A
column with no row gets a greyed-out "i" icon rather than an error.

---

## 3. The tabs

| Tab | What it shows |
|---|---|
| **Home** | Home SLA Overview: server monitor, services monitor with `OverallStatus`, uptime report, services reliability. Works with no SAS session. |
| **Data Grid** | The joined 34-column result. "i" icon on every column, pencil on `JobName`. |
| **SQL** | Edit and re-run the statement. `SELECT` only. |
| **Dashboard** | KPI tiles and a group-by summary of the current result set. |
| **Chart** | Bar chart of the current result set, built from the hamburger **Chart Data** item. |
| **Visualize Performance** | Report performance from **two sources**: per-file rollup (program count, run count, CPU by program, CPU over time, ranked detail) and per-step detail from the cgs_ai scanner (CPU by step type, slowest individual steps). |
| **Security** | Software delivery and hotfix cycle, drawn inline as SVG. |
| **Ask DMW** | Embedded Copilot Studio agent. |
| **Connection** | Effective settings and the SAS OLE DB snippet. |
| **Console** | Request URL, status, timing, errors. Start here when something fails. |

> **The two chart tabs are distinct.** *Chart* plots whatever is currently in
> the Data Grid. *Visualize Performance* plots SAS log performance from the two
> scanners in `ps1\` and is independent of the grid.

### Home tab specifics

- **Server monitor** — one card per endpoint per environment, from the most
  recent sweep. `UP` / `DEGRADED` / `DOWN`. A reachable but slow endpoint is
  DEGRADED, so you see amber before anything breaks.
- **Services monitor** — the `OverallStatus` per environment is the headline:
  `HEALTHY`, `DEGRADED`, `CRITICAL` or `UNKNOWN`. It is the worst state among
  the monitored services, where the Object Spawner and Metadata Server are
  marked Critical and a stopped Connect Spawner only degrades.
- **Uptime report** — availability over time, with red dots on failed checks
  and an availability percentage per endpoint.
- **Services reliability** — the same shape for services.
- **Login prompt, upper right** — the page states what signing in adds, and
  offers the Logon button.
- A sweep older than `STATUS_STALE_MINUTES` raises a **STALE** badge, so an
  all-green board from a scanner that died three days ago cannot mislead.

---

## 3.1 The two SAS-log sources

The Visualize Performance tab reads two CSVs with **different grains**. They are
complementary — neither derives the other.

| | `Export-SasLogRollup.ps1` | `Export-SasLogSteps.ps1` (wraps cgs_ai) |
|---|---|---|
| Grain | one row per **log file** | one row per **SAS step** |
| Columns | 22: session totals, memory, step counts, error/warning counts, top PROC, librefs | 6, in the cgs_ai names: `FullPath, ProgramName, StepIndex, StepLabel, RealTimeSec, CpuTimeSec` |
| Answers | which programs cost most, which are failing | where the time goes inside a program |
| Required? | yes — the tab needs it | optional — the tab degrades cleanly without it |

**`Export-SasLogSteps.ps1` does not re-implement the parse.** It runs
`cgs_ai\src\ps\scanFileSystem.ps1 -metric_profile sas_log`, which is the single
source of truth for step timings and has a Python twin, then flattens its
`Metrics` sheet to CSV.

Two constraints make the wrapper necessary rather than optional:

- The cgs_ai script writes metrics to an **XLSX sheet**. Its own code is
  explicit: with an output path that does not end in `.xlsx` it writes only the
  match rows and the metric rows are **silently dropped**. Asking it for a
  `.csv` produces a file with no timings in it.
- A browser cannot read `.xlsx`.

The wrapper needs nothing installed: it uses `ImportExcel` when present, and
otherwise unzips the workbook and parses the sheet XML with built-in .NET.

Point it at your checkout with `-ScanScript`. The cgs_ai script dot-sources
`cgsUtils.ps1` from its own folder, so it must be run where it lives — copying
it alone to a share will fail.

> **Why `Export-SasLogRollup.ps1` was renamed.** It previously shipped as
> `scanFileSystem.ps1`, which collided with the cgs_ai tool of that name while
> producing an incompatible output contract. The name is corrected and both now
> coexist. Two differences in the rollup are deliberate: it reads
> **Windows-1252** (SAS on Windows writes ANSI; the cgs_ai fallback to
> ISO-8859-1 mangles smart quotes and em dashes), and it strips **trailing**
> schedule/version suffixes so repeated runs collapse to one program name,
> where cgs_ai strips a leading keyword token.

## 4. Replacing the sample data

The CSVs in `data/` are synthetic. They contain deliberately scripted
incidents — three endpoint outages and three service outages over 14 days —
so the charts show something other than a flat line. Delete them and run the
scanners for real data; the schemas match exactly, so nothing else changes.

---

## 5. Design decisions

**The write path is separate from the read path.** `dmw_run.sas` keeps its
`SELECT`-only guard and is untouched. `dmw_update.sas` can change one
column on one row and accepts no SQL.

**The grid shows `View_Scheduler_PBI_Dashboard.JobName`, but the update writes
`[dbo].[Job]`.** A view is usually not updatable, and the job name lives on
the base table. The view reflects the change on the next read.

**The author is never sent by the browser.** All writes arrive as
`SVC_CGSUSASSRV`, so SQL Server cannot say who renamed a job. The stored
process reads `_METAPERSON` server-side. A browser-supplied username can be
forged.

**The new name never becomes SQL.** It is carried through `SYMGET` into a SAS
dataset and applied through an OLE DB libname, so the driver parameterises the
value. This avoids both SQL injection and the SAS trap where a name containing
`&` or `%` is macro-resolved into garbage.

**Optimistic concurrency.** The editor sends `old_name`. If someone else
renamed the job since the grid loaded, the save is refused with `CONFLICT`
rather than overwriting their change.

**Writes POST, reads GET.** A name in a query string is written verbatim into
the web server log, and a browser refresh can silently re-submit a GET.

**Charts are inline SVG.** No charting library, so the page stays air-gapped
and adds no external asset.

**Ask DMW is an iframe, not an API integration.** No API key, no token
handling, no CORS configuration in the page. Note that an embeddable agent
requires **Copilot Studio** — the free agent builder bundled with M365 Copilot
cannot be embedded, it only runs inside the Microsoft 365 Copilot app.

---

## 6. Troubleshooting

| Symptom | Cause |
|---|---|
| All "i" icons grey | `column_metadata.csv` was not fetched. Check the Console tab. Serving from `file://` always fails. |
| Home tab empty, "Monitoring data unavailable" | The `data/*.csv` files are missing or the paths in `CONFIG` are wrong. |
| Home tab shows **STALE** | The scheduled task is not running, or cannot write to `data/`. |
| Ask DMW panel blank | The agent's host refuses to be framed (`X-Frame-Options` / `frame-ancestors`), or you are not signed in to the tenant. Open the URL in a normal tab to confirm. |
| Saving a job name returns `CONFLICT` | Someone else renamed it. Reload the grid. |
| Saving returns `FAILED` | The service account lacks `UPDATE` on `[dbo].[Job].[JobName]`. |
| `ERROR 180-322` in the SAS log | An OLE DB init string was built with `%LET`. Fixed in the current files. |

---

## 7. What was tested, and what was not

Exercised headless over http, zero console or page errors:

- All ten tabs render; the tab bar reads
  Home | Data Grid | SQL | Dashboard | Chart | Visualize Performance | Security | Ask DMW | Connection | Console
- 34 metadata definitions load from the external CSV; 12 visible columns show a
  live "i" icon, none greyed
- The metadata panel opens with all six fields populated
- One pencil icon, on `JobName`; the cell opens the editor
- The save is a **POST**, `application/x-www-form-urlencoded`, carrying
  `job_id`, `new_name`, `old_name` — and **no username field**
- `DME O'Brien 50% & <b>new</b>` survives encoding intact
- Home renders 4 server cards, 2 `OverallStatus` cards, the service table, and
  two time-series charts; availability computes to 99.08–100% against the
  sample data
- Visualize Performance renders both sources: 16 programs / 864 runs / 74.7 h CPU /
  41 error runs from the rollup, and 348 step rows / 12 step types / 1.7 h CPU from
  the cgs_ai source, with a 15-row slowest-steps table
- With `sas_log_steps.csv` returning 404 the tab still renders the rollup half and
  explains how to produce the other
- Security renders as one SVG with no overlapping elements
- Ask DMW renders the iframe with the configured URL

**Not tested here** — there is no SAS, SQL Server, or Windows in the build
environment. Unverified: the real `UPDATE`, `_METAPERSON` resolution, the join
against the live `Job` table, all four PowerShell scripts against real hosts, including the cgs_ai wrapper,
and whether the Copilot agent permits framing from your origin.
