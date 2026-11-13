# DMW — IT Management Solution v1.5Beta

`dmw_it_management_v1_5Beta.html` (version **1.5Beta**, shown in the page header
beside the database name). Adds a Home SLA overview, a report-performance
Data Pipeline Performance tab, a Security diagram, an embedded Ask DMW agent, column
metadata with an "i" icon on every column, and an editable `JobName`.

The Data Grid query now joins `View_Scheduler_PBI_Dashboard` to `Job` on
`JobID`, so all 34 scheduler variables are available in one row.

---

## 1. What is in the folder

Keep this structure. Every path below is relative to the HTML file.

```
v2/
├── dmw_it_management_v1_5Beta.html     the application
├── column_metadata.csv           34 column definitions — MUST sit beside the HTML
├── dmw_update.sas        stored process: the only thing that writes
├── scripts/ag-grid-community.min.js
├── styles/ag-grid.css, ag-theme-alpine.css
├── data/                         CSVs the Home and Data Pipeline Performance tabs read
│   ├── sas_server_status.csv         <- Check-SASServers_v2.0.ps1
│   ├── objectspawner_status.csv      <- Check-ObjectSpawner_v1.0.ps1
│   ├── sas_log_rollup.csv            <- Export-SasLogRollup.ps1
│   ├── sas_log_steps.csv             <- Export-SasLogSteps.ps1 (cgs_ai)
│   └── sas_server_latency_daily.csv  <- Export-ServerLatencyDaily.ps1
└── ps1/
    ├── Check-SASServers_v2.0.ps1
    ├── Check-ObjectSpawner_v1.0.ps1
    ├── Export-SasLogRollup.ps1       one row per LOG FILE
    ├── Export-SasLogSteps.ps1        one row per SAS STEP (wraps cgs_ai)
    └── Export-ServerLatencyDaily.ps1 one row per ENDPOINT per DAY
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

then open `http://localhost:8080/dmw_it_management_v1_5Beta.html`.

### Step 2 — Check the configuration block

Everything adjustable is in `CONFIG` at the very top of the HTML, in eleven
commented sections. At minimum confirm:

| Key | Default | Note |
|---|---|---|
| `SAS_BASE_URL` | `https://cgssaswebu.a70admed.com/SASStoredProcess/do?_program=` | must end in `_program=` |
| `SAS_PROGRAM` | `/User Folders/731o/My Folder/dmw_run` | the read stored process |
| `METADATA_CSV_PATH` | `column_metadata.csv` | relative to the HTML |
| `SERVER_STATUS_CSV` | `data/sas_server_status.csv` | Home tab; Charts scatter, loaded on demand |
| `SERVER_LATENCY_DAILY_CSV` | `data/sas_server_latency_daily.csv` | Charts time series and summary. Empty = compute from raw |
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
| `Export-ServerLatencyDaily.ps1` | `sas_server_latency_daily.csv` | daily, 00:20 |

Try each by hand first:

```powershell
.\Check-SASServers_v2.0.ps1   -TestMode
.\Check-ObjectSpawner_v1.0.ps1 -TestMode
.\Export-SasLogRollup.ps1 -metric_profile sas_log -TestMode -MaxFiles 50
.\Export-SasLogSteps.ps1   -TestMode
.\Export-ServerLatencyDaily.ps1 -TestMode
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
| **Scheduler Overview** | KPI tiles plus row counts by `DASHBOARD_GROUP_BY`. Populates automatically when the query returns. |
| **Server Performance** | Response-time charts for the monitored HTTP endpoints. Hover for values, click the legend to filter endpoints, and set a date range. |
| **Data Pipeline Performance** | Cross-filtered by date and by SAS program: selecting either re-scopes every section of the tab. Built from **two sources**: per-file rollup (program count, run count, CPU by program, CPU over time, ranked detail) and per-step detail from the cgs_ai scanner (CPU by step type, slowest individual steps). |
| **Security** | Software delivery and hotfix cycle, drawn inline as SVG. |
| **Ask DMW Agent** | Embedded Copilot Studio agent, with an always-visible "open in a new tab" fallback. |
| **Connection** | Effective settings and the SAS OLE DB snippet. |
| **Console** | Request URL, status, timing, errors. Start here when something fails. |

> **The three performance tabs are distinct.** *Scheduler Overview* summarises
> what is currently in
> the Data Grid. *Data Pipeline Performance* plots SAS log performance from the two
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

The Data Pipeline Performance tab reads two CSVs with **different grains**. They are
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

## 3.3 Server response time — why `ResponseTimeMs`, and no log column

`Check-SASServers_v2.0.ps1` already records everything these charts need.
**`ResponseTimeMs` is the right field** — it is the wall-clock duration of the
HTTP check, written on every sweep.

**Do not add a log column to the CSV.** The logarithm is a *presentation*
choice, not a fact about the check:

- it duplicates state that can drift from the value it came from,
- it discards precision,
- nobody opening the CSV can read it in milliseconds,
- and it bakes one chart's needs into a file that feeds several views.

The transform belongs to the axis. Both charts scale the axis instead, so the
tick labels and the summary table still read in real milliseconds.

**Why a log axis is needed at all.** Latency is log-normal: a healthy check is
a few hundred milliseconds and a timeout is 30 000. On a linear axis the
timeouts flatten everything else onto the floor and a doubling from 300 ms to
600 ms — which is what you actually want to catch — is invisible.

**Timed-out checks are handled separately.** A `DOWN` row records roughly the
timeout ceiling, not a measurement. Averaged in, it drags the mean toward
30 000 ms and hides the real movement in healthy latency. So:

- the scatter plots them **in red**, so outages stay visible;
- the daily median line and the summary table **exclude** them.

**The two charts**

| Chart | What it answers |
|---|---|
| Every check, by endpoint (scatter, log axis) | What does the spread look like, and when did it break? |
| Daily median by endpoint (line, log axis) | Is normal latency drifting over months? |

The summary table gives p50, p95, p99 and worst per endpoint. **Prefer
percentiles to the mean**: one 20-second stall moves the mean and tells you
nothing, whereas p95 moving from 1.8 s to 3 s is a real change in what users
experience.

If you later want this cheaper to chart, the right move is not a log column —
it is a second CSV of **daily pre-aggregated percentiles** written by the
scanner, so the page reads 4 rows per day instead of 96.

## 3.4 The daily latency roll-up

`Check-SASServers_v2.0.ps1` writes one row per endpoint per sweep. At one sweep
every five minutes that is **288 rows per endpoint per day** — roughly 420,000
rows a year for four endpoints, all of which the browser was downloading and
aggregating on every visit to the Server Performance tab.

`Export-ServerLatencyDaily.ps1` does that arithmetic once, on the server:

```
Date, Environment, ServiceName, Checks, UpChecks, Timeouts, Degraded,
AvailabilityPct, MinMs, P50Ms, P90Ms, P95Ms, P99Ms, MaxMs, MeanMs
```

On the sample data that is **484 rows instead of 11,524 — 96% smaller**, and
43 KB instead of 1.3 MB.

**The Server Performance tab prefers it automatically.** The two time series and the
summary table read the daily file. The per-check file is fetched **only** when
you press *Load per-check detail* for the scatter, which is the one view that
genuinely needs every point. If the daily file is missing, the tab falls back
to computing everything from raw and says so.

Run it daily, after midnight, so the previous day is complete. It is
idempotent — it rewrites the whole output from the input each run, so a missed
day is repaired by the next one rather than leaving a hole.

### Two details that matter

**The percentile definition is shared.** Both the script and the page use
nearest-rank, `index = min(n - 1, floor(n * p))`. Two different definitions
would make the daily file and the raw file disagree by a millisecond or two on
the same data, and someone would eventually spend an afternoon on it. If you
change one, change the other.

**Percentiles do not average.** The 95th percentile of a quarter is *not* the
mean of 90 daily 95th percentiles. So the summary built from the daily file
labels its columns for what they actually are — *typical day p50*, *worst day
p50*, *worst day p95* — plus *worst single check*, which is the one figure that
does survive aggregation. For an exact period-wide percentile, load the
per-check detail.

> **Still on the list.** The Home tab reads the per-check file directly,
> because it needs the most recent sweep and a 14-day availability view. That
> download is unaffected by this change. If it becomes a problem, the fix is to
> have the scanner also write a small "latest status" file; say the word.

## 3.6 Chart interaction

### Tooltips are no longer `<title>`

v1.3Beta put a native SVG `<title>` on every point. Free to write, poor to
use: the browser waits about a second before showing it, it disappears while
the pointer moves, and a dense scatter needed **one DOM node per point** —
968 on the sample data.

v1.4Beta has **zero** `<title>` nodes. Each chart keeps a plain array of its
plotted points, one `mousemove` listener finds the nearest by squared
distance, and a single shared `div` follows the cursor. The tooltip appears on
the same frame as the movement — measured at **60 ms**, which is the test's
own polling interval rather than a delay in the page — and a highlight ring
marks the point being read.

### Click the legend to filter

Every legend entry is a toggle. Hidden series grey out and strike through, the
axes rescale to what is left, and a **show all** button appears once anything
is hidden. The same mechanism drives the donut, where **clicking a segment**
hides that program — the gesture people try first.

### Date range

A shared control on both tabs: presets for 7, 30 and 90 days, 6 months and
all, plus two date inputs for a span the presets do not cover. State is held
per tab, so filtering one does not disturb the other.

The row count reads *"124 of 484 rows · showing 2026-10-13 → 2026-11-12 · data
runs 2026-07-15 → 2026-11-12"*. Printing the full span beside a filtered count
would read as if the filter had not applied, so both are shown.

### More x-axis values

Up from 3–5 labels to up to 12, with the density capped by available width so
they cannot overlap. The format follows the span — times within two days,
day-and-month up to a year, month-and-year beyond — and a second grey line
carries the year when a range crosses one.

## 3.7 Cross-filtering on Data Pipeline Performance

Selecting a program on the donut used to hide one slice and leave the KPIs,
the time series, the detail table and the step-level section still reporting
on everything — so the tab quietly disagreed with itself.

The selection now lives beside the date range in the panel's own state, and
**every section is built from the same filtered set**. Filtering
`DMEC_ADM_UPDATE` moves all of it together:

```
before   16 programs / 864 runs / 74.7 h CPU / 41 error runs
         348 step rows / 16 programs / 1.7 h step CPU
after    15 programs / 834 runs / 68.1 h CPU / 40 error runs
         323 step rows / 15 programs / 1.5 h step CPU
```

**Three ways to filter, one state.** Click a donut segment, click a legend
entry, or click a row in the program detail table. All three add to the same
set.

**Every active filter is visible.** A chip bar lists each hidden program with
an `×` to restore it and a *clear all* button, and states that the programs
are hidden from every section. A filter you cannot see is a filter you will
misread the whole tab through.

**Date and program filters compose** and are independent: changing the date
range keeps the program selection, and clearing the chips keeps the date
range.

Two notes on behaviour:

- The aggregate **"Other (N)"** slice stands for several programs at once, so
  there is nothing single to filter on. Clicking it writes an explanation to
  the Console tab rather than doing something arbitrary.
- Filtering every program leaves the tab with nothing to show, so it renders
  the filter bars and an explanation instead of empty charts.

Chart sizes on this tab: the donut's viewBox went from 420×190 to **560×330**
with the centre following the radius, and the CPU time series from 230 to
**330** high.

## 4. Replacing the sample data

The CSVs in `data/` are synthetic. They contain deliberately scripted
incidents — three endpoint outages and three service outages over 14 days —
so the charts show something other than a flat line. Delete them and run the
scanners for real data; the schemas match exactly, so nothing else changes.

---

## 3.5 Diagnosing a failed job-name save

Renaming a job reported **"Your SAS session has expired"** while the user was
signed in. The cause was in this page, not in SAS: the HTML-response check
tested for the string `SASLogon` first, and a SAS **"Stored Process Error"**
page carries the standard SAS chrome, which includes a `SASLogon` logout link.
So a stored process that was simply not registered was reported as an expired
session.

Two changes:

- **Order and precision.** Specific failures are identified first, and a logon
  page is now recognised only by something a logon page actually has — a
  password input, or a form posting to `SASLogon`. Each failure now names its
  own cause, and the message includes the metadata path that was called.
- **The response is logged before it is interpreted.** Every save writes the
  HTTP status, content type and first 300 characters of the reply to the
  **Console** tab. When a save fails, what the server actually said is the
  thing worth seeing — an interpretation can be wrong, the bytes cannot.

### The GET fallback

`CONFIG.EDIT_FALLBACK_TO_GET` (default **true**) retries the save as a GET when
the POST comes back as HTML. Some Stored Process Web Application setups accept
the read path but not a form POST, and the read path is already proven to work
here.

The tradeoff is real and the page does not hide it: a GET puts the new job name
in the query string, where the web server logs it verbatim, and a browser
refresh can re-submit it. Every use writes a warning to the Console tab. Set it
to `false` to require POST and fail loudly instead.

> If the save still fails, the message now tells you which of three things is
> wrong. The most likely is that `dmw_update` is not registered at
> `CONFIG.EDITABLE_COLUMNS.JobName.program` — check that path in SAS Management
> Console.

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
| All "i" icons grey | `column_metadata.csv` was not fetched. Check the Console tab. Serving from `file://` always fails. (A *slow* fetch is handled — the headers rebuild when the metadata lands.) |
| Home tab empty, "Monitoring data unavailable" | The `data/*.csv` files are missing or the paths in `CONFIG` are wrong. |
| Home tab shows **STALE** | The scheduled task is not running, or cannot write to `data/`. |
| Ask DMW panel blank | The agent's host refuses to be framed (`X-Frame-Options` / `frame-ancestors`), or you are not signed in to the tenant. Open the URL in a normal tab to confirm. |
| Saving a job name returns `CONFLICT` | Someone else renamed it. Reload the grid. |
| Saving returns `FAILED` | The service account lacks `UPDATE` on `[dbo].[Job].[JobName]`. |
| Saving a job name reported `Unexpected token '<'` | Fixed. SAS answered with an HTML page — an expired session, a wrong metadata path, or Result type not `Streaming` — and the old code fed it to `JSON.parse`. The page now names the cause instead. |
| Ask DMW Agent: *This content is blocked* | The agent's host refuses to be framed (`X-Frame-Options` / `frame-ancestors`). **Not fixable from this page** — those are headers on Microsoft's server. Publish the agent via *Copilot Studio → Channels → Custom website*, add this page's origin, and put that embed URL in `CONFIG.ASK_DMW_URL`. The configured URL ends in `source=embedded-builder`, which is the authoring preview, not a published embed endpoint. Use **Open in a new tab** meanwhile. |
| `ERROR 180-322` in the SAS log | An OLE DB init string was built with `%LET`. Fixed in the current files. |

---

## 7. What was tested, and what was not

Exercised headless over http, zero console or page errors:

- All ten tabs render; the tab bar reads
  Home | Data Grid | SQL | Scheduler Overview | Server Performance | Data Pipeline Performance | Security | Ask DMW Agent | Connection | Console
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
- Ask DMW Agent renders the iframe with the configured URL and an always-visible fallback
- Tabs read Home | Data Grid | SQL | Scheduler Overview | Server Performance |
  Data Pipeline Performance | Security | Ask DMW Agent | Connection | Console
- Row counts render in Scheduler Overview, not Server Performance
- Charts carry no `<title>` nodes at all (was 968); the shared tooltip renders
  within 60 ms of the pointer moving
- Legend clicks filter series and a show-all button appears; the date presets
  cut 484 daily rows to 124 for "Last 30 days"
- On Data Pipeline Performance a program selection re-scopes the whole tab:
  KPIs, donut, time series, detail table and the step-level section all move
  together, and the chips survive a date change
- A Stored Process Error page containing a SASLogon link is reported as an
  unregistered stored process, a page with a password field as an expired
  session, and ODS HTML as a Result type problem — three inputs, three messages
- The Server Performance tab reads the 43 KB daily roll-up and does **not** fetch the 1.3 MB
  per-check file until *Load per-check detail* is pressed; with the daily file
  returning 404 it falls back to computing from raw and says so
- Column headers show a sort indicator that cycles ascending → descending →
  cleared, and the metadata icons survive a slow CSV: with the fetch delayed
  2.5 s past the first grid render, the icons start greyed and recover to live
  once the metadata lands

**Not tested here** — there is no SAS, SQL Server, or Windows in the build
environment. Unverified: the real `UPDATE`, `_METAPERSON` resolution, the join
against the live `Job` table, all four PowerShell scripts against real hosts, including the cgs_ai wrapper,
and whether the Copilot agent permits framing from your origin.
