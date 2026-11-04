# DMW (Data Mart Workbench) -- IT Management Solution

A self-contained web GUI that runs a SQL statement against Microsoft SQL Server and
renders the result set as a navigable grid in the browser, using an existing SAS
Stored Process Web Application as the middleware.

```
Browser (dmw_it_management.html)
      |
      |  HTTPS GET  .../SASStoredProcess/do?_program=<path>&dmq_query=<URL-encoded SQL>
      v
SAS Stored Process Web Application  ->  dmw_run.sas
      |
      |  SAS/ACCESS to OLE DB, explicit pass-through, TDS port 1433
      v
SQL Server  A70TUCGSDATA012.A70ADMED.COM / OHM_CGSScheduler
```

## Why a middleware layer is required

A browser cannot connect to SQL Server. SQL Server speaks the TDS protocol over a raw
TCP socket on port 1433, and browser JavaScript has no raw TCP socket API -- only
HTTP(S), WebSocket and WebRTC. Windows Authentication additionally needs OS-level
credential delegation that a sandboxed page is never granted. Node.js drivers such as
`tedious` require `net`/`tls`/`dns` and ship no browser build; `msnodesqlv8` is a
compiled C++ addon. A server-side bridge is therefore mandatory, and `dmw_run.sas` is
that bridge. Nothing is installed on the workstation.

## Contents

| File | Purpose |
|---|---|
| `dmw_it_management.html` | The application. All configuration is in a `CONFIG` block at the top of the file. |
| `dmw_run.sas` | The SAS Stored Process middleware. Renamed from `dmq_run`; the connection is driven by a `DMW_*` configuration block at the top. |
| `scripts/ag-grid-community.min.js` | AG Grid Community 35.3.1, vendored locally. No CDN is contacted at run time. |
| `styles/ag-grid.css`, `styles/ag-theme-alpine.css` | AG Grid stylesheets, vendored locally. |

Keep the folder structure intact -- the HTML references `scripts/` and `styles/` by
relative path.

## Installation

### 1. Register the stored process in SAS Management Console

1. Copy `dmw_run.sas` to the stored process source folder on the SAS server, e.g.
   `\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\Web_Stored_Processes`.
2. In SMC: **New Stored Process**, name `dmw_run`.
3. **Source code location**: the folder from step 1. **Source file**: `dmw_run.sas`.
4. **Result type: Streaming.** This is required. If it is left at the default, SAS
   wraps the output in ODS HTML and the browser receives a web page where it expects
   JSON. The page detects this case and reports it, but the query will not work.
5. Add two input parameters, both Text and not required at the prompt:
   - `dmq_query` -- the SQL statement
   - `debug_mode` -- `0` or `1`
6. Note the metadata path, e.g. `/User Folders/731o/My Folder/dmw_run`.

### 2. Point the stored process at the database

Edit the configuration block at the top of `dmw_run.sas`:

| Macro variable | Default | Notes |
|---|---|---|
| `DMW_SERVER` | `A70TUCGSDATA012.A70ADMED.COM` | Server Name |
| `DMW_PORT` | `1433` | |
| `DMW_CATALOG` | `OHM_CGSScheduler` | Database Name / Initial Catalog |
| `DMW_PROVIDER` | `MSOLEDBSQL19` | Try `MSOLEDBSQL` then `SQLNCLI11` if the provider is not registered |
| `DMW_TRUSTED` | `1` | `1` = Windows Authentication (Integrated Security = SSPI); `0` = SQL Server authentication using `DMW_USER`/`DMW_PASS` |
| `DMW_MAXROWS` | `5000` | Row cap; `0` disables it |
| `DMW_LOGFOLDER` | UNC log path | Empty value disables server-side logging |

### The database identity is the SAS service account, not the signed-in user

With `DMW_TRUSTED = 1` the connection is made as the **SAS Stored Process Server
service account** -- in UNIT, `A70ADMED\SVC_CGSUSASSRV`. It is **not** made as the
person who signed in to the web application. This is the single most common reason a
first query fails:

```
SYSDBMSG = OLEDB: Unable to Initialize: Login failed for user 'A70ADMED\SVC_CGSUSASSRV'.
SYSDBRC  = 28000
```

The remedy is a one-time grant from the DBA on `A70TUCGSDATA012.A70ADMED.COM`:

```sql
USE [master];
CREATE LOGIN [A70ADMED\SVC_CGSUSASSRV] FROM WINDOWS;

USE [OHM_CGSScheduler];
CREATE USER [A70ADMED\SVC_CGSUSASSRV] FOR LOGIN [A70ADMED\SVC_CGSUSASSRV];
GRANT SELECT ON [dbo].[View_Scheduler_PBI_Dashboard] TO [A70ADMED\SVC_CGSUSASSRV];
```

Confirm the account name against your own SAS log before asking for the grant -- run
the page once with `SAS_DEBUG: 1` and read `SYSDBMSG`; the failing account is named in
the message. Grant `SELECT` on the target view and nothing more.

Two consequences follow. Every DMW user shares one database identity, so the database
cannot distinguish who ran a statement -- use the server-side log at `DMW_LOGFOLDER`,
which records the SAS `_METAPERSON`, for that. And per-user database identity is not
achievable by configuration here; it would require Kerberos constrained delegation
across the browser, the mid-tier and the stored process server.

If the DBA will not grant the service account a login, set `DMW_TRUSTED = 0` and
supply `DMW_USER` / `DMW_PASS` -- preferably through a SAS authdomain rather than
literals in the file.

No credentials are stored in `dmw_run.sas` in either case.

### 3. Deploy the web page

Copy the whole `DMW_UNIT` folder to a web server directory, ideally the same server
that hosts SAS so the SAS session cookie is same-origin. Then edit the `CONFIG` block
at the top of `dmw_it_management.html`:

| Key | Notes |
|---|---|
| `SERVER_HOST`, `PORT`, `INITIAL_CATALOG`, `AUTH_MODE`, `AUTH_USER` | Display only. They must agree with the `DMW_*` values above; changing them here alone does not re-point the connection. |
| `SAS_BASE_URL` | Stored Process Web Application entry point. Must end in `_program=`. |
| `SAS_PROGRAM` | The SMC metadata path from step 1, copied exactly. |
| `SAS_QUERY_PARAM` | `dmq_query`. Keep as-is unless you rename the parameter in the stored process too. |
| `SAS_LOGON_PROGRAM`, `SAS_LOGOUT_URL` | SAS logon redirect and logout endpoints. |
| `SAS_DEBUG` | `1` writes the statement, row count and elapsed time to the SAS log (never to the HTTP response). |
| `REQUEST_TIMEOUT_MS` | Raise for queries over large views. |
| `DISABLE_CACHE`, `AUTO_RUN_ON_LOGON` | |
| `INITIAL_SQL` | The statement that loads on start-up. Fully editable at run time in the SQL pane. |
| `SAVED_QUERIES` | Toolbar dropdown. Index 0 uses `INITIAL_SQL` (`sql: null`). |
| `GRID_PAGE_SIZE`, `GRID_PAGE_SIZES`, `GRID_THEME`, `DASHBOARD_GROUP_BY`, `VISUALIZE_GROUP_BY` | Grid and chart behaviour. |

The page **must be served over http(s)**. Opening it from `file://` will always be
blocked by browser origin rules for the live call, because the SAS session cookie
cannot be attached to a request from a `file://` origin.

## Usage

1. Open the page. The header shows the target server, database and authentication mode.
2. Press **Logon** (top right). You are redirected to SAS Logon and returned here with
   an active session; the page then sends your SAS session cookie automatically with
   `credentials: 'include'`. **No password is ever collected or stored by this page.**
3. With `AUTO_RUN_ON_LOGON: true` the initial query runs immediately. Otherwise press
   **Run Query**.

### Tabs

- **Data Grid** -- AG Grid results: column filters, floating filters, sorting,
  multi-row selection, pagination, quick search across all columns, Auto-Size.
- **SQL** -- edit the statement and re-run it. Only `SELECT` statements are accepted;
  the stored process returns an empty array for anything else.
- **Dashboard** -- KPI tiles (rows returned, distinct jobs, completed, failed, not yet
  completed) plus a row count by `DASHBOARD_GROUP_BY`.
- **Visualize** -- horizontal bar chart of row counts by `VISUALIZE_GROUP_BY`.
- **Connection** -- the effective connection settings and the equivalent SAS OLE DB
  pass-through snippet, for copying into a SAS program.
- **Console** -- the request URL, HTTP status, timing and any error returned by the
  stored process. Start here when something does not work.

### Hamburger menu (top right)

- **Generate Dashboard** -- builds the Dashboard tab from the current result set.
- **Visualize Data** -- builds the Visualize tab.
- **Export to CSV** -- exports the rows currently visible in the grid, i.e. filters
  and sorting are respected. Filename is
  `DMW_<database>_<MMDDYYYY>_<HHMMSS>.csv`.

**About** and **Help** open in-page modals. **Clear Cache & Reload** forces a fresh
copy of the page and its assets after you edit `CONFIG`.

## Output contract

`dmw_run.sas` returns `Content-type: application/json` and a bare JSON array of row
objects:

```json
[ { "RequestID": 1, "JobName": "NIGHTLY_LOAD", "Status": "Completed" } ]
```

- `[]` -- the statement contained no `SELECT`.
- `{"error": "..."}` -- the pass-through failed; the message from `SYSDBMSG` is
  escaped and included, and the page surfaces it on the Console tab.

The page also accepts `{rows:[...]}`, `{data:[...]}`, `{records:[...]}`,
`{results:[...]}` and a PROC JSON wrapper, so it tolerates a stored process that was
registered without `nosastags`.

## Troubleshooting

| Symptom | Cause |
|---|---|
| A SAS logon page appears in the Console tab instead of data | No active SAS session. Press **Logon**. |
| HTML wrapped around the JSON | Result type is not **Streaming** in SMC. |
| `provider not registered` | Change `DMW_PROVIDER` in `dmw_run.sas`. |
| `Login failed for user 'A70ADMED\SVC_CGSUSASSRV'` (`SYSDBRC = 28000`) | The SAS service account has no SQL Server login. Not a code defect — see "The database identity is the SAS service account" above for the grant. |
| `ERROR 180-322: Statement is not valid or it is used out of proper order`, pointing at `Integrated Security=SSPI` | An OLE DB init string was built with `%LET`. A semicolon ends a `%LET` statement, so only `Provider=...` is assigned and the rest runs as SAS statements. Fixed in the current `dmw_run.sas`, which builds the string in a DATA step; if you see this you are running an older copy. |
| The browser shows a SAS "Stored Process Error" page instead of the JSON `{"error":...}` | A SAS syntax error aborted the program before it could write to `_webout`. Press **Show SAS Log** on that page and read the first red `ERROR`. |
| Nothing happens from `file://` | Serve the page over http(s). |
| Stale behaviour after editing `CONFIG` | **Clear Cache & Reload**, or Ctrl+F5. |

## Security notes

- The stored process executes SQL supplied over HTTP. The `SELECT` check is a guard,
  not a sandbox. Grant the stored process identity only the database permissions it
  genuinely needs.
- Every DMW user shares one database identity, the SAS service account. The database
  audit trail cannot attribute a statement to a person; the `DMW_LOGFOLDER` log can.
- `debug_mode` output goes to the SAS log only, never to the HTTP response, and the
  init string is echoed with any `Password=` value masked.
- Lower `DMW_MAXROWS` if the page should not be able to pull large result sets.
