# SAS Auth + Server-Side State Test (`sas_auth_state_test`)

A foundation prototype for moving report/dashboard state **out of the browser**
and into a per-user, server-side datastore that is read and written through a
**SAS Stored Process** after the user authenticates via SAS Web Auth.

The end-to-end goal it proves out:

> Authenticate → see a **data-driven dropdown** → **add a new item and save it**
> → **sign out** → **sign back in** → the new item is **still there** (because it
> was persisted server-side, not in the browser).

## Files (all prefixed `sas_auth_state_test`)

| File | Role |
|---|---|
| `v3/security/sas_auth_state_test.html` | The portal page: SAS logon round-trip, session display, and the data-driven dropdown with Add / Save / Update / Delete / Reload. |
| `v3/security/sas_auth_state_test_redirect.sas` | Auth stored process: captures `&_USERNAME`, derives a folder-safe userid, **ensures `meta/profiles/<userid>/data/` exists**, and 302-redirects back to the HTML with `?userid=…&loginTime=…&sasAuth=1`. |
| `v3/security/sas_auth_state_test_api.sas` | State stored process: `action=read` streams the user's JSON store; `action=save` writes it. |
| `meta/profiles/<userid>/data/sas_auth_state_test_state.json` | The per-user datastore. A seed for user `731o` is committed so the flow works out of the box. |

## Datastore shape

```json
{
  "userid": "731o",
  "items": [ { "id": 1, "value": "Region: Northeast" } ],
  "updated": "02FEB2026:09:15:00"
}
```

## Backups (taken before every overwrite)

Because a Save is a **full overwrite**, the API copies the current file to a
timestamped backup **before** writing the new version, so every save is
recoverable:

```
meta/profiles/<userid>/data/
├── sas_auth_state_test_state.json                       ← live store
└── backups/
    ├── sas_auth_state_test_state_20260202_091500.json   ← prior versions
    └── sas_auth_state_test_state_20260202_142233.json      (YYYYMMDD_HHMMSS)
```

- The backup runs only when a live file already exists (the very first save has
  nothing to copy).
- It uses `fcopy()` between two `filename()` filerefs; the `backups/` folder is
  created on demand with `dcreate()`.
- The save response JSON includes a `"backup"` field with the path written
  (or `null` on the first save), and the server log records it.
- To restore, copy the desired backup back over
  `sas_auth_state_test_state.json`. (Pruning old backups is left to a separate
  housekeeping job; the API never deletes them.)

## How the three pieces talk

```
 sas_auth_state_test.html
        │  (1) click "Sign in with SAS"
        ▼
 SASStoredProcess/do?_program=…/sas_auth_state_test_redirect&returnUrl=<this page>
        │  (2) SAS authenticates, &_USERNAME captured, folder ensured
        ▼  302 redirect
 sas_auth_state_test.html?userid=731o&loginTime=…&sasAuth=1
        │  (3) page reads userid, then calls the API
        ▼
 SASStoredProcess/do?_program=…/sas_auth_state_test_api&action=read&userid=731o
        │  (4) returns {items:[…]} → fills the dropdown
        │      "Save to server" → action=save&items=<json>
        ▼
 meta/profiles/731o/data/sas_auth_state_test_state.json   ← state lives here
```

## The browser does NOT write the file — SAS does

A common point of confusion: browser JavaScript is sandboxed and **cannot write
to the filesystem**. That rule is about the **client machine**, and this design
respects it. The page never touches disk. All it does is make an HTTP request;
the **SAS stored process running on the server** performs the actual file write
under the SAS process's OS account — exactly like any REST API's "Save" button.

```
 Browser  (sas_auth_state_test.html)            SAS server  (sas_auth_state_test_api.sas)
 ------------------------------------           --------------------------------------------
 fetch(".../sas_auth_state_test_api
        &action=save
        &userid=731o
        &items=[{...}]")            ───────────►  receives the request, then runs:

                                                    data _null_;
                                                      file "<...>/sas_auth_state_test_state.json"
                                                           lrecl=32767;     /* <- opens the file */
                                                      put '{';              /* <- writes each line */
                                                      put '  "items": ' ...;
                                                      put '}';
                                                    run;                    /* <- file is written */

       ◄───────────────────────────────────────   {"status":"saved", "items":[...], "updated":"..."}
```

So "Save is a full overwrite" means the **server-side SAS program overwrites the
server-side file**. No client-side disk access is involved at any point.

The only thing the browser ever persists locally is the **offline fallback**,
and even that is **`localStorage`** (the browser's own sandboxed key-value
store), not a file on disk. The fallback exists solely so the workflow is
demoable without the live SAS server; in production every write goes through
SAS.

### What SAS code does the writing

The write is plain Base-SAS I/O inside a `DATA _NULL_` step — no special
package required:

| SAS element | Role |
|---|---|
| `FILE "<path>" lrecl=32767;` | **FILE statement** — directs the step's output to the external file (this is what "opens" the target for writing). `lrecl` sets the max line length so long JSON lines aren't truncated. |
| `PUT '...';` | **PUT statement** — writes one line of text to the file currently named by `FILE`. The JSON document is emitted line by line with successive `PUT`s. |
| `FILE _webout;` | A reserved fileref for the **HTTP response stream** — the same `FILE`/`PUT` mechanism is reused to send the JSON reply back to the browser. |
| `dcreate('data', udir)` | **DCREATE function** — creates the `…/<userid>/data` directory if missing (idempotent, guarded by `fileexist`). |
| `fileexist(path)` | **FILEEXIST function** — tests whether the folder/file already exists before creating or reading. |
| `symget('items_in')`, `strip()` | Pull the URL-decoded `items` payload from the macro variable and trim it before writing. |

For **read**, the step uses `INFILE "<path>"` + `INPUT` to stream the existing
file's bytes back out through `_webout`. So the entire datastore is managed with
ordinary `FILE`/`PUT` (write) and `INFILE`/`INPUT` (read) statements — the same
primitives SAS has used for flat-file I/O for decades.

## Configure for your environment

In **`sas_auth_state_test.html`** (`CONFIG` and `STATE_API` blocks):
- `portalUrl` — the deployed URL of this HTML page.
- `authProgramPath` — SMC path of `sas_auth_state_test_redirect`.
- `STATE_API.apiProgramPath` — SMC path of `sas_auth_state_test_api`.
- `STATE_API.enabled` — leave `true` to use the SAS API; the page automatically
  falls back to offline mode if the API can't be reached.

In **both `.sas` files**:
- `profile_root` — the server filesystem directory that the web path
  `meta/profiles/` maps to. Both stored processes must agree on this value.

Register each `.sas` as a **Streaming** stored process and grant its identity
read/write on `profile_root`.

## Offline / local demo mode (no SAS server)

The page is built so you can exercise the full UX without the live server:

- **Reads** fall back to the committed seed file
  (`meta/profiles/<userid>/data/sas_auth_state_test_state.json`, fetched at
  `../../meta/profiles/<userid>/data/…` relative to the page) and then to a
  per-user `localStorage` overlay.
- **Writes** fall back to `localStorage` (keyed by userid), so Add → Save →
  sign out → sign back in shows the persisted item within the same browser.

The badge and footer note in the State card show which backend is live
(**SAS API** vs **local fallback**). To force a real round-trip test you need
the deployed stored processes; locally you can simulate the post-auth redirect
by opening:

```
sas_auth_state_test.html?userid=731o&sasAuth=1
```

## Update / Delete

Beyond the required Add + Save, the dropdown also supports **Update selected**
(replace the highlighted item with the New-item value) and **Delete selected**.
Both modify the working list; **Save to server** persists the whole list.

## Security notes (before production)

- **`PRODUCTION_MODE`** (in the CONFIG block of `sas_auth_state_test_api.sas`)
  controls whose store the API may touch:
  - `1` (ship with this) — the API **ignores the `userid` query param** and
    derives the folder strictly from **`&_USERNAME`**, so a user can only ever
    reach their own `meta/profiles/<their-id>/data/` store. Editing `userid=`
    in the URL has no effect.
  - `0` — accepts the `userid` param (falling back to `&_USERNAME`); a
    convenience for offline/local testing only. Do not deploy with `0`.
- Validate/escape the `items` payload; the demo writes the received JSON text
  after a light bracket check.
- Serve over HTTPS; the SAS session cookie carries the authenticated identity.

## About the `.accdb`

The uploaded `sas_portal.accdb` (Microsoft Access) can serve as an alternative
production backend if you prefer a relational store over per-user JSON files —
the API stored process would `PROC SQL` against it via an ODBC/PC Files engine
instead of reading/writing the JSON file. This prototype implements the
JSON-file store specified by the task (`meta/profiles/<userid>/data/`); swapping
in the Access backend only changes the body of `sas_auth_state_test_api.sas`.
