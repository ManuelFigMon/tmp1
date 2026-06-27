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

- The API stored process accepts a `userid` query param **for offline testing
  convenience**. In production, ignore it and trust **only `&_USERNAME`** so a
  user cannot read or write another user's store. (Marked inline in
  `sas_auth_state_test_api.sas`.)
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
