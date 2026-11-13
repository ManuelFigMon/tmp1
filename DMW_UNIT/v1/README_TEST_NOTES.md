# DMW Job Notes — TEST build

An editable notes table attached to the DMW dashboard, keyed on `JobID`. Every
filename carries **`test`** so none of it can be confused with the running
solution.

> **The read path does not change.** `dmw_run.sas` keeps its `SELECT`-only guard
> and is not involved in writing. Notes go through two separate stored processes
> that accept no SQL from the caller at all.

## Files

| File | What it is | Permanent? |
|---|---|---|
| `dmw_it_management_test.html` | The dashboard with the notes feature added | Becomes the new page if you keep it |
| `dmw_test_notes_table.sql` | Table + index + grants. **Give this to the DBA** | The table is permanent |
| `dmw_test_create_notes_table.sas` | Creates the table as the service account, if you have no DBA route | **Delete after one successful run** |
| `dmw_test_note_list.sas` | Read notes. No SQL input | Permanent |
| `dmw_test_note_save.sas` | Write/retract one note. No SQL input | Permanent |

---

## Answering your question first: should the service account create the table?

**No — I was recommending the DBA create it** with `dmw_test_notes_table.sql`,
and grant the service account rights on it afterwards.

But you said SAS is the only way you can invoke `A70ADMED\SVC_CGSUSASSRV`, so
`dmw_test_create_notes_table.sas` exists for that route. One thing to understand
before you run it:

**It will probably fail, and that is informative rather than a problem.** The
grant you obtained earlier was `SELECT` on one view. Creating a table needs
`CREATE TABLE` in `OHM_CGSScheduler` plus `ALTER` on the schema. So the program
runs a **permission probe first** and tells you which situation you are in,
without attempting any DDL:

```json
{"step":"probe","status":"BLOCKED","can_create_table":0,
 "connected_as":"A70ADMED\\SVC_CGSUSASSRV",
 "message":"This account cannot create tables ... Send dmw_test_notes_table.sql to the DBA instead ..."}
```

If you get that, you need the DBA regardless — but you now know the exact
account name and permission to ask about, which is worth the five minutes.

Two further cautions, because this is a production database you don't own:

- `OHM_CGSScheduler` belongs to someone else. Creating objects in it is a
  governance decision. Tell its owner the table exists — an undocumented table
  in a production database makes the next outage harder to diagnose.
- A stored process that can issue DDL should not stay reachable from a browser.
  **Unregister it after one successful run.** The program also has a safety
  interlock (`DMW_TEST_CONFIRM`) that must be set to `1` deliberately.

---

## Step-by-step

### Step 0 — Try the interface before building anything (15 minutes)

`dmw_it_management_test.html` ships with `NOTES_BACKEND: "local"`, which stores
notes in browser `localStorage`. Nothing to install, no database object, no
stored process.

1. Copy `dmw_it_management_test.html` next to the existing page, keeping
   `scripts/` and `styles/` beside it.
2. Open it and press **Logon**, then **Run Query**.
3. Click the **💬 add** badge in the pinned **Notes** column on any row.
4. Type a note and save. Check the **Notes** tab for the cross-job view.

This proves out the whole interface. Decide whether you want the feature at all
before asking a DBA for anything.

> **`localStorage` is not shared storage.** Notes are per-browser and
> per-machine, invisible to colleagues, and deleted when site data is cleared.
> The page says so in an orange banner inside the notes panel. Leave that banner
> alone — a demo mistaken for working shared storage is worse than no demo.

### Step 1 — Create the table

**Preferred:** send `dmw_test_notes_table.sql` to the DBA. It creates the table,
the index, and grants `SELECT, INSERT, UPDATE` on **that table only** to the
service account. `DELETE` is deliberately not granted. It ends with two
verification queries — ask for the output.

**Fallback,** if no DBA route exists:

1. Register `dmw_test_create_notes_table.sas` in SMC, **Result type = Streaming**,
   no input parameters.
2. Run it once as-is. It returns `{"status":"NOT_CONFIRMED"}` and changes
   nothing — this confirms the plumbing works.
3. Set `%let DMW_TEST_CONFIRM = 1;` at the top, save, run again.
4. Read the JSON. `OK` means the table exists and was verified by re-reading
   `sys.tables` — it does not merely assume the DDL worked.
5. **Unregister the stored process.** Then get the grants applied (section 3 of
   the `.sql` file).

### Step 2 — Register the two runtime stored processes

For each of `dmw_test_note_list.sas` and `dmw_test_note_save.sas`:

1. Copy to the stored process source folder.
2. New Stored Process in SMC. **Result type = Streaming.**
3. Input parameters, all Text, none required at prompt:
   - `dmw_test_note_list` → `job_id`, `max_rows`, `debug_mode`
   - `dmw_test_note_save` → `action`, `job_id`, `request_id`, `note_text`,
     `note_id`, `debug_mode`
4. Note each metadata path.

Check the `DMW_*` configuration block at the top of each matches `dmw_run.sas`.

### Step 3 — Point the page at them

In the `CONFIG` block (section 6, near the top of the file):

```js
NOTES_BACKEND: "sas",                                              // was "local"
SAS_NOTE_LIST_PROGRAM: "/User Folders/731o/My Folder/dmw_test_note_list",
SAS_NOTE_SAVE_PROGRAM: "/User Folders/731o/My Folder/dmw_test_note_save",
```

Save, then press **Clear Cache & Reload**. The orange demo banner disappears —
that is your confirmation the switch took effect.

### Step 4 — Verify

1. Add a note. The author should be **your** SAS username, not `(this browser)`.
2. Ask a colleague to open the page. They should see your note.
3. Add a note containing `O'Brien 50% & "quoted"`. It must come back **byte for
   byte**. If it is mangled or truncated, something is building SQL as a string
   — see the next section.
4. Retract a note. The text stays, struck through and greyed.
5. Check `DMW_LOGFOLDER\dmw_notes_YYYYMMDD.log` for the write record.

---

## Design decisions worth knowing

**The author is never sent by the browser.** Every database write arrives as
`SVC_CGSUSASSRV`, so SQL Server cannot tell who wrote a note. SAS can:
`_METAPERSON` holds the signed-in user. `dmw_test_note_save.sas` reads it
**server-side** and stamps it into `CreatedBy`. If the page supplied it, anyone
could post a note under a colleague's name and the audit trail would be
worthless. This is the single detail that makes the feature trustworthy.

**The note text never becomes SQL.** `PROC APPEND` over an OLE DB libname lets
the driver parameterise the insert. There is no statement to inject into. Two
SAS-specific traps are avoided this way:

- *SQL injection* — note text is free text; an apostrophe alone breaks a
  hand-built `INSERT`.
- *Macro resolution* — a note containing `&PROD` or `50% complete` would be
  resolved by the macro processor if it reached SQL as `&note_text`. The
  program uses `symget`, which returns the raw value.

**Writes POST, reads GET.** A 2,000-character note may not survive a query
string; query strings are written verbatim into web server logs, so note text
would leak there; and a browser refresh can silently re-submit a GET and
duplicate the note.

**Append-only, soft delete.** No lost updates when two people type at once, no
concurrency control needed, and a complete audit trail. Retraction preserves the
original text.

**`JobID`, with optional `RequestID`.** A `JobID` note is durable knowledge —
*"fails every month-end by design"* — and shows against every run of that job.
A `RequestID` note is incident-specific. The column is nullable and populated
automatically from the row you clicked, so you get both without choosing.

**Editing is in a panel, not in the grid cell.** AG Grid supports editable
cells, but a 2,000-character field does not belong in a 120-pixel column, and
in-cell editing drags in per-cell save semantics for no benefit.

---

## What was tested, and what was not

Exercised headless against a mocked stored process, both backends, zero console
or page errors:

- Notes column renders, badge opens the panel, counts update after save
- `localStorage` notes survive a page reload
- Save is a **POST**, `Content-Type: application/x-www-form-urlencoded`,
  `credentials: include`, note text in the **body** and not the URL
- No username field is sent — confirmed by inspecting the request body
- The author displayed comes from the server response
- `O'Brien: 50% & <i>x</i> "q"` round-trips exactly, with no HTML injected
- Empty and over-length notes rejected client-side
- Retract posts `action=retract&note_id=…` and renders struck through
- After save the page re-reads from the server rather than patching locally

**Not tested, because there is no SAS or SQL Server in my environment:** the
actual `CREATE TABLE`, the real `PROC APPEND` insert, the `UPDATE` for
retraction, `_METAPERSON` resolution, and the permission probe against a live
server. Steps 1 and 4 above are the real test.
