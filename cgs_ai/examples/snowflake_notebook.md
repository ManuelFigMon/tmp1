# Using `cgs_ai` in a Snowflake Notebook — CMS-2022-0193 to CSV + JSON

This walks through pulling public comments for docket **CMS-2022-0193** (the CMS
Interoperability and Prior Authorization Final Rule, CMS-0057-F) from inside a
Snowflake Notebook and writing them to both a CSV and a JSON file on a stage.

Because `cgs_ai` is standard-library only, there are no packages to install —
you only need to make the module importable, turn on outbound access, and store
the API key as a secret.

---

## One-time setup (run once, as an admin role)

The `SECRET`, `NETWORK RULE`, and `STAGE` are **schema-level** objects, so the
worksheet must have an active database + schema first. If you skip this you get
`Cannot perform CREATE SECRET. This session does not have a current database.`
Replace `<your_db>.<your_schema>` below with your own (e.g.
`ADM_PRD.CMS_ADM_CGS_MAC`).

```sql
-- Set context FIRST — this is what avoids the "no current database" error.
USE ROLE ACCOUNTADMIN;             -- or a role with the CREATE INTEGRATION privilege
USE DATABASE <your_db>;
USE SCHEMA <your_schema>;
-- USE WAREHOUSE <your_wh>;        -- only if you have no default warehouse

-- A stage to hold the code and the output files (schema-level).
CREATE STAGE IF NOT EXISTS cgs_ai_stage
    DIRECTORY = (ENABLE = TRUE)
    ENCRYPTION = (TYPE = 'SNOWFLAKE_SSE');

-- Store the regulations.gov API key as a secret (never pass it as a parameter).
CREATE OR REPLACE SECRET regulations_gov_api_key
    TYPE = GENERIC_STRING
    SECRET_STRING = 'PASTE_YOUR_REGULATIONS_GOV_API_KEY_HERE';

-- Allow outbound HTTPS to the API host only (schema-level).
CREATE OR REPLACE NETWORK RULE regulations_gov_network_rule
    MODE = EGRESS
    TYPE = HOST_PORT
    VALUE_LIST = ('api.regulations.gov');

-- Bind the rule + secret into an external access integration.
-- The EAI is ACCOUNT-level (not affected by USE DATABASE) and needs the
-- CREATE INTEGRATION privilege (usually ACCOUNTADMIN). Fully-qualify the
-- objects it references since they live in a schema.
CREATE OR REPLACE EXTERNAL ACCESS INTEGRATION regulations_gov_access_integration
    ALLOWED_NETWORK_RULES = (<your_db>.<your_schema>.regulations_gov_network_rule)
    ALLOWED_AUTHENTICATION_SECRETS = (<your_db>.<your_schema>.regulations_gov_api_key)
    ENABLED = TRUE;

-- Grant usage to the role your notebook runs as.
GRANT USAGE ON SECRET <your_db>.<your_schema>.regulations_gov_api_key TO ROLE <your_notebook_role>;
GRANT USAGE ON INTEGRATION regulations_gov_access_integration TO ROLE <your_notebook_role>;
GRANT READ, WRITE ON STAGE <your_db>.<your_schema>.cgs_ai_stage TO ROLE <your_notebook_role>;
```

## Make the `cgs_ai` package importable

Stage the package as a zip, then put it on the notebook's `sys.path` (Cell 0
below). Python imports the package straight out of the zip, so
`from cgs_ai import ...` works exactly as it does locally.

Build and upload the zip once (from the project root, e.g. via SnowSQL):

```bash
# The zip root must be the `cgs_ai/` package dir so `import cgs_ai` resolves.
( cd src && zip -r ../cgs_ai.zip cgs_ai -x '*__pycache__*' )
snowsql -q "PUT file://cgs_ai.zip @cgs_ai_stage OVERWRITE=TRUE AUTO_COMPRESS=FALSE"
```

You can also `PUT` it directly from a notebook cell if you have the file in the
notebook's Files panel:
`session.file.put("file://cgs_ai.zip", "@cgs_ai_stage", auto_compress=False, overwrite=True)`.

## Attach the external access integration to the notebook

In Snowsight: open the notebook → **⋮** (top-right) → **Notebook settings** →
**External access** → toggle on `REGULATIONS_GOV_ACCESS_INTEGRATION`. Without
this, the outbound call to `api.regulations.gov` is blocked and `get_comments`
will fail. (Restart the notebook session after enabling it.)

---

## Notebook cells

### Cell 0 — put the staged `cgs_ai` package on the path

```python
import sys
from snowflake.snowpark.context import get_active_session
session = get_active_session()

# Download the staged zip and import cgs_ai straight out of it (zipimport).
session.file.get("@cgs_ai_stage/cgs_ai.zip", "/tmp/")
if "/tmp/cgs_ai.zip" not in sys.path:
    sys.path.insert(0, "/tmp/cgs_ai.zip")

import cgs_ai
print("cgs_ai", cgs_ai.__version__)
```

### Cell 1 — read the key from the secret and pull the docket

```python
import _snowflake
from cgs_ai import get_comments, write_output, build_metadata, write_metadata

api_key = _snowflake.get_generic_secret_string("regulations_gov_api_key")

records = get_comments(
    agency="CMS",
    docket_id="CMS-2022-0193",
    api_key=api_key,
    download_type="all",   # comment bodies + attachment metadata
)
print(f"Retrieved {len(records)} comments")
```

### Cell 2 — write CSV and JSON locally

```python
write_output(records, "cms_2022_0193_comments.csv")   # dispatch by extension
write_output(records, "cms_2022_0193_comments.json")
write_metadata(
    build_metadata(agency="CMS", docket_id="CMS-2022-0193", records=records),
    "cms_2022_0193_comments.metadata.json",
)
```

### Cell 3 — persist the files to the stage

A notebook's local filesystem is ephemeral, so copy the files to the stage to
keep them:

```python
from snowflake.snowpark.context import get_active_session
session = get_active_session()

for name in ("cms_2022_0193_comments.csv",
             "cms_2022_0193_comments.json",
             "cms_2022_0193_comments.metadata.json"):
    session.file.put(f"file://{name}", "@cgs_ai_stage",
                     auto_compress=False, overwrite=True)

# Confirm they landed.
session.sql("LS @cgs_ai_stage").show()
```

### Cell 4 (optional) — load into a table for Cortex

If the point is to run Cortex functions (sentiment, summarize) on the comments,
load the records straight into a table:

```python
df = session.create_dataframe(
    [{k: r.get(k) for k in
      ("id", "agencyId", "postedDate", "title", "comment", "docketId", "attachmentCount")}
     for r in records]
)
df.write.mode("overwrite").save_as_table("CMS_2022_0193_COMMENTS")

# Example Cortex sentiment over the comment bodies:
session.sql("""
    SELECT id, title,
           SNOWFLAKE.CORTEX.SENTIMENT(comment) AS sentiment
    FROM CMS_2022_0193_COMMENTS
    WHERE comment <> ''
    LIMIT 20
""").show()
```

---

## Notes

- The API key is read from the secret and never printed, logged, or written to
  any output/metadata file.
- `download_type="all"` fetches comment bodies plus attachment metadata. Use
  `"metadata"` for a fast identifiers-only pull, or `"comments"` for bodies
  without attachments.
- To download the staged files to your machine:
  `GET @cgs_ai_stage/cms_2022_0193_comments.csv file:///tmp/`.
```
