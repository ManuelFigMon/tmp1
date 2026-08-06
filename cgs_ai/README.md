# cgs_ai

`cgs_ai` is an installable Python package of CGS AI utilities designed to run
**inside Snowflake** (Snowpark UDFs and stored procedures) as well as on a
normal workstation. It is intentionally **standard-library only** — no
third-party dependencies — so it survives Snowflake's restricted Python
sandbox. The package is meant to grow over time; its first module,
`cgs_ai.regulations`, retrieves public comments from the
[Regulations.gov API (v4)](https://open.gsa.gov/api/regulationsgov/).

## Install

```bash
pip install -e .
```

Run the tests (they mock all HTTP and never hit the network):

```bash
pip install -e ".[test]"
pytest
```

## Quickstart — CMS-2022-0193

Pull the CMS *Interoperability and Prior Authorization* Final Rule
(**CMS-0057-F**, docket **CMS-2022-0193**) comments to CSV.

### As a library

```python
import os
from cgs_ai import get_comments, write_output, build_metadata, write_metadata

records = get_comments(agency="CMS", docket_id="CMS-2022-0193",
                       api_key=os.environ["REGULATIONS_GOV_API_KEY"],
                       download_type="all")
write_output(records, "cms_2022_0193_comments.csv")
write_metadata(build_metadata(agency="CMS", docket_id="CMS-2022-0193", records=records),
               "cms_2022_0193_comments.metadata.json")
```

### From the CLI

The key is read from the `REGULATIONS_GOV_API_KEY` environment variable:

```bash
export REGULATIONS_GOV_API_KEY=your_key_here
python -m cgs_ai.regulations --agency CMS --docket-id CMS-2022-0193 \
    --output cms_2022_0193_comments.csv
```

A `cms_2022_0193_comments.metadata.json` sidecar is written automatically. The
output format is inferred from the `--output` extension (`.json`, `.csv`,
`.xml`); use `--format` to override it.

## Inputs / outputs reference

### `get_comments(...)` inputs

| Argument | Type | Required | Description |
|---|---|---|---|
| `agency` | str | yes | Agency acronym, e.g. `CMS`. |
| `start_date` | str | no | `YYYY-MM-DD` lower bound on `postedDate`; filter omitted if unset. |
| `end_date` | str | no | `YYYY-MM-DD` upper bound on `postedDate`; filter omitted if unset. |
| `api_key` | str | yes* | Regulations.gov key. *Resolved from env/secret when `None` (see below). |
| `keyword` | str | no | Free-text search term (`filter[searchTerm]`). |
| `include_attachments` | bool | no | Include attachment download URLs (default `False`). |
| `download_type` | str | no | `comments` \| `attachments` \| `metadata` \| `all` (default `all`). |
| `docket_id` | str | no | Docket id, e.g. `CMS-2022-0193` (`filter[docketId]`). |

**`download_type` semantics**

| Value | What is retrieved / returned |
|---|---|
| `comments` | Comment body plus core fields. |
| `metadata` | Identifiers / dates / docket only — no full body, no detail calls. |
| `attachments` | Attachment metadata (and content URLs when `include_attachments=True`). |
| `all` | Everything combined (default). |

### Comment record fields

Each returned record is a `dict`:

| Field | Description |
|---|---|
| `id` | Unique comment identifier (e.g. `CMS-2022-0193-0001`). |
| `agencyId` | Owning agency acronym. |
| `postedDate` | Public posting date (ISO-8601). |
| `title` | Comment title/subject. |
| `comment` | Full text body (empty for `metadata` pulls). |
| `docketId` | Docket identifier. |
| `attachmentCount` | Number of attachments. |

With attachment pulls each record also gets an `attachments` list.

### Outputs

| Function | Output |
|---|---|
| `write_output(records, filename)` | Dispatches by extension → `.json` / `.csv` / `.xml`. |
| `write_json` / `write_csv` / `write_xml` | Explicit format writers. |
| `write_metadata(meta, filename)` | Writes a metadata JSON sidecar. |
| `build_metadata(...)` | Builds the pull-metadata record (never contains the key). |

The metadata record contains: `agency`, `docket_id`, `start_date`, `end_date`,
`keyword`, `download_type`, `record_count`, `retrieved_at` (UTC ISO-8601),
`source` (API base URL), and `api_version`.

## API key handling (secret)

The API key is treated as a **secret**: it is never printed, logged, written to
any output/metadata file, or included in error messages. `resolve_api_key()`
looks for it in this order:

1. The explicit `api_key` argument.
2. The `REGULATIONS_GOV_API_KEY` environment variable.
3. A Snowflake secret named `regulations_gov_api_key` (read via the
   `_snowflake` module inside a Snowpark handler).

If none is found, a clear `RuntimeError` is raised (with no key material in it).

## Deploying `cgs_ai` in Snowflake

Because `cgs_ai` is stdlib-only, no external packages need to be installed —
that is exactly what makes it easy to run in Snowflake's sandbox. Two things
still require configuration: getting the code onto Snowflake, and allowing the
outbound call to `api.regulations.gov` with the key stored as a secret.

### 1. Package & stage the code

Build a zip of the importable package and upload it to a stage, then reference
it from a handler via `IMPORTS`.

```sql
CREATE STAGE IF NOT EXISTS cgs_ai_stage;
```

```bash
# From the project root: zip the import root so `import cgs_ai` works.
( cd src && zip -r ../cgs_ai.zip cgs_ai )
snowsql -q "PUT file://cgs_ai.zip @cgs_ai_stage OVERWRITE=TRUE AUTO_COMPRESS=FALSE"
```

Alternatively, in a Snowpark session you can `session.add_import("cgs_ai.zip")`
before registering the UDF/procedure. No `packages=[...]` are needed since the
code is pure standard library.

### 2. Allow outbound access + store the key as a secret

Snowflake blocks outbound network calls unless a **network rule** and an
**external access integration** are configured, so `get_comments` cannot reach
`api.regulations.gov` without them. The API key should live in a Snowflake
**secret**, not be passed as a parameter.

```sql
-- (a) Store the regulations.gov API key as a secret.
CREATE OR REPLACE SECRET regulations_gov_api_key
    TYPE = GENERIC_STRING
    SECRET_STRING = 'your_regulations_gov_api_key';

-- (b) Allow egress to the API host only.
CREATE OR REPLACE NETWORK RULE regulations_gov_network_rule
    MODE = EGRESS
    TYPE = HOST_PORT
    VALUE_LIST = ('api.regulations.gov');

-- (c) External access integration binding the rule + secret together.
CREATE OR REPLACE EXTERNAL ACCESS INTEGRATION regulations_gov_access_integration
    ALLOWED_NETWORK_RULES = (regulations_gov_network_rule)
    ALLOWED_AUTHENTICATION_SECRETS = (regulations_gov_api_key)
    ENABLED = TRUE;

-- (d) Grant usage as needed.
GRANT USAGE ON SECRET regulations_gov_api_key TO ROLE my_role;
GRANT USAGE ON INTEGRATION regulations_gov_access_integration TO ROLE my_role;
```

### 3. Register a UDF or stored procedure

The handler reads the key from the secret (never a parameter) and calls
`get_comments`. Inside a Snowpark handler, `resolve_api_key()` will pick up the
secret automatically when it is bound as `regulations_gov_api_key`.

```sql
CREATE OR REPLACE FUNCTION fetch_regulations_comments(agency STRING, docket_id STRING)
    RETURNS VARIANT
    LANGUAGE PYTHON
    RUNTIME_VERSION = '3.9'
    HANDLER = 'handler'
    IMPORTS = ('@cgs_ai_stage/cgs_ai.zip')
    EXTERNAL_ACCESS_INTEGRATIONS = (regulations_gov_access_integration)
    SECRETS = ('regulations_gov_api_key' = regulations_gov_api_key)
AS
$$
import _snowflake
from cgs_ai import get_comments

def handler(agency, docket_id):
    api_key = _snowflake.get_generic_secret_string('regulations_gov_api_key')
    return get_comments(agency=agency, docket_id=docket_id,
                        api_key=api_key, download_type="all")
$$;
```

A stored procedure is analogous — swap `CREATE FUNCTION` for
`CREATE PROCEDURE ... RETURNS VARIANT ... HANDLER='handler'` and keep the same
`EXTERNAL_ACCESS_INTEGRATIONS` and `SECRETS` clauses.

```sql
SELECT fetch_regulations_comments('CMS', 'CMS-2022-0193');
```
