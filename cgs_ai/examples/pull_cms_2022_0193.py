"""Pull the CMS Interoperability and Prior Authorization Final Rule comments.

Docket CMS-2022-0193 (rule CMS-0057-F), written to CSV with a metadata sidecar.

Run with the API key in the environment::

    export REGULATIONS_GOV_API_KEY=your_key_here
    python examples/pull_cms_2022_0193.py
"""

import os

from cgs_ai import build_metadata, get_comments, write_metadata, write_output

records = get_comments(
    agency="CMS",
    docket_id="CMS-2022-0193",
    api_key=os.environ["REGULATIONS_GOV_API_KEY"],
    download_type="all",
)

write_output(records, "cms_2022_0193_comments.csv")
write_metadata(
    build_metadata(agency="CMS", docket_id="CMS-2022-0193", records=records),
    "cms_2022_0193_comments.metadata.json",
)

print(f"Retrieved {len(records)} comment(s).")
