#!/usr/bin/env python3
"""
fetch_regulations_comments.py

Downloads public comments from regulations.gov for a given agency,
date range, and optional keyword search. Outputs JSON (default) or CSV,
ready for consumption by the NLP Sentiment Dashboard.

Uses only the Python standard library (urllib, json, csv, argparse).

Example (as a library)
----------------------
    from fetch_regulations_comments import fetch_comments, write_json

    comments = fetch_comments(
        agency="CMS",
        start_date="2026-01-01",
        end_date="2026-03-31",
        api_key="YOUR_API_KEY_HERE",
        keyword="telehealth reimbursement",
        include_attachments=False,
    )

    write_json(comments, "cms_telehealth_comments.json")
    print(f"Downloaded {len(comments)} comments")

Example (command line)
----------------------
    python fetch_regulations_comments.py \\
        --agency CMS \\
        --start-date 2026-01-01 \\
        --end-date 2026-03-31 \\
        --api-key YOUR_API_KEY_HERE \\
        --keyword "telehealth reimbursement" \\
        --output cms_comments.json \\
        --format json
"""

import argparse
import json
import csv
import time
import urllib.request
import urllib.parse
import urllib.error
from datetime import datetime

REGULATIONS_GOV_BASE = "https://api.regulations.gov/v4/comments"

SUPPORTED_AGENCIES = {
    "CMS": "CMS",   # Centers for Medicare & Medicaid Services
}


def fetch_comments(agency: str,
                   start_date: str,
                   end_date: str,
                   api_key: str,
                   keyword: str = None,
                   include_attachments: bool = False,
                   page_size: int = 250,
                   max_pages: int = 20,
                   request_delay_sec: float = 0.4) -> list:
    """
    Fetch public comments from regulations.gov for the given agency and date range.

    Parameters
    ----------
    agency : str
        Agency acronym. Currently only 'CMS' is supported.
    start_date : str
        Start of date range, format 'YYYY-MM-DD'.
    end_date : str
        End of date range, format 'YYYY-MM-DD'.
    api_key : str
        Your regulations.gov API key (https://open.gsa.gov/api/regulationsgov/).
    keyword : str, optional
        Free-text keyword filter applied via the API's searchTerm parameter.
    include_attachments : bool, optional
        If False (default), attachment metadata/content is not fetched —
        only the comment body text and core metadata are retrieved.
    page_size : int, optional
        Results per page (API max is 250).
    max_pages : int, optional
        Safety cap on number of pages fetched (250 * max_pages comments max).
    request_delay_sec : float, optional
        Delay between paginated requests to respect rate limits.

    Returns
    -------
    list of dict
        Each dict represents one comment with normalized fields:
        id, agencyId, postedDate, title, comment, docketId, attachmentCount.

    Raises
    ------
    ValueError
        If agency is not supported or dates are malformed.
    urllib.error.HTTPError
        If the API request fails (e.g., bad API key, rate limit).
    """
    if agency not in SUPPORTED_AGENCIES:
        raise ValueError(f"Unsupported agency '{agency}'. Supported: {list(SUPPORTED_AGENCIES)}")

    # Validate date format
    for d in (start_date, end_date):
        datetime.strptime(d, "%Y-%m-%d")

    all_comments = []
    page = 1

    while page <= max_pages:
        params = {
            "filter[agencyId]": SUPPORTED_AGENCIES[agency],
            "filter[postedDate][ge]": start_date,
            "filter[postedDate][le]": end_date,
            "page[size]": page_size,
            "page[number]": page,
            "sort": "postedDate",
            "api_key": api_key,
        }
        if keyword:
            params["filter[searchTerm]"] = keyword

        url = REGULATIONS_GOV_BASE + "?" + urllib.parse.urlencode(params)

        req = urllib.request.Request(url, headers={"Accept": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                body = json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            raise urllib.error.HTTPError(
                e.url, e.code, f"regulations.gov API error: {e.reason}", e.headers, e.fp
            )

        data = body.get("data", [])
        if not data:
            break

        for item in data:
            attrs = item.get("attributes", {})
            record = {
                "id": item.get("id"),
                "agencyId": attrs.get("agencyId"),
                "postedDate": attrs.get("postedDate"),
                "title": attrs.get("title"),
                "comment": attrs.get("comment") or attrs.get("highlightedContent") or "",
                "docketId": attrs.get("docketId"),
                "attachmentCount": attrs.get("attachmentCount", 0),
            }
            if include_attachments:
                # Note: fetching attachment content requires a separate API call
                # per comment (GET /comments/{id}?include=attachments). Not
                # implemented by default to keep requests minimal.
                record["attachmentsRequested"] = True
            all_comments.append(record)

        # Stop if fewer results than page_size were returned (last page)
        if len(data) < page_size:
            break

        page += 1
        time.sleep(request_delay_sec)

    return all_comments


def write_json(comments: list, out_path: str):
    """Write comments list to a JSON file (array of objects)."""
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(comments, f, indent=2, ensure_ascii=False)


def write_csv(comments: list, out_path: str):
    """Write comments list to a CSV file."""
    if not comments:
        with open(out_path, "w", encoding="utf-8") as f:
            f.write("")
        return
    fieldnames = list(comments[0].keys())
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(comments)


def main():
    parser = argparse.ArgumentParser(description="Download CMS public comments from regulations.gov")
    parser.add_argument("--agency", default="CMS", choices=list(SUPPORTED_AGENCIES))
    parser.add_argument("--start-date", required=True, help="YYYY-MM-DD")
    parser.add_argument("--end-date", required=True, help="YYYY-MM-DD")
    parser.add_argument("--api-key", required=True)
    parser.add_argument("--keyword", default=None)
    parser.add_argument("--include-attachments", action="store_true", default=False)
    parser.add_argument("--output", default="comments.json")
    parser.add_argument("--format", choices=["json", "csv"], default="json")
    args = parser.parse_args()

    comments = fetch_comments(
        agency=args.agency,
        start_date=args.start_date,
        end_date=args.end_date,
        api_key=args.api_key,
        keyword=args.keyword,
        include_attachments=args.include_attachments,
    )

    if args.format == "csv":
        write_csv(comments, args.output)
    else:
        write_json(comments, args.output)

    print(f"Wrote {len(comments)} comments to {args.output}")


if __name__ == "__main__":
    main()
