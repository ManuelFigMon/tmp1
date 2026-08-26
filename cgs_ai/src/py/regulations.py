"""Retrieve public comments from the Regulations.gov API (v4).

This is the first module of the ``cgs_ai`` package. It is written using only
the Python standard library so that it can run inside Snowflake's restricted
Python environment (Snowpark UDFs / stored procedures) without pulling in any
third-party dependency.

Comment record data dictionary
------------------------------
Every record returned by :func:`get_comments` is a plain ``dict`` with the
following fields:

============== =====================================================================
Field          Description
============== =====================================================================
``id``            Unique Regulations.gov comment identifier (e.g. ``CMS-2022-0193-0001``).
``agencyId``      Acronym of the agency that owns the docket (e.g. ``CMS``).
``postedDate``    Date the comment was posted publicly (ISO-8601 date).
``title``         Title/subject line of the comment.
``comment``       Full text body of the comment (empty for ``metadata`` pulls).
``docketId``      Identifier of the docket the comment belongs to (e.g. ``CMS-2022-0193``).
``attachmentCount`` Number of attachments associated with the comment.
============== =====================================================================

When ``download_type="attachments"`` or ``"all"`` each record additionally
carries an ``attachments`` list describing each attachment (title, format, and,
when ``include_attachments=True``, the download ``url``).

The Regulations.gov API key is treated as a secret: it is never printed,
logged, embedded in output files, or included in metadata or error messages.
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional

# --------------------------------------------------------------------------- #
# Constants
# --------------------------------------------------------------------------- #

API_BASE_URL = "https://api.regulations.gov/v4"
API_VERSION = "v4"

#: Name of the environment variable that may hold the API key.
API_KEY_ENV_VAR = "REGULATIONS_GOV_API_KEY"

#: Default name of the Snowflake secret that may hold the API key.
SNOWFLAKE_SECRET_NAME = "regulations_gov_api_key"

#: Maximum page size supported by the Regulations.gov API.
DEFAULT_PAGE_SIZE = 250

#: Safety cap on paginated requests (DEFAULT_PAGE_SIZE * DEFAULT_MAX_PAGES max).
DEFAULT_MAX_PAGES = 20

#: Delay between paginated requests to respect the API rate limits.
DEFAULT_REQUEST_DELAY_SEC = 0.4

#: Known agency acronyms -> the ``filter[agencyId]`` value to send. The mapping
#: is permissive: any agency not listed here is passed through unchanged so the
#: package can be expanded to new agencies without a code change.
SUPPORTED_AGENCIES = {
    "CMS": "CMS",  # Centers for Medicare & Medicaid Services
}

#: The set of valid ``download_type`` values.
DOWNLOAD_TYPES = ("comments", "attachments", "metadata", "all")

#: The ordered core fields present on every comment record.
CORE_FIELDS = ("id", "agencyId", "postedDate", "title", "comment", "docketId",
               "attachmentCount")


# --------------------------------------------------------------------------- #
# API key resolution (secret handling)
# --------------------------------------------------------------------------- #

def resolve_api_key(
    api_key: Optional[str] = None,
    *,
    secret_name: str = SNOWFLAKE_SECRET_NAME,
) -> str:
    """Resolve the Regulations.gov API key from the first available source.

    Resolution order:

    1. The explicit ``api_key`` argument.
    2. The ``REGULATIONS_GOV_API_KEY`` environment variable.
    3. A Snowflake secret named ``secret_name`` (read via the ``_snowflake``
       module that is available inside Snowpark handlers).

    Raises:
        RuntimeError: if no key can be found. The error message never contains
            any key material.
    """
    if api_key:
        return api_key

    env_key = os.environ.get(API_KEY_ENV_VAR)
    if env_key:
        return env_key

    snowflake_key = _read_snowflake_secret(secret_name)
    if snowflake_key:
        return snowflake_key

    raise RuntimeError(
        "No Regulations.gov API key found. Provide it explicitly, set the "
        f"{API_KEY_ENV_VAR} environment variable, or configure the Snowflake "
        f"secret '{secret_name}' (see the README)."
    )


def _read_snowflake_secret(secret_name: str) -> Optional[str]:
    """Best-effort read of the API key from a Snowflake secret.

    Returns ``None`` when not running inside Snowflake or when the secret is
    not configured. Never raises so it can be used as a silent fallback.
    """
    try:
        import _snowflake  # type: ignore  # provided by the Snowflake runtime
    except Exception:
        return None
    try:
        value = _snowflake.get_generic_secret_string(secret_name)
        return value or None
    except Exception:
        return None


# --------------------------------------------------------------------------- #
# HTTP primitives (mockable in tests)
# --------------------------------------------------------------------------- #

def _http_get_json(url: str, headers: Dict[str, str]) -> Dict[str, Any]:
    """Perform an HTTP GET and decode the JSON response.

    This is the single network primitive; tests monkeypatch it so that no
    real network call is ever made. The API key travels in ``headers`` and is
    never included in exception messages raised here.
    """
    request = urllib.request.Request(url, headers=headers, method="GET")
    with urllib.request.urlopen(request) as response:  # noqa: S310 (trusted host)
        charset = response.headers.get_content_charset() or "utf-8"
        return json.loads(response.read().decode(charset))


def _api_get(path: str, params: Dict[str, Any], api_key: str) -> Dict[str, Any]:
    """Call a Regulations.gov endpoint and return the parsed JSON body.

    The API key is sent in the ``X-Api-Key`` header (never in the URL/query
    string) so it does not leak into logs or proxies that record URLs. This is
    an intentional change from the original script, which passed the key as an
    ``api_key`` query parameter.
    """
    query = urllib.parse.urlencode(params, doseq=True)
    url = f"{API_BASE_URL}/{path.lstrip('/')}"
    if query:
        url = f"{url}?{query}"
    headers = {"X-Api-Key": api_key, "Accept": "application/vnd.api+json"}
    try:
        return _http_get_json(url, headers)
    except urllib.error.HTTPError as exc:
        # Surface a friendly message. The key lives in a header, so neither the
        # URL nor this message can contain it.
        raise urllib.error.HTTPError(
            exc.url, exc.code, f"regulations.gov API error: {exc.reason}",
            exc.headers, exc.fp,
        )


# --------------------------------------------------------------------------- #
# Request construction
# --------------------------------------------------------------------------- #

def _normalize_agency(agency: str) -> str:
    """Map a known agency acronym to its ``filter[agencyId]`` value.

    Unknown agencies are passed through unchanged so new agencies work without
    a code change (the original script hard-failed on anything but CMS).
    """
    return SUPPORTED_AGENCIES.get(agency, agency)


def _validate_date(value: Optional[str], label: str) -> None:
    """Validate that an optional date string is ``YYYY-MM-DD``.

    ``None`` is allowed (the filter is simply omitted).
    """
    if value is None:
        return
    try:
        datetime.strptime(value, "%Y-%m-%d")
    except ValueError:
        raise ValueError(f"{label} must be in YYYY-MM-DD format, got {value!r}")


def build_filters(
    agency: str,
    *,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    keyword: Optional[str] = None,
    docket_id: Optional[str] = None,
) -> Dict[str, str]:
    """Build the ``filter[...]`` query parameters for a comments request.

    Date filters are only added when the corresponding argument is supplied,
    which is what makes ``start_date``/``end_date`` optional.
    """
    filters: Dict[str, str] = {"filter[agencyId]": _normalize_agency(agency)}
    if docket_id:
        filters["filter[docketId]"] = docket_id
    if keyword:
        filters["filter[searchTerm]"] = keyword
    if start_date:
        filters["filter[postedDate][ge]"] = start_date
    if end_date:
        filters["filter[postedDate][le]"] = end_date
    return filters


# --------------------------------------------------------------------------- #
# Record shaping
# --------------------------------------------------------------------------- #

def _summary_record(item: Dict[str, Any]) -> Dict[str, Any]:
    """Turn a list-endpoint item into a core (metadata-level) record."""
    attrs = item.get("attributes", {}) or {}
    return {
        "id": item.get("id"),
        "agencyId": attrs.get("agencyId"),
        "postedDate": attrs.get("postedDate"),
        "title": attrs.get("title"),
        # Keyword searches surface a snippet in highlightedContent even on the
        # list endpoint; fall back to it so metadata pulls aren't empty.
        "comment": attrs.get("highlightedContent") or "",
        "docketId": attrs.get("docketId"),
        "attachmentCount": attrs.get("attachmentCount", 0) or 0,
    }


def _extract_attachments(detail: Dict[str, Any], include_content: bool) -> List[Dict[str, Any]]:
    """Extract attachment metadata (and download URLs when requested)."""
    attachments: List[Dict[str, Any]] = []
    for inc in detail.get("included", []) or []:
        if inc.get("type") != "attachments":
            continue
        attrs = inc.get("attributes", {}) or {}
        for fmt in attrs.get("fileFormats", []) or []:
            att = {
                "title": attrs.get("title"),
                "format": fmt.get("format"),
                "size": fmt.get("size"),
            }
            if include_content:
                att["url"] = fmt.get("fileUrl")
            attachments.append(att)
    return attachments


# --------------------------------------------------------------------------- #
# Public API
# --------------------------------------------------------------------------- #

def get_comments(
    agency: str,
    *,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    api_key: Optional[str] = None,
    keyword: Optional[str] = None,
    include_attachments: bool = False,
    download_type: str = "all",
    docket_id: Optional[str] = None,
    max_records: Optional[int] = None,
    page_size: int = DEFAULT_PAGE_SIZE,
    max_pages: int = DEFAULT_MAX_PAGES,
    request_delay_sec: float = DEFAULT_REQUEST_DELAY_SEC,
) -> List[Dict[str, Any]]:
    """Retrieve comments from Regulations.gov.

    Args:
        agency: Agency acronym, e.g. ``"CMS"`` (required).
        start_date: Optional ``YYYY-MM-DD`` lower bound on ``postedDate``.
        end_date: Optional ``YYYY-MM-DD`` upper bound on ``postedDate``.
        api_key: The Regulations.gov API key. When ``None`` it is resolved via
            :func:`resolve_api_key` (env var / Snowflake secret).
        keyword: Optional free-text search term.
        include_attachments: When ``True`` and attachments are being pulled,
            include the attachment download URLs in each record.
        download_type: One of ``"comments"``, ``"attachments"``, ``"metadata"``
            or ``"all"``. Controls what is retrieved and returned:

            * ``"metadata"``  — identifiers/dates/docket only, no full body.
            * ``"comments"``  — comment body plus core fields.
            * ``"attachments"`` — attachment metadata (+ content URLs when
              ``include_attachments`` is set).
            * ``"all"``       — everything combined.
        docket_id: Optional docket identifier, e.g. ``"CMS-2022-0193"``. When
            given, results are filtered to that docket.
        max_records: Optional cap on the number of records returned.
        page_size: Results per page (API max is 250).
        max_pages: Safety cap on the number of pages fetched.
        request_delay_sec: Delay between paginated requests (rate limiting).

    Returns:
        A list of comment record dicts (see the module data dictionary).

    Raises:
        ValueError: if ``download_type`` is invalid or a date is malformed.
    """
    if download_type not in DOWNLOAD_TYPES:
        raise ValueError(
            f"download_type must be one of {DOWNLOAD_TYPES}, got {download_type!r}"
        )
    _validate_date(start_date, "start_date")
    _validate_date(end_date, "end_date")

    api_key = resolve_api_key(api_key)
    filters = build_filters(
        agency,
        start_date=start_date,
        end_date=end_date,
        keyword=keyword,
        docket_id=docket_id,
    )

    summaries = _paginate_comments(
        filters,
        api_key,
        max_records=max_records,
        page_size=page_size,
        max_pages=max_pages,
        request_delay_sec=request_delay_sec,
    )

    # A pure metadata pull needs no per-comment detail calls.
    if download_type == "metadata":
        return summaries

    want_body = download_type in ("comments", "all")
    want_attachments = download_type in ("attachments", "all")

    records: List[Dict[str, Any]] = []
    for record in summaries:
        detail = _fetch_detail(
            record["id"], api_key, include_attachments=want_attachments
        )
        attrs = detail.get("data", {}).get("attributes", {}) or {}
        if want_body:
            record["comment"] = (
                attrs.get("comment")
                or attrs.get("highlightedContent")
                or record.get("comment")
                or ""
            )
        if want_attachments:
            atts = _extract_attachments(detail, include_content=include_attachments)
            record["attachments"] = atts
            record["attachmentCount"] = len(atts)
        records.append(record)
    return records


def _paginate_comments(
    filters: Dict[str, str],
    api_key: str,
    *,
    max_records: Optional[int] = None,
    page_size: int = DEFAULT_PAGE_SIZE,
    max_pages: int = DEFAULT_MAX_PAGES,
    request_delay_sec: float = DEFAULT_REQUEST_DELAY_SEC,
) -> List[Dict[str, Any]]:
    """Page through the comments list endpoint and return core records."""
    records: List[Dict[str, Any]] = []
    page_number = 1
    while page_number <= max_pages:
        params = dict(filters)
        params["page[size]"] = page_size
        params["page[number]"] = page_number
        params["sort"] = "postedDate"
        body = _api_get("comments", params, api_key)

        data = body.get("data", []) or []
        if not data:
            break
        for item in data:
            records.append(_summary_record(item))
            if max_records is not None and len(records) >= max_records:
                return records[:max_records]

        # Stop on the last page: either the API says so, or the page was short.
        meta = body.get("meta", {}) or {}
        if not meta.get("hasNextPage") or len(data) < page_size:
            break
        page_number += 1
        if request_delay_sec:
            time.sleep(request_delay_sec)
    return records


def _fetch_detail(comment_id: str, api_key: str, *, include_attachments: bool) -> Dict[str, Any]:
    """Fetch the full detail document for a single comment."""
    params: Dict[str, Any] = {}
    if include_attachments:
        params["include"] = "attachments"
    return _api_get(f"comments/{comment_id}", params, api_key)


# --------------------------------------------------------------------------- #
# Metadata
# --------------------------------------------------------------------------- #

def build_metadata(
    *,
    agency: str,
    docket_id: Optional[str] = None,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    keyword: Optional[str] = None,
    download_type: str = "all",
    records: Optional[List[Dict[str, Any]]] = None,
) -> Dict[str, Any]:
    """Build a metadata record describing a pull.

    The API key is deliberately *not* a parameter and never appears here.
    """
    return {
        "agency": agency,
        "docket_id": docket_id,
        "start_date": start_date,
        "end_date": end_date,
        "keyword": keyword,
        "download_type": download_type,
        "record_count": len(records) if records is not None else 0,
        "retrieved_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "source": API_BASE_URL,
        "api_version": API_VERSION,
    }


# --------------------------------------------------------------------------- #
# Output writers (stdlib only)
# --------------------------------------------------------------------------- #

def write_json(records: List[Dict[str, Any]], filename: str) -> None:
    """Write records as a JSON array."""
    with open(filename, "w", encoding="utf-8") as fh:
        json.dump(records, fh, indent=2, ensure_ascii=False)


def write_csv(records: List[Dict[str, Any]], filename: str) -> None:
    """Write records as CSV using the core fields as columns."""
    fieldnames = list(CORE_FIELDS)
    with open(filename, "w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        for record in records:
            writer.writerow({key: record.get(key, "") for key in fieldnames})


def write_xml(records: List[Dict[str, Any]], filename: str) -> None:
    """Write records as XML: a ``<comments>`` root with ``<comment>`` children."""
    root = ET.Element("comments")
    for record in records:
        comment_el = ET.SubElement(root, "comment")
        for key in CORE_FIELDS:
            child = ET.SubElement(comment_el, key)
            value = record.get(key, "")
            child.text = "" if value is None else str(value)
    tree = ET.ElementTree(root)
    ET.indent(tree, space="  ")
    tree.write(filename, encoding="utf-8", xml_declaration=True)


#: Map of file extension -> writer function used by :func:`write_output`.
_WRITERS = {
    ".json": write_json,
    ".csv": write_csv,
    ".xml": write_xml,
}


def format_for_filename(filename: str) -> str:
    """Return the output format (``json``/``csv``/``xml``) implied by ``filename``."""
    ext = os.path.splitext(filename)[1].lower()
    if ext not in _WRITERS:
        raise ValueError(
            f"Unsupported output extension {ext!r}; expected one of "
            f"{sorted(_WRITERS)}"
        )
    return ext.lstrip(".")


def write_output(records: List[Dict[str, Any]], filename: str) -> None:
    """Dispatch to the correct writer based on the filename extension."""
    ext = os.path.splitext(filename)[1].lower()
    try:
        writer = _WRITERS[ext]
    except KeyError:
        raise ValueError(
            f"Unsupported output extension {ext!r}; expected one of "
            f"{sorted(_WRITERS)}"
        )
    writer(records, filename)


def write_metadata(meta: Dict[str, Any], filename: str) -> None:
    """Write a metadata record as a JSON sidecar file."""
    with open(filename, "w", encoding="utf-8") as fh:
        json.dump(meta, fh, indent=2, ensure_ascii=False)


def metadata_filename(output_filename: str) -> str:
    """Derive the metadata sidecar filename from an output filename.

    ``cms.csv`` -> ``cms.metadata.json``
    """
    stem = os.path.splitext(output_filename)[0]
    return f"{stem}.metadata.json"


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #

def _build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="python -m cgs_ai.regulations",
        description="Fetch public comments from Regulations.gov.",
    )
    parser.add_argument("--agency", required=True, help="Agency acronym, e.g. CMS")
    parser.add_argument("--docket-id", help="Docket id, e.g. CMS-2022-0193")
    parser.add_argument("--start-date", help="Lower bound postedDate (YYYY-MM-DD)")
    parser.add_argument("--end-date", help="Upper bound postedDate (YYYY-MM-DD)")
    parser.add_argument("--keyword", help="Free-text search term")
    parser.add_argument(
        "--download-type",
        choices=DOWNLOAD_TYPES,
        default="all",
        help="What to retrieve (default: all)",
    )
    parser.add_argument(
        "--include-attachments",
        action="store_true",
        help="Include attachment download URLs (with attachment pulls)",
    )
    parser.add_argument(
        "--api-key",
        help=(
            "Regulations.gov API key. If omitted, falls back to the "
            f"{API_KEY_ENV_VAR} environment variable (or a Snowflake secret)."
        ),
    )
    parser.add_argument(
        "--output",
        required=True,
        help="Output file; format inferred from extension (.json/.csv/.xml)",
    )
    parser.add_argument(
        "--format",
        choices=sorted(w.lstrip(".") for w in _WRITERS),
        help="Override the output format inferred from --output",
    )
    parser.add_argument(
        "--max-records",
        type=int,
        help="Optional cap on the number of records retrieved",
    )
    return parser


def main(argv: Optional[List[str]] = None) -> int:
    """Command-line entry point."""
    args = _build_arg_parser().parse_args(argv)

    # Resolve the key without ever echoing it.
    api_key = resolve_api_key(args.api_key)

    records = get_comments(
        agency=args.agency,
        start_date=args.start_date,
        end_date=args.end_date,
        api_key=api_key,
        keyword=args.keyword,
        include_attachments=args.include_attachments,
        download_type=args.download_type,
        docket_id=args.docket_id,
        max_records=args.max_records,
    )

    # Honour an explicit --format override by adjusting the output extension.
    output = args.output
    if args.format:
        stem = os.path.splitext(output)[0]
        output = f"{stem}.{args.format}"

    write_output(records, output)

    meta = build_metadata(
        agency=args.agency,
        docket_id=args.docket_id,
        start_date=args.start_date,
        end_date=args.end_date,
        keyword=args.keyword,
        download_type=args.download_type,
        records=records,
    )
    meta_path = metadata_filename(output)
    write_metadata(meta, meta_path)

    print(f"Wrote {len(records)} record(s) to {output}")
    print(f"Wrote metadata to {meta_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
