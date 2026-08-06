"""Tests for cgs_ai.regulations.

All HTTP is mocked; these tests never touch the network.
"""

import csv
import json
import os
import xml.etree.ElementTree as ET

import pytest

from cgs_ai import regulations as reg

SECRET = "SECRET-API-KEY-should-never-appear"

FIXTURE = [
    {
        "id": "CMS-2022-0193-0001",
        "agencyId": "CMS",
        "postedDate": "2023-01-15",
        "title": "Comment on Prior Authorization",
        "comment": "I support this rule.",
        "docketId": "CMS-2022-0193",
        "attachmentCount": 1,
        "attachments": [{"title": "letter", "format": "pdf", "url": "https://x/att.pdf"}],
    },
    {
        "id": "CMS-2022-0193-0002",
        "agencyId": "CMS",
        "postedDate": "2023-01-16",
        "title": "Another comment",
        "comment": "Please reconsider, ünïcode ok.",
        "docketId": "CMS-2022-0193",
        "attachmentCount": 0,
    },
]


# --------------------------------------------------------------------------- #
# filename -> format dispatch
# --------------------------------------------------------------------------- #

@pytest.mark.parametrize(
    "filename,expected",
    [("a.json", "json"), ("a.csv", "csv"), ("a.xml", "xml"),
     ("dir/sub/report.JSON", "json")],
)
def test_format_for_filename(filename, expected):
    assert reg.format_for_filename(filename) == expected


def test_format_for_filename_rejects_unknown():
    with pytest.raises(ValueError):
        reg.format_for_filename("data.txt")


def test_write_output_dispatches_by_extension(tmp_path):
    for ext, loader in (("json", None), ("csv", None), ("xml", None)):
        path = tmp_path / f"out.{ext}"
        reg.write_output(FIXTURE, str(path))
        assert path.exists()


def test_write_output_rejects_unknown_extension(tmp_path):
    with pytest.raises(ValueError):
        reg.write_output(FIXTURE, str(tmp_path / "out.txt"))


# --------------------------------------------------------------------------- #
# writer round-trips
# --------------------------------------------------------------------------- #

def test_write_json_round_trip(tmp_path):
    path = tmp_path / "out.json"
    reg.write_json(FIXTURE, str(path))
    loaded = json.loads(path.read_text(encoding="utf-8"))
    assert loaded == FIXTURE


def test_write_csv_round_trip(tmp_path):
    path = tmp_path / "out.csv"
    reg.write_csv(FIXTURE, str(path))
    with open(path, encoding="utf-8", newline="") as fh:
        rows = list(csv.DictReader(fh))
    assert [r["id"] for r in rows] == [r["id"] for r in FIXTURE]
    assert rows[0]["comment"] == "I support this rule."
    assert set(rows[0].keys()) == set(reg.CORE_FIELDS)


def test_write_xml_round_trip(tmp_path):
    path = tmp_path / "out.xml"
    reg.write_xml(FIXTURE, str(path))
    root = ET.parse(path).getroot()
    assert root.tag == "comments"
    comments = root.findall("comment")
    assert len(comments) == len(FIXTURE)
    assert comments[0].findtext("id") == "CMS-2022-0193-0001"
    assert comments[1].findtext("comment") == "Please reconsider, ünïcode ok."


# --------------------------------------------------------------------------- #
# request / filter construction
# --------------------------------------------------------------------------- #

def test_optional_dates_omitted_when_absent():
    filters = reg.build_filters("CMS")
    assert filters == {"filter[agencyId]": "CMS"}
    assert "filter[postedDate][ge]" not in filters
    assert "filter[postedDate][le]" not in filters


def test_dates_included_when_provided():
    filters = reg.build_filters("CMS", start_date="2023-01-01", end_date="2023-02-01")
    assert filters["filter[postedDate][ge]"] == "2023-01-01"
    assert filters["filter[postedDate][le]"] == "2023-02-01"


def test_docket_id_filter_included():
    filters = reg.build_filters("CMS", docket_id="CMS-2022-0193")
    assert filters["filter[docketId]"] == "CMS-2022-0193"


def test_keyword_filter_included():
    filters = reg.build_filters("CMS", keyword="prior authorization")
    assert filters["filter[searchTerm]"] == "prior authorization"


def test_agency_passthrough_for_unknown_agency():
    # Original hard-failed on non-CMS; refactor passes unknown agencies through.
    filters = reg.build_filters("EPA")
    assert filters["filter[agencyId]"] == "EPA"


def test_invalid_date_raises(monkeypatch):
    _fake_api(monkeypatch)
    with pytest.raises(ValueError):
        reg.get_comments("CMS", api_key=SECRET, start_date="2023/01/01")


def test_valid_optional_date_accepted(monkeypatch):
    _fake_api(monkeypatch)
    # Well-formed dates should not raise.
    reg.get_comments("CMS", api_key=SECRET, download_type="metadata",
                     start_date="2023-01-01", end_date="2023-02-01")


def test_pagination_stops_on_short_page(monkeypatch):
    """A page shorter than page_size ends pagination (original behavior)."""
    calls = {"n": 0}

    def fake_api_get(path, params, api_key):
        calls["n"] += 1
        # One record on the only page; page_size is 250 so it's a short page.
        return {"data": [{"id": "X-1", "attributes": {"agencyId": "CMS"}}],
                "meta": {"hasNextPage": True}}

    monkeypatch.setattr(reg, "_api_get", fake_api_get)
    records = reg.get_comments("CMS", api_key=SECRET, download_type="metadata")
    assert len(records) == 1
    assert calls["n"] == 1  # did not request a second page


# --------------------------------------------------------------------------- #
# get_comments with mocked HTTP
# --------------------------------------------------------------------------- #

def _fake_api(monkeypatch, capture=None):
    """Install a fake _api_get that returns canned list + detail responses."""
    list_body = {
        "data": [
            {"id": "CMS-2022-0193-0001", "attributes": {
                "agencyId": "CMS", "postedDate": "2023-01-15",
                "title": "T1", "docketId": "CMS-2022-0193"}},
            {"id": "CMS-2022-0193-0002", "attributes": {
                "agencyId": "CMS", "postedDate": "2023-01-16",
                "title": "T2", "docketId": "CMS-2022-0193"}},
        ],
        "meta": {"hasNextPage": False},
    }

    def fake_api_get(path, params, api_key):
        if capture is not None:
            capture.append((path, dict(params)))
        if path == "comments":
            return list_body
        # detail endpoint
        cid = path.split("/", 1)[1]
        return {
            "data": {"id": cid, "attributes": {"comment": f"body of {cid}"}},
            "included": [
                {"type": "attachments", "attributes": {
                    "title": "att", "fileFormats": [
                        {"format": "pdf", "size": 10, "fileUrl": "https://x/a.pdf"}]}}
            ],
        }

    monkeypatch.setattr(reg, "_api_get", fake_api_get)


def test_get_comments_metadata_skips_detail(monkeypatch):
    capture = []
    _fake_api(monkeypatch, capture)
    records = reg.get_comments("CMS", api_key=SECRET, download_type="metadata",
                               docket_id="CMS-2022-0193")
    assert len(records) == 2
    assert all(r["comment"] == "" for r in records)
    # Only the list endpoint should have been called.
    assert all(path == "comments" for path, _ in capture)


def test_get_comments_all_includes_body_and_attachments(monkeypatch):
    _fake_api(monkeypatch)
    records = reg.get_comments("CMS", api_key=SECRET, download_type="all",
                               include_attachments=True, docket_id="CMS-2022-0193")
    assert records[0]["comment"] == "body of CMS-2022-0193-0001"
    assert records[0]["attachmentCount"] == 1
    assert records[0]["attachments"][0]["url"] == "https://x/a.pdf"


def test_get_comments_passes_docket_filter(monkeypatch):
    capture = []
    _fake_api(monkeypatch, capture)
    reg.get_comments("CMS", api_key=SECRET, download_type="metadata",
                     docket_id="CMS-2022-0193")
    list_params = [p for path, p in capture if path == "comments"][0]
    assert list_params["filter[docketId]"] == "CMS-2022-0193"


def test_get_comments_rejects_bad_download_type(monkeypatch):
    _fake_api(monkeypatch)
    with pytest.raises(ValueError):
        reg.get_comments("CMS", api_key=SECRET, download_type="bogus")


# --------------------------------------------------------------------------- #
# API key resolution + secrecy
# --------------------------------------------------------------------------- #

def test_resolve_api_key_prefers_explicit(monkeypatch):
    monkeypatch.setenv(reg.API_KEY_ENV_VAR, "env-key")
    assert reg.resolve_api_key("explicit") == "explicit"


def test_resolve_api_key_env_fallback(monkeypatch):
    monkeypatch.setenv(reg.API_KEY_ENV_VAR, "env-key")
    assert reg.resolve_api_key(None) == "env-key"


def test_resolve_api_key_raises_without_source(monkeypatch):
    monkeypatch.delenv(reg.API_KEY_ENV_VAR, raising=False)
    monkeypatch.setattr(reg, "_read_snowflake_secret", lambda name: None)
    with pytest.raises(RuntimeError) as exc:
        reg.resolve_api_key(None)
    # The error must not leak any key material.
    assert "env-key" not in str(exc.value)


def test_build_metadata_excludes_api_key():
    meta = reg.build_metadata(agency="CMS", docket_id="CMS-2022-0193",
                              records=FIXTURE, download_type="all")
    assert "api_key" not in meta
    assert SECRET not in json.dumps(meta)
    assert meta["record_count"] == 2
    assert meta["source"] == reg.API_BASE_URL
    assert meta["api_version"] == reg.API_VERSION


def test_api_key_never_appears_in_any_output(tmp_path, monkeypatch):
    """End-to-end: the secret key must not leak into output or metadata files."""
    _fake_api(monkeypatch)
    records = reg.get_comments("CMS", api_key=SECRET, download_type="all",
                               docket_id="CMS-2022-0193")

    for ext in ("json", "csv", "xml"):
        out = tmp_path / f"out.{ext}"
        reg.write_output(records, str(out))
        assert SECRET not in out.read_text(encoding="utf-8")

    meta = reg.build_metadata(agency="CMS", docket_id="CMS-2022-0193",
                              records=records)
    meta_path = tmp_path / "out.metadata.json"
    reg.write_metadata(meta, str(meta_path))
    assert SECRET not in meta_path.read_text(encoding="utf-8")


def test_metadata_filename_derivation():
    assert reg.metadata_filename("a/b/out.csv") == "a/b/out.metadata.json"
