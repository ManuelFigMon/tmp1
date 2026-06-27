# NLP Sentiment Dashboard for CMS Public Comments

A self-contained, **air-gapped** HTML dashboard for analyzing public comments on
proposed CMS (Centers for Medicare & Medicaid Services) rules, plus a
standard-library Python downloader for pulling those comments from
[regulations.gov](https://www.regulations.gov/).

Part of the DMQ report family — it reuses the same theme, configuration loading,
SAS session/auth, and report-save infrastructure as `correlation_matrix.html`.

## Files

| File | Description |
|---|---|
| `meta/profiles/default/reports/nlp_sentiment_dashboard.html` | The dashboard. D3 v7 is **fully inlined** — no CDN, no external/local script references except local `html2canvas.min.js`. Open it from within the DMQ `reports/` folder so it can load `data_config.json` / `datasets.json` / `sas_config.json` one level up. |
| `fetch_regulations_comments.py` | Standard-library-only downloader for regulations.gov comments (JSON or CSV output). |

## Dashboard features

**Left input panel**
- **Dataset Selection** — warns inline if the chosen dataset has no free-text column.
- **Column Selection** — a required single-select **Text Column** (free-text comment body), optional multi-select **Category Column(s)**, and an info-icon explainer. Text columns are detected with a heuristic: average string length > 40 characters across a sample of fetched rows.
- **Filters** — the standard DMQ WHERE-condition builder (column / operator / value).
- **Configuration Selection** — toggle the word cloud and summaries, pick the sentiment method (lexicon or heuristic), edit the comma-separated category list, set a minimum comment length, and cap word-cloud words.

**Right visualization panel** (all four sub-panels recompute once per `renderDashboard()` from a single `lastAnalysisResult` global)
1. **KPI strip** — total comments, % Supportive / Neutral / Opposed (green/gray/red), and the most-discussed category.
2. **Sentiment & Categorization Matrix (heatmap)** — categories (rows) × Supportive/Neutral/Opposed (columns). Per-column sequential color scales (Greens / Greys / Reds); the row-maximum cell in the **Opposed** column gets a 2px gold border to flag the dominant complaint driver. Hover for counts, % of row, and up to 3 example excerpts. Click a cell to cross-filter the Summaries panel. A `# Count / % of Row` toggle switches the cell labels.
3. **Word cloud** — hand-written spiral placement (no plugin), sized by term frequency after stopword removal, colored by each term's dominant associated sentiment.
4. **Comment summaries** — scrollable table of excerpt, sentiment badge, category, and confidence; supports the heatmap cross-filter with a "Clear cell filter" link.

**Action buttons** — Export Image (html2canvas), Export to CSV (per-comment: text, category, sentiment, confidence), Reset Form.

## How classification works (offline, no ML)

- **Sentiment** — a built-in lexicon of positive/negative healthcare & regulatory terms with negation handling (e.g., *"not beneficial"* flips a positive word). Net balance → Supportive / Neutral / Opposed, with a 0–100% confidence from the score margin.
- **Categories** — if you select real category columns, those values are used as heatmap rows. Otherwise a keyword rule set assigns each comment to the category whose terms appear most often (defaults tuned for Medicare/CMS rulemaking language: Healthcare Impact, Implementation Cost, Legal Authority, Administrative Burden, Beneficiary Access).

This is intentionally a transparent, auditable lexicon/keyword approach so it can run with **zero network access**. Always review the underlying comment text before drawing conclusions from any single cell.

## Python downloader

Standard library only (`urllib`, `json`, `csv`, `argparse`, `datetime`, `time`) — no `requests`, no `pandas`. Get a free API key at <https://open.gsa.gov/api/regulationsgov/>.

```bash
python fetch_regulations_comments.py \
    --agency CMS \
    --start-date 2026-01-01 \
    --end-date 2026-03-31 \
    --api-key YOUR_API_KEY_HERE \
    --keyword "telehealth reimbursement" \
    --output cms_comments.json \
    --format json
```

As a library:

```python
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
```

The output JSON (an array of `{id, agencyId, postedDate, title, comment, docketId, attachmentCount}`) can be wired into the dashboard as a `file`-type dataset via `data_config.json` / `datasets.json`, using a `comment` column as the dashboard's Text Column.

## Air-gapped guarantees

- D3 v7.9.0 is pasted inline into a `<script>` tag — the HTML makes **no** external network calls for libraries.
- The Python script uses only the standard library.
- The only runtime fetches the HTML performs are to the local DMQ config/data files (and, at the user's explicit action, the SAS logon endpoint) — never to a third-party CDN.

## License

MIT — see `LICENSE`.
