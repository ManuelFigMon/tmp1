# cgs_ai

**Version 1.0beta** — Claims intelligence toolkit for SAS, Python and PowerShell.

> **Full documentation is in [`README.docx`](README.docx)**, which covers the
> folder structure and why it matters in data engineering, a reference for every
> function (parameters and claims-processing uses), the `.env` configuration
> rules, a Visual Studio Code tutorial, and the version-control table.
>
> **[`CGS_AI_Presentation.pptx`](CGS_AI_Presentation.pptx)** is the briefing deck
> ("From GUI to Code using CGS_AI"), regenerated with
> `python src/utils/build_presentation.py <template.pptx> CGS_AI_Presentation.pptx`.

## Quickstart

```bash
python -m venv .venv
.venv\Scripts\activate            # Linux/macOS: source .venv/bin/activate
pip install -r requirements.txt
copy .env.example .env             # then edit the paths
python Examples.py                 # confirms the package imports
```

```python
import cgs_ai
cgs_ai.scanFileSystem(input_folder_root=[r"\\srv\logs"],
                      extract_keyword=["real time", "cpu time"],
                      output_file_path="data/scan.csv",
                      metric_profile="sas_log")   # -> Excel + Metrics sheet
```

## Functions

| Function | Purpose |
|---|---|
| `scanFileSystem` | Keyword scan of directory roots; **one row per match** with context lines and extracted tokens |
| `runSQLServerQuery` | SQL Server query via Windows Integrated Security (no password) |
| `formatCSV` | CSV to styled Excel with a SAS ODS look and feel |
| `downloadBulkFiles` | Bulk HTTP download from a CSV link column |
| `sendEmail` | SMTP alert to one or many recipients |
| `convertSAS2Pandas` | sas7bdat to pandas DataFrame |
| `copyExcelSheet2CSV` | Excel worksheet to CSV, validated before writing |
| `collectSystemMetrics` | Host metrics to a CSV time series; fails gracefully |
| `zipFolder` | Folder plus extra files into one archive |
| `runFilescanPipeline` | scan → format → email, end to end |
| `get_comments` | Retrieve public comments from Regulations.gov |

Every function exists in **Python** (`src/py`), **PowerShell** (`src/ps`) and as a
**SAS wrapper macro** (`src/sas`), with identical function and parameter names.

## Dependencies

The package imports with **nothing beyond the Python standard library**. Optional
packages are imported lazily by the one function that needs them:
`openpyxl` (Excel), `pandas` (SAS conversion), `pyodbc` (SQL Server, Python only —
the PowerShell twin needs no module), `psutil` (richer host metrics).

## Tests

```bash
pytest -q
```
