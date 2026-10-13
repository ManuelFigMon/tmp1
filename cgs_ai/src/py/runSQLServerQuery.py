"""
=====================================================================
  Program Name  : runSQLServerQuery.py
  Author        : Manuel Figallo
  Purpose       : Run a SQL Server query against a LOB catalog using
                  Windows Integrated Security and return the result set.
  Version       : 1.0beta
  Created       : 2026-08-26
  Last Modified : 2026-08-26

  Dependencies:
    pyodbc, imported lazily. The PowerShell twin needs no module at all
    (it uses the built-in .NET System.Data.OleDb), so prefer the PowerShell
    version on a locked-down server where pip installs are not permitted.

  Description:
    Python equivalent of this SAS pass-through block:

        proc sql noprint;
          connect to oledb as myconn
          (init_string="Provider=MSOLEDBSQL19;
           Integrated Security=SSPI;
           Persist Security Info=True;
           Initial Catalog=&LOB_Catalog;
           Data Source=my.dbserver.CCOM, 1433");
          create table work.dmq_result as
            select * from connection to myconn (&SQL_Statement);
          disconnect from myconn;
        quit;

    Integrated Security is used, so NO password is ever handled, stored or
    logged by this function.

  Input Parameters (required first):
    SQL_Statement (REQUIRED, str) - the query to execute.
    LOB_Catalog   (REQUIRED, str) - Initial Catalog, e.g. DataMartKYA.
    DataSource    (optional, str) - server,port; defaults to SQL_DATA_SOURCE
                                    in .env, else my.dbserver.CCOM,1433.
    OutputCsvPath (optional, str) - when given, the rows are also written here.
=====================================================================
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

_HERE = Path(__file__).resolve()
if str(_HERE.parent.parent.parent) not in sys.path:
    sys.path.insert(0, str(_HERE.parent.parent.parent))

from src.utils.config import getConfig            # noqa: E402
from src.utils.helpers import writeCsv            # noqa: E402
from src.utils.logger import logInfo, logWarn     # noqa: E402

__version__ = "1.0beta"

DEFAULT_DATA_SOURCE = "my.dbserver.CCOM,1433"
DEFAULT_PROVIDER = "MSOLEDBSQL19"
DEFAULT_ODBC_DRIVER = "ODBC Driver 18 for SQL Server"


def buildConnectionString(LOB_Catalog: str, DataSource: str,
                          Driver: str = DEFAULT_ODBC_DRIVER,
                          TrustServerCertificate: bool = True) -> str:
    """Build the ODBC connection string (Integrated Security, no password).

    Parameters:
        LOB_Catalog (str)  - Initial Catalog / database name.
        DataSource (str)   - "server,port".
        Driver (str)       - ODBC driver name.
        TrustServerCertificate (bool) - set for self-signed server certs.
    Returns:
        str connection string. Contains no credentials -- safe to log.
    """
    parts = [f"DRIVER={{{Driver}}}", f"SERVER={DataSource}",
             f"DATABASE={LOB_Catalog}", "Trusted_Connection=yes"]
    if TrustServerCertificate:
        parts.append("TrustServerCertificate=yes")
    return ";".join(parts) + ";"


def runSQLServerQuery(SQL_Statement: str, LOB_Catalog: str,
                      DataSource: Optional[str] = None,
                      OutputCsvPath: Optional[str] = None,
                      Driver: str = DEFAULT_ODBC_DRIVER,
                      TimeoutSeconds: int = 300) -> Dict[str, Any]:
    """Execute a SQL Server query and return its rows.

    Parameters:
        SQL_Statement (str)  - REQUIRED query text, exactly what the SAS
                               version passes as &query.
        LOB_Catalog (str)    - REQUIRED Initial Catalog, e.g. DataMartKYA.
        DataSource (str)     - "server,port"; defaults to SQL_DATA_SOURCE
                               from .env, else my.dbserver.CCOM,1433.
        OutputCsvPath (str)  - optional CSV to also write the rows to.
        Driver (str)         - ODBC driver name.
        TimeoutSeconds (int) - query timeout.
    Returns:
        dict with Rows (list[dict]), Columns (list[str]), RowCount,
        LOB_Catalog, DataSource and OutputCsvPath.
    Raises:
        ValueError  - a required parameter is missing.
        ImportError - pyodbc is not installed; message names the pip command
                      and points at the PowerShell twin as the no-install
                      alternative.
        Exception   - driver/connection errors propagate from pyodbc.

    Use in claims processing:
        Pull a claims or denial extract straight from the LOB data mart into
        Python or CSV, using the same catalog and credentials the SAS
        pass-through block uses, with no password anywhere in the code.
    """
    for name, value in (("SQL_Statement", SQL_Statement),
                        ("LOB_Catalog", LOB_Catalog)):
        if not value or not str(value).strip():
            raise ValueError(f"required parameter '{name}' is missing or empty")

    dataSource = DataSource or getConfig("SQL_DATA_SOURCE", DEFAULT_DATA_SOURCE)
    try:
        import pyodbc
    except ImportError:
        raise ImportError(
            "runSQLServerQuery requires pyodbc. Install it with "
            "'pip install pyodbc'. Alternatively use the PowerShell twin "
            "(src/ps/runSQLServerQuery.ps1), which needs no module because it "
            "uses the built-in .NET System.Data.OleDb provider.")

    connectionString = buildConnectionString(LOB_Catalog, dataSource, Driver)
    logInfo(f"connecting to catalog '{LOB_Catalog}' on {dataSource} "
            f"(Integrated Security)")

    rows: List[Dict[str, Any]] = []
    columns: List[str] = []
    with pyodbc.connect(connectionString, timeout=TimeoutSeconds) as connection:
        cursor = connection.cursor()
        cursor.execute(SQL_Statement)
        if cursor.description:
            columns = [c[0] for c in cursor.description]
            for record in cursor.fetchall():
                rows.append(dict(zip(columns, record)))
        else:
            logWarn("the statement returned no result set (not a SELECT?)")

    logInfo(f"returned {len(rows)} row(s) x {len(columns)} column(s)")
    written = None
    if OutputCsvPath:
        written = writeCsv(rows, columns, OutputCsvPath)
        logInfo(f"wrote result set to {written}")

    return {"Rows": rows, "Columns": columns, "RowCount": len(rows),
            "LOB_Catalog": LOB_Catalog, "DataSource": dataSource,
            "OutputCsvPath": written}
