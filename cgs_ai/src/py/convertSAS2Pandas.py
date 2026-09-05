"""
=====================================================================
  Program Name  : convertSAS2Pandas.py
  Author        : Manuel Figallo
  Purpose       : Read a SAS sas7bdat data set into a pandas DataFrame and
                  persist it in a portable format.
  Version       : 1.0beta
  Created       : 2026-08-26
  Last Modified : 2026-08-26

  Dependencies:
    pandas (required) plus a sas7bdat reader. pandas.read_sas covers this
    natively; pyreadstat is used when installed because it is faster and
    preserves column labels. Both are imported LAZILY, so importing cgs_ai
    does not require them -- only calling this function does.

    NOTE: this is the one function that cannot be standard-library only.
    A sas7bdat file is a proprietary binary format.

  Input Parameters (required first):
    InputSas7bdatPath (REQUIRED, str) - source .sas7bdat.
    OutputPath        (REQUIRED, str) - destination; the format follows the
                        extension: .parquet (default, needs pyarrow), .csv,
                        or .pkl.
=====================================================================
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any, Dict

_HERE = Path(__file__).resolve()
if str(_HERE.parent.parent.parent) not in sys.path:
    sys.path.insert(0, str(_HERE.parent.parent.parent))

from src.utils.helpers import ensureParent  # noqa: E402
from src.utils.logger import logInfo, logWarn  # noqa: E402

__version__ = "1.0beta"

SUPPORTED_OUTPUTS = (".parquet", ".csv", ".pkl")


def convertSAS2Pandas(InputSas7bdatPath: str, OutputPath: str,
                      Encoding: str = "latin-1") -> Dict[str, Any]:
    """Convert a sas7bdat data set to a pandas DataFrame and save it.

    Parameters:
        InputSas7bdatPath (str) - REQUIRED path to the .sas7bdat file.
        OutputPath (str)        - REQUIRED destination path. The extension
                                  selects the writer: .parquet (needs
                                  pyarrow), .csv, or .pkl.
        Encoding (str)          - character encoding of the SAS file;
                                  latin-1 suits most Windows SAS output.
    Returns:
        dict with OutputPath, RowCount, ColumnCount, Columns and DataFrame.
    Raises:
        ValueError  - a required parameter is missing or the output
                      extension is unsupported.
        ImportError - pandas (or pyarrow for .parquet) is not installed;
                      the message names the exact pip command.
        OSError     - the input file cannot be read.

    Use in claims processing:
        Bring a SAS claims extract into the Python/Snowflake side of the
        stack without a manual export step, so the same data can feed
        pandas analysis or a Snowflake load.
    """
    for name, value in (("InputSas7bdatPath", InputSas7bdatPath),
                        ("OutputPath", OutputPath)):
        if not value or not str(value).strip():
            raise ValueError(f"required parameter '{name}' is missing or empty")

    source = Path(InputSas7bdatPath)
    if not source.is_file():
        raise OSError(f"sas7bdat file not found: {source}")

    suffix = Path(OutputPath).suffix.lower()
    if suffix not in SUPPORTED_OUTPUTS:
        raise ValueError(f"unsupported output extension {suffix!r}; expected one "
                         f"of: {', '.join(SUPPORTED_OUTPUTS)}")
    try:
        import pandas as pd
    except ImportError:
        raise ImportError(
            "convertSAS2Pandas requires pandas. Install it with "
            "'pip install pandas'. (This is the one cgs_ai function that "
            "cannot be standard-library only: sas7bdat is a binary format.)")

    frame = None
    try:                                   # preferred: faster, keeps labels
        import pyreadstat
        frame, _meta = pyreadstat.read_sas7bdat(str(source))
        logInfo("read via pyreadstat")
    except ImportError:
        logWarn("pyreadstat not installed; falling back to pandas.read_sas")
        frame = pd.read_sas(str(source), format="sas7bdat", encoding=Encoding)

    ensureParent(OutputPath)
    if suffix == ".parquet":
        try:
            frame.to_parquet(OutputPath, index=False)
        except ImportError:
            raise ImportError(
                "Writing .parquet requires pyarrow. Install it with "
                "'pip install pyarrow', or choose a .csv output path.")
    elif suffix == ".csv":
        frame.to_csv(OutputPath, index=False)
    else:
        frame.to_pickle(OutputPath)

    logInfo(f"converted {len(frame)} row(s) x {len(frame.columns)} column(s) "
            f"-> {OutputPath}")
    return {"OutputPath": OutputPath, "RowCount": int(len(frame)),
            "ColumnCount": int(len(frame.columns)),
            "Columns": list(frame.columns), "DataFrame": frame}
