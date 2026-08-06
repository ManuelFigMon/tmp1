"""cgs_ai — CGS AI utilities designed to run inside Snowflake.

Standard-library-only building blocks that survive Snowflake's restricted
Python sandbox. The first module, :mod:`cgs_ai.regulations`, retrieves public
comments from the Regulations.gov API.
"""

from .regulations import (
    API_BASE_URL,
    API_VERSION,
    DOWNLOAD_TYPES,
    build_filters,
    build_metadata,
    format_for_filename,
    get_comments,
    metadata_filename,
    resolve_api_key,
    write_csv,
    write_json,
    write_metadata,
    write_output,
    write_xml,
)

__version__ = "0.1.0"

__all__ = [
    "__version__",
    "API_BASE_URL",
    "API_VERSION",
    "DOWNLOAD_TYPES",
    "build_filters",
    "build_metadata",
    "format_for_filename",
    "get_comments",
    "metadata_filename",
    "resolve_api_key",
    "write_csv",
    "write_json",
    "write_metadata",
    "write_output",
    "write_xml",
]
