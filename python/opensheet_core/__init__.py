"""OpenSheet Core - Fast, memory-efficient spreadsheet I/O for Python."""

from opensheet_core._native import (
    version,
    read_xlsx,
    read_sheet,
    sheet_names,
    defined_names,
    XlsxWriter,
    Formula,
    FormattedCell,
    CellStyle,
    StyledCell,
)
__version__ = version()
__all__ = [
    "__version__",
    "read_xlsx",
    "read_sheet",
    "sheet_names",
    "defined_names",
    "XlsxWriter",
    "Formula",
    "FormattedCell",
    "CellStyle",
    "StyledCell",
    "read_xlsx_df",
    "to_xlsx",
    "xlsx_to_markdown",
    "xlsx_to_text",
    "xlsx_to_chunks",
]


def read_xlsx_df(*args, **kwargs):
    """Read an XLSX sheet into a pandas DataFrame.

    Requires pandas to be installed.
    """
    from opensheet_core.pandas import read_xlsx_df as _read_xlsx_df
    return _read_xlsx_df(*args, **kwargs)


def to_xlsx(*args, **kwargs):
    """Write a pandas DataFrame to an XLSX file.

    Requires pandas to be installed.
    """
    from opensheet_core.pandas import to_xlsx as _to_xlsx
    return _to_xlsx(*args, **kwargs)


def xlsx_to_markdown(*args, **kwargs):
    """Convert an XLSX file to markdown table(s) for LLM consumption."""
    from opensheet_core.extract import xlsx_to_markdown as _fn
    return _fn(*args, **kwargs)


def xlsx_to_text(*args, **kwargs):
    """Convert an XLSX file to plain text for search indexes."""
    from opensheet_core.extract import xlsx_to_text as _fn
    return _fn(*args, **kwargs)


def xlsx_to_chunks(*args, **kwargs):
    """Convert an XLSX file to embedding-sized markdown chunks for RAG."""
    from opensheet_core.extract import xlsx_to_chunks as _fn
    return _fn(*args, **kwargs)
