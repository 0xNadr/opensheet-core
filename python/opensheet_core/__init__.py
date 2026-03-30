"""OpenSheet Core - Fast, memory-efficient spreadsheet I/O for Python."""

from opensheet_core._native import (
    version,
    read_xlsx,
    read_sheet,
    sheet_names,
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
    "XlsxWriter",
    "Formula",
    "FormattedCell",
    "CellStyle",
    "StyledCell",
    "read_xlsx_df",
    "to_xlsx",
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
