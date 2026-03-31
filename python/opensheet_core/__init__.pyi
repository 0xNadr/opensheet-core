"""Type stubs for the opensheet_core public API."""

import datetime
from typing import Any

from opensheet_core._native import (
    CellStyle as CellStyle,
    FormattedCell as FormattedCell,
    Formula as Formula,
    StyledCell as StyledCell,
    XlsxWriter as XlsxWriter,
    document_properties as document_properties,
    read_sheet as read_sheet,
    read_xlsx as read_xlsx,
    sheet_names as sheet_names,
    version as version,
)

__version__: str
__all__: list[str]

def read_xlsx_df(
    path: str,
    sheet_name: str | None = None,
    sheet_index: int | None = None,
    header: bool = True,
) -> Any:
    """Read an XLSX sheet into a pandas DataFrame."""
    ...

def to_xlsx(
    df: Any,
    path: str,
    sheet_name: str = "Sheet1",
    header: bool = True,
    index: bool = False,
) -> None:
    """Write a pandas DataFrame to an XLSX file."""
    ...

def xlsx_to_markdown(
    path: str,
    sheet_name: str | None = None,
    sheet_index: int | None = None,
    header: bool = True,
) -> str:
    """Convert an XLSX file to markdown table(s) for LLM consumption."""
    ...

def xlsx_to_text(
    path: str,
    sheet_name: str | None = None,
    sheet_index: int | None = None,
    delimiter: str = "\t",
) -> str:
    """Convert an XLSX file to plain text for search indexes."""
    ...

def xlsx_to_chunks(
    path: str,
    sheet_name: str | None = None,
    sheet_index: int | None = None,
    max_rows: int = 50,
    header: bool = True,
) -> list[str]:
    """Convert an XLSX file to embedding-sized markdown chunks for RAG."""
    ...
