"""Type stubs for the native Rust extension module."""

import datetime
from typing import Any

def version() -> str:
    """Return the version string of the native core."""
    ...

def read_xlsx(path: str) -> list[dict[str, Any]]:
    """Read an XLSX file and return a list of sheet dicts.

    Each dict has keys:
      - ``"name"``: sheet name (str)
      - ``"rows"``: list of lists of cell values
      - ``"merges"``: list of merged cell range strings (e.g. ``"A1:C1"``)
      - ``"column_widths"``: dict mapping 0-based column index to width in character units
      - ``"row_heights"``: dict mapping 0-based row index to height in points
    """
    ...

def read_sheet(
    path: str,
    sheet_name: str | None = None,
    sheet_index: int | None = None,
) -> list[list[str | int | float | bool | datetime.date | datetime.datetime | Formula | None]]:
    """Read a single sheet by name or index.

    Returns the first sheet by default.
    """
    ...

def sheet_names(path: str) -> list[str]:
    """Return the list of sheet names in a workbook."""
    ...

class XlsxWriter:
    """Streaming XLSX writer.

    Use as a context manager::

        with XlsxWriter("output.xlsx") as writer:
            writer.add_sheet("Sheet1")
            writer.write_row(["Name", "Age"])
    """

    def __init__(self, path: str) -> None: ...
    def add_sheet(self, name: str) -> None:
        """Create a new worksheet."""
        ...
    def write_row(
        self,
        values: list[str | int | float | bool | datetime.date | datetime.datetime | Formula | None],
    ) -> None:
        """Write a row of values to the current sheet."""
        ...
    def merge_cells(self, range: str) -> None:
        """Merge a range of cells (e.g. ``"A1:C1"``)."""
        ...
    def set_column_width(self, column: str | int, width: float) -> None:
        """Set the width of a column in character units.

        ``column`` can be a letter (e.g. ``"A"``, ``"AA"``) or a 0-based integer index.
        Must be called after ``add_sheet()`` but before any ``write_row()`` calls on that sheet.
        """
        ...
    def set_row_height(self, row: int, height: float) -> None:
        """Set the height of a row in points.

        ``row`` is a 1-based row number (matching Excel convention).
        """
        ...
    def close(self) -> None:
        """Finalize and close the XLSX file."""
        ...
    def __enter__(self) -> XlsxWriter: ...
    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_val: BaseException | None,
        exc_tb: Any | None,
    ) -> bool: ...

class Formula:
    """A spreadsheet formula with optional cached value.

    Args:
        formula: The formula string (e.g. ``"SUM(A1:A10)"``).
        cached_value: Optional pre-computed value for the formula.
    """

    formula: str
    cached_value: Any | None

    def __init__(self, formula: str, cached_value: Any | None = None) -> None: ...
    def __repr__(self) -> str: ...
    def __eq__(self, other: object) -> bool: ...
