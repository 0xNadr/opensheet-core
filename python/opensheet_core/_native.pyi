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
      - ``"freeze_pane"``: tuple of (rows_frozen, cols_frozen) or None
      - ``"auto_filter"``: auto-filter range string (e.g. ``"A1:C1"``) or None
    """
    ...

def read_sheet(
    path: str,
    sheet_name: str | None = None,
    sheet_index: int | None = None,
) -> list[list[str | int | float | bool | datetime.date | datetime.datetime | Formula | FormattedCell | StyledCell | None]]:
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
        values: list[str | int | float | bool | datetime.date | datetime.datetime | Formula | FormattedCell | StyledCell | None],
    ) -> None:
        """Write a row of values to the current sheet."""
        ...
    def write_rows(
        self,
        rows: list[list[str | int | float | bool | datetime.date | datetime.datetime | Formula | FormattedCell | StyledCell | None]],
    ) -> None:
        """Write multiple rows at once, minimizing FFI overhead.

        Each element of ``rows`` should be a list of cell values.
        Equivalent to calling ``write_row()`` for each row, but faster
        because it crosses the Python→Rust boundary only once.
        """
        ...
    def merge_cells(self, range: str) -> None:
        """Merge a range of cells (e.g. ``"A1:C1"``)."""
        ...
    def auto_filter(self, range: str) -> None:
        """Set an auto-filter on a range (e.g. ``"A1:C1"``)."""
        ...
    def freeze_panes(self, row: int = 0, col: int = 0) -> None:
        """Freeze the top ``row`` rows and left ``col`` columns.

        Must be called after ``add_sheet()`` but before any ``write_row()`` calls on that sheet.
        """
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

class FormattedCell:
    """A cell value with a custom number format.

    Args:
        value: The numeric value.
        number_format: Excel number format code (e.g. ``"$#,##0.00"``, ``"0.00%"``).
    """

    value: Any
    number_format: str

    def __init__(self, value: Any, number_format: str) -> None: ...
    def __repr__(self) -> str: ...
    def __eq__(self, other: object) -> bool: ...

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

class CellStyle:
    """Cell style properties for fonts, fills, borders, and alignment.

    All parameters are keyword-only. The ``border`` parameter is a shorthand
    that sets all four border sides at once.

    Args:
        bold: Bold font.
        italic: Italic font.
        underline: Underlined font.
        font_name: Font family name (e.g. ``"Arial"``).
        font_size: Font size in points (e.g. ``12.0``).
        font_color: Font color as hex string (``"RRGGBB"`` or ``"AARRGGBB"``).
        fill_color: Solid fill color as hex string.
        border: Shorthand to set all four border sides (e.g. ``"thin"``).
        border_left: Left border style (e.g. ``"thin"``, ``"medium"``, ``"thick"``).
        border_right: Right border style.
        border_top: Top border style.
        border_bottom: Bottom border style.
        border_color: Border color as hex string (shared for all sides).
        horizontal_alignment: Horizontal alignment (``"left"``, ``"center"``, ``"right"``).
        vertical_alignment: Vertical alignment (``"top"``, ``"center"``, ``"bottom"``).
        wrap_text: Enable text wrapping.
        text_rotation: Text rotation in degrees (0-180).
        number_format: Excel number format code (e.g. ``"$#,##0.00"``).
    """

    bold: bool
    italic: bool
    underline: bool
    font_name: str | None
    font_size: float | None
    font_color: str | None
    fill_color: str | None
    border_left: str | None
    border_right: str | None
    border_top: str | None
    border_bottom: str | None
    border_color: str | None
    horizontal_alignment: str | None
    vertical_alignment: str | None
    wrap_text: bool
    text_rotation: int | None
    number_format: str | None

    def __init__(
        self,
        *,
        bold: bool = False,
        italic: bool = False,
        underline: bool = False,
        font_name: str | None = None,
        font_size: float | None = None,
        font_color: str | None = None,
        fill_color: str | None = None,
        border: str | None = None,
        border_left: str | None = None,
        border_right: str | None = None,
        border_top: str | None = None,
        border_bottom: str | None = None,
        border_color: str | None = None,
        horizontal_alignment: str | None = None,
        vertical_alignment: str | None = None,
        wrap_text: bool = False,
        text_rotation: int | None = None,
        number_format: str | None = None,
    ) -> None: ...
    def __repr__(self) -> str: ...
    def __eq__(self, other: object) -> bool: ...

class StyledCell:
    """A cell value with styling (fonts, fills, borders, alignment).

    Args:
        value: The cell value (str, int, float, bool, date, datetime, Formula, or None).
        style: A :class:`CellStyle` instance.
    """

    value: Any
    style: CellStyle

    def __init__(self, value: Any, style: CellStyle) -> None: ...
    def __repr__(self) -> str: ...
    def __eq__(self, other: object) -> bool: ...
