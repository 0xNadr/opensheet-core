"""AI/RAG extraction functions for OpenSheet Core.

Provides xlsx_to_markdown(), xlsx_to_text(), and xlsx_to_chunks() for
converting spreadsheets into LLM-friendly formats.
"""

import datetime
import math

from opensheet_core._native import (
    read_xlsx,
    read_sheet,
    Formula,
    FormattedCell,
    StyledCell,
)


def _unwrap_cell(val):
    """Unwrap wrapper types to their plain display value."""
    if isinstance(val, StyledCell):
        return _unwrap_cell(val.value)
    if isinstance(val, Formula):
        return _unwrap_cell(val.cached_value) if val.cached_value is not None else ""
    if isinstance(val, FormattedCell):
        return _unwrap_cell(val.value)
    return val


def _cell_to_str(val):
    """Convert a cell value to its string representation."""
    val = _unwrap_cell(val)
    if val is None:
        return ""
    if isinstance(val, bool):
        return str(val)
    if isinstance(val, datetime.datetime):
        return val.isoformat()
    if isinstance(val, datetime.date):
        return val.isoformat()
    if isinstance(val, float):
        if math.isinf(val) or math.isnan(val):
            return str(val)
        # Drop trailing .0 for whole numbers
        if val == int(val):
            return str(int(val))
        return str(val)
    s = str(val)
    # Replace newlines (common in multi-line Excel cells) with spaces
    # to preserve the one-row-per-line invariant in all output formats.
    if "\n" in s or "\r" in s:
        s = s.replace("\r\n", " ").replace("\r", " ").replace("\n", " ")
    return s


def _pad_rows(rows, ncols):
    """Pad all rows to ncols width."""
    return [row + [None] * (ncols - len(row)) for row in rows]


def _max_cols(rows):
    """Find the maximum column count across all rows."""
    return max((len(row) for row in rows), default=0)


def _rows_to_markdown(rows, header=True):
    """Convert a list of rows to a markdown table string.

    Args:
        rows: List of lists of cell values.
        header: If True, treat the first row as a header row.

    Returns:
        A markdown table string.
    """
    if not rows:
        return ""

    ncols = _max_cols(rows)
    rows = _pad_rows(rows, ncols)

    # Convert all cells to strings, escaping pipes for markdown
    str_rows = [[_cell_to_str(cell).replace("\\", "\\\\").replace("|", "\\|") for cell in row] for row in rows]

    # Calculate column widths for alignment
    col_widths = [0] * ncols
    for row in str_rows:
        for i, cell in enumerate(row):
            col_widths[i] = max(col_widths[i], len(cell))
    # Minimum width of 3 for separator dashes
    col_widths = [max(w, 3) for w in col_widths]

    lines = []
    if header and len(str_rows) >= 1:
        # Header row
        header_cells = [cell.ljust(col_widths[i]) for i, cell in enumerate(str_rows[0])]
        lines.append("| " + " | ".join(header_cells) + " |")
        # Separator
        lines.append("| " + " | ".join("-" * w for w in col_widths) + " |")
        # Data rows
        data_start = 1
    else:
        # No header — still need a separator for valid markdown table
        # Use column indices as header
        header_cells = [f"Col {i}".ljust(col_widths[i]) for i in range(ncols)]
        lines.append("| " + " | ".join(header_cells) + " |")
        lines.append("| " + " | ".join("-" * w for w in col_widths) + " |")
        data_start = 0

    for row in str_rows[data_start:]:
        cells = [cell.ljust(col_widths[i]) for i, cell in enumerate(row)]
        lines.append("| " + " | ".join(cells) + " |")

    return "\n".join(lines)


def xlsx_to_markdown(path, sheet_name=None, sheet_index=None, header=True):
    """Convert an XLSX file to markdown table(s).

    Args:
        path: Path to the XLSX file.
        sheet_name: Name of a specific sheet to convert.
        sheet_index: 0-based index of a specific sheet to convert.
        header: If True (default), treat the first row of each sheet as a
            header row for the markdown table.

    Returns:
        A string containing one or more markdown tables. When multiple sheets
        are converted, each table is preceded by a ``## Sheet Name`` heading.
    """
    if sheet_name is not None or sheet_index is not None:
        rows = read_sheet(path, sheet_name=sheet_name, sheet_index=sheet_index)
        return _rows_to_markdown(rows, header=header)

    sheets = read_xlsx(path)
    if len(sheets) == 1:
        return _rows_to_markdown(sheets[0]["rows"], header=header)

    parts = []
    for sheet in sheets:
        table = _rows_to_markdown(sheet["rows"], header=header)
        if table:
            parts.append(f"## {sheet['name']}\n\n{table}")
    return "\n\n".join(parts)


def xlsx_to_text(path, sheet_name=None, sheet_index=None, delimiter="\t"):
    """Convert an XLSX file to plain text.

    Each row becomes a line with cells separated by the delimiter.
    Suitable for search indexes and simple text pipelines.

    Args:
        path: Path to the XLSX file.
        sheet_name: Name of a specific sheet to convert.
        sheet_index: 0-based index of a specific sheet to convert.
        delimiter: Cell separator (default: tab character).

    Returns:
        A plain text string with one row per line.
    """
    def _rows_to_lines(rows):
        ncols = _max_cols(rows) if rows else 0
        padded = _pad_rows(rows, ncols)
        return [delimiter.join(_cell_to_str(cell) for cell in row) for row in padded]

    if sheet_name is not None or sheet_index is not None:
        rows = read_sheet(path, sheet_name=sheet_name, sheet_index=sheet_index)
        return "\n".join(_rows_to_lines(rows))

    sheets = read_xlsx(path)
    parts = []
    for sheet in sheets:
        if len(sheets) > 1:
            parts.append(f"--- {sheet['name']} ---")
        parts.extend(_rows_to_lines(sheet["rows"]))
    return "\n".join(parts)


def xlsx_to_chunks(
    path,
    sheet_name=None,
    sheet_index=None,
    max_rows=50,
    header=True,
):
    """Convert an XLSX file to embedding-sized markdown chunks.

    Splits each sheet into chunks of at most ``max_rows`` data rows.
    When ``header=True``, the header row is repeated at the top of each chunk
    so that every chunk is a self-contained markdown table.

    Args:
        path: Path to the XLSX file.
        sheet_name: Name of a specific sheet to convert.
        sheet_index: 0-based index of a specific sheet to convert.
        max_rows: Maximum number of data rows per chunk (default: 50).
        header: If True (default), attach the header row to each chunk.

    Returns:
        A list of markdown table strings, each containing at most ``max_rows``
        data rows.
    """
    if max_rows < 1:
        raise ValueError("max_rows must be at least 1")

    def _chunk_sheet(rows, sheet_label=None):
        if not rows:
            return []

        if header and len(rows) >= 1:
            header_row = [rows[0]]
            data_rows = rows[1:]
        else:
            header_row = []
            data_rows = rows

        chunks = []
        for i in range(0, max(len(data_rows), 1), max_rows):
            batch = data_rows[i : i + max_rows]
            if not batch:
                continue
            chunk_rows = header_row + batch
            table = _rows_to_markdown(chunk_rows, header=header)
            if sheet_label:
                table = f"## {sheet_label}\n\n{table}"
            chunks.append(table)
        return chunks

    if sheet_name is not None or sheet_index is not None:
        rows = read_sheet(path, sheet_name=sheet_name, sheet_index=sheet_index)
        return _chunk_sheet(rows)

    sheets = read_xlsx(path)
    multi = len(sheets) > 1
    all_chunks = []
    for sheet in sheets:
        label = sheet["name"] if multi else None
        all_chunks.extend(_chunk_sheet(sheet["rows"], sheet_label=label))
    return all_chunks
