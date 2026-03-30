"""Pandas integration for OpenSheet Core.

Provides read_xlsx_df() and to_xlsx() for DataFrame I/O.
Pandas is an optional dependency — these functions raise ImportError
if pandas is not installed.
"""

import datetime

from opensheet_core._native import (
    read_sheet,
    sheet_names as _sheet_names,
    XlsxWriter,
    Formula,
    FormattedCell,
    StyledCell,
)


def _check_pandas():
    try:
        import pandas as pd
        return pd
    except ImportError:
        raise ImportError(
            "pandas is required for this function. "
            "Install it with: pip install opensheet-core[pandas]"
        ) from None


def read_xlsx_df(
    path,
    sheet_name=None,
    sheet_index=None,
    header=True,
):
    """Read an XLSX sheet into a pandas DataFrame.

    Args:
        path: Path to the XLSX file.
        sheet_name: Name of the sheet to read. Reads the first sheet by default.
        sheet_index: 0-based index of the sheet to read.
        header: If True (default), use the first row as column names.

    Returns:
        A pandas DataFrame.
    """
    pd = _check_pandas()

    rows = read_sheet(path, sheet_name=sheet_name, sheet_index=sheet_index)

    if not rows:
        return pd.DataFrame()

    # Unwrap Formula/FormattedCell to plain values
    def _unwrap(val):
        if isinstance(val, StyledCell):
            return _unwrap(val.value)
        if isinstance(val, Formula):
            return val.cached_value
        if isinstance(val, FormattedCell):
            return val.value
        return val

    rows = [[_unwrap(cell) for cell in row] for row in rows]

    if header and len(rows) >= 1:
        columns = rows[0]
        data = rows[1:]
        # Pad short rows to match column count (reader trims trailing empties)
        ncols = len(columns)
        data = [row + [None] * (ncols - len(row)) for row in data]
        return pd.DataFrame(data, columns=columns)
    else:
        # Pad all rows to the max length
        max_len = max((len(row) for row in rows), default=0)
        rows = [row + [None] * (max_len - len(row)) for row in rows]
        return pd.DataFrame(rows)


def to_xlsx(
    df,
    path,
    sheet_name="Sheet1",
    header=True,
    index=False,
):
    """Write a pandas DataFrame to an XLSX file.

    Args:
        df: The pandas DataFrame to write.
        path: Output file path.
        sheet_name: Name of the worksheet (default "Sheet1").
        header: If True (default), write column names as the first row.
        index: If True, write the DataFrame index as the first column(s).
    """
    pd = _check_pandas()
    import numpy as np

    def _convert_value(val):
        """Convert a pandas/numpy value to a type XlsxWriter accepts."""
        if val is None or (isinstance(val, float) and np.isnan(val)):
            return None
        if pd.isna(val):
            return None
        if isinstance(val, (np.integer,)):
            return int(val)
        if isinstance(val, (np.floating,)):
            return float(val)
        if isinstance(val, (np.bool_,)):
            return bool(val)
        if isinstance(val, (pd.Timestamp, np.datetime64)):
            ts = pd.Timestamp(val)
            if ts.hour == 0 and ts.minute == 0 and ts.second == 0 and ts.microsecond == 0:
                return datetime.date(ts.year, ts.month, ts.day)
            return datetime.datetime(
                ts.year, ts.month, ts.day,
                ts.hour, ts.minute, ts.second, ts.microsecond,
            )
        if isinstance(val, (datetime.datetime, datetime.date)):
            return val
        if isinstance(val, (str, int, float, bool)):
            return val
        # Fallback: stringify
        return str(val)

    with XlsxWriter(path) as writer:
        writer.add_sheet(sheet_name)

        if header:
            col_names = list(df.columns)
            if index:
                index_names = list(df.index.names)
                index_names = [n if n is not None else "index" for n in index_names]
                col_names = index_names + col_names
            writer.write_row([_convert_value(c) for c in col_names])

        for row_idx in range(len(df)):
            row_values = []
            if index:
                idx_val = df.index[row_idx]
                if isinstance(idx_val, tuple):
                    row_values.extend([_convert_value(v) for v in idx_val])
                else:
                    row_values.append(_convert_value(idx_val))
            row_values.extend([_convert_value(v) for v in df.iloc[row_idx]])
            writer.write_row(row_values)
