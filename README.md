<p align="center">
  <img src="https://raw.githubusercontent.com/0xNadr/opensheet-core/main/assets/banner.svg" alt="OpenSheet Core — Fast, memory-efficient spreadsheet I/O for Python, powered by Rust" width="100%">
</p>

<p align="center">
  <a href="https://github.com/0xNadr/opensheet-core/actions/workflows/ci.yml"><img src="https://github.com/0xNadr/opensheet-core/actions/workflows/ci.yml/badge.svg" alt="CI"></a>
  <a href="https://pypi.org/project/opensheet-core/"><img src="https://img.shields.io/pypi/v/opensheet-core.svg" alt="PyPI"></a>
  <a href="https://github.com/0xNadr/opensheet-core/blob/main/LICENSE"><img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="License: MIT"></a>
  <a href="https://www.python.org/downloads/"><img src="https://img.shields.io/badge/python-3.9%E2%80%933.13-blue.svg" alt="Python 3.9–3.13"></a>
  <a href="https://codecov.io/gh/0xNadr/opensheet-core"><img src="https://codecov.io/gh/0xNadr/opensheet-core/branch/main/graph/badge.svg" alt="Coverage"></a>
</p>

<p align="center">
  <a href="#features">Features</a> &nbsp;&bull;&nbsp;
  <a href="#benchmarks">Benchmarks</a> &nbsp;&bull;&nbsp;
  <a href="#installation">Installation</a> &nbsp;&bull;&nbsp;
  <a href="#quick-start">Quick Start</a> &nbsp;&bull;&nbsp;
  <a href="#api-reference">API</a> &nbsp;&bull;&nbsp;
  <a href="#roadmap">Roadmap</a> &nbsp;&bull;&nbsp;
  <a href="#contributing">Contributing</a>
</p>

---

## Why OpenSheet Core?

Existing Python spreadsheet libraries force you to choose between performance, memory efficiency, broad format support, and easy installation. OpenSheet Core eliminates that tradeoff with a native Rust core exposed through a clean Python API — installable with a single `pip install`.

## Features

- **Streaming XLSX reader** — row-by-row iteration without loading the entire file into memory
- **Streaming XLSX writer** — write millions of rows with constant memory usage
- **Formula support** — read and write formulas with optional cached values
- **Date/time support** — read and write `datetime.date` and `datetime.datetime` cells with automatic Excel serial number conversion
- **Merged cells** — read and write merged cell ranges
- **Column widths & row heights** — set and read custom column widths and row heights
- **Freeze panes** — freeze rows and/or columns so they stay visible when scrolling
- **Auto-filter** — add drop-down filter controls to column headers
- **Number formats** — write and read cells with custom number formats (currency, percentage, custom format strings)
- **Cell styling** — fonts (bold, italic, underline, name, size, color), fills, borders (thin, medium, thick, dashed, dotted, double), alignment (horizontal, vertical, wrap text, rotation), and number formats on styled cells
- **Typed cell extraction** — strings, numbers, booleans, dates, datetimes, formulas, and empty cells are returned as native Python types
- **Context manager support** — Pythonic `with` statement for safe resource management
- **AI/RAG-ready** — convert spreadsheets to markdown, chunked text, or plain text for LLM pipelines (planned)
- **Cross-platform** — tested on Linux, macOS, and Windows across Python 3.9–3.13
- **Pandas integration** — read XLSX files into DataFrames and write DataFrames to XLSX (`pip install opensheet-core[pandas]`)
- **Zero Python dependencies** — single native extension, no dependency tree to manage

## Benchmarks

Benchmarked against [openpyxl](https://openpyxl.readthedocs.io/) 3.1.5 on a 100,000-row x 10-column dataset (1M cells), 5 interleaved runs, current RSS measurement (not high-water mark):

| Operation | OpenSheet Core | openpyxl | Speedup | Memory (RSS delta) |
|-----------|---------------|----------|---------|---------------------|
| **Write** | 2.3s | 3.7s | **1.6x faster** | **1.7x less** (1.2 MB vs 2.1 MB) |
| **Read** | 253ms | 3.5s | **13.8x faster** | **2.5x less** (13.5 MB vs 33.3 MB) |

OpenSheet Core is faster and uses less memory for both reads and writes. The speed advantage comes from a Rust streaming parser with deferred shared-string resolution — strings are stored as indices during parsing and only converted to Python objects at the boundary. Write memory is low because the Rust writer streams data directly to disk.

> Run it yourself: `python benchmarks/benchmark.py`
>
> See the [Benchmarking Methodology](docs/benchmarking.md) doc for details on how we measure and avoid common benchmarking pitfalls.

## Installation

```bash
pip install opensheet-core

# With pandas support
pip install opensheet-core[pandas]
```

### From source (requires Rust toolchain)

```bash
pip install maturin
git clone https://github.com/0xNadr/opensheet-core
cd opensheet-core
maturin develop --release
```

## Quick Start

### Reading an XLSX file

```python
from opensheet_core import read_xlsx, read_sheet

# Read all sheets
sheets = read_xlsx("report.xlsx")
for sheet in sheets:
    print(f"Sheet: {sheet['name']}")
    for row in sheet["rows"]:
        print(row)  # List of typed Python values

# Read a specific sheet
rows = read_sheet("report.xlsx", sheet_name="Data")
```

### Writing an XLSX file

```python
from opensheet_core import XlsxWriter

with XlsxWriter("output.xlsx") as writer:
    writer.add_sheet("Data")
    writer.write_row(["Name", "Age", "Active"])
    writer.write_row(["Alice", 30, True])
    writer.write_row(["Bob", 25, False])
```

### Writing dates

```python
import datetime
from opensheet_core import XlsxWriter

with XlsxWriter("output.xlsx") as writer:
    writer.add_sheet("Events")
    writer.write_row(["Event", "Date", "Timestamp"])
    writer.write_row(["Launch", datetime.date(2025, 3, 15), datetime.datetime(2025, 3, 15, 14, 30)])
```

### Merging cells

```python
from opensheet_core import XlsxWriter

with XlsxWriter("output.xlsx") as writer:
    writer.add_sheet("Report")
    writer.write_row(["Title spanning three columns", "", ""])
    writer.write_row(["A", "B", "C"])
    writer.merge_cells("A1:C1")
```

### Column widths and row heights

```python
from opensheet_core import XlsxWriter

with XlsxWriter("output.xlsx") as writer:
    writer.add_sheet("Data")
    writer.set_column_width("A", 25.0)   # By letter
    writer.set_column_width(1, 15.0)     # By 0-based index
    writer.set_row_height(1, 30.0)       # Row 1 (1-based)
    writer.write_row(["Name", "Age"])
    writer.write_row(["Alice", 30])
```

### Freeze panes

```python
from opensheet_core import XlsxWriter

with XlsxWriter("output.xlsx") as writer:
    writer.add_sheet("Data")
    writer.freeze_panes(row=1, col=0)    # Freeze top row
    writer.write_row(["Header1", "Header2", "Header3"])
    writer.write_row(["data", "data", "data"])
```

### Auto-filter

```python
from opensheet_core import XlsxWriter

with XlsxWriter("output.xlsx") as writer:
    writer.add_sheet("Data")
    writer.write_row(["Name", "Age", "City"])
    writer.write_row(["Alice", 30, "NYC"])
    writer.write_row(["Bob", 25, "LA"])
    writer.auto_filter("A1:C1")
```

### Number formats

```python
from opensheet_core import XlsxWriter, FormattedCell

with XlsxWriter("output.xlsx") as writer:
    writer.add_sheet("Finance")
    writer.write_row(["Item", "Price", "Tax Rate"])
    writer.write_row([
        "Widget",
        FormattedCell(19.99, "$#,##0.00"),   # Currency
        FormattedCell(0.08, "0.00%"),         # Percentage
    ])
```

### Cell styling

```python
from opensheet_core import XlsxWriter, CellStyle, StyledCell

with XlsxWriter("output.xlsx") as writer:
    writer.add_sheet("Report")
    # Bold header with fill color
    writer.write_row([
        StyledCell("Name", CellStyle(bold=True, fill_color="4472C4", font_color="FFFFFF")),
        StyledCell("Score", CellStyle(bold=True, fill_color="4472C4", font_color="FFFFFF")),
    ])
    # Data with borders and alignment
    writer.write_row([
        StyledCell("Alice", CellStyle(border="thin", horizontal_alignment="left")),
        StyledCell(95, CellStyle(border="thin", number_format="0.0")),
    ])
```

### Writing formulas

```python
from opensheet_core import XlsxWriter, Formula

with XlsxWriter("output.xlsx") as writer:
    writer.add_sheet("Budget")
    writer.write_row(["Item", "Cost"])
    writer.write_row(["Rent", 1200])
    writer.write_row(["Food", 400])
    writer.write_row(["Total", Formula("SUM(B2:B3)", cached_value=1600)])
```

### Pandas integration

```python
import pandas as pd
from opensheet_core import read_xlsx_df, to_xlsx

# Read XLSX into a DataFrame
df = read_xlsx_df("data.xlsx", sheet_name="Sheet1")

# Write a DataFrame to XLSX
df = pd.DataFrame({"Name": ["Alice", "Bob"], "Age": [30, 25]})
to_xlsx(df, "output.xlsx", sheet_name="Results")
```

## API Reference

### `read_xlsx(path: str) -> list[dict]`

Reads an XLSX file and returns a list of dicts with `"name"` (str), `"rows"` (list of lists), `"merges"` (list of range strings like `"A1:C1"`), `"column_widths"` (dict of 0-based col index to width), `"row_heights"` (dict of 0-based row index to height), `"freeze_pane"` (tuple of `(rows_frozen, cols_frozen)` or `None`), and `"auto_filter"` (range string like `"A1:C1"` or `None`). Each cell is a typed Python value (`str`, `int`, `float`, `bool`, `datetime.date`, `datetime.datetime`, `Formula`, `FormattedCell`, or `None`).

### `read_sheet(path, sheet_name=None, sheet_index=None) -> list[list]`

Reads a single sheet by name or index. Returns the first sheet by default.

### `sheet_names(path: str) -> list[str]`

Returns the list of sheet names in a workbook.

### `XlsxWriter(path: str)`

Streaming XLSX writer. Use as a context manager.

| Method | Description |
|--------|-------------|
| `add_sheet(name: str)` | Create a new worksheet |
| `write_row(values: list)` | Write a row of values to the current sheet |
| `merge_cells(range: str)` | Merge a range of cells (e.g. `"A1:C1"`) |
| `set_column_width(column, width)` | Set column width (`column` is a letter or 0-based int) |
| `set_row_height(row, height)` | Set row height in points (`row` is 1-based) |
| `freeze_panes(row=0, col=0)` | Freeze top `row` rows and left `col` columns |
| `auto_filter(range)` | Set auto-filter on a range (e.g. `"A1:C1"`) |
| `close()` | Finalize and close the file |

### `read_xlsx_df(path, sheet_name=None, sheet_index=None, header=True)`

Reads a single XLSX sheet into a pandas DataFrame. Requires `pip install opensheet-core[pandas]`. When `header=True` (default), the first row is used as column names. Formulas are unwrapped to cached values, `FormattedCell` values are unwrapped to plain numbers.

### `to_xlsx(df, path, sheet_name="Sheet1", header=True, index=False)`

Writes a pandas DataFrame to an XLSX file. Handles numpy int/float/bool types, `NaN`/`NaT` (written as empty cells), and `datetime64`/`Timestamp` columns. Set `index=True` to include the DataFrame index as column(s).

### `CellStyle(**kwargs)`

Style properties for a cell. All parameters are keyword-only. Properties: `bold` (bool), `italic` (bool), `underline` (bool), `font_name` (str), `font_size` (float), `font_color` (str, hex RGB), `fill_color` (str, hex RGB), `border` (str, shorthand for all 4 sides), `border_left`/`border_right`/`border_top`/`border_bottom` (str: `"thin"`, `"medium"`, `"thick"`, `"dashed"`, `"dotted"`, `"double"`), `border_color` (str, hex RGB), `horizontal_alignment` (str: `"left"`, `"center"`, `"right"`), `vertical_alignment` (str: `"top"`, `"center"`, `"bottom"`), `wrap_text` (bool), `text_rotation` (int, 0-180), `number_format` (str, Excel format code).

### `StyledCell(value, style: CellStyle)`

A cell value with styling. Pass as a cell value in `write_row()`. Returned by `read_xlsx()` and `read_sheet()` for cells that have visual styling. The inner `value` can be a string, number, bool, date, datetime, or formula.

### `FormattedCell(value, number_format: str)`

A numeric value with a custom Excel number format code. Pass as a cell value in `write_row()`. Returned by `read_xlsx()` for cells with non-default number formats. Common format codes: `"$#,##0.00"` (currency), `"0.00%"` (percentage), `"#,##0"` (thousands separator).

### `Formula(formula: str, cached_value=None)`

Represents a spreadsheet formula. Pass as a cell value when writing, and received when reading cells that contain formulas.

## Architecture

```
┌──────────────────────────┐
│      Python API          │  ← opensheet_core (PyO3 bindings)
├──────────────────────────┤
│      Rust Core           │  ← Streaming parser & writer
│  ┌────────┐ ┌──────────┐ │
│  │ Reader │ │  Writer  │ │
│  │ (SAX)  │ │ (Stream) │ │
│  └────────┘ └──────────┘ │
├──────────────────────────┤
│  quick-xml  │    zip     │  ← Dependencies
└──────────────────────────┘
```

## Feature Comparison vs openpyxl

OpenSheet Core is designed to be a faster, memory-efficient alternative to openpyxl for the most common spreadsheet workflows. Here's where we stand:

### What we already do better

| | OpenSheet Core | openpyxl |
|---|---|---|
| **Write 1M cells** | ~2.3s | ~3.7s |
| **Read 1M cells** | ~0.25s | ~3.5s |
| **Write memory** | 1.2 MB RSS delta | 2.1 MB RSS delta |
| **Read memory** | 13.5 MB RSS delta | 33.3 MB RSS delta |
| **Python dependencies** | Zero | Several |
| **Architecture** | Rust streaming core | Pure Python DOM |

> Memory optimization: shared strings are stored as indices during parsing and resolved to Python objects at the boundary via pre-interned lookup, avoiding duplicate string allocations. A future streaming iterator API will bring constant-memory reads.

### Feature coverage

| Category | Feature | openpyxl | OpenSheet Core |
|----------|---------|:--------:|:--------------:|
| **Formats** | .xlsx read/write | Yes | Yes |
| | .xlsm (macro-enabled) | Yes | Planned |
| | .xltx/.xltm (templates) | Yes | — |
| **Cell Types** | Strings, numbers, booleans | Yes | Yes |
| | Dates and datetimes | Yes | Yes |
| | Formulas with cached values | Yes | Yes |
| | Rich text | Yes | Planned |
| | Error values | Yes | Planned |
| **Styling** | Fonts (name, size, bold, italic, color) | Yes | Yes |
| | Fill (solid, pattern, gradient) | Yes | Solid |
| | Borders (14 styles) | Yes | 6 styles |
| | Alignment (horizontal, vertical, wrap, rotation) | Yes | Yes |
| | Number formats (30+ builtins + custom) | Yes | Yes |
| | Named styles | Yes | Planned |
| | Conditional formatting (6 rule types) | Yes | Planned |
| **Worksheet** | Merged cells | Yes | Yes |
| | Freeze panes | Yes | Yes |
| | Auto-filter | Yes | Yes |
| | Column widths / row heights | Yes | Yes |
| | Data validation (7 types) | Yes | Planned |
| | Sheet protection | Yes | Planned |
| | Row/column insert/delete | Yes | — |
| | Print settings | Yes | Planned |
| | Row/column grouping | Yes | — |
| **Workbook** | Named ranges / defined names | Yes | Planned |
| | Document properties | Yes | Planned |
| | Workbook protection | Yes | — |
| | Multiple sheet states (hidden, veryHidden) | Yes | Planned |
| **Charts** | 12+ chart types (bar, line, pie, scatter, etc.) | Yes | Planned |
| | 3D variants and combined charts | Yes | — |
| **Images** | Embed PNG/JPEG | Yes | Planned |
| **Tables** | Structured tables with styles | Yes | Planned |
| **Pivot Tables** | Read/preserve existing | Yes | — |
| **VBA/Macros** | Preserve on load (.xlsm) | Yes | Planned |
| **Integration** | Pandas DataFrame I/O | Yes | Yes |
| | NumPy type support | Yes | Yes |
| **AI/RAG** | Markdown/text extraction for LLMs | — | Planned |
| | Embedding-sized chunking | — | Planned |
| | LangChain / LlamaIndex loaders | — | Planned |
| **Performance** | Streaming read (constant memory) | Yes (read_only mode) | Yes (default) |
| | Streaming write (constant memory) | Yes (write_only mode) | Yes (default) |

> **Legend:** Yes = implemented, Planned = on the roadmap, — = not planned for now

### Our approach

We are not trying to clone openpyxl. We are building a **fast, safe, memory-efficient core** for the most common Excel workflows. The goal is to cover the ~80% of features that people use day-to-day, while being up to 14x faster and using 2–3x less memory. Streaming is the default, not an opt-in mode.

## Roadmap

### Done

- [x] XLSX reading with typed cell extraction
- [x] Streaming XLSX writing with low memory usage
- [x] Formula read/write support with cached values
- [x] Date/time cell support with automatic serial number conversion
- [x] Merged cell metadata (read and write)
- [x] Python bindings via PyO3
- [x] Type stubs (`.pyi`) and `py.typed` marker for IDE autocomplete
- [x] CI across Linux, macOS, Windows (Python 3.9–3.13)
- [x] Prebuilt wheels on PyPI
- [x] Benchmarks vs openpyxl
- [x] Runnable benchmark script (`python benchmarks/benchmark.py`)
- [x] Zero Python dependencies
- [x] Column widths and row heights
- [x] Freeze panes
- [x] Auto-filter
- [x] Number formats (currency, percentage, custom format strings)
- [x] Pandas DataFrame integration (`read_xlsx_df` / `to_xlsx`)
- [x] Basic cell styling (fonts, fills, borders, alignment)

### Phase 1.5 — AI/RAG integration

- [ ] `xlsx_to_markdown()` — convert sheets to structured markdown tables for LLM consumption
- [ ] `xlsx_to_chunks()` — yield embedding-sized chunks with header attachment and merge-aware boundaries
- [ ] `xlsx_to_text()` — plain text extraction for simple pipelines and search indexes
- [ ] LangChain `OpenSheetLoader` document loader
- [ ] LlamaIndex `OpenSheetReader` data connector

### Phase 2 — Broader compatibility

- [ ] Named ranges / defined names
- [ ] Data validation
- [ ] Comments and hyperlinks
- [ ] .xlsm read support (preserve macros)
- [ ] Sheet protection
- [ ] Structured tables with styles
- [ ] Multiple sheet states (hidden, veryHidden)

### Phase 3 — Rich content and ecosystem

- [ ] Charts (bar, line, pie, scatter — most common types)
- [ ] Image embedding (PNG, JPEG)
- [ ] Conditional formatting
- [ ] Document and custom properties
- [ ] NumPy type support
- [ ] Broader test corpus and fuzzing
- [ ] Security hardening (XML attack prevention)

### Docs & community

- [ ] Migration guide: openpyxl → opensheet-core (side-by-side code comparisons)
- [ ] FastAPI/Flask streaming XLSX download examples
- [x] Benchmark methodology documentation
- [ ] Dedicated benchmark page with chart visualizations

## Project Status

**v0.2.0** — streaming reader and writer with formula, date/time, merged cell, column width/row height, freeze pane, auto-filter, number format, cell styling, and pandas DataFrame support. 135 passing tests and prebuilt wheels on PyPI. The API may change before 1.0.

## Contributing

Contributions are welcome! Here are some great ways to get involved:

- Report bugs or real-world spreadsheet edge cases
- Submit representative sample files for testing
- Suggest benchmark scenarios
- Improve documentation
- Open PRs for roadmap items

## License

[MIT](LICENSE)

---

<p align="center">
  <sub>Built with Rust and PyO3 &nbsp;|&nbsp; Open digital infrastructure for the Python ecosystem</sub>
</p>
