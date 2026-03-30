# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.2.0] - 2026-03-30

### Added
- Cell styling support: `CellStyle` and `StyledCell` types for fonts (bold, italic, underline, name, size, color), fills (solid color), borders (thin, medium, thick, dashed, dotted, double with per-side control), alignment (horizontal, vertical, wrap text, rotation), and number formats on styled cells
- Reader returns `StyledCell` for cells with visual styling; plain cells remain unchanged (backward compatible)
- Style deduplication in writer (identical styles share a single XF entry)
- `border` shorthand parameter on `CellStyle` sets all four sides at once, with per-side overrides
- Column width support: `set_column_width()` on writer, `"column_widths"` in reader output
- Row height support: `set_row_height()` on writer, `"row_heights"` in reader output
- Freeze panes support: `freeze_panes()` on writer, `"freeze_pane"` in reader output
- Auto-filter support: `auto_filter()` on writer, `"auto_filter"` in reader output
- Number format support: `FormattedCell(value, format_code)` for writing cells with custom number formats (currency, percentage, custom format strings); reader returns `FormattedCell` for formatted number cells
- Pandas integration: `read_xlsx_df()` reads XLSX sheets into DataFrames, `to_xlsx()` writes DataFrames to XLSX; install with `pip install opensheet-core[pandas]`

## [0.1.1] - 2026-03-30

### Changed
- Limit Dependabot Cargo updates to patch-only, monthly
- Add Codecov test results reporting to CI
- Add project infrastructure (CONTRIBUTING.md, CHANGELOG.md, issue templates)

## [0.1.0] - 2026-03-29

### Added
- Streaming XLSX reader with typed cell extraction (strings, numbers, booleans, dates, datetimes, formulas, empty cells)
- Streaming XLSX writer with constant memory usage
- Formula read/write support with optional cached values
- Date and datetime cell support with automatic Excel serial number conversion
- Merged cell metadata (read and write)
- `read_xlsx()` to read all sheets from a workbook
- `read_sheet()` to read a single sheet by name or index
- `sheet_names()` to list sheet names in a workbook
- `XlsxWriter` with context manager support
- `Formula` type for formula cells
- Python type stubs (`.pyi`) and `py.typed` marker for IDE autocomplete
- Prebuilt wheels for Linux (x86_64, aarch64), macOS (x86_64, aarch64), and Windows (x64)
- CI across Linux, macOS, and Windows for Python 3.9-3.13
- Benchmarks vs openpyxl with runnable benchmark script
- Zero Python dependencies

[Unreleased]: https://github.com/0xNadr/opensheet-core/compare/v0.2.0...HEAD
[0.2.0]: https://github.com/0xNadr/opensheet-core/compare/v0.1.1...v0.2.0
[0.1.1]: https://github.com/0xNadr/opensheet-core/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/0xNadr/opensheet-core/releases/tag/v0.1.0
