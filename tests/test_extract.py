"""Tests for AI/RAG extraction functions (xlsx_to_markdown, xlsx_to_text, xlsx_to_chunks)."""

import datetime
import pytest
import opensheet_core
from opensheet_core import (
    xlsx_to_markdown,
    xlsx_to_text,
    xlsx_to_chunks,
    XlsxWriter,
    Formula,
    FormattedCell,
    CellStyle,
    StyledCell,
)


@pytest.fixture
def tmp_xlsx(tmp_path):
    return str(tmp_path / "output.xlsx")


def _write_basic(path):
    """Write a simple 2-column sheet for testing."""
    with XlsxWriter(path) as w:
        w.add_sheet("Data")
        w.write_row(["Name", "Age"])
        w.write_row(["Alice", 30])
        w.write_row(["Bob", 25])


def _write_multi_sheet(path):
    """Write a workbook with two sheets."""
    with XlsxWriter(path) as w:
        w.add_sheet("Users")
        w.write_row(["Name", "Age"])
        w.write_row(["Alice", 30])
        w.add_sheet("Items")
        w.write_row(["Item", "Price"])
        w.write_row(["Widget", 9.99])


# ── xlsx_to_markdown ─────────────────────────────────────────────


class TestXlsxToMarkdown:
    def test_basic_table(self, tmp_xlsx):
        _write_basic(tmp_xlsx)
        md = xlsx_to_markdown(tmp_xlsx)
        lines = md.strip().split("\n")
        assert len(lines) == 4  # header + separator + 2 data rows
        assert "Name" in lines[0]
        assert "Age" in lines[0]
        assert "---" in lines[1]
        assert "Alice" in lines[2] and "30" in lines[2]
        assert "Bob" in lines[3] and "25" in lines[3]

    def test_pipe_structure(self, tmp_xlsx):
        _write_basic(tmp_xlsx)
        md = xlsx_to_markdown(tmp_xlsx)
        for line in md.strip().split("\n"):
            assert line.startswith("|")
            assert line.endswith("|")

    def test_single_sheet_no_heading(self, tmp_xlsx):
        _write_basic(tmp_xlsx)
        md = xlsx_to_markdown(tmp_xlsx)
        assert "##" not in md

    def test_multi_sheet_headings(self, tmp_xlsx):
        _write_multi_sheet(tmp_xlsx)
        md = xlsx_to_markdown(tmp_xlsx)
        assert "## Users" in md
        assert "## Items" in md

    def test_sheet_by_name(self, tmp_xlsx):
        _write_multi_sheet(tmp_xlsx)
        md = xlsx_to_markdown(tmp_xlsx, sheet_name="Items")
        assert "Widget" in md
        assert "Alice" not in md
        assert "##" not in md

    def test_sheet_by_index(self, tmp_xlsx):
        _write_multi_sheet(tmp_xlsx)
        md = xlsx_to_markdown(tmp_xlsx, sheet_index=1)
        assert "Widget" in md
        assert "Alice" not in md

    def test_no_header(self, tmp_xlsx):
        _write_basic(tmp_xlsx)
        md = xlsx_to_markdown(tmp_xlsx, header=False)
        # All rows should be data rows (with auto-generated column header)
        assert "Col 0" in md
        assert "Name" in md  # first row is now data

    def test_empty_sheet(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Empty")
            w.write_row(["Header"])
        md = xlsx_to_markdown(tmp_xlsx)
        assert "Header" in md

    def test_formulas_unwrapped(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Budget")
            w.write_row(["Item", "Cost"])
            w.write_row(["Rent", 1200])
            w.write_row(["Total", Formula("SUM(B2:B2)", cached_value=1200)])
        md = xlsx_to_markdown(tmp_xlsx)
        assert "1200" in md
        assert "SUM" not in md

    def test_formatted_cells_unwrapped(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Finance")
            w.write_row(["Item", "Price"])
            w.write_row(["Widget", FormattedCell(19.99, "$#,##0.00")])
        md = xlsx_to_markdown(tmp_xlsx)
        assert "19.99" in md

    def test_styled_cells_unwrapped(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Report")
            w.write_row([
                StyledCell("Name", CellStyle(bold=True)),
                StyledCell("Score", CellStyle(bold=True)),
            ])
            w.write_row(["Alice", 95])
        md = xlsx_to_markdown(tmp_xlsx)
        assert "Name" in md
        assert "Score" in md
        assert "Alice" in md

    def test_dates(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Dates")
            w.write_row(["Event", "Date"])
            w.write_row(["Launch", datetime.date(2025, 3, 15)])
        md = xlsx_to_markdown(tmp_xlsx)
        assert "2025-03-15" in md

    def test_datetimes(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Times")
            w.write_row(["Event", "Timestamp"])
            w.write_row(["Deploy", datetime.datetime(2025, 3, 15, 14, 30)])
        md = xlsx_to_markdown(tmp_xlsx)
        assert "2025-03-15T14:30:00" in md

    def test_none_cells(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Sparse")
            w.write_row(["A", "B", "C"])
            w.write_row(["x", None, "z"])
        md = xlsx_to_markdown(tmp_xlsx)
        lines = md.strip().split("\n")
        # The data row should have an empty cell in the middle
        assert lines[2].count("|") == 4  # 3 cells = 4 pipes

    def test_booleans(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Bools")
            w.write_row(["Flag", "Value"])
            w.write_row(["active", True])
            w.write_row(["deleted", False])
        md = xlsx_to_markdown(tmp_xlsx)
        assert "True" in md
        assert "False" in md

    def test_mixed_types(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Mix")
            w.write_row(["str", "int", "float", "bool", "none"])
            w.write_row(["hello", 42, 3.14, True, None])
        md = xlsx_to_markdown(tmp_xlsx)
        assert "hello" in md
        assert "42" in md
        assert "3.14" in md
        assert "True" in md


# ── xlsx_to_text ─────────────────────────────────────────────────


class TestXlsxToText:
    def test_basic(self, tmp_xlsx):
        _write_basic(tmp_xlsx)
        text = xlsx_to_text(tmp_xlsx)
        lines = text.strip().split("\n")
        assert lines[0] == "Name\tAge"
        assert lines[1] == "Alice\t30"
        assert lines[2] == "Bob\t25"

    def test_custom_delimiter(self, tmp_xlsx):
        _write_basic(tmp_xlsx)
        text = xlsx_to_text(tmp_xlsx, delimiter=",")
        assert "Name,Age" in text
        assert "Alice,30" in text

    def test_single_sheet_no_separator(self, tmp_xlsx):
        _write_basic(tmp_xlsx)
        text = xlsx_to_text(tmp_xlsx)
        assert "---" not in text

    def test_multi_sheet_separator(self, tmp_xlsx):
        _write_multi_sheet(tmp_xlsx)
        text = xlsx_to_text(tmp_xlsx)
        assert "--- Users ---" in text
        assert "--- Items ---" in text

    def test_sheet_by_name(self, tmp_xlsx):
        _write_multi_sheet(tmp_xlsx)
        text = xlsx_to_text(tmp_xlsx, sheet_name="Items")
        assert "Widget" in text
        assert "Alice" not in text
        assert "---" not in text

    def test_formulas_unwrapped(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Budget")
            w.write_row(["Total", Formula("SUM(A1:A2)", cached_value=100)])
        text = xlsx_to_text(tmp_xlsx)
        assert "100" in text
        assert "SUM" not in text

    def test_none_is_empty_string(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Sparse")
            w.write_row(["a", None, "c"])
        text = xlsx_to_text(tmp_xlsx)
        assert "a\t\tc" in text


# ── xlsx_to_chunks ───────────────────────────────────────────────


class TestXlsxToChunks:
    def test_single_chunk(self, tmp_xlsx):
        _write_basic(tmp_xlsx)
        chunks = xlsx_to_chunks(tmp_xlsx)
        assert len(chunks) == 1
        assert "Name" in chunks[0]
        assert "Alice" in chunks[0]

    def test_splits_by_max_rows(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Data")
            w.write_row(["ID", "Value"])
            for i in range(10):
                w.write_row([i, i * 10])
        chunks = xlsx_to_chunks(tmp_xlsx, max_rows=3)
        assert len(chunks) == 4  # ceil(10/3) = 4 chunks
        # Each chunk should have the header
        for chunk in chunks:
            assert "ID" in chunk
            assert "Value" in chunk
            assert "---" in chunk

    def test_header_repeated(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Data")
            w.write_row(["Name", "Score"])
            for i in range(6):
                w.write_row([f"student_{i}", i * 10])
        chunks = xlsx_to_chunks(tmp_xlsx, max_rows=2)
        assert len(chunks) == 3
        for chunk in chunks:
            lines = chunk.strip().split("\n")
            assert "Name" in lines[0]
            assert "---" in lines[1]

    def test_no_header_mode(self, tmp_xlsx):
        _write_basic(tmp_xlsx)
        chunks = xlsx_to_chunks(tmp_xlsx, header=False, max_rows=1)
        assert len(chunks) == 3  # All 3 rows are data
        # Auto-generated headers
        for chunk in chunks:
            assert "Col 0" in chunk

    def test_multi_sheet(self, tmp_xlsx):
        _write_multi_sheet(tmp_xlsx)
        chunks = xlsx_to_chunks(tmp_xlsx)
        assert len(chunks) == 2  # One chunk per sheet
        assert "## Users" in chunks[0]
        assert "## Items" in chunks[1]

    def test_multi_sheet_chunked(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("A")
            w.write_row(["X"])
            for i in range(4):
                w.write_row([i])
            w.add_sheet("B")
            w.write_row(["Y"])
            for i in range(4):
                w.write_row([i])
        chunks = xlsx_to_chunks(tmp_xlsx, max_rows=2)
        assert len(chunks) == 4  # 2 chunks per sheet
        # First sheet chunks
        assert "## A" in chunks[0]
        assert "## A" in chunks[1]
        # Second sheet chunks
        assert "## B" in chunks[2]
        assert "## B" in chunks[3]

    def test_single_sheet_by_name_no_label(self, tmp_xlsx):
        _write_multi_sheet(tmp_xlsx)
        chunks = xlsx_to_chunks(tmp_xlsx, sheet_name="Users")
        assert len(chunks) == 1
        assert "##" not in chunks[0]

    def test_max_rows_validation(self, tmp_xlsx):
        _write_basic(tmp_xlsx)
        with pytest.raises(ValueError, match="max_rows"):
            xlsx_to_chunks(tmp_xlsx, max_rows=0)

    def test_exact_boundary(self, tmp_xlsx):
        """When data rows == max_rows, should produce exactly 1 chunk."""
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Data")
            w.write_row(["H"])
            for i in range(5):
                w.write_row([i])
        chunks = xlsx_to_chunks(tmp_xlsx, max_rows=5)
        assert len(chunks) == 1


# ── Edge cases ───────────────────────────────────────────────────


class TestEdgeCases:
    def test_file_not_found(self):
        with pytest.raises(FileNotFoundError):
            xlsx_to_markdown("/nonexistent/file.xlsx")
        with pytest.raises(FileNotFoundError):
            xlsx_to_text("/nonexistent/file.xlsx")
        with pytest.raises(FileNotFoundError):
            xlsx_to_chunks("/nonexistent/file.xlsx")

    def test_whole_float_displayed_as_int(self, tmp_xlsx):
        """Floats like 30.0 should display as '30', not '30.0'."""
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Data")
            w.write_row(["Val"])
            w.write_row([30.0])
        md = xlsx_to_markdown(tmp_xlsx)
        assert "30 " in md or "30|" in md.replace(" ", "")
        assert "30.0" not in md

    def test_roundtrip_with_fixture(self):
        """Test with the existing fixture file."""
        import os
        fixture = os.path.join(os.path.dirname(__file__), "fixtures", "basic.xlsx")
        if os.path.exists(fixture):
            md = xlsx_to_markdown(fixture)
            assert len(md) > 0
            text = xlsx_to_text(fixture)
            assert len(text) > 0
            chunks = xlsx_to_chunks(fixture)
            assert len(chunks) > 0
