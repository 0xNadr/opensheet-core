"""Test the reader against a diverse corpus of XLSX files.

Covers valid files (different features, data types, structural variations)
and malformed files (truncated, missing parts, invalid XML) to ensure
the reader either succeeds or returns a clean error — never panics.
"""

import os
import pytest
from pathlib import Path

import opensheet_core

FIXTURES = Path(__file__).parent / "fixtures"


# ---- Helpers ----

def read_or_error(path: str):
    """Try reading; return result or exception (never raises)."""
    try:
        return opensheet_core.read_xlsx(path)
    except Exception as e:
        return e


# ---- Valid corpus: must parse successfully ----

class TestValidCorpus:
    """Files that are valid XLSX and must parse without error."""

    def test_empty_workbook(self):
        sheets = opensheet_core.read_xlsx(str(FIXTURES / "empty_workbook.xlsx"))
        assert len(sheets) == 1
        assert sheets[0]["rows"] == []

    def test_single_cell(self):
        sheets = opensheet_core.read_xlsx(str(FIXTURES / "single_cell.xlsx"))
        rows = sheets[0]["rows"]
        assert len(rows) == 1
        assert rows[0][0] == 42.0

    def test_many_data_types(self):
        sheets = opensheet_core.read_xlsx(str(FIXTURES / "many_data_types.xlsx"))
        sheet = sheets[0]
        row1 = sheet["rows"][0]
        # Number
        assert row1[0] == pytest.approx(3.14)
        # Bool
        assert row1[1] is True
        # Shared strings
        assert row1[2] == "hello"
        assert row1[3] == "world"

    def test_multiple_sheets(self):
        sheets = opensheet_core.read_xlsx(str(FIXTURES / "multiple_sheets.xlsx"))
        assert len(sheets) == 5
        # Check sheet states
        states = [s["state"] for s in sheets]
        assert states[0] == "visible"
        assert states[2] == "hidden"
        assert states[4] == "veryHidden"

    def test_merged_cells(self):
        sheets = opensheet_core.read_xlsx(str(FIXTURES / "merged_cells.xlsx"))
        sheet = sheets[0]
        assert len(sheet["merges"]) == 2
        assert "A1:B2" in sheet["merges"]

    def test_sparse_rows(self):
        """Sparse rows should produce a grid with gaps filled by None."""
        sheets = opensheet_core.read_xlsx(str(FIXTURES / "sparse_rows.xlsx"))
        rows = sheets[0]["rows"]
        assert len(rows) >= 1000  # Should have rows up to 1000

    def test_unicode_strings(self):
        sheets = opensheet_core.read_xlsx(str(FIXTURES / "unicode_strings.xlsx"))
        row = sheets[0]["rows"][0]
        assert "🌍" in row[0]  # Emoji
        assert "日本語" in row[1]  # CJK

    def test_rich_text_strings(self):
        sheets = opensheet_core.read_xlsx(str(FIXTURES / "rich_text_strings.xlsx"))
        row = sheets[0]["rows"][0]
        assert "Bold" in row[0]
        assert "Normal" in row[0]

    def test_freeze_panes(self):
        sheets = opensheet_core.read_xlsx(str(FIXTURES / "freeze_panes.xlsx"))
        sheet = sheets[0]
        assert sheet["freeze_pane"] is not None

    def test_auto_filter(self):
        sheets = opensheet_core.read_xlsx(str(FIXTURES / "auto_filter.xlsx"))
        sheet = sheets[0]
        assert sheet["auto_filter"] == "A1:B2"

    def test_defined_names(self):
        names = opensheet_core.defined_names(str(FIXTURES / "defined_names.xlsx"))
        assert any(n["name"] == "MyRange" for n in names)

    def test_large_shared_strings(self):
        sheets = opensheet_core.read_xlsx(str(FIXTURES / "large_shared_strings.xlsx"))
        sheet = sheets[0]
        assert len(sheet["rows"]) > 0

    def test_date_cells(self):
        sheets = opensheet_core.read_xlsx(str(FIXTURES / "date_cells.xlsx"))
        rows = sheets[0]["rows"]
        assert len(rows) > 0

    def test_column_widths(self):
        sheets = opensheet_core.read_xlsx(str(FIXTURES / "column_widths.xlsx"))
        widths = sheets[0]["column_widths"]
        assert len(widths) > 0

    def test_inline_strings(self):
        sheets = opensheet_core.read_xlsx(str(FIXTURES / "inline_strings.xlsx"))
        row = sheets[0]["rows"][0]
        assert row[0] == "Inline A"
        assert row[1] == "Inline B"

    def test_document_properties(self):
        sheets = opensheet_core.read_xlsx(str(FIXTURES / "document_properties.xlsx"))
        assert len(sheets) > 0

    def test_sheet_names(self):
        """sheet_names() should work on all valid files."""
        for name in ["multiple_sheets.xlsx", "defined_names.xlsx", "empty_workbook.xlsx"]:
            names = opensheet_core.sheet_names(str(FIXTURES / name))
            assert isinstance(names, list)
            assert len(names) > 0

    def test_read_sheet_by_index(self):
        rows = opensheet_core.read_sheet(str(FIXTURES / "multiple_sheets.xlsx"), sheet_index=1)
        assert isinstance(rows, list)

    def test_read_sheet_by_name(self):
        rows = opensheet_core.read_sheet(str(FIXTURES / "multiple_sheets.xlsx"), sheet_name="Sheet3")
        assert isinstance(rows, list)


# ---- Malformed corpus: must not panic ----

class TestMalformedCorpus:
    """Files that are invalid or malformed. The reader must return
    a clean error (exception), never panic or hang."""

    def test_truncated_zip(self):
        r = read_or_error(str(FIXTURES / "truncated_zip.xlsx"))
        assert isinstance(r, Exception)

    def test_empty_zip(self):
        r = read_or_error(str(FIXTURES / "empty_zip.xlsx"))
        assert isinstance(r, Exception)

    def test_missing_workbook(self):
        r = read_or_error(str(FIXTURES / "missing_workbook.xlsx"))
        assert isinstance(r, Exception)

    def test_malformed_xml(self):
        """Malformed XML should produce an error or degrade gracefully."""
        r = read_or_error(str(FIXTURES / "malformed_xml.xlsx"))
        # May succeed partially or error — just must not panic
        assert r is not None

    def test_wrong_string_index(self):
        """Out-of-bounds shared string index should not panic."""
        r = read_or_error(str(FIXTURES / "wrong_string_index.xlsx"))
        assert r is not None

    def test_negative_row_number(self):
        """Row number 0 should not cause a panic."""
        r = read_or_error(str(FIXTURES / "negative_row_number.xlsx"))
        assert r is not None

    def test_huge_row_gap(self):
        """Rows at 1 and 1048576 should not OOM or panic."""
        r = read_or_error(str(FIXTURES / "huge_row_gap.xlsx"))
        assert r is not None

    def test_random_bytes(self):
        """Completely random data should produce a clean error."""
        import tempfile
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            f.write(os.urandom(1024))
            f.flush()
            r = read_or_error(f.name)
            assert isinstance(r, Exception)
        os.unlink(f.name)

    def test_nonexistent_file(self):
        r = read_or_error("/nonexistent/path/file.xlsx")
        assert isinstance(r, Exception)


# ---- Roundtrip: write then read ----

class TestRoundtrip:
    """Write diverse data and verify we can read it back."""

    def test_roundtrip_all_types(self):
        import tempfile
        path = tempfile.mktemp(suffix=".xlsx")
        try:
            w = opensheet_core.XlsxWriter(path)
            w.add_sheet("Types")
            w.write_row([1, 2.5, True, False, "text", None, ""])
            w.write_row([0, -1, 1e10, 1e-10, "🎉", " ", "\t"])
            w.close()

            sheets = opensheet_core.read_xlsx(path)
            rows = sheets[0]["rows"]
            assert rows[0][0] == 1.0
            assert rows[0][2] is True
            assert rows[0][4] == "text"
            assert rows[1][4] == "🎉"
        finally:
            os.unlink(path)

    def test_roundtrip_many_sheets(self):
        import tempfile
        path = tempfile.mktemp(suffix=".xlsx")
        try:
            w = opensheet_core.XlsxWriter(path)
            for i in range(10):
                w.add_sheet(f"Sheet{i}")
                w.write_row([i, f"data_{i}"])
            w.close()

            names = opensheet_core.sheet_names(path)
            assert len(names) == 10
            assert names[0] == "Sheet0"
            assert names[9] == "Sheet9"
        finally:
            os.unlink(path)
