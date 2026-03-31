"""Tests for structured tables (read + write) -- issue #12."""

import os
import tempfile

import opensheet_core


def test_write_and_read_basic_table():
    """Write a structured table and read it back."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.add_sheet("Sheet1")
            w.write_row(["Name", "Age", "Score"])
            w.write_row(["Alice", 30, 95])
            w.write_row(["Bob", 25, 88])
            w.add_table("A1:C3", ["Name", "Age", "Score"], name="People")

        sheets = opensheet_core.read_xlsx(path)
        tables = sheets[0]["tables"]
        assert len(tables) == 1
        t = tables[0]
        assert t["name"] == "People"
        assert t["display_name"] == "People"
        assert t["ref"] == "A1:C3"
        assert len(t["columns"]) == 3
        assert t["columns"][0]["name"] == "Name"
        assert t["columns"][1]["name"] == "Age"
        assert t["columns"][2]["name"] == "Score"
        assert t["has_auto_filter"] is True
    finally:
        os.unlink(path)


def test_table_with_style():
    """Write a table with a style name."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.add_sheet("Sheet1")
            w.write_row(["A", "B"])
            w.write_row([1, 2])
            w.add_table(
                "A1:B2",
                ["A", "B"],
                name="MyTable",
                style="TableStyleMedium2",
            )

        sheets = opensheet_core.read_xlsx(path)
        tables = sheets[0]["tables"]
        assert len(tables) == 1
        assert tables[0]["style"] == "TableStyleMedium2"
    finally:
        os.unlink(path)


def test_table_default_name():
    """Table without explicit name gets auto-generated name."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.add_sheet("Sheet1")
            w.write_row(["X", "Y"])
            w.write_row([1, 2])
            w.add_table("A1:B2", ["X", "Y"])

        sheets = opensheet_core.read_xlsx(path)
        tables = sheets[0]["tables"]
        assert len(tables) == 1
        assert tables[0]["name"] == "Table1"
    finally:
        os.unlink(path)


def test_no_tables():
    """Sheet with no tables returns empty list."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.add_sheet("Sheet1")
            w.write_row(["data"])

        sheets = opensheet_core.read_xlsx(path)
        assert sheets[0]["tables"] == []
    finally:
        os.unlink(path)


def test_table_requires_open_sheet():
    """Adding a table without an open sheet raises an error."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        w = opensheet_core.XlsxWriter(path)
        try:
            w.add_table("A1:B2", ["A", "B"])
            assert False, "Should have raised"
        except Exception as e:
            assert "No sheet is open" in str(e)
        w.close()
    finally:
        os.unlink(path)


def test_multiple_tables():
    """Multiple tables on the same sheet."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.add_sheet("Sheet1")
            w.write_row(["A", "B", "", "C", "D"])
            w.write_row([1, 2, None, 3, 4])
            w.add_table("A1:B2", ["A", "B"], name="Table1")
            w.add_table("D1:E2", ["C", "D"], name="Table2")

        sheets = opensheet_core.read_xlsx(path)
        tables = sheets[0]["tables"]
        assert len(tables) == 2
        assert tables[0]["name"] == "Table1"
        assert tables[1]["name"] == "Table2"
    finally:
        os.unlink(path)
