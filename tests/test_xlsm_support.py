"""Tests for .xlsm read support -- issue #10."""

import os
import tempfile
import zipfile

import opensheet_core


def _create_xlsm(path):
    """Create a minimal .xlsm file (same as .xlsx but with VBA content type)."""
    # First, create a normal xlsx via the writer
    xlsx_path = path + ".xlsx"
    with opensheet_core.XlsxWriter(xlsx_path) as w:
        w.add_sheet("Sheet1")
        w.write_row(["Hello", 42])
        w.write_row(["World", 99])

    # Re-pack as .xlsm: just copy the zip entries and update content type
    with zipfile.ZipFile(xlsx_path, "r") as zin, zipfile.ZipFile(path, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                # Replace .sheet.main with .sheet.macroEnabled.main
                data = data.replace(
                    b"spreadsheetml.sheet.main+xml",
                    b"spreadsheetml.sheet.macroEnabled.main+xml",
                )
            zout.writestr(item, data)
        # Add a dummy vbaProject.bin
        zout.writestr("xl/vbaProject.bin", b"dummy VBA project data")

    os.unlink(xlsx_path)


def test_read_xlsm_basic():
    """Reading .xlsm files works the same as .xlsx."""
    with tempfile.NamedTemporaryFile(suffix=".xlsm", delete=False) as f:
        path = f.name
    try:
        _create_xlsm(path)
        sheets = opensheet_core.read_xlsx(path)
        assert len(sheets) == 1
        assert sheets[0]["name"] == "Sheet1"
        assert len(sheets[0]["rows"]) == 2
        assert sheets[0]["rows"][0][0] == "Hello"
        assert sheets[0]["rows"][0][1] == 42
    finally:
        os.unlink(path)


def test_read_xlsm_sheet_names():
    """sheet_names() works on .xlsm files."""
    with tempfile.NamedTemporaryFile(suffix=".xlsm", delete=False) as f:
        path = f.name
    try:
        _create_xlsm(path)
        names = opensheet_core.sheet_names(path)
        assert names == ["Sheet1"]
    finally:
        os.unlink(path)


def test_read_xlsm_read_sheet():
    """read_sheet() works on .xlsm files."""
    with tempfile.NamedTemporaryFile(suffix=".xlsm", delete=False) as f:
        path = f.name
    try:
        _create_xlsm(path)
        rows = opensheet_core.read_sheet(path)
        assert len(rows) == 2
        assert rows[1][0] == "World"
    finally:
        os.unlink(path)
