"""Tests for sheet protection (read + write) -- issue #11."""

import os
import tempfile

import opensheet_core


def test_write_and_read_basic_protection():
    """Write basic sheet protection and read it back."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.add_sheet("Sheet1")
            w.protect_sheet()
            w.write_row(["Protected"])

        sheets = opensheet_core.read_xlsx(path)
        prot = sheets[0]["protection"]
        assert prot is not None
        assert prot["sheet"] is True
        assert prot["objects"] is True
        assert prot["scenarios"] is True
        assert prot["password_hash"] is None
    finally:
        os.unlink(path)


def test_write_and_read_protection_with_password():
    """Write sheet protection with password and read it back."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.add_sheet("Sheet1")
            w.protect_sheet(password="secret")
            w.write_row(["Protected"])

        sheets = opensheet_core.read_xlsx(path)
        prot = sheets[0]["protection"]
        assert prot is not None
        assert prot["sheet"] is True
        assert prot["password_hash"] is not None
        assert len(prot["password_hash"]) == 4  # 4-char hex
    finally:
        os.unlink(path)


def test_write_and_read_protection_with_options():
    """Write sheet protection with custom options."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.add_sheet("Sheet1")
            w.protect_sheet(
                sheet=True,
                objects=False,
                scenarios=False,
                format_cells=True,
                insert_rows=True,
                delete_rows=True,
                sort=True,
            )
            w.write_row(["Protected"])

        sheets = opensheet_core.read_xlsx(path)
        prot = sheets[0]["protection"]
        assert prot is not None
        assert prot["sheet"] is True
        assert prot["objects"] is False
        assert prot["scenarios"] is False
        assert prot["format_cells"] is True
        assert prot["insert_rows"] is True
        assert prot["delete_rows"] is True
        assert prot["sort"] is True
        assert prot["format_columns"] is False
    finally:
        os.unlink(path)


def test_no_protection():
    """Sheet without protection returns None."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.add_sheet("Sheet1")
            w.write_row(["No protection"])

        sheets = opensheet_core.read_xlsx(path)
        assert sheets[0]["protection"] is None
    finally:
        os.unlink(path)


def test_protection_requires_open_sheet():
    """Protecting without an open sheet raises an error."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        w = opensheet_core.XlsxWriter(path)
        try:
            w.protect_sheet()
            assert False, "Should have raised"
        except Exception as e:
            assert "No sheet is open" in str(e)
        w.close()
    finally:
        os.unlink(path)


def test_password_hash_deterministic():
    """Same password produces the same hash each time."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path1 = f.name
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path2 = f.name
    try:
        for path in [path1, path2]:
            with opensheet_core.XlsxWriter(path) as w:
                w.add_sheet("Sheet1")
                w.protect_sheet(password="test123")
                w.write_row(["data"])

        sheets1 = opensheet_core.read_xlsx(path1)
        sheets2 = opensheet_core.read_xlsx(path2)
        assert sheets1[0]["protection"]["password_hash"] == sheets2[0]["protection"]["password_hash"]
    finally:
        os.unlink(path1)
        os.unlink(path2)
