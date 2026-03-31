"""Tests for security hardening — issue #20."""

import opensheet_core


def test_read_normal_file():
    """Normal files still work with security checks enabled."""
    import os
    import tempfile

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.add_sheet("Sheet1")
            for i in range(1000):
                w.write_row([f"row_{i}", i, i * 1.5])

        sheets = opensheet_core.read_xlsx(path)
        assert len(sheets[0]["rows"]) == 1000
    finally:
        os.unlink(path)


def test_document_properties_available():
    """document_properties function is importable and callable."""
    assert callable(opensheet_core.document_properties)


def test_data_validations_in_read_xlsx():
    """read_xlsx output includes data_validations key."""
    import os
    import tempfile

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.add_sheet("Sheet1")
            w.write_row(["data"])

        sheets = opensheet_core.read_xlsx(path)
        assert "data_validations" in sheets[0]
        assert sheets[0]["data_validations"] == []
    finally:
        os.unlink(path)
