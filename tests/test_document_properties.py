"""Tests for document properties (core + custom) — issue #17."""

import os
import tempfile

import opensheet_core


def test_write_and_read_core_properties():
    """Write core document properties and read them back."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.set_document_property("title", "Test Spreadsheet")
            w.set_document_property("subject", "Unit Testing")
            w.set_document_property("creator", "OpenSheet Core Tests")
            w.set_document_property("keywords", "test, xlsx, python")
            w.set_document_property("description", "A test file")
            w.set_document_property("last_modified_by", "Test Runner")
            w.set_document_property("category", "Testing")
            w.add_sheet("Sheet1")
            w.write_row(["hello"])

        props = opensheet_core.document_properties(path)
        core = props["core"]
        assert core["title"] == "Test Spreadsheet"
        assert core["subject"] == "Unit Testing"
        assert core["creator"] == "OpenSheet Core Tests"
        assert core["keywords"] == "test, xlsx, python"
        assert core["description"] == "A test file"
        assert core["last_modified_by"] == "Test Runner"
        assert core["category"] == "Testing"
    finally:
        os.unlink(path)


def test_write_and_read_custom_properties():
    """Write custom properties and read them back."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.set_custom_property("Project", "OpenSheet")
            w.set_custom_property("Version", "1.0")
            w.set_custom_property("Reviewed", "true")
            w.add_sheet("Sheet1")
            w.write_row(["data"])

        props = opensheet_core.document_properties(path)
        custom = props["custom"]
        assert len(custom) == 3
        names = {p["name"]: p["value"] for p in custom}
        assert names["Project"] == "OpenSheet"
        assert names["Version"] == "1.0"
        assert names["Reviewed"] == "true"
    finally:
        os.unlink(path)


def test_no_properties():
    """File with no properties returns empty core dict and empty custom list."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.add_sheet("Sheet1")
            w.write_row(["no props"])

        props = opensheet_core.document_properties(path)
        assert props["core"] == {}
        assert props["custom"] == []
    finally:
        os.unlink(path)


def test_special_characters_in_properties():
    """Properties with XML special characters are handled correctly."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.set_document_property("title", 'Test & "Quotes" <Tags>')
            w.set_custom_property("HTML", "<b>bold</b>")
            w.add_sheet("Sheet1")
            w.write_row(["data"])

        props = opensheet_core.document_properties(path)
        assert props["core"]["title"] == 'Test & "Quotes" <Tags>'
        assert props["custom"][0]["value"] == "<b>bold</b>"
    finally:
        os.unlink(path)


def test_invalid_property_key():
    """Setting an invalid core property key raises an error."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            try:
                w.set_document_property("invalid_key", "value")
                assert False, "Should have raised"
            except Exception as e:
                assert "Unknown document property key" in str(e)
    finally:
        os.unlink(path)
