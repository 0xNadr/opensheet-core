"""Tests for comments and hyperlinks (read + write) -- issue #9."""

import os
import tempfile

import opensheet_core


def test_write_and_read_comment():
    """Write a comment and read it back."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.add_sheet("Sheet1")
            w.write_row(["Data"])
            w.add_comment("A1", "Author1", "This is a comment")

        sheets = opensheet_core.read_xlsx(path)
        comments = sheets[0]["comments"]
        assert len(comments) == 1
        assert comments[0]["cell"] == "A1"
        assert comments[0]["author"] == "Author1"
        assert comments[0]["text"] == "This is a comment"
    finally:
        os.unlink(path)


def test_multiple_comments():
    """Multiple comments on different cells."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.add_sheet("Sheet1")
            w.write_row(["A", "B", "C"])
            w.add_comment("A1", "Alice", "First comment")
            w.add_comment("B1", "Bob", "Second comment")
            w.add_comment("C1", "Alice", "Third comment")

        sheets = opensheet_core.read_xlsx(path)
        comments = sheets[0]["comments"]
        assert len(comments) == 3
        assert comments[0]["author"] == "Alice"
        assert comments[1]["author"] == "Bob"
        assert comments[2]["author"] == "Alice"
    finally:
        os.unlink(path)


def test_no_comments():
    """Sheet with no comments returns empty list."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.add_sheet("Sheet1")
            w.write_row(["data"])

        sheets = opensheet_core.read_xlsx(path)
        assert sheets[0]["comments"] == []
    finally:
        os.unlink(path)


def test_write_and_read_hyperlink():
    """Write a hyperlink and read it back."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.add_sheet("Sheet1")
            w.write_row(["Click here"])
            w.add_hyperlink("A1", "https://example.com", tooltip="Example")

        sheets = opensheet_core.read_xlsx(path)
        hyperlinks = sheets[0]["hyperlinks"]
        assert len(hyperlinks) == 1
        assert hyperlinks[0]["cell"] == "A1"
        assert hyperlinks[0]["url"] == "https://example.com"
        assert hyperlinks[0]["tooltip"] == "Example"
    finally:
        os.unlink(path)


def test_hyperlink_without_tooltip():
    """Write a hyperlink without tooltip."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.add_sheet("Sheet1")
            w.write_row(["Link"])
            w.add_hyperlink("A1", "https://example.com")

        sheets = opensheet_core.read_xlsx(path)
        hyperlinks = sheets[0]["hyperlinks"]
        assert len(hyperlinks) == 1
        assert hyperlinks[0]["cell"] == "A1"
        assert hyperlinks[0]["url"] == "https://example.com"
        assert hyperlinks[0]["tooltip"] is None
    finally:
        os.unlink(path)


def test_multiple_hyperlinks():
    """Multiple hyperlinks on different cells."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.add_sheet("Sheet1")
            w.write_row(["Link1", "Link2"])
            w.add_hyperlink("A1", "https://example.com")
            w.add_hyperlink("B1", "https://other.com", tooltip="Other")

        sheets = opensheet_core.read_xlsx(path)
        hyperlinks = sheets[0]["hyperlinks"]
        assert len(hyperlinks) == 2
        assert hyperlinks[0]["url"] == "https://example.com"
        assert hyperlinks[1]["url"] == "https://other.com"
    finally:
        os.unlink(path)


def test_no_hyperlinks():
    """Sheet with no hyperlinks returns empty list."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.add_sheet("Sheet1")
            w.write_row(["data"])

        sheets = opensheet_core.read_xlsx(path)
        assert sheets[0]["hyperlinks"] == []
    finally:
        os.unlink(path)


def test_comment_requires_open_sheet():
    """Adding a comment without an open sheet raises an error."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        w = opensheet_core.XlsxWriter(path)
        try:
            w.add_comment("A1", "Author", "text")
            assert False, "Should have raised"
        except Exception as e:
            assert "No sheet is open" in str(e)
        w.close()
    finally:
        os.unlink(path)


def test_hyperlink_requires_open_sheet():
    """Adding a hyperlink without an open sheet raises an error."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        w = opensheet_core.XlsxWriter(path)
        try:
            w.add_hyperlink("A1", "https://example.com")
            assert False, "Should have raised"
        except Exception as e:
            assert "No sheet is open" in str(e)
        w.close()
    finally:
        os.unlink(path)
