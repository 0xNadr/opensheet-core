"""Tests for LlamaIndex OpenSheetReader integration."""

import sys
import pytest
import opensheet_core
from opensheet_core import XlsxWriter


@pytest.fixture
def tmp_xlsx(tmp_path):
    return str(tmp_path / "output.xlsx")


def _write_basic(path):
    with XlsxWriter(path) as w:
        w.add_sheet("Data")
        w.write_row(["Name", "Age"])
        w.write_row(["Alice", 30])
        w.write_row(["Bob", 25])


def _write_multi_sheet(path):
    with XlsxWriter(path) as w:
        w.add_sheet("Users")
        w.write_row(["Name", "Age"])
        w.write_row(["Alice", 30])
        w.add_sheet("Items")
        w.write_row(["Item", "Price"])
        w.write_row(["Widget", 9.99])


# Check if llama_index is available
try:
    from llama_index.core.readers.base import BaseReader
    from llama_index.core.schema import Document
    HAS_LLAMAINDEX = True
except ImportError:
    HAS_LLAMAINDEX = False


@pytest.mark.skipif(not HAS_LLAMAINDEX, reason="llama-index-core not installed")
class TestOpenSheetReader:
    def test_load_markdown_default(self, tmp_xlsx):
        from opensheet_core.llamaindex import OpenSheetReader
        _write_basic(tmp_xlsx)
        reader = OpenSheetReader()
        docs = reader.load_data(tmp_xlsx)
        assert len(docs) == 1
        assert "Name" in docs[0].text
        assert "Alice" in docs[0].text
        assert docs[0].metadata["file_name"] == tmp_xlsx

    def test_load_text_mode(self, tmp_xlsx):
        from opensheet_core.llamaindex import OpenSheetReader
        _write_basic(tmp_xlsx)
        reader = OpenSheetReader(mode="text")
        docs = reader.load_data(tmp_xlsx)
        assert len(docs) == 1
        assert "Name\tAge" in docs[0].text

    def test_load_chunks_mode(self, tmp_xlsx):
        from opensheet_core.llamaindex import OpenSheetReader
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Data")
            w.write_row(["ID", "Val"])
            for i in range(10):
                w.write_row([i, i * 10])
        reader = OpenSheetReader(mode="chunks", max_rows=3)
        docs = reader.load_data(tmp_xlsx)
        assert len(docs) == 4
        for i, doc in enumerate(docs):
            assert "ID" in doc.text
            assert doc.metadata["chunk_index"] == i

    def test_sheet_by_name(self, tmp_xlsx):
        from opensheet_core.llamaindex import OpenSheetReader
        _write_multi_sheet(tmp_xlsx)
        reader = OpenSheetReader()
        docs = reader.load_data(tmp_xlsx, sheet_name="Items")
        assert "Widget" in docs[0].text
        assert "Alice" not in docs[0].text
        assert docs[0].metadata["sheet_name"] == "Items"

    def test_sheet_by_index(self, tmp_xlsx):
        from opensheet_core.llamaindex import OpenSheetReader
        _write_multi_sheet(tmp_xlsx)
        reader = OpenSheetReader()
        docs = reader.load_data(tmp_xlsx, sheet_index=1)
        assert "Widget" in docs[0].text
        assert docs[0].metadata["sheet_index"] == 1

    def test_extra_info(self, tmp_xlsx):
        from opensheet_core.llamaindex import OpenSheetReader
        _write_basic(tmp_xlsx)
        reader = OpenSheetReader()
        docs = reader.load_data(tmp_xlsx, extra_info={"author": "test"})
        assert docs[0].metadata["author"] == "test"
        assert docs[0].metadata["file_name"] == tmp_xlsx

    def test_invalid_mode(self):
        from opensheet_core.llamaindex import OpenSheetReader
        with pytest.raises(ValueError, match="Invalid mode"):
            OpenSheetReader(mode="invalid")

    def test_custom_delimiter(self, tmp_xlsx):
        from opensheet_core.llamaindex import OpenSheetReader
        _write_basic(tmp_xlsx)
        reader = OpenSheetReader(mode="text", delimiter=",")
        docs = reader.load_data(tmp_xlsx)
        assert "Name,Age" in docs[0].text

    def test_no_header(self, tmp_xlsx):
        from opensheet_core.llamaindex import OpenSheetReader
        _write_basic(tmp_xlsx)
        reader = OpenSheetReader(header=False)
        docs = reader.load_data(tmp_xlsx)
        assert "Col 0" in docs[0].text

    def test_path_object(self, tmp_xlsx):
        """Accepts Path objects as well as strings."""
        from pathlib import Path
        from opensheet_core.llamaindex import OpenSheetReader
        _write_basic(tmp_xlsx)
        reader = OpenSheetReader()
        docs = reader.load_data(Path(tmp_xlsx))
        assert len(docs) == 1
        assert "Alice" in docs[0].text


class TestLlamaIndexImportError:
    def test_import_error_without_llamaindex(self, tmp_xlsx, monkeypatch):
        _write_basic(tmp_xlsx)
        # Temporarily hide llama_index
        monkeypatch.setitem(sys.modules, "llama_index", None)
        monkeypatch.setitem(sys.modules, "llama_index.core", None)
        monkeypatch.setitem(sys.modules, "llama_index.core.readers", None)
        monkeypatch.setitem(sys.modules, "llama_index.core.readers.base", None)
        monkeypatch.setitem(sys.modules, "llama_index.core.schema", None)
        from opensheet_core.llamaindex import OpenSheetReader
        with pytest.raises(ImportError, match="llama-index-core"):
            OpenSheetReader()
