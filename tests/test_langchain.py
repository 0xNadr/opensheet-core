"""Tests for LangChain OpenSheetLoader integration."""

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


# Check if langchain_core is available
try:
    from langchain_core.document_loaders import BaseLoader
    from langchain_core.documents import Document
    HAS_LANGCHAIN = True
except ImportError:
    HAS_LANGCHAIN = False


@pytest.mark.skipif(not HAS_LANGCHAIN, reason="langchain-core not installed")
class TestOpenSheetLoader:
    def test_load_markdown_default(self, tmp_xlsx):
        from opensheet_core.langchain import OpenSheetLoader
        _write_basic(tmp_xlsx)
        loader = OpenSheetLoader(tmp_xlsx)
        docs = loader.load()
        assert len(docs) == 1
        assert "Name" in docs[0].page_content
        assert "Alice" in docs[0].page_content
        assert docs[0].metadata["source"] == tmp_xlsx

    def test_load_text_mode(self, tmp_xlsx):
        from opensheet_core.langchain import OpenSheetLoader
        _write_basic(tmp_xlsx)
        loader = OpenSheetLoader(tmp_xlsx, mode="text")
        docs = loader.load()
        assert len(docs) == 1
        assert "Name\tAge" in docs[0].page_content

    def test_load_chunks_mode(self, tmp_xlsx):
        from opensheet_core.langchain import OpenSheetLoader
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Data")
            w.write_row(["ID", "Val"])
            for i in range(10):
                w.write_row([i, i * 10])
        loader = OpenSheetLoader(tmp_xlsx, mode="chunks", max_rows=3)
        docs = loader.load()
        assert len(docs) == 4
        for i, doc in enumerate(docs):
            assert "ID" in doc.page_content
            assert doc.metadata["chunk_index"] == i
            assert doc.metadata["source"] == tmp_xlsx

    def test_lazy_load(self, tmp_xlsx):
        from opensheet_core.langchain import OpenSheetLoader
        _write_basic(tmp_xlsx)
        loader = OpenSheetLoader(tmp_xlsx)
        docs = list(loader.lazy_load())
        assert len(docs) == 1

    def test_sheet_by_name(self, tmp_xlsx):
        from opensheet_core.langchain import OpenSheetLoader
        _write_multi_sheet(tmp_xlsx)
        loader = OpenSheetLoader(tmp_xlsx, sheet_name="Items")
        docs = loader.load()
        assert "Widget" in docs[0].page_content
        assert "Alice" not in docs[0].page_content
        assert docs[0].metadata["sheet_name"] == "Items"

    def test_sheet_by_index(self, tmp_xlsx):
        from opensheet_core.langchain import OpenSheetLoader
        _write_multi_sheet(tmp_xlsx)
        loader = OpenSheetLoader(tmp_xlsx, sheet_index=1)
        docs = loader.load()
        assert "Widget" in docs[0].page_content
        assert docs[0].metadata["sheet_index"] == 1

    def test_invalid_mode(self, tmp_xlsx):
        from opensheet_core.langchain import OpenSheetLoader
        with pytest.raises(ValueError, match="Invalid mode"):
            OpenSheetLoader(tmp_xlsx, mode="invalid")

    def test_custom_delimiter(self, tmp_xlsx):
        from opensheet_core.langchain import OpenSheetLoader
        _write_basic(tmp_xlsx)
        loader = OpenSheetLoader(tmp_xlsx, mode="text", delimiter=",")
        docs = loader.load()
        assert "Name,Age" in docs[0].page_content

    def test_no_header(self, tmp_xlsx):
        from opensheet_core.langchain import OpenSheetLoader
        _write_basic(tmp_xlsx)
        loader = OpenSheetLoader(tmp_xlsx, header=False)
        docs = loader.load()
        assert "Col 0" in docs[0].page_content


class TestLangChainImportError:
    def test_import_error_without_langchain(self, tmp_xlsx, monkeypatch):
        _write_basic(tmp_xlsx)
        # Temporarily hide langchain_core
        monkeypatch.setitem(sys.modules, "langchain_core", None)
        monkeypatch.setitem(sys.modules, "langchain_core.document_loaders", None)
        monkeypatch.setitem(sys.modules, "langchain_core.documents", None)
        from opensheet_core.langchain import OpenSheetLoader
        with pytest.raises(ImportError, match="langchain-core"):
            OpenSheetLoader(tmp_xlsx)
