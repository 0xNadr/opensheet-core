"""LangChain document loader integration for OpenSheet Core.

Provides OpenSheetLoader, a LangChain-compatible document loader that
converts XLSX sheets into Document objects using opensheet-core's fast
Rust-powered reader.

Requires langchain-core to be installed:
    pip install langchain-core
"""

from __future__ import annotations

from typing import Iterator, Optional

from opensheet_core.extract import xlsx_to_markdown, xlsx_to_text, xlsx_to_chunks


def _check_langchain():
    try:
        from langchain_core.document_loaders import BaseLoader
        from langchain_core.documents import Document
        return BaseLoader, Document
    except ImportError:
        raise ImportError(
            "langchain-core is required for OpenSheetLoader. "
            "Install it with: pip install langchain-core"
        ) from None


class OpenSheetLoader:
    """LangChain document loader for XLSX files.

    Converts spreadsheets into LangChain Document objects using
    opensheet-core's streaming Rust reader.

    Args:
        file_path: Path to the XLSX file.
        mode: Output mode — ``"markdown"`` (default), ``"text"``, or
            ``"chunks"``. Markdown produces structured tables suitable for
            LLMs, text produces plain tab-separated output, and chunks
            splits into embedding-sized markdown tables.
        sheet_name: Name of a specific sheet to load.
        sheet_index: 0-based index of a specific sheet to load.
        header: If True (default), treat the first row as a header.
        max_rows: Maximum data rows per chunk (only used in chunks mode,
            default 50).
        delimiter: Cell separator for text mode (default tab).

    Example::

        from opensheet_core.langchain import OpenSheetLoader

        loader = OpenSheetLoader("data.xlsx")
        docs = loader.load()

        # Chunked for RAG
        loader = OpenSheetLoader("data.xlsx", mode="chunks", max_rows=25)
        docs = loader.load()
    """

    def __init__(
        self,
        file_path: str,
        mode: str = "markdown",
        sheet_name: Optional[str] = None,
        sheet_index: Optional[int] = None,
        header: bool = True,
        max_rows: int = 50,
        delimiter: str = "\t",
    ):
        _check_langchain()
        if mode not in ("markdown", "text", "chunks"):
            raise ValueError(
                f"Invalid mode {mode!r}. Must be 'markdown', 'text', or 'chunks'."
            )
        self.file_path = file_path
        self.mode = mode
        self.sheet_name = sheet_name
        self.sheet_index = sheet_index
        self.header = header
        self.max_rows = max_rows
        self.delimiter = delimiter

    def lazy_load(self) -> Iterator:
        """Yield Document objects lazily."""
        _, Document = _check_langchain()

        base_metadata = {"source": self.file_path}
        if self.sheet_name is not None:
            base_metadata["sheet_name"] = self.sheet_name
        if self.sheet_index is not None:
            base_metadata["sheet_index"] = self.sheet_index

        if self.mode == "chunks":
            chunks = xlsx_to_chunks(
                self.file_path,
                sheet_name=self.sheet_name,
                sheet_index=self.sheet_index,
                max_rows=self.max_rows,
                header=self.header,
            )
            for i, chunk in enumerate(chunks):
                metadata = {**base_metadata, "chunk_index": i}
                yield Document(page_content=chunk, metadata=metadata)
        elif self.mode == "text":
            content = xlsx_to_text(
                self.file_path,
                sheet_name=self.sheet_name,
                sheet_index=self.sheet_index,
                delimiter=self.delimiter,
            )
            yield Document(page_content=content, metadata=base_metadata)
        else:
            content = xlsx_to_markdown(
                self.file_path,
                sheet_name=self.sheet_name,
                sheet_index=self.sheet_index,
                header=self.header,
            )
            yield Document(page_content=content, metadata=base_metadata)

    def load(self) -> list:
        """Load all documents into a list."""
        return list(self.lazy_load())
