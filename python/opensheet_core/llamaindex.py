"""LlamaIndex data reader integration for OpenSheet Core.

Provides OpenSheetReader, a LlamaIndex-compatible data reader that
converts XLSX sheets into Document objects using opensheet-core's fast
Rust-powered reader.

Requires llama-index-core to be installed:
    pip install llama-index-core
"""

from __future__ import annotations

from typing import List, Optional

from opensheet_core.extract import xlsx_to_markdown, xlsx_to_text, xlsx_to_chunks


def _check_llamaindex():
    try:
        from llama_index.core.readers.base import BaseReader
        from llama_index.core.schema import Document
        return BaseReader, Document
    except ImportError:
        raise ImportError(
            "llama-index-core is required for OpenSheetReader. "
            "Install it with: pip install llama-index-core"
        ) from None


class OpenSheetReader:
    """LlamaIndex data reader for XLSX files.

    Converts spreadsheets into LlamaIndex Document objects using
    opensheet-core's streaming Rust reader.

    Args:
        mode: Output mode — ``"markdown"`` (default), ``"text"``, or
            ``"chunks"``. Markdown produces structured tables suitable for
            LLMs, text produces plain tab-separated output, and chunks
            splits into embedding-sized markdown tables.
        header: If True (default), treat the first row as a header.
        max_rows: Maximum data rows per chunk (only used in chunks mode,
            default 50).
        delimiter: Cell separator for text mode (default tab).

    Example::

        from opensheet_core.llamaindex import OpenSheetReader

        reader = OpenSheetReader()
        docs = reader.load_data("data.xlsx")

        # Chunked for RAG
        reader = OpenSheetReader(mode="chunks", max_rows=25)
        docs = reader.load_data("data.xlsx")

        # Use with SimpleDirectoryReader
        from llama_index.core import SimpleDirectoryReader
        reader = SimpleDirectoryReader(
            input_dir="./data",
            file_extractor={".xlsx": OpenSheetReader()},
        )
    """

    def __init__(
        self,
        mode: str = "markdown",
        header: bool = True,
        max_rows: int = 50,
        delimiter: str = "\t",
    ):
        _check_llamaindex()
        if mode not in ("markdown", "text", "chunks"):
            raise ValueError(
                f"Invalid mode {mode!r}. Must be 'markdown', 'text', or 'chunks'."
            )
        self.mode = mode
        self.header = header
        self.max_rows = max_rows
        self.delimiter = delimiter

    def load_data(
        self,
        file: str,
        sheet_name: Optional[str] = None,
        sheet_index: Optional[int] = None,
        extra_info: Optional[dict] = None,
    ) -> List:
        """Load an XLSX file and return a list of Document objects.

        Args:
            file: Path to the XLSX file (str or Path).
            sheet_name: Name of a specific sheet to load.
            sheet_index: 0-based index of a specific sheet to load.
            extra_info: Additional metadata to attach to each document.

        Returns:
            A list of LlamaIndex Document objects.
        """
        _, Document = _check_llamaindex()
        file_path = str(file)

        base_metadata = {"file_name": file_path}
        if sheet_name is not None:
            base_metadata["sheet_name"] = sheet_name
        if sheet_index is not None:
            base_metadata["sheet_index"] = sheet_index
        if extra_info:
            base_metadata.update(extra_info)

        if self.mode == "chunks":
            chunks = xlsx_to_chunks(
                file_path,
                sheet_name=sheet_name,
                sheet_index=sheet_index,
                max_rows=self.max_rows,
                header=self.header,
            )
            return [
                Document(text=chunk, metadata={**base_metadata, "chunk_index": i})
                for i, chunk in enumerate(chunks)
            ]
        elif self.mode == "text":
            content = xlsx_to_text(
                file_path,
                sheet_name=sheet_name,
                sheet_index=sheet_index,
                delimiter=self.delimiter,
            )
            return [Document(text=content, metadata=base_metadata)]
        else:
            content = xlsx_to_markdown(
                file_path,
                sheet_name=sheet_name,
                sheet_index=sheet_index,
                header=self.header,
            )
            return [Document(text=content, metadata=base_metadata)]
