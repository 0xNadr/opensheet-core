"""Tests for named ranges / defined names (read and write)."""

import pytest
from opensheet_core import XlsxWriter, defined_names


@pytest.fixture
def tmp_xlsx(tmp_path):
    return str(tmp_path / "output.xlsx")


class TestWriteDefinedNames:
    def test_workbook_scoped_name(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Sheet1")
            w.write_row(["Rate"])
            w.write_row([0.08])
            w.define_name("TaxRate", "Sheet1!$A$2")
        names = defined_names(tmp_xlsx)
        assert len(names) == 1
        assert names[0]["name"] == "TaxRate"
        assert names[0]["value"] == "Sheet1!$A$2"
        assert names[0]["sheet_index"] is None

    def test_sheet_scoped_name(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Sheet1")
            w.write_row(["data"])
            w.define_name("LocalName", "Sheet1!$A$1", sheet_index=0)
        names = defined_names(tmp_xlsx)
        assert len(names) == 1
        assert names[0]["name"] == "LocalName"
        assert names[0]["value"] == "Sheet1!$A$1"
        assert names[0]["sheet_index"] == 0

    def test_multiple_names(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Data")
            w.write_row(["A", "B", "C"])
            w.write_row([1, 2, 3])
            w.define_name("ColA", "Data!$A:$A")
            w.define_name("ColB", "Data!$B:$B")
            w.define_name("Header", "Data!$A$1:$C$1")
        names = defined_names(tmp_xlsx)
        assert len(names) == 3
        name_map = {n["name"]: n["value"] for n in names}
        assert name_map["ColA"] == "Data!$A:$A"
        assert name_map["ColB"] == "Data!$B:$B"
        assert name_map["Header"] == "Data!$A$1:$C$1"

    def test_mixed_scopes(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Sheet1")
            w.write_row(["a"])
            w.add_sheet("Sheet2")
            w.write_row(["b"])
            w.define_name("Global", "Sheet1!$A$1")
            w.define_name("Local1", "Sheet1!$A$1", sheet_index=0)
            w.define_name("Local2", "Sheet2!$A$1", sheet_index=1)
        names = defined_names(tmp_xlsx)
        assert len(names) == 3
        by_name = {n["name"]: n for n in names}
        assert by_name["Global"]["sheet_index"] is None
        assert by_name["Local1"]["sheet_index"] == 0
        assert by_name["Local2"]["sheet_index"] == 1

    def test_define_name_before_sheets(self, tmp_xlsx):
        """define_name can be called before or after adding sheets."""
        with XlsxWriter(tmp_xlsx) as w:
            w.define_name("Early", "Sheet1!$A$1")
            w.add_sheet("Sheet1")
            w.write_row(["data"])
        names = defined_names(tmp_xlsx)
        assert len(names) == 1
        assert names[0]["name"] == "Early"

    def test_define_name_after_writes(self, tmp_xlsx):
        """define_name can be called after writing rows."""
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Sheet1")
            w.write_row(["data"])
            w.define_name("Late", "Sheet1!$A$1")
        names = defined_names(tmp_xlsx)
        assert len(names) == 1
        assert names[0]["name"] == "Late"

    def test_empty_name_raises(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Sheet1")
            with pytest.raises(Exception, match="empty"):
                w.define_name("", "Sheet1!$A$1")

    def test_empty_value_raises(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Sheet1")
            with pytest.raises(Exception, match="empty"):
                w.define_name("Name", "")

    def test_no_defined_names(self, tmp_xlsx):
        """File with no defined names returns empty list."""
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Sheet1")
            w.write_row(["data"])
        names = defined_names(tmp_xlsx)
        assert names == []


class TestReadDefinedNames:
    def test_roundtrip_workbook_scoped(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Sales")
            w.write_row(["Revenue"])
            w.write_row([1000])
            w.define_name("TotalRevenue", "Sales!$A$2")
        names = defined_names(tmp_xlsx)
        assert len(names) == 1
        assert names[0] == {"name": "TotalRevenue", "value": "Sales!$A$2", "sheet_index": None}

    def test_roundtrip_sheet_scoped(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Config")
            w.write_row(["Rate"])
            w.write_row([0.05])
            w.define_name("Rate", "Config!$A$2", sheet_index=0)
        names = defined_names(tmp_xlsx)
        assert len(names) == 1
        assert names[0] == {"name": "Rate", "value": "Config!$A$2", "sheet_index": 0}

    def test_roundtrip_multiple(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Sheet1")
            w.write_row(["a", "b"])
            w.add_sheet("Sheet2")
            w.write_row(["c", "d"])
            w.define_name("AllData", "Sheet1!$A$1:$B$1")
            w.define_name("SheetLocal", "Sheet2!$A$1", sheet_index=1)
        names = defined_names(tmp_xlsx)
        assert len(names) == 2
        by_name = {n["name"]: n for n in names}
        assert by_name["AllData"]["value"] == "Sheet1!$A$1:$B$1"
        assert by_name["AllData"]["sheet_index"] is None
        assert by_name["SheetLocal"]["value"] == "Sheet2!$A$1"
        assert by_name["SheetLocal"]["sheet_index"] == 1

    def test_special_characters_in_name(self, tmp_xlsx):
        """Names with special XML characters should round-trip correctly."""
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Sheet1")
            w.write_row(["data"])
            w.define_name("Tax_Rate", "Sheet1!$A$1")
        names = defined_names(tmp_xlsx)
        assert names[0]["name"] == "Tax_Rate"

    def test_range_value(self, tmp_xlsx):
        """Named range referencing a multi-cell range."""
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Data")
            w.write_row(["a", "b", "c"])
            w.write_row([1, 2, 3])
            w.define_name("DataRange", "Data!$A$1:$C$2")
        names = defined_names(tmp_xlsx)
        assert names[0]["value"] == "Data!$A$1:$C$2"
