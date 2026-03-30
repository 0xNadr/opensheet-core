"""Tests for sheet visibility states (visible, hidden, veryHidden)."""

import pytest
from opensheet_core import XlsxWriter, read_xlsx, read_sheet, sheet_names


@pytest.fixture
def tmp_xlsx(tmp_path):
    return str(tmp_path / "output.xlsx")


class TestWriteSheetState:
    def test_default_state_is_visible(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Sheet1")
            w.write_row(["data"])
        sheets = read_xlsx(tmp_xlsx)
        assert sheets[0]["state"] == "visible"

    def test_hidden_sheet(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Visible")
            w.write_row(["visible data"])
            w.add_sheet("Hidden")
            w.set_sheet_state("hidden")
            w.write_row(["hidden data"])
        sheets = read_xlsx(tmp_xlsx)
        assert sheets[0]["state"] == "visible"
        assert sheets[1]["state"] == "hidden"

    def test_very_hidden_sheet(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Visible")
            w.write_row(["data"])
            w.add_sheet("Secret")
            w.set_sheet_state("veryHidden")
            w.write_row(["secret data"])
        sheets = read_xlsx(tmp_xlsx)
        assert sheets[0]["state"] == "visible"
        assert sheets[1]["state"] == "veryHidden"

    def test_mixed_states(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Public")
            w.write_row(["public"])
            w.add_sheet("Hidden")
            w.set_sheet_state("hidden")
            w.write_row(["hidden"])
            w.add_sheet("VeryHidden")
            w.set_sheet_state("veryHidden")
            w.write_row(["very hidden"])
            w.add_sheet("AlsoPublic")
            w.write_row(["also public"])
        sheets = read_xlsx(tmp_xlsx)
        assert sheets[0]["state"] == "visible"
        assert sheets[1]["state"] == "hidden"
        assert sheets[2]["state"] == "veryHidden"
        assert sheets[3]["state"] == "visible"

    def test_set_state_before_write(self, tmp_xlsx):
        """set_sheet_state can be called right after add_sheet."""
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Sheet1")
            w.set_sheet_state("hidden")
            w.write_row(["data"])
        sheets = read_xlsx(tmp_xlsx)
        assert sheets[0]["state"] == "hidden"

    def test_set_state_after_write(self, tmp_xlsx):
        """set_sheet_state can be called after writing rows."""
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Sheet1")
            w.write_row(["data"])
            w.set_sheet_state("hidden")
        sheets = read_xlsx(tmp_xlsx)
        assert sheets[0]["state"] == "hidden"

    def test_invalid_state_raises(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Sheet1")
            with pytest.raises(Exception, match="Invalid sheet state"):
                w.set_sheet_state("invisible")

    def test_set_state_without_sheet_raises(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            with pytest.raises(Exception, match="No sheet is open"):
                w.set_sheet_state("hidden")


class TestReadSheetState:
    def test_read_xlsx_includes_state(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Data")
            w.write_row(["val"])
        sheets = read_xlsx(tmp_xlsx)
        assert "state" in sheets[0]
        assert sheets[0]["state"] == "visible"

    def test_hidden_sheet_data_readable(self, tmp_xlsx):
        """Hidden sheets should still have their data readable."""
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Visible")
            w.write_row(["A", "B"])
            w.add_sheet("Hidden")
            w.set_sheet_state("hidden")
            w.write_row(["C", "D"])
        sheets = read_xlsx(tmp_xlsx)
        assert sheets[1]["rows"][0] == ["C", "D"]
        assert sheets[1]["state"] == "hidden"

    def test_read_sheet_by_name_hidden(self, tmp_xlsx):
        """read_sheet works for hidden sheets by name."""
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Visible")
            w.write_row(["A"])
            w.add_sheet("Hidden")
            w.set_sheet_state("hidden")
            w.write_row(["B"])
        rows = read_sheet(tmp_xlsx, sheet_name="Hidden")
        assert rows[0] == ["B"]

    def test_read_sheet_by_index_hidden(self, tmp_xlsx):
        """read_sheet works for hidden sheets by index."""
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Visible")
            w.write_row(["A"])
            w.add_sheet("Hidden")
            w.set_sheet_state("veryHidden")
            w.write_row(["B"])
        rows = read_sheet(tmp_xlsx, sheet_index=1)
        assert rows[0] == ["B"]

    def test_sheet_names_includes_hidden(self, tmp_xlsx):
        """sheet_names returns all sheets including hidden ones."""
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Visible")
            w.write_row(["A"])
            w.add_sheet("Hidden")
            w.set_sheet_state("hidden")
            w.write_row(["B"])
        names = sheet_names(tmp_xlsx)
        assert names == ["Visible", "Hidden"]


class TestRoundTrip:
    def test_roundtrip_preserves_state(self, tmp_xlsx, tmp_path):
        """Write with states, read, verify states preserved."""
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Visible")
            w.write_row(["data"])
            w.add_sheet("Hidden")
            w.set_sheet_state("hidden")
            w.write_row(["hidden"])
            w.add_sheet("VeryHidden")
            w.set_sheet_state("veryHidden")
            w.write_row(["secret"])

        sheets = read_xlsx(tmp_xlsx)
        states = [(s["name"], s["state"]) for s in sheets]
        assert states == [
            ("Visible", "visible"),
            ("Hidden", "hidden"),
            ("VeryHidden", "veryHidden"),
        ]
