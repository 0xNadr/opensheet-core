import datetime
import pytest
import opensheet_core
from opensheet_core import CellStyle, StyledCell, XlsxWriter, Formula


@pytest.fixture
def tmp_xlsx(tmp_path):
    return str(tmp_path / "styled.xlsx")


class TestCellStyleCreation:
    def test_default_style(self):
        style = CellStyle()
        assert style.bold is False
        assert style.italic is False
        assert style.underline is False
        assert style.font_name is None
        assert style.font_size is None
        assert style.font_color is None
        assert style.fill_color is None
        assert style.border_left is None
        assert style.border_right is None
        assert style.border_top is None
        assert style.border_bottom is None
        assert style.border_color is None
        assert style.horizontal_alignment is None
        assert style.vertical_alignment is None
        assert style.wrap_text is False
        assert style.text_rotation is None
        assert style.number_format is None

    def test_bold_style(self):
        style = CellStyle(bold=True)
        assert style.bold is True

    def test_font_properties(self):
        style = CellStyle(font_name="Arial", font_size=14.0, font_color="FF0000")
        assert style.font_name == "Arial"
        assert style.font_size == 14.0
        assert style.font_color == "FF0000"

    def test_border_shorthand(self):
        """The `border` param sets all four sides."""
        style = CellStyle(border="thin")
        assert style.border_left == "thin"
        assert style.border_right == "thin"
        assert style.border_top == "thin"
        assert style.border_bottom == "thin"

    def test_border_shorthand_with_override(self):
        """Individual border side overrides the shorthand."""
        style = CellStyle(border="thin", border_left="medium")
        assert style.border_left == "medium"
        assert style.border_right == "thin"
        assert style.border_top == "thin"
        assert style.border_bottom == "thin"

    def test_alignment(self):
        style = CellStyle(
            horizontal_alignment="center",
            vertical_alignment="top",
            wrap_text=True,
            text_rotation=45,
        )
        assert style.horizontal_alignment == "center"
        assert style.vertical_alignment == "top"
        assert style.wrap_text is True
        assert style.text_rotation == 45

    def test_number_format(self):
        style = CellStyle(number_format="$#,##0.00")
        assert style.number_format == "$#,##0.00"

    def test_repr(self):
        style = CellStyle(bold=True, fill_color="FFFF00")
        r = repr(style)
        assert "bold=True" in r
        assert "fill_color=" in r

    def test_equality(self):
        s1 = CellStyle(bold=True, fill_color="FF0000")
        s2 = CellStyle(bold=True, fill_color="FF0000")
        s3 = CellStyle(bold=False, fill_color="FF0000")
        assert s1 == s2
        assert s1 != s3


class TestStyledCellCreation:
    def test_basic(self):
        style = CellStyle(bold=True)
        cell = StyledCell("Hello", style)
        assert cell.value == "Hello"
        assert cell.style.bold is True

    def test_with_number(self):
        style = CellStyle(fill_color="FFFF00")
        cell = StyledCell(42, style)
        assert cell.value == 42

    def test_repr(self):
        cell = StyledCell("x", CellStyle(italic=True))
        r = repr(cell)
        assert "StyledCell" in r

    def test_equality(self):
        s = CellStyle(bold=True)
        c1 = StyledCell("a", s)
        c2 = StyledCell("a", CellStyle(bold=True))
        c3 = StyledCell("b", CellStyle(bold=True))
        assert c1 == c2
        assert c1 != c3


class TestWriteStyledCells:
    def test_bold_roundtrip(self, tmp_xlsx):
        """Bold text writes and reads back as StyledCell."""
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Test")
            w.write_row([StyledCell("Bold", CellStyle(bold=True))])

        rows = opensheet_core.read_sheet(tmp_xlsx)
        cell = rows[0][0]
        assert isinstance(cell, StyledCell)
        assert cell.value == "Bold"
        assert cell.style.bold is True

    def test_italic_roundtrip(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Test")
            w.write_row([StyledCell("Italic", CellStyle(italic=True))])

        rows = opensheet_core.read_sheet(tmp_xlsx)
        cell = rows[0][0]
        assert isinstance(cell, StyledCell)
        assert cell.style.italic is True

    def test_underline_roundtrip(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Test")
            w.write_row([StyledCell("Under", CellStyle(underline=True))])

        rows = opensheet_core.read_sheet(tmp_xlsx)
        cell = rows[0][0]
        assert isinstance(cell, StyledCell)
        assert cell.style.underline is True

    def test_font_name_roundtrip(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Test")
            w.write_row([StyledCell("Arial", CellStyle(font_name="Arial"))])

        rows = opensheet_core.read_sheet(tmp_xlsx)
        cell = rows[0][0]
        assert isinstance(cell, StyledCell)
        assert cell.style.font_name == "Arial"

    def test_font_size_roundtrip(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Test")
            w.write_row([StyledCell("Big", CellStyle(font_size=16.0))])

        rows = opensheet_core.read_sheet(tmp_xlsx)
        cell = rows[0][0]
        assert isinstance(cell, StyledCell)
        assert cell.style.font_size == 16.0

    def test_font_color_roundtrip(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Test")
            w.write_row([StyledCell("Red", CellStyle(font_color="FF0000"))])

        rows = opensheet_core.read_sheet(tmp_xlsx)
        cell = rows[0][0]
        assert isinstance(cell, StyledCell)
        # Color is normalized to AARRGGBB
        assert cell.style.font_color == "FFFF0000"

    def test_fill_color_roundtrip(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Test")
            w.write_row([StyledCell(42, CellStyle(fill_color="FFFF00"))])

        rows = opensheet_core.read_sheet(tmp_xlsx)
        cell = rows[0][0]
        assert isinstance(cell, StyledCell)
        assert cell.value == 42
        assert cell.style.fill_color == "FFFFFF00"

    def test_border_thin_roundtrip(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Test")
            w.write_row([
                StyledCell("Bordered", CellStyle(border="thin", border_color="000000"))
            ])

        rows = opensheet_core.read_sheet(tmp_xlsx)
        cell = rows[0][0]
        assert isinstance(cell, StyledCell)
        assert cell.style.border_left == "thin"
        assert cell.style.border_right == "thin"
        assert cell.style.border_top == "thin"
        assert cell.style.border_bottom == "thin"

    def test_border_mixed_styles(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Test")
            w.write_row([
                StyledCell(
                    "Mixed",
                    CellStyle(
                        border_left="thin",
                        border_right="medium",
                        border_top="thick",
                        border_bottom="dashed",
                    ),
                )
            ])

        rows = opensheet_core.read_sheet(tmp_xlsx)
        cell = rows[0][0]
        assert isinstance(cell, StyledCell)
        assert cell.style.border_left == "thin"
        assert cell.style.border_right == "medium"
        assert cell.style.border_top == "thick"
        assert cell.style.border_bottom == "dashed"

    def test_alignment_roundtrip(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Test")
            w.write_row([
                StyledCell(
                    "Centered",
                    CellStyle(
                        horizontal_alignment="center",
                        vertical_alignment="top",
                        wrap_text=True,
                    ),
                )
            ])

        rows = opensheet_core.read_sheet(tmp_xlsx)
        cell = rows[0][0]
        assert isinstance(cell, StyledCell)
        assert cell.style.horizontal_alignment == "center"
        assert cell.style.vertical_alignment == "top"
        assert cell.style.wrap_text is True

    def test_text_rotation_roundtrip(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Test")
            w.write_row([StyledCell("Rotated", CellStyle(text_rotation=90))])

        rows = opensheet_core.read_sheet(tmp_xlsx)
        cell = rows[0][0]
        assert isinstance(cell, StyledCell)
        assert cell.style.text_rotation == 90

    def test_combined_style_roundtrip(self, tmp_xlsx):
        """Multiple style properties applied together."""
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Test")
            w.write_row([
                StyledCell(
                    "Full",
                    CellStyle(
                        bold=True,
                        italic=True,
                        font_color="0000FF",
                        fill_color="FFFF00",
                        border="thin",
                        border_color="000000",
                        horizontal_alignment="center",
                    ),
                )
            ])

        rows = opensheet_core.read_sheet(tmp_xlsx)
        cell = rows[0][0]
        assert isinstance(cell, StyledCell)
        assert cell.style.bold is True
        assert cell.style.italic is True
        assert cell.style.font_color == "FF0000FF"
        assert cell.style.fill_color == "FFFFFF00"
        assert cell.style.border_left == "thin"
        assert cell.style.horizontal_alignment == "center"

    def test_styled_number_roundtrip(self, tmp_xlsx):
        """Number value inside a StyledCell."""
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Test")
            w.write_row([StyledCell(3.14, CellStyle(bold=True))])

        rows = opensheet_core.read_sheet(tmp_xlsx)
        cell = rows[0][0]
        assert isinstance(cell, StyledCell)
        assert abs(cell.value - 3.14) < 0.001

    def test_styled_bool_roundtrip(self, tmp_xlsx):
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Test")
            w.write_row([StyledCell(True, CellStyle(bold=True))])

        rows = opensheet_core.read_sheet(tmp_xlsx)
        cell = rows[0][0]
        assert isinstance(cell, StyledCell)
        assert cell.value == True

    def test_styled_date_roundtrip(self, tmp_xlsx):
        """Date value inside a StyledCell."""
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Test")
            w.write_row([
                StyledCell(datetime.date(2025, 3, 15), CellStyle(bold=True))
            ])

        rows = opensheet_core.read_sheet(tmp_xlsx)
        cell = rows[0][0]
        assert isinstance(cell, StyledCell)
        assert cell.value == datetime.date(2025, 3, 15)

    def test_styled_with_number_format(self, tmp_xlsx):
        """StyledCell with number_format in the style."""
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Test")
            w.write_row([
                StyledCell(
                    1234.56,
                    CellStyle(bold=True, number_format="$#,##0.00"),
                )
            ])

        rows = opensheet_core.read_sheet(tmp_xlsx)
        cell = rows[0][0]
        assert isinstance(cell, StyledCell)
        assert abs(cell.value - 1234.56) < 0.01
        assert cell.style.bold is True
        assert cell.style.number_format == "$#,##0.00"

    def test_unstyled_cells_unchanged(self, tmp_xlsx):
        """Plain cells (no styling) don't become StyledCell on read."""
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Test")
            w.write_row(["Plain", 42, True])

        rows = opensheet_core.read_sheet(tmp_xlsx)
        assert rows[0][0] == "Plain"
        assert rows[0][1] == 42
        assert rows[0][2] == True

    def test_mixed_styled_and_plain(self, tmp_xlsx):
        """Row with both styled and plain cells."""
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Test")
            w.write_row([
                StyledCell("Styled", CellStyle(bold=True)),
                "Plain",
                StyledCell(99, CellStyle(fill_color="00FF00")),
            ])

        rows = opensheet_core.read_sheet(tmp_xlsx)
        assert isinstance(rows[0][0], StyledCell)
        assert rows[0][0].style.bold is True
        assert rows[0][1] == "Plain"
        assert isinstance(rows[0][2], StyledCell)
        assert rows[0][2].value == 99

    def test_style_dedup(self, tmp_xlsx):
        """Same style used for multiple cells only creates one xf entry."""
        style = CellStyle(bold=True)
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Test")
            w.write_row([
                StyledCell("A", style),
                StyledCell("B", style),
            ])

        rows = opensheet_core.read_sheet(tmp_xlsx)
        assert isinstance(rows[0][0], StyledCell)
        assert isinstance(rows[0][1], StyledCell)
        assert rows[0][0].style.bold is True
        assert rows[0][1].style.bold is True

    def test_multiple_different_styles(self, tmp_xlsx):
        """Different styles create separate xf entries."""
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Test")
            w.write_row([
                StyledCell("Bold", CellStyle(bold=True)),
                StyledCell("Italic", CellStyle(italic=True)),
                StyledCell("Both", CellStyle(bold=True, italic=True)),
            ])

        rows = opensheet_core.read_sheet(tmp_xlsx)
        assert rows[0][0].style.bold is True
        assert rows[0][0].style.italic is False
        assert rows[0][1].style.bold is False
        assert rows[0][1].style.italic is True
        assert rows[0][2].style.bold is True
        assert rows[0][2].style.italic is True


class TestReadXlsxStyling:
    def test_read_xlsx_returns_styled_cells(self, tmp_xlsx):
        """read_xlsx returns StyledCell in rows for styled cells."""
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Styled")
            w.write_row([StyledCell("Bold", CellStyle(bold=True))])

        sheets = opensheet_core.read_xlsx(tmp_xlsx)
        cell = sheets[0]["rows"][0][0]
        assert isinstance(cell, StyledCell)
        assert cell.style.bold is True


class TestPandasIntegration:
    def test_styled_cell_unwrapped_in_dataframe(self, tmp_xlsx):
        """Pandas read_xlsx_df should unwrap StyledCell to plain values."""
        pd = pytest.importorskip("pandas")
        with XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Test")
            w.write_row(["Name", "Value"])
            w.write_row([
                StyledCell("Alice", CellStyle(bold=True)),
                StyledCell(42, CellStyle(fill_color="FFFF00")),
            ])

        df = opensheet_core.read_xlsx_df(tmp_xlsx)
        assert df.iloc[0]["Name"] == "Alice"
        assert df.iloc[0]["Value"] == 42
