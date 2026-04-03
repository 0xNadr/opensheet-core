#!/usr/bin/env python3
"""Generate diverse XLSX test corpus for broader coverage and fuzz seeding.

Creates files that test edge cases across different Excel features,
data types, and structural variations. Each file is a valid XLSX that
exercises a specific code path in the reader.
"""

import os
import struct
import zipfile
from io import BytesIO
from pathlib import Path

FIXTURES = Path(__file__).parent / "fixtures"


def _minimal_xlsx(sheets_xml: dict[str, str], shared_strings: str | None = None,
                  workbook_xml: str | None = None, extra_files: dict[str, str] | None = None) -> bytes:
    """Build a minimal XLSX from raw XML parts."""
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        # [Content_Types].xml
        overrides = ""
        for name in sheets_xml:
            idx = int(name.replace("sheet", "").replace(".xml", ""))
            overrides += f'<Override PartName="/xl/worksheets/{name}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>\n'
        if shared_strings:
            overrides += '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>\n'
        zf.writestr("[Content_Types].xml", f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
{overrides}
</Types>""")

        # _rels/.rels
        zf.writestr("_rels/.rels", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>""")

        # xl/_rels/workbook.xml.rels
        rels = ""
        for i, name in enumerate(sheets_xml, 1):
            rels += f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/{name}"/>\n'
        if shared_strings:
            rels += f'<Relationship Id="rId{len(sheets_xml)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>\n'
        zf.writestr("xl/_rels/workbook.xml.rels", f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{rels}
</Relationships>""")

        # xl/workbook.xml
        if workbook_xml is None:
            sheet_entries = ""
            for i, name in enumerate(sheets_xml, 1):
                sheet_entries += f'<sheet name="Sheet{i}" sheetId="{i}" r:id="rId{i}"/>\n'
            workbook_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>{sheet_entries}</sheets>
</workbook>"""
        zf.writestr("xl/workbook.xml", workbook_xml)

        # Worksheets
        for name, xml in sheets_xml.items():
            zf.writestr(f"xl/worksheets/{name}", xml)

        # Shared strings
        if shared_strings:
            zf.writestr("xl/sharedStrings.xml", shared_strings)

        # Extra files
        if extra_files:
            for path, content in extra_files.items():
                zf.writestr(path, content)

    return buf.getvalue()


def _sheet_xml(rows_xml: str, merges: str = "", extras: str = "") -> str:
    """Wrap row XML in a full worksheet element."""
    merge_block = f"<mergeCells>{merges}</mergeCells>" if merges else ""
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
{extras}
<sheetData>{rows_xml}</sheetData>
{merge_block}
</worksheet>"""


def gen_empty_workbook():
    """Workbook with a single sheet containing no data."""
    return _minimal_xlsx({"sheet1.xml": _sheet_xml("")})


def gen_single_cell():
    """Workbook with exactly one cell."""
    rows = '<row r="1"><c r="A1"><v>42</v></c></row>'
    return _minimal_xlsx({"sheet1.xml": _sheet_xml(rows)})


def gen_many_data_types():
    """Cells with every data type: number, bool, string, inline string, formula, empty."""
    ss = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
<si><t>hello</t></si>
<si><t>world</t></si>
</sst>"""
    rows = """
<row r="1">
  <c r="A1"><v>3.14</v></c>
  <c r="B1" t="b"><v>1</v></c>
  <c r="C1" t="s"><v>0</v></c>
  <c r="D1" t="s"><v>1</v></c>
  <c r="E1" t="str"><v>inline</v></c>
</row>
<row r="2">
  <c r="A2"><f>SUM(A1:A1)</f><v>3.14</v></c>
  <c r="B2" t="e"><v>#REF!</v></c>
  <c r="C2"/>
</row>"""
    return _minimal_xlsx({"sheet1.xml": _sheet_xml(rows)}, shared_strings=ss)


def gen_multiple_sheets():
    """Workbook with 5 sheets including hidden ones."""
    sheets = {}
    for i in range(1, 6):
        rows = f'<row r="1"><c r="A1"><v>{i}</v></c></row>'
        sheets[f"sheet{i}.xml"] = _sheet_xml(rows)

    sheet_entries = ""
    for i in range(1, 6):
        state = ' state="hidden"' if i == 3 else (' state="veryHidden"' if i == 5 else "")
        sheet_entries += f'<sheet name="Sheet{i}" sheetId="{i}" r:id="rId{i}"{state}/>\n'
    wb = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>{sheet_entries}</sheets>
</workbook>"""
    return _minimal_xlsx(sheets, workbook_xml=wb)


def gen_merged_cells():
    """Sheet with merged cell regions."""
    rows = """
<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>
<row r="2"><c r="A2"><v>3</v></c></row>
<row r="3"><c r="A3"><v>4</v></c></row>"""
    merges = '<mergeCell ref="A1:B2"/><mergeCell ref="A3:B3"/>'
    return _minimal_xlsx({"sheet1.xml": _sheet_xml(rows, merges=merges)})


def gen_sparse_rows():
    """Sheet with gaps in row numbers (rows 1, 5, 1000)."""
    rows = """
<row r="1"><c r="A1"><v>1</v></c></row>
<row r="5"><c r="C5"><v>5</v></c></row>
<row r="1000"><c r="Z1000"><v>999</v></c></row>"""
    return _minimal_xlsx({"sheet1.xml": _sheet_xml(rows)})


def gen_unicode_strings():
    """Shared strings with Unicode: emoji, CJK, RTL, combining chars."""
    strings = [
        "Hello 🌍🎉",
        "日本語テスト",
        "العربية",
        "café résumé naïve",
        "Z̤͔ͧ̑a̴̬l̶̞g̗̲ͧo̙̫",  # Combining characters
        "",  # Empty string
        " ",  # Space only
        "\t\n\r",  # Whitespace
        "a" * 32767,  # Max Excel cell length
    ]
    si_entries = "".join(f"<si><t xml:space=\"preserve\">{s}</t></si>" for s in strings)
    ss = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{len(strings)}" uniqueCount="{len(strings)}">
{si_entries}
</sst>"""
    cells = "".join(f'<c r="{chr(65+i)}1" t="s"><v>{i}</v></c>' for i in range(len(strings)))
    rows = f'<row r="1">{cells}</row>'
    return _minimal_xlsx({"sheet1.xml": _sheet_xml(rows)}, shared_strings=ss)


def gen_rich_text_strings():
    """Shared strings with rich text runs (<r> elements)."""
    ss = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
<si><r><rPr><b/><sz val="12"/></rPr><t>Bold</t></r><r><t> Normal</t></r></si>
<si><r><rPr><i/><color rgb="FFFF0000"/></rPr><t>Red Italic</t></r></si>
</sst>"""
    rows = '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>'
    return _minimal_xlsx({"sheet1.xml": _sheet_xml(rows)}, shared_strings=ss)


def gen_freeze_panes():
    """Sheet with freeze panes set."""
    extras = '<sheetViews><sheetView tabSelected="1"><pane ySplit="1" xSplit="2" topLeftCell="C2" state="frozen"/></sheetView></sheetViews>'
    rows = '<row r="1"><c r="A1"><v>1</v></c></row>'
    return _minimal_xlsx({"sheet1.xml": _sheet_xml(rows, extras=extras)})


def gen_auto_filter():
    """Sheet with auto-filter on a range."""
    extras = ""
    rows = """
<row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Age</t></is></c></row>
<row r="2"><c r="A2" t="inlineStr"><is><t>Alice</t></is></c><c r="B2"><v>30</v></c></row>"""
    af = '<autoFilter ref="A1:B2"/>'
    sheet = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>{rows}</sheetData>
{af}
</worksheet>"""
    return _minimal_xlsx({"sheet1.xml": sheet})


def gen_defined_names():
    """Workbook with named ranges / defined names."""
    sheets = {"sheet1.xml": _sheet_xml('<row r="1"><c r="A1"><v>1</v></c></row>')}
    wb = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
<definedNames>
<definedName name="MyRange">Sheet1!$A$1:$C$10</definedName>
<definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$Z$100</definedName>
<definedName name="Hidden" hidden="1">Sheet1!$B$2</definedName>
</definedNames>
</workbook>"""
    return _minimal_xlsx(sheets, workbook_xml=wb)


def gen_large_shared_strings():
    """Large shared string table (1000 entries)."""
    entries = "".join(f"<si><t>string_{i:04d}</t></si>" for i in range(1000))
    ss = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1000" uniqueCount="1000">
{entries}
</sst>"""
    cells = "".join(f'<c r="{chr(65 + (i % 26))}{i // 26 + 1}" t="s"><v>{i}</v></c>' for i in range(100))
    rows = ""
    for r in range(1, 5):
        row_cells = "".join(f'<c r="{chr(65 + c)}{r}" t="s"><v>{(r-1)*26+c}</v></c>' for c in range(26))
        rows += f'<row r="{r}">{row_cells}</row>'
    return _minimal_xlsx({"sheet1.xml": _sheet_xml(rows)}, shared_strings=ss)


def gen_date_cells():
    """Cells with date serial numbers and a styles.xml with date formats."""
    styles = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>
<fonts count="1"><font><sz val="11"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border/></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="3">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
<xf numFmtId="14" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
</cellXfs>
</styleSheet>"""
    rows = """
<row r="1">
  <c r="A1" s="1"><v>44197</v></c>
  <c r="B1" s="2"><v>44197</v></c>
  <c r="C1" s="1"><v>1</v></c>
  <c r="D1"><v>44197</v></c>
</row>"""
    extra = {"xl/styles.xml": styles}
    content_type_override = '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    return _minimal_xlsx({"sheet1.xml": _sheet_xml(rows)}, extra_files=extra)


def gen_column_widths():
    """Sheet with custom column widths."""
    extras = '<cols><col min="1" max="1" width="20" customWidth="1"/><col min="3" max="5" width="8.5" customWidth="1"/></cols>'
    rows = '<row r="1"><c r="A1"><v>1</v></c></row>'
    sheet = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
{extras}
<sheetData>{rows}</sheetData>
</worksheet>"""
    return _minimal_xlsx({"sheet1.xml": sheet})


def gen_inline_strings():
    """Cells using inline string type (t="inlineStr") with <is><t>...</t></is>."""
    rows = """
<row r="1">
  <c r="A1" t="inlineStr"><is><t>Inline A</t></is></c>
  <c r="B1" t="inlineStr"><is><t>Inline B</t></is></c>
  <c r="C1" t="inlineStr"><is><r><t>Rich </t></r><r><rPr><b/></rPr><t>inline</t></r></is></c>
</row>"""
    return _minimal_xlsx({"sheet1.xml": _sheet_xml(rows)})


def gen_document_properties():
    """Workbook with core and custom document properties."""
    core = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/">
<dc:title>Test Workbook</dc:title>
<dc:creator>Test Author</dc:creator>
<dc:description>A test description</dc:description>
<dcterms:created>2024-01-15T10:30:00Z</dcterms:created>
</cp:coreProperties>"""
    custom = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Department">
<vt:lpwstr>Engineering</vt:lpwstr>
</property>
</Properties>"""
    extras = {"docProps/core.xml": core, "docProps/custom.xml": custom}
    rows = '<row r="1"><c r="A1"><v>1</v></c></row>'
    return _minimal_xlsx({"sheet1.xml": _sheet_xml(rows)}, extra_files=extras)


# ---- Malformed / edge-case files for robustness testing ----

def gen_truncated_zip():
    """A valid ZIP header followed by truncated data."""
    valid = gen_single_cell()
    return valid[:len(valid) // 2]  # Cut in half


def gen_empty_zip():
    """A valid but empty ZIP file (no entries)."""
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w"):
        pass
    return buf.getvalue()


def gen_missing_workbook():
    """ZIP with worksheet but no workbook.xml."""
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
        zf.writestr("xl/worksheets/sheet1.xml", _sheet_xml('<row r="1"><c r="A1"><v>1</v></c></row>'))
    return buf.getvalue()


def gen_malformed_xml():
    """XLSX with syntactically invalid XML in the worksheet."""
    rows = '<row r="1"><c r="A1"><v>1</v></c></row><UNCLOSED'
    return _minimal_xlsx({"sheet1.xml": _sheet_xml(rows) + "<<<GARBAGE>>>"})


def gen_wrong_string_index():
    """Shared string reference pointing beyond table size."""
    ss = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
<si><t>only_one</t></si>
</sst>"""
    rows = '<row r="1"><c r="A1" t="s"><v>999</v></c></row>'
    return _minimal_xlsx({"sheet1.xml": _sheet_xml(rows)}, shared_strings=ss)


def gen_negative_row_number():
    """Row with negative or zero row number."""
    rows = '<row r="0"><c r="A0"><v>0</v></c></row>'
    return _minimal_xlsx({"sheet1.xml": _sheet_xml(rows)})


def gen_huge_row_gap():
    """Rows at 1 and 1048576 (Excel max), testing sparse allocation."""
    rows = '<row r="1"><c r="A1"><v>1</v></c></row><row r="1048576"><c r="A1048576"><v>2</v></c></row>'
    return _minimal_xlsx({"sheet1.xml": _sheet_xml(rows)})


GENERATORS = {
    # Valid files
    "empty_workbook.xlsx": gen_empty_workbook,
    "single_cell.xlsx": gen_single_cell,
    "many_data_types.xlsx": gen_many_data_types,
    "multiple_sheets.xlsx": gen_multiple_sheets,
    "merged_cells.xlsx": gen_merged_cells,
    "sparse_rows.xlsx": gen_sparse_rows,
    "unicode_strings.xlsx": gen_unicode_strings,
    "rich_text_strings.xlsx": gen_rich_text_strings,
    "freeze_panes.xlsx": gen_freeze_panes,
    "auto_filter.xlsx": gen_auto_filter,
    "defined_names.xlsx": gen_defined_names,
    "large_shared_strings.xlsx": gen_large_shared_strings,
    "date_cells.xlsx": gen_date_cells,
    "column_widths.xlsx": gen_column_widths,
    "inline_strings.xlsx": gen_inline_strings,
    "document_properties.xlsx": gen_document_properties,
    # Malformed / edge cases
    "truncated_zip.xlsx": gen_truncated_zip,
    "empty_zip.xlsx": gen_empty_zip,
    "missing_workbook.xlsx": gen_missing_workbook,
    "malformed_xml.xlsx": gen_malformed_xml,
    "wrong_string_index.xlsx": gen_wrong_string_index,
    "negative_row_number.xlsx": gen_negative_row_number,
    "huge_row_gap.xlsx": gen_huge_row_gap,
}


def main():
    FIXTURES.mkdir(parents=True, exist_ok=True)
    for name, gen in GENERATORS.items():
        path = FIXTURES / name
        data = gen()
        path.write_bytes(data)
        print(f"  {name} ({len(data):,} bytes)")
    print(f"\nGenerated {len(GENERATORS)} files in {FIXTURES}")


if __name__ == "__main__":
    main()
