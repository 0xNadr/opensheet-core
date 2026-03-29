use std::io::{Seek, Write};

use zip::write::SimpleFileOptions;
use zip::ZipWriter;

use crate::types::CellValue;

/// Errors that can occur during XLSX writing.
#[derive(Debug)]
pub enum XlsxWriteError {
    Zip(zip::result::ZipError),
    Io(std::io::Error),
    InvalidState(String),
}

impl std::fmt::Display for XlsxWriteError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            XlsxWriteError::Zip(e) => write!(f, "ZIP error: {e}"),
            XlsxWriteError::Io(e) => write!(f, "IO error: {e}"),
            XlsxWriteError::InvalidState(msg) => write!(f, "Invalid state: {msg}"),
        }
    }
}

impl From<zip::result::ZipError> for XlsxWriteError {
    fn from(e: zip::result::ZipError) -> Self {
        XlsxWriteError::Zip(e)
    }
}

impl From<std::io::Error> for XlsxWriteError {
    fn from(e: std::io::Error) -> Self {
        XlsxWriteError::Io(e)
    }
}

impl From<XlsxWriteError> for pyo3::PyErr {
    fn from(e: XlsxWriteError) -> Self {
        pyo3::exceptions::PyIOError::new_err(e.to_string())
    }
}

/// Tracks info about each sheet for writing workbook metadata at the end.
struct SheetEntry {
    name: String,
    index: usize,
}

/// A streaming XLSX writer that writes rows directly to a ZIP archive.
///
/// Uses inline strings instead of a shared string table for true streaming
/// without needing to buffer all string values.
pub struct StreamingXlsxWriter<W: Write + Seek> {
    zip: Option<ZipWriter<W>>,
    sheets: Vec<SheetEntry>,
    current_row: u32,
    sheet_open: bool,
}

impl<W: Write + Seek> StreamingXlsxWriter<W> {
    /// Create a new streaming XLSX writer.
    pub fn new(writer: W) -> Self {
        StreamingXlsxWriter {
            zip: Some(ZipWriter::new(writer)),
            sheets: Vec::new(),
            current_row: 0,
            sheet_open: false,
        }
    }

    /// Get a mutable reference to the inner ZipWriter, or error if closed.
    fn zip(&mut self) -> Result<&mut ZipWriter<W>, XlsxWriteError> {
        self.zip
            .as_mut()
            .ok_or_else(|| XlsxWriteError::InvalidState("Writer is already closed".to_string()))
    }

    /// Add a new sheet. If a sheet is currently open, it will be closed first.
    pub fn add_sheet(&mut self, name: &str) -> Result<(), XlsxWriteError> {
        self.zip()?; // Check not closed

        // Close the previous sheet if one is open
        if self.sheet_open {
            self.close_sheet()?;
        }

        let index = self.sheets.len() + 1;
        self.sheets.push(SheetEntry {
            name: name.to_string(),
            index,
        });

        // Start the worksheet XML file in the ZIP
        let path = format!("xl/worksheets/sheet{index}.xml");
        let options =
            SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
        self.zip()?.start_file(path, options)?;

        // Write worksheet XML header
        write!(
            self.zip()?,
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
             <worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" \
             xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\
             <sheetData>"
        )?;

        self.current_row = 0;
        self.sheet_open = true;
        Ok(())
    }

    /// Write a row of cell values to the current sheet.
    pub fn write_row(&mut self, cells: &[CellValue]) -> Result<(), XlsxWriteError> {
        if !self.sheet_open {
            return Err(XlsxWriteError::InvalidState(
                "No sheet is open. Call add_sheet() first.".to_string(),
            ));
        }

        self.current_row += 1;
        let row_num = self.current_row;

        write!(self.zip()?, "<row r=\"{row_num}\">")?;

        for (col_idx, cell) in cells.iter().enumerate() {
            let col_letter = col_index_to_letter(col_idx);
            let cell_ref = format!("{col_letter}{row_num}");

            match cell {
                CellValue::String(s) => {
                    let escaped = xml_escape(s);
                    write!(
                        self.zip()?,
                        "<c r=\"{cell_ref}\" t=\"inlineStr\"><is><t>{escaped}</t></is></c>"
                    )?;
                }
                CellValue::Number(n) => {
                    write!(self.zip()?, "<c r=\"{cell_ref}\"><v>{n}</v></c>")?;
                }
                CellValue::Bool(b) => {
                    let val = if *b { "1" } else { "0" };
                    write!(self.zip()?, "<c r=\"{cell_ref}\" t=\"b\"><v>{val}</v></c>")?;
                }
                CellValue::Empty => {}
            }
        }

        write!(self.zip()?, "</row>")?;
        Ok(())
    }

    /// Close the current sheet's XML.
    fn close_sheet(&mut self) -> Result<(), XlsxWriteError> {
        if self.sheet_open {
            write!(self.zip()?, "</sheetData></worksheet>")?;
            self.sheet_open = false;
        }
        Ok(())
    }

    /// Finalize the XLSX file: close any open sheet, write workbook metadata,
    /// content types, and relationships.
    pub fn close(mut self) -> Result<(), XlsxWriteError> {
        self.finalize()
    }

    fn finalize(&mut self) -> Result<(), XlsxWriteError> {
        if self.zip.is_none() {
            return Ok(());
        }

        // Close the current sheet if open
        self.close_sheet()?;

        // If no sheets were added, add a default empty one
        if self.sheets.is_empty() {
            self.add_sheet("Sheet1")?;
            self.close_sheet()?;
        }

        let options =
            SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);

        // Write [Content_Types].xml
        self.zip()?.start_file("[Content_Types].xml", options)?;
        write!(
            self.zip()?,
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
             <Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\
             <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\
             <Default Extension=\"xml\" ContentType=\"application/xml\"/>\
             <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\
             <Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>"
        )?;
        for i in 0..self.sheets.len() {
            let index = self.sheets[i].index;
            write!(
                self.zip()?,
                "<Override PartName=\"/xl/worksheets/sheet{index}.xml\" \
                 ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
            )?;
        }
        write!(self.zip()?, "</Types>")?;

        // Write _rels/.rels
        self.zip()?.start_file("_rels/.rels", options)?;
        write!(
            self.zip()?,
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
             <Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
             <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>\
             </Relationships>"
        )?;

        // Write xl/workbook.xml
        self.zip()?.start_file("xl/workbook.xml", options)?;
        write!(
            self.zip()?,
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
             <workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" \
             xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\
             <sheets>"
        )?;
        for i in 0..self.sheets.len() {
            let escaped_name = xml_escape(&self.sheets[i].name);
            let index = self.sheets[i].index;
            write!(
                self.zip()?,
                "<sheet name=\"{escaped_name}\" sheetId=\"{index}\" r:id=\"rId{index}\"/>"
            )?;
        }
        write!(self.zip()?, "</sheets></workbook>")?;

        // Write xl/_rels/workbook.xml.rels
        self.zip()?
            .start_file("xl/_rels/workbook.xml.rels", options)?;
        write!(
            self.zip()?,
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
             <Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        )?;
        for i in 0..self.sheets.len() {
            let index = self.sheets[i].index;
            write!(
                self.zip()?,
                "<Relationship Id=\"rId{index}\" \
                 Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" \
                 Target=\"worksheets/sheet{index}.xml\"/>"
            )?;
        }
        let styles_id = self.sheets.len() + 1;
        write!(
            self.zip()?,
            "<Relationship Id=\"rId{styles_id}\" \
             Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" \
             Target=\"styles.xml\"/>\
             </Relationships>"
        )?;

        // Write xl/styles.xml (minimal required styles)
        self.zip()?.start_file("xl/styles.xml", options)?;
        write!(
            self.zip()?,
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
             <styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\
             <fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>\
             <fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>\
             <borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>\
             <cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>\
             <cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>\
             </styleSheet>"
        )?;

        // Take ownership of the ZipWriter to call finish()
        let zip = self.zip.take().unwrap();
        zip.finish()?;
        Ok(())
    }
}

/// Convert a 0-based column index to an Excel-style column letter (A, B, ..., Z, AA, AB, ...).
fn col_index_to_letter(index: usize) -> String {
    let mut result = String::new();
    let mut n = index;
    loop {
        result.insert(0, (b'A' + (n % 26) as u8) as char);
        if n < 26 {
            break;
        }
        n = n / 26 - 1;
    }
    result
}

/// Escape special XML characters in a string.
fn xml_escape(s: &str) -> String {
    let mut result = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '&' => result.push_str("&amp;"),
            '<' => result.push_str("&lt;"),
            '>' => result.push_str("&gt;"),
            '"' => result.push_str("&quot;"),
            '\'' => result.push_str("&apos;"),
            _ => result.push(c),
        }
    }
    result
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_col_index_to_letter() {
        assert_eq!(col_index_to_letter(0), "A");
        assert_eq!(col_index_to_letter(1), "B");
        assert_eq!(col_index_to_letter(25), "Z");
        assert_eq!(col_index_to_letter(26), "AA");
        assert_eq!(col_index_to_letter(27), "AB");
        assert_eq!(col_index_to_letter(51), "AZ");
        assert_eq!(col_index_to_letter(52), "BA");
        assert_eq!(col_index_to_letter(701), "ZZ");
        assert_eq!(col_index_to_letter(702), "AAA");
    }

    #[test]
    fn test_xml_escape() {
        assert_eq!(xml_escape("hello"), "hello");
        assert_eq!(xml_escape("a & b"), "a &amp; b");
        assert_eq!(xml_escape("<tag>"), "&lt;tag&gt;");
        assert_eq!(xml_escape("it's \"fine\""), "it&apos;s &quot;fine&quot;");
    }

    #[test]
    fn test_write_and_read_roundtrip() {
        use crate::reader::xlsx::read_xlsx;
        use std::io::Cursor;

        let mut buf = Cursor::new(Vec::new());

        {
            let mut writer = StreamingXlsxWriter::new(&mut buf);
            writer.add_sheet("TestSheet").unwrap();
            writer
                .write_row(&[
                    CellValue::String("Name".to_string()),
                    CellValue::String("Value".to_string()),
                ])
                .unwrap();
            writer
                .write_row(&[
                    CellValue::String("Alice".to_string()),
                    CellValue::Number(42.0),
                ])
                .unwrap();
            writer
                .write_row(&[CellValue::Bool(true), CellValue::Empty])
                .unwrap();
            writer.close().unwrap();
        }

        buf.set_position(0);
        let sheets = read_xlsx(buf).unwrap();

        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].name, "TestSheet");
        assert_eq!(sheets[0].rows.len(), 3);

        // Row 1: header
        match &sheets[0].rows[0][0] {
            CellValue::String(s) => assert_eq!(s, "Name"),
            other => panic!("expected string, got {other:?}"),
        }

        // Row 2: mixed
        match &sheets[0].rows[1][1] {
            CellValue::Number(n) => assert!((n - 42.0).abs() < f64::EPSILON),
            other => panic!("expected number, got {other:?}"),
        }

        // Row 3: bool
        match &sheets[0].rows[2][0] {
            CellValue::Bool(b) => assert!(*b),
            other => panic!("expected bool, got {other:?}"),
        }
    }
}
