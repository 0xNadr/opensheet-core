use std::io::{Seek, Write};

use zip::write::SimpleFileOptions;
use zip::ZipWriter;

use std::collections::HashMap;

use crate::types::{datetime_to_excel_serial, CellValue};

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

/// Pending merge ranges for the current sheet.
struct PendingMerges {
    ranges: Vec<String>,
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
    sheet_data_started: bool,
    has_dates: bool,
    has_datetimes: bool,
    pending_merges: PendingMerges,
    pending_columns: HashMap<u32, f64>,
    pending_row_heights: HashMap<u32, f64>,
}

impl<W: Write + Seek> StreamingXlsxWriter<W> {
    /// Create a new streaming XLSX writer.
    pub fn new(writer: W) -> Self {
        StreamingXlsxWriter {
            zip: Some(ZipWriter::new(writer)),
            sheets: Vec::new(),
            current_row: 0,
            sheet_open: false,
            sheet_data_started: false,
            has_dates: false,
            has_datetimes: false,
            pending_merges: PendingMerges { ranges: Vec::new() },
            pending_columns: HashMap::new(),
            pending_row_heights: HashMap::new(),
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

        // Write worksheet XML header (but NOT <sheetData> yet — deferred until first row
        // so that <cols> can be written before it if column widths are set)
        write!(
            self.zip()?,
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
             <worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" \
             xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        )?;

        self.current_row = 0;
        self.sheet_open = true;
        self.sheet_data_started = false;
        Ok(())
    }

    /// Flush any pending column widths and write the <sheetData> opening tag.
    /// Called lazily before the first row is written.
    fn start_sheet_data(&mut self) -> Result<(), XlsxWriteError> {
        if self.sheet_data_started {
            return Ok(());
        }

        // Write <cols> if any column widths are set
        let columns = std::mem::take(&mut self.pending_columns);
        if !columns.is_empty() {
            write!(self.zip()?, "<cols>")?;
            let mut sorted_cols: Vec<_> = columns.into_iter().collect();
            sorted_cols.sort_by_key(|(idx, _)| *idx);
            for (col_idx, width) in sorted_cols {
                let col_num = col_idx + 1; // XLSX uses 1-based column numbers
                write!(
                    self.zip()?,
                    "<col min=\"{col_num}\" max=\"{col_num}\" width=\"{width}\" customWidth=\"1\"/>"
                )?;
            }
            write!(self.zip()?, "</cols>")?;
        }

        write!(self.zip()?, "<sheetData>")?;
        self.sheet_data_started = true;
        Ok(())
    }

    /// Write a row of cell values to the current sheet.
    pub fn write_row(&mut self, cells: &[CellValue]) -> Result<(), XlsxWriteError> {
        if !self.sheet_open {
            return Err(XlsxWriteError::InvalidState(
                "No sheet is open. Call add_sheet() first.".to_string(),
            ));
        }

        self.start_sheet_data()?;

        self.current_row += 1;
        let row_num = self.current_row;

        // Check for custom row height (0-based index)
        let row_idx = row_num - 1;
        let custom_height = self.pending_row_heights.get(&row_idx).copied();
        if let Some(height) = custom_height {
            write!(
                self.zip()?,
                "<row r=\"{row_num}\" ht=\"{height}\" customHeight=\"1\">"
            )?;
        } else {
            write!(self.zip()?, "<row r=\"{row_num}\">")?;
        }

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
                CellValue::Formula {
                    formula,
                    cached_value,
                } => {
                    let escaped_formula = xml_escape(formula);
                    match cached_value.as_deref() {
                        Some(CellValue::Number(n)) => {
                            write!(
                                self.zip()?,
                                "<c r=\"{cell_ref}\"><f>{escaped_formula}</f><v>{n}</v></c>"
                            )?;
                        }
                        Some(CellValue::String(s)) => {
                            let escaped_val = xml_escape(s);
                            write!(
                                self.zip()?,
                                "<c r=\"{cell_ref}\" t=\"str\"><f>{escaped_formula}</f><v>{escaped_val}</v></c>"
                            )?;
                        }
                        _ => {
                            write!(
                                self.zip()?,
                                "<c r=\"{cell_ref}\"><f>{escaped_formula}</f></c>"
                            )?;
                        }
                    }
                }
                CellValue::Date { year, month, day } => {
                    let serial = datetime_to_excel_serial(*year, *month, *day, 0, 0, 0, 0);
                    // Style index 1 = date format
                    write!(
                        self.zip()?,
                        "<c r=\"{cell_ref}\" s=\"1\"><v>{serial}</v></c>"
                    )?;
                    self.has_dates = true;
                }
                CellValue::DateTime {
                    year,
                    month,
                    day,
                    hour,
                    minute,
                    second,
                    microsecond,
                } => {
                    let serial = datetime_to_excel_serial(
                        *year,
                        *month,
                        *day,
                        *hour,
                        *minute,
                        *second,
                        *microsecond,
                    );
                    // Style index 2 = datetime format
                    write!(
                        self.zip()?,
                        "<c r=\"{cell_ref}\" s=\"2\"><v>{serial}</v></c>"
                    )?;
                    self.has_datetimes = true;
                }
                CellValue::Empty => {}
            }
        }

        write!(self.zip()?, "</row>")?;
        Ok(())
    }

    /// Mark a range of cells as merged (e.g. "A1:B2").
    pub fn merge_cells(&mut self, range: &str) -> Result<(), XlsxWriteError> {
        if !self.sheet_open {
            return Err(XlsxWriteError::InvalidState(
                "No sheet is open. Call add_sheet() first.".to_string(),
            ));
        }
        self.pending_merges.ranges.push(range.to_string());
        Ok(())
    }

    /// Set the width of a column (0-based index) in character units.
    /// Must be called before any rows are written (i.e., after add_sheet but before write_row).
    pub fn set_column_width(&mut self, col_index: u32, width: f64) -> Result<(), XlsxWriteError> {
        if !self.sheet_open {
            return Err(XlsxWriteError::InvalidState(
                "No sheet is open. Call add_sheet() first.".to_string(),
            ));
        }
        if self.sheet_data_started {
            return Err(XlsxWriteError::InvalidState(
                "Column widths must be set before writing any rows.".to_string(),
            ));
        }
        self.pending_columns.insert(col_index, width);
        Ok(())
    }

    /// Set the height of a row (0-based index) in points.
    pub fn set_row_height(&mut self, row_index: u32, height: f64) -> Result<(), XlsxWriteError> {
        if !self.sheet_open {
            return Err(XlsxWriteError::InvalidState(
                "No sheet is open. Call add_sheet() first.".to_string(),
            ));
        }
        self.pending_row_heights.insert(row_index, height);
        Ok(())
    }

    /// Close the current sheet's XML.
    fn close_sheet(&mut self) -> Result<(), XlsxWriteError> {
        if self.sheet_open {
            // Ensure <sheetData> was opened (even for sheets with no rows)
            self.start_sheet_data()?;
            write!(self.zip()?, "</sheetData>")?;

            // Clear per-sheet state
            self.pending_row_heights.clear();

            // Write mergeCells if any
            let merges = std::mem::take(&mut self.pending_merges.ranges);
            if !merges.is_empty() {
                let count = merges.len();
                write!(self.zip()?, "<mergeCells count=\"{count}\">")?;
                for range in &merges {
                    let escaped = xml_escape(range);
                    write!(self.zip()?, "<mergeCell ref=\"{escaped}\"/>")?;
                }
                write!(self.zip()?, "</mergeCells>")?;
            }

            write!(self.zip()?, "</worksheet>")?;
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

        // Write xl/styles.xml with date/datetime formats
        self.zip()?.start_file("xl/styles.xml", options)?;
        write!(
            self.zip()?,
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
             <styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\
             <numFmts count=\"2\">\
             <numFmt numFmtId=\"164\" formatCode=\"yyyy\\-mm\\-dd\"/>\
             <numFmt numFmtId=\"165\" formatCode=\"yyyy\\-mm\\-dd\\ hh:mm:ss\"/>\
             </numFmts>\
             <fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>\
             <fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>\
             <borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>\
             <cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>\
             <cellXfs count=\"3\">\
             <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>\
             <xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>\
             <xf numFmtId=\"165\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>\
             </cellXfs>\
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
