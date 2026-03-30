use std::io::{Seek, Write};

use zip::write::SimpleFileOptions;
use zip::ZipWriter;

use std::collections::HashMap;

use crate::types::{datetime_to_excel_serial, CellStyle, CellValue};

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

// ---------- Internal style registry types ----------

#[derive(Debug, Clone, PartialEq)]
struct FontDef {
    bold: bool,
    italic: bool,
    underline: bool,
    name: String,
    size: f64,
    color: Option<String>,
}

impl Default for FontDef {
    fn default() -> Self {
        FontDef {
            bold: false,
            italic: false,
            underline: false,
            name: "Calibri".to_string(),
            size: 11.0,
            color: None,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
struct FillDef {
    pattern_type: String,
    fg_color: Option<String>,
}

#[derive(Debug, Clone, PartialEq)]
struct BorderSideDef {
    style: String,
    color: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Default)]
struct BorderDef {
    left: Option<BorderSideDef>,
    right: Option<BorderSideDef>,
    top: Option<BorderSideDef>,
    bottom: Option<BorderSideDef>,
}

#[derive(Debug, Clone, PartialEq)]
struct AlignmentDef {
    horizontal: Option<String>,
    vertical: Option<String>,
    wrap_text: bool,
    text_rotation: Option<u16>,
}

#[derive(Debug, Clone, PartialEq)]
struct XfDef {
    font_id: usize,
    fill_id: usize,
    border_id: usize,
    num_fmt_id: u32,
    alignment: Option<AlignmentDef>,
}

/// Normalize a hex color string to AARRGGBB format.
fn normalize_color(color: &str) -> String {
    let color = color.trim_start_matches('#');
    if color.len() == 6 {
        format!("FF{}", color.to_uppercase())
    } else {
        color.to_uppercase()
    }
}

// ---------- Writer types ----------

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
    pending_freeze_pane: Option<(u32, u32)>,
    pending_auto_filter: Option<String>,
    // Style registries
    fonts: Vec<FontDef>,
    fills: Vec<FillDef>,
    borders: Vec<BorderDef>,
    xfs: Vec<XfDef>,
    num_fmts: Vec<(u32, String)>,
    num_fmt_map: HashMap<String, u32>,
    next_num_fmt_id: u32,
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
            pending_freeze_pane: None,
            pending_auto_filter: None,
            fonts: vec![FontDef::default()],
            fills: vec![
                FillDef {
                    pattern_type: "none".to_string(),
                    fg_color: None,
                },
                FillDef {
                    pattern_type: "gray125".to_string(),
                    fg_color: None,
                },
            ],
            borders: vec![BorderDef::default()],
            xfs: vec![
                // XF 0: general
                XfDef {
                    font_id: 0,
                    fill_id: 0,
                    border_id: 0,
                    num_fmt_id: 0,
                    alignment: None,
                },
                // XF 1: date format (numFmt 164)
                XfDef {
                    font_id: 0,
                    fill_id: 0,
                    border_id: 0,
                    num_fmt_id: 164,
                    alignment: None,
                },
                // XF 2: datetime format (numFmt 165)
                XfDef {
                    font_id: 0,
                    fill_id: 0,
                    border_id: 0,
                    num_fmt_id: 165,
                    alignment: None,
                },
            ],
            num_fmts: vec![],
            num_fmt_map: HashMap::new(),
            next_num_fmt_id: 166,
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
        self.pending_freeze_pane = None;
        Ok(())
    }

    /// Flush any pending column widths and write the <sheetData> opening tag.
    /// Called lazily before the first row is written.
    fn start_sheet_data(&mut self) -> Result<(), XlsxWriteError> {
        if self.sheet_data_started {
            return Ok(());
        }

        // Write <sheetViews> with freeze pane if set
        if let Some((row_split, col_split)) = self.pending_freeze_pane.take() {
            if row_split > 0 || col_split > 0 {
                let top_left_cell = format!(
                    "{}{}",
                    col_index_to_letter(col_split as usize),
                    row_split + 1
                );
                let active_pane = match (row_split > 0, col_split > 0) {
                    (true, true) => "bottomRight",
                    (true, false) => "bottomLeft",
                    (false, true) => "topRight",
                    (false, false) => unreachable!(),
                };
                write!(
                    self.zip()?,
                    "<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\">\
                     <pane{}{}topLeftCell=\"{top_left_cell}\" activePane=\"{active_pane}\" state=\"frozen\"/>\
                     <selection pane=\"{active_pane}\"/>\
                     </sheetView></sheetViews>",
                    if row_split > 0 {
                        format!(" ySplit=\"{row_split}\"")
                    } else {
                        String::new()
                    },
                    if col_split > 0 {
                        format!(" xSplit=\"{col_split}\"")
                    } else {
                        String::new()
                    },
                )?;
            }
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

            // Unwrap StyledCell to get inner value and style override
            let (inner, style_id) = match cell {
                CellValue::StyledCell { value, style } => {
                    // For Date/DateTime inner values, ensure the style includes
                    // the appropriate date/datetime number format
                    let mut style = style.clone();
                    match value.as_ref() {
                        CellValue::Date { .. } if style.number_format.is_none() => {
                            style.number_format = Some("yyyy\\-mm\\-dd".to_string());
                        }
                        CellValue::DateTime { .. } if style.number_format.is_none() => {
                            style.number_format = Some("yyyy\\-mm\\-dd\\ hh:mm:ss".to_string());
                        }
                        _ => {}
                    }
                    let id = self.register_cell_style(&style);
                    (value.as_ref(), Some(id))
                }
                other => (other, None),
            };

            self.write_single_cell(&cell_ref, inner, style_id)?;
        }

        write!(self.zip()?, "</row>")?;
        Ok(())
    }

    /// Write a single cell value with an optional style override.
    fn write_single_cell(
        &mut self,
        cell_ref: &str,
        cell: &CellValue,
        style_id: Option<u32>,
    ) -> Result<(), XlsxWriteError> {
        let s_attr = style_id
            .map(|id| format!(" s=\"{id}\""))
            .unwrap_or_default();

        match cell {
            CellValue::String(s) => {
                let escaped = xml_escape(s);
                write!(
                    self.zip()?,
                    "<c r=\"{cell_ref}\" t=\"inlineStr\"{s_attr}><is><t>{escaped}</t></is></c>"
                )?;
            }
            CellValue::Number(n) => {
                write!(self.zip()?, "<c r=\"{cell_ref}\"{s_attr}><v>{n}</v></c>")?;
            }
            CellValue::Bool(b) => {
                let val = if *b { "1" } else { "0" };
                write!(
                    self.zip()?,
                    "<c r=\"{cell_ref}\" t=\"b\"{s_attr}><v>{val}</v></c>"
                )?;
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
                            "<c r=\"{cell_ref}\"{s_attr}><f>{escaped_formula}</f><v>{n}</v></c>"
                        )?;
                    }
                    Some(CellValue::String(s)) => {
                        let escaped_val = xml_escape(s);
                        write!(
                            self.zip()?,
                            "<c r=\"{cell_ref}\" t=\"str\"{s_attr}><f>{escaped_formula}</f><v>{escaped_val}</v></c>"
                        )?;
                    }
                    _ => {
                        write!(
                            self.zip()?,
                            "<c r=\"{cell_ref}\"{s_attr}><f>{escaped_formula}</f></c>"
                        )?;
                    }
                }
            }
            CellValue::Date { year, month, day } => {
                let serial = datetime_to_excel_serial(*year, *month, *day, 0, 0, 0, 0);
                let sid = style_id.unwrap_or(1); // Default: XF 1 = date format
                write!(
                    self.zip()?,
                    "<c r=\"{cell_ref}\" s=\"{sid}\"><v>{serial}</v></c>"
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
                let sid = style_id.unwrap_or(2); // Default: XF 2 = datetime format
                write!(
                    self.zip()?,
                    "<c r=\"{cell_ref}\" s=\"{sid}\"><v>{serial}</v></c>"
                )?;
                self.has_datetimes = true;
            }
            CellValue::FormattedNumber { value, format_code } => {
                let xf_id = if let Some(id) = style_id {
                    id
                } else {
                    self.register_format(format_code)
                };
                write!(
                    self.zip()?,
                    "<c r=\"{cell_ref}\" s=\"{xf_id}\"><v>{value}</v></c>"
                )?;
            }
            CellValue::StyledCell { .. } => {
                // Should not happen — StyledCell is unwrapped before calling this
            }
            CellValue::Empty => {}
        }
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

    /// Set freeze panes: freeze the top `row` rows and left `col` columns.
    /// Must be called after add_sheet() but before write_row().
    pub fn freeze_panes(&mut self, row: u32, col: u32) -> Result<(), XlsxWriteError> {
        if !self.sheet_open {
            return Err(XlsxWriteError::InvalidState(
                "No sheet is open. Call add_sheet() first.".to_string(),
            ));
        }
        if self.sheet_data_started {
            return Err(XlsxWriteError::InvalidState(
                "Freeze panes must be set before writing any rows.".to_string(),
            ));
        }
        self.pending_freeze_pane = Some((row, col));
        Ok(())
    }

    /// Set an auto-filter on a range (e.g. "A1:C1").
    pub fn auto_filter(&mut self, range: &str) -> Result<(), XlsxWriteError> {
        if !self.sheet_open {
            return Err(XlsxWriteError::InvalidState(
                "No sheet is open. Call add_sheet() first.".to_string(),
            ));
        }
        self.pending_auto_filter = Some(range.to_string());
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

    // ---------- Style registry methods ----------

    /// Register a custom number format and return its numFmtId.
    fn register_num_fmt(&mut self, format_code: &str) -> u32 {
        if let Some(&id) = self.num_fmt_map.get(format_code) {
            return id;
        }
        let id = self.next_num_fmt_id;
        self.num_fmts.push((id, format_code.to_string()));
        self.num_fmt_map.insert(format_code.to_string(), id);
        self.next_num_fmt_id += 1;
        id
    }

    /// Register an XF entry and return its index. Deduplicates.
    fn register_xf(&mut self, xf: &XfDef) -> u32 {
        if let Some(idx) = self.xfs.iter().position(|x| x == xf) {
            return idx as u32;
        }
        let idx = self.xfs.len() as u32;
        self.xfs.push(xf.clone());
        idx
    }

    /// Register a number format and return its xf index (for FormattedNumber backward compat).
    fn register_format(&mut self, format_code: &str) -> u32 {
        let num_fmt_id = self.register_num_fmt(format_code);
        let xf = XfDef {
            font_id: 0,
            fill_id: 0,
            border_id: 0,
            num_fmt_id,
            alignment: None,
        };
        self.register_xf(&xf)
    }

    /// Register a font from a CellStyle and return its fontId.
    fn register_font(&mut self, style: &CellStyle) -> usize {
        let font = FontDef {
            bold: style.bold,
            italic: style.italic,
            underline: style.underline,
            name: style
                .font_name
                .clone()
                .unwrap_or_else(|| "Calibri".to_string()),
            size: style.font_size.unwrap_or(11.0),
            color: style.font_color.as_ref().map(|c| normalize_color(c)),
        };
        if let Some(idx) = self.fonts.iter().position(|f| f == &font) {
            return idx;
        }
        self.fonts.push(font);
        self.fonts.len() - 1
    }

    /// Register a fill from a CellStyle and return its fillId.
    fn register_fill(&mut self, style: &CellStyle) -> usize {
        if style.fill_color.is_none() {
            return 0; // Default "none" fill
        }
        let fill = FillDef {
            pattern_type: "solid".to_string(),
            fg_color: style.fill_color.as_ref().map(|c| normalize_color(c)),
        };
        if let Some(idx) = self.fills.iter().position(|f| f == &fill) {
            return idx;
        }
        self.fills.push(fill);
        self.fills.len() - 1
    }

    /// Register a border from a CellStyle and return its borderId.
    fn register_border(&mut self, style: &CellStyle) -> usize {
        if style.border_left.is_none()
            && style.border_right.is_none()
            && style.border_top.is_none()
            && style.border_bottom.is_none()
        {
            return 0; // Default empty border
        }
        let color = style.border_color.as_ref().map(|c| normalize_color(c));
        let make_side = |s: &Option<String>| -> Option<BorderSideDef> {
            s.as_ref().map(|st| BorderSideDef {
                style: st.clone(),
                color: color.clone(),
            })
        };
        let border = BorderDef {
            left: make_side(&style.border_left),
            right: make_side(&style.border_right),
            top: make_side(&style.border_top),
            bottom: make_side(&style.border_bottom),
        };
        if let Some(idx) = self.borders.iter().position(|b| b == &border) {
            return idx;
        }
        self.borders.push(border);
        self.borders.len() - 1
    }

    /// Register a full cell style and return its xf index.
    fn register_cell_style(&mut self, style: &CellStyle) -> u32 {
        let font_id = self.register_font(style);
        let fill_id = self.register_fill(style);
        let border_id = self.register_border(style);
        let num_fmt_id = if let Some(ref fmt) = style.number_format {
            self.register_num_fmt(fmt)
        } else {
            0
        };
        let alignment = if style.horizontal_alignment.is_some()
            || style.vertical_alignment.is_some()
            || style.wrap_text
            || style.text_rotation.is_some()
        {
            Some(AlignmentDef {
                horizontal: style.horizontal_alignment.clone(),
                vertical: style.vertical_alignment.clone(),
                wrap_text: style.wrap_text,
                text_rotation: style.text_rotation,
            })
        } else {
            None
        };
        let xf = XfDef {
            font_id,
            fill_id,
            border_id,
            num_fmt_id,
            alignment,
        };
        self.register_xf(&xf)
    }

    // ---------- Sheet/file lifecycle ----------

    /// Close the current sheet's XML.
    fn close_sheet(&mut self) -> Result<(), XlsxWriteError> {
        if self.sheet_open {
            // Ensure <sheetData> was opened (even for sheets with no rows)
            self.start_sheet_data()?;
            write!(self.zip()?, "</sheetData>")?;

            // Clear per-sheet state
            self.pending_row_heights.clear();

            // Write autoFilter if set
            if let Some(ref range) = self.pending_auto_filter.take() {
                let escaped = xml_escape(range);
                write!(self.zip()?, "<autoFilter ref=\"{escaped}\"/>")?;
            }

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

        // Write xl/styles.xml
        self.zip()?.start_file("xl/styles.xml", options)?;
        self.write_styles_xml()?;

        // Take ownership of the ZipWriter to call finish()
        let zip = self.zip.take().unwrap();
        zip.finish()?;
        Ok(())
    }

    /// Write the xl/styles.xml file with all registered styles.
    fn write_styles_xml(&mut self) -> Result<(), XlsxWriteError> {
        // Take ownership of style data
        let fonts = std::mem::take(&mut self.fonts);
        let fills = std::mem::take(&mut self.fills);
        let borders = std::mem::take(&mut self.borders);
        let xfs = std::mem::take(&mut self.xfs);
        let num_fmts = std::mem::take(&mut self.num_fmts);

        // numFmts: always include date (164) and datetime (165) plus custom ones
        let total_num_fmts = 2 + num_fmts.len();

        write!(
            self.zip()?,
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
             <styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\
             <numFmts count=\"{total_num_fmts}\">\
             <numFmt numFmtId=\"164\" formatCode=\"yyyy\\-mm\\-dd\"/>\
             <numFmt numFmtId=\"165\" formatCode=\"yyyy\\-mm\\-dd\\ hh:mm:ss\"/>"
        )?;
        for (num_fmt_id, format_code) in &num_fmts {
            let escaped = xml_escape(format_code);
            write!(
                self.zip()?,
                "<numFmt numFmtId=\"{num_fmt_id}\" formatCode=\"{escaped}\"/>"
            )?;
        }
        write!(self.zip()?, "</numFmts>")?;

        // Fonts
        write!(self.zip()?, "<fonts count=\"{}\">", fonts.len())?;
        for font in &fonts {
            write!(self.zip()?, "<font>")?;
            if font.bold {
                write!(self.zip()?, "<b/>")?;
            }
            if font.italic {
                write!(self.zip()?, "<i/>")?;
            }
            if font.underline {
                write!(self.zip()?, "<u/>")?;
            }
            write!(self.zip()?, "<sz val=\"{}\"/>", font.size)?;
            if let Some(ref color) = font.color {
                write!(self.zip()?, "<color rgb=\"{color}\"/>")?;
            }
            let escaped_name = xml_escape(&font.name);
            write!(self.zip()?, "<name val=\"{escaped_name}\"/>")?;
            write!(self.zip()?, "</font>")?;
        }
        write!(self.zip()?, "</fonts>")?;

        // Fills
        write!(self.zip()?, "<fills count=\"{}\">", fills.len())?;
        for fill in &fills {
            if let Some(ref color) = fill.fg_color {
                write!(
                    self.zip()?,
                    "<fill><patternFill patternType=\"{}\"><fgColor rgb=\"{color}\"/></patternFill></fill>",
                    fill.pattern_type
                )?;
            } else {
                write!(
                    self.zip()?,
                    "<fill><patternFill patternType=\"{}\"/></fill>",
                    fill.pattern_type
                )?;
            }
        }
        write!(self.zip()?, "</fills>")?;

        // Borders
        write!(self.zip()?, "<borders count=\"{}\">", borders.len())?;
        for border in &borders {
            write!(self.zip()?, "<border>")?;
            write_border_side(self.zip()?, "left", &border.left)?;
            write_border_side(self.zip()?, "right", &border.right)?;
            write_border_side(self.zip()?, "top", &border.top)?;
            write_border_side(self.zip()?, "bottom", &border.bottom)?;
            write!(self.zip()?, "<diagonal/></border>")?;
        }
        write!(self.zip()?, "</borders>")?;

        // cellStyleXfs
        write!(
            self.zip()?,
            "<cellStyleXfs count=\"1\">\
             <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>\
             </cellStyleXfs>"
        )?;

        // cellXfs
        write!(self.zip()?, "<cellXfs count=\"{}\">", xfs.len())?;
        for xf in &xfs {
            write!(
                self.zip()?,
                "<xf numFmtId=\"{}\" fontId=\"{}\" fillId=\"{}\" borderId=\"{}\" xfId=\"0\"",
                xf.num_fmt_id,
                xf.font_id,
                xf.fill_id,
                xf.border_id
            )?;
            if xf.num_fmt_id != 0 {
                write!(self.zip()?, " applyNumberFormat=\"1\"")?;
            }
            if xf.font_id != 0 {
                write!(self.zip()?, " applyFont=\"1\"")?;
            }
            if xf.fill_id != 0 {
                write!(self.zip()?, " applyFill=\"1\"")?;
            }
            if xf.border_id != 0 {
                write!(self.zip()?, " applyBorder=\"1\"")?;
            }
            if xf.alignment.is_some() {
                write!(self.zip()?, " applyAlignment=\"1\"")?;
            }
            if let Some(ref align) = xf.alignment {
                write!(self.zip()?, "><alignment")?;
                if let Some(ref h) = align.horizontal {
                    write!(self.zip()?, " horizontal=\"{h}\"")?;
                }
                if let Some(ref v) = align.vertical {
                    write!(self.zip()?, " vertical=\"{v}\"")?;
                }
                if align.wrap_text {
                    write!(self.zip()?, " wrapText=\"1\"")?;
                }
                if let Some(rot) = align.text_rotation {
                    write!(self.zip()?, " textRotation=\"{rot}\"")?;
                }
                write!(self.zip()?, "/></xf>")?;
            } else {
                write!(self.zip()?, "/>")?;
            }
        }
        write!(self.zip()?, "</cellXfs></styleSheet>")?;

        Ok(())
    }
}

/// Write a single border side element.
fn write_border_side<W: Write>(
    w: &mut W,
    name: &str,
    side: &Option<BorderSideDef>,
) -> Result<(), std::io::Error> {
    match side {
        Some(def) => {
            write!(w, "<{name} style=\"{}\">", def.style)?;
            if let Some(ref color) = def.color {
                write!(w, "<color rgb=\"{color}\"/>")?;
            }
            write!(w, "</{name}>")?;
        }
        None => {
            write!(w, "<{name}/>")?;
        }
    }
    Ok(())
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
    fn test_normalize_color() {
        assert_eq!(normalize_color("FF0000"), "FFFF0000");
        assert_eq!(normalize_color("#00FF00"), "FF00FF00");
        assert_eq!(normalize_color("FFFF0000"), "FFFF0000");
        assert_eq!(normalize_color("ff0000"), "FFFF0000");
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

    #[test]
    fn test_freeze_panes_top_row() {
        use crate::reader::xlsx::read_xlsx;
        use std::io::Cursor;

        let mut buf = Cursor::new(Vec::new());
        {
            let mut writer = StreamingXlsxWriter::new(&mut buf);
            writer.add_sheet("Frozen").unwrap();
            writer.freeze_panes(1, 0).unwrap();
            writer
                .write_row(&[CellValue::String("Header".to_string())])
                .unwrap();
            writer
                .write_row(&[CellValue::String("Data".to_string())])
                .unwrap();
            writer.close().unwrap();
        }

        buf.set_position(0);
        let sheets = read_xlsx(buf).unwrap();
        assert_eq!(sheets[0].freeze_pane, Some((1, 0)));
    }

    #[test]
    fn test_freeze_panes_left_column() {
        use crate::reader::xlsx::read_xlsx;
        use std::io::Cursor;

        let mut buf = Cursor::new(Vec::new());
        {
            let mut writer = StreamingXlsxWriter::new(&mut buf);
            writer.add_sheet("Frozen").unwrap();
            writer.freeze_panes(0, 2).unwrap();
            writer
                .write_row(&[CellValue::String("A".to_string())])
                .unwrap();
            writer.close().unwrap();
        }

        buf.set_position(0);
        let sheets = read_xlsx(buf).unwrap();
        assert_eq!(sheets[0].freeze_pane, Some((0, 2)));
    }

    #[test]
    fn test_freeze_panes_both() {
        use crate::reader::xlsx::read_xlsx;
        use std::io::Cursor;

        let mut buf = Cursor::new(Vec::new());
        {
            let mut writer = StreamingXlsxWriter::new(&mut buf);
            writer.add_sheet("Frozen").unwrap();
            writer.freeze_panes(2, 1).unwrap();
            writer
                .write_row(&[CellValue::String("A".to_string())])
                .unwrap();
            writer.close().unwrap();
        }

        buf.set_position(0);
        let sheets = read_xlsx(buf).unwrap();
        assert_eq!(sheets[0].freeze_pane, Some((2, 1)));
    }

    #[test]
    fn test_no_freeze_panes() {
        use crate::reader::xlsx::read_xlsx;
        use std::io::Cursor;

        let mut buf = Cursor::new(Vec::new());
        {
            let mut writer = StreamingXlsxWriter::new(&mut buf);
            writer.add_sheet("Plain").unwrap();
            writer
                .write_row(&[CellValue::String("Data".to_string())])
                .unwrap();
            writer.close().unwrap();
        }

        buf.set_position(0);
        let sheets = read_xlsx(buf).unwrap();
        assert_eq!(sheets[0].freeze_pane, None);
    }

    #[test]
    fn test_auto_filter_roundtrip() {
        use crate::reader::xlsx::read_xlsx;
        use std::io::Cursor;

        let mut buf = Cursor::new(Vec::new());
        {
            let mut writer = StreamingXlsxWriter::new(&mut buf);
            writer.add_sheet("Filtered").unwrap();
            writer
                .write_row(&[
                    CellValue::String("Name".to_string()),
                    CellValue::String("Age".to_string()),
                ])
                .unwrap();
            writer
                .write_row(&[
                    CellValue::String("Alice".to_string()),
                    CellValue::Number(30.0),
                ])
                .unwrap();
            writer.auto_filter("A1:B1").unwrap();
            writer.close().unwrap();
        }

        buf.set_position(0);
        let sheets = read_xlsx(buf).unwrap();
        assert_eq!(sheets[0].auto_filter, Some("A1:B1".to_string()));
    }

    #[test]
    fn test_formatted_number_roundtrip() {
        use crate::reader::xlsx::read_xlsx;
        use std::io::Cursor;

        let mut buf = Cursor::new(Vec::new());
        {
            let mut writer = StreamingXlsxWriter::new(&mut buf);
            writer.add_sheet("Formats").unwrap();
            writer
                .write_row(&[
                    CellValue::String("Price".to_string()),
                    CellValue::String("Percentage".to_string()),
                ])
                .unwrap();
            writer
                .write_row(&[
                    CellValue::FormattedNumber {
                        value: 1234.56,
                        format_code: "$#,##0.00".to_string(),
                    },
                    CellValue::FormattedNumber {
                        value: 0.75,
                        format_code: "0.00%".to_string(),
                    },
                ])
                .unwrap();
            writer.close().unwrap();
        }

        buf.set_position(0);
        let sheets = read_xlsx(buf).unwrap();
        assert_eq!(sheets.len(), 1);

        // Row 2, Col 0: formatted number with currency
        match &sheets[0].rows[1][0] {
            CellValue::FormattedNumber { value, format_code } => {
                assert!((value - 1234.56).abs() < f64::EPSILON);
                assert_eq!(format_code, "$#,##0.00");
            }
            other => panic!("expected FormattedNumber, got {other:?}"),
        }

        // Row 2, Col 1: formatted number with percentage
        match &sheets[0].rows[1][1] {
            CellValue::FormattedNumber { value, format_code } => {
                assert!((value - 0.75).abs() < f64::EPSILON);
                assert_eq!(format_code, "0.00%");
            }
            other => panic!("expected FormattedNumber, got {other:?}"),
        }
    }

    #[test]
    fn test_formatted_number_dedup() {
        use crate::reader::xlsx::read_xlsx;
        use std::io::Cursor;

        let mut buf = Cursor::new(Vec::new());
        {
            let mut writer = StreamingXlsxWriter::new(&mut buf);
            writer.add_sheet("Dedup").unwrap();
            // Same format code used for two cells — should reuse the same xf index
            writer
                .write_row(&[
                    CellValue::FormattedNumber {
                        value: 100.0,
                        format_code: "#,##0".to_string(),
                    },
                    CellValue::FormattedNumber {
                        value: 200.0,
                        format_code: "#,##0".to_string(),
                    },
                ])
                .unwrap();
            writer.close().unwrap();
        }

        buf.set_position(0);
        let sheets = read_xlsx(buf).unwrap();

        match &sheets[0].rows[0][0] {
            CellValue::FormattedNumber { value, format_code } => {
                assert!((value - 100.0).abs() < f64::EPSILON);
                assert_eq!(format_code, "#,##0");
            }
            other => panic!("expected FormattedNumber, got {other:?}"),
        }
        match &sheets[0].rows[0][1] {
            CellValue::FormattedNumber { value, format_code } => {
                assert!((value - 200.0).abs() < f64::EPSILON);
                assert_eq!(format_code, "#,##0");
            }
            other => panic!("expected FormattedNumber, got {other:?}"),
        }
    }

    #[test]
    fn test_no_auto_filter() {
        use crate::reader::xlsx::read_xlsx;
        use std::io::Cursor;

        let mut buf = Cursor::new(Vec::new());
        {
            let mut writer = StreamingXlsxWriter::new(&mut buf);
            writer.add_sheet("Plain").unwrap();
            writer
                .write_row(&[CellValue::String("Data".to_string())])
                .unwrap();
            writer.close().unwrap();
        }

        buf.set_position(0);
        let sheets = read_xlsx(buf).unwrap();
        assert_eq!(sheets[0].auto_filter, None);
    }

    #[test]
    fn test_styled_cell_bold_roundtrip() {
        use crate::reader::xlsx::read_xlsx;
        use std::io::Cursor;

        let mut buf = Cursor::new(Vec::new());
        {
            let mut writer = StreamingXlsxWriter::new(&mut buf);
            writer.add_sheet("Styled").unwrap();
            writer
                .write_row(&[CellValue::StyledCell {
                    value: Box::new(CellValue::String("Bold text".to_string())),
                    style: Box::new(CellStyle {
                        bold: true,
                        ..CellStyle::default()
                    }),
                }])
                .unwrap();
            writer.close().unwrap();
        }

        buf.set_position(0);
        let sheets = read_xlsx(buf).unwrap();
        match &sheets[0].rows[0][0] {
            CellValue::StyledCell { value, style } => {
                match value.as_ref() {
                    CellValue::String(s) => assert_eq!(s, "Bold text"),
                    other => panic!("expected String inside StyledCell, got {other:?}"),
                }
                assert!(style.bold);
            }
            other => panic!("expected StyledCell, got {other:?}"),
        }
    }

    #[test]
    fn test_styled_cell_fill_roundtrip() {
        use crate::reader::xlsx::read_xlsx;
        use std::io::Cursor;

        let mut buf = Cursor::new(Vec::new());
        {
            let mut writer = StreamingXlsxWriter::new(&mut buf);
            writer.add_sheet("Styled").unwrap();
            writer
                .write_row(&[CellValue::StyledCell {
                    value: Box::new(CellValue::Number(42.0)),
                    style: Box::new(CellStyle {
                        fill_color: Some("FFFF00".to_string()),
                        ..CellStyle::default()
                    }),
                }])
                .unwrap();
            writer.close().unwrap();
        }

        buf.set_position(0);
        let sheets = read_xlsx(buf).unwrap();
        match &sheets[0].rows[0][0] {
            CellValue::StyledCell { value, style } => {
                match value.as_ref() {
                    CellValue::Number(n) => assert!((n - 42.0).abs() < f64::EPSILON),
                    other => panic!("expected Number inside StyledCell, got {other:?}"),
                }
                assert_eq!(style.fill_color.as_deref(), Some("FFFFFF00"));
            }
            other => panic!("expected StyledCell, got {other:?}"),
        }
    }

    #[test]
    fn test_styled_cell_border_roundtrip() {
        use crate::reader::xlsx::read_xlsx;
        use std::io::Cursor;

        let mut buf = Cursor::new(Vec::new());
        {
            let mut writer = StreamingXlsxWriter::new(&mut buf);
            writer.add_sheet("Styled").unwrap();
            writer
                .write_row(&[CellValue::StyledCell {
                    value: Box::new(CellValue::String("Bordered".to_string())),
                    style: Box::new(CellStyle {
                        border_left: Some("thin".to_string()),
                        border_right: Some("thin".to_string()),
                        border_top: Some("thin".to_string()),
                        border_bottom: Some("thin".to_string()),
                        border_color: Some("000000".to_string()),
                        ..CellStyle::default()
                    }),
                }])
                .unwrap();
            writer.close().unwrap();
        }

        buf.set_position(0);
        let sheets = read_xlsx(buf).unwrap();
        match &sheets[0].rows[0][0] {
            CellValue::StyledCell { value, style } => {
                match value.as_ref() {
                    CellValue::String(s) => assert_eq!(s, "Bordered"),
                    other => panic!("expected String inside StyledCell, got {other:?}"),
                }
                assert_eq!(style.border_left.as_deref(), Some("thin"));
                assert_eq!(style.border_right.as_deref(), Some("thin"));
                assert_eq!(style.border_top.as_deref(), Some("thin"));
                assert_eq!(style.border_bottom.as_deref(), Some("thin"));
            }
            other => panic!("expected StyledCell, got {other:?}"),
        }
    }

    #[test]
    fn test_styled_cell_alignment_roundtrip() {
        use crate::reader::xlsx::read_xlsx;
        use std::io::Cursor;

        let mut buf = Cursor::new(Vec::new());
        {
            let mut writer = StreamingXlsxWriter::new(&mut buf);
            writer.add_sheet("Styled").unwrap();
            writer
                .write_row(&[CellValue::StyledCell {
                    value: Box::new(CellValue::String("Centered".to_string())),
                    style: Box::new(CellStyle {
                        horizontal_alignment: Some("center".to_string()),
                        vertical_alignment: Some("top".to_string()),
                        wrap_text: true,
                        ..CellStyle::default()
                    }),
                }])
                .unwrap();
            writer.close().unwrap();
        }

        buf.set_position(0);
        let sheets = read_xlsx(buf).unwrap();
        match &sheets[0].rows[0][0] {
            CellValue::StyledCell { value, style } => {
                match value.as_ref() {
                    CellValue::String(s) => assert_eq!(s, "Centered"),
                    other => panic!("expected String inside StyledCell, got {other:?}"),
                }
                assert_eq!(style.horizontal_alignment.as_deref(), Some("center"));
                assert_eq!(style.vertical_alignment.as_deref(), Some("top"));
                assert!(style.wrap_text);
            }
            other => panic!("expected StyledCell, got {other:?}"),
        }
    }
}
