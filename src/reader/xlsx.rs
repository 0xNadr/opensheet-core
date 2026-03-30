use std::collections::HashMap;
use std::io::{BufReader, Read, Seek};

use quick_xml::events::Event;
use quick_xml::Reader;
use zip::ZipArchive;

use crate::types::{excel_serial_to_datetime, CellValue, Sheet};

/// Errors that can occur during XLSX reading.
#[derive(Debug)]
pub enum XlsxError {
    Zip(zip::result::ZipError),
    Xml(quick_xml::Error),
    Io(std::io::Error),
    InvalidStructure(String),
}

impl std::fmt::Display for XlsxError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            XlsxError::Zip(e) => write!(f, "ZIP error: {e}"),
            XlsxError::Xml(e) => write!(f, "XML error: {e}"),
            XlsxError::Io(e) => write!(f, "IO error: {e}"),
            XlsxError::InvalidStructure(msg) => write!(f, "Invalid XLSX: {msg}"),
        }
    }
}

impl From<zip::result::ZipError> for XlsxError {
    fn from(e: zip::result::ZipError) -> Self {
        XlsxError::Zip(e)
    }
}

impl From<quick_xml::Error> for XlsxError {
    fn from(e: quick_xml::Error) -> Self {
        XlsxError::Xml(e)
    }
}

impl From<std::io::Error> for XlsxError {
    fn from(e: std::io::Error) -> Self {
        XlsxError::Io(e)
    }
}

impl From<XlsxError> for pyo3::PyErr {
    fn from(e: XlsxError) -> Self {
        pyo3::exceptions::PyValueError::new_err(e.to_string())
    }
}

/// Represents the relationship between sheet IDs and their file paths.
struct SheetInfo {
    name: String,
    path: String,
}

/// Read an XLSX file and return all sheets.
pub fn read_xlsx<R: Read + Seek>(reader: R) -> Result<Vec<Sheet>, XlsxError> {
    let mut archive = ZipArchive::new(reader)?;

    // 1. Parse relationships to map rId -> file path
    let rels = parse_workbook_rels(&mut archive)?;

    // 2. Parse workbook.xml to get sheet names and their rIds
    let sheet_infos = parse_workbook(&mut archive, &rels)?;

    // 3. Parse shared strings
    let shared_strings = parse_shared_strings(&mut archive)?;

    // 4. Parse styles for date detection
    let date_styles = parse_styles(&mut archive)?;

    // 5. Parse each worksheet
    let mut sheets = Vec::with_capacity(sheet_infos.len());
    for info in &sheet_infos {
        let data = parse_worksheet(&mut archive, &info.path, &shared_strings, &date_styles)?;
        sheets.push(Sheet {
            name: info.name.clone(),
            rows: data.rows,
            merges: data.merges,
            column_widths: data.column_widths,
            row_heights: data.row_heights,
            freeze_pane: data.freeze_pane,
        });
    }

    Ok(sheets)
}

/// Parse xl/_rels/workbook.xml.rels to get rId -> target path mapping.
fn parse_workbook_rels<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
) -> Result<HashMap<String, String>, XlsxError> {
    let mut rels = HashMap::new();

    let file = match archive.by_name("xl/_rels/workbook.xml.rels") {
        Ok(f) => f,
        Err(_) => return Ok(rels),
    };

    let mut reader = Reader::from_reader(BufReader::new(file));
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e))
                if e.name().as_ref() == b"Relationship" =>
            {
                let mut id = String::new();
                let mut target = String::new();

                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"Id" => id = String::from_utf8_lossy(&attr.value).to_string(),
                        b"Target" => target = String::from_utf8_lossy(&attr.value).to_string(),
                        _ => {}
                    }
                }

                if !id.is_empty() && !target.is_empty() {
                    // Targets are relative to xl/
                    let full_path = if let Some(stripped) = target.strip_prefix('/') {
                        stripped.to_string()
                    } else {
                        format!("xl/{target}")
                    };
                    rels.insert(id, full_path);
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(rels)
}

/// Parse xl/workbook.xml to get sheet names and their relationship IDs.
fn parse_workbook<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    rels: &HashMap<String, String>,
) -> Result<Vec<SheetInfo>, XlsxError> {
    let file = archive.by_name("xl/workbook.xml")?;
    let mut reader = Reader::from_reader(BufReader::new(file));
    let mut buf = Vec::new();
    let mut sheets = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) if e.name().as_ref() == b"sheet" => {
                let mut name = String::new();
                let mut r_id = String::new();

                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"name" => name = String::from_utf8_lossy(&attr.value).to_string(),
                        b"r:id" => r_id = String::from_utf8_lossy(&attr.value).to_string(),
                        _ => {}
                    }
                }

                if !name.is_empty() {
                    let path = rels
                        .get(&r_id)
                        .cloned()
                        .unwrap_or_else(|| format!("xl/worksheets/sheet{}.xml", sheets.len() + 1));
                    sheets.push(SheetInfo { name, path });
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(sheets)
}

/// Parse xl/sharedStrings.xml to build the shared string table.
fn parse_shared_strings<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
) -> Result<Vec<String>, XlsxError> {
    let file = match archive.by_name("xl/sharedStrings.xml") {
        Ok(f) => f,
        Err(_) => return Ok(Vec::new()),
    };

    let mut reader = Reader::from_reader(BufReader::new(file));
    let mut buf = Vec::new();
    let mut strings = Vec::new();
    let mut current_string = String::new();
    let mut in_si = false;
    let mut in_t = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match e.name().as_ref() {
                b"si" => {
                    in_si = true;
                    current_string.clear();
                }
                b"t" if in_si => {
                    in_t = true;
                }
                _ => {}
            },
            Ok(Event::End(ref e)) => match e.name().as_ref() {
                b"si" => {
                    in_si = false;
                    strings.push(std::mem::take(&mut current_string));
                }
                b"t" => {
                    in_t = false;
                }
                _ => {}
            },
            Ok(Event::Text(ref e)) if in_t => {
                let text = e.unescape().unwrap_or_default();
                current_string.push_str(&text);
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(strings)
}

/// Parse xl/styles.xml to determine which cell style indices use date number formats.
/// Returns a set of xf indices (0-based) that are date-formatted.
fn parse_styles<R: Read + Seek>(archive: &mut ZipArchive<R>) -> Result<Vec<bool>, XlsxError> {
    let file = match archive.by_name("xl/styles.xml") {
        Ok(f) => f,
        Err(_) => return Ok(Vec::new()),
    };

    let mut reader = Reader::from_reader(BufReader::new(file));
    let mut buf = Vec::new();

    // Custom number formats: numFmtId -> formatCode
    let mut custom_formats: HashMap<u32, String> = HashMap::new();
    // The numFmtId for each xf in cellXfs
    let mut xf_num_fmt_ids: Vec<u32> = Vec::new();
    let mut in_num_fmts = false;
    let mut in_cell_xfs = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match e.name().as_ref() {
                b"numFmts" => in_num_fmts = true,
                b"cellXfs" => in_cell_xfs = true,
                b"xf" if in_cell_xfs => {
                    let mut num_fmt_id: u32 = 0;
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"numFmtId" {
                            num_fmt_id = String::from_utf8_lossy(&attr.value).parse().unwrap_or(0);
                        }
                    }
                    xf_num_fmt_ids.push(num_fmt_id);
                }
                _ => {}
            },
            Ok(Event::Empty(ref e)) => match e.name().as_ref() {
                b"numFmt" if in_num_fmts => {
                    let mut id: u32 = 0;
                    let mut code = String::new();
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"numFmtId" => {
                                id = String::from_utf8_lossy(&attr.value).parse().unwrap_or(0);
                            }
                            b"formatCode" => {
                                code = String::from_utf8_lossy(&attr.value).to_string();
                            }
                            _ => {}
                        }
                    }
                    if id > 0 {
                        custom_formats.insert(id, code);
                    }
                }
                b"xf" if in_cell_xfs => {
                    let mut num_fmt_id: u32 = 0;
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"numFmtId" {
                            num_fmt_id = String::from_utf8_lossy(&attr.value).parse().unwrap_or(0);
                        }
                    }
                    xf_num_fmt_ids.push(num_fmt_id);
                }
                _ => {}
            },
            Ok(Event::End(ref e)) => match e.name().as_ref() {
                b"numFmts" => in_num_fmts = false,
                b"cellXfs" => in_cell_xfs = false,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    // Build a bool vec: is_date[xf_index] = true if the format is a date format
    let is_date: Vec<bool> = xf_num_fmt_ids
        .iter()
        .map(|&fmt_id| is_date_format(fmt_id, &custom_formats))
        .collect();

    Ok(is_date)
}

/// Check if a number format ID represents a date/time format.
/// Built-in date format IDs: 14-22, 27-36, 45-47, 50-58.
/// Custom formats are checked for date/time pattern characters.
fn is_date_format(num_fmt_id: u32, custom_formats: &HashMap<u32, String>) -> bool {
    // Built-in date/time formats
    match num_fmt_id {
        14..=22 | 27..=36 | 45..=47 | 50..=58 => return true,
        0 => return false,      // General
        1..=13 => return false, // Number formats
        _ => {}
    }

    // Check custom format codes
    if let Some(code) = custom_formats.get(&num_fmt_id) {
        is_date_format_code(code)
    } else {
        false
    }
}

/// Check if a format code string looks like a date/time format.
fn is_date_format_code(code: &str) -> bool {
    let lower = code.to_lowercase();
    // Skip text in quotes and brackets
    let mut cleaned = String::new();
    let mut in_quotes = false;
    let mut in_bracket = false;
    for ch in lower.chars() {
        match ch {
            '"' => in_quotes = !in_quotes,
            '[' if !in_quotes => in_bracket = true,
            ']' if !in_quotes => in_bracket = false,
            _ if !in_quotes && !in_bracket => cleaned.push(ch),
            _ => {}
        }
    }

    // Look for date/time characters (y, m, d, h, s) but not pure number formats
    let has_date = cleaned.contains('y') || cleaned.contains('d');
    let has_time = cleaned.contains('h') || cleaned.contains('s');
    // 'm' alone could be minutes (if near h/s) or months
    let has_m = cleaned.contains('m');

    has_date || has_time || (has_m && !cleaned.contains('#') && !cleaned.contains('0'))
}

/// Parse a column reference like "A1", "AA5", "BZ100" and return the 0-based column index.
fn col_to_index(col_ref: &str) -> usize {
    let mut index: usize = 0;
    for b in col_ref.bytes() {
        if b.is_ascii_alphabetic() {
            index = index * 26 + (b.to_ascii_uppercase() - b'A') as usize + 1;
        } else {
            break;
        }
    }
    if index == 0 {
        0
    } else {
        index - 1
    }
}

/// Parse a cell reference like "A1" and return (row_0based, col_0based).
fn parse_cell_ref(cell_ref: &str) -> (usize, usize) {
    let col_end = cell_ref
        .bytes()
        .position(|b| b.is_ascii_digit())
        .unwrap_or(cell_ref.len());
    let col = col_to_index(&cell_ref[..col_end]);
    let row: usize = cell_ref[col_end..].parse().unwrap_or(1);
    (row.saturating_sub(1), col)
}

/// Parsed worksheet data: rows, merge ranges, column widths, row heights, and freeze pane.
struct WorksheetData {
    rows: Vec<Vec<CellValue>>,
    merges: Vec<String>,
    column_widths: HashMap<u32, f64>,
    row_heights: HashMap<u32, f64>,
    freeze_pane: Option<(u32, u32)>,
}

/// Parse a single worksheet XML file and return rows of cell values and merge ranges.
fn parse_worksheet<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    path: &str,
    shared_strings: &[String],
    date_styles: &[bool],
) -> Result<WorksheetData, XlsxError> {
    let file = archive.by_name(path)?;
    let mut reader = Reader::from_reader(BufReader::new(file));
    let mut buf = Vec::new();

    let mut rows: Vec<Vec<CellValue>> = Vec::new();
    let mut current_row: usize = 0;
    let mut current_col: usize = 0;
    let mut cell_type = String::new();
    let mut cell_style: usize = 0;
    let mut in_cell = false;
    let mut in_value = false;
    let mut in_formula = false;
    let mut in_inline_str = false;
    let mut cell_value_text = String::new();
    let mut cell_formula_text = String::new();
    let mut merges: Vec<String> = Vec::new();
    let mut in_merge_cells = false;
    let mut column_widths: HashMap<u32, f64> = HashMap::new();
    let mut row_heights: HashMap<u32, f64> = HashMap::new();
    let mut in_cols = false;
    let mut freeze_pane: Option<(u32, u32)> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                match e.name().as_ref() {
                    b"pane" => {
                        let mut y_split: u32 = 0;
                        let mut x_split: u32 = 0;
                        let mut state = String::new();
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"ySplit" => {
                                    y_split =
                                        String::from_utf8_lossy(&attr.value).parse().unwrap_or(0);
                                }
                                b"xSplit" => {
                                    x_split =
                                        String::from_utf8_lossy(&attr.value).parse().unwrap_or(0);
                                }
                                b"state" => {
                                    state = String::from_utf8_lossy(&attr.value).to_string();
                                }
                                _ => {}
                            }
                        }
                        if state == "frozen" && (y_split > 0 || x_split > 0) {
                            freeze_pane = Some((y_split, x_split));
                        }
                    }
                    b"cols" => {
                        in_cols = true;
                    }
                    b"col" if in_cols => {
                        let mut min: u32 = 0;
                        let mut max: u32 = 0;
                        let mut width: f64 = 0.0;
                        let mut custom_width = false;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"min" => {
                                    min = String::from_utf8_lossy(&attr.value).parse().unwrap_or(0);
                                }
                                b"max" => {
                                    max = String::from_utf8_lossy(&attr.value).parse().unwrap_or(0);
                                }
                                b"width" => {
                                    width =
                                        String::from_utf8_lossy(&attr.value).parse().unwrap_or(0.0);
                                }
                                b"customWidth" => {
                                    custom_width = String::from_utf8_lossy(&attr.value) == "1";
                                }
                                _ => {}
                            }
                        }
                        if custom_width && min > 0 && max >= min {
                            for col in min..=max {
                                column_widths.insert(col - 1, width); // 0-based
                            }
                        }
                    }
                    b"mergeCells" => {
                        in_merge_cells = true;
                    }
                    b"row" => {
                        // Get row number and optional custom height from attributes
                        let mut custom_height = false;
                        let mut height: f64 = 0.0;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"r" => {
                                    let row_num: usize = String::from_utf8_lossy(&attr.value)
                                        .parse()
                                        .unwrap_or(current_row + 2);
                                    current_row = row_num - 1;
                                }
                                b"ht" => {
                                    height =
                                        String::from_utf8_lossy(&attr.value).parse().unwrap_or(0.0);
                                }
                                b"customHeight" => {
                                    custom_height = String::from_utf8_lossy(&attr.value) == "1";
                                }
                                _ => {}
                            }
                        }
                        if custom_height && height > 0.0 {
                            row_heights.insert(current_row as u32, height);
                        }
                        // Ensure rows vec is large enough
                        while rows.len() <= current_row {
                            rows.push(Vec::new());
                        }
                    }
                    b"c" => {
                        in_cell = true;
                        cell_type.clear();
                        cell_style = 0;
                        cell_value_text.clear();
                        cell_formula_text.clear();

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"r" => {
                                    let cell_ref = String::from_utf8_lossy(&attr.value).to_string();
                                    let (_, col) = parse_cell_ref(&cell_ref);
                                    current_col = col;
                                }
                                b"t" => {
                                    cell_type = String::from_utf8_lossy(&attr.value).to_string();
                                }
                                b"s" => {
                                    cell_style =
                                        String::from_utf8_lossy(&attr.value).parse().unwrap_or(0);
                                }
                                _ => {}
                            }
                        }
                    }
                    b"v" if in_cell => {
                        in_value = true;
                        cell_value_text.clear();
                    }
                    b"f" if in_cell => {
                        in_formula = true;
                        cell_formula_text.clear();
                    }
                    b"is" if in_cell => {
                        in_inline_str = true;
                    }
                    b"t" if in_inline_str => {
                        in_value = true;
                        cell_value_text.clear();
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                match e.name().as_ref() {
                    b"c" => {
                        if in_cell {
                            let is_date = date_styles.get(cell_style).copied().unwrap_or(false);
                            let value = if !cell_formula_text.is_empty() {
                                // Cell has a formula
                                let cached = resolve_cell_value(
                                    &cell_type,
                                    &cell_value_text,
                                    shared_strings,
                                    is_date,
                                );
                                let cached_value = match cached {
                                    CellValue::Empty => None,
                                    other => Some(Box::new(other)),
                                };
                                CellValue::Formula {
                                    formula: std::mem::take(&mut cell_formula_text),
                                    cached_value,
                                }
                            } else {
                                resolve_cell_value(
                                    &cell_type,
                                    &cell_value_text,
                                    shared_strings,
                                    is_date,
                                )
                            };

                            // Ensure the row has enough columns
                            if let Some(row) = rows.get_mut(current_row) {
                                while row.len() <= current_col {
                                    row.push(CellValue::Empty);
                                }
                                row[current_col] = value;
                            }

                            in_cell = false;
                            in_inline_str = false;
                        }
                    }
                    b"v" => {
                        in_value = false;
                    }
                    b"f" => {
                        in_formula = false;
                    }
                    b"t" if in_inline_str => {
                        in_value = false;
                    }
                    b"is" => {
                        in_inline_str = false;
                    }
                    b"cols" => {
                        in_cols = false;
                    }
                    b"mergeCells" => {
                        in_merge_cells = false;
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) if e.name().as_ref() == b"pane" => {
                let mut y_split: u32 = 0;
                let mut x_split: u32 = 0;
                let mut state = String::new();
                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"ySplit" => {
                            y_split = String::from_utf8_lossy(&attr.value).parse().unwrap_or(0);
                        }
                        b"xSplit" => {
                            x_split = String::from_utf8_lossy(&attr.value).parse().unwrap_or(0);
                        }
                        b"state" => {
                            state = String::from_utf8_lossy(&attr.value).to_string();
                        }
                        _ => {}
                    }
                }
                if state == "frozen" && (y_split > 0 || x_split > 0) {
                    freeze_pane = Some((y_split, x_split));
                }
            }
            Ok(Event::Empty(ref e)) if in_cols && e.name().as_ref() == b"col" => {
                let mut min: u32 = 0;
                let mut max: u32 = 0;
                let mut width: f64 = 0.0;
                let mut custom_width = false;
                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"min" => {
                            min = String::from_utf8_lossy(&attr.value).parse().unwrap_or(0);
                        }
                        b"max" => {
                            max = String::from_utf8_lossy(&attr.value).parse().unwrap_or(0);
                        }
                        b"width" => {
                            width = String::from_utf8_lossy(&attr.value).parse().unwrap_or(0.0);
                        }
                        b"customWidth" => {
                            custom_width = String::from_utf8_lossy(&attr.value) == "1";
                        }
                        _ => {}
                    }
                }
                if custom_width && min > 0 && max >= min {
                    for col in min..=max {
                        column_widths.insert(col - 1, width); // 0-based
                    }
                }
            }
            Ok(Event::Empty(ref e)) if in_merge_cells && e.name().as_ref() == b"mergeCell" => {
                for attr in e.attributes().flatten() {
                    if attr.key.as_ref() == b"ref" {
                        merges.push(String::from_utf8_lossy(&attr.value).to_string());
                    }
                }
            }
            Ok(Event::Text(ref e)) if in_formula => {
                let text = e.unescape().unwrap_or_default();
                cell_formula_text.push_str(&text);
            }
            Ok(Event::Text(ref e)) if in_value => {
                let text = e.unescape().unwrap_or_default();
                cell_value_text.push_str(&text);
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(WorksheetData {
        rows,
        merges,
        column_widths,
        row_heights,
        freeze_pane,
    })
}

/// Convert raw cell value text into a typed CellValue based on the cell type attribute.
fn resolve_cell_value(
    cell_type: &str,
    raw: &str,
    shared_strings: &[String],
    is_date: bool,
) -> CellValue {
    if raw.is_empty() && cell_type != "inlineStr" {
        return CellValue::Empty;
    }

    match cell_type {
        // Shared string
        "s" => {
            if let Ok(idx) = raw.parse::<usize>() {
                shared_strings
                    .get(idx)
                    .map(|s| CellValue::String(s.clone()))
                    .unwrap_or(CellValue::Empty)
            } else {
                CellValue::String(raw.to_string())
            }
        }
        // Boolean
        "b" => CellValue::Bool(raw == "1" || raw.eq_ignore_ascii_case("true")),
        // Inline string
        "inlineStr" => CellValue::String(raw.to_string()),
        // Error
        "e" => CellValue::String(raw.to_string()),
        // String (str type)
        "str" => CellValue::String(raw.to_string()),
        // Number (default or explicit "n") — may be a date if style says so
        _ => {
            if let Ok(n) = raw.parse::<f64>() {
                if is_date {
                    if let Some((y, mo, d, h, mi, s, us)) = excel_serial_to_datetime(n) {
                        if h == 0 && mi == 0 && s == 0 && us == 0 && n.fract() == 0.0 {
                            CellValue::Date {
                                year: y,
                                month: mo,
                                day: d,
                            }
                        } else {
                            CellValue::DateTime {
                                year: y,
                                month: mo,
                                day: d,
                                hour: h,
                                minute: mi,
                                second: s,
                                microsecond: us,
                            }
                        }
                    } else {
                        CellValue::Number(n)
                    }
                } else {
                    CellValue::Number(n)
                }
            } else {
                CellValue::String(raw.to_string())
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_col_to_index() {
        assert_eq!(col_to_index("A"), 0);
        assert_eq!(col_to_index("B"), 1);
        assert_eq!(col_to_index("Z"), 25);
        assert_eq!(col_to_index("AA"), 26);
        assert_eq!(col_to_index("AZ"), 51);
        assert_eq!(col_to_index("BA"), 52);
    }

    #[test]
    fn test_parse_cell_ref() {
        assert_eq!(parse_cell_ref("A1"), (0, 0));
        assert_eq!(parse_cell_ref("B3"), (2, 1));
        assert_eq!(parse_cell_ref("AA10"), (9, 26));
    }

    #[test]
    fn test_resolve_cell_value() {
        let shared = vec!["hello".to_string(), "world".to_string()];

        match resolve_cell_value("s", "0", &shared, false) {
            CellValue::String(s) => assert_eq!(s, "hello"),
            _ => panic!("expected string"),
        }

        match resolve_cell_value("", "42.5", &shared, false) {
            CellValue::Number(n) => assert!((n - 42.5).abs() < f64::EPSILON),
            _ => panic!("expected number"),
        }

        match resolve_cell_value("b", "1", &shared, false) {
            CellValue::Bool(b) => assert!(b),
            _ => panic!("expected bool"),
        }

        assert!(matches!(
            resolve_cell_value("", "", &shared, false),
            CellValue::Empty
        ));

        // Date detection
        match resolve_cell_value("", "44197", &shared, true) {
            CellValue::Date { year, month, day } => {
                assert_eq!((year, month, day), (2021, 1, 1));
            }
            other => panic!("expected date, got {other:?}"),
        }
    }
}
