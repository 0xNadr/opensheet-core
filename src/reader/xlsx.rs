use std::collections::HashMap;
use std::io::{BufReader, Read, Seek};

use quick_xml::events::Event;
use quick_xml::Reader;
use zip::ZipArchive;

use crate::types::{CellValue, Sheet};

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

    // 4. Parse each worksheet
    let mut sheets = Vec::with_capacity(sheet_infos.len());
    for info in &sheet_infos {
        let rows = parse_worksheet(&mut archive, &info.path, &shared_strings)?;
        sheets.push(Sheet {
            name: info.name.clone(),
            rows,
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

/// Parse a single worksheet XML file and return rows of cell values.
fn parse_worksheet<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    path: &str,
    shared_strings: &[String],
) -> Result<Vec<Vec<CellValue>>, XlsxError> {
    let file = archive.by_name(path)?;
    let mut reader = Reader::from_reader(BufReader::new(file));
    let mut buf = Vec::new();

    let mut rows: Vec<Vec<CellValue>> = Vec::new();
    let mut current_row: usize = 0;
    let mut current_col: usize = 0;
    let mut cell_type = String::new();
    let mut in_cell = false;
    let mut in_value = false;
    let mut in_formula = false;
    let mut in_inline_str = false;
    let mut cell_value_text = String::new();
    let mut cell_formula_text = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                match e.name().as_ref() {
                    b"row" => {
                        // Get row number from r attribute
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"r" {
                                let row_num: usize = String::from_utf8_lossy(&attr.value)
                                    .parse()
                                    .unwrap_or(current_row + 2);
                                current_row = row_num - 1;
                            }
                        }
                        // Ensure rows vec is large enough
                        while rows.len() <= current_row {
                            rows.push(Vec::new());
                        }
                    }
                    b"c" => {
                        in_cell = true;
                        cell_type.clear();
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
                            let value = if !cell_formula_text.is_empty() {
                                // Cell has a formula
                                let cached = resolve_cell_value(
                                    &cell_type,
                                    &cell_value_text,
                                    shared_strings,
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
                                resolve_cell_value(&cell_type, &cell_value_text, shared_strings)
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
                    _ => {}
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

    Ok(rows)
}

/// Convert raw cell value text into a typed CellValue based on the cell type attribute.
fn resolve_cell_value(cell_type: &str, raw: &str, shared_strings: &[String]) -> CellValue {
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
        // Number (default or explicit "n")
        _ => {
            if let Ok(n) = raw.parse::<f64>() {
                CellValue::Number(n)
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

        match resolve_cell_value("s", "0", &shared) {
            CellValue::String(s) => assert_eq!(s, "hello"),
            _ => panic!("expected string"),
        }

        match resolve_cell_value("", "42.5", &shared) {
            CellValue::Number(n) => assert!((n - 42.5).abs() < f64::EPSILON),
            _ => panic!("expected number"),
        }

        match resolve_cell_value("b", "1", &shared) {
            CellValue::Bool(b) => assert!(b),
            _ => panic!("expected bool"),
        }

        assert!(matches!(
            resolve_cell_value("", "", &shared),
            CellValue::Empty
        ));
    }
}
