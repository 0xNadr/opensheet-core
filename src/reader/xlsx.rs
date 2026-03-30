use std::collections::HashMap;
use std::io::{BufReader, Read, Seek};

use quick_xml::events::Event;
use quick_xml::Reader;
use zip::ZipArchive;

use crate::types::{excel_serial_to_datetime, CellStyle, CellValue, Sheet};

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
    /// Sheet visibility: "visible", "hidden", or "veryHidden".
    state: String,
}

/// A defined name (named range) from the workbook.
#[derive(Debug, Clone)]
pub struct DefinedName {
    pub name: String,
    pub value: String,
    /// If Some, the name is sheet-scoped (0-based sheet index). If None, workbook-scoped.
    pub sheet_index: Option<usize>,
}

/// Read an XLSX file and return all sheets plus the shared string table.
///
/// Cells referencing shared strings are stored as `CellValue::SharedString(index)`
/// to avoid cloning. The caller must resolve them using the returned shared strings.
#[allow(clippy::type_complexity)]
pub fn read_xlsx<R: Read + Seek>(
    reader: R,
) -> Result<(Vec<Sheet>, Vec<String>, Vec<DefinedName>), XlsxError> {
    let mut archive = ZipArchive::new(reader)?;

    // 1. Parse relationships to map rId -> file path
    let rels = parse_workbook_rels(&mut archive)?;

    // 2. Parse workbook.xml to get sheet names, rIds, and defined names
    let (sheet_infos, defined_names) = parse_workbook(&mut archive, &rels)?;

    // 3. Parse shared strings
    let shared_strings = parse_shared_strings(&mut archive)?;

    // 4. Parse styles for date detection, number formats, and cell styling
    let styles = parse_styles(&mut archive)?;

    // 5. Parse each worksheet
    let mut sheets = Vec::with_capacity(sheet_infos.len());
    for info in &sheet_infos {
        let data = parse_worksheet(&mut archive, &info.path, &shared_strings, &styles)?;
        sheets.push(Sheet {
            name: info.name.clone(),
            rows: data.rows,
            merges: data.merges,
            column_widths: data.column_widths,
            row_heights: data.row_heights,
            freeze_pane: data.freeze_pane,
            auto_filter: data.auto_filter,
            state: info.state.clone(),
        });
    }

    Ok((sheets, shared_strings, defined_names))
}

/// Read a single sheet by name or index, without parsing other worksheets.
///
/// Returns the sheet and the shared string table for resolving SharedString cells.
pub fn read_single_sheet<R: Read + Seek>(
    reader: R,
    sheet_name: Option<&str>,
    sheet_index: Option<usize>,
) -> Result<(Sheet, Vec<String>), XlsxError> {
    let mut archive = ZipArchive::new(reader)?;

    let rels = parse_workbook_rels(&mut archive)?;
    let (sheet_infos, _defined_names) = parse_workbook(&mut archive, &rels)?;
    let shared_strings = parse_shared_strings(&mut archive)?;
    let styles = parse_styles(&mut archive)?;

    let info = if let Some(name) = sheet_name {
        sheet_infos
            .iter()
            .find(|s| s.name == name)
            .ok_or_else(|| XlsxError::InvalidStructure(format!("Sheet '{name}' not found")))?
    } else if let Some(idx) = sheet_index {
        sheet_infos.get(idx).ok_or_else(|| {
            XlsxError::InvalidStructure(format!(
                "Sheet index {idx} out of range (file has {} sheets)",
                sheet_infos.len()
            ))
        })?
    } else {
        sheet_infos
            .first()
            .ok_or_else(|| XlsxError::InvalidStructure("No sheets found in file".to_string()))?
    };

    let data = parse_worksheet(&mut archive, &info.path, &shared_strings, &styles)?;
    let sheet = Sheet {
        name: info.name.clone(),
        rows: data.rows,
        merges: data.merges,
        column_widths: data.column_widths,
        row_heights: data.row_heights,
        freeze_pane: data.freeze_pane,
        auto_filter: data.auto_filter,
        state: info.state.clone(),
    };

    Ok((sheet, shared_strings))
}

/// List sheet names without parsing worksheet data.
pub fn read_sheet_names<R: Read + Seek>(reader: R) -> Result<Vec<String>, XlsxError> {
    let mut archive = ZipArchive::new(reader)?;
    let rels = parse_workbook_rels(&mut archive)?;
    let (sheet_infos, _) = parse_workbook(&mut archive, &rels)?;
    Ok(sheet_infos.into_iter().map(|s| s.name).collect())
}

/// Read defined names (named ranges) without parsing worksheet data.
pub fn read_defined_names<R: Read + Seek>(reader: R) -> Result<Vec<DefinedName>, XlsxError> {
    let mut archive = ZipArchive::new(reader)?;
    let rels = parse_workbook_rels(&mut archive)?;
    let (_, defined_names) = parse_workbook(&mut archive, &rels)?;
    Ok(defined_names)
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
) -> Result<(Vec<SheetInfo>, Vec<DefinedName>), XlsxError> {
    let file = archive.by_name("xl/workbook.xml")?;
    let mut reader = Reader::from_reader(BufReader::new(file));
    let mut buf = Vec::new();
    let mut sheets = Vec::new();
    let mut defined_names = Vec::new();

    // State for parsing <definedName> elements
    let mut in_defined_name = false;
    let mut dn_name = String::new();
    let mut dn_sheet_index: Option<usize> = None;
    let mut dn_value = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) if e.name().as_ref() == b"sheet" => {
                let mut name = String::new();
                let mut r_id = String::new();
                let mut state = String::from("visible");

                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"name" => name = String::from_utf8_lossy(&attr.value).to_string(),
                        b"r:id" => r_id = String::from_utf8_lossy(&attr.value).to_string(),
                        b"state" => state = String::from_utf8_lossy(&attr.value).to_string(),
                        _ => {}
                    }
                }

                if !name.is_empty() {
                    let path = rels
                        .get(&r_id)
                        .cloned()
                        .unwrap_or_else(|| format!("xl/worksheets/sheet{}.xml", sheets.len() + 1));
                    sheets.push(SheetInfo { name, path, state });
                }
            }
            Ok(Event::Start(ref e)) if e.name().as_ref() == b"definedName" => {
                in_defined_name = true;
                dn_name.clear();
                dn_sheet_index = None;
                dn_value.clear();

                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"name" => dn_name = String::from_utf8_lossy(&attr.value).to_string(),
                        b"localSheetId" => {
                            dn_sheet_index =
                                String::from_utf8_lossy(&attr.value).parse::<usize>().ok();
                        }
                        _ => {}
                    }
                }
            }
            Ok(Event::Text(ref e)) if in_defined_name => {
                if let Ok(text) = e.unescape() {
                    dn_value.push_str(&text);
                }
            }
            Ok(Event::End(ref e)) if e.name().as_ref() == b"definedName" => {
                if in_defined_name && !dn_name.is_empty() && !dn_value.is_empty() {
                    defined_names.push(DefinedName {
                        name: std::mem::take(&mut dn_name),
                        value: std::mem::take(&mut dn_value),
                        sheet_index: dn_sheet_index.take(),
                    });
                }
                in_defined_name = false;
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok((sheets, defined_names))
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

// ---------- Style parsing types ----------

/// Parsed font info from styles.xml.
#[derive(Debug, Clone, Default)]
struct ParsedFont {
    bold: bool,
    italic: bool,
    underline: bool,
    name: Option<String>,
    size: Option<f64>,
    color: Option<String>,
}

/// Parsed fill info from styles.xml.
#[derive(Debug, Clone, Default)]
struct ParsedFill {
    pattern_type: Option<String>,
    fg_color: Option<String>,
}

/// Parsed border side info.
#[derive(Debug, Clone, Default)]
struct ParsedBorderSide {
    style: Option<String>,
    color: Option<String>,
}

/// Parsed border info from styles.xml.
#[derive(Debug, Clone, Default)]
struct ParsedBorder {
    left: ParsedBorderSide,
    right: ParsedBorderSide,
    top: ParsedBorderSide,
    bottom: ParsedBorderSide,
}

/// Parsed alignment info from styles.xml.
#[derive(Debug, Clone, Default)]
struct ParsedAlignment {
    horizontal: Option<String>,
    vertical: Option<String>,
    wrap_text: bool,
    text_rotation: Option<u16>,
}

/// Style info for a single xf entry.
#[derive(Debug, Clone)]
struct StyleInfo {
    is_date: bool,
    /// The format code string, if it's a custom (non-built-in, non-date) format.
    format_code: Option<String>,
    /// Cell style info, populated when the xf has non-default styling.
    cell_style: Option<CellStyle>,
}

/// Parse xl/styles.xml to determine which cell style indices use date number formats
/// and to extract font/fill/border/alignment information for cell styling.
fn parse_styles<R: Read + Seek>(archive: &mut ZipArchive<R>) -> Result<Vec<StyleInfo>, XlsxError> {
    let file = match archive.by_name("xl/styles.xml") {
        Ok(f) => f,
        Err(_) => return Ok(Vec::new()),
    };

    let mut reader = Reader::from_reader(BufReader::new(file));
    let mut buf = Vec::new();

    // Custom number formats: numFmtId -> formatCode
    let mut custom_formats: HashMap<u32, String> = HashMap::new();

    // Parsed style components
    let mut parsed_fonts: Vec<ParsedFont> = Vec::new();
    let mut parsed_fills: Vec<ParsedFill> = Vec::new();
    let mut parsed_borders: Vec<ParsedBorder> = Vec::new();

    // Per-xf data from cellXfs
    struct XfData {
        num_fmt_id: u32,
        font_id: usize,
        fill_id: usize,
        border_id: usize,
        alignment: Option<ParsedAlignment>,
    }
    let mut xf_entries: Vec<XfData> = Vec::new();

    // Section tracking flags
    let mut in_num_fmts = false;
    let mut in_cell_xfs = false;
    let mut in_fonts = false;
    let mut in_font = false;
    let mut in_fills = false;
    let mut in_fill = false;
    let mut in_pattern_fill = false;
    let mut in_borders = false;
    let mut in_border = false;
    let mut in_xf = false;

    // Current element being built
    let mut current_font = ParsedFont::default();
    let mut current_fill = ParsedFill::default();
    let mut current_border = ParsedBorder::default();
    let mut current_xf_data: Option<XfData> = None;

    // Border side tracking: which side are we inside?
    let mut current_border_side: u8 = 0; // 0=none, 1=left, 2=right, 3=top, 4=bottom

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match e.name().as_ref() {
                b"numFmts" => in_num_fmts = true,
                b"fonts" => in_fonts = true,
                b"font" if in_fonts => {
                    in_font = true;
                    current_font = ParsedFont::default();
                }
                b"fills" => in_fills = true,
                b"fill" if in_fills => {
                    in_fill = true;
                    current_fill = ParsedFill::default();
                }
                b"patternFill" if in_fill => {
                    in_pattern_fill = true;
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"patternType" {
                            current_fill.pattern_type =
                                Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                }
                b"borders" => in_borders = true,
                b"border" if in_borders => {
                    in_border = true;
                    current_border = ParsedBorder::default();
                }
                b"left" if in_border => {
                    current_border_side = 1;
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"style" {
                            current_border.left.style =
                                Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                }
                b"right" if in_border => {
                    current_border_side = 2;
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"style" {
                            current_border.right.style =
                                Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                }
                b"top" if in_border => {
                    current_border_side = 3;
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"style" {
                            current_border.top.style =
                                Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                }
                b"bottom" if in_border => {
                    current_border_side = 4;
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"style" {
                            current_border.bottom.style =
                                Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                }
                b"cellXfs" => in_cell_xfs = true,
                b"xf" if in_cell_xfs => {
                    in_xf = true;
                    let mut num_fmt_id: u32 = 0;
                    let mut font_id: usize = 0;
                    let mut fill_id: usize = 0;
                    let mut border_id: usize = 0;
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"numFmtId" => {
                                num_fmt_id = parse_u32_from_bytes(&attr.value);
                            }
                            b"fontId" => {
                                font_id = parse_u32_from_bytes(&attr.value) as usize;
                            }
                            b"fillId" => {
                                fill_id = parse_u32_from_bytes(&attr.value) as usize;
                            }
                            b"borderId" => {
                                border_id = parse_u32_from_bytes(&attr.value) as usize;
                            }
                            _ => {}
                        }
                    }
                    current_xf_data = Some(XfData {
                        num_fmt_id,
                        font_id,
                        fill_id,
                        border_id,
                        alignment: None,
                    });
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
                                id = parse_u32_from_bytes(&attr.value);
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
                // Font child elements (empty tags)
                b"b" if in_font => current_font.bold = true,
                b"i" if in_font => current_font.italic = true,
                b"u" if in_font => current_font.underline = true,
                b"sz" if in_font => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"val" {
                            current_font.size = std::str::from_utf8(&attr.value)
                                .ok()
                                .and_then(|s| s.parse().ok());
                        }
                    }
                }
                b"name" if in_font => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"val" {
                            current_font.name =
                                Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                }
                b"color" if in_font => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"rgb" {
                            current_font.color =
                                Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                }
                // Fill child: patternFill as empty element
                b"patternFill" if in_fill => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"patternType" {
                            current_fill.pattern_type =
                                Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                }
                // fgColor inside patternFill
                b"fgColor" if in_pattern_fill => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"rgb" {
                            current_fill.fg_color =
                                Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                }
                // Border sides as empty elements
                b"left" if in_border => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"style" {
                            current_border.left.style =
                                Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                }
                b"right" if in_border => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"style" {
                            current_border.right.style =
                                Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                }
                b"top" if in_border => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"style" {
                            current_border.top.style =
                                Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                }
                b"bottom" if in_border => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"style" {
                            current_border.bottom.style =
                                Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                }
                // Color inside a border side
                b"color" if current_border_side > 0 => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"rgb" {
                            let color = Some(String::from_utf8_lossy(&attr.value).to_string());
                            match current_border_side {
                                1 => current_border.left.color = color,
                                2 => current_border.right.color = color,
                                3 => current_border.top.color = color,
                                4 => current_border.bottom.color = color,
                                _ => {}
                            }
                        }
                    }
                }
                // Alignment inside xf
                b"alignment" if in_xf => {
                    if let Some(ref mut xf) = current_xf_data {
                        let mut align = ParsedAlignment::default();
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"horizontal" => {
                                    align.horizontal =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                                b"vertical" => {
                                    align.vertical =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                                b"wrapText" => {
                                    align.wrap_text = attr.value.as_ref() == b"1";
                                }
                                b"textRotation" => {
                                    align.text_rotation = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                _ => {}
                            }
                        }
                        xf.alignment = Some(align);
                    }
                }
                // xf as empty element in cellXfs
                b"xf" if in_cell_xfs && !in_xf => {
                    let mut num_fmt_id: u32 = 0;
                    let mut font_id: usize = 0;
                    let mut fill_id: usize = 0;
                    let mut border_id: usize = 0;
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"numFmtId" => {
                                num_fmt_id = parse_u32_from_bytes(&attr.value);
                            }
                            b"fontId" => {
                                font_id = parse_u32_from_bytes(&attr.value) as usize;
                            }
                            b"fillId" => {
                                fill_id = parse_u32_from_bytes(&attr.value) as usize;
                            }
                            b"borderId" => {
                                border_id = parse_u32_from_bytes(&attr.value) as usize;
                            }
                            _ => {}
                        }
                    }
                    xf_entries.push(XfData {
                        num_fmt_id,
                        font_id,
                        fill_id,
                        border_id,
                        alignment: None,
                    });
                }
                _ => {}
            },
            Ok(Event::End(ref e)) => match e.name().as_ref() {
                b"numFmts" => in_num_fmts = false,
                b"fonts" => in_fonts = false,
                b"font" if in_font => {
                    in_font = false;
                    parsed_fonts.push(std::mem::take(&mut current_font));
                }
                b"fills" => in_fills = false,
                b"fill" if in_fill => {
                    in_fill = false;
                    in_pattern_fill = false;
                    parsed_fills.push(std::mem::take(&mut current_fill));
                }
                b"patternFill" => in_pattern_fill = false,
                b"borders" => in_borders = false,
                b"border" if in_border => {
                    in_border = false;
                    current_border_side = 0;
                    parsed_borders.push(std::mem::take(&mut current_border));
                }
                b"left" if in_border => current_border_side = 0,
                b"right" if in_border => current_border_side = 0,
                b"top" if in_border => current_border_side = 0,
                b"bottom" if in_border => current_border_side = 0,
                b"cellXfs" => in_cell_xfs = false,
                b"xf" if in_xf => {
                    in_xf = false;
                    if let Some(xf) = current_xf_data.take() {
                        xf_entries.push(xf);
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    // Build style info per xf index
    let default_font = ParsedFont {
        bold: false,
        italic: false,
        underline: false,
        name: Some("Calibri".to_string()),
        size: Some(11.0),
        color: None,
    };

    let styles: Vec<StyleInfo> = xf_entries
        .iter()
        .map(|xf| {
            let is_date = is_date_format(xf.num_fmt_id, &custom_formats);
            let format_code = if is_date || xf.num_fmt_id == 0 {
                None
            } else {
                get_format_code(xf.num_fmt_id, &custom_formats)
            };

            // Build CellStyle if there's non-default styling
            let font = parsed_fonts.get(xf.font_id);
            let fill = parsed_fills.get(xf.fill_id);
            let border = parsed_borders.get(xf.border_id);

            let has_font_styling = font.is_some_and(|f| {
                f.bold
                    || f.italic
                    || f.underline
                    || (f.name.is_some() && f.name.as_deref() != Some("Calibri"))
                    || (f.size.is_some() && f.size != Some(11.0))
                    || f.color.is_some()
            });

            let has_fill_styling = fill.is_some_and(|f| {
                f.pattern_type.as_deref() == Some("solid") && f.fg_color.is_some()
            });

            let has_border_styling = border.is_some_and(|b| {
                b.left.style.is_some()
                    || b.right.style.is_some()
                    || b.top.style.is_some()
                    || b.bottom.style.is_some()
            });

            let has_alignment = xf.alignment.is_some();

            let cell_style =
                if has_font_styling || has_fill_styling || has_border_styling || has_alignment {
                    let f = font.unwrap_or(&default_font);
                    let mut style = CellStyle::default();

                    // Font
                    if has_font_styling {
                        style.bold = f.bold;
                        style.italic = f.italic;
                        style.underline = f.underline;
                        if f.name.as_deref() != Some("Calibri") {
                            style.font_name = f.name.clone();
                        }
                        if f.size != Some(11.0) {
                            style.font_size = f.size;
                        }
                        style.font_color = f.color.clone();
                    }

                    // Fill
                    if has_fill_styling {
                        if let Some(fl) = fill {
                            style.fill_color = fl.fg_color.clone();
                        }
                    }

                    // Border
                    if has_border_styling {
                        if let Some(b) = border {
                            style.border_left = b.left.style.clone();
                            style.border_right = b.right.style.clone();
                            style.border_top = b.top.style.clone();
                            style.border_bottom = b.bottom.style.clone();
                            // Use the first non-None border color as the shared color
                            let first_color = b
                                .left
                                .color
                                .as_ref()
                                .or(b.right.color.as_ref())
                                .or(b.top.color.as_ref())
                                .or(b.bottom.color.as_ref());
                            style.border_color = first_color.cloned();
                        }
                    }

                    // Alignment
                    if let Some(ref align) = xf.alignment {
                        style.horizontal_alignment = align.horizontal.clone();
                        style.vertical_alignment = align.vertical.clone();
                        style.wrap_text = align.wrap_text;
                        style.text_rotation = align.text_rotation;
                    }

                    // Number format (include in style if there's also visual styling)
                    if format_code.is_some() {
                        style.number_format = format_code.clone();
                    }

                    Some(style)
                } else {
                    None
                };

            StyleInfo {
                is_date,
                format_code,
                cell_style,
            }
        })
        .collect();

    Ok(styles)
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

/// Get the format code string for a numFmtId.
/// Returns None for general (0) format. Returns built-in code for IDs 1-49,
/// or looks up custom formats for IDs >= 164.
fn get_format_code(num_fmt_id: u32, custom_formats: &HashMap<u32, String>) -> Option<String> {
    // Built-in number format codes (non-date ones)
    let builtin = match num_fmt_id {
        0 => return None, // General
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        5 => "$#,##0_);($#,##0)",
        6 => "$#,##0_);[Red]($#,##0)",
        7 => "$#,##0.00_);($#,##0.00)",
        8 => "$#,##0.00_);[Red]($#,##0.00)",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        12 => "# ?/?",
        13 => "# ??/??",
        37 => "#,##0_);(#,##0)",
        38 => "#,##0_);[Red](#,##0)",
        39 => "#,##0.00_);(#,##0.00)",
        40 => "#,##0.00_);[Red](#,##0.00)",
        41 => r#"_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)"#,
        42 => r#"_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)"#,
        43 => r#"_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)"#,
        44 => r#"_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)"#,
        48 => "##0.0E+0",
        49 => "@",
        _ => {
            // Custom format or date format (dates already filtered out)
            return custom_formats.get(&num_fmt_id).cloned();
        }
    };
    Some(builtin.to_string())
}

/// Extract the column index directly from a cell reference byte slice (e.g. b"A1", b"AB12")
/// without allocating a String.
fn col_from_cell_ref_bytes(bytes: &[u8]) -> usize {
    let col_end = bytes
        .iter()
        .position(|b| b.is_ascii_digit())
        .unwrap_or(bytes.len());
    let mut col: usize = 0;
    for &b in &bytes[..col_end] {
        let c = b.to_ascii_uppercase();
        if c.is_ascii_uppercase() {
            col = col * 26 + (c - b'A') as usize + 1;
        }
    }
    col.saturating_sub(1)
}

/// Parse a usize directly from a byte slice without String allocation.
fn parse_usize_from_bytes(bytes: &[u8]) -> usize {
    let mut n: usize = 0;
    for &b in bytes {
        if b.is_ascii_digit() {
            n = n * 10 + (b - b'0') as usize;
        }
    }
    n
}

/// Parse a u32 directly from a byte slice without String allocation.
fn parse_u32_from_bytes(bytes: &[u8]) -> u32 {
    let mut n: u32 = 0;
    for &b in bytes {
        if b.is_ascii_digit() {
            n = n * 10 + (b - b'0') as u32;
        }
    }
    n
}

/// Parse an f64 from a byte slice, falling back to 0.0 on failure.
fn parse_f64_from_bytes(bytes: &[u8]) -> f64 {
    // Fast path: try to parse directly from UTF-8 bytes
    std::str::from_utf8(bytes)
        .ok()
        .and_then(|s| s.parse().ok())
        .unwrap_or(0.0)
}

/// Parsed worksheet data: rows, merge ranges, column widths, row heights, freeze pane, and auto-filter.
struct WorksheetData {
    rows: Vec<Vec<CellValue>>,
    merges: Vec<String>,
    column_widths: HashMap<u32, f64>,
    row_heights: HashMap<u32, f64>,
    freeze_pane: Option<(u32, u32)>,
    auto_filter: Option<String>,
}

/// Parse a single worksheet XML file and return rows of cell values and merge ranges.
fn parse_worksheet<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    path: &str,
    shared_strings: &[String],
    styles: &[StyleInfo],
) -> Result<WorksheetData, XlsxError> {
    // First pass: quickly scan for <dimension> tag to pre-allocate rows
    let mut estimated_rows: usize = 0;
    {
        let dim_file = archive.by_name(path)?;
        let mut dim_reader = Reader::from_reader(BufReader::new(dim_file));
        let mut dim_buf = Vec::new();
        loop {
            match dim_reader.read_event_into(&mut dim_buf) {
                Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e))
                    if e.name().as_ref() == b"dimension" =>
                {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"ref" {
                            // Parse "A1:J100000" -> extract row number after the colon
                            if let Some(colon_pos) = attr.value.iter().position(|&b| b == b':') {
                                let after_colon = &attr.value[colon_pos + 1..];
                                let row_start = after_colon
                                    .iter()
                                    .position(|b| b.is_ascii_digit())
                                    .unwrap_or(0);
                                estimated_rows = parse_usize_from_bytes(&after_colon[row_start..]);
                            }
                        }
                    }
                    break;
                }
                Ok(Event::Start(ref e)) if e.name().as_ref() == b"sheetData" => break,
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            dim_buf.clear();
        }
    }

    // Main parse
    let file = archive.by_name(path)?;
    let mut reader = Reader::from_reader(BufReader::new(file));
    let mut buf = Vec::new();

    let mut rows: Vec<Vec<CellValue>> = if estimated_rows > 0 {
        Vec::with_capacity(estimated_rows)
    } else {
        Vec::new()
    };
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
    let mut auto_filter: Option<String> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                match e.name().as_ref() {
                    b"pane" => {
                        let mut y_split: u32 = 0;
                        let mut x_split: u32 = 0;
                        let mut is_frozen = false;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"ySplit" => {
                                    y_split = parse_u32_from_bytes(&attr.value);
                                }
                                b"xSplit" => {
                                    x_split = parse_u32_from_bytes(&attr.value);
                                }
                                b"state" => {
                                    is_frozen = attr.value.as_ref() == b"frozen";
                                }
                                _ => {}
                            }
                        }
                        if is_frozen && (y_split > 0 || x_split > 0) {
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
                                    min = parse_u32_from_bytes(&attr.value);
                                }
                                b"max" => {
                                    max = parse_u32_from_bytes(&attr.value);
                                }
                                b"width" => {
                                    width = parse_f64_from_bytes(&attr.value);
                                }
                                b"customWidth" => {
                                    custom_width = attr.value.as_ref() == b"1";
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
                    b"autoFilter" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"ref" {
                                auto_filter =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
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
                                    let row_num = parse_usize_from_bytes(&attr.value);
                                    current_row = if row_num > 0 {
                                        row_num - 1
                                    } else {
                                        current_row + 1
                                    };
                                }
                                b"ht" => {
                                    height = parse_f64_from_bytes(&attr.value);
                                }
                                b"customHeight" => {
                                    custom_height = attr.value.as_ref() == b"1";
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
                                    // Parse column directly from bytes without String allocation
                                    current_col = col_from_cell_ref_bytes(&attr.value);
                                }
                                b"t" => {
                                    // cell_type is only matched against known short ASCII strings,
                                    // write directly from bytes without Cow/String intermediary
                                    cell_type.clear();
                                    if let Ok(s) = std::str::from_utf8(&attr.value) {
                                        cell_type.push_str(s);
                                    }
                                }
                                b"s" => {
                                    // Parse integer directly from bytes without String allocation
                                    cell_style = parse_usize_from_bytes(&attr.value);
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
                            let style_info = styles.get(cell_style);
                            let is_date = style_info.map(|s| s.is_date).unwrap_or(false);
                            let has_cell_style =
                                style_info.map(|s| s.cell_style.is_some()).unwrap_or(false);

                            // Only clone format_code/cell_style when actually needed
                            let resolve_fmt = if has_cell_style {
                                &None
                            } else {
                                // Borrow instead of clone when possible
                                match style_info.and_then(|s| s.format_code.as_ref()) {
                                    Some(_) => &style_info.unwrap().format_code,
                                    None => &None,
                                }
                            };

                            let value = if !cell_formula_text.is_empty() {
                                // Cell has a formula
                                let cached = resolve_cell_value(
                                    &cell_type,
                                    &cell_value_text,
                                    shared_strings,
                                    is_date,
                                    &None,
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
                                    resolve_fmt,
                                )
                            };

                            // Wrap in StyledCell if there's non-default styling
                            // Only clone the style here (rare path — most cells have no style)
                            let final_value = if has_cell_style {
                                let style = style_info.unwrap().cell_style.clone().unwrap();
                                CellValue::StyledCell {
                                    value: Box::new(value),
                                    style: Box::new(style),
                                }
                            } else {
                                value
                            };

                            // Ensure the row has enough columns
                            if let Some(row) = rows.get_mut(current_row) {
                                while row.len() <= current_col {
                                    row.push(CellValue::Empty);
                                }
                                row[current_col] = final_value;
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
            Ok(Event::Empty(ref e)) if e.name().as_ref() == b"autoFilter" => {
                for attr in e.attributes().flatten() {
                    if attr.key.as_ref() == b"ref" {
                        auto_filter = Some(String::from_utf8_lossy(&attr.value).to_string());
                    }
                }
            }
            Ok(Event::Empty(ref e)) if e.name().as_ref() == b"pane" => {
                let mut y_split: u32 = 0;
                let mut x_split: u32 = 0;
                let mut is_frozen = false;
                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"ySplit" => {
                            y_split = parse_u32_from_bytes(&attr.value);
                        }
                        b"xSplit" => {
                            x_split = parse_u32_from_bytes(&attr.value);
                        }
                        b"state" => {
                            is_frozen = attr.value.as_ref() == b"frozen";
                        }
                        _ => {}
                    }
                }
                if is_frozen && (y_split > 0 || x_split > 0) {
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
                            min = parse_u32_from_bytes(&attr.value);
                        }
                        b"max" => {
                            max = parse_u32_from_bytes(&attr.value);
                        }
                        b"width" => {
                            width = parse_f64_from_bytes(&attr.value);
                        }
                        b"customWidth" => {
                            custom_width = attr.value.as_ref() == b"1";
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
        auto_filter,
    })
}

/// Convert raw cell value text into a typed CellValue based on the cell type attribute.
fn resolve_cell_value(
    cell_type: &str,
    raw: &str,
    shared_strings: &[String],
    is_date: bool,
    format_code: &Option<String>,
) -> CellValue {
    if raw.is_empty() && cell_type != "inlineStr" {
        return CellValue::Empty;
    }

    match cell_type {
        // Shared string — store index to avoid cloning; resolved at Python conversion time
        "s" => {
            if let Ok(idx) = raw.parse::<usize>() {
                if idx < shared_strings.len() {
                    CellValue::SharedString(idx)
                } else {
                    CellValue::Empty
                }
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
                } else if let Some(code) = format_code {
                    CellValue::FormattedNumber {
                        value: n,
                        format_code: code.clone(),
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
    fn test_col_from_cell_ref_bytes() {
        assert_eq!(col_from_cell_ref_bytes(b"A1"), 0);
        assert_eq!(col_from_cell_ref_bytes(b"B1"), 1);
        assert_eq!(col_from_cell_ref_bytes(b"Z1"), 25);
        assert_eq!(col_from_cell_ref_bytes(b"AA1"), 26);
        assert_eq!(col_from_cell_ref_bytes(b"AZ1"), 51);
        assert_eq!(col_from_cell_ref_bytes(b"BA1"), 52);
    }

    #[test]
    fn test_parse_usize_from_bytes() {
        assert_eq!(parse_usize_from_bytes(b"0"), 0);
        assert_eq!(parse_usize_from_bytes(b"1"), 1);
        assert_eq!(parse_usize_from_bytes(b"42"), 42);
        assert_eq!(parse_usize_from_bytes(b"100000"), 100000);
    }

    #[test]
    fn test_resolve_cell_value() {
        let shared = vec!["hello".to_string(), "world".to_string()];

        let none_fmt = None;

        match resolve_cell_value("s", "0", &shared, false, &none_fmt) {
            CellValue::SharedString(idx) => assert_eq!(shared[idx], "hello"),
            _ => panic!("expected SharedString"),
        }

        match resolve_cell_value("", "42.5", &shared, false, &none_fmt) {
            CellValue::Number(n) => assert!((n - 42.5).abs() < f64::EPSILON),
            _ => panic!("expected number"),
        }

        match resolve_cell_value("b", "1", &shared, false, &none_fmt) {
            CellValue::Bool(b) => assert!(b),
            _ => panic!("expected bool"),
        }

        assert!(matches!(
            resolve_cell_value("", "", &shared, false, &none_fmt),
            CellValue::Empty
        ));

        // Date detection
        match resolve_cell_value("", "44197", &shared, true, &none_fmt) {
            CellValue::Date { year, month, day } => {
                assert_eq!((year, month, day), (2021, 1, 1));
            }
            other => panic!("expected date, got {other:?}"),
        }

        // Formatted number detection
        let pct_fmt = Some("0.00%".to_string());
        match resolve_cell_value("", "0.75", &shared, false, &pct_fmt) {
            CellValue::FormattedNumber { value, format_code } => {
                assert!((value - 0.75).abs() < f64::EPSILON);
                assert_eq!(format_code, "0.00%");
            }
            other => panic!("expected FormattedNumber, got {other:?}"),
        }
    }
}
