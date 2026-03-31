use pyo3::prelude::*;
use pyo3::types::{
    PyBool, PyDate, PyDateAccess, PyDateTime, PyDict, PyFloat, PyInt, PyList, PyNone, PyString,
    PyTimeAccess,
};
use std::fs::File;
use std::io::{BufReader, BufWriter};

mod reader;
mod types;
mod writer;

use types::CellValue;
use writer::xlsx::StreamingXlsxWriter;

/// Convert a CellValue to a Python object.
///
/// `py_shared_strings` contains pre-converted Python str objects for the shared string table.
/// SharedString(idx) cells look up into this table instead of creating new Python strings.
fn cell_to_py(
    py: Python<'_>,
    cell: CellValue,
    py_shared_strings: &[Py<PyAny>],
) -> PyResult<Py<PyAny>> {
    match cell {
        CellValue::String(s) => Ok(s.into_pyobject(py)?.into_any().unbind()),
        CellValue::SharedString(idx) => {
            // Reuse pre-converted Python string — just increment refcount
            if let Some(py_str) = py_shared_strings.get(idx) {
                Ok(py_str.clone_ref(py))
            } else {
                Ok(PyNone::get(py).to_owned().into_any().unbind())
            }
        }
        CellValue::Number(n) => {
            if n.fract() == 0.0 && n.abs() < i64::MAX as f64 {
                Ok((n as i64).into_pyobject(py)?.into_any().unbind())
            } else {
                Ok(n.into_pyobject(py)?.into_any().unbind())
            }
        }
        CellValue::Bool(b) => Ok(b.into_pyobject(py)?.to_owned().into_any().unbind()),
        CellValue::Formula {
            formula,
            cached_value,
        } => {
            let cached_py = match cached_value {
                Some(v) => Some(cell_to_py(py, *v, py_shared_strings)?),
                None => None,
            };
            let f = Formula {
                formula,
                cached_value: cached_py,
            };
            Ok(f.into_pyobject(py)?.into_any().unbind())
        }
        CellValue::Date { year, month, day } => {
            let date = PyDate::new(py, year, month as u8, day as u8)?;
            Ok(date.into_any().unbind())
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
            let dt = PyDateTime::new(
                py,
                year,
                month as u8,
                day as u8,
                hour as u8,
                minute as u8,
                second as u8,
                microsecond,
                None,
            )?;
            Ok(dt.into_any().unbind())
        }
        CellValue::FormattedNumber { value, format_code } => {
            let py_value: Py<PyAny> = if value.fract() == 0.0 && value.abs() < i64::MAX as f64 {
                (value as i64).into_pyobject(py)?.into_any().unbind()
            } else {
                value.into_pyobject(py)?.into_any().unbind()
            };
            let fc = FormattedCell {
                value: py_value,
                number_format: format_code,
            };
            Ok(fc.into_pyobject(py)?.into_any().unbind())
        }
        CellValue::StyledCell { value, style } => {
            let inner_py = cell_to_py(py, *value, py_shared_strings)?;
            let py_style = cell_style_to_py(py, &style)?;
            let sc = StyledCell {
                value: inner_py,
                style: Py::new(py, py_style)?,
            };
            Ok(sc.into_pyobject(py)?.into_any().unbind())
        }
        CellValue::Empty => Ok(PyNone::get(py).to_owned().into_any().unbind()),
    }
}

/// Convert a Rust CellStyle to a Python CellStyle.
fn cell_style_to_py(_py: Python<'_>, style: &types::CellStyle) -> PyResult<CellStyle> {
    Ok(CellStyle {
        bold: style.bold,
        italic: style.italic,
        underline: style.underline,
        font_name: style.font_name.clone(),
        font_size: style.font_size,
        font_color: style.font_color.clone(),
        fill_color: style.fill_color.clone(),
        border_left: style.border_left.clone(),
        border_right: style.border_right.clone(),
        border_top: style.border_top.clone(),
        border_bottom: style.border_bottom.clone(),
        border_color: style.border_color.clone(),
        horizontal_alignment: style.horizontal_alignment.clone(),
        vertical_alignment: style.vertical_alignment.clone(),
        wrap_text: style.wrap_text,
        text_rotation: style.text_rotation,
        number_format: style.number_format.clone(),
    })
}

/// Pre-convert shared strings to Python str objects for reuse.
///
/// Each SharedString(idx) cell can then cheaply clone_ref the Python object
/// instead of creating a new Python string from scratch.
fn intern_shared_strings(py: Python<'_>, shared_strings: &[String]) -> Vec<Py<PyAny>> {
    shared_strings
        .iter()
        .map(|s| {
            s.into_pyobject(py)
                .expect("string conversion should not fail")
                .into_any()
                .unbind()
        })
        .collect()
}

/// Convert rows to a Python list of lists, consuming the Rust data.
///
/// Takes ownership of rows so each Rust row is freed after conversion,
/// reducing peak memory (Rust + Python data don't fully overlap).
fn rows_to_py(
    py: Python<'_>,
    rows: Vec<Vec<CellValue>>,
    py_shared_strings: &[Py<PyAny>],
) -> PyResult<Py<PyAny>> {
    let mut outer: Vec<Py<PyAny>> = Vec::with_capacity(rows.len());
    for row in rows {
        let mut py_cells: Vec<Py<PyAny>> = Vec::with_capacity(row.len());
        for cell in row {
            py_cells.push(cell_to_py(py, cell, py_shared_strings)?);
        }
        let py_row = PyList::new(py, &py_cells)?;
        outer.push(py_row.into_any().unbind());
    }
    let py_rows = PyList::new(py, &outer)?;
    Ok(py_rows.into_any().unbind())
}

/// Convert a Python value to a CellValue.
///
/// Type checks are ordered by frequency: most common types first to minimize
/// failed extract() calls in the hot path.
fn py_to_cell(obj: &Bound<'_, PyAny>) -> CellValue {
    if obj.is_none() {
        return CellValue::Empty;
    }
    // --- Common types first (Bool must precede Int since bool is a subclass of int in Python) ---
    if let Ok(b) = obj.cast::<PyBool>() {
        return CellValue::Bool(b.is_true());
    }
    if let Ok(i) = obj.cast::<PyInt>() {
        return match i.extract::<i64>() {
            Ok(v) => CellValue::Number(v as f64),
            Err(_) => CellValue::String(i.to_string()),
        };
    }
    if let Ok(f) = obj.cast::<PyFloat>() {
        let v = f.value();
        if v.is_nan() {
            return CellValue::Empty;
        }
        if v.is_infinite() {
            return CellValue::String(if v.is_sign_positive() {
                "Infinity".to_string()
            } else {
                "-Infinity".to_string()
            });
        }
        return CellValue::Number(v);
    }
    if let Ok(s) = obj.cast::<PyString>() {
        return CellValue::String(s.to_string());
    }
    // --- Date types (DateTime must precede Date since datetime is a subclass of date) ---
    if let Ok(dt) = obj.cast::<PyDateTime>() {
        // Use PyDateAccess/PyTimeAccess direct C-level accessors (no Python getattr overhead)
        return CellValue::DateTime {
            year: dt.get_year(),
            month: dt.get_month() as u32,
            day: dt.get_day() as u32,
            hour: dt.get_hour() as u32,
            minute: dt.get_minute() as u32,
            second: dt.get_second() as u32,
            microsecond: dt.get_microsecond(),
        };
    }
    if let Ok(d) = obj.cast::<PyDate>() {
        return CellValue::Date {
            year: d.get_year(),
            month: d.get_month() as u32,
            day: d.get_day() as u32,
        };
    }
    // --- NumPy type support ---
    // numpy 2.x scalars are NOT subclasses of Python int/float/bool.
    // Detect them by checking if the type's module is "numpy".
    if let Ok(module) = obj.get_type().getattr("__module__") {
        if let Ok(mod_str) = module.extract::<String>() {
            if mod_str == "numpy" {
                if let Ok(type_name) = obj.get_type().qualname() {
                    let name = type_name.to_cow().unwrap_or_default();
                    match name.as_ref() {
                        // numpy.bool_ (qualname is "bool" in numpy 2.x)
                        "bool_" | "bool" => {
                            if let Ok(b) = obj.is_truthy() {
                                return CellValue::Bool(b);
                            }
                        }
                        // numpy integer types
                        "int8" | "int16" | "int32" | "int64" | "uint8" | "uint16" | "uint32"
                        | "uint64" | "intp" | "uintp" | "intc" | "long" | "longlong" => {
                            if let Ok(item) = obj.call_method0("item") {
                                if let Ok(v) = item.extract::<i64>() {
                                    return CellValue::Number(v as f64);
                                }
                            }
                        }
                        // numpy float types
                        "float16" | "float32" | "float64" | "float128" | "half" | "single"
                        | "double" => {
                            if let Ok(item) = obj.call_method0("item") {
                                if let Ok(v) = item.extract::<f64>() {
                                    if v.is_nan() {
                                        return CellValue::Empty;
                                    }
                                    if v.is_infinite() {
                                        return CellValue::String(if v.is_sign_positive() {
                                            "Infinity".to_string()
                                        } else {
                                            "-Infinity".to_string()
                                        });
                                    }
                                    return CellValue::Number(v);
                                }
                            }
                        }
                        // numpy.datetime64
                        "datetime64" => {
                            if let Ok(dt_obj) = obj.call_method0("item") {
                                if let Ok(dt) = dt_obj.cast::<PyDateTime>() {
                                    return CellValue::DateTime {
                                        year: dt.get_year(),
                                        month: dt.get_month() as u32,
                                        day: dt.get_day() as u32,
                                        hour: dt.get_hour() as u32,
                                        minute: dt.get_minute() as u32,
                                        second: dt.get_second() as u32,
                                        microsecond: dt.get_microsecond(),
                                    };
                                }
                                if let Ok(d) = dt_obj.cast::<PyDate>() {
                                    return CellValue::Date {
                                        year: d.get_year(),
                                        month: d.get_month() as u32,
                                        day: d.get_day() as u32,
                                    };
                                }
                            }
                        }
                        // numpy string types
                        "str_" | "bytes_" => {
                            return CellValue::String(obj.to_string());
                        }
                        _ => {
                            // Unknown numpy type — try generic conversion via .item()
                            if let Ok(item) = obj.call_method0("item") {
                                return py_to_cell(&item);
                            }
                        }
                    }
                }
            }
        }
    }
    // --- Rare wrapper types last ---
    if let Ok(sc) = obj.extract::<PyRef<'_, StyledCell>>() {
        let py = obj.py();
        let inner = sc.value.bind(py);
        let inner_cell = py_to_cell(inner);
        let style_ref: PyRef<'_, CellStyle> = sc.style.bind(py).extract().unwrap();
        let cell_style = py_cell_style_to_rust(&style_ref);
        return CellValue::StyledCell {
            value: Box::new(inner_cell),
            style: Box::new(cell_style),
        };
    }
    if let Ok(fc) = obj.extract::<PyRef<'_, FormattedCell>>() {
        let py = obj.py();
        let inner = fc.value.bind(py);
        let value = if let Ok(i) = inner.cast::<PyInt>() {
            match i.extract::<i64>() {
                Ok(v) => v as f64,
                Err(_) => 0.0,
            }
        } else if let Ok(f) = inner.cast::<PyFloat>() {
            f.value()
        } else {
            0.0
        };
        return CellValue::FormattedNumber {
            value,
            format_code: fc.number_format.clone(),
        };
    }
    if let Ok(f) = obj.extract::<PyRef<'_, Formula>>() {
        let py = obj.py();
        let cached = f.cached_value.as_ref().and_then(|v| {
            let bound = v.bind(py);
            if bound.is_none() {
                None
            } else {
                Some(Box::new(py_to_cell(bound)))
            }
        });
        return CellValue::Formula {
            formula: f.formula.clone(),
            cached_value: cached,
        };
    }
    CellValue::String(obj.to_string())
}

/// Convert a Python CellStyle to a Rust CellStyle.
fn py_cell_style_to_rust(style: &CellStyle) -> types::CellStyle {
    types::CellStyle {
        bold: style.bold,
        italic: style.italic,
        underline: style.underline,
        font_name: style.font_name.clone(),
        font_size: style.font_size,
        font_color: style.font_color.clone(),
        fill_color: style.fill_color.clone(),
        border_left: style.border_left.clone(),
        border_right: style.border_right.clone(),
        border_top: style.border_top.clone(),
        border_bottom: style.border_bottom.clone(),
        border_color: style.border_color.clone(),
        horizontal_alignment: style.horizontal_alignment.clone(),
        vertical_alignment: style.vertical_alignment.clone(),
        wrap_text: style.wrap_text,
        text_rotation: style.text_rotation,
        number_format: style.number_format.clone(),
    }
}

/// Convert a column letter (e.g. "A", "AA", "BZ") to a 0-based column index.
fn col_letter_to_index(col: &str) -> PyResult<u32> {
    let col = col.trim().to_uppercase();
    if col.is_empty() || !col.bytes().all(|b| b.is_ascii_alphabetic()) {
        return Err(pyo3::exceptions::PyValueError::new_err(format!(
            "Invalid column letter: '{col}'"
        )));
    }
    let mut index: u32 = 0;
    for b in col.bytes() {
        index = index * 26 + (b - b'A') as u32 + 1;
    }
    Ok(index - 1)
}

/// Returns version information about the native core.
#[pyfunction]
fn version() -> &'static str {
    env!("CARGO_PKG_VERSION")
}

/// Read an XLSX file and return a list of sheets.
///
/// Each sheet is a dict with:
///   - "name": sheet name (str)
///   - "rows": list of lists of cell values
///   - "merges": list of merged cell range strings (e.g. ["A1:B2"])
#[pyfunction]
fn read_xlsx(py: Python<'_>, path: &str) -> PyResult<Py<PyAny>> {
    let file = File::open(path)
        .map_err(|e| pyo3::exceptions::PyFileNotFoundError::new_err(format!("{path}: {e}")))?;
    let reader = BufReader::new(file);
    let (sheets, shared_strings, _defined_names) = reader::xlsx::read_xlsx(reader)?;

    // Pre-convert shared strings to Python objects for reuse
    let py_shared = intern_shared_strings(py, &shared_strings);
    // Drop the Rust shared strings — they've been converted to Python
    drop(shared_strings);

    let result = PyList::empty(py);
    for sheet in sheets {
        let dict = PyDict::new(py);
        dict.set_item("name", &sheet.name)?;
        dict.set_item("rows", rows_to_py(py, sheet.rows, &py_shared)?)?;
        let merges_list = PyList::new(py, &sheet.merges)?;
        dict.set_item("merges", merges_list)?;

        // Column widths: {0-based col index -> width}
        let col_widths_dict = PyDict::new(py);
        for (col_idx, width) in &sheet.column_widths {
            col_widths_dict.set_item(col_idx, width)?;
        }
        dict.set_item("column_widths", col_widths_dict)?;

        // Row heights: {0-based row index -> height}
        let row_heights_dict = PyDict::new(py);
        for (row_idx, height) in &sheet.row_heights {
            row_heights_dict.set_item(row_idx, height)?;
        }
        dict.set_item("row_heights", row_heights_dict)?;

        // Freeze pane: (rows_frozen, cols_frozen) or None
        match sheet.freeze_pane {
            Some((row, col)) => {
                let tuple = (row, col);
                dict.set_item("freeze_pane", tuple)?;
            }
            None => {
                dict.set_item("freeze_pane", py.None())?;
            }
        }

        // Auto-filter range or None
        match &sheet.auto_filter {
            Some(range) => dict.set_item("auto_filter", range)?,
            None => dict.set_item("auto_filter", py.None())?,
        }

        // Sheet visibility state
        dict.set_item("state", &sheet.state)?;

        // Data validations
        let dv_list = PyList::empty(py);
        for dv in &sheet.data_validations {
            let dv_dict = PyDict::new(py);
            dv_dict.set_item("type", &dv.validation_type)?;
            dv_dict.set_item("sqref", &dv.sqref)?;
            match &dv.operator {
                Some(op) => dv_dict.set_item("operator", op)?,
                None => dv_dict.set_item("operator", py.None())?,
            }
            match &dv.formula1 {
                Some(f) => dv_dict.set_item("formula1", f)?,
                None => dv_dict.set_item("formula1", py.None())?,
            }
            match &dv.formula2 {
                Some(f) => dv_dict.set_item("formula2", f)?,
                None => dv_dict.set_item("formula2", py.None())?,
            }
            dv_dict.set_item("allow_blank", dv.allow_blank)?;
            dv_dict.set_item("show_input_message", dv.show_input_message)?;
            dv_dict.set_item("show_error_message", dv.show_error_message)?;
            match &dv.prompt_title {
                Some(t) => dv_dict.set_item("prompt_title", t)?,
                None => dv_dict.set_item("prompt_title", py.None())?,
            }
            match &dv.prompt {
                Some(p) => dv_dict.set_item("prompt", p)?,
                None => dv_dict.set_item("prompt", py.None())?,
            }
            match &dv.error_title {
                Some(t) => dv_dict.set_item("error_title", t)?,
                None => dv_dict.set_item("error_title", py.None())?,
            }
            match &dv.error_message {
                Some(m) => dv_dict.set_item("error_message", m)?,
                None => dv_dict.set_item("error_message", py.None())?,
            }
            match &dv.error_style {
                Some(s) => dv_dict.set_item("error_style", s)?,
                None => dv_dict.set_item("error_style", py.None())?,
            }
            dv_list.append(dv_dict)?;
        }
        dict.set_item("data_validations", dv_list)?;

        // Comments
        let comments_list = PyList::empty(py);
        for c in &sheet.comments {
            let c_dict = PyDict::new(py);
            c_dict.set_item("cell", &c.cell)?;
            c_dict.set_item("author", &c.author)?;
            c_dict.set_item("text", &c.text)?;
            comments_list.append(c_dict)?;
        }
        dict.set_item("comments", comments_list)?;

        // Hyperlinks
        let hyperlinks_list = PyList::empty(py);
        for h in &sheet.hyperlinks {
            let h_dict = PyDict::new(py);
            h_dict.set_item("cell", &h.cell)?;
            h_dict.set_item("url", &h.url)?;
            match &h.tooltip {
                Some(t) => h_dict.set_item("tooltip", t)?,
                None => h_dict.set_item("tooltip", py.None())?,
            }
            hyperlinks_list.append(h_dict)?;
        }
        dict.set_item("hyperlinks", hyperlinks_list)?;

        // Protection
        match &sheet.protection {
            Some(prot) => {
                let p_dict = PyDict::new(py);
                p_dict.set_item("sheet", prot.sheet)?;
                p_dict.set_item("objects", prot.objects)?;
                p_dict.set_item("scenarios", prot.scenarios)?;
                match &prot.password_hash {
                    Some(h) => p_dict.set_item("password_hash", h)?,
                    None => p_dict.set_item("password_hash", py.None())?,
                }
                p_dict.set_item("format_cells", prot.format_cells)?;
                p_dict.set_item("format_columns", prot.format_columns)?;
                p_dict.set_item("format_rows", prot.format_rows)?;
                p_dict.set_item("insert_columns", prot.insert_columns)?;
                p_dict.set_item("insert_rows", prot.insert_rows)?;
                p_dict.set_item("insert_hyperlinks", prot.insert_hyperlinks)?;
                p_dict.set_item("delete_columns", prot.delete_columns)?;
                p_dict.set_item("delete_rows", prot.delete_rows)?;
                p_dict.set_item("sort", prot.sort)?;
                p_dict.set_item("auto_filter", prot.auto_filter)?;
                p_dict.set_item("pivot_tables", prot.pivot_tables)?;
                p_dict.set_item("select_locked_cells", prot.select_locked_cells)?;
                p_dict.set_item("select_unlocked_cells", prot.select_unlocked_cells)?;
                dict.set_item("protection", p_dict)?;
            }
            None => {
                dict.set_item("protection", py.None())?;
            }
        }

        // Tables
        let tables_list = PyList::empty(py);
        for t in &sheet.tables {
            let t_dict = PyDict::new(py);
            t_dict.set_item("name", &t.name)?;
            t_dict.set_item("display_name", &t.display_name)?;
            t_dict.set_item("ref", &t.reference)?;
            let cols_list = PyList::empty(py);
            for col in &t.columns {
                let col_dict = PyDict::new(py);
                col_dict.set_item("id", col.id)?;
                col_dict.set_item("name", &col.name)?;
                cols_list.append(col_dict)?;
            }
            t_dict.set_item("columns", cols_list)?;
            match &t.style {
                Some(s) => t_dict.set_item("style", s)?,
                None => t_dict.set_item("style", py.None())?,
            }
            t_dict.set_item("has_auto_filter", t.has_auto_filter)?;
            tables_list.append(t_dict)?;
        }
        dict.set_item("tables", tables_list)?;

        result.append(dict)?;
    }

    Ok(result.into_any().unbind())
}

/// Read defined names (named ranges) from an XLSX file.
///
/// Returns a list of dicts with:
///   - "name": the defined name (str)
///   - "value": the reference/formula (str)
///   - "sheet_index": 0-based sheet index if sheet-scoped, or None if workbook-scoped
#[pyfunction]
fn defined_names(py: Python<'_>, path: &str) -> PyResult<Py<PyAny>> {
    let file = File::open(path)
        .map_err(|e| pyo3::exceptions::PyFileNotFoundError::new_err(format!("{path}: {e}")))?;
    let reader = BufReader::new(file);
    let names = reader::xlsx::read_defined_names(reader)?;

    let result = PyList::empty(py);
    for dn in names {
        let dict = PyDict::new(py);
        dict.set_item("name", &dn.name)?;
        dict.set_item("value", &dn.value)?;
        match dn.sheet_index {
            Some(idx) => dict.set_item("sheet_index", idx)?,
            None => dict.set_item("sheet_index", py.None())?,
        }
        result.append(dict)?;
    }

    Ok(result.into_any().unbind())
}

/// Read document properties from an XLSX file.
///
/// Returns a dict with:
///   - "core": dict of core properties (title, subject, creator, etc.)
///   - "custom": list of dicts with "name" and "value" keys
#[pyfunction]
fn document_properties(py: Python<'_>, path: &str) -> PyResult<Py<PyAny>> {
    let file = File::open(path)
        .map_err(|e| pyo3::exceptions::PyFileNotFoundError::new_err(format!("{path}: {e}")))?;
    let reader = BufReader::new(file);
    let (core, custom) = reader::xlsx::read_document_properties(reader)?;

    let result = PyDict::new(py);

    let core_dict = PyDict::new(py);
    if let Some(ref v) = core.title {
        core_dict.set_item("title", v)?;
    }
    if let Some(ref v) = core.subject {
        core_dict.set_item("subject", v)?;
    }
    if let Some(ref v) = core.creator {
        core_dict.set_item("creator", v)?;
    }
    if let Some(ref v) = core.keywords {
        core_dict.set_item("keywords", v)?;
    }
    if let Some(ref v) = core.description {
        core_dict.set_item("description", v)?;
    }
    if let Some(ref v) = core.last_modified_by {
        core_dict.set_item("last_modified_by", v)?;
    }
    if let Some(ref v) = core.category {
        core_dict.set_item("category", v)?;
    }
    if let Some(ref v) = core.created {
        core_dict.set_item("created", v)?;
    }
    if let Some(ref v) = core.modified {
        core_dict.set_item("modified", v)?;
    }
    result.set_item("core", core_dict)?;

    let custom_list = PyList::empty(py);
    for prop in custom {
        let d = PyDict::new(py);
        d.set_item("name", &prop.name)?;
        d.set_item("value", &prop.value)?;
        custom_list.append(d)?;
    }
    result.set_item("custom", custom_list)?;

    Ok(result.into_any().unbind())
}

/// Read a specific sheet by name or index from an XLSX file.
///
/// Returns a list of rows (list of lists of cell values).
/// Only parses the requested sheet, skipping others for efficiency.
#[pyfunction]
#[pyo3(signature = (path, sheet_name=None, sheet_index=None))]
fn read_sheet(
    py: Python<'_>,
    path: &str,
    sheet_name: Option<&str>,
    sheet_index: Option<usize>,
) -> PyResult<Py<PyAny>> {
    let file = File::open(path)
        .map_err(|e| pyo3::exceptions::PyFileNotFoundError::new_err(format!("{path}: {e}")))?;
    let reader = BufReader::new(file);
    let (sheet, shared_strings) = reader::xlsx::read_single_sheet(reader, sheet_name, sheet_index)?;

    let py_shared = intern_shared_strings(py, &shared_strings);
    drop(shared_strings);

    rows_to_py(py, sheet.rows, &py_shared)
}

/// List sheet names in an XLSX file.
#[pyfunction]
fn sheet_names(path: &str) -> PyResult<Vec<String>> {
    let file = File::open(path)
        .map_err(|e| pyo3::exceptions::PyFileNotFoundError::new_err(format!("{path}: {e}")))?;
    let reader = BufReader::new(file);
    let names = reader::xlsx::read_sheet_names(reader)?;
    Ok(names)
}

/// A spreadsheet formula with optional cached value.
///
/// Usage:
///     from opensheet_core import Formula
///     f = Formula("SUM(A1:A10)")
///     f = Formula("A1+B1", cached_value=42)
#[pyclass(skip_from_py_object)]
struct Formula {
    #[pyo3(get, set)]
    formula: String,
    #[pyo3(get, set)]
    cached_value: Option<Py<PyAny>>,
}

#[pymethods]
impl Formula {
    #[new]
    #[pyo3(signature = (formula, cached_value=None))]
    fn new(formula: String, cached_value: Option<Py<PyAny>>) -> Self {
        Formula {
            formula,
            cached_value,
        }
    }

    fn __repr__(&self) -> String {
        match &self.cached_value {
            Some(_) => format!("Formula('{}', cached_value=...)", self.formula),
            None => format!("Formula('{}')", self.formula),
        }
    }

    fn __eq__(&self, other: &Formula) -> PyResult<bool> {
        if self.formula != other.formula {
            return Ok(false);
        }
        match (&self.cached_value, &other.cached_value) {
            (None, None) => Ok(true),
            (Some(a), Some(b)) => {
                Python::try_attach(|py| a.bind(py).eq(b.bind(py))).unwrap_or(Ok(false))
            }
            _ => Ok(false),
        }
    }
}

/// A cell value with a custom number format.
///
/// Usage:
///     from opensheet_core import FormattedCell
///     cell = FormattedCell(1234.56, "$#,##0.00")
///     cell = FormattedCell(0.75, "0%")
#[pyclass]
struct FormattedCell {
    #[pyo3(get, set)]
    value: Py<PyAny>,
    #[pyo3(get, set)]
    number_format: String,
}

#[pymethods]
impl FormattedCell {
    #[new]
    fn new(value: Py<PyAny>, number_format: String) -> Self {
        FormattedCell {
            value,
            number_format,
        }
    }

    fn __repr__(&self) -> String {
        format!("FormattedCell(..., '{}')", self.number_format)
    }

    fn __eq__(&self, other: &FormattedCell) -> PyResult<bool> {
        if self.number_format != other.number_format {
            return Ok(false);
        }
        Python::try_attach(|py| self.value.bind(py).eq(other.value.bind(py))).unwrap_or(Ok(false))
    }
}

/// Cell style properties for fonts, fills, borders, and alignment.
///
/// Usage:
///     from opensheet_core import CellStyle
///     style = CellStyle(bold=True, fill_color="FFFF00")
///     style = CellStyle(border="thin", border_color="000000", horizontal_alignment="center")
#[pyclass(skip_from_py_object)]
#[derive(Clone)]
struct CellStyle {
    #[pyo3(get, set)]
    bold: bool,
    #[pyo3(get, set)]
    italic: bool,
    #[pyo3(get, set)]
    underline: bool,
    #[pyo3(get, set)]
    font_name: Option<String>,
    #[pyo3(get, set)]
    font_size: Option<f64>,
    #[pyo3(get, set)]
    font_color: Option<String>,
    #[pyo3(get, set)]
    fill_color: Option<String>,
    #[pyo3(get, set)]
    border_left: Option<String>,
    #[pyo3(get, set)]
    border_right: Option<String>,
    #[pyo3(get, set)]
    border_top: Option<String>,
    #[pyo3(get, set)]
    border_bottom: Option<String>,
    #[pyo3(get, set)]
    border_color: Option<String>,
    #[pyo3(get, set)]
    horizontal_alignment: Option<String>,
    #[pyo3(get, set)]
    vertical_alignment: Option<String>,
    #[pyo3(get, set)]
    wrap_text: bool,
    #[pyo3(get, set)]
    text_rotation: Option<u16>,
    #[pyo3(get, set)]
    number_format: Option<String>,
}

#[pymethods]
impl CellStyle {
    #[new]
    #[pyo3(signature = (
        *,
        bold = false,
        italic = false,
        underline = false,
        font_name = None,
        font_size = None,
        font_color = None,
        fill_color = None,
        border = None,
        border_left = None,
        border_right = None,
        border_top = None,
        border_bottom = None,
        border_color = None,
        horizontal_alignment = None,
        vertical_alignment = None,
        wrap_text = false,
        text_rotation = None,
        number_format = None,
    ))]
    #[allow(clippy::too_many_arguments)]
    fn new(
        bold: bool,
        italic: bool,
        underline: bool,
        font_name: Option<String>,
        font_size: Option<f64>,
        font_color: Option<String>,
        fill_color: Option<String>,
        border: Option<String>,
        border_left: Option<String>,
        border_right: Option<String>,
        border_top: Option<String>,
        border_bottom: Option<String>,
        border_color: Option<String>,
        horizontal_alignment: Option<String>,
        vertical_alignment: Option<String>,
        wrap_text: bool,
        text_rotation: Option<u16>,
        number_format: Option<String>,
    ) -> Self {
        // If `border` shorthand is set, apply it to any unset sides
        let border_left = border_left.or_else(|| border.clone());
        let border_right = border_right.or_else(|| border.clone());
        let border_top = border_top.or_else(|| border.clone());
        let border_bottom = border_bottom.or(border);

        CellStyle {
            bold,
            italic,
            underline,
            font_name,
            font_size,
            font_color,
            fill_color,
            border_left,
            border_right,
            border_top,
            border_bottom,
            border_color,
            horizontal_alignment,
            vertical_alignment,
            wrap_text,
            text_rotation,
            number_format,
        }
    }

    fn __repr__(&self) -> String {
        let mut parts = Vec::new();
        if self.bold {
            parts.push("bold=True".to_string());
        }
        if self.italic {
            parts.push("italic=True".to_string());
        }
        if self.underline {
            parts.push("underline=True".to_string());
        }
        if let Some(ref name) = self.font_name {
            parts.push(format!("font_name='{name}'"));
        }
        if let Some(size) = self.font_size {
            parts.push(format!("font_size={size}"));
        }
        if let Some(ref color) = self.font_color {
            parts.push(format!("font_color='{color}'"));
        }
        if let Some(ref color) = self.fill_color {
            parts.push(format!("fill_color='{color}'"));
        }
        if self.border_left.is_some()
            || self.border_right.is_some()
            || self.border_top.is_some()
            || self.border_bottom.is_some()
        {
            parts.push("border=...".to_string());
        }
        if let Some(ref h) = self.horizontal_alignment {
            parts.push(format!("horizontal_alignment='{h}'"));
        }
        if let Some(ref v) = self.vertical_alignment {
            parts.push(format!("vertical_alignment='{v}'"));
        }
        if self.wrap_text {
            parts.push("wrap_text=True".to_string());
        }
        if let Some(rot) = self.text_rotation {
            parts.push(format!("text_rotation={rot}"));
        }
        if let Some(ref fmt) = self.number_format {
            parts.push(format!("number_format='{fmt}'"));
        }
        format!("CellStyle({})", parts.join(", "))
    }

    fn __eq__(&self, other: &CellStyle) -> bool {
        self.bold == other.bold
            && self.italic == other.italic
            && self.underline == other.underline
            && self.font_name == other.font_name
            && self.font_size == other.font_size
            && self.font_color == other.font_color
            && self.fill_color == other.fill_color
            && self.border_left == other.border_left
            && self.border_right == other.border_right
            && self.border_top == other.border_top
            && self.border_bottom == other.border_bottom
            && self.border_color == other.border_color
            && self.horizontal_alignment == other.horizontal_alignment
            && self.vertical_alignment == other.vertical_alignment
            && self.wrap_text == other.wrap_text
            && self.text_rotation == other.text_rotation
            && self.number_format == other.number_format
    }
}

/// A cell value with styling (fonts, fills, borders, alignment).
///
/// Usage:
///     from opensheet_core import StyledCell, CellStyle
///     cell = StyledCell("Hello", CellStyle(bold=True))
///     cell = StyledCell(42, CellStyle(fill_color="FFFF00"))
#[pyclass]
struct StyledCell {
    #[pyo3(get, set)]
    value: Py<PyAny>,
    #[pyo3(get, set)]
    style: Py<CellStyle>,
}

#[pymethods]
impl StyledCell {
    #[new]
    fn new(value: Py<PyAny>, style: Py<CellStyle>) -> Self {
        StyledCell { value, style }
    }

    fn __repr__(&self) -> String {
        Python::try_attach(|py| {
            let style_repr = self.style.bind(py).call_method0("__repr__")?;
            Ok::<String, PyErr>(format!(
                "StyledCell(..., {})",
                style_repr.extract::<String>()?
            ))
        })
        .unwrap_or(Ok("StyledCell(...)".to_string()))
        .unwrap_or("StyledCell(...)".to_string())
    }

    fn __eq__(&self, other: &StyledCell) -> PyResult<bool> {
        let values_eq = Python::try_attach(|py| self.value.bind(py).eq(other.value.bind(py)))
            .unwrap_or(Ok(false))?;
        if !values_eq {
            return Ok(false);
        }
        Python::try_attach(|py| {
            let s1: PyRef<'_, CellStyle> = self.style.bind(py).extract()?;
            let s2: PyRef<'_, CellStyle> = other.style.bind(py).extract()?;
            Ok(s1.__eq__(&s2))
        })
        .unwrap_or(Ok(false))
    }
}

/// Streaming XLSX writer.
///
/// Usage:
///     writer = XlsxWriter("output.xlsx")
///     writer.add_sheet("Sheet1")
///     writer.write_row(["Name", "Age", "Score"])
///     writer.write_row(["Alice", 30, 95.5])
///     writer.close()
#[pyclass]
struct XlsxWriter {
    inner: Option<StreamingXlsxWriter<BufWriter<File>>>,
}

#[pymethods]
impl XlsxWriter {
    #[new]
    fn new(path: &str) -> PyResult<Self> {
        let file = File::create(path)
            .map_err(|e| pyo3::exceptions::PyIOError::new_err(format!("{path}: {e}")))?;
        let writer = BufWriter::new(file);
        Ok(XlsxWriter {
            inner: Some(StreamingXlsxWriter::new(writer)),
        })
    }

    /// Add a new sheet to the workbook.
    fn add_sheet(&mut self, name: &str) -> PyResult<()> {
        let w = self
            .inner
            .as_mut()
            .ok_or_else(|| pyo3::exceptions::PyRuntimeError::new_err("Writer is already closed"))?;
        w.add_sheet(name)?;
        Ok(())
    }

    /// Merge a range of cells (e.g. "A1:B2").
    fn merge_cells(&mut self, range: &str) -> PyResult<()> {
        let w = self
            .inner
            .as_mut()
            .ok_or_else(|| pyo3::exceptions::PyRuntimeError::new_err("Writer is already closed"))?;
        w.merge_cells(range)?;
        Ok(())
    }

    /// Set an auto-filter on a range (e.g. "A1:C1").
    fn auto_filter(&mut self, range: &str) -> PyResult<()> {
        let w = self
            .inner
            .as_mut()
            .ok_or_else(|| pyo3::exceptions::PyRuntimeError::new_err("Writer is already closed"))?;
        w.auto_filter(range)?;
        Ok(())
    }

    /// Define a named range (defined name) for the workbook.
    ///
    /// - `name`: The defined name (e.g. "TaxRate").
    /// - `value`: The reference (e.g. "Sheet1!$B$2").
    /// - `sheet_index`: If provided, the name is scoped to that sheet (0-based).
    #[pyo3(signature = (name, value, sheet_index=None))]
    fn define_name(&mut self, name: &str, value: &str, sheet_index: Option<usize>) -> PyResult<()> {
        let w = self
            .inner
            .as_mut()
            .ok_or_else(|| pyo3::exceptions::PyRuntimeError::new_err("Writer is already closed"))?;
        w.define_name(name, value, sheet_index)?;
        Ok(())
    }

    /// Set the visibility state of the current sheet.
    ///
    /// Valid states: "visible" (default), "hidden", "veryHidden".
    fn set_sheet_state(&mut self, state: &str) -> PyResult<()> {
        let w = self
            .inner
            .as_mut()
            .ok_or_else(|| pyo3::exceptions::PyRuntimeError::new_err("Writer is already closed"))?;
        w.set_sheet_state(state)?;
        Ok(())
    }

    /// Set a core document property (title, subject, creator, etc.).
    fn set_document_property(&mut self, key: &str, value: &str) -> PyResult<()> {
        let w = self
            .inner
            .as_mut()
            .ok_or_else(|| pyo3::exceptions::PyRuntimeError::new_err("Writer is already closed"))?;
        w.set_document_property(key, value)?;
        Ok(())
    }

    /// Set a custom document property (arbitrary key-value pair).
    fn set_custom_property(&mut self, name: &str, value: &str) -> PyResult<()> {
        let w = self
            .inner
            .as_mut()
            .ok_or_else(|| pyo3::exceptions::PyRuntimeError::new_err("Writer is already closed"))?;
        w.set_custom_property(name, value)?;
        Ok(())
    }

    /// Add a data validation rule to the current sheet.
    #[pyo3(signature = (
        validation_type,
        sqref,
        formula1=None,
        formula2=None,
        operator=None,
        allow_blank=false,
        show_input_message=false,
        show_error_message=false,
        prompt_title=None,
        prompt=None,
        error_title=None,
        error_message=None,
        error_style=None,
    ))]
    #[allow(clippy::too_many_arguments)]
    fn add_data_validation(
        &mut self,
        validation_type: &str,
        sqref: &str,
        formula1: Option<&str>,
        formula2: Option<&str>,
        operator: Option<&str>,
        allow_blank: bool,
        show_input_message: bool,
        show_error_message: bool,
        prompt_title: Option<&str>,
        prompt: Option<&str>,
        error_title: Option<&str>,
        error_message: Option<&str>,
        error_style: Option<&str>,
    ) -> PyResult<()> {
        let w = self
            .inner
            .as_mut()
            .ok_or_else(|| pyo3::exceptions::PyRuntimeError::new_err("Writer is already closed"))?;
        w.add_data_validation(
            validation_type,
            sqref,
            formula1,
            formula2,
            operator,
            allow_blank,
            show_input_message,
            show_error_message,
            prompt_title,
            prompt,
            error_title,
            error_message,
            error_style,
        )?;
        Ok(())
    }

    /// Add a comment to a cell in the current sheet.
    fn add_comment(&mut self, cell_ref: &str, author: &str, text: &str) -> PyResult<()> {
        let w = self
            .inner
            .as_mut()
            .ok_or_else(|| pyo3::exceptions::PyRuntimeError::new_err("Writer is already closed"))?;
        w.add_comment(cell_ref, author, text)?;
        Ok(())
    }

    /// Add a hyperlink to a cell in the current sheet.
    #[pyo3(signature = (cell_ref, url, tooltip=None))]
    fn add_hyperlink(&mut self, cell_ref: &str, url: &str, tooltip: Option<&str>) -> PyResult<()> {
        let w = self
            .inner
            .as_mut()
            .ok_or_else(|| pyo3::exceptions::PyRuntimeError::new_err("Writer is already closed"))?;
        w.add_hyperlink(cell_ref, url, tooltip)?;
        Ok(())
    }

    /// Protect the current sheet with optional password and configurable options.
    #[pyo3(signature = (
        password=None,
        sheet=true,
        objects=true,
        scenarios=true,
        format_cells=false,
        format_columns=false,
        format_rows=false,
        insert_columns=false,
        insert_rows=false,
        insert_hyperlinks=false,
        delete_columns=false,
        delete_rows=false,
        sort=false,
        auto_filter=false,
        pivot_tables=false,
        select_locked_cells=false,
        select_unlocked_cells=false,
    ))]
    #[allow(clippy::too_many_arguments)]
    fn protect_sheet(
        &mut self,
        password: Option<&str>,
        sheet: bool,
        objects: bool,
        scenarios: bool,
        format_cells: bool,
        format_columns: bool,
        format_rows: bool,
        insert_columns: bool,
        insert_rows: bool,
        insert_hyperlinks: bool,
        delete_columns: bool,
        delete_rows: bool,
        sort: bool,
        auto_filter: bool,
        pivot_tables: bool,
        select_locked_cells: bool,
        select_unlocked_cells: bool,
    ) -> PyResult<()> {
        let w = self
            .inner
            .as_mut()
            .ok_or_else(|| pyo3::exceptions::PyRuntimeError::new_err("Writer is already closed"))?;
        w.protect_sheet(
            password,
            sheet,
            objects,
            scenarios,
            format_cells,
            format_columns,
            format_rows,
            insert_columns,
            insert_rows,
            insert_hyperlinks,
            delete_columns,
            delete_rows,
            sort,
            auto_filter,
            pivot_tables,
            select_locked_cells,
            select_unlocked_cells,
        )?;
        Ok(())
    }

    /// Add a structured table to the current sheet.
    #[pyo3(signature = (reference, columns, name=None, style=None))]
    fn add_table(
        &mut self,
        reference: &str,
        columns: Vec<String>,
        name: Option<&str>,
        style: Option<&str>,
    ) -> PyResult<()> {
        let w = self
            .inner
            .as_mut()
            .ok_or_else(|| pyo3::exceptions::PyRuntimeError::new_err("Writer is already closed"))?;
        w.add_table(reference, &columns, name, style)?;
        Ok(())
    }

    /// Freeze the top `row` rows and left `col` columns.
    ///
    /// Must be called after add_sheet() but before any write_row() calls on that sheet.
    #[pyo3(signature = (row=0, col=0))]
    fn freeze_panes(&mut self, row: u32, col: u32) -> PyResult<()> {
        let w = self
            .inner
            .as_mut()
            .ok_or_else(|| pyo3::exceptions::PyRuntimeError::new_err("Writer is already closed"))?;
        w.freeze_panes(row, col)?;
        Ok(())
    }

    /// Set the width of a column in character units.
    ///
    /// `column` can be a letter (e.g. "A", "AA") or a 0-based integer index.
    /// Must be called after add_sheet() but before any write_row() calls on that sheet.
    #[pyo3(signature = (column, width))]
    fn set_column_width(&mut self, column: &Bound<'_, PyAny>, width: f64) -> PyResult<()> {
        let w = self
            .inner
            .as_mut()
            .ok_or_else(|| pyo3::exceptions::PyRuntimeError::new_err("Writer is already closed"))?;

        let col_index = if let Ok(s) = column.extract::<String>() {
            col_letter_to_index(&s)?
        } else if let Ok(i) = column.extract::<u32>() {
            i
        } else {
            return Err(pyo3::exceptions::PyTypeError::new_err(
                "column must be a string (e.g. 'A') or an integer index",
            ));
        };

        w.set_column_width(col_index, width)?;
        Ok(())
    }

    /// Set the height of a row in points.
    ///
    /// `row` is a 1-based row number (matching Excel convention).
    fn set_row_height(&mut self, row: u32, height: f64) -> PyResult<()> {
        let w = self
            .inner
            .as_mut()
            .ok_or_else(|| pyo3::exceptions::PyRuntimeError::new_err("Writer is already closed"))?;
        if row == 0 {
            return Err(pyo3::exceptions::PyValueError::new_err(
                "Row number must be 1-based (1, 2, 3, ...)",
            ));
        }
        w.set_row_height(row - 1, height)?;
        Ok(())
    }

    /// Write a row of values to the current sheet.
    ///
    /// Values can be: str, int, float, bool, or None.
    fn write_row(&mut self, row: &Bound<'_, PyList>) -> PyResult<()> {
        let w = self
            .inner
            .as_mut()
            .ok_or_else(|| pyo3::exceptions::PyRuntimeError::new_err("Writer is already closed"))?;

        let cells: Vec<CellValue> = row.iter().map(|item| py_to_cell(&item)).collect();
        w.write_row(&cells)?;
        Ok(())
    }

    /// Write multiple rows at once, minimizing Python→Rust FFI crossings.
    ///
    /// Each element of `rows` should be a list of cell values.
    fn write_rows(&mut self, rows: &Bound<'_, PyList>) -> PyResult<()> {
        let w = self
            .inner
            .as_mut()
            .ok_or_else(|| pyo3::exceptions::PyRuntimeError::new_err("Writer is already closed"))?;

        for row_obj in rows.iter() {
            let row_list = row_obj.cast::<PyList>().map_err(|_| {
                pyo3::exceptions::PyTypeError::new_err("Each element of rows must be a list")
            })?;
            let cells: Vec<CellValue> = row_list.iter().map(|item| py_to_cell(&item)).collect();
            w.write_row(&cells)?;
        }
        Ok(())
    }

    /// Close the writer and finalize the XLSX file.
    fn close(&mut self) -> PyResult<()> {
        let w = self
            .inner
            .take()
            .ok_or_else(|| pyo3::exceptions::PyRuntimeError::new_err("Writer is already closed"))?;
        w.close()?;
        Ok(())
    }

    fn __enter__(slf: Py<Self>) -> Py<Self> {
        slf
    }

    fn __exit__(
        &mut self,
        _exc_type: Option<&Bound<'_, PyAny>>,
        _exc_val: Option<&Bound<'_, PyAny>>,
        _exc_tb: Option<&Bound<'_, PyAny>>,
    ) -> PyResult<bool> {
        if self.inner.is_some() {
            self.close()?;
        }
        Ok(false)
    }
}

/// A Python module implemented in Rust.
#[pymodule]
fn _native(m: &Bound<'_, PyModule>) -> PyResult<()> {
    // Initialize Python's datetime C API so that cast::<PyDate>/cast::<PyDateTime>
    // work correctly in py_to_cell (needed before first use of PyDate_Check etc.)
    unsafe { pyo3::ffi::PyDateTime_IMPORT() };

    m.add_function(wrap_pyfunction!(version, m)?)?;
    m.add_function(wrap_pyfunction!(read_xlsx, m)?)?;
    m.add_function(wrap_pyfunction!(read_sheet, m)?)?;
    m.add_function(wrap_pyfunction!(sheet_names, m)?)?;
    m.add_function(wrap_pyfunction!(defined_names, m)?)?;
    m.add_function(wrap_pyfunction!(document_properties, m)?)?;
    m.add_class::<XlsxWriter>()?;
    m.add_class::<Formula>()?;
    m.add_class::<FormattedCell>()?;
    m.add_class::<CellStyle>()?;
    m.add_class::<StyledCell>()?;
    Ok(())
}
