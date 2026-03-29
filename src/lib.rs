use pyo3::prelude::*;
use pyo3::types::{PyBool, PyDict, PyFloat, PyInt, PyList, PyNone, PyString};
use std::fs::File;
use std::io::{BufReader, BufWriter};

mod reader;
mod types;
mod writer;

use types::CellValue;
use writer::xlsx::StreamingXlsxWriter;

/// Convert a CellValue to a Python object.
fn cell_to_py(py: Python<'_>, cell: &CellValue) -> PyResult<Py<PyAny>> {
    match cell {
        CellValue::String(s) => Ok(s.into_pyobject(py)?.into_any().unbind()),
        CellValue::Number(n) => {
            if n.fract() == 0.0 && n.abs() < i64::MAX as f64 {
                Ok((*n as i64).into_pyobject(py)?.into_any().unbind())
            } else {
                Ok(n.into_pyobject(py)?.into_any().unbind())
            }
        }
        CellValue::Bool(b) => Ok(b.into_pyobject(py)?.to_owned().into_any().unbind()),
        CellValue::Empty => Ok(PyNone::get(py).to_owned().into_any().unbind()),
    }
}

/// Convert rows to a Python list of lists.
fn rows_to_py(py: Python<'_>, rows: &[Vec<CellValue>]) -> PyResult<Py<PyAny>> {
    let py_rows = PyList::empty(py);
    for row in rows {
        let py_row = PyList::empty(py);
        for cell in row {
            py_row.append(cell_to_py(py, cell)?)?;
        }
        py_rows.append(py_row)?;
    }
    Ok(py_rows.into_any().unbind())
}

/// Convert a Python value to a CellValue.
fn py_to_cell(obj: &Bound<'_, PyAny>) -> CellValue {
    if obj.is_none() {
        CellValue::Empty
    } else if let Ok(b) = obj.cast::<PyBool>() {
        CellValue::Bool(b.is_true())
    } else if let Ok(i) = obj.cast::<PyInt>() {
        match i.extract::<i64>() {
            Ok(v) => CellValue::Number(v as f64),
            Err(_) => CellValue::String(i.to_string()),
        }
    } else if let Ok(f) = obj.cast::<PyFloat>() {
        CellValue::Number(f.value())
    } else if let Ok(s) = obj.cast::<PyString>() {
        CellValue::String(s.to_string())
    } else {
        CellValue::String(obj.to_string())
    }
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
#[pyfunction]
fn read_xlsx(py: Python<'_>, path: &str) -> PyResult<Py<PyAny>> {
    let file = File::open(path)
        .map_err(|e| pyo3::exceptions::PyFileNotFoundError::new_err(format!("{path}: {e}")))?;
    let reader = BufReader::new(file);
    let sheets = reader::xlsx::read_xlsx(reader)?;

    let result = PyList::empty(py);
    for sheet in sheets {
        let dict = PyDict::new(py);
        dict.set_item("name", &sheet.name)?;
        dict.set_item("rows", rows_to_py(py, &sheet.rows)?)?;
        result.append(dict)?;
    }

    Ok(result.into_any().unbind())
}

/// Read a specific sheet by name or index from an XLSX file.
///
/// Returns a list of rows (list of lists of cell values).
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
    let sheets = reader::xlsx::read_xlsx(reader)?;

    let sheet = if let Some(name) = sheet_name {
        sheets.iter().find(|s| s.name == name).ok_or_else(|| {
            pyo3::exceptions::PyValueError::new_err(format!("Sheet '{name}' not found"))
        })?
    } else if let Some(idx) = sheet_index {
        sheets.get(idx).ok_or_else(|| {
            pyo3::exceptions::PyValueError::new_err(format!(
                "Sheet index {idx} out of range (file has {} sheets)",
                sheets.len()
            ))
        })?
    } else {
        sheets
            .first()
            .ok_or_else(|| pyo3::exceptions::PyValueError::new_err("No sheets found in file"))?
    };

    rows_to_py(py, &sheet.rows)
}

/// List sheet names in an XLSX file.
#[pyfunction]
fn sheet_names(path: &str) -> PyResult<Vec<String>> {
    let file = File::open(path)
        .map_err(|e| pyo3::exceptions::PyFileNotFoundError::new_err(format!("{path}: {e}")))?;
    let reader = BufReader::new(file);
    let sheets = reader::xlsx::read_xlsx(reader)?;
    Ok(sheets.iter().map(|s| s.name.clone()).collect())
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
    m.add_function(wrap_pyfunction!(version, m)?)?;
    m.add_function(wrap_pyfunction!(read_xlsx, m)?)?;
    m.add_function(wrap_pyfunction!(read_sheet, m)?)?;
    m.add_function(wrap_pyfunction!(sheet_names, m)?)?;
    m.add_class::<XlsxWriter>()?;
    Ok(())
}
