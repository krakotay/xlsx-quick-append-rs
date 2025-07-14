use pyo3::prelude::*;
use pyo3::exceptions::PyRuntimeError;
use std::path::PathBuf;
use rust_core::scan;
use rust_core::XlsxEditor;


#[pyfunction]
fn scan_excel(path: PathBuf) -> PyResult<Vec<String>> {
    scan(&path).map_err(|e| PyRuntimeError::new_err(e.to_string()))
}

#[pyclass]
struct PyXlsxEditor {
    editor: XlsxEditor,
}

#[pymethods]
impl PyXlsxEditor {
    #[new]
    fn new(path: PathBuf, sheet_name: &str) -> PyResult<Self> {
        let openned = XlsxEditor::open(path, sheet_name).map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(PyXlsxEditor { editor: openned })
    }

    fn append_row(&mut self, cells: Vec<String>) -> PyResult<()> {
        self.editor.append_row(cells).map_err(|e| PyRuntimeError::new_err(e.to_string()))
    }

    fn append_table_at(&mut self, cells: Vec<Vec<String>>, start_cell: &str) -> PyResult<()> {
        self.editor.append_table_at(start_cell, cells).map_err(|e| PyRuntimeError::new_err(e.to_string()))
    }

    fn save(&mut self, path: PathBuf) -> PyResult<()> {
        self.editor.save(path).map_err(|e| PyRuntimeError::new_err(e.to_string()))
    }
}

#[pymodule]
fn xlsx_append_py(_py: Python<'_>, m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<PyXlsxEditor>()?;
    m.add_function(wrap_pyfunction!(scan_excel, m)?)?;
    Ok(())
}
