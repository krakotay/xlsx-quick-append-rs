use pyo3::exceptions::PyRuntimeError;
use pyo3::prelude::*;

use pyo3::PyRefMut;
use rust_core::{scan, XlsxEditor};
use std::path::PathBuf;

#[cfg(feature = "polars")]
use pyo3_polars::PyDataFrame;

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
    #[pyo3(signature = (path, sheet_name))]
    fn new(path: PathBuf, sheet_name: &str) -> PyResult<Self> {
        let openned = XlsxEditor::open(path, sheet_name)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(PyXlsxEditor { editor: openned })
    }
    #[pyo3(signature = (sheet_name))]
    fn add_worksheet(&mut self, sheet_name: &str) -> PyResult<()> {
        self.editor
            .add_worksheet(sheet_name)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))
    }
    fn append_row(&mut self, cells: Vec<String>) -> PyResult<()> {
        self.editor
            .append_row(cells)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))
    }

    fn append_table_at(&mut self, cells: Vec<Vec<String>>, start_cell: &str) -> PyResult<()> {
        self.editor
            .append_table_at(start_cell, cells)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))
    }
    fn last_row_index(&mut self, col_name: String) -> PyResult<u32> {
        self.editor
            .get_last_row_index(&col_name)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))
    }
    fn last_rows_index(&mut self, col_name: String) -> PyResult<Vec<u32>> {
        self.editor
            .get_last_roww_index(&col_name)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))
    }

    fn save(&mut self, path: PathBuf) -> PyResult<()> {
        self.editor
            .save(path)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))
    }
    #[cfg(feature = "polars")]
    #[pyo3(signature = (py_df, start_cell = None))]
    fn with_polars(&mut self, py_df: PyDataFrame, start_cell: Option<String>) -> PyResult<()> {
        let df = py_df.into();
        self.editor
            .with_polars(&df, start_cell.as_deref())
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))
    }
    fn set_number_format<'py>(
        mut slf: PyRefMut<'py, Self>,
        range: &str,
        fmt: &str,
    ) -> PyResult<PyRefMut<'py, Self>> {
        slf.editor
            .set_number_format(range, fmt)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(slf)
    }

    fn set_fill<'py>(
        mut slf: PyRefMut<'py, Self>,
        range: &str,
        fmt: &str,
    ) -> PyResult<PyRefMut<'py, Self>> {
        slf.editor
            .set_fill(range, fmt)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(slf)
    }
    fn set_font<'py>(
        mut slf: PyRefMut<'py, Self>,
        range: &str,
        name: &str,
        size: f32,
        bold: bool,
        italic: bool,
    ) -> PyResult<PyRefMut<'py, Self>> {
        slf.editor
            .set_font(range, name, size, bold, italic)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(slf)
    }
    fn merge_cells<'py>(mut slf: PyRefMut<'py, Self>, range: &str) -> PyResult<PyRefMut<'py, Self>> {
        slf.editor
            .merge_cells(range)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(slf)
    }

    // set_font
    // fn set_number_format(&mut self, range: &str, fmt: &str) -> PyResult<()> {
    //     self.editor
    //         .set_number_format(range, fmt)
    //         .map_err(|e| PyRuntimeError::new_err(e.to_string()))
    // }
    // fn set_number_format(&mut self, range: &str, fmt: &str) -> PyResult<()> {
    //     self.editor
    //         .set_number_format(range, fmt)
    //         .map_err(|e| PyRuntimeError::new_err(e.to_string()))
    // }
}
#[pyclass]
struct PyXlsxScanner {
    path: PathBuf,
}
#[pymethods]
impl PyXlsxScanner {
    #[new]
    fn new(path: PathBuf) -> PyResult<Self> {
        Ok(PyXlsxScanner { path })
    }
    fn get_sheets(&self) -> PyResult<Vec<String>> {
        scan_excel(self.path.clone()).map_err(|e| PyRuntimeError::new_err(e.to_string()))
    }
    fn open_editor(&self, sheet_name: String) -> PyResult<PyXlsxEditor> {
        let openned = XlsxEditor::open(self.path.clone(), &sheet_name)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(PyXlsxEditor { editor: openned })
    }
}

#[pymodule]
fn xlsx_append_py(_py: Python<'_>, m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<PyXlsxEditor>()?;
    m.add_class::<PyXlsxScanner>()?;
    m.add_function(wrap_pyfunction!(scan_excel, m)?)?;
    Ok(())
}
