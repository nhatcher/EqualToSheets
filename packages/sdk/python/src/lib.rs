use pyo3::exceptions::PyException;
use pyo3::{create_exception, prelude::*, wrap_pyfunction};

use equalto_calc::model::{Environment, Model};
use equalto_xlsx::error::XlsxError;
use equalto_xlsx::export::save_to_xlsx;
use equalto_xlsx::import::load_from_excel;
use std::time::{SystemTime, UNIX_EPOCH};

create_exception!(_pycalc, WorkbookError, PyException);

#[pyclass]
pub struct PyModel {
    model: Model,
}

#[pymethods]
impl PyModel {
    pub fn get_formatted_cell_value(&self, sheet: i32, row: i32, column: i32) -> PyResult<String> {
        Ok(self
            .model
            .get_formatted_cell_value(sheet.try_into().unwrap(), row, column))
    }

    pub fn has_formula(&self, sheet: i32, row: i32, column: i32) -> PyResult<bool> {
        Ok(self
            .model
            .has_formula(sheet.try_into().unwrap(), row, column))
    }

    pub fn get_formula_or_value(&self, sheet: i32, row: i32, column: i32) -> PyResult<String> {
        Ok(self
            .model
            .get_formula_or_value(sheet.try_into().unwrap(), row, column))
    }

    pub fn get_cell_value_by_index(&self, sheet: i32, row: i32, column: i32) -> PyResult<String> {
        Ok(self
            .model
            .get_cell_value_by_index(sheet.try_into().unwrap(), row, column)
            .to_json_str())
    }

    pub fn get_cell_type(&self, sheet: i32, row: i32, column: i32) -> PyResult<i32> {
        Ok(self
            .model
            .get_cell_type(sheet.try_into().unwrap(), row, column) as i32)
    }

    pub fn set_input(&mut self, sheet: i32, row: i32, column: i32, value: String) {
        let sheet: u32 = sheet.try_into().unwrap();
        let style = self.model.get_cell_style_index(sheet, row, column);
        self.model.set_input(sheet, row, column, value, style)
    }

    pub fn delete_cell(&mut self, sheet: i32, row: i32, column: i32) -> PyResult<()> {
        self.model
            .delete_cell(sheet.try_into().unwrap(), row, column)
            .unwrap();
        Ok(())
    }

    pub fn evaluate(&mut self) -> PyResult<()> {
        self.model
            .evaluate_with_error_check()
            .map_err(WorkbookError::new_err)
    }

    pub fn get_style_for_cell(&self, sheet: i32, row: i32, column: i32) -> PyResult<String> {
        Ok(serde_json::to_string(&self.model.get_style_for_cell(
            sheet.try_into().unwrap(),
            row,
            column,
        ))
        .unwrap())
    }

    pub fn set_cell_style(&mut self, sheet: i32, row: i32, column: i32, style: &str) {
        self.model
            .set_cell_style(
                sheet.try_into().unwrap(),
                row,
                column,
                &serde_json::from_str(style).unwrap(),
            )
            .unwrap()
    }

    pub fn get_worksheet_names(&self) -> PyResult<Vec<String>> {
        Ok(self.model.workbook.get_worksheet_names())
    }

    pub fn get_worksheet_ids(&self) -> PyResult<Vec<i32>> {
        Ok(self
            .model
            .workbook
            .get_worksheet_ids()
            .iter()
            .map(|&id| id.try_into().unwrap())
            .collect())
    }

    pub fn add_sheet(&mut self, name: &str) -> PyResult<()> {
        self.model.add_sheet(name).map_err(WorkbookError::new_err)
    }

    pub fn new_sheet(&mut self) -> PyResult<()> {
        self.model.new_sheet();
        Ok(())
    }

    pub fn rename_sheet(&mut self, sheet: i32, new_name: &str) -> PyResult<()> {
        self.model
            .rename_sheet_by_index(sheet.try_into().unwrap(), new_name)
            .map_err(WorkbookError::new_err)
    }

    pub fn delete_sheet_by_sheet_id(&mut self, sheet_id: i32) -> PyResult<()> {
        self.model
            .delete_sheet_by_sheet_id(sheet_id.try_into().unwrap())
            .map_err(WorkbookError::new_err)
    }

    pub fn update_cell_with_text(&mut self, sheet: i32, row: i32, column: i32, value: &str) {
        self.model
            .update_cell_with_text(sheet.try_into().unwrap(), row, column, value);
    }

    pub fn update_cell_with_number(&mut self, sheet: i32, row: i32, column: i32, value: f64) {
        self.model
            .update_cell_with_number(sheet.try_into().unwrap(), row, column, value);
    }

    pub fn update_cell_with_bool(&mut self, sheet: i32, row: i32, column: i32, value: bool) {
        self.model
            .update_cell_with_bool(sheet.try_into().unwrap(), row, column, value);
    }

    pub fn get_timezone(&self) -> PyResult<String> {
        Ok(self.model.tz.to_string())
    }

    pub fn save_to_xlsx(&self, file: &str) -> PyResult<()> {
        save_to_xlsx(&self.model, file).map_err(|e| WorkbookError::new_err(e.to_string()))
    }
}

impl WorkbookError {
    fn from_xlsx_error(error: XlsxError) -> PyErr {
        WorkbookError::new_err(error.to_string())
    }
}

// This is equivalent to the JavaScript Date.now()
fn get_milliseconds_since_epoch() -> i64 {
    let start = SystemTime::now();
    start.duration_since(UNIX_EPOCH).unwrap().as_millis() as i64
}

#[pyfunction]
pub fn load_excel(file_path: &str, locale: &str, tz: &str) -> PyResult<PyModel> {
    let env = Environment {
        get_milliseconds_since_epoch,
    };
    let model = load_from_excel(file_path, locale, tz).map_err(WorkbookError::from_xlsx_error)?;
    let s = serde_json::to_string(&model).map_err(|e| WorkbookError::new_err(e.to_string()))?;
    let model = Model::from_json(&s, env).map_err(WorkbookError::new_err)?;
    Ok(PyModel { model })
}

#[pyfunction]
pub fn create(name: &str, locale: &str, tz: &str) -> PyResult<PyModel> {
    let env = Environment {
        get_milliseconds_since_epoch,
    };
    let model = Model::new_empty(name, locale, tz, env).map_err(WorkbookError::new_err)?;
    Ok(PyModel { model })
}

/// A Python module implemented in Rust.
#[pymodule]
fn _pycalc(py: Python, m: &PyModule) -> PyResult<()> {
    m.add("__version__", env!("CARGO_PKG_VERSION"))?;
    m.add_function(wrap_pyfunction!(create, m)?).unwrap();
    m.add_function(wrap_pyfunction!(load_excel, m)?).unwrap();
    m.add("WorkbookError", py.get_type::<WorkbookError>())?;
    Ok(())
}
