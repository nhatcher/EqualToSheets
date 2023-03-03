use pyo3::exceptions::{PyException, PyValueError};
use pyo3::{create_exception, prelude::*, wrap_pyfunction};

use equalto_calc::expressions::utils;
use equalto_calc::model::{Environment, Model};
use equalto_calc::types::CellType;
use equalto_calc::types::Worksheet;
use equalto_xlsx::error::XlsxError;
use equalto_xlsx::export::save_to_xlsx;
use equalto_xlsx::import::load_model_from_xlsx;

create_exception!(_equalto, WorkbookError, PyException);

#[pyclass]
pub struct PyModel {
    model: Model,
}

impl PyModel {
    #[inline]
    fn worksheet(&self, sheet: i32) -> Result<&Worksheet, PyErr> {
        self.model
            .workbook
            .worksheet(sheet.try_into().unwrap())
            .map_err(WorkbookError::new_err)
    }
}

#[pymethods]
impl PyModel {
    pub fn get_name(&self) -> PyResult<String> {
        Ok(self.model.workbook.name.clone())
    }

    pub fn get_formatted_cell_value(&self, sheet: i32, row: i32, column: i32) -> PyResult<String> {
        self.model
            .formatted_cell_value(sheet.try_into().unwrap(), row, column)
            .map_err(WorkbookError::new_err)
    }

    pub fn get_cell_formula(&self, sheet: i32, row: i32, column: i32) -> PyResult<Option<String>> {
        self.model
            .cell_formula(sheet.try_into().unwrap(), row, column)
            .map_err(WorkbookError::new_err)
    }

    pub fn get_cell_value_by_index(&self, sheet: i32, row: i32, column: i32) -> PyResult<String> {
        Ok(self
            .model
            .get_cell_value_by_index(sheet.try_into().unwrap(), row, column)
            .map_err(WorkbookError::new_err)?
            .to_json_str())
    }

    pub fn get_cell_type(&self, sheet: i32, row: i32, column: i32) -> PyResult<i32> {
        Ok(self
            .worksheet(sheet)?
            .cell(row, column)
            .map_or(CellType::Number, |cell| cell.get_type()) as i32)
    }

    pub fn set_cell_empty(&mut self, sheet: i32, row: i32, column: i32) -> PyResult<()> {
        self.model
            .set_cell_empty(sheet.try_into().unwrap(), row, column)
            .map_err(WorkbookError::new_err)
    }

    pub fn delete_cell(&mut self, sheet: i32, row: i32, column: i32) -> PyResult<()> {
        self.model
            .delete_cell(sheet.try_into().unwrap(), row, column)
            .map_err(WorkbookError::new_err)
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

    pub fn update_cell_with_formula(
        &mut self,
        sheet: i32,
        row: i32,
        column: i32,
        formula: String,
    ) -> PyResult<()> {
        self.model
            .update_cell_with_formula(sheet.try_into().unwrap(), row, column, formula)
            .map_err(WorkbookError::new_err)
    }

    pub fn set_user_input(&mut self, sheet: i32, row: i32, column: i32, value: String) {
        let sheet: u32 = sheet.try_into().unwrap();
        self.model.set_user_input(sheet, row, column, value)
    }

    pub fn get_timezone(&self) -> PyResult<String> {
        Ok(self.model.tz.to_string())
    }

    pub fn save_to_xlsx(&self, file: &str) -> PyResult<()> {
        save_to_xlsx(&self.model, file).map_err(|e| WorkbookError::new_err(e.to_string()))
    }

    pub fn to_json(&self) -> PyResult<String> {
        Ok(self.model.to_json_str())
    }
}

impl WorkbookError {
    fn from_xlsx_error(error: XlsxError) -> PyErr {
        WorkbookError::new_err(error.user_message())
    }
}

#[pyfunction]
pub fn load_excel(file_path: &str, locale: &str, tz: &str) -> PyResult<PyModel> {
    Ok(PyModel {
        model: load_model_from_xlsx(file_path, locale, tz)
            .map_err(WorkbookError::from_xlsx_error)?,
    })
}

#[pyfunction]
pub fn load_json(workbook_json: &str) -> PyResult<PyModel> {
    let env = Environment::default();
    let model = Model::from_json(workbook_json, env).map_err(WorkbookError::new_err)?;
    Ok(PyModel { model })
}

#[pyfunction]
pub fn create(name: &str, locale: &str, tz: &str) -> PyResult<PyModel> {
    let env = Environment::default();
    let model = Model::new_empty(name, locale, tz, env).map_err(WorkbookError::new_err)?;
    Ok(PyModel { model })
}

#[pyfunction]
pub fn number_to_column(col_number: i32) -> PyResult<String> {
    utils::number_to_column(col_number)
        .ok_or_else(|| PyValueError::new_err("Invalid column number"))
}

/// A Python module implemented in Rust.
#[pymodule]
fn _equalto(py: Python, m: &PyModule) -> PyResult<()> {
    m.add("__version__", env!("CARGO_PKG_VERSION"))?;

    // errors
    m.add("WorkbookError", py.get_type::<WorkbookError>())?;

    // PyModel
    m.add_function(wrap_pyfunction!(create, m)?).unwrap();
    m.add_function(wrap_pyfunction!(load_excel, m)?).unwrap();
    m.add_function(wrap_pyfunction!(load_json, m)?).unwrap();

    // utils
    m.add_function(wrap_pyfunction!(number_to_column, m)?)
        .unwrap();

    Ok(())
}
