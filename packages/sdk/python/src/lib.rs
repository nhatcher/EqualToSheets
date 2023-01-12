use pyo3::exceptions::{PyException, PyValueError};
use pyo3::{create_exception, prelude::*, wrap_pyfunction};

use equalto_calc::{
    expressions::lexer::util::get_tokens as tokenizer,
    model::{Environment, Model},
};
use equalto_xlsx::compare::compare;
use equalto_xlsx::error::XlsxError;
use equalto_xlsx::export::save_to_xlsx;
use equalto_xlsx::import::load_from_excel;
use serde_json::json;
use std::time::{SystemTime, UNIX_EPOCH};

// This is equivalent to the JavaScript Date.now()
fn get_milliseconds_since_epoch() -> i64 {
    let start = SystemTime::now();
    start.duration_since(UNIX_EPOCH).unwrap().as_millis() as i64
}

create_exception!(_pycalc, WorkbookError, PyException);

#[pyclass]
pub struct PyModel {
    model: Model,
}

#[pyclass]
pub struct Cell {
    #[pyo3(get, set)]
    pub row: i32,
    #[pyo3(get, set)]
    pub column: i32,
}

#[pymethods]
impl PyModel {
    pub fn to_string(&self) -> PyResult<String> {
        Ok(self.model.to_json_str())
    }

    pub fn get_text_at(&self, sheet: i32, row: i32, column: i32) -> PyResult<String> {
        Ok(self
            .model
            .get_text_at(sheet.try_into().unwrap(), row, column))
    }

    pub fn get_ui_cell(&self, sheet: i32, row: i32, column: i32) -> PyResult<String> {
        Ok(serde_json::to_string(
            &self
                .model
                .get_ui_cell(sheet.try_into().unwrap(), row, column),
        )
        .unwrap())
    }

    pub fn format_number(&self, value: f64, format_code: String) -> PyResult<String> {
        Ok(self.model.format_number(value, format_code).text)
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

    pub fn get_cell_value_by_ref(&self, cell_reference: &str) -> PyResult<String> {
        match self.model.get_cell_value_by_ref(cell_reference) {
            Ok(s) => Ok(s.to_json_str()),
            Err(e) => Err(PyValueError::new_err(e)),
        }
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

    pub fn set_cells_with_values_json(&mut self, input_json: &str) {
        self.model.set_cells_with_values_json(input_json).unwrap();
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

    pub fn get_column_width(&self, sheet: i32, column: i32) -> PyResult<f64> {
        Ok(self
            .model
            .get_column_width(sheet.try_into().unwrap(), column))
    }

    pub fn get_row_height(&self, sheet: i32, row: i32) -> PyResult<f64> {
        Ok(self.model.get_row_height(sheet.try_into().unwrap(), row))
    }

    pub fn set_column_width(&mut self, sheet: i32, column: i32, width: f64) {
        self.model
            .set_column_width(sheet.try_into().unwrap(), column, width);
    }

    pub fn set_row_height(&mut self, sheet: i32, row: i32, height: f64) {
        self.model
            .set_row_height(sheet.try_into().unwrap(), row, height);
    }

    pub fn get_merge_cells(&self, sheet: i32) -> PyResult<String> {
        Ok(self.model.get_merge_cells(sheet.try_into().unwrap()))
    }

    pub fn get_tabs(&self) -> PyResult<String> {
        Ok(self.model.get_tabs())
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

    pub fn get_cell_style_index(&self, sheet: i32, row: i32, column: i32) -> i32 {
        self.model
            .get_cell_style_index(sheet.try_into().unwrap(), row, column)
    }

    pub fn remove_sheet_data(&mut self, sheet: i32) -> PyResult<()> {
        match self.model.remove_sheet_data(sheet.try_into().unwrap()) {
            Ok(()) => Ok(()),
            Err(message) => Err(PyValueError::new_err(message)),
        }
    }

    pub fn insert_sheet(
        &mut self,
        sheet_name: &str,
        sheet_index: i32,
        sheet_id: Option<i32>,
    ) -> PyResult<()> {
        match self.model.insert_sheet(
            sheet_name,
            sheet_index.try_into().unwrap(),
            sheet_id.map(|sheet_id| sheet_id.try_into().unwrap()),
        ) {
            Ok(()) => Ok(()),
            Err(message) => Err(PyValueError::new_err(message)),
        }
    }

    pub fn get_new_sheet_id(&self) -> i32 {
        self.model.get_new_sheet_id().try_into().unwrap()
    }

    pub fn set_cell_style_by_name(
        &mut self,
        sheet: i32,
        row: i32,
        column: i32,
        style_name: &str,
    ) -> PyResult<()> {
        let result =
            self.model
                .set_cell_style_by_name(sheet.try_into().unwrap(), row, column, style_name);
        match result {
            Ok(()) => Ok(()),
            Err(message) => Err(PyValueError::new_err(message)),
        }
    }

    pub fn get_worksheet_names(&self) -> PyResult<Vec<String>> {
        Ok(self.model.get_worksheet_names())
    }

    pub fn get_worksheet_ids(&self) -> PyResult<Vec<i32>> {
        Ok(self
            .model
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

    pub fn set_sheet_row_style(&mut self, sheet: i32, row: i32, style_name: &str) -> PyResult<()> {
        match self
            .model
            .set_sheet_row_style(sheet.try_into().unwrap(), row, style_name)
        {
            Ok(()) => Ok(()),
            Err(s) => Err(PyValueError::new_err(s)),
        }
    }

    pub fn set_sheet_column_style(
        &mut self,
        sheet: i32,
        column: i32,
        style_name: &str,
    ) -> PyResult<()> {
        match self
            .model
            .set_sheet_column_style(sheet.try_into().unwrap(), column, style_name)
        {
            Ok(()) => Ok(()),
            Err(s) => Err(PyValueError::new_err(s)),
        }
    }

    pub fn set_sheet_style(&mut self, sheet: i32, style_name: &str) -> PyResult<()> {
        match self
            .model
            .set_sheet_style(sheet.try_into().unwrap(), style_name)
        {
            Ok(()) => Ok(()),
            Err(s) => Err(PyValueError::new_err(s)),
        }
    }

    pub fn get_style_index_by_name(&self, style_name: &str) -> PyResult<i32> {
        match self
            .model
            .workbook
            .styles
            .get_style_index_by_name(style_name)
        {
            Ok(i) => Ok(i),
            Err(_) => Ok(-1),
        }
    }

    pub fn set_sheet_color(&mut self, sheet: i32, color: &str) -> PyResult<()> {
        match self.model.set_sheet_color(sheet.try_into().unwrap(), color) {
            Ok(_) => Ok(()),
            Err(s) => Err(PyValueError::new_err(s)),
        }
    }

    pub fn get_timezone(&self) -> PyResult<String> {
        Ok(self.model.tz.to_string())
    }

    pub fn save_to_xlsx(&self, file: &str) -> PyResult<()> {
        save_to_xlsx(&self.model, file).map_err(|e| WorkbookError::new_err(e.to_string()))
    }

    pub fn test_panic(&self) -> PyResult<()> {
        panic!("This function panics for testing panic handling");
    }
}

#[pyfunction]
pub fn loads(data: String) -> PyModel {
    let env = Environment {
        get_milliseconds_since_epoch,
    };
    let model = Model::from_json(&data, env).unwrap();
    PyModel { model }
}

impl WorkbookError {
    fn from_xlsx_error(error: XlsxError) -> PyErr {
        WorkbookError::new_err(error.to_string())
    }
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

#[pyfunction]
pub fn compare_models(workbook1: &str, workbook2: &str) -> PyResult<bool> {
    let env = Environment {
        get_milliseconds_since_epoch,
    };
    let model1 = match Model::from_json(workbook1, env.clone()) {
        Ok(m1) => m1,
        Err(message) => return Err(PyValueError::new_err(message)),
    };
    let model2 = match Model::from_json(workbook2, env) {
        Ok(m2) => m2,
        Err(message) => return Err(PyValueError::new_err(message)),
    };
    Ok(compare(model1, model2).is_ok())
}

#[pyfunction]
pub fn get_tokens(formula: &str) -> PyResult<String> {
    let tokens = tokenizer(formula);
    Ok(match serde_json::to_string(&tokens) {
        Ok(s) => s,
        Err(_) => json!([]).to_string(),
    })
}

#[pyfunction]
pub fn test_panic() {
    panic!("This function panics for testing panic handling");
}

/// A Python module implemented in Rust.
#[pymodule]
fn _pycalc(py: Python, m: &PyModule) -> PyResult<()> {
    m.add("__version__", env!("CARGO_PKG_VERSION"))?;
    m.add_function(wrap_pyfunction!(loads, m)?).unwrap();
    m.add_function(wrap_pyfunction!(create, m)?).unwrap();
    m.add_function(wrap_pyfunction!(load_excel, m)?).unwrap();
    m.add_function(wrap_pyfunction!(compare_models, m)?)
        .unwrap();
    m.add_function(wrap_pyfunction!(test_panic, m)?).unwrap();
    m.add("WorkbookError", py.get_type::<WorkbookError>())?;
    // m.add_class::<PyModel>()?;
    Ok(())
}
