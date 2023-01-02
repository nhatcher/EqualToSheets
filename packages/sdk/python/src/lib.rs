use pyo3::{exceptions::PyValueError, prelude::*, wrap_pyfunction};

use equalto_calc::{
    expressions::lexer::util::get_tokens as tokenizer,
    model::{Environment, Model},
};
use equalto_xlsx::compare::compare;
use equalto_xlsx::load_from_excel;
use serde_json::json;
use std::time::{SystemTime, UNIX_EPOCH};

// This is equivalent to the JavaScript Date.now()
fn get_milliseconds_since_epoch() -> i64 {
    let start = SystemTime::now();
    start.duration_since(UNIX_EPOCH).unwrap().as_millis() as i64
}

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
        Ok(self.model.get_text_at(sheet, row, column))
    }

    pub fn get_ui_cell(&self, sheet: i32, row: i32, column: i32) -> PyResult<String> {
        Ok(serde_json::to_string(&self.model.get_ui_cell(sheet, row, column)).unwrap())
    }

    pub fn format_number(&self, value: f64, format_code: String) -> PyResult<String> {
        Ok(self.model.format_number(value, format_code).text)
    }

    pub fn extend_to(
        &self,
        sheet: i32,
        row: i32,
        column: i32,
        target_row: i32,
        target_column: i32,
    ) -> PyResult<String> {
        Ok(self
            .model
            .extend_to(sheet, row, column, target_row, target_column))
    }

    pub fn has_formula(&self, sheet: i32, row: i32, column: i32) -> PyResult<bool> {
        Ok(self.model.has_formula(sheet, row, column))
    }

    pub fn get_formula_or_value(&self, sheet: i32, row: i32, column: i32) -> PyResult<String> {
        Ok(self.model.get_formula_or_value(sheet, row, column))
    }

    pub fn get_cell_value_by_ref(&self, cell_reference: &str) -> PyResult<String> {
        match self.model.get_cell_value_by_ref(cell_reference) {
            Ok(s) => Ok(s.to_json_str()),
            Err(e) => Err(PyValueError::new_err(e)),
        }
    }

    pub fn get_cell_value_by_index(
        &self,
        sheet: i32,
        row: i32,
        column: i32,
    ) -> PyResult<String> {
        Ok(self.model.get_cell_value_by_index(sheet, row, column).to_json_str())
    }

    pub fn get_cell_type(&self, sheet: i32, row: i32, column: i32) -> PyResult<i32> {
        Ok(self.model.get_cell_type(sheet, row, column) as i32)
    }

    pub fn set_input(&mut self, sheet: i32, row: i32, column: i32, value: String) {
        let style = self.model.get_cell_style_index(sheet, row, column);
        self.model.set_input(sheet, row, column, value, style)
    }

    pub fn set_cells_with_values_json(&mut self, input_json: &str) {
        self.model.set_cells_with_values_json(input_json).unwrap();
    }

    pub fn delete_cell(&mut self, sheet: i32, row: i32, column: i32) {
        self.model.delete_cell(sheet, row, column)
    }

    pub fn evaluate(&mut self) {
        self.model.evaluate()
    }

    pub fn evaluate_with_input(
        &mut self,
        input_json: &str,
        output_refs: Vec<String>,
    ) -> PyResult<String> {
        let output: Vec<&str> = output_refs.iter().map(AsRef::as_ref).collect();
        match self.model.evaluate_with_input(input_json, &output) {
            Ok(o) => Ok(serde_json::to_string(&o).unwrap()),
            Err(e) => Err(PyValueError::new_err(e)),
        }
    }

    pub fn get_range_formatted_data(&self, range: &str) -> PyResult<String> {
        match self.model.get_range_formatted_data(range) {
            Ok(o) => Ok(serde_json::to_string(&o).unwrap()),
            Err(e) => Err(PyValueError::new_err(e)),
        }
    }

    pub fn get_column_width(&self, sheet: i32, column: i32) -> PyResult<f64> {
        Ok(self.model.get_column_width(sheet, column))
    }

    pub fn get_row_height(&self, sheet: i32, row: i32) -> PyResult<f64> {
        Ok(self.model.get_row_height(sheet, row))
    }

    pub fn set_column_width(&mut self, sheet: i32, column: i32, width: f64) {
        self.model.set_column_width(sheet, column, width);
    }

    pub fn set_row_height(&mut self, sheet: i32, row: i32, height: f64) {
        self.model.set_row_height(sheet, row, height);
    }

    pub fn get_merge_cells(&self, sheet: i32) -> PyResult<String> {
        Ok(self.model.get_merge_cells(sheet))
    }

    pub fn get_tabs(&self) -> PyResult<String> {
        Ok(self.model.get_tabs())
    }

    pub fn get_style_for_cell(&self, sheet: i32, row: i32, column: i32) -> PyResult<String> {
        Ok(serde_json::to_string(&self.model.get_style_for_cell(sheet, row, column)).unwrap())
    }

    pub fn get_navigation_right_edge(&self, sheet: i32, row: i32, column: i32) -> PyResult<i32> {
        Ok(self.model.get_navigation_right_edge(sheet, row, column))
    }

    pub fn get_navigation_left_edge(&self, sheet: i32, row: i32, column: i32) -> PyResult<i32> {
        Ok(self.model.get_navigation_left_edge(sheet, row, column))
    }

    pub fn get_navigation_top_edge(&self, sheet: i32, row: i32, column: i32) -> PyResult<i32> {
        Ok(self.model.get_navigation_top_edge(sheet, row, column))
    }

    pub fn get_navigation_bottom_edge(&self, sheet: i32, row: i32, column: i32) -> PyResult<i32> {
        Ok(self.model.get_navigation_bottom_edge(sheet, row, column))
    }

    pub fn get_navigation_home(&self, sheet: i32) -> Cell {
        let (row, column) = self.model.get_navigation_home(sheet);
        Cell { row, column }
    }

    pub fn get_navigation_end(&self, sheet: i32) -> PyResult<Cell> {
        let (row, column) = self.model.get_navigation_end(sheet);
        Ok(Cell { row, column })
    }

    pub fn set_cell_style(&mut self, sheet: i32, row: i32, column: i32, style: &str) {
        self.model
            .set_cell_style(sheet, row, column, &serde_json::from_str(style).unwrap())
            .unwrap()
    }

    pub fn get_cell_style_index(&self, sheet: i32, row: i32, column: i32) -> i32 {
        self.model.get_cell_style_index(sheet, row, column)
    }

    pub fn remove_sheet_data(&mut self, sheet_index: i32) -> PyResult<()> {
        match self.model.remove_sheet_data(sheet_index) {
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
        match self.model.insert_sheet(sheet_name, sheet_index, sheet_id) {
            Ok(()) => Ok(()),
            Err(message) => Err(PyValueError::new_err(message)),
        }
    }

    pub fn get_new_sheet_id(&self) -> i32 {
        self.model.get_new_sheet_id()
    }

    pub fn swap_cells_in_row(
        &mut self,
        sheet: i32,
        row: i32,
        column1: i32,
        column2: i32,
    ) -> PyResult<()> {
        match self.model.swap_cells_in_row(sheet, row, column1, column2) {
            Ok(()) => Ok(()),
            Err(message) => Err(PyValueError::new_err(message)),
        }
    }

    pub fn move_column_action(&mut self, sheet: i32, column: i32, delta: i32) -> PyResult<()> {
        match self.model.move_column_action(sheet, column, delta) {
            Ok(()) => Ok(()),
            Err(message) => Err(PyValueError::new_err(message)),
        }
    }

    pub fn shift_cells_right(
        &mut self,
        sheet: i32,
        row: i32,
        column: i32,
        cell_count: i32,
    ) -> PyResult<String> {
        let result = self.model.shift_cells_right(sheet, row, column, cell_count);
        match result {
            Ok(()) => Ok("{\"success\":true}".to_string()),
            Err(message) => Ok(format!(
                "{{\"success\": false, \"message\":\"{}\"}}",
                message
            )),
        }
    }

    pub fn shift_cells_left(
        &mut self,
        sheet: i32,
        row: i32,
        column: i32,
        cell_count: i32,
    ) -> PyResult<String> {
        let result = self.model.shift_cells_left(sheet, row, column, cell_count);
        match result {
            Ok(()) => Ok("{\"success\":true}".to_string()),
            Err(message) => Ok(format!(
                "{{\"success\": false, \"message\":\"{}\"}}",
                message
            )),
        }
    }

    pub fn shift_cells_down(
        &mut self,
        sheet: i32,
        row: i32,
        column: i32,
        cell_count: i32,
    ) -> PyResult<String> {
        let result = self.model.shift_cells_down(sheet, row, column, cell_count);
        match result {
            Ok(()) => Ok("{\"success\":true}".to_string()),
            Err(message) => Ok(format!(
                "{{\"success\": false, \"message\":\"{}\"}}",
                message
            )),
        }
    }

    pub fn shift_cells_up(
        &mut self,
        sheet: i32,
        row: i32,
        column: i32,
        cell_count: i32,
    ) -> PyResult<String> {
        let result = self.model.shift_cells_up(sheet, row, column, cell_count);
        match result {
            Ok(()) => Ok("{\"success\":true}".to_string()),
            Err(message) => Ok(format!(
                "{{\"success\": false, \"message\":\"{}\"}}",
                message
            )),
        }
    }

    pub fn insert_columns(
        &mut self,
        sheet: i32,
        column: i32,
        column_count: i32,
    ) -> PyResult<String> {
        let result = self.model.insert_columns(sheet, column, column_count);
        match result {
            Ok(()) => Ok("{\"success\":true}".to_string()),
            Err(message) => Ok(format!(
                "{{\"success\": false, \"message\":\"{}\"}}",
                message
            )),
        }
    }

    pub fn delete_columns(
        &mut self,
        sheet: i32,
        column: i32,
        column_count: i32,
    ) -> PyResult<String> {
        let result = self.model.delete_columns(sheet, column, column_count);
        match result {
            Ok(()) => Ok("{\"success\":true}".to_string()),
            Err(message) => Ok(format!(
                "{{\"success\": false, \"message\":\"{}\"}}",
                message
            )),
        }
    }

    pub fn insert_rows(&mut self, sheet: i32, row: i32, row_count: i32) -> PyResult<String> {
        let result = self.model.insert_rows(sheet, row, row_count);
        match result {
            Ok(()) => Ok("{\"success\":true}".to_string()),
            Err(message) => Ok(format!(
                "{{\"success\": false, \"message\":\"{}\"}}",
                message
            )),
        }
    }

    pub fn delete_rows(&mut self, sheet: i32, row: i32, row_count: i32) -> PyResult<String> {
        let result = self.model.delete_rows(sheet, row, row_count);
        match result {
            Ok(()) => Ok("{\"success\":true}".to_string()),
            Err(message) => Ok(format!(
                "{{\"success\": false, \"message\":\"{}\"}}",
                message
            )),
        }
    }

    pub fn cell_independent_of_sheets_and_cells(
        &self,
        input_cell: &str,
        sheets: Vec<&str>,
        cells: Vec<&str>,
    ) -> PyResult<bool> {
        match self
            .model
            .cell_independent_of_sheets_and_cells(input_cell, &sheets, &cells)
        {
            Ok(value) => Ok(value),
            Err(s) => Err(PyValueError::new_err(s)),
        }
    }

    pub fn set_cell_style_by_name(
        &mut self,
        sheet: i32,
        row: i32,
        column: i32,
        style_name: &str,
    ) -> PyResult<()> {
        let result = self
            .model
            .set_cell_style_by_name(sheet, row, column, style_name);
        match result {
            Ok(()) => Ok(()),
            Err(message) => Err(PyValueError::new_err(message)),
        }
    }

    pub fn get_worksheet_names(&self) -> PyResult<Vec<String>> {
        Ok(self.model.get_worksheet_names())
    }

    pub fn get_worksheet_ids(&self) -> PyResult<Vec<i32>> {
        Ok(self.model.get_worksheet_ids())
    }

    pub fn add_sheet(&mut self, new_name: &str) -> PyResult<bool> {
        Ok(self.model.add_sheet(new_name).is_ok())
    }

    pub fn rename_sheet(&mut self, old_name: &str, new_name: &str) -> PyResult<bool> {
        Ok(self.model.rename_sheet(old_name, new_name).is_ok())
    }

    pub fn delete_sheet_by_name(&mut self, name: &str) -> PyResult<bool> {
        Ok(self.model.delete_sheet_by_name(name).is_ok())
    }

    pub fn delete_sheet_by_sheet_id(&mut self, sheet_id: i32) -> PyResult<bool> {
        Ok(self.model.delete_sheet_by_sheet_id(sheet_id).is_ok())
    }

    pub fn update_cell_with_text(&mut self, sheet: i32, row: i32, column: i32, value: &str) {
        self.model.update_cell_with_text(sheet, row, column, value);
    }

    pub fn update_cell_with_number(&mut self, sheet: i32, row: i32, column: i32, value: f64) {
        self.model.update_cell_with_number(sheet, row, column, value);
    }

    pub fn update_cell_with_bool(&mut self, sheet: i32, row: i32, column: i32, value: bool) {
        self.model.update_cell_with_bool(sheet, row, column, value);
    }

    pub fn set_sheet_row_style(&mut self, sheet: i32, row: i32, style_name: &str) -> PyResult<()> {
        match self.model.set_sheet_row_style(sheet, row, style_name) {
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
        match self.model.set_sheet_column_style(sheet, column, style_name) {
            Ok(()) => Ok(()),
            Err(s) => Err(PyValueError::new_err(s)),
        }
    }

    pub fn set_sheet_style(&mut self, sheet: i32, style_name: &str) -> PyResult<()> {
        match self.model.set_sheet_style(sheet, style_name) {
            Ok(()) => Ok(()),
            Err(s) => Err(PyValueError::new_err(s)),
        }
    }

    pub fn get_style_index_by_name(&self, style_name: &str) -> PyResult<i32> {
        match self.model.get_style_index_by_name(style_name) {
            Ok(i) => Ok(i),
            Err(_) => Ok(-1),
        }
    }

    pub fn set_sheet_color(&mut self, sheet: i32, color: &str) -> PyResult<()> {
        match self.model.set_sheet_color(sheet, color) {
            Ok(_) => Ok(()),
            Err(s) => Err(PyValueError::new_err(s)),
        }
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

#[pyfunction]
pub fn load_excel(file_path: &str, locale: &str, tz: &str) -> PyModel {
    let env = Environment {
        get_milliseconds_since_epoch,
    };
    let model = load_from_excel(file_path, locale, tz);
    let s = serde_json::to_string(&model).unwrap();
    let model = Model::from_json(&s, env).unwrap();
    PyModel { model }
}

#[pyfunction]
pub fn create(name: &str, locale: &str, tz: &str) -> PyModel {
    let env = Environment {
        get_milliseconds_since_epoch,
    };
    let model = Model::new_empty(name, locale, tz, env).unwrap();
    PyModel { model }
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
fn _pycalc(_: Python, m: &PyModule) -> PyResult<()> {
    m.add("__version__", env!("CARGO_PKG_VERSION"))?;
    m.add_function(wrap_pyfunction!(loads, m)?).unwrap();
    m.add_function(wrap_pyfunction!(create, m)?).unwrap();
    m.add_function(wrap_pyfunction!(load_excel, m)?).unwrap();
    m.add_function(wrap_pyfunction!(compare_models, m)?)
        .unwrap();
    m.add_function(wrap_pyfunction!(test_panic, m)?).unwrap();
    // m.add_class::<PyModel>()?;
    Ok(())
}
