use gloo_utils::format::JsValueSerdeExt;
use js_sys::Date;
use serde::{Deserialize, Serialize};
use serde_json::json;
use wasm_bindgen::prelude::{wasm_bindgen, JsValue};

use equalto_calc::{
    expressions::{
        lexer::util::get_tokens as tokenizer,
        types::{Area, CellReferenceIndex},
    },
    model::{ExcelValue, Model, Style},
    number_format,
};

use equalto_calc::model::Environment;

#[wasm_bindgen]
pub struct JSModel {
    model: Model,
}

#[wasm_bindgen]
pub struct Cell {
    pub row: i32,
    pub column: i32,
}

#[wasm_bindgen]
pub struct IntegerResult {
    pub success: bool,
    pub value: i32,
    message: String,
}

#[wasm_bindgen]
impl IntegerResult {
    #[wasm_bindgen(getter)]
    pub fn message(&self) -> String {
        self.message.clone()
    }
    fn get_success(value: i32) -> IntegerResult {
        IntegerResult {
            success: true,
            value,
            message: "".to_string(),
        }
    }
    fn get_error(message: &str) -> IntegerResult {
        IntegerResult {
            success: false,
            value: -1,
            message: message.to_string(),
        }
    }
}

#[wasm_bindgen]
pub struct IndexResult {
    pub success: bool,
    pub index: i32,
    message: String,
}

#[wasm_bindgen]
impl IndexResult {
    #[wasm_bindgen(getter)]
    pub fn message(&self) -> String {
        self.message.clone()
    }
    fn get_success(index: i32) -> IndexResult {
        IndexResult {
            success: true,
            index,
            message: "".to_string(),
        }
    }
    fn get_error(message: &str) -> IndexResult {
        IndexResult {
            success: false,
            index: -1,
            message: message.to_string(),
        }
    }
}

#[wasm_bindgen]
pub struct JsResult {
    pub success: bool,
    message: String,
}

#[wasm_bindgen]
impl JsResult {
    #[wasm_bindgen(getter)]
    pub fn message(&self) -> String {
        self.message.clone()
    }

    fn get_success() -> JsResult {
        JsResult {
            success: true,
            message: "".to_string(),
        }
    }

    fn get_error(message: &str) -> JsResult {
        JsResult {
            success: false,
            message: message.to_string(),
        }
    }
}

#[wasm_bindgen]
#[derive(Serialize, Deserialize)]
pub struct AreaJs {
    pub sheet: i32,
    pub row: i32,
    pub column: i32,
    pub width: i32,
    pub height: i32,
}

#[wasm_bindgen]
#[derive(Serialize, Deserialize)]
struct CellIndexJs {
    pub sheet: i32,
    pub row: i32,
    pub column: i32,
}
#[wasm_bindgen]
#[derive(Serialize, Deserialize)]
struct MoveData {
    source: CellIndexJs,
    value: String,
    target: CellIndexJs,
    area: AreaJs,
}

#[wasm_bindgen]
#[derive(Serialize, Deserialize)]
struct ExtendToData {
    source: CellIndexJs,
    value: String,
    target: CellIndexJs,
}

#[wasm_bindgen]
#[derive(Serialize, Deserialize)]
pub struct SheetDimension {
    pub min_row: i32,
    pub min_column: i32,
    pub max_row: i32,
    pub max_column: i32,
}

#[wasm_bindgen]
pub fn format_number(value: f64, format_code: &str, locale: &str) -> String {
    number_format::format_number(value, format_code, locale).text
}

/// Return a JSON string with a list of all the tokens from a formula
/// This is used by the UI to color them according to a theme.
#[wasm_bindgen]
pub fn get_tokens(formula: &str) -> String {
    let tokens = tokenizer(formula);
    match serde_json::to_string(&tokens) {
        Ok(s) => s,
        Err(_) => json!([]).to_string(),
    }
}

fn get_milliseconds_since_epoch() -> i64 {
    Date::now() as i64
}

#[wasm_bindgen]
impl JSModel {
    pub fn new(val: &JsValue) -> JSModel {
        let env = Environment {
            get_milliseconds_since_epoch,
        };
        let data: String = JsValueSerdeExt::into_serde(val).unwrap();
        let model = Model::from_json(&data, env).unwrap();
        JSModel { model }
    }

    pub fn new_empty(name: &str, locale: &str, timezone: &str) -> JSModel {
        let env = Environment {
            get_milliseconds_since_epoch,
        };
        let model = Model::new_empty(name, locale, timezone, env).unwrap();
        JSModel { model }
    }

    pub fn to_json_string(&self) -> String {
        self.model.to_json_str()
    }

    pub fn get_text_at(&self, sheet: i32, row: i32, column: i32) -> String {
        self.model.get_text_at(sheet, row, column)
    }

    pub fn get_ui_cell(&self, sheet: i32, row: i32, column: i32) -> String {
        let c = self.model.get_ui_cell(sheet, row, column);
        // FIXME: This should return an object not a string
        json!({"kind": c.kind, "value": c.value, "details": c.details}).to_string()
    }

    pub fn get_cell_value_by_ref(&self, cel_ref: &str) -> JsValue {
        match self.model.get_cell_value_by_ref(cel_ref) {
            Ok(result) => match result {
                ExcelValue::String(s) => JsValue::from_str(&s),
                ExcelValue::Boolean(b) => JsValue::from_bool(b),
                ExcelValue::Number(n) => JsValue::from_f64(n),
            },
            Err(_) => JsValue::from_str(""),
        }
    }

    pub fn get_range_data(&self, range: &str) -> String {
        match self.model.get_range_data(range) {
            Ok(result) => {
                if let Ok(result_str) = serde_json::to_string(&result) {
                    return json!({"success": true, "value": result_str}).to_string();
                }
                json!({"success": false, "message": "Failed processing data"}).to_string()
            }
            Err(s) => json!({"success": false, "message": s}).to_string(),
        }
    }

    pub fn format_number(&self, value: f64, format_code: String) -> String {
        self.model.format_number(value, format_code).text
    }

    pub fn extend_to(
        &self,
        sheet: i32,
        row: i32,
        column: i32,
        target_row: i32,
        target_column: i32,
    ) -> String {
        self.model
            .extend_to(sheet, row, column, target_row, target_column)
    }

    pub fn get_formula_or_value(&self, sheet: i32, row: i32, column: i32) -> String {
        self.model.get_formula_or_value(sheet, row, column)
    }

    pub fn has_formula(&self, sheet: i32, row: i32, column: i32) -> bool {
        self.model.has_formula(sheet, row, column)
    }

    pub fn set_input(&mut self, sheet: i32, row: i32, column: i32, value: String, style: i32) {
        self.model.set_input(sheet, row, column, value, style)
    }

    pub fn update_cell_with_text(&mut self, sheet: i32, row: i32, column: i32, value: &str) {
        self.model.update_cell_with_text(sheet, row, column, value)
    }

    pub fn update_cell_with_number(&mut self, sheet: i32, row: i32, column: i32, value: f64) {
        self.model
            .update_cell_with_number(sheet, row, column, value)
    }

    pub fn update_cell_with_bool(&mut self, sheet: i32, row: i32, column: i32, value: bool) {
        self.model.update_cell_with_bool(sheet, row, column, value)
    }

    pub fn set_cells_with_values_json(&mut self, input_json: &str) {
        self.model.set_cells_with_values_json(input_json).unwrap();
    }

    pub fn delete_cell(&mut self, sheet: i32, row: i32, column: i32) {
        self.model.delete_cell(sheet, row, column).unwrap();
    }

    pub fn remove_cell(&mut self, sheet: i32, row: i32, column: i32) {
        self.model.remove_cell(sheet, row, column).unwrap();
    }

    pub fn evaluate(&mut self) {
        self.model.evaluate()
    }

    pub fn get_column_width(&self, sheet: i32, column: i32) -> f64 {
        self.model.get_column_width(sheet, column)
    }

    pub fn get_row_height(&self, sheet: i32, row: i32) -> f64 {
        self.model.get_row_height(sheet, row)
    }

    pub fn set_column_width(&mut self, sheet: i32, column: i32, width: f64) {
        self.model.set_column_width(sheet, column, width);
    }

    pub fn set_row_height(&mut self, sheet: i32, row: i32, height: f64) {
        self.model.set_row_height(sheet, row, height);
    }

    pub fn get_frozen_rows(&self, sheet: i32) -> IntegerResult {
        match self.model.get_frozen_rows(sheet) {
            Ok(value) => IntegerResult::get_success(value),
            Err(message) => IntegerResult::get_error(&message),
        }
    }

    pub fn get_frozen_columns(&self, sheet: i32) -> IntegerResult {
        match self.model.get_frozen_columns(sheet) {
            Ok(value) => IntegerResult::get_success(value),
            Err(message) => IntegerResult::get_error(&message),
        }
    }

    pub fn set_frozen_columns(&mut self, sheet: i32, frozen_columns: i32) -> JsResult {
        match self.model.set_frozen_columns(sheet, frozen_columns) {
            Ok(()) => JsResult::get_success(),
            Err(message) => JsResult::get_error(&message),
        }
    }

    pub fn set_frozen_rows(&mut self, sheet: i32, frozen_rows: i32) -> JsResult {
        match self.model.set_frozen_rows(sheet, frozen_rows) {
            Ok(()) => JsResult::get_success(),
            Err(message) => JsResult::get_error(&message),
        }
    }

    pub fn get_merge_cells(&self, sheet: i32) -> String {
        // FIXME: This should return an object not a string
        self.model.get_merge_cells(sheet)
    }

    pub fn get_tabs(&self) -> String {
        // FIXME: This should not return a string but an object
        self.model.get_tabs()
    }

    pub fn create_named_style(&mut self, style_name: &str, style_js: &JsValue) -> JsResult {
        let style: Style = match JsValueSerdeExt::into_serde(style_js) {
            Ok(s) => s,
            Err(error) => return JsResult::get_error(&error.to_string()),
        };
        match self.model.create_named_style(style_name, &style) {
            Ok(()) => JsResult::get_success(),
            Err(message) => JsResult::get_error(&message),
        }
    }

    pub fn get_sheet_dimension(&self, sheet: i32) -> SheetDimension {
        let (min_row, min_column, max_row, max_column) = self.model.get_sheet_dimension(sheet);
        SheetDimension {
            min_row,
            min_column,
            max_row,
            max_column,
        }
    }

    pub fn set_sheet_style(&mut self, sheet: i32, style_name: &str) -> JsResult {
        match self.model.set_sheet_style(sheet, style_name) {
            Ok(()) => JsResult::get_success(),
            Err(message) => JsResult::get_error(&message),
        }
    }

    pub fn set_sheet_row_style(&mut self, sheet: i32, row: i32, style_name: &str) -> JsResult {
        match self.model.set_sheet_row_style(sheet, row, style_name) {
            Ok(()) => JsResult::get_success(),
            Err(message) => JsResult::get_error(&message),
        }
    }

    pub fn set_sheet_column_style(
        &mut self,
        sheet: i32,
        column: i32,
        style_name: &str,
    ) -> JsResult {
        match self.model.set_sheet_column_style(sheet, column, style_name) {
            Ok(()) => JsResult::get_success(),
            Err(message) => JsResult::get_error(&message),
        }
    }

    pub fn get_style_for_cell(&self, sheet: i32, row: i32, column: i32) -> String {
        // FIXME: This should not return a string but an object
        serde_json::to_string(&self.model.get_style_for_cell(sheet, row, column)).unwrap()
    }

    pub fn get_navigation_right_edge(&self, sheet: i32, row: i32, column: i32) -> i32 {
        self.model
            .get_navigation_right_edge(sheet, row, column)
            .unwrap()
    }

    pub fn get_navigation_left_edge(&self, sheet: i32, row: i32, column: i32) -> i32 {
        self.model
            .get_navigation_left_edge(sheet, row, column)
            .unwrap()
    }

    pub fn get_navigation_top_edge(&self, sheet: i32, row: i32, column: i32) -> i32 {
        self.model
            .get_navigation_top_edge(sheet, row, column)
            .unwrap()
    }

    pub fn get_navigation_bottom_edge(&self, sheet: i32, row: i32, column: i32) -> i32 {
        self.model
            .get_navigation_bottom_edge(sheet, row, column)
            .unwrap()
    }

    pub fn get_navigation_home(&self, sheet: i32) -> Cell {
        let (row, column) = self.model.get_navigation_home(sheet);
        Cell { row, column }
    }

    pub fn get_navigation_end(&self, sheet: i32) -> Cell {
        let (row, column) = self.model.get_navigation_end(sheet);
        Cell { row, column }
    }

    pub fn new_sheet(&mut self) {
        self.model.new_sheet();
    }

    pub fn rename_sheet(&mut self, sheet_index: i32, new_name: &str) -> JsResult {
        match self.model.rename_sheet_by_index(sheet_index, new_name) {
            Ok(()) => JsResult::get_success(),
            Err(message) => JsResult::get_error(&message),
        }
    }

    pub fn delete_sheet(&mut self, sheet_index: i32) -> JsResult {
        match self.model.delete_sheet(sheet_index) {
            Ok(()) => JsResult::get_success(),
            Err(message) => JsResult::get_error(&message),
        }
    }

    pub fn delete_sheet_by_name(&mut self, name: &str) -> JsResult {
        match self.model.delete_sheet_by_name(name) {
            Ok(()) => JsResult::get_success(),
            Err(message) => JsResult::get_error(&message),
        }
    }

    pub fn set_cell_style(
        &mut self,
        sheet: i32,
        row: i32,
        column: i32,
        style_js: &JsValue,
    ) -> JsResult {
        let style: Style = match JsValueSerdeExt::into_serde(style_js) {
            Ok(s) => s,
            Err(error) => return JsResult::get_error(&error.to_string()),
        };
        match self.model.set_cell_style(sheet, row, column, &style) {
            Ok(()) => JsResult::get_success(),
            Err(message) => JsResult::get_error(&message),
        }
    }

    pub fn set_cell_style_by_name(
        &mut self,
        sheet: i32,
        row: i32,
        column: i32,
        style_name: &str,
    ) -> JsResult {
        match self
            .model
            .set_cell_style_by_name(sheet, row, column, style_name)
        {
            Ok(()) => JsResult::get_success(),
            Err(message) => JsResult::get_error(&message),
        }
    }

    pub fn get_cell_style_index(&self, sheet: i32, row: i32, column: i32) -> i32 {
        self.model.get_cell_style_index(sheet, row, column)
    }

    pub fn shift_cells_right(
        &mut self,
        sheet: i32,
        row: i32,
        column: i32,
        cell_count: i32,
    ) -> JsResult {
        let result = self.model.shift_cells_right(sheet, row, column, cell_count);
        match result {
            Ok(()) => JsResult::get_success(),
            Err(message) => JsResult::get_error(message),
        }
    }

    pub fn shift_cells_left(
        &mut self,
        sheet: i32,
        row: i32,
        column: i32,
        cell_count: i32,
    ) -> JsResult {
        let result = self.model.shift_cells_left(sheet, row, column, cell_count);
        match result {
            Ok(()) => JsResult::get_success(),
            Err(message) => JsResult::get_error(message),
        }
    }

    pub fn shift_cells_down(
        &mut self,
        sheet: i32,
        row: i32,
        column: i32,
        cell_count: i32,
    ) -> JsResult {
        let result = self.model.shift_cells_down(sheet, row, column, cell_count);
        match result {
            Ok(()) => JsResult::get_success(),
            Err(message) => JsResult::get_error(message),
        }
    }

    pub fn shift_cells_up(
        &mut self,
        sheet: i32,
        row: i32,
        column: i32,
        cell_count: i32,
    ) -> JsResult {
        let result = self.model.shift_cells_up(sheet, row, column, cell_count);
        match result {
            Ok(()) => JsResult::get_success(),
            Err(message) => JsResult::get_error(message),
        }
    }

    pub fn insert_columns(&mut self, sheet: i32, column: i32, column_count: i32) -> JsResult {
        let result = self.model.insert_columns(sheet, column, column_count);
        match result {
            Ok(()) => JsResult::get_success(),
            Err(message) => JsResult::get_error(message),
        }
    }

    pub fn delete_columns(&mut self, sheet: i32, column: i32, column_count: i32) -> JsResult {
        let result = self.model.delete_columns(sheet, column, column_count);
        match result {
            Ok(()) => JsResult::get_success(),
            Err(message) => JsResult::get_error(message),
        }
    }

    pub fn insert_rows(&mut self, sheet: i32, row: i32, row_count: i32) -> JsResult {
        let result = self.model.insert_rows(sheet, row, row_count);
        match result {
            Ok(()) => JsResult::get_success(),
            Err(message) => JsResult::get_error(message),
        }
    }

    pub fn delete_rows(&mut self, sheet: i32, row: i32, row_count: i32) -> JsResult {
        let result = self.model.delete_rows(sheet, row, row_count);
        match result {
            Ok(()) => JsResult::get_success(),
            Err(message) => JsResult::get_error(message),
        }
    }

    pub fn is_row_read_only(&self, sheet: i32, row: i32) -> bool {
        self.model.is_row_read_only(sheet, row)
    }

    pub fn get_row_undo_data(&self, sheet: i32, row: i32) -> String {
        match self.model.get_row_undo_data(sheet, row) {
            Ok(s) => s,
            Err(_) => json!({"success": false}).to_string(),
        }
    }

    pub fn set_row_undo_data(&mut self, sheet: i32, row: i32, row_data_str: &str) -> JsResult {
        match self.model.set_row_undo_data(sheet, row, row_data_str) {
            Ok(_) => JsResult::get_success(),
            Err(message) => JsResult::get_error(&message),
        }
    }

    pub fn forward_references(
        &mut self,
        source_data: &str,
        target_sheet: i32,
        target_row: i32,
        target_column: i32,
    ) -> String {
        let source: Area = match serde_json::from_str(source_data) {
            Ok(v) => v,
            Err(e) => return json!({"success": false, "message": e.to_string()}).to_string(),
        };
        match self
            .model
            .forward_references(&source, target_sheet, target_row, target_column)
        {
            Ok(diff_list) => json!({"success": true, "diffList": diff_list}).to_string(),
            Err(e) => json!({"success": false, "error": e}).to_string(),
        }
    }

    pub fn move_cell_value_to_area(&mut self, move_data: &str) -> String {
        let data: MoveData = match serde_json::from_str(move_data) {
            Ok(v) => v,
            Err(e) => return json!({"success": false, "message": e.to_string()}).to_string(),
        };
        let source = CellReferenceIndex {
            sheet: data.source.sheet,
            column: data.source.column,
            row: data.source.row,
        };
        let target = CellReferenceIndex {
            sheet: data.target.sheet,
            column: data.target.column,
            row: data.target.row,
        };
        let value = data.value;
        let area = Area {
            sheet: data.area.sheet,
            row: data.area.row,
            column: data.area.column,
            width: data.area.width,
            height: data.area.height,
        };
        match self
            .model
            .move_cell_value_to_area(&source, &value, &target, &area)
        {
            Ok(s) => json!({"success": true, "value": s}).to_string(),
            Err(e) => json!({"success": false, "error": e}).to_string(),
        }
    }

    /// Takes 'value' and extends it from source to target
    pub fn extend_formula_to(&mut self, extend_data: &str) -> String {
        let data: ExtendToData = match serde_json::from_str(extend_data) {
            Ok(v) => v,
            Err(e) => return json!({"success": false, "message": e.to_string()}).to_string(),
        };
        let source = CellReferenceIndex {
            sheet: data.source.sheet,
            column: data.source.column,
            row: data.source.row,
        };
        let target = CellReferenceIndex {
            sheet: data.target.sheet,
            column: data.target.column,
            row: data.target.row,
        };
        let value = data.value;
        match self.model.extend_formula_to(&source, &value, &target) {
            Ok(s) => json!({"success": true, "value": s}).to_string(),
            Err(e) => json!({"success": false, "message": e}).to_string(),
        }
    }

    /// Returns the index of style if present otherwise creates the style and returns it's index
    pub fn get_style_index_or_create(&mut self, style_js: &JsValue) -> IndexResult {
        let style: Style = match JsValueSerdeExt::into_serde(style_js) {
            Ok(v) => v,
            Err(message) => return IndexResult::get_error(&message.to_string()),
        };
        if let Some(index) = self.model.get_style_index(&style) {
            IndexResult::get_success(index)
        } else {
            IndexResult::get_success(self.model.create_new_style(&style))
        }
    }

    pub fn set_sheet_color(&mut self, sheet: i32, color: &str) -> JsResult {
        match self.model.set_sheet_color(sheet, color) {
            Ok(()) => JsResult::get_success(),
            Err(message) => JsResult::get_error(&message),
        }
    }
}
