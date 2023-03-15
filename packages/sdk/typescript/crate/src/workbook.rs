use wasm_bindgen::{
    prelude::{wasm_bindgen, JsError},
    JsValue,
};

use equalto_calc::{
    cell::CellValue,
    expressions::types::{Area, CellReferenceIndex},
    model::{Environment, Model},
    worksheet::NavigationDirection,
};

#[cfg(feature = "xlsx")]
use equalto_xlsx::import::load_xlsx_from_memory;

#[cfg(feature = "xlsx")]
use equalto_xlsx::export::save_xlsx_to_writer;

use crate::error::WorkbookError;

#[wasm_bindgen]
pub enum WasmNavigationDirection {
    Left,
    Right,
    Up,
    Down,
}

#[wasm_bindgen]
pub struct WasmLocalCellCoordinate {
    pub row: i32,
    pub column: i32,
}

impl From<WasmNavigationDirection> for NavigationDirection {
    fn from(value: WasmNavigationDirection) -> Self {
        match value {
            WasmNavigationDirection::Left => NavigationDirection::Left,
            WasmNavigationDirection::Right => NavigationDirection::Right,
            WasmNavigationDirection::Up => NavigationDirection::Up,
            WasmNavigationDirection::Down => NavigationDirection::Down,
        }
    }
}

#[wasm_bindgen]
pub struct WasmCellReferenceIndex {
    pub sheet: u32,
    pub row: i32,
    pub column: i32,
}

#[wasm_bindgen]
impl WasmCellReferenceIndex {
    #[wasm_bindgen(constructor)]
    pub fn new(sheet: u32, row: i32, column: i32) -> Self {
        Self { sheet, row, column }
    }
}

impl From<WasmCellReferenceIndex> for CellReferenceIndex {
    fn from(value: WasmCellReferenceIndex) -> Self {
        CellReferenceIndex {
            sheet: value.sheet,
            row: value.row,
            column: value.column,
        }
    }
}

#[wasm_bindgen]
pub struct WasmArea {
    pub sheet: u32,
    pub row: i32,
    pub column: i32,
    pub width: i32,
    pub height: i32,
}

#[wasm_bindgen]
impl WasmArea {
    #[wasm_bindgen(constructor)]
    pub fn new(sheet: u32, row: i32, column: i32, width: i32, height: i32) -> Self {
        Self {
            sheet,
            row,
            column,
            width,
            height,
        }
    }
}

impl From<WasmArea> for Area {
    fn from(value: WasmArea) -> Self {
        Area {
            sheet: value.sheet,
            row: value.row,
            column: value.column,
            width: value.width,
            height: value.height,
        }
    }
}

#[wasm_bindgen(js_name = "SheetDimension")]
pub struct WasmSheetDimension {
    #[wasm_bindgen(js_name = "minColumn")]
    pub min_column: i32,
    #[wasm_bindgen(js_name = "maxColumn")]
    pub max_column: i32,
    #[wasm_bindgen(js_name = "minRow")]
    pub min_row: i32,
    #[wasm_bindgen(js_name = "maxRow")]
    pub max_row: i32,
}

#[wasm_bindgen]
pub struct WasmWorkbook {
    model: Model,
}

#[wasm_bindgen]
impl WasmWorkbook {
    #[wasm_bindgen(constructor)]
    pub fn new(locale: &str, timezone: &str) -> Result<WasmWorkbook, JsError> {
        let env = Environment::default();
        let model =
            Model::new_empty("workbook", locale, timezone, env).map_err(WorkbookError::from)?;
        Ok(WasmWorkbook { model })
    }

    #[wasm_bindgen(js_name=loadFromMemory)]
    #[cfg(feature = "xlsx")]
    pub fn load_from_memory(
        data: &mut [u8],
        locale: &str,
        timezone: &str,
    ) -> Result<WasmWorkbook, JsError> {
        let env = Environment::default();
        let model = load_xlsx_from_memory("workbook", data, locale, timezone, env)
            .map_err(WorkbookError::from)?;
        Ok(WasmWorkbook { model })
    }

    #[wasm_bindgen(js_name=loadFromJson)]
    pub fn load_from_json(workbook_json: &str) -> Result<WasmWorkbook, JsError> {
        let env = Environment::default();
        let model = Model::from_json(workbook_json, env).map_err(WorkbookError::from)?;
        Ok(WasmWorkbook { model })
    }

    #[wasm_bindgen(js_name=saveToMemory)]
    #[cfg(feature = "xlsx")]
    pub fn save_xlsx_to_memory(&self) -> Result<js_sys::Uint8Array, JsError> {
        use js_sys::Uint8Array;
        use std::io::Cursor;

        let memory_buffer = Vec::new();
        let memory_writer = Cursor::new(memory_buffer);
        let memory_writer = save_xlsx_to_writer(&self.model, memory_writer)?;
        let memory_buffer = memory_writer.into_inner();

        let byte_array = Uint8Array::new_with_length(memory_buffer.len() as u32);
        byte_array.copy_from(&memory_buffer);

        Ok(byte_array)
    }

    pub fn evaluate(&mut self) -> Result<(), JsError> {
        self.model.evaluate();
        Ok(())
    }

    #[wasm_bindgen(js_name = "getWorksheetNames")]
    pub fn get_worksheet_names(&self) -> Result<js_sys::Array, JsError> {
        Ok(self
            .model
            .workbook
            .get_worksheet_names()
            .into_iter()
            .map(JsValue::from)
            .collect())
    }

    #[wasm_bindgen(js_name = "getWorksheetIds")]
    pub fn get_worksheet_ids(&self) -> Result<js_sys::Array, JsError> {
        Ok(self
            .model
            .workbook
            .get_worksheet_ids()
            .into_iter()
            .map(JsValue::from)
            .collect())
    }

    #[wasm_bindgen(js_name = "addSheet")]
    pub fn add_sheet(&mut self, name: &str) -> Result<(), JsError> {
        self.model
            .add_sheet(name)
            .map_err(WorkbookError::from)
            .map_err(JsError::from)
    }

    #[wasm_bindgen(js_name = "newSheet")]
    pub fn new_sheet(&mut self) -> Result<(), JsError> {
        self.model.new_sheet();
        Ok(())
    }

    // TODO: Should be by sheetId
    #[wasm_bindgen(js_name = "renameSheetBySheetIndex")]
    pub fn rename_sheet_by_sheet_index(
        &mut self,
        sheet: i32,
        new_name: &str,
    ) -> Result<(), JsError> {
        self.model
            .rename_sheet_by_index(sheet.try_into().unwrap(), new_name)
            .map_err(WorkbookError::from)
            .map_err(JsError::from)
    }

    #[wasm_bindgen(js_name = "deleteSheetBySheetId")]
    pub fn delete_sheet_by_sheet_id(&mut self, sheet_id: i32) -> Result<(), JsError> {
        self.model
            .delete_sheet_by_sheet_id(sheet_id.try_into().unwrap())
            .map_err(WorkbookError::from)
            .map_err(JsError::from)
    }

    #[wasm_bindgen(js_name = "getCellValueByIndex")]
    pub fn get_cell_value_by_index(
        &self,
        sheet: i32,
        row: i32,
        column: i32,
    ) -> Result<JsValue, JsError> {
        Ok(
            match self
                .model
                .get_cell_value_by_index(sheet.try_into().unwrap(), row, column)
                .map_err(WorkbookError::from)?
            {
                CellValue::String(s) => JsValue::from(s),
                CellValue::Number(f) => JsValue::from(f),
                CellValue::Boolean(b) => JsValue::from(b),
            },
        )
    }

    #[wasm_bindgen(js_name = "getFormattedCellValue")]
    pub fn formatted_cell_value(
        &self,
        sheet: i32,
        row: i32,
        column: i32,
    ) -> Result<String, JsError> {
        self.model
            .formatted_cell_value(sheet.try_into().unwrap(), row, column)
            .map_err(WorkbookError::from)
            .map_err(JsError::from)
    }

    #[wasm_bindgen(js_name = "updateCellWithText")]
    pub fn update_cell_with_text(
        &mut self,
        sheet: i32,
        row: i32,
        column: i32,
        value: &str,
    ) -> Result<(), JsError> {
        self.model
            .update_cell_with_text(sheet.try_into().unwrap(), row, column, value);
        Ok(())
    }

    #[wasm_bindgen(js_name = "updateCellWithNumber")]
    pub fn update_cell_with_number(
        &mut self,
        sheet: i32,
        row: i32,
        column: i32,
        value: f64,
    ) -> Result<(), JsError> {
        self.model
            .update_cell_with_number(sheet.try_into().unwrap(), row, column, value);
        Ok(())
    }

    #[wasm_bindgen(js_name = "updateCellWithBool")]
    pub fn update_cell_with_bool(
        &mut self,
        sheet: i32,
        row: i32,
        column: i32,
        value: bool,
    ) -> Result<(), JsError> {
        self.model
            .update_cell_with_bool(sheet.try_into().unwrap(), row, column, value);
        Ok(())
    }

    #[wasm_bindgen(js_name = "getCellFormula")]
    pub fn cell_formula(
        &self,
        sheet: i32,
        row: i32,
        column: i32,
    ) -> Result<Option<String>, JsError> {
        self.model
            .cell_formula(sheet.try_into().unwrap(), row, column)
            .map_err(WorkbookError::from)
            .map_err(JsError::from)
    }

    #[wasm_bindgen(js_name = "updateCellWithFormula")]
    pub fn update_cell_with_formula(
        &mut self,
        sheet: i32,
        row: i32,
        column: i32,
        formula: String,
    ) -> Result<(), JsError> {
        self.model
            .update_cell_with_formula(sheet.try_into().unwrap(), row, column, formula)
            .map_err(WorkbookError::from)
            .map_err(JsError::from)
    }

    #[wasm_bindgen(js_name = "setUserInput")]
    pub fn set_user_input(
        &mut self,
        sheet: i32,
        row: i32,
        column: i32,
        input: String,
    ) -> Result<(), JsError> {
        self.model
            .set_user_input(sheet.try_into().unwrap(), row, column, input);
        // FIXME: set_user_input should return result
        Ok(())
    }

    #[wasm_bindgen(js_name = "setCellEmpty")]
    pub fn set_cell_empty(&mut self, sheet: i32, row: i32, column: i32) -> Result<(), JsError> {
        self.model
            .set_cell_empty(sheet.try_into().unwrap(), row, column)
            .map_err(WorkbookError::from)
            .map_err(JsError::from)
    }

    #[wasm_bindgen(js_name = "deleteCell")]
    pub fn delete_cell(&mut self, sheet: i32, row: i32, column: i32) -> Result<(), JsError> {
        self.model
            .delete_cell(sheet.try_into().unwrap(), row, column)
            .map_err(WorkbookError::from)
            .map_err(JsError::from)
    }

    #[wasm_bindgen(js_name = "getColumnWidth")]
    pub fn column_width(&self, sheet_index: u32, column: i32) -> Result<f64, JsError> {
        self.model
            .workbook
            .worksheet(sheet_index)
            .map_err(WorkbookError::from)?
            .column_width(column)
            .map_err(WorkbookError::from)
            .map_err(JsError::from)
    }

    #[wasm_bindgen(js_name = "getColumnCellReferences")]
    pub fn column_cell_references(&self, sheet_index: u32, column: i32) -> Result<String, JsError> {
        let cell_references = self
            .model
            .workbook
            .worksheet(sheet_index)
            .map_err(WorkbookError::from)?
            .column_cell_references(column)
            .map_err(WorkbookError::from)?;
        Ok(serde_json::to_string(&cell_references)
            .map_err(|_| "Could not stringify style to JSON.".to_string())
            .map_err(WorkbookError::from)?)
    }

    #[wasm_bindgen(js_name = "getRowCellReferences")]
    pub fn row_cell_references(&self, sheet_index: u32, row: i32) -> Result<String, JsError> {
        let cell_references = self
            .model
            .workbook
            .worksheet(sheet_index)
            .map_err(WorkbookError::from)?
            .row_cell_references(row)
            .map_err(WorkbookError::from)?;
        Ok(serde_json::to_string(&cell_references)
            .map_err(|_| "Could not stringify style to JSON.".to_string())
            .map_err(WorkbookError::from)?)
    }

    #[wasm_bindgen(js_name = "getCellReferences")]
    pub fn cell_references(&self, sheet_index: u32) -> Result<String, JsError> {
        let cell_references = self
            .model
            .workbook
            .worksheet(sheet_index)
            .map_err(WorkbookError::from)?
            .cell_references()
            .map_err(WorkbookError::from)?;
        Ok(serde_json::to_string(&cell_references)
            .map_err(|_| "Could not stringify style to JSON.".to_string())
            .map_err(WorkbookError::from)?)
    }

    #[wasm_bindgen(js_name = "getRowHeight")]
    pub fn row_height(&self, sheet_index: u32, row: i32) -> Result<f64, JsError> {
        self.model
            .workbook
            .worksheet(sheet_index)
            .map_err(WorkbookError::from)?
            .row_height(row)
            .map_err(WorkbookError::from)
            .map_err(JsError::from)
    }

    #[wasm_bindgen(js_name = "setColumnWidth")]
    pub fn set_column_width(
        &mut self,
        sheet_index: u32,
        column: i32,
        width: f64,
    ) -> Result<(), JsError> {
        self.model
            .workbook
            .worksheet_mut(sheet_index)
            .map_err(WorkbookError::from)?
            .set_column_width(column, width)
            .map_err(WorkbookError::from)
            .map_err(JsError::from)
    }

    #[wasm_bindgen(js_name = "setRowHeight")]
    pub fn set_row_height(
        &mut self,
        sheet_index: u32,
        row: i32,
        height: f64,
    ) -> Result<(), JsError> {
        self.model
            .workbook
            .worksheet_mut(sheet_index)
            .map_err(WorkbookError::from)?
            .set_row_height(row, height)
            .map_err(WorkbookError::from)
            .map_err(JsError::from)
    }

    #[wasm_bindgen(js_name = "getSheetDimensions")]
    pub fn sheet_dimensions(&self, sheet_index: u32) -> Result<WasmSheetDimension, JsError> {
        let dimension = self
            .model
            .workbook
            .worksheet(sheet_index)
            .map_err(WorkbookError::from)?
            .dimension();

        Ok(WasmSheetDimension {
            min_row: dimension.min_row,
            min_column: dimension.min_column,
            max_row: dimension.max_row,
            max_column: dimension.max_column,
        })
    }

    #[wasm_bindgen(js_name = "navigateToEdgeInDirection")]
    pub fn navigate_to_edge_in_direction(
        &self,
        sheet_index: u32,
        row: i32,
        column: i32,
        direction: WasmNavigationDirection,
    ) -> Result<WasmLocalCellCoordinate, JsError> {
        let (row, column) = self
            .model
            .workbook
            .worksheet(sheet_index)
            .map_err(WorkbookError::from)?
            .navigate_to_edge_in_direction(row, column, direction.into())
            .map_err(WorkbookError::from)?;
        Ok(WasmLocalCellCoordinate { row, column })
    }

    #[wasm_bindgen(js_name = "getExtendedValue")]
    pub fn get_extended_value(
        &self,
        sheet_index: u32,
        source_row: i32,
        source_column: i32,
        target_row: i32,
        target_column: i32,
    ) -> Result<String, JsError> {
        Ok(self
            .model
            .extend_to(
                sheet_index,
                source_row,
                source_column,
                target_row,
                target_column,
            )
            .map_err(WorkbookError::from)?)
    }

    #[wasm_bindgen(js_name = "getCopiedValueExtended")]
    pub fn get_copied_value_extended(
        &mut self,
        value: &str,
        source_sheet_name: &str,
        source: WasmCellReferenceIndex,
        target: WasmCellReferenceIndex,
    ) -> Result<String, JsError> {
        Ok(self
            .model
            .extend_copied_value(value, source_sheet_name, &source.into(), &target.into())
            .map_err(WorkbookError::from)?)
    }

    #[wasm_bindgen(js_name = "getCutValueMoved")]
    pub fn get_cut_value_moved(
        &mut self,
        value: &str,
        source: WasmCellReferenceIndex,
        target: WasmCellReferenceIndex,
        source_area: WasmArea,
    ) -> Result<String, JsError> {
        Ok(self
            .model
            .move_cell_value_to_area(value, &source.into(), &target.into(), &source_area.into())
            .map_err(WorkbookError::from)?)
    }

    #[wasm_bindgen(js_name = "forwardReferences")]
    pub fn forward_references(
        &mut self,
        source_area: WasmArea,
        target: WasmCellReferenceIndex,
    ) -> Result<(), JsError> {
        self.model
            .forward_references(&source_area.into(), &target.into())
            .map_err(WorkbookError::from)?;
        Ok(())
    }

    #[wasm_bindgen(js_name = "getCellStyle")]
    pub fn get_cell_style(
        &self,
        sheet_index: u32,
        row: i32,
        column: i32,
    ) -> Result<String, JsError> {
        Ok(
            serde_json::to_string(&self.model.get_style_for_cell(sheet_index, row, column))
                .map_err(|_| "Could not stringify style to JSON.".to_string())
                .map_err(WorkbookError::from)?,
        )
    }

    #[wasm_bindgen(js_name = "setCellStyle")]
    pub fn set_cell_style(
        &mut self,
        sheet_index: u32,
        row: i32,
        column: i32,
        style: &str,
    ) -> Result<(), JsError> {
        self.model
            .set_cell_style(
                sheet_index,
                row,
                column,
                &serde_json::from_str(style)
                    .map_err(|_| "Could not parse data transfer blob for style.".to_string())
                    .map_err(WorkbookError::from)?,
            )
            .map_err(WorkbookError::from)
            .map_err(JsError::from)
    }

    #[wasm_bindgen(js_name = "insertRows")]
    pub fn insert_rows(
        &mut self,
        sheet_index: u32,
        row: i32,
        row_count: i32,
    ) -> Result<(), JsError> {
        self.model
            .insert_rows(sheet_index, row, row_count)
            .map_err(WorkbookError::from)
            .map_err(JsError::from)
    }

    #[wasm_bindgen(js_name = "deleteRows")]
    pub fn delete_rows(
        &mut self,
        sheet_index: u32,
        row: i32,
        row_count: i32,
    ) -> Result<(), JsError> {
        self.model
            .delete_rows(sheet_index, row, row_count)
            .map_err(WorkbookError::from)
            .map_err(JsError::from)
    }

    #[wasm_bindgen(js_name = "insertColumns")]
    pub fn insert_columns(
        &mut self,
        sheet_index: u32,
        column: i32,
        column_count: i32,
    ) -> Result<(), JsError> {
        self.model
            .insert_columns(sheet_index, column, column_count)
            .map_err(WorkbookError::from)
            .map_err(JsError::from)
    }

    #[wasm_bindgen(js_name = "deleteColumns")]
    pub fn delete_columns(
        &mut self,
        sheet_index: u32,
        column: i32,
        column_count: i32,
    ) -> Result<(), JsError> {
        self.model
            .delete_columns(sheet_index, column, column_count)
            .map_err(WorkbookError::from)
            .map_err(JsError::from)
    }

    #[wasm_bindgen(js_name = "copyCellStyle")]
    pub fn copy_cell_style(
        &mut self,
        source_sheet_index: u32,
        source_row: i32,
        source_column: i32,
        destination_sheet_index: u32,
        destination_source_row: i32,
        destination_source_column: i32,
    ) -> Result<(), JsError> {
        self.model
            .copy_cell_style(
                (source_sheet_index, source_row, source_column),
                (
                    destination_sheet_index,
                    destination_source_row,
                    destination_source_column,
                ),
            )
            .map_err(WorkbookError::from)
            .map_err(JsError::from)
    }

    #[wasm_bindgen(js_name = "toJson")]
    pub fn to_json(&self) -> Result<String, JsError> {
        Ok(self.model.to_json_str())
    }
}
