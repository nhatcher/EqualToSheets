use js_sys::Date;
use wasm_bindgen::prelude::{wasm_bindgen, JsError};

use equalto_calc::model::{Environment, Model};

#[cfg(feature = "xlsx")]
use equalto_xlsx::import::load_xlsx_from_memory;

use crate::error::WorkbookError;

#[wasm_bindgen]
pub struct WasmWorkbook {
    model: Model,
}

#[wasm_bindgen]
impl WasmWorkbook {
    #[wasm_bindgen(constructor)]
    pub fn new(locale: &str, timezone: &str) -> Result<WasmWorkbook, JsError> {
        let env = Environment {
            get_milliseconds_since_epoch,
        };
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
        let env = Environment {
            get_milliseconds_since_epoch,
        };
        let model = load_xlsx_from_memory("workbook", data, locale, timezone, env)
            .map_err(WorkbookError::from)?;
        Ok(WasmWorkbook { model })
    }

    #[wasm_bindgen(js_name = "setInput")]
    pub fn set_input(
        &mut self,
        sheet: u32,
        row: i32,
        column: i32,
        value: String,
        style: i32,
    ) -> Result<(), JsError> {
        self.model.set_input(sheet, row, column, value, style);
        Ok(())
    }

    #[wasm_bindgen(js_name=getFormattedCellValue)]
    pub fn formatted_cell_value(
        &self,
        sheet: u32,
        row: i32,
        column: i32,
    ) -> Result<String, JsError> {
        Ok(self
            .model
            .formatted_cell_value(sheet, row, column)
            .map_err(WorkbookError::from)?)
    }

    pub fn evaluate(&mut self) -> Result<(), JsError> {
        self.model.evaluate();
        Ok(())
    }
}

fn get_milliseconds_since_epoch() -> i64 {
    Date::now() as i64
}
