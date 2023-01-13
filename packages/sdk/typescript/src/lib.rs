use gloo_utils::format::JsValueSerdeExt;
use js_sys::Date;
use wasm_bindgen::prelude::{wasm_bindgen, JsValue};

use equalto_calc::model::{Environment, Model};

#[cfg(feature = "xlsx")]
use equalto_xlsx::import::load_xlsx_from_memory;

fn get_milliseconds_since_epoch() -> i64 {
    Date::now() as i64
}

#[wasm_bindgen]
pub struct JSModel {
    model: Model,
}

#[wasm_bindgen]
impl JSModel {
    #[wasm_bindgen(js_name=newFromJson)]
    pub fn new_from_json(val: &JsValue) -> JSModel {
        let env = Environment {
            get_milliseconds_since_epoch,
        };
        let data: String = JsValueSerdeExt::into_serde(val).unwrap();
        let model = Model::from_json(&data, env).unwrap();
        JSModel { model }
    }

    #[wasm_bindgen(js_name=newFromExcelFile)]
    #[cfg(feature = "xlsx")]
    pub fn new_from_excel_file(
        name: &str,
        data: &mut [u8],
        locale: &str,
        timezone: &str,
    ) -> JSModel {
        let env = Environment {
            get_milliseconds_since_epoch,
        };
        let model = load_xlsx_from_memory(name, data, locale, timezone, env).unwrap();
        JSModel { model }
    }

    #[wasm_bindgen(js_name=newEmpty)]
    pub fn new_empty(name: &str, locale: &str, timezone: &str) -> JSModel {
        let env = Environment {
            get_milliseconds_since_epoch,
        };
        let model = Model::new_empty(name, locale, timezone, env).unwrap();
        JSModel { model }
    }

    pub fn set_input(&mut self, sheet: u32, row: i32, column: i32, value: String, style: i32) {
        self.model.set_input(sheet, row, column, value, style)
    }

    pub fn get_text_at(&self, sheet: u32, row: i32, column: i32) -> String {
        self.model.get_text_at(sheet, row, column)
    }

    pub fn evaluate(&mut self) {
        self.model.evaluate()
    }
}
