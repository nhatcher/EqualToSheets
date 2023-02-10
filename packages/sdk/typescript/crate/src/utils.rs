use equalto_calc::expressions::lexer::util::get_tokens;
use wasm_bindgen::{prelude::wasm_bindgen, JsError};

use crate::error::WorkbookError;

/// Return a JSON string with a list of all the tokens from a formula
/// This is used by the UI to color them according to a theme.
#[wasm_bindgen(js_name = "getFormulaTokens")]
#[allow(dead_code)] // code is not dead, for some reason wasm_bindgen doesn't mark it as in use
pub fn get_formula_tokens(formula: &str) -> Result<String, JsError> {
    let tokens = get_tokens(formula);
    serde_json::to_string(&tokens)
        .map_err(|serde_error| WorkbookError::PlainString(serde_error.to_string()))
        .map_err(JsError::from)
}
