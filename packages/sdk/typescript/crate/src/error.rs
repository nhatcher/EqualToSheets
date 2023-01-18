use wasm_bindgen::JsError;

#[cfg(feature = "xlsx")]
use equalto_xlsx::error::XlsxError;

#[derive(Debug, Clone)]
pub enum WorkbookError {
    PlainString(String),
    #[cfg(feature = "xlsx")]
    XlsxError(String),
}

impl From<String> for WorkbookError {
    fn from(error: String) -> WorkbookError {
        WorkbookError::PlainString(error)
    }
}

#[cfg(feature = "xlsx")]
impl From<XlsxError> for WorkbookError {
    fn from(error: XlsxError) -> WorkbookError {
        WorkbookError::XlsxError(error.to_string())
    }
}

impl From<WorkbookError> for JsError {
    fn from(error: WorkbookError) -> JsError {
        let (kind, description) = match error {
            WorkbookError::PlainString(description) => ("PlainString", description),
            #[cfg(feature = "xlsx")]
            WorkbookError::XlsxError(description) => ("XlsxError", description),
        };

        JsError::new(
            &serde_json::json!({
                "kind": kind,
                "description": description,
            })
            .to_string(),
        )
    }
}
