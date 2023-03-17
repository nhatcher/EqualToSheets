use equalto_xlsx::error::XlsxError;
use thiserror::Error;

#[derive(Error, Debug, PartialEq, Eq)]
#[error("{message}")]
pub struct WorkbookError {
    pub message: String,
}

impl WorkbookError {
    pub fn new(message: String) -> Self {
        Self { message }
    }
}

impl From<String> for WorkbookError {
    fn from(message: String) -> Self {
        Self::new(message)
    }
}

impl From<Vec<String>> for WorkbookError {
    fn from(errors: Vec<String>) -> Self {
        // We just use the first error from the list
        let error = errors.first().expect("empty error list").to_string();
        Self::new(error)
    }
}

impl From<XlsxError> for WorkbookError {
    fn from(error: XlsxError) -> Self {
        Self::new(error.to_string())
    }
}
