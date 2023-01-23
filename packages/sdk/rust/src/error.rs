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

impl From<XlsxError> for WorkbookError {
    fn from(error: XlsxError) -> Self {
        Self::new(error.to_string())
    }
}
