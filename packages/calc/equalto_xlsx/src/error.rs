use std::io;
use std::num::{ParseFloatError, ParseIntError};
use thiserror::Error;
use zip::result::ZipError;

#[derive(Error, Debug, PartialEq, Eq)]
pub enum XlsxError {
    #[error("I/O Error: {0}")]
    IO(String),
    #[error("Zip Error: {0}")]
    Zip(String),
    #[error("XML Error: {0}")]
    Xml(String),
    #[error("{0}")]
    Workbook(String),
}

impl From<io::Error> for XlsxError {
    fn from(error: io::Error) -> Self {
        XlsxError::IO(error.to_string())
    }
}

impl From<ZipError> for XlsxError {
    fn from(error: ZipError) -> Self {
        XlsxError::Zip(error.to_string())
    }
}

impl From<ParseIntError> for XlsxError {
    fn from(error: ParseIntError) -> Self {
        XlsxError::Xml(error.to_string())
    }
}

impl From<ParseFloatError> for XlsxError {
    fn from(error: ParseFloatError) -> Self {
        XlsxError::Xml(error.to_string())
    }
}

impl From<roxmltree::Error> for XlsxError {
    fn from(error: roxmltree::Error) -> Self {
        XlsxError::Xml(error.to_string())
    }
}
