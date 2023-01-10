use roxmltree;
use std::io;
use std::num::{ParseFloatError, ParseIntError};
use thiserror::Error;
use zip::result::ZipError;

#[derive(Error, Debug)]
pub enum XlsxError {
    #[error("I/O Error: {0}")]
    IOError(String),
    #[error("Zip Error: {0}")]
    ZipError(String),
    #[error("XML Error: {0}")]
    XMLError(String),
}

impl From<io::Error> for XlsxError {
    fn from(error: io::Error) -> Self {
        XlsxError::IOError(error.to_string())
    }
}

impl From<ZipError> for XlsxError {
    fn from(error: ZipError) -> Self {
        XlsxError::ZipError(error.to_string())
    }
}

impl From<ParseIntError> for XlsxError {
    fn from(error: ParseIntError) -> Self {
        XlsxError::XMLError(error.to_string())
    }
}

impl From<ParseFloatError> for XlsxError {
    fn from(error: ParseFloatError) -> Self {
        XlsxError::XMLError(error.to_string())
    }
}

impl From<roxmltree::Error> for XlsxError {
    fn from(error: roxmltree::Error) -> Self {
        XlsxError::XMLError(error.to_string())
    }
}
