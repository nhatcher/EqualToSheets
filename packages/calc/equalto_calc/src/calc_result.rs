use std::cmp::Ordering;

use crate::expressions::token::Error;

#[derive(Clone, PartialEq, Eq)]
pub struct CellReference {
    pub sheet: i32,
    pub column: i32,
    pub row: i32,
}

#[derive(Clone)]
pub struct Range {
    pub left: CellReference,
    pub right: CellReference,
}

#[derive(Clone)]
pub(crate) enum CalcResult {
    String(String),
    Number(f64),
    Boolean(bool),
    Error {
        error: Error,
        origin: CellReference,
        message: String,
    },
    Range {
        left: CellReference,
        right: CellReference,
    },
    EmptyCell,
    EmptyArg,
}

impl CalcResult {
    pub fn new_error(
        error: Error,
        sheet: i32,
        row: i32,
        column: i32,
        message: String,
    ) -> CalcResult {
        CalcResult::Error {
            error,
            origin: CellReference { sheet, column, row },
            message,
        }
    }
    pub fn is_error(&self) -> bool {
        matches!(self, CalcResult::Error { .. })
    }
}

impl Ord for CalcResult {
    // ..., -2, -1, 0, 1, 2, ..., A-Z, FALSE, TRUE, empty;
    fn cmp(&self, other: &Self) -> Ordering {
        match (self, other) {
            (CalcResult::Number(value1), CalcResult::Number(value2)) => {
                if (value2 - value1).abs() < f64::EPSILON {
                    return Ordering::Equal;
                }
                if value1 < value2 {
                    return Ordering::Less;
                }
                Ordering::Greater
            }
            (CalcResult::Number(_value1), CalcResult::String(_value2)) => Ordering::Less,
            (CalcResult::Number(_value1), CalcResult::Boolean(_value2)) => Ordering::Less,
            (CalcResult::String(value1), CalcResult::String(value2)) => {
                let value1 = value1.to_uppercase();
                let value2 = value2.to_uppercase();
                value1.cmp(&value2)
            }
            (CalcResult::String(_value1), CalcResult::Boolean(_value2)) => Ordering::Less,
            (CalcResult::Boolean(value1), CalcResult::Boolean(value2)) => {
                if value1 == value2 {
                    return Ordering::Equal;
                }
                if *value1 {
                    return Ordering::Greater;
                }
                Ordering::Less
            }
            (CalcResult::EmptyCell, CalcResult::String(_value2)) => Ordering::Greater,
            (CalcResult::EmptyArg, CalcResult::String(_value2)) => Ordering::Greater,
            (CalcResult::String(_value1), CalcResult::EmptyCell) => Ordering::Less,
            (CalcResult::EmptyCell, CalcResult::Number(_value2)) => Ordering::Greater,
            (CalcResult::EmptyArg, CalcResult::Number(_value2)) => Ordering::Greater,
            (CalcResult::Number(_value1), CalcResult::EmptyCell) => Ordering::Less,
            (CalcResult::EmptyCell, CalcResult::EmptyCell) => Ordering::Equal,
            (CalcResult::EmptyCell, CalcResult::EmptyArg) => Ordering::Equal,
            (CalcResult::EmptyArg, CalcResult::EmptyCell) => Ordering::Equal,
            (CalcResult::EmptyArg, CalcResult::EmptyArg) => Ordering::Equal,
            // NOTE: Errors and Ranges are not covered
            (_, _) => Ordering::Greater,
        }
    }
}

impl PartialOrd for CalcResult {
    fn partial_cmp(&self, other: &Self) -> Option<Ordering> {
        Some(self.cmp(other))
    }
}

impl PartialEq for CalcResult {
    fn eq(&self, other: &Self) -> bool {
        self.cmp(other) == Ordering::Equal
    }
}

impl Eq for CalcResult {}
