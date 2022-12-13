use crate::{
    expressions::token::Error, language::Language, number_format::to_excel_precision_str, types::*,
};
use serde::{Deserialize, Serialize};
use serde_json::json;

#[derive(Serialize, Deserialize, Debug, PartialEq)]
#[serde(untagged, deny_unknown_fields)]
pub enum UIValue {
    Text(String),
    Number(f64),
}
#[derive(Serialize, Deserialize, Debug, PartialEq)]
pub struct UICell {
    pub kind: String,
    pub value: UIValue,
    pub details: String,
}

impl Cell {
    /// Creates a new Cell with a shared string (`si` is the string index)
    pub fn new_string(si: i32, s: i32) -> Cell {
        Cell::SharedString {
            t: "s".to_string(),
            si,
            s,
        }
    }

    /// Creates a new Cell with a number
    pub fn new_number(v: f64, s: i32) -> Cell {
        Cell::NumberCell {
            t: "n".to_string(),
            v,
            s,
        }
    }

    /// Creates a new Cell with a boolean
    pub fn new_boolean(v: bool, s: i32) -> Cell {
        Cell::BooleanCell {
            t: "b".to_string(),
            v,
            s,
        }
    }

    /// Creates a new Cell with an error value
    pub fn new_error(ei: Error, s: i32) -> Cell {
        Cell::ErrorCell {
            t: "e".to_string(),
            ei,
            s,
        }
    }

    /// Creates a new Cell with an unevaluated formula
    pub fn new_formula(f: i32, s: i32) -> Cell {
        Cell::CellFormula {
            t: "u".to_string(),
            f,
            s,
        }
    }

    /// Returns the formula of a cell if any.
    pub fn get_formula(&self) -> Option<i32> {
        match self {
            Cell::CellFormula { f, .. } => Some(*f),
            Cell::CellFormulaBoolean { f, .. } => Some(*f),
            Cell::CellFormulaNumber { f, .. } => Some(*f),
            Cell::CellFormulaString { f, .. } => Some(*f),
            Cell::CellFormulaError { f, .. } => Some(*f),
            _ => None,
        }
    }

    pub fn set_style(&mut self, style: i32) {
        match self {
            Cell::EmptyCell { s, .. } => *s = style,
            Cell::BooleanCell { s, .. } => *s = style,
            Cell::NumberCell { s, .. } => *s = style,
            Cell::ErrorCell { s, .. } => *s = style,
            Cell::SharedString { s, .. } => *s = style,
            Cell::CellFormula { s, .. } => *s = style,
            Cell::CellFormulaBoolean { s, .. } => *s = style,
            Cell::CellFormulaNumber { s, .. } => *s = style,
            Cell::CellFormulaString { s, .. } => *s = style,
            Cell::CellFormulaError { s, .. } => *s = style,
        };
    }

    pub fn get_style(&self) -> i32 {
        match self {
            Cell::EmptyCell { s, .. } => *s,
            Cell::BooleanCell { s, .. } => *s,
            Cell::NumberCell { s, .. } => *s,
            Cell::ErrorCell { s, .. } => *s,
            Cell::SharedString { s, .. } => *s,
            Cell::CellFormula { s, .. } => *s,
            Cell::CellFormulaBoolean { s, .. } => *s,
            Cell::CellFormulaNumber { s, .. } => *s,
            Cell::CellFormulaString { s, .. } => *s,
            Cell::CellFormulaError { s, .. } => *s,
        }
    }

    /// Return a JSON string representation of the Cell
    pub fn to_json(&self) -> String {
        match self {
            Cell::EmptyCell { t, s } => json!({"t":t, "s":s}).to_string(),
            Cell::BooleanCell { t, v, s } => json!({"t":t, "v": v,"s":s}).to_string(),
            Cell::NumberCell { t, v, s } => json!({"t":t, "v": v,"s":s}).to_string(),
            Cell::ErrorCell { t, ei, s } => json!({"t":t, "ei": ei,"s":s}).to_string(),
            Cell::SharedString { t, si, s } => json!({"t":t, "si": si,"s":s}).to_string(),
            Cell::CellFormula { t, f, s } => json!({"t":t, "f": f,"s":s}).to_string(),
            Cell::CellFormulaBoolean { t, f, v, s } => {
                json!({"t":t, "f": f, "v": v,"s":s}).to_string()
            }
            Cell::CellFormulaNumber { t, f, v, s } => {
                json!({"t":t, "f": f, "v": v,"s":s}).to_string()
            }
            Cell::CellFormulaString { t, f, v, s } => {
                json!({"t":t, "f":f, "v": v,"s":s}).to_string()
            }
            Cell::CellFormulaError { t, f, ei, s, o, m } => {
                json!({"t":t, "f": f, "ei": ei, "o":o, "m":m,"s":s}).to_string()
            }
        }
    }

    pub fn get_text(&self, shared_strings: &[String], language: &Language) -> String {
        match self {
            Cell::CellFormula { f: _, .. } => "".to_string(),
            Cell::CellFormulaBoolean { v, .. } => v.to_string().to_uppercase(),
            Cell::CellFormulaNumber { v, .. } => to_excel_precision_str(*v),
            Cell::CellFormulaString { v, .. } => v.clone(),
            Cell::CellFormulaError { ei, .. } => ei.to_localized_error_string(language),
            Cell::EmptyCell { t: _, s: _ } => "".to_string(),
            Cell::BooleanCell { v, .. } => v.to_string().to_uppercase(),
            Cell::NumberCell { v, .. } => to_excel_precision_str(*v),
            Cell::ErrorCell { ei, .. } => ei.to_localized_error_string(language),
            Cell::SharedString { si, .. } => {
                let s = shared_strings.get(*si as usize);
                match s {
                    Some(str) => str.clone(),
                    None => "".to_string(),
                }
            }
        }
    }

    pub fn get_ui_cell(&self, shared_strings: &[String], language: &Language) -> UICell {
        match self {
            Cell::CellFormula { f: _, .. } => UICell {
                kind: "formula".to_string(),
                value: UIValue::Text("".to_string()),
                details: "".to_string(),
            },
            Cell::CellFormulaBoolean { v, .. } => UICell {
                kind: "bool".to_string(),
                value: UIValue::Text(v.to_string().to_uppercase()),
                details: "".to_string(),
            },
            Cell::CellFormulaNumber { v, .. } => UICell {
                kind: "number".to_string(),
                value: UIValue::Number(*v),
                details: "".to_string(),
            },
            Cell::CellFormulaString { v, .. } => UICell {
                kind: "text".to_string(),
                value: UIValue::Text(v.clone()),
                details: "".to_string(),
            },
            Cell::CellFormulaError { ei, .. } => UICell {
                kind: "error".to_string(),
                value: UIValue::Text(ei.to_localized_error_string(language)),
                details: "".to_string(),
            },
            Cell::EmptyCell { t: _, s: _ } => UICell {
                kind: "empty".to_string(),
                value: UIValue::Text("".to_string()),
                details: "".to_string(),
            },
            Cell::BooleanCell { v, .. } => UICell {
                kind: "bool".to_string(),
                value: UIValue::Text(v.to_string().to_uppercase()),
                details: "".to_string(),
            },
            Cell::NumberCell { v, .. } => UICell {
                kind: "number".to_string(),
                value: UIValue::Number(*v),
                details: "".to_string(),
            },
            Cell::ErrorCell { ei, .. } => UICell {
                kind: "error".to_string(),
                value: UIValue::Text(ei.to_localized_error_string(language)),
                details: "".to_string(),
            },
            Cell::SharedString { si, .. } => {
                let s = shared_strings.get(*si as usize);
                match s {
                    Some(str) => UICell {
                        kind: "text".to_string(),
                        value: UIValue::Text(str.clone()),
                        details: "".to_string(),
                    },
                    None => UICell {
                        kind: "text".to_string(),
                        value: UIValue::Text("".to_string()),
                        details: "".to_string(),
                    },
                }
            }
        }
    }
}
