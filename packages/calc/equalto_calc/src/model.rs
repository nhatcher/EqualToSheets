use chrono_tz::Tz;
use serde::{Deserialize, Serialize};
use serde_json::json;

use std::collections::HashMap;
use std::vec::Vec;

use crate::{
    calc_result::{CalcResult, CellReference, Range},
    cell::{UICell, UIValue},
    expressions::parser::{
        stringify::{to_rc_format, to_string, to_string_displaced, DisplaceData},
        Node, Parser,
    },
    expressions::{
        lexer::Lexer,
        parser::move_formula::{move_formula, MoveContext},
        token::{self, get_error_by_name},
        types::*,
        utils,
    },
    expressions::{
        lexer::LexerMode,
        token::{Error, OpCompare, OpProduct, OpSum, OpUnary},
    },
    formatter,
    formatter::format::{format_number, Formatted},
    functions::util::compare_values,
    language::{get_language, Language},
    locale::{get_locale, Locale},
    number_format::get_num_fmt,
    types::*,
};

#[derive(Clone)]
pub struct Environment {
    /// Returns the number of milliseconds January 1, 1970 00:00:00 UTC.
    pub get_milliseconds_since_epoch: fn() -> i64,
}

/// A model includes:
///     * A Workbook: An internal representation of and Excel workbook
///     * Parsed Formulas: All the formulas in the workbook are parsed here (runtime only)
///     * A list of cells with its status (evaluating, evaluated, not evaluated)
#[derive(Clone)]
pub struct Model {
    pub workbook: Workbook,
    pub parsed_formulas: Vec<Vec<Node>>,
    pub parser: Parser,
    pub cells: HashMap<String, char>,
    pub locale: Locale,
    pub language: Language,
    pub env: Environment,
    pub tz: Tz,
}

pub struct CellIndex {
    pub index: i32,
    pub row: i32,
    pub column: i32,
}

pub struct CellRowInput {
    pub style: i32,
    pub text: String,
    pub column: i32,
}

/// Used for the eval_workbook binary. A ExcelValue is the Excel representation of the cell content.
#[derive(Serialize, Deserialize, Debug, PartialEq)]
#[serde(untagged)]
pub enum ExcelValue {
    String(String),
    Number(f64),
    Boolean(bool),
}

#[derive(Serialize, Deserialize, Debug, PartialEq)]
#[serde(untagged)]
pub enum ExcelValueOrRange {
    Value(ExcelValue),
    Range(Vec<Vec<ExcelValue>>),
}

impl ExcelValue {
    pub fn to_json_str(&self) -> String {
        match &self {
            ExcelValue::String(s) => json!(s).to_string(),
            ExcelValue::Number(f) => json!(f).to_string(),
            ExcelValue::Boolean(b) => json!(b).to_string(),
        }
    }
}

/// sheet, row, column, value
pub type InputData = HashMap<i32, HashMap<i32, HashMap<i32, ExcelValue>>>;

#[derive(Serialize, Deserialize, Debug, PartialEq, Clone)]
pub struct Style {
    pub horizontal_alignment: String,
    pub read_only: bool,
    pub num_fmt: String,
    pub fill: Fill,
    pub font: Font,
    pub border: Border,
    pub quote_prefix: bool,
}

/// Excel compatibility values
/// COLUMN_WIDTH and ROW_HEIGHT are pixel values
/// A column width of Excel value `w` will result in `w * COLUMN_WIDTH_FACTOR` pixels
/// Note that these constants are inlined
pub(crate) const DEFAULT_COLUMN_WIDTH: f64 = 100.0;
pub(crate) const DEFAULT_ROW_HEIGHT: f64 = 21.0;
pub(crate) const COLUMN_WIDTH_FACTOR: f64 = 12.0;
pub(crate) const ROW_HEIGHT_FACTOR: f64 = 2.0;

/// Returns true if the string value could be interpreted as:
///  * a formula
///  * a number
///  * a boolean
///  * an error (i.e "#VALUE!")
fn value_needs_quoting(value: &str, language: &Language) -> bool {
    value.starts_with('=')
        || value.parse::<f64>().is_ok()
        || value.to_lowercase().parse::<bool>().is_ok()
        || get_error_by_name(&value.to_uppercase(), language).is_some()
}

// valid hex colors are #FFAABB
// #fff is not valid
fn is_valid_hex_color(color: &str) -> bool {
    if color.chars().count() != 7 {
        return false;
    }
    if !color.starts_with('#') {
        return false;
    }
    if let Ok(z) = i32::from_str_radix(&color[1..], 16) {
        if (0..=0xffffff).contains(&z) {
            return true;
        }
    }
    false
}

impl Model {
    fn get_string_index(&self, str: &str) -> Option<usize> {
        self.workbook.shared_strings.iter().position(|r| r == str)
    }

    fn get_range(&self, left: &Node, right: &Node, cell: CellReference) -> CalcResult {
        match (left, right) {
            (
                Node::ReferenceKind {
                    sheet_name: _,
                    sheet_index: sheet_left,
                    absolute_row: absolute_row_left,
                    absolute_column: absolute_column_left,
                    row: row_left,
                    column: column_left,
                },
                Node::ReferenceKind {
                    sheet_name: _,
                    sheet_index: _sheet_right,
                    absolute_row: absolute_row_right,
                    absolute_column: absolute_column_right,
                    row: row_right,
                    column: column_right,
                },
            ) => {
                let mut row1 = *row_left;
                let mut column1 = *column_left;
                if !absolute_row_left {
                    row1 += cell.row;
                }
                if !absolute_column_left {
                    column1 += cell.column;
                }
                let mut row2 = *row_right;
                let mut column2 = *column_right;
                if !absolute_row_right {
                    row2 += cell.row;
                }
                if !absolute_column_right {
                    column2 += cell.column;
                }
                // FIXME: HACK. The parser is currently parsing Sheet3!A1:A10 as Sheet3!A1:(present sheet)!A10
                CalcResult::Range {
                    left: CellReference {
                        sheet: *sheet_left as i32,
                        row: row1,
                        column: column1,
                    },
                    right: CellReference {
                        sheet: *sheet_left as i32,
                        row: row2,
                        column: column2,
                    },
                }
            }
            _ => CalcResult::Error {
                error: Error::NIMPL,
                origin: cell,
                message: "Function not implemented".to_string(),
            },
        }
    }

    pub(crate) fn evaluate_node_in_context(
        &mut self,
        node: &Node,
        cell: CellReference,
    ) -> CalcResult {
        use Node::*;
        match node {
            OpSumKind { kind, left, right } => {
                // In the future once the feature try trait stabilizes we could use the '?' operator for this :)
                // See: https://play.rust-lang.org/?version=nightly&mode=debug&edition=2018&gist=236044e8321a1450988e6ffe5a27dab5
                let l = match self.get_number(left, cell) {
                    Ok(f) => f,
                    Err(s) => {
                        return s;
                    }
                };
                let r = match self.get_number(right, cell) {
                    Ok(f) => f,
                    Err(s) => {
                        return s;
                    }
                };
                let result = match kind {
                    OpSum::Add => l + r,
                    OpSum::Minus => l - r,
                };
                CalcResult::Number(result)
            }
            NumberKind(value) => CalcResult::Number(*value),
            StringKind(value) => CalcResult::String(value.clone()),
            BooleanKind(value) => CalcResult::Boolean(*value),
            ReferenceKind {
                sheet_name: _,
                sheet_index,
                absolute_row,
                absolute_column,
                row,
                column,
            } => {
                let mut row1 = *row;
                let mut column1 = *column;
                if !absolute_row {
                    row1 += cell.row;
                }
                if !absolute_column {
                    column1 += cell.column;
                }
                self.evaluate_cell(CellReference {
                    sheet: *sheet_index,
                    row: row1,
                    column: column1,
                })
            }
            WrongReferenceKind { .. } => {
                CalcResult::new_error(Error::REF, cell, "Wrong reference".to_string())
            }
            OpRangeKind { left, right } => self.get_range(left, right, cell),
            WrongRangeKind { .. } => {
                CalcResult::new_error(Error::REF, cell, "Wrong range".to_string())
            }
            RangeKind {
                sheet_index,
                row1,
                column1,
                row2,
                column2,
                absolute_column1,
                absolute_row2,
                absolute_row1,
                absolute_column2,
                sheet_name: _,
            } => CalcResult::Range {
                left: CellReference {
                    sheet: *sheet_index,
                    row: if *absolute_row1 {
                        *row1
                    } else {
                        *row1 + cell.row
                    },
                    column: if *absolute_column1 {
                        *column1
                    } else {
                        *column1 + cell.column
                    },
                },
                right: CellReference {
                    sheet: *sheet_index,
                    row: if *absolute_row2 {
                        *row2
                    } else {
                        *row2 + cell.row
                    },
                    column: if *absolute_column2 {
                        *column2
                    } else {
                        *column2 + cell.column
                    },
                },
            },
            OpConcatenateKind { left, right } => {
                let l = match self.get_string(left, cell) {
                    Ok(f) => f,
                    Err(s) => {
                        return s;
                    }
                };
                let r = match self.get_string(right, cell) {
                    Ok(f) => f,
                    Err(s) => {
                        return s;
                    }
                };
                let result = format!("{}{}", l, r);
                CalcResult::String(result)
            }
            OpProductKind { kind, left, right } => {
                let l = match self.get_number(left, cell) {
                    Ok(f) => f,
                    Err(s) => {
                        return s;
                    }
                };
                let r = match self.get_number(right, cell) {
                    Ok(f) => f,
                    Err(s) => {
                        return s;
                    }
                };
                let result = match kind {
                    OpProduct::Times => l * r,
                    OpProduct::Divide => {
                        if r == 0.0 {
                            return CalcResult::new_error(
                                Error::DIV,
                                cell,
                                "Divide by Zero".to_string(),
                            );
                        }
                        l / r
                    }
                };
                CalcResult::Number(result)
            }
            OpPowerKind { left, right } => {
                let l = match self.get_number(left, cell) {
                    Ok(f) => f,
                    Err(s) => {
                        return s;
                    }
                };
                let r = match self.get_number(right, cell) {
                    Ok(f) => f,
                    Err(s) => {
                        return s;
                    }
                };
                // Deal with errors properly
                CalcResult::Number(l.powf(r))
            }
            FunctionKind { name, args } => match name.as_str() {
                // Logical
                "AND" => self.fn_and(args, cell),
                "FALSE" => CalcResult::Boolean(false),
                "IF" => self.fn_if(args, cell),
                "IFERROR" => self.fn_iferror(args, cell),
                "IFNA" => self.fn_ifna(args, cell),
                "IFS" => self.fn_ifs(args, cell),
                "NOT" => self.fn_not(args, cell),
                "OR" => self.fn_or(args, cell),
                "SWITCH" => self.fn_switch(args, cell),
                "TRUE" => CalcResult::Boolean(true),
                "XOR" => self.fn_xor(args, cell),
                // Math and trigonometry
                "SIN" => self.fn_sin(args, cell),
                "COS" => self.fn_cos(args, cell),
                "TAN" => self.fn_tan(args, cell),

                "ASIN" => self.fn_asin(args, cell),
                "ACOS" => self.fn_acos(args, cell),
                "ATAN" => self.fn_atan(args, cell),

                "SINH" => self.fn_sinh(args, cell),
                "COSH" => self.fn_cosh(args, cell),
                "TANH" => self.fn_tanh(args, cell),

                "ASINH" => self.fn_asinh(args, cell),
                "ACOSH" => self.fn_acosh(args, cell),
                "ATANH" => self.fn_atanh(args, cell),

                "PI" => self.fn_pi(args, cell),

                "MAX" => self.fn_max(args, cell),
                "MIN" => self.fn_min(args, cell),
                "ROUND" => self.fn_round(args, cell),
                "ROUNDDOWN" => self.fn_rounddown(args, cell),
                "ROUNDUP" => self.fn_roundup(args, cell),
                "SUM" => self.fn_sum(args, cell),
                "SUMIF" => self.fn_sumif(args, cell),
                "SUMIFS" => self.fn_sumifs(args, cell),
                // Lookup and Reference
                "COLUMN" => self.fn_column(args, cell),
                "COLUMNS" => self.fn_columns(args, cell),
                "INDEX" => self.fn_index(args, cell),
                "INDIRECT" => self.fn_indirect(args, cell),
                "HLOOKUP" => self.fn_hlookup(args, cell),
                "LOOKUP" => self.fn_lookup(args, cell),
                "MATCH" => self.fn_match(args, cell),
                "OFFSET" => self.fn_offset(args, cell),
                "ROW" => self.fn_row(args, cell),
                "ROWS" => self.fn_rows(args, cell),
                "VLOOKUP" => self.fn_vlookup(args, cell),
                "XLOOKUP" => self.fn_xlookup(args, cell),
                // Text
                "CONCAT" => self.fn_concat(args, cell),
                "FIND" => self.fn_find(args, cell),
                "LEFT" => self.fn_left(args, cell),
                "LEN" => self.fn_len(args, cell),
                "LOWER" => self.fn_lower(args, cell),
                "MID" => self.fn_mid(args, cell),
                "RIGHT" => self.fn_right(args, cell),
                "SEARCH" => self.fn_search(args, cell),
                "TEXT" => self.fn_text(args, cell),
                "TRIM" => self.fn_trim(args, cell),
                "UPPER" => self.fn_upper(args, cell),
                // Information
                "ISNUMBER" => self.fn_isnumber(args, cell),
                "ISNONTEXT" => self.fn_isnontext(args, cell),
                "ISTEXT" => self.fn_istext(args, cell),
                "ISLOGICAL" => self.fn_islogical(args, cell),
                "ISBLANK" => self.fn_isblank(args, cell),
                "ISERR" => self.fn_iserr(args, cell),
                "ISERROR" => self.fn_iserror(args, cell),
                "ISNA" => self.fn_isna(args, cell),
                "NA" => CalcResult::new_error(Error::NA, cell, "".to_string()),
                // Statistical
                "AVERAGE" => self.fn_average(args, cell),
                "AVERAGEA" => self.fn_averagea(args, cell),
                "AVERAGEIF" => self.fn_averageif(args, cell),
                "AVERAGEIFS" => self.fn_averageifs(args, cell),
                "COUNT" => self.fn_count(args, cell),
                "COUNTA" => self.fn_counta(args, cell),
                "COUNTBLANK" => self.fn_countblank(args, cell),
                "COUNTIF" => self.fn_countif(args, cell),
                "COUNTIFS" => self.fn_countifs(args, cell),
                "MAXIFS" => self.fn_maxifs(args, cell),
                "MINIFS" => self.fn_minifs(args, cell),
                // Date and Time
                "YEAR" => self.fn_year(args, cell),
                "DAY" => self.fn_day(args, cell),
                "MONTH" => self.fn_month(args, cell),
                "DATE" => self.fn_date(args, cell),
                "EDATE" => self.fn_edate(args, cell),
                "TODAY" => self.fn_today(args, cell),
                _ => {
                    CalcResult::new_error(Error::ERROR, cell, format!("Invalid function: {}", name))
                }
            },
            ArrayKind(_) => {
                // TODO: NOT IMPLEMENTED
                CalcResult::new_error(Error::NIMPL, cell, "Arrays not implemented".to_string())
            }
            VariableKind(_) => {
                // TODO: NOT IMPLEMENTED
                CalcResult::new_error(
                    Error::NIMPL,
                    cell,
                    "Defined names not implemented".to_string(),
                )
            }
            CompareKind { kind, left, right } => {
                let l = self.evaluate_node_in_context(left, cell);
                if l.is_error() {
                    return l;
                }
                let r = self.evaluate_node_in_context(right, cell);
                if r.is_error() {
                    return r;
                }
                let compare = compare_values(&l, &r);
                match kind {
                    OpCompare::Equal => {
                        if compare == 0 {
                            CalcResult::Boolean(true)
                        } else {
                            CalcResult::Boolean(false)
                        }
                    }
                    OpCompare::LessThan => {
                        if compare == -1 {
                            CalcResult::Boolean(true)
                        } else {
                            CalcResult::Boolean(false)
                        }
                    }
                    OpCompare::GreaterThan => {
                        if compare == 1 {
                            CalcResult::Boolean(true)
                        } else {
                            CalcResult::Boolean(false)
                        }
                    }
                    OpCompare::LessOrEqualThan => {
                        if compare < 1 {
                            CalcResult::Boolean(true)
                        } else {
                            CalcResult::Boolean(false)
                        }
                    }
                    OpCompare::GreaterOrEqualThan => {
                        if compare > -1 {
                            CalcResult::Boolean(true)
                        } else {
                            CalcResult::Boolean(false)
                        }
                    }
                    OpCompare::NonEqual => {
                        if compare != 0 {
                            CalcResult::Boolean(true)
                        } else {
                            CalcResult::Boolean(false)
                        }
                    }
                }
            }
            UnaryKind { kind, right } => {
                let r = match self.get_number(right, cell) {
                    Ok(f) => f,
                    Err(s) => {
                        return s;
                    }
                };
                match kind {
                    OpUnary::Minus => CalcResult::Number(-r),
                    OpUnary::Percentage => CalcResult::Number(r / 100.0),
                }
            }
            ErrorKind(kind) => CalcResult::new_error(kind.clone(), cell, "".to_string()),
            ParseErrorKind {
                formula,
                message,
                position: _,
            } => CalcResult::new_error(
                Error::ERROR,
                cell,
                format!("Error parsing {}: {}", formula, message),
            ),
            EmptyArgKind => CalcResult::EmptyArg,
        }
    }

    fn cell_reference_to_string(&self, cell_reference: &CellReference) -> Result<String, String> {
        let sheet = match self.workbook.worksheets.get(cell_reference.sheet as usize) {
            Some(s) => s,
            None => return Err("Invalid sheet".to_string()),
        };
        let column = match utils::number_to_column(cell_reference.column) {
            Some(c) => c,
            None => return Err("Invalid column".to_string()),
        };
        Ok(format!("{}!{}{}", sheet.name, column, cell_reference.row))
    }
    /// Sets `result` in the cell given by `sheet` sheet index, row and column
    /// Note that will panic if the cell does not exist
    /// It will do nothing if the cell does not have a formula
    fn set_cell_value(&mut self, cell_reference: CellReference, result: &CalcResult) {
        let CellReference { sheet, column, row } = cell_reference;
        let cell = &self.workbook.worksheets[sheet as usize].sheet_data[&row][&column];
        let s = cell.get_style();
        if let Some(f) = cell.get_formula() {
            match result {
                CalcResult::Number(value) => {
                    *self.workbook.worksheets[sheet as usize]
                        .sheet_data
                        .get_mut(&row)
                        .expect("expected a row")
                        .get_mut(&column)
                        .expect("expected a column") = Cell::CellFormulaNumber {
                        t: "n".to_string(),
                        f,
                        s,
                        v: *value,
                    };
                }
                CalcResult::String(value) => {
                    *self.workbook.worksheets[sheet as usize]
                        .sheet_data
                        .get_mut(&row)
                        .expect("expected a row")
                        .get_mut(&column)
                        .expect("expected a column") = Cell::CellFormulaString {
                        t: "str".to_string(),
                        f,
                        s,
                        v: value.clone(),
                    };
                }
                CalcResult::Boolean(value) => {
                    *self.workbook.worksheets[sheet as usize]
                        .sheet_data
                        .get_mut(&row)
                        .expect("expected a row")
                        .get_mut(&column)
                        .expect("expected a column") = Cell::CellFormulaBoolean {
                        t: "b".to_string(),
                        f,
                        s,
                        v: *value,
                    };
                }
                CalcResult::Error {
                    error,
                    origin,
                    message,
                } => {
                    let o = match self.cell_reference_to_string(origin) {
                        Ok(s) => s,
                        Err(_) => "".to_string(),
                    };
                    *self.workbook.worksheets[sheet as usize]
                        .sheet_data
                        .get_mut(&row)
                        .expect("expected a row")
                        .get_mut(&column)
                        .expect("expected a column") = Cell::CellFormulaError {
                        t: "e".to_string(),
                        f,
                        s,
                        o,
                        m: message.to_string(),
                        ei: error.clone(),
                    };
                }
                CalcResult::Range { left, right } => {
                    let range = Range {
                        left: *left,
                        right: *right,
                    };
                    if let Some(intersection_cell) =
                        self.implicit_intersection(&cell_reference, &range)
                    {
                        let v = self.evaluate_cell(intersection_cell);
                        self.set_cell_value(cell_reference, &v);
                    } else {
                        let o = match self.cell_reference_to_string(&cell_reference) {
                            Ok(s) => s,
                            Err(_) => "".to_string(),
                        };
                        *self.workbook.worksheets[sheet as usize]
                            .sheet_data
                            .get_mut(&row)
                            .expect("expected a row")
                            .get_mut(&column)
                            .expect("expected a column") = Cell::CellFormulaError {
                            t: "e".to_string(),
                            f,
                            s,
                            o,
                            m: "Invalid reference".to_string(),
                            ei: Error::VALUE,
                        };
                    }
                }
                CalcResult::EmptyCell | CalcResult::EmptyArg => {
                    *self.workbook.worksheets[sheet as usize]
                        .sheet_data
                        .get_mut(&row)
                        .expect("expected a row")
                        .get_mut(&column)
                        .expect("expected a column") = Cell::CellFormulaNumber {
                        t: "n".to_string(),
                        f,
                        s,
                        v: 0.0,
                    };
                }
            }
        }
    }

    pub fn set_sheet_color(&mut self, sheet: i32, color: &str) -> Result<(), String> {
        if let Some(worksheet) = self.workbook.worksheets.get_mut(sheet as usize) {
            if color.is_empty() {
                worksheet.color = Color::None;
                return Ok(());
            } else if is_valid_hex_color(color) {
                worksheet.color = Color::RGB(color.to_string());
                return Ok(());
            }
            Err(format!("Invalid color: {}", color))
        } else {
            Err("Invalid sheet".to_string())
        }
    }

    pub fn get_frozen_rows(&self, sheet: i32) -> Result<i32, String> {
        if let Some(worksheet) = self.workbook.worksheets.get(sheet as usize) {
            Ok(worksheet.frozen_rows)
        } else {
            Err("Invalid sheet".to_string())
        }
    }

    pub fn get_frozen_columns(&self, sheet: i32) -> Result<i32, String> {
        if let Some(worksheet) = self.workbook.worksheets.get(sheet as usize) {
            Ok(worksheet.frozen_columns)
        } else {
            Err("Invalid sheet".to_string())
        }
    }

    pub fn set_frozen_rows(&mut self, sheet: i32, frozen_rows: i32) -> Result<(), String> {
        if let Some(worksheet) = self.workbook.worksheets.get_mut(sheet as usize) {
            if frozen_rows < 0 {
                return Err("Frozen rows cannot be negative".to_string());
            } else if frozen_rows >= utils::LAST_ROW {
                return Err("Too many rows".to_string());
            }
            worksheet.frozen_rows = frozen_rows;
            Ok(())
        } else {
            Err("Invalid sheet".to_string())
        }
    }

    pub fn set_frozen_columns(&mut self, sheet: i32, frozen_columns: i32) -> Result<(), String> {
        if let Some(worksheet) = self.workbook.worksheets.get_mut(sheet as usize) {
            if frozen_columns < 0 {
                return Err("Frozen columns cannot be negative".to_string());
            } else if frozen_columns >= utils::LAST_COLUMN {
                return Err("Too many columns".to_string());
            }
            worksheet.frozen_columns = frozen_columns;
            Ok(())
        } else {
            Err("Invalid sheet".to_string())
        }
    }

    fn get_cell_value(&self, cell: &Cell, cell_reference: CellReference) -> CalcResult {
        use Cell::*;
        match cell {
            EmptyCell { .. } => CalcResult::EmptyCell,
            BooleanCell { v, .. } => CalcResult::Boolean(*v),
            NumberCell { v, .. } => CalcResult::Number(*v),
            ErrorCell { ei, .. } => {
                let message = ei.to_localized_error_string(&self.language);
                CalcResult::new_error(ei.clone(), cell_reference, message)
            }
            SharedString { si, .. } => {
                if let Some(s) = self.workbook.shared_strings.get(*si as usize) {
                    CalcResult::String(s.clone())
                } else {
                    let message = "Invalid shared string".to_string();
                    CalcResult::new_error(Error::ERROR, cell_reference, message)
                }
            }
            CellFormula { .. } => CalcResult::Error {
                error: Error::ERROR,
                origin: cell_reference,
                message: "Unevaluated formula".to_string(),
            },
            CellFormulaBoolean { v, .. } => CalcResult::Boolean(*v),
            CellFormulaNumber { v, .. } => CalcResult::Number(*v),
            CellFormulaString { v, .. } => CalcResult::String(v.clone()),
            CellFormulaError { ei, o, m, .. } => {
                if let Some(cell_reference) = self.parse_reference(o) {
                    CalcResult::new_error(ei.clone(), cell_reference, m.clone())
                } else {
                    CalcResult::Error {
                        error: ei.clone(),
                        origin: cell_reference,
                        message: ei.to_localized_error_string(&self.language),
                    }
                }
            }
        }
    }

    pub fn get_cell(&self, sheet: i32, row: i32, column: i32) -> Option<&Cell> {
        let worksheet = self.workbook.worksheets.get(sheet as usize)?;
        let sheet_data = &worksheet.sheet_data;
        let data_row = sheet_data.get(&row)?;
        data_row.get(&column)
    }

    pub fn get_cell_mut(&mut self, sheet: i32, row: i32, column: i32) -> Option<&mut Cell> {
        let worksheet = self.workbook.worksheets.get_mut(sheet as usize)?;
        let data_row = worksheet.sheet_data.get_mut(&row)?;
        data_row.get_mut(&column)
    }

    pub fn is_empty_cell(&self, sheet: i32, row: i32, column: i32) -> bool {
        let worksheet = &self.workbook.worksheets[sheet as usize];
        let sheet_data = &worksheet.sheet_data;
        if let Some(data_row) = sheet_data.get(&row) {
            if let Some(cell) = data_row.get(&column) {
                match cell {
                    Cell::EmptyCell { .. } => return true,
                    _ => return false,
                }
            }
            true
        } else {
            true
        }
    }

    pub(crate) fn evaluate_cell(&mut self, cell_reference: CellReference) -> CalcResult {
        let row_data = match self.workbook.worksheets[cell_reference.sheet as usize]
            .sheet_data
            .get(&cell_reference.row)
        {
            Some(r) => r,
            None => return CalcResult::EmptyCell,
        };
        let cell = match row_data.get(&cell_reference.column) {
            Some(c) => c,
            None => {
                return CalcResult::EmptyCell;
            }
        };

        match cell.get_formula() {
            Some(f) => {
                let key = format!(
                    "{}!{}!{}",
                    cell_reference.sheet, cell_reference.column, cell_reference.row,
                );
                match self.cells.get(&key) {
                    Some('c') => {
                        return CalcResult::new_error(
                            Error::CIRC,
                            cell_reference,
                            "Circular reference detected".to_string(),
                        );
                    }
                    Some('t') => {
                        return self.get_cell_value(cell, cell_reference);
                    }
                    _ => {
                        self.cells.insert(key.clone(), 'c');
                    }
                }
                // mark cell as being evaluated
                let node = &self.parsed_formulas[cell_reference.sheet as usize][f as usize].clone();
                let result = self.evaluate_node_in_context(node, cell_reference);
                self.set_cell_value(cell_reference, &result);
                self.cells.insert(key, 't');
                result
            }
            None => self.get_cell_value(cell, cell_reference),
        }
    }

    pub(crate) fn get_sheet_index_by_name(&self, name: &str) -> Option<usize> {
        let worksheets = &self.workbook.worksheets;
        for (index, worksheet) in worksheets.iter().enumerate() {
            if worksheet.get_name().to_uppercase() == name.to_uppercase() {
                return Some(index);
            }
        }
        None
    }

    // Public API
    /// Returns a model from a String representation of a workbook
    pub fn from_json(s: &str, env: Environment) -> Result<Model, String> {
        let workbook: Workbook = match serde_json::from_str(s) {
            Ok(w) => w,
            Err(_) => return Err("Error parsing workbook".to_string()),
        };
        let parsed_formulas = Vec::new();
        let worksheets = &workbook.worksheets;
        let parser = Parser::new(worksheets.iter().map(|s| s.get_name()).collect());
        let cells = HashMap::new();
        let locale = match get_locale(&workbook.settings.locale) {
            Ok(l) => l.clone(),
            Err(_) => return Err("Invalid locale".to_string()),
        };
        let tz: Tz = match &workbook.settings.tz.parse() {
            Ok(t) => *t,
            Err(_) => {
                return Err(format!("Invalid timezone: {}", &workbook.settings.tz));
            }
        };

        // FIXME: Add support for display languages
        let language = get_language("en").expect("").clone();

        let mut model = Model {
            workbook,
            parsed_formulas,
            parser,
            cells,
            language,
            locale,
            env,
            tz,
        };
        model.parse_formulas();
        Ok(model)
    }

    /// Parses a reference like "Sheet1!B4" into {0, 2, 4}
    pub(crate) fn parse_reference(&self, s: &str) -> Option<CellReference> {
        let bytes = s.as_bytes();
        let mut sheet_name = "".to_string();
        let mut column = "".to_string();
        let mut row = "".to_string();
        let mut state = "sheet"; // "sheet", "col", "row"
        for &byte in bytes {
            match state {
                "sheet" => {
                    if byte == b'!' {
                        state = "col"
                    } else {
                        sheet_name.push(byte as char);
                    }
                }
                "col" => {
                    if byte.is_ascii_alphabetic() {
                        column.push(byte as char);
                    } else {
                        state = "row";
                        row.push(byte as char);
                    }
                }
                _ => {
                    row.push(byte as char);
                }
            }
        }
        let sheet = match self.get_sheet_index_by_name(&sheet_name) {
            Some(s) => s,
            None => return None,
        } as i32;
        let row = match row.parse::<i32>() {
            Ok(r) => r,
            Err(_) => return None,
        };
        if !(1..=utils::LAST_ROW).contains(&row) {
            return None;
        }
        if column.is_empty() {
            return None;
        }
        let column = utils::column_to_number(&column);
        if !(1..=utils::LAST_COLUMN).contains(&column) {
            return None;
        }
        Some(CellReference { sheet, row, column })
    }

    /// moves the value in area from source to target.
    pub fn move_cell_value_to_area(
        &mut self,
        source: &CellReferenceIndex,
        value: &str,
        target: &CellReferenceIndex,
        area: &Area,
    ) -> Result<String, String> {
        let source_sheet_name = match self.workbook.worksheets.get(source.sheet as usize) {
            Some(ws) => ws.get_name(),
            None => {
                return Err("Invalid worksheet index".to_owned());
            }
        };
        if source.sheet != area.sheet {
            return Err("Source and area are in different sheets".to_string());
        }
        if source.row < area.row || source.row >= area.row + area.height {
            return Err("Source is outside the area".to_string());
        }
        if source.column < area.column || source.column >= area.column + area.width {
            return Err("Source is outside the area".to_string());
        }
        let target_sheet_name = match self.workbook.worksheets.get(target.sheet as usize) {
            Some(ws) => ws.get_name(),
            None => {
                return Err("Invalid worksheet index".to_owned());
            }
        };
        if let Some(formula) = value.strip_prefix('=') {
            let cell_reference = CellReferenceRC {
                sheet: source_sheet_name.to_owned(),
                row: source.row,
                column: source.column,
            };
            let formula_str = move_formula(
                &self.parser.parse(formula, &Some(cell_reference)),
                &MoveContext {
                    source_sheet_name: &source_sheet_name,
                    row: source.row,
                    column: source.column,
                    area,
                    target_sheet_name: &target_sheet_name,
                    row_delta: target.row - source.row,
                    column_delta: target.column - source.column,
                },
            );
            Ok(format!("={}", formula_str))
        } else {
            Ok(value.to_string())
        }
    }

    /// 'Extends' the value from cell [sheet, row, column] to [target_row, target_column]
    pub fn extend_to(
        &self,
        sheet: i32,
        row: i32,
        column: i32,
        target_row: i32,
        target_column: i32,
    ) -> String {
        match self.get_cell(sheet, row, column) {
            Some(cell) => match cell.get_formula() {
                None => cell.get_text(&self.workbook.shared_strings, &self.language),
                Some(i) => {
                    let formula = &self.parsed_formulas[sheet as usize][i as usize];
                    let cell_ref = CellReferenceRC {
                        sheet: self.workbook.worksheets[sheet as usize].get_name(),
                        row: target_row,
                        column: target_column,
                    };
                    format!("={}", to_string(formula, &cell_ref))
                }
            },
            None => "".to_string(),
        }
    }

    /// 'Extends' value from cell [sheet, row, column] to [target_row, target_column]
    pub fn extend_formula_to(
        &mut self,
        source: &CellReferenceIndex,
        value: &str,
        target: &CellReferenceIndex,
    ) -> Result<String, String> {
        let source_sheet_name = match self.workbook.worksheets.get(source.sheet as usize) {
            Some(ws) => ws.get_name(),
            None => {
                return Err("Invalid worksheet index".to_owned());
            }
        };
        let target_sheet_name = match self.workbook.worksheets.get(target.sheet as usize) {
            Some(ws) => ws.get_name(),
            None => {
                return Err("Invalid worksheet index".to_owned());
            }
        };
        if let Some(formula_str) = value.strip_prefix('=') {
            let cell_reference = CellReferenceRC {
                sheet: source_sheet_name,
                row: source.row,
                column: source.column,
            };
            let formula = &self.parser.parse(formula_str, &Some(cell_reference));
            let cell_reference = CellReferenceRC {
                sheet: target_sheet_name,
                row: target.row,
                column: target.column,
            };
            return Ok(format!("={}", to_string(formula, &cell_reference)));
        };
        Ok(value.to_string())
    }

    /// Returns a formula if the cell has one or the value of the cell
    pub fn get_formula_or_value(&self, sheet: i32, row: i32, column: i32) -> String {
        match self.get_cell(sheet, row, column) {
            Some(cell) => match cell.get_formula() {
                None => cell.get_text(&self.workbook.shared_strings, &self.language),
                Some(i) => {
                    let formula = &self.parsed_formulas[sheet as usize][i as usize];
                    let cell_ref = CellReferenceRC {
                        sheet: self.workbook.worksheets[sheet as usize].get_name(),
                        row,
                        column,
                    };
                    format!("={}", to_string(formula, &cell_ref))
                }
            },
            None => "".to_string(),
        }
    }

    /// Checks if cell has formula
    pub fn has_formula(&self, sheet: i32, row: i32, column: i32) -> bool {
        match self.get_cell(sheet, row, column) {
            Some(cell) => cell.get_formula().is_some(),
            None => false,
        }
    }

    /// Returns a text representation of the value of the cell
    pub fn get_text_at(&self, sheet: i32, row: i32, column: i32) -> String {
        match self.get_cell(sheet, row, column) {
            Some(cell) => cell.get_text(&self.workbook.shared_strings, &self.language),
            None => "".to_string(),
        }
    }

    /// Returns the information needed to display a cell in the UI
    pub fn get_ui_cell(&self, sheet: i32, row: i32, column: i32) -> UICell {
        return match self.get_cell(sheet, row, column) {
            Some(cell) => cell.get_ui_cell(&self.workbook.shared_strings, &self.language),
            None => UICell {
                kind: "empty".to_string(),
                value: UIValue::Text("".to_string()),
                details: "".to_string(),
            },
        };
    }

    /// Used by the dashboard project
    pub fn get_range_data(&self, range: &str) -> Result<Vec<Vec<ExcelValue>>, String> {
        let mut lx = Lexer::new(
            range,
            LexerMode::A1,
            &self.locale,
            get_language("en").expect(""),
        );
        match lx.next_token() {
            token::TokenType::RANGE { sheet, left, right } => {
                if lx.next_token() != token::TokenType::EOF {
                    return Err("Invalid range".to_string());
                }
                if let Some(sheet_name) = sheet {
                    if let Some(sheet_index) = self.get_sheet_index_by_name(&sheet_name) {
                        let mut data = Vec::new();
                        for row in left.row..=right.row {
                            let mut row_data = Vec::new();
                            for column in left.column..=right.column {
                                let excel_value =
                                    self.get_cell_value_by_index(sheet_index as i32, row, column);
                                row_data.push(excel_value);
                            }
                            data.push(row_data);
                        }

                        return Ok(data);
                    }
                    return Err("Invalid sheet name".to_string());
                };
                Err("Expected sheet name".to_string())
            }
            _ => Err("Invalid range".to_string()),
        }
    }

    pub fn get_range_formatted_data(&self, range: &str) -> Result<Vec<Vec<String>>, String> {
        let mut lx = Lexer::new(
            range,
            LexerMode::A1,
            &self.locale,
            get_language("en").expect(""),
        );
        match lx.next_token() {
            token::TokenType::RANGE { sheet, left, right } => {
                if lx.next_token() != token::TokenType::EOF {
                    return Err("Invalid range".to_string());
                }
                if let Some(sheet_name) = sheet {
                    if let Some(sheet_index) = self.get_sheet_index_by_name(&sheet_name) {
                        let mut data = Vec::new();
                        for row in left.row..=right.row {
                            let mut row_data = Vec::new();
                            for column in left.column..=right.column {
                                let excel_value =
                                    self.get_cell_value_by_index(sheet_index as i32, row, column);

                                let value = match excel_value {
                                    ExcelValue::String(s) => s,
                                    ExcelValue::Number(number) => {
                                        let style = self.get_style_for_cell(
                                            sheet_index as i32,
                                            row,
                                            column,
                                        );
                                        format_number(number, &style.num_fmt, &self.locale).text
                                    }
                                    ExcelValue::Boolean(b) => format!("{}", b).to_uppercase(),
                                };
                                row_data.push(value);
                            }
                            data.push(row_data);
                        }

                        return Ok(data);
                    }
                    return Err("Invalid sheet name".to_string());
                };
                Err("Expected sheet name".to_string())
            }
            _ => Err("Invalid range".to_string()),
        }
    }

    pub fn format_number(&self, value: f64, format_code: String) -> Formatted {
        formatter::format::format_number(value, &format_code, &self.locale)
    }

    pub(crate) fn get_cell_formula_index(&self, sheet: i32, row: i32, column: i32) -> Option<i32> {
        match &self.workbook.worksheets[sheet as usize]
            .sheet_data
            .get(&row)
        {
            Some(full_row) => match full_row.get(&column) {
                Some(cell) => cell.get_formula(),
                None => None,
            },
            None => None,
        }
    }

    pub(crate) fn shift_cell_formula(
        &mut self,
        sheet: i32,
        row: i32,
        column: i32,
        displace_data: &DisplaceData,
    ) {
        if let Some(f) = self.get_cell_formula_index(sheet, row, column) {
            let node = &self.parsed_formulas[sheet as usize][f as usize].clone();
            let cell_reference = CellReferenceRC {
                sheet: self.workbook.worksheets[sheet as usize].get_name(),
                row,
                column,
            };
            // FIXME: This is not a very performant way if the formula has changed :S.
            let formula = to_string(node, &cell_reference);
            let formula_displaced = to_string_displaced(node, &cell_reference, displace_data);
            if formula != formula_displaced {
                self.set_input_with_formula(sheet, row, column, &formula_displaced);
            }
        }
    }

    // This assumes the cell exists. Do not make public
    pub(crate) fn set_input_with_formula(
        &mut self,
        sheet: i32,
        row: i32,
        column: i32,
        formula: &str,
    ) {
        let style = self.get_cell_style_index(sheet, row, column);
        let worksheets = &mut self.workbook.worksheets;
        let worksheet = &mut worksheets[sheet as usize];
        let cell_reference = CellReferenceRC {
            sheet: worksheet.get_name(),
            row,
            column,
        };
        let shared_formulas = &mut worksheet.shared_formulas;
        let node = self.parser.parse(formula, &Some(cell_reference));
        let formula_rc = to_rc_format(&node);
        let mut formula_index: i32 = -1;
        if let Some(index) = shared_formulas.iter().position(|x| x == &formula_rc) {
            formula_index = index as i32;
        }
        if formula_index == -1 {
            shared_formulas.push(formula_rc);
            self.parsed_formulas[sheet as usize].push(node);
            formula_index = (shared_formulas.len() as i32) - 1;
        }
        worksheet.set_cell_with_formula(row, column, formula_index, style);
    }

    /// Updates the value of a cell with some text
    /// It does not change the style unless needs to add "quoting"
    pub fn update_cell_with_text(&mut self, sheet: i32, row: i32, column: i32, value: &str) {
        let style_index = self.get_cell_style_index(sheet, row, column);
        let new_style_index;
        if value_needs_quoting(value, &self.language) {
            new_style_index = self.get_style_with_quote_prefix(style_index);
        } else if self.style_is_quote_prefix(style_index) {
            new_style_index = self.get_style_without_quote_prefix(style_index);
        } else {
            new_style_index = style_index;
        }
        self.set_cell_with_string(sheet, row, column, value, new_style_index);
    }

    /// Updates the value of a cell with a boolean value
    /// It does not change the style
    pub fn update_cell_with_bool(&mut self, sheet: i32, row: i32, column: i32, value: bool) {
        let style_index = self.get_cell_style_index(sheet, row, column);
        let new_style_index = if self.style_is_quote_prefix(style_index) {
            self.get_style_without_quote_prefix(style_index)
        } else {
            style_index
        };
        let worksheet = &mut self.workbook.worksheets[sheet as usize];
        worksheet.set_cell_with_boolean(row, column, value, new_style_index);
    }

    /// Updates the value of a cell with a number
    /// It does not change the style
    pub fn update_cell_with_number(&mut self, sheet: i32, row: i32, column: i32, value: f64) {
        let style_index = self.get_cell_style_index(sheet, row, column);
        let new_style_index = if self.style_is_quote_prefix(style_index) {
            self.get_style_without_quote_prefix(style_index)
        } else {
            style_index
        };
        let worksheet = &mut self.workbook.worksheets[sheet as usize];
        worksheet.set_cell_with_number(row, column, value, new_style_index);
    }

    /// Sets a cell parametrized by (`sheet`, `row`, `column`) with `value` and `style`
    /// The value is always a string, so we need to try to cast it into numbers/bools/errors
    pub fn set_input(
        &mut self,
        sheet: i32,
        row: i32,
        column: i32,
        value: String,
        style_index: i32,
    ) {
        // FIXME: This is a bad API.
        // If value starts with "'" then we force the style to be quote_prefix
        // If it is not we force the styleNOT to be quote_prefix
        // We probably should have two separate methods (set_input and set_style)
        if let Some(new_value) = value.strip_prefix('\'') {
            // First check if it needs quoting
            let new_style = if value_needs_quoting(new_value, &self.language) {
                self.get_style_with_quote_prefix(style_index)
            } else {
                style_index
            };
            self.set_cell_with_string(sheet, row, column, new_value, new_style);
        } else {
            let mut new_style_index = style_index;
            if self.style_is_quote_prefix(style_index) {
                new_style_index = self.get_style_without_quote_prefix(style_index);
            }
            let worksheets = &mut self.workbook.worksheets;
            let worksheet = &mut worksheets[sheet as usize];
            if let Some(formula) = value.strip_prefix('=') {
                let cell_reference = CellReferenceRC {
                    sheet: worksheet.get_name(),
                    row,
                    column,
                };
                let shared_formulas = &mut worksheet.shared_formulas;
                let mut parsed_formula = self.parser.parse(formula, &Some(cell_reference.clone()));
                // If the formula fails to parse try adding a parenthesis
                // SUM(A1:A3  => SUM(A1:A3)
                if let Node::ParseErrorKind { .. } = parsed_formula {
                    let new_parsed_formula = self
                        .parser
                        .parse(&format!("{})", formula), &Some(cell_reference));
                    match new_parsed_formula {
                        Node::ParseErrorKind { .. } => {}
                        _ => parsed_formula = new_parsed_formula,
                    }
                }

                let s = to_rc_format(&parsed_formula);
                let mut formula_index: i32 = -1;
                if let Some(index) = shared_formulas.iter().position(|x| x == &s) {
                    formula_index = index as i32;
                }
                if formula_index == -1 {
                    shared_formulas.push(s);
                    self.parsed_formulas[sheet as usize].push(parsed_formula);
                    formula_index = (shared_formulas.len() as i32) - 1;
                }
                worksheet.set_cell_with_formula(row, column, formula_index, new_style_index);
            } else {
                let worksheets = &mut self.workbook.worksheets;
                let worksheet = &mut worksheets[sheet as usize];
                // We try to parse as number
                if let Ok(v) = value.parse::<f64>() {
                    worksheet.set_cell_with_number(row, column, v, new_style_index);
                    return;
                }
                // We try to parse as boolean
                if let Ok(v) = value.to_lowercase().parse::<bool>() {
                    worksheet.set_cell_with_boolean(row, column, v, new_style_index);
                    return;
                }
                // Check is it is error value
                let upper = value.to_uppercase();
                match get_error_by_name(&upper, &self.language) {
                    Some(error) => {
                        worksheet.set_cell_with_error(row, column, error, new_style_index);
                    }
                    None => {
                        self.set_cell_with_string(sheet, row, column, &value, new_style_index);
                    }
                }
            }
        }
    }

    fn set_cell_with_string(&mut self, sheet: i32, row: i32, column: i32, value: &str, style: i32) {
        let worksheets = &mut self.workbook.worksheets;
        let worksheet = &mut worksheets[sheet as usize];
        let shared_strings = &mut self.workbook.shared_strings;
        let index = shared_strings.iter().position(|r| r == value);
        match index {
            Some(string_index) => {
                worksheet.set_cell_with_string(row, column, string_index as i32, style);
            }
            None => {
                shared_strings.push(value.to_string());
                let string_index = shared_strings.len() as i32 - 1;
                worksheet.set_cell_with_string(row, column, string_index, style);
            }
        }
    }

    /// Gets the Excel Value (Bool, Number, String) of a cell
    pub fn get_cell_value_by_ref(&self, cell_ref: &str) -> Result<ExcelValue, String> {
        let cell_reference = match self.parse_reference(cell_ref) {
            Some(c) => c,
            None => return Err(format!("Error parsing reference: '{cell_ref}'")),
        };
        // get the worksheet
        let sheet_index = cell_reference.sheet;

        let column = cell_reference.column;
        let row = cell_reference.row;
        Ok(self.get_cell_value_by_index(sheet_index, row, column))
    }

    fn get_cell_value_by_index(&self, sheet_index: i32, row: i32, column: i32) -> ExcelValue {
        let cell = self.get_cell_at(sheet_index as i32, row, column);

        match cell {
            Cell::EmptyCell { .. } => ExcelValue::String("".to_string()),
            Cell::BooleanCell { t: _, v, s: _ } => ExcelValue::Boolean(v),
            Cell::NumberCell { t: _, v, s: _ } => ExcelValue::Number(v),
            Cell::ErrorCell { ei, .. } => {
                let v = ei.to_localized_error_string(&self.language);
                ExcelValue::String(v)
            }
            Cell::SharedString { si, .. } => {
                let v = self.workbook.shared_strings[si as usize].clone();
                ExcelValue::String(v)
            }
            Cell::CellFormula { .. } => ExcelValue::String("#ERROR!".to_string()),
            Cell::CellFormulaBoolean { v, .. } => ExcelValue::Boolean(v),
            Cell::CellFormulaNumber { v, .. } => ExcelValue::Number(v),
            Cell::CellFormulaString { v, .. } => ExcelValue::String(v),
            Cell::CellFormulaError { ei, .. } => {
                let v = ei.to_localized_error_string(&self.language);
                ExcelValue::String(v)
            }
        }
    }

    pub fn set_cells_with_values_json(&mut self, input_json: &str) -> Result<(), String> {
        let values: HashMap<String, ExcelValue> = match serde_json::from_str(input_json) {
            Ok(v) => v,
            Err(_) => return Err("Cannot parse input".to_string()),
        };
        self.set_cells_with_values(values)
    }

    fn set_cells_with_parsed_references(&mut self, input_data: &InputData) -> Result<(), String> {
        for (&sheet_index, sheet_data) in input_data {
            for (&row, row_data) in sheet_data {
                for (&column, value) in row_data {
                    let style_index = self.get_cell_style_index(sheet_index, row, column);
                    match &value {
                        ExcelValue::String(v) => {
                            let upper = v.to_uppercase();
                            match get_error_by_name(&upper, &self.language) {
                                Some(index) => {
                                    self.workbook.worksheets[sheet_index as usize]
                                        .set_cell_with_error(row, column, index, style_index);
                                }
                                None => {
                                    let si = match self.get_string_index(v) {
                                        Some(i) => i as i32,
                                        None => {
                                            self.workbook.shared_strings.push(v.clone());
                                            self.workbook.shared_strings.len() as i32 - 1
                                        }
                                    };
                                    self.workbook.worksheets[sheet_index as usize]
                                        .set_cell_with_string(row, column, si, style_index);
                                }
                            }
                        }
                        ExcelValue::Number(v) => {
                            self.workbook.worksheets[sheet_index as usize].set_cell_with_number(
                                row,
                                column,
                                *v,
                                style_index,
                            );
                        }
                        ExcelValue::Boolean(v) => {
                            self.workbook.worksheets[sheet_index as usize].set_cell_with_boolean(
                                row,
                                column,
                                *v,
                                style_index,
                            );
                        }
                    }
                }
            }
        }
        Ok(())
    }

    /// Sets cells with the Excel Values (Used by eval_workbook and calc)
    pub fn set_cells_with_values(
        &mut self,
        data: HashMap<String, ExcelValue>,
    ) -> Result<(), String> {
        let mut input_data: InputData = HashMap::new();
        for (key, value) in data {
            // TODO: If the reference is wrong this function should return an error
            let cell_reference = match self.parse_reference(&key) {
                Some(c) => c,
                None => return Err("Invalid reference".to_string()),
            };

            // get the worksheet index
            let sheet = cell_reference.sheet;

            let column = cell_reference.column;
            let row = cell_reference.row;
            let sheet_data = input_data.entry(sheet).or_default();
            let row_data = sheet_data.entry(row).or_default();
            row_data.insert(column, value);
        }
        self.set_cells_with_parsed_references(&input_data)
    }

    /// Returns a list of all cells
    pub fn get_all_cells(&self) -> Vec<CellIndex> {
        let mut cells = Vec::new();
        for (index, sheet) in self.workbook.worksheets.iter().enumerate() {
            let mut sorted_rows: Vec<_> = sheet.sheet_data.keys().collect();
            sorted_rows.sort_unstable();
            for row in sorted_rows {
                let row_data = &sheet.sheet_data[row];
                let mut sorted_columns: Vec<_> = row_data.keys().collect();
                sorted_columns.sort_unstable();
                for column in sorted_columns {
                    cells.push(CellIndex {
                        index: index as i32,
                        row: *row,
                        column: *column,
                    });
                }
            }
        }
        cells
    }

    /// Returns dimension of the sheet: (min_row, min_column, max_row, max_column)
    pub fn get_sheet_dimension(&self, sheet: i32) -> (i32, i32, i32, i32) {
        // FIXME, this should be read from the worksheet:
        // self.workbook.worksheets[sheet as usize].dimension
        let mut min_column = -1;
        let mut max_column = -1;
        let worksheet = &self.workbook.worksheets[sheet as usize];
        let mut sorted_rows: Vec<_> = worksheet.sheet_data.keys().collect();
        sorted_rows.sort_unstable();
        if sorted_rows.is_empty() {
            return (1, 1, 1, 1);
        }
        let min_row = *sorted_rows[0];
        let max_row = *sorted_rows[sorted_rows.len() - 1];
        for row in sorted_rows {
            let row_data = &worksheet.sheet_data[row];
            let mut sorted_columns: Vec<_> = row_data.keys().collect();
            sorted_columns.sort_unstable();
            if sorted_columns.is_empty() {
                continue;
            }
            if min_column == -1 {
                min_column = *sorted_columns[0];
                max_column = *sorted_columns[sorted_columns.len() - 1];
            } else {
                min_column = min_column.min(*sorted_columns[0]);
                max_column = max_column.max(*sorted_columns[sorted_columns.len() - 1]);
            }
        }
        if min_column == -1 {
            return (1, 1, 1, 1);
        }
        (min_row, min_column, max_row, max_column)
    }

    /// Returns the Cell. Used in tests
    pub fn get_cell_at(&self, sheet: i32, row: i32, column: i32) -> Cell {
        match self.get_cell(sheet, row, column) {
            Some(cell) => cell.clone(),
            None => Cell::EmptyCell {
                t: "empty".to_string(),
                s: 0,
            },
        }
    }

    /// Evaluates the model with a top-down recursive algorithm
    pub fn evaluate(&mut self) {
        // clear all computation artifacts
        self.cells.clear();

        let cells = self.get_all_cells();

        for cell in cells {
            self.evaluate_cell(CellReference {
                sheet: cell.index,
                row: cell.row,
                column: cell.column,
            });
        }
    }

    /// Uses the present state of the model to evaluate the input without changing it's state
    pub fn evaluate_with_input(
        &mut self,
        input_json: &str,
        output_refs: &[&str],
    ) -> Result<HashMap<String, ExcelValueOrRange>, String> {
        let input_data: InputData = match serde_json::from_str(input_json) {
            Ok(v) => v,
            Err(_) => return Err("Cannot parse input".to_string()),
        };

        let mut model = self.clone();

        model
            .set_cells_with_parsed_references(&input_data)
            .expect("Could not set input cells");
        model.evaluate();

        let mut output: HashMap<String, ExcelValueOrRange> = HashMap::new();
        for output_ref in output_refs {
            match model.get_cell_value_by_ref(output_ref) {
                Ok(value) => {
                    output.insert(output_ref.to_string(), ExcelValueOrRange::Value(value));
                }
                Err(message) => {
                    if let Ok(result) = model.get_range_data(output_ref) {
                        output.insert(output_ref.to_string(), ExcelValueOrRange::Range(result));
                    } else {
                        return Err(message);
                    }
                }
            }
        }

        Ok(output)
    }

    /// Return the width of a column in pixels
    pub fn get_column_width(&self, sheet: i32, column: i32) -> f64 {
        let cols = &self.workbook.worksheets[sheet as usize].cols;
        for col in cols {
            let min = col.min;
            let max = col.max;
            if column >= min && column <= max {
                if col.custom_width {
                    return col.width * COLUMN_WIDTH_FACTOR;
                } else {
                    break;
                }
            }
        }
        DEFAULT_COLUMN_WIDTH
    }

    /// Returns the height of a row in pixels
    pub fn get_row_height(&self, sheet: i32, row: i32) -> f64 {
        let rows = &self.workbook.worksheets[sheet as usize].rows;
        for r in rows {
            if r.r == row {
                return r.height * ROW_HEIGHT_FACTOR;
            }
        }
        DEFAULT_ROW_HEIGHT
    }

    // FIXME: This should return an object
    /// Returns a JSON string with (name, state and visibility) of each sheet
    pub fn get_tabs(&self) -> String {
        let mut tabs = Vec::new();
        let worksheets = &self.workbook.worksheets;
        for (index, worksheet) in worksheets.iter().enumerate() {
            tabs.push(Tab {
                name: worksheet.get_name(),
                state: worksheet.state.clone(),
                color: worksheet.color.clone(),
                index: index as i32,
                sheet_id: worksheet.sheet_id,
            });
        }
        match serde_json::to_string(&tabs) {
            Ok(s) => s,
            Err(_) => json!([]).to_string(),
        }
    }

    // FIXME: This should return an object
    /// Returns a JSON string with the set of merge cells
    pub fn get_merge_cells(&self, sheet: i32) -> String {
        let merge_cells = &self.workbook.worksheets[sheet as usize].merge_cells;
        match serde_json::to_string(&merge_cells) {
            Ok(s) => s,
            Err(_) => json!([]).to_string(),
        }
    }

    /// Deletes a cell by setting it empty.
    /// TODO: A better name would be set_cell_empty or remove_cell_contents
    pub fn delete_cell(&mut self, sheet: i32, row: i32, column: i32) {
        let worksheets = &mut self.workbook.worksheets;
        let worksheet = &mut worksheets[sheet as usize];
        worksheet.set_cell_empty(row, column);
    }

    pub fn remove_cell(&mut self, sheet: i32, row: i32, column: i32) {
        let worksheets = &mut self.workbook.worksheets;
        let worksheet = &mut worksheets[sheet as usize];
        let sheet_data = &mut worksheet.sheet_data;
        if sheet_data.contains_key(&row) {
            sheet_data
                .get_mut(&row)
                .expect("expected row")
                .remove(&column);
        }
    }

    /// Changes the height of a row.
    ///   * If the row does not a have a style we add it.
    ///   * If it has we modify the height and make sure it is applied.
    pub fn set_row_height(&mut self, sheet: i32, row: i32, height: f64) {
        let rows = &mut self.workbook.worksheets[sheet as usize].rows;
        for r in rows.iter_mut() {
            if r.r == row {
                r.height = height / ROW_HEIGHT_FACTOR;
                r.custom_height = true;
                return;
            }
        }
        rows.push(Row {
            height: height / ROW_HEIGHT_FACTOR,
            r: row,
            custom_format: false,
            custom_height: true,
            s: 0,
        })
    }
    /// Changes the width of a column.
    ///   * If the column does not a have a width we simply add it
    ///   * If it has, it might be part of a range and we ned to split the range.
    pub fn set_column_width(&mut self, sheet: i32, column: i32, width: f64) {
        let cols = &mut self.workbook.worksheets[sheet as usize].cols;
        let mut col = Col {
            min: column,
            max: column,
            width: width / COLUMN_WIDTH_FACTOR,
            custom_width: true,
            style: None,
        };
        let mut index = 0;
        let mut split = false;
        for c in cols.iter_mut() {
            let min = c.min;
            let max = c.max;
            if min <= column && column <= max {
                if min == column && max == column {
                    c.width = width / COLUMN_WIDTH_FACTOR;
                    return;
                } else {
                    // We need to split the result
                    split = true;
                    break;
                }
            }
            if column < min {
                // We passed, we should insert at index
                break;
            }
            index += 1;
        }
        if split {
            let min = cols[index].min;
            let max = cols[index].max;
            let pre = Col {
                min,
                max: column - 1,
                width: cols[index].width,
                custom_width: cols[index].custom_width,
                style: cols[index].style,
            };
            let post = Col {
                min: column + 1,
                max,
                width: cols[index].width,
                custom_width: cols[index].custom_width,
                style: cols[index].style,
            };
            col.style = cols[index].style;
            let index = index as usize;
            cols.remove(index);
            if column != max {
                cols.insert(index, post);
            }
            cols.insert(index, col);
            if column != min {
                cols.insert(index, pre);
            }
        } else {
            cols.insert(index, col);
        }
    }

    /// A sheet is read only iff all columns are read only
    fn is_sheet_read_only(&self, sheet: i32) -> bool {
        let cols = &self.workbook.worksheets[sheet as usize].cols;
        let mut last_col = 0;
        for col in cols {
            if col.min == last_col + 1 {
                let is_read_only = match col.style {
                    Some(style) => self.get_style(style).read_only,
                    None => false,
                };
                if !is_read_only {
                    return false;
                }
                last_col = col.max;
                if col.max == utils::LAST_COLUMN {
                    return true;
                }
            } else {
                return false;
            }
        }
        false
    }

    /// Returns true if the row is read only or if the whole sheet is read only
    pub fn is_row_read_only(&self, sheet: i32, row: i32) -> bool {
        let rows = &self.workbook.worksheets[sheet as usize].rows;
        for r in rows {
            if r.r == row {
                if self.get_style(r.s).read_only {
                    return true;
                }
                break;
            }
        }
        self.is_sheet_read_only(sheet)
    }

    pub fn get_cell_style_index(&self, sheet: i32, row: i32, column: i32) -> i32 {
        // First check the cell, then row, the column
        match self.get_cell(sheet, row, column) {
            Some(cell) => cell.get_style() as i32,
            None => {
                let rows = &self.workbook.worksheets[sheet as usize].rows;
                for r in rows {
                    if r.r == row {
                        if r.custom_format {
                            return r.s;
                        } else {
                            break;
                        }
                    }
                }
                let cols = &self.workbook.worksheets[sheet as usize].cols;
                for c in cols.iter() {
                    let min = c.min;
                    let max = c.max;
                    if column >= min && column <= max {
                        return c.style.unwrap_or(0);
                    }
                }
                0
            }
        }
    }

    pub fn get_style_for_cell(&self, sheet: i32, row: i32, column: i32) -> Style {
        self.get_style(self.get_cell_style_index(sheet, row, column))
    }

    fn style_is_quote_prefix(&self, index: i32) -> bool {
        let styles = &self.workbook.styles;
        let cell_xf = &styles.cell_xfs[index as usize];
        cell_xf.quote_prefix
    }

    fn get_style(&self, index: i32) -> Style {
        let styles = &self.workbook.styles;
        let cell_xf = &styles.cell_xfs[index as usize];

        let border_id = cell_xf.border_id as usize;
        let fill_id = cell_xf.fill_id as usize;
        let font_id = cell_xf.font_id as usize;
        let num_fmt_id = cell_xf.num_fmt_id;
        let read_only = cell_xf.read_only;
        let quote_prefix = cell_xf.quote_prefix;
        let horizontal_alignment = cell_xf.horizontal_alignment.clone();

        Style {
            horizontal_alignment,
            read_only,
            num_fmt: get_num_fmt(num_fmt_id, &styles.num_fmts),
            fill: styles.fills[fill_id].clone(),
            font: styles.fonts[font_id].clone(),
            border: styles.borders[border_id].clone(),
            quote_prefix,
        }
    }

    fn get_style_with_quote_prefix(&mut self, index: i32) -> i32 {
        let mut style = self.get_style(index);
        style.quote_prefix = true;
        // Check if style exist. If so sets style cell number to that otherwise create a new style.
        if let Some(index) = self.get_style_index(&style) {
            index
        } else {
            self.create_new_style(&style)
        }
    }

    fn get_style_without_quote_prefix(&mut self, index: i32) -> i32 {
        let mut style = self.get_style(index);
        style.quote_prefix = false;
        // Check if style exist. If so sets style cell number to that otherwise create a new style.
        if let Some(index) = self.get_style_index(&style) {
            index
        } else {
            self.create_new_style(&style)
        }
    }

    /// Returns a list with all the names of the worksheets
    pub fn get_worksheet_names(&self) -> Vec<String> {
        self.workbook.get_worksheet_names()
    }

    /// Returns a JSON string of the workbook
    pub fn to_json_str(&self) -> String {
        match serde_json::to_string(&self.workbook) {
            Ok(s) => s,
            Err(_) => {
                // TODO, is this branch possible at all?
                json!({"error": "Error stringifying workbook"}).to_string()
            }
        }
    }

    /// Removes all data on a sheet, including the cell styles
    pub fn remove_sheet_data(&mut self, sheet_index: i32) -> Result<(), String> {
        let worksheet = match self.workbook.worksheets.get_mut(sheet_index as usize) {
            Some(s) => s,
            None => return Err("Wrong worksheet index".to_string()),
        };

        // Remove all data
        worksheet.sheet_data = HashMap::new();
        Ok(())
    }
}
