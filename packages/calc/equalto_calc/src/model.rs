use chrono_tz::Tz;
use serde::{Deserialize, Serialize};
use serde_json::json;

use std::collections::HashMap;
use std::vec::Vec;

use crate::{
    calc_result::{CalcResult, CellReference, Range},
    cell::CellValue,
    constants,
    expressions::token::{Error, OpCompare, OpProduct, OpSum, OpUnary},
    expressions::{
        parser::move_formula::{move_formula, MoveContext},
        token::get_error_by_name,
        types::*,
        utils::{self, is_valid_row},
    },
    expressions::{
        parser::{
            stringify::{to_rc_format, to_string},
            Node, Parser,
        },
        utils::is_valid_column_number,
    },
    formatter::format::format_number,
    functions::util::compare_values,
    implicit_intersection::implicit_intersection,
    language::{get_language, Language},
    locale::{get_locale, Locale},
    types::*,
    utils as common,
};
#[cfg(feature = "wasm-build")]
use js_sys::Date;
#[cfg(not(feature = "wasm-build"))]
use std::time::{SystemTime, UNIX_EPOCH};

#[derive(Clone)]
pub struct Environment {
    /// Returns the number of milliseconds January 1, 1970 00:00:00 UTC.
    pub get_milliseconds_since_epoch: fn() -> i64,
}

fn get_milliseconds_since_epoch() -> i64 {
    #[cfg(not(feature = "wasm-build"))]
    return SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("problem with system time")
        .as_millis() as i64;
    #[cfg(feature = "wasm-build")]
    return Date::now() as i64;
}

impl Default for Environment {
    fn default() -> Self {
        Self {
            get_milliseconds_since_epoch,
        }
    }
}

#[derive(Clone)]
pub enum CellState {
    Evaluated,
    Evaluating,
}

#[derive(Debug, Clone)]
pub enum ParsedDefinedName {
    CellReference(CellReference),
    RangeReference(Range),
    InvalidDefinedNameFormula,
    // TODO: Support constants in defined names
    // TODO: Support formulas in defined names
    // TODO: Support tables in defined names
}

/// A model includes:
///     * A Workbook: An internal representation of and Excel workbook
///     * Parsed Formulas: All the formulas in the workbook are parsed here (runtime only)
///     * A list of cells with its status (evaluating, evaluated, not evaluated)
#[derive(Clone)]
pub struct Model {
    pub workbook: Workbook,
    pub parsed_formulas: Vec<Vec<Node>>,
    pub parsed_defined_names: HashMap<(Option<u32>, String), ParsedDefinedName>,
    pub parser: Parser,
    pub cells: HashMap<(u32, i32, i32), CellState>,
    pub locale: Locale,
    pub language: Language,
    pub env: Environment,
    pub tz: Tz,
}

pub struct CellIndex {
    pub index: u32,
    pub row: i32,
    pub column: i32,
}

#[derive(Serialize, Deserialize, Debug, PartialEq, Eq, Clone)]
pub struct Style {
    pub alignment: Option<Alignment>,
    pub num_fmt: String,
    pub fill: Fill,
    pub font: Font,
    pub border: Border,
    pub quote_prefix: bool,
}

impl Model {
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
                        sheet: *sheet_left,
                        row: row1,
                        column: column1,
                    },
                    right: CellReference {
                        sheet: *sheet_left,
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
            FunctionKind { kind, args } => self.evaluate_function(kind, args, cell),
            InvalidFunctionKind { name, args: _ } => {
                CalcResult::new_error(Error::ERROR, cell, format!("Invalid function: {}", name))
            }
            ArrayKind(_) => {
                // TODO: NOT IMPLEMENTED
                CalcResult::new_error(Error::NIMPL, cell, "Arrays not implemented".to_string())
            }
            VariableKind(defined_name) => {
                let parsed_defined_name = self
                    .parsed_defined_names
                    .get(&(Some(cell.sheet), defined_name.to_lowercase())) // try getting local defined name
                    .or_else(|| {
                        self.parsed_defined_names
                            .get(&(None, defined_name.to_lowercase()))
                    }); // fallback to global

                if let Some(parsed_defined_name) = parsed_defined_name {
                    match parsed_defined_name {
                        ParsedDefinedName::CellReference(reference) => {
                            self.evaluate_cell(*reference)
                        }
                        ParsedDefinedName::RangeReference(range) => CalcResult::Range {
                            left: range.left,
                            right: range.right,
                        },
                        ParsedDefinedName::InvalidDefinedNameFormula => CalcResult::new_error(
                            Error::NIMPL,
                            cell,
                            format!("Defined name \"{}\" is not a reference.", defined_name),
                        ),
                    }
                } else {
                    CalcResult::new_error(
                        Error::NAME,
                        cell,
                        format!("Defined name \"{}\" not found.", defined_name),
                    )
                }
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
        let sheet = self.workbook.worksheet(cell_reference.sheet)?;
        let column = utils::number_to_column(cell_reference.column)
            .ok_or_else(|| "Invalid column".to_string())?;
        if !is_valid_row(cell_reference.row) {
            return Err("Invalid row".to_string());
        }
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
                        .expect("expected a column") = Cell::CellFormulaNumber { f, s, v: *value };
                }
                CalcResult::String(value) => {
                    *self.workbook.worksheets[sheet as usize]
                        .sheet_data
                        .get_mut(&row)
                        .expect("expected a row")
                        .get_mut(&column)
                        .expect("expected a column") = Cell::CellFormulaString {
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
                        .expect("expected a column") = Cell::CellFormulaBoolean { f, s, v: *value };
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
                    if let Some(intersection_cell) = implicit_intersection(&cell_reference, &range)
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
                        .expect("expected a column") = Cell::CellFormulaNumber { f, s, v: 0.0 };
                }
            }
        }
    }

    pub fn set_sheet_color(&mut self, sheet: u32, color: &str) -> Result<(), String> {
        let mut worksheet = self.workbook.worksheet_mut(sheet)?;
        if color.is_empty() {
            worksheet.color = None;
            return Ok(());
        } else if common::is_valid_hex_color(color) {
            worksheet.color = Some(color.to_string());
            return Ok(());
        }
        Err(format!("Invalid color: {}", color))
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

    /// Returns true if cell is completely empty.
    /// Cell with formula that evaluates to empty string is not considered empty.
    pub fn is_empty_cell(&self, sheet: u32, row: i32, column: i32) -> Result<bool, String> {
        let worksheet = self.workbook.worksheet(sheet)?;
        let sheet_data = &worksheet.sheet_data;

        let is_empty = if let Some(data_row) = sheet_data.get(&row) {
            if let Some(cell) = data_row.get(&column) {
                matches!(cell, Cell::EmptyCell { .. })
            } else {
                true
            }
        } else {
            true
        };

        Ok(is_empty)
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
                let key = (
                    cell_reference.sheet,
                    cell_reference.row,
                    cell_reference.column,
                );
                match self.cells.get(&key) {
                    Some(CellState::Evaluating) => {
                        return CalcResult::new_error(
                            Error::CIRC,
                            cell_reference,
                            "Circular reference detected".to_string(),
                        );
                    }
                    Some(CellState::Evaluated) => {
                        return self.get_cell_value(cell, cell_reference);
                    }
                    _ => {
                        // mark cell as being evaluated
                        self.cells.insert(key, CellState::Evaluating);
                    }
                }
                let node = &self.parsed_formulas[cell_reference.sheet as usize][f as usize].clone();
                let result = self.evaluate_node_in_context(node, cell_reference);
                self.set_cell_value(cell_reference, &result);
                // mark cell as evaluated
                self.cells.insert(key, CellState::Evaluated);
                result
            }
            None => self.get_cell_value(cell, cell_reference),
        }
    }

    pub(crate) fn get_sheet_index_by_name(&self, name: &str) -> Option<u32> {
        let worksheets = &self.workbook.worksheets;
        for (index, worksheet) in worksheets.iter().enumerate() {
            if worksheet.get_name().to_uppercase() == name.to_uppercase() {
                return Some(index as u32);
            }
        }
        None
    }

    // Public API
    /// Returns a model from a String representation of a workbook
    pub fn from_json(s: &str, env: Environment) -> Result<Model, String> {
        let workbook: Workbook =
            serde_json::from_str(s).map_err(|_| "Error parsing workbook".to_string())?;
        Model::from_workbook(workbook, env)
    }

    pub fn from_workbook(workbook: Workbook, env: Environment) -> Result<Model, String> {
        let parsed_formulas = Vec::new();
        let worksheets = &workbook.worksheets;
        let parser = Parser::new(worksheets.iter().map(|s| s.get_name()).collect());
        let cells = HashMap::new();
        let locale = get_locale(&workbook.settings.locale)
            .map_err(|_| "Invalid locale".to_string())?
            .clone();
        let tz: Tz = workbook
            .settings
            .tz
            .parse()
            .map_err(|_| format!("Invalid timezone: {}", workbook.settings.tz))?;

        // FIXME: Add support for display languages
        let language = get_language("en").expect("").clone();

        let mut model = Model {
            workbook,
            parsed_formulas,
            parsed_defined_names: HashMap::new(),
            parser,
            cells,
            language,
            locale,
            env,
            tz,
        };

        model.parse_formulas();
        model.parse_defined_names();

        Ok(model)
    }

    /// Parses a reference like "Sheet1!B4" into {0, 2, 4}
    pub fn parse_reference(&self, s: &str) -> Option<CellReference> {
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
        };
        let row = match row.parse::<i32>() {
            Ok(r) => r,
            Err(_) => return None,
        };
        if !(1..=constants::LAST_ROW).contains(&row) {
            return None;
        }

        let column = match utils::column_to_number(&column) {
            Ok(column) => {
                if is_valid_column_number(column) {
                    column
                } else {
                    return None;
                }
            }
            Err(_) => return None,
        };

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
        let source_sheet_name = self
            .workbook
            .worksheet(source.sheet)
            .map_err(|e| format!("Could not find source worksheet: {}", e))?
            .get_name();
        if source.sheet != area.sheet {
            return Err("Source and area are in different sheets".to_string());
        }
        if source.row < area.row || source.row >= area.row + area.height {
            return Err("Source is outside the area".to_string());
        }
        if source.column < area.column || source.column >= area.column + area.width {
            return Err("Source is outside the area".to_string());
        }
        let target_sheet_name = self
            .workbook
            .worksheet(target.sheet)
            .map_err(|e| format!("Could not find target worksheet: {}", e))?
            .get_name();
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
        sheet: u32,
        row: i32,
        column: i32,
        target_row: i32,
        target_column: i32,
    ) -> Result<String, String> {
        let cell = self.workbook.worksheet(sheet)?.cell(row, column);
        let result = match cell {
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
        };
        Ok(result)
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

    pub fn cell_formula(
        &self,
        sheet: u32,
        row: i32,
        column: i32,
    ) -> Result<Option<String>, String> {
        let worksheet = self.workbook.worksheet(sheet)?;
        Ok(worksheet.cell(row, column).and_then(|cell| {
            cell.get_formula().map(|formula_index| {
                let formula = &self.parsed_formulas[sheet as usize][formula_index as usize];
                let cell_ref = CellReferenceRC {
                    sheet: worksheet.get_name(),
                    row,
                    column,
                };
                format!("={}", to_string(formula, &cell_ref))
            })
        }))
    }

    /// Updates the value of a cell with some text
    /// It does not change the style unless needs to add "quoting"
    pub fn update_cell_with_text(&mut self, sheet: u32, row: i32, column: i32, value: &str) {
        let style_index = self.get_cell_style_index(sheet, row, column);
        let new_style_index;
        if common::value_needs_quoting(value, &self.language) {
            new_style_index = self
                .workbook
                .styles
                .get_style_with_quote_prefix(style_index);
        } else if self.workbook.styles.style_is_quote_prefix(style_index) {
            new_style_index = self
                .workbook
                .styles
                .get_style_without_quote_prefix(style_index);
        } else {
            new_style_index = style_index;
        }
        self.set_cell_with_string(sheet, row, column, value, new_style_index);
    }

    /// Updates the value of a cell with a boolean value
    /// It does not change the style
    pub fn update_cell_with_bool(&mut self, sheet: u32, row: i32, column: i32, value: bool) {
        let style_index = self.get_cell_style_index(sheet, row, column);
        let new_style_index = if self.workbook.styles.style_is_quote_prefix(style_index) {
            self.workbook
                .styles
                .get_style_without_quote_prefix(style_index)
        } else {
            style_index
        };
        let worksheet = &mut self.workbook.worksheets[sheet as usize];
        worksheet.set_cell_with_boolean(row, column, value, new_style_index);
    }

    /// Updates the value of a cell with a number
    /// It does not change the style
    pub fn update_cell_with_number(&mut self, sheet: u32, row: i32, column: i32, value: f64) {
        let style_index = self.get_cell_style_index(sheet, row, column);
        let new_style_index = if self.workbook.styles.style_is_quote_prefix(style_index) {
            self.workbook
                .styles
                .get_style_without_quote_prefix(style_index)
        } else {
            style_index
        };
        let worksheet = &mut self.workbook.worksheets[sheet as usize];
        worksheet.set_cell_with_number(row, column, value, new_style_index);
    }

    /// Updates the formula of given cell
    /// It does not change the style unless needs to add "quoting"
    /// Expects the formula to start with "="
    pub fn update_cell_with_formula(
        &mut self,
        sheet: u32,
        row: i32,
        column: i32,
        formula: String,
    ) -> Result<(), String> {
        let mut style_index = self.get_cell_style_index(sheet, row, column);
        if self.workbook.styles.style_is_quote_prefix(style_index) {
            style_index = self
                .workbook
                .styles
                .get_style_without_quote_prefix(style_index);
        }
        let formula = formula
            .strip_prefix('=')
            .ok_or_else(|| format!("\"{formula}\" is not a valid formula"))?;
        self.set_cell_with_formula(sheet, row, column, formula, style_index)
    }

    /// Sets a cell parametrized by (`sheet`, `row`, `column`) with `value` and `style`
    /// The value is always a string, so we need to try to cast it into numbers/bools/errors
    pub fn set_input(
        &mut self,
        sheet: u32,
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
            let new_style = if common::value_needs_quoting(new_value, &self.language) {
                self.workbook
                    .styles
                    .get_style_with_quote_prefix(style_index)
            } else {
                style_index
            };
            self.set_cell_with_string(sheet, row, column, new_value, new_style);
        } else {
            let mut new_style_index = style_index;
            if self.workbook.styles.style_is_quote_prefix(style_index) {
                new_style_index = self
                    .workbook
                    .styles
                    .get_style_without_quote_prefix(style_index);
            }
            if let Some(formula) = value.strip_prefix('=') {
                self.set_cell_with_formula(sheet, row, column, formula, new_style_index)
                    .expect("could not set the cell formula")
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

    fn set_cell_with_formula(
        &mut self,
        sheet: u32,
        row: i32,
        column: i32,
        formula: &str,
        style: i32,
    ) -> Result<(), String> {
        let worksheet = self.workbook.worksheet_mut(sheet)?;
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
        worksheet.set_cell_with_formula(row, column, formula_index, style);
        Ok(())
    }

    fn set_cell_with_string(&mut self, sheet: u32, row: i32, column: i32, value: &str, style: i32) {
        // Interestingly, `self.workbook.worksheet()` cannot be used because it would create two
        // mutable borrows of worksheet. However, I suspect that lexical lifetimes silently help
        // here, so there is no issue with inlined call.
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

    // FIXME: Can't put it in Workbook, because language is outside of workbook, sic!
    /// Gets the Excel Value (Bool, Number, String) of a cell
    pub fn get_cell_value_by_ref(&self, cell_ref: &str) -> Result<CellValue, String> {
        let cell_reference = match self.parse_reference(cell_ref) {
            Some(c) => c,
            None => return Err(format!("Error parsing reference: '{cell_ref}'")),
        };
        let sheet_index = cell_reference.sheet;
        let column = cell_reference.column;
        let row = cell_reference.row;

        self.get_cell_value_by_index(sheet_index, row, column)
    }

    // FIXME: Can't put it in Workbook, because language is outside of workbook, sic!
    pub fn get_cell_value_by_index(
        &self,
        sheet_index: u32,
        row: i32,
        column: i32,
    ) -> Result<CellValue, String> {
        let cell = self
            .workbook
            .worksheet(sheet_index)?
            .cell(row, column)
            .cloned()
            .unwrap_or_default();
        let cell_value = cell.value(&self.workbook.shared_strings, &self.language);
        Ok(cell_value)
    }

    // FIXME: Can't put it in Workbook, because locale and language are outside of workbook, sic!
    pub fn formatted_cell_value(
        &self,
        sheet_index: u32,
        row: i32,
        column: i32,
    ) -> Result<String, String> {
        let format = self.get_style_for_cell(sheet_index, row, column).num_fmt;
        let cell = self
            .workbook
            .worksheet(sheet_index)?
            .cell(row, column)
            .cloned()
            .unwrap_or_default();
        let formatted_value =
            cell.formatted_value(&self.workbook.shared_strings, &self.language, |value| {
                format_number(value, &format, &self.locale).text
            });
        Ok(formatted_value)
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
                        index: index as u32,
                        row: *row,
                        column: *column,
                    });
                }
            }
        }
        cells
    }

    /// Returns dimension of the sheet: (min_row, min_column, max_row, max_column)
    pub fn get_sheet_dimension(&self, sheet: u32) -> (i32, i32, i32, i32) {
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

    /// Evaluates the model with a top-down recursive algorithm
    /// Returns an error instead of using #N/IMPL!, #CIRC! or #ERROR! values.
    pub fn evaluate_with_error_check(&mut self) -> Result<(), String> {
        // clear all computation artifacts
        self.cells.clear();

        let cells = self.get_all_cells();

        let mut result = Ok(());

        for cell in cells {
            let calc_result = self.evaluate_cell(CellReference {
                sheet: cell.index,
                row: cell.row,
                column: cell.column,
            });
            if result.is_err() {
                continue;
            }
            if let CalcResult::Error {
                error: Error::CIRC | Error::NIMPL | Error::ERROR,
                origin,
                message,
            } = calc_result
            {
                // if the error origin formula is None then it's a cell with an explicitly set
                // error value (i.e. "#ERROR!"), in that case we don't want to break the evaluation
                if let Some(formula) = self
                    .cell_formula(origin.sheet, origin.row, origin.column)
                    .expect("expected a valid error origin cell")
                {
                    let cell_text_reference = self
                        .cell_reference_to_string(&origin)
                        .expect("expected a valid error origin cell");
                    result = Err(format!("{cell_text_reference} ('{formula}'): {message}"));
                }
            }
        }

        result
    }

    /// Sets cell to empty. Can be used to delete value without affecting style.
    pub fn set_cell_empty(&mut self, sheet: u32, row: i32, column: i32) -> Result<(), String> {
        let worksheet = self.workbook.worksheet_mut(sheet)?;
        worksheet.set_cell_empty(row, column);
        Ok(())
    }

    /// Deletes a cell by removing it from worksheet data.
    pub fn delete_cell(&mut self, sheet: u32, row: i32, column: i32) -> Result<(), String> {
        let worksheet = self.workbook.worksheet_mut(sheet)?;

        let sheet_data = &mut worksheet.sheet_data;
        if let Some(row_data) = sheet_data.get_mut(&row) {
            row_data.remove(&column);
        }

        Ok(())
    }

    // FIXME: expect
    pub fn get_cell_style_index(&self, sheet: u32, row: i32, column: i32) -> i32 {
        // First check the cell, then row, the column
        let cell = self
            .workbook
            .worksheet(sheet)
            .expect("Invalid sheet")
            .cell(row, column);
        match cell {
            Some(cell) => cell.get_style(),
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

    pub fn get_style_for_cell(&self, sheet: u32, row: i32, column: i32) -> Style {
        self.workbook
            .styles
            .get_style(self.get_cell_style_index(sheet, row, column))
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
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::test::util::new_empty_model;

    #[test]
    fn test_cell_reference_to_string() {
        let model = new_empty_model();
        let reference = CellReference {
            sheet: 0,
            row: 32,
            column: 16,
        };
        assert_eq!(
            model.cell_reference_to_string(&reference),
            Ok("Sheet1!P32".to_string())
        )
    }

    #[test]
    fn test_cell_reference_to_string_invalid_worksheet() {
        let model = new_empty_model();
        let reference = CellReference {
            sheet: 10,
            row: 1,
            column: 1,
        };
        assert_eq!(
            model.cell_reference_to_string(&reference),
            Err("Invalid sheet index".to_string())
        )
    }

    #[test]
    fn test_cell_reference_to_string_invalid_column() {
        let model = new_empty_model();
        let reference = CellReference {
            sheet: 0,
            row: 1,
            column: 20_000,
        };
        assert_eq!(
            model.cell_reference_to_string(&reference),
            Err("Invalid column".to_string())
        )
    }

    #[test]
    fn test_cell_reference_to_string_invalid_row() {
        let model = new_empty_model();
        let reference = CellReference {
            sheet: 0,
            row: 2_000_000,
            column: 1,
        };
        assert_eq!(
            model.cell_reference_to_string(&reference),
            Err("Invalid row".to_string())
        )
    }

    #[test]
    fn test_get_cell() {
        let mut model = new_empty_model();
        model._set("A1", "35");
        model._set("A2", "");
        let worksheet = model.workbook.worksheet(0).expect("Invalid sheet");

        assert_eq!(
            worksheet.cell(1, 1),
            Some(&Cell::NumberCell { v: 35.0, s: 0 })
        );

        assert_eq!(
            worksheet.cell(2, 1),
            Some(&Cell::SharedString { si: 0, s: 0 })
        );
        assert_eq!(worksheet.cell(3, 1), None)
    }

    #[test]
    fn test_get_cell_invalid_sheet() {
        let model = new_empty_model();
        assert_eq!(
            model.workbook.worksheet(5),
            Err("Invalid sheet index".to_string()),
        )
    }
}
