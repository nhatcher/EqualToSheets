use crate::{
    calc_result::{CalcResult, CellReference, Range},
    expressions::{parser::Node, token::Error},
    model::Model,
};

impl Model {
    pub(crate) fn get_number(
        &mut self,
        node: &Node,
        sheet: i32,
        column: i32,
        row: i32,
    ) -> Result<f64, CalcResult> {
        let c = self.evaluate_node_in_context(node, sheet, column, row);
        self.cast_to_number(c, sheet, column, row)
    }

    fn cast_to_number(
        &mut self,
        result: CalcResult,
        sheet: i32,
        column: i32,
        row: i32,
    ) -> Result<f64, CalcResult> {
        match result {
            CalcResult::Number(f) => Ok(f),
            CalcResult::String(s) => match s.parse::<f64>() {
                Ok(f) => Ok(f),
                _ => Err(CalcResult::new_error(
                    Error::VALUE,
                    sheet,
                    row,
                    column,
                    "Expecting number".to_string(),
                )),
            },
            CalcResult::Boolean(f) => {
                if f {
                    Ok(1.0)
                } else {
                    Ok(0.0)
                }
            }
            CalcResult::EmptyCell | CalcResult::EmptyArg => Ok(0.0),
            CalcResult::Error {
                error,
                origin,
                message,
            } => Err(CalcResult::Error {
                error,
                origin,
                message,
            }),
            CalcResult::Range { left, right } => {
                match self.implicit_intersection(
                    &CellReference { sheet, column, row },
                    &Range { left, right },
                ) {
                    Some(cell_reference) => {
                        let c = self.evaluate_cell(
                            cell_reference.sheet,
                            cell_reference.row,
                            cell_reference.column,
                        );
                        self.cast_to_number(
                            c,
                            cell_reference.sheet,
                            cell_reference.row,
                            cell_reference.column,
                        )
                    }
                    None => Err(CalcResult::Error {
                        error: Error::VALUE,
                        origin: CellReference { sheet, column, row },
                        message: "Invalid reference".to_string(),
                    }),
                }
            }
        }
    }

    pub(crate) fn get_string(
        &mut self,
        node: &Node,
        sheet: i32,
        column: i32,
        row: i32,
    ) -> Result<String, CalcResult> {
        let c = self.evaluate_node_in_context(node, sheet, column, row);
        self.cast_to_string(c, sheet, column, row)
    }

    fn cast_to_string(
        &mut self,
        result: CalcResult,
        sheet: i32,
        column: i32,
        row: i32,
    ) -> Result<String, CalcResult> {
        match result {
            CalcResult::Number(f) => Ok(format!("{}", f)),
            CalcResult::String(s) => Ok(s),
            CalcResult::Boolean(f) => {
                if f {
                    Ok("TRUE".to_string())
                } else {
                    Ok("FALSE".to_string())
                }
            }
            CalcResult::EmptyCell | CalcResult::EmptyArg => Ok("".to_string()),
            CalcResult::Error {
                error,
                origin,
                message,
            } => Err(CalcResult::Error {
                error,
                origin,
                message,
            }),
            CalcResult::Range { left, right } => {
                match self.implicit_intersection(
                    &CellReference { sheet, column, row },
                    &Range { left, right },
                ) {
                    Some(cell_reference) => {
                        let c = self.evaluate_cell(
                            cell_reference.sheet,
                            cell_reference.row,
                            cell_reference.column,
                        );
                        self.cast_to_string(
                            c,
                            cell_reference.sheet,
                            cell_reference.row,
                            cell_reference.column,
                        )
                    }
                    None => Err(CalcResult::Error {
                        error: Error::VALUE,
                        origin: CellReference { sheet, column, row },
                        message: "Invalid reference".to_string(),
                    }),
                }
            }
        }
    }

    pub(crate) fn get_boolean(
        &mut self,
        node: &Node,
        sheet: i32,
        column: i32,
        row: i32,
    ) -> Result<bool, CalcResult> {
        let c = self.evaluate_node_in_context(node, sheet, column, row);
        self.cast_to_bool(c, sheet, column, row)
    }

    fn cast_to_bool(
        &mut self,
        result: CalcResult,
        sheet: i32,
        column: i32,
        row: i32,
    ) -> Result<bool, CalcResult> {
        match result {
            CalcResult::Number(f) => {
                if f == 0.0 {
                    return Ok(false);
                }
                Ok(true)
            }
            CalcResult::String(s) => {
                if s.to_lowercase() == *"true" {
                    return Ok(true);
                } else if s.to_lowercase() == *"false" {
                    return Ok(false);
                }
                Err(CalcResult::Error {
                    error: Error::VALUE,
                    origin: CellReference { sheet, column, row },
                    message: "Expected boolean".to_string(),
                })
            }
            CalcResult::Boolean(b) => Ok(b),
            CalcResult::EmptyCell | CalcResult::EmptyArg => Ok(false),
            CalcResult::Error {
                error,
                origin,
                message,
            } => Err(CalcResult::Error {
                error,
                origin,
                message,
            }),
            CalcResult::Range { left, right } => {
                match self.implicit_intersection(
                    &CellReference { sheet, column, row },
                    &Range { left, right },
                ) {
                    Some(cell_reference) => {
                        let c = self.evaluate_cell(
                            cell_reference.sheet,
                            cell_reference.row,
                            cell_reference.column,
                        );
                        self.cast_to_bool(
                            c,
                            cell_reference.sheet,
                            cell_reference.row,
                            cell_reference.column,
                        )
                    }
                    None => Err(CalcResult::Error {
                        error: Error::VALUE,
                        origin: CellReference { sheet, column, row },
                        message: "Invalid reference".to_string(),
                    }),
                }
            }
        }
    }

    // tries to return a reference. That is either a reference or a formula that evaluates to a range/reference
    pub(crate) fn get_reference(
        &mut self,
        node: &Node,
        sheet_ref: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> Result<Range, CalcResult> {
        match node {
            Node::ReferenceKind {
                column,
                absolute_column,
                row,
                absolute_row,
                sheet_index,
                sheet_name: _,
            } => {
                let left = CellReference {
                    sheet: *sheet_index,
                    row: if *absolute_row { *row } else { *row + row_ref },
                    column: if *absolute_column {
                        *column
                    } else {
                        *column + column_ref
                    },
                };

                Ok(Range {
                    left: left.clone(),
                    right: left,
                })
            }
            _ => {
                let value = self.evaluate_node_in_context(node, sheet_ref, column_ref, row_ref);
                if value.is_error() {
                    return Err(value);
                }
                if let CalcResult::Range { left, right } = value {
                    Ok(Range { left, right })
                } else {
                    Err(CalcResult::Error {
                        error: Error::VALUE,
                        origin: CellReference {
                            sheet: sheet_ref,
                            row: row_ref,
                            column: column_ref,
                        },
                        message: "Expected reference".to_string(),
                    })
                }
            }
        }
    }
}
