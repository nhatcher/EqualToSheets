use crate::{
    calc_result::{CalcResult, CellReference},
    expressions::parser::Node,
    expressions::token::Error,
    model::Model,
};

impl Model {
    pub(crate) fn fn_isnumber(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        if args.len() == 1 {
            match self.evaluate_node_in_context(&args[0], sheet, column_ref, row_ref) {
                CalcResult::Number(_) => return CalcResult::Boolean(true),
                _ => {
                    return CalcResult::Boolean(false);
                }
            };
        }
        CalcResult::Error {
            error: Error::ERROR,
            origin: CellReference {
                sheet,
                row: row_ref,
                column: column_ref,
            },
            message: "Wrong number of arguments".to_string(),
        }
    }
    pub(crate) fn fn_istext(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        if args.len() == 1 {
            match self.evaluate_node_in_context(&args[0], sheet, column_ref, row_ref) {
                CalcResult::String(_) => return CalcResult::Boolean(true),
                _ => {
                    return CalcResult::Boolean(false);
                }
            };
        }
        CalcResult::Error {
            error: Error::ERROR,
            origin: CellReference {
                sheet,
                row: row_ref,
                column: column_ref,
            },
            message: "Wrong number of arguments".to_string(),
        }
    }
    pub(crate) fn fn_isnontext(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        if args.len() == 1 {
            match self.evaluate_node_in_context(&args[0], sheet, column_ref, row_ref) {
                CalcResult::String(_) => return CalcResult::Boolean(false),
                _ => {
                    return CalcResult::Boolean(true);
                }
            };
        }
        CalcResult::Error {
            error: Error::ERROR,
            origin: CellReference {
                sheet,
                row: row_ref,
                column: column_ref,
            },
            message: "Wrong number of arguments".to_string(),
        }
    }
    pub(crate) fn fn_islogical(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        if args.len() == 1 {
            match self.evaluate_node_in_context(&args[0], sheet, column_ref, row_ref) {
                CalcResult::Boolean(_) => return CalcResult::Boolean(true),
                _ => {
                    return CalcResult::Boolean(false);
                }
            };
        }
        CalcResult::Error {
            error: Error::ERROR,
            origin: CellReference {
                sheet,
                row: row_ref,
                column: column_ref,
            },
            message: "Wrong number of arguments".to_string(),
        }
    }
    pub(crate) fn fn_isblank(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        if args.len() == 1 {
            match self.evaluate_node_in_context(&args[0], sheet, column_ref, row_ref) {
                CalcResult::EmptyCell => return CalcResult::Boolean(true),
                _ => {
                    return CalcResult::Boolean(false);
                }
            };
        }
        CalcResult::Error {
            error: Error::ERROR,
            origin: CellReference {
                sheet,
                row: row_ref,
                column: column_ref,
            },
            message: "Wrong number of arguments".to_string(),
        }
    }
    pub(crate) fn fn_iserror(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        if args.len() == 1 {
            match self.evaluate_node_in_context(&args[0], sheet, column_ref, row_ref) {
                CalcResult::Error { .. } => return CalcResult::Boolean(true),
                _ => {
                    return CalcResult::Boolean(false);
                }
            };
        }
        CalcResult::Error {
            error: Error::ERROR,
            origin: CellReference {
                sheet,
                row: row_ref,
                column: column_ref,
            },
            message: "Wrong number of arguments".to_string(),
        }
    }
    pub(crate) fn fn_iserr(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        if args.len() == 1 {
            match self.evaluate_node_in_context(&args[0], sheet, column_ref, row_ref) {
                CalcResult::Error { error, .. } => {
                    if Error::NA == error {
                        return CalcResult::Boolean(false);
                    } else {
                        return CalcResult::Boolean(true);
                    }
                }
                _ => {
                    return CalcResult::Boolean(false);
                }
            };
        }
        CalcResult::Error {
            error: Error::ERROR,
            origin: CellReference {
                sheet,
                row: row_ref,
                column: column_ref,
            },
            message: "Wrong number of arguments".to_string(),
        }
    }
    pub(crate) fn fn_isna(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        if args.len() == 1 {
            match self.evaluate_node_in_context(&args[0], sheet, column_ref, row_ref) {
                CalcResult::Error { error, .. } => {
                    if error == Error::NA {
                        return CalcResult::Boolean(true);
                    } else {
                        return CalcResult::Boolean(false);
                    }
                }
                _ => {
                    return CalcResult::Boolean(false);
                }
            };
        }
        CalcResult::Error {
            error: Error::ERROR,
            origin: CellReference {
                sheet,
                row: row_ref,
                column: column_ref,
            },
            message: "Wrong number of arguments".to_string(),
        }
    }
}
