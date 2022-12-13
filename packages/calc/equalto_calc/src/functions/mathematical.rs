use crate::{
    calc_result::{CalcResult, CellReference},
    expressions::parser::Node,
    expressions::{
        token::Error,
        utils::{LAST_COLUMN, LAST_ROW},
    },
    model::Model,
};

impl Model {
    pub(crate) fn fn_min(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        let mut result = f64::NAN;
        for arg in args {
            match self.evaluate_node_in_context(arg, sheet, column_ref, row_ref) {
                CalcResult::Number(value) => result = value.min(result),
                CalcResult::Range { left, right } => {
                    if left.sheet != right.sheet {
                        return CalcResult::new_error(
                            Error::VALUE,
                            sheet,
                            row_ref,
                            column_ref,
                            "Ranges are in different sheets".to_string(),
                        );
                    }
                    for row in left.row..(right.row + 1) {
                        for column in left.column..(right.column + 1) {
                            match self.evaluate_cell(left.sheet, row, column) {
                                CalcResult::Number(value) => {
                                    result = value.min(result);
                                }
                                CalcResult::Error {
                                    error,
                                    origin,
                                    message,
                                } => {
                                    return CalcResult::Error {
                                        error,
                                        origin,
                                        message,
                                    };
                                }
                                _ => {
                                    // We ignore booleans and strings
                                }
                            }
                        }
                    }
                }
                CalcResult::Error {
                    error,
                    origin,
                    message,
                } => {
                    return CalcResult::Error {
                        error,
                        origin,
                        message,
                    };
                }
                _ => {
                    // We ignore booleans and strings
                }
            };
        }
        if result.is_nan() {
            return CalcResult::Number(0.0);
        }
        CalcResult::Number(result)
    }

    pub(crate) fn fn_max(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        let mut result = f64::NAN;
        for arg in args {
            match self.evaluate_node_in_context(arg, sheet, column_ref, row_ref) {
                CalcResult::Number(value) => result = value.max(result),
                CalcResult::Range { left, right } => {
                    if left.sheet != right.sheet {
                        return CalcResult::new_error(
                            Error::VALUE,
                            sheet,
                            row_ref,
                            column_ref,
                            "Ranges are in different sheets".to_string(),
                        );
                    }
                    for row in left.row..(right.row + 1) {
                        for column in left.column..(right.column + 1) {
                            match self.evaluate_cell(left.sheet, row, column) {
                                CalcResult::Number(value) => {
                                    result = value.max(result);
                                }
                                CalcResult::Error {
                                    error,
                                    origin,
                                    message,
                                } => {
                                    return CalcResult::Error {
                                        error,
                                        origin,
                                        message,
                                    };
                                }
                                _ => {
                                    // We ignore booleans and strings
                                }
                            }
                        }
                    }
                }
                CalcResult::Error {
                    error,
                    origin,
                    message,
                } => {
                    return CalcResult::Error {
                        error,
                        origin,
                        message,
                    };
                }
                _ => {
                    // We ignore booleans and strings
                }
            };
        }
        if result.is_nan() {
            return CalcResult::Number(0.0);
        }
        CalcResult::Number(result)
    }

    pub(crate) fn fn_sum(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        let mut result = 0.0;
        for arg in args {
            match self.evaluate_node_in_context(arg, sheet, column_ref, row_ref) {
                CalcResult::Number(value) => result += value,
                CalcResult::Range { left, right } => {
                    if left.sheet != right.sheet {
                        return CalcResult::new_error(
                            Error::VALUE,
                            sheet,
                            row_ref,
                            column_ref,
                            "Ranges are in different sheets".to_string(),
                        );
                    }
                    // TODO: We should do this for all functions that run through ranges
                    // Running cargo test for the equalto_xlsx takes around .8 seconds with this speedup
                    // and ~ 3.5 seconds without it. Note that once properly in place get_sheet_dimension should be almost a noop
                    let row1 = left.row;
                    let mut row2 = right.row;
                    let column1 = left.column;
                    let mut column2 = right.column;
                    if row1 == 1 && row2 == LAST_ROW {
                        let (_, _, row_max, _) = self.get_sheet_dimension(left.sheet);
                        row2 = row_max;
                    }
                    if column1 == 1 && column2 == LAST_COLUMN {
                        let (_, _, _, column_max) = self.get_sheet_dimension(left.sheet);
                        column2 = column_max;
                    }
                    for row in row1..row2 + 1 {
                        for column in column1..(column2 + 1) {
                            match self.evaluate_cell(left.sheet, row, column) {
                                CalcResult::Number(value) => {
                                    result += value;
                                }
                                CalcResult::Error {
                                    error,
                                    origin,
                                    message,
                                } => {
                                    return CalcResult::Error {
                                        error,
                                        origin,
                                        message,
                                    };
                                }
                                _ => {
                                    // We ignore booleans and strings
                                }
                            }
                        }
                    }
                }
                CalcResult::Error {
                    error,
                    origin,
                    message,
                } => {
                    return CalcResult::Error {
                        error,
                        origin,
                        message,
                    };
                }
                _ => {
                    // We ignore booleans and strings
                }
            };
        }
        CalcResult::Number(result)
    }

    /// SUMIF(criteria_range, criteria, [sum_range])
    /// if sum_rage is missing then criteria_range will be used
    pub(crate) fn fn_sumif(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        if args.len() == 2 {
            let arguments = vec![args[0].clone(), args[0].clone(), args[1].clone()];
            self.fn_sumifs(&arguments, sheet, column_ref, row_ref)
        } else if args.len() == 3 {
            let arguments = vec![args[2].clone(), args[0].clone(), args[1].clone()];
            self.fn_sumifs(&arguments, sheet, column_ref, row_ref)
        } else {
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

    /// SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
    pub(crate) fn fn_sumifs(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        let mut total = 0.0;
        let sum = |value| total += value;
        if let Err(e) = self.apply_ifs(args, sheet, column_ref, row_ref, sum) {
            return e;
        }
        CalcResult::Number(total)
    }

    pub(crate) fn fn_round(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        if args.len() != 2 {
            // Incorrect number of arguments
            return CalcResult::Error {
                error: Error::ERROR,
                origin: CellReference {
                    sheet,
                    row: row_ref,
                    column: column_ref,
                },
                message: "Wrong number of arguments".to_string(),
            };
        }
        let value = match self.get_number(&args[0], sheet, column_ref, row_ref) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let number_of_digits = match self.get_number(&args[1], sheet, column_ref, row_ref) {
            Ok(f) => {
                if f > 0.0 {
                    f.floor()
                } else {
                    f.ceil()
                }
            }
            Err(s) => return s,
        };
        let scale = 10.0_f64.powf(number_of_digits);
        CalcResult::Number((value * scale).round() / scale)
    }
    pub(crate) fn fn_roundup(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        if args.len() != 2 {
            // Incorrect number of arguments
            return CalcResult::Error {
                error: Error::ERROR,
                origin: CellReference {
                    sheet,
                    row: row_ref,
                    column: column_ref,
                },
                message: "Wrong number of arguments".to_string(),
            };
        }
        let value = match self.get_number(&args[0], sheet, column_ref, row_ref) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let number_of_digits = match self.get_number(&args[1], sheet, column_ref, row_ref) {
            Ok(f) => {
                if f > 0.0 {
                    f.floor()
                } else {
                    f.ceil()
                }
            }
            Err(s) => return s,
        };
        let scale = 10.0_f64.powf(number_of_digits);
        if value > 0.0 {
            CalcResult::Number((value * scale).ceil() / scale)
        } else {
            CalcResult::Number((value * scale).floor() / scale)
        }
    }
    pub(crate) fn fn_rounddown(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        if args.len() != 2 {
            // Incorrect number of arguments
            return CalcResult::Error {
                error: Error::ERROR,
                origin: CellReference {
                    sheet,
                    row: row_ref,
                    column: column_ref,
                },
                message: "Wrong number of arguments".to_string(),
            };
        }
        let value = match self.get_number(&args[0], sheet, column_ref, row_ref) {
            Ok(f) => f,
            Err(s) => return s,
        };
        let number_of_digits = match self.get_number(&args[1], sheet, column_ref, row_ref) {
            Ok(f) => {
                if f > 0.0 {
                    f.floor()
                } else {
                    f.ceil()
                }
            }
            Err(s) => return s,
        };
        let scale = 10.0_f64.powf(number_of_digits);
        if value > 0.0 {
            CalcResult::Number((value * scale).floor() / scale)
        } else {
            CalcResult::Number((value * scale).ceil() / scale)
        }
    }
}
