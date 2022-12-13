use crate::{
    calc_result::{CalcResult, CellReference, Range},
    expressions::parser::Node,
    expressions::{
        token::Error,
        utils::{LAST_COLUMN, LAST_ROW},
    },
    model::Model,
};

use super::util::build_criteria;

impl Model {
    pub(crate) fn fn_average(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        if args.is_empty() {
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
        let mut count = 0.0;
        let mut sum = 0.0;
        for arg in args {
            match self.evaluate_node_in_context(arg, sheet, column_ref, row_ref) {
                CalcResult::Number(value) => {
                    count += 1.0;
                    sum += value;
                }
                CalcResult::Boolean(b) => {
                    if let Node::ReferenceKind { .. } = arg {
                    } else {
                        sum += if b { 1.0 } else { 0.0 };
                        count += 1.0;
                    }
                }
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
                                    count += 1.0;
                                    sum += value;
                                }
                                CalcResult::Error {
                                    error,
                                    origin,
                                    message,
                                } => {
                                    return CalcResult::new_error(
                                        error,
                                        origin.sheet,
                                        origin.row,
                                        origin.column,
                                        message,
                                    );
                                }
                                CalcResult::Range { .. } => {
                                    return CalcResult::new_error(
                                        Error::ERROR,
                                        sheet,
                                        row_ref,
                                        column_ref,
                                        "Unexpected Range".to_string(),
                                    );
                                }
                                _ => {}
                            }
                        }
                    }
                }
                CalcResult::Error {
                    error,
                    origin,
                    message,
                } => {
                    return CalcResult::new_error(
                        error,
                        origin.sheet,
                        origin.row,
                        origin.column,
                        message,
                    );
                }
                CalcResult::String(s) => {
                    if let Node::ReferenceKind { .. } = arg {
                        // Do nothing
                    } else if let Ok(t) = s.parse::<f64>() {
                        sum += t;
                        count += 1.0;
                    } else {
                        return CalcResult::Error {
                            error: Error::VALUE,
                            origin: CellReference {
                                sheet,
                                row: row_ref,
                                column: column_ref,
                            },
                            message: "Argument cannot be cast into number".to_string(),
                        };
                    }
                }
                _ => {
                    // Ignore everything else
                }
            };
        }
        if count == 0.0 {
            return CalcResult::Error {
                error: Error::DIV,
                origin: CellReference {
                    sheet,
                    row: row_ref,
                    column: column_ref,
                },
                message: "Division by Zero".to_string(),
            };
        }
        CalcResult::Number(sum / count)
    }

    pub(crate) fn fn_averagea(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        if args.is_empty() {
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
        let mut count = 0.0;
        let mut sum = 0.0;
        for arg in args {
            match self.evaluate_node_in_context(arg, sheet, column_ref, row_ref) {
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
                                CalcResult::String(_) => count += 1.0,
                                CalcResult::Number(value) => {
                                    count += 1.0;
                                    sum += value;
                                }
                                CalcResult::Boolean(b) => {
                                    if b {
                                        sum += 1.0;
                                    }
                                    count += 1.0;
                                }
                                CalcResult::Error {
                                    error,
                                    origin,
                                    message,
                                } => {
                                    return CalcResult::new_error(
                                        error,
                                        origin.sheet,
                                        origin.row,
                                        origin.column,
                                        message,
                                    );
                                }
                                CalcResult::Range { .. } => {
                                    return CalcResult::new_error(
                                        Error::ERROR,
                                        sheet,
                                        row_ref,
                                        column_ref,
                                        "Unexpected Range".to_string(),
                                    );
                                }
                                CalcResult::EmptyCell | CalcResult::EmptyArg => {}
                            }
                        }
                    }
                }
                CalcResult::Number(value) => {
                    count += 1.0;
                    sum += value;
                }
                CalcResult::String(s) => {
                    if let Node::ReferenceKind { .. } = arg {
                        // Do nothing
                        count += 1.0;
                    } else if let Ok(t) = s.parse::<f64>() {
                        sum += t;
                        count += 1.0;
                    } else {
                        return CalcResult::Error {
                            error: Error::VALUE,
                            origin: CellReference {
                                sheet,
                                row: row_ref,
                                column: column_ref,
                            },
                            message: "Argument cannot be cast into number".to_string(),
                        };
                    }
                }
                CalcResult::Boolean(b) => {
                    count += 1.0;
                    if b {
                        sum += 1.0;
                    }
                }
                CalcResult::Error {
                    error,
                    origin,
                    message,
                } => {
                    return CalcResult::new_error(
                        error,
                        origin.sheet,
                        origin.row,
                        origin.column,
                        message,
                    );
                }
                CalcResult::EmptyCell | CalcResult::EmptyArg => {}
            };
        }
        if count == 0.0 {
            return CalcResult::Error {
                error: Error::DIV,
                origin: CellReference {
                    sheet,
                    row: row_ref,
                    column: column_ref,
                },
                message: "Division by Zero".to_string(),
            };
        }
        CalcResult::Number(sum / count)
    }

    pub(crate) fn fn_count(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        if args.is_empty() {
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
        let mut result = 0.0;
        for arg in args {
            match self.evaluate_node_in_context(arg, sheet, column_ref, row_ref) {
                CalcResult::Number(_) => {
                    result += 1.0;
                }
                CalcResult::Boolean(_) => {
                    if !matches!(arg, Node::ReferenceKind { .. }) {
                        result += 1.0;
                    }
                }
                CalcResult::String(s) => {
                    if !matches!(arg, Node::ReferenceKind { .. }) && s.parse::<f64>().is_ok() {
                        result += 1.0;
                    }
                }
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
                            if let CalcResult::Number(_) =
                                self.evaluate_cell(left.sheet, row, column)
                            {
                                result += 1.0;
                            }
                        }
                    }
                }
                _ => {
                    // Ignore everything else
                }
            };
        }
        CalcResult::Number(result)
    }

    pub(crate) fn fn_counta(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        if args.is_empty() {
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
        let mut result = 0.0;
        for arg in args {
            match self.evaluate_node_in_context(arg, sheet, column_ref, row_ref) {
                CalcResult::EmptyCell | CalcResult::EmptyArg => {}
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
                                CalcResult::EmptyCell | CalcResult::EmptyArg => {}
                                _ => {
                                    result += 1.0;
                                }
                            }
                        }
                    }
                }
                _ => {
                    result += 1.0;
                }
            };
        }
        CalcResult::Number(result)
    }

    pub(crate) fn fn_countblank(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        // COUNTBLANK requires only one argument
        if args.len() != 1 {
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
        let mut result = 0.0;
        for arg in args {
            match self.evaluate_node_in_context(arg, sheet, column_ref, row_ref) {
                CalcResult::EmptyCell | CalcResult::EmptyArg => result += 1.0,
                CalcResult::String(s) => {
                    if s.is_empty() {
                        result += 1.0
                    }
                }
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
                                CalcResult::EmptyCell | CalcResult::EmptyArg => result += 1.0,
                                CalcResult::String(s) => {
                                    if s.is_empty() {
                                        result += 1.0
                                    }
                                }
                                _ => {}
                            }
                        }
                    }
                }
                _ => {}
            };
        }
        CalcResult::Number(result)
    }

    pub(crate) fn fn_countif(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        if args.len() == 2 {
            let arguments = vec![args[0].clone(), args[1].clone()];
            self.fn_countifs(&arguments, sheet, column_ref, row_ref)
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

    /// AVERAGEIF(criteria_range, criteria, [average_range])
    /// if average_rage is missing then criteria_range will be used
    pub(crate) fn fn_averageif(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        if args.len() == 2 {
            let arguments = vec![args[0].clone(), args[0].clone(), args[1].clone()];
            self.fn_averageifs(&arguments, sheet, column_ref, row_ref)
        } else if args.len() == 3 {
            let arguments = vec![args[2].clone(), args[0].clone(), args[1].clone()];
            self.fn_averageifs(&arguments, sheet, column_ref, row_ref)
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

    // FIXME: This function shares a lot of code with apply_ifs. Can we merge them?
    pub(crate) fn fn_countifs(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        let args_count = args.len();
        if args_count < 2 || args_count % 2 == 1 {
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

        let case_count = args_count / 2;
        // NB: this is a beautiful example of the borrow checker
        // The order of these two definitions cannot be swapped.
        let mut criteria = Vec::new();
        let mut fn_criteria = Vec::new();
        let ranges = &mut Vec::new();
        for case_index in 0..case_count {
            let criterion = self.evaluate_node_in_context(
                &args[case_index * 2 + 1],
                sheet,
                column_ref,
                row_ref,
            );
            criteria.push(criterion);
            // NB: We cannot do:
            // fn_criteria.push(build_criteria(&criterion));
            // because criterion doesn't live long enough
            let result =
                self.evaluate_node_in_context(&args[case_index * 2], sheet, column_ref, row_ref);
            if result.is_error() {
                return result;
            }
            if let CalcResult::Range { left, right } = result {
                if left.sheet != right.sheet {
                    return CalcResult::new_error(
                        Error::VALUE,
                        sheet,
                        row_ref,
                        column_ref,
                        "Ranges are in different sheets".to_string(),
                    );
                }
                // TODO test ranges are of the same size as sum_range
                ranges.push(Range { left, right });
            } else {
                return CalcResult::new_error(
                    Error::VALUE,
                    sheet,
                    row_ref,
                    column_ref,
                    "Expected a range".to_string(),
                );
            }
        }
        for criterion in criteria.iter() {
            fn_criteria.push(build_criteria(criterion));
        }

        let mut total = 0.0;
        let first_range = &ranges[0];
        let left_row = first_range.left.row;
        let left_column = first_range.left.column;
        let right_row = first_range.right.row;
        let right_column = first_range.right.column;
        let (_, _, row_max, column_max) = self.get_sheet_dimension(first_range.left.sheet);

        let open_row = left_row == 1 && right_row == LAST_ROW;
        let open_column = left_column == 1 && right_column == LAST_COLUMN;

        for row in left_row..right_row + 1 {
            if open_row && row > row_max {
                // If the row is larger than the max row in the sheet then all cells are empty.
                // We compute it only once
                let mut is_true = true;
                for fn_criterion in fn_criteria.iter() {
                    if !fn_criterion(&CalcResult::EmptyCell) {
                        is_true = false;
                        break;
                    }
                }
                if is_true {
                    total += ((LAST_ROW - row_max) * (right_column - left_column + 1)) as f64;
                }
                break;
            }
            for column in left_column..right_column + 1 {
                if open_column && column > column_max {
                    // If the column is larger than the max column in the sheet then all cells are empty.
                    // We compute it only once
                    let mut is_true = true;
                    for fn_criterion in fn_criteria.iter() {
                        if !fn_criterion(&CalcResult::EmptyCell) {
                            is_true = false;
                            break;
                        }
                    }
                    if is_true {
                        total += (LAST_COLUMN - column_max) as f64;
                    }
                    break;
                }
                let mut is_true = true;
                for case_index in 0..case_count {
                    // We check if value in range n meets criterion n
                    let range = &ranges[case_index];
                    let fn_criterion = &fn_criteria[case_index];
                    let value = self.evaluate_cell(
                        range.left.sheet,
                        range.left.row + row - first_range.left.row,
                        range.left.column + column - first_range.left.column,
                    );
                    if !fn_criterion(&value) {
                        is_true = false;
                        break;
                    }
                }
                if is_true {
                    total += 1.0;
                }
            }
        }
        CalcResult::Number(total)
    }

    pub(crate) fn apply_ifs<F>(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
        mut apply: F,
    ) -> Result<(), CalcResult>
    where
        F: FnMut(f64),
    {
        let args_count = args.len();
        if args_count < 3 || args_count % 2 == 0 {
            // Incorrect number of arguments
            return Err(CalcResult::Error {
                error: Error::ERROR,
                origin: CellReference {
                    sheet,
                    row: row_ref,
                    column: column_ref,
                },
                message: "Wrong number of arguments".to_string(),
            });
        }
        let arg_0 = self.evaluate_node_in_context(&args[0], sheet, column_ref, row_ref);
        if arg_0.is_error() {
            return Err(arg_0);
        }
        let sum_range = if let CalcResult::Range { left, right } = arg_0 {
            if left.sheet != right.sheet {
                return Err(CalcResult::new_error(
                    Error::VALUE,
                    sheet,
                    row_ref,
                    column_ref,
                    "Ranges are in different sheets".to_string(),
                ));
            }
            Range { left, right }
        } else {
            return Err(CalcResult::new_error(
                Error::VALUE,
                sheet,
                row_ref,
                column_ref,
                "Expected a range".to_string(),
            ));
        };

        let case_count = (args_count - 1) / 2;
        // NB: this is a beautiful example of the borrow checker
        // The order of these two definitions cannot be swapped.
        let mut criteria = Vec::new();
        let mut fn_criteria = Vec::new();
        let ranges = &mut Vec::new();
        for case_index in 1..=case_count {
            let criterion =
                self.evaluate_node_in_context(&args[case_index * 2], sheet, column_ref, row_ref);
            // NB: criterion might be an error. That's ok
            criteria.push(criterion);
            // NB: We cannot do:
            // fn_criteria.push(build_criteria(&criterion));
            // because criterion doesn't live long enough
            let result = self.evaluate_node_in_context(
                &args[case_index * 2 - 1],
                sheet,
                column_ref,
                row_ref,
            );
            if result.is_error() {
                return Err(result);
            }
            if let CalcResult::Range { left, right } = result {
                if left.sheet != right.sheet {
                    return Err(CalcResult::new_error(
                        Error::VALUE,
                        sheet,
                        row_ref,
                        column_ref,
                        "Ranges are in different sheets".to_string(),
                    ));
                }
                // TODO test ranges are of the same size as sum_range
                ranges.push(Range { left, right });
            } else {
                return Err(CalcResult::new_error(
                    Error::VALUE,
                    sheet,
                    row_ref,
                    column_ref,
                    "Expected a range".to_string(),
                ));
            }
        }
        for criterion in criteria.iter() {
            fn_criteria.push(build_criteria(criterion));
        }

        let left_row = sum_range.left.row;
        let left_column = sum_range.left.column;
        let mut right_row = sum_range.right.row;
        let mut right_column = sum_range.right.column;

        if left_row == 1 && right_row == LAST_ROW {
            let (_, _, row_max, _) = self.get_sheet_dimension(sum_range.left.sheet);
            right_row = row_max;
        }
        if left_column == 1 && right_column == LAST_COLUMN {
            let (_, _, _, column_max) = self.get_sheet_dimension(sum_range.left.sheet);
            right_column = column_max;
        }

        for row in left_row..right_row + 1 {
            for column in left_column..right_column + 1 {
                let mut is_true = true;
                for case_index in 0..case_count {
                    // We check if value in range n meets criterion n
                    let range = &ranges[case_index];
                    let fn_criterion = &fn_criteria[case_index];
                    let value = self.evaluate_cell(
                        range.left.sheet,
                        range.left.row + row - sum_range.left.row,
                        range.left.column + column - sum_range.left.column,
                    );
                    if !fn_criterion(&value) {
                        is_true = false;
                        break;
                    }
                }
                if is_true {
                    let v = self.evaluate_cell(sum_range.left.sheet, row, column);
                    match v {
                        CalcResult::Number(n) => apply(n),
                        CalcResult::Error { .. } => return Err(v),
                        _ => {}
                    }
                }
            }
        }
        Ok(())
    }

    pub(crate) fn fn_averageifs(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        let mut total = 0.0;
        let mut count = 0.0;

        let average = |value: f64| {
            total += value;
            count += 1.0;
        };
        if let Err(e) = self.apply_ifs(args, sheet, column_ref, row_ref, average) {
            return e;
        }

        if count == 0.0 {
            return CalcResult::Error {
                error: Error::DIV,
                origin: CellReference {
                    sheet,
                    row: row_ref,
                    column: column_ref,
                },
                message: "division by 0".to_string(),
            };
        }
        CalcResult::Number(total / count)
    }

    pub(crate) fn fn_minifs(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        let mut min = f64::INFINITY;
        let apply_min = |value: f64| min = value.min(min);
        if let Err(e) = self.apply_ifs(args, sheet, column_ref, row_ref, apply_min) {
            return e;
        }

        if min.is_infinite() {
            min = 0.0;
        }
        CalcResult::Number(min)
    }

    pub(crate) fn fn_maxifs(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        let mut max = -f64::INFINITY;
        let apply_max = |value: f64| max = value.max(max);
        if let Err(e) = self.apply_ifs(args, sheet, column_ref, row_ref, apply_max) {
            return e;
        }
        if max.is_infinite() {
            max = 0.0;
        }
        CalcResult::Number(max)
    }
}
