use crate::{
    calc_result::{CalcResult, CellReference},
    expressions::parser::Node,
    expressions::token::Error,
    model::Model,
};

impl Model {
    pub(crate) fn fn_isnumber(&mut self, args: &[Node], cell: CellReference) -> CalcResult {
        if args.len() == 1 {
            match self.evaluate_node_in_context(&args[0], cell) {
                CalcResult::Number(_) => return CalcResult::Boolean(true),
                _ => {
                    return CalcResult::Boolean(false);
                }
            };
        }
        CalcResult::new_args_number_error(cell)
    }
    pub(crate) fn fn_istext(&mut self, args: &[Node], cell: CellReference) -> CalcResult {
        if args.len() == 1 {
            match self.evaluate_node_in_context(&args[0], cell) {
                CalcResult::String(_) => return CalcResult::Boolean(true),
                _ => {
                    return CalcResult::Boolean(false);
                }
            };
        }
        CalcResult::new_args_number_error(cell)
    }
    pub(crate) fn fn_isnontext(&mut self, args: &[Node], cell: CellReference) -> CalcResult {
        if args.len() == 1 {
            match self.evaluate_node_in_context(&args[0], cell) {
                CalcResult::String(_) => return CalcResult::Boolean(false),
                _ => {
                    return CalcResult::Boolean(true);
                }
            };
        }
        CalcResult::new_args_number_error(cell)
    }
    pub(crate) fn fn_islogical(&mut self, args: &[Node], cell: CellReference) -> CalcResult {
        if args.len() == 1 {
            match self.evaluate_node_in_context(&args[0], cell) {
                CalcResult::Boolean(_) => return CalcResult::Boolean(true),
                _ => {
                    return CalcResult::Boolean(false);
                }
            };
        }
        CalcResult::new_args_number_error(cell)
    }
    pub(crate) fn fn_isblank(&mut self, args: &[Node], cell: CellReference) -> CalcResult {
        if args.len() == 1 {
            match self.evaluate_node_in_context(&args[0], cell) {
                CalcResult::EmptyCell => return CalcResult::Boolean(true),
                _ => {
                    return CalcResult::Boolean(false);
                }
            };
        }
        CalcResult::new_args_number_error(cell)
    }
    pub(crate) fn fn_iserror(&mut self, args: &[Node], cell: CellReference) -> CalcResult {
        if args.len() == 1 {
            match self.evaluate_node_in_context(&args[0], cell) {
                CalcResult::Error { .. } => return CalcResult::Boolean(true),
                _ => {
                    return CalcResult::Boolean(false);
                }
            };
        }
        CalcResult::new_args_number_error(cell)
    }
    pub(crate) fn fn_iserr(&mut self, args: &[Node], cell: CellReference) -> CalcResult {
        if args.len() == 1 {
            match self.evaluate_node_in_context(&args[0], cell) {
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
        CalcResult::new_args_number_error(cell)
    }
    pub(crate) fn fn_isna(&mut self, args: &[Node], cell: CellReference) -> CalcResult {
        if args.len() == 1 {
            match self.evaluate_node_in_context(&args[0], cell) {
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
        CalcResult::new_args_number_error(cell)
    }
}
