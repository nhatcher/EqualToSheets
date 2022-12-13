use crate::{
    calc_result::{CalcResult, CellReference},
    expressions::parser::Node,
    expressions::token::Error,
    model::Model,
};

use super::util::from_wildcard_to_regex;

/// Finds the first instance of 'search_for' in text starting at char index start
fn find(search_for: &str, text: &str, start: usize) -> Option<i32> {
    let ch = text.chars();
    let mut byte_index = 0;
    for (char_index, c) in ch.enumerate() {
        if char_index + 1 >= start && text[byte_index..].starts_with(search_for) {
            return Some((char_index + 1) as i32);
        }
        byte_index += c.len_utf8();
    }
    None
}

/// You can use the wildcard characters — the question mark (?) and asterisk (*) — in the find_text argument.
/// * A question mark matches any single character.
/// * An asterisk matches any sequence of characters.
/// * If you want to find an actual question mark or asterisk, type a tilde (~) before the character.
fn search(search_for: &str, text: &str, start: usize) -> Option<i32> {
    let re = match from_wildcard_to_regex(search_for, false) {
        Ok(r) => r,
        Err(_) => return None,
    };

    let ch = text.chars();
    let mut byte_index = 0;
    for (char_index, c) in ch.enumerate() {
        if char_index + 1 >= start {
            if let Some(m) = re.find(&text[byte_index..]) {
                return Some((text[0..(m.start() + byte_index)].chars().count() as i32) + 1);
            } else {
                return None;
            }
        }
        byte_index += c.len_utf8();
    }
    None
}

impl Model {
    pub(crate) fn fn_concat(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        let mut result = "".to_string();
        for arg in args {
            match self.evaluate_node_in_context(arg, sheet, column_ref, row_ref) {
                CalcResult::String(value) => result = format!("{}{}", result, value),
                CalcResult::Number(value) => result = format!("{}{}", result, value),
                CalcResult::EmptyCell | CalcResult::EmptyArg => {}
                CalcResult::Boolean(value) => {
                    if value {
                        result = format!("{}TRUE", result);
                    } else {
                        result = format!("{}FALSE", result);
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
                                CalcResult::String(value) => {
                                    result = format!("{}{}", result, value);
                                }
                                CalcResult::Number(value) => {
                                    result = format!("{}{}", result, value)
                                }
                                CalcResult::Boolean(value) => {
                                    if value {
                                        result = format!("{}TRUE", result);
                                    } else {
                                        result = format!("{}FALSE", result);
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
                                CalcResult::EmptyCell | CalcResult::EmptyArg => {}
                                CalcResult::Range { .. } => {}
                            }
                        }
                    }
                }
            };
        }
        CalcResult::String(result)
    }
    pub(crate) fn fn_text(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        if args.len() == 2 {
            let value = match self.evaluate_node_in_context(&args[0], sheet, column_ref, row_ref) {
                CalcResult::Number(f) => f,
                CalcResult::String(s) => {
                    return CalcResult::String(s);
                }
                CalcResult::Boolean(b) => {
                    return CalcResult::Boolean(b);
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
                CalcResult::Range { .. } => {
                    // Implicit Intersection not implemented
                    return CalcResult::Error {
                        error: Error::NIMPL,
                        origin: CellReference {
                            sheet,
                            row: row_ref,
                            column: column_ref,
                        },
                        message: "Implicit Intersection not implemented".to_string(),
                    };
                }
                CalcResult::EmptyCell | CalcResult::EmptyArg => 0.0,
            };
            let format_code = match self.get_string(&args[1], sheet, column_ref, row_ref) {
                Ok(s) => s,
                Err(s) => return s,
            };
            let d = self.format_number(value, format_code);
            if let Some(_e) = d.error {
                return CalcResult::Error {
                    error: Error::VALUE,
                    origin: CellReference {
                        sheet,
                        row: row_ref,
                        column: column_ref,
                    },
                    message: "Invalid format code".to_string(),
                };
            }
            CalcResult::String(d.text)
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

    /// FIND(find_text, within_text, [start_num])
    ///  * FIND and FINDB are case sensitive and don't allow wildcard characters.
    ///  * If find_text is "" (empty text), FIND matches the first character in the search string (that is, the character numbered start_num or 1).
    ///  * Find_text cannot contain any wildcard characters.
    ///  * If find_text does not appear in within_text, FIND and FINDB return the #VALUE! error value.
    ///  * If start_num is not greater than zero, FIND and FINDB return the #VALUE! error value.
    ///  * If start_num is greater than the length of within_text, FIND and FINDB return the #VALUE! error value.
    /// NB: FINDB is not implemented. It is the same as FIND function unless locale is a DBCS (Double Byte Character Set)
    pub(crate) fn fn_find(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        if args.len() < 2 || args.len() > 3 {
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
        let find_text = match self.get_string(&args[0], sheet, column_ref, row_ref) {
            Ok(s) => s,
            Err(s) => return s,
        };
        let within_text = match self.get_string(&args[1], sheet, column_ref, row_ref) {
            Ok(s) => s,
            Err(s) => return s,
        };
        let start_num = if args.len() == 3 {
            match self.get_number(&args[2], sheet, column_ref, row_ref) {
                Ok(s) => s.floor(),
                Err(s) => return s,
            }
        } else {
            1.0
        };

        if start_num < 1.0 {
            return CalcResult::Error {
                error: Error::VALUE,
                origin: CellReference {
                    sheet,
                    row: row_ref,
                    column: column_ref,
                },
                message: "Start num must be >= 1".to_string(),
            };
        }
        let start_num = start_num as usize;

        if start_num > within_text.len() {
            return CalcResult::Error {
                error: Error::VALUE,
                origin: CellReference {
                    sheet,
                    row: row_ref,
                    column: column_ref,
                },
                message: "Start num greater than length".to_string(),
            };
        }
        if let Some(s) = find(&find_text, &within_text, start_num) {
            CalcResult::Number(s as f64)
        } else {
            CalcResult::Error {
                error: Error::VALUE,
                origin: CellReference {
                    sheet,
                    row: row_ref,
                    column: column_ref,
                },
                message: "Text not found".to_string(),
            }
        }
    }

    /// Same API as FIND but:
    ///  * Allows wildcards
    ///  * It is case insensitive
    /// SEARCH(find_text, within_text, [start_num])
    pub(crate) fn fn_search(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        if args.len() < 2 || args.len() > 3 {
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
        let find_text = match self.get_string(&args[0], sheet, column_ref, row_ref) {
            Ok(s) => s,
            Err(s) => return s,
        };
        let within_text = match self.get_string(&args[1], sheet, column_ref, row_ref) {
            Ok(s) => s,
            Err(s) => return s,
        };
        let start_num = if args.len() == 3 {
            match self.get_number(&args[2], sheet, column_ref, row_ref) {
                Ok(s) => s.floor(),
                Err(s) => return s,
            }
        } else {
            1.0
        };

        if start_num < 1.0 {
            return CalcResult::Error {
                error: Error::VALUE,
                origin: CellReference {
                    sheet,
                    row: row_ref,
                    column: column_ref,
                },
                message: "Start num must be >= 1".to_string(),
            };
        }
        let start_num = start_num as usize;

        if start_num > within_text.len() {
            return CalcResult::Error {
                error: Error::VALUE,
                origin: CellReference {
                    sheet,
                    row: row_ref,
                    column: column_ref,
                },
                message: "Start num greater than length".to_string(),
            };
        }
        // SEARCH is case insensitive
        if let Some(s) = search(
            &find_text.to_lowercase(),
            &within_text.to_lowercase(),
            start_num,
        ) {
            CalcResult::Number(s as f64)
        } else {
            CalcResult::Error {
                error: Error::VALUE,
                origin: CellReference {
                    sheet,
                    row: row_ref,
                    column: column_ref,
                },
                message: "Text not found".to_string(),
            }
        }
    }

    // LEN, LEFT, RIGHT, MID, LOWER, UPPER, TRIM
    pub(crate) fn fn_len(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        if args.len() == 1 {
            let s = match self.evaluate_node_in_context(&args[0], sheet, column_ref, row_ref) {
                CalcResult::Number(v) => format!("{}", v),
                CalcResult::String(v) => v,
                CalcResult::Boolean(b) => {
                    if b {
                        "TRUE".to_string()
                    } else {
                        "FALSE".to_string()
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
                    }
                }
                CalcResult::Range { .. } => {
                    // Implicit Intersection not implemented
                    return CalcResult::Error {
                        error: Error::NIMPL,
                        origin: CellReference {
                            sheet,
                            row: row_ref,
                            column: column_ref,
                        },
                        message: "Implicit Intersection not implemented".to_string(),
                    };
                }
                CalcResult::EmptyCell | CalcResult::EmptyArg => "".to_string(),
            };
            return CalcResult::Number(s.chars().count() as f64);
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

    pub(crate) fn fn_trim(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        if args.len() == 1 {
            let s = match self.evaluate_node_in_context(&args[0], sheet, column_ref, row_ref) {
                CalcResult::Number(v) => format!("{}", v),
                CalcResult::String(v) => v,
                CalcResult::Boolean(b) => {
                    if b {
                        "TRUE".to_string()
                    } else {
                        "FALSE".to_string()
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
                    }
                }
                CalcResult::Range { .. } => {
                    // Implicit Intersection not implemented
                    return CalcResult::Error {
                        error: Error::NIMPL,
                        origin: CellReference {
                            sheet,
                            row: row_ref,
                            column: column_ref,
                        },
                        message: "Implicit Intersection not implemented".to_string(),
                    };
                }
                CalcResult::EmptyCell | CalcResult::EmptyArg => "".to_string(),
            };
            return CalcResult::String(s.trim().to_owned());
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

    pub(crate) fn fn_lower(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        if args.len() == 1 {
            let s = match self.evaluate_node_in_context(&args[0], sheet, column_ref, row_ref) {
                CalcResult::Number(v) => format!("{}", v),
                CalcResult::String(v) => v,
                CalcResult::Boolean(b) => {
                    if b {
                        "TRUE".to_string()
                    } else {
                        "FALSE".to_string()
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
                    }
                }
                CalcResult::Range { .. } => {
                    // Implicit Intersection not implemented
                    return CalcResult::Error {
                        error: Error::NIMPL,
                        origin: CellReference {
                            sheet,
                            row: row_ref,
                            column: column_ref,
                        },
                        message: "Implicit Intersection not implemented".to_string(),
                    };
                }
                CalcResult::EmptyCell | CalcResult::EmptyArg => "".to_string(),
            };
            return CalcResult::String(s.to_lowercase());
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

    pub(crate) fn fn_upper(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        if args.len() == 1 {
            let s = match self.evaluate_node_in_context(&args[0], sheet, column_ref, row_ref) {
                CalcResult::Number(v) => format!("{}", v),
                CalcResult::String(v) => v,
                CalcResult::Boolean(b) => {
                    if b {
                        "TRUE".to_string()
                    } else {
                        "FALSE".to_string()
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
                    }
                }
                CalcResult::Range { .. } => {
                    // Implicit Intersection not implemented
                    return CalcResult::Error {
                        error: Error::NIMPL,
                        origin: CellReference {
                            sheet,
                            row: row_ref,
                            column: column_ref,
                        },
                        message: "Implicit Intersection not implemented".to_string(),
                    };
                }
                CalcResult::EmptyCell | CalcResult::EmptyArg => "".to_string(),
            };
            return CalcResult::String(s.to_uppercase());
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

    pub(crate) fn fn_left(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        if args.len() > 2 || args.is_empty() {
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
        let s = match self.evaluate_node_in_context(&args[0], sheet, column_ref, row_ref) {
            CalcResult::Number(v) => format!("{}", v),
            CalcResult::String(v) => v,
            CalcResult::Boolean(b) => {
                if b {
                    "TRUE".to_string()
                } else {
                    "FALSE".to_string()
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
                }
            }
            CalcResult::Range { .. } => {
                // Implicit Intersection not implemented
                return CalcResult::Error {
                    error: Error::NIMPL,
                    origin: CellReference {
                        sheet,
                        row: row_ref,
                        column: column_ref,
                    },
                    message: "Implicit Intersection not implemented".to_string(),
                };
            }
            CalcResult::EmptyCell | CalcResult::EmptyArg => "".to_string(),
        };
        let num_chars = if args.len() == 2 {
            match self.evaluate_node_in_context(&args[1], sheet, column_ref, row_ref) {
                CalcResult::Number(v) => {
                    if v < 0.0 {
                        return CalcResult::Error {
                            error: Error::VALUE,
                            origin: CellReference {
                                sheet,
                                row: row_ref,
                                column: column_ref,
                            },
                            message: "Number must be >= 0".to_string(),
                        };
                    }
                    v.floor() as usize
                }
                CalcResult::String(_) => {
                    return CalcResult::Error {
                        error: Error::VALUE,
                        origin: CellReference {
                            sheet,
                            row: row_ref,
                            column: column_ref,
                        },
                        message: "Expecting number".to_string(),
                    };
                }
                CalcResult::Boolean(_) => {
                    return CalcResult::Error {
                        error: Error::VALUE,
                        origin: CellReference {
                            sheet,
                            row: row_ref,
                            column: column_ref,
                        },
                        message: "Expecting number".to_string(),
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
                    }
                }
                CalcResult::Range { .. } => {
                    // Implicit Intersection not implemented
                    return CalcResult::Error {
                        error: Error::NIMPL,
                        origin: CellReference {
                            sheet,
                            row: row_ref,
                            column: column_ref,
                        },
                        message: "Implicit Intersection not implemented".to_string(),
                    };
                }
                CalcResult::EmptyCell | CalcResult::EmptyArg => 0,
            }
        } else {
            1
        };
        let mut result = "".to_string();
        for (index, ch) in s.chars().enumerate() {
            if index >= num_chars {
                break;
            }
            result.push(ch);
        }
        CalcResult::String(result)
    }

    pub(crate) fn fn_right(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        if args.len() > 2 || args.is_empty() {
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
        let s = match self.evaluate_node_in_context(&args[0], sheet, column_ref, row_ref) {
            CalcResult::Number(v) => format!("{}", v),
            CalcResult::String(v) => v,
            CalcResult::Boolean(b) => {
                if b {
                    "TRUE".to_string()
                } else {
                    "FALSE".to_string()
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
                }
            }
            CalcResult::Range { .. } => {
                // Implicit Intersection not implemented
                return CalcResult::Error {
                    error: Error::NIMPL,
                    origin: CellReference {
                        sheet,
                        row: row_ref,
                        column: column_ref,
                    },
                    message: "Implicit Intersection not implemented".to_string(),
                };
            }
            CalcResult::EmptyCell | CalcResult::EmptyArg => "".to_string(),
        };
        let num_chars = if args.len() == 2 {
            match self.evaluate_node_in_context(&args[1], sheet, column_ref, row_ref) {
                CalcResult::Number(v) => {
                    if v < 0.0 {
                        return CalcResult::Error {
                            error: Error::VALUE,
                            origin: CellReference {
                                sheet,
                                row: row_ref,
                                column: column_ref,
                            },
                            message: "Number must be >= 0".to_string(),
                        };
                    }
                    v.floor() as usize
                }
                CalcResult::String(_) => {
                    return CalcResult::Error {
                        error: Error::VALUE,
                        origin: CellReference {
                            sheet,
                            row: row_ref,
                            column: column_ref,
                        },
                        message: "Expecting number".to_string(),
                    };
                }
                CalcResult::Boolean(_) => {
                    return CalcResult::Error {
                        error: Error::VALUE,
                        origin: CellReference {
                            sheet,
                            row: row_ref,
                            column: column_ref,
                        },
                        message: "Expecting number".to_string(),
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
                    }
                }
                CalcResult::Range { .. } => {
                    // Implicit Intersection not implemented
                    return CalcResult::Error {
                        error: Error::NIMPL,
                        origin: CellReference {
                            sheet,
                            row: row_ref,
                            column: column_ref,
                        },
                        message: "Implicit Intersection not implemented".to_string(),
                    };
                }
                CalcResult::EmptyCell | CalcResult::EmptyArg => 0,
            }
        } else {
            1
        };
        let mut result = "".to_string();
        for (index, ch) in s.chars().rev().enumerate() {
            if index >= num_chars {
                break;
            }
            result.push(ch);
        }
        return CalcResult::String(result.chars().rev().collect::<String>());
    }

    pub(crate) fn fn_mid(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        if args.len() != 3 {
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
        let s = match self.evaluate_node_in_context(&args[0], sheet, column_ref, row_ref) {
            CalcResult::Number(v) => format!("{}", v),
            CalcResult::String(v) => v,
            CalcResult::Boolean(b) => {
                if b {
                    "TRUE".to_string()
                } else {
                    "FALSE".to_string()
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
                }
            }
            CalcResult::Range { .. } => {
                // Implicit Intersection not implemented
                return CalcResult::Error {
                    error: Error::NIMPL,
                    origin: CellReference {
                        sheet,
                        row: row_ref,
                        column: column_ref,
                    },
                    message: "Implicit Intersection not implemented".to_string(),
                };
            }
            CalcResult::EmptyCell | CalcResult::EmptyArg => "".to_string(),
        };
        let start_num = match self.evaluate_node_in_context(&args[1], sheet, column_ref, row_ref) {
            CalcResult::Number(v) => {
                if v < 1.0 {
                    return CalcResult::Error {
                        error: Error::VALUE,
                        origin: CellReference {
                            sheet,
                            row: row_ref,
                            column: column_ref,
                        },
                        message: "Number must be >= 1".to_string(),
                    };
                }
                v.floor() as usize
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
                }
            }
            CalcResult::Range { .. } => {
                // Implicit Intersection not implemented
                return CalcResult::Error {
                    error: Error::NIMPL,
                    origin: CellReference {
                        sheet,
                        row: row_ref,
                        column: column_ref,
                    },
                    message: "Implicit Intersection not implemented".to_string(),
                };
            }
            _ => {
                return CalcResult::Error {
                    error: Error::VALUE,
                    origin: CellReference {
                        sheet,
                        row: row_ref,
                        column: column_ref,
                    },
                    message: "Expecting number".to_string(),
                };
            }
        };
        let num_chars = match self.evaluate_node_in_context(&args[2], sheet, column_ref, row_ref) {
            CalcResult::Number(v) => {
                if v < 0.0 {
                    return CalcResult::Error {
                        error: Error::VALUE,
                        origin: CellReference {
                            sheet,
                            row: row_ref,
                            column: column_ref,
                        },
                        message: "Number must be >= 0".to_string(),
                    };
                }
                v.floor() as usize
            }
            CalcResult::String(_) => {
                return CalcResult::Error {
                    error: Error::VALUE,
                    origin: CellReference {
                        sheet,
                        row: row_ref,
                        column: column_ref,
                    },
                    message: "Expecting number".to_string(),
                };
            }
            CalcResult::Boolean(_) => {
                return CalcResult::Error {
                    error: Error::VALUE,
                    origin: CellReference {
                        sheet,
                        row: row_ref,
                        column: column_ref,
                    },
                    message: "Expecting number".to_string(),
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
                }
            }
            CalcResult::Range { .. } => {
                // Implicit Intersection not implemented
                return CalcResult::Error {
                    error: Error::NIMPL,
                    origin: CellReference {
                        sheet,
                        row: row_ref,
                        column: column_ref,
                    },
                    message: "Implicit Intersection not implemented".to_string(),
                };
            }
            CalcResult::EmptyCell | CalcResult::EmptyArg => 0,
        };
        let mut result = "".to_string();
        let mut count: usize = 0;
        for (index, ch) in s.chars().enumerate() {
            if count >= num_chars {
                break;
            }
            if index + 1 >= start_num {
                result.push(ch);
                count += 1;
            }
        }
        CalcResult::String(result)
    }
}
