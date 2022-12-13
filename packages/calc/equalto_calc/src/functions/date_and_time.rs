use chrono::Datelike;
use chrono::Months;
use chrono::NaiveDate;

use chrono::NaiveDateTime;
use chrono::TimeZone;

use crate::{
    calc_result::{CalcResult, CellReference},
    expressions::parser::Node,
    expressions::token::Error,
    formatter::dates::from_excel_date,
    model::Model,
};

impl Model {
    pub(crate) fn fn_day(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        let args_count = args.len();
        if args_count != 1 {
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
        let serial_number = match self.get_number(&args[0], sheet, column_ref, row_ref) {
            Ok(c) => {
                let t = c.floor() as i64;
                if t < 0 {
                    return CalcResult::Error {
                        error: Error::NUM,
                        origin: CellReference {
                            sheet,
                            row: row_ref,
                            column: column_ref,
                        },
                        message: "Function DAY parameter 1 value is negative. It should be positive or zero.".to_string(),
                    };
                }
                t
            }
            Err(s) => return s,
        };
        let date = from_excel_date(serial_number);
        let day = date.day() as f64;
        CalcResult::Number(day)
    }

    pub(crate) fn fn_month(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        let args_count = args.len();
        if args_count != 1 {
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
        let serial_number = match self.get_number(&args[0], sheet, column_ref, row_ref) {
            Ok(c) => {
                let t = c.floor() as i64;
                if t < 0 {
                    return CalcResult::Error {
                        error: Error::NUM,
                        origin: CellReference {
                            sheet,
                            row: row_ref,
                            column: column_ref,
                        },
                        message: "Function MONTH parameter 1 value is negative. It should be positive or zero.".to_string(),
                    };
                }
                t
            }
            Err(s) => return s,
        };
        let date = from_excel_date(serial_number);
        let month = date.month() as f64;
        CalcResult::Number(month)
    }

    // year, month, day
    pub(crate) fn fn_date(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        let args_count = args.len();
        if args_count != 3 {
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
        let year = match self.get_number(&args[0], sheet, column_ref, row_ref) {
            Ok(c) => {
                let t = c.floor() as i32;
                if t < 0 {
                    return CalcResult::Error {
                        error: Error::NUM,
                        origin: CellReference {
                            sheet,
                            row: row_ref,
                            column: column_ref,
                        },
                        message: "Out of range parameters for date".to_string(),
                    };
                }
                t
            }
            Err(s) => return s,
        };
        let month = match self.get_number(&args[1], sheet, column_ref, row_ref) {
            Ok(c) => {
                let t = c.floor();
                if t < 0.0 {
                    return CalcResult::Error {
                        error: Error::NUM,
                        origin: CellReference {
                            sheet,
                            row: row_ref,
                            column: column_ref,
                        },
                        message: "Out of range parameters for date".to_string(),
                    };
                }
                t as u32
            }
            Err(s) => return s,
        };
        let day = match self.get_number(&args[2], sheet, column_ref, row_ref) {
            Ok(c) => {
                let t = c.floor();
                if t < 0.0 {
                    return CalcResult::Error {
                        error: Error::NUM,
                        origin: CellReference {
                            sheet,
                            row: row_ref,
                            column: column_ref,
                        },
                        message: "Out of range parameters for date".to_string(),
                    };
                }
                t as u32
            }
            Err(s) => return s,
        };
        let time_delta = NaiveDate::from_ymd(1900, 1, 1).num_days_from_ce() - 2;
        let serial_number = match NaiveDate::from_ymd_opt(year, month, day) {
            Some(native_date) => native_date.num_days_from_ce() - time_delta,
            None => {
                return CalcResult::Error {
                    error: Error::NUM,
                    origin: CellReference {
                        sheet,
                        row: row_ref,
                        column: column_ref,
                    },
                    message: "Out of range parameters for date".to_string(),
                };
            }
        };
        CalcResult::Number(serial_number as f64)
    }

    pub(crate) fn fn_year(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        let args_count = args.len();
        if args_count != 1 {
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
        let serial_number = match self.get_number(&args[0], sheet, column_ref, row_ref) {
            Ok(c) => {
                let t = c.floor() as i64;
                if t < 0 {
                    return CalcResult::Error {
                        error: Error::NUM,
                        origin: CellReference {
                            sheet,
                            row: row_ref,
                            column: column_ref,
                        },
                        message: "Function YEAR parameter 1 value is negative. It should be positive or zero.".to_string(),
                    };
                }
                t
            }
            Err(s) => return s,
        };
        let date = from_excel_date(serial_number);
        let year = date.year() as f64;
        CalcResult::Number(year)
    }

    pub(crate) fn fn_today(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        let args_count = args.len();
        if args_count != 0 {
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
        // milliseconds since January 1, 1970 00:00:00 UTC.
        let milliseconds = (self.env.get_milliseconds_since_epoch)();
        let seconds = milliseconds / 1000;
        let dt = NaiveDateTime::from_timestamp(seconds, 0);
        let local_time = self.tz.from_utc_datetime(&dt);
        // 693_594 is computed as:
        // NaiveDate::from_ymd(1900, 1, 1).num_days_from_ce() - 2
        // The 2 days offset is because of Excel 1900 bug
        let days_from_1900 = local_time.num_days_from_ce() - 693_594;

        CalcResult::Number(days_from_1900 as f64)
    }

    // date, months
    pub(crate) fn fn_edate(
        &mut self,
        args: &[Node],
        sheet: i32,
        column_ref: i32,
        row_ref: i32,
    ) -> CalcResult {
        let args_count = args.len();
        if args_count != 2 {
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
        let serial_number = match self.get_number(&args[0], sheet, column_ref, row_ref) {
            Ok(c) => {
                let t = c.floor() as i64;
                if t < 0 {
                    return CalcResult::Error {
                        error: Error::NUM,
                        origin: CellReference {
                            sheet,
                            row: row_ref,
                            column: column_ref,
                        },
                        message: "Parameter 1 value is negative. It should be positive or zero."
                            .to_string(),
                    };
                }
                t
            }
            Err(s) => return s,
        };

        let months = match self.get_number(&args[1], sheet, column_ref, row_ref) {
            Ok(c) => {
                let t = c.trunc();
                t as i32
            }
            Err(s) => return s,
        };

        let months_abs = months.unsigned_abs();

        let native_date = if months > 0 {
            from_excel_date(serial_number) + Months::new(months_abs)
        } else {
            from_excel_date(serial_number) - Months::new(months_abs)
        };

        let time_delta = NaiveDate::from_ymd(1900, 1, 1).num_days_from_ce() - 2;
        let serial_number = native_date.num_days_from_ce() - time_delta;
        if serial_number < 0 {
            return CalcResult::Error {
                error: Error::NUM,
                origin: CellReference {
                    sheet,
                    row: row_ref,
                    column: column_ref,
                },
                message: "EDATE out of bounds".to_string(),
            };
        }
        CalcResult::Number(serial_number as f64)
    }
}
