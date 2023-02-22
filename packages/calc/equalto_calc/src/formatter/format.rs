use chrono::Datelike;

use crate::{locale::Locale, number_format::to_precision};

use super::{
    dates::from_excel_date,
    parser::{ParsePart, Parser, TextToken},
};

pub struct Formatted {
    pub color: Option<i32>,
    pub text: String,
    pub error: Option<String>,
}

/// Returns the vector of chars of the fractional part of a *positive* number:
/// 3.1415926 ==> ['1', '4', '1', '5', '9', '2', '6']
fn get_fract_part(value: f64, precision: i32) -> Vec<char> {
    let b = format!("{:.1$}", value.fract(), precision as usize)
        .chars()
        .collect::<Vec<char>>();
    let l = b.len() - 1;
    let mut last_non_zero = b.len() - 1;
    for i in 0..l {
        if b[l - i] != '0' {
            last_non_zero = l - i + 1;
            break;
        }
    }
    if last_non_zero < 2 {
        return vec![];
    }
    b[2..last_non_zero].to_vec()
}

/// Return true if we need to add a separator in position digit_index
/// It normally happens if if digit_index -1 is 3, 6, 9,... digit_index ≡ 1 mod 3
fn use_group_separator(use_thousands: bool, digit_index: i32, group_sizes: &str) -> bool {
    if use_thousands {
        if group_sizes == "#,##0.###" {
            if digit_index > 1 && (digit_index - 1) % 3 == 0 {
                return true;
            }
        } else if group_sizes == "#,##,##0.###"
            && (digit_index == 3 || (digit_index > 3 && digit_index % 2 == 0))
        {
            return true;
        }
    }
    false
}

pub fn format_number(value_original: f64, format: &str, locale: &Locale) -> Formatted {
    let mut parser = Parser::new(format);
    parser.parse();
    let parts = parser.parts;
    // There are four parts:
    // 1) Positive numbers
    // 2) Negative numbers
    // 3) Zero
    // 4) Text
    // If you specify only one section of format code, the code in that section is used for all numbers.
    // If you specify two sections of format code, the first section of code is used
    // for positive numbers and zeros, and the second section of code is used for negative numbers.
    // When you skip code sections in your number format,
    // you must include a semicolon for each of the missing sections of code.
    // You can use the ampersand (&) text operator to join, or concatenate, two values.
    let mut value = value_original;
    let part;
    match parts.len() {
        1 => {
            part = &parts[0];
        }
        2 => {
            if value >= 0.0 {
                part = &parts[0]
            } else {
                value = -value;
                part = &parts[1];
            }
        }
        3 => {
            if value > 0.0 {
                part = &parts[0]
            } else if value < 0.0 {
                value = -value;
                part = &parts[1];
            } else {
                value = 0.0;
                part = &parts[2];
            }
        }
        4 => {
            if value > 0.0 {
                part = &parts[0]
            } else if value < 0.0 {
                value = -value;
                part = &parts[1];
            } else {
                value = 0.0;
                part = &parts[2];
            }
        }
        _ => {
            return Formatted {
                text: "#VALUE!".to_owned(),
                color: None,
                error: Some("Too many parts".to_owned()),
            };
        }
    }
    match part {
        ParsePart::Error(..) => Formatted {
            text: "#VALUE!".to_owned(),
            color: None,
            error: Some("Problem parsing format string".to_owned()),
        },
        ParsePart::General(..) => {
            // FIXME: This is "General formatting"
            // We should have different codepaths for general formatting and errors
            let value_abs = value.abs();
            if (1.0e-8..1.0e+11).contains(&value_abs) {
                let mut text = format!("{:.9}", value);
                text = text.trim_end_matches('0').trim_end_matches('.').to_string();
                Formatted {
                    text,
                    color: None,
                    error: None,
                }
            } else {
                if value_abs == 0.0 {
                    return Formatted {
                        text: "0".to_string(),
                        color: None,
                        error: None,
                    };
                }
                let exponent = value_abs.log10().floor();
                value /= 10.0_f64.powf(exponent);
                let sign = if exponent < 0.0 { '-' } else { '+' };
                let s = format!("{:.5}", value);
                Formatted {
                    text: format!(
                        "{}E{}{:02}",
                        s.trim_end_matches('0').trim_end_matches('.'),
                        sign,
                        exponent.abs()
                    ),
                    color: None,
                    error: None,
                }
            }
        }
        ParsePart::Date(p) => {
            let tokens = &p.tokens;
            let mut text = "".to_string();
            if !(1.0..=2_958_465.0).contains(&value) {
                // 2_958_465 is 31 December 9999
                return Formatted {
                    text: "#VALUE!".to_owned(),
                    color: None,
                    error: Some("Date negative or too long".to_owned()),
                };
            }
            let date = from_excel_date(value as i64);
            for token in tokens {
                match token {
                    TextToken::Literal(c) => {
                        text = format!("{}{}", text, c);
                    }
                    TextToken::Text(t) => {
                        text = format!("{}{}", text, t);
                    }
                    TextToken::Ghost(_) => {
                        // we just leave a whitespace
                        // This is what the TEXT function does
                        text = format!("{} ", text);
                    }
                    TextToken::Spacer(_) => {
                        // we just leave a whitespace
                        // This is what the TEXT function does
                        text = format!("{} ", text);
                    }
                    TextToken::Raw => {
                        text = format!("{}{}", text, value);
                    }
                    TextToken::Digit(_) => {}
                    TextToken::Period => {}
                    TextToken::Day => {
                        let day = date.day() as usize;
                        text = format!("{}{}", text, day);
                    }
                    TextToken::DayPadded => {
                        let day = date.day() as usize;
                        text = format!("{}{:02}", text, day);
                    }
                    TextToken::DayNameShort => {
                        let mut day = date.weekday().number_from_monday() as usize;
                        if day == 7 {
                            day = 0;
                        }
                        text = format!("{}{}", text, &locale.dates.day_names_short[day]);
                    }
                    TextToken::DayName => {
                        let mut day = date.weekday().number_from_monday() as usize;
                        if day == 7 {
                            day = 0;
                        }
                        text = format!("{}{}", text, &locale.dates.day_names[day]);
                    }
                    TextToken::Month => {
                        let month = date.month() as usize;
                        text = format!("{}{}", text, month);
                    }
                    TextToken::MonthPadded => {
                        let month = date.month() as usize;
                        text = format!("{}{:02}", text, month);
                    }
                    TextToken::MonthNameShort => {
                        let month = date.month() as usize;
                        text = format!("{}{}", text, &locale.dates.months_short[month - 1]);
                    }
                    TextToken::MonthName => {
                        let month = date.month() as usize;
                        text = format!("{}{}", text, &locale.dates.months[month - 1]);
                    }
                    TextToken::MonthLetter => {
                        let month = date.month() as usize;
                        let months_letter = &locale.dates.months_letter[month - 1];
                        text = format!("{}{}", text, months_letter);
                    }
                    TextToken::YearShort => {
                        text = format!("{}{}", text, date.format("%y"));
                    }
                    TextToken::Year => {
                        text = format!("{}{}", text, date.year());
                    }
                }
            }
            Formatted {
                text,
                color: p.color,
                error: None,
            }
        }
        ParsePart::Number(p) => {
            let mut text = "".to_string();
            let tokens = &p.tokens;
            value = value * 100.0_f64.powi(p.percent) / (1000.0_f64.powi(p.comma));
            // p.precision is the number of significant digits _after_ the decimal point
            value = to_precision(
                value,
                (p.precision as usize) + format!("{}", value.abs().floor()).len(),
            );
            let mut value_abs = value.abs();
            let mut exponent_part: Vec<char> = vec![];
            let mut exponent_is_negative = value_abs < 10.0;
            if p.is_scientific {
                if value_abs == 0.0 {
                    exponent_part = vec!['0'];
                    exponent_is_negative = false;
                } else {
                    // TODO: Implement engineering formatting.
                    let exponent = value_abs.log10().floor();
                    exponent_part = format!("{}", exponent.abs()).chars().collect();
                    value /= 10.0_f64.powf(exponent);
                    value_abs = value.abs();
                }
            }
            let l_exp = exponent_part.len() as i32;
            let mut int_part: Vec<char> = format!("{}", value_abs.floor()).chars().collect();
            if value_abs as i64 == 0 {
                int_part = vec![];
            }
            let fract_part = get_fract_part(value_abs, p.precision);
            // ln is the number of digits of the integer part of the value
            let ln = int_part.len() as i32;
            // digit count is the number of digit tokens ('0', '?' and '#') to the left of the decimal point
            let digit_count = p.digit_count;
            // digit_index points to the digit index in value that we have already formatted
            let mut digit_index = 0;

            let symbols = &locale.numbers.symbols;
            let group_sizes = locale.numbers.decimal_formats.standard.to_owned();
            let group_separator = symbols.group.to_owned();
            let decimal_separator = symbols.decimal.to_owned();
            // There probably are better ways to check if a number at a given precision is negative :/
            let is_negative = value < -(10.0_f64.powf(-(p.precision as f64)));

            for token in tokens {
                match token {
                    TextToken::Literal(c) => {
                        text = format!("{}{}", text, c);
                    }
                    TextToken::Text(t) => {
                        text = format!("{}{}", text, t);
                    }
                    TextToken::Ghost(_) => {
                        // we just leave a whitespace
                        // This is what the TEXT function does
                        text = format!("{} ", text);
                    }
                    TextToken::Spacer(_) => {
                        // we just leave a whitespace
                        // This is what the TEXT function does
                        text = format!("{} ", text);
                    }
                    TextToken::Raw => {
                        text = format!("{}{}", text, value);
                    }
                    TextToken::Period => {
                        text = format!("{}{}", text, decimal_separator);
                    }
                    TextToken::Digit(digit) => {
                        if digit.number == 'i' {
                            // 1. Integer part
                            let index = digit.index;
                            let number_index = ln - digit_count + index;
                            if index == 0 && is_negative {
                                text = format!("{}-", text);
                            }
                            if ln <= digit_count {
                                // The number of digits is less or equal than the number of digit tokens
                                // i.e. the value is 123 and the format_code is ##### (ln = 3 and digit_count = 5)
                                if !(number_index < 0 && digit.kind == '#') {
                                    let c = if number_index < 0 {
                                        if digit.kind == '0' {
                                            '0'
                                        } else {
                                            // digit.kind = '?'
                                            ' '
                                        }
                                    } else {
                                        int_part[number_index as usize]
                                    };
                                    let sep = if use_group_separator(
                                        p.use_thousands,
                                        ln - digit_index,
                                        &group_sizes,
                                    ) {
                                        &group_separator
                                    } else {
                                        ""
                                    };
                                    text = format!("{}{}{}", text, c, sep);
                                }
                                digit_index += 1;
                            } else {
                                // The number is larger than the formatting code 12345 and 0##
                                // We just hit the first formatting digit (0 in the example above) so we write as many digits as we can (123 in the example)
                                for i in digit_index..number_index + 1 {
                                    let sep = if use_group_separator(
                                        p.use_thousands,
                                        ln - i,
                                        &group_sizes,
                                    ) {
                                        &group_separator
                                    } else {
                                        ""
                                    };
                                    text = format!("{}{}{}", text, int_part[i as usize], sep);
                                }
                                digit_index = number_index + 1;
                            }
                        } else if digit.number == 'd' {
                            // 2. After the decimal point
                            let index = digit.index as usize;
                            if index < fract_part.len() {
                                text = format!("{}{}", text, fract_part[index]);
                            } else if digit.kind == '0' {
                                text = format!("{}0", text);
                            } else if digit.kind == '?' {
                                text = format!("{} ", text);
                            }
                        } else if digit.number == 'e' {
                            // 3. Exponent part
                            let index = digit.index;
                            if index == 0 {
                                if exponent_is_negative {
                                    text = format!("{}E-", text);
                                } else {
                                    text = format!("{}E+", text);
                                }
                            }
                            let number_index = l_exp - (p.exponent_digit_count - index);
                            if l_exp <= p.exponent_digit_count {
                                if !(number_index < 0 && digit.kind == '#') {
                                    let c = if number_index < 0 {
                                        if digit.kind == '?' {
                                            ' '
                                        } else {
                                            '0'
                                        }
                                    } else {
                                        exponent_part[number_index as usize]
                                    };

                                    text = format!("{}{}", text, c);
                                }
                            } else {
                                for i in 0..number_index + 1 {
                                    text = format!("{}{}", text, exponent_part[i as usize]);
                                }
                                digit_index += number_index + 1;
                            }
                        }
                    }
                    // Date tokens should not be present
                    TextToken::Day => {}
                    TextToken::DayPadded => {}
                    TextToken::DayNameShort => {}
                    TextToken::DayName => {}
                    TextToken::Month => {}
                    TextToken::MonthPadded => {}
                    TextToken::MonthNameShort => {}
                    TextToken::MonthName => {}
                    TextToken::MonthLetter => {}
                    TextToken::YearShort => {}
                    TextToken::Year => {}
                }
            }
            Formatted {
                text,
                color: p.color,
                error: None,
            }
        }
    }
}

// FIXME: Right now the locale does not include a currency
// The problem is that https://github.com/unicode-org/cldr-json does NOT include a currency
// It does include all the names for currencies in the language
// and the formatting of those currencies in that language

/// Parses a formatted number, returning the numeric value together with the format
/// "$ 123,345.678" => (123345.678, "$ #,##0.00")
/// "30.34%" => (0.3034, "0.00%")
/// 100€ => 100,
pub(crate) fn parse_formatted_number(value: &str) -> Result<(f64, Option<String>), String> {
    let currency = "$";
    let currency_format_standard = "$#,##0";
    let currency_format_standard_with_decimals = "$#,##0.00";
    let percentage_format_standard = "#,##0%";
    let percentage_format_with_decimals = "#,##0.00%";
    let format_with_thousand_separator = "#,##0";
    let format_with_thousand_separator_and_decimals = "#,##0.00";
    let value = value.trim();
    if let Some(p) = value.strip_suffix('%') {
        let (f, options) = parse_number(p.trim())?;
        // We ignore the separator
        if options.decimal_digits > 0 {
            return Ok((f / 100.0, Some(percentage_format_with_decimals.to_string())));
        }
        return Ok((f / 100.0, Some(percentage_format_standard.to_string())));
    } else if let Some(p) = value.strip_prefix(currency) {
        let (f, options) = parse_number(p.trim())?;
        if options.decimal_digits > 0 {
            return Ok((f, Some(currency_format_standard_with_decimals.to_string())));
        }
        return Ok((f, Some(currency_format_standard.to_string())));
    } else if let Some(p) = value.strip_suffix(currency) {
        let (f, options) = parse_number(p.trim())?;
        if options.decimal_digits > 0 {
            return Ok((f, Some(currency_format_standard_with_decimals.to_string())));
        }
        return Ok((f, Some(currency_format_standard.to_string())));
    }
    let (f, options) = parse_number(value)?;
    if options.has_commas {
        if options.decimal_digits > 0 {
            return Ok((
                f,
                Some(format_with_thousand_separator_and_decimals.to_string()),
            ));
        }
        return Ok((f, Some(format_with_thousand_separator.to_string())));
    }
    Ok((f, None))
}

#[allow(dead_code)]
struct NumberOptions {
    has_commas: bool,
    is_scientific: bool,
    decimal_digits: usize,
}

// tries to parse 'value' as a number.
// If it is a number it either uses commas as thousands separator or it does not
fn parse_number(value: &str) -> Result<(f64, NumberOptions), String> {
    let mut position = 0;
    let bytes = value.as_bytes();
    let len = bytes.len();
    let mut chars = String::from("");
    let decimal_separator = b'.';
    let group_separator = b',';
    let mut group_separator_index = Vec::new();
    // numbers before the decimal point
    while position < len {
        let x = bytes[position];
        if x.is_ascii_digit() {
            chars.push(x as char);
        } else if x == group_separator {
            group_separator_index.push(chars.len());
        } else {
            break;
        }
        position += 1;
    }
    // Check the group separator is in multiples of three
    for index in &group_separator_index {
        if (chars.len() - index) % 3 != 0 {
            return Err("Cannot parse number".to_string());
        }
    }
    let mut decimal_digits = 0;
    if position < len && bytes[position] == decimal_separator {
        // numbers after the decimal point
        chars.push('.');
        position += 1;
        let start_position = 0;
        while position < len {
            let x = bytes[position];
            if x.is_ascii_digit() {
                chars.push(x as char);
            } else {
                break;
            }
            position += 1;
        }
        decimal_digits = position - start_position;
    }
    let mut is_scientific = false;
    if position + 1 < len && (bytes[position] == b'e' || bytes[position] == b'E') {
        // exponential side
        is_scientific = true;
        let x = bytes[position + 1];
        if x == b'-' || x == b'+' || x.is_ascii_digit() {
            chars.push('e');
            chars.push(x as char);
            position += 2;
            while position < len {
                let x = bytes[position];
                if x.is_ascii_digit() {
                    chars.push(x as char);
                } else {
                    break;
                }
                position += 1;
            }
        }
    }
    if position != len {
        return Err("Could not parse number".to_string());
    };
    match chars.parse::<f64>() {
        Err(_) => Err("Failed to parse to double".to_string()),
        Ok(v) => Ok((
            v,
            NumberOptions {
                has_commas: !group_separator_index.is_empty(),
                is_scientific,
                decimal_digits,
            },
        )),
    }
}
