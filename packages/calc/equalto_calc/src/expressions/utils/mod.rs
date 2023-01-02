use super::types::*;

#[cfg(test)]
mod test;

pub const LAST_COLUMN: i32 = 16_384;
pub const LAST_ROW: i32 = 1_048_576;

pub fn column_to_number(column: &str) -> i32 {
    let bytes = column.as_bytes();
    let mut column_number = 0;
    let mut factor = 26_i32.pow((bytes.len() - 1) as u32);
    for &byte in bytes {
        let i = byte as i32;
        column_number += (i - 64) * factor;
        factor /= 26;
    }
    column_number
}

pub fn number_to_column(mut i: i32) -> Option<String> {
    if !(1..=LAST_COLUMN).contains(&i) {
        return None;
    }
    let mut column = "".to_string();
    while i > 0 {
        let r = ((i - 1) % 26) as u8;
        column.insert(0, (65 + r) as char);
        i = (i - 1) / 26;
    }
    Some(column)
}

pub fn is_valid_column(column: &str) -> bool {
    // last column XFD
    if column.len() > 3 {
        return false;
    }
    column_to_number(column) <= LAST_COLUMN
}

pub fn is_valid_row(row: i32) -> bool {
    row <= LAST_ROW
}

fn is_valid_row_str(row: &str) -> bool {
    match row.parse::<i32>() {
        Ok(r) => is_valid_row(r),
        Err(_r) => false,
    }
}

pub fn parse_reference_r1c1(r: &str) -> Option<ParsedReference> {
    let chars = r.as_bytes();
    let len = chars.len();
    let absolute_column;
    let absolute_row;
    let mut row = "".to_string();
    let mut column = "".to_string();
    if len < 4 {
        return None;
    }
    if chars[0] != b'R' {
        return None;
    }
    let mut i = 1;
    if chars[i] == b'[' {
        i += 1;
        absolute_row = false;
        if chars[i] == b'-' {
            i += 1;
            row.push('-');
        }
    } else {
        absolute_row = true;
    }
    while i < len {
        let ch = chars[i];
        if ch.is_ascii_digit() {
            row.push(ch as char);
        } else {
            break;
        }
        i += 1;
    }
    if !absolute_row {
        if i >= len || chars[i] != b']' {
            return None;
        };
        i += 1;
    }
    if i >= len || chars[i] != b'C' {
        return None;
    };
    i += 1;
    if i < len && chars[i] == b'[' {
        absolute_column = false;
        i += 1;
        if i < len && chars[i] == b'-' {
            i += 1;
            column.push('-');
        }
    } else {
        absolute_column = true;
    }
    while i < len {
        let ch = chars[i];
        if ch.is_ascii_digit() {
            column.push(ch as char);
        } else {
            break;
        }
        i += 1;
    }
    if !absolute_column {
        if i >= len || chars[i] != b']' {
            return None;
        };
        i += 1;
    }
    if i != len {
        return None;
    }
    Some(ParsedReference {
        row: row.parse::<i32>().unwrap_or(0),
        column: column.parse::<i32>().unwrap_or(0),
        absolute_column,
        absolute_row,
    })
}

pub fn parse_reference_a1(r: &str) -> Option<ParsedReference> {
    let chars = r.chars();
    let mut absolute_column = false;
    let mut absolute_row = false;
    let mut row = "".to_string();
    let mut column = "".to_string();
    let mut state = 1; // 1(colum), 2(row)

    for ch in chars {
        match ch {
            'A'..='Z' => {
                if state == 1 {
                    column.push(ch);
                } else {
                    return None;
                }
            }
            '0'..='9' => {
                if state == 1 {
                    state = 2
                }
                row.push(ch);
            }
            '$' => {
                if column == *"" {
                    absolute_column = true;
                } else if state == 1 {
                    absolute_row = true;
                    state = 2;
                } else {
                    return None;
                }
            }
            _ => {
                return None;
            }
        }
    }
    if !is_valid_column(&column) {
        return None;
    }
    if !is_valid_row_str(&row) {
        return None;
    }
    let row = match row.parse::<i32>() {
        Ok(r) => r,
        Err(_) => return None,
    };
    Some(ParsedReference {
        row,
        column: column_to_number(&column),
        absolute_column,
        absolute_row,
    })
}

pub fn is_valid_identifier(name: &str) -> bool {
    // https://support.microsoft.com/en-us/office/names-in-formulas-fc2935f9-115d-4bef-a370-3aa8bb4c91f1
    // https://github.com/MartinTrummer/excel-names/
    // NOTE: We are being much more restrictive than Excel.
    // In particular we do not support non ascii characters.
    let upper = name.to_ascii_uppercase();
    let bytes = upper.as_bytes();
    let len = bytes.len();
    if len > 255 || len == 0 {
        return false;
    }
    let first = bytes[0] as char;
    // The first character of a name must be a letter, an underscore character (_), or a backslash (\).
    if !(first.is_ascii_alphabetic() || first == '_' || first == '\\') {
        return false;
    }
    // You cannot use the uppercase and lowercase characters "C", "c", "R", or "r" as a defined name
    if len == 1 && (first == 'R' || first == 'C') {
        return false;
    }
    if upper == *"TRUE" || upper == *"FALSE" {
        return false;
    }
    if parse_reference_a1(name).is_some() {
        return false;
    }
    if parse_reference_r1c1(name).is_some() {
        return false;
    }
    let mut i = 1;
    while i < len {
        let ch = bytes[i] as char;
        match ch {
            'a'..='z' => {}
            'A'..='Z' => {}
            '0'..='9' => {}
            '_' => {}
            '.' => {}
            _ => {
                return false;
            }
        }
        i += 1;
    }

    true
}

fn name_needs_quoting(name: &str) -> bool {
    let chars = name.chars();
    // it contains any of these characters: ()'$,;-+{} or space
    for char in chars {
        if [' ', '(', ')', '\'', '$', ',', ';', '-', '+', '{', '}'].contains(&char) {
            return true;
        }
    }
    // TODO:
    // cell reference in A1 notation, e.g. B1048576 is quoted, B1048577 is not
    // cell reference in R1C1 notation, e.g. RC, RC2, R5C, R-4C, RC-8, R, C
    // integers
    false
}

/// Quotes a string sheet name if it needs to
/// NOTE: Invalid characters in a sheet name \, /, *, \[, \], :, ?
pub fn quote_name(name: &str) -> String {
    if name_needs_quoting(name) {
        return format!("'{}'", name.replace('\'', "''"));
    };
    name.to_string()
}
