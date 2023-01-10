use crate::expressions::token::get_error_by_name;
use crate::language::Language;

/// Returns true if the string value could be interpreted as:
///  * a formula
///  * a number
///  * a boolean
///  * an error (i.e "#VALUE!")
pub(crate) fn value_needs_quoting(value: &str, language: &Language) -> bool {
    value.starts_with('=')
        || value.parse::<f64>().is_ok()
        || value.to_lowercase().parse::<bool>().is_ok()
        || get_error_by_name(&value.to_uppercase(), language).is_some()
}

/// Valid hex colors are #FFAABB
/// #fff is not valid
pub(crate) fn is_valid_hex_color(color: &str) -> bool {
    if color.chars().count() != 7 {
        return false;
    }
    if !color.starts_with('#') {
        return false;
    }
    if let Ok(z) = i32::from_str_radix(&color[1..], 16) {
        if (0..=0xffffff).contains(&z) {
            return true;
        }
    }
    false
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::language::get_language;

    #[test]
    fn test_value_needs_quoting() {
        let en_language = get_language("en").expect("en language expected");

        assert!(!value_needs_quoting("", en_language));
        assert!(!value_needs_quoting("hello", en_language));

        assert!(value_needs_quoting("12", en_language));
        assert!(value_needs_quoting("true", en_language));
        assert!(value_needs_quoting("False", en_language));

        assert!(value_needs_quoting("=A1", en_language));

        assert!(value_needs_quoting("#REF!", en_language));
        assert!(value_needs_quoting("#NAME?", en_language));
    }

    #[test]
    fn test_is_valid_hex_color() {
        assert!(is_valid_hex_color("#000000"));
        assert!(is_valid_hex_color("#ffffff"));

        assert!(!is_valid_hex_color("000000"));
        assert!(!is_valid_hex_color("ffffff"));

        assert!(!is_valid_hex_color("#gggggg"));

        // Not obvious cases unrecognized as colors
        assert!(!is_valid_hex_color("#ffffff "));
        assert!(!is_valid_hex_color("#fff")); // CSS shorthand
        assert!(!is_valid_hex_color("#ffffff00")); // with alpha channel
    }
}
