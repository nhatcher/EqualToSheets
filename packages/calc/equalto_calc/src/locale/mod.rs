use serde::{Deserialize, Serialize};

use std::collections::HashMap;

#[derive(Serialize, Deserialize, Clone)]
pub struct Locale {
    pub dates: Dates,
    pub numbers: NumbersProperties,
}

#[derive(Serialize, Deserialize, Clone)]
pub struct NumbersProperties {
    #[serde(rename = "symbols-numberSystem-latn")]
    pub symbols: NumbersSymbols,
    #[serde(rename = "decimalFormats-numberSystem-latn")]
    pub decimal_formats: DecimalFormats,
}

#[derive(Serialize, Deserialize, Clone)]
pub struct Dates {
    pub day_names: Vec<String>,
    pub day_names_short: Vec<String>,
    pub months: Vec<String>,
    pub months_short: Vec<String>,
    pub months_letter: Vec<String>,
}

#[derive(Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct NumbersSymbols {
    pub decimal: String,
    pub group: String,
    pub list: String,
    pub percent_sign: String,
    pub plus_sign: String,
    pub minus_sign: String,
    pub approximately_sign: String,
    pub exponential: String,
    pub superscripting_exponent: String,
    pub per_mille: String,
    pub infinity: String,
    pub nan: String,
    pub time_separator: String,
}

#[derive(Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct DecimalFormats {
    pub standard: String,
}

lazy_static! {
    static ref LOCALES: HashMap<String, Locale> =
        serde_json::from_str(include_str!("locales.json")).expect("Failed parsing locale");
}

pub fn get_locale(_id: &str) -> Result<&Locale, String> {
    // TODO: pass the locale once we implement locales in Rust
    let locale = LOCALES.get("en").ok_or("Invalid locale")?;
    Ok(locale)
}

// TODO: Remove this function one we implement locales properly
pub fn get_locale_fix(id: &str) -> Result<&Locale, String> {
    let locale = LOCALES.get(id).ok_or("Invalid locale")?;
    Ok(locale)
}
