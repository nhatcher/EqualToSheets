use serde::{Deserialize, Serialize};

pub const LOCAL_TYPE: &str = "modern"; // or "full"

#[derive(Serialize, Deserialize)]
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

#[derive(Serialize, Deserialize)]
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
