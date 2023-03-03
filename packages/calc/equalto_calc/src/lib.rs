#![deny(clippy::unwrap_used)]
pub mod calc_result;
pub mod cell;
pub mod expressions;
pub mod formatter;
pub mod language;
pub mod locale;
pub mod model;
pub mod new_empty;
pub mod number_format;
#[cfg(not(feature = "timezones"))]
pub mod timezones_utc;
pub mod types;
pub mod worksheet;

mod functions;

mod cast;
mod constants;
mod styles;

mod diffs;
mod implicit_intersection;

mod units;
mod utils;
mod workbook;

#[cfg(test)]
mod test;
