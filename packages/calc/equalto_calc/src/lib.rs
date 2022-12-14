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
pub mod types;
pub mod worksheet;

mod functions;

mod actions;
mod cast;
mod diffs;
mod navigation;
mod styles;

mod graph_static;
mod implicit_intersection;

mod workbook;

#[cfg(test)]
mod test;
