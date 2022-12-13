#![allow(clippy::unwrap_used)]

use crate::{
    model::{Environment, Model},
    types::WorkbookType,
};

// 8 November 2022 12:13 Berlin time
pub fn mock_get_milliseconds_since_epoch() -> i64 {
    1667906008578
}
pub fn new_empty_model() -> Model {
    Model::new_empty(
        "model",
        "en",
        "Europe/Berlin",
        WorkbookType::Standard,
        Environment {
            get_milliseconds_since_epoch: mock_get_milliseconds_since_epoch,
        },
    )
    .unwrap()
}

impl Model {
    pub fn _set(&mut self, cell: &str, value: &str) {
        let cell_reference = if cell.contains('!') {
            self.parse_reference(cell).unwrap()
        } else {
            self.parse_reference(&format!("Sheet1!{}", cell)).unwrap()
        };
        let column = cell_reference.column;
        let row = cell_reference.row;
        self.set_input(cell_reference.sheet, row, column, value.to_string(), 0);
    }
    pub fn _get_formula(&self, cell: &str) -> String {
        let cell_reference = if cell.contains('!') {
            self.parse_reference(cell).unwrap()
        } else {
            self.parse_reference(&format!("Sheet1!{}", cell)).unwrap()
        };
        let column = cell_reference.column;
        let row = cell_reference.row;
        self.get_formula_or_value(cell_reference.sheet, row, column)
    }
    pub fn _get_text(&self, cell: &str) -> String {
        let cell_reference = if cell.contains('!') {
            self.parse_reference(cell).unwrap()
        } else {
            self.parse_reference(&format!("Sheet1!{}", cell)).unwrap()
        };
        let column = cell_reference.column;
        let row = cell_reference.row;
        self.get_text_at(cell_reference.sheet, row, column)
    }
}
