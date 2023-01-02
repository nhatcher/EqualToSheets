#![allow(clippy::unwrap_used)]

use crate::calc_result::CellReference;
use crate::model::{Environment, Model};
use crate::types::Cell;

// 8 November 2022 12:13 Berlin time
pub fn mock_get_milliseconds_since_epoch() -> i64 {
    1667906008578
}
pub fn new_empty_model() -> Model {
    Model::new_empty(
        "model",
        "en",
        "Europe/Berlin",
        Environment {
            get_milliseconds_since_epoch: mock_get_milliseconds_since_epoch,
        },
    )
    .unwrap()
}

impl Model {
    fn _parse_reference(&self, cell: &str) -> CellReference {
        if cell.contains('!') {
            self.parse_reference(cell).unwrap()
        } else {
            self.parse_reference(&format!("Sheet1!{}", cell)).unwrap()
        }
    }
    pub fn _set(&mut self, cell: &str, value: &str) {
        let cell_reference = self._parse_reference(cell);
        let column = cell_reference.column;
        let row = cell_reference.row;
        self.set_input(cell_reference.sheet, row, column, value.to_string(), 0);
    }
    pub fn _get_formula(&self, cell: &str) -> String {
        let cell_reference = self._parse_reference(cell);
        let column = cell_reference.column;
        let row = cell_reference.row;
        self.get_formula_or_value(cell_reference.sheet, row, column)
    }
    pub fn _get_text(&self, cell: &str) -> String {
        let cell_reference = self._parse_reference(cell);
        let column = cell_reference.column;
        let row = cell_reference.row;
        self.get_text_at(cell_reference.sheet, row, column)
    }
    pub fn _get_cell(&self, cell: &str) -> &Cell {
        let cell_reference = self._parse_reference(cell);
        let column = cell_reference.column;
        let row = cell_reference.row;
        self.get_cell(cell_reference.sheet, row, column)
            .unwrap()
            .unwrap()
    }
}
