use crate::constants;
use crate::{expressions::token::Error, types::*};

use std::collections::HashMap;

impl Worksheet {
    pub fn get_name(&self) -> String {
        self.name.clone()
    }

    pub fn get_sheet_id(&self) -> i32 {
        self.sheet_id
    }

    pub fn set_name(&mut self, name: &str) {
        self.name = name.to_string();
    }

    fn update_cell(&mut self, row: i32, column: i32, new_cell: Cell) {
        match self.sheet_data.get_mut(&row) {
            Some(column_data) => match column_data.get(&column) {
                Some(_cell) => {
                    column_data.insert(column, new_cell);
                }
                None => {
                    column_data.insert(column, new_cell);
                }
            },
            None => {
                let mut column_data = HashMap::new();
                column_data.insert(column, new_cell);
                self.sheet_data.insert(row, column_data);
            }
        }
    }

    // TODO [MVP]: Pass the cell style from the model
    // See: get_style_for_cell
    fn get_row_column_style(&self, row_index: i32, column_index: i32) -> i32 {
        let rows = &self.rows;
        for row in rows {
            if row.r == row_index {
                if row.custom_format {
                    return row.s;
                } else {
                    break;
                }
            }
        }
        let cols = &self.cols;
        for column in cols.iter() {
            let min = column.min;
            let max = column.max;
            if column_index >= min && column_index <= max {
                return column.style.unwrap_or(0);
            }
        }
        0
    }

    fn get_style(&mut self, row: i32, column: i32) -> i32 {
        match self.sheet_data.get_mut(&row) {
            Some(column_data) => match column_data.get(&column) {
                Some(cell) => cell.get_style(),
                None => self.get_row_column_style(row, column),
            },
            None => self.get_row_column_style(row, column),
        }
    }
    pub fn set_cell_with_formula(&mut self, row: i32, column: i32, index: i32, style: i32) {
        let cell = Cell::new_formula(index, style);
        self.update_cell(row, column, cell);
    }
    pub fn set_cell_with_number(&mut self, row: i32, column: i32, value: f64, style: i32) {
        let cell = Cell::new_number(value, style);
        self.update_cell(row, column, cell);
    }
    pub fn set_cell_with_string(&mut self, row: i32, column: i32, index: i32, style: i32) {
        let cell = Cell::new_string(index, style);
        self.update_cell(row, column, cell);
    }
    pub fn set_cell_with_boolean(&mut self, row: i32, column: i32, value: bool, style: i32) {
        let cell = Cell::new_boolean(value, style);
        self.update_cell(row, column, cell);
    }
    pub fn set_cell_with_error(&mut self, row: i32, column: i32, error: Error, style: i32) {
        let cell = Cell::new_error(error, style);
        self.update_cell(row, column, cell);
    }
    pub fn set_cell_empty(&mut self, row: i32, column: i32) {
        let s = self.get_style(row, column);
        let cell = Cell::EmptyCell {
            t: "empty".to_string(),
            s,
        };
        self.update_cell(row, column, cell);
    }
    pub fn set_cell_empty_with_style(&mut self, row: i32, column: i32, style: i32) {
        let cell = Cell::EmptyCell {
            t: "empty".to_string(),
            s: style,
        };
        self.update_cell(row, column, cell);
    }

    pub fn set_frozen_rows(&mut self, frozen_rows: i32) -> Result<(), String> {
        if frozen_rows < 0 {
            return Err("Frozen rows cannot be negative".to_string());
        } else if frozen_rows >= constants::LAST_ROW {
            return Err("Too many rows".to_string());
        }
        self.frozen_rows = frozen_rows;
        Ok(())
    }

    pub fn set_frozen_columns(&mut self, frozen_columns: i32) -> Result<(), String> {
        if frozen_columns < 0 {
            return Err("Frozen columns cannot be negative".to_string());
        } else if frozen_columns >= constants::LAST_COLUMN {
            return Err("Too many columns".to_string());
        }
        self.frozen_columns = frozen_columns;
        Ok(())
    }
}
