use crate::constants;
use crate::expressions::utils::{is_valid_column_number, is_valid_row};
use crate::{expressions::token::Error, types::*};

use std::collections::HashMap;

#[derive(Debug, PartialEq, Eq)]
pub struct WorksheetDimension {
    pub min_row: i32,
    pub max_row: i32,
    pub min_column: i32,
    pub max_column: i32,
}

impl Worksheet {
    pub fn get_name(&self) -> String {
        self.name.clone()
    }

    pub fn get_sheet_id(&self) -> u32 {
        self.sheet_id
    }

    pub fn set_name(&mut self, name: &str) {
        self.name = name.to_string();
    }

    pub fn cell(&self, row: i32, column: i32) -> Option<&Cell> {
        self.sheet_data.get(&row)?.get(&column)
    }

    pub(crate) fn cell_mut(&mut self, row: i32, column: i32) -> Option<&mut Cell> {
        self.sheet_data.get_mut(&row)?.get_mut(&column)
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

    pub fn set_style(&mut self, style_index: i32) -> Result<(), String> {
        self.cols = vec![Col {
            min: 1,
            max: constants::LAST_COLUMN,
            width: constants::DEFAULT_COLUMN_WIDTH / constants::COLUMN_WIDTH_FACTOR,
            custom_width: true,
            style: Some(style_index),
        }];
        Ok(())
    }

    pub fn set_column_style(&mut self, column: i32, style_index: i32) -> Result<(), String> {
        let cols = &mut self.cols;
        let col = Col {
            min: column,
            max: column,
            width: constants::DEFAULT_COLUMN_WIDTH / constants::COLUMN_WIDTH_FACTOR,
            custom_width: true,
            style: Some(style_index),
        };
        let mut index = 0;
        let mut split = false;
        for c in cols.iter_mut() {
            let min = c.min;
            let max = c.max;
            if min <= column && column <= max {
                if min == column && max == column {
                    c.style = Some(style_index);
                    return Ok(());
                } else {
                    // We need to split the result
                    split = true;
                    break;
                }
            }
            if column < min {
                // We passed, we should insert at index
                break;
            }
            index += 1;
        }
        if split {
            let min = cols[index].min;
            let max = cols[index].max;
            let pre = Col {
                min,
                max: column - 1,
                width: cols[index].width,
                custom_width: cols[index].custom_width,
                style: cols[index].style,
            };
            let post = Col {
                min: column + 1,
                max,
                width: cols[index].width,
                custom_width: cols[index].custom_width,
                style: cols[index].style,
            };
            cols.remove(index);
            if column != max {
                cols.insert(index, post);
            }
            cols.insert(index, col);
            if column != min {
                cols.insert(index, pre);
            }
        } else {
            cols.insert(index, col);
        }
        Ok(())
    }

    pub fn set_row_style(&mut self, row: i32, style_index: i32) -> Result<(), String> {
        for r in self.rows.iter_mut() {
            if r.r == row {
                r.s = style_index;
                r.custom_format = true;
                return Ok(());
            }
        }
        self.rows.push(Row {
            height: constants::DEFAULT_ROW_HEIGHT / constants::ROW_HEIGHT_FACTOR,
            r: row,
            custom_format: true,
            custom_height: true,
            s: style_index,
        });
        Ok(())
    }

    pub fn set_cell_style(&mut self, row: i32, column: i32, style_index: i32) {
        match self.cell_mut(row, column) {
            Some(cell) => {
                cell.set_style(style_index);
            }
            None => {
                self.set_cell_empty_with_style(row, column, style_index);
            }
        }

        // TODO: cleanup check if the old cell style is still in use
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
        let cell = Cell::EmptyCell { s };
        self.update_cell(row, column, cell);
    }

    pub fn set_cell_empty_with_style(&mut self, row: i32, column: i32, style: i32) {
        let cell = Cell::EmptyCell { s: style };
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

    /// Changes the height of a row.
    ///   * If the row does not a have a style we add it.
    ///   * If it has we modify the height and make sure it is applied.
    /// Fails if column index is outside allowed range.
    pub fn set_row_height(&mut self, row: i32, height: f64) -> Result<(), String> {
        if !is_valid_row(row) {
            return Err(format!("Row number '{row}' is not valid."));
        }

        let rows = &mut self.rows;
        for r in rows.iter_mut() {
            if r.r == row {
                r.height = height / constants::ROW_HEIGHT_FACTOR;
                r.custom_height = true;
                return Ok(());
            }
        }
        rows.push(Row {
            height: height / constants::ROW_HEIGHT_FACTOR,
            r: row,
            custom_format: false,
            custom_height: true,
            s: 0,
        });
        Ok(())
    }
    /// Changes the width of a column.
    ///   * If the column does not a have a width we simply add it
    ///   * If it has, it might be part of a range and we ned to split the range.
    /// Fails if column index is outside allowed range.
    pub fn set_column_width(&mut self, column: i32, width: f64) -> Result<(), String> {
        if !is_valid_column_number(column) {
            return Err(format!("Column number '{column}' is not valid."));
        }
        let cols = &mut self.cols;
        let mut col = Col {
            min: column,
            max: column,
            width: width / constants::COLUMN_WIDTH_FACTOR,
            custom_width: true,
            style: None,
        };
        let mut index = 0;
        let mut split = false;
        for c in cols.iter_mut() {
            let min = c.min;
            let max = c.max;
            if min <= column && column <= max {
                if min == column && max == column {
                    c.width = width / constants::COLUMN_WIDTH_FACTOR;
                    return Ok(());
                } else {
                    // We need to split the result
                    split = true;
                    break;
                }
            }
            if column < min {
                // We passed, we should insert at index
                break;
            }
            index += 1;
        }
        if split {
            let min = cols[index].min;
            let max = cols[index].max;
            let pre = Col {
                min,
                max: column - 1,
                width: cols[index].width,
                custom_width: cols[index].custom_width,
                style: cols[index].style,
            };
            let post = Col {
                min: column + 1,
                max,
                width: cols[index].width,
                custom_width: cols[index].custom_width,
                style: cols[index].style,
            };
            col.style = cols[index].style;
            cols.remove(index);
            if column != max {
                cols.insert(index, post);
            }
            cols.insert(index, col);
            if column != min {
                cols.insert(index, pre);
            }
        } else {
            cols.insert(index, col);
        }
        Ok(())
    }

    /// Return the width of a column in pixels
    pub fn column_width(&self, column: i32) -> Result<f64, String> {
        if !is_valid_column_number(column) {
            return Err(format!("Column number '{column}' is not valid."));
        }

        let cols = &self.cols;
        for col in cols {
            let min = col.min;
            let max = col.max;
            if column >= min && column <= max {
                if col.custom_width {
                    return Ok(col.width * constants::COLUMN_WIDTH_FACTOR);
                } else {
                    break;
                }
            }
        }
        Ok(constants::DEFAULT_COLUMN_WIDTH)
    }

    /// Returns the height of a row in pixels
    pub fn row_height(&self, row: i32) -> Result<f64, String> {
        if !is_valid_row(row) {
            return Err(format!("Row number '{row}' is not valid."));
        }

        let rows = &self.rows;
        for r in rows {
            if r.r == row {
                return Ok(r.height * constants::ROW_HEIGHT_FACTOR);
            }
        }
        Ok(constants::DEFAULT_ROW_HEIGHT)
    }

    /// Calculates dimension of the sheet. This function isn't cheap to calculate.
    pub fn dimension(&self) -> WorksheetDimension {
        // FIXME: It's probably better to just track the size as operations happen.
        if self.sheet_data.is_empty() {
            return WorksheetDimension {
                min_row: 1,
                max_row: 1,
                min_column: 1,
                max_column: 1,
            };
        }

        let mut row_range: Option<(i32, i32)> = None;
        let mut column_range: Option<(i32, i32)> = None;

        for (row_index, columns) in &self.sheet_data {
            row_range = if let Some((current_min, current_max)) = row_range {
                Some((current_min.min(*row_index), current_max.max(*row_index)))
            } else {
                Some((*row_index, *row_index))
            };

            for column_index in columns.keys() {
                column_range = if let Some((current_min, current_max)) = column_range {
                    Some((
                        current_min.min(*column_index),
                        current_max.max(*column_index),
                    ))
                } else {
                    Some((*column_index, *column_index))
                }
            }
        }

        let dimension = if let Some((min_row, max_row)) = row_range {
            if let Some((min_column, max_column)) = column_range {
                Some(WorksheetDimension {
                    min_row,
                    min_column,
                    max_row,
                    max_column,
                })
            } else {
                None
            }
        } else {
            None
        };

        dimension.unwrap_or(WorksheetDimension {
            min_row: 1,
            max_row: 1,
            min_column: 1,
            max_column: 1,
        })
    }
}
