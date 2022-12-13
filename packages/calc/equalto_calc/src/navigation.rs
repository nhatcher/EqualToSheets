use crate::model::Model;

// FIXME: Some speedups could be done here once we put into use the dimension property of the spreadsheets

impl Model {
    pub fn get_navigation_right_edge(&self, sheet: i32, row: i32, column: i32) -> i32 {
        let last_column = 16384;
        if column == last_column {
            column
        } else if self.is_empty_cell(sheet, row, column) {
            // We look for the next one that is not empty
            for col in column + 1..last_column {
                if !self.is_empty_cell(sheet, row, col) {
                    return col;
                }
            }
            // Did not find it, we go to the end
            last_column
        } else {
            // Cell is not empty
            if self.is_empty_cell(sheet, row, column + 1) {
                // The next one is so we look for the next one that is not empty
                for col in column + 2..last_column {
                    if !self.is_empty_cell(sheet, row, col) {
                        return col;
                    }
                }
                // Did not find it, we go to the end
            } else {
                // We go to the last one that is not empty
                for col in column + 1..last_column {
                    if self.is_empty_cell(sheet, row, col) {
                        return col - 1;
                    }
                }
                // Did not find it, we go to the end
            }
            last_column
        }
    }

    pub fn get_navigation_left_edge(&self, sheet: i32, row: i32, column: i32) -> i32 {
        if column == 1 {
            column
        } else if self.is_empty_cell(sheet, row, column) {
            // We look for the previous one that is not empty
            for col in (1..column).rev() {
                if !self.is_empty_cell(sheet, row, col) {
                    return col;
                }
            }
            // Did not find it, we go to the beginning
            1
        } else {
            // Cell is not empty
            if self.is_empty_cell(sheet, row, column - 1) {
                // The next one is so we look for the next one that is not empty
                for col in (1..column - 1).rev() {
                    if !self.is_empty_cell(sheet, row, col) {
                        return col;
                    }
                }
                // Did not find it, we go to the beginning
                1
            } else {
                // We go to the last one that is not empty
                for col in (1..column).rev() {
                    if self.is_empty_cell(sheet, row, col) {
                        return col + 1;
                    }
                }
                // Did not find it, we go to the beginning
                1
            }
        }
    }

    pub fn get_navigation_top_edge(&self, sheet: i32, row: i32, column: i32) -> i32 {
        if row == 1 {
            row
        } else if self.is_empty_cell(sheet, row, column) {
            // We look for the next one that is not empty
            for r in (1..row).rev() {
                if !self.is_empty_cell(sheet, r, column) {
                    return r;
                }
            }
            // Did not find it, we go to the end
            1
        } else {
            // Cell is not empty
            if self.is_empty_cell(sheet, row - 1, column) {
                // The next one is so we look for the next one that is not empty
                for r in (1..row - 1).rev() {
                    if !self.is_empty_cell(sheet, r, column) {
                        return r;
                    }
                }
                // Did not find it, we go to the end
                1
            } else {
                // We go to the last one that is not empty
                for r in (1..row).rev() {
                    if self.is_empty_cell(sheet, r, column) {
                        return r + 1;
                    }
                }
                // Did not find it, we go to the end
                1
            }
        }
    }

    pub fn get_navigation_bottom_edge(&self, sheet: i32, row: i32, column: i32) -> i32 {
        let last_row = 1048576;
        if row == last_row {
            row
        } else if self.is_empty_cell(sheet, row, column) {
            // We look for the next one that is not empty
            for r in row + 1..last_row {
                if !self.is_empty_cell(sheet, r, column) {
                    return r;
                }
            }
            // Did not find it, we go to the end
            last_row
        } else {
            // Cell is not empty
            if self.is_empty_cell(sheet, row + 1, column) {
                // The next one is so we look for the next one that is not empty
                for r in row + 2..last_row {
                    if !self.is_empty_cell(sheet, r, column) {
                        return r;
                    }
                }
                // Did not find it, we go to the bottom
            } else {
                // We go to the last one that is not empty
                for r in row + 1..last_row {
                    if self.is_empty_cell(sheet, r, column) {
                        return r - 1;
                    }
                }
                // Did not find it, we go to the bottom
            }
            last_row
        }
    }

    pub fn get_navigation_home(&self, sheet: i32) -> (i32, i32) {
        let dimension = self.get_sheet_dimension(sheet);
        (dimension.0, dimension.1)
    }

    pub fn get_navigation_end(&self, sheet: i32) -> (i32, i32) {
        let dimension = self.get_sheet_dimension(sheet);
        (dimension.2, dimension.3)
    }
}
