use crate::expressions::parser::stringify::to_string;
use crate::expressions::parser::walk::swap_references;
use crate::expressions::types::{CellReferenceIndex, CellReferenceRC};
use crate::expressions::{parser::stringify::DisplaceData, utils::LAST_ROW};
use crate::{expressions::utils::LAST_COLUMN, model::Model};

// NOTE: There is a difference with Excel behaviour when deleting cells/rows/columns
// In Excel if the whole range is deleted then it will substitute for #REF!
// In EqualTo, if one of the edges of the range is deleted will replace the edge with #REF!
// I feel this is unimportant for now.

impl Model {
    fn displace_cells(&mut self, displace_data: &DisplaceData) {
        let cells = self.get_all_cells();
        for cell in cells {
            self.shift_cell_formula(cell.index, cell.row, cell.column, displace_data);
        }
    }
    /// Returns the list of columns in row
    fn get_columns_for_row(&self, sheet: u32, row: i32, descending: bool) -> Vec<i32> {
        let worksheet = &self.workbook.worksheets[sheet as usize];
        if let Some(row_data) = worksheet.sheet_data.get(&row) {
            let mut columns: Vec<i32> = row_data.keys().copied().collect();
            columns.sort_unstable();
            if descending {
                columns.reverse();
            }
            columns
        } else {
            vec![]
        }
    }

    /// Returns the list of row in a column
    fn get_rows_for_column(&self, sheet: u32, column: i32, descending: bool) -> Vec<i32> {
        let worksheet = &self.workbook.worksheets[sheet as usize];
        let mut rows: Vec<i32> = vec![];
        let all_rows: Vec<_> = worksheet.sheet_data.keys().collect();
        for row in all_rows {
            if worksheet
                .sheet_data
                .get(row)
                .expect("expected row")
                .contains_key(&column)
            {
                rows.push(*row);
            }
        }
        rows.sort_unstable();
        if descending {
            rows.reverse();
        }
        rows
    }

    /// Moves the contents of cell (source_row, source_column) tp (target_row, target_column)
    /// Assumes that cell exists, otherwise will panic.
    fn move_cell(
        &mut self,
        sheet: u32,
        source_row: i32,
        source_column: i32,
        target_row: i32,
        target_column: i32,
    ) {
        let value = self.get_formula_or_value(sheet, source_row, source_column);
        let style = self.get_cell_style_index(sheet, source_row, source_column);
        self.set_input(sheet, target_row, target_column, value, style);
        self.remove_cell(sheet, source_row, source_column)
            .expect("Expected cell to exist");
    }

    pub fn insert_columns(
        &mut self,
        sheet: u32,
        column: i32,
        column_count: i32,
    ) -> Result<(), &'static str> {
        if column_count <= 0 {
            return Err("Cannot add a negative number of cells :)");
        }
        // check if it is possible:
        let dimensions = self.get_sheet_dimension(sheet);
        let last_column = dimensions.3 + column_count;
        if last_column > LAST_COLUMN {
            return Err("Cannot shift cells because that would delete cells at the end of a row");
        }
        let worksheet = &self.workbook.worksheets[sheet as usize];
        let all_rows: Vec<i32> = worksheet.sheet_data.keys().copied().collect();
        for row in all_rows {
            let sorted_columns = self.get_columns_for_row(sheet, row, true);
            for col in sorted_columns {
                if col >= column {
                    self.move_cell(sheet, row, col, row, col + column_count);
                } else {
                    // They are in descending order
                    break;
                }
            }
        }

        // Update all formulas in the workbook
        self.displace_cells(&DisplaceData::Column {
            sheet,
            column,
            delta: column_count,
        });

        Ok(())
    }

    pub fn delete_columns(
        &mut self,
        sheet: u32,
        column: i32,
        column_count: i32,
    ) -> Result<(), &'static str> {
        if column_count <= 0 {
            return Err("Please use insert columns instead");
        }

        // Move cells
        let worksheet = &self.workbook.worksheets[sheet as usize];
        let mut all_rows: Vec<i32> = worksheet.sheet_data.keys().copied().collect();
        // We do not need to do that, but it is safer to eliminate sources of randomness in the algorithm
        all_rows.sort_unstable();

        for r in all_rows {
            let columns: Vec<i32> = self.get_columns_for_row(sheet, r, false);
            for col in columns {
                if col >= column {
                    if col >= column + column_count {
                        self.move_cell(sheet, r, col, r, col - column_count);
                    } else {
                        self.remove_cell(sheet, r, col)
                            .expect("Expected cell to exist");
                    }
                }
            }
        }
        // Update all formulas in the workbook

        self.displace_cells(&DisplaceData::Column {
            sheet,
            column,
            delta: -column_count,
        });

        Ok(())
    }

    pub fn insert_rows(
        &mut self,
        sheet: u32,
        row: i32,
        row_count: i32,
    ) -> Result<(), &'static str> {
        if row_count <= 0 {
            return Err("Cannot add a negative number of cells :)");
        }
        // Check if it is possible:
        let dimensions = self.get_sheet_dimension(sheet);
        let last_row = dimensions.2 + row_count;
        if last_row > LAST_ROW {
            return Err(
                "Cannot shift cells because that would delete cells at the end of a column",
            );
        }

        // Move cells
        let worksheet = &self.workbook.worksheets[sheet as usize];
        let mut all_rows: Vec<i32> = worksheet.sheet_data.keys().copied().collect();
        all_rows.sort_unstable();
        all_rows.reverse();
        for r in all_rows {
            if r >= row {
                // We do not really need the columns in any order
                let columns: Vec<i32> = self.get_columns_for_row(sheet, r, false);
                for column in columns {
                    self.move_cell(sheet, r, column, r + row_count, column);
                }
            } else {
                // Rows are in descending order
                break;
            }
        }
        // In the list of rows styles:
        // * Add all rows above the rows we are inserting unchanged
        // * Shift the ones below
        let rows = &self.workbook.worksheets[sheet as usize].rows;
        let mut new_rows = vec![];
        for r in rows {
            if r.r < row {
                new_rows.push(r.clone());
            } else if r.r >= row {
                let mut new_row = r.clone();
                new_row.r = r.r + row_count;
                new_rows.push(new_row);
            }
        }
        self.workbook.worksheets[sheet as usize].rows = new_rows;

        // Update all formulas in the workbook
        self.displace_cells(&DisplaceData::Row {
            sheet,
            row,
            delta: row_count,
        });

        Ok(())
    }

    pub fn delete_rows(
        &mut self,
        sheet: u32,
        row: i32,
        row_count: i32,
    ) -> Result<(), &'static str> {
        if row_count <= 0 {
            return Err("Please use insert rows instead");
        }
        // Move cells
        let worksheet = &self.workbook.worksheets[sheet as usize];
        let mut all_rows: Vec<i32> = worksheet.sheet_data.keys().copied().collect();
        all_rows.sort_unstable();

        for r in all_rows {
            if r >= row {
                // We do not need ordered, but it is safer to eliminate sources of randomness in the algorithm
                let columns: Vec<i32> = self.get_columns_for_row(sheet, r, false);
                if r >= row + row_count {
                    // displace all cells in column
                    for column in columns {
                        self.move_cell(sheet, r, column, r - row_count, column);
                    }
                } else {
                    // remove all cells in row
                    // FIXME: We could just remove the entire row in one go
                    for column in columns {
                        self.remove_cell(sheet, r, column)
                            .expect("Expected cell to exist");
                    }
                }
            }
        }
        // In the list of rows styles:
        // * Add all rows above the rows we are deleting unchanged
        // * Skip all those we are deleting
        // * Shift the ones below
        let rows = &self.workbook.worksheets[sheet as usize].rows;
        let mut new_rows = vec![];
        for r in rows {
            if r.r < row {
                new_rows.push(r.clone());
            } else if r.r >= row + row_count {
                let mut new_row = r.clone();
                new_row.r = r.r - row_count;
                new_rows.push(new_row);
            }
        }
        self.workbook.worksheets[sheet as usize].rows = new_rows;
        self.displace_cells(&DisplaceData::Row {
            sheet,
            row,
            delta: -row_count,
        });
        Ok(())
    }

    // Adds cell_count cells to the right of cell (row, column)
    // Returns an error if there is content in some of the last cell_count columns.
    pub fn shift_cells_right(
        &mut self,
        sheet: u32,
        row: i32,
        column: i32,
        cell_count: i32,
    ) -> Result<(), &'static str> {
        if cell_count <= 0 {
            return Err("Cannot add a negative number of cells :)");
        }
        // Check if it is possible
        let sorted_columns = self.get_columns_for_row(sheet, row, true);
        if sorted_columns.is_empty() {
            // Nothing to do.
            return Ok(());
        }
        let last_column = sorted_columns[0] + cell_count;
        if last_column > LAST_COLUMN {
            return Err("Cannot shift cells because that would delete cells at the end of the row");
        }

        // Move cells
        for col in sorted_columns {
            if col >= column {
                self.move_cell(sheet, row, col, row, col + cell_count);
            } else {
                // They are in descending order
                break;
            }
        }

        // Update all formulas in the workbook
        self.displace_cells(&DisplaceData::CellHorizontal {
            sheet,
            row,
            column,
            delta: cell_count,
        });

        Ok(())
    }

    pub fn shift_cells_down(
        &mut self,
        sheet: u32,
        row: i32,
        column: i32,
        cell_count: i32,
    ) -> Result<(), &'static str> {
        if cell_count <= 0 {
            return Err("Cannot add a negative number of cells :)");
        }
        // Check if it is possible
        let sorted_rows = self.get_rows_for_column(sheet, column, true);
        if sorted_rows.is_empty() {
            // Nothing to do.
            return Ok(());
        }
        let last_row = sorted_rows[0] + cell_count;
        if last_row > LAST_ROW {
            return Err(
                "Cannot shift cells because that would delete cells at the end of the column",
            );
        }

        // Move cells
        for r in sorted_rows {
            if r >= row {
                self.move_cell(sheet, r, column, r + cell_count, column);
            } else {
                // They are in descending order
                break;
            }
        }

        // Update all formulas in the workbook
        self.displace_cells(&DisplaceData::CellVertical {
            sheet,
            row,
            column,
            delta: cell_count,
        });
        Ok(())
    }

    // delete cells
    pub fn shift_cells_up(
        &mut self,
        sheet: u32,
        row: i32,
        column: i32,
        cell_count: i32,
    ) -> Result<(), &'static str> {
        if cell_count <= 0 {
            return Err("Cannot shift a negative number of cells, please use 'shift cells down' for that purpose");
        }

        let sorted_rows = self.get_rows_for_column(sheet, column, false);
        if sorted_rows.is_empty() {
            // Nothing to do.
            return Ok(());
        }

        // Move cells
        for r in sorted_rows {
            if r >= row {
                if r < row + cell_count {
                    // delete the cell
                    self.remove_cell(sheet, r, column)
                        .expect("Expected cell to exist");
                } else {
                    self.move_cell(sheet, r, column, r - cell_count, column);
                }
            }
        }

        // Update all formulas in the workbook
        self.displace_cells(&DisplaceData::CellVertical {
            sheet,
            row,
            column,
            delta: -cell_count,
        });

        Ok(())
    }

    // delete cells
    pub fn shift_cells_left(
        &mut self,
        sheet: u32,
        row: i32,
        column: i32,
        cell_count: i32,
    ) -> Result<(), &'static str> {
        if cell_count <= 0 {
            return Err("Cannot shift a negative number of cells, please use 'shift cells right' for that purpose");
        }

        let sorted_columns = self.get_columns_for_row(sheet, row, false);
        if sorted_columns.is_empty() {
            // Nothing to do
            return Ok(());
        }

        // Move cells
        for col in sorted_columns {
            if col >= column {
                if col < column + cell_count {
                    self.remove_cell(sheet, row, col)
                        .expect("Expected cell to exist");
                } else {
                    self.move_cell(sheet, row, col, row, col - cell_count);
                }
            }
        }

        // Update all formulas in the workbook
        self.displace_cells(&DisplaceData::CellHorizontal {
            sheet,
            row,
            column,
            delta: -cell_count,
        });

        Ok(())
    }

    pub fn swap_cells_in_row(
        &mut self,
        sheet: u32,
        row: i32,
        column1: i32,
        column2: i32,
    ) -> Result<(), &'static str> {
        // First swap styles and values
        let value1 = self.get_formula_or_value(sheet, row, column1);
        let style_index1 = self.get_cell_style_index(sheet, row, column1);
        let value2 = self.get_formula_or_value(sheet, row, column2);
        let style_index2 = self.get_cell_style_index(sheet, row, column2);
        self.set_input(sheet, row, column1, value2, style_index2);
        self.set_input(sheet, row, column2, value1, style_index1);

        // Now walk over every formula swapping the references
        let cells = self.get_all_cells();
        for cell in cells {
            if let Some(f) = self
                .get_cell_formula_index(cell.index, cell.row, cell.column)
                .expect("Expected cell formula index")
            {
                // Get a copy of the AST
                let node = &mut self.parsed_formulas[sheet as usize][f as usize].clone();
                let cell_reference = CellReferenceRC {
                    sheet: self.workbook.worksheets[sheet as usize].get_name(),
                    column: cell.column,
                    row: cell.row,
                };
                let formula = to_string(node, &cell_reference);

                // update the AST
                swap_references(
                    node,
                    &CellReferenceIndex {
                        sheet: cell.index,
                        column: cell.column,
                        row: cell.row,
                    },
                    sheet,
                    row,
                    column1,
                    column2,
                );

                // If the string representation of the formula has changed update the cell
                let updated_formula = to_string(node, &cell_reference);
                if formula != updated_formula {
                    self.set_input_with_formula(
                        cell.index,
                        cell.row,
                        cell.column,
                        &updated_formula,
                    );
                }
            }
        }

        // Now every formula pointing a cell1 needs to point at cell2 and all the other way around
        Ok(())
    }

    /// Displaces cells due to a move column action
    /// from initial_column to target_column = initial_column + column_delta
    /// References will be updated following:
    /// Cell references:
    ///    * All cell references to initial_column will go to target_column
    ///    * All cell references to columns in between (initial_column, target_column] will be displaced one to the left
    ///    * All other cell references are left unchanged
    /// Ranges. This is the tricky bit:
    ///    * Column is one of the extremes of the range. The new extreme would be target_column.
    ///      Range is then normalized
    ///    * Any other case, range is left unchanged.
    /// NOTE: This does NOT move the data in the columns or move the colum styles
    pub fn move_column_action(
        &mut self,
        sheet: u32,
        column: i32,
        delta: i32,
    ) -> Result<(), &'static str> {
        // Check boundaries
        let target_column = column + delta;
        if !(1..=LAST_COLUMN).contains(&target_column) {
            return Err("Target column out of boundaries");
        }
        if !(1..=LAST_COLUMN).contains(&column) {
            return Err("Initial column out of boundaries");
        }

        // TODO: Add the actual displacement of data and styles

        // Update all formulas in the workbook
        self.displace_cells(&DisplaceData::ColumnMove {
            sheet,
            column,
            delta,
        });

        Ok(())
    }
}
