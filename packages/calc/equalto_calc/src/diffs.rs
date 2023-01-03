use crate::{
    expressions::{
        parser::{move_formula::ref_is_in_area, stringify::to_string, walk::forward_references},
        types::{Area, CellReferenceIndex, CellReferenceRC},
    },
    model::Model,
    types::Cell,
};
use serde::{Deserialize, Serialize};
use serde_json::json;
use std::collections::HashMap;

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Eq)]
#[serde(untagged, deny_unknown_fields)]
pub enum CellValue {
    Value(String),
    None,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Eq)]
#[serde(tag = "type", rename_all = "snake_case", deny_unknown_fields)]
pub enum Diff {
    #[serde(rename_all = "camelCase")]
    SetCellValue {
        sheet: u32,
        column: i32,
        row: i32,
        new_value: CellValue,
        new_style: i32,
        old_value: CellValue,
        old_style: i32,
    },
    // TODO: Rest of the diffs
}

impl Model {
    /// FIXME: This two methods only make sense in the front-end.
    /// All the undo/redo logic should be completely rewritten in Rust
    /// Returns the data of all cells in a row. Enough to reconstruct it.
    pub fn get_row_undo_data(&self, sheet: u32, row: i32) -> Result<String, String> {
        if let Some(worksheet) = self.workbook.worksheets.get(sheet as usize) {
            if let Some(row_data) = worksheet.sheet_data.get(&row) {
                match serde_json::to_string(row_data) {
                    Ok(s) => Ok(s),
                    Err(_) => Err("Cannot process data".to_string()),
                }
            } else {
                Ok(json!({}).to_string())
            }
        } else {
            Err("Invalid worksheet".to_string())
        }
    }

    /// Sets the row undo data
    pub fn set_row_undo_data(
        &mut self,
        sheet: u32,
        row: i32,
        row_data_str: &str,
    ) -> Result<(), String> {
        let row_data: HashMap<i32, Cell> = match serde_json::from_str(row_data_str) {
            Ok(row_data) => row_data,
            Err(_) => return Err("Cannot parse data".to_string()),
        };
        let sheet_data = &mut self.workbook.worksheets[sheet as usize].sheet_data;
        sheet_data.insert(row, row_data);
        Ok(())
    }

    pub fn forward_references(
        &mut self,
        source_area: &Area,
        target_sheet: u32,
        target_row: i32,
        target_column: i32,
    ) -> Result<Vec<Diff>, String> {
        let mut diff_list: Vec<Diff> = Vec::new();
        let target_area = &Area {
            sheet: target_sheet,
            row: target_row,
            column: target_column,
            width: source_area.width,
            height: source_area.height,
        };
        // Walk over every formula
        let cells = self.get_all_cells();
        for cell in cells {
            if let Some(f) = self
                .get_cell_formula_index(cell.index, cell.row, cell.column)
                .expect("Expected cell formula index")
            {
                let sheet = cell.index;
                let row = cell.row;
                let column = cell.column;

                // If cell is in the source or target area, skip
                if ref_is_in_area(sheet, row, column, source_area)
                    || ref_is_in_area(sheet, row, column, target_area)
                {
                    continue;
                }

                // Get the formula
                // Get a copy of the AST
                let node = &mut self.parsed_formulas[sheet as usize][f as usize].clone();
                let cell_reference = CellReferenceRC {
                    sheet: self.workbook.worksheets[sheet as usize].get_name(),
                    column: cell.column,
                    row: cell.row,
                };
                let context = CellReferenceIndex { sheet, column, row };
                let formula = to_string(node, &cell_reference);
                let target_sheet_name = &self.workbook.worksheets[target_sheet as usize].name;
                forward_references(
                    node,
                    &context,
                    source_area,
                    target_sheet,
                    target_sheet_name,
                    target_row,
                    target_column,
                );

                // If the string representation of the formula has changed update the cell
                let updated_formula = to_string(node, &cell_reference);
                if formula != updated_formula {
                    self.set_input_with_formula(sheet, row, column, &updated_formula);
                    // Update the diff list
                    let style = self.get_cell_style_index(sheet, row, column);
                    diff_list.push(Diff::SetCellValue {
                        sheet,
                        column,
                        row,
                        new_value: CellValue::Value(format!("={}", updated_formula)),
                        new_style: style,
                        old_value: CellValue::Value(format!("={}", formula)),
                        old_style: style,
                    });
                }
            }
        }
        Ok(diff_list)
    }
}
