use std::path::Path;

use equalto_calc::types::*;
use equalto_calc::{
    expressions::utils::number_to_column, model::Model, number_format::to_precision,
};

use crate::export::save_to_xlsx;
use crate::import::load_model_from_xlsx;

pub struct CompareError {
    message: String,
}

type CompareResult<T> = std::result::Result<T, CompareError>;

pub struct Diff {
    pub sheet_name: String,
    pub row: i32,
    pub column: i32,
    pub value1: String,
    pub value2: String,
    pub reason: String,
}

/// Compares two Models in the internal representation and returns a list of differences
pub fn compare(m1: Model, m2: Model) -> CompareResult<Vec<Diff>> {
    let ws1 = m1.get_worksheet_names();
    let ws2 = m2.get_worksheet_names();
    if ws1.len() != ws2.len() {
        return Err(CompareError {
            message: "Different number of sheets".to_string(),
        });
    }
    let mut diffs = Vec::new();
    let cells = m1.get_all_cells();
    for cell in cells {
        let sheet = cell.index;
        let row = cell.row;
        let column = cell.column;
        let cell1 = &m1.get_cell_at(sheet, row, column);
        let cell2 = &m2.get_cell_at(sheet, row, column);
        match (cell1, cell2) {
            (Cell::EmptyCell { .. }, Cell::EmptyCell { .. }) => {}
            (Cell::NumberCell { .. }, Cell::NumberCell { .. }) => {}
            (Cell::BooleanCell { .. }, Cell::BooleanCell { .. }) => {}
            (Cell::ErrorCell { .. }, Cell::ErrorCell { .. }) => {}
            (Cell::SharedString { .. }, Cell::SharedString { .. }) => {}
            (
                Cell::CellFormulaNumber { v: value1, .. },
                Cell::CellFormulaNumber { v: value2, .. },
            ) => {
                if (to_precision(*value1, 14) - to_precision(*value2, 14)).abs() > f64::EPSILON {
                    diffs.push(Diff {
                        sheet_name: ws1[cell.index as usize].clone(),
                        row,
                        column,
                        value1: cell1.to_json(),
                        value2: cell2.to_json(),
                        reason: "Numbers are different".to_string(),
                    });
                }
            }
            (
                Cell::CellFormulaString { v: value1, .. },
                Cell::CellFormulaString { v: value2, .. },
            ) => {
                // FIXME: We should compare the actual value, not just the index
                if value1 != value2 {
                    diffs.push(Diff {
                        sheet_name: ws1[cell.index as usize].clone(),
                        row,
                        column,
                        value1: cell1.to_json(),
                        value2: cell2.to_json(),
                        reason: "Strings are different".to_string(),
                    });
                }
            }
            (
                Cell::CellFormulaBoolean { v: value1, .. },
                Cell::CellFormulaBoolean { v: value2, .. },
            ) => {
                // FIXME: We should compare the actual value, not just the index
                if value1 != value2 {
                    diffs.push(Diff {
                        sheet_name: ws1[cell.index as usize].clone(),
                        row,
                        column,
                        value1: cell1.to_json(),
                        value2: cell2.to_json(),
                        reason: "Booleans are different".to_string(),
                    });
                }
            }
            (
                Cell::CellFormulaError { ei: index1, .. },
                Cell::CellFormulaError { ei: index2, .. },
            ) => {
                // FIXME: We should compare the actual value, not just the index
                if index1 != index2 {
                    diffs.push(Diff {
                        sheet_name: ws1[cell.index as usize].clone(),
                        row,
                        column,
                        value1: cell1.to_json(),
                        value2: cell2.to_json(),
                        reason: "Errors are different".to_string(),
                    });
                }
            }
            (_, _) => {
                diffs.push(Diff {
                    sheet_name: ws1[cell.index as usize].clone(),
                    row,
                    column,
                    value1: cell1.to_json(),
                    value2: cell2.to_json(),
                    reason: "Types are different".to_string(),
                });
            }
        }
    }
    Ok(diffs)
}

fn compare_models(m1: Model, m2: Model) -> Result<(), String> {
    match compare(m1, m2) {
        Ok(diffs) => {
            if diffs.is_empty() {
                println!("Models are equivalent!");
            } else {
                let mut message = "".to_string();
                for diff in diffs {
                    message = format!(
                        "{}\n.Diff: {}!{}{}, value1: {}, value2 {}\n {}",
                        message,
                        diff.sheet_name,
                        number_to_column(diff.column).unwrap(),
                        diff.row,
                        diff.value1,
                        diff.value2,
                        diff.reason
                    );
                }
                panic!("Models are different: {}", message);
            }
        }
        Err(r) => {
            panic!("Models are different: {}", r.message);
        }
    }
    Ok(())
}

/// Tests that file in file_path produces the same results in Excel and in EqualTo Calc.
pub fn test_file(file_path: &str) -> Result<(), String> {
    let model1 = load_model_from_xlsx(file_path, "en", "UTC").unwrap();
    let mut model2 = load_model_from_xlsx(file_path, "en", "UTC").unwrap();
    model2.evaluate();
    compare_models(model1, model2)
}

/// Tests that file in file_path can be converted to xlsx and read again
pub fn test_load_and_saving(file_path: &str, temp_dir_name: &Path) -> Result<(), String> {
    let model1 = load_model_from_xlsx(file_path, "en", "UTC").unwrap();

    let base_name = Path::new(file_path).file_name().unwrap().to_str().unwrap();

    let temp_path_buff = temp_dir_name.join(base_name);
    let temp_file_name = temp_path_buff.to_str().unwrap();
    let temp_file_path = &format!("{}.xlsx", &temp_file_name);
    // test can save
    save_to_xlsx(&model1, temp_file_name).unwrap();
    // test can open
    let mut model2 = load_model_from_xlsx(temp_file_path, "en", "UTC").unwrap();
    model2.evaluate();
    compare_models(model1, model2)
}
