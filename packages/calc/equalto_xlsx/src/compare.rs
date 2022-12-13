use crate::load_from_excel;
use equalto_calc::model::Environment;
use equalto_calc::types::*;
use equalto_calc::{
    expressions::utils::number_to_column, model::Model, number_format::to_precision,
};

// Not used
fn mock_get_milliseconds_since_epoch() -> i64 {
    1
}

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
        let sheet = cell.index as i32;
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
                if (to_precision(*value1, 15) - to_precision(*value2, 15)).abs() > f64::EPSILON {
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

/// Tests that file in file_path produces the same results in Excel than in EqualTo Calc.
pub fn test_file(file_path: &str) -> Result<(), String> {
    let env = Environment {
        get_milliseconds_since_epoch: mock_get_milliseconds_since_epoch,
    };
    let model = load_from_excel(file_path, "en", "UTC");
    let s1 = serde_json::to_string(&model).unwrap();
    let m1 = match Model::from_json(&s1, env.clone()) {
        Ok(model1) => model1,
        Err(_) => return Err("Failed loading model".to_string()),
    };
    let s2 = serde_json::to_string(&model).unwrap();

    let mut m2 = match Model::from_json(&s2, env) {
        Ok(model2) => model2,
        Err(_) => return Err("Failed loading model".to_string()),
    };
    m2.evaluate();
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
