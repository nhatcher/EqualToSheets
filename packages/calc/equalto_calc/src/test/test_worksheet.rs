#![allow(clippy::unwrap_used)]

use crate::{test::util::new_empty_model, worksheet::WorksheetDimension};

#[test]
fn test_worksheet_dimension_empty_sheet() {
    let model = new_empty_model();
    assert_eq!(
        model.workbook.worksheet(0).unwrap().dimension(),
        WorksheetDimension {
            min_row: 1,
            min_column: 1,
            max_row: 1,
            max_column: 1
        }
    );
}

#[test]
fn test_worksheet_dimension_single_cell() {
    let mut model = new_empty_model();
    model._set("W11", "1");
    assert_eq!(
        model.workbook.worksheet(0).unwrap().dimension(),
        WorksheetDimension {
            min_row: 11,
            min_column: 23,
            max_row: 11,
            max_column: 23
        }
    );
}

#[test]
fn test_worksheet_dimension_single_cell_set_empty() {
    let mut model = new_empty_model();
    model._set("W11", "1");
    model.set_cell_empty(0, 11, 23).unwrap();
    assert_eq!(
        model.workbook.worksheet(0).unwrap().dimension(),
        WorksheetDimension {
            min_row: 11,
            min_column: 23,
            max_row: 11,
            max_column: 23
        }
    );
}

#[test]
fn test_worksheet_dimension_single_cell_deleted() {
    let mut model = new_empty_model();
    model._set("W11", "1");
    model.delete_cell(0, 11, 23).unwrap();
    assert_eq!(
        model.workbook.worksheet(0).unwrap().dimension(),
        WorksheetDimension {
            min_row: 1,
            min_column: 1,
            max_row: 1,
            max_column: 1
        }
    );
}

#[test]
fn test_worksheet_dimension_multiple_cells() {
    let mut model = new_empty_model();
    model._set("W11", "1");
    model._set("E11", "1");
    model._set("AA17", "1");
    model._set("G17", "1");
    model._set("B19", "1");
    model.delete_cell(0, 11, 23).unwrap();
    assert_eq!(
        model.workbook.worksheet(0).unwrap().dimension(),
        WorksheetDimension {
            min_row: 11,
            min_column: 2,
            max_row: 19,
            max_column: 27
        }
    );
}

#[test]
fn test_worksheet_dimension_progressive() {
    let mut model = new_empty_model();
    assert_eq!(
        model.workbook.worksheet(0).unwrap().dimension(),
        WorksheetDimension {
            min_row: 1,
            min_column: 1,
            max_row: 1,
            max_column: 1
        }
    );

    model.set_input(0, 30, 50, "Hello World".to_string(), 0);
    assert_eq!(
        model.workbook.worksheet(0).unwrap().dimension(),
        WorksheetDimension {
            min_row: 30,
            min_column: 50,
            max_row: 30,
            max_column: 50
        }
    );

    model.set_input(0, 10, 15, "Hello World".to_string(), 0);
    assert_eq!(
        model.workbook.worksheet(0).unwrap().dimension(),
        WorksheetDimension {
            min_row: 10,
            min_column: 15,
            max_row: 30,
            max_column: 50
        }
    );

    model.set_input(0, 5, 25, "Hello World".to_string(), 0);
    assert_eq!(
        model.workbook.worksheet(0).unwrap().dimension(),
        WorksheetDimension {
            min_row: 5,
            min_column: 15,
            max_row: 30,
            max_column: 50
        }
    );

    model.set_input(0, 10, 250, "Hello World".to_string(), 0);
    assert_eq!(
        model.workbook.worksheet(0).unwrap().dimension(),
        WorksheetDimension {
            min_row: 5,
            min_column: 15,
            max_row: 30,
            max_column: 250
        }
    );
}
