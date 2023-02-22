#![allow(clippy::unwrap_used)]

use crate::{cell::CellValue, test::util::new_empty_model};

#[test]
fn test_general() {
    let mut model = new_empty_model();
    model.set_user_input(0, 1, 1, "$100.348".to_string());
    model.set_user_input(0, 1, 2, "=ISNUMBER(A1)".to_string());

    model.set_user_input(0, 2, 1, "$ 100.348".to_string());
    model.set_user_input(0, 2, 2, "=ISNUMBER(A2)".to_string());

    model.set_user_input(0, 3, 1, "100$".to_string());
    model.set_user_input(0, 3, 2, "=ISNUMBER(A3)".to_string());

    model.set_user_input(0, 10, 1, "50%".to_string());
    model.set_user_input(0, 10, 2, "=ISNUMBER(A10)".to_string());
    model.set_user_input(0, 11, 1, "55.759%".to_string());

    model.set_user_input(0, 20, 1, "50,123.549".to_string());
    model.set_user_input(0, 21, 1, "50,12.549".to_string());
    model.set_user_input(0, 22, 1, "1,234567".to_string());

    model.evaluate();

    // two decimal rounded up
    assert_eq!(model._get_text("A1"), "$100.35");
    assert_eq!(model._get_text("B1"), *"TRUE");
    assert_eq!(
        model.get_cell_value_by_ref("Sheet1!A1"),
        Ok(CellValue::Number(100.348))
    );
    // No space
    assert_eq!(model._get_text("A2"), "$100.35");
    assert_eq!(
        model.get_cell_value_by_ref("Sheet1!A2"),
        Ok(CellValue::Number(100.348))
    );
    assert_eq!(model._get_text("B2"), *"TRUE");

    // Dollar is on the left
    assert_eq!(model._get_text("A3"), "$100");
    assert_eq!(model._get_text("B3"), *"TRUE");

    assert_eq!(model._get_text("B10"), *"TRUE");
    assert_eq!(
        model.get_cell_value_by_ref("Sheet1!A10"),
        Ok(CellValue::Number(0.5))
    );

    // Two decimal places
    assert_eq!(model._get_text("A11"), "55.76%");

    // Two decimal places
    assert_eq!(model._get_text("A20"), "50,123.55");
    assert_eq!(
        model.get_cell_value_by_ref("Sheet1!A20"),
        Ok(CellValue::Number(50123.549))
    );

    // This is a string
    assert_eq!(model._get_text("A21"), "50,12.549");
    assert_eq!(
        model.get_cell_value_by_ref("Sheet1!A21"),
        Ok(CellValue::String("50,12.549".to_string()))
    );

    // Commas in all places
    assert_eq!(model._get_text("A22"), "1,234,567");
    assert_eq!(
        model.get_cell_value_by_ref("Sheet1!A22"),
        Ok(CellValue::Number(1234567.0))
    );
}

#[test]
fn test_numbers() {
    let mut model = new_empty_model();
    model.set_user_input(0, 1, 1, "1,000,000".to_string());

    model.evaluate();

    assert_eq!(
        model.get_cell_value_by_ref("Sheet1!A1"),
        Ok(CellValue::Number(1000000.0))
    );
}
