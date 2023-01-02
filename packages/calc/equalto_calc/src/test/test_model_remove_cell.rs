#![allow(clippy::unwrap_used)]
use crate::test::util::new_empty_model;

#[test]
fn test_remove_cell_non_existing_sheet() {
    let mut model = new_empty_model();
    assert_eq!(
        model.remove_cell(13, 1, 1),
        Err("Invalid sheet index".to_string())
    );
}

#[test]
fn test_remove_cell_unset_cell() {
    let mut model = new_empty_model();
    assert!(model.remove_cell(0, 1, 1).is_ok());
}

#[test]
fn test_remove_cell_with_value() {
    let mut model = new_empty_model();
    model._set("A1", "hello");
    model.evaluate();

    assert_eq!(model.get_text_at(0, 1, 1), "hello");
    assert_eq!(model.is_empty_cell(0, 1, 1), Ok(false));

    model.remove_cell(0, 1, 1).unwrap();
    model.evaluate();

    assert_eq!(model.get_text_at(0, 1, 1), "");
    assert_eq!(model.is_empty_cell(0, 1, 1), Ok(true));
}

#[test]
fn test_remove_cell_referenced_elsewhere() {
    let mut model = new_empty_model();
    model._set("A1", "35");
    model._set("A2", "=2*A1");
    model.evaluate();

    assert_eq!(model.get_text_at(0, 1, 1), "35");
    assert_eq!(model.get_text_at(0, 2, 1), "70");
    assert_eq!(model.is_empty_cell(0, 1, 1), Ok(false));
    assert_eq!(model.is_empty_cell(0, 2, 1), Ok(false));

    model.remove_cell(0, 1, 1).unwrap();
    model.evaluate();

    assert_eq!(model.get_text_at(0, 1, 1), "");
    assert_eq!(model.get_text_at(0, 2, 1), "0");
    assert_eq!(model.is_empty_cell(0, 1, 1), Ok(true));
    assert_eq!(model.is_empty_cell(0, 2, 1), Ok(false));
}
