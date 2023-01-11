#![allow(clippy::unwrap_used)]

use crate::test::util::new_empty_model;

#[test]
fn test_circ() {
    let mut model = new_empty_model();
    model._set("A1", "=A1+1");
    assert_eq!(
        model.evaluate_with_error_check(),
        Err("Sheet1!A1 ('=A1+1'): Circular reference detected".to_string()),
    )
}

#[test]
fn test_nimpl() {
    let mut model = new_empty_model();
    model._set("B2", "={{1}}");
    assert_eq!(
        model.evaluate_with_error_check(),
        Err("Sheet1!B2 ('={{1}}'): Arrays not implemented".to_string()),
    )
}

#[test]
fn test_error() {
    let mut model = new_empty_model();
    model._set("C3", "=INVALID()");
    assert_eq!(
        model.evaluate_with_error_check(),
        Err("Sheet1!C3 ('=INVALID()'): Invalid function: INVALID".to_string()),
    )
}
