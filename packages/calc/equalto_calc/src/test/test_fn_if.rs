#![allow(clippy::unwrap_used)]

use crate::test::util::new_empty_model;

#[test]
fn test_fn_if_arguments() {
    let mut model = new_empty_model();
    model._set("A1", "=IF()");
    model._set("A2", "=IF(1, 2, 3, 4)");
    model.evaluate();

    assert_eq!(model._get_text("A1"), *"#ERROR!");
    assert_eq!(model._get_text("A2"), *"#ERROR!");
}

#[test]
fn test_fn_if_2_args() {
    let mut model = new_empty_model();
    model._set("A1", "=IF(2 > 3, TRUE)");
    model.evaluate();
    assert_eq!(model._get_text("A1"), *"FALSE");
}
