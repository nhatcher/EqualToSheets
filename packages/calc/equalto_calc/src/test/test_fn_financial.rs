#![allow(clippy::unwrap_used)]

use crate::test::util::new_empty_model;

#[test]
fn test_fn_arguments() {
    let mut model = new_empty_model();
    model._set("A1", "=PMT()");
    model._set("A2", "=PMT(1,1)");
    model._set("A3", "=PMT(1,1,1,1,1,1)");

    model._set("B1", "=FV()");
    model._set("B2", "=FV(1,1)");
    model._set("B3", "=FV(1,1,1,1,1,1)");

    model._set("C1", "=PV()");
    model._set("C2", "=PV(1,1)");
    model._set("C3", "=PV(1,1,1,1,1,1)");

    model._set("D1", "=NPER()");
    model._set("D2", "=NPER(1,1)");
    model._set("D3", "=NPER(1,1,1,1,1,1)");

    model._set("E1", "=RATE()");
    model._set("E2", "=RATE(1,1)");
    model._set("E3", "=RATE(1,1,1,1,1,1)");

    model.evaluate();

    assert_eq!(model._get_text("A1"), *"#ERROR!");
    assert_eq!(model._get_text("A2"), *"#ERROR!");
    assert_eq!(model._get_text("A3"), *"#ERROR!");

    assert_eq!(model._get_text("B1"), *"#ERROR!");
    assert_eq!(model._get_text("B2"), *"#ERROR!");
    assert_eq!(model._get_text("B3"), *"#ERROR!");

    assert_eq!(model._get_text("C1"), *"#ERROR!");
    assert_eq!(model._get_text("C2"), *"#ERROR!");
    assert_eq!(model._get_text("C3"), *"#ERROR!");

    assert_eq!(model._get_text("D1"), *"#ERROR!");
    assert_eq!(model._get_text("D2"), *"#ERROR!");
    assert_eq!(model._get_text("D3"), *"#ERROR!");

    assert_eq!(model._get_text("E1"), *"#ERROR!");
    assert_eq!(model._get_text("E2"), *"#ERROR!");
    assert_eq!(model._get_text("E3"), *"#ERROR!");
}

#[test]
fn test_fn_impmt_ppmt_arguments() {
    let mut model = new_empty_model();
    model._set("A1", "=IPMT()");
    model._set("A2", "=IPMT(1,1,1)");
    model._set("A3", "=IPMT(1,1,1,1,1,1,1)");

    model._set("B1", "=PPMT()");
    model._set("B2", "=PPMT(1,1,1)");
    model._set("B3", "=PPMT(1,1,1,1,1,1,1)");

    model.evaluate();

    assert_eq!(model._get_text("A1"), *"#ERROR!");
    assert_eq!(model._get_text("A2"), *"#ERROR!");
    assert_eq!(model._get_text("A3"), *"#ERROR!");

    assert_eq!(model._get_text("B1"), *"#ERROR!");
    assert_eq!(model._get_text("B2"), *"#ERROR!");
    assert_eq!(model._get_text("B3"), *"#ERROR!");
}
