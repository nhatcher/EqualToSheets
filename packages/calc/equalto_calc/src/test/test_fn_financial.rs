#![allow(clippy::unwrap_used)]

use crate::{cell::CellValue, test::util::new_empty_model};

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

#[test]
fn test_fn_irr_npv_arguments() {
    let mut model = new_empty_model();
    model._set("A1", "=NPV()");
    model._set("A2", "=NPV(1,1)");

    model._set("C1", "-2"); // v0
    model._set("C2", "5"); // v1
    model._set("B1", "=IRR()");
    model._set("B3", "=IRR(1, 2, 3, 4)");
    // r such that v0 + v1/(1+r) = 0
    // r = -v1/v0 - 1
    model._set("B4", "=IRR(C1:C2)");

    model.evaluate();

    assert_eq!(model._get_text("A1"), *"#ERROR!");
    assert_eq!(model._get_text("A2"), *"$0.50");

    assert_eq!(model._get_text("B1"), *"#ERROR!");
    assert_eq!(model._get_text("B3"), *"#ERROR!");
    // r = 5/2-1 = 1.5
    assert_eq!(model._get_text("B4"), *"150%");
}

#[test]
fn test_fn_mirr() {
    let mut model = new_empty_model();
    model._set("A2", "-120000");
    model._set("A3", "39000");
    model._set("A4", "30000");
    model._set("A5", "21000");
    model._set("A6", "37000");
    model._set("A7", "46000");
    model._set("A8", "0.1");
    model._set("A9", "0.12");

    model._set("B1", "=MIRR(A2:A7, A8, A9)");
    model._set("B2", "=MIRR(A2:A5, A8, A9)");

    model.evaluate();
    assert_eq!(
        model.get_cell_value_by_ref("Sheet1!B1"),
        Ok(CellValue::Number(0.1260941303659051))
    );
    assert_eq!(model._get_text("B1"), "13%");
    assert_eq!(model._get_text("B2"), "-5%");
}

#[test]
fn test_fn_mirr_div_0() {
    // This test produces #DIV/0! in Excel (but it is incorrect)
    let mut model = new_empty_model();
    model._set("A2", "-30");
    model._set("A3", "-20");
    model._set("A4", "-10");
    model._set("A5", "5");
    model._set("A6", "5");
    model._set("A7", "5");
    model._set("A8", "-1");
    model._set("A9", "2");

    model._set("B1", "=MIRR(A2:A7, A8, A9)");

    model.evaluate();
    assert_eq!(
        model.get_cell_value_by_ref("Sheet1!B1"),
        Ok(CellValue::Number(-1.0))
    );
    assert_eq!(model._get_text("B1"), "-100%");
}

#[test]
fn test_fn_ispmt() {
    let mut model = new_empty_model();
    model._set("A1", "1"); // rate
    model._set("A2", "2"); // per
    model._set("A3", "5"); // nper
    model._set("A4", "4"); // pv

    model._set("B1", "=ISPMT(A1, A2, A3, A4)");
    model._set("B2", "=ISPMT(A1, A2, A3, A4, 1)");
    model._set("B3", "=ISPMT(A1, A2, A3)");

    model.evaluate();

    assert_eq!(model._get_text("B1"), "-2.4");
    assert_eq!(model._get_text("B2"), *"#ERROR!");
    assert_eq!(model._get_text("B3"), *"#ERROR!");
}

#[test]
fn test_fn_rri() {
    let mut model = new_empty_model();
    model._set("A1", "1"); // nper
    model._set("A2", "2"); // pv
    model._set("A3", "3"); // fv

    model._set("B1", "=RRI(A1, A2, A3)");
    model._set("B2", "=RRI(A1, A2)");
    model._set("B3", "=RRI(A1, A2, A3, 1)");

    model.evaluate();

    assert_eq!(model._get_text("B1"), "0.5");
    assert_eq!(model._get_text("B2"), *"#ERROR!");
    assert_eq!(model._get_text("B3"), *"#ERROR!");
}
