#![allow(clippy::unwrap_used)]

use crate::model::{Environment, Model};
use crate::test::util::{mock_get_milliseconds_since_epoch, new_empty_model};
use crate::types::WorkbookType;

#[test]
fn test_today_basic() {
    let mut model = new_empty_model();
    model._set("A1", "=TODAY()");
    model._set("A2", "=TEXT(A1, \"yyyy/m/d\")");
    model.evaluate();

    assert_eq!(model._get_text("A1"), *"44873");
    assert_eq!(model._get_text("A2"), *"2022/11/8");
}

#[test]
fn test_today_equalto_calculation_tz() {
    let mut model = Model::new_empty(
        "model",
        "en",
        "Europe/Berlin",
        WorkbookType::EqualToPlanCalculation,
        Environment {
            get_milliseconds_since_epoch: mock_get_milliseconds_since_epoch,
        },
    )
    .unwrap();
    model._set("A1", "=TODAY()");
    model._set("A2", "=TEXT(A1, \"yyyy/m/d\")");
    model.evaluate();

    assert_eq!(model._get_text("A1"), *"#ERROR!");
    assert_eq!(model._get_text("A2"), *"#ERROR!");
}

#[test]
fn test_today_equalto_payout_tz() {
    let mut model = Model::new_empty(
        "model",
        "en",
        "Europe/Berlin",
        WorkbookType::EqualToPayoutProfile,
        Environment {
            get_milliseconds_since_epoch: mock_get_milliseconds_since_epoch,
        },
    )
    .unwrap();
    model._set("A1", "=TODAY()");
    model._set("A2", "=TEXT(A1, \"yyyy/m/d\")");
    model.evaluate();

    assert_eq!(model._get_text("A1"), *"#ERROR!");
    assert_eq!(model._get_text("A2"), *"#ERROR!");
}

#[test]
fn test_today_equalto_analysis_tz() {
    let mut model = Model::new_empty(
        "model",
        "en",
        "Europe/Berlin",
        WorkbookType::EqualToPlanAnalysis,
        Environment {
            get_milliseconds_since_epoch: mock_get_milliseconds_since_epoch,
        },
    )
    .unwrap();
    model._set("A1", "=TODAY()");
    model._set("A2", "=TEXT(A1, \"yyyy/m/d\")");
    model.evaluate();

    assert_eq!(model._get_text("A1"), *"44873");
    assert_eq!(model._get_text("A2"), *"2022/11/8");
}

#[test]
fn test_today_with_wrong_tz() {
    let model = Model::new_empty(
        "model",
        "en",
        "Wrong Timezone",
        WorkbookType::Standard,
        Environment {
            get_milliseconds_since_epoch: mock_get_milliseconds_since_epoch,
        },
    );
    assert!(model.is_err());
}
