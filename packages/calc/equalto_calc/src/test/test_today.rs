#![allow(clippy::unwrap_used)]

use crate::model::{Environment, Model};
use crate::test::util::{mock_get_milliseconds_since_epoch, new_empty_model};

pub fn mock_get_milliseconds_since_epoch_2023() -> i64 {
    // 14:44 20 Mar 2023 Berlin
    1679319865208
}

#[test]
fn today_basic() {
    let mut model = new_empty_model();
    model._set("A1", "=TODAY()");
    model._set("A2", "=TEXT(A1, \"yyyy/m/d\")");
    model.evaluate();

    assert_eq!(model._get_text("A1"), *"08/11/2022");
    assert_eq!(model._get_text("A2"), *"2022/11/8");
}

#[test]
fn today_with_wrong_tz() {
    let model = Model::new_empty(
        "model",
        "en",
        "Wrong Timezone",
        Environment {
            get_milliseconds_since_epoch: mock_get_milliseconds_since_epoch,
        },
    );
    assert!(model.is_err());
}

#[test]
fn now_basic_utc() {
    let mut model = Model::new_empty(
        "model",
        "en",
        "UTC",
        Environment {
            get_milliseconds_since_epoch: mock_get_milliseconds_since_epoch_2023,
        },
    )
    .unwrap();
    model._set("A1", "=TODAY()");
    model._set("A2", "=NOW()");
    model.evaluate();

    assert_eq!(model._get_text("A1"), *"20/03/2023");
    assert_eq!(model._get_text("A2"), *"45005.572511574");
}

#[test]
fn now_basic_europe_berlin() {
    let mut model = Model::new_empty(
        "model",
        "en",
        "Europe/Berlin",
        Environment {
            get_milliseconds_since_epoch: mock_get_milliseconds_since_epoch_2023,
        },
    )
    .unwrap();
    model._set("A1", "=TODAY()");
    model._set("A2", "=NOW()");
    model.evaluate();

    assert_eq!(model._get_text("A1"), *"20/03/2023");
    // This is UTC + 1 hour: 45005.572511574 + 1/24
    assert_eq!(model._get_text("A2"), *"45005.614178241");
}
