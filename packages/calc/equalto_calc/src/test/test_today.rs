#![allow(clippy::unwrap_used)]

use crate::model::{Environment, Model};
use crate::test::util::{mock_get_milliseconds_since_epoch, new_empty_model};

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
fn test_today_with_wrong_tz() {
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
