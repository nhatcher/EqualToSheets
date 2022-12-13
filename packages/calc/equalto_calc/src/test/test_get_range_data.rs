#![allow(clippy::unwrap_used)]
use crate::model::ExcelValue;
use crate::test::util::new_empty_model;

#[test]
fn test_get_range_data() {
    let mut model = new_empty_model();
    model._set("R33", "123");
    model._set("S33", "true");
    model._set("T33", "=R33+1");

    model._set("B4", "Love,");
    model._set("B5", "love will tear us apart,");
    model._set("B6", "again");
    model.evaluate();

    let data = model.get_range_data("Sheet1!R33:T33").unwrap();
    assert_eq!(
        data,
        [[
            ExcelValue::Number(123.0),
            ExcelValue::Boolean(true),
            ExcelValue::Number(124.0)
        ]]
    );

    let data = model.get_range_data("Sheet1!R33");
    assert!(data.is_err());

    let data = model.get_range_data("Sheet1!A1:A2:A3");
    assert!(data.is_err());

    // Invalid sheet
    let data = model.get_range_data("Sheet3!A1:A2");
    assert!(data.is_err());

    // you need to specify sheet name
    let data = model.get_range_data("A1:A2");
    assert!(data.is_err());

    let data = model.get_range_data("Sheet1!B4:B7").unwrap();
    assert_eq!(
        data,
        [
            [ExcelValue::String("Love,".to_string())],
            [ExcelValue::String("love will tear us apart,".to_string())],
            [ExcelValue::String("again".to_string())],
            [ExcelValue::String("".to_string())]
        ]
    );
}

#[test]
fn test_get_range_formatted_data() {
    let mut model = new_empty_model();
    model._set("R33", "123");
    model._set("S33", "true");
    model._set("T33", "=R33+1");

    model._set("B4", "Love,");
    model._set("B5", "love will tear us apart,");
    model._set("B6", "again");
    model.evaluate();

    let data = model.get_range_formatted_data("Sheet1!R33:T33").unwrap();
    assert_eq!(data, [["123", "TRUE", "124"]]);

    let data = model.get_range_formatted_data("Sheet1!R33");
    assert!(data.is_err());

    let data = model.get_range_formatted_data("Sheet1!A1:A2:A3");
    assert!(data.is_err());

    // Invalid sheet
    let data = model.get_range_formatted_data("Sheet3!A1:A2");
    assert!(data.is_err());

    // you need to specify sheet name
    let data = model.get_range_formatted_data("A1:A2");
    assert!(data.is_err());

    let data = model.get_range_formatted_data("Sheet1!B4:B7").unwrap();
    assert_eq!(
        data,
        [
            ["Love,".to_string()],
            ["love will tear us apart,".to_string()],
            ["again".to_string()],
            ["".to_string()]
        ]
    );
}
