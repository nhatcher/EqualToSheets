use std::fs;

use equalto_calc::model::{Environment, Model};

use crate::error::XlsxError;
use crate::{export::save_to_xlsx, import::load_model_from_xlsx};

// 8 November 2022 12:13 Berlin time
pub fn mock_get_milliseconds_since_epoch() -> i64 {
    1667906008578
}

pub fn new_empty_model() -> Model {
    Model::new_empty(
        "model",
        "en",
        "Europe/Berlin",
        Environment {
            get_milliseconds_since_epoch: mock_get_milliseconds_since_epoch,
        },
    )
    .unwrap()
}

#[test]
fn test_values() {
    let mut model = new_empty_model();
    // numbers
    model.set_input(0, 1, 1, "123.456".to_string(), 0);
    // strings
    model.set_input(0, 2, 1, "Hello world!".to_string(), 0);
    model.set_input(0, 3, 1, "Hello world!".to_string(), 0);
    model.set_input(0, 4, 1, "你好世界！".to_string(), 0);
    // booleans
    model.set_input(0, 5, 1, "TRUE".to_string(), 0);
    model.set_input(0, 6, 1, "FALSE".to_string(), 0);
    // errors
    model.set_input(0, 7, 1, "#VALUE!".to_string(), 0);

    // noop
    model.evaluate();

    let temp_file_name = "temp_file_test_values.xlsx";
    save_to_xlsx(&model, temp_file_name).unwrap();

    let model = load_model_from_xlsx(temp_file_name, "en", "Europe/Berlin").unwrap();
    assert_eq!(model.get_text_at(0, 1, 1), "123.456");
    assert_eq!(model.get_text_at(0, 2, 1), "Hello world!");
    assert_eq!(model.get_text_at(0, 3, 1), "Hello world!");
    assert_eq!(model.get_text_at(0, 4, 1), "你好世界！");
    assert_eq!(model.get_text_at(0, 5, 1), "TRUE");
    assert_eq!(model.get_text_at(0, 6, 1), "FALSE");
    assert_eq!(model.get_text_at(0, 7, 1), "#VALUE!");

    fs::remove_file(temp_file_name).unwrap();
}

#[test]
fn test_formulas() {
    let mut model = new_empty_model();
    model.set_input(0, 1, 1, "5.5".to_string(), 0);
    model.set_input(0, 2, 1, "6.5".to_string(), 0);
    model.set_input(0, 3, 1, "7.5".to_string(), 0);

    model.set_input(0, 1, 2, "=A1*2".to_string(), 0);
    model.set_input(0, 2, 2, "=A2*2".to_string(), 0);
    model.set_input(0, 3, 2, "=A3*2".to_string(), 0);
    model.set_input(0, 4, 2, "=SUM(A1:B3)".to_string(), 0);

    model.evaluate();
    let temp_file_name = "temp_file_test_formulas.xlsx";
    save_to_xlsx(&model, temp_file_name).unwrap();

    let model = load_model_from_xlsx(temp_file_name, "en", "Europe/Berlin").unwrap();
    assert_eq!(model.get_text_at(0, 1, 2), "11");
    assert_eq!(model.get_text_at(0, 2, 2), "13");
    assert_eq!(model.get_text_at(0, 3, 2), "15");
    assert_eq!(model.get_text_at(0, 4, 2), "58.5");
    fs::remove_file(temp_file_name).unwrap();
}

#[test]
fn test_sheets() {
    let mut model = new_empty_model();
    model.add_sheet("With space").unwrap();
    // xml escaped
    model.add_sheet("Tango & Cash").unwrap();
    model.add_sheet("你好世界").unwrap();

    // noop
    model.evaluate();

    let temp_file_name = "temp_file_test_sheets.xlsx";
    save_to_xlsx(&model, temp_file_name).unwrap();

    let model = load_model_from_xlsx(temp_file_name, "en", "Europe/Berlin").unwrap();
    assert_eq!(
        model.workbook.get_worksheet_names(),
        vec!["Sheet1", "With space", "Tango & Cash", "你好世界"]
    );
    fs::remove_file(temp_file_name).unwrap();
}

#[test]
fn test_named_styles() {
    let mut model = new_empty_model();
    model.set_input(0, 1, 1, "5.5".to_string(), 0);
    let mut style = model.get_style_for_cell(0, 1, 1);
    style.font.b = true;
    style.font.i = true;
    assert!(model.set_cell_style(0, 1, 1, &style).is_ok());
    let bold_style_index = model.get_cell_style_index(0, 1, 1);
    let e = model
        .workbook
        .styles
        .add_named_cell_style("bold & italics", bold_style_index);
    assert!(e.is_ok());

    // noop
    model.evaluate();

    let temp_file_name = "temp_file_test_named_styles.xlsx";
    save_to_xlsx(&model, temp_file_name).unwrap();

    let model = load_model_from_xlsx(temp_file_name, "en", "Europe/Berlin").unwrap();
    assert!(model
        .workbook
        .styles
        .get_style_index_by_name("bold & italics")
        .is_ok());
    fs::remove_file(temp_file_name).unwrap();
}

#[test]
fn test_existing_file() {
    let file_name = "existing_file.xlsx";
    fs::File::create(file_name).unwrap();

    assert_eq!(
        save_to_xlsx(&new_empty_model(), file_name),
        Err(XlsxError::IO(
            "file existing_file.xlsx already exists".to_string()
        )),
    );

    fs::remove_file(file_name).unwrap();
}
