use std::{env, fs, io};
use uuid::Uuid;

use equalto_calc::model::{Environment, Model};
use equalto_calc::types::Workbook;
use equalto_xlsx::compare::{test_file, test_load_and_saving};
use equalto_xlsx::export::save_to_xlsx;
use equalto_xlsx::import::{load_from_excel, load_model_from_xlsx};

// This is a functional test.
// We check that the output of example.xlsx is what we expect.

#[test]
fn test_example() {
    let model = load_from_excel("tests/example.xlsx", "en", "Europe/Berlin").unwrap();
    assert_eq!(model.worksheets[0].frozen_rows, 0);
    assert_eq!(model.worksheets[0].frozen_columns, 0);
    let contents =
        fs::read_to_string("tests/example.json").expect("Something went wrong reading the file");
    let model2: Workbook = serde_json::from_str(&contents).unwrap();
    assert_eq!(model, model2);
}

#[test]
fn test_save_to_xlsx() {
    let mut model = load_model_from_xlsx("tests/example.xlsx", "en", "Europe/Berlin").unwrap();
    model.evaluate();
    let temp_file_name = "temp_file_example.xlsx";
    // test can safe
    save_to_xlsx(&model, temp_file_name).unwrap();
    // test can open
    let model = load_model_from_xlsx(temp_file_name, "en", "Europe/Berlin").unwrap();
    let metadata = &model.workbook.metadata;
    assert_eq!(metadata.application, "EqualTo Sheets");
    // FIXME: This will need to be updated once we fix versioning
    assert_eq!(metadata.app_version, "10.0000");
    // TODO: can we show it is the 'same' model?
    fs::remove_file(temp_file_name).unwrap();
}

#[test]
fn test_freeze() {
    // freeze has 3 frozen columns and 2 frozen rows
    let model = load_from_excel("tests/freeze.xlsx", "en", "Europe/Berlin").unwrap();
    assert_eq!(model.worksheets[0].frozen_rows, 2);
    assert_eq!(model.worksheets[0].frozen_columns, 3);
}

#[test]
fn test_split() {
    // We test that a workbook with split panes do not produce frozen rows and columns
    let model = load_from_excel("tests/split.xlsx", "en", "Europe/Berlin").unwrap();
    assert_eq!(model.worksheets[0].frozen_rows, 0);
    assert_eq!(model.worksheets[0].frozen_columns, 0);
}

#[test]
fn test_defined_names_casing() {
    let test_file_path = "tests/calc_tests/defined_names_for_unit_test.xlsx";
    let loaded_workbook = load_from_excel(test_file_path, "en", "UTC").unwrap();
    let mut model = Model::from_json(
        &serde_json::to_string(&loaded_workbook).unwrap(),
        Environment {
            get_milliseconds_since_epoch: || 1,
        },
    )
    .unwrap();

    let (row, column) = (2, 13); // B13
    let test_cases = [
        ("=named1", "11"),
        ("=NAMED1", "11"),
        ("=NaMeD1", "11"),
        ("=named2", "22"),
        ("=NAMED2", "22"),
        ("=NaMeD2", "22"),
        ("=named3", "33"),
        ("=NAMED3", "33"),
        ("=NaMeD3", "33"),
    ];
    for (formula, expected_value) in test_cases {
        model.set_input(0, row, column, formula.to_string(), 0);
        model.evaluate();
        assert_eq!(
            model.formatted_cell_value(0, row, column).unwrap(),
            expected_value
        );
    }
}

#[test]
fn test_xlsx() {
    let mut entries = fs::read_dir("tests/calc_tests/")
        .unwrap()
        .map(|res| res.map(|e| e.path()))
        .collect::<Result<Vec<_>, io::Error>>()
        .unwrap();
    entries.sort();
    let temp_folder = env::temp_dir();
    let path = format!("{}", Uuid::new_v4());
    let dir = temp_folder.join(path);
    fs::create_dir(&dir).unwrap();
    for file_path in entries {
        let file_name_str = file_path.file_name().unwrap().to_str().unwrap();
        let file_path_str = file_path.to_str().unwrap();
        println!("Testing file: {}", file_path_str);
        if file_name_str.ends_with(".xlsx") && !file_name_str.starts_with('~') {
            assert!(test_file(file_path_str).is_ok());
            assert!(test_load_and_saving(file_path_str, &dir).is_ok());
        } else {
            println!("skipping");
        }
    }
    fs::remove_dir_all(&dir).unwrap();
}
