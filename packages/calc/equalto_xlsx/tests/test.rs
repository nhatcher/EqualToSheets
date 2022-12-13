use std::{fs, io};

use equalto_calc::types::Workbook;
use equalto_xlsx::compare::test_file;
use equalto_xlsx::load_from_excel;

// This is a functional test.
// We check that the output of example.xlsx is what we expect.

#[test]
fn test_example() {
    let model = load_from_excel("tests/example.xlsx", "en", "Europe/Berlin");
    assert_eq!(model.worksheets[0].frozen_rows, 0);
    assert_eq!(model.worksheets[0].frozen_columns, 0);
    let contents =
        fs::read_to_string("tests/example.json").expect("Something went wrong reading the file");
    let model2: Workbook = serde_json::from_str(&contents).unwrap();
    assert_eq!(model, model2);
}

#[test]
fn test_freeze() {
    // freeze has 3 frozen columns and 2 frozen rows
    let model = load_from_excel("tests/freeze.xlsx", "en", "Europe/Berlin");
    assert_eq!(model.worksheets[0].frozen_rows, 2);
    assert_eq!(model.worksheets[0].frozen_columns, 3);
}

#[test]
fn test_split() {
    // We test that a workbook with split panes do not produce frozen rows and columns
    let model = load_from_excel("tests/split.xlsx", "en", "Europe/Berlin");
    assert_eq!(model.worksheets[0].frozen_rows, 0);
    assert_eq!(model.worksheets[0].frozen_columns, 0);
}

#[test]
fn test_xlsx() {
    let mut entries = fs::read_dir("tests/calc_tests/")
        .unwrap()
        .map(|res| res.map(|e| e.path()))
        .collect::<Result<Vec<_>, io::Error>>()
        .unwrap();
    entries.sort();
    for file_name in entries {
        println!("{}", file_name.to_string_lossy());
        assert!(test_file(file_name.to_str().unwrap()).is_ok());
    }
}
