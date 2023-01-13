#![allow(clippy::unwrap_used)]

use crate::model::ExcelValue::{self};

use crate::{number_format::to_excel_precision_str, types::Color};

use crate::test::util::new_empty_model;

#[test]
fn test_empty_model() {
    let model = new_empty_model();
    let names = model.get_worksheet_names();
    assert_eq!(names.len(), 1);
    assert_eq!(names[0], "Sheet1");
}

#[test]
fn test_model_simple_evaluation() {
    let mut model = new_empty_model();
    model.set_input(0, 1, 1, "= 1 + 3".to_string(), 0);
    model.evaluate();
    let result = model.get_text_at(0, 1, 1);
    assert_eq!(result, *"4");
    let result = model.get_formula_or_value(0, 1, 1);
    assert_eq!(result, *"=1+3");
}

#[test]
fn test_model_simple_evaluation_order() {
    let mut model = new_empty_model();
    model._set("A1", "=1/2/3");
    model._set("A2", "=(1/2)/3");
    model._set("A3", "=1/(2/3)");
    model.evaluate();
    assert_eq!(model._get_text("A1"), *"0.166666666666667");
    assert_eq!(model._get_text("A2"), *"0.166666666666667");
    assert_eq!(model._get_text("A3"), *"1.5");
    // Unnecessary parenthesis are lost
    assert_eq!(model._get_formula("A2"), *"=1/2/3");
    assert_eq!(model._get_formula("A3"), *"=1/(2/3)");
}

#[test]
fn test_model_invalid_formula() {
    let mut model = new_empty_model();
    model.set_input(0, 1, 1, "= 1 +".to_string(), 0);
    model.evaluate();
    let result = model.get_text_at(0, 1, 1);
    assert_eq!(result, *"#ERROR!");
    let result = model.get_formula_or_value(0, 1, 1);
    assert_eq!(result, *"= 1 +");
}

#[test]
fn test_model_dependencies() {
    let mut model = new_empty_model();
    model.set_input(0, 1, 1, "23".to_string(), 0); // A1
    model.set_input(0, 1, 2, "= A1* 2-4".to_string(), 0); // B1
    model.evaluate();
    let result = model.get_text_at(0, 1, 1);
    assert_eq!(result, *"23");
    let result = model.get_formula_or_value(0, 1, 1);
    assert_eq!(result, *"23");
    let result = model.get_text_at(0, 1, 2);
    assert_eq!(result, *"42");
    let result = model.get_formula_or_value(0, 1, 2);
    assert_eq!(result, *"=A1*2-4");

    model.set_input(0, 2, 1, "=SUM(A1, B1)".to_string(), 0); // A2
    model.evaluate();
    let result = model.get_text_at(0, 2, 1);
    assert_eq!(result, *"65");
}

#[test]
fn test_model_strings() {
    let mut model = new_empty_model();
    model.set_input(0, 1, 1, "Hello World".to_string(), 0);
    model.set_input(0, 1, 2, "=A1".to_string(), 0);
    model.evaluate();
    let result = model.get_text_at(0, 1, 1);
    assert_eq!(result, *"Hello World");
    let result = model.get_text_at(0, 1, 2);
    assert_eq!(result, *"Hello World");
}

#[test]
fn test_get_sheet_index_by_sheet_id() {
    let mut model = new_empty_model();
    model.new_sheet();

    assert_eq!(model.get_sheet_index_by_sheet_id(1), Some(0));
    assert_eq!(model.get_sheet_index_by_sheet_id(2), Some(1));
    assert_eq!(model.get_sheet_index_by_sheet_id(1337), None);
}

#[test]
fn test_set_row_height() {
    let mut model = new_empty_model();
    model.set_row_height(0, 5, 25.0);
    assert!((25.0 - model.get_row_height(0, 5)).abs() < f64::EPSILON);

    model.set_row_height(0, 5, 5.0);
    assert!((5.0 - model.get_row_height(0, 5)).abs() < f64::EPSILON);
}

#[test]
fn test_to_excel_precision_str() {
    struct TestCase<'a> {
        value: f64,
        str: &'a str,
    }
    let test_cases = vec![
        TestCase {
            value: 2e-23,
            str: "2e-23",
        },
        TestCase {
            value: 42.0,
            str: "42",
        },
        TestCase {
            value: 200.0e-23,
            str: "2e-21",
        },
        TestCase {
            value: -200e-23,
            str: "-2e-21",
        },
        TestCase {
            value: 10.002,
            str: "10.002",
        },
        TestCase {
            value: f64::INFINITY,
            str: "inf",
        },
        TestCase {
            value: f64::NAN,
            str: "NaN",
        },
    ];
    for test_case in test_cases {
        let str = to_excel_precision_str(test_case.value);
        assert_eq!(str, test_case.str);
    }
}

#[test]
fn test_booleans() {
    let mut model = new_empty_model();
    model.set_input(0, 1, 1, "true".to_string(), 0);
    model.set_input(0, 2, 1, "TRUE".to_string(), 0);
    model.set_input(0, 3, 1, "True".to_string(), 0);
    model.set_input(0, 4, 1, "false".to_string(), 0);
    model.set_input(0, 5, 1, "FALSE".to_string(), 0);
    model.set_input(0, 6, 1, "False".to_string(), 0);

    model.set_input(0, 1, 2, "=ISLOGICAL(A1)".to_string(), 0);
    model.set_input(0, 2, 2, "=ISLOGICAL(A2)".to_string(), 0);
    model.set_input(0, 3, 2, "=ISLOGICAL(A3)".to_string(), 0);
    model.set_input(0, 4, 2, "=ISLOGICAL(A4)".to_string(), 0);
    model.set_input(0, 5, 2, "=ISLOGICAL(A5)".to_string(), 0);
    model.set_input(0, 6, 2, "=ISLOGICAL(A6)".to_string(), 0);

    model.set_input(0, 1, 5, "=IF(false, True, FALSe)".to_string(), 0);

    model.evaluate();

    assert_eq!(model.get_text_at(0, 1, 1), *"TRUE");
    assert_eq!(model.get_text_at(0, 2, 1), *"TRUE");
    assert_eq!(model.get_text_at(0, 3, 1), *"TRUE");

    assert_eq!(model.get_text_at(0, 4, 1), *"FALSE");
    assert_eq!(model.get_text_at(0, 5, 1), *"FALSE");
    assert_eq!(model.get_text_at(0, 6, 1), *"FALSE");

    assert_eq!(model.get_text_at(0, 1, 2), *"TRUE");
    assert_eq!(model.get_text_at(0, 2, 2), *"TRUE");
    assert_eq!(model.get_text_at(0, 3, 2), *"TRUE");
    assert_eq!(model.get_text_at(0, 4, 2), *"TRUE");
    assert_eq!(model.get_text_at(0, 5, 2), *"TRUE");
    assert_eq!(model.get_text_at(0, 6, 2), *"TRUE");

    assert_eq!(
        model.get_formula_or_value(0, 1, 5),
        *"=IF(FALSE,TRUE,FALSE)"
    );
}

#[test]
fn test_get_sheet_dimensions() {
    let mut model = new_empty_model();
    assert_eq!(model.get_sheet_dimension(0), (1, 1, 1, 1));

    model.set_input(0, 30, 50, "Hello World".to_string(), 0);
    assert_eq!(model.get_sheet_dimension(0), (30, 50, 30, 50));

    model.set_input(0, 10, 15, "Hello World".to_string(), 0);
    assert_eq!(model.get_sheet_dimension(0), (10, 15, 30, 50));

    model.set_input(0, 5, 25, "Hello World".to_string(), 0);
    assert_eq!(model.get_sheet_dimension(0), (5, 15, 30, 50));

    model.set_input(0, 10, 250, "Hello World".to_string(), 0);
    assert_eq!(model.get_sheet_dimension(0), (5, 15, 30, 250));
}

#[test]
fn test_set_cell_style() {
    let mut model = new_empty_model();
    let mut style = model.get_style_for_cell(0, 1, 1);
    assert!(!style.font.b);

    style.font.b = true;
    assert!(model.set_cell_style(0, 1, 1, &style).is_ok());

    let mut style = model.get_style_for_cell(0, 1, 1);
    assert!(style.font.b);

    style.font.b = false;
    assert!(model.set_cell_style(0, 1, 1, &style).is_ok());

    let style = model.get_style_for_cell(0, 1, 1);
    assert!(!style.font.b);
}

#[test]
fn test_get_cell_style_index() {
    let mut model = new_empty_model();

    let mut style = model.get_style_for_cell(0, 1, 1);
    let style_index = model.get_cell_style_index(0, 1, 1);
    assert_eq!(style_index, 0);
    assert!(!style.font.b);

    style.font.b = true;
    assert!(model.set_cell_style(0, 1, 1, &style).is_ok());

    let style_index = model.get_cell_style_index(0, 1, 1);
    assert_eq!(style_index, 1);
}

#[test]
fn test_model_set_cells_with_values_styles() {
    let mut model = new_empty_model();
    // Inputs
    model.set_input(0, 1, 1, "21".to_string(), 0); // A1
    model.set_input(0, 2, 1, "2".to_string(), 0); // A2

    let style_index = model.get_cell_style_index(0, 1, 1);
    assert_eq!(style_index, 0);
    let mut style = model.get_style_for_cell(0, 1, 1);
    style.font.b = true;
    assert!(model.set_cell_style(0, 1, 1, &style).is_ok());
    assert!(model.set_cell_style(0, 2, 1, &style).is_ok());
    let style_index = model.get_cell_style_index(0, 1, 1);
    assert_eq!(style_index, 1);
    let style_index = model.get_cell_style_index(0, 2, 1);
    assert_eq!(style_index, 1);

    model.update_cell_with_number(0, 1, 2, 1.0);
    model.update_cell_with_number(0, 2, 1, 2.0);

    model.evaluate();

    // Styles are not modified
    let style_index = model.get_cell_style_index(0, 1, 1);
    assert_eq!(style_index, 1);
    let style_index = model.get_cell_style_index(0, 2, 1);
    assert_eq!(style_index, 1);
}

#[test]
fn test_style_fmt_id() {
    let mut model = new_empty_model();

    let mut style = model.get_style_for_cell(0, 1, 1);
    style.num_fmt = "#.##".to_string();
    assert!(model.set_cell_style(0, 1, 1, &style).is_ok());
    let style = model.get_style_for_cell(0, 1, 1);
    assert_eq!(style.num_fmt, "#.##");

    let mut style = model.get_style_for_cell(0, 10, 1);
    style.num_fmt = "$$#,##0.0000".to_string();
    assert!(model.set_cell_style(0, 10, 1, &style).is_ok());
    let style = model.get_style_for_cell(0, 10, 1);
    assert_eq!(style.num_fmt, "$$#,##0.0000");

    // Make sure old style is not touched
    let style = model.get_style_for_cell(0, 1, 1);
    assert_eq!(style.num_fmt, "#.##");
}

#[test]
fn test_set_sheet_color() {
    let mut model = new_empty_model();
    assert_eq!(model.workbook.worksheet(0).unwrap().color, Color::None);
    assert!(model.set_sheet_color(0, "#FFFAAA").is_ok());

    // Test new tab color is properly set
    assert_eq!(
        model.workbook.worksheet(0).unwrap().color,
        Color::RGB("#FFFAAA".to_string())
    );

    // Test we can remove it
    assert!(model.set_sheet_color(0, "").is_ok());
    assert_eq!(model.workbook.worksheet(0).unwrap().color, Color::None);
}

#[test]
fn test_set_sheet_color_invalid_sheet() {
    let mut model = new_empty_model();
    assert_eq!(
        model.set_sheet_color(10, "#FFFAAA"),
        Err("Invalid sheet index".to_string())
    );
}

#[test]
fn test_set_sheet_color_invalid() {
    let mut model = new_empty_model();
    // Boundaries
    assert!(model.set_sheet_color(0, "#FFFFFF").is_ok());
    assert!(model.set_sheet_color(0, "#000000").is_ok());

    assert_eq!(
        model.set_sheet_color(0, "#FFF"),
        Err("Invalid color: #FFF".to_string())
    );
    assert_eq!(
        model.set_sheet_color(0, "-#FFF"),
        Err("Invalid color: -#FFF".to_string())
    );
    assert_eq!(
        model.set_sheet_color(0, "#-FFF"),
        Err("Invalid color: #-FFF".to_string())
    );
    assert_eq!(
        model.set_sheet_color(0, "2FFFFFF"),
        Err("Invalid color: 2FFFFFF".to_string())
    );
    assert_eq!(
        model.set_sheet_color(0, "#FFFFFF1"),
        Err("Invalid color: #FFFFFF1".to_string())
    );
}

#[test]
fn set_input_autocomplete() {
    let mut model = new_empty_model();
    model._set("A1", "1");
    model._set("A2", "2");
    model.set_input(0, 3, 1, "=SUM(A1:A2".to_string(), 0);
    // This will fail anyway
    model.set_input(0, 4, 1, "=SUM(A1*".to_string(), 0);
    model.evaluate();

    assert_eq!(model._get_formula("A3"), "=SUM(A1:A2)");
    assert_eq!(model._get_text("A3"), "3");

    assert_eq!(model._get_formula("A4"), "=SUM(A1*");
    assert_eq!(model._get_text("A4"), "#ERROR!");
}

#[test]
fn test_get_cell_value_by_ref() {
    let mut model = new_empty_model();
    model._set("A1", "1");
    model._set("A2", "2");
    model.evaluate();

    // Correct
    assert_eq!(
        model.get_cell_value_by_ref("Sheet1!A1"),
        Ok(ExcelValue::Number(1.0))
    );

    // You need to specify full reference
    assert_eq!(
        model.get_cell_value_by_ref("A1"),
        Err("Error parsing reference: 'A1'".to_string())
    );

    // Error, it has a trailing space
    assert_eq!(
        model.get_cell_value_by_ref("Sheet1!A1 "),
        Err("Error parsing reference: 'Sheet1!A1 '".to_string())
    );
}

#[test]
fn test_get_formatted_cell_value() {
    let mut model = new_empty_model();
    model._set("A1", "foobar");
    model._set("A2", "true");
    model._set("A3", "");
    model._set("A4", "123.456");
    model._set("A5", "123.456");

    // change A5 format
    let mut style = model.get_style_for_cell(0, 5, 1);
    style.num_fmt = "$#,##0.00".to_string();
    model.set_cell_style(0, 5, 1, &style).unwrap();

    model.evaluate();

    assert_eq!(model.get_formatted_cell_value(0, 1, 1), "foobar");
    assert_eq!(model.get_formatted_cell_value(0, 2, 1), "TRUE");
    assert_eq!(model.get_formatted_cell_value(0, 3, 1), "");
    assert_eq!(model.get_formatted_cell_value(0, 4, 1), "123.456");
    assert_eq!(model.get_formatted_cell_value(0, 5, 1), "$123.46");
}
