#![allow(clippy::unwrap_used)]

use crate::model::Style;
use crate::test::util::new_empty_model;

#[test]
fn test_model_set_cells_with_values_styles() {
    let mut model = new_empty_model();
    // Inputs
    model.set_input(0, 1, 1, "21".to_string(), 0); // A1
    model.set_input(0, 2, 1, "42".to_string(), 0); // A2

    let style_base = model.get_style_for_cell(0, 1, 1);
    let mut style = style_base.clone();
    style.font.b = true;
    style.num_fmt = "#,##0.00".to_string();
    assert!(model.set_cell_style(0, 1, 1, &style).is_ok());

    let mut style = style_base;
    style.num_fmt = "#,##0.00".to_string();
    assert!(model.set_cell_style(0, 2, 1, &style).is_ok());
    let style: Style = model.get_style_for_cell(0, 2, 1);
    assert_eq!(style.num_fmt, "#,##0.00".to_string());
}

#[test]
fn test_named_styles() {
    let mut model = new_empty_model();
    model._set("A1", "42");
    let mut style = model.get_style_for_cell(0, 1, 1);
    style.font.b = true;
    assert!(model.set_cell_style(0, 1, 1, &style).is_ok());
    let bold_style_index = model.get_cell_style_index(0, 1, 1);
    let e = model.add_named_cell_style("bold", bold_style_index);
    assert!(e.is_ok());
    model._set("A2", "420");
    let a2_style_index = model.get_cell_style_index(0, 2, 1);
    assert!(a2_style_index != bold_style_index);
    let e = model.set_cell_style_by_name(0, 2, 1, "bold");
    assert!(e.is_ok());
    assert_eq!(model.get_cell_style_index(0, 2, 1), bold_style_index);
}

#[test]
fn test_create_named_style() {
    let mut model = new_empty_model();
    model._set("A1", "42");

    let mut style = model.get_style_for_cell(0, 1, 1);
    assert!(!style.font.b);

    style.font.b = true;
    let e = model.create_named_style("bold", &style);
    assert!(e.is_ok());

    let e = model.set_cell_style_by_name(0, 1, 1, "bold");
    assert!(e.is_ok());

    let style = model.get_style_for_cell(0, 1, 1);
    assert!(style.font.b);
}

#[test]
fn test_update_cell() -> Result<(), String> {
    let mut model = new_empty_model();
    model.add_equalto_read_only_styles()?;
    model._set("A1", "42");
    model.set_cell_style_by_name(0, 1, 1, "read_only_body")?;
    assert!(model.get_style_for_cell(0, 1, 1).read_only);
    model.update_cell_with_number(0, 1, 1, 43.0);
    assert!(model.get_style_for_cell(0, 1, 1).read_only);
    assert_eq!(model._get_text("A1"), "43");
    model.update_cell_with_text(0, 1, 1, "Hello World!");
    assert!(model.get_style_for_cell(0, 1, 1).read_only);
    assert_eq!(model._get_text("A1"), "Hello World!");
    model.update_cell_with_bool(0, 1, 1, true);
    assert!(model.get_style_for_cell(0, 1, 1).read_only);
    assert_eq!(model._get_text("A1"), "TRUE");
    Ok(())
}

#[test]
fn test_read_only_equalto_date_style() -> Result<(), String> {
    let mut model = new_empty_model();
    model.add_equalto_read_only_styles()?;
    model._set("A1", "1234");
    model.set_cell_style_by_name(0, 1, 1, "read_only_date")?;
    assert_eq!(model.get_style_for_cell(0, 1, 1).num_fmt, "dd/mm/yyyy;@");
    Ok(())
}

#[test]
fn test_set_sheet_column_style() -> Result<(), String> {
    let mut model = new_empty_model();
    model.add_equalto_read_only_styles()?;
    let read_only_body = model.get_style_index_by_name("read_only_body")?;
    let read_only_header = model.get_style_index_by_name("read_only_header")?;
    model.set_sheet_style(0, "read_only_body")?;
    model.set_sheet_column_style(0, 5, "read_only_header")?;
    // there should be three groups of cols from 1 to 4, 5 and from 6 to last one
    assert_eq!(model.get_cell_style_index(0, 1, 1), read_only_body);
    assert_eq!(model.get_cell_style_index(0, 13, 4), read_only_body);
    assert_eq!(model.get_cell_style_index(0, 15, 5), read_only_header);
    assert_eq!(model.get_cell_style_index(0, 7, 6), read_only_body);
    assert_eq!(model.get_cell_style_index(0, 31, 5), read_only_header);
    Ok(())
}

#[test]
fn test_column_styles_branches() -> Result<(), String> {
    let mut model = new_empty_model();
    model.add_equalto_read_only_styles()?;
    model.set_sheet_column_style(0, 5, "read_only_header")?;
    model.set_sheet_column_style(0, 4, "read_only_header")?;
    model.set_sheet_column_style(0, 6, "read_only_header")?;
    model.set_sheet_column_style(0, 5, "read_only_date")?;
    Ok(())
}

#[test]
fn test_column_styles_branches_2() -> Result<(), String> {
    let mut model = new_empty_model();
    model.add_equalto_read_only_styles()?;
    model.set_sheet_style(0, "read_only_body")?;
    model.set_sheet_column_style(0, 5, "read_only_header")?;
    model.set_sheet_column_style(0, 4, "read_only_header")?;
    model.set_sheet_column_style(0, 6, "read_only_header")?;
    Ok(())
}

#[test]
fn test_set_sheet_row_style() -> Result<(), String> {
    let mut model = new_empty_model();
    model.add_equalto_read_only_styles()?;
    let read_only_body = model.get_style_index_by_name("read_only_body")?;
    let read_only_header = model.get_style_index_by_name("read_only_header")?;
    model.set_sheet_style(0, "read_only_body")?;
    model.set_sheet_row_style(0, 5, "read_only_header")?;
    // there should be three groups of cols from 1 to 4, 5 and from 6 to last one
    assert_eq!(model.get_cell_style_index(0, 1, 1), read_only_body);
    assert_eq!(model.get_cell_style_index(0, 4, 41), read_only_body);
    assert_eq!(model.get_cell_style_index(0, 5, 51), read_only_header);
    assert_eq!(model.get_cell_style_index(0, 6, 61), read_only_body);
    assert_eq!(model.get_cell_style_index(0, 5, 1), read_only_header);

    // change the existing row style
    model.set_sheet_row_style(0, 5, "read_only_body")?;
    assert_eq!(model.get_cell_style_index(0, 1, 1), read_only_body);
    assert_eq!(model.get_cell_style_index(0, 4, 41), read_only_body);
    assert_eq!(model.get_cell_style_index(0, 5, 51), read_only_body);
    assert_eq!(model.get_cell_style_index(0, 6, 61), read_only_body);
    assert_eq!(model.get_cell_style_index(0, 5, 1), read_only_body);
    Ok(())
}

#[test]
#[should_panic(expected = "Style \'read_only_bdy\' not found")]
fn test_get_wrong_style() {
    let mut model = new_empty_model();
    model.add_equalto_read_only_styles().unwrap();
    let _ = model.get_style_index_by_name("read_only_bdy").unwrap();
}

#[test]
fn test_get_sheet_row_style() -> Result<(), String> {
    let mut model = new_empty_model();
    assert!(!model.is_row_read_only(0, 5));

    model.add_equalto_read_only_styles()?;
    assert!(!model.is_row_read_only(0, 1));
    // Set all the columns to read only
    model.set_sheet_style(0, "read_only_body")?;
    assert!(model.is_row_read_only(0, 1));
    assert!(model.is_row_read_only(0, 5));
    model.set_sheet_row_style(0, 5, "read_only_header")?;
    assert!(model.is_row_read_only(0, 5));
    Ok(())
}
