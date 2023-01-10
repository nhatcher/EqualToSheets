#![allow(clippy::unwrap_used)]

use crate::constants::{LAST_COLUMN, LAST_ROW};
use crate::model::Model;
use crate::test::util::new_empty_model;

#[test]
fn test_shift_cells_right_empty() {
    let mut model = new_empty_model();
    let r = model.shift_cells_right(0, 1, 2, 1);
    assert!(r.is_ok());
}

#[test]
fn test_shift_cells_right() {
    let mut model = new_empty_model();
    // We populate cells A1 to C1
    model._set("A1", "1");
    model._set("B1", "2");
    model._set("C1", "=B1*2");

    model._set("S1", "300");
    model._set("T1", "=S1*5");

    model._set("Y1", "=A1");

    model._set("A3", "=A1");

    model._set("H1", "=D5");
    model._set("L1", "=$D$5");

    // And another cell outside from that
    model._set("AA30", "=C1*10");

    model.evaluate();
    assert_eq!(model._get_text("C1"), *"4");
    // We push to the right from cell B1
    let r = model.shift_cells_right(0, 1, 2, 1);
    assert!(r.is_ok());
    model.evaluate();

    // Check B1 is now empty
    assert!(model.is_empty_cell(0, 1, 2).unwrap());

    // C1 is the old B1
    assert_eq!(model._get_text("C1"), *"2");

    // Check D1 has the right displaced formula
    assert_eq!(model._get_text("D1"), *"4");
    assert_eq!(model._get_formula("D1"), *"=C1*2");

    assert_eq!(model._get_text("AA30"), *"40");
    assert_eq!(model._get_formula("AA30"), *"=D1*10");

    // Cell A3 is NOT displaced
    assert_eq!(model._get_formula("A3"), *"=A1");

    // Old H1 is I1
    assert_eq!(model._get_formula("I1"), *"=D5");
    // Old L1 is M1
    assert_eq!(model._get_formula("M1"), *"=$D$5");

    assert_eq!(model._get_formula("Z1"), *"=A1");
    assert_eq!(model._get_formula("U1"), *"=T1*5");
    assert_eq!(model._get_text("U1"), *"1500");
}

#[test]
fn test_shift_cells_left_empty() {
    let mut model = new_empty_model();
    let r = model.shift_cells_left(0, 1, 2, 1);
    assert!(r.is_ok());
}

#[test]
fn test_shift_cells_left() {
    let mut model = new_empty_model();
    // We populate cells A1 to C1
    model._set("A1", "1");
    model._set("B1", "2");
    model._set("C1", "=B1*2");

    model._set("S1", "300");
    model._set("T1", "=S1*5");

    model._set("Y1", "=A1");

    model._set("A3", "=A1");

    // And another cell outside from that
    model._set("AA30", "=C1*10");
    model.evaluate();
    let result = model._get_text("C1");
    assert_eq!(result, *"4");

    // Remove cell B1
    let r = model.shift_cells_left(0, 1, 2, 1);
    assert!(r.is_ok());
    model.evaluate();

    // the old C1 is moved to B1 that used to refer to B1 that does not exist anymore
    assert_eq!(model._get_formula("B1"), *"=#REF!*2");
    // S1 and T1 are just displaced to the left
    assert_eq!(model._get_formula("R1"), *"300");
    assert_eq!(model._get_formula("S1"), *"=R1*5");
    // Old Y1 is now X1 but formula unchanged
    assert_eq!(model._get_formula("X1"), *"=A1");
    // A3 should be unchanged
    assert_eq!(model._get_formula("A3"), *"=A1");
    // AA30 use to refer to C1 now should be B1
    assert_eq!(model._get_formula("AA30"), *"=B1*10");
}

#[test]
fn test_shift_cells_down_empty() {
    let mut model = new_empty_model();
    let r = model.shift_cells_down(0, 6, 3, 3);
    assert!(r.is_ok());
}

#[test]
fn test_shift_cells_down() {
    let mut model = new_empty_model();
    // We populate cells in column C
    model._set("C5", "1");
    model._set("C6", "2");
    model._set("C7", "=C6*2");

    model._set("C100", "300");
    model._set("C101", "=C100*5");

    model._set("C200", "=C1");

    model._set("A3", "=C1");

    model._set("C30", "=D5");
    model._set("C31", "=$D$5");

    // And another cell outside from that
    model._set("AA30", "=C1*10");

    model.evaluate();

    assert_eq!(model._get_text("C7"), *"4");

    // We push down three cells from cell C6
    let r = model.shift_cells_down(0, 6, 3, 3);
    assert!(r.is_ok());
    model.evaluate();

    // C7 is now C10
    assert_eq!(model._get_formula("C10"), *"=C9*2");

    // C100 and C101 are C103 and C104
    assert_eq!(model._get_formula("C103"), *"300");
    assert_eq!(model._get_formula("C104"), *"=C103*5");

    // C200 would be C203
    assert_eq!(model._get_formula("C203"), *"=C1");

    // A3 would be unchanged
    assert_eq!(model._get_formula("A3"), *"=C1");

    assert_eq!(model._get_formula("C33"), "=D5");
    assert_eq!(model._get_formula("C34"), "=$D$5");

    assert_eq!(model._get_formula("AA30"), "=C1*10");

    // inserting a negative number of rows fails:
    let r = model.insert_rows(0, 6, -5);
    assert!(r.is_err());
    let r = model.insert_rows(0, 6, 0);
    assert!(r.is_err());

    // If you have data at the very end it fails
    model._set(&format!("X{}", LAST_ROW - 1), "300");
    let r = model.insert_rows(0, 6, 2);
    assert!(r.is_err());

    // You can insert 1 in this case :)
    let r = model.insert_rows(0, 6, 1);
    assert!(r.is_ok());
}

#[test]
fn test_shift_cells_up_empty() {
    let mut model = new_empty_model();
    let r = model.shift_cells_up(0, 6, 3, 1);
    assert!(r.is_ok());
}

#[test]
fn test_shift_cells_up() {
    let mut model = new_empty_model();
    // We populate cells in column C
    model._set("C5", "1");
    model._set("C6", "2");
    model._set("C7", "=C6*2");

    model._set("C100", "300");
    model._set("C101", "=C100*5");

    model._set("C200", "=C1");

    model._set("A3", "=C1");

    model._set("C30", "=D5");
    model._set("C31", "=$D$5");

    // And another cell outside from that
    model._set("AA30", "=C1*10");

    model.evaluate();

    assert_eq!(model._get_text("C7"), *"4");

    // We delete cell C6 (push up)
    let r = model.shift_cells_up(0, 6, 3, 1);
    assert!(r.is_ok());
    model.evaluate();

    // C7 is now C6
    assert_eq!(model._get_formula("C6"), *"=#REF!*2");

    // C100 and C101 are C99 and C100
    assert_eq!(model._get_formula("C99"), *"300");
    assert_eq!(model._get_formula("C100"), *"=C99*5");

    // C200 would be C199
    assert_eq!(model._get_formula("C199"), *"=C1");

    // A3 would be unchanged
    assert_eq!(model._get_formula("A3"), *"=C1");

    assert_eq!(model._get_formula("C29"), "=D5");
    assert_eq!(model._get_formula("C30"), "=$D$5");

    assert_eq!(model._get_formula("AA30"), "=C1*10");
}

#[test]
fn test_insert_columns() {
    let mut model = new_empty_model();
    // We populate cells A1 to C1
    model._set("A1", "1");
    model._set("B1", "2");
    model._set("C1", "=B1*2");

    model._set("F1", "=B1");

    model._set("L11", "300");
    model._set("M11", "=L11*5");

    model.evaluate();
    assert_eq!(model._get_text("C1"), *"4");

    // Let's insert 5 columns in column F (6)
    let r = model.insert_columns(0, 6, 5);
    assert!(r.is_ok());
    model.evaluate();

    // Check F1 is now empty
    assert!(model.is_empty_cell(0, 1, 6).unwrap());

    // The old F1 is K1
    assert_eq!(model._get_formula("K1"), *"=B1");

    // L11 and M11 are Q11 and R11
    assert_eq!(model._get_formula("Q11"), *"300");
    assert_eq!(model._get_formula("R11"), *"=Q11*5");

    assert_eq!(model._get_formula("C1"), "=B1*2");
    assert_eq!(model._get_formula("A1"), "1");

    // inserting a negative number of columns fails:
    let r = model.insert_columns(0, 6, -5);
    assert!(r.is_err());
    let r = model.insert_columns(0, 6, -5);
    assert!(r.is_err());

    // If you have data at the very ebd it fails
    model._set("XFC12", "300");
    let r = model.insert_columns(0, 6, 5);
    assert!(r.is_err());
}

#[test]
fn test_insert_rows() {
    let mut model = new_empty_model();

    model._set("C4", "3");
    model._set("C5", "7");
    model._set("C6", "=C5");

    model._set("H11", "=C4");

    model._set("R10", "=C6");

    model.evaluate();

    // Let's insert 5 rows in row 6
    let r = model.insert_rows(0, 6, 5);
    assert!(r.is_ok());
    model.evaluate();

    // Check C6 is now empty
    assert!(model.is_empty_cell(0, 6, 3).unwrap());

    // Old C6 is now C11
    assert_eq!(model._get_formula("C11"), *"=C5");
    assert_eq!(model._get_formula("H16"), *"=C4");

    assert_eq!(model._get_formula("R15"), *"=C11");
    assert_eq!(model._get_text("C4"), *"3");
    assert_eq!(model._get_text("C5"), *"7");
}

#[test]
fn test_insert_rows_styles() {
    let mut model = new_empty_model();

    assert!((21.0 - model.get_row_height(0, 10)).abs() < f64::EPSILON);
    // sets height 42 in row 10
    model.set_row_height(0, 10, 42.0);
    assert!((42.0 - model.get_row_height(0, 10)).abs() < f64::EPSILON);

    // Let's insert 5 rows in row 3
    let r = model.insert_rows(0, 3, 5);
    assert!(r.is_ok());

    // Row 10 has the default height
    assert!((21.0 - model.get_row_height(0, 10)).abs() < f64::EPSILON);

    // Row 10 is now row 15
    assert!((42.0 - model.get_row_height(0, 15)).abs() < f64::EPSILON);
}

#[test]
fn test_delete_rows_styles() {
    let mut model = new_empty_model();

    assert!((21.0 - model.get_row_height(0, 10)).abs() < f64::EPSILON);
    // sets height 42 in row 10
    model.set_row_height(0, 10, 42.0);
    assert!((42.0 - model.get_row_height(0, 10)).abs() < f64::EPSILON);

    // Let's delete 5 rows in row 3 (3-8)
    let r = model.delete_rows(0, 3, 5);
    assert!(r.is_ok());

    // Row 10 has the default height
    assert!((21.0 - model.get_row_height(0, 10)).abs() < f64::EPSILON);

    // Row 10 is now row 5
    assert!((42.0 - model.get_row_height(0, 5)).abs() < f64::EPSILON);
}

#[test]
fn test_delete_columns() {
    let mut model = new_empty_model();

    model._set("C4", "3");
    model._set("D4", "7");
    model._set("E4", "=D4");
    model._set("F4", "=C4");

    model._set("H11", "=D4");

    model._set("R10", "=C6");

    model._set("M5", "300");
    model._set("N5", "=M5*6");

    model._set("A1", "=SUM(M5:N5)");
    model._set("A2", "=SUM(C4:M4)");
    model._set("A3", "=SUM(E4:M4)");

    model.evaluate();

    // We delete columns D and E
    let r = model.delete_columns(0, 4, 2);
    assert!(r.is_ok());
    model.evaluate();

    // Old H11 will be F11 and contain =#REF!
    assert_eq!(model._get_formula("F11"), *"=#REF!");

    // Old F4 will be D4 now
    assert_eq!(model._get_formula("D4"), *"=C4");

    // Old N5 will be L5
    assert_eq!(model._get_formula("L5"), *"=K5*6");

    // Range in A1 is displaced correctly
    assert_eq!(model._get_formula("A1"), *"=SUM(K5:L5)");

    // Note that range in A2 would contain some of the deleted cells
    // A long as the borders of the range are not included that's ok.
    assert_eq!(model._get_formula("A2"), *"=SUM(C4:K4)");

    // FIXME: In Excel this would be (lower limit won't change)
    // assert_eq!(model._get_formula("A3"), *"=SUM(E4:K4)");
    assert_eq!(model._get_formula("A3"), *"=SUM(#REF!:K4)");
}

#[test]
fn test_delete_rows() {
    let mut model = new_empty_model();

    model._set("C4", "4");
    model._set("C5", "5");
    model._set("C6", "6");
    model._set("C7", "=C6*2");

    model._set("C72", "=C1*3");

    model.evaluate();

    // We delete rows 5, 6
    let r = model.delete_rows(0, 5, 2);
    assert!(r.is_ok());
    model.evaluate();

    assert_eq!(model._get_formula("C5"), *"=#REF!*2");
    assert_eq!(model._get_formula("C70"), *"=C1*3");
}

#[test]
fn test_swap_cells_with_formulas() {
    let mut model = new_empty_model();
    model._set("B1", "=H2+H5");
    model._set("C1", "=M5");

    let r = model.swap_cells_in_row(0, 1, 2, 3);
    assert!(r.is_ok());
    model.evaluate();
    assert_eq!(model._get_formula("B1"), "=M5");
    assert_eq!(model._get_formula("C1"), "=H2+H5");
}

#[test]
fn test_swap_cells_with_self_referencing_formulas() {
    let mut model = new_empty_model();
    model._set("B1", "=H2*C1");
    model._set("C1", "=B1");

    let r = model.swap_cells_in_row(0, 1, 2, 3);
    assert!(r.is_ok());
    model.evaluate();
    assert_eq!(model._get_formula("B1"), "=C1");
    assert_eq!(model._get_formula("C1"), "=H2*B1");
}

#[test]
fn test_swap_cells() {
    let mut model = new_empty_model();
    model._set("B4", "4");
    model._set("C4", "5");

    // Simple
    model._set("R5", "=C4");
    model._set("R6", "=B4");

    // In operations
    model._set("D6", "=C4*$C$4+B4*$B$4");

    // In functions
    model._set("AD45", "=SUM(C4, $C$4, B4, $B$4)");

    // Different cells
    model._set("F8", "=D6");
    model.evaluate();
    let r = model.swap_cells_in_row(0, 4, 2, 3);
    assert!(r.is_ok());
    model.evaluate();

    // Cells themselves are swapped
    assert_eq!(model._get_text("B4"), *"5");
    assert_eq!(model._get_text("C4"), *"4");

    // Simple
    model._set("R5", "=B4");
    model._set("R6", "=C4");

    // In operations
    assert_eq!(model._get_formula("D6"), "=B4*$B$4+C4*$C$4");

    // In functions
    model._set("AD45", "=SUM(B4, $B$4, C4, $C$4)");

    // Other references are not changed
    assert_eq!(model._get_formula("F8"), "=D6");
}

#[test]
fn test_swap_cells_with_styles() {
    let mut model = new_empty_model();
    model._set("B4", "4");
    model._set("C4", "5");

    // create two styles
    let style_base = model.get_style_for_cell(0, 1, 1);

    let mut style = style_base.clone();
    style.font.b = true;
    // B4
    assert!(model.set_cell_style(0, 4, 2, &style).is_ok());

    let mut style = style_base;
    style.num_fmt = "#,##0.00".to_string();
    // C4
    assert!(model.set_cell_style(0, 4, 3, &style).is_ok());
    model.evaluate();

    let index_b4 = model.get_cell_style_index(0, 4, 2);
    let index_c4 = model.get_cell_style_index(0, 4, 3);
    assert!(index_b4 != index_c4);

    // swap cells
    let r = model.swap_cells_in_row(0, 4, 2, 3);
    assert!(r.is_ok());

    model.evaluate();

    let index_b4_new = model.get_cell_style_index(0, 4, 2);
    let index_c4_new = model.get_cell_style_index(0, 4, 3);

    assert_eq!(index_b4, index_c4_new);
    assert_eq!(index_c4, index_b4_new);
}

// E	F	G	H	I	J	K
// 			3	1	1	2
// 			4	2	5	8
// 			-2	3	6	7
fn populate_table(model: &mut Model) {
    model._set("G1", "3");
    model._set("H1", "1");
    model._set("I1", "1");
    model._set("J1", "2");

    model._set("G2", "4");
    model._set("H2", "2");
    model._set("I2", "5");
    model._set("J2", "8");

    model._set("G3", "-2");
    model._set("H3", "3");
    model._set("I3", "6");
    model._set("J3", "7");
}

#[test]
fn test_move_column_right() {
    let mut model = new_empty_model();
    populate_table(&mut model);
    model._set("E3", "=G3");
    model._set("E4", "=H3");
    model._set("E5", "=SUM(G3:J7)");
    model._set("E6", "=SUM(G3:G7)");
    model._set("E7", "=SUM(H3:H7)");
    model.evaluate();

    // Wee swap column G with column H
    let result = model.move_column_action(0, 7, 1);
    assert!(result.is_ok());
    model.evaluate();

    assert_eq!(model._get_formula("E3"), "=H3");
    assert_eq!(model._get_formula("E4"), "=G3");
    assert_eq!(model._get_formula("E5"), "=SUM(H3:J7)");
    assert_eq!(model._get_formula("E6"), "=SUM(H3:H7)");
    assert_eq!(model._get_formula("E7"), "=SUM(G3:G7)");
}

#[test]
fn tets_move_column_error() {
    let mut model = new_empty_model();
    model.evaluate();

    let result = model.move_column_action(0, 7, -10);
    assert!(result.is_err());

    let result = model.move_column_action(0, -7, 20);
    assert!(result.is_err());

    let result = model.move_column_action(0, LAST_COLUMN, 1);
    assert!(result.is_err());

    let result = model.move_column_action(0, LAST_COLUMN + 1, -10);
    assert!(result.is_err());

    // This works
    let result = model.move_column_action(0, LAST_COLUMN, -1);
    assert!(result.is_ok());
}

// A  B  C  D  E  F  G   H  I  J   K   L   M   N   O   P   Q   R
// 1  2  3  4  5  6  7   8  9  10  11  12  13  14  15  16  17  18
