#![allow(clippy::unwrap_used)]

use crate::calc_result::{CellReference, Range};
use crate::expressions::utils::number_to_column;
use crate::graph_static::cell_is_in_range;
use crate::{graph_static::StaticDependencies, model::Model};

use crate::constants::LAST_ROW;
use crate::test::util::new_empty_model;

impl Model {
    fn _cell_reference_to_string(&self, cell_reference: &CellReference) -> String {
        let sheet = self.workbook.worksheets.get(cell_reference.sheet as usize);
        let column = number_to_column(cell_reference.column).unwrap();
        if cell_reference.sheet != 0 {
            format!("{}!{}{}", sheet.unwrap().name, column, cell_reference.row)
        } else {
            format!("{}{}", column, cell_reference.row)
        }
    }
    fn _cell_range_to_string(&self, range: &Range) -> String {
        let sheet = self.workbook.worksheets.get(range.left.sheet as usize);
        let column1 = number_to_column(range.left.column).unwrap();
        let column2 = number_to_column(range.right.column).unwrap();
        if range.left.sheet != 0 {
            format!(
                "{}!{}{}:{}{}",
                sheet.unwrap().name,
                column1,
                range.left.row,
                column2,
                range.right.row
            )
        } else {
            format!(
                "{}{}:{}{}",
                column1, range.left.row, column2, range.right.row
            )
        }
    }
    fn _check_expected(
        &self,
        t: StaticDependencies,
        non_strict: bool,
        cells: Vec<&str>,
        ranges: Vec<&str>,
    ) {
        assert_eq!(t.non_strict, non_strict);
        let mut t_cells: Vec<String> = t
            .cells
            .into_iter()
            .map(|r| self._cell_reference_to_string(&r))
            .collect();
        t_cells.sort();
        let mut cells = cells.clone();
        cells.sort_unstable();
        assert_eq!(cells, t_cells);
        let mut t_ranges: Vec<String> = t
            .ranges
            .into_iter()
            .map(|r| self._cell_range_to_string(&r))
            .collect();
        t_ranges.sort();
        let mut ranges = ranges.clone();
        ranges.sort_unstable();
        assert_eq!(ranges, t_ranges);
    }
}

#[test]
fn test_indirect_dependencies() {
    let mut model = new_empty_model();
    model._set("A1", "=B1+B2");
    model._set("B1", "=C1+C2");
    model.evaluate();

    let mut deps = StaticDependencies {
        non_strict: false,
        cells: vec![],
        ranges: vec![],
    };

    model.add_static_dependencies(
        &CellReference {
            sheet: 0,
            column: 1,
            row: 1,
        },
        &mut deps,
        &mut Vec::new(),
    );
    let dep_vec = deps
        .cells
        .iter()
        .map(|x| model._cell_reference_to_string(x))
        .collect::<Vec<_>>();
    assert_eq!(dep_vec, vec!["B1", "C1", "C2", "B2"]);
}

#[test]
fn test_indirect_dependencies_ranges() {
    let mut model = new_empty_model();
    model._set("A1", "=B1*SUM(B2:B5)");
    model._set("B3", "=C1+IF(C2, C3, SUM(S5:S10))");
    model._set("S7", "=AA4");
    model.evaluate();

    let mut deps = StaticDependencies {
        non_strict: false,
        cells: vec![],
        ranges: vec![],
    };

    model.add_static_dependencies(
        &CellReference {
            sheet: 0,
            column: 1,
            row: 1,
        },
        &mut deps,
        &mut Vec::new(),
    );
    let dep_vec = deps
        .cells
        .iter()
        .map(|x| model._cell_reference_to_string(x))
        .collect::<Vec<_>>();
    let dep_ranges = deps
        .ranges
        .iter()
        .map(|x| model._cell_range_to_string(x))
        .collect::<Vec<_>>();
    assert_eq!(dep_vec, vec!["B1", "C1", "C2", "C3", "AA4"]);
    assert_eq!(dep_ranges, vec!["S5:S10", "B2:B5"]);
}

#[test]
fn test_cell_does_not_depend_on_sheet() {
    let mut model = new_empty_model();
    model.add_sheet("Sheet3").unwrap();
    model.add_sheet("Sheet5").unwrap();
    model._set("A1", "=Sheet5!B1*SUM(B2:B5)");
    model._set("A2", "=B1+INDIRECT(\"A1\")");
    model.evaluate();

    // A1 does not depend on Sheet3
    assert!(model
        .cell_independent_of_sheets_and_cells("Sheet1!A1", &["Sheet3"], &[])
        .unwrap());

    // But it does depend on Sheet5
    assert!(!model
        .cell_independent_of_sheets_and_cells("Sheet1!A1", &["Sheet5"], &[])
        .unwrap());
    assert!(!model
        .cell_independent_of_sheets_and_cells("Sheet1!A1", &["Sheet3", "Sheet5"], &[])
        .unwrap());

    // Because INDIRECT is non-strict we cannot be sure with _static_ analysis that it does not depend on any sheet
    assert!(!model
        .cell_independent_of_sheets_and_cells("Sheet1!A2", &["Sheet3"], &[])
        .unwrap());
}

#[test]
fn test_cell_does_not_depend_on_cells() {
    let mut model = new_empty_model();
    model._set("A1", "=B1*SUM(B2:B5)");
    model._set("G7", "=B1*SUM(H1:ZZ1048576)*A1");
    model.evaluate();

    // A1 does not depend on C1
    assert!(model
        .cell_independent_of_sheets_and_cells("Sheet1!A1", &[], &["Sheet1!C1"])
        .unwrap());

    // It does depend on B1 though
    assert!(!model
        .cell_independent_of_sheets_and_cells("Sheet1!A1", &[], &["Sheet1!B1"])
        .unwrap());
    assert!(!model
        .cell_independent_of_sheets_and_cells("Sheet1!A1", &[], &["Sheet1!B2"])
        .unwrap());
    assert!(!model
        .cell_independent_of_sheets_and_cells("Sheet1!A1", &[], &["Sheet1!C1", "Sheet1!B1"])
        .unwrap());
}

#[test]
fn test_open_ranges() {
    let mut model = new_empty_model();
    model._set("A1", "=B1*SUM(B:B)");
    model._set("B123", "=C23");
    model._set("G7", "=B1*SUM(H:ZZ)*A1");
    model.evaluate();

    let mut deps = StaticDependencies {
        non_strict: false,
        cells: vec![],
        ranges: vec![],
    };

    model.add_static_dependencies(
        &CellReference {
            sheet: 0,
            column: 7, // "G"
            row: 7,
        },
        &mut deps,
        &mut Vec::new(),
    );

    let dep_vec = deps
        .cells
        .iter()
        .map(|x| model._cell_reference_to_string(x))
        .collect::<Vec<_>>();
    assert_eq!(dep_vec, vec!["B1", "A1", "C23"]);

    let dep_ranges = deps
        .ranges
        .iter()
        .map(|x| model._cell_range_to_string(x))
        .collect::<Vec<_>>();
    assert_eq!(
        dep_ranges,
        vec![format!("B1:B{}", LAST_ROW), format!("H1:ZZ{}", LAST_ROW)]
    );
}

#[test]
fn test_cyclic_dependencies() {
    let mut model = new_empty_model();
    model._set("A1", "=B1*C1");
    // There is an _static_ cyclic dependency but it is not a cyclic _runtime_ dependency
    model._set("B1", "=IF(TRUE, D1, B1*C2)");

    model.evaluate();

    let mut deps = StaticDependencies {
        non_strict: false,
        cells: vec![],
        ranges: vec![],
    };

    model.add_static_dependencies(
        &CellReference {
            sheet: 0,
            column: 1,
            row: 1,
        },
        &mut deps,
        &mut Vec::new(),
    );

    let dep_vec = deps
        .cells
        .iter()
        .map(|x| model._cell_reference_to_string(x))
        .collect::<Vec<_>>();

    assert_eq!(dep_vec, vec!["B1", "D1", "C2", "C1"]);
}

#[test]
fn test_cyclic_range_dependencies() {
    let mut model = new_empty_model();
    model._set("A1", "=SUM(B1:B7)");
    // There is an _static_ cyclic dependency but it is not a cyclic _runtime_ dependency
    model._set("B1", "=IF(TRUE, D1, SUM(B1:B7))");

    model.evaluate();

    let mut deps = StaticDependencies {
        non_strict: false,
        cells: vec![],
        ranges: vec![],
    };

    model.add_static_dependencies(
        &CellReference {
            sheet: 0,
            column: 1,
            row: 1,
        },
        &mut deps,
        &mut Vec::new(),
    );

    let dep_vec = deps
        .cells
        .iter()
        .map(|x| model._cell_reference_to_string(x))
        .collect::<Vec<_>>();

    let dep_ranges = deps
        .ranges
        .iter()
        .map(|x| model._cell_range_to_string(x))
        .collect::<Vec<_>>();

    assert_eq!(dep_vec, vec!["D1"]);
    // The range will be detected twice, and we keep it as is
    // although it would be easy no to repeat ranges it would be a bit more
    // difficult to ensure some ranges are not contained in others
    assert_eq!(dep_ranges, vec!["B1:B7", "B1:B7"]);
}

#[test]
fn test_cell_independent_from_sheets_and_cells() {
    let mut earth = new_empty_model();
    earth._set("A1", "=B1*SUM(Sheet3!B2:B5)");
    earth.add_sheet("Sheet3").unwrap();
    earth.add_sheet("Sheet5").unwrap();

    earth.evaluate();

    // There is Future on earth
    let result =
        earth.cell_independent_of_sheets_and_cells("Sheet1!A1", &["Future"], &["Sheet1!C1"]);
    assert!(result.is_err());

    // You need to specify the sheet of the reference
    let result = earth.cell_independent_of_sheets_and_cells("Sheet1!A1", &[], &["C1"]);
    assert!(result.is_err());

    assert!(!earth
        .cell_independent_of_sheets_and_cells("Sheet1!A1", &["Sheet3"], &[])
        .unwrap());
}

#[test]
fn test_cell_in_ranges() {
    let mut model = new_empty_model();
    model._set("A1", "=SUM(C1:C5)");

    model.evaluate();

    assert!(!model
        .cell_independent_of_sheets_and_cells("Sheet1!A1", &[], &["Sheet1!C1"])
        .unwrap());
}

#[test]
fn test_cell_is_in_range() {
    let cell = CellReference {
        sheet: 0,
        row: 5,
        column: 7,
    };
    let range = Range {
        left: CellReference {
            sheet: 0,
            column: 1,
            row: 1,
        },
        right: CellReference {
            sheet: 0,
            column: 10,
            row: 10,
        },
    };
    assert!(cell_is_in_range(&cell, &range));

    // different sheet
    assert!(!cell_is_in_range(
        &CellReference {
            sheet: 1,
            column: 5,
            row: 7
        },
        &range
    ));

    assert!(!cell_is_in_range(
        &CellReference {
            sheet: 0,
            column: 15,
            row: 17
        },
        &range
    ));

    // testing the border
    assert!(cell_is_in_range(
        &CellReference {
            sheet: 0,
            column: 10,
            row: 10
        },
        &range
    ));
}
