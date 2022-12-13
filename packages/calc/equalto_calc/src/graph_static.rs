use crate::{
    calc_result::{CellReference, Range},
    expressions::parser::Node,
    model::Model,
};

/// Static, _declared_ or _strict_ dependencies in a formula are the set of cells and ranges that are explicitly used.
/// That is, in the AST of the Node, they appear as a Reference or a Range Node.
/// Note that static dependencies might not be dependencies at all. For instance:
///  =IF(TRUE, A1, A2)
/// A2 would be a static dependency but would never get evaluated (remember that Excel does lazy evaluation)
/// Another example (@ is the implicit intersection operator):
///  =@A1:A10
/// Depending on _where_ this formula is evaluated it might have no dependency at all or just one.
/// 'Real' dependencies are called _dynamic_ dependencies or simply dependencies.
/// They depend (no pun intended) on the last evaluation and might be random:
///   =INDIRECT(CONCAT("A", RANDBETWEEN(1, 10))
/// This formula has no static dependencies and the only dynamic dependency will change every time.
/// Formulas whose dependencies are not known at 'compile time' are called _non-strict_.
/// Functions whose result does not depend only on it's arguments and the context are called _volatile_.
///   * RANDBETWEEN is strict but _volatile_.
///     We don't know the result of RANDBETWEEN(1, 100) even though we know all the arguments
///   * INDIRECT is non-strict and volatile.
///     We don't know the result of INDIRECT("A1") even though we know the arguments
///   * ROW() is strict and non-volatile.
///     Although we do not know the result by inspecting the arguments, we know it if we add the context.
/// IF (and most functions) is strict and non-volatile
/// IF(A1, A2, A3) (If we know the values of the arguments A1, A2 and A3 we can compute the result)
/// Note that the arguments of IF in the example are the values of A1, A2 and A3 but the argument of INDIRECT is the string "A1"
/// Another typical example of volatile function is TODAY
///
/// Two 'theorems':
/// * If a formula has no non-strict functions then all its dynamic dependencies are a subset of the static dependencies.
/// * The composition of strict functions is itself strict
///
/// As a corollary, although we cannot know the dynamic dependencies of a formula a compile time (ie without evaluation)
/// we can sometimes guarantee that they are in a subset of static dependencies.
///
/// We take a conservative approach in which non_strict might be true in the object but not in reality.
/// For instance IF(TRUE, 2, INDIRECT("A1")) is non_strict by us,
/// because you will need an evaluation to check the non_strict status.

/// Object representing the strict dependencies of a formula/cell
#[derive(Clone)]
pub struct StaticDependencies {
    pub non_strict: bool,
    pub cells: Vec<CellReference>,
    pub ranges: Vec<Range>,
}

impl StaticDependencies {
    // adds the dependencies to self
    fn add(&mut self, right: &StaticDependencies) {
        self.non_strict = self.non_strict || right.non_strict;
        self.cells.append(&mut right.cells.clone());
        self.ranges.append(&mut right.ranges.clone());
    }
}

// Returns true if the function is non-strict
fn is_non_strict(name: &str) -> bool {
    let list = vec!["OFFSET", "INDIRECT"];
    if list.contains(&name) {
        return true;
    }
    false
}

pub(crate) fn cell_is_in_range(cell: &CellReference, range: &Range) -> bool {
    if cell.sheet != range.left.sheet {
        return false;
    }
    if cell.row < range.left.row || cell.row > range.right.row {
        return false;
    }
    if cell.column < range.left.column || cell.column > range.right.column {
        return false;
    }
    true
}

fn get_node_static_direct_dependencies(
    node: &Node,
    column_ref: i32,
    row_ref: i32,
) -> StaticDependencies {
    match node {
        Node::ReferenceKind {
            sheet_name: _,
            sheet_index,
            absolute_row,
            absolute_column,
            row,
            column,
        } => {
            let mut row1 = *row;
            let mut column1 = *column;
            if !absolute_row {
                row1 += row_ref;
            }
            if !absolute_column {
                column1 += column_ref;
            }
            StaticDependencies {
                non_strict: false,
                cells: vec![CellReference {
                    sheet: *sheet_index,
                    column: column1,
                    row: row1,
                }],
                ranges: vec![],
            }
        }
        Node::RangeKind {
            sheet_name: _,
            sheet_index,
            absolute_row1,
            absolute_column1,
            row1,
            column1,
            absolute_row2,
            absolute_column2,
            row2,
            column2,
        } => {
            let r = Range {
                left: CellReference {
                    sheet: *sheet_index,
                    row: if *absolute_row1 {
                        *row1
                    } else {
                        *row1 + row_ref
                    },
                    column: if *absolute_column1 {
                        *column1
                    } else {
                        *column1 + column_ref
                    },
                },
                right: CellReference {
                    sheet: *sheet_index,
                    row: if *absolute_row2 {
                        *row2
                    } else {
                        *row2 + row_ref
                    },
                    column: if *absolute_column2 {
                        *column2
                    } else {
                        *column2 + column_ref
                    },
                },
            };
            StaticDependencies {
                non_strict: false,
                cells: vec![],
                ranges: vec![r],
            }
        }
        Node::OpRangeKind { left, right } => {
            // The range operator produces a new range that we don't control, that's non-strict
            let mut l = get_node_static_direct_dependencies(left, column_ref, row_ref);
            let r = get_node_static_direct_dependencies(right, column_ref, row_ref);
            l.add(&r);
            l.non_strict = true;
            l
        }
        Node::OpConcatenateKind { left, right } => {
            let mut l = get_node_static_direct_dependencies(left, column_ref, row_ref);
            let r = get_node_static_direct_dependencies(right, column_ref, row_ref);
            l.add(&r);
            l
        }
        Node::OpSumKind {
            kind: _,
            left,
            right,
        } => {
            let mut l = get_node_static_direct_dependencies(left, column_ref, row_ref);
            let r = get_node_static_direct_dependencies(right, column_ref, row_ref);
            l.add(&r);
            l
        }
        Node::OpProductKind {
            kind: _,
            left,
            right,
        } => {
            let mut l = get_node_static_direct_dependencies(left, column_ref, row_ref);
            let r = get_node_static_direct_dependencies(right, column_ref, row_ref);
            l.add(&r);
            l
        }
        Node::OpPowerKind { left, right } => {
            let mut l = get_node_static_direct_dependencies(left, column_ref, row_ref);
            let r = get_node_static_direct_dependencies(right, column_ref, row_ref);
            l.add(&r);
            l
        }
        Node::FunctionKind { name, args } => {
            let mut deps = StaticDependencies {
                non_strict: false,
                cells: vec![],
                ranges: vec![],
            };
            for arg in args {
                let a = get_node_static_direct_dependencies(arg, column_ref, row_ref);
                deps.add(&a);
            }
            if is_non_strict(name) {
                deps.non_strict = true;
            }
            deps
        }
        Node::CompareKind {
            kind: _,
            left,
            right,
        } => {
            let mut l = get_node_static_direct_dependencies(left, column_ref, row_ref);
            let r = get_node_static_direct_dependencies(right, column_ref, row_ref);
            l.add(&r);
            l
        }
        Node::UnaryKind { kind: _, right } => {
            get_node_static_direct_dependencies(right, column_ref, row_ref)
        }
        // Node::ArrayKind(_) => {}
        // Node::VariableKind(_) => {}
        // Node::WrongReferenceKind { sheet_name, absolute_row, absolute_column, row, column } => {}
        // Node::BooleanKind(_) => {}
        // Node::NumberKind(_) => {}
        // Node::StringKind(_) => {}
        // Node::ErrorKind(_) => {}
        // Node::ParseErrorKind { formula, message, position } => {}
        _ => StaticDependencies {
            non_strict: false,
            cells: vec![],
            ranges: vec![],
        },
    }
}

impl Model {
    /// Recursively adds all static dependencies of "cell" to "dependencies"
    /// For simplicity we carry a list of visited cells. Note that those are cells with formulas only.
    /// NOTE: dependencies.ranges might contain duplicated ranges
    pub(crate) fn add_static_dependencies(
        &self,
        cell: &CellReference,
        dependencies: &mut StaticDependencies,
        visited_cells: &mut Vec<CellReference>,
    ) {
        let sheet = cell.sheet;
        let row = cell.row;
        let column = cell.column;
        // If the cell has no formula, then has no dependencies
        if let Some(f) = self.get_cell_at(sheet, row, column).get_formula() {
            // NOTE: This will not stop at cyclic dependencies.
            // Static cyclic dependencies are not necessarily an error
            if visited_cells.contains(cell) {
                return;
            }
            visited_cells.push(cell.clone());

            // We first get all explicit dependencies. That is dependencies that appear on the formula
            let node = &self.parsed_formulas[sheet as usize][f as usize].clone();
            let direct_dependencies = get_node_static_direct_dependencies(node, column, row);

            let non_strict = direct_dependencies.non_strict;
            dependencies.non_strict = non_strict || dependencies.non_strict;

            // Add all direct dependencies and recurse into them if we don't have them yet
            for cell in &direct_dependencies.cells {
                if !dependencies.cells.contains(cell) {
                    dependencies.cells.push(cell.clone());
                    self.add_static_dependencies(cell, dependencies, visited_cells);
                }
            }

            // Loop over all direct dependent ranges
            // And the recurse into every cell in the range if it is not one of our dependencies already
            for range in &direct_dependencies.ranges {
                let cell_start = &range.left;
                let cell_end = &range.right;
                let sheet = cell_start.sheet;

                // constrain loping to sheet size.
                let (_, _, row_max, column_max) = self.get_sheet_dimension(sheet);

                let last_row = cell_end.row.min(row_max);
                let last_column = cell_end.column.min(column_max);

                for cell_row in cell_start.row..=last_row {
                    for cell_column in cell_start.column..=last_column {
                        let cell = CellReference {
                            sheet,
                            row: cell_row,
                            column: cell_column,
                        };
                        if !dependencies.cells.contains(&cell) {
                            self.add_static_dependencies(&cell, dependencies, visited_cells);
                        }
                    }
                }
                // NOTE: We do not check if range is in dependencies.ranges or if it is contained in another range
                dependencies.ranges.push(range.clone());
            }
        }
    }

    /// Returns true if we can assure that the formula on cell does not depend on any of the sheets
    /// and the given cells by doing static analysis.
    /// That is returns true if the formula does not have a static dependency on the sheet
    /// and there are no non-strict functions.
    /// Note that it is possible that the formula does not depend on a sheet but this function returns false
    /// The formula =IF(TRUE, 1, Sheet1!A2) does not depend on `Sheet2` but this function will return false
    /// The formula =INDIRECT("A1") does not depend on Sheet1 but the function will return false
    pub fn cell_independent_of_sheets_and_cells(
        &self,
        input_cell: &str,
        sheets: &[&str],
        cells: &[&str],
    ) -> Result<bool, String> {
        // First parse the input cell
        let parsed_input_cell = self
            .parse_reference(input_cell)
            .ok_or(format!("Can't parse reference: {}", input_cell))?;

        // Get a list with all the "forbidden" sheet indices from the sheet names
        let mut sheet_indices: Vec<i32> = vec![];
        for sheet_name in sheets {
            if let Some(sheet_index) = self.get_sheet_index_by_name(sheet_name) {
                sheet_indices.push(sheet_index as i32);
            } else {
                return Err(format!("Bad sheet name: {}", sheet_name));
            }
        }

        // Parse all the "forbidden" cells
        let mut parsed_cells = vec![];
        for &cell_str in cells {
            let parsed_input = self
                .parse_reference(cell_str)
                .ok_or(format!("Can't parse reference: {}", cell_str))?;
            parsed_cells.push(parsed_input);
        }

        // Get all the static dependencies of input cell
        let mut dependencies = StaticDependencies {
            non_strict: false,
            cells: vec![],
            ranges: vec![],
        };
        self.add_static_dependencies(&parsed_input_cell, &mut dependencies, &mut Vec::new());

        // 1. Make sure the are no non-strict functions
        if dependencies.non_strict {
            return Ok(false);
        }

        // 2. Make sure none of the cells in the dependencies is in a "forbidden" sheet.
        //    And that that no cell in the dependencies is one of the "forbidden" cells.
        for cell in dependencies.cells {
            if sheet_indices.contains(&cell.sheet) {
                return Ok(false);
            }
            if parsed_cells.contains(&cell) {
                return Ok(false);
            }
        }
        // 3. Ensure that no range is in a "forbidden" sheet.
        //    And that no "forbidden" cell is in one of the dependent ranges.
        for range in dependencies.ranges {
            if sheet_indices.contains(&range.left.sheet)
                || sheet_indices.contains(&range.right.sheet)
            {
                return Ok(false);
            }
            for parsed_cell in &parsed_cells {
                if cell_is_in_range(parsed_cell, &range) {
                    return Ok(false);
                }
            }
        }
        Ok(true)
    }
}
