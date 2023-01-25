use std::{collections::HashMap, io::Read, num::ParseIntError};

use equalto_calc::{
    expressions::{
        parser::{stringify::to_rc_format, Parser},
        token::{get_error_by_english_name, Error},
        types::CellReferenceRC,
        utils::column_to_number,
    },
    types::{Cell, Col, Color, Comment, DefinedName, Row, SheetData, SheetState, Worksheet},
};
use roxmltree::Node;
use serde::{Deserialize, Serialize};

use crate::error::XlsxError;

use super::util::{get_attribute, get_color, get_number};

#[derive(Serialize, Deserialize, Debug)]
pub(crate) struct Sheet {
    pub(crate) name: String,
    pub(crate) sheet_id: u32,
    pub(crate) id: String,
    pub(crate) state: SheetState,
}

#[derive(Serialize, Deserialize, Debug)]
pub(crate) struct WorkbookXML {
    pub(crate) worksheets: Vec<Sheet>,
    pub(crate) defined_names: Vec<DefinedName>,
}

#[derive(Serialize, Deserialize, Debug)]
pub(crate) struct Relationship {
    pub(crate) target: String,
    pub(crate) rel_type: String,
}

fn get_column_from_ref(s: &str) -> String {
    let cs = s.chars();
    let mut column = Vec::<char>::new();
    for c in cs {
        if !c.is_ascii_digit() {
            column.push(c);
        }
    }
    column.into_iter().collect()
}

fn load_dimension(ws: Node) -> String {
    // <dimension ref="A1:O18"/>
    let application_nodes = ws
        .children()
        .filter(|n| n.has_tag_name("dimension"))
        .collect::<Vec<Node>>();
    if application_nodes.len() == 1 {
        application_nodes[0]
            .attribute("ref")
            .unwrap_or("A1")
            .to_string()
    } else {
        "A1".to_string()
    }
}

fn load_columns(ws: Node) -> Result<Vec<Col>, XlsxError> {
    // cols
    // <cols>
    //     <col min="5" max="5" width="38.26953125" customWidth="1"/>
    //     <col min="6" max="6" width="9.1796875" style="1"/>
    //     <col min="8" max="8" width="4" customWidth="1"/>
    // </cols>
    let mut cols = Vec::new();
    let columns = ws
        .children()
        .filter(|n| n.has_tag_name("cols"))
        .collect::<Vec<Node>>();
    if columns.len() == 1 {
        for col in columns[0].children() {
            let min = get_attribute(&col, "min")?;
            let min = min.parse::<i32>()?;
            let max = get_attribute(&col, "max")?;
            let max = max.parse::<i32>()?;
            let width = get_attribute(&col, "width")?;
            let width = width.parse::<f64>()?;
            let custom_width = matches!(col.attribute("customWidth"), Some("1"));
            let style = col
                .attribute("style")
                .map(|s| s.parse::<i32>().unwrap_or(0));
            cols.push(Col {
                min,
                max,
                width,
                custom_width,
                style,
            })
        }
    }
    Ok(cols)
}

fn load_merge_cells(ws: Node) -> Result<Vec<String>, XlsxError> {
    // 18.3.1.55 Merge Cells
    // <mergeCells count="1">
    //    <mergeCell ref="K7:L10"/>
    // </mergeCells>
    let mut merge_cells = Vec::new();
    let merge_cells_nodes = ws
        .children()
        .filter(|n| n.has_tag_name("mergeCells"))
        .collect::<Vec<Node>>();
    if merge_cells_nodes.len() == 1 {
        for merge_cell in merge_cells_nodes[0].children() {
            let reference = get_attribute(&merge_cell, "ref")?.to_string();
            merge_cells.push(reference);
        }
    }
    Ok(merge_cells)
}

fn load_sheet_color(ws: Node) -> Result<Color, XlsxError> {
    // <sheetPr>
    //     <tabColor theme="5" tint="-0.249977111117893"/>
    // </sheetPr>
    let mut color = Color::None;
    let sheet_pr = ws
        .children()
        .filter(|n| n.has_tag_name("sheetPr"))
        .collect::<Vec<Node>>();
    if sheet_pr.len() == 1 {
        let tabs = sheet_pr[0]
            .children()
            .filter(|n| n.has_tag_name("tabColor"))
            .collect::<Vec<Node>>();
        if tabs.len() == 1 {
            color = get_color(tabs[0])?;
        }
    }
    Ok(color)
}

fn load_comments<R: Read + std::io::Seek>(
    archive: &mut zip::read::ZipArchive<R>,
    path: &str,
) -> Result<Vec<Comment>, XlsxError> {
    let mut comments = Vec::new();
    let mut file = archive.by_name(path)?;
    let mut text = String::new();
    file.read_to_string(&mut text)?;
    let doc = roxmltree::Document::parse(&text)?;
    let ws = doc
        .root()
        .first_child()
        .ok_or_else(|| XlsxError::Xml("Corrupt XML structure".to_string()))?;
    let comment_list = ws
        .children()
        .filter(|n| n.has_tag_name("commentList"))
        .collect::<Vec<Node>>();
    if comment_list.len() == 1 {
        for comment in comment_list[0].children() {
            let text = comment
                .descendants()
                .filter(|n| n.has_tag_name("t"))
                .map(|n| n.text().unwrap().to_string())
                .collect::<Vec<String>>()
                .join("");
            let cell_ref = get_attribute(&comment, "ref")?.to_string();
            // TODO: Read author_name from the list of authors
            let author_name = "".to_string();
            comments.push(Comment {
                text,
                author_name,
                author_id: None,
                cell_ref,
            });
        }
    }

    Ok(comments)
}

// This parses Sheet1!AS23 into sheet, column and row
// FIXME: This is buggy. Does not check that is a valid sheet name or column
// There is a similar named function in equalto_calc. We probably should fix both at the same time.
// NB: Maybe use regexes for this?
fn parse_reference(s: &str) -> Result<CellReferenceRC, ParseIntError> {
    let bytes = s.as_bytes();
    let mut sheet_name = "".to_string();
    let mut column = "".to_string();
    let mut row = "".to_string();
    let mut state = "sheet"; // "sheet", "col", "row"
    for &byte in bytes {
        match state {
            "sheet" => {
                if byte == b'!' {
                    state = "col"
                } else {
                    sheet_name.push(byte as char);
                }
            }
            "col" => {
                if byte.is_ascii_alphabetic() {
                    column.push(byte as char);
                } else {
                    state = "row";
                    row.push(byte as char);
                }
            }
            _ => {
                row.push(byte as char);
            }
        }
    }
    Ok(CellReferenceRC {
        sheet: sheet_name,
        row: row.parse::<i32>()?,
        column: column_to_number(&column),
    })
}

fn from_a1_to_rc(
    formula: String,
    worksheets: &[String],
    context: String,
) -> Result<String, XlsxError> {
    let mut parser = Parser::new(worksheets.to_owned());
    let cell_reference = parse_reference(&context)?;
    let t = parser.parse(&formula, &Some(cell_reference));
    Ok(to_rc_format(&t))
}

fn get_formula_index(formula: &str, shared_formulas: &[String]) -> Option<i32> {
    for (index, f) in shared_formulas.iter().enumerate() {
        if f == formula {
            return Some(index as i32);
        }
    }
    None
}

fn get_cell_from_excel(
    cell_value: Option<&str>,
    cell_type: &str,
    cell_style: i32,
    formula_index: i32,
    sheet_name: &str,
    cell_ref: &str,
) -> Cell {
    // Possible cell types:
    // 18.18.11 ST_CellType (Cell Type)
    //   b (Boolean)
    //   d (Date)
    //   e (Error)
    //   inlineStr (Inline String)
    //   n (Number)
    //   s (Shared String)
    //   str (String)

    if formula_index == -1 {
        match cell_type {
            "b" => Cell::BooleanCell {
                v: cell_value == Some("1"),
                s: cell_style,
            },
            "n" => Cell::NumberCell {
                v: cell_value.unwrap_or("0").parse::<f64>().unwrap_or(0.0),
                s: cell_style,
            },
            "e" => Cell::ErrorCell {
                ei: get_error_by_english_name(cell_value.unwrap_or("#ERROR!"))
                    .unwrap_or(Error::ERROR),
                s: cell_style,
            },
            "s" => Cell::SharedString {
                si: cell_value.unwrap_or("0").parse::<i32>().unwrap_or(0),
                s: cell_style,
            },
            "str" => {
                // We are assuming that all strings in cells without a formula in Excel are shared strings.
                // Not implemented
                println!("Invalid type (str) in {}!{}", sheet_name, cell_ref);
                Cell::ErrorCell {
                    ei: Error::NIMPL,
                    s: cell_style,
                }
            }
            "d" => {
                // Not implemented
                println!("Invalid type (d) in {}!{}", sheet_name, cell_ref);
                Cell::ErrorCell {
                    ei: Error::NIMPL,
                    s: cell_style,
                }
            }
            "inlineStr" => {
                // Not implemented
                println!("Invalid type (inlineStr) in {}!{}", sheet_name, cell_ref);
                Cell::ErrorCell {
                    ei: Error::NIMPL,
                    s: cell_style,
                }
            }
            "empty" => Cell::EmptyCell { s: cell_style },
            _ => {
                // error
                println!(
                    "Unexpected type ({}) in {}!{}",
                    cell_type, sheet_name, cell_ref
                );
                Cell::ErrorCell {
                    ei: Error::ERROR,
                    s: cell_style,
                }
            }
        }
    } else {
        match cell_type {
            "b" => Cell::CellFormulaBoolean {
                f: formula_index,
                v: cell_value == Some("1"),
                s: cell_style,
            },
            "n" => Cell::CellFormulaNumber {
                f: formula_index,
                v: cell_value.unwrap_or("0").parse::<f64>().unwrap_or(0.0),
                s: cell_style,
            },
            "e" => Cell::CellFormulaError {
                f: formula_index,
                ei: get_error_by_english_name(cell_value.unwrap_or("#ERROR!"))
                    .unwrap_or(Error::ERROR),
                s: cell_style,
                o: format!("{}!{}", sheet_name, cell_ref),
                m: cell_value.unwrap_or("#ERROR!").to_string(),
            },
            "s" => {
                // Not implemented
                let o = format!("{}!{}", sheet_name, cell_ref);
                let m = Error::NIMPL.to_string();
                println!("Invalid type (s) in {}!{}", sheet_name, cell_ref);
                Cell::CellFormulaError {
                    f: formula_index,
                    ei: Error::NIMPL,
                    s: cell_style,
                    o,
                    m,
                }
            }
            "str" => {
                // In Excel and in EqualTo all strings in cells result of a formula are *not* shared strings.
                Cell::CellFormulaString {
                    f: formula_index,
                    v: cell_value.unwrap_or("").to_string(),
                    s: cell_style,
                }
            }
            "d" => {
                // Not implemented
                println!("Invalid type (d) in {}!{}", sheet_name, cell_ref);
                let o = format!("{}!{}", sheet_name, cell_ref);
                let m = Error::NIMPL.to_string();
                Cell::CellFormulaError {
                    f: formula_index,
                    ei: Error::NIMPL,
                    s: cell_style,
                    o,
                    m,
                }
            }
            "inlineStr" => {
                // Not implemented
                let o = format!("{}!{}", sheet_name, cell_ref);
                let m = Error::NIMPL.to_string();
                println!("Invalid type (inlineStr) in {}!{}", sheet_name, cell_ref);
                Cell::CellFormulaError {
                    f: formula_index,
                    ei: Error::NIMPL,
                    s: cell_style,
                    o,
                    m,
                }
            }
            _ => {
                // error
                println!(
                    "Unexpected type ({}) in {}!{}",
                    cell_type, sheet_name, cell_ref
                );
                let o = format!("{}!{}", sheet_name, cell_ref);
                let m = Error::ERROR.to_string();
                Cell::CellFormulaError {
                    f: formula_index,
                    ei: Error::ERROR,
                    s: cell_style,
                    o,
                    m,
                }
            }
        }
    }
}

fn load_sheet_rels<R: Read + std::io::Seek>(
    archive: &mut zip::read::ZipArchive<R>,
    path: &str,
) -> Result<Vec<Comment>, XlsxError> {
    // ...xl/worksheets/sheet6.xml -> xl/worksheets/_rels/sheet6.xml.rels
    let mut comments = Vec::new();
    let v: Vec<&str> = path.split("/worksheets/").collect();
    let mut path = v[0].to_string();
    path.push_str("/worksheets/_rels/");
    path.push_str(v[1]);
    path.push_str(".rels");
    let file = archive.by_name(&path);
    if file.is_err() {
        return Ok(vec![]);
    }
    let mut text = String::new();
    file.unwrap().read_to_string(&mut text)?;
    let doc = roxmltree::Document::parse(&text)?;

    let rels = doc
        .root()
        .first_child()
        .ok_or_else(|| XlsxError::Xml("Corrupt XML structure".to_string()))?
        .children()
        .collect::<Vec<Node>>();
    for rel in rels {
        let t = get_attribute(&rel, "Type")?.to_string();
        if t.ends_with("comments") {
            let mut target = get_attribute(&rel, "Target")?.to_string();
            // Target="../comments1.xlsx"
            target.replace_range(..2, v[0]);
            comments = load_comments(archive, &target)?;
            break;
        }
    }
    Ok(comments)
}

pub(super) fn load_sheet<R: Read + std::io::Seek>(
    archive: &mut zip::read::ZipArchive<R>,
    path: &str,
    sheet_name: &str,
    sheet_id: u32,
    state: &SheetState,
    worksheets: &[String],
) -> Result<Worksheet, XlsxError> {
    let comments = load_sheet_rels(archive, path)?;
    let mut file = archive.by_name(path)?;
    let mut text = String::new();
    file.read_to_string(&mut text)?;
    let doc = roxmltree::Document::parse(&text)?;
    let ws = doc
        .root()
        .first_child()
        .ok_or_else(|| XlsxError::Xml("Corrupt XML structure".to_string()))?;
    let mut shared_formulas = Vec::new();

    let dimension = load_dimension(ws);

    // <sheetViews>
    //   <sheetView workbookViewId="0">
    //     <selection activeCell="E10" sqref="E10"/>
    //   </sheetView>
    // </sheetViews>
    // <sheetFormatPr defaultRowHeight="14.5" x14ac:dyDescent="0.35"/>

    // If we have frozen rows and columns:

    // <sheetView tabSelected="1" workbookViewId="0">
    //   <pane xSplit="3" ySplit="2" topLeftCell="D3" activePane="bottomRight" state="frozen"/>
    //   <selection pane="topRight" activeCell="D1" sqref="D1"/>
    //   <selection pane="bottomLeft" activeCell="A3" sqref="A3"/>
    //   <selection pane="bottomRight" activeCell="K16" sqref="K16"/>
    // </sheetView>

    // 18.18.52 ST_Pane (Pane Types)
    // bottomLeft, bottomRight, topLeft, topRight

    // NB: bottomLeft is used when only rows are frozen, etc
    // Calc ignores all those.

    let mut frozen_rows = 0;
    let mut frozen_columns = 0;

    // In Calc there can only be one sheetView
    let sheet_view = ws
        .children()
        .filter(|n| n.has_tag_name("sheetViews"))
        .collect::<Vec<Node>>()[0]
        .children()
        .filter(|n| n.has_tag_name("sheetView"))
        .collect::<Vec<Node>>()[0];
    let pane = sheet_view
        .children()
        .filter(|n| n.has_tag_name("pane"))
        .collect::<Vec<Node>>();

    // 18.18.53 ST_PaneState (Pane State)
    // frozen, frozenSplit, split
    if pane.len() == 1 && pane[0].attribute("state").unwrap_or("split") == "frozen" {
        // TODO: Should we assert that topLeft is consistent?
        // let top_left_cell = pane[0].attribute("topLeftCell").unwrap_or("A1").to_string();

        frozen_columns = get_number(pane[0], "xSplit");
        frozen_rows = get_number(pane[0], "ySplit");
    }

    let cols = load_columns(ws)?;
    let color = load_sheet_color(ws)?;

    // sheetData
    // <row r="1" spans="1:15" x14ac:dyDescent="0.35">
    //     <c r="A1" t="s">
    //         <v>0</v>
    //     </c>
    //     <c r="D1">
    //         <f>C1+1</f>
    //     </c>
    // </row>

    // holds the row heights
    let mut rows = Vec::new();
    let mut sheet_data = SheetData::new();
    let sheet_data_nodes = ws
        .children()
        .filter(|n| n.has_tag_name("sheetData"))
        .collect::<Vec<Node>>()[0];

    let default_row_height = 14.5;

    // holds a map from the formula index in Excel to the index in EqualTo
    let mut index_map = HashMap::new();
    for row in sheet_data_nodes.children() {
        // This is the row number 1-indexed
        let row_index = get_attribute(&row, "r")?.parse::<i32>()?;
        // `spans` is not used in EqualTo at the moment (it's an optimization)
        // let spans = row.attribute("spans");
        // This is the height of the row
        let has_height_attribute;
        let height = match row.attribute("ht") {
            Some(s) => {
                has_height_attribute = true;
                s.parse::<f64>().unwrap_or(default_row_height)
            }
            None => {
                has_height_attribute = false;
                default_row_height
            }
        };
        let custom_height = matches!(row.attribute("customHeight"), Some("1"));
        // The height of the row is always the visible height of the row
        // If custom_height is false that means the height was calculated automatically:
        // for example because a cell has many lines or a larger font

        let row_style = match row.attribute("s") {
            Some(s) => s.parse::<i32>().unwrap_or(0),
            None => 0,
        };
        let custom_format = matches!(row.attribute("customFormat"), Some("1"));
        if custom_height || custom_format || row_style != 0 || has_height_attribute {
            rows.push(Row {
                r: row_index,
                height,
                s: row_style,
                custom_height,
                custom_format,
            });
        }

        // Unused attributes:
        // * thickBot, thickTop, ph, collapsed, outlineLevel

        let mut data_row = HashMap::new();

        // 18.3.1.4 c (Cell)
        // Child Elements:
        // * v: Cell value
        // * is: Rich Text Inline (not used in EqualTo)
        // * f: Formula
        // Attributes:
        // r: reference. A1 style
        // s: style index
        // t: cell type
        // Unused attributes
        // cm (cell metadata), ph (Show Phonetic), vm (value metadata)
        for cell in row.children() {
            let cell_ref = get_attribute(&cell, "r")?;
            let column_letter = get_column_from_ref(cell_ref);
            let column = column_to_number(column_letter.as_str());

            // We check the value "v" child.
            let vs: Vec<Node> = cell.children().filter(|n| n.has_tag_name("v")).collect();
            let cell_value = if vs.len() == 1 {
                Some(vs[0].text().unwrap_or(""))
            } else {
                None
            };

            // type, the default type being "n" for number
            // If the cell does not have a value is an empty cell
            let cell_type = match cell.attribute("t") {
                Some(t) => t,
                None => {
                    if cell_value.is_none() {
                        "empty"
                    } else {
                        "n"
                    }
                }
            };

            // style index, the default style is 0
            let cell_style = match cell.attribute("s") {
                Some(s) => s.parse::<i32>().unwrap_or(0),
                None => 0,
            };

            // Check for formula
            // In Excel some formulas are shared and some are not, but in EqualTo all formulas are shared
            // A cell with a "non-shared" formula is like:
            // <c r="E3">
            //   <f>C2+1</f>
            //   <v>3</v>
            // </c>
            // A cell with a shared formula will be either a "mother" cell:
            // <c r="D2">
            //   <f t="shared" ref="D2:D3" si="0">C2+1</f>
            //   <v>3</v>
            // </c>
            // Or a "daughter" cell:
            // <c r="D3">
            //   <f t="shared" si="0"/>
            //   <v>4</v>
            // </c>
            // In EqualTo two cells have the same formula iff the R1C1 representation is the same
            // TODO: This algorithm could end up with "repeated" shared formulas
            //       We could solve that with a second transversal.
            let fs: Vec<Node> = cell.children().filter(|n| n.has_tag_name("f")).collect();
            let mut formula_index = -1;
            if fs.len() == 1 {
                // formula types:
                // 18.18.6 ST_CellFormulaType (Formula Type)
                // array (Array Formula) Formula is an array formula.
                // dataTable (Table Formula) Formula is a data table formula.
                // normal (Normal) Formula is a regular cell formula. (Default)
                // shared (Shared Formula) Formula is part of a shared formula.
                let formula_type = fs[0].attribute("t").unwrap_or("normal");
                match formula_type {
                    "shared" => {
                        // We have a shared formula
                        let si = get_attribute(&fs[0], "si")?;
                        let si = si.parse::<i32>()?;
                        match fs[0].attribute("ref") {
                            Some(_) => {
                                // It's the mother cell. We do not use the ref attribute in EqualTo
                                let formula = fs[0].text().unwrap_or("").to_string();
                                let context = format!("{}!{}", sheet_name, cell_ref);
                                let formula = from_a1_to_rc(formula, worksheets, context)?;
                                match index_map.get(&si) {
                                    Some(index) => {
                                        // The index for that formula already exists meaning we bumped into a daughter cell first
                                        // TODO: Worth assert the content is a placeholder?
                                        formula_index = *index;
                                        shared_formulas.insert(formula_index as usize, formula);
                                    }
                                    None => {
                                        // We haven't met any of the daughter cells
                                        match get_formula_index(&formula, &shared_formulas) {
                                            // The formula is already present, use that index
                                            Some(index) => {
                                                formula_index = index;
                                            }
                                            None => {
                                                shared_formulas.push(formula);
                                                formula_index = shared_formulas.len() as i32 - 1;
                                            }
                                        };
                                        index_map.insert(si, formula_index);
                                    }
                                }
                            }
                            None => {
                                // It's a daughter cell
                                match index_map.get(&si) {
                                    Some(index) => {
                                        formula_index = *index;
                                    }
                                    None => {
                                        // Haven't bumped into the mother cell yet. We insert a placeholder.
                                        // Note that it is perfectly possible that the formula of the mother cell
                                        // is already in the set of array formulas. This will lead to the above mention duplicity.
                                        // This is not a problem
                                        let placeholder = "".to_string();
                                        shared_formulas.push(placeholder);
                                        formula_index = shared_formulas.len() as i32 - 1;
                                        index_map.insert(si, formula_index);
                                    }
                                }
                            }
                        }
                    }
                    "array" => {
                        return Err(XlsxError::Workbook(
                            "Array formulas are not supported.".to_string(),
                        ));
                    }
                    "dataTable" => {
                        return Err(XlsxError::Workbook(
                            "Data table formulas are not supported.".to_string(),
                        ));
                    }
                    "normal" => {
                        // Its a cell with a simple formula
                        let formula = fs[0].text().unwrap_or("").to_string();
                        let context = format!("{}!{}", sheet_name, cell_ref);
                        let formula = from_a1_to_rc(formula, worksheets, context)?;
                        match get_formula_index(&formula, &shared_formulas) {
                            Some(index) => formula_index = index,
                            None => {
                                shared_formulas.push(formula);
                                formula_index = shared_formulas.len() as i32 - 1;
                            }
                        }
                    }
                    _ => {
                        return Err(XlsxError::Workbook(format!(
                            "Invalid formula type {:?}.",
                            formula_type,
                        )));
                    }
                }
            }
            let cell = get_cell_from_excel(
                cell_value,
                cell_type,
                cell_style,
                formula_index,
                sheet_name,
                cell_ref,
            );
            data_row.insert(column, cell);
        }
        sheet_data.insert(row_index, data_row);
    }

    let merge_cells = load_merge_cells(ws)?;

    // Conditional Formatting
    // <conditionalFormatting sqref="B1:B9">
    //     <cfRule type="colorScale" priority="1">
    //         <colorScale>
    //             <cfvo type="min"/>
    //             <cfvo type="max"/>
    //             <color rgb="FFF8696B"/>
    //             <color rgb="FFFCFCFF"/>
    //         </colorScale>
    //     </cfRule>
    // </conditionalFormatting>
    // pageSetup
    // <pageSetup orientation="portrait" r:id="rId1"/>

    Ok(Worksheet {
        dimension,
        cols,
        rows,
        shared_formulas,
        sheet_data,
        name: sheet_name.to_string(),
        sheet_id,
        state: state.to_owned(),
        color,
        merge_cells,
        comments,
        frozen_rows,
        frozen_columns,
    })
}

pub(super) fn load_sheets<R: Read + std::io::Seek>(
    archive: &mut zip::read::ZipArchive<R>,
    rels: &HashMap<String, Relationship>,
    workbook: &WorkbookXML,
) -> Result<Vec<Worksheet>, XlsxError> {
    let worksheets = &workbook.worksheets;
    let ws_list: Vec<String> = worksheets.iter().map(|s| s.name.clone()).collect();
    let mut sheets = Vec::new();
    for sheet in &workbook.worksheets {
        let name = sheet.name.clone();
        let rel_id = sheet.id.clone();
        let state = sheet.state.clone();
        let rel = &rels[&rel_id];
        if rel.rel_type.ends_with("worksheet") {
            let path = &rel.target;
            // let sheet_index = get_index_from_rel_id(rel_id.as_str());
            sheets.push(load_sheet(
                archive,
                &("xl/".to_owned() + path),
                &name,
                sheet.sheet_id, //&rel_id,
                &state,
                &ws_list,
            )?);
        }
    }
    Ok(sheets)
}
