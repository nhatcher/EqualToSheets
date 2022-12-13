//! This cate reads an Excel file and transforms it into an internal representation ([`Model`]).
//! An `xlsx` is a zip file containing a set of folders and `xml` files. The EqualTo json structure mimics the relevant parts of the Excel zip.
//! Although the xlsx structure is quite complicated, it's essentials regarding the spreadsheet technology are easier to grasp.
//!
//! The simplest workbook folder structure might look like this:
//!
//! ```text
//! docProps
//!     app.xml
//!     core.xml
//!
//! _rels
//!     .rels
//!
//! xl
//!     _rels
//!         workbook.xml.rels
//!     theme
//!         theme1.xml
//!     worksheets
//!         sheet1.xml
//!     calcChain.xml
//!     styles.xml
//!     workbook.xml
//!     sharedStrings.xml
//!
//! [Content_Types].xml
//! ```
//!
//! Note that more complicated workbooks will have many more files and folders.
//! For instance charts, pivot tables, comments, tables,...
//!
//! The relevant json structure in EqualTo will be:
//!
//! ```json
//! {
//!     "name": "Workbook1",
//!     "defined_names": [],
//!     "shared_strings": [],
//!     "worksheets": [],
//!     "styles": {
//!         "num_fmts": [],
//!         "fonts": [],
//!         "fills": [],
//!         "borders": [],
//!         "cell_style_xfs": [],
//!         "cell_styles" : [],
//!         "cell_xfs": []
//!     }
//! }
//! ```
//!
//! Note that there is not a 1-1 correspondence but there is a close resemblance.
//!
//! [`Model`]: ../equalto_xlsx/struct.Model.html

use equalto_calc::expressions::{
    parser::{stringify::to_rc_format, Parser},
    token::{get_error_by_english_name, Error},
};
use equalto_calc::expressions::{types::CellReferenceRC, utils::column_to_number};
use equalto_calc::types::*;

use roxmltree::Node;
use serde::{Deserialize, Serialize};

use std::fs;
use std::{
    collections::HashMap,
    io::{BufReader, Read, Write},
};

mod colors;
pub mod compare;
mod shared_strings;
mod types;

use crate::colors::{get_indexed_color, get_themed_color};
use crate::shared_strings::read_shared_strings;
use crate::types::ExcelArchive;

#[cfg(test)]
mod test;

#[derive(Serialize, Deserialize, Debug)]
struct Sheet {
    name: String,
    sheet_id: i32,
    id: String,
    state: String,
}

#[derive(Serialize, Deserialize, Debug)]
struct WorkbookXML {
    worksheets: Vec<Sheet>,
    defined_names: Vec<DefinedName>,
}

#[derive(Serialize, Deserialize, Debug)]
struct Relationship {
    target: String,
    rel_type: String,
}

// Private methods

fn get_color(node: Node) -> Color {
    // 18.3.1.15 color (Data Bar Color)
    if node.has_attribute("rgb") {
        let mut val = node.attribute("rgb").unwrap().to_string();
        // FIXME the two first values is normally the alpha.
        if val.len() == 8 {
            val = format!("#{}", &val[2..8]);
        }
        Color::RGB(val)
    } else if node.has_attribute("indexed") {
        let index = node.attribute("indexed").unwrap().parse::<i32>().unwrap();
        let rgb = get_indexed_color(index);
        Color::RGB(rgb)
    // Color::Indexed(val)
    } else if node.has_attribute("theme") {
        let theme = node.attribute("theme").unwrap().parse::<i32>().unwrap();
        let tint = match node.attribute("tint") {
            Some(t) => t.parse::<f64>().unwrap_or(0.0),
            None => 0.0,
        };
        let rgb = get_themed_color(theme, tint);
        Color::RGB(rgb)
    // Color::Theme { theme, tint }
    } else if node.has_attribute("auto") {
        // TODO: Is this correct?
        // A boolean value indicating the color is automatic and system color dependent.
        Color::None
    } else {
        println!("Unexpected color node {:?}", node);
        Color::None
    }
}

fn get_border(node: Node, name: &str) -> BorderItem {
    let style;
    let color;
    let border_nodes = node
        .children()
        .filter(|n| n.has_tag_name(name))
        .collect::<Vec<Node>>();
    if border_nodes.len() == 1 {
        let border = border_nodes[0];
        style = match border.attribute("s") {
            Some(s) => s.to_string(),
            None => "none".to_string(),
        };

        let color_node = border
            .children()
            .filter(|n| n.has_tag_name("color"))
            .collect::<Vec<Node>>();
        if color_node.len() == 1 {
            color = get_color(color_node[0]);
        } else {
            color = Color::None;
        }
    } else {
        style = "none".to_string();
        color = Color::None;
    }
    BorderItem { style, color }
}

fn get_bool(node: Node, s: &str) -> bool {
    // defaults to true
    !matches!(node.attribute(s), Some("0"))
}

fn get_bool_false(node: Node, s: &str) -> bool {
    // defaults to false
    matches!(node.attribute(s), Some("1"))
}

fn get_number(node: Node, s: &str) -> i32 {
    node.attribute(s).unwrap_or("0").parse::<i32>().unwrap_or(0)
}

fn load_styles(archive: &mut ExcelArchive) -> Styles {
    let mut file = archive.by_name("xl/styles.xml").unwrap();
    let mut text = String::new();
    file.read_to_string(&mut text).unwrap();
    let doc = roxmltree::Document::parse(&text).unwrap();
    let style_sheet = doc.root().first_child().unwrap();

    let mut num_fmts = Vec::new();
    let num_fmts_nodes = style_sheet
        .children()
        .filter(|n| n.has_tag_name("numFmts"))
        .collect::<Vec<Node>>();
    if num_fmts_nodes.len() == 1 {
        for num_fmt in num_fmts_nodes[0].children() {
            let num_fmt_id = get_number(num_fmt, "numFmtId");
            let format_code = num_fmt.attribute("formatCode").unwrap_or("").to_string();
            num_fmts.push(NumFmt {
                num_fmt_id,
                format_code,
            });
        }
    }

    let mut fonts = Vec::new();
    let font_nodes = style_sheet
        .children()
        .filter(|n| n.has_tag_name("fonts"))
        .collect::<Vec<Node>>()[0];
    for font in font_nodes.children() {
        let mut sz = 11;
        let mut name = "Calibri".to_string();
        let mut family = 2;
        let mut scheme = "minor".to_string();
        // NOTE: In Excel you can have simple underline or double underline
        // In EqualTo convert double underline to simple
        // This in excel is u with a value of "double"
        let mut u = false;
        let mut b = false;
        let mut i = false;
        let mut strike = false;
        let mut color = Color::RGB("FFFFFF00".to_string());
        for feature in font.children() {
            match feature.tag_name().name() {
                "sz" => {
                    sz = feature
                        .attribute("val")
                        .unwrap_or("11")
                        .parse::<i32>()
                        .unwrap_or(11);
                }
                "color" => {
                    color = get_color(feature);
                }
                "u" => {
                    u = true;
                }
                "b" => {
                    b = true;
                }
                "i" => {
                    i = true;
                }
                "strike" => {
                    strike = true;
                }
                "name" => name = feature.attribute("val").unwrap_or("Calibri").to_string(),
                "family" => {
                    family = feature
                        .attribute("val")
                        .unwrap_or("2")
                        .parse::<i32>()
                        .unwrap_or(2);
                }
                "scheme" => scheme = feature.attribute("val").unwrap_or("minor").to_string(),
                "charset" => {}
                _ => {
                    println!("Unexpected feature {:?}", feature);
                }
            }
        }
        fonts.push(Font {
            strike,
            u,
            b,
            i,
            sz,
            color,
            name,
            family,
            scheme,
        });
    }

    let mut fills = Vec::new();
    let fill_nodes = style_sheet
        .children()
        .filter(|n| n.has_tag_name("fills"))
        .collect::<Vec<Node>>()[0];
    for fill in fill_nodes.children() {
        let pattern_fill = fill
            .children()
            .filter(|n| n.has_tag_name("patternFill"))
            .collect::<Vec<Node>>()[0];
        let pattern_type = pattern_fill
            .attribute("patternType")
            .unwrap_or("none")
            .to_string();
        let mut fg_color = Color::None;
        let mut bg_color = Color::None;
        for feature in pattern_fill.children() {
            match feature.tag_name().name() {
                "fgColor" => {
                    fg_color = get_color(feature);
                }
                "bgColor" => {
                    bg_color = get_color(feature);
                }
                _ => {
                    println!("Unexpected pattern");
                    dbg!(feature);
                }
            }
        }
        fills.push(Fill {
            pattern_type,
            fg_color,
            bg_color,
        })
    }

    let mut borders = Vec::new();
    let border_nodes = style_sheet
        .children()
        .filter(|n| n.has_tag_name("borders"))
        .collect::<Vec<Node>>()[0];
    for border in border_nodes.children() {
        let diagonal_up = get_bool_false(border, "diagonal_up");
        let diagonal_down = get_bool_false(border, "diagonal_down");
        let left = get_border(border, "left");
        let right = get_border(border, "right");
        let top = get_border(border, "top");
        let bottom = get_border(border, "bottom");
        let diagonal = get_border(border, "diagonal");
        borders.push(Border {
            diagonal_up,
            diagonal_down,
            left,
            right,
            top,
            bottom,
            diagonal,
        });
    }

    let mut cell_style_xfs = Vec::new();
    let cell_style_xfs_nodes = style_sheet
        .children()
        .filter(|n| n.has_tag_name("cellStyleXfs"))
        .collect::<Vec<Node>>()[0];
    for xfs in cell_style_xfs_nodes.children() {
        let num_fmt_id = get_number(xfs, "numFmtId");
        let font_id = get_number(xfs, "fontId");
        let fill_id = get_number(xfs, "fillId");
        let border_id = get_number(xfs, "borderId");
        let apply_number_format = get_bool(xfs, "applyNumberFormat");
        let apply_border = get_bool(xfs, "applyBorder");
        let apply_alignment = get_bool(xfs, "applyAlignment");
        let apply_protection = get_bool(xfs, "applyProtection");
        let apply_font = get_bool(xfs, "applyFont");
        let apply_fill = get_bool(xfs, "applyFill");
        cell_style_xfs.push(CellStyleXfs {
            num_fmt_id,
            font_id,
            fill_id,
            border_id,
            apply_number_format,
            apply_border,
            apply_alignment,
            apply_protection,
            apply_font,
            apply_fill,
        });
    }

    let mut cell_styles = Vec::new();
    let mut style_names = HashMap::new();
    let cell_style_nodes = style_sheet
        .children()
        .filter(|n| n.has_tag_name("cellStyles"))
        .collect::<Vec<Node>>()[0];
    for cell_style in cell_style_nodes.children() {
        let name = cell_style.attribute("name").unwrap().to_string();
        let xf_id = get_number(cell_style, "xfId");
        let builtin_id = get_number(cell_style, "builtinId");
        style_names.insert(xf_id, name.clone());
        cell_styles.push(CellStyles {
            name,
            xf_id,
            builtin_id,
        })
    }

    let mut cell_xfs = Vec::new();
    let cell_xfs_nodes = style_sheet
        .children()
        .filter(|n| n.has_tag_name("cellXfs"))
        .collect::<Vec<Node>>()[0];
    for xfs in cell_xfs_nodes.children() {
        let xf_id = xfs.attribute("xfId").unwrap().parse::<i32>().unwrap();
        let num_fmt_id = get_number(xfs, "numFmtId");
        let font_id = get_number(xfs, "fontId");
        let fill_id = get_number(xfs, "fillId");
        let border_id = get_number(xfs, "borderId");
        let apply_number_format = get_bool_false(xfs, "applyNumberFormat");
        let apply_border = get_bool_false(xfs, "applyBorder");
        let apply_alignment = get_bool_false(xfs, "applyAlignment");
        let apply_protection = get_bool_false(xfs, "applyProtection");
        let apply_font = get_bool_false(xfs, "applyFont");
        let apply_fill = get_bool_false(xfs, "applyFill");
        let quote_prefix = get_bool_false(xfs, "quotePrefix");

        // TODO: Pivot Tables
        // let pivotButton = get_bool(xfs, "pivotButton");
        let mut horizontal_alignment = "default".to_string();
        let alignment_node = xfs
            .children()
            .filter(|n| n.has_tag_name("alignment"))
            .collect::<Vec<Node>>();
        if alignment_node.len() == 1 {
            horizontal_alignment = alignment_node[0]
                .attribute("horizontal")
                .unwrap_or("default")
                .to_string();
        }
        let style_name = style_names
            .get(&xf_id)
            .unwrap_or(&"".to_string())
            .to_lowercase();
        let read_only = style_name.contains("readonly") || style_name.contains("read_only");
        cell_xfs.push(CellXfs {
            xf_id,
            num_fmt_id,
            font_id,
            fill_id,
            border_id,
            horizontal_alignment,
            read_only,
            apply_number_format,
            apply_border,
            apply_alignment,
            apply_protection,
            apply_font,
            apply_fill,
            quote_prefix,
        });
    }

    // TODO
    // let mut dxfs = Vec::new();
    // let mut tableStyles = Vec::new();
    // let mut colors = Vec::new();
    // <colors>
    //     <mruColors>
    //         <color rgb="FFB1BB4D"/>
    //         <color rgb="FFFF99CC"/>
    //         <color rgb="FF6C56DC"/>
    //         <color rgb="FFFF66CC"/>
    //     </mruColors>
    // </colors>

    Styles {
        num_fmts,
        fonts,
        fills,
        borders,
        cell_style_xfs,
        cell_xfs,
        cell_styles,
    }
}

fn load_relationships(archive: &mut ExcelArchive) -> HashMap<String, Relationship> {
    let mut file = archive.by_name("xl/_rels/workbook.xml.rels").unwrap();
    let mut text = String::new();
    file.read_to_string(&mut text).unwrap();
    let doc = roxmltree::Document::parse(&text).unwrap();
    let nodes: Vec<Node> = doc
        .descendants()
        .filter(|n| n.has_tag_name("Relationship"))
        .collect();
    let mut rels = HashMap::new();
    for node in nodes {
        rels.insert(
            node.attribute("Id").unwrap().to_string(),
            Relationship {
                rel_type: node.attribute("Type").unwrap().to_string(),
                target: node.attribute("Target").unwrap().to_string(),
            },
        );
    }
    rels
}

fn load_workbook(archive: &mut ExcelArchive) -> WorkbookXML {
    let mut file = archive.by_name("xl/workbook.xml").unwrap();
    let mut text = String::new();
    file.read_to_string(&mut text).unwrap();
    let doc = roxmltree::Document::parse(&text).unwrap();
    let mut defined_names = Vec::new();
    let mut sheets = Vec::new();
    // Get the sheets
    let sheet_nodes: Vec<Node> = doc
        .descendants()
        .filter(|n| n.has_tag_name("sheet"))
        .collect();
    for sheet in sheet_nodes {
        let name = sheet.attribute("name").unwrap().to_string();
        let sheet_id = sheet.attribute("sheetId").unwrap().to_string();
        let sheet_id = sheet_id.parse::<i32>().unwrap();
        let id = sheet
            .attribute((
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                "id",
            ))
            .unwrap()
            .to_string();
        let state = match sheet.attribute("state") {
            Some(s) => s.to_string(),
            None => "visible".to_string(),
        };
        sheets.push(Sheet {
            name,
            sheet_id,
            id,
            state,
        });
    }
    // Get the defined names
    let name_nodes: Vec<Node> = doc
        .descendants()
        .filter(|n| n.has_tag_name("definedName"))
        .collect();
    for node in name_nodes {
        let name = node.attribute("name").unwrap().to_string();
        let formula = node.text().unwrap().to_string();
        let sheet_id = match node.attribute("localSheetId") {
            Some(s) => {
                let index = s.parse::<usize>().unwrap();
                Some(sheets[index].sheet_id)
            }
            None => None,
        };
        defined_names.push(DefinedName {
            name,
            formula,
            sheet_id,
        })
    }
    // read the relationships file
    WorkbookXML {
        worksheets: sheets,
        defined_names,
    }
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

fn load_comments(archive: &mut ExcelArchive, path: &str) -> Vec<Comment> {
    let mut comments = Vec::new();
    let mut file = archive.by_name(path).unwrap();
    let mut text = String::new();
    file.read_to_string(&mut text).unwrap();
    let doc = roxmltree::Document::parse(&text).unwrap();
    let ws = doc.root().first_child().unwrap();
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
            let cell_ref = comment.attribute("ref").unwrap().to_string();
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

    comments
}

fn load_sheet_rels(archive: &mut ExcelArchive, path: &str) -> Vec<Comment> {
    // ...xl/worksheets/sheet6.xml -> xl/worksheets/_rels/sheet6.xml.rels
    let mut comments = Vec::new();
    let v: Vec<&str> = path.split("/worksheets/").collect();
    let mut path = v[0].to_string();
    path.push_str("/worksheets/_rels/");
    path.push_str(v[1]);
    path.push_str(".rels");
    let file = archive.by_name(&path);
    if file.is_err() {
        return vec![];
    }
    let mut text = String::new();
    file.unwrap().read_to_string(&mut text).unwrap();
    let doc = roxmltree::Document::parse(&text).unwrap();

    let rels = doc
        .root()
        .first_child()
        .unwrap()
        .children()
        .collect::<Vec<Node>>();
    for rel in rels {
        let t = rel.attribute("Type").unwrap().to_string();
        if t.ends_with("comments") {
            let mut target = rel.attribute("Target").unwrap().to_string();
            // Target="../comments1.xlsx"
            target.replace_range(..2, v[0]);
            comments = load_comments(archive, &target);
            break;
        }
    }
    comments
}

fn get_formula_index(formula: &str, shared_formulas: &[String]) -> Option<i32> {
    for (index, f) in shared_formulas.iter().enumerate() {
        if f == formula {
            return Some(index as i32);
        }
    }
    None
}

fn load_columns(ws: Node) -> Vec<Col> {
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
            let min = col.attribute("min").unwrap();
            let min = min.parse::<i32>().unwrap();
            let max = col.attribute("max").unwrap();
            let max = max.parse::<i32>().unwrap();
            let width = col.attribute("width").unwrap();
            let width = width.parse::<f64>().unwrap();
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
    cols
}

fn load_merge_cells(ws: Node) -> Vec<String> {
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
            let reference = merge_cell.attribute("ref").unwrap().to_string();
            merge_cells.push(reference);
        }
    }
    merge_cells
}

fn load_sheet_color(ws: Node) -> Color {
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
            color = get_color(tabs[0]);
        }
    }
    color
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
                t: "b".to_string(),
                v: cell_value == Some("1"),
                s: cell_style,
            },
            "n" => Cell::NumberCell {
                t: "n".to_string(),
                v: cell_value.unwrap_or("0").parse::<f64>().unwrap_or(0.0),
                s: cell_style,
            },
            "e" => Cell::ErrorCell {
                t: "e".to_string(),
                ei: get_error_by_english_name(cell_value.unwrap_or("#ERROR!"))
                    .unwrap_or(Error::ERROR),
                s: cell_style,
            },
            "s" => Cell::SharedString {
                t: "s".to_string(),
                si: cell_value.unwrap_or("0").parse::<i32>().unwrap_or(0),
                s: cell_style,
            },
            "str" => {
                // We are assuming that all strings in cells without a formula in Excel are shared strings.
                // Not implemented
                println!("Invalid type (str) in {}!{}", sheet_name, cell_ref);
                Cell::ErrorCell {
                    t: "e".to_string(),
                    ei: Error::NIMPL,
                    s: cell_style,
                }
            }
            "d" => {
                // Not implemented
                println!("Invalid type (d) in {}!{}", sheet_name, cell_ref);
                Cell::ErrorCell {
                    t: "e".to_string(),
                    ei: Error::NIMPL,
                    s: cell_style,
                }
            }
            "inlineStr" => {
                // Not implemented
                println!("Invalid type (inlineStr) in {}!{}", sheet_name, cell_ref);
                Cell::ErrorCell {
                    t: "e".to_string(),
                    ei: Error::NIMPL,
                    s: cell_style,
                }
            }
            "empty" => Cell::EmptyCell {
                t: "empty".to_string(),
                s: cell_style,
            },
            _ => {
                // error
                println!(
                    "Unexpected type ({}) in {}!{}",
                    cell_type, sheet_name, cell_ref
                );
                Cell::ErrorCell {
                    t: "e".to_string(),
                    ei: Error::ERROR,
                    s: cell_style,
                }
            }
        }
    } else {
        match cell_type {
            "b" => Cell::CellFormulaBoolean {
                t: "b".to_string(),
                f: formula_index,
                v: cell_value == Some("1"),
                s: cell_style,
            },
            "n" => Cell::CellFormulaNumber {
                t: "n".to_string(),
                f: formula_index,
                v: cell_value.unwrap_or("0").parse::<f64>().unwrap_or(0.0),
                s: cell_style,
            },
            "e" => Cell::CellFormulaError {
                t: "e".to_string(),
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
                    t: "e".to_string(),
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
                    t: "str".to_string(),
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
                    t: "e".to_string(),
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
                    t: "e".to_string(),
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
                    t: "e".to_string(),
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

fn load_dimension(ws: Node) -> String {
    // <dimension ref="A1:O18"/>
    let dimension_nodes = ws
        .children()
        .filter(|n| n.has_tag_name("dimension"))
        .collect::<Vec<Node>>();
    if dimension_nodes.len() == 1 {
        dimension_nodes[0]
            .attribute("ref")
            .unwrap_or("A1")
            .to_string()
    } else {
        "A1".to_string()
    }
}

fn load_sheet(
    archive: &mut ExcelArchive,
    path: &str,
    sheet_name: &str,
    sheet_id: i32,
    state: &str,
    worksheets: &[String],
) -> Worksheet {
    let comments = load_sheet_rels(archive, path);
    let mut file = archive.by_name(path).unwrap();
    let mut text = String::new();
    file.read_to_string(&mut text).unwrap();
    let doc = roxmltree::Document::parse(&text).unwrap();
    let ws = doc.root().first_child().unwrap();
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

    let cols = load_columns(ws);
    let color = load_sheet_color(ws);

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
        let row_index = row.attribute("r").unwrap().parse::<i32>().unwrap();
        // `spans` is not used in EqualTo at the moment (it's an optimization)
        // let spans = row.attribute("spans");
        // This is the height of the row
        let height = match row.attribute("ht") {
            Some(s) => s.parse::<f64>().unwrap_or(default_row_height),
            None => default_row_height,
        };
        let custom_height = matches!(row.attribute("customHeight"), Some("1"));

        let row_style = match row.attribute("s") {
            Some(s) => s.parse::<i32>().unwrap_or(0),
            None => 0,
        };
        let custom_format = matches!(row.attribute("customFormat"), Some("1"));
        if custom_height || custom_format {
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
            let cell_ref = cell.attribute("r").unwrap();
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
                        let si = fs[0].attribute("si").unwrap();
                        let si = si.parse::<i32>().unwrap();
                        match fs[0].attribute("ref") {
                            Some(_) => {
                                // It's the mother cell. We do not use the ref attribute in EqualTo
                                let formula = fs[0].text().unwrap().to_string();
                                let context = format!("{}!{}", sheet_name, cell_ref);
                                let formula = from_a1_to_rc(formula, worksheets, context);
                                match index_map.get(&si) {
                                    Some(index) => {
                                        // The index for that formula already exists meaning we bumped into a daughter cell first
                                        // TODO: Worth assert the content is a placeholder?
                                        formula_index = *index as i32;
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
                                        formula_index = *index as i32;
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
                        panic!("Array formulas are not supported at the moment. Aborting");
                    }
                    "dataTable" => {
                        panic!("Data table formulas are not supported. Aborting");
                    }
                    "normal" => {
                        // Its a cell with a simple formula
                        let formula = fs[0].text().unwrap().to_string();
                        let context = format!("{}!{}", sheet_name, cell_ref);
                        let formula = from_a1_to_rc(formula, worksheets, context);
                        match get_formula_index(&formula, &shared_formulas) {
                            Some(index) => formula_index = index,
                            None => {
                                shared_formulas.push(formula);
                                formula_index = shared_formulas.len() as i32 - 1;
                            }
                        }
                    }
                    _ => {
                        panic!("Invalid formula type {:?}. Aborting", formula_type);
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

    let merge_cells = load_merge_cells(ws);

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

    Worksheet {
        dimension,
        cols,
        rows,
        shared_formulas,
        sheet_data,
        name: sheet_name.to_string(),
        sheet_id,
        state: state.to_string(),
        color,
        merge_cells,
        comments,
        frozen_rows,
        frozen_columns,
    }
}

fn load_sheets(
    archive: &mut ExcelArchive,
    rels: &HashMap<String, Relationship>,
    workbook: &WorkbookXML,
) -> Vec<Worksheet> {
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
            ));
        }
    }
    sheets
}

fn from_a1_to_rc(formula: String, worksheets: &[String], context: String) -> String {
    let mut parser = Parser::new(worksheets.to_owned());
    let cell_reference = parse_reference(&context);
    let t = parser.parse(&formula, &Some(cell_reference));
    to_rc_format(&t)
}

// This parses Sheet1!AS23 into sheet, column and row
// FIXME: This is buggy. Does not check that is a valid sheet name or column
// There is a similar named function in equalto_calc. We probably should fix both at the same time.
// NB: Maybe use regexes for this?
fn parse_reference(s: &str) -> CellReferenceRC {
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
    CellReferenceRC {
        sheet: sheet_name,
        row: row.parse::<i32>().unwrap(),
        column: column_to_number(&column),
    }
}
// Public methods

/// Imports a file from disk into an internal representation
pub fn load_from_excel(file_name: &str, locale: &str, tz: &str, wb_type: WorkbookType) -> Workbook {
    let file_path = std::path::Path::new(file_name);
    let file = fs::File::open(file_path).unwrap();
    let reader = BufReader::new(file);
    let name = file_path.file_stem().unwrap().to_string_lossy().to_string();

    let mut archive = zip::ZipArchive::new(reader).unwrap();

    let shared_strings = read_shared_strings(&mut archive);
    let workbook = load_workbook(&mut archive);
    let rels = load_relationships(&mut archive);

    let worksheets = load_sheets(&mut archive, &rels, &workbook);
    let styles = load_styles(&mut archive);
    Workbook {
        shared_strings,
        defined_names: workbook.defined_names,
        worksheets,
        styles,
        name,
        settings: WorkbookSettings {
            tz: tz.to_string(),
            locale: locale.to_string(),
        },
        wb_type,
    }
}

/// Exports an internal representation of a workbook into an equivalent EqualTo json format
pub fn save_to_json(model: Workbook, output: &str) {
    let s = serde_json::to_string(&model).unwrap();
    let file_path = std::path::Path::new(output);
    let mut file = fs::File::create(file_path).unwrap();
    file.write_all(s.as_bytes()).unwrap();
}
