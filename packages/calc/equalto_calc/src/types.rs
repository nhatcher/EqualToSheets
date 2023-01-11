use serde::{Deserialize, Serialize};
use std::collections::HashMap;

use crate::expressions::token::Error;

// Useful for `#[serde(default = "default_as_true")]`
fn default_as_true() -> bool {
    true
}

fn default_as_false() -> bool {
    false
}

fn default_as_default() -> String {
    "default".to_string()
}

// Useful for `#[serde(skip_serializing_if = "is_true")]`
fn is_true(b: &bool) -> bool {
    *b
}

fn is_false(b: &bool) -> bool {
    !*b
}

fn is_default(s: &str) -> bool {
    s == "default"
}

fn is_zero(num: &i32) -> bool {
    *num == 0
}

/// Information need to show a sheet tab in the UI
/// The color is serialized only if it is not Color::None
#[derive(Serialize, Deserialize, Debug, PartialEq)]
pub struct Tab {
    pub name: String,
    pub state: String,
    pub index: i32,
    pub sheet_id: i32,
    #[serde(default = "Color::new")]
    #[serde(skip_serializing_if = "Color::is_none")]
    pub color: Color,
}

#[derive(Serialize, Deserialize, Debug, PartialEq, Eq, Clone)]
pub struct WorkbookSettings {
    pub tz: String,
    pub locale: String,
}
/// An internal representation of an EqualTo Workbook
#[derive(Serialize, Deserialize, Debug, PartialEq, Clone)]
#[serde(deny_unknown_fields)]
pub struct Workbook {
    pub shared_strings: Vec<String>,
    pub defined_names: Vec<DefinedName>,
    pub worksheets: Vec<Worksheet>,
    pub styles: Styles,
    pub name: String,
    pub settings: WorkbookSettings,
}

/// A defined name. The `sheet_id` is the sheet index in case the name is local
#[derive(Serialize, Deserialize, Debug, PartialEq, Eq, Clone)]
pub struct DefinedName {
    pub name: String,
    pub formula: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sheet_id: Option<i32>,
}

// TODO: Move to worksheet.rs make frozen_rows/columns private and u32
/// Internal representation of a worksheet Excel object
/// * state:
///    18.18.68 ST_SheetState (Sheet Visibility Types)
///    hidden, veryHidden, visible
// TODO: Maybe use an enum for the state values here?
#[derive(Serialize, Deserialize, Debug, PartialEq, Clone)]
pub struct Worksheet {
    pub dimension: String,
    pub cols: Vec<Col>,
    pub rows: Vec<Row>,
    pub name: String,
    pub sheet_data: SheetData,
    pub shared_formulas: Vec<String>,
    pub sheet_id: i32,
    pub state: String,
    #[serde(default = "Color::new")]
    #[serde(skip_serializing_if = "Color::is_none")]
    pub color: Color,
    pub merge_cells: Vec<String>,
    pub comments: Vec<Comment>,
    #[serde(default)]
    #[serde(skip_serializing_if = "is_zero")]
    pub frozen_rows: i32,
    #[serde(default)]
    #[serde(skip_serializing_if = "is_zero")]
    pub frozen_columns: i32,
}

/// Internal representation of Excel's sheet_data
/// It is row first and because of this all of our API's should be row first
pub type SheetData = HashMap<i32, HashMap<i32, Cell>>;

#[derive(Serialize, Deserialize, Debug, PartialEq, Clone)]
pub struct Row {
    pub height: f64,
    pub r: i32,
    pub custom_format: bool,
    pub custom_height: bool,
    pub s: i32,
}

#[derive(Serialize, Deserialize, Debug, PartialEq, Clone)]
pub struct Col {
    pub min: i32,
    pub max: i32,
    pub width: f64,
    pub custom_width: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style: Option<i32>,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
pub enum Color {
    None,
    RGB(String),
    Indexed(i32),
    Theme { theme: i32, tint: f64 },
}

impl Color {
    pub fn is_none(&self) -> bool {
        matches!(self, Color::None)
    }
    pub fn new() -> Color {
        Color::None
    }
}

impl Default for Color {
    fn default() -> Self {
        Self::new()
    }
}

/// Cell type enum matching Excel TYPE() function values.
#[derive(Debug, Eq, PartialEq)]
pub enum CellType {
    Number = 1,
    Text = 2,
    LogicalValue = 4,
    ErrorValue = 16,
    Array = 64,
    CompoundData = 128,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(untagged, deny_unknown_fields)]
pub enum Cell {
    EmptyCell {
        t: String,
        s: i32,
    },
    BooleanCell {
        t: String,
        v: bool,
        s: i32,
    },
    NumberCell {
        t: String,
        v: f64,
        s: i32,
    },
    ErrorCell {
        t: String,
        ei: Error,
        s: i32,
    },
    // Always a shared string
    SharedString {
        t: String,
        si: i32,
        s: i32,
    },
    // Non evaluated Formula
    CellFormula {
        t: String,
        f: i32,
        s: i32,
    },
    CellFormulaBoolean {
        t: String,
        f: i32,
        v: bool,
        s: i32,
    },
    CellFormulaNumber {
        t: String,
        f: i32,
        v: f64,
        s: i32,
    },
    // always inline string
    CellFormulaString {
        t: String,
        f: i32,
        v: String,
        s: i32,
    },
    CellFormulaError {
        t: String,
        f: i32,
        ei: Error,
        s: i32,
        // Origin: Sheet3!C4
        o: String,
        // Error Message: "Not implemented function"
        m: String,
    },
    // TODO: Array formulas
}

#[derive(Serialize, Deserialize, Debug, PartialEq, Eq, Clone)]
pub struct Comment {
    pub text: String,
    pub author_name: String,
    pub author_id: Option<String>,
    pub cell_ref: String,
}
#[derive(Serialize, Deserialize, Debug, PartialEq, Clone)]
pub struct Styles {
    pub num_fmts: Vec<NumFmt>,
    pub fonts: Vec<Font>,
    pub fills: Vec<Fill>,
    pub borders: Vec<Border>,
    pub cell_style_xfs: Vec<CellStyleXfs>,
    pub cell_xfs: Vec<CellXfs>,
    pub cell_styles: Vec<CellStyles>,
}

impl Default for Styles {
    fn default() -> Self {
        Styles {
            num_fmts: vec![],
            fonts: vec![Default::default()],
            fills: vec![Default::default()],
            borders: vec![Default::default()],
            cell_style_xfs: vec![Default::default()],
            cell_xfs: vec![Default::default()],
            cell_styles: vec![Default::default()],
        }
    }
}

#[derive(Serialize, Deserialize, Debug, PartialEq, Eq, Clone)]
pub struct NumFmt {
    pub num_fmt_id: i32,
    pub format_code: String,
}

impl Default for NumFmt {
    fn default() -> Self {
        NumFmt {
            num_fmt_id: 0,
            format_code: "general".to_string(),
        }
    }
}
#[derive(Serialize, Deserialize, Debug, PartialEq, Clone)]
pub struct Font {
    #[serde(default = "default_as_false")]
    #[serde(skip_serializing_if = "is_false")]
    pub strike: bool,
    #[serde(default = "default_as_false")]
    #[serde(skip_serializing_if = "is_false")]
    pub u: bool,
    pub b: bool,
    pub i: bool,
    pub sz: i32,
    pub color: Color,
    pub name: String,
    pub family: i32,
    pub scheme: String,
}

impl Default for Font {
    fn default() -> Self {
        Font {
            strike: false,
            u: false,
            b: false,
            i: false,
            sz: 11,
            color: Color::RGB("#000000".to_string()),
            name: "Calibri".to_string(),
            family: 2,
            scheme: "minor".to_string(),
        }
    }
}

// TODO: Maybe use an enum for the pattern_type values here?
#[derive(Serialize, Deserialize, Debug, PartialEq, Clone)]
pub struct Fill {
    pub pattern_type: String,
    #[serde(default = "Color::new")]
    #[serde(skip_serializing_if = "Color::is_none")]
    pub fg_color: Color,
    #[serde(default = "Color::new")]
    #[serde(skip_serializing_if = "Color::is_none")]
    pub bg_color: Color,
}

impl Default for Fill {
    fn default() -> Self {
        Fill {
            pattern_type: "none".to_string(),
            fg_color: Default::default(),
            bg_color: Default::default(),
        }
    }
}

#[derive(Serialize, Deserialize, Debug, PartialEq, Eq, Clone)]
pub struct CellStyleXfs {
    pub num_fmt_id: i32,
    pub font_id: i32,
    pub fill_id: i32,
    pub border_id: i32,
    #[serde(default = "default_as_true")]
    #[serde(skip_serializing_if = "is_true")]
    pub apply_number_format: bool,
    #[serde(default = "default_as_true")]
    #[serde(skip_serializing_if = "is_true")]
    pub apply_border: bool,
    #[serde(default = "default_as_true")]
    #[serde(skip_serializing_if = "is_true")]
    pub apply_alignment: bool,
    #[serde(default = "default_as_true")]
    #[serde(skip_serializing_if = "is_true")]
    pub apply_protection: bool,
    #[serde(default = "default_as_true")]
    #[serde(skip_serializing_if = "is_true")]
    pub apply_font: bool,
    #[serde(default = "default_as_true")]
    #[serde(skip_serializing_if = "is_true")]
    pub apply_fill: bool,
}

impl Default for CellStyleXfs {
    fn default() -> Self {
        CellStyleXfs {
            num_fmt_id: 0,
            font_id: 0,
            fill_id: 0,
            border_id: 0,
            apply_number_format: true,
            apply_border: true,
            apply_alignment: true,
            apply_protection: true,
            apply_font: true,
            apply_fill: true,
        }
    }
}

#[derive(Serialize, Deserialize, Debug, PartialEq, Eq, Clone)]
pub struct CellXfs {
    pub xf_id: i32,
    pub num_fmt_id: i32,
    pub font_id: i32,
    pub fill_id: i32,
    pub border_id: i32,
    #[serde(default = "default_as_default")]
    #[serde(skip_serializing_if = "is_default")]
    pub horizontal_alignment: String,
    #[serde(default = "default_as_false")]
    #[serde(skip_serializing_if = "is_false")]
    pub read_only: bool,
    #[serde(default = "default_as_false")]
    #[serde(skip_serializing_if = "is_false")]
    pub apply_number_format: bool,
    #[serde(default = "default_as_false")]
    #[serde(skip_serializing_if = "is_false")]
    pub apply_border: bool,
    #[serde(default = "default_as_false")]
    #[serde(skip_serializing_if = "is_false")]
    pub apply_alignment: bool,
    #[serde(default = "default_as_false")]
    #[serde(skip_serializing_if = "is_false")]
    pub apply_protection: bool,
    #[serde(default = "default_as_false")]
    #[serde(skip_serializing_if = "is_false")]
    pub apply_font: bool,
    #[serde(default = "default_as_false")]
    #[serde(skip_serializing_if = "is_false")]
    pub apply_fill: bool,
    #[serde(default = "default_as_false")]
    #[serde(skip_serializing_if = "is_false")]
    pub quote_prefix: bool,
}

impl Default for CellXfs {
    fn default() -> Self {
        CellXfs {
            xf_id: 0,
            num_fmt_id: 0,
            font_id: 0,
            fill_id: 0,
            border_id: 0,
            horizontal_alignment: "default".to_string(),
            read_only: false,
            apply_number_format: false,
            apply_border: false,
            apply_alignment: false,
            apply_protection: false,
            apply_font: false,
            apply_fill: false,
            quote_prefix: false,
        }
    }
}

#[derive(Serialize, Deserialize, Debug, PartialEq, Eq, Clone)]
pub struct CellStyles {
    pub name: String,
    pub xf_id: i32,
    pub builtin_id: i32,
}

impl Default for CellStyles {
    fn default() -> Self {
        CellStyles {
            name: "normal".to_string(),
            xf_id: 0,
            builtin_id: 0,
        }
    }
}

// TODO: Maybe use an enum for the values here?
#[derive(Serialize, Deserialize, Debug, PartialEq, Clone)]
pub struct BorderItem {
    // think, thick, slantDashDot, none, mediumDashed, mediumDashDotDot, mediumDashDot,
    // medium, double, dotted
    pub style: String,
    pub color: Color,
}

impl Default for BorderItem {
    fn default() -> Self {
        BorderItem {
            style: "none".to_string(),
            color: Color::None,
        }
    }
}

#[derive(Serialize, Deserialize, Debug, PartialEq, Clone, Default)]
pub struct Border {
    #[serde(default = "default_as_false")]
    #[serde(skip_serializing_if = "is_false")]
    pub diagonal_up: bool,
    #[serde(default = "default_as_false")]
    #[serde(skip_serializing_if = "is_false")]
    pub diagonal_down: bool,
    pub left: BorderItem,
    pub right: BorderItem,
    pub top: BorderItem,
    pub bottom: BorderItem,
    pub diagonal: BorderItem,
}
