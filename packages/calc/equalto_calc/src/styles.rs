use crate::{
    expressions::utils::LAST_COLUMN,
    model::{
        Model, Style, COLUMN_WIDTH_FACTOR, DEFAULT_COLUMN_WIDTH, DEFAULT_ROW_HEIGHT,
        ROW_HEIGHT_FACTOR,
    },
    number_format::{get_new_num_fmt_index, get_num_fmt},
    types::{Border, CellStyles, CellXfs, Col, Fill, Font, NumFmt, Row},
};

impl Model {
    fn get_font_index(&self, font: &Font) -> Option<i32> {
        let fonts = &self.workbook.styles.fonts;
        for (font_index, item) in fonts.iter().enumerate() {
            if item == font {
                return Some(font_index as i32);
            }
        }
        None
    }
    fn get_fill_index(&self, fill: &Fill) -> Option<i32> {
        let fills = &self.workbook.styles.fills;
        for (fill_index, item) in fills.iter().enumerate() {
            if item == fill {
                return Some(fill_index as i32);
            }
        }
        None
    }
    fn get_border_index(&self, border: &Border) -> Option<i32> {
        let borders = &self.workbook.styles.borders;
        for (border_index, item) in borders.iter().enumerate() {
            if item == border {
                return Some(border_index as i32);
            }
        }
        None
    }
    fn get_num_fmt_index(&self, format_code: &str) -> Option<i32> {
        let num_fmts = &self.workbook.styles.num_fmts;
        for item in num_fmts.iter() {
            if item.format_code == format_code {
                return Some(item.num_fmt_id as i32);
            }
        }
        None
    }

    pub fn create_new_style(&mut self, style: &Style) -> i32 {
        let font = &style.font;
        let font_id = if let Some(index) = self.get_font_index(font) {
            index
        } else {
            self.workbook.styles.fonts.push(font.clone());
            self.workbook.styles.fonts.len() as i32 - 1
        };
        let fill = &style.fill;
        let fill_id = if let Some(index) = self.get_fill_index(fill) {
            index
        } else {
            self.workbook.styles.fills.push(fill.clone());
            self.workbook.styles.fills.len() as i32 - 1
        };
        let border = &style.border;
        let border_id = if let Some(index) = self.get_border_index(border) {
            index
        } else {
            self.workbook.styles.borders.push(border.clone());
            self.workbook.styles.borders.len() as i32 - 1
        };
        let num_fmt = &style.num_fmt;
        let num_fmt_id;
        if let Some(index) = self.get_num_fmt_index(num_fmt) {
            num_fmt_id = index;
        } else {
            num_fmt_id = get_new_num_fmt_index(&self.workbook.styles.num_fmts);
            self.workbook.styles.num_fmts.push(NumFmt {
                format_code: num_fmt.to_string(),
                num_fmt_id,
            });
        }
        let styles = &mut self.workbook.styles;

        styles.cell_xfs.push(CellXfs {
            xf_id: 0,
            num_fmt_id,
            font_id,
            fill_id,
            border_id,
            horizontal_alignment: style.horizontal_alignment.clone(),
            read_only: style.read_only,
            apply_number_format: false,
            apply_border: false,
            apply_alignment: false,
            apply_protection: false,
            apply_font: false,
            apply_fill: false,
            quote_prefix: style.quote_prefix,
        });

        styles.cell_xfs.len() as i32 - 1
    }

    pub fn get_style_index(&self, style: &Style) -> Option<i32> {
        let styles = &self.workbook.styles;
        for (index, cell_xf) in styles.cell_xfs.iter().enumerate() {
            let border_id = cell_xf.border_id as usize;
            let fill_id = cell_xf.fill_id as usize;
            let font_id = cell_xf.font_id as usize;
            let num_fmt_id = cell_xf.num_fmt_id;
            let read_only = cell_xf.read_only;
            let quote_prefix = cell_xf.quote_prefix;
            let horizontal_alignment = cell_xf.horizontal_alignment.clone();
            if style
                == &(Style {
                    horizontal_alignment,
                    read_only,
                    num_fmt: get_num_fmt(num_fmt_id, &styles.num_fmts),
                    fill: styles.fills[fill_id].clone(),
                    font: styles.fonts[font_id].clone(),
                    border: styles.borders[border_id].clone(),
                    quote_prefix,
                })
            {
                return Some(index as i32);
            }
        }
        None
    }

    pub fn set_cell_style(
        &mut self,
        sheet: i32,
        row: i32,
        column: i32,
        style: &Style,
    ) -> Result<(), String> {
        // Check if style exist. If so sets style cell number to that otherwise create a new style.
        let style_index = if let Some(index) = self.get_style_index(style) {
            index
        } else {
            self.create_new_style(style)
        };

        match self.get_cell_mut(sheet, row, column) {
            Some(cell) => {
                cell.set_style(style_index);
            }
            None => {
                // The cell does not exist
                self.workbook.worksheets[sheet as usize].set_cell_empty_with_style(
                    row,
                    column,
                    style_index,
                );
            }
        };
        Ok(())

        // Cleanup: check if the old cell style is still in use
        // TODO
    }

    /// Sets the style "style_name" in cell
    pub fn set_cell_style_by_name(
        &mut self,
        sheet: i32,
        row: i32,
        column: i32,
        style_name: &str,
    ) -> Result<(), String> {
        let style_index = self.get_style_index_by_name(style_name)?;
        match self.get_cell_mut(sheet, row, column) {
            Some(cell) => {
                cell.set_style(style_index);
            }
            None => {
                // The cell does not exist
                self.workbook.worksheets[sheet as usize].set_cell_empty_with_style(
                    row,
                    column,
                    style_index,
                );
            }
        }
        Ok(())
    }

    pub fn set_sheet_style(&mut self, sheet: i32, style_name: &str) -> Result<(), String> {
        let style_index = self.get_style_index_by_name(style_name)?;
        self.workbook.worksheets[sheet as usize].cols = vec![Col {
            min: 1,
            max: LAST_COLUMN,
            width: DEFAULT_COLUMN_WIDTH / COLUMN_WIDTH_FACTOR,
            custom_width: true,
            style: Some(style_index),
        }];
        Ok(())
    }

    pub fn set_sheet_row_style(
        &mut self,
        sheet: i32,
        row: i32,
        style_name: &str,
    ) -> Result<(), String> {
        let style_index = self.get_style_index_by_name(style_name)?;
        let rows = &mut self.workbook.worksheets[sheet as usize].rows;
        for r in rows.iter_mut() {
            if r.r == row {
                r.s = style_index;
                r.custom_format = true;
                return Ok(());
            }
        }
        rows.push(Row {
            height: DEFAULT_ROW_HEIGHT / ROW_HEIGHT_FACTOR,
            r: row,
            custom_format: true,
            custom_height: true,
            s: style_index,
        });
        Ok(())
    }

    pub fn set_sheet_column_style(
        &mut self,
        sheet: i32,
        column: i32,
        style_name: &str,
    ) -> Result<(), String> {
        let style_index = self.get_style_index_by_name(style_name)?;
        let worksheet = match self.workbook.worksheets.get_mut(sheet as usize) {
            Some(s) => s,
            None => return Err("Wrong sheet index".to_string()),
        };
        let cols = &mut worksheet.cols;
        let col = Col {
            min: column,
            max: column,
            width: DEFAULT_COLUMN_WIDTH / COLUMN_WIDTH_FACTOR,
            custom_width: true,
            style: Some(style_index),
        };
        let mut index = 0;
        let mut split = false;
        for c in cols.iter_mut() {
            let min = c.min;
            let max = c.max;
            if min <= column && column <= max {
                if min == column && max == column {
                    c.style = Some(style_index);
                    return Ok(());
                } else {
                    // We need to split the result
                    split = true;
                    break;
                }
            }
            if column < min {
                // We passed, we should insert at index
                break;
            }
            index += 1;
        }
        if split {
            let min = cols[index].min;
            let max = cols[index].max;
            let pre = Col {
                min,
                max: column - 1,
                width: cols[index].width,
                custom_width: cols[index].custom_width,
                style: cols[index].style,
            };
            let post = Col {
                min: column + 1,
                max,
                width: cols[index].width,
                custom_width: cols[index].custom_width,
                style: cols[index].style,
            };
            let index = index as usize;
            cols.remove(index);
            if column != max {
                cols.insert(index, post);
            }
            cols.insert(index, col);
            if column != min {
                cols.insert(index, pre);
            }
        } else {
            cols.insert(index, col);
        }
        Ok(())
    }

    /// Adds a named cell style from an existing index
    /// Fails if the named style already exists or if there is not a style with that index
    pub fn add_named_cell_style(
        &mut self,
        style_name: &str,
        style_index: i32,
    ) -> Result<(), String> {
        if self.get_style_index_by_name(style_name).is_ok() {
            return Err("A style with that name already exists".to_string());
        }
        if self.workbook.styles.cell_xfs.len() < style_index as usize {
            return Err("There is no style with that index".to_string());
        }
        let cell_style = CellStyles {
            name: style_name.to_string(),
            xf_id: style_index,
            builtin_id: 0,
        };
        self.workbook.styles.cell_styles.push(cell_style);
        Ok(())
    }

    // Returns the index of the style or fails.
    // NB: this method is case sensitive
    pub fn get_style_index_by_name(&self, style_name: &str) -> Result<i32, String> {
        for cell_style in &self.workbook.styles.cell_styles {
            if cell_style.name == style_name {
                return Ok(cell_style.xf_id);
            }
        }
        Err(format!("Style '{}' not found", style_name))
    }

    pub fn create_named_style(&mut self, style_name: &str, style: &Style) -> Result<(), String> {
        let style_index = self.create_new_style(style);
        self.add_named_cell_style(style_name, style_index)?;
        Ok(())
    }
}
