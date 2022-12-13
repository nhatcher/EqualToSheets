use std::collections::HashMap;

use crate::{
    model::{Model, Style},
    types::{Color, Fill, Font},
};

/// EqualTo specific methods.
/// This file contains methods that are unlikely to be used outside of EqualTo
/// Note that there might be some other methods that are very EqualTo specific in other files.
/// We will keep the `equalto` name in the API name

impl Model {
    /// Adds the read-only styles needed in an EqualTo workbook
    pub fn add_equalto_read_only_styles(&mut self) -> Result<(), String> {
        let read_only_fill = Fill {
            pattern_type: "solid".to_string(),
            fg_color: Color::RGB("#F1F2F8".to_string()),
            bg_color: Color::RGB("#21243A".to_string()),
        };
        let read_only_body = Style {
            horizontal_alignment: "default".to_string(),
            read_only: true,
            num_fmt: "general".to_string(),
            fill: read_only_fill.clone(),
            font: Default::default(),
            border: Default::default(),
            quote_prefix: false,
        };
        let read_only_date = Style {
            horizontal_alignment: "default".to_string(),
            read_only: true,
            num_fmt: "dd/mm/yyyy;@".to_string(),
            fill: read_only_fill.clone(),
            font: Default::default(),
            border: Default::default(),
            quote_prefix: false,
        };
        let bold_font = Font {
            b: true,
            ..Default::default()
        };
        let read_only_header = Style {
            horizontal_alignment: "default".to_string(),
            read_only: true,
            num_fmt: "general".to_string(),
            fill: read_only_fill,
            font: bold_font,
            border: Default::default(),
            quote_prefix: false,
        };
        let style_index = self.create_new_style(&read_only_body);
        self.add_named_cell_style("read_only_body", style_index)?;
        let style_index = self.create_new_style(&read_only_date);
        self.add_named_cell_style("read_only_date", style_index)?;
        let style_index = self.create_new_style(&read_only_header);
        self.add_named_cell_style("read_only_header", style_index)?;
        Ok(())
    }

    /// Removes all data on a sheet, including the cell styles
    pub fn remove_sheet_data(&mut self, sheet_index: i32) -> Result<(), String> {
        let worksheet = match self.workbook.worksheets.get_mut(sheet_index as usize) {
            Some(s) => s,
            None => return Err("Wrong worksheet index".to_string()),
        };

        // Remove all data
        worksheet.sheet_data = HashMap::new();
        Ok(())
    }
}
