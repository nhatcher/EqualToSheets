use std::vec::Vec;

use crate::types::*;

impl Workbook {
    pub fn get_worksheet_names(&self) -> Vec<String> {
        self.worksheets
            .iter()
            .map(|worksheet| worksheet.get_name())
            .collect()
    }
    pub fn get_worksheet_ids(&self) -> Vec<i32> {
        self.worksheets
            .iter()
            .map(|worksheet| worksheet.get_sheet_id())
            .collect()
    }
}
