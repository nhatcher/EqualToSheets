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

    pub fn worksheet(&self, worksheet_index: i32) -> Result<&Worksheet, String> {
        let index =
            usize::try_from(worksheet_index).map_err(|_| "Invalid sheet index".to_string())?;
        self.worksheets
            .get(index)
            .ok_or_else(|| "Invalid sheet index".to_string())
    }

    pub fn worksheet_mut(&mut self, worksheet_index: i32) -> Result<&mut Worksheet, String> {
        let index =
            usize::try_from(worksheet_index).map_err(|_| "Invalid sheet index".to_string())?;
        self.worksheets
            .get_mut(index)
            .ok_or_else(|| "Invalid sheet index".to_string())
    }
}
