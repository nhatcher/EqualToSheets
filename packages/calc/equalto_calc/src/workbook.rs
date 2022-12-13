use std::vec::Vec;

use crate::types::*;

impl Workbook {
    pub fn get_worksheet_names(&self) -> Vec<String> {
        let mut names = Vec::new();
        for worksheet in &self.worksheets {
            names.push(worksheet.get_name());
        }
        names
    }
}
