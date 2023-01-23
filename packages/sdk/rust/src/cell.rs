use crate::error::WorkbookError;
use crate::workbook::Workbook;
use equalto_calc::{calc_result, cell};

pub enum CellReference {
    Text(String), // i.e. Sheet1!A1
    Index { sheet: u32, row: u32, column: u32 },
}

impl From<&str> for CellReference {
    fn from(text: &str) -> Self {
        Self::Text(text.to_string())
    }
}

impl From<(u32, u32, u32)> for CellReference {
    fn from((sheet, row, column): (u32, u32, u32)) -> Self {
        Self::Index { sheet, row, column }
    }
}

impl Workbook {
    pub fn value<C>(&self, cell: C) -> Result<cell::CellValue, WorkbookError>
    where
        C: Into<CellReference>,
    {
        let cell = self.parse_cell_reference(cell)?;
        Ok(self
            .calc_model
            .get_cell_value_by_index(cell.sheet, cell.row, cell.column)?)
    }

    pub fn set_value<C, V>(&mut self, cell: C, value: V) -> Result<(), WorkbookError>
    where
        C: Into<CellReference>,
        V: Into<cell::CellValue>,
    {
        let cell = self.parse_cell_reference(cell)?;
        match value.into() {
            cell::CellValue::Number(number) => self.set_number(&cell, number),
            cell::CellValue::String(text) => self.set_text(&cell, text),
            cell::CellValue::Boolean(boolean) => self.set_bool(&cell, boolean),
        }
        self.calc_model.evaluate_with_error_check()?;
        Ok(())
    }

    pub fn formula<C>(&mut self, cell: C) -> Result<Option<String>, WorkbookError>
    where
        C: Into<CellReference>,
    {
        let cell = self.parse_cell_reference(cell)?;
        Ok(self
            .calc_model
            .cell_formula(cell.sheet, cell.row, cell.column)?)
    }

    pub fn set_formula<C>(&mut self, cell: C, formula: &str) -> Result<(), WorkbookError>
    where
        C: Into<CellReference>,
    {
        let cell = self.parse_cell_reference(cell)?;
        self.calc_model.update_cell_with_formula(
            cell.sheet,
            cell.row,
            cell.column,
            formula.to_string(),
        )?;
        self.calc_model.evaluate_with_error_check()?;
        Ok(())
    }

    fn set_number(&mut self, cell: &calc_result::CellReference, value: f64) {
        self.calc_model
            .update_cell_with_number(cell.sheet, cell.row, cell.column, value);
    }

    fn set_text(&mut self, cell: &calc_result::CellReference, value: String) {
        self.calc_model
            .update_cell_with_text(cell.sheet, cell.row, cell.column, value.as_str());
    }

    fn set_bool(&mut self, cell: &calc_result::CellReference, value: bool) {
        self.calc_model
            .update_cell_with_bool(cell.sheet, cell.row, cell.column, value);
    }

    fn parse_cell_reference<C>(&self, cell: C) -> Result<calc_result::CellReference, WorkbookError>
    where
        C: Into<CellReference>,
    {
        match cell.into() {
            CellReference::Index { sheet, row, column } => Ok(calc_result::CellReference {
                sheet,
                row: row as i32,
                column: column as i32,
            }),
            CellReference::Text(cell) => {
                self.calc_model
                    .parse_reference(cell.as_str())
                    .ok_or_else(|| WorkbookError {
                        message: format!("invalid cell reference: '{cell}'"),
                    })
            }
        }
    }
}
