use chrono_tz::Tz;

use std::collections::HashMap;

use crate::{
    expressions::{
        lexer::LexerMode,
        parser::stringify::{rename_sheet_in_node, to_rc_format},
        parser::Parser,
        types::CellReferenceRC,
    },
    language::get_language,
    locale::get_locale,
    model::{Environment, Model},
    types::{SheetState, Workbook, WorkbookSettings, Worksheet},
};

/// You can use all alphanumeric characters but not the following special characters:
/// \ , / , * , ? , : , [ , ].
fn is_valid_sheet_name(name: &str) -> bool {
    let invalid = ['\\', '/', '*', '?', ':', '[', ']'];
    if name.contains(&invalid[..]) {
        return false;
    }
    true
}

impl Model {
    /// Creates a new worksheet. Note that it does not check if the name or the sheet_id exists
    fn new_empty_worksheet(name: &str, sheet_id: i32) -> Worksheet {
        Worksheet {
            cols: vec![],
            rows: vec![],
            comments: vec![],
            dimension: "A1".to_string(),
            merge_cells: vec![],
            name: name.to_string(),
            shared_formulas: vec![],
            sheet_data: Default::default(),
            sheet_id,
            state: SheetState::Visible,
            color: Default::default(),
            frozen_columns: 0,
            frozen_rows: 0,
        }
    }

    pub fn get_new_sheet_id(&self) -> i32 {
        let mut index = 1;
        let worksheets = &self.workbook.worksheets;
        for worksheet in worksheets {
            index = index.max(worksheet.sheet_id);
        }
        index + 1
    }
    pub(crate) fn parse_formulas(&mut self) {
        self.parser.set_lexer_mode(LexerMode::R1C1);
        let worksheets = &self.workbook.worksheets;
        for worksheet in worksheets {
            let shared_formulas = &worksheet.shared_formulas;
            let cell_reference = &Some(CellReferenceRC {
                sheet: worksheet.get_name(),
                row: 1,
                column: 1,
            });
            let mut parse_formula = Vec::new();
            for formula in shared_formulas {
                let t = self.parser.parse(formula, cell_reference);
                parse_formula.push(t);
            }
            self.parsed_formulas.push(parse_formula);
        }
        self.parser.set_lexer_mode(LexerMode::A1);
    }

    // Reparses all formulas
    fn reset_formulas(&mut self) {
        self.parser.set_worksheets(self.get_worksheet_names());
        self.parsed_formulas = vec![];
        self.parse_formulas();
        self.evaluate();
    }

    /// Adds a sheet with a automatically generated name
    pub fn new_sheet(&mut self) {
        // First we find a name

        // TODO: When/if we support i18n the name could depend on the locale
        let base_name = "Sheet";
        let base_name_uppercase = base_name.to_uppercase();
        let mut index = 1;
        while self
            .get_worksheet_names()
            .iter()
            .map(|s| s.to_uppercase())
            .any(|x| x == format!("{}{}", base_name_uppercase, index))
        {
            index += 1;
        }
        let sheet_name = format!("{}{}", base_name, index);
        // Now we need a sheet_id
        let sheet_id = self.get_new_sheet_id();
        let worksheet = Model::new_empty_worksheet(&sheet_name, sheet_id);
        self.workbook.worksheets.push(worksheet);
        self.reset_formulas();
    }

    /// Inserts a sheet with a particular index
    /// Fails if a worksheet with that name already exists or the name is invalid
    /// Fails if the index is too large
    pub fn insert_sheet(
        &mut self,
        sheet_name: &str,
        sheet_index: u32,
        sheet_id: Option<i32>,
    ) -> Result<(), String> {
        if !is_valid_sheet_name(sheet_name) {
            return Err(format!("Invalid name for a sheet: '{}'", sheet_name));
        }
        if self
            .get_worksheet_names()
            .iter()
            .map(|s| s.to_uppercase())
            .any(|x| x == sheet_name.to_uppercase())
        {
            return Err("A worksheet already exists with that name".to_string());
        }
        let sheet_id = match sheet_id {
            Some(id) => id,
            None => self.get_new_sheet_id(),
        };
        let worksheet = Model::new_empty_worksheet(sheet_name, sheet_id);
        if sheet_index as usize > self.workbook.worksheets.len() {
            return Err("Sheet index out of range".to_string());
        }
        self.workbook
            .worksheets
            .insert(sheet_index as usize, worksheet);
        self.reset_formulas();
        Ok(())
    }

    /// Adds a sheet with a specific name
    /// Fails if a worksheet with that name already exists or the name is invalid
    pub fn add_sheet(&mut self, sheet_name: &str) -> Result<(), String> {
        self.insert_sheet(sheet_name, self.workbook.worksheets.len() as u32, None)
    }

    /// Renames a sheet and updates all existing references to that sheet.
    /// It can fail if:
    ///   * The original sheet does not exists
    ///   * The target sheet already exists
    ///   * The target sheet name is invalid
    pub fn rename_sheet(&mut self, old_name: &str, new_name: &str) -> Result<(), String> {
        if let Some(sheet_index) = self.get_sheet_index_by_name(old_name) {
            return self.rename_sheet_by_index(sheet_index, new_name);
        }
        Err(format!("Could not find sheet {}", old_name))
    }

    /// Renames a sheet and updates all existing references to that sheet.
    /// It can fail if:
    ///   * The original index is too large
    ///   * The target sheet name already exists
    ///   * The target sheet name is invalid
    pub fn rename_sheet_by_index(
        &mut self,
        sheet_index: u32,
        new_name: &str,
    ) -> Result<(), String> {
        if !is_valid_sheet_name(new_name) {
            return Err(format!("Invalid name for a sheet: '{}'", new_name));
        }
        if self.get_sheet_index_by_name(new_name).is_some() {
            return Err(format!("Sheet already exists: '{}'", new_name));
        }
        let worksheets = &self.workbook.worksheets;
        let sheet_count = worksheets.len() as u32;
        if sheet_index >= sheet_count {
            return Err("Sheet index out of bounds".to_string());
        }
        // Parse all formulas with the old name
        // All internal formulas are R1C1
        self.parser.set_lexer_mode(LexerMode::R1C1);
        // We use iter because the default would be a mut_iter and we don't need a mutable reference
        let worksheets = &mut self.workbook.worksheets;
        for worksheet in worksheets {
            let cell_reference = &Some(CellReferenceRC {
                sheet: worksheet.get_name(),
                row: 1,
                column: 1,
            });
            let mut formulas = Vec::new();
            for formula in &worksheet.shared_formulas {
                let mut t = self.parser.parse(formula, cell_reference);
                rename_sheet_in_node(&mut t, sheet_index, new_name);
                formulas.push(to_rc_format(&t));
            }
            worksheet.shared_formulas = formulas;
        }
        // Se the mode back to A1
        self.parser.set_lexer_mode(LexerMode::A1);
        // Update the name of the worksheet
        let worksheets = &mut self.workbook.worksheets;
        worksheets[sheet_index as usize].set_name(new_name);
        self.reset_formulas();
        Ok(())
    }

    /// Deletes a sheet by index. Fails if:
    ///   * The sheet does not exists
    ///   * It is the last sheet
    pub fn delete_sheet(&mut self, sheet_index: u32) -> Result<(), String> {
        let worksheets = &self.workbook.worksheets;
        let sheet_count = worksheets.len() as u32;
        if sheet_count == 1 {
            return Err("Cannot delete only sheet".to_string());
        };
        if sheet_index > sheet_count {
            return Err("Sheet index too large".to_string());
        }
        self.workbook.worksheets.remove(sheet_index as usize);
        self.reset_formulas();
        Ok(())
    }

    /// Deletes a sheet by name. Fails if:
    ///   * The sheet does not exists
    ///   * It is the last sheet
    pub fn delete_sheet_by_name(&mut self, name: &str) -> Result<(), String> {
        if let Some(sheet_index) = self.get_sheet_index_by_name(name) {
            self.delete_sheet(sheet_index)
        } else {
            Err("Sheet not found".to_string())
        }
    }

    /// Deletes a sheet by sheet_id. Fails if:
    ///   * The sheet by sheet_id does not exists
    ///   * It is the last sheet
    pub fn delete_sheet_by_sheet_id(&mut self, sheet_id: i32) -> Result<(), String> {
        if let Some(sheet_index) = self.get_sheet_index_by_sheet_id(sheet_id) {
            self.delete_sheet(sheet_index)
        } else {
            Err("Sheet not found".to_string())
        }
    }

    pub(crate) fn get_sheet_index_by_sheet_id(&self, sheet_id: i32) -> Option<u32> {
        let worksheets = &self.workbook.worksheets;
        for (index, worksheet) in worksheets.iter().enumerate() {
            if worksheet.sheet_id == sheet_id {
                return Some(index as u32);
            }
        }
        None
    }

    /// Creates a new workbook with one empty sheet
    pub fn new_empty(
        name: &str,
        locale_id: &str,
        timezone: &str,
        env: Environment,
    ) -> Result<Model, String> {
        let tz: Tz = match &timezone.parse() {
            Ok(tz) => *tz,
            Err(_) => return Err(format!("Invalid timezone: {}", &timezone)),
        };
        let locale = match get_locale(locale_id) {
            Ok(l) => l.clone(),
            Err(_) => return Err(format!("Invalid locale: {}", locale_id)),
        };
        // String versions of the locale are added here to simplify the serialize/deserialize logic
        let workbook = Workbook {
            shared_strings: vec![],
            defined_names: vec![],
            worksheets: vec![Model::new_empty_worksheet("Sheet1", 1)],
            styles: Default::default(),
            name: name.to_string(),
            settings: WorkbookSettings {
                tz: timezone.to_string(),
                locale: locale_id.to_string(),
            },
        };
        let parsed_formulas = Vec::new();
        let worksheets = &workbook.worksheets;
        let parser = Parser::new(worksheets.iter().map(|s| s.get_name()).collect());
        let cells = HashMap::new();

        // FIXME: Add support for display languages
        let language = get_language("en").expect("").clone();

        let mut model = Model {
            workbook,
            parsed_formulas,
            parser,
            cells,
            locale,
            language,
            env,
            tz,
        };
        model.parse_formulas();
        Ok(model)
    }
}
