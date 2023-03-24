use crate::error::WorkbookError;
use equalto_calc::model::Model;
use equalto_xlsx::import::load_from_excel;

pub struct Workbook {
    pub(crate) calc_model: Model,
}

impl Workbook {
    pub fn new() -> Result<Self, WorkbookError> {
        let calc_model = Model::new_empty("workbook", "en", "UTC")?;
        Ok(Self { calc_model })
    }

    pub fn load(file_path: &str) -> Result<Self, WorkbookError> {
        let model = load_from_excel(file_path, "en", "UTC")?;
        let s = serde_json::to_string(&model).map_err(|e| e.to_string())?;
        let calc_model = Model::from_json(&s)?;
        Ok(Self { calc_model })
    }
}
