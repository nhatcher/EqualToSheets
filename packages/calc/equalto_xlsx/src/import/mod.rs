mod colors;
mod metadata;
mod shared_strings;
mod styles;
mod tables;
mod util;
mod workbook;
mod worksheets;

use std::{
    collections::HashMap,
    fs,
    io::{BufReader, Read},
};

use roxmltree::Node;

use equalto_calc::{
    model::Model,
    types::{Workbook, WorkbookSettings},
};

use crate::compare::compare_models;
use crate::error::XlsxError;

use shared_strings::read_shared_strings;

use metadata::load_metadata;
use styles::load_styles;
use util::get_attribute;
use workbook::load_workbook;
use worksheets::{load_sheets, Relationship};

fn load_relationships<R: Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
) -> Result<HashMap<String, Relationship>, XlsxError> {
    let mut file = archive.by_name("xl/_rels/workbook.xml.rels")?;
    let mut text = String::new();
    file.read_to_string(&mut text)?;
    let doc = roxmltree::Document::parse(&text)?;
    let nodes: Vec<Node> = doc
        .descendants()
        .filter(|n| n.has_tag_name("Relationship"))
        .collect();
    let mut rels = HashMap::new();
    for node in nodes {
        rels.insert(
            get_attribute(&node, "Id")?.to_string(),
            Relationship {
                rel_type: get_attribute(&node, "Type")?.to_string(),
                target: get_attribute(&node, "Target")?.to_string(),
            },
        );
    }
    Ok(rels)
}

fn load_xlsx_from_reader<R: Read + std::io::Seek>(
    name: String,
    reader: R,
    locale: &str,
    tz: &str,
) -> Result<Workbook, XlsxError> {
    let mut archive = zip::ZipArchive::new(reader)?;

    let mut shared_strings = read_shared_strings(&mut archive)?;
    let workbook = load_workbook(&mut archive)?;
    let rels = load_relationships(&mut archive)?;
    let mut tables = HashMap::new();
    let worksheets = load_sheets(
        &mut archive,
        &rels,
        &workbook,
        &mut tables,
        &mut shared_strings,
    )?;
    let styles = load_styles(&mut archive)?;
    let metadata = load_metadata(&mut archive)?;
    Ok(Workbook {
        shared_strings,
        defined_names: workbook.defined_names,
        worksheets,
        styles,
        name,
        settings: WorkbookSettings {
            tz: tz.to_string(),
            locale: locale.to_string(),
        },
        metadata,
        tables,
    })
}

// Public methods

/// Imports a file from disk into an internal representation
pub fn load_from_excel(file_name: &str, locale: &str, tz: &str) -> Result<Workbook, XlsxError> {
    let file_path = std::path::Path::new(file_name);
    let file = fs::File::open(file_path)?;
    let reader = BufReader::new(file);
    let name = file_path
        .file_stem()
        .ok_or_else(|| XlsxError::IO("Could not extract workbook name".to_string()))?
        .to_string_lossy()
        .to_string();
    load_xlsx_from_reader(name, reader, locale, tz)
}

pub fn load_xlsx_from_memory(
    name: &str,
    data: &mut [u8],
    locale: &str,
    tz: &str,
) -> Result<Model, XlsxError> {
    let reader = std::io::Cursor::new(data);
    let workbook = load_xlsx_from_reader(name.to_string(), reader, locale, tz)?;
    let mut model = Model::from_workbook(workbook).map_err(XlsxError::Workbook)?;
    check_model_support(&mut model)?;
    Ok(model)
}

pub fn load_model_from_xlsx(file_name: &str, locale: &str, tz: &str) -> Result<Model, XlsxError> {
    let mut model = load_model_from_xlsx_without_support_check(file_name, locale, tz)?;
    check_model_support(&mut model)?;
    Ok(model)
}

pub fn load_model_from_xlsx_without_support_check(
    file_name: &str,
    locale: &str,
    tz: &str,
) -> Result<Model, XlsxError> {
    let workbook = load_from_excel(file_name, locale, tz)?;
    Model::from_workbook(workbook).map_err(XlsxError::Workbook)
}

/// Checks if imported model can be safely used in our system.
/// Doesn't provide full support check, but tries to catch basic errors as soon as possible.
pub fn check_model_support(model: &mut Model) -> Result<(), XlsxError> {
    let model_copy = model.clone();

    // Checks if next evaluation of model won't change it's values. Returns Ok(()) if no changes
    // are made, otherwise returns text report with differences.
    // It is useful when testing if XLSX import can be safely used in our calculation engine.
    // (XLSX contains pre-evaluated values for cells, so it's possible to compare against other
    // spreadsheet applications).

    model
        .evaluate_with_error_check()
        // We want to report issues with unsupported formulas as early as possible.
        .map_err(XlsxError::Evaluation)?;

    compare_models(model, &model_copy).map_err(XlsxError::Comparison)?;

    Ok(())
}
