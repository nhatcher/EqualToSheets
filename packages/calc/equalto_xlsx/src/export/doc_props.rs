use equalto_calc::types::Workbook;

// Application-Defined File Properties part
pub(crate) fn get_app_xml(_model: &Workbook) -> String {
    // contains application name and version

    // The next few are not needed:
    // security. It is password protected (not implemented)
    // Scale
    // Titles of parts

    // Read those from somewhere else
    let application = "Equalto Sheets";
    let app_version = "1.0.0";

    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" \
            xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">\
  <Application>{}</Application>\
  <AppVersion>{}</AppVersion>\
</Properties>",
        application, app_version
    )
}

// Core File Properties part
pub(crate) fn get_core_xml(_model: &Workbook) -> String {
    // contains the name of the creator, last modified and date
    let creator = "EqualTo User";
    let last_modified_by = "EqualTo User";
    let created = "2020-11-27T10:08:29Z";
    // FIXME add now
    let last_modified = "2020-11-27T10:56:45Z";
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<cp:coreProperties \
 xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" \
 xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" \
 xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" \
 xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"> \
<dc:title></dc:title><dc:subject></dc:subject>\
<dc:creator>{}</dc:creator>\
<cp:keywords></cp:keywords>\
<dc:description></dc:description>\
<cp:lastModifiedBy>{}</cp:lastModifiedBy>\
<cp:revision></cp:revision>\
<dcterms:created xsi:type=\"dcterms:W3CDTF\">{}</dcterms:created>\
<dcterms:modified xsi:type=\"dcterms:W3CDTF\">{}</dcterms:modified>\
<cp:category></cp:category>\
<cp:contentStatus></cp:contentStatus>\
</cp:coreProperties>",
        creator, last_modified_by, created, last_modified
    )
}
