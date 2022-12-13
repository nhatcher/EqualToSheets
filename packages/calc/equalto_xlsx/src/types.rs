pub(crate) type ExcelArchive = zip::read::ZipArchive<std::io::BufReader<std::fs::File>>;
