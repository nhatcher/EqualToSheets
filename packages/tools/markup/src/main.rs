use equalto_xlsx::import::load_model_from_xlsx;
use std::env;

/// Prints markup of the first sheet in the input file
fn main() {
    let args: Vec<String> = env::args().collect();
    let file_path = &args[1];
    let model = load_model_from_xlsx(file_path, "en", "UTC").unwrap();
    println!("{}", model.sheet_markup(0).unwrap());
}
