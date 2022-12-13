//! Test a n Excel xlsx file.
//! Returns a list of differences in json format.
//!
//! Usage: test file.xlsx

use equalto_xlsx::compare::test_file;

fn main() {
    let args: Vec<_> = std::env::args().collect();
    if args.len() != 2 {
        panic!("Usage: {} <file.xlsx>", args[0]);
    }
    assert!(test_file(&args[1]).is_ok());
}
