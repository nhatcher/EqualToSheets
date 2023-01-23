pub mod cell;
pub mod error;
pub mod workbook;

#[cfg(test)]
mod tests {
    use equalto_calc::cell::CellValue;

    use crate::workbook::Workbook;

    #[test]
    fn test_new() {
        let mut workbook = Workbook::new().unwrap();

        workbook.set_value("Sheet1!A1", 100.0).unwrap();
        workbook.set_value((0, 2, 1), "foobar").unwrap();
        workbook.set_value("Sheet1!A3", true).unwrap();
        workbook.set_formula("Sheet1!A4", "=A1*2").unwrap();

        assert_eq!(workbook.value((0, 1, 1)).unwrap(), CellValue::Number(100.0),);
        assert_eq!(
            workbook.value("Sheet1!A2").unwrap(),
            CellValue::String("foobar".to_string()),
        );
        assert_eq!(
            workbook.value("Sheet1!A3").unwrap(),
            CellValue::Boolean(true),
        );
        assert_eq!(
            workbook.value("Sheet1!A4").unwrap(),
            CellValue::Number(200.0),
        );

        assert_eq!(workbook.formula("Sheet1!A3").unwrap(), None);
        assert_eq!(
            workbook.formula("Sheet1!A4").unwrap(),
            Some("=A1*2".to_string()),
        );
        assert_eq!(
            workbook.formula((0, 4, 1)).unwrap(),
            Some("=A1*2".to_string()),
        );
    }

    #[test]
    fn test_load() {
        let workbook = Workbook::load("tests/example.xlsx").unwrap();

        assert_eq!(
            workbook.value("Sheet1!A1").unwrap(),
            CellValue::String("A string".to_string()),
        );
        assert_eq!(
            workbook.value("Sheet1!A2").unwrap(),
            CellValue::Number(222.0),
        );
        assert_eq!(
            workbook.value("Second!A1").unwrap(),
            CellValue::String("Tres".to_string()),
        );
    }
}
