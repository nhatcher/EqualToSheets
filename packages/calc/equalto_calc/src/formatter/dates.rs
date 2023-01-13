use chrono::Duration;
use chrono::NaiveDate;

pub fn from_excel_date(days: i64) -> NaiveDate {
    let dt = NaiveDate::from_ymd_opt(1900, 1, 1).expect("problem with chrono::NaiveDate");
    dt + Duration::days(days - 2)
}
