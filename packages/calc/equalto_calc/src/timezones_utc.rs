use std::str::FromStr;

#[derive(Clone, Copy)]
pub enum Tz {
    UTC,
}

pub struct ParseTzError;

impl FromStr for Tz {
    type Err = ParseTzError;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        if s == "UTC" {
            return Ok(Tz::UTC);
        }
        Err(ParseTzError)
    }
}
