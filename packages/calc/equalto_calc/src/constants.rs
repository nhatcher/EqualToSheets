/// Excel compatibility values
/// COLUMN_WIDTH and ROW_HEIGHT are pixel values
/// A column width of Excel value `w` will result in `w * COLUMN_WIDTH_FACTOR` pixels
/// Note that these constants are inlined
pub(crate) const DEFAULT_COLUMN_WIDTH: f64 = 100.0;
pub(crate) const DEFAULT_ROW_HEIGHT: f64 = 21.0;
pub(crate) const COLUMN_WIDTH_FACTOR: f64 = 12.0;
pub(crate) const ROW_HEIGHT_FACTOR: f64 = 2.0;

pub(crate) const LAST_COLUMN: i32 = 16_384;
pub(crate) const LAST_ROW: i32 = 1_048_576;
