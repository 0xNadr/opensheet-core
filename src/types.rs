/// Cell style properties for fonts, fills, borders, and alignment.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CellStyle {
    // Font properties
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub font_name: Option<std::string::String>,
    pub font_size: Option<f64>,
    pub font_color: Option<std::string::String>,

    // Fill properties (solid fill)
    pub fill_color: Option<std::string::String>,

    // Border properties (style: "thin", "medium", "thick", "dashed", "dotted", "double", etc.)
    pub border_left: Option<std::string::String>,
    pub border_right: Option<std::string::String>,
    pub border_top: Option<std::string::String>,
    pub border_bottom: Option<std::string::String>,
    pub border_color: Option<std::string::String>,

    // Alignment properties
    pub horizontal_alignment: Option<std::string::String>,
    pub vertical_alignment: Option<std::string::String>,
    pub wrap_text: bool,
    pub text_rotation: Option<u16>,

    // Number format (e.g. "$#,##0.00", "0.00%")
    pub number_format: Option<std::string::String>,
}

impl CellStyle {
    /// Returns true if any visual style property (font/fill/border/alignment) is set.
    #[allow(dead_code)]
    pub fn has_visual_styling(&self) -> bool {
        self.bold
            || self.italic
            || self.underline
            || self.font_name.is_some()
            || self.font_size.is_some()
            || self.font_color.is_some()
            || self.fill_color.is_some()
            || self.border_left.is_some()
            || self.border_right.is_some()
            || self.border_top.is_some()
            || self.border_bottom.is_some()
            || self.horizontal_alignment.is_some()
            || self.vertical_alignment.is_some()
            || self.wrap_text
            || self.text_rotation.is_some()
    }
}

/// Represents a single cell value from a spreadsheet.
#[derive(Debug, Clone)]
pub enum CellValue {
    String(std::string::String),
    /// A reference to the shared string table by index.
    /// Avoids cloning the string during parsing; resolved to a Python str at conversion time.
    SharedString(usize),
    Number(f64),
    Bool(bool),
    /// A formula with its text and optional cached value.
    Formula {
        formula: std::string::String,
        cached_value: Option<Box<CellValue>>,
    },
    /// A date (year, month, day).
    Date {
        year: i32,
        month: u32,
        day: u32,
    },
    /// A date with time (year, month, day, hour, minute, second, microsecond).
    DateTime {
        year: i32,
        month: u32,
        day: u32,
        hour: u32,
        minute: u32,
        second: u32,
        microsecond: u32,
    },
    /// A number with a custom number format code (e.g. "$#,##0.00", "0.00%").
    FormattedNumber {
        value: f64,
        format_code: std::string::String,
    },
    /// A cell with styling (fonts, fills, borders, alignment).
    StyledCell {
        value: Box<CellValue>,
        style: Box<CellStyle>,
    },
    Empty,
}

/// A worksheet with a name and rows of cell values.
#[derive(Debug, Clone)]
pub struct Sheet {
    pub name: std::string::String,
    pub rows: Vec<Vec<CellValue>>,
    /// Merged cell ranges, e.g. ["A1:B2", "C3:D4"].
    pub merges: Vec<std::string::String>,
    /// Column widths: 0-based column index -> width in character units.
    pub column_widths: std::collections::HashMap<u32, f64>,
    /// Row heights: 0-based row index -> height in points.
    pub row_heights: std::collections::HashMap<u32, f64>,
    /// Freeze pane split point: (rows_frozen, cols_frozen).
    /// None if no freeze panes are set.
    pub freeze_pane: Option<(u32, u32)>,
    /// Auto-filter range (e.g. "A1:C1"), or None if not set.
    pub auto_filter: Option<std::string::String>,
    /// Sheet visibility state: "visible", "hidden", or "veryHidden".
    pub state: std::string::String,
}

/// Convert an Excel serial number to (year, month, day, hour, minute, second, microsecond).
/// Returns None if the serial number is negative.
pub fn excel_serial_to_datetime(serial: f64) -> Option<(i32, u32, u32, u32, u32, u32, u32)> {
    if serial < 0.0 {
        return None;
    }

    let mut day_serial = serial.floor() as i64;
    let time_frac = serial - serial.floor();

    // Handle the Lotus 1-2-3 bug: Excel thinks 1900-02-29 exists (serial 60)
    if day_serial == 0 {
        return Some((1899, 12, 31, 0, 0, 0, 0));
    }
    if day_serial >= 60 {
        day_serial -= 1; // Skip the phantom Feb 29
    }
    // Now day_serial is 1-based days since 1900-01-01
    day_serial -= 1; // Make 0-based

    let (y, m, d) = days_to_ymd(1900, day_serial);

    // Convert fractional day to time
    let total_seconds = (time_frac * 86400.0).round() as u64;
    let hour = (total_seconds / 3600) as u32;
    let minute = ((total_seconds % 3600) / 60) as u32;
    let second = (total_seconds % 60) as u32;
    let microsecond = ((time_frac * 86400.0 - total_seconds as f64) * 1_000_000.0)
        .round()
        .abs() as u32;

    Some((y, m, d, hour, minute, second, microsecond))
}

/// Convert a date/time to an Excel serial number.
pub fn datetime_to_excel_serial(
    year: i32,
    month: u32,
    day: u32,
    hour: u32,
    minute: u32,
    second: u32,
    microsecond: u32,
) -> f64 {
    let mut days = ymd_to_days(1900, year, month, day) + 1; // 1-based

    // Re-add the Lotus 1-2-3 phantom day
    if days >= 60 {
        days += 1;
    }

    let time_frac = (hour as f64 * 3600.0
        + minute as f64 * 60.0
        + second as f64
        + microsecond as f64 / 1_000_000.0)
        / 86400.0;

    days as f64 + time_frac
}

/// Convert days since base_year-01-01 (0-based) to (year, month, day).
fn days_to_ymd(base_year: i32, days: i64) -> (i32, u32, u32) {
    let mut y = base_year;
    let mut remaining = days;

    loop {
        let days_in_y = if is_leap_year(y) { 366 } else { 365 };
        if remaining < days_in_y {
            break;
        }
        remaining -= days_in_y;
        y += 1;
    }

    let leap = is_leap_year(y);
    let month_days: [u32; 12] = if leap {
        [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    } else {
        [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    };

    let mut m = 0;
    for (i, &md) in month_days.iter().enumerate() {
        if remaining < md as i64 {
            m = i;
            break;
        }
        remaining -= md as i64;
    }

    (y, (m + 1) as u32, (remaining + 1) as u32)
}

/// Convert (year, month, day) to days since base_year-01-01 (0-based).
fn ymd_to_days(base_year: i32, year: i32, month: u32, day: u32) -> i64 {
    let mut days: i64 = 0;

    for y in base_year..year {
        days += if is_leap_year(y) { 366 } else { 365 };
    }

    let leap = is_leap_year(year);
    let month_days: [u32; 12] = if leap {
        [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    } else {
        [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    };

    for &md in month_days.iter().take(month as usize - 1) {
        days += md as i64;
    }
    days += (day as i64) - 1;

    days
}

fn is_leap_year(y: i32) -> bool {
    (y % 4 == 0 && y % 100 != 0) || y % 400 == 0
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_excel_serial_to_date() {
        // 1 = 1900-01-01
        assert_eq!(
            excel_serial_to_datetime(1.0),
            Some((1900, 1, 1, 0, 0, 0, 0))
        );
        // 59 = 1900-02-28
        assert_eq!(
            excel_serial_to_datetime(59.0),
            Some((1900, 2, 28, 0, 0, 0, 0))
        );
        // 61 = 1900-03-01 (60 is the phantom Feb 29)
        assert_eq!(
            excel_serial_to_datetime(61.0),
            Some((1900, 3, 1, 0, 0, 0, 0))
        );
        // 44197 = 2021-01-01
        assert_eq!(
            excel_serial_to_datetime(44197.0),
            Some((2021, 1, 1, 0, 0, 0, 0))
        );
        // 45658 = 2025-01-01
        assert_eq!(
            excel_serial_to_datetime(45658.0),
            Some((2025, 1, 1, 0, 0, 0, 0))
        );
    }

    #[test]
    fn test_excel_serial_with_time() {
        // 44197.5 = 2021-01-01 12:00:00
        let result = excel_serial_to_datetime(44197.5).unwrap();
        assert_eq!((result.0, result.1, result.2), (2021, 1, 1));
        assert_eq!(result.3, 12); // hour
        assert_eq!(result.4, 0); // minute
    }

    #[test]
    fn test_roundtrip() {
        let serial = datetime_to_excel_serial(2025, 3, 15, 14, 30, 0, 0);
        let (y, m, d, h, min, s, _us) = excel_serial_to_datetime(serial).unwrap();
        assert_eq!((y, m, d, h, min, s), (2025, 3, 15, 14, 30, 0));
    }

    #[test]
    fn test_date_to_serial() {
        // 2021-01-01 should be 44197
        let serial = datetime_to_excel_serial(2021, 1, 1, 0, 0, 0, 0);
        assert!((serial - 44197.0).abs() < 0.001);
    }
}
