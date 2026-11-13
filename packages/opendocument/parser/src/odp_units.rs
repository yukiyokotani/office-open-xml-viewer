//! Lengths, percentages, angles and colours of ODF attributes (Part 3 §18).

// Colours, percentages and opacity are parsed by the helpers the ODT and ODS
// readers share (one copy in the ODP WASM).
pub(crate) use crate::odf_fill::with_alpha;
pub(crate) use crate::odf_props::{hex_color as color, percent};

pub const EMU_PER_CM: f64 = 360_000.0;
pub const EMU_PER_PT: f64 = 12_700.0;
/// 1/100 mm, the unit of `logwidth` / `logheight` (Part 3 §19.171).
pub const EMU_PER_HMM: f64 = 360.0;

/// A `length` (§18.3.18) in EMU: a number with one of the units cm, mm, in,
/// pt, pc or px (CSS 96 px per inch). `None` for anything else.
pub fn length_emu(value: &str) -> Option<f64> {
    let value = value.trim();
    let split = value
        .find(|c: char| c.is_ascii_alphabetic() || c == '%')
        .unwrap_or(value.len());
    let (number, unit) = value.split_at(split);
    let number: f64 = number.trim().parse().ok()?;
    if !number.is_finite() {
        return None;
    }
    let factor = match unit {
        "cm" => EMU_PER_CM,
        "mm" => EMU_PER_CM / 10.0,
        "in" | "inch" => 914_400.0,
        "pt" => EMU_PER_PT,
        "pc" => EMU_PER_PT * 12.0,
        "px" => 9_525.0,
        // A unitless coordinate (svg:x1 without unit is not conforming, but
        // drawing coordinates in 1/100 mm appear in some producers).
        "" => EMU_PER_HMM,
        _ => return None,
    };
    Some(number * factor)
}

#[cfg(test)]
fn length_i64(value: &str) -> Option<i64> {
    length_emu(value).map(|v| v.round() as i64)
}

/// An `angle` (§18.3.1) in degrees: a plain number is degrees; `deg`,
/// `rad` and `grad` units are converted.
pub fn angle_degrees(value: &str) -> Option<f64> {
    let value = value.trim();
    let (number, factor) = if let Some(n) = value.strip_suffix("deg") {
        (n, 1.0)
    } else if let Some(n) = value.strip_suffix("grad") {
        (n, 0.9)
    } else if let Some(n) = value.strip_suffix("rad") {
        (n, 180.0 / std::f64::consts::PI)
    } else {
        (value, 1.0)
    };
    let number: f64 = number.trim().parse().ok()?;
    number.is_finite().then_some(number * factor)
}

pub fn rgb(color: &str) -> (f64, f64, f64) {
    let channel =
        |i: usize| u8::from_str_radix(color.get(i..i + 2).unwrap_or("00"), 16).unwrap_or(0) as f64;
    (channel(0), channel(2), channel(4))
}

pub fn hex(r: f64, g: f64, b: f64) -> String {
    let c = |v: f64| v.round().clamp(0.0, 255.0) as u8;
    format!("{:02X}{:02X}{:02X}", c(r), c(g), c(b))
}

/// Scale a colour by an intensity fraction (`draw:start-intensity`, §19.214).
pub fn intensity(color: &str, factor: f64) -> String {
    if (factor - 1.0).abs() < 1e-9 {
        return color.to_string();
    }
    let (r, g, b) = rgb(color);
    hex(r * factor, g * factor, b * factor)
}

/// LibreOffice's `Color::IsDark` luminance test, used for automatic font
/// colours (`style:use-window-font-color`): luminance at or below 62.
pub fn is_dark(color: &str) -> bool {
    let (r, g, b) = rgb(color);
    (b * 29.0 + g * 151.0 + r * 76.0) / 256.0 <= 62.0
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_lengths_angles_and_colours() {
        assert_eq!(length_i64("1cm"), Some(360_000));
        assert_eq!(length_i64("10mm"), Some(360_000));
        assert_eq!(length_i64("1in"), Some(914_400));
        assert_eq!(length_i64("2pt"), Some(25_400));
        assert_eq!(length_i64("-0.5cm"), Some(-180_000));
        assert_eq!(length_i64("1em"), None);
        assert_eq!(percent("45%"), Some(0.45));
        assert_eq!(angle_degrees("30deg"), Some(30.0));
        assert_eq!(angle_degrees("300"), Some(300.0));
        assert!((angle_degrees("3.14159265rad").unwrap() - 180.0).abs() < 1e-5);
        assert_eq!(color("#b13dc8").as_deref(), Some("B13DC8"));
        assert_eq!(with_alpha("FF0000", 0.5), "FF000080");
        assert!(is_dark("000000") && !is_dark("FFFFFF"));
    }
}
