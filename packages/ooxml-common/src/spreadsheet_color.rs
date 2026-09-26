//! SpreadsheetML color resolution shared by archive parsers and direct legacy
//! spreadsheet projection. ECMA-376 Part 1 §§18.8.3, 18.8.19, 18.8.27.

use std::borrow::Cow;

// Excel built-in indexed color palette (indices 0-63).
const INDEXED_COLORS: &[&str] = &[
    "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF",
    "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF",
    "#800000", "#008000", "#000080", "#808000", "#800080", "#008080", "#C0C0C0", "#808080",
    "#9999FF", "#993366", "#FFFFCC", "#CCFFFF", "#660066", "#FF8080", "#0066CC", "#CCCCFF",
    "#000080", "#FF00FF", "#FFFF00", "#00FFFF", "#800080", "#800000", "#008080", "#0000FF",
    "#00CCFF", "#CCFFFF", "#CCFFCC", "#FFFF99", "#99CCFF", "#FF99CC", "#CC99FF", "#FFCC99",
    "#3366FF", "#33CCCC", "#99CC00", "#FFCC00", "#FF9900", "#FF6600", "#666699", "#969696",
    "#003366", "#339966", "#003300", "#333300", "#993300", "#993366", "#333399", "#333333",
];

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SpreadsheetColor {
    Auto,
    Indexed(u32),
    Argb([u8; 4]),
    Theme(u32),
}

/// Resolve a typed SpreadsheetML color using the XLSX parser's existing policy.
///
/// `Auto` remains unresolved for its caller to interpret. ARGB alpha is dropped
/// because the existing renderer color contract is RGB. Indexed values 64 and
/// 65 resolve to black and white, and invalid indexed values resolve to black;
/// these are compatibility policies carried over from the parser rather than
/// general claims about every SpreadsheetML consumer.
pub fn resolve_color(
    color: SpreadsheetColor,
    tint: Option<f64>,
    theme_colors: &[String],
) -> Option<String> {
    let base: Cow<'_, str> = match color {
        SpreadsheetColor::Auto => return None,
        SpreadsheetColor::Argb([_, r, g, b]) => format!("#{r:02X}{g:02X}{b:02X}").into(),
        SpreadsheetColor::Indexed(index) => indexed_color(index as usize).into(),
        SpreadsheetColor::Theme(index) => theme_colors.get(theme_index(index as usize))?.into(),
    };
    Some(finish(base, valid_tint(tint)))
}

/// Attribute adapter preserving the XLSX parser's established precedence and
/// malformed-input behavior. Direct producers should use [`resolve_color`].
pub fn resolve_color_attrs(
    rgb: Option<&str>,
    theme: Option<&str>,
    tint: Option<&str>,
    indexed: Option<&str>,
    theme_colors: &[String],
) -> Option<String> {
    let tint = valid_tint(tint.and_then(|value| value.trim().parse().ok()));
    if let Some(rgb) = rgb {
        let rgb = if rgb.len() == 8 { rgb.get(2..)? } else { rgb };
        return Some(finish(format!("#{}", rgb.to_uppercase()).into(), tint));
    }
    if let Some(index) = theme.and_then(|value| value.parse::<usize>().ok()) {
        if let Some(base) = theme_colors.get(theme_index(index)) {
            return Some(finish(base.as_str().into(), tint));
        }
    }
    if let Some(index) = indexed.and_then(|value| value.parse::<usize>().ok()) {
        return Some(finish(indexed_color(index).into(), tint));
    }
    None
}

fn valid_tint(tint: Option<f64>) -> f64 {
    tint.filter(|value| (-1.0..=1.0).contains(value))
        .unwrap_or(0.0)
}

// SpreadsheetML theme references use light/dark ordering, whereas the theme
// array stores dk1, lt1, dk2, lt2, accent1..accent6, hlink, folHlink. Preserve
// the parser's index remap (ECMA-376 Part 1 18.8.3 and 22.1.2.7); this is not
// DrawingML's logical-name clrMap and must not use SCHEME_DEFAULT_SLOTS.
fn theme_index(index: usize) -> usize {
    match index {
        0 => 1,
        1 => 0,
        2 => 3,
        3 => 2,
        other => other,
    }
}

fn indexed_color(index: usize) -> &'static str {
    match index {
        64 => "#000000",
        65 => "#FFFFFF",
        _ => INDEXED_COLORS.get(index).copied().unwrap_or("#000000"),
    }
}

fn finish(base: Cow<'_, str>, tint: f64) -> String {
    let hex = base.strip_prefix('#').unwrap_or(&base);
    if tint != 0.0 && hex.len() == 6 && hex.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        apply_tint(&base, tint)
    } else {
        base.into_owned()
    }
}

fn apply_tint(hex: &str, tint: f64) -> String {
    let hex = hex.trim_start_matches('#');
    if hex.len() < 6 {
        return format!("#{hex}");
    }
    let r = u8::from_str_radix(&hex[0..2], 16).unwrap_or(0) as f64 / 255.0;
    let g = u8::from_str_radix(&hex[2..4], 16).unwrap_or(0) as f64 / 255.0;
    let b = u8::from_str_radix(&hex[4..6], 16).unwrap_or(0) as f64 / 255.0;
    let max = r.max(g).max(b);
    let min = r.min(g).min(b);
    let l = (max + min) / 2.0;
    let s = if max == min {
        0.0
    } else if l < 0.5 {
        (max - min) / (max + min)
    } else {
        (max - min) / (2.0 - max - min)
    };
    let h = if max == min {
        0.0
    } else if max == r {
        (g - b) / (max - min) / 6.0
    } else if max == g {
        ((b - r) / (max - min) + 2.0) / 6.0
    } else {
        ((r - g) / (max - min) + 4.0) / 6.0
    };
    let h = if h < 0.0 { h + 1.0 } else { h };
    let new_l = if tint > 0.0 {
        l * (1.0 - tint) + tint
    } else {
        l * (1.0 + tint)
    };
    let (r, g, b) = hls_to_rgb(h, new_l, s);
    format!(
        "#{:02X}{:02X}{:02X}",
        (r * 255.0).round() as u8,
        (g * 255.0).round() as u8,
        (b * 255.0).round() as u8
    )
}

fn hls_to_rgb(h: f64, l: f64, s: f64) -> (f64, f64, f64) {
    if s == 0.0 {
        return (l, l, l);
    }
    let q = if l < 0.5 {
        l * (1.0 + s)
    } else {
        l + s - l * s
    };
    let p = 2.0 * l - q;
    (
        hue_to_rgb(p, q, h + 1.0 / 3.0),
        hue_to_rgb(p, q, h),
        hue_to_rgb(p, q, h - 1.0 / 3.0),
    )
}

fn hue_to_rgb(p: f64, q: f64, mut t: f64) -> f64 {
    if t < 0.0 {
        t += 1.0;
    }
    if t > 1.0 {
        t -= 1.0;
    }
    if t < 1.0 / 6.0 {
        p + (q - p) * 6.0 * t
    } else if t < 0.5 {
        q
    } else if t < 2.0 / 3.0 {
        p + (q - p) * (2.0 / 3.0 - t) * 6.0
    } else {
        p
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn typed_argb_drops_alpha_and_auto_stays_unresolved() {
        assert_eq!(
            resolve_color(SpreadsheetColor::Argb([0x12, 0x34, 0x56, 0x78]), None, &[]).as_deref(),
            Some("#345678")
        );
        assert_eq!(resolve_color(SpreadsheetColor::Auto, None, &[]), None);
    }

    #[test]
    fn typed_indexed_specials_fallback_and_tint_match_existing_policy() {
        assert_eq!(
            resolve_color(SpreadsheetColor::Indexed(0), None, &[]).as_deref(),
            Some("#000000")
        );
        assert_eq!(
            resolve_color(SpreadsheetColor::Indexed(65), None, &[]).as_deref(),
            Some("#FFFFFF")
        );
        assert_eq!(
            resolve_color(SpreadsheetColor::Indexed(999), None, &[]).as_deref(),
            Some("#000000")
        );
        assert_eq!(
            resolve_color(SpreadsheetColor::Indexed(23), Some(0.5), &[]).as_deref(),
            Some("#C0C0C0")
        );
    }

    #[test]
    fn typed_theme_uses_spreadsheet_light_dark_index_map() {
        let theme = ["#000000", "#FFFFFF", "#111111", "#EEEEEE"].map(str::to_owned);
        assert_eq!(
            resolve_color(SpreadsheetColor::Theme(0), None, &theme).as_deref(),
            Some("#FFFFFF")
        );
        assert_eq!(
            resolve_color(SpreadsheetColor::Theme(3), None, &theme).as_deref(),
            Some("#111111")
        );
        assert_eq!(
            resolve_color(SpreadsheetColor::Theme(12), None, &theme),
            None
        );
    }

    #[test]
    fn typed_and_attribute_indexed_resolution_match_all_defined_indices() {
        for index in 0_u32..=65 {
            assert_eq!(
                resolve_color(SpreadsheetColor::Indexed(index), None, &[]),
                resolve_color_attrs(None, None, None, Some(&index.to_string()), &[]),
                "indexed color {index}"
            );
        }
    }

    #[test]
    fn tint_boundaries_apply_and_invalid_values_are_ignored() {
        let base = SpreadsheetColor::Argb([0x80, 0x12, 0x34, 0x56]);
        assert_eq!(
            resolve_color(base, Some(-1.0), &[]).as_deref(),
            Some("#000000")
        );
        assert_eq!(
            resolve_color(base, Some(1.0), &[]).as_deref(),
            Some("#FFFFFF")
        );
        for tint in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY, -1.01, 1.01] {
            assert_eq!(
                resolve_color(base, Some(tint), &[]).as_deref(),
                Some("#123456")
            );
        }
        for tint in ["NaN", "inf", "-inf", "-1.01", "1.01"] {
            assert_eq!(
                resolve_color_attrs(Some("80123456"), None, Some(tint), None, &[]).as_deref(),
                Some("#123456")
            );
        }
    }

    #[test]
    fn malformed_unicode_argb_preserves_safe_attribute_behavior() {
        // Eight UTF-8 bytes, but byte offset two is not a character boundary.
        assert_eq!(
            resolve_color_attrs(Some("€abcde"), None, None, None, &[]),
            None
        );
    }
}
