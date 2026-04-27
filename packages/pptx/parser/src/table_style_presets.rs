// Built-in PowerPoint table style presets.
//
// Color-transform logic adapted from:
//   LibreOffice oox/source/drawingml/table/predefined-table-styles.cxx
//   Mozilla Public License v. 2.0 — https://www.mozilla.org/en-US/MPL/2.0/
//
// GUID → (family, accent) catalog verified against:
//   https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/hh273476

use std::collections::HashMap;

use crate::{Fill, Stroke, TableStyleDef};

// ── Color helpers ────────────────────────────────────────────────────────────

fn rgb_to_hls(r: f64, g: f64, b: f64) -> (f64, f64, f64) {
    let max = r.max(g).max(b);
    let min = r.min(g).min(b);
    let l = (max + min) / 2.0;
    let d = max - min;
    if d < 1e-10 {
        return (0.0, l, 0.0);
    }
    let s = if l < 0.5 { d / (max + min) } else { d / (2.0 - max - min) };
    let h = if (r - max).abs() < 1e-10 {
        (g - b) / d + if g < b { 6.0 } else { 0.0 }
    } else if (g - max).abs() < 1e-10 {
        (b - r) / d + 2.0
    } else {
        (r - g) / d + 4.0
    } / 6.0;
    (h, l, s)
}

fn hue_to_rgb(p: f64, q: f64, t: f64) -> f64 {
    let t = t.rem_euclid(1.0);
    if t < 1.0 / 6.0 { p + (q - p) * 6.0 * t }
    else if t < 1.0 / 2.0 { q }
    else if t < 2.0 / 3.0 { p + (q - p) * (2.0 / 3.0 - t) * 6.0 }
    else { p }
}

fn hls_to_rgb(h: f64, l: f64, s: f64) -> (f64, f64, f64) {
    if s < 1e-10 { return (l, l, l); }
    let q = if l < 0.5 { l * (1.0 + s) } else { l + s - l * s };
    let p = 2.0 * l - q;
    (hue_to_rgb(p, q, h + 1.0 / 3.0), hue_to_rgb(p, q, h), hue_to_rgb(p, q, h - 1.0 / 3.0))
}

fn srgb_to_linear(c: f64) -> f64 {
    if c <= 0.04045 { c / 12.92 } else { ((c + 0.055) / 1.055).powf(2.4) }
}

fn linear_to_srgb(c: f64) -> f64 {
    if c <= 0.0031308 { 12.92 * c } else { 1.055 * c.powf(1.0 / 2.4) - 0.055 }
}

/// Apply a sequence of (name, val) color transforms to a hex color.
/// val uses OOXML units (100_000 = 100%).
pub(crate) fn apply_transforms(hex: &str, transforms: &[(&str, i64)]) -> String {
    if hex.len() < 6 { return hex.to_owned(); }
    let r = u8::from_str_radix(&hex[0..2], 16).unwrap_or(0);
    let g = u8::from_str_radix(&hex[2..4], 16).unwrap_or(0);
    let b = u8::from_str_radix(&hex[4..6], 16).unwrap_or(0);
    let mut rf = r as f64 / 255.0;
    let mut gf = g as f64 / 255.0;
    let mut bf = b as f64 / 255.0;
    let mut alpha = 1.0_f64;
    for (name, val) in transforms {
        let v = *val as f64 / 100_000.0;
        match *name {
            "lumMod" => {
                let (h, l, s) = rgb_to_hls(rf, gf, bf);
                let (nr, ng, nb) = hls_to_rgb(h, (l * v).min(1.0), s);
                rf = nr; gf = ng; bf = nb;
            }
            "lumOff" => {
                let (h, l, s) = rgb_to_hls(rf, gf, bf);
                let (nr, ng, nb) = hls_to_rgb(h, (l + v).clamp(0.0, 1.0), s);
                rf = nr; gf = ng; bf = nb;
            }
            "shade" => { rf *= v; gf *= v; bf *= v; }
            "tint" => {
                let lr = srgb_to_linear(rf);
                let lg = srgb_to_linear(gf);
                let lb = srgb_to_linear(bf);
                rf = linear_to_srgb((lr + (1.0 - lr) * v).clamp(0.0, 1.0));
                gf = linear_to_srgb((lg + (1.0 - lg) * v).clamp(0.0, 1.0));
                bf = linear_to_srgb((lb + (1.0 - lb) * v).clamp(0.0, 1.0));
            }
            "alpha" => { alpha = v; }
            _ => {}
        }
    }
    let ro = (rf.clamp(0.0, 1.0) * 255.0).round() as u8;
    let go = (gf.clamp(0.0, 1.0) * 255.0).round() as u8;
    let bo = (bf.clamp(0.0, 1.0) * 255.0).round() as u8;
    if (alpha - 1.0).abs() < 0.004 {
        format!("{:02X}{:02X}{:02X}", ro, go, bo)
    } else {
        let a = (alpha.clamp(0.0, 1.0) * 255.0).round() as u8;
        format!("{:02X}{:02X}{:02X}{:02X}", ro, go, bo, a)
    }
}

// ── Theme color helpers ──────────────────────────────────────────────────────

fn accent(theme: &HashMap<String, String>, idx: Option<u8>) -> Option<String> {
    let key = idx.map(|n| format!("accent{n}")).unwrap_or_else(|| "dk1".into());
    theme.get(&key).cloned()
}

fn lt1(theme: &HashMap<String, String>) -> Option<String> {
    theme.get("lt1").cloned()
}

fn dk1(theme: &HashMap<String, String>) -> Option<String> {
    theme.get("dk1").cloned()
}

fn solid(color: &str) -> Fill {
    Fill::Solid { color: color.to_owned() }
}

fn stroke(color: &str) -> Stroke {
    Stroke { color: color.to_owned(), width: 12700, dash_style: None, head_end: None, tail_end: None }
}

// ── Family generators ────────────────────────────────────────────────────────

fn themed_style_1(theme: &HashMap<String, String>, accent_idx: Option<u8>) -> TableStyleDef {
    let Some(a) = accent(theme, accent_idx) else { return TableStyleDef::default(); };
    let lt = lt1(theme).unwrap_or_else(|| "FFFFFF".into());
    let border = Some(stroke(&a));
    let first_row_fill = Some(solid(&a));
    let band1h_color = apply_transforms(&a, &[("alpha", 40000)]);
    let band1h_fill = Some(solid(&band1h_color));
    TableStyleDef {
        first_row_fill,
        band1h_fill,
        whole_outer_h: border.clone(),
        whole_outer_v: border.clone(),
        whole_inside_h: border.clone(),
        whole_inside_v: border,
        first_row_border_b: Some(stroke(&lt)),
        ..Default::default()
    }
}

fn themed_style_2(theme: &HashMap<String, String>, accent_idx: Option<u8>) -> TableStyleDef {
    if let Some(a) = accent(theme, accent_idx) {
        let lt = lt1(theme).unwrap_or_else(|| "FFFFFF".into());
        let outer_color = apply_transforms(&a, &[("tint", 50000)]);
        let outer = Some(stroke(&outer_color));
        TableStyleDef {
            whole_fill: Some(solid(&a)),
            whole_outer_h: outer.clone(),
            whole_outer_v: outer,
            first_row_border_b: Some(stroke(&lt)),
            ..Default::default()
        }
    } else {
        let tx = dk1(theme).unwrap_or_else(|| "000000".into());
        let outer_color = apply_transforms(&tx, &[("tint", 50000)]);
        let outer = Some(stroke(&outer_color));
        let inside = Some(stroke(&tx));
        TableStyleDef {
            whole_outer_h: outer.clone(),
            whole_outer_v: outer,
            whole_inside_h: inside.clone(),
            whole_inside_v: inside,
            ..Default::default()
        }
    }
}

fn light_style_1(theme: &HashMap<String, String>, accent_idx: Option<u8>) -> TableStyleDef {
    let a = accent(theme, accent_idx)
        .or_else(|| dk1(theme))
        .unwrap_or_else(|| "000000".into());
    let band1h_color = apply_transforms(&a, &[("alpha", 20000)]);
    TableStyleDef {
        band1h_fill: Some(solid(&band1h_color)),
        whole_outer_h: Some(stroke(&a)),
        first_row_border_b: Some(stroke(&a)),
        ..Default::default()
    }
}

fn light_style_2(theme: &HashMap<String, String>, accent_idx: Option<u8>) -> TableStyleDef {
    let a = accent(theme, accent_idx)
        .or_else(|| dk1(theme))
        .unwrap_or_else(|| "000000".into());
    let outer = Some(stroke(&a));
    TableStyleDef {
        first_row_fill: Some(solid(&a)),
        whole_outer_h: outer.clone(),
        whole_outer_v: outer,
        ..Default::default()
    }
}

fn light_style_3(theme: &HashMap<String, String>, accent_idx: Option<u8>) -> TableStyleDef {
    let a = accent(theme, accent_idx)
        .or_else(|| dk1(theme))
        .unwrap_or_else(|| "000000".into());
    let border = Some(stroke(&a));
    let band1h_color = apply_transforms(&a, &[("alpha", 20000)]);
    TableStyleDef {
        band1h_fill: Some(solid(&band1h_color)),
        whole_outer_h: border.clone(),
        whole_outer_v: border.clone(),
        whole_inside_h: border.clone(),
        whole_inside_v: border,
        first_row_border_b: Some(stroke(&a)),
        ..Default::default()
    }
}

fn medium_style_1(theme: &HashMap<String, String>, accent_idx: Option<u8>) -> TableStyleDef {
    let a = accent(theme, accent_idx)
        .or_else(|| dk1(theme))
        .unwrap_or_else(|| "000000".into());
    let lt = lt1(theme).unwrap_or_else(|| "FFFFFF".into());
    let border = Some(stroke(&a));
    let band1h_color = apply_transforms(&a, &[("tint", 20000)]);
    TableStyleDef {
        whole_fill: Some(solid(&lt)),
        first_row_fill: Some(solid(&a)),
        last_row_fill: Some(solid(&lt)),
        band1h_fill: Some(solid(&band1h_color)),
        whole_outer_h: border.clone(),
        whole_outer_v: border,
        whole_inside_h: Some(stroke(&a)),
        ..Default::default()
    }
}

fn medium_style_2(theme: &HashMap<String, String>, accent_idx: Option<u8>) -> TableStyleDef {
    let a = accent(theme, accent_idx)
        .or_else(|| dk1(theme))
        .unwrap_or_else(|| "000000".into());
    let lt = lt1(theme).unwrap_or_else(|| "FFFFFF".into());
    let border = Some(stroke(&lt));
    let whole_color = apply_transforms(&a, &[("tint", 20000)]);
    let band1h_color = apply_transforms(&a, &[("tint", 40000)]);
    TableStyleDef {
        whole_fill: Some(solid(&whole_color)),
        first_row_fill: Some(solid(&a)),
        last_row_fill: Some(solid(&a)),
        first_col_fill: Some(solid(&a)),
        last_col_fill: Some(solid(&a)),
        band1h_fill: Some(solid(&band1h_color)),
        whole_outer_h: border.clone(),
        whole_outer_v: border.clone(),
        whole_inside_h: border.clone(),
        whole_inside_v: border,
        first_row_border_b: Some(stroke(&lt)),
        ..Default::default()
    }
}

fn medium_style_3(theme: &HashMap<String, String>, accent_idx: Option<u8>) -> TableStyleDef {
    let a = accent(theme, accent_idx)
        .or_else(|| dk1(theme))
        .unwrap_or_else(|| "000000".into());
    let dk = dk1(theme).unwrap_or_else(|| "000000".into());
    let lt = lt1(theme).unwrap_or_else(|| "FFFFFF".into());
    let band1h_color = apply_transforms(&dk, &[("tint", 20000)]);
    TableStyleDef {
        whole_fill: Some(solid(&lt)),
        first_row_fill: Some(solid(&a)),
        last_row_fill: Some(solid(&lt)),
        first_col_fill: Some(solid(&a)),
        last_col_fill: Some(solid(&a)),
        band1h_fill: Some(solid(&band1h_color)),
        whole_outer_h: Some(stroke(&dk)),
        first_row_border_b: Some(stroke(&dk)),
        ..Default::default()
    }
}

fn medium_style_4(theme: &HashMap<String, String>, accent_idx: Option<u8>) -> TableStyleDef {
    let a = accent(theme, accent_idx)
        .or_else(|| dk1(theme))
        .unwrap_or_else(|| "000000".into());
    let dk = dk1(theme).unwrap_or_else(|| "000000".into());
    let border = Some(stroke(&a));
    let whole_color = apply_transforms(&a, &[("tint", 20000)]);
    let first_row_color = apply_transforms(&a, &[("tint", 20000)]);
    let last_row_color = apply_transforms(&dk, &[("tint", 20000)]);
    let band1h_color = apply_transforms(&a, &[("tint", 40000)]);
    TableStyleDef {
        whole_fill: Some(solid(&whole_color)),
        first_row_fill: Some(solid(&first_row_color)),
        last_row_fill: Some(solid(&last_row_color)),
        band1h_fill: Some(solid(&band1h_color)),
        whole_outer_h: border.clone(),
        whole_outer_v: border.clone(),
        whole_inside_h: border.clone(),
        whole_inside_v: border,
        ..Default::default()
    }
}

fn dark_style_1(theme: &HashMap<String, String>, accent_idx: Option<u8>) -> TableStyleDef {
    if let Some(a) = accent(theme, accent_idx) {
        let lt = lt1(theme).unwrap_or_else(|| "FFFFFF".into());
        let dk = dk1(theme).unwrap_or_else(|| "000000".into());
        let whole_color = apply_transforms(&a, &[("shade", 20000)]);
        let band1h_color = apply_transforms(&a, &[("shade", 40000)]);
        let col_color = apply_transforms(&a, &[("shade", 60000)]);
        TableStyleDef {
            whole_fill: Some(solid(&whole_color)),
            first_row_fill: Some(solid(&dk)),
            last_row_fill: Some(solid(&col_color)),
            first_col_fill: Some(solid(&col_color)),
            last_col_fill: Some(solid(&col_color)),
            band1h_fill: Some(solid(&band1h_color)),
            first_row_border_b: Some(stroke(&lt)),
            ..Default::default()
        }
    } else {
        let dk = dk1(theme).unwrap_or_else(|| "000000".into());
        let lt = lt1(theme).unwrap_or_else(|| "FFFFFF".into());
        let whole_color = apply_transforms(&dk, &[("tint", 20000)]);
        let band1h_color = apply_transforms(&dk, &[("tint", 40000)]);
        let col_color = apply_transforms(&dk, &[("tint", 60000)]);
        TableStyleDef {
            whole_fill: Some(solid(&whole_color)),
            first_row_fill: Some(solid(&dk)),
            last_row_fill: Some(solid(&col_color)),
            first_col_fill: Some(solid(&col_color)),
            last_col_fill: Some(solid(&col_color)),
            band1h_fill: Some(solid(&band1h_color)),
            first_row_border_b: Some(stroke(&lt)),
            ..Default::default()
        }
    }
}

fn dark_style_2(theme: &HashMap<String, String>, accent_idx: Option<u8>) -> TableStyleDef {
    let a = accent(theme, accent_idx)
        .or_else(|| dk1(theme))
        .unwrap_or_else(|| "000000".into());
    let dk = dk1(theme).unwrap_or_else(|| "000000".into());
    // firstRow fill = paired accent (odd→even offset: 1→2, 3→4, 5→6, none→dk1)
    let first_row_key = match accent_idx {
        Some(1) => "accent2",
        Some(3) => "accent4",
        Some(5) => "accent6",
        _ => "dk1",
    };
    let first_row_hex = theme.get(first_row_key).cloned().unwrap_or_else(|| dk.clone());
    let whole_color = apply_transforms(&a, &[("tint", 20000)]);
    let band1h_color = apply_transforms(&a, &[("tint", 40000)]);
    let last_row_color = apply_transforms(&a, &[("tint", 20000)]);
    TableStyleDef {
        whole_fill: Some(solid(&whole_color)),
        first_row_fill: Some(solid(&first_row_hex)),
        last_row_fill: Some(solid(&last_row_color)),
        band1h_fill: Some(solid(&band1h_color)),
        ..Default::default()
    }
}

// ── GUID catalog ─────────────────────────────────────────────────────────────

#[derive(Clone, Copy)]
enum Family {
    ThemedStyle1,
    ThemedStyle2,
    LightStyle1,
    LightStyle2,
    LightStyle3,
    MediumStyle1,
    MediumStyle2,
    MediumStyle3,
    MediumStyle4,
    DarkStyle1,
    DarkStyle2,
}

// (GUID, family, accent_index: 0=no accent/dk1, 1-6=accent1-6)
const CATALOG: &[(&str, Family, u8)] = &[
    // Themed Style 1
    ("{2D5ABB26-0587-4C30-8999-92F81FD0307C}", Family::ThemedStyle1, 0),
    ("{3C2FFA5D-87B4-456A-9821-1D502468CF0F}", Family::ThemedStyle1, 1),
    ("{284E427A-3D55-4303-BF80-6455036E1DE7}", Family::ThemedStyle1, 2),
    ("{69C7853C-536D-4A76-A0AE-DD22124D55A5}", Family::ThemedStyle1, 3),
    ("{775DCB02-9BB8-47FD-8907-85C794F793BA}", Family::ThemedStyle1, 4),
    ("{35758FB7-9AC5-4552-8A53-C91805E547FA}", Family::ThemedStyle1, 5),
    ("{08FB837D-C827-4EFA-A057-4D05807E0F7C}", Family::ThemedStyle1, 6),
    // Themed Style 2
    ("{5940675A-B579-460E-94D1-54222C63F5DA}", Family::ThemedStyle2, 0),
    ("{D113A9D2-9D6B-4929-AA2D-F23B5EE8CBE7}", Family::ThemedStyle2, 1),
    ("{18603FDC-E32A-4AB5-989C-0864C3EAD2B8}", Family::ThemedStyle2, 2),
    ("{306799F8-075E-4A3A-A7F6-7FBC6576F1A4}", Family::ThemedStyle2, 3),
    ("{E269D01E-BC32-4049-B463-5C60D7B0CCD2}", Family::ThemedStyle2, 4),
    ("{327F97BB-C833-4FB7-BDE5-3F7075034690}", Family::ThemedStyle2, 5),
    ("{638B1855-1B75-4FBE-930C-398BA8C253C6}", Family::ThemedStyle2, 6),
    // Light Style 1
    ("{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}", Family::LightStyle1, 0),
    ("{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}", Family::LightStyle1, 1),
    ("{0E3FDE45-AF77-4B5C-9715-49D594BDF05E}", Family::LightStyle1, 2),
    ("{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}", Family::LightStyle1, 3),
    ("{D27102A9-8310-4765-A935-A1911B00CA55}", Family::LightStyle1, 4),
    ("{5FD0F851-EC5A-4D38-B0AD-8093EC10F338}", Family::LightStyle1, 5),
    ("{68D230F3-CF80-4859-8CE7-A43EE81993B5}", Family::LightStyle1, 6),
    // Light Style 2
    ("{7E9639D4-E3E2-4D34-9284-5A2195B3D0D7}", Family::LightStyle2, 0),
    ("{69012ECD-51FC-41F1-AA8D-1B2483CD663E}", Family::LightStyle2, 1),
    ("{72833802-FEF1-4C79-8D5D-14CF1EAF98D9}", Family::LightStyle2, 2),
    ("{F2DE63D5-997A-4646-A377-4702673A728D}", Family::LightStyle2, 3),
    ("{17292A2E-F333-43FB-9621-5CBBE7FDCDCB}", Family::LightStyle2, 4),
    ("{5A111915-BE36-4E01-A7E5-04B1672EAD32}", Family::LightStyle2, 5),
    ("{912C8C85-51F0-491E-9774-3900AFEF0FD7}", Family::LightStyle2, 6),
    // Light Style 3
    ("{616DA210-FB5B-4158-B5E0-FEB733F419BA}", Family::LightStyle3, 0),
    ("{BC89EF96-8CEA-46FF-86C4-4CE0E7609802}", Family::LightStyle3, 1),
    ("{5DA37D80-6434-44D0-A028-1B22A696006F}", Family::LightStyle3, 2),
    ("{8799B23B-EC83-4686-B30A-512413B5E67A}", Family::LightStyle3, 3),
    ("{ED083AE6-46FA-4A59-8FB0-9F97EB10719F}", Family::LightStyle3, 4),
    ("{BDBED569-4797-4DF1-A0F4-6AAB3CD982D8}", Family::LightStyle3, 5),
    ("{E8B1032C-EA38-4F05-BA0D-38AFFFC7BED3}", Family::LightStyle3, 6),
    // Medium Style 1
    ("{793D81CF-94F2-401A-BA57-92F5A7B2D0C5}", Family::MediumStyle1, 0),
    ("{B301B821-A1FF-4177-AEE7-76D212191A09}", Family::MediumStyle1, 1),
    ("{9DCAF9ED-07DC-4A11-8D7F-57B35C25682E}", Family::MediumStyle1, 2),
    ("{1FECB4D8-DB02-4DC6-A0A2-4F2EBAE1DC90}", Family::MediumStyle1, 3),
    ("{1E171933-4619-4E11-9A3F-F7608DF75F80}", Family::MediumStyle1, 4),
    ("{FABFCF23-3B69-468F-B69F-88F6DE6A72F2}", Family::MediumStyle1, 5),
    ("{10A1B5D5-9B99-4C35-A422-299274C87663}", Family::MediumStyle1, 6),
    // Medium Style 2
    ("{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}", Family::MediumStyle2, 0),
    ("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}", Family::MediumStyle2, 1),
    ("{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}", Family::MediumStyle2, 2),
    ("{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}", Family::MediumStyle2, 3),
    ("{00A15C55-8517-42AA-B614-E9B94910E393}", Family::MediumStyle2, 4),
    ("{7DF18680-E054-41AD-8BC1-D1AEF772440D}", Family::MediumStyle2, 5),
    ("{93296810-A885-4BE3-A3E7-6D5BEEA58F35}", Family::MediumStyle2, 6),
    // Medium Style 3
    ("{8EC20E35-A176-4012-BC5E-935CFFF8708E}", Family::MediumStyle3, 0),
    ("{6E25E649-3F16-4E02-A733-19D2CDBF48F0}", Family::MediumStyle3, 1),
    ("{85BE263C-DBD7-4A20-BB59-AAB30ACAA65A}", Family::MediumStyle3, 2),
    ("{EB344D84-9AFB-497E-A393-DC336BA19D2E}", Family::MediumStyle3, 3),
    ("{EB9631B5-78F2-41C9-869B-9F39066F8104}", Family::MediumStyle3, 4),
    ("{74C1A8A3-306A-4EB7-A6B1-4F7E0EB9C5D6}", Family::MediumStyle3, 5),
    ("{2A488322-F2BA-4B5B-9748-0D474271808F}", Family::MediumStyle3, 6),
    // Medium Style 4
    ("{D7AC3CCA-C797-4891-BE02-D94E43425B78}", Family::MediumStyle4, 0),
    ("{69CF1AB2-1976-4502-BF36-3FF5EA218861}", Family::MediumStyle4, 1),
    ("{8A107856-5554-42FB-B03E-39F5DBC370BA}", Family::MediumStyle4, 2),
    ("{0505E3EF-67EA-436B-97B2-0124C06EBD24}", Family::MediumStyle4, 3),
    ("{C4B1156A-380E-4F78-BDF5-A606A8083BF9}", Family::MediumStyle4, 4),
    ("{22838BEF-8BB2-4498-84A7-C5851F593DF1}", Family::MediumStyle4, 5),
    ("{16D9F66E-5EB9-4882-86FB-DCBF35E3C3E4}", Family::MediumStyle4, 6),
    // Dark Style 1
    ("{E8034E78-7F5D-4C2E-B375-FC64B27BC917}", Family::DarkStyle1, 0),
    ("{125E5076-3810-47DD-B79F-674D7AD40C01}", Family::DarkStyle1, 1),
    ("{37CE84F3-28C3-443E-9E96-99CF82512B78}", Family::DarkStyle1, 2),
    ("{D03447BB-5D67-496B-8E87-E561075AD55C}", Family::DarkStyle1, 3),
    ("{E929F9F4-4A8F-4326-A1B4-22849713DDAB}", Family::DarkStyle1, 4),
    ("{8FD4443E-F989-4FC4-A0C8-D5A2AF1F390B}", Family::DarkStyle1, 5),
    ("{AF606853-7671-496A-8E4F-DF71F8EC918B}", Family::DarkStyle1, 6),
    // Dark Style 2
    ("{5202B0CA-FC54-4496-8BCA-5EF66A818D29}", Family::DarkStyle2, 0),
    ("{0660B408-B3CF-4A94-85FC-2B1E0A45F4A2}", Family::DarkStyle2, 1),
    ("{91EBBBCC-DAD2-459C-BE2E-F6DE35CF9A28}", Family::DarkStyle2, 3),
    ("{46F890A9-2807-4EBB-B81D-B2AA78EC7F39}", Family::DarkStyle2, 5),
];

// ── Public entry point ───────────────────────────────────────────────────────

pub fn lookup_builtin_table_style(
    guid: &str,
    theme: &HashMap<String, String>,
) -> Option<TableStyleDef> {
    let (_, family, accent_u8) = CATALOG.iter().find(|(g, _, _)| *g == guid)?;
    let accent_idx = if *accent_u8 == 0 { None } else { Some(*accent_u8) };
    Some(match family {
        Family::ThemedStyle1  => themed_style_1(theme, accent_idx),
        Family::ThemedStyle2  => themed_style_2(theme, accent_idx),
        Family::LightStyle1   => light_style_1(theme, accent_idx),
        Family::LightStyle2   => light_style_2(theme, accent_idx),
        Family::LightStyle3   => light_style_3(theme, accent_idx),
        Family::MediumStyle1  => medium_style_1(theme, accent_idx),
        Family::MediumStyle2  => medium_style_2(theme, accent_idx),
        Family::MediumStyle3  => medium_style_3(theme, accent_idx),
        Family::MediumStyle4  => medium_style_4(theme, accent_idx),
        Family::DarkStyle1    => dark_style_1(theme, accent_idx),
        Family::DarkStyle2    => dark_style_2(theme, accent_idx),
    })
}
