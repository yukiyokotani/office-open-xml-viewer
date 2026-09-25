//! Built-in PivotTable styles as model formats.
//!
//! The preset table lives in `ooxml_common::spreadsheet_style_presets`; this
//! builds the model `Dxf` of each element exactly as the XLSX parser reads a
//! `<dxf>` (font size 11, a lone background mirrored into the foreground)
//! and keeps a style-less edge as a `none` edge, which clears an earlier
//! PivotTable style element's edge (ECMA-376 §18.8.6 style default `none`,
//! §18.8.41 layering).

use crate::resolve_color_attrs;
use crate::types::{Border, BorderEdge, Dxf, Fill, Font, PivotTableStyleElement};
use ooxml_common::spreadsheet_style_presets::{self, PresetColor, PresetDxf, PresetEdge};

/// A preset theme reference resolved exactly as a `<color theme tint>`
/// attribute pair in `styles.xml` (§18.8.3, including the theme index remap).
fn color(color: Option<PresetColor>, theme_colors: &[String]) -> Option<String> {
    let color = color?;
    let theme = color.theme.to_string();
    let tint = color.tint.map(|tint| tint.to_string());
    resolve_color_attrs(None, Some(&theme), tint.as_deref(), None, theme_colors)
}

fn edge(edge: Option<PresetEdge>, theme_colors: &[String]) -> Option<BorderEdge> {
    let edge = edge?;
    Some(match edge.style {
        Some(style) => BorderEdge {
            style: style.to_string(),
            color: color(edge.color, theme_colors),
        },
        None => BorderEdge {
            style: "none".into(),
            color: None,
        },
    })
}

/// The model format of a preset `<dxf>` under `theme_colors` (the workbook
/// theme in clrScheme order, as the XLSX parser holds it).
pub(crate) fn preset_dxf(dxf: &PresetDxf, theme_colors: &[String]) -> Dxf {
    Dxf {
        font: dxf.font.map(|font| Font {
            bold: font.bold,
            size: 11.0,
            color: color(font.color, theme_colors),
            ..Default::default()
        }),
        fill: dxf.fill.map(|fill| {
            let bg_color = color(fill.bg, theme_colors);
            Fill {
                pattern_type: fill.pattern_type.to_string(),
                fg_color: color(fill.fg, theme_colors).or_else(|| bg_color.clone()),
                bg_color,
                gradient: None,
            }
        }),
        border: dxf.border.map(|border| Border {
            left: edge(border.left, theme_colors),
            right: edge(border.right, theme_colors),
            top: edge(border.top, theme_colors),
            bottom: edge(border.bottom, theme_colors),
            horizontal: edge(border.horizontal, theme_colors),
            vertical: edge(border.vertical, theme_colors),
            ..Default::default()
        }),
        ..Default::default()
    }
}

/// The elements of the built-in PivotTable style `name`, or `None` when it
/// is not a built-in style.
pub(crate) fn pivot_style_elements(
    name: &str,
    theme_colors: &[String],
) -> Option<Vec<PivotTableStyleElement>> {
    Some(
        spreadsheet_style_presets::pivot_style(name)?
            .into_iter()
            .map(|element| PivotTableStyleElement {
                kind: element.kind.to_string(),
                size: element.size,
                dxf: preset_dxf(element.dxf, theme_colors),
            })
            .collect(),
    )
}
