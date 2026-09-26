//! Built-in SpreadsheetML PivotTable styles (ECMA-376 Part 1 §18.8.40,
//! Annex G presetTableStyles.xml) as typed differential formats.
//!
//! A workbook names a built-in style without defining it (pivotTableStyleInfo
//! in XLSX, SXAddl_SXCView_SXDTableStyleClient in XLS), so every spreadsheet
//! producer resolves it from this one table. Colors stay theme references
//! (§18.8.3); callers resolve them against the workbook theme.

/// A theme color reference (`theme`, optional `tint`) of a preset format.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct PresetColor {
    pub theme: u32,
    pub tint: Option<f64>,
}

/// `<font>`: the preset formats only set bold and a color.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct PresetFont {
    pub bold: bool,
    pub color: Option<PresetColor>,
}

/// `<fill><patternFill>`.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct PresetFill {
    pub pattern_type: &'static str,
    pub fg: Option<PresetColor>,
    pub bg: Option<PresetColor>,
}

/// A border edge; `style` is `None` when the edge element has no style
/// attribute (an explicitly present, style-less edge).
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct PresetEdge {
    pub style: Option<&'static str>,
    pub color: Option<PresetColor>,
}

/// `<border>` edges, including the region-interior `horizontal`/`vertical`.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct PresetBorder {
    pub left: Option<PresetEdge>,
    pub right: Option<PresetEdge>,
    pub top: Option<PresetEdge>,
    pub bottom: Option<PresetEdge>,
    pub horizontal: Option<PresetEdge>,
    pub vertical: Option<PresetEdge>,
}

/// One preset `<dxf>`.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct PresetDxf {
    pub font: Option<PresetFont>,
    pub fill: Option<PresetFill>,
    pub border: Option<PresetBorder>,
}

/// One `tableStyleElement` of a preset style.
#[derive(Clone, Copy, Debug)]
pub struct PresetElement {
    /// ST_TableStyleType (§18.18.82), e.g. `firstRowStripe`.
    pub kind: &'static str,
    /// Band size (`size`, default 1).
    pub size: u32,
    pub dxf: &'static PresetDxf,
}

impl PresetDxf {
    /// Theme indices this format references.
    pub fn theme_indices(&self) -> impl Iterator<Item = u32> + '_ {
        let font = self.font.and_then(|font| font.color);
        let fill = self.fill.into_iter().flat_map(|fill| [fill.fg, fill.bg]);
        let border = self.border.into_iter().flat_map(|border| {
            [
                border.left,
                border.right,
                border.top,
                border.bottom,
                border.horizontal,
                border.vertical,
            ]
            .into_iter()
            .map(|edge| edge.and_then(|edge| edge.color))
        });
        std::iter::once(font)
            .chain(fill)
            .chain(border)
            .flatten()
            .map(|color| color.theme)
    }
}

/// A generated style entry: (name, [(element type, band size, format index)]).
type PresetStyleEntry = (&'static str, &'static [(&'static str, u32, u16)]);

include!("spreadsheet_style_presets.generated.rs");

/// The elements of the built-in PivotTable style `name` (e.g.
/// `PivotStyleLight16`), or `None` when `name` is not a built-in style.
pub fn pivot_style(name: &str) -> Option<Vec<PresetElement>> {
    let (_, elements) = PIVOT_STYLES.iter().find(|(preset, _)| *preset == name)?;
    Some(
        elements
            .iter()
            .map(|&(kind, size, index)| PresetElement {
                kind,
                size,
                dxf: &PIVOT_DXFS[usize::from(index)],
            })
            .collect(),
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn built_in_pivot_styles_resolve_by_name() {
        let light16 = pivot_style("PivotStyleLight16").unwrap();
        assert!(light16.iter().any(|element| element.kind == "headerRow"));
        assert!(pivot_style("PivotStyleDark28").is_some());
        assert!(pivot_style("NoSuchStyle").is_none());
        assert_eq!(PIVOT_STYLES.len(), 84);
    }
}
