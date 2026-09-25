//! BIFF8 Font/Format/Palette/XF -> the XLSX renderer style model.
//! [MS-XLS] 2.4.122, 2.4.126, 2.4.188, 2.4.353, 2.5.20,
//! 2.5.129; ECMA-376 Part 1 18.8. Cell XFs contain complete properties;
//! fAtr* controls later style updates, not inheritance during display.

use super::{decode_biff_chars, parse_biff_string, u16_at, u32_at, unsupported, Record};
use std::collections::BTreeMap;
mod color;
mod extensions;
mod font;

use color::ColorIdentity;
use font::{ResolvedFont, Script, Underline};

pub(super) const PATTERNS: [&str; 19] = [
    "none",
    "solid",
    "mediumGray",
    "darkGray",
    "lightGray",
    "darkHorizontal",
    "darkVertical",
    "darkDown",
    "darkUp",
    "darkGrid",
    "darkTrellis",
    "lightHorizontal",
    "lightVertical",
    "lightDown",
    "lightUp",
    "lightGrid",
    "lightTrellis",
    "gray125",
    "gray0625",
];
pub(super) const BORDERS: [&str; 14] = [
    "none",
    "thin",
    "medium",
    "dashed",
    "dotted",
    "thick",
    "double",
    "hair",
    "mediumDashed",
    "dashDot",
    "mediumDashDot",
    "dashDotDot",
    "mediumDashDotDot",
    "slantDashDot",
];

pub(super) struct Styles<'a> {
    fonts: Vec<&'a [u8]>,
    xfs: Vec<&'a [u8]>,
    formats: BTreeMap<u16, String>,
    palette: Option<&'a [u8]>,
    extensions: extensions::Extensions,
}

#[derive(Clone, Debug, PartialEq)]
pub(crate) struct NormalFont {
    pub name: String,
    pub size_points: f64,
    pub bold: bool,
    pub italic: bool,
}

pub(super) struct ResolvedStyleSheet {
    minimal: bool,
    // Rich-text FontIndex addresses only authored FONT records; XF-local
    // variants appended during resolution are not valid rich-run fonts.
    original_font_count: usize,
    fonts: Vec<ResolvedStyleFont>,
    fills: Vec<ResolvedFill>,
    borders: Vec<ResolvedBorder>,
    xfs: Vec<ResolvedXf>,
    formats: BTreeMap<u16, String>,
    /// Conditional-formatting and table differential formats.
    dxfs: Vec<xlsx_model::Dxf>,
    /// XFExt gradient fills by XF index; each replaces its XF's palette
    /// pattern fill in the model.
    gradients: Vec<(usize, ResolvedGradient)>,
}

struct ResolvedGradient {
    path: bool,
    degree: f64,
    left: f64,
    right: f64,
    top: f64,
    bottom: f64,
    stops: Vec<(f64, String)>,
}

#[derive(Clone)]
struct ResolvedStyleFont {
    font: ResolvedFont,
    color: ColorIdentity,
}

#[derive(Clone, PartialEq, Eq, PartialOrd, Ord)]
enum ResolvedFill {
    None,
    Gray125,
    Pattern {
        pattern: &'static str,
        foreground: ColorIdentity,
        background: ColorIdentity,
    },
}

#[derive(Clone, PartialEq, Eq, PartialOrd, Ord)]
struct ResolvedEdge {
    style: &'static str,
    color: ColorIdentity,
}

#[derive(Clone, PartialEq, Eq, PartialOrd, Ord)]
struct ResolvedBorder {
    seed: bool,
    diagonal_down: bool,
    diagonal_up: bool,
    left: Option<ResolvedEdge>,
    right: Option<ResolvedEdge>,
    top: Option<ResolvedEdge>,
    bottom: Option<ResolvedEdge>,
    diagonal: Option<ResolvedEdge>,
}

#[derive(Clone)]
struct ResolvedXf {
    num_fmt_id: u16,
    font_id: usize,
    fill_id: usize,
    border_id: usize,
    horizontal: &'static str,
    vertical: &'static str,
    wrap_text: bool,
    text_rotation: u8,
    indent: u16,
    shrink_to_fit: bool,
    reading_order: u8,
}

#[derive(Clone, PartialEq, Eq, PartialOrd, Ord)]
struct FontKey {
    name: String,
    size_twips: u16,
    color: ColorIdentity,
    family: u8,
    charset: u8,
    bold: bool,
    italic: bool,
    strike: bool,
    outline: bool,
    shadow: bool,
    condense: bool,
    extend: bool,
    underline: Underline,
    script: Script,
}

impl ResolvedStyleFont {
    fn key(&self) -> FontKey {
        FontKey {
            name: self.font.name.clone(),
            size_twips: self.font.size_twips,
            color: self.color,
            family: self.font.family,
            charset: self.font.charset,
            bold: self.font.weight == 700,
            italic: self.font.italic,
            strike: self.font.strike,
            outline: self.font.outline,
            shadow: self.font.shadow,
            condense: self.font.condense,
            extend: self.font.extend,
            underline: self.font.underline,
            script: self.font.script,
        }
    }
}

impl ResolvedStyleSheet {
    pub(super) fn default_font(&self) -> Option<(&str, f64)> {
        let font = self.fonts.get(self.xfs.first()?.font_id)?;
        Some((&font.font.name, f64::from(font.font.size_twips) / 20.0))
    }
    pub(super) fn into_model(self) -> xlsx_model::Styles {
        let gradients = self.gradients;
        // The synthesized minimal stylesheet has no BIFF FONT record, so its
        // font carries no authored bCharSet.
        let minimal = self.minimal;
        let mut model = xlsx_model::Styles {
            fonts: self
                .fonts
                .into_iter()
                .map(|v| {
                    let mut font = v.font.model(v.color.model());
                    if minimal {
                        font.charset = None;
                    }
                    font
                })
                .collect(),
            fills: self.fills.into_iter().map(fill_model).collect(),
            borders: self.borders.into_iter().map(border_model).collect(),
            cell_xfs: self.xfs.into_iter().map(xf_model).collect(),
            num_fmts: self
                .formats
                .into_iter()
                .map(|(num_fmt_id, format_code)| xlsx_model::NumFmt {
                    num_fmt_id: num_fmt_id.into(),
                    format_code,
                })
                .collect(),
            dxfs: self.dxfs,
        };
        // ECMA-376 18.8.24 gradientFill: one fill per XF gradient, which the
        // XF then uses instead of its palette pattern fallback.
        for (index, gradient) in gradients {
            let Some(xf) = model.cell_xfs.get_mut(index) else {
                continue;
            };
            xf.fill_id = model.fills.len() as u32;
            model.fills.push(xlsx_model::Fill {
                gradient: Some(xlsx_model::GradientFillSpec {
                    gradient_type: if gradient.path { "path" } else { "linear" }.into(),
                    degree: gradient.degree,
                    left: gradient.left,
                    right: gradient.right,
                    top: gradient.top,
                    bottom: gradient.bottom,
                    stops: gradient
                        .stops
                        .into_iter()
                        .map(|(position, color)| xlsx_model::GradientStopSpec { position, color })
                        .collect(),
                }),
                ..Default::default()
            });
        }
        model
    }

    pub(super) fn set_dxfs(&mut self, dxfs: Vec<xlsx_model::Dxf>) {
        self.dxfs = dxfs;
    }

    pub(super) fn into_model_bounded(
        self,
        budget: &mut usize,
    ) -> Result<xlsx_model::Styles, String> {
        let mut bytes = self
            .fonts
            .len()
            .checked_mul(std::mem::size_of::<xlsx_model::Font>())
            .and_then(|n| n.checked_add(self.fills.len() * std::mem::size_of::<xlsx_model::Fill>()))
            .and_then(|n| {
                n.checked_add(self.borders.len() * std::mem::size_of::<xlsx_model::Border>())
            })
            .and_then(|n| n.checked_add(self.xfs.len() * std::mem::size_of::<xlsx_model::CellXf>()))
            .and_then(|n| {
                n.checked_add(self.formats.len() * std::mem::size_of::<xlsx_model::NumFmt>())
            })
            .ok_or_else(|| unsupported("XLS style model byte budget exceeded"))?;
        for font in &self.fonts {
            let underline = match font.font.underline {
                Underline::None | Underline::Single => 0,
                value => value.model_value().len(),
            };
            let script = match font.font.script {
                Script::Baseline => 0,
                value => value.model_value().len(),
            };
            bytes = bytes
                .checked_add(font.font.name.len())
                .and_then(|n| n.checked_add(7))
                .and_then(|n| n.checked_add(underline + script))
                .ok_or_else(|| unsupported("XLS style model byte budget exceeded"))?;
        }
        for fill in &self.fills {
            let owned = match fill {
                ResolvedFill::None => "none".len(),
                ResolvedFill::Gray125 => "gray125".len(),
                ResolvedFill::Pattern { pattern, .. } => pattern.len() + 14,
            };
            bytes = bytes
                .checked_add(owned)
                .ok_or_else(|| unsupported("XLS style model byte budget exceeded"))?;
        }
        for border in &self.borders {
            for edge in [
                &border.left,
                &border.right,
                &border.top,
                &border.bottom,
                &border.diagonal,
            ]
            .into_iter()
            .flatten()
            {
                bytes = bytes
                    .checked_add(edge.style.len() + 7)
                    .ok_or_else(|| unsupported("XLS style model byte budget exceeded"))?;
            }
        }
        for border in &self.borders {
            // border_model clones the projected diagonal for diagonal_up.
            if border.diagonal_up {
                if let Some(edge) = &border.diagonal {
                    bytes = bytes
                        .checked_add(edge.style.len() + 7)
                        .ok_or_else(|| unsupported("XLS style model byte budget exceeded"))?;
                }
            }
        }
        for xf in &self.xfs {
            bytes = bytes
                .checked_add(xf.horizontal.len() + xf.vertical.len())
                .ok_or_else(|| unsupported("XLS style model byte budget exceeded"))?;
        }
        for format in self.formats.values() {
            bytes = bytes
                .checked_add(format.capacity())
                .ok_or_else(|| unsupported("XLS style model byte budget exceeded"))?;
        }
        for (_, gradient) in &self.gradients {
            bytes = bytes
                .checked_add(
                    std::mem::size_of::<xlsx_model::Fill>()
                        + std::mem::size_of::<xlsx_model::GradientFillSpec>()
                        + gradient.stops.len()
                            * (std::mem::size_of::<xlsx_model::GradientStopSpec>() + 7),
                )
                .ok_or_else(|| unsupported("XLS style model byte budget exceeded"))?;
        }
        for dxf in &self.dxfs {
            bytes = bytes
                .checked_add(dxf_bytes(dxf))
                .ok_or_else(|| unsupported("XLS style model byte budget exceeded"))?;
        }
        *budget = budget
            .checked_sub(bytes)
            .ok_or_else(|| unsupported("XLS style model byte budget exceeded"))?;
        Ok(self.into_model())
    }

    fn run_font(&self, index: u16) -> Result<&ResolvedStyleFont, String> {
        let offset = usize::from(index - u16::from(index > 4));
        self.fonts
            .get(offset)
            .filter(|_| index != 4 && offset < self.original_font_count)
            .ok_or_else(|| unsupported("BIFF rich-text font index out of range"))
    }

    pub(super) fn validate_run_font(&self, index: u16) -> Result<(), String> {
        self.run_font(index).map(|_| ())
    }

    pub(super) fn run_font_model(
        &self,
        index: u16,
        budget: &mut usize,
    ) -> Result<xlsx_model::RunFont, String> {
        let value = self.run_font(index)?;
        let underline_style = match value.font.underline {
            Underline::None | Underline::Single => None,
            underline => Some(underline.model_value()),
        };
        let vert_align = match value.font.script {
            Script::Baseline => None,
            script => Some(script.model_value()),
        };
        // The owning Run Vec slot already includes its inline RunFont. Charge
        // only allocations owned behind that value, before materializing them.
        let color_bytes = usize::from(!matches!(value.color, ColorIdentity::Auto)) * 7;
        let owned_bytes = value
            .font
            .name
            .len()
            .checked_add(color_bytes)
            .and_then(|n| n.checked_add(underline_style.map_or(0, str::len)))
            .and_then(|n| n.checked_add(vert_align.map_or(0, str::len)))
            .ok_or_else(|| "OUTPUT_TOO_LARGE".to_string())?;
        *budget = budget
            .checked_sub(owned_bytes)
            .ok_or_else(|| "OUTPUT_TOO_LARGE".to_string())?;
        let color = value.color.model();
        Ok(xlsx_model::RunFont {
            bold: value.font.weight == 700,
            italic: value.font.italic,
            underline: value.font.underline != Underline::None,
            strike: value.font.strike,
            size: Some(f64::from(value.font.size_twips) / 20.0),
            color,
            name: Some(value.font.name.clone()),
            underline_style: underline_style.map(str::to_owned),
            vert_align: vert_align.map(str::to_owned),
        })
    }
}

/// Drawing-shape text font resolved from a BIFF Font record (MS-XLS 2.4.122)
/// selected by a TxO formatting run's FontIndex (2.5.129).
#[derive(Debug, Clone)]
pub(super) struct ShapeFont {
    pub name: String,
    pub size_twips: u16,
    pub weight: u16,
    pub italic: bool,
    pub underline: bool,
    pub strike: bool,
    /// Superscript/subscript, outline, shadow, condense or extend.
    pub other_effects: bool,
    /// Resolved palette color; `None` for the automatic color (0x7FFF).
    pub color: Option<String>,
    pub automatic_color: bool,
}

/// Chart text font resolved from a BIFF Font record.
#[derive(Debug, Clone)]
pub(super) struct ChartFont {
    pub name: String,
    pub size_twips: u16,
    pub bold: bool,
    pub italic: bool,
    pub color: Option<String>,
}

impl<'a> Styles<'a> {
    /// MS-XLS 2.2.6.1.2.2: Normal references XF zero, not FONT zero.
    /// Return no measurement request for font variants we cannot reproduce.
    pub(super) fn normal_font(&self) -> Option<NormalFont> {
        let xf = self.xfs.first()?;
        if u16_at(xf, 4).ok()? & 4 == 0 {
            return None;
        }
        let index = u16_at(xf, 0).ok()?;
        if index == 4 {
            return None;
        }
        let font = self.fonts.get(usize::from(index - u16::from(index > 4)))?;
        if font.len() < 16 || font[2] & 0xf0 != 0 || u16_at(font, 8).ok()? != 0 {
            return None;
        }
        let weight = u16_at(font, 6).ok()?;
        let size = u16_at(font, 0).ok()?;
        if !matches!(weight, 400 | 700) || size == 0 {
            return None;
        }
        let (name, _) =
            decode_biff_chars(font, 16, usize::from(font[14]), font[15] & 1 != 0).ok()?;
        if name.is_empty() {
            return None;
        }
        Some(NormalFont {
            name,
            size_points: f64::from(size) / 20.0,
            bold: weight == 700,
            italic: font[2] & 2 != 0,
        })
    }
    pub fn parse(records: &[Record<'a>]) -> Result<Self, String> {
        let mut styles = Self {
            fonts: vec![],
            xfs: vec![],
            formats: BTreeMap::new(),
            palette: None,
            extensions: extensions::Extensions::default(),
        };
        for record in records.iter().take_while(|r| r.kind != super::EOF) {
            match record.kind {
                0x0031 => {
                    if styles.fonts.len() >= 1022 {
                        return Err(unsupported("too many BIFF fonts"));
                    }
                    styles.fonts.push(record.data);
                }
                0x00e0 => {
                    if record.data.len() != 20 || styles.xfs.len() >= 4096 {
                        return Err(unsupported("invalid or excessive BIFF XF records"));
                    }
                    styles.xfs.push(record.data);
                }
                0x041e => {
                    let id = u16_at(record.data, 0)?;
                    let (value, _) = parse_biff_string(&record.data[2..])?;
                    if styles.formats.insert(id, value).is_some() {
                        return Err(unsupported("duplicate BIFF number format"));
                    }
                }
                0x0092 => {
                    if u16_at(record.data, 0)? != 56 || record.data.len() != 226 {
                        return Err(unsupported("invalid BIFF palette"));
                    }
                    styles.palette = Some(&record.data[2..]);
                }
                _ => {}
            }
        }
        styles.extensions = extensions::Extensions::parse(records, &styles.xfs)?;
        Ok(styles)
    }

    /// Whether an XF extension carries formatting the model cannot resolve.
    pub(super) fn extensions_unrepresented(&self) -> bool {
        self.extensions.unrepresented
    }

    pub fn validate_xf(&self, index: u16) -> Result<(), String> {
        if usize::from(index) >= self.xfs.len().max(1) {
            return Err(unsupported("BIFF cell XF index out of range"));
        }
        Ok(())
    }

    /// Number of Font records in the Globals Substream (FontX indexing, 2.4.123).
    pub(super) fn font_count(&self) -> usize {
        self.fonts.len()
    }

    /// Decode a Font record (2.4.122) for chart text: name, twip size,
    /// weight/italic and palette color.
    pub(super) fn chart_font(&self, data: &[u8]) -> Option<ChartFont> {
        let font = ResolvedFont::decode(data).ok()?;
        Some(ChartFont {
            name: font.name,
            size_twips: font.size_twips,
            bold: font.weight >= 700,
            italic: font.italic,
            color: self.chart_color(font.color_index),
        })
    }

    /// A shape text run's font by FontIndex: 4 is reserved and indices above
    /// it are one-based (MS-XLS 2.5.129).
    pub(super) fn shape_font(&self, index: u16) -> Result<ShapeFont, String> {
        let data = self
            .fonts
            .get(usize::from(index - u16::from(index > 4)))
            .filter(|_| index != 4)
            .ok_or_else(|| unsupported("BIFF shape text font index out of range"))?;
        let font = ResolvedFont::decode(data)?;
        let automatic_color = matches!(self.color(font.color_index), ColorIdentity::Auto);
        Ok(ShapeFont {
            color: self.chart_color(font.color_index),
            automatic_color,
            size_twips: font.size_twips,
            weight: font.weight,
            italic: font.italic,
            underline: font.underline != Underline::None,
            strike: font.strike,
            other_effects: font.script != Script::Baseline
                || font.outline
                || font.shadow
                || font.condense
                || font.extend,
            name: font.name,
        })
    }

    /// FontX.iFont (2.4.123) one-based index into the global Font records.
    pub(super) fn global_font(&self, index: u16) -> Option<ChartFont> {
        let data = *self.fonts.get(usize::from(index).checked_sub(1)?)?;
        self.chart_font(data)
    }

    /// Chart element color for an Icv (MS-XLS 2.5.161): the workbook Palette
    /// record when it overrides the index, else the built-in indexed palette.
    pub(super) fn chart_color(&self, index: u16) -> Option<String> {
        use ooxml_common::spreadsheet_color::{resolve_color, SpreadsheetColor};
        match self.color(index) {
            ColorIdentity::Argb(argb) => resolve_color(SpreadsheetColor::Argb(argb), None, &[]),
            ColorIdentity::Indexed(index) if index < 64 => {
                resolve_color(SpreadsheetColor::Indexed(u32::from(index)), None, &[])
            }
            _ => None,
        }
    }

    /// An IcvXF/IcvFont color as the XLSX model resolves cell-style colors.
    pub(super) fn icv_model(&self, index: u16) -> Option<String> {
        self.color(index).model()
    }

    /// The FORMAT record (2.4.126) id whose string equals `code`.
    pub(super) fn format_id(&self, code: &str) -> Option<u16> {
        self.formats
            .iter()
            .find_map(|(id, value)| (value == code).then_some(*id))
    }

    /// The FORMAT record string for `id`, if the workbook defines one.
    pub(super) fn format_code(&self, id: u16) -> Option<String> {
        self.formats.get(&id).cloned()
    }

    fn color(&self, index: u16) -> ColorIdentity {
        if let Some(palette) = self.palette {
            if (8..64).contains(&index) {
                let offset = usize::from(index - 8) * 4;
                return ColorIdentity::Argb([
                    0xff,
                    palette[offset],
                    palette[offset + 1],
                    palette[offset + 2],
                ]);
            }
        }
        if index == 0x7fff {
            ColorIdentity::Auto
        } else {
            ColorIdentity::Indexed(index)
        }
    }

    pub(super) fn resolve(&self) -> Result<ResolvedStyleSheet, String> {
        if self.xfs.is_empty() && self.fonts.is_empty() {
            return Ok(minimal_resolved());
        }
        let mut fonts: Vec<ResolvedStyleFont> = Vec::new();
        for font in &self.fonts {
            let resolved = ResolvedFont::decode(font)?;
            let color = self.color(resolved.color_index);
            fonts.push(ResolvedStyleFont {
                font: resolved,
                color,
            });
        }
        if fonts.is_empty() {
            return Err(unsupported("BIFF styles reference missing fonts"));
        }
        let original_font_count = fonts.len();
        let mut font_ids = BTreeMap::new();
        for (id, font) in fonts.iter().enumerate() {
            font_ids.entry(font.key()).or_insert(id);
        }
        // Keep original font indices stable for shared-string rich runs. XF-local
        // color overrides append an interned variant, never mutate a shared font.
        let mut fills = vec![ResolvedFill::None, ResolvedFill::Gray125];
        let mut fill_ids = BTreeMap::from([(ResolvedFill::None, 0), (ResolvedFill::Gray125, 1)]);
        let mut borders = vec![ResolvedBorder::seed()];
        let mut border_ids = BTreeMap::from([(ResolvedBorder::seed(), 0)]);
        let mut xfs = Vec::new();
        for (index, data) in self.xfs.iter().enumerate() {
            let ifnt = u16_at(data, 0)?;
            let mut font = usize::from(ifnt - u16::from(ifnt > 4));
            if ifnt == 4 || font >= self.fonts.len() {
                return Err(unsupported("BIFF XF font index out of range"));
            }
            if let Some(color) = self.extensions.color(index, 13) {
                let resolved = ResolvedFont::decode(self.fonts[font])?;
                let variant = ResolvedStyleFont {
                    font: resolved,
                    color,
                };
                let key = variant.key();
                font = if let Some(id) = font_ids.get(&key) {
                    *id
                } else {
                    let id = fonts.len();
                    fonts.push(variant);
                    font_ids.insert(key, id);
                    id
                };
            }
            let color = |property, fallback| -> ColorIdentity {
                self.extensions
                    .color(index, property)
                    .unwrap_or_else(|| self.color(fallback))
            };
            let b1 = u32_at(data, 10)?;
            let b2 = u32_at(data, 14)?;
            let colors = u16_at(data, 18)?;
            let pattern = PATTERNS
                .get((b2 >> 26) as usize)
                .ok_or_else(|| unsupported("invalid BIFF fill pattern"))?;
            let fill = if *pattern == "none" {
                ResolvedFill::None
            } else {
                ResolvedFill::Pattern {
                    pattern,
                    foreground: color(4, colors & 127),
                    background: color(5, (colors >> 7) & 127),
                }
            };
            let fill_id = intern_typed(fill, &mut fills, &mut fill_ids);
            let mut edges = Vec::with_capacity(5);
            for (style, palette_color, property) in [
                (b1 & 15, (b1 >> 16) & 127, 9),
                ((b1 >> 4) & 15, (b1 >> 23) & 127, 10),
                ((b1 >> 8) & 15, b2 & 127, 7),
                ((b1 >> 12) & 15, (b2 >> 7) & 127, 8),
                ((b2 >> 21) & 15, (b2 >> 14) & 127, 11),
            ] {
                let style = BORDERS
                    .get(style as usize)
                    .ok_or_else(|| unsupported("invalid BIFF border style"))?;
                edges.push((*style != "none").then(|| ResolvedEdge {
                    style,
                    color: color(property, palette_color as u16),
                }));
            }
            let border_id = intern_typed(
                ResolvedBorder {
                    seed: false,
                    diagonal_down: b1 >> 30 & 1 != 0,
                    diagonal_up: b1 >> 31 & 1 != 0,
                    left: edges[0].clone(),
                    right: edges[1].clone(),
                    top: edges[2].clone(),
                    bottom: edges[3].clone(),
                    diagonal: edges[4].clone(),
                },
                &mut borders,
                &mut border_ids,
            );
            let horizontal = [
                "general",
                "left",
                "center",
                "right",
                "fill",
                "justify",
                "centerContinuous",
                "distributed",
            ][usize::from(data[6] & 7)];
            let vertical = ["top", "center", "bottom", "justify", "distributed"]
                .get(usize::from((data[6] >> 4) & 7))
                .ok_or_else(|| unsupported("invalid BIFF vertical alignment"))?;
            if (181..255).contains(&data[7]) || data[8] >> 6 > 2 {
                return Err(unsupported("invalid BIFF text alignment"));
            }
            let indent = self
                .extensions
                .indent(index)
                .unwrap_or(u16::from(data[8] & 15));
            let reading = data[8] >> 6;
            xfs.push(ResolvedXf {
                num_fmt_id: u16_at(data, 2)?,
                font_id: font,
                fill_id,
                border_id,
                horizontal,
                vertical,
                wrap_text: data[6] >> 3 & 1 != 0,
                text_rotation: data[7],
                indent,
                shrink_to_fit: data[8] >> 4 & 1 != 0,
                reading_order: reading,
            });
        }
        if xfs.is_empty() {
            xfs.push(ResolvedXf::default_xf());
        }
        let mut gradients = Vec::new();
        for index in 0..self.xfs.len() {
            let Some(gradient) = self.extensions.gradient(index) else {
                continue;
            };
            let mut stops = Vec::with_capacity(gradient.stops.len());
            for (position, color) in &gradient.stops {
                use ooxml_common::spreadsheet_color::{resolve_color, SpreadsheetColor};
                let color = match *color {
                    extensions::StopColor::Argb(argb) => {
                        resolve_color(SpreadsheetColor::Argb(argb), None, &[])
                    }
                    extensions::StopColor::Indexed(icv, tint) => match self.color(icv) {
                        ColorIdentity::Argb(argb) => resolve_color(
                            SpreadsheetColor::Argb(argb),
                            (tint != 0.0).then_some(tint),
                            &[],
                        ),
                        ColorIdentity::Indexed(index) if index < 64 => resolve_color(
                            SpreadsheetColor::Indexed(index.into()),
                            (tint != 0.0).then_some(tint),
                            &[],
                        ),
                        _ => None,
                    },
                }
                .ok_or_else(|| unsupported("unsupported BIFF gradient stop color"))?;
                stops.push((*position, color));
            }
            gradients.push((
                index,
                ResolvedGradient {
                    path: gradient.path,
                    degree: gradient.degree,
                    left: gradient.left,
                    right: gradient.right,
                    top: gradient.top,
                    bottom: gradient.bottom,
                    stops,
                },
            ));
        }
        Ok(ResolvedStyleSheet {
            minimal: false,
            original_font_count,
            fonts,
            fills,
            borders,
            xfs,
            formats: self.formats.clone(),
            dxfs: Vec::new(),
            gradients,
        })
    }

    #[cfg(test)]
    fn model(&self) -> Result<xlsx_model::Styles, String> {
        Ok(self.resolve()?.into_model())
    }
}

impl ResolvedBorder {
    fn seed() -> Self {
        Self {
            seed: true,
            diagonal_down: false,
            diagonal_up: false,
            left: None,
            right: None,
            top: None,
            bottom: None,
            diagonal: None,
        }
    }
}

impl ResolvedXf {
    fn default_xf() -> Self {
        Self {
            num_fmt_id: 0,
            font_id: 0,
            fill_id: 0,
            border_id: 0,
            horizontal: "",
            vertical: "",
            wrap_text: false,
            text_rotation: 0,
            indent: 0,
            shrink_to_fit: false,
            reading_order: 0,
        }
    }
}

pub(super) fn minimal_resolved() -> ResolvedStyleSheet {
    let font = ResolvedFont::minimal_calibri();
    ResolvedStyleSheet {
        minimal: true,
        original_font_count: 0,
        fonts: vec![ResolvedStyleFont {
            font,
            color: ColorIdentity::Auto,
        }],
        fills: vec![ResolvedFill::None, ResolvedFill::Gray125],
        borders: vec![ResolvedBorder::seed()],
        xfs: vec![ResolvedXf::default_xf()],
        formats: BTreeMap::new(),
        dxfs: Vec::new(),
        gradients: Vec::new(),
    }
}

fn intern_typed<T: Clone + Ord>(
    value: T,
    values: &mut Vec<T>,
    ids: &mut BTreeMap<T, usize>,
) -> usize {
    if let Some(id) = ids.get(&value) {
        return *id;
    }
    let id = values.len();
    values.push(value.clone());
    ids.insert(value, id);
    id
}

/// Retained model bytes of one conditional-formatting dxf.
fn dxf_bytes(dxf: &xlsx_model::Dxf) -> usize {
    let color = |value: &Option<String>| value.as_ref().map_or(0, String::len);
    let edge = |value: &Option<xlsx_model::BorderEdge>| {
        value
            .as_ref()
            .map_or(0, |e| e.style.len() + color(&e.color))
    };
    std::mem::size_of::<xlsx_model::Dxf>()
        + dxf.font.as_ref().map_or(0, |f| {
            std::mem::size_of::<xlsx_model::Font>()
                + color(&f.color)
                + color(&f.name)
                + color(&f.underline_style)
                + color(&f.vert_align)
        })
        + dxf.fill.as_ref().map_or(0, |f| {
            std::mem::size_of::<xlsx_model::Fill>()
                + f.pattern_type.len()
                + color(&f.fg_color)
                + color(&f.bg_color)
        })
        + dxf.border.as_ref().map_or(0, |b| {
            std::mem::size_of::<xlsx_model::Border>()
                + edge(&b.left)
                + edge(&b.right)
                + edge(&b.top)
                + edge(&b.bottom)
                + edge(&b.diagonal_up)
                + edge(&b.diagonal_down)
        })
        + dxf.num_fmt.as_ref().map_or(0, |n| {
            std::mem::size_of::<xlsx_model::NumFmt>() + n.format_code.len()
        })
}

fn fill_model(value: ResolvedFill) -> xlsx_model::Fill {
    match value {
        ResolvedFill::None => xlsx_model::Fill {
            pattern_type: "none".into(),
            ..Default::default()
        },
        ResolvedFill::Gray125 => xlsx_model::Fill {
            pattern_type: "gray125".into(),
            ..Default::default()
        },
        ResolvedFill::Pattern {
            pattern,
            foreground,
            background,
        } => xlsx_model::Fill {
            pattern_type: pattern.into(),
            fg_color: foreground.model(),
            bg_color: background.model(),
            gradient: None,
        },
    }
}

fn edge_model(value: Option<ResolvedEdge>) -> Option<xlsx_model::BorderEdge> {
    value.map(|e| xlsx_model::BorderEdge {
        style: e.style.into(),
        color: e.color.model(),
    })
}

fn border_model(value: ResolvedBorder) -> xlsx_model::Border {
    let diagonal = edge_model(value.diagonal);
    xlsx_model::Border {
        left: edge_model(value.left),
        right: edge_model(value.right),
        top: edge_model(value.top),
        bottom: edge_model(value.bottom),
        diagonal_up: value.diagonal_up.then(|| diagonal.clone()).flatten(),
        diagonal_down: value.diagonal_down.then_some(diagonal).flatten(),
        horizontal: None,
        vertical: None,
    }
}

fn xf_model(value: ResolvedXf) -> xlsx_model::CellXf {
    xlsx_model::CellXf {
        font_id: value.font_id as u32,
        fill_id: value.fill_id as u32,
        border_id: value.border_id as u32,
        num_fmt_id: value.num_fmt_id.into(),
        // BIFF8 XF alc 0 is General, the SpreadsheetML default (ECMA-376
        // §18.8.1); the shared model represents it as omission so the renderer
        // applies the value-type rule of §18.18.40.
        align_h: (!matches!(value.horizontal, "" | "general")).then(|| value.horizontal.into()),
        align_v: (!value.vertical.is_empty()).then(|| value.vertical.into()),
        wrap_text: value.wrap_text,
        indent: (value.indent != 0).then_some(value.indent.into()),
        text_rotation: (value.text_rotation != 0).then_some(value.text_rotation.into()),
        shrink_to_fit: value.shrink_to_fit,
        reading_order: (value.reading_order != 0).then_some(value.reading_order.into()),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A model budget that never binds.
    fn unbounded() -> usize {
        usize::MAX
    }
    fn crc(bytes: impl Iterator<Item = u8>) -> u32 {
        let mut value = 0u32;
        for byte in bytes {
            value ^= u32::from(byte) << 24;
            for _ in 0..8 {
                value = (value << 1) ^ if value & 0x8000_0000 != 0 { 0xaf } else { 0 };
            }
        }
        value
    }
    fn extended_indent(xfs: &[[u8; 20]], index: u16, value: u16) -> ([u8; 20], Vec<u8>) {
        let mut check = [0; 20];
        check[..2].copy_from_slice(&0x087cu16.to_le_bytes());
        check[14..16].copy_from_slice(&(xfs.len() as u16).to_le_bytes());
        check[16..].copy_from_slice(&crc(xfs.iter().flatten().copied()).to_le_bytes());
        let mut ext = vec![0; 20];
        ext[..2].copy_from_slice(&0x087du16.to_le_bytes());
        ext[14..16].copy_from_slice(&index.to_le_bytes());
        ext[18..20].copy_from_slice(&1u16.to_le_bytes());
        ext.extend_from_slice(&0x000fu16.to_le_bytes());
        ext.extend_from_slice(&6u16.to_le_bytes());
        ext.extend_from_slice(&value.to_le_bytes());
        (check, ext)
    }
    fn extended_color(
        xfs: &[[u8; 20]],
        index: u16,
        kind: u16,
        argb: [u8; 4],
    ) -> ([u8; 20], Vec<u8>) {
        let mut check = [0; 20];
        check[..2].copy_from_slice(&0x087cu16.to_le_bytes());
        check[14..16].copy_from_slice(&(xfs.len() as u16).to_le_bytes());
        check[16..].copy_from_slice(&crc(xfs.iter().flatten().copied()).to_le_bytes());
        let mut ext = vec![0; 20];
        ext[..2].copy_from_slice(&0x087du16.to_le_bytes());
        ext[14..16].copy_from_slice(&index.to_le_bytes());
        ext[18..20].copy_from_slice(&1u16.to_le_bytes());
        ext.extend_from_slice(&kind.to_le_bytes());
        ext.extend_from_slice(&20u16.to_le_bytes());
        ext.extend_from_slice(&[2, 0, 0, 0, argb[1], argb[2], argb[3], argb[0]]);
        ext.extend_from_slice(&[0; 8]);
        (check, ext)
    }
    fn font() -> Vec<u8> {
        let mut data = vec![0; 16];
        data[..2].copy_from_slice(&200u16.to_le_bytes());
        data[6..8].copy_from_slice(&400u16.to_le_bytes());
        data[14] = 1;
        data.push(b'F');
        data
    }
    #[test]
    fn bounded_model_accounts_for_font_variants_and_both_diagonals() {
        fn fixture() -> ResolvedStyleSheet {
            let mut source = minimal_resolved();
            source.fonts[0].font.underline = Underline::Double;
            source.fonts[0].font.script = Script::Superscript;
            source.borders[0].diagonal_up = true;
            source.borders[0].diagonal_down = true;
            source.borders[0].diagonal = Some(ResolvedEdge {
                style: "double",
                color: ColorIdentity::Auto,
            });
            source
        }
        let mut budget = usize::MAX;
        let model = fixture().into_model_bounded(&mut budget).unwrap();
        let required = usize::MAX - budget;
        assert_eq!(model.fonts[0].underline_style.as_deref(), Some("double"));
        assert_eq!(model.fonts[0].vert_align.as_deref(), Some("superscript"));
        assert!(model.borders[0].diagonal_up.is_some());
        assert!(model.borders[0].diagonal_down.is_some());
        let mut plain_budget = usize::MAX;
        minimal_resolved()
            .into_model_bounded(&mut plain_budget)
            .unwrap();
        assert_eq!(
            plain_budget - budget,
            "double".len() + "superscript".len() + 2 * ("double".len() + 7)
        );
        let mut exact = required;
        fixture().into_model_bounded(&mut exact).unwrap();
        assert_eq!(exact, 0);
        assert!(fixture().into_model_bounded(&mut (required - 1)).is_err());
    }

    #[test]
    fn minimal_native_model_has_the_synthesized_defaults() {
        let s = Styles {
            fonts: vec![],
            xfs: vec![],
            formats: BTreeMap::new(),
            palette: None,
            extensions: extensions::Extensions::default(),
        };
        let model = s.model().unwrap();
        assert_eq!(
            (
                model.fonts.len(),
                model.fills.len(),
                model.borders.len(),
                model.cell_xfs.len()
            ),
            (1, 2, 1, 1)
        );
        assert_eq!(model.fonts[0].name.as_deref(), Some("Calibri"));
        assert_eq!(model.fonts[0].size, 11.0);
        assert_eq!(model.fonts[0].color, None);
        // The synthesized font has no authored FONT bCharSet.
        assert_eq!(model.fonts[0].charset, None);
        assert_eq!(model.fills[0].pattern_type, "none");
        assert_eq!(model.fills[1].pattern_type, "gray125");
        assert_eq!(model.cell_xfs[0].align_v, None);
    }
    #[test]
    fn normal_measurement_resolves_style_xf_and_reserved_font_gap() {
        let first = font();
        let mut last = font();
        last[16] = b'Z';
        last[2] = 2;
        last[6..8].copy_from_slice(&700u16.to_le_bytes());
        let mut xf = [0; 20];
        xf[0] = 5;
        xf[4] = 4;
        let mut records: Vec<_> = [&first, &first, &first, &first, &last]
            .into_iter()
            .map(|f| Record {
                kind: 0x31,
                offset: 0,
                data: f,
            })
            .collect();
        records.push(Record {
            kind: 0xe0,
            offset: 0,
            data: &xf,
        });
        let s = Styles::parse(&records).unwrap();
        assert_eq!(
            s.normal_font(),
            Some(NormalFont {
                name: "Z".into(),
                size_points: 10.0,
                bold: true,
                italic: true
            })
        );
        for index in [4, 6] {
            let mut bad = xf;
            bad[0] = index;
            let s = Styles::parse(&[Record {
                kind: 0xe0,
                offset: 0,
                data: &bad,
            }])
            .unwrap();
            assert_eq!(s.normal_font(), None);
        }
        for (offset, value) in [(4, 0), (0, 4)] {
            let mut bad = xf;
            bad[offset] = value;
            let mut local = records[..5].to_vec();
            local.push(Record {
                kind: 0xe0,
                offset: 0,
                data: &bad,
            });
            assert_eq!(Styles::parse(&local).unwrap().normal_font(), None);
        }
    }

    #[test]
    fn unsupported_font_variants_do_not_request_guessed_measurements() {
        for (offset, value) in [(2, 64), (2, 128), (8, 1), (6, 0)] {
            let mut font = font();
            font[offset] = value;
            if offset == 6 {
                font[7] = 0;
            }
            let mut xf = [0; 20];
            xf[4] = 4;
            let s = Styles::parse(&[
                Record {
                    kind: 0x31,
                    offset: 0,
                    data: &font,
                },
                Record {
                    kind: 0xe0,
                    offset: 0,
                    data: &xf,
                },
            ])
            .unwrap();
            assert_eq!(s.normal_font(), None);
        }
    }
    #[test]
    fn checksum_bound_extended_rgb_colors_override_only_the_owned_xf() {
        let font = font();
        let mut xf = [0; 20];
        xf[17] = 6; // Solid fill and CellXF.fHasXFExt.
        let mut crc = [0; 20];
        crc[..2].copy_from_slice(&0x087cu16.to_le_bytes());
        crc[14..16].copy_from_slice(&16u16.to_le_bytes());
        crc[16..].copy_from_slice(&0x344d21a3u32.to_le_bytes());
        let mut ext = vec![0; 20];
        ext[..2].copy_from_slice(&0x087du16.to_le_bytes());
        ext[14] = 1;
        ext[18] = 2;
        for kind in [4u16, 13] {
            ext.extend_from_slice(&kind.to_le_bytes());
            ext.extend_from_slice(&20u16.to_le_bytes());
            ext.extend_from_slice(&[2, 0, 0, 0, 0x12, 0x34, 0x56, 0xff]);
            ext.extend_from_slice(&[0; 8]);
        }
        let mut records = vec![Record {
            kind: 0x31,
            offset: 0,
            data: &font,
        }];
        records.extend((0..16).map(|_| Record {
            kind: 0xe0,
            offset: 0,
            data: &xf,
        }));
        records.push(Record {
            kind: 0x87c,
            offset: 0,
            data: &crc,
        });
        records.push(Record {
            kind: 0x87d,
            offset: 0,
            data: &ext,
        });
        let model = Styles::parse(&records).unwrap().model().unwrap();
        let owned = &model.cell_xfs[1];
        assert_eq!(
            model.fills[owned.fill_id as usize].fg_color.as_deref(),
            Some("#123456")
        );
        assert_eq!(
            model.fills[model.cell_xfs[2].fill_id as usize]
                .fg_color
                .as_deref(),
            Some("#000000")
        );
        assert_eq!(model.fonts.len(), 2);
        assert_eq!(model.fonts[1].color.as_deref(), Some("#123456"));
        assert_eq!(owned.font_id, 1);
        assert_eq!(
            model.cell_xfs.iter().filter(|xf| xf.font_id == 1).count(),
            1
        );
    }
    #[test]
    fn extended_theme_colors_zero_to_three_follow_excels_light_dark_order() {
        const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
        const R: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
        const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        let theme = format!("<a:theme xmlns:a=\"{A}\"><a:themeElements><a:clrScheme name=\"T\"><a:dk1><a:srgbClr val=\"111111\"/></a:dk1><a:lt1><a:srgbClr val=\"EEEEEE\"/></a:lt1><a:dk2><a:srgbClr val=\"222222\"/></a:dk2><a:lt2><a:srgbClr val=\"DDDDDD\"/></a:lt2></a:clrScheme></a:themeElements></a:theme>");
        let package = super::super::theme::test_zip(&[
            ("_rels/.rels", format!("<Relationships xmlns=\"{R}\"><Relationship Id=\"main\" Type=\"{REL}/officeDocument\" Target=\"theme/manager.xml\"/></Relationships>")),
            ("theme/manager.xml", format!("<a:themeManager xmlns:a=\"{A}\"/>")),
            ("theme/_rels/manager.xml.rels", format!("<Relationships xmlns=\"{R}\"><Relationship Id=\"t\" Type=\"{REL}/theme\" Target=\"theme1.xml\"/></Relationships>")),
            ("theme/theme1.xml", theme),
        ]);
        let mut theme_record = vec![0; 16];
        theme_record[..2].copy_from_slice(&0x0896u16.to_le_bytes());
        theme_record.extend_from_slice(&package);
        let font = font();
        let mut xf = [0; 20];
        xf[17] = 6; // Solid fill and CellXF.fHasXFExt.
        let mut crc = [0; 20];
        crc[..2].copy_from_slice(&0x087cu16.to_le_bytes());
        crc[14..16].copy_from_slice(&16u16.to_le_bytes());
        crc[16..].copy_from_slice(&0x344d21a3u32.to_le_bytes());
        for (index, expected) in [
            (0u32, "EEEEEE"),
            (1, "111111"),
            (2, "DDDDDD"),
            (3, "222222"),
        ] {
            let mut ext = vec![0; 20];
            ext[..2].copy_from_slice(&0x087du16.to_le_bytes());
            ext[14] = 1;
            ext[18] = 1;
            ext.extend_from_slice(&4u16.to_le_bytes());
            ext.extend_from_slice(&20u16.to_le_bytes());
            ext.extend_from_slice(&[3, 0, 0, 0]);
            ext.extend_from_slice(&index.to_le_bytes());
            ext.extend_from_slice(&[0; 8]);
            let mut records = vec![
                Record {
                    kind: 0x0896,
                    offset: 0,
                    data: &theme_record,
                },
                Record {
                    kind: 0x31,
                    offset: 0,
                    data: &font,
                },
            ];
            records.extend((0..16).map(|_| Record {
                kind: 0xe0,
                offset: 0,
                data: &xf,
            }));
            records.push(Record {
                kind: 0x87c,
                offset: 0,
                data: &crc,
            });
            records.push(Record {
                kind: 0x87d,
                offset: 0,
                data: &ext,
            });
            let model = Styles::parse(&records).unwrap().model().unwrap();
            assert_eq!(
                model.fills[model.cell_xfs[1].fill_id as usize].fg_color,
                Some(format!("#{expected}")),
                "theme {index}"
            );
        }
    }

    #[test]
    fn extended_color_tint_uses_the_spreadsheetml_tint_algorithm() {
        let font = font();
        let mut xf = [0; 20];
        xf[17] = 6; // Solid fill and CellXF.fHasXFExt.
        let mut crc = [0; 20];
        crc[..2].copy_from_slice(&0x087cu16.to_le_bytes());
        crc[14..16].copy_from_slice(&16u16.to_le_bytes());
        crc[16..].copy_from_slice(&0x344d21a3u32.to_le_bytes());
        let mut ext = vec![0; 20];
        ext[..2].copy_from_slice(&0x087du16.to_le_bytes());
        ext[14] = 1;
        ext[18] = 1;
        ext.extend_from_slice(&4u16.to_le_bytes());
        ext.extend_from_slice(&20u16.to_le_bytes());
        // RGB 808080 with nTintShade 16383 (tint 0.5, lighten toward white).
        ext.extend_from_slice(&[2, 0]);
        ext.extend_from_slice(&16383i16.to_le_bytes());
        ext.extend_from_slice(&[0x80, 0x80, 0x80, 0xff]);
        ext.extend_from_slice(&[0; 8]);
        let mut records = vec![Record {
            kind: 0x31,
            offset: 0,
            data: &font,
        }];
        records.extend((0..16).map(|_| Record {
            kind: 0xe0,
            offset: 0,
            data: &xf,
        }));
        records.push(Record {
            kind: 0x87c,
            offset: 0,
            data: &crc,
        });
        records.push(Record {
            kind: 0x87d,
            offset: 0,
            data: &ext,
        });
        let model = Styles::parse(&records).unwrap().model().unwrap();
        // HLS luminance 0.502 -> 0.502 * 0.5 + 0.5 = 0.751 (ECMA-376 §18.8.19).
        assert_eq!(
            model.fills[model.cell_xfs[1].fill_id as usize]
                .fg_color
                .as_deref(),
            Some("#BFBFBF")
        );
    }
    #[test]
    fn checksum_bound_extended_indent_overrides_only_the_owned_cell_xf() {
        for value in [0u16, 15, 16, 250] {
            let font = font();
            let mut owned = [0u8; 20];
            owned[6] = 1; // left alignment
            owned[8] = 7; // distinguish the four-bit base cIndent
            owned[17] = 2; // CellXF.fHasXFExt
            let mut xfs = [owned; 16];
            xfs[0][4] = 4; // Normal StyleXF; its XFExt ownership is fStyle.
            xfs[0][17] = 0;
            let (check, ext) = extended_indent(&xfs, 1, value);
            let mut records = vec![Record {
                kind: 0x31,
                offset: 0,
                data: &font,
            }];
            records.extend(xfs.iter().map(|xf| Record {
                kind: 0xe0,
                offset: 0,
                data: xf,
            }));
            records.push(Record {
                kind: 0x087c,
                offset: 0,
                data: &check,
            });
            records.push(Record {
                kind: 0x087d,
                offset: 0,
                data: &ext,
            });
            let model = Styles::parse(&records).unwrap().model().unwrap();
            let indents: Vec<_> = model.cell_xfs.iter().map(|xf| xf.indent).collect();
            // The style XF plus 15 unextended cell XFs keep the base cIndent.
            assert_eq!(indents.iter().filter(|v| **v == Some(7)).count(), 15);
            assert_eq!(indents[1], (value != 0).then_some(u32::from(value)));
        }
    }
    #[test]
    fn extended_indent_requires_current_ownership_and_isolates_style_and_cell_xfs() {
        let font = font();
        for case in ["unowned", "stale", "cell", "style"] {
            let mut xf = [0u8; 20];
            xf[6] = 1;
            xf[8] = 7;
            let mut xfs = [xf; 16];
            xfs[0][4] = 4;
            xfs[0][8] = 0x87; // StyleXF cIndent=7, iReadOrder=2.
            if matches!(case, "stale" | "cell") {
                xfs[1][17] = 2;
            }
            let target = if case == "style" { 0 } else { 1 };
            let (mut check, ext) = extended_indent(&xfs, target, 250);
            if case == "stale" {
                check[16] ^= 1;
            }
            let mut records = vec![Record {
                kind: 0x31,
                offset: 0,
                data: &font,
            }];
            records.extend(xfs.iter().map(|xf| Record {
                kind: 0xe0,
                offset: 0,
                data: xf,
            }));
            records.push(Record {
                kind: 0x087c,
                offset: 0,
                data: &check,
            });
            records.push(Record {
                kind: 0x087d,
                offset: 0,
                data: &ext,
            });
            let model = Styles::parse(&records).unwrap().model().unwrap();
            let expected = match case {
                "cell" => 1,
                "style" => 1,
                _ => 0,
            };
            let count = |value| {
                model
                    .cell_xfs
                    .iter()
                    .filter(|xf| xf.indent == value)
                    .count()
            };
            assert_eq!(count(Some(250)), expected, "{case}");
            assert_eq!(
                model
                    .cell_xfs
                    .iter()
                    .filter(|xf| xf.reading_order == Some(2))
                    .count(),
                1,
                "{case}"
            );
        }
    }
    #[test]
    fn rejects_malformed_extended_indent() {
        let font = font();
        let mut xf = [0u8; 20];
        xf[17] = 2;
        let xfs = [xf; 16];
        for case in ["short", "long", "range", "duplicate"] {
            let (check, mut ext) = extended_indent(&xfs, 1, if case == "range" { 251 } else { 16 });
            match case {
                "short" => {
                    ext[22..24].copy_from_slice(&5u16.to_le_bytes());
                    ext.pop();
                }
                "long" => {
                    ext[22..24].copy_from_slice(&7u16.to_le_bytes());
                    ext.push(0);
                }
                "duplicate" => {
                    ext[18..20].copy_from_slice(&2u16.to_le_bytes());
                    ext.extend_from_within(20..26);
                }
                _ => {}
            }
            let mut records = vec![Record {
                kind: 0x31,
                offset: 0,
                data: &font,
            }];
            records.extend(xfs.iter().map(|xf| Record {
                kind: 0xe0,
                offset: 0,
                data: xf,
            }));
            records.push(Record {
                kind: 0x087c,
                offset: 0,
                data: &check,
            });
            records.push(Record {
                kind: 0x087d,
                offset: 0,
                data: &ext,
            });
            assert!(Styles::parse(&records).is_err(), "{case}");
        }
    }
    #[test]
    fn remaps_font_gap_and_deduplicates_repeated_fills() {
        let font = font();
        let mut xf = [0u8; 20];
        xf[0] = 5;
        let s = Styles {
            fonts: vec![&font; 5],
            xfs: vec![&xf; 100],
            formats: BTreeMap::new(),
            palette: None,
            extensions: extensions::Extensions::default(),
        };
        let model = s.model().unwrap();
        assert_eq!(
            (model.fonts.len(), model.fills.len(), model.borders.len()),
            (5, 2, 2)
        );
        assert_eq!(model.cell_xfs.len(), 100);
        assert_eq!(model.cell_xfs[0].font_id, 4);
        assert_eq!(model.cell_xfs[0].fill_id, 0);
        assert_eq!(model.cell_xfs[0].border_id, 1);
        // alc 0 (General) is the SpreadsheetML default and is omitted.
        assert_eq!(model.cell_xfs[0].align_h, None);
        assert_eq!(model.cell_xfs[0].align_v.as_deref(), Some("top"));
        assert_eq!(model.cell_xfs[0].indent, None);
        assert_eq!(model.cell_xfs[0].text_rotation, None);
        assert_eq!(model.cell_xfs[0].reading_order, None);
    }
    #[test]
    fn typed_font_projects_cell_and_run_font_properties() {
        let mut font = vec![0; 16];
        font[..2].copy_from_slice(&240u16.to_le_bytes());
        font[2] = 2 | 8 | 16 | 32 | 64 | 128;
        font[4..6].copy_from_slice(&10u16.to_le_bytes());
        font[6..8].copy_from_slice(&700u16.to_le_bytes());
        font[8..10].copy_from_slice(&1u16.to_le_bytes());
        font[10] = 0x21;
        font[11] = 3;
        font[12] = 0x80;
        font[14] = 3;
        font.extend_from_slice(b"A&B");
        let s = Styles {
            fonts: vec![&font],
            xfs: vec![],
            formats: BTreeMap::new(),
            palette: None,
            extensions: extensions::Extensions::default(),
        };
        let model = s.model().unwrap();
        assert_eq!(model.fonts.len(), 1);
        assert_eq!(model.fonts[0].name.as_deref(), Some("A&B"));
        assert_eq!(model.fonts[0].color.as_deref(), Some("#FF0000"));
        assert!(model.fonts[0].bold && model.fonts[0].italic && model.fonts[0].strike);
        assert_eq!(
            model.fonts[0].underline_style.as_deref(),
            Some("singleAccounting")
        );
        assert_eq!(model.fonts[0].vert_align.as_deref(), Some("superscript"));
        assert_eq!(model.fonts[0].charset, Some(0x80));
        let resolved = s.resolve().unwrap();
        // MS-XLS 2.5.129: FontIndex 4 is reserved.
        assert!(resolved.validate_run_font(4).is_err());
        let run = resolved.run_font_model(0, &mut unbounded()).unwrap();
        assert!(run.bold && run.italic && run.underline && run.strike);
        assert_eq!((run.size, run.name.as_deref()), (Some(12.0), Some("A&B")));
        assert_eq!(run.color.as_deref(), Some("#FF0000"));
        assert_eq!(run.underline_style.as_deref(), Some("singleAccounting"));
        assert_eq!(run.vert_align.as_deref(), Some("superscript"));

        let mut plain = font.clone();
        plain[2] = 0;
        plain[6..8].copy_from_slice(&650u16.to_le_bytes());
        plain[8..10].copy_from_slice(&0u16.to_le_bytes());
        plain[10] = 0;
        let plain_styles = Styles {
            fonts: vec![&plain],
            xfs: vec![],
            formats: BTreeMap::new(),
            palette: None,
            extensions: extensions::Extensions::default(),
        };
        // A plain run font states every property explicitly off, so it
        // resets the cell's font rather than inheriting from it.
        let run = plain_styles
            .resolve()
            .unwrap()
            .run_font_model(0, &mut unbounded())
            .unwrap();
        assert!(!run.bold && !run.italic && !run.underline && !run.strike);
        assert_eq!((run.underline_style, run.vert_align), (None, None));
    }
    #[test]
    fn resolves_custom_palette_and_preserves_automatic_font_color() {
        let mut palette = vec![0; 224];
        palette[8..12].copy_from_slice(&[0x12, 0x34, 0x56, 0]);
        let s = Styles {
            fonts: vec![],
            xfs: vec![],
            formats: BTreeMap::new(),
            palette: Some(&palette),
            extensions: extensions::Extensions::default(),
        };
        assert_eq!(s.color(10), ColorIdentity::Argb([0xff, 0x12, 0x34, 0x56]));
        assert_eq!(s.color(0x7fff), ColorIdentity::Auto);
        assert_eq!(s.color(65), ColorIdentity::Indexed(65));
        assert_eq!(s.color(10).model().as_deref(), Some("#123456"));
        assert_eq!(s.color(0x7fff).model(), None);
    }
    #[test]
    fn palette_and_parsed_extended_argb_share_typed_identity() {
        let mut font = font();
        font[4..6].copy_from_slice(&10u16.to_le_bytes());
        let mut owned = [0u8; 20];
        owned[17] = 2;
        let xfs = [owned; 16];
        let (check, ext) = extended_color(&xfs, 1, 13, [0xff, 0x12, 0x34, 0x56]);
        let mut palette = vec![0; 226];
        palette[..2].copy_from_slice(&56u16.to_le_bytes());
        palette[10..14].copy_from_slice(&[0x12, 0x34, 0x56, 0]);
        let mut records = vec![Record {
            kind: 0x31,
            offset: 0,
            data: &font,
        }];
        records.extend(xfs.iter().map(|xf| Record {
            kind: 0xe0,
            offset: 0,
            data: xf,
        }));
        records.extend([
            Record {
                kind: 0x0092,
                offset: 0,
                data: &palette,
            },
            Record {
                kind: 0x087c,
                offset: 0,
                data: &check,
            },
            Record {
                kind: 0x087d,
                offset: 0,
                data: &ext,
            },
        ]);
        let styles = Styles::parse(&records).unwrap();
        assert_eq!(styles.color(10), styles.extensions.color(1, 13).unwrap());
        let model = styles.model().unwrap();
        // The extension resolves to the same typed ARGB identity as the
        // palette-backed base font, so interning keeps one font.
        assert_eq!(model.fonts.len(), 1);
        assert_eq!(model.fonts[0].color.as_deref(), Some("#123456"));
        assert!(model.cell_xfs.iter().all(|xf| xf.font_id == 0));
    }
    #[test]
    fn font_override_interns_equivalent_weights_without_removing_source_slots() {
        let mut normal = font();
        normal[4..6].copy_from_slice(&10u16.to_le_bytes());
        let mut other = normal.clone();
        other[6..8].copy_from_slice(&650u16.to_le_bytes());
        let mut xf = [0u8; 20];
        xf[0] = 1;
        xf[17] = 2;
        let mut xfs = [xf; 16];
        xfs[0][0] = 0;
        let (check, ext) = extended_color(&xfs, 1, 13, [0xff, 0x12, 0x34, 0x56]);
        let mut palette = vec![0; 226];
        palette[..2].copy_from_slice(&56u16.to_le_bytes());
        palette[10..14].copy_from_slice(&[0x12, 0x34, 0x56, 0]);
        let mut records = vec![
            Record {
                kind: 0x31,
                offset: 0,
                data: &normal,
            },
            Record {
                kind: 0x31,
                offset: 0,
                data: &other,
            },
        ];
        records.extend(xfs.iter().map(|xf| Record {
            kind: 0xe0,
            offset: 0,
            data: xf,
        }));
        records.extend([
            Record {
                kind: 0x92,
                offset: 0,
                data: &palette,
            },
            Record {
                kind: 0x87c,
                offset: 0,
                data: &check,
            },
            Record {
                kind: 0x87d,
                offset: 0,
                data: &ext,
            },
        ]);
        let resolved = Styles::parse(&records).unwrap().resolve().unwrap();
        assert_eq!(resolved.fonts.len(), 2);
        assert_eq!(resolved.fonts[1].font.weight, 650);
        let model = resolved.into_model();
        assert_eq!(model.cell_xfs[1].font_id, 0);
        assert_eq!(model.cell_xfs[2].font_id, 1);
    }

    #[test]
    fn extended_argb_alpha_remains_distinct_for_font_interning() {
        let font = font();
        let mut owned = [0u8; 20];
        owned[17] = 2;
        let xfs = [owned; 16];
        let (check, first) = extended_color(&xfs, 1, 13, [0x80, 0x12, 0x34, 0x56]);
        let (_, second) = extended_color(&xfs, 2, 13, [0xff, 0x12, 0x34, 0x56]);
        let mut records = vec![Record {
            kind: 0x31,
            offset: 0,
            data: &font,
        }];
        records.extend(xfs.iter().map(|xf| Record {
            kind: 0xe0,
            offset: 0,
            data: xf,
        }));
        records.extend([
            Record {
                kind: 0x087c,
                offset: 0,
                data: &check,
            },
            Record {
                kind: 0x087d,
                offset: 0,
                data: &first,
            },
            Record {
                kind: 0x087d,
                offset: 0,
                data: &second,
            },
        ]);
        let model = Styles::parse(&records).unwrap().model().unwrap();
        // Alpha is dropped by color resolution but still distinguishes the
        // interned source fonts.
        assert_eq!(model.fonts.len(), 3);
        assert_eq!(model.fonts[1].color, model.fonts[2].color);
        assert_eq!(
            (model.cell_xfs[1].font_id, model.cell_xfs[2].font_id),
            (1, 2)
        );
    }
    #[test]
    fn rejects_invalid_font_references_and_truncated_records() {
        let font = font();
        for index in [4u8, 5, 255] {
            let mut xf = [0u8; 20];
            xf[0] = index;
            let s = Styles {
                fonts: vec![&font],
                xfs: vec![&xf],
                formats: BTreeMap::new(),
                palette: None,
                extensions: extensions::Extensions::default(),
            };
            assert!(s.resolve().is_err());
        }
        for kind in [0x00e0, 0x041e, 0x0092] {
            assert!(Styles::parse(&[Record {
                kind,
                offset: 0,
                data: &[0]
            }])
            .is_err());
        }
    }

    #[test]
    fn rich_font_index_cannot_address_an_appended_xf_variant() {
        let font = font();
        let xf = [0u8; 20];
        let mut resolved = Styles {
            fonts: vec![&font],
            xfs: vec![&xf],
            formats: BTreeMap::new(),
            palette: None,
            extensions: extensions::Extensions::default(),
        }
        .resolve()
        .unwrap();
        resolved.fonts.push(resolved.fonts[0].clone());
        assert!(resolved.validate_run_font(0).is_ok());
        assert!(resolved.validate_run_font(1).is_err());
    }
    #[test]
    fn keeps_custom_formats_verbatim_without_evaluating_them() {
        let font = font();
        let xf = [0u8; 20];
        let s = Styles {
            fonts: vec![&font],
            xfs: vec![&xf],
            formats: BTreeMap::from([(164, "[Red][<0]0.0\"&\"".into())]),
            palette: None,
            extensions: extensions::Extensions::default(),
        };
        let model = s.model().unwrap();
        assert_eq!(model.num_fmts.len(), 1);
        assert_eq!(model.num_fmts[0].num_fmt_id, 164);
        assert_eq!(model.num_fmts[0].format_code, "[Red][<0]0.0\"&\"");
    }
}
