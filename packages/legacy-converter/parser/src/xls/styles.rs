//! BIFF8 Font/Format/Palette/XF -> SpreadsheetML styles.
//! [MS-XLS] 2.4.122, 2.4.126, 2.4.188, 2.4.353, 2.5.20,
//! 2.5.129; ECMA-376 Part 1 18.8. Cell XFs contain complete properties;
//! fAtr* controls later style updates, not inheritance during display.

use super::{
    decode_biff_chars, minimal_styles, parse_biff_string, u16_at, u32_at, unsupported, Record,
};
use crate::ooxml::xml_attr;
use std::collections::BTreeMap;
mod color;
mod extensions;
mod font;

use color::ColorIdentity;
use font::{ResolvedFont, Script, Underline};

const PATTERNS: [&str; 19] = [
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
const BORDERS: [&str; 14] = [
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
    pub extensions_omitted: bool,
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
    fonts: Vec<ResolvedStyleFont>,
    fills: Vec<ResolvedFill>,
    borders: Vec<ResolvedBorder>,
    xfs: Vec<ResolvedXf>,
    formats: BTreeMap<u16, String>,
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
    quote_prefix: bool,
    horizontal: &'static str,
    vertical: &'static str,
    wrap_text: bool,
    text_rotation: u8,
    indent: u16,
    shrink_to_fit: bool,
    reading_order: u8,
    justify_last_line: bool,
    locked: bool,
    hidden: bool,
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
    pub(super) fn xml(&self) -> String {
        if self.minimal {
            return minimal_styles();
        }
        let fonts: Vec<_> = self.fonts.iter().map(font_style_xml).collect();
        let fills: Vec<_> = self.fills.iter().map(fill_xml).collect();
        let borders: Vec<_> = self.borders.iter().map(border_xml).collect();
        let xfs: Vec<_> = self.xfs.iter().map(xf_xml).collect();
        let normal = xfs[0].replace(" xfId=\"0\"", "");
        let formats: String = self
            .formats
            .iter()
            .map(|(id, code)| {
                format!(
                    "<numFmt numFmtId=\"{id}\" formatCode=\"{}\"/>",
                    xml_attr(code)
                )
            })
            .collect();
        format!("<?xml version=\"1.0\" encoding=\"UTF-8\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><numFmts count=\"{}\">{formats}</numFmts><fonts count=\"{}\">{}</fonts><fills count=\"{}\">{}</fills><borders count=\"{}\">{}</borders><cellStyleXfs count=\"1\">{normal}</cellStyleXfs><cellXfs count=\"{}\">{}</cellXfs><cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles></styleSheet>", self.formats.len(), fonts.len(), fonts.join(""), fills.len(), fills.join(""), borders.len(), borders.join(""), xfs.len(), xfs.join(""))
    }

    #[allow(dead_code)] // Consumed by the direct XLS session in the next unit.
    pub(super) fn into_model(self) -> xlsx_model::Styles {
        xlsx_model::Styles {
            fonts: self
                .fonts
                .into_iter()
                .map(|v| v.font.model(v.color.model()))
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
            dxfs: Vec::new(),
        }
    }
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
            extensions_omitted: false,
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
                // XFExt/StyleExt may contain true colors/gradients beyond BIFF8 XF.
                0x087d | 0x0892 => styles.extensions_omitted = true,
                _ => {}
            }
        }
        styles.extensions = extensions::Extensions::parse(records, &styles.xfs)?;
        Ok(styles)
    }

    pub fn validate_xf(&self, index: u16) -> Result<(), String> {
        if usize::from(index) >= self.xfs.len().max(1) {
            return Err(unsupported("BIFF cell XF index out of range"));
        }
        Ok(())
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

    #[cfg(test)]
    fn font(&self, data: &[u8]) -> Result<String, String> {
        let font = ResolvedFont::decode(data)?;
        Ok(font_xml_value(&font, false, self.color(font.color_index)))
    }

    pub(super) fn run_font(&self, index: u16) -> Result<String, String> {
        // MS-XLS 2.5.129: FontIndex 4 is reserved, indices above it are one-based.
        let offset = usize::from(index - u16::from(index > 4));
        let data = self
            .fonts
            .get(offset)
            .filter(|_| index != 4)
            .ok_or_else(|| unsupported("BIFF rich-text font index out of range"))?;
        let font = ResolvedFont::decode(data)?;
        Ok(font_xml_value(&font, true, self.color(font.color_index)))
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
            let flags = u16_at(data, 4)?;
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
                quote_prefix: flags >> 3 & 1 != 0,
                horizontal,
                vertical,
                wrap_text: data[6] >> 3 & 1 != 0,
                text_rotation: data[7],
                indent,
                shrink_to_fit: data[8] >> 4 & 1 != 0,
                reading_order: reading,
                justify_last_line: data[6] >> 7 != 0,
                locked: flags & 1 != 0,
                hidden: flags >> 1 & 1 != 0,
            });
        }
        if xfs.is_empty() {
            xfs.push(ResolvedXf::default_xf());
        }
        Ok(ResolvedStyleSheet {
            minimal: false,
            fonts,
            fills,
            borders,
            xfs,
            formats: self.formats.clone(),
        })
    }

    pub fn xml(&self) -> Result<String, String> {
        Ok(self.resolve()?.xml())
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
            quote_prefix: false,
            horizontal: "",
            vertical: "",
            wrap_text: false,
            text_rotation: 0,
            indent: 0,
            shrink_to_fit: false,
            reading_order: 0,
            justify_last_line: false,
            locked: false,
            hidden: false,
        }
    }
}

fn minimal_resolved() -> ResolvedStyleSheet {
    let font = ResolvedFont::minimal_calibri();
    ResolvedStyleSheet {
        minimal: true,
        fonts: vec![ResolvedStyleFont {
            font,
            color: ColorIdentity::Auto,
        }],
        fills: vec![ResolvedFill::None, ResolvedFill::Gray125],
        borders: vec![ResolvedBorder::seed()],
        xfs: vec![ResolvedXf::default_xf()],
        formats: BTreeMap::new(),
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

fn font_style_xml(value: &ResolvedStyleFont) -> String {
    font_xml_value(&value.font, false, value.color)
}

fn font_xml_value(font: &ResolvedFont, run: bool, color: ColorIdentity) -> String {
    let (tag, name_tag) = if run {
        ("rPr", "rFont")
    } else {
        ("font", "name")
    };
    let mut xml = format!("<{tag}><{name_tag} val=\"{}\"/><sz val=\"{}\"/><color {}/><family val=\"{}\"/><charset val=\"{}\"/>", xml_attr(&font.name), f64::from(font.size_twips) / 20.0, color.xml(), font.family, font.charset);
    if font.weight == 700 {
        xml.push_str("<b/>");
    } else if run {
        xml.push_str("<b val=\"0\"/>");
    }
    for (enabled, tag) in [
        (font.italic, "i"),
        (font.strike, "strike"),
        (font.outline, "outline"),
        (font.shadow, "shadow"),
        (font.condense, "condense"),
        (font.extend, "extend"),
    ] {
        if enabled {
            xml.push_str(&format!("<{tag}/>"));
        } else if run {
            xml.push_str(&format!("<{tag} val=\"0\"/>"));
        }
    }
    if font.underline != Underline::None || run {
        xml.push_str(&format!("<u val=\"{}\"/>", font.underline.xml_value()));
    }
    if font.script != Script::Baseline || run {
        xml.push_str(&format!("<vertAlign val=\"{}\"/>", font.script.xml_value()));
    }
    xml.push_str(&format!("</{tag}>"));
    xml
}

fn fill_xml(value: &ResolvedFill) -> String {
    match value {
        ResolvedFill::None => "<fill><patternFill patternType=\"none\"/></fill>".into(),
        ResolvedFill::Gray125 => "<fill><patternFill patternType=\"gray125\"/></fill>".into(),
        ResolvedFill::Pattern { pattern, foreground, background } => format!("<fill><patternFill patternType=\"{pattern}\"><fgColor {}/><bgColor {}/></patternFill></fill>", foreground.xml(), background.xml()),
    }
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

fn edge_xml(tag: &str, edge: &Option<ResolvedEdge>) -> String {
    edge.as_ref()
        .map(|e| {
            format!(
                "<{tag} style=\"{}\"><color {}/></{tag}>",
                e.style,
                e.color.xml()
            )
        })
        .unwrap_or_else(|| format!("<{tag}/>"))
}

fn border_xml(value: &ResolvedBorder) -> String {
    if value.seed {
        return "<border><left/><right/><top/><bottom/><diagonal/></border>".into();
    }
    format!(
        "<border diagonalDown=\"{}\" diagonalUp=\"{}\">{}{}{}{}{}</border>",
        value.diagonal_down as u8,
        value.diagonal_up as u8,
        edge_xml("left", &value.left),
        edge_xml("right", &value.right),
        edge_xml("top", &value.top),
        edge_xml("bottom", &value.bottom),
        edge_xml("diagonal", &value.diagonal)
    )
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

fn xf_xml(value: &ResolvedXf) -> String {
    if value.horizontal.is_empty() {
        return "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>".into();
    }
    format!("<xf numFmtId=\"{}\" fontId=\"{}\" fillId=\"{}\" borderId=\"{}\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\" applyProtection=\"1\" quotePrefix=\"{}\"><alignment horizontal=\"{}\" vertical=\"{}\" wrapText=\"{}\" textRotation=\"{}\" indent=\"{}\" shrinkToFit=\"{}\" readingOrder=\"{}\" justifyLastLine=\"{}\"/><protection locked=\"{}\" hidden=\"{}\"/></xf>", value.num_fmt_id, value.font_id, value.fill_id, value.border_id, value.quote_prefix as u8, value.horizontal, value.vertical, value.wrap_text as u8, value.text_rotation, value.indent, value.shrink_to_fit as u8, value.reading_order, value.justify_last_line as u8, value.locked as u8, value.hidden as u8)
}

fn xf_model(value: ResolvedXf) -> xlsx_model::CellXf {
    xlsx_model::CellXf {
        font_id: value.font_id as u32,
        fill_id: value.fill_id as u32,
        border_id: value.border_id as u32,
        num_fmt_id: value.num_fmt_id.into(),
        align_h: (!value.horizontal.is_empty()).then(|| value.horizontal.into()),
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
    fn extended_indent<'a>(xfs: &'a [[u8; 20]], index: u16, value: u16) -> ([u8; 20], Vec<u8>) {
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
    fn minimal_native_model_matches_minimal_styles_xml_defaults() {
        let s = Styles {
            fonts: vec![],
            xfs: vec![],
            formats: BTreeMap::new(),
            palette: None,
            extensions_omitted: false,
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
        assert_eq!(model.fills[0].pattern_type, "none");
        assert_eq!(model.fills[1].pattern_type, "gray125");
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
        let xml = Styles::parse(&records).unwrap().xml().unwrap();
        assert!(xml.contains("<fgColor rgb=\"FF123456\"/>"));
        assert!(xml.contains("<color rgb=\"FF123456\"/>"));
        assert!(xml.contains("<fonts count=\"2\">"));
        assert_eq!(xml.matches("fontId=\"1\"").count(), 1);
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
            let xml = Styles::parse(&records).unwrap().xml().unwrap();
            assert_eq!(xml.matches("indent=\"7\"").count(), 16); // style XF plus 15 unextended cell XFs
            assert_eq!(
                xml.matches(&format!("indent=\"{value}\"")).count(),
                usize::from(value != 7)
            );
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
            let xml = Styles::parse(&records).unwrap().xml().unwrap();
            let expected = match case {
                "cell" => 1,
                "style" => 2,
                _ => 0,
            };
            assert_eq!(xml.matches("indent=\"250\"").count(), expected, "{case}");
            assert_eq!(xml.matches("readingOrder=\"2\"").count(), 2, "{case}");
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
                    ext.extend_from_slice(&ext[20..26].to_vec());
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
            extensions_omitted: false,
            extensions: extensions::Extensions::default(),
        };
        let xml = s.xml().unwrap();
        assert!(xml.contains("fontId=\"4\""));
        assert!(xml.contains("<fills count=\"2\">"));
        assert!(xml.contains("<cellXfs count=\"100\">"));
        let model = s.model().unwrap();
        assert_eq!(
            (model.fonts.len(), model.fills.len(), model.borders.len()),
            (5, 2, 2)
        );
        assert_eq!(model.cell_xfs.len(), 100);
        assert_eq!(model.cell_xfs[0].font_id, 4);
        assert_eq!(model.cell_xfs[0].fill_id, 0);
        assert_eq!(model.cell_xfs[0].border_id, 1);
        assert_eq!(model.cell_xfs[0].align_h.as_deref(), Some("general"));
        assert_eq!(model.cell_xfs[0].align_v.as_deref(), Some("top"));
        assert_eq!(model.cell_xfs[0].indent, None);
        assert_eq!(model.cell_xfs[0].text_rotation, None);
        assert_eq!(model.cell_xfs[0].reading_order, None);
    }
    #[test]
    fn typed_font_keeps_legacy_cell_and_run_xml_byte_exact() {
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
            extensions_omitted: false,
            extensions: extensions::Extensions::default(),
        };
        assert_eq!(s.font(&font).unwrap(), "<font><name val=\"A&amp;B\"/><sz val=\"12\"/><color indexed=\"10\"/><family val=\"3\"/><charset val=\"128\"/><b/><i/><strike/><outline/><shadow/><condense/><extend/><u val=\"singleAccounting\"/><vertAlign val=\"superscript\"/></font>");
        assert_eq!(s.run_font(0).unwrap(), "<rPr><rFont val=\"A&amp;B\"/><sz val=\"12\"/><color indexed=\"10\"/><family val=\"3\"/><charset val=\"128\"/><b/><i/><strike/><outline/><shadow/><condense/><extend/><u val=\"singleAccounting\"/><vertAlign val=\"superscript\"/></rPr>");
        assert!(s.run_font(4).is_err());
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
            extensions_omitted: false,
            extensions: extensions::Extensions::default(),
        };
        assert_eq!(plain_styles.run_font(0).unwrap(), "<rPr><rFont val=\"A&amp;B\"/><sz val=\"12\"/><color indexed=\"10\"/><family val=\"3\"/><charset val=\"128\"/><b val=\"0\"/><i val=\"0\"/><strike val=\"0\"/><outline val=\"0\"/><shadow val=\"0\"/><condense val=\"0\"/><extend val=\"0\"/><u val=\"none\"/><vertAlign val=\"baseline\"/></rPr>");
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
            extensions_omitted: false,
            extensions: extensions::Extensions::default(),
        };
        assert_eq!(s.color(10).xml(), "rgb=\"FF123456\"");
        assert_eq!(s.color(0x7fff).xml(), "auto=\"1\"");
        assert_eq!(s.color(65).xml(), "indexed=\"65\"");
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
        let xml = styles.xml().unwrap();
        // The extension resolves to the same typed ARGB identity as the
        // palette-backed base font, so exact XML interning keeps one font.
        assert!(xml.contains("<fonts count=\"1\">"));
        assert_eq!(xml.matches("rgb=\"FF123456\"").count(), 1);
    }
    #[test]
    fn font_override_interns_emission_equivalent_weights_without_removing_source_slots() {
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
        assert!(resolved.xml().contains("<fonts count=\"2\">"));
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
        let xml = Styles::parse(&records).unwrap().xml().unwrap();
        assert!(xml.contains("<fonts count=\"3\">"));
        assert!(xml.contains("<color rgb=\"80123456\"/>"));
        assert!(xml.contains("<color rgb=\"FF123456\"/>"));
        assert_eq!(xml.matches("fontId=\"1\"").count(), 1);
        assert_eq!(xml.matches("fontId=\"2\"").count(), 1);
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
                extensions_omitted: false,
                extensions: extensions::Extensions::default(),
            };
            assert!(s.xml().is_err());
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
    fn escapes_custom_formats_without_evaluating_them() {
        let font = font();
        let xf = [0u8; 20];
        let s = Styles {
            fonts: vec![&font],
            xfs: vec![&xf],
            formats: BTreeMap::from([(164, "[Red][<0]0.0\"&\"".into())]),
            palette: None,
            extensions_omitted: false,
            extensions: extensions::Extensions::default(),
        };
        assert!(s
            .xml()
            .unwrap()
            .contains("[Red][&lt;0]0.0&quot;&amp;&quot;"));
    }
}
