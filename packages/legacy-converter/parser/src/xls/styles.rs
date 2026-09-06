//! BIFF8 Font/Format/Palette/XF -> SpreadsheetML styles.
//! [MS-XLS] 2.4.122, 2.4.126, 2.4.188, 2.4.353, 2.5.20,
//! 2.5.129; ECMA-376 Part 1 18.8. Cell XFs contain complete properties;
//! fAtr* controls later style updates, not inheritance during display.

use super::{
    decode_biff_chars, minimal_styles, parse_biff_string, u16_at, u32_at, unsupported, Record,
};
use crate::ooxml::xml_attr;
use std::collections::BTreeMap;
mod extensions;

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

impl<'a> Styles<'a> {
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

    fn color(&self, index: u16) -> String {
        if let Some(palette) = self.palette {
            if (8..64).contains(&index) {
                let offset = usize::from(index - 8) * 4;
                return format!(
                    "rgb=\"FF{:02X}{:02X}{:02X}\"",
                    palette[offset],
                    palette[offset + 1],
                    palette[offset + 2]
                );
            }
        }
        if index == 0x7fff {
            "auto=\"1\"".into()
        } else {
            format!("indexed=\"{index}\"")
        }
    }

    fn font(&self, data: &[u8]) -> Result<String, String> {
        self.font_xml(data, false, None)
    }

    pub(super) fn run_font(&self, index: u16) -> Result<String, String> {
        // MS-XLS 2.5.129: FontIndex 4 is reserved, indices above it are one-based.
        let offset = usize::from(index - u16::from(index > 4));
        let data = self
            .fonts
            .get(offset)
            .filter(|_| index != 4)
            .ok_or_else(|| unsupported("BIFF rich-text font index out of range"))?;
        self.font_xml(data, true, None)
    }

    fn font_xml(&self, data: &[u8], run: bool, color: Option<&str>) -> Result<String, String> {
        if data.len() < 16 {
            return Err(unsupported("truncated BIFF font"));
        }
        let (name, _) = decode_biff_chars(data, 16, usize::from(data[14]), data[15] & 1 != 0)?;
        let (tag, name_tag) = if run {
            ("rPr", "rFont")
        } else {
            ("font", "name")
        };
        let base_color = self.color(u16_at(data, 4)?);
        let mut xml = format!("<{tag}><{name_tag} val=\"{}\"/><sz val=\"{}\"/><color {}/><family val=\"{}\"/><charset val=\"{}\"/>", xml_attr(&name), f64::from(u16_at(data, 0)?) / 20.0, color.unwrap_or(&base_color), data[11], data[12]);
        // OOXML exposes only bold/normal, not arbitrary LOGFONT weight.
        if u16_at(data, 6)? == 700 {
            xml.push_str("<b/>");
        } else if run {
            xml.push_str("<b val=\"0\"/>");
        }
        for (mask, tag) in [
            (2, "i"),
            (8, "strike"),
            (16, "outline"),
            (32, "shadow"),
            (64, "condense"),
            (128, "extend"),
        ] {
            if data[2] & mask != 0 {
                xml.push_str(&format!("<{tag}/>"));
            } else if run {
                xml.push_str(&format!("<{tag} val=\"0\"/>"));
            }
        }
        let underline = match data[10] {
            0 => {
                if run {
                    Some("none")
                } else {
                    None
                }
            }
            1 => Some("single"),
            2 => Some("double"),
            0x21 => Some("singleAccounting"),
            0x22 => Some("doubleAccounting"),
            _ => return Err(unsupported("invalid BIFF underline")),
        };
        if let Some(value) = underline {
            xml.push_str(&format!("<u val=\"{value}\"/>"));
        }
        match u16_at(data, 8)? {
            0 => {
                if run {
                    xml.push_str("<vertAlign val=\"baseline\"/>");
                }
            }
            1 => xml.push_str("<vertAlign val=\"superscript\"/>"),
            2 => xml.push_str("<vertAlign val=\"subscript\"/>"),
            _ => return Err(unsupported("invalid BIFF font script")),
        }
        xml.push_str(&format!("</{tag}>"));
        Ok(xml)
    }

    pub fn xml(&self) -> Result<String, String> {
        if self.xfs.is_empty() && self.fonts.is_empty() {
            return Ok(minimal_styles());
        }
        let mut fonts = Vec::new();
        for font in &self.fonts {
            fonts.push(self.font(font)?);
        }
        if fonts.is_empty() {
            return Err(unsupported("BIFF styles reference missing fonts"));
        }
        // Keep original font indices stable for shared-string rich runs. XF-local
        // color overrides append an interned variant, never mutate a shared font.
        let mut font_ids = BTreeMap::new();
        for (id, xml) in fonts.iter().enumerate() {
            font_ids.entry(xml.clone()).or_insert(id);
        }
        let mut fills = vec![
            "<fill><patternFill patternType=\"none\"/></fill>".to_string(),
            "<fill><patternFill patternType=\"gray125\"/></fill>".to_string(),
        ];
        let mut borders =
            vec!["<border><left/><right/><top/><bottom/><diagonal/></border>".to_string()];
        let mut fill_ids = BTreeMap::from([(fills[0].clone(), 0usize), (fills[1].clone(), 1)]);
        let mut border_ids = BTreeMap::from([(borders[0].clone(), 0usize)]);
        let mut xfs = Vec::new();
        for (index, data) in self.xfs.iter().enumerate() {
            let ifnt = u16_at(data, 0)?;
            let mut font = usize::from(ifnt - u16::from(ifnt > 4));
            if ifnt == 4 || font >= self.fonts.len() {
                return Err(unsupported("BIFF XF font index out of range"));
            }
            if let Some(color) = self.extensions.color(index, 13) {
                font = intern(
                    self.font_xml(self.fonts[font], false, Some(color))?,
                    &mut fonts,
                    &mut font_ids,
                );
            }
            let color = |property, fallback| {
                self.extensions
                    .color(index, property)
                    .map(str::to_owned)
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
                fills[0].clone()
            } else {
                format!("<fill><patternFill patternType=\"{pattern}\"><fgColor {}/><bgColor {}/></patternFill></fill>", color(4, colors & 127), color(5, (colors >> 7) & 127))
            };
            let fill_id = intern(fill, &mut fills, &mut fill_ids);
            let mut border = format!(
                "<border diagonalDown=\"{}\" diagonalUp=\"{}\">",
                (b1 >> 30) & 1,
                (b1 >> 31) & 1
            );
            for (tag, style, palette_color, property) in [
                ("left", b1 & 15, (b1 >> 16) & 127, 9),
                ("right", (b1 >> 4) & 15, (b1 >> 23) & 127, 10),
                ("top", (b1 >> 8) & 15, b2 & 127, 7),
                ("bottom", (b1 >> 12) & 15, (b2 >> 7) & 127, 8),
                ("diagonal", (b2 >> 21) & 15, (b2 >> 14) & 127, 11),
            ] {
                let style = BORDERS
                    .get(style as usize)
                    .ok_or_else(|| unsupported("invalid BIFF border style"))?;
                if *style == "none" {
                    border.push_str(&format!("<{tag}/>"));
                } else {
                    border.push_str(&format!(
                        "<{tag} style=\"{style}\"><color {}/></{tag}>",
                        color(property, palette_color as u16)
                    ));
                }
            }
            border.push_str("</border>");
            let border_id = intern(border, &mut borders, &mut border_ids);
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
            if (181..255).contains(&data[7]) || (flags & 4 == 0 && data[8] >> 6 > 2) {
                return Err(unsupported("invalid BIFF text alignment"));
            }
            // StyleXF reserves cIndent/iReadOrder. Ignore those bits for styles.
            let indent = if flags & 4 == 0 { data[8] & 15 } else { 0 };
            let reading = if flags & 4 == 0 { data[8] >> 6 } else { 0 };
            xfs.push(format!("<xf numFmtId=\"{}\" fontId=\"{font}\" fillId=\"{fill_id}\" borderId=\"{border_id}\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\" applyProtection=\"1\" quotePrefix=\"{}\"><alignment horizontal=\"{horizontal}\" vertical=\"{vertical}\" wrapText=\"{}\" textRotation=\"{}\" indent=\"{indent}\" shrinkToFit=\"{}\" readingOrder=\"{reading}\" justifyLastLine=\"{}\"/><protection locked=\"{}\" hidden=\"{}\"/></xf>", u16_at(data, 2)?, (flags >> 3) & 1, (data[6] >> 3) & 1, data[7], (data[8] >> 4) & 1, data[6] >> 7, flags & 1, (flags >> 1) & 1));
        }
        if xfs.is_empty() {
            xfs.push(
                "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>".into(),
            );
        }
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
        Ok(format!("<?xml version=\"1.0\" encoding=\"UTF-8\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><numFmts count=\"{}\">{formats}</numFmts><fonts count=\"{}\">{}</fonts><fills count=\"{}\">{}</fills><borders count=\"{}\">{}</borders><cellStyleXfs count=\"1\">{normal}</cellStyleXfs><cellXfs count=\"{}\">{}</cellXfs><cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles></styleSheet>", self.formats.len(), fonts.len(), fonts.join(""), fills.len(), fills.join(""), borders.len(), borders.join(""), xfs.len(), xfs.join("")))
    }
}

fn intern(value: String, values: &mut Vec<String>, ids: &mut BTreeMap<String, usize>) -> usize {
    if let Some(index) = ids.get(&value) {
        return *index;
    }
    let index = values.len();
    ids.insert(value.clone(), index);
    values.push(value);
    index
}

#[cfg(test)]
mod tests {
    use super::*;
    fn font() -> Vec<u8> {
        let mut data = vec![0; 16];
        data[..2].copy_from_slice(&200u16.to_le_bytes());
        data[6..8].copy_from_slice(&400u16.to_le_bytes());
        data[14] = 1;
        data.push(b'F');
        data
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
        assert_eq!(s.color(10), "rgb=\"FF123456\"");
        assert_eq!(s.color(0x7fff), "auto=\"1\"");
        assert_eq!(s.color(65), "indexed=\"65\"");
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
