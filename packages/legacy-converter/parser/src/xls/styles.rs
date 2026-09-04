//! BIFF8 Font/Format/Palette/XF -> SpreadsheetML styles.
//! [MS-XLS] 2.4.122, 2.4.126, 2.4.188, 2.4.353, 2.5.20,
//! 2.5.129; ECMA-376 Part 1 18.8. Cell XFs contain complete properties;
//! fAtr* controls later style updates, not inheritance during display.

use super::{
    decode_biff_chars, minimal_styles, parse_biff_string, u16_at, u32_at, unsupported, Record,
};
use crate::ooxml::xml_attr;
use std::collections::BTreeMap;

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
    pub extensions_omitted: bool,
}

impl<'a> Styles<'a> {
    pub fn parse(records: &[Record<'a>]) -> Result<Self, String> {
        let mut styles = Self {
            fonts: vec![],
            xfs: vec![],
            formats: BTreeMap::new(),
            palette: None,
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
        if data.len() < 16 {
            return Err(unsupported("truncated BIFF font"));
        }
        let (name, _) = decode_biff_chars(data, 16, usize::from(data[14]), data[15] & 1 != 0)?;
        let mut xml = format!("<font><name val=\"{}\"/><sz val=\"{}\"/><color {}/><family val=\"{}\"/><charset val=\"{}\"/>", xml_attr(&name), f64::from(u16_at(data, 0)?) / 20.0, self.color(u16_at(data, 4)?), data[11], data[12]);
        // OOXML exposes only bold/normal, not arbitrary LOGFONT weight.
        if u16_at(data, 6)? == 700 {
            xml.push_str("<b/>");
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
            }
        }
        let underline = match data[10] {
            0 => None,
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
            0 => {}
            1 => xml.push_str("<vertAlign val=\"superscript\"/>"),
            2 => xml.push_str("<vertAlign val=\"subscript\"/>"),
            _ => return Err(unsupported("invalid BIFF font script")),
        }
        xml.push_str("</font>");
        Ok(xml)
    }

    pub fn xml(&self) -> Result<String, String> {
        if self.xfs.is_empty() && self.fonts.is_empty() {
            return Ok(minimal_styles());
        }
        let mut fonts = String::new();
        for font in &self.fonts {
            fonts.push_str(&self.font(font)?);
        }
        if fonts.is_empty() {
            return Err(unsupported("BIFF styles reference missing fonts"));
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
        for data in &self.xfs {
            let ifnt = u16_at(data, 0)?;
            let font = usize::from(ifnt - u16::from(ifnt > 4));
            if ifnt == 4 || font >= self.fonts.len() {
                return Err(unsupported("BIFF XF font index out of range"));
            }
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
                format!("<fill><patternFill patternType=\"{pattern}\"><fgColor {}/><bgColor {}/></patternFill></fill>", self.color(colors & 127), self.color((colors >> 7) & 127))
            };
            let fill_id = intern(fill, &mut fills, &mut fill_ids);
            let mut border = format!(
                "<border diagonalDown=\"{}\" diagonalUp=\"{}\">",
                (b1 >> 30) & 1,
                (b1 >> 31) & 1
            );
            for (tag, style, color) in [
                ("left", b1 & 15, (b1 >> 16) & 127),
                ("right", (b1 >> 4) & 15, (b1 >> 23) & 127),
                ("top", (b1 >> 8) & 15, b2 & 127),
                ("bottom", (b1 >> 12) & 15, (b2 >> 7) & 127),
                ("diagonal", (b2 >> 21) & 15, (b2 >> 14) & 127),
            ] {
                let style = BORDERS
                    .get(style as usize)
                    .ok_or_else(|| unsupported("invalid BIFF border style"))?;
                if *style == "none" {
                    border.push_str(&format!("<{tag}/>"));
                } else {
                    border.push_str(&format!(
                        "<{tag} style=\"{style}\"><color {}/></{tag}>",
                        self.color(color as u16)
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
        Ok(format!("<?xml version=\"1.0\" encoding=\"UTF-8\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><numFmts count=\"{}\">{formats}</numFmts><fonts count=\"{}\">{fonts}</fonts><fills count=\"{}\">{}</fills><borders count=\"{}\">{}</borders><cellStyleXfs count=\"1\">{normal}</cellStyleXfs><cellXfs count=\"{}\">{}</cellXfs><cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles></styleSheet>", self.formats.len(), self.fonts.len(), fills.len(), fills.join(""), borders.len(), borders.join(""), xfs.len(), xfs.join("")))
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
        };
        assert!(s
            .xml()
            .unwrap()
            .contains("[Red][&lt;0]0.0&quot;&amp;&quot;"));
    }
}
