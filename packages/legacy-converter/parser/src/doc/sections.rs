//! [MS-DOC] 2.8.26 PlcfSed, 2.9.243 Sed, 2.9.245 Sepx,
//! 2.2.5.1 Sprm, 2.6.4 section properties. ECMA-376 17.6.
use super::{u16_at, u32_at, unsupported};

pub(super) struct Section {
    pub end: usize,
    pub xml: String,
    pub incomplete_margins: bool,
}

pub(super) fn read(word: &[u8], table: &[u8], ccp_text: usize) -> Result<Vec<Section>, String> {
    // FibRgFcLcb97, fc/lcbPlcfSed. A zero lcb makes the fc undefined.
    let size = u32_at(word, 0xce)? as usize;
    if size == 0 {
        return Ok(Vec::new());
    }
    if size < 20 || !(size - 4).is_multiple_of(16) {
        return Err(unsupported("invalid Word section table length"));
    }
    let count = (size - 4) / 16;
    // Resource policy, not a format limit.
    if count > 16384 {
        return Err(unsupported("too many Word sections"));
    }
    let offset = u32_at(word, 0xca)? as usize;
    let bytes = table
        .get(offset..)
        .and_then(|tail| tail.get(..size))
        .ok_or_else(|| unsupported("Word section table is outside its stream"))?;
    if u32_at(bytes, 0)? != 0 {
        return Err(unsupported("Word section table does not start at CP zero"));
    }
    let mut previous = 0;
    let mut budget = 1_000_000usize;
    let mut sections = Vec::with_capacity(count);
    for index in 0..count {
        let end = u32_at(bytes, (index + 1) * 4)? as usize;
        if end <= previous
            || (index + 1 < count && end >= ccp_text)
            || (index + 1 == count && end < ccp_text)
        {
            return Err(unsupported("invalid Word section character range"));
        }
        let sepx = u32_at(bytes, (count + 1) * 4 + index * 12 + 2)? as usize;
        let tail = word
            .get(sepx..)
            .ok_or_else(|| unsupported("Word section properties are outside WordDocument"))?;
        let size = u16_at(tail, 0)? as usize;
        if size > i16::MAX as usize {
            return Err(unsupported("negative Word section property size"));
        }
        let data = tail
            .get(2..2 + size)
            .ok_or_else(|| unsupported("truncated Word section properties"))?;
        let properties = Properties::parse(data, &mut budget)?;
        sections.push(Section {
            end,
            xml: properties.xml()?,
            incomplete_margins: properties.margins.iter().any(Option::is_none),
        });
        previous = end;
    }
    Ok(sections)
}

// CPs count UTF-16 units, not Rust characters or UTF-8 bytes. A form feed at
// a section boundary is consumed by sectPr; other form feeds remain page breaks.
pub(super) fn split_story<'a>(text: &'a str, sections: &[Section]) -> Result<Vec<&'a str>, String> {
    if sections.is_empty() {
        return Ok(vec![text]);
    }
    let mut result = Vec::with_capacity(sections.len());
    let mut cp = 0;
    let mut start = 0;
    let mut section = 0;
    for (byte, character) in text.char_indices() {
        cp += character.len_utf16();
        if section + 1 < sections.len() {
            let end = sections[section].end;
            if cp > end {
                return Err(unsupported("Word section ends inside a Unicode character"));
            }
            if cp == end {
                if character != '\u{c}' {
                    return Err(unsupported("Word section boundary lacks a section break"));
                }
                result.push(&text[start..byte]);
                start = byte + 1;
                section += 1;
            }
        }
    }
    if section + 1 != sections.len() {
        return Err(unsupported("Word section range exceeds main story"));
    }
    result.push(&text[start..]);
    Ok(result)
}

struct Properties {
    size: (u16, u16),
    orientation: u8,
    kind: u8,
    margins: [Option<i32>; 6],
    gutter: u16,
    columns: u16,
    equal: bool,
    spacing: Option<u16>,
    separator: bool,
    widths: [Option<u16>; 44],
    spaces: [u16; 44],
    vertical: u8,
    title: bool,
    bidi: bool,
    rtl_gutter: bool,
    grid: u16,
    line_pitch: Option<u16>,
    char_space: i32,
}

impl Properties {
    fn parse(mut bytes: &[u8], budget: &mut usize) -> Result<Self, String> {
        let mut p = Self {
            size: (12240, 15840),
            orientation: 1,
            kind: 2,
            margins: [None; 6],
            gutter: 0,
            columns: 1,
            equal: true,
            spacing: None,
            separator: false,
            widths: [None; 44],
            spaces: [0; 44],
            vertical: 0,
            title: false,
            bidi: false,
            rtl_gutter: false,
            grid: 0,
            line_pitch: None,
            char_space: 0,
        };
        while !bytes.is_empty() {
            *budget = budget
                .checked_sub(1)
                .ok_or_else(|| unsupported("Word section property budget exceeded"))?;
            let sprm = u16_at(bytes, 0)?;
            if (sprm >> 10) & 7 != 4 {
                return Err(unsupported("non-section SPRM in Word Sepx"));
            }
            bytes = &bytes[2..];
            let size =
                match sprm >> 13 {
                    0 | 1 => 1,
                    2 | 4 | 5 => 2,
                    3 => 4,
                    7 => 3,
                    // The two variable-length exceptions apply to paragraph/table
                    // SPRMs, which cannot reach this section-only decoder.
                    _ => {
                        1 + usize::from(*bytes.first().ok_or_else(|| {
                            unsupported("truncated variable Word section property")
                        })?)
                    }
                };
            let value = bytes
                .get(..size)
                .ok_or_else(|| unsupported("truncated Word section SPRM"))?;
            bytes = &bytes[size..];
            match sprm {
                0x3009 => {
                    if value[0] > 4 {
                        return Err(unsupported("invalid Word section break type"));
                    }
                    p.kind = value[0];
                }
                0x301d => {
                    if !(1..=2).contains(&value[0]) {
                        return Err(unsupported("invalid Word page orientation"));
                    }
                    p.orientation = value[0];
                }
                0xb01f | 0xb020 => {
                    let v = u16_at(value, 0)?;
                    if !(144..=31680).contains(&v) {
                        return Err(unsupported("invalid Word page size"));
                    }
                    if sprm == 0xb01f {
                        p.size.0 = v;
                    } else {
                        p.size.1 = v;
                    }
                }
                0xb021 | 0xb022 | 0xb017 | 0xb018 => {
                    let v = u16_at(value, 0)?;
                    if v > 31680 {
                        return Err(unsupported("invalid Word horizontal/header margin"));
                    }
                    let index = match sprm {
                        0xb021 => 3,
                        0xb022 => 1,
                        0xb017 => 4,
                        _ => 5,
                    };
                    p.margins[index] = Some(i32::from(v));
                }
                0x9023 | 0x9024 => {
                    let v = i32::from(u16_at(value, 0)? as i16);
                    if !(-31665..=31665).contains(&v) {
                        return Err(unsupported("invalid Word signed margin"));
                    }
                    p.margins[if sprm == 0x9023 { 0 } else { 2 }] = Some(v);
                }
                0xb025 => p.gutter = u16_at(value, 0)?,
                0x500b => {
                    let v = u16_at(value, 0)?;
                    if v > 43 {
                        return Err(unsupported("invalid Word column count"));
                    }
                    p.columns = v + 1;
                }
                0x900c => {
                    let v = u16_at(value, 0)?;
                    if v > 31680 {
                        return Err(unsupported("invalid Word column spacing"));
                    }
                    p.spacing = Some(v);
                }
                0xf203 | 0xf204 => {
                    let index = usize::from(value[0]);
                    let v = u16_at(value, 1)?;
                    // SDxaColWidthOperand requires >= 718 twips; the
                    // adjacent SDxaColSpacingOperand explicitly allows zero.
                    if index >= 44 || v > 31680 || (sprm == 0xf203 && v < 718) {
                        return Err(unsupported("invalid Word column geometry"));
                    }
                    if sprm == 0xf203 {
                        p.widths[index] = Some(v);
                    } else {
                        p.spaces[index] = v;
                    }
                }
                0x3005 | 0x3019 | 0x300a | 0x3228 | 0x322a => {
                    if value[0] > 1 {
                        return Err(unsupported("invalid Word section Boolean"));
                    }
                    let v = value[0] != 0;
                    match sprm {
                        0x3005 => p.equal = v,
                        0x3019 => p.separator = v,
                        0x300a => p.title = v,
                        0x3228 => p.bidi = v,
                        _ => p.rtl_gutter = v,
                    }
                }
                0x301a => {
                    if value[0] > 3 {
                        return Err(unsupported("invalid Word vertical alignment"));
                    }
                    p.vertical = value[0];
                }
                0x5032 => {
                    let v = u16_at(value, 0)?;
                    if v > 3 {
                        return Err(unsupported("invalid Word document grid"));
                    }
                    p.grid = v;
                }
                0x9031 => {
                    let v = u16_at(value, 0)?;
                    if !(1..=31680).contains(&v) {
                        return Err(unsupported("invalid Word line pitch"));
                    }
                    p.line_pitch = Some(v);
                }
                0x7030 => {
                    let v = u32_at(value, 0)? as i32;
                    if !(-670925..=6488064).contains(&v) {
                        return Err(unsupported("invalid Word character grid pitch"));
                    }
                    p.char_space = v;
                }
                _ => {} // Remaining properties are covered by the lossy-conversion warning.
            }
        }
        Ok(p)
    }

    fn xml(&self) -> Result<String, String> {
        let kind = [
            "continuous",
            "nextColumn",
            "nextPage",
            "evenPage",
            "oddPage",
        ][usize::from(self.kind)];
        let mut xml = format!(
            "<w:sectPr><w:type w:val=\"{kind}\"/><w:pgSz w:w=\"{}\" w:h=\"{}\" w:orient=\"{}\"/>",
            self.size.0,
            self.size.1,
            if self.orientation == 2 {
                "landscape"
            } else {
                "portrait"
            }
        );
        if let [Some(top), Some(right), Some(bottom), Some(left), header, footer] = self.margins {
            // Conversion policy: headers/footers are not emitted yet. Preserve
            // known body margins even when locale-dependent header/footer
            // distances are absent, using neutral zero distances for these
            // absent stories to satisfy CT_PageMar's required attributes.
            // This does not claim to recover the producer's locale defaults.
            let header = header.unwrap_or(0);
            let footer = footer.unwrap_or(0);
            xml.push_str(&format!("<w:pgMar w:top=\"{top}\" w:right=\"{right}\" w:bottom=\"{bottom}\" w:left=\"{left}\" w:header=\"{header}\" w:footer=\"{footer}\" w:gutter=\"{}\"/>",self.gutter));
        }
        xml.push_str(&format!(
            "<w:cols w:num=\"{}\" w:equalWidth=\"{}\" w:sep=\"{}\"",
            self.columns,
            u8::from(self.equal),
            u8::from(self.separator)
        ));
        if let Some(space) = self.spacing {
            xml.push_str(&format!(" w:space=\"{space}\""));
        }
        xml.push('>');
        if !self.equal {
            for index in 0..usize::from(self.columns) {
                let width = self.widths[index]
                    .ok_or_else(|| unsupported("Word unequal columns lack a width"))?;
                xml.push_str(&format!(
                    "<w:col w:w=\"{width}\" w:space=\"{}\"/>",
                    self.spaces[index]
                ));
            }
        }
        xml.push_str("</w:cols>");
        xml.push_str(&format!(
            "<w:vAlign w:val=\"{}\"/>",
            ["top", "center", "both", "bottom"][usize::from(self.vertical)]
        ));
        for (name, value) in [
            ("titlePg", self.title),
            ("bidi", self.bidi),
            ("rtlGutter", self.rtl_gutter),
        ] {
            xml.push_str(&format!("<w:{name} w:val=\"{}\"/>", u8::from(value)));
        }
        if self.grid != 0 {
            let pitch = self
                .line_pitch
                .ok_or_else(|| unsupported("Word document grid lacks line pitch"))?;
            xml.push_str(&format!(
                "<w:docGrid w:type=\"{}\" w:linePitch=\"{pitch}\" w:charSpace=\"{}\"/>",
                ["default", "linesAndChars", "lines", "snapToChars"][usize::from(self.grid)],
                self.char_space
            ));
        }
        xml.push_str("</w:sectPr>");
        Ok(xml)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn rejects_invalid_section_ranges_and_out_of_stream_properties() {
        let mut word = vec![0; 512];
        word[0xce..0xd2].copy_from_slice(&20u32.to_le_bytes());
        let mut table = vec![0; 20];
        table[4..8].copy_from_slice(&3u32.to_le_bytes());
        table[10..14].copy_from_slice(&300u32.to_le_bytes());
        assert_eq!(read(&word, &table, 3).unwrap().len(), 1);
        assert!(read(&word, &table, 4).is_err());
        table[10..14].copy_from_slice(&u32::MAX.to_le_bytes());
        assert!(read(&word, &table, 3).is_err());
        word[0xce..0xd2].copy_from_slice(&(4u32 + 16 * 16385).to_le_bytes());
        assert!(read(&word, &[], 3).is_err());
    }

    #[test]
    fn preserves_unequal_columns_and_requires_each_width() {
        let mut bytes = vec![0x05, 0x30, 0];
        bytes.extend(prl(0x500b, 1));
        assert!(Properties::parse(&bytes, &mut 100).unwrap().xml().is_err());
        for (sprm, index, value) in [
            (0xf203u16, 0u8, 2000u16),
            (0xf203, 1, 4000),
            (0xf204, 0, 300),
        ] {
            bytes.extend(sprm.to_le_bytes());
            bytes.push(index);
            bytes.extend(value.to_le_bytes());
        }
        let xml = Properties::parse(&bytes, &mut 100).unwrap().xml().unwrap();
        assert!(xml
            .contains("<w:col w:w=\"2000\" w:space=\"300\"/><w:col w:w=\"4000\" w:space=\"0\"/>"));
    }

    #[test]
    fn column_width_has_its_own_normative_minimum_not_a_spacing_minimum() {
        let operand = |sprm: u16, width: u16| {
            [
                sprm.to_le_bytes().to_vec(),
                vec![0],
                width.to_le_bytes().to_vec(),
            ]
            .concat()
        };
        assert!(Properties::parse(&operand(0xf203, 717), &mut 100).is_err());
        assert!(Properties::parse(&operand(0xf203, 718), &mut 100).is_ok());
        assert!(Properties::parse(&operand(0xf204, 0), &mut 100).is_ok());
    }
    #[test]
    fn missing_header_distance_does_not_discard_known_body_margins() {
        let mut bytes = Vec::new();
        for (sprm, value) in [
            (0x9023, 1800),
            (0x9024, 1440),
            (0xb021, 2160),
            (0xb022, 1440),
            (0xb025, 720),
        ] {
            bytes.extend(prl(sprm, value));
        }
        let xml = Properties::parse(&bytes, &mut 100).unwrap().xml().unwrap();
        assert!(xml.contains("w:top=\"1800\""));
        assert!(xml.contains("w:left=\"2160\""));
        assert!(xml.contains("w:header=\"0\" w:footer=\"0\" w:gutter=\"720\""));
    }
    fn prl(sprm: u16, value: i16) -> Vec<u8> {
        [sprm.to_le_bytes(), value.to_le_bytes()].concat()
    }
    #[test]
    fn keeps_signed_margins_and_document_grid_units() {
        let mut bytes = Vec::new();
        for (sprm, value) in [
            (0xb01f, 11906),
            (0xb020, 16838),
            (0x9023, -500),
            (0x9024, -300),
            (0xb021, 1200),
            (0xb022, 1300),
            (0xb017, 700),
            (0xb018, 600),
            (0x5032, 1),
            (0x9031, 330),
        ] {
            bytes.extend(prl(sprm, value));
        }
        bytes.extend(0x7030u16.to_le_bytes());
        bytes.extend(6144i32.to_le_bytes());
        let p = Properties::parse(&bytes, &mut 100).unwrap();
        let xml = p.xml().unwrap();
        assert!(xml.contains("w:w=\"11906\" w:h=\"16838\""));
        assert!(xml.contains("w:top=\"-500\"") && xml.contains("w:bottom=\"-300\""));
        assert!(xml.contains("w:linePitch=\"330\" w:charSpace=\"6144\""));
    }
    #[test]
    fn section_boundaries_count_surrogate_pairs_and_do_not_double_page_breaks() {
        let sections = [
            Section {
                end: 4,
                xml: String::new(),
                incomplete_margins: false,
            },
            Section {
                end: 9,
                xml: String::new(),
                incomplete_margins: false,
            },
        ];
        assert_eq!(
            split_story("A😀\u{c}B\u{c}C", &sections).unwrap(),
            ["A😀", "B\u{c}C"]
        );
        assert!(split_story("A😀XY", &sections).is_err());
    }
    #[test]
    fn rejects_truncated_properties_and_charges_repeated_work() {
        assert!(Properties::parse(&[0x1f, 0xb0, 0], &mut 100).is_err());
        assert!(Properties::parse(&prl(0xb01f, 100), &mut 100).is_err());
        assert!(Properties::parse(&prl(0xb01f, 12000), &mut 0).is_err());
        assert!(Properties::parse(&[0x34, 0xd2, 8, 0], &mut 100).is_err());
    }
}
