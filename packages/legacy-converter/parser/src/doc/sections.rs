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
        let mut properties = Properties::parse(data, &mut budget)?;
        // FibBase.lid records the producer's installation language (MS-DOC
        // 2.5.2). Apply only the LCIDs explicitly listed in 2.6.4, not the
        // document text's language, the converter host locale or a guessed
        // base language. Explicit distances, including zero, take priority.
        let default = header_distance_for_install_lid(u16_at(word, 6)?);
        for distance in &mut properties.margins[4..] {
            *distance = distance.or(default);
        }
        sections.push(Section {
            end,
            xml: properties.xml()?,
            incomplete_margins: properties.margins.iter().any(Option::is_none),
        });
        previous = end;
    }
    Ok(sections)
}

// MS-DOC 2.6.4 sprmSDyaHdrTop / sprmSDyaHdrBottom share this default table.
// Unlisted/unknown installation languages remain unresolved and warned about.
fn header_distance_for_install_lid(lid: u16) -> Option<i32> {
    match lid {
        1025 | 1028 | 1031 | 1032 | 1033 | 1034 | 1036 | 1037 | 1040 | 1041 | 1042 | 1046
        | 1049 | 1050 | 1053 | 1062 | 1086 | 1104 | 2052 | 2070 => Some(720),
        1026 | 1027 | 1029 | 1030 | 1035 | 1038 | 1039 | 1043 | 1044 | 1045 | 1048 | 1051
        | 1055 | 1058 | 1059 | 1060 | 1061 | 1067 | 1068 | 1069 | 1078 | 1079 | 1087 | 1088
        | 1089 | 1092 | 2074 => Some(708),
        1063 => Some(567),
        _ => None,
    }
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
    text_flow: Option<&'static str>,
    page_format: &'static str,
    page_restart: bool,
    page_start: u32,
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
            text_flow: None,
            page_format: "decimal",
            page_restart: false,
            page_start: 0,
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
                0x300e => p.page_format = page_number_format(value[0])?,
                // MS-DOC 2.2.5: later modifiers of the same property win.
                // The older 16-bit operand has a SHOULD, not MUST, maximum
                // of 32766. Do not reinterpret it as signed or clamp it.
                0x501c => p.page_start = u32::from(u16_at(value, 0)?),
                0x7044 => p.page_start = u32_at(value, 0)?,
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
                0x3005 | 0x3019 | 0x300a | 0x3011 | 0x3228 | 0x322a => {
                    if value[0] > 1 {
                        return Err(unsupported("invalid Word section Boolean"));
                    }
                    let v = value[0] != 0;
                    match sprm {
                        0x3005 => p.equal = v,
                        0x3019 => p.separator = v,
                        0x300a => p.title = v,
                        0x3011 => p.page_restart = v,
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
                0x5033 => {
                    // MS-DOC 2.6.4 sprmSTextFlow -> MS-ODRAW 2.4.5 MSOTXFL.
                    // TtoBA: glyph sequence top-to-bottom, columns right-to-left.
                    // ECMA-376 17.6.20 / 17.18.93 and Part 4 14.11.7:
                    // Transitional tbRl is the matching section-flow token.
                    p.text_flow = match u16_at(value, 0)? {
                        0 => None,
                        1 => Some("tbRl"),
                        // Rotation variants and the Word-version-dependent
                        // VertN column direction are not inferred from this
                        // base-flow mapping. Covered by the existing advanced
                        // section-property omission warning.
                        2..=5 => None,
                        _ => return Err(unsupported("invalid Word section text flow")),
                    };
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
            // Recovery policy: retain known body margins when the stored
            // installation language cannot resolve a missing header/footer
            // distance. Zero satisfies CT_PageMar's required attributes;
            // incomplete_margins keeps this unresolved recovery diagnostic.
            let header = header.unwrap_or(0);
            let footer = footer.unwrap_or(0);
            xml.push_str(&format!("<w:pgMar w:top=\"{top}\" w:right=\"{right}\" w:bottom=\"{bottom}\" w:left=\"{left}\" w:header=\"{header}\" w:footer=\"{footer}\" w:gutter=\"{}\"/>",self.gutter));
        }
        // MS-DOC 2.6.4 SNfcPgn/SFPgnRestart/SPgnStart97/SPgnStart ->
        // ECMA-376 17.6.12 pgNumType. A dormant start MUST be ignored.
        // Validate only the effective restart, after all modifiers are applied.
        // Omission preserves decimal continuation, independently per section.
        if self.page_restart || self.page_format != "decimal" {
            xml.push_str(&format!("<w:pgNumType w:fmt=\"{}\"", self.page_format));
            if self.page_restart {
                if self.page_start > 2147483646 {
                    return Err(unsupported("invalid Word page-number restart"));
                }
                xml.push_str(&format!(" w:start=\"{}\"", self.page_start));
            }
            xml.push_str("/>");
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
        xml.push_str(&format!("<w:titlePg w:val=\"{}\"/>", u8::from(self.title)));
        if let Some(flow) = self.text_flow {
            xml.push_str(&format!("<w:textDirection w:val=\"{flow}\"/>"));
        }
        for (name, value) in [("bidi", self.bidi), ("rtlGutter", self.rtl_gutter)] {
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

// MS-OSHARED 2.2.1.3 MSONFC -> ECMA-376 17.18.59 ST_NumberFormat.
// Tokens use the actual OOXML schema spelling (lowercase aiueo/iroha).
// For the non-counting bullet format, use the decimal fallback permitted by
// MS-DOC 2.6.4 SNfcPgn and documented for Word in its implementation note.
// Preserve `none`, whose specified meaning is to suppress the number.
fn page_number_format(value: u8) -> Result<&'static str, String> {
    const FORMATS: [&str; 60] = [
        "decimal",
        "upperRoman",
        "lowerRoman",
        "upperLetter",
        "lowerLetter",
        "ordinal",
        "cardinalText",
        "ordinalText",
        "hex",
        "chicago",
        "ideographDigital",
        "japaneseCounting",
        "aiueo",
        "iroha",
        "decimalFullWidth",
        "decimalHalfWidth",
        "japaneseLegal",
        "japaneseDigitalTenThousand",
        "decimalEnclosedCircle",
        "decimalFullWidth2",
        "aiueoFullWidth",
        "irohaFullWidth",
        "decimalZero",
        "decimal",
        "ganada",
        "chosung",
        "decimalEnclosedFullstop",
        "decimalEnclosedParen",
        "decimalEnclosedCircleChinese",
        "ideographEnclosedCircle",
        "ideographTraditional",
        "ideographZodiac",
        "ideographZodiacTraditional",
        "taiwaneseCounting",
        "ideographLegalTraditional",
        "taiwaneseCountingThousand",
        "taiwaneseDigital",
        "chineseCounting",
        "chineseLegalSimplified",
        "chineseCountingThousand",
        "decimal",
        "koreanDigital",
        "koreanCounting",
        "koreanLegal",
        "koreanDigital2",
        "hebrew1",
        "arabicAlpha",
        "hebrew2",
        "arabicAbjad",
        "hindiVowels",
        "hindiConsonants",
        "hindiNumbers",
        "hindiCounting",
        "thaiLetters",
        "thaiNumbers",
        "thaiCounting",
        "vietnameseCounting",
        "numberInDash",
        "russianLower",
        "russianUpper",
    ];
    if value == 0xff {
        return Ok("none");
    }
    FORMATS
        .get(usize::from(value))
        .copied()
        .ok_or_else(|| unsupported("invalid Word page-number format"))
}

#[cfg(test)]
mod tests {
    use super::*;
    fn section_xml(bytes: &[u8]) -> String {
        Properties::parse(bytes, &mut 100).unwrap().xml().unwrap()
    }

    #[test]
    fn page_numbering_continues_by_default_and_ignores_dormant_start_values() {
        for bytes in [
            vec![],
            prl(0x501c, 25),
            [vec![0x11, 0x30, 1], prl(0x501c, 25), vec![0x11, 0x30, 0]].concat(),
        ] {
            assert!(!section_xml(&bytes).contains("pgNumType"));
        }
        assert!(section_xml(&[0x11, 0x30, 1])
            .contains("<w:pgNumType w:fmt=\"decimal\" w:start=\"0\"/>"));
        assert!(Properties::parse(&[0x11, 0x30, 2], &mut 100).is_err());
    }

    #[test]
    fn page_number_start_uses_unsigned_widths_and_the_last_applied_modifier() {
        for start in [0u32, 1, 32766, 32767, 65535, 65536, 2147483646] {
            let mut bytes = vec![0x11, 0x30, 1];
            if start <= u16::MAX as u32 {
                bytes.extend(prl(0x501c, start as u16 as i16));
                assert!(section_xml(&bytes).contains(&format!("w:start=\"{start}\"")));
            }
            bytes.extend(0x7044u16.to_le_bytes());
            bytes.extend(start.to_le_bytes());
            let xml = section_xml(&bytes);
            assert!(xml.contains(&format!("w:start=\"{start}\"")));
            assert!(xml.find("<w:pgNumType").unwrap() < xml.find("<w:cols").unwrap());
        }
        let bytes = [
            vec![0x11, 0x30, 1],
            prl(0x501c, 9),
            0x7044u16.to_le_bytes().to_vec(),
            123456u32.to_le_bytes().to_vec(),
        ]
        .concat();
        assert!(section_xml(&bytes).contains("w:start=\"123456\""));
        for start in [2147483647u32, u32::MAX] {
            let bytes = [vec![0x44, 0x70], start.to_le_bytes().to_vec()].concat();
            // MUST ignore a dormant start, including an out-of-range value.
            assert!(!section_xml(&bytes).contains("pgNumType"));
            assert!(
                Properties::parse(&[bytes, vec![0x11, 0x30, 1]].concat(), &mut 100)
                    .unwrap()
                    .xml()
                    .is_err()
            );
        }
    }

    #[test]
    fn section_number_formats_map_to_ooxml_without_changing_continuation() {
        for (value, name) in [
            (1, "upperRoman"),
            (2, "lowerRoman"),
            (3, "upperLetter"),
            (4, "lowerLetter"),
            (0x0c, "aiueo"),
            (0x0d, "iroha"),
            (0x16, "decimalZero"),
            (0x28, "decimal"),
            (0x3b, "russianUpper"),
            (0xff, "none"),
        ] {
            let xml = section_xml(&[0x0e, 0x30, value]);
            if name == "decimal" {
                assert!(!xml.contains("pgNumType"));
            } else {
                assert!(xml.contains(&format!("<w:pgNumType w:fmt=\"{name}\"/>")));
            }
            assert!(!xml.contains("w:start="));
        }
        assert!(!section_xml(&[0x0e, 0x30, 2, 0x0e, 0x30, 0]).contains("pgNumType"));
        assert!(Properties::parse(&[0x0e, 0x30, 0x60], &mut 100).is_err());
    }

    #[test]
    fn stored_install_language_resolves_only_missing_header_distances() {
        // MS-DOC 2.5.2 FibBase.lid and 2.6.4 sprmSDyaHdrTop/Bottom.
        for (lid, default) in [
            (1033u16, 720),
            (1036, 720),
            (1041, 720),
            (1043, 708),
            (1063, 567),
            (0, 0),
            (0xffff, 0),
        ] {
            for explicit in [None, Some(0), Some(900)] {
                let mut properties = [
                    prl(0x9023, 1440),
                    prl(0x9024, 1440),
                    prl(0xb021, 1440),
                    prl(0xb022, 1440),
                ]
                .concat();
                if let Some(value) = explicit {
                    properties.extend(prl(0xb017, value));
                }
                let mut word = vec![0; 512];
                word[6..8].copy_from_slice(&lid.to_le_bytes());
                word[0xce..0xd2].copy_from_slice(&20u32.to_le_bytes());
                word[300..302].copy_from_slice(&(properties.len() as u16).to_le_bytes());
                word[302..302 + properties.len()].copy_from_slice(&properties);
                let mut table = vec![0; 20];
                table[4..8].copy_from_slice(&3u32.to_le_bytes());
                table[10..14].copy_from_slice(&300u32.to_le_bytes());
                let sections = read(&word, &table, 3).unwrap();
                assert!(
                    sections[0].xml.contains(&format!(
                        "w:header=\"{}\" w:footer=\"{default}\"",
                        explicit.unwrap_or(default)
                    )),
                    "lid={lid}, explicit={explicit:?}: {}",
                    sections[0].xml
                );
                assert_eq!(sections[0].incomplete_margins, default == 0);
            }
        }
    }

    #[test]
    fn preserves_section_vertical_flow_and_explicit_horizontal_reset() {
        let vertical = prl(0x5033, 1);
        let xml = Properties::parse(&vertical, &mut 100)
            .unwrap()
            .xml()
            .unwrap();
        assert!(xml.contains("<w:textDirection w:val=\"tbRl\"/>"));
        assert!(xml.find("<w:titlePg").unwrap() < xml.find("<w:textDirection").unwrap());
        assert!(xml.find("<w:textDirection").unwrap() < xml.find("<w:bidi").unwrap());
        let horizontal = [vertical, prl(0x5033, 0)].concat();
        let xml = Properties::parse(&horizontal, &mut 100)
            .unwrap()
            .xml()
            .unwrap();
        assert!(!xml.contains("tbRl"));
        assert!(Properties::parse(&prl(0x5033, 6), &mut 100).is_err());
        assert!(Properties::parse(&prl(0x5033, -1), &mut 100).is_err());
    }

    #[test]
    fn unsupported_rotation_variants_do_not_retain_a_previous_flow() {
        for flow in 2..=5 {
            let bytes = [prl(0x5033, 1), prl(0x5033, flow)].concat();
            let xml = Properties::parse(&bytes, &mut 100).unwrap().xml().unwrap();
            assert!(!xml.contains("textDirection"));
        }
    }

    #[test]
    fn each_section_resolves_its_flow_independently() {
        for flows in [[0, 1], [1, 0]] {
            let mut word = vec![0; 512];
            word[0xce..0xd2].copy_from_slice(&36u32.to_le_bytes());
            let mut table = vec![0; 36];
            table[4..8].copy_from_slice(&2u32.to_le_bytes());
            table[8..12].copy_from_slice(&4u32.to_le_bytes());
            for (index, flow) in flows.into_iter().enumerate() {
                let offset = 400 + index * 10;
                table[14 + index * 12..18 + index * 12]
                    .copy_from_slice(&(offset as u32).to_le_bytes());
                word[offset..offset + 2].copy_from_slice(&4u16.to_le_bytes());
                word[offset + 2..offset + 6].copy_from_slice(&prl(0x5033, flow));
            }
            let sections = read(&word, &table, 4).unwrap();
            for (section, flow) in sections.iter().zip(flows) {
                assert_eq!(section.xml.contains("tbRl"), flow == 1);
            }
        }
    }
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
