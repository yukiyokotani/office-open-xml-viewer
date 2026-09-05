//! Paragraph layout from [MS-DOC] 2.6.2, LSPD and XAS to ECMA-376
//! 17.3.1 / CT_PPrBase. Values retain their units; no font/pixel fitting.

use super::{border::Border, u16_at, unsupported};
use std::collections::BTreeMap;

#[derive(Clone)]
pub struct Properties {
    tabs: super::tabs::Stops,
    flags: BTreeMap<&'static str, bool>,
    line: (u16, &'static str),
    before: u16,
    after: u16,
    before_lines: Option<i16>,
    after_lines: Option<i16>,
    before_auto: bool,
    after_auto: bool,
    indents: [Option<(i32, u32)>; 4],
    indent_order: u32,
    first: i32,
    nest: Option<(i32, bool)>,
    chars: [Option<i16>; 3],
    alignment: (u8, bool),
    text_alignment: Option<&'static str>,
    borders: [Option<Border>; 5],
}

impl Default for Properties {
    fn default() -> Self {
        Self {
            tabs: super::tabs::Stops::default(),
            flags: BTreeMap::from([
                ("widowControl", true),
                ("kinsoku", true),
                ("wordWrap", true),
                ("overflowPunct", true),
                ("autoSpaceDE", true),
                ("autoSpaceDN", true),
                ("snapToGrid", true),
                ("adjustRightInd", true),
            ]),
            line: (240, "auto"),
            before: 0,
            after: 0,
            before_lines: None,
            after_lines: None,
            before_auto: false,
            after_auto: false,
            indents: [None; 4],
            indent_order: 0,
            first: 0,
            nest: None,
            chars: [None; 3],
            alignment: (0, false),
            text_alignment: None,
            borders: std::array::from_fn(|_| None),
        }
    }
}

fn signed(bytes: &[u8]) -> Result<i32, String> {
    let value = u16_at(bytes, 0)? as i16 as i32;
    if !(-31680..=31680).contains(&value) {
        return Err(unsupported("invalid Word paragraph distance"));
    }
    Ok(value)
}

impl Properties {
    pub fn apply(&mut self, code: u16, operand: &[u8]) -> Result<bool, String> {
        let flag = match code {
            0x2405 => Some("keepLines"),
            0x2406 => Some("keepNext"),
            0x2407 => Some("pageBreakBefore"),
            0x240c => Some("suppressLineNumbers"),
            0x242a => Some("suppressAutoHyphens"),
            0x2431 => Some("widowControl"),
            0x2433 => Some("kinsoku"),
            0x2434 => Some("wordWrap"),
            0x2435 => Some("overflowPunct"),
            0x2436 => Some("topLinePunct"),
            0x2437 => Some("autoSpaceDE"),
            0x2438 => Some("autoSpaceDN"),
            0x2441 => Some("bidi"),
            0x2447 => Some("snapToGrid"),
            0x2448 => Some("adjustRightInd"),
            _ => None,
        };
        if let Some(flag) = flag {
            self.flags.insert(flag, bool8(operand[0])?);
            return Ok(true);
        }
        match code {
            0x6424..=0x6428 => {
                let side = usize::from(code - 0x6424);
                self.borders[side] = Some(Border::paragraph(operand, true, side)?);
            }
            0xc64e..=0xc652 => {
                // MS-DOC 2.9.21: BrcOperand.cb MUST be 8, excluding cb.
                if operand.len() != 9 || operand[0] != 8 {
                    return Err(unsupported("invalid Word paragraph border operand"));
                }
                let side = usize::from(code - 0xc64e);
                self.borders[side] = Some(Border::paragraph(&operand[1..], false, side)?);
            }
            0xc60d | 0xc615 => self.tabs.apply(operand, code == 0xc615)?,
            0x6412 => {
                let line = signed(operand)?;
                let multiple = u16_at(operand, 2)?;
                if multiple > 1 {
                    return Err(unsupported("invalid Word line spacing multiplier"));
                }
                self.line = if line < 0 {
                    ((-line) as u16, "exact")
                } else {
                    (line as u16, if multiple == 1 { "auto" } else { "atLeast" })
                };
            }
            0xa413 | 0xa414 => {
                let value = u16_at(operand, 0)?;
                if value > 31680 {
                    return Err(unsupported("invalid Word paragraph spacing"));
                }
                if code == 0xa413 {
                    self.before = value;
                } else {
                    self.after = value;
                }
            }
            0x4458 | 0x4459 => {
                let value = u16_at(operand, 0)? as i16;
                if !(-20..=31680).contains(&value) {
                    return Err(unsupported("invalid Word line-unit paragraph spacing"));
                }
                if code == 0x4458 {
                    self.before_lines = Some(value);
                } else {
                    self.after_lines = Some(value);
                }
            }
            0x245b => self.before_auto = bool8(operand[0])?,
            0x245c => self.after_auto = bool8(operand[0])?,
            0x840e | 0x845d | 0x840f | 0x845e => {
                let index = match code {
                    0x840f => 0,
                    0x840e => 1,
                    0x845e => 2,
                    _ => 3,
                };
                self.indent_order += 1;
                self.indents[index] = Some((signed(operand)?, self.indent_order));
            }
            0x8411 | 0x8460 => self.first = signed(operand)?,
            0x4610 | 0x465f => {
                // PNest supersedes PNest80, irrespective of their order.
                if code == 0x465f || !self.nest.is_some_and(|(_, physical)| !physical) {
                    self.nest = Some((signed(operand)?, code == 0x4610));
                }
            }
            0x4455..=0x4457 => {
                let index = match code {
                    0x4455 => 1,
                    0x4456 => 0,
                    _ => 2,
                };
                self.chars[index] = Some(u16_at(operand, 0)? as i16);
            }
            0x2403 | 0x2461 => {
                let maximum = if code == 0x2403 { 5 } else { 9 };
                if operand[0] > maximum {
                    return Err(unsupported("invalid Word paragraph alignment"));
                }
                self.alignment = (operand[0], code == 0x2403);
            }
            0x4439 => {
                self.text_alignment = Some(match u16_at(operand, 0)? {
                    0 => "top",
                    1 => "center",
                    2 => "baseline",
                    3 => "bottom",
                    4 => "auto",
                    _ => return Err(unsupported("invalid Word paragraph text alignment")),
                });
            }
            _ => return Ok(false),
        }
        Ok(true)
    }

    pub fn xml(&self) -> String {
        let mut xml = String::new();
        // CT_PPrBase has a sequence, not an arbitrary element map.
        for key in [
            "keepNext",
            "keepLines",
            "pageBreakBefore",
            "widowControl",
            "suppressLineNumbers",
            "suppressAutoHyphens",
            "kinsoku",
            "wordWrap",
            "overflowPunct",
            "topLinePunct",
            "autoSpaceDE",
            "autoSpaceDN",
            "bidi",
            "adjustRightInd",
            "snapToGrid",
        ] {
            if key == "suppressAutoHyphens" {
                // MS-DOC 2.6.2 PBrcLeft/Right are logical. ECMA-376
                // 17.3.1.17/28 left/right are physical, so resolve after bidi.
                // Keep grouping as ordinary pBdr; the existing OOXML layout
                // owns adjacency, spacing and between-border decisions.
                if self.borders.iter().any(Option::is_some) {
                    xml.push_str("<w:pBdr>");
                    let rtl = self.flags.get("bidi") == Some(&true);
                    for (side, index) in [
                        ("top", 0),
                        ("left", if rtl { 3 } else { 1 }),
                        ("bottom", 2),
                        ("right", if rtl { 1 } else { 3 }),
                        ("between", 4),
                    ] {
                        if let Some(border) = &self.borders[index] {
                            xml.push_str(&border.xml(side));
                        }
                    }
                    xml.push_str("</w:pBdr>");
                }
                // CT_PPrBase places tabs after shading/borders and before this flag.
                xml.push_str(&self.tabs.xml());
            }
            if let Some(value) = self.flags.get(key) {
                xml.push_str(&format!("<w:{key} w:val=\"{}\"/>", u8::from(*value)));
            }
        }
        xml.push_str(&format!("<w:spacing w:before=\"{}\" w:after=\"{}\" w:line=\"{}\" w:lineRule=\"{}\" w:beforeAutospacing=\"{}\" w:afterAutospacing=\"{}\"",
            self.before, self.after, self.line.0, self.line.1, u8::from(self.before_auto), u8::from(self.after_auto)));
        for (name, value) in [
            ("beforeLines", self.before_lines),
            ("afterLines", self.after_lines),
        ] {
            if let Some(value) = value {
                xml.push_str(&format!(" w:{name}=\"{value}\""));
            }
        }
        xml.push_str("/>");
        let bidi = self.flags.get("bidi") == Some(&true);
        // Resolve physical/logical coordinates after bidi is known. Keep the
        // last assignment to either logical side, including mixed old/new SPRMs.
        let mut indents = [0, 0];
        let mut writes: Vec<_> = self
            .indents
            .iter()
            .enumerate()
            .filter_map(|(i, value)| value.map(|(v, order)| (order, i, v)))
            .collect();
        writes.sort_unstable_by_key(|(order, _, _)| *order);
        for (_, i, value) in writes {
            indents[(i % 2) ^ usize::from(i < 2 && bidi)] = value;
        }
        let [mut left, mut right] = indents;
        if let Some((value, physical)) = self.nest {
            if physical && bidi {
                right += value;
            } else {
                left += value;
            }
        }
        xml.push_str(&format!(
            "<w:ind w:left=\"{left}\" w:right=\"{right}\" w:{}=\"{}\"",
            if self.first < 0 {
                "hanging"
            } else {
                "firstLine"
            },
            self.first.abs()
        ));
        for (name, value) in [("leftChars", self.chars[0]), ("rightChars", self.chars[1])] {
            if let Some(value) = value {
                xml.push_str(&format!(" w:{name}=\"{value}\""));
            }
        }
        if let Some(value) = self.chars[2] {
            xml.push_str(&format!(
                " w:{}=\"{}\"",
                if value < 0 {
                    "hangingChars"
                } else {
                    "firstLineChars"
                },
                i32::from(value).abs()
            ));
        }
        xml.push_str("/>");
        let (value, physical) = self.alignment;
        let alignment = if physical {
            match value {
                0 if bidi => "right",
                2 if bidi => "left",
                0 => "left",
                1 => "center",
                2 => "right",
                3 => "both",
                4 => "mediumKashida",
                _ => "highKashida",
            }
        } else {
            [
                "left",
                "center",
                "right",
                "both",
                "distribute",
                "mediumKashida",
                "numTab",
                "highKashida",
                "lowKashida",
                "thaiDistribute",
            ][value as usize]
        };
        xml.push_str(&format!("<w:jc w:val=\"{alignment}\"/>"));
        if let Some(value) = self.text_alignment {
            xml.push_str(&format!("<w:textAlignment w:val=\"{value}\"/>"));
        }
        xml
    }
}

fn bool8(value: u8) -> Result<bool, String> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(unsupported("invalid Word paragraph boolean")),
    }
}

pub fn prm0(prm: u16) -> Option<[u8; 3]> {
    let code: u16 = match (prm >> 1) & 127 {
        0x05 => 0x2461,
        0x07 => 0x2405,
        0x08 => 0x2406,
        0x09 => 0x2407,
        0x0e => 0x240c,
        0x2c => 0x242a,
        0x33 => 0x2431,
        0x35 => 0x2433,
        0x36 => 0x2434,
        0x37 => 0x2435,
        0x38 => 0x2436,
        0x39 => 0x2437,
        0x3a => 0x2438,
        _ => return None,
    };
    let [a, b] = code.to_le_bytes();
    Some([a, b, (prm >> 8) as u8])
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn paragraph_borders_preserve_each_side_and_schema_order() {
        let mut p = Properties::default();
        assert!(!p.xml().contains("pBdr"));
        for code in 0x6424..=0x6428 {
            assert!(p.apply(code, &[8, 1, 2, 3]).unwrap());
        }
        let xml = p.xml();
        let border = xml
            .split("<w:pBdr>")
            .nth(1)
            .unwrap()
            .split("</w:pBdr>")
            .next()
            .unwrap();
        let mut offset = 0;
        for side in ["top", "left", "bottom", "right", "between"] {
            let at = border.find(&format!("<w:{side} ")).unwrap();
            assert!(at >= offset);
            offset = at;
            assert!(border[at..].starts_with(&format!(
                "<w:{side} w:val=\"single\" w:sz=\"8\" w:color=\"0000FF\" w:space=\"3\""
            )));
        }
        assert!(xml.find("</w:pBdr>").unwrap() < xml.find("<w:kinsoku").unwrap());
    }

    #[test]
    fn modern_border_replaces_only_its_side_and_explicit_none_clears_it() {
        let mut p = Properties::default();
        p.apply(0x6424, &[8, 1, 2, 0]).unwrap();
        p.apply(0x6426, &[8, 1, 2, 0]).unwrap();
        assert!(p
            .apply(0xc650, &[8, 0x12, 0x34, 0x56, 0, 16, 3, 7, 0])
            .unwrap());
        let xml = p.xml();
        assert!(xml.contains("<w:top w:val=\"single\""));
        assert!(
            xml.contains("<w:bottom w:val=\"double\" w:sz=\"16\" w:color=\"123456\" w:space=\"7\"")
        );
        p.apply(0xc650, &[8, 0, 0, 0, 0xff, 0, 0, 0, 0]).unwrap();
        assert!(p.xml().contains("<w:bottom w:val=\"none\""));
        assert!(p.xml().contains("<w:top w:val=\"single\""));
        p.apply(0xc650, &[8, 0, 1, 2, 3, 255, 255, 255, 255])
            .unwrap();
        assert!(p.xml().contains("<w:bottom w:val=\"nil\"/>"));
        assert!(p.xml().contains("<w:top w:val=\"single\""));
        for bad in [&[0][..], &[7, 0, 0, 0, 0, 0, 0, 0], &[8, 0, 0]] {
            assert!(p.apply(0xc650, bad).is_err());
        }
    }

    #[test]
    fn paragraph_borders_resolve_logical_sides_after_the_final_bidi_setting() {
        let mut p = Properties::default();
        p.apply(0x6425, &[8, 1, 2, 0]).unwrap();
        p.apply(0xc651, &[8, 255, 0, 0, 0, 8, 3, 0, 0]).unwrap();
        p.apply(0x2441, &[1]).unwrap();
        let xml = p.xml();
        assert!(xml.contains("<w:left w:val=\"double\""));
        assert!(xml.contains("<w:right w:val=\"single\""));
        p.apply(0x2441, &[0]).unwrap();
        assert!(p.xml().contains("<w:left w:val=\"single\""));
        assert!(p.xml().contains("<w:right w:val=\"double\""));
    }

    #[test]
    fn negative_line_spacing_is_exact_even_when_multiplier_flag_is_set() {
        let mut p = Properties::default();
        p.apply(0x6412, &[0xd4, 0xfe, 1, 0]).unwrap(); // -300 twips
        assert!(p.xml().contains("w:line=\"300\" w:lineRule=\"exact\""));
        p.apply(0x6412, &[0x68, 1, 1, 0]).unwrap();
        assert!(p.xml().contains("w:line=\"360\" w:lineRule=\"auto\""));
        p.apply(0x6412, &[0x68, 1, 0, 0]).unwrap();
        assert!(p.xml().contains("w:line=\"360\" w:lineRule=\"atLeast\""));
        assert!(p.apply(0x6412, &[0, 0, 2, 0]).is_err());
    }

    #[test]
    fn preserves_signed_hanging_indents_and_explicit_spacing_resets() {
        let mut p = Properties::default();
        p.apply(0x845e, &720u16.to_le_bytes()).unwrap();
        p.apply(0x8460, &(-360i16).to_le_bytes()).unwrap();
        p.apply(0xa413, &240u16.to_le_bytes()).unwrap();
        p.apply(0xa413, &[0, 0]).unwrap();
        assert!(p
            .xml()
            .contains("w:left=\"720\" w:right=\"0\" w:hanging=\"360\""));
        assert!(p.xml().contains("w:before=\"0\""));
    }

    #[test]
    fn late_bidi_property_does_not_turn_physical_indents_into_logical_ones() {
        let mut p = Properties::default();
        p.apply(0x840f, &720u16.to_le_bytes()).unwrap();
        p.apply(0x840e, &360u16.to_le_bytes()).unwrap();
        p.apply(0x2441, &[1]).unwrap();
        assert!(p.xml().contains("w:left=\"360\" w:right=\"720\""));
        p.apply(0x845d, &240u16.to_le_bytes()).unwrap();
        assert!(p.xml().contains("w:left=\"360\" w:right=\"240\""));
        let mut single = Properties::default();
        single.apply(0x840f, &720u16.to_le_bytes()).unwrap();
        single.apply(0x2441, &[1]).unwrap();
        assert!(single.xml().contains("w:left=\"0\" w:right=\"720\""));
    }
}
