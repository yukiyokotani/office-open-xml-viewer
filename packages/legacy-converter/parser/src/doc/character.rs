//! A bounded, explicit character-property subset. Unknown SPRMs are skipped
//! using their encoded operand size, never interpreted as text or executed.
//! [MS-DOC] 2.2.5, 2.6.1, 2.9.327; ECMA-376 17.3.2 (run properties).

use super::{u16_at, u32_at, unsupported};
use crate::ooxml::xml_attr;
use std::collections::BTreeMap;

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Properties {
    values: BTreeMap<&'static str, String>,
    pub fonts: [Option<usize>; 4],
    pub picture: Picture,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct Picture {
    pub special: bool,
    pub data: bool,
    pub ole: bool,
    pub object: bool,
    pub location: Option<i32>,
}

impl Picture {
    pub fn passive_special(self) -> bool {
        self.special && !self.data && !self.ole && !self.object
    }
    /// MS-DOC 2.6.1: only a special U+0001 without binary-data/OLE flags
    /// denotes a PICFAndOfficeArtData. The caller checks the character code.
    pub fn inline_location(self) -> Result<Option<usize>, String> {
        if !self.passive_special() {
            return Ok(None);
        }
        let location = self
            .location
            .ok_or_else(|| unsupported("Word picture lacks a data location"))?;
        usize::try_from(location)
            .map(Some)
            .map_err(|_| unsupported("negative Word picture data location"))
    }
}

impl Default for Properties {
    fn default() -> Self {
        Self {
            values: BTreeMap::from([("sz", "20".into())]),
            fonts: [None; 4],
            picture: Picture::default(),
        }
    }
}

impl Properties {
    /// A resolved style patch, without injecting the document's default size.
    /// Used when list-linked styles are layered onto a paragraph mark.
    pub fn sparse() -> Self {
        Self {
            values: BTreeMap::new(),
            fonts: [None; 4],
            picture: Picture::default(),
        }
    }

    pub fn overlay_visible(&mut self, patch: &Self) {
        self.values
            .extend(patch.values.iter().map(|(k, v)| (*k, v.clone())));
        for (current, added) in self.fonts.iter_mut().zip(patch.fonts) {
            if added.is_some() {
                *current = added;
            }
        }
        // Object/special flags are not visual run formatting and must not
        // turn numbering text into a picture or an executable object.
    }

    pub fn reset_to(&mut self, paragraph: &Self, preserve_object: bool) {
        // Of the reset exceptions in sprmCPlain/CIstd, these are the supported
        // properties. Revision metadata remains omitted. CIstd preserves
        // CFObj; CPlain does not include it in its exception list.
        let mut picture = self.picture;
        if !preserve_object {
            picture.object = paragraph.picture.object;
        }
        let preserved: Vec<_> = ["rtl", "cs", "highlight", "webHidden"]
            .iter()
            .map(|key| (*key, self.values.get(key).cloned()))
            .collect();
        *self = paragraph.clone();
        self.picture = picture;
        for (key, value) in preserved {
            if let Some(value) = value {
                self.values.insert(key, value);
            } else {
                self.values.remove(key);
            }
        }
    }

    pub fn apply(&mut self, code: u16, operand: &[u8], style: &Self) -> Result<bool, String> {
        match code {
            0x6a03 => {
                self.picture.location = Some(u32_at(operand, 0)? as i32);
                return Ok(true);
            }
            0x0855 => {
                self.picture.special = match operand[0] {
                    0 => false,
                    1 => true,
                    0x80 => style.picture.special,
                    0x81 => !style.picture.special,
                    _ => return Err(unsupported("invalid Word special-character toggle")),
                };
                return Ok(true);
            }
            0x0806 | 0x080a | 0x0856 => {
                if operand[0] > 1 {
                    return Err(unsupported("invalid Word picture/object flag"));
                }
                let value = operand[0] != 0;
                match code {
                    0x0806 => self.picture.data = value,
                    0x080a => self.picture.ole = value,
                    _ => self.picture.object = value,
                }
                return Ok(true);
            }
            _ => {}
        }
        let flag = match code {
            0x0835 => Some("b"),
            0x0836 => Some("i"),
            0x0837 => Some("strike"),
            0x0838 => Some("outline"),
            0x0839 => Some("shadow"),
            0x083a => Some("smallCaps"),
            0x083b => Some("caps"),
            0x083c => Some("vanish"),
            0x2a53 => Some("dstrike"),
            0x085a => Some("rtl"),
            0x085c => Some("bCs"),
            0x085d => Some("iCs"),
            0x0882 => Some("cs"),
            _ => None,
        };
        if let Some(key) = flag {
            let base = style.values.get(key).is_some_and(|v| v == "1");
            let value = match operand[0] {
                0 => false,
                1 => true,
                0x80 => base,
                0x81 => !base,
                _ => return Err(unsupported("invalid Word character toggle")),
            };
            self.values
                .insert(key, if value { "1" } else { "0" }.into());
            return Ok(true);
        }
        let (key, value) = match code {
            0x4a4f | 0x4a50 | 0x4a51 | 0x4a5e => {
                let slot = match code {
                    0x4a4f => 0,
                    0x4a50 => 1,
                    0x4a51 => 2,
                    _ => 3,
                };
                let index = u16_at(operand, 0)?;
                if index > i16::MAX as u16 {
                    return Err(unsupported("negative Word font index"));
                }
                self.fonts[slot] = Some(index as usize);
                return Ok(true);
            }
            0x4a43 | 0x4a61 => {
                let size = u16_at(operand, 0)?;
                if size > 3276 || (code == 0x4a43 && size < 2) {
                    return Err(unsupported("invalid Word character size"));
                }
                (if code == 0x4a43 { "sz" } else { "szCs" }, size.to_string())
            }
            0x8840 => ("spacing", (u16_at(operand, 0)? as i16).to_string()),
            0x4845 => {
                let position = u16_at(operand, 0)? as i16;
                if !(-3168..=3168).contains(&position) {
                    return Err(unsupported("invalid Word character position"));
                }
                ("position", position.to_string())
            }
            0x484b => {
                let kern = u16_at(operand, 0)?;
                if kern > 3276 {
                    return Err(unsupported("invalid Word kerning size"));
                }
                ("kern", kern.to_string())
            }
            0x4852 => {
                let value = u16_at(operand, 0)?;
                if !(1..=600).contains(&value) {
                    return Err(unsupported("invalid Word character scaling"));
                }
                ("w", value.to_string())
            }
            0x2a48 => (
                "vertAlign",
                match operand[0] {
                    0 => "baseline",
                    1 => "superscript",
                    2 => "subscript",
                    _ => return Err(unsupported("invalid Word character vertical alignment")),
                }
                .into(),
            ),
            0x2a3e => (
                "u",
                match operand[0] {
                    0 => "none",
                    1 => "single",
                    2 => "words",
                    3 => "double",
                    4 => "dotted",
                    6 => "thick",
                    7 => "dash",
                    9 => "dotDash",
                    10 => "dotDotDash",
                    11 => "wave",
                    20 => "dottedHeavy",
                    23 => "dashedHeavy",
                    25 => "dashDotHeavy",
                    26 => "dashDotDotHeavy",
                    27 => "wavyHeavy",
                    39 => "dashLong",
                    43 => "wavyDouble",
                    55 => "dashLongHeavy",
                    _ => return Err(unsupported("invalid Word underline kind")),
                }
                .into(),
            ),
            0x6870 => (
                "color",
                match operand[3] {
                    0 => format!("{:02X}{:02X}{:02X}", operand[0], operand[1], operand[2]),
                    255 => "auto".into(),
                    _ => return Err(unsupported("invalid Word COLORREF")),
                },
            ),
            _ => return Ok(false),
        };
        self.values.insert(key, value);
        Ok(true)
    }

    pub fn xml(&self, fonts: &[String]) -> Result<String, String> {
        let mut xml = String::from("<w:rPr>");
        if self.fonts.iter().any(Option::is_some) && !fonts.is_empty() {
            xml.push_str("<w:rFonts");
            for (key, index) in ["ascii", "eastAsia", "hAnsi", "cs"].iter().zip(self.fonts) {
                if let Some(index) = index {
                    let name = fonts
                        .get(index)
                        .ok_or_else(|| unsupported("Word font index outside font table"))?;
                    xml.push_str(&format!(" w:{key}=\"{}\"", xml_attr(name)));
                }
            }
            xml.push_str("/>");
        } else if self.fonts.iter().flatten().any(|index| *index != 0) {
            return Err(unsupported("Word font index outside empty font table"));
        }
        for (key, value) in &self.values {
            xml.push_str(&format!("<w:{key} w:val=\"{value}\"/>"));
        }
        xml.push_str("</w:rPr>");
        Ok(xml)
    }
}

/// Character subset of [MS-DOC] Prm0.isprm. Non-character properties cannot
/// accidentally enter the character interpreter; unsupported entries warn.
pub fn prm0(prm: u16) -> Option<[u8; 3]> {
    let code: u16 = match (prm >> 1) & 127 {
        0x53 => 0x2a33,
        0x55 => 0x0835,
        0x56 => 0x0836,
        0x57 => 0x0837,
        0x58 => 0x0838,
        0x59 => 0x0839,
        0x5a => 0x083a,
        0x5b => 0x083b,
        0x5c => 0x083c,
        0x5e => 0x2a3e,
        0x68 => 0x2a48,
        0x73 => 0x2a53,
        _ => return None,
    };
    let [a, b] = code.to_le_bytes();
    Some([a, b, (prm >> 8) as u8])
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn picture_metadata_survives_style_reset_without_becoming_run_xml() {
        let base = Properties::default();
        let mut props = base.clone();
        for (code, bytes) in [
            (0x0855, vec![1]),
            (0x6a03, 123i32.to_le_bytes().to_vec()),
            (0x0856, vec![1]),
        ] {
            assert!(props.apply(code, &bytes, &base).unwrap());
        }
        props.reset_to(&base, true);
        assert!(props.picture.object);
        assert_eq!(props.picture.inline_location().unwrap(), None);
        props.reset_to(&base, false);
        assert_eq!(props.picture.inline_location().unwrap(), Some(123));
        assert!(!props.xml(&[]).unwrap().contains("123"));
        for code in [0x0806, 0x080a, 0x0856] {
            let mut active = props.clone();
            active.apply(code, &[1], &base).unwrap();
            active.apply(0x6a03, &(-1i32).to_le_bytes(), &base).unwrap();
            assert_eq!(
                active.picture.inline_location().unwrap(),
                None,
                "active objects must not dereference even an invalid offset"
            );
            assert!(active.apply(code, &[2], &base).is_err());
        }
        props.apply(0x6a03, &(-1i32).to_le_bytes(), &base).unwrap();
        assert!(props.picture.inline_location().is_err());
        props.apply(0x0855, &[0], &base).unwrap();
        assert_eq!(props.picture.inline_location().unwrap(), None);
    }
    #[test]
    fn repeated_toggle_is_relative_to_style_not_previous_direct_value() {
        let mut style = Properties::default();
        style.apply(0x0835, &[1], &Properties::default()).unwrap();
        let mut direct = style.clone();
        direct.apply(0x0835, &[0x81], &style).unwrap();
        direct.apply(0x0835, &[0x81], &style).unwrap();
        assert!(direct.xml(&[]).unwrap().contains("<w:b w:val=\"0\"/>"));
        direct.apply(0x0835, &[0x80], &style).unwrap();
        assert_eq!(direct, style);
        assert!(direct.apply(0x0835, &[2], &style).is_err());
    }

    #[test]
    fn preserves_font_slots_and_escapes_names_without_embedding_fonts() {
        let mut p = Properties::default();
        let base = p.clone();
        p.apply(0x4a4f, &[0, 0], &base).unwrap();
        p.apply(0x4a50, &[1, 0], &base).unwrap();
        p.apply(0x4a43, &[24, 0], &base).unwrap();
        let xml = p.xml(&["A & \"B\"".into(), "CJK".into()]).unwrap();
        assert!(xml.contains("w:ascii=\"A &amp; &quot;B&quot;\" w:eastAsia=\"CJK\""));
        assert!(xml.contains("w:sz w:val=\"24\""));
        assert!(p.xml(&[]).is_err());
    }

    #[test]
    fn style_reset_preserves_direction_but_removes_direct_font_size() {
        let base = Properties::default();
        let mut p = base.clone();
        p.apply(0x085a, &[1], &base).unwrap();
        p.apply(0x4a43, &[40, 0], &base).unwrap();
        p.reset_to(&base, false);
        let xml = p.xml(&[]).unwrap();
        assert!(xml.contains("w:sz w:val=\"20\""));
        assert!(xml.contains("w:rtl w:val=\"1\""));
    }
}
