//! A bounded, explicit character-property subset. Unknown SPRMs are skipped
//! using their encoded operand size, never interpreted as text or executed.
//! [MS-DOC] 2.2.5, 2.6.1, 2.9.327; ECMA-376 17.3.2 (run properties).

use super::{border::ICO_COLORS, u16_at, u32_at, unsupported};
use crate::ooxml::xml_attr;
use std::collections::BTreeMap;

// MS-DOC 2.6.1 sprmCHighlight / 2.9.119 Ico. This is deliberately separate
// from ICO_COLORS: highlight is a symbolic OOXML palette, and controlled
// Word DOC roundtrips distinguish 0x0C darkMagenta from 0x0D darkRed despite
// the duplicated RGB entries in the published Ico table. The control covered
// direct CHPX operands in two fonts; it does not establish style inheritance.
const HIGHLIGHT_COLORS: [&str; 17] = [
    "none",
    "black",
    "blue",
    "cyan",
    "green",
    "magenta",
    "red",
    "yellow",
    "white",
    "darkBlue",
    "darkCyan",
    "darkGreen",
    "darkMagenta",
    "darkRed",
    "darkYellow",
    "darkGray",
    "lightGray",
];

#[cfg(feature = "direct-doc")]
mod direct;

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Properties {
    values: BTreeMap<&'static str, String>,
    pub fonts: [Option<usize>; 4],
    font_hint: Option<FontHint>,
    // Sparse style patches must distinguish an absent hint from 0xFF, which
    // explicitly removes inherited guidance but has no ST_Hint equivalent.
    font_hint_present: bool,
    // Raw MS-DOC LID. Resolution is deferred so an assigned locale and an
    // unknown/custom identifier never collapse into the same absent value.
    lang_bidi_lid: Option<u16>,
    // Raw language axes from MS-DOC 2.6.1. Modern properties supply the
    // projected language; compatibility properties remain acquisition metadata,
    // never a guessed fallback or an implicit noProof toggle.
    lang_default_80_lid: Option<u16>,
    lang_east_asia_80_lid: Option<u16>,
    lang_default_lid: Option<u16>,
    lang_east_asia_lid: Option<u16>,
    pub picture: Picture,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum FontHint {
    Default,
    EastAsia,
    ComplexScript,
}

enum LanguageResolution {
    Absent,
    Assigned(&'static str),
    Unsupported(u16),
}

#[derive(Clone, Copy)]
pub(super) struct ResolvedLanguages {
    pub(super) default: Option<&'static str>,
    pub(super) east_asia: Option<&'static str>,
    pub(super) bidi: Option<&'static str>,
}

impl FontHint {
    fn xml_value(self) -> &'static str {
        match self {
            Self::Default => "default",
            Self::EastAsia => "eastAsia",
            Self::ComplexScript => "cs",
        }
    }
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
            font_hint: None,
            font_hint_present: false,
            lang_bidi_lid: None,
            lang_default_80_lid: None,
            lang_east_asia_80_lid: None,
            lang_default_lid: None,
            lang_east_asia_lid: None,
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
            font_hint: None,
            font_hint_present: false,
            lang_bidi_lid: None,
            lang_default_80_lid: None,
            lang_east_asia_80_lid: None,
            lang_default_lid: None,
            lang_east_asia_lid: None,
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
        if patch.font_hint_present {
            self.font_hint = patch.font_hint;
            self.font_hint_present = true;
        }
        if patch.lang_bidi_lid.is_some() {
            self.lang_bidi_lid = patch.lang_bidi_lid;
        }
        for (current, added) in [
            &mut self.lang_default_80_lid,
            &mut self.lang_east_asia_80_lid,
            &mut self.lang_default_lid,
            &mut self.lang_east_asia_lid,
        ]
        .into_iter()
        .zip([
            patch.lang_default_80_lid,
            patch.lang_east_asia_80_lid,
            patch.lang_default_lid,
            patch.lang_east_asia_lid,
        ]) {
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
        let font_hint = self.font_hint;
        let font_hint_present = self.font_hint_present;
        if !preserve_object {
            picture.object = paragraph.picture.object;
        }
        let preserved: Vec<_> = ["rtl", "cs", "highlight", "webHidden"]
            .iter()
            .map(|key| (*key, self.values.get(key).cloned()))
            .collect();
        *self = paragraph.clone();
        self.picture = picture;
        // MS-DOC 2.6.1: sprmCPlain and sprmCIstd both leave any previous
        // sprmCIdctHint operand unaffected, including 0xFF (no guidance).
        self.font_hint = font_hint;
        self.font_hint_present = font_hint_present;
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
            // MS-DOC 2.6.1 sprmCFNoProof / 2.9.327 ToggleOperand maps to
            // ECMA-376 17.3.2.21. The byte route retains it; the display-only
            // direct viewer validates it without introducing proofing UI state.
            0x0875 => Some("noProof"),
            0x0882 => Some("cs"),
            _ => None,
        };
        if let Some(key) = flag {
            if operand.len() != 1 {
                return Err(unsupported("invalid Word character toggle"));
            }
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
            0x485f => {
                // MS-DOC 2.6.1 sprmCLidBi / 2.9.134 LID: this axis is used for
                // RTL or complex-script presentation. The language itself is
                // not required to belong to either family.
                if operand.len() != 2 {
                    return Err(unsupported("invalid Word complex-script language ID"));
                }
                self.lang_bidi_lid = Some(u16_at(operand, 0)?);
                return Ok(true);
            }
            0x486d | 0x486e | 0x4873 | 0x4874 => {
                // MS-DOC 2.6.1 sprmCRgLid0_80, sprmCRgLid1_80,
                // sprmCRgLid0, and sprmCRgLid1; their operand is the exact
                // two-byte LID from 2.9.134. Keep all four raw axes distinct.
                // Controlled Word roundtrips over the tested flag/order
                // combinations confirm that modern LIDs govern the emitted
                // language. Compatibility LIDs remain typed acquisition
                // metadata only; they never synthesize language or noProof.
                if operand.len() != 2 {
                    return Err(unsupported("invalid Word character language ID"));
                }
                let lid = Some(u16_at(operand, 0)?);
                match code {
                    0x486d => self.lang_default_80_lid = lid,
                    0x486e => self.lang_east_asia_80_lid = lid,
                    0x4873 => self.lang_default_lid = lid,
                    _ => self.lang_east_asia_lid = lid,
                }
                return Ok(true);
            }
            0x6815..=0x6817 => {
                // MS-DOC 2.6.1 identifies these as revision-session IDs for
                // character formatting, inserted text, and deleted text. They
                // do not themselves establish a revision mark. Viewer policy:
                // validate and omit this nonvisual provenance; the actual
                // CFRMarkIns/CFRMarkDel revision properties remain unsupported.
                if operand.len() != 4 {
                    return Err(unsupported("invalid Word character revision session ID"));
                }
                let _ = u32_at(operand, 0)?;
                return Ok(true);
            }
            0x2a0c => {
                let index = usize::from(
                    *operand
                        .first()
                        .ok_or_else(|| unsupported("short Word highlight index"))?,
                );
                (
                    "highlight",
                    HIGHLIGHT_COLORS
                        .get(index)
                        .ok_or_else(|| unsupported("invalid Word highlight index"))?
                        .to_string(),
                )
            }
            0x2a42 => {
                // MS-DOC 2.6.1 sprmCIco / 2.9.119 Ico: fixed palette index,
                // with zero representing automatic text color.
                let index = usize::from(
                    *operand
                        .first()
                        .ok_or_else(|| unsupported("short Word text color index"))?,
                );
                (
                    "color",
                    ICO_COLORS
                        .get(index)
                        .ok_or_else(|| unsupported("invalid Word text color index"))?
                        .to_string(),
                )
            }
            0x286f => {
                // MS-DOC 2.6.1 sprmCIdctHint. 0xFF is an explicit absence of
                // guidance and therefore cancels an inherited ST_Hint value.
                self.font_hint = match operand[0] {
                    0 => Some(FontHint::Default),
                    1 => Some(FontHint::EastAsia),
                    2 => Some(FontHint::ComplexScript),
                    0xff => None,
                    _ => return Err(unsupported("invalid Word character font hint")),
                };
                self.font_hint_present = true;
                return Ok(true);
            }
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
            0x6870 => ("color", colorref(operand)?),
            0x6877 => ("uColor", colorref(operand)?),
            _ => return Ok(false),
        };
        self.values.insert(key, value);
        Ok(true)
    }

    pub fn xml(&self, fonts: &[String]) -> Result<String, String> {
        self.xml_with_language_policy(fonts, false)
            .map(|(xml, _)| xml)
    }

    /// Preserve the byte adapter's warning-based omission policy for unresolved
    /// languages only. The caller must surface the returned omission flag;
    /// direct model projection remains strict and never guesses a locale.
    pub(in crate::doc) fn byte_xml(&self, fonts: &[String]) -> Result<(String, bool), String> {
        self.xml_with_language_policy(fonts, true)
    }

    fn xml_with_language_policy(
        &self,
        fonts: &[String],
        omit_unsupported_language: bool,
    ) -> Result<(String, bool), String> {
        let mut xml = String::from("<w:rPr>");
        let has_font_index = self.fonts.iter().any(Option::is_some);
        if self.font_hint.is_some() || (has_font_index && !fonts.is_empty()) {
            xml.push_str("<w:rFonts");
            if let Some(hint) = self.font_hint {
                xml.push_str(&format!(" w:hint=\"{}\"", hint.xml_value()));
            }
            if !fonts.is_empty() {
                for (key, index) in ["ascii", "eastAsia", "hAnsi", "cs"].iter().zip(self.fonts) {
                    if let Some(index) = index {
                        let name = fonts
                            .get(index)
                            .ok_or_else(|| unsupported("Word font index outside font table"))?;
                        xml.push_str(&format!(" w:{key}=\"{}\"", xml_attr(name)));
                    }
                }
            }
            xml.push_str("/>");
        } else if self.fonts.iter().flatten().any(|index| *index != 0) {
            return Err(unsupported("Word font index outside empty font table"));
        }
        for (key, value) in &self.values {
            if *key == "uColor" {
                continue;
            }
            if *key == "u" {
                xml.push_str(&format!("<w:u w:val=\"{value}\""));
                if let Some(color) = self.values.get("uColor") {
                    xml.push_str(&format!(" w:color=\"{color}\""));
                }
                xml.push_str("/>");
            } else {
                xml.push_str(&format!("<w:{key} w:val=\"{value}\"/>"));
            }
        }
        if !self.values.contains_key("u") {
            if let Some(color) = self.values.get("uColor") {
                // MS-DOC 2.6.1 sprmCCvUl retains underline color independently.
                // MS-OI29500 2.1.100(c), refining ECMA-376 17.3.2.40: absent
                // w:u@val inherits and ultimately means no underline; color
                // alone therefore preserves authorship without activating it.
                xml.push_str(&format!("<w:u w:color=\"{color}\"/>"));
            }
        }
        let axes = [
            ("val", "default", self.language(self.lang_default_lid)),
            (
                "eastAsia",
                "East Asian",
                self.language(self.lang_east_asia_lid),
            ),
            ("bidi", "complex-script", self.bidi_language()),
        ];
        let mut omitted_language = false;
        let mut language_started = false;
        for (attribute, name, value) in axes {
            match value {
                LanguageResolution::Absent => {}
                LanguageResolution::Assigned(language) => {
                    if !language_started {
                        xml.push_str("<w:lang");
                        language_started = true;
                    }
                    xml.push_str(&format!(" w:{attribute}=\"{language}\""));
                }
                LanguageResolution::Unsupported(_) if omit_unsupported_language => {
                    omitted_language = true;
                }
                LanguageResolution::Unsupported(lid) => {
                    return Err(unsupported(format!(
                        "unsupported Word {name} language ID 0x{lid:04X}"
                    )));
                }
            }
        }
        if language_started {
            xml.push_str("/>");
        }
        xml.push_str("</w:rPr>");
        Ok((xml, omitted_language))
    }

    pub(super) fn resolved_languages(&self) -> Result<ResolvedLanguages, String> {
        Ok(ResolvedLanguages {
            default: self.resolve_language_axis(self.lang_default_lid, "default")?,
            east_asia: self.resolve_language_axis(self.lang_east_asia_lid, "East Asian")?,
            bidi: self.resolve_language_axis(self.lang_bidi_lid, "complex-script")?,
        })
    }

    fn resolve_language_axis(
        &self,
        lid: Option<u16>,
        name: &str,
    ) -> Result<Option<&'static str>, String> {
        match self.language(lid) {
            LanguageResolution::Absent => Ok(None),
            LanguageResolution::Assigned(language) => Ok(Some(language)),
            LanguageResolution::Unsupported(lid) => Err(unsupported(format!(
                "unsupported Word {name} language ID 0x{lid:04X}"
            ))),
        }
    }

    fn bidi_language(&self) -> LanguageResolution {
        self.language(self.lang_bidi_lid)
    }

    fn language(&self, lid: Option<u16>) -> LanguageResolution {
        let Some(lid) = lid else {
            return LanguageResolution::Absent;
        };
        match crate::lcid::resolve(u32::from(lid)) {
            crate::lcid::Resolution::Assigned(language) => LanguageResolution::Assigned(language),
            _ => LanguageResolution::Unsupported(lid),
        }
    }
}

fn colorref(operand: &[u8]) -> Result<String, String> {
    match operand.get(3).copied() {
        Some(0) => Ok(format!(
            "{:02X}{:02X}{:02X}",
            operand[0], operand[1], operand[2]
        )),
        Some(0xff) => Ok("auto".into()),
        Some(_) => Err(unsupported("invalid Word COLORREF")),
        None => Err(unsupported("short Word COLORREF")),
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
    use crate::doc::sprm::{Budget, Sprms};

    #[test]
    fn indexed_text_color_uses_shared_palette_and_normal_cascade_order() {
        let base = Properties::default();
        for (index, expected) in ICO_COLORS.iter().enumerate() {
            let mut value = base.clone();
            assert!(value.apply(0x2a42, &[index as u8], &base).unwrap());
            assert!(value
                .xml(&[])
                .unwrap()
                .contains(&format!("<w:color w:val=\"{expected}\"/>")));
        }
        for index in 17..=255u8 {
            assert!(base.clone().apply(0x2a42, &[index], &base).is_err());
        }
        assert!(base.clone().apply(0x2a42, &[], &base).is_err());

        let mut style = base.clone();
        style.apply(0x2a42, &[2], &base).unwrap();
        let mut value = style.clone();
        value.apply(0x6870, &[0x12, 0x34, 0x56, 0], &style).unwrap();
        assert!(value.xml(&[]).unwrap().contains("w:val=\"123456\""));
        value.apply(0x2a42, &[6], &style).unwrap();
        assert!(value.xml(&[]).unwrap().contains("w:val=\"FF0000\""));
        value.reset_to(&style, false);
        assert!(value.xml(&[]).unwrap().contains("w:val=\"0000FF\""));
    }

    #[test]
    fn highlight_uses_symbolic_palette_and_explicit_none() {
        let base = Properties::default();
        for (index, expected) in HIGHLIGHT_COLORS.iter().enumerate() {
            let mut value = base.clone();
            assert!(value.apply(0x2a0c, &[index as u8], &base).unwrap());
            assert!(value
                .xml(&[])
                .unwrap()
                .contains(&format!("<w:highlight w:val=\"{expected}\"/>")));
        }
        for index in 17..=255u8 {
            assert!(base.clone().apply(0x2a0c, &[index], &base).is_err());
        }
        assert!(base.clone().apply(0x2a0c, &[], &base).is_err());

        let mut style = base.clone();
        style.apply(0x2a0c, &[12], &base).unwrap();
        let mut value = style.clone();
        value.apply(0x2a0c, &[13], &style).unwrap();
        assert!(value.xml(&[]).unwrap().contains("w:val=\"darkRed\""));
        value.apply(0x2a0c, &[0], &style).unwrap();
        assert!(value.xml(&[]).unwrap().contains("w:val=\"none\""));

        // MS-DOC 2.6.1 explicitly excludes highlight from both CPlain and
        // CIstd resets. This tests the shared reset primitive used by both.
        value.reset_to(&base, false);
        assert!(value.xml(&[]).unwrap().contains("w:val=\"none\""));
        value.apply(0x2a0c, &[12], &base).unwrap();
        value.reset_to(&base, true);
        assert!(value.xml(&[]).unwrap().contains("w:val=\"darkMagenta\""));
    }

    #[test]
    fn revision_session_ids_are_validated_nonvisual_viewer_provenance() {
        let base = Properties::default();
        for code in 0x6815..=0x6817 {
            for value in [0, 0x7856_3412, u32::MAX] {
                let mut properties = base.clone();
                assert!(properties.apply(code, &value.to_le_bytes(), &base).unwrap());
                assert_eq!(properties.xml(&[]).unwrap(), base.xml(&[]).unwrap());
            }
            assert!(base.clone().apply(code, &[0; 3], &base).is_err());
            assert!(base.clone().apply(code, &[0; 5], &base).is_err());
        }

        let mut adjacent = base.clone();
        adjacent
            .apply(0x6816, &0x7856_3412u32.to_le_bytes(), &base)
            .unwrap();
        adjacent.apply(0x0835, &[1], &base).unwrap();
        assert!(adjacent.xml(&[]).unwrap().contains("<w:b w:val=\"1\"/>"));
        adjacent.reset_to(&base, false);
        assert_eq!(adjacent.xml(&[]).unwrap(), base.xml(&[]).unwrap());
        // A session ID associated with deletion cannot manufacture a deletion.
        assert!(!adjacent.apply(0x0800, &[1], &base).unwrap());

        let truncated = [0x16, 0x68, 0x12, 0x34, 0x56];
        assert!(Sprms::new(&truncated).next(&mut Budget::default()).is_err());
    }

    #[test]
    fn complex_script_language_is_typed_cascading_and_strictly_resolved() {
        let base = Properties::default();
        let mut style = Properties::sparse();
        assert!(style
            .apply(0x485f, &0x0401u16.to_le_bytes(), &base)
            .unwrap());
        assert!(style
            .xml(&[])
            .unwrap()
            .contains("<w:lang w:bidi=\"ar-SA\"/>"));

        let mut patch = Properties::sparse();
        patch
            .apply(0x485f, &0x0411u16.to_le_bytes(), &base)
            .unwrap();
        style.overlay_visible(&patch);
        assert!(style.xml(&[]).unwrap().contains("w:bidi=\"ja-JP\""));
        style.overlay_visible(&Properties::sparse());
        assert!(style.xml(&[]).unwrap().contains("w:bidi=\"ja-JP\""));
        style
            .apply(0x485f, &0x0409u16.to_le_bytes(), &base)
            .unwrap();
        assert!(style.xml(&[]).unwrap().contains("w:bidi=\"en-US\""));

        // Both CPlain and CIstd reset language: neither preserved-property
        // list in MS-DOC 2.6.1 includes language.
        let mut paragraph = base.clone();
        paragraph
            .apply(0x485f, &0x0411u16.to_le_bytes(), &base)
            .unwrap();
        style.reset_to(&paragraph, false);
        assert!(style.xml(&[]).unwrap().contains("w:bidi=\"ja-JP\""));
        style
            .apply(0x485f, &0x0401u16.to_le_bytes(), &paragraph)
            .unwrap();
        style.reset_to(&paragraph, true);
        assert!(style.xml(&[]).unwrap().contains("w:bidi=\"ja-JP\""));
        for preserve_object in [false, true] {
            let mut reset = style.clone();
            reset.reset_to(&base, preserve_object);
            assert!(!reset.xml(&[]).unwrap().contains("<w:lang"));
        }

        assert!(base.clone().apply(0x485f, &[1], &base).is_err());
        assert!(base.clone().apply(0x485f, &[1, 2, 3], &base).is_err());
        for lid in [u16::MAX, 0x1000, 0x0400, 0x007f, 0x0467, 0x040a] {
            let mut unresolved = base.clone();
            unresolved.apply(0x485f, &lid.to_le_bytes(), &base).unwrap();
            assert!(unresolved.xml(&[]).is_err(), "LID {lid:04x}");
        }

        let mut truncated = Sprms::new(&[0x5f, 0x48, 0x01]);
        assert!(truncated.next(&mut Budget::default()).is_err());
        let framed = [0x5f, 0x48, 0x01, 0x04, 0x35, 0x08, 0x01];
        let mut framed = Sprms::new(&framed);
        let mut budget = Budget::default();
        assert_eq!(
            framed.next(&mut budget).unwrap(),
            Some((0x485f, &[0x01, 0x04][..]))
        );
        assert_eq!(framed.next(&mut budget).unwrap(), Some((0x0835, &[1][..])));
    }

    #[test]
    fn modern_languages_project_while_compatibility_languages_remain_metadata() {
        let base = Properties::default();
        let cases: [(u16, u16, usize); 4] = [
            (0x486d, 0x0400, 0),
            (0x486e, 0x007f, 1),
            (0x4873, 0x040c, 2),
            (0x4874, 0x0411, 3),
        ];
        let slots = |p: &Properties| {
            [
                p.lang_default_80_lid,
                p.lang_east_asia_80_lid,
                p.lang_default_lid,
                p.lang_east_asia_lid,
            ]
        };

        let mut properties = Properties::sparse();
        for (code, lid, slot) in cases {
            assert!(properties.apply(code, &lid.to_le_bytes(), &base).unwrap());
            assert_eq!(slots(&properties)[slot], Some(lid));
        }
        assert_eq!(
            slots(&properties),
            [Some(0x0400), Some(0x007f), Some(0x040c), Some(0x0411)]
        );
        assert_eq!(
            properties.xml(&[]).unwrap(),
            "<w:rPr><w:lang w:val=\"fr-FR\" w:eastAsia=\"ja-JP\"/></w:rPr>"
        );

        // Repeated properties retain their ordered last value without merging
        // the four source axes.
        assert!(properties
            .apply(0x4873, &0x1234u16.to_le_bytes(), &base)
            .unwrap());
        assert_eq!(slots(&properties)[2], Some(0x1234));
        assert_eq!(slots(&properties)[3], Some(0x0411));

        let mut inherited = Properties::sparse();
        inherited
            .apply(0x486d, &0x1111u16.to_le_bytes(), &base)
            .unwrap();
        inherited
            .apply(0x4874, &0x0404u16.to_le_bytes(), &base)
            .unwrap();
        let mut patch = Properties::sparse();
        patch
            .apply(0x486e, &0x3333u16.to_le_bytes(), &base)
            .unwrap();
        inherited.overlay_visible(&patch);
        assert_eq!(
            slots(&inherited),
            [Some(0x1111), Some(0x3333), None, Some(0x0404)]
        );
        inherited.overlay_visible(&Properties::sparse());
        assert_eq!(
            slots(&inherited),
            [Some(0x1111), Some(0x3333), None, Some(0x0404)]
        );

        // CPlain and CIstd use this reset primitive. Language is not among
        // their preserved properties, so all four axes come from the reset
        // target while bidi/noProof remain independent properties.
        let mut paragraph = Properties::sparse();
        paragraph
            .apply(0x4873, &0x0409u16.to_le_bytes(), &base)
            .unwrap();
        inherited
            .apply(0x485f, &0x0401u16.to_le_bytes(), &base)
            .unwrap();
        inherited.apply(0x0875, &[1], &base).unwrap();
        let independent_xml = inherited.xml(&[]).unwrap();
        assert!(independent_xml.contains("w:eastAsia=\"zh-TW\""));
        assert!(independent_xml.contains("w:bidi=\"ar-SA\""));
        assert!(independent_xml.contains("<w:noProof w:val=\"1\"/>"));
        for preserve_object in [false, true] {
            let mut reset = inherited.clone();
            reset.reset_to(&paragraph, preserve_object);
            assert_eq!(slots(&reset), [None, None, Some(0x0409), None]);
            assert_eq!(reset.lang_bidi_lid, paragraph.lang_bidi_lid);
            assert!(!reset.xml(&[]).unwrap().contains("<w:noProof"));
        }
    }

    #[test]
    fn western_and_east_asian_language_ids_require_exact_operands() {
        let base = Properties::default();
        for code in [0x486d, 0x486e, 0x4873, 0x4874] {
            assert!(base.clone().apply(code, &[], &base).is_err());
            assert!(base.clone().apply(code, &[1], &base).is_err());
            assert!(base.clone().apply(code, &[1, 2, 3], &base).is_err());
        }

        let mut truncated = Sprms::new(&[0x6d, 0x48, 0x01]);
        assert!(truncated.next(&mut Budget::default()).is_err());
        let framed = [0x6d, 0x48, 0x01, 0x04, 0x35, 0x08, 0x01];
        let mut framed = Sprms::new(&framed);
        let mut budget = Budget::default();
        assert_eq!(
            framed.next(&mut budget).unwrap(),
            Some((0x486d, &[0x01, 0x04][..]))
        );
        assert_eq!(framed.next(&mut budget).unwrap(), Some((0x0835, &[1][..])));
    }

    #[test]
    fn underline_color_is_an_attribute_and_follows_style_resets() {
        let base = Properties::default();
        let mut style = base.clone();
        style.apply(0x2a3e, &[11], &base).unwrap();
        style.apply(0x6877, &[0x12, 0x34, 0x56, 0], &base).unwrap();
        assert!(style
            .xml(&[])
            .unwrap()
            .contains("<w:u w:val=\"wave\" w:color=\"123456\"/>"));
        let mut direct = style.clone();
        direct.apply(0x6877, &[0, 0, 0, 0xff], &style).unwrap();
        assert!(direct.xml(&[]).unwrap().contains("w:color=\"auto\""));
        direct.reset_to(&base, false);
        assert!(!direct.xml(&[]).unwrap().contains("w:color="));
        assert!(base.clone().apply(0x6877, &[], &base).is_err());
        assert!(base.clone().apply(0x6877, &[1, 2, 3, 1], &base).is_err());

        let mut color_only = base.clone();
        color_only
            .apply(0x6877, &[0x12, 0x34, 0x56, 0], &base)
            .unwrap();
        assert!(color_only
            .xml(&[])
            .unwrap()
            .contains("<w:u w:color=\"123456\"/>"));
        color_only.apply(0x2a3e, &[4], &base).unwrap();
        assert!(color_only
            .xml(&[])
            .unwrap()
            .contains("<w:u w:val=\"dotted\" w:color=\"123456\"/>"));
    }

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
        assert!(direct.apply(0x0835, &[], &style).is_err());
        assert!(direct.apply(0x0835, &[1, 0], &style).is_err());
    }

    #[test]
    fn no_proof_is_a_validated_style_relative_toggle_and_resets_normally() {
        let base = Properties::default();
        let mut style = base.clone();
        assert!(style.apply(0x0875, &[1], &base).unwrap());
        assert!(style.xml(&[]).unwrap().contains("<w:noProof w:val=\"1\"/>"));

        for (operand, expected) in [(0, "0"), (1, "1"), (0x80, "1"), (0x81, "0")] {
            let mut direct = style.clone();
            assert!(direct.apply(0x0875, &[operand], &style).unwrap());
            assert!(direct
                .xml(&[])
                .unwrap()
                .contains(&format!("<w:noProof w:val=\"{expected}\"/>")));
        }

        let mut false_style = base.clone();
        false_style.apply(0x0875, &[0], &base).unwrap();
        let mut opposite = false_style.clone();
        opposite.apply(0x0875, &[0x81], &false_style).unwrap();
        assert!(opposite
            .xml(&[])
            .unwrap()
            .contains("<w:noProof w:val=\"1\"/>"));
        opposite.apply(0x0875, &[0], &false_style).unwrap();
        assert!(opposite
            .xml(&[])
            .unwrap()
            .contains("<w:noProof w:val=\"0\"/>"));

        let mut inherited = style.clone();
        inherited.overlay_visible(&Properties::sparse());
        assert!(inherited
            .xml(&[])
            .unwrap()
            .contains("<w:noProof w:val=\"1\"/>"));
        for preserve_object in [false, true] {
            let mut reset = inherited.clone();
            reset.reset_to(&false_style, preserve_object);
            assert!(reset.xml(&[]).unwrap().contains("<w:noProof w:val=\"0\"/>"));
        }

        for operand in [vec![], vec![2], vec![0x82], vec![1, 0]] {
            assert!(base.clone().apply(0x0875, &operand, &style).is_err());
        }

        assert!(Sprms::new(&[0x75, 0x08])
            .next(&mut Budget::default())
            .is_err());
        let mut adjacent = Sprms::new(&[0x75, 0x08, 1, 0x35, 0x08, 1]);
        let mut budget = Budget::default();
        assert_eq!(
            adjacent.next(&mut budget).unwrap(),
            Some((0x0875, &[1][..]))
        );
        assert_eq!(
            adjacent.next(&mut budget).unwrap(),
            Some((0x0835, &[1][..]))
        );
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
    fn preserves_all_documented_font_hints_without_changing_font_slots() {
        let mut p = Properties::default();
        let base = p.clone();
        for (code, bytes) in [
            (0x4a4f, [0, 0]),
            (0x4a50, [1, 0]),
            (0x4a51, [2, 0]),
            (0x4a5e, [3, 0]),
        ] {
            assert!(p.apply(code, &bytes, &base).unwrap());
        }
        let fonts = ["ASCII", "East Asia", "High ANSI", "Complex Script"].map(String::from);
        for (value, expected) in [(0, "default"), (1, "eastAsia"), (2, "cs")] {
            assert!(p.apply(0x286f, &[value], &base).unwrap());
            let xml = p.xml(&fonts).unwrap();
            assert!(xml.contains(&format!("w:hint=\"{expected}\"")), "{xml}");
            for (slot, name) in [
                ("ascii", "ASCII"),
                ("eastAsia", "East Asia"),
                ("hAnsi", "High ANSI"),
                ("cs", "Complex Script"),
            ] {
                assert!(xml.contains(&format!("w:{slot}=\"{name}\"")), "{xml}");
            }
        }
    }

    #[test]
    fn no_guidance_cancels_inherited_hint_and_absence_emits_no_font_element() {
        let base = Properties::default();
        assert!(!base.xml(&[]).unwrap().contains("<w:rFonts"));

        let mut inherited = base.clone();
        inherited.apply(0x286f, &[1], &base).unwrap();
        assert!(inherited.xml(&[]).unwrap().contains("w:hint=\"eastAsia\""));
        inherited.apply(0x286f, &[0xff], &base).unwrap();
        assert!(!inherited.xml(&[]).unwrap().contains("<w:rFonts"));
        assert!(inherited.apply(0x286f, &[3], &base).is_err());
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
