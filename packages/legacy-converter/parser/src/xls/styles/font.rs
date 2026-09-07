//! Owned BIFF8 font semantics shared by legacy XML and future native projection.
//! [MS-XLS] 2.4.122 (Font, including family/charset) and 2.5.129 (FontIndex).

use super::super::{decode_biff_chars, u16_at, unsupported};

#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
pub(super) enum Underline {
    None,
    Single,
    Double,
    SingleAccounting,
    DoubleAccounting,
}

impl Underline {
    pub(super) fn xml_value(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Single => "single",
            Self::Double => "double",
            Self::SingleAccounting => "singleAccounting",
            Self::DoubleAccounting => "doubleAccounting",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
pub(super) enum Script {
    Baseline,
    Superscript,
    Subscript,
}

impl Script {
    pub(super) fn xml_value(self) -> &'static str {
        match self {
            Self::Baseline => "baseline",
            Self::Superscript => "superscript",
            Self::Subscript => "subscript",
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(super) struct ResolvedFont {
    pub(super) name: String,
    pub(super) size_twips: u16,
    pub(super) color_index: u16,
    pub(super) weight: u16,
    pub(super) family: u8,
    pub(super) charset: u8,
    pub(super) italic: bool,
    pub(super) strike: bool,
    pub(super) outline: bool,
    pub(super) shadow: bool,
    pub(super) condense: bool,
    pub(super) extend: bool,
    pub(super) underline: Underline,
    pub(super) script: Script,
}

impl ResolvedFont {
    pub(super) fn minimal_calibri() -> Self {
        Self {
            name: "Calibri".into(),
            size_twips: 220,
            color_index: 0x7fff,
            weight: 400,
            family: 0,
            charset: 0,
            italic: false,
            strike: false,
            outline: false,
            shadow: false,
            condense: false,
            extend: false,
            underline: Underline::None,
            script: Script::Baseline,
        }
    }

    pub(super) fn decode(data: &[u8]) -> Result<Self, String> {
        if data.len() < 16 {
            return Err(unsupported("truncated BIFF font"));
        }
        let (name, _) = decode_biff_chars(data, 16, usize::from(data[14]), data[15] & 1 != 0)?;
        let size_twips = u16_at(data, 0)?;
        let color_index = u16_at(data, 4)?;
        let weight = u16_at(data, 6)?;
        let flags = data[2];
        let underline = match data[10] {
            0 => Underline::None,
            1 => Underline::Single,
            2 => Underline::Double,
            0x21 => Underline::SingleAccounting,
            0x22 => Underline::DoubleAccounting,
            _ => return Err(unsupported("invalid BIFF underline")),
        };
        let script = match u16_at(data, 8)? {
            0 => Script::Baseline,
            1 => Script::Superscript,
            2 => Script::Subscript,
            _ => return Err(unsupported("invalid BIFF font script")),
        };
        Ok(Self {
            name,
            size_twips,
            color_index,
            weight,
            family: data[11],
            charset: data[12],
            italic: flags & 2 != 0,
            strike: flags & 8 != 0,
            outline: flags & 16 != 0,
            shadow: flags & 32 != 0,
            condense: flags & 64 != 0,
            extend: flags & 128 != 0,
            underline,
            script,
        })
    }

    pub(super) fn model(&self, color: Option<String>) -> xlsx_model::Font {
        xlsx_model::Font {
            bold: self.weight == 700,
            italic: self.italic,
            underline: self.underline != Underline::None,
            strike: self.strike,
            size: f64::from(self.size_twips) / 20.0,
            color,
            name: Some(self.name.clone()),
            underline_style: match self.underline {
                Underline::None | Underline::Single => None,
                value => Some(value.xml_value().to_string()),
            },
            vert_align: match self.script {
                Script::Baseline => None,
                value => Some(value.xml_value().to_string()),
            },
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn font(name: &[u8]) -> Vec<u8> {
        let mut data = vec![0; 16];
        data[0..2].copy_from_slice(&240u16.to_le_bytes());
        data[4..6].copy_from_slice(&10u16.to_le_bytes());
        data[6..8].copy_from_slice(&700u16.to_le_bytes());
        data[14] = name.len() as u8;
        data.extend_from_slice(name);
        data
    }

    #[test]
    fn decodes_supported_flags_underline_script_and_identity() {
        for (underline, expected) in [
            (0, Underline::None),
            (1, Underline::Single),
            (2, Underline::Double),
            (0x21, Underline::SingleAccounting),
            (0x22, Underline::DoubleAccounting),
        ] {
            let mut data = font(b"Arial");
            data[2] = 2 | 8 | 16 | 32 | 64 | 128;
            data[8..10].copy_from_slice(&2u16.to_le_bytes());
            data[10] = underline;
            data[11] = 3;
            data[12] = 0x80;
            let decoded = ResolvedFont::decode(&data).unwrap();
            assert_eq!(decoded.name, "Arial");
            assert_eq!(
                (decoded.size_twips, decoded.color_index, decoded.weight),
                (240, 10, 700)
            );
            assert_eq!((decoded.family, decoded.charset), (3, 0x80));
            assert!(decoded.italic && decoded.strike && decoded.outline);
            assert!(decoded.shadow && decoded.condense && decoded.extend);
            assert_eq!(decoded.underline, expected);
            assert_eq!(decoded.script, Script::Subscript);
        }
    }

    #[test]
    fn decodes_unicode_name_and_rejects_malformed_boundaries() {
        let mut unicode = font(&[]);
        unicode[14] = 2;
        unicode[15] = 1;
        unicode.extend_from_slice(&[0xE5, 0x65, 0x2C, 0x67]);
        assert_eq!(ResolvedFont::decode(&unicode).unwrap().name, "日本");

        assert!(ResolvedFont::decode(&[0; 15]).is_err());
        let mut truncated_name = font(b"A");
        truncated_name[14] = 2;
        assert!(ResolvedFont::decode(&truncated_name).is_err());
        let mut bad_underline = font(b"A");
        bad_underline[10] = 3;
        assert!(ResolvedFont::decode(&bad_underline).is_err());
        let mut bad_script = font(b"A");
        bad_script[8..10].copy_from_slice(&3u16.to_le_bytes());
        assert!(ResolvedFont::decode(&bad_script).is_err());
    }
}
