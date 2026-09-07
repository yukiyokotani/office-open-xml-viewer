//! Neutral BIFF/SML color identity retained before renderer color resolution.
//! Identity matters for style-table interning even when two colors paint alike.

#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
pub(super) enum ColorIdentity {
    Auto,
    Indexed(u16),
    Argb([u8; 4]),
}

impl ColorIdentity {
    pub(super) fn xml(self) -> String {
        match self {
            Self::Auto => "auto=\"1\"".into(),
            Self::Indexed(index) => format!("indexed=\"{index}\""),
            Self::Argb([a, r, g, b]) => format!("rgb=\"{a:02X}{r:02X}{g:02X}{b:02X}\""),
        }
    }

    pub(super) fn model(self) -> Option<String> {
        use ooxml_common::spreadsheet_color::{resolve_color, SpreadsheetColor};
        let color = match self {
            Self::Auto => SpreadsheetColor::Auto,
            Self::Indexed(index) => SpreadsheetColor::Indexed(index.into()),
            Self::Argb(argb) => SpreadsheetColor::Argb(argb),
        };
        resolve_color(color, None, &[])
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn identities_preserve_source_distinctions_and_exact_xml() {
        assert_ne!(ColorIdentity::Indexed(0), ColorIdentity::Indexed(8));
        assert_ne!(ColorIdentity::Auto, ColorIdentity::Indexed(0x7fff));
        assert_ne!(
            ColorIdentity::Argb([0x80, 0x12, 0x34, 0x56]),
            ColorIdentity::Argb([0xff, 0x12, 0x34, 0x56]),
        );
        assert_eq!(ColorIdentity::Auto.xml(), "auto=\"1\"");
        assert_eq!(ColorIdentity::Indexed(8).xml(), "indexed=\"8\"");
        assert_eq!(
            ColorIdentity::Argb([0xab, 0x01, 0x2c, 0xff]).xml(),
            "rgb=\"AB012CFF\"",
        );
    }
}
