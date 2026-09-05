//! MS-PPT 2.9.20-22 bullet masks and values inherit independently.
use super::*;

#[derive(Clone, Default, PartialEq)]
pub(super) struct Bullet {
    enabled: Option<bool>,
    has_font: Option<bool>,
    has_color: Option<bool>,
    has_size: Option<bool>,
    character: Option<u16>,
    font: Option<u16>,
    size: Option<i16>,
    color: Option<u32>,
}
impl Bullet {
    pub fn read(reader: &mut Reader<'_, '_>, mask: u32) -> Result<Self, String> {
        let flags = reader.optional16(mask, 15)?.unwrap_or(0);
        let flag = |bit| (mask & bit != 0).then_some(flags & bit as u16 != 0);
        Ok(Self {
            enabled: flag(1),
            has_font: flag(2),
            has_color: flag(4),
            has_size: flag(8),
            character: reader.optional16(mask, 0x80)?,
            font: reader.optional16(mask, 0x10)?,
            size: reader.optional16(mask, 0x40)?.map(|v| v as i16),
            color: if mask & 0x20 != 0 {
                Some(reader.u32()?)
            } else {
                None
            },
        })
    }
    pub fn inherit(&self, base: &Self) -> Self {
        Self {
            enabled: self.enabled.or(base.enabled),
            has_font: self.has_font.or(base.has_font),
            has_color: self.has_color.or(base.has_color),
            has_size: self.has_size.or(base.has_size),
            character: self.character.or(base.character),
            font: self.font.or(base.font),
            size: self.size.or(base.size),
            color: self.color.or(base.color),
        }
    }
    pub fn xml(&self, context: Context<'_>) -> String {
        if self.enabled == Some(false) {
            return "<a:buNone/>".into();
        }
        if self.enabled != Some(true) {
            return String::new();
        }
        // One BMP scalar is representable; controls, isolated surrogates and
        // invalid XML characters are unsupported, not replacement bullet glyphs.
        let Some(character) = self
            .character
            .and_then(|v| char::from_u32(u32::from(v)))
            .filter(|c| !c.is_control() && !matches!(*c, '\u{fffe}' | '\u{ffff}'))
        else {
            return "<a:buNone/>".into();
        };
        let mut xml = String::new();
        // DrawingML CT_TextParagraphProperties requires color, size, font,
        // then the bullet choice in this order. False flags override inheritance.
        if self.has_color == Some(false) {
            xml.push_str("<a:buClrTx/>");
        } else if self.has_color == Some(true) {
            if let Some(color) = self.color.and_then(|c| scheme::text(c, context.scheme)) {
                xml.push_str(&format!(
                    "<a:buClr><a:srgbClr val=\"{:02X}{:02X}{:02X}\"/></a:buClr>",
                    color & 255,
                    (color >> 8) & 255,
                    (color >> 16) & 255
                ));
            }
        }
        if self.has_size == Some(false) {
            xml.push_str("<a:buSzTx/>");
        } else if self.has_size == Some(true) {
            // MS-PPT 2.2.3: positive percentages; negative absolute points.
            // Omit values outside the supported DrawingML range, never clamp.
            match self.size {
                Some(size @ 25..=400) => {
                    xml.push_str(&format!("<a:buSzPct val=\"{}\"/>", i32::from(size) * 1000))
                }
                Some(size @ -4000..=-1) => {
                    xml.push_str(&format!("<a:buSzPts val=\"{}\"/>", -i32::from(size) * 100))
                }
                _ => {}
            }
        }
        if self.has_font == Some(false) {
            xml.push_str("<a:buFontTx/>");
        } else if self.has_font == Some(true) {
            if let Some(font) = self.font.and_then(|f| context.fonts.get(usize::from(f))) {
                xml.push_str(&format!(
                    "<a:buFont typeface=\"{}\"/>",
                    crate::ooxml::xml_attr(font)
                ));
            }
        }
        xml.push_str(&format!(
            "<a:buChar char=\"{}\"/>",
            crate::ooxml::xml_attr(&character.to_string())
        ));
        xml
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    fn read(mask: u32, bytes: &[u8]) -> Bullet {
        Bullet::read(
            &mut Reader {
                bytes,
                pos: 0,
                budget: &mut 100,
            },
            mask,
        )
        .unwrap()
    }
    fn base() -> Bullet {
        Bullet {
            enabled: Some(true),
            has_color: Some(true),
            has_font: Some(true),
            has_size: Some(true),
            character: Some(0x2022),
            font: Some(0),
            size: Some(75),
            color: Some(0x01000000),
        }
    }
    #[test]
    fn flags_and_values_inherit_independently_in_schema_order() {
        // Only the enabled bit is valid; zero bits for the other flags in this
        // same word must not override inherited font/color/size flags.
        let direct = read(1, &[1, 0]);
        let fonts = ["Arial".to_string()];
        let scheme = [0x123456; 8];
        let context = Context {
            fonts: &fonts,
            scheme: Some(&scheme),
            levels: None,
            ..Context::default()
        };
        let inherited = direct.inherit(&base());
        let xml = inherited.xml(context);
        assert!(xml.contains("<a:srgbClr val=\"563412\"/>"));
        assert!(xml.contains("<a:buSzPct val=\"75000\"/>"));
        assert!(xml.contains("<a:buFont typeface=\"Arial\"/>"));
        let positions = ["buClr>", "buSzPct", "buFont ", "buChar"].map(|s| xml.find(s).unwrap());
        assert!(positions.windows(2).all(|w| w[0] < w[1]));
        let disabled = read(1, &[0, 0]).inherit(&inherited);
        assert_eq!(disabled.xml(context), "<a:buNone/>");
        let follow = read(14, &[0, 0]).inherit(&inherited).xml(context);
        assert!(follow.contains("<a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buChar"));
        assert!(!follow.contains("563412"));
        let changed = read(0x40, &50u16.to_le_bytes())
            .inherit(&inherited)
            .xml(context);
        assert!(changed.contains("<a:buSzPct val=\"50000\"/>"));
        // A glyph without a valid enabled flag must not invent a list.
        assert!(read(0x80, &0x2022u16.to_le_bytes()).xml(context).is_empty());
    }
    #[test]
    fn bullet_size_boundaries_and_unsupported_glyphs_never_get_clamped() {
        let context = Context::default();
        for (size, tag, val) in [
            (25, "buSzPct", 25000),
            (400, "buSzPct", 400000),
            (-1, "buSzPts", 100),
            (-4000, "buSzPts", 400000),
        ] {
            let mut bullet = base();
            bullet.size = Some(size);
            assert!(bullet
                .xml(context)
                .contains(&format!("<a:{tag} val=\"{val}\"/>")));
        }
        for size in [i16::MIN, -4001, 0, 24, 401, i16::MAX] {
            let mut bullet = base();
            bullet.size = Some(size);
            assert!(!bullet.xml(context).contains("buSz"));
        }
        for character in [0, 9, 10, 13, 0xd800, 0xdfff, 0xfffe, 0xffff] {
            let mut bullet = base();
            bullet.character = Some(character);
            assert_eq!(bullet.xml(context), "<a:buNone/>");
        }
        let mut bullet = base();
        bullet.character = Some('"' as u16);
        assert!(bullet.xml(context).contains("char=\"&quot;\""));
        bullet.font = Some(u16::MAX);
        assert!(!bullet.xml(context).contains("buFont"));
        assert!(!bullet.xml(context).contains("buClr>")); // no scheme guessed
        for bytes in [vec![], vec![0], vec![1, 0, 0]] {
            assert!(Bullet::read(
                &mut Reader {
                    bytes: &bytes,
                    pos: 0,
                    budget: &mut 100
                },
                0xff
            )
            .is_err());
        }
    }
}
