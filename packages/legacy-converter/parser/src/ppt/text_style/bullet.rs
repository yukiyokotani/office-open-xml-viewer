//! MS-PPT 2.9.20-22 bullet masks and values inherit independently.
use super::*;

#[derive(Clone, Default, PartialEq)]
pub(super) struct Bullet {
    pub(in crate::ppt::text_style) enabled: Option<bool>,
    pub(in crate::ppt::text_style) has_font: Option<bool>,
    pub(in crate::ppt::text_style) has_color: Option<bool>,
    pub(in crate::ppt::text_style) has_size: Option<bool>,
    pub(in crate::ppt::text_style) character: Option<u16>,
    pub(in crate::ppt::text_style) font: Option<u16>,
    pub(in crate::ppt::text_style) size: Option<i16>,
    pub(in crate::ppt::text_style) color: Option<u32>,
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
}
