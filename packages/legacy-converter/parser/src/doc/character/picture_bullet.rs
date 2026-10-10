//! [MS-DOC] 2.6.1 sprmCPbiIBullet / sprmCPbiGrf and 2.9.176 PbiGrfOperand:
//! typed picture-bullet facts retained with presence and winning layer.
//!
//! Both SPRMs are character properties of a paragraph mark (IBullet also of a
//! cell or section mark). The legal layers are the mark's direct CHPX, its
//! piece PRM and the list level's grpprlChpx. [MS-DOC] 2.9.336 forbids both
//! in a style's UpxChpx; a referenced style's value is kept with a `Style`
//! origin only to decide ownership: disabled it changes nothing, a legal
//! later layer replaces it, and an enabled winning one is refused. sprmCPlain
//! and sprmCIstd do not list them among their preserved properties, so either
//! reset returns them to the paragraph (or character) style value.
//!
//! Strict display keeps the atomic unsupported refusal: no Word size is
//! derived from fNoAutoSize. An explicit stored-size reading consumer may
//! retain the raw operands and legal source owners independently of that
//! unresolved Word metric.

use super::super::{u32_at, unsupported};

/// The cascade layer that last wrote a picture-bullet operand. `Decoded`
/// marks a write whose caller supplied no layer (unit decoding).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(in crate::doc) enum BulletOrigin {
    Decoded,
    Direct {
        fc: usize,
    },
    Piece {
        fc: usize,
        prm: u16,
    },
    ListLevel {
        instance: usize,
        list: usize,
        level: u8,
    },
    /// A referenced style's UpxChpx, where [MS-DOC] 2.9.336 forbids both
    /// operands. Retained only to decide ownership; never a legal source.
    Style(usize),
}

/// [MS-DOC] 2.9.176 PbiGrfOperand without its 14 ignored fUnused bits.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(in crate::doc) struct PbiGrf {
    /// fPicBullet: the bullet is a picture.
    pub(in crate::doc) picture: bool,
    /// fNoAutoSize: the picture is not resized to the following text. No
    /// automatic-size metric is specified or derived here.
    pub(in crate::doc) no_auto_size: bool,
}

/// An enabled picture bullet: its CP relative to the start of the hidden
/// `_PictureBullets` bookmark and its flags.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(in crate::doc) struct EnabledPictureBullet {
    pub(in crate::doc) relative_cp: u32,
    pub(in crate::doc) flags: PbiGrf,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub(in crate::doc) struct PictureBullet {
    index: Option<(u32, BulletOrigin)>,
    flags: Option<(PbiGrf, BulletOrigin)>,
    raw_flags: u16,
}

impl PictureBullet {
    /// Decode one operand. Returns `None` for other codes, otherwise whether
    /// the write leaves the run displayable by the current consumer: an
    /// enabled fPicBullet is not.
    pub(super) fn apply(&mut self, code: u16, operand: &[u8]) -> Result<Option<bool>, String> {
        match code {
            0x6887 => {
                // A CP value that MUST be greater than or equal to zero.
                if operand.len() != 4 || (u32_at(operand, 0)? as i32) < 0 {
                    return Err(unsupported("invalid Word picture bullet position"));
                }
                self.index = Some((u32_at(operand, 0)?, BulletOrigin::Decoded));
                Ok(Some(true))
            }
            0x4888 => {
                if operand.len() != 2 {
                    return Err(unsupported("invalid Word picture bullet flags"));
                }
                let flags = PbiGrf {
                    picture: operand[0] & 1 != 0,
                    no_auto_size: operand[0] & 2 != 0,
                };
                self.raw_flags = u16::from_le_bytes([operand[0], operand[1]]);
                self.flags = Some((flags, BulletOrigin::Decoded));
                Ok(Some(!flags.picture))
            }
            _ => Ok(None),
        }
    }

    pub(in crate::doc) fn reading_source_owners(
        &self,
    ) -> Result<
        (
            u16,
            docx_model::NativePictureBulletOrigin,
            docx_model::NativePictureBulletOrigin,
        ),
        String,
    > {
        fn owner(origin: BulletOrigin) -> Result<docx_model::NativePictureBulletOrigin, String> {
            Ok(match origin {
                BulletOrigin::Direct { fc } => docx_model::NativePictureBulletOrigin::Direct { fc },
                BulletOrigin::Piece { fc, prm } => {
                    docx_model::NativePictureBulletOrigin::Piece { fc, prm }
                }
                BulletOrigin::ListLevel {
                    instance,
                    list,
                    level,
                } => docx_model::NativePictureBulletOrigin::ListLevel {
                    instance,
                    list,
                    level,
                },
                _ => {
                    return Err(unsupported(
                        "Word picture bullet lacks a legal source owner",
                    ))
                }
            })
        }
        self.enabled()?
            .ok_or_else(|| unsupported("disabled Word reading picture bullet"))?;
        let flags = self
            .flags
            .ok_or_else(|| unsupported("Word picture bullet lacks flags"))?;
        let index = self
            .index
            .ok_or_else(|| unsupported("Word picture bullet lacks position"))?;
        Ok((self.raw_flags, owner(flags.1)?, owner(index.1)?))
    }

    /// Attribute the operand `code` just applied to its cascade layer.
    pub(super) fn stamp(&mut self, code: u16, origin: BulletOrigin) {
        match code {
            0x6887 => {
                if let Some((_, slot)) = &mut self.index {
                    *slot = origin;
                }
            }
            0x4888 => {
                if let Some((_, slot)) = &mut self.flags {
                    *slot = origin;
                }
            }
            _ => {}
        }
    }

    #[allow(
        dead_code,
        reason = "native acquisition fact; picture-bullet consumer pending"
    )]
    pub(in crate::doc) fn index(&self) -> Option<(u32, BulletOrigin)> {
        self.index
    }

    #[allow(
        dead_code,
        reason = "native acquisition fact; picture-bullet consumer pending"
    )]
    pub(in crate::doc) fn flags(&self) -> Option<(PbiGrf, BulletOrigin)> {
        self.flags
    }

    /// Whether the winning fPicBullet is an enabled style-placed value.
    pub(in crate::doc) fn style_enabled(&self) -> bool {
        matches!(self.flags, Some((flags, BulletOrigin::Style(_))) if flags.picture)
    }

    /// The enabled picture bullet, if any. [MS-DOC] 2.6.1: when a picture
    /// bullet is used, sprmCPbiIBullet MUST locate it; an enabled flag
    /// without a position is malformed rather than a text-bullet fallback.
    pub(in crate::doc) fn enabled(&self) -> Result<Option<EnabledPictureBullet>, String> {
        let Some((flags, flags_origin)) = self.flags.filter(|(flags, _)| flags.picture) else {
            return Ok(None);
        };
        let (relative_cp, index_origin) = self
            .index
            .ok_or_else(|| unsupported("Word picture bullet lacks its picture position"))?;
        if matches!(flags_origin, BulletOrigin::Style(_))
            || matches!(index_origin, BulletOrigin::Style(_))
        {
            return Err(unsupported("Word picture bullet is placed in a style"));
        }
        Ok(Some(EnabledPictureBullet { relative_cp, flags }))
    }
}
