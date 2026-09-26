//! Local passive picture references, not complete drawing visibility/layout.
//! MS-XLS 2.4.181, 2.5.142/151; MS-ODRAW 2.3.23.1-9/2.3.18.5.
use super::*;
use crate::officeart::{properties, Record as ArtRecord};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PictureReference {
    /// One-based global BStore index; zero and complex BLIPs are not projected.
    pub store_index: u32,
    /// Raw signed 16.16 fractions: top, bottom, left, right.
    pub crop: [i32; 4],
    /// Raw signed 16.16 degrees. No inferred geometry transform.
    pub rotation: i32,
    pub clipboard_format: u16,
    /// FtPioGrbit fAutoPict, retained for later view/aspect handling.
    pub auto_picture: bool,
}

#[derive(Default)]
pub(super) struct Properties {
    seen: bool,
    index: Option<u32>,
    crop: [i32; 4],
    rotation: i32,
    excluded: bool,
}

impl Properties {
    pub(super) fn excluded(&self) -> bool {
        // A rotated parent requires group-transform composition, which the
        // current ordinary-picture projection deliberately does not perform.
        self.excluded || self.rotation != 0
    }
    pub(super) fn read(&mut self, record: ArtRecord<'_>, work: &mut usize) -> Result<(), String> {
        if self.seen {
            return Err(unsupported("duplicate BIFF picture property table"));
        }
        self.seen = true;
        let mut seen = HashSet::new();
        properties::visit(record, work, |p| {
            let id = p.opid & 0x3fff;
            if matches!(id, 4 | 0x100..=0x104 | 0x106 | 0x3bf) && !seen.insert(id) {
                return Err(unsupported("duplicate BIFF picture property"));
            }
            match id {
                0x104 => {
                    // fComplex means an inline BLIP, NOT a global store index.
                    if p.complex.is_some() {
                        self.excluded = true;
                    } else if p.opid != 0x4104 {
                        return Err(unsupported("invalid BIFF picture index flags"));
                    } else {
                        self.index = (p.value != 0).then_some(p.value);
                    }
                }
                4 | 0x100..=0x103 | 0x106 | 0x3bf => {
                    if p.opid != id {
                        return Err(unsupported("invalid BIFF picture scalar flags"));
                    }
                    match id {
                        4 => self.rotation = p.value as i32,
                        0x100..=0x103 => self.crop[(id - 0x100) as usize] = p.value as i32,
                        // MS-ODRAW 2.4.8: admit the embedded/comment form only;
                        // do not follow filenames, URLs or linked image updates.
                        0x106 => self.excluded |= p.value != 0,
                        0x3bf => {
                            // Hidden / script-anchor values apply only with
                            // their corresponding fUse bits (2.3.4.44).
                            self.excluded |= [1, 7].into_iter().any(|bit| {
                                p.value & (1 << (bit + 16)) != 0 && p.value & (1 << bit) != 0
                            });
                        }
                        _ => unreachable!(),
                    }
                }
                _ => {} // No text, script, link or nested object is interpreted.
            }
            Ok(())
        })
    }

    pub(super) fn reference(
        &self,
        shape_flags: u32,
        object: Option<&[u8]>,
    ) -> Result<Option<PictureReference>, String> {
        // Group, patriarch, deleted, OLE and background shapes are not ordinary
        // local picture frames. Ancestor visibility/layout is a later stage.
        if self.excluded || shape_flags & (1 | 4 | 8 | 16 | 1024) != 0 {
            return Ok(None);
        }
        let (Some(index), Some(object)) = (self.index, object) else {
            return Ok(None);
        };
        if !passive_object(object)? {
            return Ok(None);
        }
        Ok(Some(PictureReference {
            store_index: index,
            crop: self.crop,
            rotation: self.rotation,
            clipboard_format: u16_at(object, 26)?,
            auto_picture: u16_at(object, 32)? & 1 != 0,
        }))
    }
}

fn passive_object(data: &[u8]) -> Result<bool, String> {
    // The owning reader has already validated FtCmo. Only picture Obj forms
    // carry the two required fields below, in this exact order.
    if u16_at(data, 4)? != 8 {
        return Ok(false);
    }
    if data.len() < 38
        || u16_at(data, 22)? != 7
        || u16_at(data, 24)? != 2
        || u16_at(data, 28)? != 8
        || u16_at(data, 30)? != 2
    {
        return Err(unsupported("invalid BIFF picture object fields"));
    }
    if !matches!(u16_at(data, 26)?, 2 | 9 | 0xffff) {
        return Err(unsupported("invalid BIFF picture clipboard format"));
    }
    let flags = u16_at(data, 32)?;
    // fDde and fCtl are mutually exclusive, even in an unsupported object.
    if flags & 0x12 == 0x12 {
        return Err(unsupported("invalid BIFF DDE/control picture flags"));
    }
    // Explicit-sized plain pictures only: DDE, print recalculation, icon,
    // ActiveX, controls-stream, camera, default-size and auto-load are excluded.
    // Undefined bits (6, 10..15) are ignored as specified, not rejection flags.
    if flags & 0x03be != 0 || u16_at(data, 8)? & 4 != 0 {
        return Ok(false);
    }
    // Extra FtMacro/FtPictFmla fields are intentionally not evaluated/copied.
    // The ordinary picture form ends with four reserved, ignored bytes.
    Ok(data.len() == 38)
}

#[cfg(test)]
mod tests;
