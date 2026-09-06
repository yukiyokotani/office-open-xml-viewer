//! Native preparation evidence, not drawing emission. MS-XLS 2.4.170/181,
//! 2.5.143/193-195; MS-ODRAW 2.2.13-17/40. Keep raw cell fractions intact.
use super::{
    parse_bound_sheet, u16_at, u32_at, unsupported, Record, BIFF8, BOF, BOUNDSHEET8, EOF, FILEPASS,
    WORKBOOK_GLOBALS, WORKSHEET,
};
use crate::officeart::record_with_end;
use std::collections::{BTreeMap, HashSet};

const MAX_BYTES: usize = 128 * 1024 * 1024;
const MAX_OBJECTS: usize = 65_536;
const MAX_DEPTH: usize = 32;

mod picture;
pub use picture::PictureReference;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CellCorner {
    pub column: u16,
    pub row: u16,
    /// Signed 1/1024 column-width units, not pixels or EMUs.
    pub dx: i16,
    /// Signed 1/256 row-height units, not pixels or EMUs.
    pub dy: i16,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DrawingAnchor {
    /// Zero-based BoundSheet tab order, including non-worksheet tabs.
    pub sheet: usize,
    pub shape_id: u32,
    pub shape_flags: u32,
    pub object_id: u16,
    pub object_type: u16,
    pub object_flags: u16,
    /// Number of enclosing Spgr containers; includes the patriarch group.
    pub group_depth: usize,
    /// Raw fMove/fSize bits only. No inferred DrawingML editAs mapping.
    pub behavior: u16,
    pub from: CellCorner,
    pub to: CellCorner,
    pub picture: Option<PictureReference>,
}

pub(super) fn workbook(records: &[Record<'_>]) -> Result<Vec<DrawingAnchor>, String> {
    let first = records
        .first()
        .ok_or_else(|| unsupported("empty BIFF workbook"))?;
    if first.kind != BOF
        || u16_at(first.data, 0)? != BIFF8
        || u16_at(first.data, 2)? != WORKBOOK_GLOBALS
    {
        return Err(unsupported("anchor inspection requires BIFF8 globals"));
    }
    if records.iter().any(|r| r.kind == FILEPASS) {
        return Err(unsupported("encrypted BIFF drawing anchors"));
    }
    let globals_end = records
        .iter()
        .position(|r| r.kind == EOF)
        .ok_or_else(|| unsupported("missing BIFF global EOF"))?;
    let mut starts = BTreeMap::new();
    for record in &records[..globals_end] {
        if record.kind == BOUNDSHEET8 {
            let sheet = parse_bound_sheet(record.data)?;
            let index = records
                .binary_search_by_key(&sheet.offset, |r| r.offset)
                .map_err(|_| unsupported("BIFF sheet offset is not a record boundary"))?;
            if index <= globals_end || starts.len() >= MAX_OBJECTS {
                return Err(unsupported("invalid BIFF drawing sheet range"));
            }
            let tab = starts.len();
            if starts.insert(index, (tab, sheet.sheet_type)).is_some() {
                return Err(unsupported("duplicate BIFF drawing sheet offset"));
            }
        }
    }
    let starts: Vec<_> = starts.into_iter().collect();
    let mut work = 2_000_000;
    let mut remaining = MAX_BYTES;
    let mut output = Vec::new();
    // Disjoint physical ranges prevent overlapping BoundSheet pointers from
    // making us repeatedly scan the same worksheet or its embedded charts.
    for (ordinal, &(start, (tab, kind))) in starts.iter().enumerate() {
        if kind != 0 {
            continue;
        }
        let end = starts.get(ordinal + 1).map_or(records.len(), |s| s.0);
        let data = assemble(&records[start..end], &mut work, &mut remaining)?;
        if let Some(mut drawing) = data {
            walk(&mut drawing, tab, &mut work, &mut output)?;
        }
    }
    Ok(output)
}

struct Drawing<'a> {
    bytes: Vec<u8>,
    /// Assembled byte boundary -> the immediately following native client.
    clients: BTreeMap<usize, Record<'a>>,
}

fn spend(work: &mut usize) -> Result<(), String> {
    *work = work
        .checked_sub(1)
        .ok_or_else(|| unsupported("BIFF drawing work budget exceeded"))?;
    Ok(())
}

fn assemble<'a>(
    records: &[Record<'a>],
    work: &mut usize,
    remaining: &mut usize,
) -> Result<Option<Drawing<'a>>, String> {
    let first = records
        .first()
        .ok_or_else(|| unsupported("empty BIFF drawing sheet"))?;
    if first.kind != BOF || u16_at(first.data, 0)? != BIFF8 || u16_at(first.data, 2)? != WORKSHEET {
        return Err(unsupported("invalid BIFF drawing worksheet BOF"));
    }
    let (mut depth, mut active, mut length, mut complete) = (0usize, false, 0usize, false);
    let mut fragments = Vec::new();
    let mut clients = BTreeMap::new();
    for &record in &records[1..] {
        spend(work)?;
        if record.kind == BOF {
            depth += 1;
            if depth > MAX_DEPTH {
                return Err(unsupported("BIFF substream depth exceeded"));
            }
            active = false;
            continue;
        }
        if record.kind == EOF {
            if depth == 0 {
                complete = true;
                break;
            }
            depth -= 1;
            active = false;
            continue;
        }
        if depth != 0 {
            continue;
        }
        if record.kind == 0x00ec || (active && record.kind == 0x003c) {
            if record.data.len() > 8224 {
                return Err(unsupported("oversized BIFF sheet drawing fragment"));
            }
            length = length
                .checked_add(record.data.len())
                .filter(|n| *n <= *remaining)
                .ok_or_else(|| unsupported("BIFF sheet drawing byte budget exceeded"))?;
            fragments.push(record.data);
            active = true;
        } else {
            if active && matches!(record.kind, 0x005d | 0x01b6) {
                if clients.len() >= MAX_OBJECTS || clients.insert(length, record).is_some() {
                    return Err(unsupported("ambiguous or excessive BIFF drawing clients"));
                }
            }
            // This subset assigns post-Obj/TxO continuations to their native
            // client. Reclassifying interleaved producer continuations needs
            // complete client-length ownership, not an OfficeArt header sniff.
            active = false;
        }
    }
    if !complete {
        return Err(unsupported("missing BIFF drawing worksheet EOF"));
    }
    if fragments.is_empty() {
        return Ok(None);
    }
    *remaining -= length;
    let mut bytes = Vec::with_capacity(length);
    for fragment in fragments {
        bytes.extend_from_slice(fragment);
    }
    Ok(Some(Drawing { bytes, clients }))
}

fn corner(bytes: &[u8], offset: usize) -> Result<CellCorner, String> {
    let column = u16_at(bytes, offset)?;
    if column > 256 {
        return Err(unsupported("invalid BIFF anchor column"));
    }
    Ok(CellCorner {
        column,
        dx: u16_at(bytes, offset + 2)? as i16,
        row: u16_at(bytes, offset + 4)?,
        dy: u16_at(bytes, offset + 6)? as i16,
    })
}

fn walk(
    drawing: &mut Drawing<'_>,
    sheet: usize,
    work: &mut usize,
    output: &mut Vec<DrawingAnchor>,
) -> Result<(), String> {
    let (root, end) = record_with_end(&drawing.bytes, 0, work, "XLS drawing")?;
    if root.kind != 0xf002 || root.version != 15 || root.instance != 0 || end != drawing.bytes.len()
    {
        return Err(unsupported("invalid BIFF sheet drawing root"));
    }
    let mut stack = vec![(8usize, end, 0usize, false, true)];
    let mut ids = HashSet::new();
    let mut objects = HashSet::new();
    while let Some((mut at, end, depth, group, mut first)) = stack.pop() {
        if depth > MAX_DEPTH {
            return Err(unsupported("BIFF drawing group depth exceeded"));
        }
        while at < end {
            let (record, next) = record_with_end(&drawing.bytes[..end], at, work, "XLS drawing")?;
            if group
                && ((first && record.kind != 0xf004) || !matches!(record.kind, 0xf003 | 0xf004))
            {
                return Err(unsupported("invalid BIFF drawing group child"));
            }
            first = false;
            if matches!(record.kind, 0xf003 | 0xf004)
                && (record.version != 15 || record.instance != 0)
            {
                return Err(unsupported("invalid BIFF shape container"));
            }
            if record.kind == 0xf003 {
                // Resume siblings after the owned group; bounded stack, no
                // recursive descent into arbitrary application-specific data.
                stack.push((next, end, depth, group, false));
                stack.push((at + 8, next, depth + 1, true, true));
                break;
            }
            if record.kind == 0xf004 {
                let (mut position, mut shape, mut anchor, mut object, mut textbox) =
                    (at + 8, None, None, None, false);
                let mut object_data = None;
                let mut picture = picture::Properties::default();
                while position < next {
                    let (child, child_end) =
                        record_with_end(&drawing.bytes[..next], position, work, "XLS shape")?;
                    match child.kind {
                        0xf00b => picture.read(child, work)?,
                        0xf00a => {
                            if shape.is_some() || child.version != 2 || child.payload.len() != 8 {
                                return Err(unsupported("invalid BIFF shape identity"));
                            }
                            let id = u32_at(child.payload, 0)?;
                            if ids.len() >= MAX_OBJECTS || !ids.insert(id) {
                                return Err(unsupported("duplicate or excessive BIFF shapes"));
                            }
                            shape = Some((id, u32_at(child.payload, 4)?));
                        }
                        0xf010 => {
                            if anchor.is_some()
                                || child.version != 0
                                || child.instance != 0
                                || child.payload.len() != 18
                            {
                                return Err(unsupported("invalid BIFF cell anchor"));
                            }
                            let flags = u16_at(child.payload, 0)? & 3;
                            if flags == 1 {
                                return Err(unsupported("invalid BIFF anchor movement flags"));
                            }
                            anchor = Some((
                                flags,
                                corner(child.payload, 2)?,
                                corner(child.payload, 10)?,
                            ));
                        }
                        0xf011 | 0xf00d => {
                            if child.version != 0
                                || child.instance != 0
                                || !child.payload.is_empty()
                            {
                                return Err(unsupported("invalid BIFF drawing client marker"));
                            }
                            let client = drawing.clients.remove(&child_end).ok_or_else(|| {
                                unsupported("BIFF drawing client is not at its fragment boundary")
                            })?;
                            if child.kind == 0xf011 {
                                if object.is_some()
                                    || client.kind != 0x005d
                                    || client.data.len() < 22
                                    || u16_at(client.data, 0)? != 0x15
                                    || u16_at(client.data, 2)? != 0x12
                                {
                                    return Err(unsupported("invalid BIFF drawing object owner"));
                                }
                                let kind = u16_at(client.data, 4)?;
                                if !matches!(kind, 0..=9 | 11..=20 | 25 | 30) {
                                    return Err(unsupported("invalid BIFF object type"));
                                }
                                let id = u16_at(client.data, 6)?;
                                if !objects.insert(id) {
                                    return Err(unsupported("duplicate BIFF drawing object id"));
                                }
                                object = Some((id, kind, u16_at(client.data, 8)?));
                                object_data = Some(client.data);
                            } else {
                                if textbox || client.kind != 0x01b6 {
                                    return Err(unsupported("invalid BIFF drawing textbox owner"));
                                }
                                textbox = true;
                            }
                        }
                        _ => {} // No formula, action, OLE or hyperlink decoding.
                    }
                    position = child_end;
                }
                let (shape_id, shape_flags) =
                    shape.ok_or_else(|| unsupported("missing BIFF shape identity"))?;
                if let Some((behavior, from, to)) = anchor {
                    let (object_id, object_type, object_flags) = object
                        .ok_or_else(|| unsupported("BIFF cell anchor has no owned object"))?;
                    if output.len() >= MAX_OBJECTS {
                        return Err(unsupported("BIFF retained anchor budget exceeded"));
                    }
                    output.push(DrawingAnchor {
                        sheet,
                        shape_id,
                        shape_flags,
                        object_id,
                        object_type,
                        object_flags,
                        group_depth: depth,
                        behavior,
                        from,
                        to,
                        picture: picture.reference(shape_flags, object_data)?,
                    });
                }
            }
            at = next;
        }
        if group && first {
            return Err(unsupported("empty BIFF drawing group"));
        }
    }
    if !drawing.clients.is_empty() {
        return Err(unsupported("unowned BIFF drawing clients"));
    }
    Ok(())
}

#[cfg(test)]
mod tests;
