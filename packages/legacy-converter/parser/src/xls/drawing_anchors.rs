//! Owned sheet anchor preparation, not drawing emission. MS-XLS 2.4.170/181,
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

#[derive(Debug, Clone, PartialEq)]
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
    /// Absolute record range (BOF..=EOF) of the chart substream that follows
    /// this anchor's Obj record (MS-XLS 2.1.7.20.6 `OBJ = Obj *Continue *CHART`).
    pub chart: Option<(usize, usize)>,
    /// One-based position of this shape container among all shape
    /// containers of the sheet drawing, in document (paint) order.
    pub order: u64,
    /// OfficeArt facts of a rectangle or text box (Obj ot 2 or 6), whose
    /// interpretation belongs to the projecting host.
    pub shape: Option<ShapeSource>,
    /// Members of a group placed on the sheet (Obj ot 0), in document order.
    /// Only the strict walk flattens groups.
    pub members: Vec<GroupMember>,
}

/// An OfficeArt shape's own facts: its MS-ODRAW 2.4.24 shape type (the FSP
/// instance), its primary then tertiary FOPT entries (full opid with fBid and
/// fComplex, and value) and the stream offset of its TxO client (MS-XLS
/// 2.4.329), if any.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ShapeSource {
    pub kind: u16,
    pub properties: Vec<(u16, u32)>,
    /// Complex data of the geometry arrays (MS-ODRAW 2.3.6: pVertices,
    /// pSegmentInfo, pConnectionSites and their directions, pAdjustHandles,
    /// pGuides and pInscribe), by property id.
    pub complex: Vec<(u16, Vec<u8>)>,
    pub text: Option<usize>,
}

/// One enclosing group of a group member (MS-ODRAW 2.2.38 OfficeArtFSPGR,
/// 2.2.39 child anchor, 2.3.18.5 rotation, 2.2.40 flips), without any
/// interpretation of how Office composes them.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GroupFrame {
    /// Left, top, right, bottom in the parent group's coordinates; `None`
    /// for the sheet-anchored group, which its cell anchor places.
    pub anchor: Option<[i32; 4]>,
    /// The group's own coordinate rectangle: left, top, right, bottom.
    pub rect: [i32; 4],
    /// Raw signed 16.16 degrees.
    pub rotation: i32,
    pub flip_h: bool,
    pub flip_v: bool,
}

/// A leaf of a sheet-anchored group: its child anchor in the innermost
/// group's coordinates and its enclosing groups, outermost first.
#[derive(Debug, Clone, PartialEq)]
pub struct GroupMember {
    pub order: u64,
    pub shape_id: u32,
    pub shape_flags: u32,
    pub object_id: u16,
    pub object_type: u16,
    pub anchor: [i32; 4],
    pub groups: Vec<GroupFrame>,
    pub shape: Option<ShapeSource>,
    pub picture: Option<PictureReference>,
}

/// Which anchors a walk returns. `Projectable` skips shapes whose sheet
/// placement this subset does not flatten (nested groups, child anchors,
/// excluded patriarchs); `Strict` rejects them instead, for readers that must
/// not drop drawn content.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Policy {
    #[cfg(any(test, all(feature = "inspection", not(target_arch = "wasm32"))))]
    All,
    Projectable,
    Strict,
}

#[cfg(any(test, all(feature = "inspection", not(target_arch = "wasm32"))))]
pub(super) fn workbook(records: &[Record<'_>]) -> Result<Vec<DrawingAnchor>, String> {
    workbook_with_policy(records, Policy::All)
}

/// Every sheet-anchored drawing object, rejecting any drawn shape whose
/// placement is not projected (grouped shapes and their children).
pub(super) fn strict(records: &[Record<'_>]) -> Result<Vec<DrawingAnchor>, String> {
    workbook_with_policy(records, Policy::Strict)
}

pub(super) fn projectable(records: &[Record<'_>]) -> Result<Vec<DrawingAnchor>, String> {
    workbook_with_policy(records, Policy::Projectable)
}

fn workbook_with_policy(
    records: &[Record<'_>],
    policy: Policy,
) -> Result<Vec<DrawingAnchor>, String> {
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
        let data = assemble(&records[start..end], start, &mut work, &mut remaining)?;
        if let Some(mut drawing) = data {
            walk_with_policy(&mut drawing, tab, &mut work, &mut output, policy)?;
        }
    }
    Ok(output)
}

struct Drawing<'a> {
    bytes: Vec<u8>,
    /// Assembled byte boundary -> the immediately following native client.
    clients: BTreeMap<usize, Record<'a>>,
    /// Client boundary -> absolute chart substream record range.
    charts: BTreeMap<usize, (usize, usize)>,
}

fn spend(work: &mut usize) -> Result<(), String> {
    *work = work
        .checked_sub(1)
        .ok_or_else(|| unsupported("BIFF drawing work budget exceeded"))?;
    Ok(())
}

fn assemble<'a>(
    records: &[Record<'a>],
    base: usize,
    work: &mut usize,
    remaining: &mut usize,
) -> Result<Option<Drawing<'a>>, String> {
    let first = records
        .first()
        .ok_or_else(|| unsupported("empty BIFF drawing sheet"))?;
    if first.kind != BOF || u16_at(first.data, 0)? != BIFF8 || u16_at(first.data, 2)? != WORKSHEET {
        return Err(unsupported("invalid BIFF drawing worksheet BOF"));
    }
    let (mut depth, mut length, mut complete) = (0usize, 0usize, false);
    let mut owner = Owner::None;
    let mut fragments = Vec::new();
    let mut clients = BTreeMap::new();
    let mut charts = BTreeMap::new();
    // The Obj client that may own an immediately following chart substream.
    let mut last_object: Option<usize> = None;
    let mut chart_start: Option<(usize, usize)> = None;
    for (offset, &record) in records.iter().enumerate().skip(1) {
        spend(work)?;
        if record.kind == BOF {
            depth += 1;
            if depth > MAX_DEPTH {
                return Err(unsupported("BIFF substream depth exceeded"));
            }
            if depth == 1 && u16_at(record.data, 2)? == 0x0020 {
                chart_start = last_object.map(|client| (client, base + offset));
            }
            owner = Owner::None;
            continue;
        }
        if record.kind == EOF {
            if depth == 0 {
                complete = true;
                break;
            }
            depth -= 1;
            owner = Owner::None;
            if depth == 0 {
                if let Some((client, start)) = chart_start.take() {
                    charts.insert(client, (start, base + offset));
                    // OBJ = Obj *Continue *CHART is complete.
                    owner = Owner::Complete;
                }
                last_object = None;
            }
            continue;
        }
        if depth != 0 {
            continue;
        }
        let drawing = match (record.kind, &mut owner) {
            (0x00ec, _) | (0x003c, Owner::Drawing | Owner::Complete) => true,
            (0x003c, Owner::Text { characters, runs }) => {
                // TEXTOBJECT = TxO *Continue: the text fragments, then the
                // formatting runs (MS-XLS 2.4.329).
                if *characters > 0 {
                    let (flag, chars) = record
                        .data
                        .split_first()
                        .ok_or_else(|| unsupported("empty BIFF TxO text fragment"))?;
                    let count = if *flag & 1 != 0 {
                        chars.len() / 2
                    } else {
                        chars.len()
                    };
                    *characters = characters.saturating_sub(count);
                } else {
                    *runs = runs.saturating_sub(record.data.len());
                }
                if *characters == 0 && *runs == 0 {
                    owner = Owner::Complete;
                }
                false
            }
            _ => false,
        };
        if drawing {
            if record.data.len() > 8224 {
                return Err(unsupported("oversized BIFF sheet drawing fragment"));
            }
            length = length
                .checked_add(record.data.len())
                .filter(|n| *n <= *remaining)
                .ok_or_else(|| unsupported("BIFF sheet drawing byte budget exceeded"))?;
            fragments.push(record.data);
            owner = Owner::Drawing;
            continue;
        }
        if record.kind == 0x003c {
            // A continuation its native owner still claims (or an unowned
            // one) is never drawing data.
            continue;
        }
        if owner == Owner::Drawing
            && matches!(record.kind, 0x005d | 0x01b6)
            && (clients.len() >= MAX_OBJECTS || clients.insert(length, record).is_some())
        {
            return Err(unsupported("ambiguous or excessive BIFF drawing clients"));
        }
        last_object = (record.kind == 0x005d).then_some(length);
        // Excel continues the sheet's OfficeArt stream in Continue records
        // after an Obj or TxO once that native record is complete, which the
        // OBJECTS grammar (MS-XLS 2.1.7.20.5) would assign to the Obj/TxO.
        // Ownership follows the native record's own length: a complete Obj
        // (its subrecords end with FtEnd at the record end) or a TxO whose
        // text and runs are consumed owns no further Continue. Any Obj whose
        // length cannot be established keeps its continuations.
        owner = match record.kind {
            0x005d if object_complete(record.data) => Owner::Complete,
            0x01b6 if record.data.len() >= 14 => {
                let characters = usize::from(u16_at(record.data, 10)?);
                let runs = usize::from(u16_at(record.data, 12)?);
                if characters == 0 && runs == 0 {
                    Owner::Complete
                } else {
                    Owner::Text { characters, runs }
                }
            }
            _ => Owner::None,
        };
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
    Ok(Some(Drawing {
        bytes,
        clients,
        charts,
    }))
}

/// Who owns a Continue record at this point of the worksheet substream.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Owner {
    None,
    /// The sheet's OfficeArt stream (MsoDrawing and its continuations).
    Drawing,
    /// A complete Obj or TxO: a following Continue resumes the drawing.
    Complete,
    /// A TxO still owning this many characters and formatting-run bytes.
    Text {
        characters: usize,
        runs: usize,
    },
}

/// Whether an Obj record's subrecords (MS-XLS 2.4.181) end with FtEnd exactly
/// at the record end. FtLbsData (ft 0x0013) does not state its own length,
/// so an Obj containing it is never taken as complete.
fn object_complete(data: &[u8]) -> bool {
    let mut at = 0usize;
    while at + 4 <= data.len() {
        let kind = u16::from_le_bytes([data[at], data[at + 1]]);
        let size = usize::from(u16::from_le_bytes([data[at + 2], data[at + 3]]));
        match kind {
            0x0000 => return size == 0 && at + 4 == data.len(),
            0x0013 => return false,
            _ => at += 4 + size,
        }
    }
    false
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

#[cfg(test)]
fn walk(
    drawing: &mut Drawing<'_>,
    sheet: usize,
    work: &mut usize,
    output: &mut Vec<DrawingAnchor>,
) -> Result<(), String> {
    walk_with_policy(drawing, sheet, work, output, Policy::All)
}

/// One OfficeArtSpContainer's owned facts (MS-ODRAW 2.2.14), with its native
/// Obj/TxO clients bound by fragment boundary.
struct Facts<'a> {
    shape_id: u32,
    shape_flags: u32,
    kind: u16,
    anchor: Option<(u16, CellCorner, CellCorner)>,
    /// OfficeArtChildAnchor (2.2.39): left, top, right, bottom.
    child_anchor: Option<[i32; 4]>,
    /// OfficeArtFSPGR (2.2.38): the group's own coordinate rectangle.
    group_rect: Option<[i32; 4]>,
    object: Option<(u16, u16, u16)>,
    object_data: Option<&'a [u8]>,
    chart: Option<(usize, usize)>,
    text: Option<usize>,
    picture: picture::Properties,
    properties: Vec<(u16, u32)>,
    complex: Vec<(u16, Vec<u8>)>,
}

impl Facts<'_> {
    fn source(&self) -> ShapeSource {
        ShapeSource {
            kind: self.kind,
            properties: self.properties.clone(),
            complex: self.complex.clone(),
            text: self.text,
        }
    }

    fn rotation(&self) -> u32 {
        self.properties
            .iter()
            .find(|(id, _)| *id == 0x0004)
            .map_or(0, |(_, value)| *value)
    }

    /// This group's frame; `anchor` is its child anchor, if nested.
    fn frame(&self, anchor: Option<[i32; 4]>) -> Result<GroupFrame, String> {
        let rect = self
            .group_rect
            .filter(|[left, top, right, bottom]| right > left && bottom > top)
            .ok_or_else(|| unsupported("empty BIFF drawing group rectangle"))?;
        Ok(GroupFrame {
            anchor,
            rect,
            rotation: self.rotation() as i32,
            flip_h: self.shape_flags & 0x40 != 0,
            flip_v: self.shape_flags & 0x80 != 0,
        })
    }
}

/// Geometry arrays retained for hosts that draw OfficeArt freeforms.
const GEOMETRY_ARRAYS: [u16; 7] = [0x145, 0x146, 0x151, 0x152, 0x155, 0x156, 0x157];
const MAX_SHAPE_COMPLEX_BYTES: usize = 4 * 1024 * 1024;

fn rectangle(payload: &[u8]) -> Result<[i32; 4], String> {
    if payload.len() != 16 {
        return Err(unsupported("invalid BIFF drawing rectangle"));
    }
    Ok(std::array::from_fn(|i| {
        i32::from_le_bytes(payload[i * 4..i * 4 + 4].try_into().unwrap())
    }))
}

fn read_shape<'a>(
    drawing: &mut Drawing<'a>,
    start: usize,
    next: usize,
    work: &mut usize,
    ids: &mut HashSet<u32>,
    objects: &mut HashSet<u16>,
) -> Result<Facts<'a>, String> {
    let mut position = start;
    let mut shape = None;
    let mut facts = Facts {
        shape_id: 0,
        shape_flags: 0,
        kind: 0,
        anchor: None,
        child_anchor: None,
        group_rect: None,
        object: None,
        object_data: None,
        chart: None,
        text: None,
        picture: picture::Properties::default(),
        properties: Vec::new(),
        complex: Vec::new(),
    };
    let mut complex_bytes = 0usize;
    let mut textbox = false;
    while position < next {
        let (child, child_end) =
            record_with_end(&drawing.bytes[..next], position, work, "XLS shape")?;
        match child.kind {
            0xf00f => {
                if facts.child_anchor.is_some() || child.version != 0 || child.instance != 0 {
                    return Err(unsupported("invalid BIFF child anchor"));
                }
                facts.child_anchor = Some(rectangle(child.payload)?);
            }
            0xf009 => {
                if facts.group_rect.is_some() || child.version != 1 || child.instance != 0 {
                    return Err(unsupported("invalid BIFF group rectangle"));
                }
                facts.group_rect = Some(rectangle(child.payload)?);
            }
            0xf00b | 0xf122 => {
                if child.kind == 0xf00b {
                    facts.picture.read(child, work)?;
                }
                let retain = |p: crate::officeart::properties::Property<'_>| {
                    if facts.properties.len() >= 4096 {
                        return Err(unsupported("excessive BIFF shape properties"));
                    }
                    facts.properties.push((p.opid, p.value));
                    if let Some(data) = p.complex {
                        if GEOMETRY_ARRAYS.contains(&(p.opid & 0x3fff)) {
                            complex_bytes += data.len();
                            if complex_bytes > MAX_SHAPE_COMPLEX_BYTES {
                                return Err(unsupported("oversized BIFF shape geometry"));
                            }
                            facts.complex.push((p.opid & 0x3fff, data.to_vec()));
                        }
                    }
                    Ok(())
                };
                if child.kind == 0xf00b {
                    crate::officeart::properties::visit(child, work, retain)?;
                } else {
                    crate::officeart::properties::visit_tertiary(child, work, retain)?;
                }
            }
            0xf00a => {
                if shape.is_some() || child.version != 2 || child.payload.len() != 8 {
                    return Err(unsupported("invalid BIFF shape identity"));
                }
                let id = u32_at(child.payload, 0)?;
                if ids.len() >= MAX_OBJECTS || !ids.insert(id) {
                    return Err(unsupported("duplicate or excessive BIFF shapes"));
                }
                shape = Some((id, u32_at(child.payload, 4)?, child.instance));
            }
            0xf010 => {
                if facts.anchor.is_some()
                    || child.version != 0
                    || child.instance != 0
                    || child.payload.len() != 18
                {
                    return Err(unsupported("invalid BIFF cell anchor"));
                }
                // fMove without fSize is invalid (MS-XLS 2.5.193); the owner
                // check below admits it only for AutoFilter drop-downs.
                let flags = u16_at(child.payload, 0)? & 3;
                facts.anchor = Some((flags, corner(child.payload, 2)?, corner(child.payload, 10)?));
            }
            0xf011 | 0xf00d => {
                if child.version != 0 || child.instance != 0 || !child.payload.is_empty() {
                    return Err(unsupported("invalid BIFF drawing client marker"));
                }
                let client = drawing.clients.remove(&child_end).ok_or_else(|| {
                    unsupported("BIFF drawing client is not at its fragment boundary")
                })?;
                if child.kind == 0xf011 {
                    facts.chart = drawing.charts.remove(&child_end);
                    if facts.object.is_some()
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
                    facts.object = Some((id, kind, u16_at(client.data, 8)?));
                    facts.object_data = Some(client.data);
                } else {
                    if textbox || client.kind != 0x01b6 {
                        return Err(unsupported("invalid BIFF drawing textbox owner"));
                    }
                    textbox = true;
                    facts.text = Some(client.offset);
                }
            }
            _ => {} // No formula, action, OLE or hyperlink decoding.
        }
        position = child_end;
    }
    let (shape_id, shape_flags, kind) =
        shape.ok_or_else(|| unsupported("missing BIFF shape identity"))?;
    facts.shape_id = shape_id;
    facts.shape_flags = shape_flags;
    facts.kind = kind;
    Ok(facts)
}

/// Members of one group container, with their enclosing group frames
/// (outermost first). Interpretation of rotation, flips and scaling belongs
/// to the host.
#[allow(clippy::too_many_arguments)]
fn read_group_members<'a>(
    drawing: &mut Drawing<'a>,
    start: usize,
    end: usize,
    groups: &[GroupFrame],
    work: &mut usize,
    ids: &mut HashSet<u32>,
    objects: &mut HashSet<u16>,
    order: &mut u64,
    members: &mut Vec<GroupMember>,
) -> Result<(), String> {
    if groups.len() > MAX_DEPTH {
        return Err(unsupported("BIFF drawing group depth exceeded"));
    }
    let checked = |rect: [i32; 4]| -> Result<[i32; 4], String> {
        let [left, top, right, bottom] = rect;
        if right < left || bottom < top {
            return Err(unsupported("invalid BIFF child anchor"));
        }
        Ok(rect)
    };
    let mut at = start;
    while at < end {
        spend(work)?;
        let (record, next) = record_with_end(&drawing.bytes[..end], at, work, "XLS drawing")?;
        if !matches!(record.kind, 0xf003 | 0xf004) || record.version != 15 || record.instance != 0 {
            return Err(unsupported("invalid BIFF drawing group child"));
        }
        if record.kind == 0xf003 {
            // A nested group: its head's child anchor places it in this
            // group, and its own rectangle is its members' coordinates.
            let (head, head_end) =
                record_with_end(&drawing.bytes[..next], at + 8, work, "XLS drawing")?;
            if head.kind != 0xf004 || head.version != 15 || head.instance != 0 {
                return Err(unsupported("invalid BIFF drawing group child"));
            }
            let facts = read_shape(drawing, at + 16, head_end, work, ids, objects)?;
            let child = facts
                .child_anchor
                .ok_or_else(|| unsupported("BIFF nested drawing group without its child anchor"))?;
            if facts.shape_flags & 0x3 != 0x3
                || facts.anchor.is_some()
                || facts.object.is_some_and(|(_, kind, _)| kind != 0)
            {
                return Err(unsupported("invalid BIFF nested drawing group"));
            }
            let mut inner = groups.to_vec();
            inner.push(facts.frame(Some(checked(child)?))?);
            read_group_members(
                drawing, head_end, next, &inner, work, ids, objects, order, members,
            )?;
        } else {
            let facts = read_shape(drawing, at + 8, next, work, ids, objects)?;
            let child = facts
                .child_anchor
                .filter(|_| facts.shape_flags & 0x3 == 0x2 && facts.anchor.is_none())
                .ok_or_else(|| unsupported("BIFF group member without its child anchor"))?;
            let (object_id, object_type, _) = facts
                .object
                .ok_or_else(|| unsupported("BIFF group member has no owned object"))?;
            if members.len() >= MAX_OBJECTS {
                return Err(unsupported("BIFF retained anchor budget exceeded"));
            }
            *order += 1;
            members.push(GroupMember {
                order: *order,
                shape_id: facts.shape_id,
                shape_flags: facts.shape_flags,
                object_id,
                object_type,
                anchor: checked(child)?,
                groups: groups.to_vec(),
                shape: (object_type != 8).then(|| facts.source()),
                picture: facts
                    .picture
                    .reference(facts.shape_flags, facts.object_data)?,
            });
        }
        at = next;
    }
    Ok(())
}

/// A group placed directly on the sheet (MS-ODRAW 2.2.13 inside the
/// patriarch): its head carries the cell anchor and an Obj of type 0.
#[allow(clippy::too_many_arguments)]
fn read_sheet_group<'a>(
    drawing: &mut Drawing<'a>,
    at: usize,
    next: usize,
    sheet: usize,
    work: &mut usize,
    ids: &mut HashSet<u32>,
    objects: &mut HashSet<u16>,
    order: &mut u64,
) -> Result<DrawingAnchor, String> {
    let (head, head_end) = record_with_end(&drawing.bytes[..next], at + 8, work, "XLS drawing")?;
    if head.kind != 0xf004 || head.version != 15 || head.instance != 0 {
        return Err(unsupported("invalid BIFF drawing group child"));
    }
    let facts = read_shape(drawing, at + 16, head_end, work, ids, objects)?;
    let Some((behavior, from, to)) = facts.anchor else {
        return Err(unsupported("BIFF drawing group without its anchor"));
    };
    if behavior == 1 {
        return Err(unsupported("invalid BIFF anchor movement flags"));
    }
    let (object_id, object_type, object_flags) = facts
        .object
        .filter(|(_, kind, _)| *kind == 0)
        .ok_or_else(|| unsupported("BIFF drawing group without its group object"))?;
    if facts.shape_flags & 0x3 != 0x1 || facts.child_anchor.is_some() {
        return Err(unsupported("invalid BIFF drawing group"));
    }
    *order += 1;
    let group_order = *order;
    let mut members = Vec::new();
    read_group_members(
        drawing,
        head_end,
        next,
        &[facts.frame(None)?],
        work,
        ids,
        objects,
        order,
        &mut members,
    )?;
    if members.is_empty() {
        return Err(unsupported("empty BIFF drawing group"));
    }
    Ok(DrawingAnchor {
        sheet,
        shape_id: facts.shape_id,
        shape_flags: facts.shape_flags,
        object_id,
        object_type,
        object_flags,
        group_depth: 2,
        behavior,
        from,
        to,
        picture: None,
        chart: None,
        order: group_order,
        shape: Some(facts.source()),
        members,
    })
}

fn walk_with_policy(
    drawing: &mut Drawing<'_>,
    sheet: usize,
    work: &mut usize,
    output: &mut Vec<DrawingAnchor>,
    policy: Policy,
) -> Result<(), String> {
    let (root, end) = record_with_end(&drawing.bytes, 0, work, "XLS drawing")?;
    if root.kind != 0xf002 || root.version != 15 || root.instance != 0 || end != drawing.bytes.len()
    {
        return Err(unsupported("invalid BIFF sheet drawing root"));
    }
    let mut stack = vec![(8usize, end, 0usize, false, true, false)];
    let mut ids = HashSet::new();
    let mut objects = HashSet::new();
    // Paint order: every shape container in document order (MS-ODRAW 2.2.13).
    let mut order = 0u64;
    while let Some((mut at, end, depth, group, mut first, mut excluded)) = stack.pop() {
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
            let group_head = group && first;
            first = false;
            if matches!(record.kind, 0xf003 | 0xf004)
                && (record.version != 15 || record.instance != 0)
            {
                return Err(unsupported("invalid BIFF shape container"));
            }
            if record.kind == 0xf003 {
                if policy == Policy::Strict && depth == 1 {
                    // A group placed on the sheet: flatten its members into
                    // its anchor. Deeper nesting is composed inside.
                    if excluded {
                        return Err(unsupported(
                            "BIFF grouped or transformed drawing shapes are not projected",
                        ));
                    }
                    if output.len() >= MAX_OBJECTS {
                        return Err(unsupported("BIFF retained anchor budget exceeded"));
                    }
                    let anchor = read_sheet_group(
                        drawing,
                        at,
                        next,
                        sheet,
                        work,
                        &mut ids,
                        &mut objects,
                        &mut order,
                    )?;
                    output.push(anchor);
                    at = next;
                    continue;
                }
                // Resume siblings after the owned group; bounded stack, no
                // recursive descent into arbitrary application-specific data.
                stack.push((next, end, depth, group, false, excluded));
                stack.push((at + 8, next, depth + 1, true, true, excluded));
                break;
            }
            if record.kind == 0xf004 {
                let facts = read_shape(drawing, at + 8, next, work, &mut ids, &mut objects)?;
                order += 1;
                if group_head {
                    // Only the top-level patriarch is transparent to sheet
                    // coordinates. Nested group transforms are not flattened.
                    excluded |= depth != 1
                        || facts.shape_flags & 5 != 5
                        || facts.shape_flags & (8 | 16 | 64 | 128 | 1024) != 0
                        || facts.picture.excluded();
                }
                if policy == Policy::Strict
                    && facts.anchor.is_none()
                    && facts.child_anchor.is_some()
                {
                    return Err(unsupported("BIFF grouped drawing shapes are not projected"));
                }
                if let Some((behavior, from, to)) = facts.anchor {
                    let (object_id, object_type, object_flags) = facts
                        .object
                        .ok_or_else(|| unsupported("BIFF cell anchor has no owned object"))?;
                    // Excel anchors the application-inserted drop-down
                    // objects of AutoFilters (Obj ot 20 with fUIObj) with
                    // fMove but not fSize; any other such anchor is invalid.
                    if behavior == 1 && !(object_type == 20 && object_flags & 0x100 != 0) {
                        return Err(unsupported("invalid BIFF anchor movement flags"));
                    }
                    if output.len() >= MAX_OBJECTS {
                        return Err(unsupported("BIFF retained anchor budget exceeded"));
                    }
                    if matches!(policy, Policy::Projectable | Policy::Strict)
                        && (excluded || depth > 1 || facts.child_anchor.is_some())
                    {
                        if policy == Policy::Strict {
                            return Err(unsupported(
                                "BIFF grouped or transformed drawing shapes are not projected",
                            ));
                        }
                        at = next;
                        continue;
                    }
                    output.push(DrawingAnchor {
                        sheet,
                        shape_id: facts.shape_id,
                        shape_flags: facts.shape_flags,
                        object_id,
                        object_type,
                        object_flags,
                        group_depth: depth,
                        behavior,
                        from,
                        to,
                        picture: facts
                            .picture
                            .reference(facts.shape_flags, facts.object_data)?,
                        chart: facts.chart.filter(|_| object_type == 5),
                        order,
                        shape: matches!(object_type, 2 | 6 | 9).then(|| facts.source()),
                        members: Vec::new(),
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
