//! Anchored OfficeArt groups for the direct model (MS-ODRAW 2.2.16, 2.2.38,
//! 2.2.39 and 2.2.40; MS-DOC 2.9.253).
//!
//! The top-level group's SPA rectangle is its frame on the page. Its
//! OfficeArtFSPGR record defines the group coordinate system, in which each
//! member's OfficeArtChildAnchor is expressed; a nested group maps its own
//! OfficeArtFSPGR onto its child anchor in the parent's system. Mapping is the
//! linear rectangle-to-rectangle transform these records define.
//!
//! Rotation (MS-ODRAW 2.3.18.5 defines only the clockwise angle about the
//! centre) follows Word's own DOCX of the same drawings in the local corpus,
//! the DOC files having been saved by Word from those DOCX files:
//! - A member whose angle, normalised to [0, 360), lies in [45, 135) or
//!   [225, 315) stores its rotated bounds as its child anchor: Word's DOCX
//!   gives the same angle with width and height exchanged about the same
//!   centre (90 degrees, three members). Members at 44.25, 136.16, 224.39 and
//!   318.31 degrees store their unrotated box unchanged. This is the rule the
//!   PowerPoint control in ppt/drawing/direct_transform.rs established for
//!   PowerPoint; Word agrees on every sampled angle. Mapping the stored bounds
//!   through the group before the exchange reproduces the scale order that
//!   the DOCX reader applies to Word's exact quarter turns.
//! - A top-level group rotated by 180 degrees keeps its unrotated SPA frame
//!   and member coordinates (Word's DOCX: `rot="10800000"` on the group
//!   transform with the same extent); members turn about the group centre.
//! Other group angles, nested rotated groups, rotated flipped members,
//! rotated pictures and rotated text shapes have no Word evidence yet and
//! stay unsupported. Members that are straight connectors keep their static
//! path; endpoint rerouting is not reconstructed.

use super::super::{u32_at, unsupported};
use super::{direct_alignment, records, shape, Anchor, Content, Placement, ResolvedDrawing, Store};
use crate::doc::pictures::Options as PictureOptions;
use crate::officeart::{properties, Record};

const MAX_MEMBERS: usize = 10_000;
const MAX_DEPTH: usize = 16;

/// One displayed group member, in source (paint) order.
pub(in crate::doc) struct Member {
    pub content: Content,
    /// Unrotated x, y, width, height in EMUs relative to the top-level group
    /// frame; `rotation` (degrees, clockwise) turns it about its centre.
    pub frame: [f64; 4],
    pub rotation: f64,
    pub flip: [bool; 2],
    pub spid: u32,
}

/// The linear map from one group's coordinate system to the top-level frame.
#[derive(Clone, Copy)]
struct Map {
    origin: [f64; 2],
    scale: [f64; 2],
    offset: [f64; 2],
}

impl Map {
    fn new(system: [i32; 4], frame: [f64; 4]) -> Result<Self, String> {
        let [left, top, right, bottom] = system.map(f64::from);
        if right <= left || bottom <= top {
            return Err(unsupported("empty Word drawing group coordinate system"));
        }
        Ok(Self {
            origin: [left, top],
            scale: [frame[2] / (right - left), frame[3] / (bottom - top)],
            offset: [frame[0], frame[1]],
        })
    }

    fn rect(&self, rect: [i32; 4]) -> Result<[f64; 4], String> {
        let [left, top, right, bottom] = rect.map(f64::from);
        if right < left || bottom < top {
            return Err(unsupported("inverted Word drawing child anchor"));
        }
        let frame = [
            self.offset[0] + (left - self.origin[0]) * self.scale[0],
            self.offset[1] + (top - self.origin[1]) * self.scale[1],
            (right - left) * self.scale[0],
            (bottom - top) * self.scale[1],
        ];
        if frame.iter().any(|value| !value.is_finite()) {
            return Err(unsupported("Word drawing child anchor exceeds its range"));
        }
        Ok(frame)
    }
}

fn rect(record: Record<'_>, name: &str) -> Result<[i32; 4], String> {
    if record.payload.len() != 16 {
        return Err(unsupported(format!("invalid Word drawing {name}")));
    }
    let mut values = [0; 4];
    for (index, value) in values.iter_mut().enumerate() {
        *value = u32_at(record.payload, index * 4)? as i32;
    }
    Ok(values)
}

/// The facts of one OfficeArtSpContainer header shared by groups and members.
struct Header<'a> {
    kind: u16,
    flags: u32,
    spid: u32,
    group_system: Option<[i32; 4]>,
    child_anchor: Option<[i32; 4]>,
    placement: Placement,
    picture: PictureOptions,
    record: Record<'a>,
}

fn header<'a>(shape: Record<'a>, budget: &mut usize) -> Result<Header<'a>, String> {
    if shape.kind != 0xf004 || shape.version != 15 {
        return Err(unsupported("invalid Word drawing group member"));
    }
    let mut result = Header {
        kind: 0,
        flags: 0,
        spid: 0,
        group_system: None,
        child_anchor: None,
        placement: Placement::default(),
        picture: PictureOptions::default(),
        record: shape,
    };
    let mut seen = false;
    for record in records(shape.payload, budget)? {
        match record.kind {
            0xf00a => {
                if seen || record.version != 2 || record.payload.len() != 8 {
                    return Err(unsupported("invalid Word floating shape properties"));
                }
                seen = true;
                result.kind = record.instance;
                result.spid = u32_at(record.payload, 0)?;
                result.flags = u32_at(record.payload, 4)?;
            }
            0xf009 => result.group_system = Some(rect(record, "group coordinate system")?),
            0xf00f => result.child_anchor = Some(rect(record, "child anchor")?),
            0xf00b | 0xf122 => {
                result.picture.apply_indexed(record, budget)?;
                result.placement.apply(record, budget)?;
            }
            _ => {}
        }
    }
    if !seen {
        return Err(unsupported("Word drawing group member lacks its shape"));
    }
    Ok(result)
}

/// Classify a group shape's own properties. Groups have no geometry or
/// paint; only identification, editing, placement and zero relative-size
/// properties are accepted.
/// Returns the group's own rotation in degrees (0 or 180; see above).
fn group_properties(shape: Record<'_>, budget: &mut usize) -> Result<f64, String> {
    let mut rotation = 0.0;
    let mut visit = |property: properties::Property<'_>| -> Result<(), String> {
        let id = property.opid & 0x3fff;
        let value = property.value;
        match id {
            0x380 | 0x381 | 0x3a9 if property.complex.is_some() => Ok(()),
            0x4 if value == 0 => Ok(()),
            0x4 if value == 180 << 16 => {
                rotation = 180.0;
                Ok(())
            }
            0x4 => Err(unsupported("rotated Word drawing groups are not supported")),
            0x40..=0x7f | 0x384..=0x388 | 0x38f..=0x392 | 0x3aa | 0x3bf | 0x7c4 | 0x7c5 => Ok(()),
            0x7c0..=0x7c3 if value == 0 => Ok(()),
            0x53f if value & 0x0001_0001 != 0x0001_0001 => Ok(()),
            _ => Err(unsupported(format!(
                "Word drawing group property {id:#06x}={value:#x} is not supported"
            ))),
        }
    };
    for record in records(shape.payload, budget)? {
        match record.kind {
            0xf00b => properties::visit(record, budget, &mut visit)?,
            0xf122 => properties::visit_tertiary(record, budget, &mut visit)?,
            _ => {}
        }
    }
    Ok(rotation)
}

/// FSP flags of a group shape: fGroup, optionally fChild when nested.
/// Patriarch, deleted, OLE, master-linked, flipped, connector and background
/// groups are rejected (MS-ODRAW 2.2.40).
fn check_group_flags(flags: u32, nested: bool) -> Result<(), String> {
    if flags & 0x1 == 0 || flags & 0x2 != u32::from(nested) << 1 {
        return Err(unsupported("invalid Word drawing group shape"));
    }
    if flags & 0xc0 != 0 {
        return Err(unsupported("flipped Word drawing groups are not supported"));
    }
    if flags & 0x53c != 0 {
        return Err(unsupported(
            "Word drawing group has unsupported shape flags",
        ));
    }
    Ok(())
}

impl Store<'_> {
    pub(super) fn resolve_group(
        &mut self,
        anchor: &Anchor,
        order: u32,
        group: Record<'_>,
    ) -> Result<Option<ResolvedDrawing>, String> {
        let children = records(group.payload, &mut self.budget)?;
        let head = header(children[0], &mut self.budget)?;
        check_group_flags(head.flags, false)?;
        if head.placement.script {
            self.omitted = true;
            return Ok(None);
        }
        if head.placement.hidden {
            // MS-ODRAW 2.3.4.44 fHidden, as for single drawings.
            return Ok(None);
        }
        let turned = group_properties(head.record, &mut self.budget)? != 0.0;
        let [left, top, right, bottom] = anchor.rect.map(i64::from);
        let extent = [(right - left) * 635, (bottom - top) * 635];
        if extent.iter().any(|value| *value <= 0) {
            return Err(unsupported("invalid Word drawing group extent"));
        }
        let system = head
            .group_system
            .ok_or_else(|| unsupported("Word drawing group lacks its coordinate system"))?;
        let map = Map::new(system, [0.0, 0.0, extent[0] as f64, extent[1] as f64])?;
        let align = direct_alignment(anchor, &head.placement)?;
        if matches!(anchor.wrapping, 0 | 4 | 5) {
            return Err(unsupported(
                "Word drawing group uses an unsupported wrap contour",
            ));
        }
        let mut members = Vec::new();
        self.group_members(&children[1..], map, 0, &mut members)?;
        if members.is_empty() {
            return Ok(None);
        }
        if turned {
            // A half turn about the group centre maps each member's centre
            // through the centre and adds 180 degrees; flips are unchanged.
            for member in &mut members {
                if matches!(member.content, Content::Picture { .. }) {
                    return Err(unsupported(
                        "pictures in rotated Word drawing groups are not supported",
                    ));
                }
                if matches!(&member.content, Content::Shape(shape) if shape.text.is_some()) {
                    return Err(unsupported("rotated Word drawing text is not supported"));
                }
                // fUseShapeAnchor 0 keeps a picture fill upright while the
                // shape turns; the DOCX shape model has no such fill.
                if matches!(&member.content, Content::Shape(shape) if shape.fill_picture.is_some_and(|(_, rotates)| !rotates))
                {
                    return Err(unsupported(
                        "upright picture fills in rotated Word shapes are not supported",
                    ));
                }
                let [x, y, width, height] = member.frame;
                member.frame = [
                    extent[0] as f64 - x - width,
                    extent[1] as f64 - y - height,
                    width,
                    height,
                ];
                member.rotation = (member.rotation + 180.0).rem_euclid(360.0);
            }
        }
        self.finish(
            anchor,
            &head.placement,
            order,
            extent,
            0,
            align,
            Content::Group(members),
        )
        .map(Some)
    }

    fn group_members(
        &mut self,
        children: &[Record<'_>],
        map: Map,
        depth: usize,
        members: &mut Vec<Member>,
    ) -> Result<(), String> {
        if depth > MAX_DEPTH {
            return Err(unsupported("Word drawing groups are nested too deeply"));
        }
        for child in children {
            if members.len() >= MAX_MEMBERS {
                return Err(unsupported("Word drawing group member budget exceeded"));
            }
            match child.kind {
                0xf003 if child.version == 15 => {
                    let nested = records(child.payload, &mut self.budget)?;
                    let first = *nested
                        .first()
                        .ok_or_else(|| unsupported("empty Word drawing group"))?;
                    let head = header(first, &mut self.budget)?;
                    check_group_flags(head.flags, true)?;
                    if head.placement.script {
                        self.omitted = true;
                        continue;
                    }
                    if head.placement.hidden {
                        continue;
                    }
                    if group_properties(head.record, &mut self.budget)? != 0.0 {
                        return Err(unsupported(
                            "rotated nested Word drawing groups are not supported",
                        ));
                    }
                    let frame = map.rect(head.child_anchor.ok_or_else(|| {
                        unsupported("nested Word drawing group lacks its child anchor")
                    })?)?;
                    let system = head.group_system.ok_or_else(|| {
                        unsupported("Word drawing group lacks its coordinate system")
                    })?;
                    self.group_members(&nested[1..], Map::new(system, frame)?, depth + 1, members)?;
                }
                0xf004 => {
                    if let Some(member) = self.group_member(*child, map)? {
                        members.push(member);
                    }
                }
                _ => return Err(unsupported("invalid Word drawing group member")),
            }
        }
        Ok(())
    }

    fn group_member(&mut self, child: Record<'_>, map: Map) -> Result<Option<Member>, String> {
        let head = header(child, &mut self.budget)?;
        if head.placement.script {
            self.omitted = true;
            return Ok(None);
        }
        if head.placement.hidden {
            return Ok(None);
        }
        let mut frame = map
            .rect(head.child_anchor.ok_or_else(|| {
                unsupported("Word drawing group member lacks its child anchor")
            })?)?;
        let extent = [frame[2].round() as i64, frame[3].round() as i64];
        let flip = [head.flags & 0x40 != 0, head.flags & 0x80 != 0];
        let content = if head.kind == 75 {
            // A passive picture frame member (MS-ODRAW 2.2.40 fChild only).
            if head.flags & 0x13d != 0 || head.flags & 0x2 == 0 {
                return Err(unsupported(
                    "Word grouped picture has unsupported shape flags",
                ));
            }
            if head.picture.rotation != 0 {
                return Err(unsupported(
                    "rotated Word grouped pictures are not supported",
                ));
            }
            let Some(image_index) = head.picture.pib else {
                self.omitted = true;
                return Ok(None);
            };
            let Some(extension) = self.load_image(image_index)? else {
                self.omitted = true;
                return Ok(None);
            };
            if extent.iter().any(|value| *value <= 0) {
                return Err(unsupported("invalid Word grouped picture extent"));
            }
            let crop = head.picture.crop;
            if crop[0] + crop[1] >= 100000 || crop[2] + crop[3] >= 100000 {
                return Err(unsupported("empty Word grouped picture crop"));
            }
            Content::Picture {
                image_index,
                extension,
                crop,
            }
        } else {
            let facts = shape::Facts::read(
                head.kind,
                head.flags,
                true,
                head.record,
                extent,
                &mut self.budget,
            )?;
            if !self.load_fill_picture(&facts)? {
                self.omitted = true;
                return Ok(None);
            }
            Content::Shape(Box::new(facts))
        };
        if matches!(&content, Content::Shape(shape) if shape.relative_size != [None; 2]) {
            return Err(unsupported(
                "relative sizes of Word drawing group members are not supported",
            ));
        }
        let rotation = match &content {
            Content::Shape(shape) => shape.rotation.rem_euclid(360.0),
            _ => 0.0,
        };
        if rotation != 0.0 {
            if flip != [false; 2] {
                return Err(unsupported(
                    "rotated flipped Word drawing members are not supported",
                ));
            }
            if matches!(&content, Content::Shape(shape) if shape.text.is_some()) {
                return Err(unsupported("rotated Word drawing text is not supported"));
            }
            // fUseShapeAnchor 0 keeps a picture fill upright while the
            // shape turns; the DOCX shape model has no such fill.
            if matches!(&content, Content::Shape(shape) if shape.fill_picture.is_some_and(|(_, rotates)| !rotates))
            {
                return Err(unsupported(
                    "upright picture fills in rotated Word shapes are not supported",
                ));
            }
            if (45.0..135.0).contains(&rotation) || (225.0..315.0).contains(&rotation) {
                // The child anchor holds the rotated bounds (see above).
                let [x, y, width, height] = frame;
                frame = [
                    x + (width - height) / 2.0,
                    y + (height - width) / 2.0,
                    height,
                    width,
                ];
            }
        }
        Ok(Some(Member {
            content,
            frame,
            rotation,
            flip,
            spid: head.spid,
        }))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn maps_child_anchors_through_nested_coordinate_systems() {
        let outer = Map::new([0, 0, 1000, 500], [0.0, 0.0, 2000.0, 2000.0]).unwrap();
        assert_eq!(
            outer.rect([100, 50, 600, 250]).unwrap(),
            [200.0, 200.0, 1000.0, 800.0]
        );
        // A nested group occupies [500, 0, 1000, 500] and defines -10..10.
        let frame = outer.rect([500, 0, 1000, 500]).unwrap();
        let nested = Map::new([-10, -10, 10, 10], frame).unwrap();
        assert_eq!(
            nested.rect([0, 0, 10, 10]).unwrap(),
            [1500.0, 1000.0, 500.0, 1000.0]
        );
        // Zero-width lines keep their axis; inverted anchors and empty systems fail.
        assert_eq!(outer.rect([10, 10, 10, 60]).unwrap()[2], 0.0);
        assert!(outer.rect([10, 10, 5, 60]).is_err());
        assert!(Map::new([5, 0, 5, 10], frame).is_err());
    }

    #[test]
    fn group_properties_reject_rotation_geometry_and_paint() {
        let table = |properties: &[(u16, u32)]| {
            let mut body = Vec::new();
            for (key, value) in properties {
                body.extend(key.to_le_bytes());
                body.extend(value.to_le_bytes());
            }
            let fopt = [
                ((properties.len() as u16) << 4 | 3)
                    .to_le_bytes()
                    .as_slice(),
                &0xf00bu16.to_le_bytes(),
                &(body.len() as u32).to_le_bytes(),
                &body,
            ]
            .concat();
            [
                15u16.to_le_bytes().as_slice(),
                &0xf004u16.to_le_bytes(),
                &(fopt.len() as u32).to_le_bytes(),
                &fopt,
            ]
            .concat()
        };
        let check = |properties: &[(u16, u32)]| {
            let bytes = table(properties);
            let (record, _) =
                crate::officeart::record_with_end(&bytes, 0, &mut 100, "test").unwrap();
            group_properties(record, &mut 100)
        };
        assert!(check(&[(0x4, 0), (0x3bf, 0x220000), (0x390, 1), (0x7c4, 0)]).is_ok());
        assert!(check(&[(0x4, 0x5a_0000)]).unwrap_err().contains("rotated"));
        for property in [(0x181, 0), (0x145, 0), (0x7c1, 5), (0x53f, 0x10001)] {
            assert!(check(&[property]).is_err(), "{property:x?}");
        }
    }

    #[test]
    fn group_shape_flags_reject_flips_and_structural_bits() {
        assert!(check_group_flags(0x201, false).is_ok());
        assert!(check_group_flags(0x203, true).is_ok());
        assert!(check_group_flags(0x203, false).is_err());
        assert!(check_group_flags(0x201, true).is_err());
        assert!(check_group_flags(0x200, false).is_err());
        for flag in [0x4, 0x8, 0x10, 0x20, 0x40, 0x80, 0x100, 0x400] {
            assert!(check_group_flags(0x201 | flag, false).is_err(), "{flag:#x}");
        }
    }
}
