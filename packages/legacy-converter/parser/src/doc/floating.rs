//! Main-story floating drawing anchors, MS-DOC 2.8.27 and 2.9.253.
use super::pictures::{Options as PictureOptions, Picture};
use super::{u16_at, u32_at, unsupported};
use crate::officeart::{
    raster::{read_store_entry, Image},
    record_with_end, Record,
};
use std::collections::BTreeMap;

pub(super) struct Store<'a> {
    anchors: Vec<Anchor>,
    shapes: BTreeMap<u32, (usize, u32, Record<'a>)>,
    entries: Vec<Record<'a>>,
    word: &'a [u8],
    images: BTreeMap<usize, Option<Image<'a>>>,
    budget: usize,
    remaining_bytes: usize,
    occurrences: u32,
    pub omitted: bool,
}

impl<'a> Store<'a> {
    pub fn read(word: &'a [u8], table: &'a [u8], main_units: usize) -> Result<Self, String> {
        let anchors = anchors(word, table, main_units)?;
        let mut result = Self {
            anchors,
            shapes: BTreeMap::new(),
            entries: Vec::new(),
            word,
            images: BTreeMap::new(),
            budget: 1_000_000,
            remaining_bytes: 128 * 1024 * 1024,
            occurrences: 0,
            omitted: false,
        };
        if result.anchors.is_empty() {
            return Ok(result);
        }
        // MS-DOC 2.9.171: the OfficeArt delay stream is WordDocument, NOT Data.
        let offset = u32_at(word, 0x22a)? as usize;
        let length = u32_at(word, 0x22e)? as usize;
        if length == 0 {
            result.omitted = true;
            return Ok(result);
        }
        let bytes = table
            .get(offset..)
            .and_then(|b| b.get(..length))
            .ok_or_else(|| unsupported("Word drawing group out of bounds"))?;
        let (group, mut position) =
            record_with_end(bytes, 0, &mut result.budget, "Word drawing group")?;
        if group.kind != 0xf000 || group.version != 15 {
            return Err(unsupported("invalid Word drawing group"));
        }
        let mut store_seen = false;
        for child in records(group.payload, &mut result.budget)? {
            if child.kind != 0xf001 {
                continue;
            }
            if store_seen || child.version != 15 {
                return Err(unsupported("invalid Word floating image store"));
            }
            store_seen = true;
            result.entries = records(child.payload, &mut result.budget)?;
            if result.entries.len() != usize::from(child.instance) {
                return Err(unsupported("Word floating image store count mismatch"));
            }
        }
        let mut main_seen = false;
        while position < bytes.len() {
            let label = bytes[position];
            let (drawing, end) =
                record_with_end(bytes, position + 1, &mut result.budget, "Word drawing")?;
            position = end;
            if label > 1 || drawing.kind != 0xf002 || drawing.version != 15 {
                return Err(unsupported("invalid Word drawing container"));
            }
            if label != 0 {
                continue;
            } // Do not leak header drawings into the body.
            if main_seen {
                return Err(unsupported("duplicate Word main drawing container"));
            }
            main_seen = true;
            for child in records(drawing.payload, &mut result.budget)? {
                if child.kind != 0xf003 {
                    continue;
                }
                if child.version != 15 {
                    return Err(unsupported("invalid Word shape group"));
                }
                for shape in records(child.payload, &mut result.budget)? {
                    // The topmost spgr contains independent shapes. Nested
                    // groups need their own coordinate transform and are not flattened.
                    if shape.kind != 0xf004 {
                        continue;
                    }
                    if shape.version != 15 {
                        return Err(unsupported("invalid Word floating shape container"));
                    }
                    let mut id = None;
                    let mut anchor_index = None;
                    for property in records(shape.payload, &mut result.budget)? {
                        match property.kind {
                            0xf00a if property.payload.len() == 8 => {
                                id = Some(u32_at(property.payload, 0)?);
                            }
                            0xf010 if property.payload.len() == 4 => {
                                anchor_index =
                                    usize::try_from(u32_at(property.payload, 0)? as i32).ok();
                            }
                            _ => {}
                        }
                    }
                    if let (Some(id), Some(anchor_index)) = (id, anchor_index) {
                        if result.shapes.len() >= 100_000 {
                            return Err(unsupported("Word floating shape budget exceeded"));
                        }
                        let order = result.shapes.len() as u32 + 1;
                        if result
                            .shapes
                            .insert(id, (anchor_index, order, shape))
                            .is_some()
                        {
                            return Err(unsupported("duplicate Word floating shape identifier"));
                        }
                    }
                }
            }
        }
        Ok(result)
    }

    pub fn drawing(&mut self, cp: usize) -> Result<String, String> {
        let Ok(index) = self.anchors.binary_search_by_key(&cp, |a| a.cp) else {
            self.omitted = true;
            return Ok(String::new());
        };
        let anchor = &self.anchors[index];
        let Some(&(anchor_index, order, shape)) = self.shapes.get(&anchor.shape_id) else {
            self.omitted = true;
            return Ok(String::new());
        };
        if anchor_index != index {
            return Err(unsupported("Word shape/anchor index mismatch"));
        }
        let mut picture = PictureOptions::default();
        let mut placement = Placement::default();
        let mut flags = None;
        let mut kind = None;
        for property in records(shape.payload, &mut self.budget)? {
            match property.kind {
                0xf00a => {
                    if flags.is_some() || property.version != 2 || property.payload.len() != 8 {
                        return Err(unsupported("invalid Word floating shape properties"));
                    }
                    flags = Some(u32_at(property.payload, 4)?);
                    kind = Some(property.instance);
                }
                0xf00b | 0xf122 => {
                    picture.apply_indexed(property, &mut self.budget)?;
                    placement.apply(property, &mut self.budget)?;
                }
                _ => {}
            }
        }
        let flags = flags.unwrap_or(0);
        if kind != Some(75)
            || flags & 0x11f != 0
            || placement.hidden
            || placement.script
            || picture.rotation != 0
            // SPA provides an explicit, host-defined coordinate origin.
            // Aligned positions require resolving OfficeArt posrelh/posrelv;
            // producer values differ from the published enumeration. Do not
            // guess a remapping or combine conflicting origins here.
            || placement.horizontal != 0
            || placement.vertical != 0
            || matches!(anchor.wrapping, 0 | 4 | 5)
        {
            self.omitted = true;
            return Ok(String::new());
        }
        let Some(image_index) = picture.pib else {
            self.omitted = true;
            return Ok(String::new());
        };
        if !self.images.contains_key(&image_index) {
            let entry = *self
                .entries
                .get(image_index)
                .ok_or_else(|| unsupported("Word floating image index out of bounds"))?;
            let image = read_store_entry(entry, Some(self.word), &mut self.budget)?;
            if let Some(image) = image {
                self.remaining_bytes = self
                    .remaining_bytes
                    .checked_sub(image.bytes.len())
                    .ok_or_else(|| unsupported("Word floating media budget exceeded"))?;
            }
            self.images.insert(image_index, image);
        }
        let Some(image) = self.images[&image_index] else {
            self.omitted = true;
            return Ok(String::new());
        };
        let [left, top, right, bottom] = anchor.rect.map(i64::from);
        let extent = [(right - left) * 635, (bottom - top) * 635];
        if extent.iter().any(|v| *v <= 0) {
            return Err(unsupported("invalid Word floating picture extent"));
        }
        if picture.crop[0] + picture.crop[1] >= 100000
            || picture.crop[2] + picture.crop[3] >= 100000
        {
            return Err(unsupported("empty Word floating picture crop"));
        }
        let x = position(anchor.horizontal, left * 635, false)?;
        let y = position(anchor.vertical, top * 635, true)?;
        let wrap = match anchor.wrapping {
            1 => "<wp:wrapTopAndBottom/>".to_string(),
            2 => format!("<wp:wrapSquare wrapText=\"{}\"/>", anchor.side),
            3 => "<wp:wrapNone/>".into(),
            _ => unreachable!(),
        };
        let [dist_l, dist_t, dist_r, dist_b] = placement.distances;
        let opening = format!(
            r#"<wp:anchor distL="{dist_l}" distT="{dist_t}" distR="{dist_r}" distB="{dist_b}" simplePos="0" relativeHeight="{}" behindDoc="{}" locked="{}" layoutInCell="{}" allowOverlap="{}"><wp:simplePos x="0" y="0"/>{x}{y}<wp:extent cx="{}" cy="{}"/>{wrap}"#,
            placement.z_order.unwrap_or(order),
            u8::from(anchor.behind),
            u8::from(anchor.locked),
            u8::from(placement.in_cell),
            u8::from(placement.overlap),
            extent[0],
            extent[1]
        );
        if self.occurrences >= 100_000 {
            return Err(unsupported("Word floating occurrence budget exceeded"));
        }
        self.occurrences += 1;
        let image = Picture {
            image,
            extent,
            crop: picture.crop,
            flip: [flags & 0x40 != 0, flags & 0x80 != 0],
            rotation: 0,
        };
        Ok(image.xml(
            1_000_000 + self.occurrences,
            &format!("rFloatImg{image_index}"),
            &opening,
            "</wp:anchor>",
        ))
    }
    pub fn relationships(&self) -> String {
        self.images.iter().filter_map(|(id,image)|image.map(|p|format!(r#"<Relationship Id="rFloatImg{id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/float{id}.{}"/>"#,p.extension))).collect()
    }
    pub fn parts(&self) -> Vec<(String, &'a [u8])> {
        self.images
            .iter()
            .filter_map(|(id, image)| {
                image.map(|p| (format!("word/media/float{id}.{}", p.extension), p.bytes))
            })
            .collect()
    }
}

fn records<'a>(bytes: &'a [u8], budget: &mut usize) -> Result<Vec<Record<'a>>, String> {
    let mut position = 0;
    let mut result = Vec::new();
    while position < bytes.len() {
        let (record, end) = record_with_end(bytes, position, budget, "Word OfficeArt")?;
        position = end;
        result.push(record);
    }
    Ok(result)
}

struct Placement {
    horizontal: u32,
    vertical: u32,
    distances: [u32; 4],
    in_cell: bool,
    overlap: bool,
    hidden: bool,
    script: bool,
    z_order: Option<u32>,
}
impl Default for Placement {
    fn default() -> Self {
        Self {
            horizontal: 0,
            vertical: 0,
            distances: [114300, 0, 114300, 0],
            in_cell: true,
            overlap: true,
            hidden: false,
            script: false,
            z_order: None,
        }
    }
}
impl Placement {
    fn apply(&mut self, property: Record<'_>, budget: &mut usize) -> Result<(), String> {
        let count = usize::from(property.instance);
        *budget = budget
            .checked_sub(count)
            .ok_or_else(|| unsupported("Word placement property budget exceeded"))?;
        let entries = property
            .payload
            .get(..count * 6)
            .ok_or_else(|| unsupported("truncated Word placement properties"))?;
        for p in entries.chunks_exact(6) {
            let key = u16_at(p, 0)?;
            let value = u32_at(p, 2)?;
            match key {
                0x384..=0x387 => {
                    if value > i32::MAX as u32 {
                        return Err(unsupported("negative Word picture wrap distance"));
                    }
                    self.distances[usize::from(key - 0x384)] = value;
                }
                0x38f => self.horizontal = value,
                0x391 => self.vertical = value,
                0x3aa if value != 0 => self.z_order = Some(value),
                0x3bf => {
                    for (bit, target) in [
                        (15, &mut self.in_cell),
                        (9, &mut self.overlap),
                        (1, &mut self.hidden),
                        (7, &mut self.script),
                    ] {
                        if value & (1 << (bit + 16)) != 0 {
                            *target = value & (1 << bit) != 0;
                        }
                    }
                }
                _ => {}
            }
        }
        Ok(())
    }
}

fn position(origin: &str, offset: i64, vertical: bool) -> Result<String, String> {
    let axis = if vertical { "V" } else { "H" };
    i32::try_from(offset)
        .map_err(|_| unsupported("Word floating position exceeds DrawingML range"))?;
    let value = format!("<wp:posOffset>{offset}</wp:posOffset>");
    Ok(format!(
        "<wp:position{axis} relativeFrom=\"{origin}\">{value}</wp:position{axis}>"
    ))
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct Anchor {
    pub cp: usize,
    pub shape_id: u32,
    pub rect: [i32; 4],
    pub horizontal: &'static str,
    pub vertical: &'static str,
    pub wrapping: u8,
    pub side: &'static str,
    pub behind: bool,
    pub locked: bool,
}

pub(super) fn anchors(word: &[u8], table: &[u8], main_units: usize) -> Result<Vec<Anchor>, String> {
    // FibRgFcLcb97 fields 40/41: main and header shape PLCs. This reader
    // intentionally does not assign header-document anchors to the main story.
    if word.len() < 0x1e2 {
        return Ok(Vec::new());
    }
    let size = u32_at(word, 0x1de)? as usize;
    if size == 0 {
        return Ok(Vec::new());
    }
    if size < 4 || !(size - 4).is_multiple_of(30) {
        return Err(unsupported("invalid Word floating-anchor PLC size"));
    }
    let count = (size - 4) / 30;
    if count > 100_000 {
        return Err(unsupported("Word floating-anchor budget exceeded"));
    }
    let offset = u32_at(word, 0x1da)? as usize;
    let bytes = table
        .get(offset..)
        .and_then(|b| b.get(..size))
        .ok_or_else(|| unsupported("Word floating-anchor PLC out of bounds"))?;
    let mut result = Vec::with_capacity(count);
    let mut ids = std::collections::BTreeSet::new();
    for i in 0..count {
        let cp = u32_at(bytes, i * 4)? as usize;
        let next = u32_at(bytes, (i + 1) * 4)? as usize;
        if cp > main_units || cp >= next {
            return Err(unsupported("invalid Word floating-anchor CP order"));
        }
        let record = &bytes[(count + 1) * 4 + i * 26..][..26];
        let shape_id = u32_at(record, 0)?;
        if !ids.insert(shape_id) {
            return Err(unsupported(
                "duplicate Word floating-anchor shape identifier",
            ));
        }
        let flags = u16_at(record, 20)?;
        let horizontal = match (flags >> 1) & 3 {
            0 => "margin",
            1 => "page",
            2 => "column",
            _ => return Err(unsupported("invalid Word floating horizontal origin")),
        };
        let vertical = match (flags >> 3) & 3 {
            0 => "margin",
            1 => "page",
            2 => "paragraph",
            _ => return Err(unsupported("invalid Word floating vertical origin")),
        };
        let wrapping = ((flags >> 5) & 15) as u8;
        if wrapping > 5 {
            return Err(unsupported("invalid Word floating wrap mode"));
        }
        let side = if matches!(wrapping, 1 | 3) {
            "bothSides"
        } else {
            match (flags >> 9) & 15 {
                0 => "bothSides",
                1 => "left",
                2 => "right",
                3 => "largest",
                _ => return Err(unsupported("invalid Word floating wrap side")),
            }
        };
        result.push(Anchor {
            cp,
            shape_id,
            rect: [
                u32_at(record, 4)? as i32,
                u32_at(record, 8)? as i32,
                u32_at(record, 12)? as i32,
                u32_at(record, 16)? as i32,
            ],
            horizontal,
            vertical,
            wrapping,
            side,
            behind: wrapping == 3 && flags & 0x4000 != 0,
            locked: flags & 0x8000 != 0,
        });
    }
    Ok(result)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(kind: u16, options: u16, body: &[u8]) -> Vec<u8> {
        [
            options.to_le_bytes().as_slice(),
            &kind.to_le_bytes(),
            &(body.len() as u32).to_le_bytes(),
            body,
        ]
        .concat()
    }
    fn drawing_input(shape_flags: u32, group_flags: u32) -> (Vec<u8>, Vec<u8>) {
        drawing_with_options(shape_flags, group_flags, &[])
    }
    fn drawing_with_options(
        shape_flags: u32,
        group_flags: u32,
        options: &[(u16, u32)],
    ) -> (Vec<u8>, Vec<u8>) {
        let (mut word, mut table) = input((1 << 1) | (1 << 3) | (2 << 5));
        let mut png = b"\x89PNG\r\n\x1a\n\0\0\0\x0dIHDR".to_vec();
        png.extend(2u32.to_be_bytes());
        png.extend(3u32.to_be_bytes());
        png.extend([8, 2, 0, 0, 0, 0, 0, 0, 0]);
        let blip = record(0xf01e, 0x6e0 << 4, &[vec![0; 17], png].concat());
        word.resize(1024, 0);
        word.extend(&blip);
        let mut bse = vec![0; 36];
        bse[0] = 6;
        bse[1] = 6;
        bse[24] = 1;
        bse[20..24].copy_from_slice(&(blip.len() as u32).to_le_bytes());
        bse[28..32].copy_from_slice(&1024u32.to_le_bytes());
        let group = record(0xf000, 15, &record(0xf001, 31, &record(0xf007, 0x62, &bse)));
        let mut props = [
            0x4104u16.to_le_bytes().as_slice(),
            &1u32.to_le_bytes(),
            &0x3bfu16.to_le_bytes(),
            &group_flags.to_le_bytes(),
        ]
        .concat();
        for (key, value) in options {
            props.extend(key.to_le_bytes());
            props.extend(value.to_le_bytes());
        }
        let shape = record(
            0xf004,
            15,
            &[
                record(
                    0xf00a,
                    (75 << 4) | 2,
                    &[1027u32.to_le_bytes(), shape_flags.to_le_bytes()].concat(),
                ),
                record(0xf00b, (((2 + options.len()) as u16) << 4) | 3, &props),
                record(0xf010, 0, &0u32.to_le_bytes()),
            ]
            .concat(),
        );
        let art = [
            group,
            vec![0],
            record(0xf002, 15, &record(0xf003, 15, &shape)),
        ]
        .concat();
        word[0x22a..0x22e].copy_from_slice(&(table.len() as u32).to_le_bytes());
        word[0x22e..0x232].copy_from_slice(&(art.len() as u32).to_le_bytes());
        table.extend(art);
        (word, table)
    }
    #[test]
    fn delayed_pictures_use_word_stream_and_share_parts_without_sharing_drawing_ids() {
        let (word, table) = drawing_input(0xa00, 0);
        let mut store = Store::read(&word, &table, 20).unwrap();
        let first = store.drawing(12).unwrap();
        let second = store.drawing(12).unwrap();
        assert!(first.contains("<wp:anchor"));
        assert!(first.contains("<wp:posOffset>-63500</wp:posOffset>"));
        assert!(first.contains("cx=\"254000\" cy=\"190500\""));
        assert!(first.contains("id=\"1000001\""));
        assert!(second.contains("id=\"1000002\""));
        assert_eq!(store.parts().len(), 1);
        assert_eq!(store.relationships().matches("<Relationship ").count(), 1);
        assert!(store.parts()[0].1.starts_with(b"\x89PNG"));
        let mut truncated = Store::read(&word[..1024], &table, 20).unwrap();
        assert!(truncated.drawing(12).is_err());
    }
    #[test]
    fn hidden_ole_and_script_shapes_are_not_dereferenced() {
        for (shape, group) in [(0xa10, 0), (0xa00, 0x00020002), (0xa00, 0x00800080)] {
            let (word, table) = drawing_input(shape, group);
            let mut store = Store::read(&word[..1024], &table, 20).unwrap();
            assert!(store.drawing(12).unwrap().is_empty());
            assert!(store.parts().is_empty());
            assert!(store.omitted);
        }
    }
    #[test]
    fn ambiguous_alignment_and_rotated_shapes_do_not_dereference_blips() {
        for key in [0x38f, 0x391, 4] {
            for value in 1..=5 {
                let operand = if key == 4 { value << 16 } else { value };
                let (word, table) = drawing_with_options(0xa00, 0, &[(key, operand)]);
                let mut store = Store::read(&word[..1024], &table, 20).unwrap();
                assert!(store.drawing(12).unwrap().is_empty());
                assert!(store.parts().is_empty());
                assert!(store.omitted);
            }
        }
    }
    #[test]
    fn enforces_media_occurrence_and_operation_budgets() {
        let (word, table) = drawing_input(0xa00, 0);
        let mut store = Store::read(&word, &table, 20).unwrap();
        store.remaining_bytes = 0;
        assert!(store.drawing(12).unwrap_err().contains("media budget"));
        assert!(store.parts().is_empty());

        let mut store = Store::read(&word, &table, 20).unwrap();
        store.occurrences = 100_000;
        assert!(store.drawing(12).unwrap_err().contains("occurrence budget"));

        let mut store = Store::read(&word, &table, 20).unwrap();
        store.budget = 0;
        assert!(store.drawing(12).is_err());
    }
    #[test]
    fn does_not_reassign_header_or_nested_group_drawings_to_main_story() {
        let (word, mut table) = drawing_input(0xa00, 0);
        let start = u32_at(&word, 0x22a).unwrap() as usize;
        let (_, group_end) = record_with_end(&table, start, &mut 100, "test").unwrap();
        table[group_end] = 1;
        let mut store = Store::read(&word, &table, 20).unwrap();
        assert!(store.drawing(12).unwrap().is_empty());
        assert!(store.parts().is_empty());

        table[group_end] = 0;
        // Replace the independent SpContainer tag with a nested SpgrContainer.
        // Its children are deliberately not traversed without a group transform.
        let shape_start = group_end + 1 + 8 + 8;
        table[shape_start + 2..shape_start + 4].copy_from_slice(&0xf003u16.to_le_bytes());
        let mut store = Store::read(&word, &table, 20).unwrap();
        assert!(store.drawing(12).unwrap().is_empty());
        assert!(store.parts().is_empty());
    }
    #[test]
    fn placement_masks_honor_explicit_false_and_ignore_unused_bits() {
        let mut placement = Placement::default();
        let apply = |p: &mut Placement, value: u32| {
            let body = [0x3bfu16.to_le_bytes().as_slice(), &value.to_le_bytes()].concat();
            p.apply(
                Record {
                    version: 3,
                    instance: 1,
                    kind: 0xf00b,
                    payload: &body,
                },
                &mut 10,
            )
            .unwrap();
        };
        apply(&mut placement, 0x00020002);
        assert!(placement.hidden);
        apply(&mut placement, 0);
        assert!(placement.hidden);
        apply(&mut placement, 0x00020000);
        assert!(!placement.hidden);
        apply(&mut placement, 0x82008000);
        assert!(placement.in_cell);
        assert!(!placement.overlap);
        assert_eq!(
            position("margin", 999, false).unwrap(),
            "<wp:positionH relativeFrom=\"margin\"><wp:posOffset>999</wp:posOffset></wp:positionH>"
        );
        assert!(position("page", i64::MAX, false).is_err());
    }
    fn input(flags: u16) -> (Vec<u8>, Vec<u8>) {
        let mut word = vec![0u8; 0x232];
        word[0x1de..0x1e2].copy_from_slice(&34u32.to_le_bytes());
        let mut table = [
            12u32.to_le_bytes(),
            30u32.to_le_bytes(),
            1027u32.to_le_bytes(),
            (-100i32).to_le_bytes(),
            200i32.to_le_bytes(),
            300i32.to_le_bytes(),
            500i32.to_le_bytes(),
        ]
        .concat();
        table.extend(flags.to_le_bytes());
        table.extend([0; 4]);
        (word, table)
    }
    #[test]
    fn preserves_signed_rectangle_origin_wrap_and_layer() {
        let (word, table) = input((1 << 1) | (2 << 3) | (3 << 5) | (1 << 14) | (1 << 15));
        let a = anchors(&word, &table, 20).unwrap();
        assert_eq!(a.len(), 1);
        assert_eq!(a[0].cp, 12);
        assert_eq!(a[0].shape_id, 1027);
        assert_eq!(a[0].rect, [-100, 200, 300, 500]);
        assert_eq!(a[0].horizontal, "page");
        assert_eq!(a[0].vertical, "paragraph");
        assert_eq!(a[0].wrapping, 3);
        assert!(a[0].behind);
        assert!(a[0].locked);
        // The final sentinel CP is undefined, apart from monotonicity. It may
        // exceed ccpText and must not be treated as a live anchor.
    }
    #[test]
    fn ignored_wrapping_fields_do_not_reject_top_bottom_anchors() {
        let (word, table) = input((1 << 5) | (15 << 9) | (1 << 14));
        let a = anchors(&word, &table, 20).unwrap();
        assert_eq!(a[0].side, "bothSides");
        assert!(!a[0].behind);
    }
    #[test]
    fn rejects_invalid_plc_size_origin_and_live_cp_but_allows_absent_table() {
        let (word, table) = input(3 << 1);
        assert!(anchors(&word, &table, 20).is_err());
        let (word, table) = input(0);
        assert!(anchors(&word, &table, 11).is_err());
        assert!(anchors(&word, &table[..table.len() - 1], 20).is_err());
        let mut word = word;
        word[0x1de..0x1e2].fill(0);
        assert!(anchors(&word, &[], 20).unwrap().is_empty());
    }
}
