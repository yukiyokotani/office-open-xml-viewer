//! DOC inline PICF/OfficeArt pictures (MS-DOC 2.9.190-193; MS-ODRAW 2.2.15).
use super::{u16_at, u32_at, unsupported};
use crate::officeart::{raster::Image, record_with_end, Record};
use std::collections::{BTreeMap, BTreeSet};

pub(super) struct Store<'a> {
    data: &'a [u8],
    cache: BTreeMap<usize, Option<Picture<'a>>>,
    part_offsets: BTreeSet<usize>,
    budget: usize,
    remaining_bytes: usize,
    occurrences: u32,
    pub omitted: bool,
}
impl<'a> Store<'a> {
    pub fn new(data: &'a [u8]) -> Self {
        Self {
            data,
            cache: BTreeMap::new(),
            part_offsets: BTreeSet::new(),
            budget: 1_000_000,
            remaining_bytes: 128 * 1024 * 1024,
            occurrences: 0,
            omitted: false,
        }
    }
    pub fn drawing(&mut self, offset: usize) -> Result<String, String> {
        if !self.cache.contains_key(&offset) {
            if self.cache.len() >= 100_000 {
                return Err(unsupported("Word picture cache budget exceeded"));
            }
            let picture = read(self.data, offset, &mut self.budget)?;
            if let Some(picture) = picture {
                self.remaining_bytes = self
                    .remaining_bytes
                    .checked_sub(picture.image.bytes.len())
                    .ok_or_else(|| unsupported("Word retained media budget exceeded"))?;
            }
            self.cache.insert(offset, picture);
        }
        let Some(picture) = self.cache[&offset] else {
            self.omitted = true;
            return Ok(String::new());
        };
        if self.occurrences >= 1_000_000 {
            return Err(unsupported("Word picture occurrence budget exceeded"));
        }
        self.occurrences += 1;
        self.part_offsets.insert(offset);
        let id = self.occurrences;
        Ok(picture.xml(
            id,
            &format!("rImg{offset}"),
            &format!(
                r#"<wp:inline><wp:extent cx="{}" cy="{}"/>"#,
                picture.extent[0], picture.extent[1]
            ),
            "</wp:inline>",
        ))
    }
    pub fn relationships(&self) -> String {
        self.part_offsets.iter().filter_map(|offset| self.cache[offset].map(|p| format!(r#"<Relationship Id="rImg{offset}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image{offset}.{}"/>"#, p.image.extension))).collect()
    }
    pub fn begin_part(&mut self) {
        self.part_offsets.clear();
    }
    pub fn parts(&self) -> Vec<(String, &'a [u8])> {
        self.cache
            .iter()
            .filter_map(|(offset, picture)| {
                picture.map(|p| {
                    (
                        format!("word/media/image{offset}.{}", p.image.extension),
                        p.image.bytes,
                    )
                })
            })
            .collect()
    }
}

#[derive(Clone, Copy)]
pub(super) struct Picture<'a> {
    pub image: Image<'a>,
    pub extent: [i64; 2],
    pub crop: [i64; 4],
    pub flip: [bool; 2],
    pub rotation: i64,
}

impl Picture<'_> {
    /// Ordinary DrawingML content, shared by inline and floating DOC pictures.
    /// All caller-supplied fragments are generated from validated numeric/enumerated values.
    pub fn xml(&self, id: u32, relationship: &str, opening: &str, closing: &str) -> String {
        let [cx, cy] = self.extent;
        let [top, bottom, left, right] = self.crop;
        let [flip_h, flip_v] = self.flip.map(u8::from);
        let rotation = self.rotation;
        format!(
            r#"<w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{opening}<wp:docPr id="{id}" name="Legacy picture {id}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="{id}" name="Legacy picture {id}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="{relationship}"/><a:srcRect l="{left}" t="{top}" r="{right}" b="{bottom}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm rot="{rotation}" flipH="{flip_h}" flipV="{flip_v}"><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>{closing}</w:drawing>"#
        )
    }
}

fn read<'a>(
    data: &'a [u8],
    offset: usize,
    budget: &mut usize,
) -> Result<Option<Picture<'a>>, String> {
    let tail = data
        .get(offset..)
        .ok_or_else(|| unsupported("Word PICF offset out of range"))?;
    let size = u32_at(tail, 0)? as i32;
    if size < 68 || u16_at(tail, 4)? != 68 {
        return Err(unsupported("invalid Word PICF header"));
    }
    let data = tail
        .get(..size as usize)
        .ok_or_else(|| unsupported("truncated Word PICF data"))?;
    let mm = u16_at(data, 6)?;
    if !matches!(mm, 100 | 102) {
        return Ok(None);
    }
    let extent = [extent(data, 28, 32)?, extent(data, 30, 34)?];
    let start = if mm == 102 {
        // An optional source path is metadata, never an instruction to open it.
        69 + usize::from(
            *data
                .get(68)
                .ok_or_else(|| unsupported("truncated Word picture name"))?,
        )
    } else {
        68
    };
    let (shape, mut position) = record_with_end(data, start, budget, "Word inline shape")?;
    if shape.kind != 0xf004 || shape.version != 15 {
        return Err(unsupported("invalid Word inline shape container"));
    }
    let mut props = Options::default();
    let mut child = 0;
    while child < shape.payload.len() {
        let (record, end) = record_with_end(shape.payload, child, budget, "Word inline shape")?;
        child = end;
        match record.kind {
            0xf00b | 0xf122 => props.apply(record, budget)?,
            0xf00a => {
                if props.shape.is_some() || record.version != 2 || record.payload.len() != 8 {
                    return Err(unsupported("invalid Word inline shape properties"));
                }
                props.shape = Some(record.instance);
                let flags = u32_at(record.payload, 4)?;
                props.passive_picture = flags & 0x11d == 0;
                props.flip = [flags & 0x40 != 0, flags & 0x80 != 0];
            }
            _ => {} // No client data, OLE, script, or external resources executed.
        }
    }
    // MS-ODRAW 2.2.40: groups, deleted shapes, OLE and connectors are not
    // passive picture frames. Do not dereference their BLIPs.
    if props.shape != Some(75) || !props.passive_picture {
        return Ok(None);
    }
    let mut selected = None;
    for index in 0..props.blips {
        let (entry, end) = record_with_end(data, position, budget, "Word inline BLIP")?;
        position = end;
        if entry.kind != 0xf007 && !(0xf018..=0xf117).contains(&entry.kind) {
            return Err(unsupported("invalid Word inline BLIP record"));
        }
        if props.pib == Some(index) {
            selected = crate::officeart::raster::read_store_entry(entry, None, budget)?;
        }
    }
    let Some(image) = selected else {
        return Ok(None);
    };
    if props.crop[0] + props.crop[1] >= 100000 || props.crop[2] + props.crop[3] >= 100000 {
        return Err(unsupported("empty Word picture crop"));
    }
    Ok(Some(Picture {
        image,
        extent,
        crop: props.crop,
        flip: props.flip,
        rotation: props.rotation,
    }))
}

fn extent(data: &[u8], goal: usize, scale: usize) -> Result<i64, String> {
    let goal = i64::from(u16_at(data, goal)? as i16);
    let scaled = goal * i64::from(u16_at(data, scale)?);
    // MS-DOC PICMID: final size is goal * scale/1000, 15..31680 twips.
    if goal <= 0 || !(15000..=31680000).contains(&scaled) {
        return Err(unsupported("invalid Word picture display size"));
    }
    Ok((scaled * 635 + 500) / 1000) // Exact twip-to-EMU factor, rounded to one EMU.
}

#[derive(Default)]
pub(super) struct Options {
    shape: Option<u16>,
    passive_picture: bool,
    blips: usize,
    pub pib: Option<usize>,
    pub crop: [i64; 4],
    flip: [bool; 2],
    pub rotation: i64,
}
impl Options {
    fn apply(&mut self, record: Record<'_>, budget: &mut usize) -> Result<(), String> {
        self.apply_mode(record, budget, true)
    }
    pub fn apply_indexed(&mut self, record: Record<'_>, budget: &mut usize) -> Result<(), String> {
        self.apply_mode(record, budget, false)
    }
    fn apply_mode(
        &mut self,
        record: Record<'_>,
        budget: &mut usize,
        inline: bool,
    ) -> Result<(), String> {
        if record.version != 3 {
            return Err(unsupported("invalid Word picture option version"));
        }
        let count = usize::from(record.instance);
        *budget = budget
            .checked_sub(count)
            .ok_or_else(|| unsupported("Word picture option budget exceeded"))?;
        let mut complex_end = count * 6;
        if complex_end > record.payload.len() {
            return Err(unsupported("truncated Word picture options"));
        }
        for entry in record.payload[..complex_end].chunks_exact(6) {
            let key = u16_at(entry, 0)?;
            let id = key & 0x3fff;
            let value = u32_at(entry, 2)?;
            // MS-ODRAW 2.2.15: all BLIP-valued properties consume a slot,
            // regardless of fBid/fComplex/op. The visible picture is pib.
            if inline
                && matches!(
                    id,
                    0x104 | 0x10f | 0x186 | 0x1c5 | 0x545 | 0x585 | 0x5c5 | 0x605
                )
            {
                if id == 0x104 {
                    self.pib = Some(self.blips);
                }
                self.blips += 1;
                continue;
            }
            if !inline && key == 0x4104 {
                self.pib = (value as usize).checked_sub(1);
                continue;
            }
            if key & 0x8000 != 0 {
                complex_end = complex_end
                    .checked_add(value as usize)
                    .ok_or_else(|| unsupported("Word picture complex option overflow"))?;
                if complex_end > record.payload.len() {
                    return Err(unsupported("truncated Word picture complex option"));
                }
                continue;
            }
            match id {
                0x100..=0x103 => {
                    let fraction = i64::from(value as i32) * 100000;
                    let percent = (fraction + fraction.signum() * 32768) / 65536;
                    i32::try_from(percent)
                        .map_err(|_| unsupported("Word crop exceeds DrawingML percentage range"))?;
                    self.crop[usize::from(id - 0x100)] = percent;
                }
                4 => self.rotation = i64::from(value as i32) * 60000 / 65536,
                _ => {}
            }
        }
        Ok(())
    }
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
    fn png() -> Vec<u8> {
        let mut b = b"\x89PNG\r\n\x1a\n\0\0\0\x0dIHDR".to_vec();
        b.extend_from_slice(&2u32.to_be_bytes());
        b.extend_from_slice(&3u32.to_be_bytes());
        b.extend_from_slice(&[8, 2, 0, 0, 0, 0, 0, 0, 0]);
        b
    }
    fn fixture(properties: &[(u16, u32)], images: &[Vec<u8>]) -> Vec<u8> {
        let mut options = Vec::new();
        for (key, value) in properties {
            options.extend_from_slice(&key.to_le_bytes());
            options.extend_from_slice(&value.to_le_bytes());
        }
        let mut shape = record(0xf00a, (75 << 4) | 2, &[1, 0, 0, 0, 0x40, 8, 0, 0]);
        shape.extend(record(
            0xf00b,
            ((properties.len() as u16) << 4) | 3,
            &options,
        ));
        let mut data = vec![0u8; 68];
        data[4..6].copy_from_slice(&68u16.to_le_bytes());
        data[6..8].copy_from_slice(&100u16.to_le_bytes());
        data[28..30].copy_from_slice(&1440u16.to_le_bytes());
        data[30..32].copy_from_slice(&720u16.to_le_bytes());
        data[32..34].copy_from_slice(&500u16.to_le_bytes());
        data[34..36].copy_from_slice(&2000u16.to_le_bytes());
        data.extend(record(0xf004, 15, &shape));
        for image in images {
            data.extend(image);
        }
        let length = data.len() as u32;
        data[..4].copy_from_slice(&length.to_le_bytes());
        data
    }
    fn raster() -> Vec<u8> {
        record(0xf01e, 0x6e0 << 4, &[vec![0; 17], png()].concat())
    }
    #[test]
    fn inline_blips_follow_property_order_not_the_ignored_index_or_flags() {
        // A fill BLIP in primary options precedes pib in tertiary options.
        // Both fComplex and the huge op value are ignored for inline pib.
        let mut data = fixture(&[(0x0186, 0)], &[record(0xf01a, 0, &[]), raster()]);
        let properties = [
            0x0102u16.to_le_bytes().as_slice(),
            &16384u32.to_le_bytes(),
            &0xc104u16.to_le_bytes(),
            &u32::MAX.to_le_bytes(),
        ]
        .concat();
        let tertiary = record(0xf122, (2 << 4) | 3, &properties);
        let shape_length = u32_at(&data, 72).unwrap() as usize;
        let end = 76 + shape_length;
        data.splice(end..end, tertiary.iter().copied());
        data[72..76].copy_from_slice(&((shape_length + tertiary.len()) as u32).to_le_bytes());
        let length = data.len() as u32;
        data[..4].copy_from_slice(&length.to_le_bytes());
        let picture = read(&data, 0, &mut 100).unwrap().expect("inline PNG");
        assert_eq!(picture.image.bytes, png());
        assert_eq!(picture.extent, [457200, 914400]);
        assert_eq!(picture.crop, [0, 0, 25000, 0]);
        assert_eq!(picture.flip, [true, false]);
    }
    #[test]
    fn non_picture_and_external_only_shapes_are_not_fabricated() {
        let data = fixture(&[(0x0106, 2)], &[]);
        assert!(read(&data, 0, &mut 100).unwrap().is_none());
        for flag in [1u32, 4, 8, 16, 256] {
            let mut data = fixture(&[(0x104, 1)], &[record(0xf01e, 0, &[])]);
            data[88..92].copy_from_slice(&flag.to_le_bytes());
            // A malformed BLIP is deliberately not dereferenced on an OLE,
            // deleted, group, patriarch or connector shape.
            assert!(read(&data, 0, &mut 100).unwrap().is_none());
        }
    }
    #[test]
    fn ranges_dimensions_and_work_are_bounded() {
        let data = fixture(&[(0x0104, 1)], &[raster()]);
        assert!(read(&data, 0, &mut 0).is_err());
        assert!(read(&data[..data.len() - 1], 0, &mut 100).is_err());
        let mut invalid = data.clone();
        invalid[32..34].fill(0);
        assert!(read(&invalid, 0, &mut 100).is_err());
        let mut invalid = data;
        invalid[4] = 67;
        assert!(read(&invalid, 0, &mut 100).is_err());
    }

    #[test]
    fn cached_images_share_a_part_but_each_occurrence_has_unique_drawing_ids() {
        let data = fixture(&[(0x0104, 1)], &[raster()]);
        let mut store = Store::new(&data);
        assert!(store.drawing(0).unwrap().contains("docPr id=\"1\""));
        let budget = store.budget;
        assert!(store.drawing(0).unwrap().contains("docPr id=\"2\""));
        assert_eq!(store.budget, budget);
        assert_eq!(store.parts().len(), 1);
        assert_eq!(store.relationships().matches("<Relationship ").count(), 1);
        store.begin_part();
        assert!(store.relationships().is_empty());
        assert_eq!(store.parts().len(), 1); // The media remains shared.
        assert!(store.drawing(0).unwrap().contains("docPr id=\"3\""));
        assert_eq!(store.budget, budget);
        assert_eq!(store.relationships().matches("<Relationship ").count(), 1);
        assert_eq!(store.remaining_bytes, 128 * 1024 * 1024 - png().len());
        let mut store = Store::new(&data);
        store.remaining_bytes = 0;
        assert!(store.drawing(0).is_err());
    }

    #[test]
    fn inline_embedded_bse_is_supported_but_delayed_or_unsupported_data_is_not_followed() {
        let blip = raster();
        let mut bse = vec![0u8; 36];
        bse[0] = 6;
        bse[1] = 6;
        bse[20..24].copy_from_slice(&(blip.len() as u32).to_le_bytes());
        bse[24] = 1;
        let data = fixture(
            &[(0x104, 1)],
            &[record(0xf007, 0x62, &[bse.clone(), blip].concat())],
        );
        assert_eq!(
            read(&data, 0, &mut 100).unwrap().unwrap().image.bytes,
            png()
        );
        bse[28..32].copy_from_slice(&u32::MAX.to_le_bytes());
        let data = fixture(&[(0x104, 1)], &[record(0xf007, 0x62, &bse)]);
        let mut store = Store::new(&data);
        assert!(store.drawing(0).unwrap().is_empty());
        assert!(store.omitted);
        assert!(store.parts().is_empty());
    }
}
