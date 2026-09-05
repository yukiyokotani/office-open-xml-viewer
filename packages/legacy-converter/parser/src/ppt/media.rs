//! Passive raster BLIPs: MS-PPT 2.1.3/2.4.3; MS-ODRAW 2.2.20-32.
use super::*;
use std::collections::{BTreeMap, BTreeSet};

// Resource policy: do not pass dimension bombs to the ordinary image decoder.
const MAX_PIXELS: u64 = 40_000_000;
const MAX_MEDIA_BYTES: usize = 128 * 1024 * 1024;

pub(super) fn catalog<'a>(
    children: &[Record<'a>],
    budget: &mut usize,
) -> Result<Vec<Record<'a>>, String> {
    let mut result = None;
    let mut group_seen = false;
    for group in children.iter().filter(|r| r.kind == 1035) {
        if group_seen || group.version != 15 {
            return Err(unsupported("invalid PowerPoint drawing group"));
        }
        group_seen = true;
        let dgg = parse_records(group.payload, budget)?;
        if dgg.len() != 1 || dgg[0].kind != 0xf000 || dgg[0].version != 15 {
            return Err(unsupported("invalid PowerPoint OfficeArt drawing group"));
        }
        for store in parse_records(dgg[0].payload, budget)?
            .iter()
            .filter(|r| r.kind == 0xf001)
        {
            if result.is_some() || store.version != 15 {
                return Err(unsupported("invalid PowerPoint image store"));
            }
            let entries = parse_records(store.payload, budget)?;
            if entries.len() != usize::from(store.instance) {
                return Err(unsupported("PowerPoint image store count mismatch"));
            }
            result = Some(entries);
        }
    }
    Ok(result.unwrap_or_default())
}

#[derive(Clone, Copy)]
struct Image<'a> {
    bytes: &'a [u8],
    extension: &'static str,
}

pub(super) struct Store<'a> {
    entries: &'a [Record<'a>],
    delayed: &'a [u8],
    images: BTreeMap<u32, Option<Image<'a>>>,
    used: BTreeSet<u32>,
    remaining: usize,
}
impl<'a> Store<'a> {
    pub fn new(entries: &'a [Record<'a>], delayed: &'a [u8]) -> Self {
        Self {
            entries,
            delayed,
            images: BTreeMap::new(),
            used: BTreeSet::new(),
            remaining: MAX_MEDIA_BYTES,
        }
    }
    pub fn begin_slide(&mut self) {
        self.used.clear();
    }
    pub fn reference(&mut self, index: u32, budget: &mut usize) -> Result<bool, String> {
        if index == 0 {
            return Ok(false);
        }
        if !self.images.contains_key(&index) {
            let entry = *self
                .entries
                .get((index - 1) as usize)
                .ok_or_else(|| unsupported("PowerPoint picture index out of range"))?;
            let image = image(entry, self.delayed, budget)?;
            if let Some(image) = image {
                self.remaining = self
                    .remaining
                    .checked_sub(image.bytes.len())
                    .ok_or_else(|| unsupported("PowerPoint retained media budget exceeded"))?;
            }
            self.images.insert(index, image);
        }
        if self.images[&index].is_none() {
            return Ok(false);
        }
        self.used.insert(index);
        Ok(true)
    }
    pub fn relationships(&self) -> String {
        let mut xml = String::new();
        for index in &self.used {
            let image = self.images[index].expect("only supported referenced images");
            xml.push_str(&format!("<Relationship Id=\"rImg{index}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image{index}.{}\"/>", image.extension));
        }
        xml
    }
    pub fn parts(&self) -> Vec<(String, &'a [u8])> {
        self.images
            .iter()
            .filter_map(|(id, image)| {
                image.map(|image| {
                    (
                        format!("ppt/media/image{id}.{}", image.extension),
                        image.bytes,
                    )
                })
            })
            .collect()
    }
}

fn image<'a>(
    entry: Record<'a>,
    delayed: &'a [u8],
    budget: &mut usize,
) -> Result<Option<Image<'a>>, String> {
    let blip = if entry.kind == 0xf007 {
        let b = entry.payload;
        if entry.version != 2
            || b.len() < 36
            || (entry.instance != u16::from(b[0]) && entry.instance != u16::from(b[1]))
        {
            return Err(unsupported("invalid PowerPoint BLIP store entry"));
        }
        let name = usize::from(b[33]);
        if name % 2 != 0 || 36 + name > b.len() {
            return Err(unsupported("invalid PowerPoint BLIP name length"));
        }
        let start = 36 + name;
        let size = u32_at(b, 20)? as usize;
        if u32_at(b, 24)? == 0 {
            return Ok(None);
        } // Empty slots retain their index.
        let source = if start < b.len() {
            b.get(start..)
                .filter(|s| s.len() == size)
                .ok_or_else(|| unsupported("PowerPoint embedded BLIP size mismatch"))?
        } else {
            let offset = u32_at(b, 28)? as usize;
            delayed
                .get(offset..)
                .and_then(|s| s.get(..size))
                .ok_or_else(|| unsupported("PowerPoint delayed BLIP range out of bounds"))?
        };
        let record = parse_record_at(source, 0, budget)?;
        if record.payload.len() + 8 != source.len() {
            return Err(unsupported("PowerPoint BLIP record size mismatch"));
        }
        record
    } else {
        entry
    };
    let (extension, prefix) = match (blip.kind, blip.instance) {
        (0xf01e, 0x6e0) => ("png", 17),
        (0xf01e, 0x6e1) => ("png", 33),
        (0xf01d, 0x46a | 0x6e2) => ("jpg", 17),
        (0xf01d, 0x46b | 0x6e3) => ("jpg", 33),
        (0xf01d | 0xf01e, _) => return Err(unsupported("invalid PowerPoint raster BLIP instance")),
        _ => return Ok(None), // No WMF/EMF/PICT/DIB or active-object decoding here.
    };
    if blip.version != 0 {
        return Err(unsupported("invalid PowerPoint raster BLIP version"));
    }
    let bytes = blip
        .payload
        .get(prefix..)
        .ok_or_else(|| unsupported("truncated PowerPoint raster BLIP"))?;
    // Admit only the advertised raster encoding. Some producer output puts a
    // different format in a PNG BLIP; omit it rather than relabel, sniff into
    // another decoder, or copy arbitrary bytes into a supported image part.
    if (extension == "png" && !bytes.starts_with(b"\x89PNG\r\n\x1a\n"))
        || (extension == "jpg" && !bytes.starts_with(&[0xff, 0xd8]))
    {
        return Ok(None);
    }
    let (width, height) = if extension == "png" {
        png_size(bytes)?
    } else {
        jpeg_size(bytes, budget)?
    };
    if width == 0
        || height == 0
        || width > 32768
        || height > 32768
        || u64::from(width) * u64::from(height) > MAX_PIXELS
    {
        return Err(unsupported(
            "PowerPoint image dimensions exceed resource limit",
        ));
    }
    Ok(Some(Image { bytes, extension }))
}
fn png_size(b: &[u8]) -> Result<(u32, u32), String> {
    if b.len() < 33 || !b.starts_with(b"\x89PNG\r\n\x1a\n\0\0\0\x0dIHDR") {
        return Err(unsupported("invalid PowerPoint PNG header"));
    }
    // PNG IHDR (W3C PNG section 11.2.2). This is a bounded header check,
    // not a replacement for the ordinary renderer's image decoder.
    let valid_depth = match b[25] {
        0 => matches!(b[24], 1 | 2 | 4 | 8 | 16),
        2 | 4 | 6 => matches!(b[24], 8 | 16),
        3 => matches!(b[24], 1 | 2 | 4 | 8),
        _ => false,
    };
    if !valid_depth || b[26] != 0 || b[27] != 0 || b[28] > 1 {
        return Err(unsupported("invalid PowerPoint PNG IHDR"));
    }
    Ok((
        u32::from_be_bytes(b[16..20].try_into().unwrap()),
        u32::from_be_bytes(b[20..24].try_into().unwrap()),
    ))
}
fn jpeg_size(b: &[u8], budget: &mut usize) -> Result<(u32, u32), String> {
    if !b.starts_with(&[0xff, 0xd8]) {
        return Err(unsupported("invalid PowerPoint JPEG header"));
    }
    let mut position = 2;
    while position < b.len() {
        *budget = budget
            .checked_sub(1)
            .ok_or_else(|| unsupported("PowerPoint JPEG marker work budget exceeded"))?;
        if b[position] != 0xff {
            return Err(unsupported("invalid PowerPoint JPEG marker"));
        }
        position += 1;
        while b.get(position) == Some(&0xff) {
            *budget = budget
                .checked_sub(1)
                .ok_or_else(|| unsupported("PowerPoint JPEG marker work budget exceeded"))?;
            position += 1;
        }
        let marker = *b
            .get(position)
            .ok_or_else(|| unsupported("truncated PowerPoint JPEG marker"))?;
        position += 1;
        if matches!(marker, 0xda | 0xd9) {
            break;
        }
        if marker == 0x01 {
            // Standalone TEM marker (ITU-T T.81, B.1.1.3).
            continue;
        }
        if matches!(marker, 0 | 0xd0..=0xd8) {
            return Err(unsupported(
                "unexpected PowerPoint JPEG marker before frame",
            ));
        }
        // Do not skip an unsupported first frame and accidentally validate the
        // dimensions of a later frame that a decoder would not choose.
        if matches!(marker, 0xc3 | 0xc5..=0xc7 | 0xc9..=0xcb | 0xcd..=0xcf | 0xde) {
            return Err(unsupported("unsupported PowerPoint JPEG frame encoding"));
        }
        let length = b
            .get(position..position + 2)
            .map(|b| u16::from_be_bytes([b[0], b[1]]) as usize)
            .ok_or_else(|| unsupported("truncated PowerPoint JPEG segment"))?;
        let segment = b
            .get(position..)
            .and_then(|b| b.get(..length))
            .filter(|b| b.len() >= 2)
            .ok_or_else(|| unsupported("invalid PowerPoint JPEG segment length"))?;
        if matches!(marker, 0xc0..=0xc2) {
            // ITU-T T.81 B.2.2: Lf = 8 + 3 * Nf. Only 8-bit Huffman
            // baseline/sequential/progressive frames are supported here.
            if segment.len() < 8
                || segment[2] != 8
                || segment[7] == 0
                || segment.len() != 8 + 3 * usize::from(segment[7])
            {
                return Err(unsupported("unsupported PowerPoint JPEG frame"));
            }
            return Ok((
                u32::from(u16::from_be_bytes([segment[5], segment[6]])),
                u32::from(u16::from_be_bytes([segment[3], segment[4]])),
            ));
        }
        position += length;
    }
    Err(unsupported("PowerPoint JPEG lacks a supported frame"))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(kind: u16, options: u16, payload: &[u8]) -> Vec<u8> {
        [
            options.to_le_bytes().as_slice(),
            kind.to_le_bytes().as_slice(),
            (payload.len() as u32).to_le_bytes().as_slice(),
            payload,
        ]
        .concat()
    }
    fn png(width: u32, height: u32) -> Vec<u8> {
        [
            b"\x89PNG\r\n\x1a\n\0\0\0\x0dIHDR".as_slice(),
            &width.to_be_bytes(),
            &height.to_be_bytes(),
            &[8, 6, 0, 0, 0, 0, 0, 0, 0],
        ]
        .concat()
    }
    fn blip(bytes: &[u8], two_uids: bool) -> Vec<u8> {
        record(
            0xf01e,
            if two_uids { 0x6e10 } else { 0x6e00 },
            &[vec![0; if two_uids { 33 } else { 17 }], bytes.to_vec()].concat(),
        )
    }
    fn bse(size: usize, offset: u32, embedded: &[u8]) -> Vec<u8> {
        let mut payload = vec![0; 36];
        payload[0] = 6;
        payload[20..24].copy_from_slice(&(size as u32).to_le_bytes());
        payload[24..28].copy_from_slice(&1u32.to_le_bytes());
        payload[28..32].copy_from_slice(&offset.to_le_bytes());
        payload.extend_from_slice(embedded);
        record(0xf007, 0x62, &payload)
    }
    fn parsed(bytes: &[u8]) -> Record<'_> {
        parse_record_at(bytes, 0, &mut 1000).unwrap()
    }

    #[test]
    fn reads_embedded_and_delayed_blips_with_one_or_two_uids() {
        for two in [false, true] {
            let png = png(7, 11);
            let blip = blip(&png, two);
            let embedded = bse(blip.len(), u32::MAX, &blip);
            assert_eq!(
                image(parsed(&embedded), &[], &mut 100)
                    .unwrap()
                    .unwrap()
                    .bytes,
                png
            );
            let delayed = [vec![0; 19], blip.clone()].concat();
            let entry = bse(blip.len(), 19, &[]);
            assert_eq!(
                image(parsed(&entry), &delayed, &mut 100)
                    .unwrap()
                    .unwrap()
                    .bytes,
                png
            );
            assert_eq!(
                image(parsed(&blip), &[], &mut 100).unwrap().unwrap().bytes,
                png
            );
        }
    }

    #[test]
    fn deduplicates_references_and_keeps_relationships_local_to_each_slide() {
        let blip = blip(&png(1, 1), false);
        let entries = [parsed(&blip)];
        let mut store = Store::new(&entries, &[]);
        store.remaining = 33;
        assert!(!store.reference(0, &mut 100).unwrap());
        assert!(store.reference(1, &mut 100).unwrap());
        assert!(store.reference(1, &mut 0).unwrap()); // Cached, no reparse/allocation.
        assert_eq!(store.parts().len(), 1);
        assert_eq!(store.remaining, 0);
        assert_eq!(store.relationships().matches("<Relationship ").count(), 1);
        store.begin_slide();
        assert!(store.relationships().is_empty());
        assert_eq!(store.parts().len(), 1);
        assert!(store.reference(1, &mut 0).unwrap());
        assert!(store.reference(2, &mut 100).is_err());
        let mut limited = Store::new(&entries, &[]);
        limited.remaining = 32;
        assert!(limited
            .reference(1, &mut 100)
            .unwrap_err()
            .contains("budget"));
        assert!(limited.parts().is_empty());
    }

    #[test]
    fn rejects_invalid_ranges_sizes_names_and_known_raster_headers() {
        let blip = blip(&png(1, 1), false);
        for entry in [
            bse(blip.len(), u32::MAX, &[]),
            bse(blip.len() - 1, 0, &blip),
            bse(blip.len() + 1, 0, &[]),
        ] {
            assert!(image(parsed(&entry), &blip, &mut 100).is_err());
        }
        let mut entry = bse(blip.len(), 0, &[]);
        entry[8 + 33] = 1;
        assert!(image(parsed(&entry), &blip, &mut 100).is_err());
        entry[8 + 33] = 0;
        entry[8 + 24..8 + 28].fill(0); // Unused slot never dereferences foDelay.
        assert!(image(parsed(&entry), &[], &mut 100).unwrap().is_none());
        let unsupported = record(0xf01a, 0, &[]);
        assert!(image(parsed(&unsupported), &[], &mut 100)
            .unwrap()
            .is_none());
        let mut invalid = blip.clone();
        invalid[0] = 1;
        assert!(image(parsed(&invalid), &[], &mut 100).is_err());
        invalid[0] = 0x20;
        assert!(image(parsed(&invalid), &[], &mut 100).is_err());
    }

    #[test]
    fn caps_dimensions_and_validates_png_ihdr_without_decoding_pixels() {
        assert!(
            image(parsed(&blip(b"GIF87a unsupported", false)), &[], &mut 100)
                .unwrap()
                .is_none()
        );
        for (width, height) in [(0, 1), (32769, 1), (8000, 8000)] {
            assert!(image(parsed(&blip(&png(width, height), false)), &[], &mut 100).is_err());
        }
        let mut header = png(1, 1);
        header[25] = 1;
        assert!(png_size(&header).is_err());
        assert!(png_size(&header[..32]).is_err());
    }

    #[test]
    fn jpeg_frame_lengths_and_marker_work_are_bounded() {
        let jpeg = [
            0xff, 0xd8, 0xff, 0x01, 0xff, 0xc0, 0, 11, 8, 0, 2, 0, 3, 1, 1, 0x11, 0,
        ];
        assert_eq!(jpeg_size(&jpeg, &mut 100).unwrap(), (3, 2));
        assert!(jpeg_size(&jpeg, &mut 1).is_err());
        assert!(jpeg_size(&jpeg[..16], &mut 100).is_err());
        let mut bad = jpeg;
        bad[13] = 2;
        assert!(jpeg_size(&bad, &mut 100).is_err());
        let unsupported_frame = [vec![0xff, 0xd8, 0xff, 0xc3, 0, 2], jpeg[2..].to_vec()].concat();
        assert!(jpeg_size(&unsupported_frame, &mut 100)
            .unwrap_err()
            .contains("encoding"));
        let mut fill = vec![0xff, 0xd8];
        fill.extend(vec![0xff; 100]);
        assert!(jpeg_size(&fill, &mut 10).unwrap_err().contains("budget"));
    }

    #[test]
    fn catalog_validates_counts_and_does_not_scan_unrelated_containers() {
        let blip = blip(&png(1, 1), false);
        let group = record(1035, 15, &record(0xf000, 15, &record(0xf001, 0x1f, &blip)));
        assert_eq!(catalog(&[parsed(&group)], &mut 100).unwrap().len(), 1);
        assert!(catalog(&[parsed(&group), parsed(&group)], &mut 100).is_err());
        let bad = record(1035, 15, &record(0xf000, 15, &record(0xf001, 0x2f, &blip)));
        assert!(catalog(&[parsed(&bad)], &mut 100).is_err());
        assert!(catalog(&[parsed(&blip)], &mut 100).unwrap().is_empty());
    }
}
