//! Passive image BLIPs: MS-PPT 2.1.3/2.4.3; MS-ODRAW 2.2.20-32.
use super::*;
use crate::officeart::raster::Image;
#[cfg(test)]
use crate::officeart::raster::{jpeg_size, png_size};
use std::collections::{BTreeMap, BTreeSet};

// Resource policy: do not pass dimension bombs to the ordinary image decoder.
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
            let image = crate::officeart::raster::read_store_entry(
                entry,
                Some(self.delayed),
                budget,
                self.remaining,
            )?;
            if let Some(image) = &image {
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
            let image = self.images[index]
                .as_ref()
                .expect("only supported referenced images");
            xml.push_str(&format!("<Relationship Id=\"rImg{index}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image{index}.{}\"/>", image.extension));
        }
        xml
    }
    pub fn parts(&self) -> Vec<(String, &[u8])> {
        self.images
            .iter()
            .filter_map(|(id, image)| {
                image.as_ref().map(|image| {
                    (
                        format!("ppt/media/image{id}.{}", image.extension),
                        image.bytes.as_ref(),
                    )
                })
            })
            .collect()
    }
}

#[cfg(test)]
fn image<'a>(
    entry: Record<'a>,
    delayed: &'a [u8],
    budget: &mut usize,
) -> Result<Option<Image<'a>>, String> {
    crate::officeart::raster::read_store_entry(entry, Some(delayed), budget, MAX_MEDIA_BYTES)
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
    fn compressed_emf_is_retained_once_and_reused_without_inflation() {
        let (source, blip) = crate::officeart::emf_test_blip();
        let entries = [parsed(&blip)];
        let mut store = Store::new(&entries, &[]);
        store.remaining = source.len();
        assert!(store.reference(1, &mut 100).unwrap());
        let pointer = store.parts()[0].1.as_ptr();
        assert_eq!(store.parts()[0].1, source);
        assert_eq!(store.remaining, 0);
        store.begin_slide();
        assert!(store.reference(1, &mut 0).unwrap());
        assert_eq!(store.parts()[0].1.as_ptr(), pointer);
        assert!(store.relationships().contains("image1.emf"));
        let mut limited = Store::new(&entries, &[]);
        limited.remaining = source.len() - 1;
        assert!(limited.reference(1, &mut 100).is_err());
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
        let unsupported = record(0xf01b, 0, &[]); // WMF is still not admitted.
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
