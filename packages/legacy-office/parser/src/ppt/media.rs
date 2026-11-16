//! Passive image BLIPs: MS-PPT 2.1.3/2.4.3; MS-ODRAW 2.2.20-32.
use super::*;
use crate::officeart::raster::StoreImageSpan;
#[cfg(test)]
use crate::officeart::raster::{jpeg_size, png_size};
use std::collections::{BTreeMap, BTreeSet};

// Resource policy: do not pass dimension bombs to the ordinary image decoder.
const MAX_MEDIA_BYTES: usize = 128 * 1024 * 1024;

pub(super) fn catalog_spans(
    document: &[u8],
    children: &[RecordSpan],
    budget: &mut usize,
) -> Result<Vec<RecordSpan>, String> {
    catalog_core(children, document, budget)
}

trait CatalogEntry<'a>: Clone {
    fn view(&self, backing: &'a [u8]) -> Result<Record<'a>, String>;
    fn children(&self, backing: &'a [u8], budget: &mut usize) -> Result<Vec<Self>, String>;
}

impl<'a> CatalogEntry<'a> for RecordSpan {
    fn view(&self, backing: &'a [u8]) -> Result<Record<'a>, String> {
        RecordSpan::view(self, backing)
    }

    fn children(&self, backing: &'a [u8], budget: &mut usize) -> Result<Vec<Self>, String> {
        parse_record_spans(backing, self.payload_span(), budget)
    }
}

fn catalog_core<'a, T: CatalogEntry<'a>>(
    children: &[T],
    backing: &'a [u8],
    budget: &mut usize,
) -> Result<Vec<T>, String> {
    let mut result = None;
    let mut group_seen = false;
    for group in children {
        let group_view = group.view(backing)?;
        if group_view.kind != 1035 {
            continue;
        }
        if group_seen || group_view.version != 15 {
            return Err(unsupported("invalid PowerPoint drawing group"));
        }
        group_seen = true;
        let dgg = group.children(backing, budget)?;
        if dgg.len() != 1 {
            return Err(unsupported("invalid PowerPoint OfficeArt drawing group"));
        }
        let dgg_view = dgg[0].view(backing)?;
        if dgg_view.kind != 0xf000 || dgg_view.version != 15 {
            return Err(unsupported("invalid PowerPoint OfficeArt drawing group"));
        }
        for store in dgg[0].children(backing, budget)? {
            let store_view = store.view(backing)?;
            if store_view.kind != 0xf001 {
                continue;
            }
            if result.is_some() || store_view.version != 15 {
                return Err(unsupported("invalid PowerPoint image store"));
            }
            let entries = store.children(backing, budget)?;
            if entries.len() != usize::from(store_view.instance) {
                return Err(unsupported("PowerPoint image store count mismatch"));
            }
            result = Some(entries);
        }
    }
    Ok(result.unwrap_or_default())
}

/// Kind of external OLE object an OLE shape refers to (MS-PPT 2.10.1
/// ExObjListContainer and its ExOleEmbedContainer 2.10.27, ExOleLinkContainer
/// 2.10.29 and ExControlContainer 2.10.10 children). Only the facts that decide
/// whether the shape's stored presentation picture may be shown are kept: the
/// object storage (ExOleObjStg 2.10.34) is never read, inflated or activated.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum OleObject {
    /// ExOleObjAtom.drawAspect is a DataViewAspectEnum ([MS-OSHARED] 2.2.1.2).
    Embedded {
        draw_aspect: u32,
    },
    Linked,
    Control,
}

// RT_ExternalObjectList, RT_ExternalOleEmbed, RT_ExternalOleLink,
// RT_ExternalOleControl and RT_ExternalOleObjectAtom (MS-PPT 2.13.24).
const EX_OBJ_LIST: u16 = 0x0409;
const EX_OLE_EMBED: u16 = 0x0fcc;
const EX_OLE_LINK: u16 = 0x0fce;
const EX_OLE_CONTROL: u16 = 0x0fee;
const EX_OLE_OBJ_ATOM: u16 = 0x0fc3;

// [MS-PPT] 2.10.16-20: the document's external-object list owns hyperlink
// targets; text actions refer to its ExHyperlinkAtom id, never to a URL in the
// OfficeArt textbox itself. Keep this catalog passive: targets are strings and
// are never fetched.
const EX_HYPERLINK: u16 = 0x0fd7;
const EX_HYPERLINK_ATOM: u16 = 0x0fd3;
const CSTRING: u16 = 0x0fba;
const MAX_HYPERLINK_TEXT_BYTES: usize = 1 << 20;

#[derive(Clone, Debug, Default)]
pub(super) struct HyperlinkCatalog {
    targets: BTreeMap<u32, Option<String>>,
    error: Option<String>,
}

impl HyperlinkCatalog {
    pub fn get(&self, id: u32) -> Result<&str, String> {
        if let Some(error) = &self.error {
            return Err(error.clone());
        }
        match self.targets.get(&id) {
            Some(Some(target)) => Ok(target.as_str()),
            Some(None) => Err(unsupported("ambiguous PowerPoint hyperlink reference")),
            None => Err(unsupported("unresolved PowerPoint hyperlink reference")),
        }
    }
}

pub(super) fn hyperlink_catalog(
    document: &[u8],
    children: &[RecordSpan],
    budget: &mut usize,
) -> HyperlinkCatalog {
    let mut catalog = HyperlinkCatalog::default();
    if let Err(error) = read_hyperlink_catalog(document, children, budget, &mut catalog.targets) {
        catalog.targets.clear();
        catalog.error = Some(error);
    }
    catalog
}

fn read_hyperlink_catalog(
    document: &[u8],
    children: &[RecordSpan],
    budget: &mut usize,
    targets: &mut BTreeMap<u32, Option<String>>,
) -> Result<(), String> {
    let mut seen = false;
    let mut remaining = MAX_HYPERLINK_TEXT_BYTES;
    for list in children {
        let view = list.view(document)?;
        if view.kind != EX_OBJ_LIST {
            continue;
        }
        if seen || view.version != 15 {
            return Err(unsupported("invalid PowerPoint external object list"));
        }
        seen = true;
        for container in parse_record_spans(document, list.payload_span(), budget)? {
            let view = container.view(document)?;
            if view.kind != EX_HYPERLINK {
                continue;
            }
            if view.version != 15 {
                return Err(unsupported("invalid PowerPoint hyperlink container"));
            }
            let mut id = None;
            let mut target = None;
            for child in parse_record_spans(document, container.payload_span(), budget)? {
                let view = child.view(document)?;
                match (view.kind, view.instance) {
                    (EX_HYPERLINK_ATOM, 0) if view.version == 0 && view.payload.len() == 4 => {
                        if id.replace(u32_at(view.payload, 0)?).is_some() {
                            return Err(unsupported("duplicate PowerPoint hyperlink id"));
                        }
                    }
                    (CSTRING, 1) if view.version == 0 && view.payload.len() % 2 == 0 => {
                        remaining = remaining.checked_sub(view.payload.len()).ok_or_else(|| {
                            crate::rejection::resource("PowerPoint hyperlink text budget exceeded")
                        })?;
                        let units = view
                            .payload
                            .as_chunks::<2>()
                            .0
                            .iter()
                            .map(|c| u16::from_le_bytes([c[0], c[1]]))
                            .collect::<Vec<_>>();
                        let value = String::from_utf16(&units)
                            .map_err(|_| unsupported("invalid PowerPoint hyperlink target"))?;
                        if target.replace(value).is_some() {
                            return Err(unsupported("duplicate PowerPoint hyperlink target"));
                        }
                    }
                    _ => {}
                }
            }
            let id = id.ok_or_else(|| unsupported("PowerPoint hyperlink lacks id"))?;
            targets
                .entry(id)
                .and_modify(|slot| *slot = None)
                .or_insert(target);
        }
    }
    Ok(())
}

/// exObjId -> OLE object. A malformed list does not reject presentations that
/// never refer to it; the recorded error is returned to the first OLE shape
/// that needs the list. Duplicate identifiers are ambiguous (MS-PPT 2.10.12:
/// each ExOleObjAtom is referred to by exactly one ExObjRefAtom).
#[derive(Clone, Debug, Default)]
pub(super) struct OleCatalog {
    objects: BTreeMap<u32, Option<OleObject>>,
    error: Option<String>,
}

impl OleCatalog {
    pub fn get(&self, id: u32) -> Result<OleObject, String> {
        if let Some(error) = &self.error {
            return Err(error.clone());
        }
        match self.objects.get(&id) {
            Some(Some(object)) => Ok(*object),
            Some(None) => Err(unsupported(
                "ambiguous or inconsistent PowerPoint OLE object reference",
            )),
            None => Err(unsupported("unresolved PowerPoint OLE object reference")),
        }
    }
}

pub(super) fn ole_catalog(
    document: &[u8],
    children: &[RecordSpan],
    budget: &mut usize,
) -> OleCatalog {
    let mut catalog = OleCatalog::default();
    if let Err(error) = read_ole_catalog(document, children, budget, &mut catalog.objects) {
        catalog.objects.clear();
        catalog.error = Some(error);
    }
    catalog
}

fn read_ole_catalog(
    document: &[u8],
    children: &[RecordSpan],
    budget: &mut usize,
    objects: &mut BTreeMap<u32, Option<OleObject>>,
) -> Result<(), String> {
    let mut seen = false;
    for list in children {
        let view = list.view(document)?;
        if view.kind != EX_OBJ_LIST {
            continue;
        }
        if seen || view.version != 15 {
            return Err(unsupported("invalid PowerPoint external object list"));
        }
        seen = true;
        for container in parse_record_spans(document, list.payload_span(), budget)? {
            let view = container.view(document)?;
            let expected_type = match view.kind {
                EX_OLE_EMBED => 0,
                EX_OLE_LINK => 1,
                EX_OLE_CONTROL => 2,
                _ => continue, // Hyperlinks and media are not OLE objects.
            };
            if view.version != 15 {
                return Err(unsupported("invalid PowerPoint OLE object container"));
            }
            let mut atom = None;
            for child in parse_record_spans(document, container.payload_span(), budget)? {
                let child = child.view(document)?;
                if child.kind != EX_OLE_OBJ_ATOM {
                    continue;
                }
                // MS-PPT 2.10.12: recVer 1, recInstance 0, recLen 0x18.
                if atom.is_some()
                    || child.version != 1
                    || child.instance != 0
                    || child.payload.len() != 24
                {
                    return Err(unsupported("invalid PowerPoint ExOleObjAtom"));
                }
                atom = Some((
                    u32_at(child.payload, 0)?,
                    u32_at(child.payload, 4)?,
                    u32_at(child.payload, 8)?,
                ));
            }
            let (draw_aspect, kind, id) =
                atom.ok_or_else(|| unsupported("PowerPoint OLE object lacks ExOleObjAtom"))?;
            let object = (kind == expected_type).then_some(match expected_type {
                0 => OleObject::Embedded { draw_aspect },
                1 => OleObject::Linked,
                _ => OleObject::Control,
            });
            objects
                .entry(id)
                .and_modify(|entry| *entry = None)
                .or_insert(object);
        }
    }
    Ok(())
}

/// A retained image: store index, extension and bytes.
#[cfg(test)]
type RetainedImage<'a> = (u32, &'static str, &'a [u8]);

/// Owned media catalog/cache for direct sessions. The session owns the backing
/// streams separately and supplies them when admitting or reading a
/// resource; cached spans record which stream must be used.
pub(super) struct SpanStore {
    entries: Vec<RecordSpan>,
    images: BTreeMap<u32, Option<StoreImageSpan>>,
    blob_images: BTreeMap<u32, (&'static str, Vec<u8>)>,
    used: BTreeSet<u32>,
    remaining: usize,
}

impl SpanStore {
    pub fn new(entries: Vec<RecordSpan>) -> Self {
        Self {
            entries,
            images: BTreeMap::new(),
            blob_images: BTreeMap::new(),
            used: BTreeSet::new(),
            remaining: MAX_MEDIA_BYTES,
        }
    }

    pub fn begin_slide(&mut self) {
        self.used.clear();
    }

    /// Alternative verification may retain bullet images before a later quota
    /// rejects the shape. Roll back only those new assets and their byte charge;
    /// binary image state and earlier successful alternatives remain owned.
    pub fn blob_checkpoint(&self) -> (usize, usize) {
        (self.blob_images.len(), self.remaining)
    }

    pub fn rollback_blobs(&mut self, (count, remaining): (usize, usize)) {
        while self.blob_images.len() > count {
            self.blob_images.pop_last();
        }
        self.remaining = remaining;
    }

    /// Retain an image from a verified metroBlob OPC relationship. Resource
    /// keys share the direct session's image namespace but cannot collide with
    /// 1-based OfficeArt BSE indices. The session budget bounds total bytes.
    pub fn admit_blob_image(
        &mut self,
        extension: &'static str,
        bytes: Vec<u8>,
    ) -> Result<u32, String> {
        if self.blob_images.len() >= 1024 {
            return Err(crate::rejection::resource(
                "too many PowerPoint blob images",
            ));
        }
        self.remaining = self.remaining.checked_sub(bytes.len()).ok_or_else(|| {
            crate::rejection::resource("PowerPoint retained media budget exceeded")
        })?;
        let index = u32::try_from(self.entries.len() + self.blob_images.len() + 1)
            .map_err(|_| unsupported("PowerPoint image index overflow"))?;
        self.blob_images.insert(index, (extension, bytes));
        Ok(index)
    }

    pub fn reference(
        &mut self,
        index: u32,
        primary: &[u8],
        pictures: Option<&[u8]>,
        budget: &mut usize,
    ) -> Result<bool, String> {
        self.reference_with_policy(index, primary, pictures, budget, false)
    }

    pub fn reference_picture(
        &mut self,
        index: u32,
        primary: &[u8],
        pictures: Option<&[u8]>,
        budget: &mut usize,
    ) -> Result<bool, String> {
        self.reference_with_policy(index, primary, pictures, budget, true)
    }

    fn reference_with_policy(
        &mut self,
        index: u32,
        primary: &[u8],
        pictures: Option<&[u8]>,
        budget: &mut usize,
        pict_placeholder: bool,
    ) -> Result<bool, String> {
        if index == 0 {
            return Ok(false);
        }
        if !self.images.contains_key(&index) {
            let entry = self
                .entries
                .get((index - 1) as usize)
                .ok_or_else(|| unsupported("PowerPoint picture index out of range"))?;
            // PowerPoint displays GIF data stored in PNG BLIPs (its PDF export
            // of a deck with such slots shows the GIF image), so PPT reads a
            // PNG slot by the GIF content signature as well.
            let image = crate::officeart::raster::read_store_entry_span_as(
                entry,
                primary,
                pictures,
                budget,
                self.remaining,
                // Cache the resource independently of its first use. The gate
                // below decides whether that caller may display a PICT frame;
                // a fill must not poison later picture references with None.
                crate::officeart::raster::Raster::PptPictures,
            )?;
            if let Some(image) = &image {
                self.remaining = self
                    .remaining
                    .checked_sub(image.view(primary, pictures)?.len())
                    .ok_or_else(|| {
                        crate::rejection::resource("PowerPoint retained media budget exceeded")
                    })?;
            }
            self.images.insert(index, image);
        }
        let Some(image) = &self.images[&index] else {
            return Ok(false);
        };
        // A picture-frame placeholder is not a tiled/stretch fill capability.
        // Keep this gate even when another picture already cached the resource.
        if matches!(
            image.image.extension,
            "pict" | crate::officeart::raster::UNDECODED_JPEG
        ) && !pict_placeholder
        {
            return Ok(false);
        }
        self.used.insert(index);
        Ok(true)
    }

    #[cfg(test)]
    pub fn used_images(&self) -> impl Iterator<Item = (u32, &StoreImageSpan)> {
        self.used.iter().map(|index| {
            (
                *index,
                self.images[index]
                    .as_ref()
                    .expect("only supported referenced images"),
            )
        })
    }

    /// Fetch an already-admitted resource without scanning the presentation's
    /// cache or decoding an arbitrary unreferenced catalog entry.
    pub fn image<'a>(
        &'a self,
        index: u32,
        primary: &'a [u8],
        pictures: Option<&'a [u8]>,
    ) -> Result<Option<(&'static str, &'a [u8])>, String> {
        if let Some((extension, bytes)) = self.blob_images.get(&index) {
            return Ok(Some((*extension, bytes)));
        }
        // A key that was never admitted is caller/session misuse. A retained
        // None below is an input image outside the passive reader subset.
        self.images
            .get(&index)
            .ok_or_else(|| "PowerPoint image was not admitted".to_string())?
            .as_ref()
            .map(|image| Ok((image.image.extension, image.view(primary, pictures)?)))
            .transpose()
    }

    #[cfg(test)]
    fn images<'a>(
        &'a self,
        primary: &'a [u8],
        pictures: Option<&'a [u8]>,
    ) -> Result<Vec<RetainedImage<'a>>, String> {
        self.images
            .iter()
            .filter_map(|(id, image)| {
                image
                    .as_ref()
                    .map(|image| Ok((*id, image.image.extension, image.view(primary, pictures)?)))
            })
            .collect()
    }
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
    fn spanned(bytes: &[u8], offset: usize) -> RecordSpan {
        record_span_with_end(bytes, offset, &mut 1000, "PowerPoint")
            .unwrap()
            .0
    }

    #[test]
    fn external_hyperlink_id_resolves_to_its_utf16_target() {
        let atom = record(EX_HYPERLINK_ATOM, 0, &7u32.to_le_bytes());
        let url = "https://example.test/path";
        let utf16 = url
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect::<Vec<_>>();
        let target = record(CSTRING, 0x10, &utf16);
        let link = record(EX_HYPERLINK, 0x0f, &[atom, target].concat());
        let document = record(EX_OBJ_LIST, 0x0f, &link);
        let catalog = hyperlink_catalog(&document, &[spanned(&document, 0)], &mut 100);
        assert_eq!(catalog.get(7).unwrap(), url);
        assert!(catalog.get(8).is_err());
    }

    /// Admit the single catalog entry `entry` through the session store, with
    /// `delayed` as the Pictures stream; `None` is an unsupported BLIP.
    fn image(entry: &[u8], delayed: &[u8], budget: &mut usize) -> Result<Option<Vec<u8>>, String> {
        let mut store = SpanStore::new(vec![spanned(entry, 0)]);
        if !store.reference(1, entry, Some(delayed), budget)? {
            return Ok(None);
        }
        Ok(store
            .image(1, entry, Some(delayed))?
            .map(|(_, bytes)| bytes.to_vec()))
    }

    #[test]
    fn pict_picture_admission_is_independent_of_fill_order_and_reuses_retained_bytes() {
        use std::io::Write;
        let source = [0, 16, 0, 0, 0, 0, 0, 1, 0, 1, 0, 0x11, 2, 0xff, 0, 0xff];
        let mut encoder =
            flate2::write::ZlibEncoder::new(Vec::new(), flate2::Compression::default());
        encoder.write_all(&source).unwrap();
        let zipped = encoder.finish().unwrap();
        let mut payload = vec![0; 16 + 34];
        payload[16..20].copy_from_slice(&(source.len() as u32).to_le_bytes());
        payload[44..48].copy_from_slice(&(zipped.len() as u32).to_le_bytes());
        payload[49] = 0xfe;
        payload.extend_from_slice(&zipped);
        let blip = record(0xf01c, 0x5420, &payload);
        let mut entry = bse(blip.len(), 0, &[]);
        entry[0] = 0x42;
        entry[8] = 4;
        let mislabeled = bse(blip.len(), 0, &[]);
        let mut wrong_slot = SpanStore::new(vec![spanned(&mislabeled, 0)]);
        assert!(wrong_slot
            .reference_picture(1, &mislabeled, Some(&blip), &mut 100)
            .unwrap_err()
            .contains("store type mismatch"));
        for fill_first in [false, true] {
            let mut store = SpanStore::new(vec![spanned(&entry, 0)]);
            store.remaining = source.len();
            if fill_first {
                assert!(!store.reference(1, &entry, Some(&blip), &mut 100).unwrap());
                assert_eq!(store.used_images().count(), 0);
            }
            assert!(store
                .reference_picture(1, &entry, Some(&blip), &mut 100)
                .unwrap());
            assert_eq!(
                store.image(1, &entry, None).unwrap(),
                Some(("pict", source.as_slice()))
            );
            assert!(!store.reference(1, &entry, None, &mut 0).unwrap());
            store.begin_slide();
            // An exhausted media/work budget still admits a cached picture on
            // another slide, without inflating or charging the resource twice.
            assert!(store.reference_picture(1, &entry, None, &mut 0).unwrap());
            assert_eq!(store.used_images().count(), 1);
            assert_eq!(
                store.image(1, &entry, None).unwrap(),
                Some(("pict", source.as_slice()))
            );
        }
    }

    #[test]
    fn reads_embedded_and_delayed_blips_with_one_or_two_uids() {
        for two in [false, true] {
            let png = png(7, 11);
            let blip = blip(&png, two);
            let embedded = bse(blip.len(), u32::MAX, &blip);
            assert_eq!(image(&embedded, &[], &mut 100).unwrap().unwrap(), png);
            let delayed = [vec![0; 19], blip.clone()].concat();
            let entry = bse(blip.len(), 19, &[]);
            assert_eq!(image(&entry, &delayed, &mut 100).unwrap().unwrap(), png);
            assert_eq!(image(&blip, &[], &mut 100).unwrap().unwrap(), png);
        }
    }

    #[test]
    fn rejects_invalid_ranges_sizes_names_and_known_raster_headers() {
        let blip = blip(&png(1, 1), false);
        for entry in [
            bse(blip.len(), u32::MAX, &[]),
            bse(blip.len() - 1, 0, &blip),
            bse(blip.len() + 1, 0, &[]),
        ] {
            assert!(image(&entry, &blip, &mut 100).is_err());
        }
        let mut entry = bse(blip.len(), 0, &[]);
        entry[8 + 33] = 1;
        assert!(image(&entry, &blip, &mut 100).is_err());
        entry[8 + 33] = 0;
        entry[8 + 24..8 + 28].fill(0); // Unused slot never dereferences foDelay.
        assert!(image(&entry, &[], &mut 100).unwrap().is_none());
        let unsupported = record(0xf01f, 0, &[]); // DIB is still not admitted.
        assert!(image(&unsupported, &[], &mut 100).unwrap().is_none());
        let malformed_pict = record(0xf01c, 0, &[]);
        assert!(image(&malformed_pict, &[], &mut 100).is_err());
        let mut invalid = blip.clone();
        invalid[0] = 1;
        assert!(image(&invalid, &[], &mut 100).is_err());
        invalid[0] = 0x20;
        assert!(image(&invalid, &[], &mut 100).is_err());
    }

    #[test]
    fn caps_dimensions_and_validates_png_ihdr_without_decoding_pixels() {
        // A PNG slot holding neither PNG nor GIF data is an unsupported BLIP.
        assert!(image(&blip(b"BM unsupported", false), &[], &mut 100)
            .unwrap()
            .is_none());
        for (width, height) in [(0, 1), (32769, 1), (8000, 8000)] {
            assert!(image(&blip(&png(width, height), false), &[], &mut 100).is_err());
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
        // A well-formed frame this reader does not decode (SOF3 lossless,
        // SOF9 arithmetic, 12-bit samples) is rejected by the size reader and
        // kept as an undecoded picture by the PowerPoint store.
        for (marker, precision, expected) in [
            (0xc3, 8, "encoding"),
            (0xc9, 8, "encoding"),
            (0xc1, 12, "JPEG frame"),
        ] {
            let mut frame = jpeg;
            frame[5] = marker;
            frame[8] = precision;
            assert!(jpeg_size(&frame, &mut 100).unwrap_err().contains(expected));
            let entry = record(0xf01d, 0x46a0, &[vec![0; 17], frame.to_vec()].concat());
            let mut store = SpanStore::new(vec![spanned(&entry, 0)]);
            // Not a fill; a picture frame keeps it as an undecoded picture.
            assert!(!store.reference(1, &entry, None, &mut 100).unwrap());
            assert!(store.reference_picture(1, &entry, None, &mut 100).unwrap());
            assert_eq!(
                store.image(1, &entry, None).unwrap().unwrap().0,
                crate::officeart::raster::UNDECODED_JPEG
            );
        }
        // A frame header whose length disagrees with its component count is
        // not a frame of any encoding.
        let truncated = [vec![0xff, 0xd8, 0xff, 0xc3, 0, 2], jpeg[2..].to_vec()].concat();
        let entry = record(0xf01d, 0x46a0, &[vec![0; 17], truncated].concat());
        assert!(image(&entry, &[], &mut 100).is_err());
        let mut fill = vec![0xff, 0xd8];
        fill.extend(vec![0xff; 100]);
        assert!(jpeg_size(&fill, &mut 10).unwrap_err().contains("budget"));
    }

    #[test]
    fn catalog_validates_counts_and_does_not_scan_unrelated_containers() {
        let blip = blip(&png(1, 1), false);
        let group = record(1035, 15, &record(0xf000, 15, &record(0xf001, 0x1f, &blip)));
        let mut document = vec![0xaa; 7];
        document.extend_from_slice(&group);
        let group_span = spanned(&document, 7);
        let entries = catalog_spans(&document, &[group_span], &mut 100).unwrap();
        assert_eq!(entries.len(), 1);
        assert_eq!(entries[0].view(&document).unwrap().payload, &blip[8..]);
        let catalog = |children: &[&[u8]]| {
            let document = children.concat();
            let mut offset = 0;
            let spans = children
                .iter()
                .map(|child| {
                    let span = spanned(&document, offset);
                    offset += child.len();
                    span
                })
                .collect::<Vec<_>>();
            catalog_spans(&document, &spans, &mut 100).map(|entries| entries.len())
        };
        assert!(catalog(&[&group, &group]).is_err());
        let bad = record(1035, 15, &record(0xf000, 15, &record(0xf001, 0x2f, &blip)));
        assert!(catalog(&[&bad]).is_err());
        assert_eq!(catalog(&[&blip]).unwrap(), 0);
    }

    #[test]
    fn span_store_decodes_once_and_keeps_cached_parts_across_slides() {
        let first = blip(&png(1, 1), false);
        let second = blip(&png(2, 2), false);
        let mut primary = vec![0xaa, 0xbb, 0xcc];
        let first_offset = primary.len();
        primary.extend_from_slice(&first);
        let second_offset = primary.len();
        primary.extend_from_slice(&second);
        let entries = vec![
            spanned(&primary, first_offset),
            spanned(&primary, second_offset),
        ];
        let mut store = SpanStore::new(entries);
        store.remaining = 66;

        assert!(store.image(0, &primary, None).is_err());
        assert!(store.image(1, &primary, None).is_err());
        assert!(!store.reference(0, &primary, None, &mut 100).unwrap());
        assert!(store
            .reference(3, &primary, None, &mut 100)
            .unwrap_err()
            .contains("index out of range"));

        assert!(store.reference(1, &primary, None, &mut 100).unwrap());
        assert!(store.reference(1, &primary, None, &mut 0).unwrap());
        assert_eq!(store.remaining, 33);
        assert_eq!(store.images(&primary, None).unwrap().len(), 1);
        store.begin_slide();
        assert_eq!(store.used_images().count(), 0);
        assert_eq!(
            store.image(1, &primary, None).unwrap().unwrap(),
            ("png", png(1, 1).as_slice())
        );
        assert!(store.image(2, &primary, None).is_err());
        assert_eq!(store.images(&primary, None).unwrap().len(), 1);
        assert!(store.reference(2, &primary, None, &mut 100).unwrap());
        assert_eq!(store.remaining, 0);
        assert_eq!(
            store.used_images().map(|(id, _)| id).collect::<Vec<_>>(),
            [2]
        );
        assert_eq!(store.images(&primary, None).unwrap().len(), 2);
        assert_eq!(
            store.image(2, &primary, None).unwrap().unwrap(),
            ("png", png(2, 2).as_slice())
        );
    }

    #[test]
    fn span_store_retains_delayed_inflation_and_enforces_media_limits() {
        for (source, blip) in [
            crate::officeart::emf_test_blip(),
            crate::officeart::wmf_test_blip(),
        ] {
            span_store_retains_delayed_metafile(source, blip);
        }
    }

    fn span_store_retains_delayed_metafile(source: Vec<u8>, blip: Vec<u8>) {
        let pictures_offset = 17;
        let entry_bytes = bse(blip.len(), pictures_offset, &[]);
        let mut primary = vec![0x55; 9];
        let entry_offset = primary.len();
        primary.extend_from_slice(&entry_bytes);
        let mut pictures = vec![0x77; pictures_offset as usize];
        pictures.extend_from_slice(&blip);
        let entry = spanned(&primary, entry_offset);

        let mut store = SpanStore::new(vec![entry.clone()]);
        store.remaining = source.len();
        assert!(store
            .reference(1, &primary, Some(&pictures), &mut 100)
            .unwrap());
        let pointer = store.images(&primary, Some(&pictures)).unwrap()[0]
            .2
            .as_ptr();
        assert_eq!(
            store.images(&primary, Some(&pictures)).unwrap()[0].2,
            source
        );
        store.begin_slide();
        assert!(store
            .reference(1, &primary, Some(&pictures), &mut 0)
            .unwrap());
        assert_eq!(
            store.images(&primary, Some(&pictures)).unwrap()[0]
                .2
                .as_ptr(),
            pointer
        );
        // Inflated bytes are owned by the cache and no longer depend on the
        // Pictures backing after the first decode.
        assert_eq!(store.images(&primary, None).unwrap()[0].2, source);

        let mut limited = SpanStore::new(vec![entry]);
        limited.remaining = source.len() - 1;
        assert!(limited
            .reference(1, &primary, Some(&pictures), &mut 100)
            .unwrap_err()
            .contains("budget"));
        assert!(limited
            .images(&primary, Some(&pictures))
            .unwrap()
            .is_empty());
    }
}
