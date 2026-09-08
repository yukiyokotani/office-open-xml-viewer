//! Owned direct-model projection of validated passive inline DOC pictures.
//! MS-DOC 2.9.192-193; MS-ODRAW 2.2.15 and 2.2.40.

use super::{unsupported, Store};

#[derive(Debug, PartialEq)]
pub(in crate::doc) struct DirectInlinePicture {
    pub resource_key: String,
    pub mime_type: &'static str,
    pub width_pt: f64,
    pub height_pt: f64,
    pub crop: Option<ooxml_common::blip::SrcRect>,
    pub rotation: f64,
    pub flip_h: bool,
    pub flip_v: bool,
}

#[derive(Debug, PartialEq, Eq)]
pub(crate) struct DirectPictureResource {
    pub key: String,
    pub mime_type: &'static str,
    pub bytes: Vec<u8>,
}

impl Store<'_> {
    /// Select an occurrence for one direct document result. Direct production
    /// must not call the OOXML part-scoping `begin_part`, which clears the
    /// selected-offset set used by `finish_direct_resources`.
    pub(in crate::doc) fn direct_inline(
        &mut self,
        offset: usize,
        remaining_bytes: &mut usize,
    ) -> Result<Option<DirectInlinePicture>, String> {
        self.load(offset)?;
        let Some(picture) = self.cache[&offset].as_ref() else {
            self.omitted = true;
            return Ok(None);
        };
        let mime_type = mime(picture.image.extension)?;
        let resource_key = key(offset);
        let required = std::mem::size_of::<DirectInlinePicture>()
            .checked_add(resource_key.capacity())
            .ok_or("OUTPUT_TOO_LARGE")?;
        *remaining_bytes = remaining_bytes
            .checked_sub(required)
            .ok_or("OUTPUT_TOO_LARGE")?;
        if self.occurrences >= 1_000_000 {
            return Err(unsupported("Word picture occurrence budget exceeded"));
        }
        self.occurrences += 1;
        self.part_offsets.insert(offset);
        let [top, bottom, left, right] = picture.crop;
        let crop = (picture.crop != [0; 4]).then_some(ooxml_common::blip::SrcRect {
            l: left as f64 / 100_000.0,
            t: top as f64 / 100_000.0,
            r: right as f64 / 100_000.0,
            b: bottom as f64 / 100_000.0,
        });
        Ok(Some(DirectInlinePicture {
            resource_key,
            mime_type,
            width_pt: picture.extent[0] as f64 / 12_700.0,
            height_pt: picture.extent[1] as f64 / 12_700.0,
            crop,
            rotation: picture.rotation as f64 / 60_000.0,
            flip_h: picture.flip[0],
            flip_v: picture.flip[1],
        }))
    }

    pub(in crate::doc) fn finish_direct_resources(
        self,
        remaining_bytes: &mut usize,
    ) -> Result<Vec<DirectPictureResource>, String> {
        let selected_count = self
            .cache
            .iter()
            .filter(|(offset, picture)| self.part_offsets.contains(offset) && picture.is_some())
            .count();
        let mut required = selected_count
            .checked_mul(std::mem::size_of::<DirectPictureResource>())
            .ok_or("OUTPUT_TOO_LARGE")?;
        for (offset, picture) in &self.cache {
            if !self.part_offsets.contains(offset) || picture.is_none() {
                continue;
            }
            let picture = picture.as_ref().expect("filtered selected picture");
            mime(picture.image.extension)?;
            let image_bytes = match &picture.image.bytes {
                std::borrow::Cow::Borrowed(bytes) => bytes.len(),
                std::borrow::Cow::Owned(bytes) => bytes.capacity(),
            };
            required = required
                .checked_add(key(*offset).capacity())
                .and_then(|bytes| bytes.checked_add(image_bytes))
                .ok_or("OUTPUT_TOO_LARGE")?;
        }
        *remaining_bytes = remaining_bytes
            .checked_sub(required)
            .ok_or("OUTPUT_TOO_LARGE")?;
        let mut resources = Vec::new();
        resources
            .try_reserve_exact(selected_count)
            .map_err(|_| "OUTPUT_TOO_LARGE".to_string())?;
        let excess = resources
            .capacity()
            .saturating_sub(selected_count)
            .checked_mul(std::mem::size_of::<DirectPictureResource>())
            .ok_or("OUTPUT_TOO_LARGE")?;
        *remaining_bytes = remaining_bytes
            .checked_sub(excess)
            .ok_or("OUTPUT_TOO_LARGE")?;
        for (offset, picture) in self.cache {
            if !self.part_offsets.contains(&offset) {
                continue;
            }
            let picture = picture.expect("selected direct picture was validated");
            let mime_type = mime(picture.image.extension)?;
            let key = key(offset);
            let bytes = match picture.image.bytes {
                std::borrow::Cow::Owned(bytes) => bytes,
                std::borrow::Cow::Borrowed(source) => {
                    let mut bytes = Vec::new();
                    bytes
                        .try_reserve_exact(source.len())
                        .map_err(|_| "OUTPUT_TOO_LARGE".to_string())?;
                    *remaining_bytes = remaining_bytes
                        .checked_sub(bytes.capacity().saturating_sub(source.len()))
                        .ok_or("OUTPUT_TOO_LARGE")?;
                    bytes.extend_from_slice(source);
                    bytes
                }
            };
            resources.push(DirectPictureResource {
                key,
                mime_type,
                bytes,
            });
        }
        Ok(resources)
    }
}

fn key(offset: usize) -> String {
    format!("legacy-doc/image/{offset}")
}

fn mime(extension: &str) -> Result<&'static str, String> {
    match extension {
        "png" | "jpg" => Ok(ooxml_common::blip::mime_from_ext(extension)),
        _ => Err(unsupported(
            "direct DOC model supports only PNG/JPEG inline pictures",
        )),
    }
}
