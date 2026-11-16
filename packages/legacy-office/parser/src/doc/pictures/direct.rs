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

/// Occurrence geometry is distinct from the shared carrier: the display box
/// comes from `picture_bullet_display_box`, never from the stored extent.
#[derive(Debug, PartialEq)]
pub(in crate::doc) struct DirectPictureBullet {
    pub picture: DirectInlinePicture,
    pub no_auto_size: bool,
}

impl DirectPictureBullet {
    /// Attach a resolved occurrence to the shared model without modifying the
    /// numbering cascade's font facts. Layout and paint then use the same box.
    pub(in crate::doc) fn install(
        self,
        numbering: &mut docx_model::NumberingInfo,
        remaining_bytes: &mut usize,
    ) -> Result<(), String> {
        if numbering.pic_bullet_image_path.is_some()
            || numbering.pic_bullet_mime_type.is_some()
            || numbering.pic_bullet_width_pt.is_some()
            || numbering.pic_bullet_height_pt.is_some()
            || numbering.pic_bullet_transform.is_some()
        {
            return Err(unsupported(
                "Word picture bullet model is already populated",
            ));
        }
        let mime_type = self.picture.mime_type.to_string();
        // Check capacity before handing over the retained strings, but leave
        // their single charge to payload::paragraph_metadata alongside the
        // NumberingInfo allocation. The projection is a temporary stack value.
        remaining_bytes
            .checked_sub(self.picture.resource_key.capacity())
            .and_then(|remaining| remaining.checked_sub(mime_type.capacity()))
            .ok_or("OUTPUT_TOO_LARGE")?;
        numbering.pic_bullet_image_path = Some(self.picture.resource_key);
        numbering.pic_bullet_mime_type = Some(mime_type);
        numbering.pic_bullet_width_pt = Some(self.picture.width_pt);
        numbering.pic_bullet_height_pt = Some(self.picture.height_pt);
        numbering.pic_bullet_transform = Some(docx_model::PictureBulletTransform {
            src_rect: self.picture.crop,
            rotation: self.picture.rotation,
            flip_h: self.picture.flip_h,
            flip_v: self.picture.flip_v,
        });
        Ok(())
    }
}

#[derive(Debug, PartialEq, Eq)]
pub(crate) struct DirectPictureResource {
    pub key: String,
    pub mime_type: &'static str,
    pub bytes: Vec<u8>,
}

impl Store<'_> {
    pub(in crate::doc) fn has_selected_direct_resources(&self) -> bool {
        !self.part_offsets.is_empty() || !self.bullet_offsets.is_empty()
    }

    /// The display box of a picture bullet, in points.
    ///
    /// Observed Windows Word (16.0) behavior, measured on Word's PDFs of
    /// same-length patches of one Word-saved DOC (zero two-export noise):
    /// - fNoAutoSize=1 displays the BLIP at PICF dxaGoal x dyaGoal (MS-DOC
    ///   2.9.193 PICMID). PICMID mx/my scaling does not apply: changing mx
    ///   from 150% to 300% leaves the PDF pixel-identical in both modes.
    /// - An OfficeArt crop (MS-ODRAW 2.3.23 cropFromTop 25%) does not apply
    ///   to a picture bullet either: pixel-identical in both modes. The
    ///   projection therefore carries no crop.
    /// - Automatic size (fNoAutoSize=0) follows the marker's own character
    ///   size (the list level's CHPX over the paragraph mark), not the
    ///   following text: changing only the following text from 12 pt to 8 or
    ///   24 pt leaves the bullet unchanged. A complex-script or right-to-left
    ///   marker (sprmCFComplexScripts / sprmCFBiDi) uses its sprmCHpsBi size,
    ///   or 10 pt with none (`picture_bullet_marker_size_pt`), and bold or
    ///   italic markers are unaffected. The displayed height is
    ///   the marker size minus 2 pt, never less than 2 pt, and the width keeps
    ///   the goal aspect ratio; the goal size itself does not matter. Each
    ///   dimension is limited to 32767 twips (1638.35 pt), which only a very
    ///   wide picture at a very large marker reaches.
    ///   Measured marker sizes and displayed heights (to Word's 0.05 pt output
    ///   rounding): 1, 2, 3 and 4 pt -> 2 pt; 5 -> 3; 6 -> 4; 7.5 -> 5.5;
    ///   8 -> 6; 9 -> 7; 10 -> 8; 10.5 -> 8.5; 11 -> 9; 11.5 -> 9.5; 12 -> 10;
    ///   13.5 -> 11.5; 14 -> 12; 16 -> 14; 18 -> 16; 20 -> 18; 24 -> 22;
    ///   28 -> 26; 36 -> 34; 48 -> 46; 72 -> 70; 96 -> 94; 144 -> 142;
    ///   288 -> 286; 720 -> 718; 1638 (the largest) -> 1636. Goals 12x6, 6x12
    ///   and 12x12 pt at a 12 pt marker display 20x10, 5x10 and 10x10 pt; a
    ///   12x6 goal displays 4x2 pt at 1 and 3 pt markers, 284x142 at 144 pt
    ///   and 1638.35x1636 at 1638 pt; a 6x12 goal displays 1x2 at 3 pt and
    ///   359x718 at 720 pt. This is an observed Office rule, not an MS-DOC
    ///   one.
    pub(in crate::doc) fn picture_bullet_display_box(
        &mut self,
        formatting: &mut super::super::formatting::Formatting<'_>,
        bullet: super::super::character::EnabledPictureBullet,
        marker_font_size_pt: Option<f64>,
    ) -> Result<Option<[f64; 2]>, String> {
        /// 32767 twips.
        const MAX_DIMENSION_PT: f64 = 1638.35;
        let picture = self.acquire_picture_bullet(formatting, bullet)?;
        let [width, height] = picture.goal;
        if width <= 0 || height <= 0 {
            return Err(unsupported("invalid Word picture bullet goal size"));
        }
        let [width, height] = [f64::from(width) / 20.0, f64::from(height) / 20.0];
        if bullet.flags.no_auto_size {
            return Ok(Some([width, height]));
        }
        Ok(marker_font_size_pt.map(|size| {
            let displayed = (size - 2.0).clamp(2.0, MAX_DIMENSION_PT);
            [
                (displayed * width / height).min(MAX_DIMENSION_PT),
                displayed,
            ]
        }))
    }

    /// Project an acquired carrier using the occurrence box from
    /// `picture_bullet_display_box`. Observed Windows Word (16.0) behavior
    /// for the picture's own OfficeArt transform, by pixel-identical PDFs of
    /// patched controls: neither a crop (cropFromTop 25%) nor a rotation
    /// (90 degrees) applies to a picture bullet, in either sizing mode, while
    /// the FSP fFlipH and fFlipV flags mirror the bullet in its unchanged box.
    pub(in crate::doc) fn direct_picture_bullet(
        &mut self,
        formatting: &mut super::super::formatting::Formatting<'_>,
        bullet: super::super::character::EnabledPictureBullet,
        display_box: [f64; 2],
        remaining_bytes: &mut usize,
    ) -> Result<DirectPictureBullet, String> {
        if !display_box
            .iter()
            .all(|value| value.is_finite() && *value > 0.0)
        {
            return Err(unsupported("invalid Word picture bullet display box"));
        }
        if self.occurrences >= 1_000_000 {
            return Err(crate::rejection::resource(
                "Word picture occurrence budget exceeded",
            ));
        }
        let picture = self.acquire_picture_bullet(formatting, bullet)?;
        let offset = picture.offset;
        let resource_key = bullet_key(offset);
        let mime_type = mime(picture.image.extension)?;
        let required = resource_key
            .capacity()
            .checked_add(mime_type.len())
            .ok_or("OUTPUT_TOO_LARGE")?;
        remaining_bytes
            .checked_sub(required)
            .ok_or("OUTPUT_TOO_LARGE")?;
        let result = DirectPictureBullet {
            picture: DirectInlinePicture {
                resource_key,
                mime_type,
                width_pt: display_box[0],
                height_pt: display_box[1],
                crop: None,
                rotation: 0.0,
                flip_h: picture.flip[0],
                flip_v: picture.flip[1],
            },
            no_auto_size: bullet.flags.no_auto_size,
        };
        // The shared paragraph payload owns and charges retained key/MIME
        // strings. Do not charge them a second time during projection.
        self.occurrences += 1;
        self.bullet_offsets.insert(offset);
        Ok(result)
    }

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
            // Word's own DOCX of the corpus documents writes nothing for the
            // placeholder of a pseudo-inline shape; the shape itself is
            // projected at its anchor character.
            if !self.placeholders.contains(&offset) {
                self.omitted
                    .defer("Word inline picture frame without a picture");
            }
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
            return Err(crate::rejection::resource(
                "Word picture occurrence budget exceeded",
            ));
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

    pub(in crate::doc) fn finish_referenced_direct_resources(
        mut self,
        references: &[&str],
        remaining_bytes: &mut usize,
    ) -> Result<Vec<DirectPictureResource>, String> {
        for reference in references {
            if let Some(suffix) = reference.strip_prefix("legacy-doc/bullet/") {
                let offset = suffix
                    .parse::<usize>()
                    .map_err(|_| unsupported("invalid direct DOC picture bullet resource key"))?;
                if bullet_key(offset) != *reference
                    || !self.bullet_offsets.contains(&offset)
                    || !self.bullets.contains_key(&offset)
                {
                    return Err(unsupported("dangling direct DOC picture bullet resource"));
                }
                continue;
            }
            let Some(suffix) = reference.strip_prefix("legacy-doc/image/") else {
                continue;
            };
            let offset = suffix
                .parse::<usize>()
                .map_err(|_| unsupported("invalid direct DOC inline picture resource key"))?;
            if key(offset) != *reference
                || !self.part_offsets.contains(&offset)
                || self.cache.get(&offset).and_then(Option::as_ref).is_none()
            {
                return Err(unsupported("dangling direct DOC inline picture resource"));
            }
        }
        let bullets = std::mem::take(&mut self.bullets);
        let mut resources = self.finish_direct_resources_with(
            |offset| {
                let value = key(offset);
                references.binary_search(&value.as_str()).is_ok()
            },
            remaining_bytes,
        )?;
        for (offset, picture) in bullets {
            let key = bullet_key(offset);
            if references.binary_search(&key.as_str()).is_err() {
                continue;
            }
            let mime_type = mime(picture.image.extension)?;
            let image_bytes = match &picture.image.bytes {
                std::borrow::Cow::Borrowed(bytes) => bytes.len(),
                std::borrow::Cow::Owned(bytes) => bytes.capacity(),
            };
            let required = key
                .capacity()
                .checked_add(image_bytes)
                .ok_or("OUTPUT_TOO_LARGE")?;
            *remaining_bytes = remaining_bytes
                .checked_sub(required)
                .ok_or("OUTPUT_TOO_LARGE")?;
            let capacity = resources.capacity();
            resources
                .try_reserve_exact(1)
                .map_err(|_| "OUTPUT_TOO_LARGE".to_string())?;
            let excess = resources
                .capacity()
                .saturating_sub(capacity)
                .checked_mul(std::mem::size_of::<DirectPictureResource>())
                .ok_or("OUTPUT_TOO_LARGE")?;
            *remaining_bytes = remaining_bytes
                .checked_sub(excess)
                .ok_or("OUTPUT_TOO_LARGE")?;
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

    #[cfg(test)]
    pub(in crate::doc) fn finish_direct_resources(
        self,
        remaining_bytes: &mut usize,
    ) -> Result<Vec<DirectPictureResource>, String> {
        let selected = self.part_offsets.clone();
        self.finish_direct_resources_with(|offset| selected.contains(&offset), remaining_bytes)
    }

    fn finish_direct_resources_with(
        self,
        selected: impl Fn(usize) -> bool,
        remaining_bytes: &mut usize,
    ) -> Result<Vec<DirectPictureResource>, String> {
        let selected_count = self
            .cache
            .iter()
            .filter(|(offset, picture)| selected(**offset) && picture.is_some())
            .count();
        let mut required = selected_count
            .checked_mul(std::mem::size_of::<DirectPictureResource>())
            .ok_or("OUTPUT_TOO_LARGE")?;
        for (offset, picture) in &self.cache {
            if !selected(*offset) || picture.is_none() {
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
            if !selected(offset) {
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

fn bullet_key(offset: usize) -> String {
    format!("legacy-doc/bullet/{offset}")
}

/// The passive BLIP reader admits PNG/JPEG rasters and validated EMF/WMF
/// metafiles (MS-ODRAW 2.2.24-25/31). They use the same extension-to-MIME
/// mapping as DOCX media parts (`image/emf`, `image/wmf`), so the shared
/// content-sniffing metafile players render them; no DOC-specific paint path.
fn mime(extension: &str) -> Result<&'static str, String> {
    match extension {
        "png" | "jpg" | "gif" | "emf" | "wmf" | "tiff" => {
            Ok(ooxml_common::blip::mime_from_ext(extension))
        }
        _ => Err(unsupported(
            "direct DOC model supports only PNG/JPEG/GIF/TIFF/EMF/WMF inline pictures",
        )),
    }
}
