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

/// Occurrence geometry is distinct from the shared carrier. PbiGrf chooses
/// whether a consumer resizes to following text; it does not specify that
/// metric. Neither AUTO nor FIXED is inferred from PICMID goal/scale here.
#[derive(Debug, PartialEq)]
#[allow(
    dead_code,
    reason = "native marker sizing consumer remains unsupported"
)]
pub(in crate::doc) struct DirectPictureBullet {
    pub picture: DirectInlinePicture,
    pub no_auto_size: bool,
}

impl DirectPictureBullet {
    /// Attach a resolved occurrence to the shared model without modifying the
    /// numbering cascade's font facts. Layout and paint then use the same box.
    #[allow(
        dead_code,
        reason = "native marker sizing consumer remains unsupported"
    )]
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

    /// Project an acquired carrier using an explicitly resolved occurrence
    /// box. This boundary preserves OfficeArt transforms and selects media;
    /// it cannot establish the native AUTO metric or a FIXED display size.
    /// The story consumer must retain its atomic refusal until it can supply
    /// a source-justified box, rather than passing stored or nominal-font size.
    #[allow(
        dead_code,
        reason = "native marker sizing consumer remains unsupported"
    )]
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
            return Err(unsupported("Word picture occurrence budget exceeded"));
        }
        let picture = self.acquire_picture_bullet(formatting, bullet)?;
        if picture.crop[0] + picture.crop[1] >= 100_000
            || picture.crop[2] + picture.crop[3] >= 100_000
        {
            return Err(unsupported("empty Word picture bullet crop"));
        }
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
        let [top, bottom, left, right] = picture.crop;
        let crop = (picture.crop != [0; 4]).then_some(ooxml_common::blip::SrcRect {
            l: left as f64 / 100_000.0,
            t: top as f64 / 100_000.0,
            r: right as f64 / 100_000.0,
            b: bottom as f64 / 100_000.0,
        });
        let result = DirectPictureBullet {
            picture: DirectInlinePicture {
                resource_key,
                mime_type,
                width_pt: display_box[0],
                height_pt: display_box[1],
                crop,
                rotation: picture.rotation as f64 / 60_000.0,
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

    /// Explicit library reading policy, not MS-DOC 2.9.176 AUTO sizing.
    /// MS-DOC 2.9.193 supplies stored PICMID extents; using those as the
    /// display box is disclosed to the reader. Source font facts are untouched.
    pub(in crate::doc) fn direct_reading_picture_bullet(
        &mut self,
        formatting: &mut super::super::formatting::Formatting<'_>,
        marker: &super::super::character::PictureBullet,
        remaining_bytes: &mut usize,
    ) -> Result<
        (
            DirectPictureBullet,
            Box<docx_model::NativeReadingPictureBullet>,
        ),
        String,
    > {
        let bullet = marker
            .enabled()?
            .ok_or_else(|| unsupported("disabled reading picture bullet"))?;
        let (raw_pbi_flags, flags_origin, index_origin) = marker.reading_source_owners()?;
        let picture = self.acquire_picture_bullet(formatting, bullet)?;
        if picture.shape != Some(75)
            || picture.client_anchor_count > 1
            || picture.pib_flags.is_some_and(|value| value.key != 0x0106)
            || picture.malformed_pib_name
            || picture.malformed_pib_flags
        {
            return Err(unsupported(
                "Word reading picture bullet has ambiguous source geometry",
            ));
        }
        if let Some(flags) = picture.pib_flags {
            // MS-ODRAW 2.4.8: Comment/File/URL are exclusive; DoNotSave
            // requires LinkToFile, which itself requires File or URL.
            let value = flags.value;
            if value & !0x0f != 0
                || value & 3 == 3
                || (value & 4 != 0 && value & 8 == 0)
                || (value & 8 != 0 && value & 3 == 0)
            {
                return Err(unsupported("invalid Word picture bullet MSOBLIPFLAGS"));
            }
            // This bounded consumer owns embedded passive data. Named or
            // linked display semantics need a separate source/name consumer.
            // Omitted pibFlags uses the normative Comment default; retain
            // the absence rather than synthesizing an encoded zero operand.
            if value != 0 {
                return Err(unsupported(
                    "named or linked Word picture bullet has no stored-size reading consumer",
                ));
            }
        }
        let needed = std::mem::size_of::<docx_model::NativeReadingPictureBullet>()
            .checked_add(picture.client_anchor.map_or(0, <[u8]>::len))
            .ok_or("OUTPUT_TOO_LARGE")?;
        remaining_bytes
            .checked_sub(needed)
            .ok_or("OUTPUT_TOO_LARGE")?;
        let client_anchor = picture
            .client_anchor
            .map(|source| {
                let mut bytes = Vec::new();
                bytes
                    .try_reserve_exact(source.len())
                    .map_err(|_| "OUTPUT_TOO_LARGE")?;
                bytes.extend_from_slice(source);
                Ok::<_, &str>(bytes)
            })
            .transpose()?;
        let mut facts = Box::new(docx_model::NativeReadingPictureBullet {
            resource_key: String::new(),
            raw_pbi_flags,
            flags_origin,
            index_origin,
            relative_cp: bullet.relative_cp,
            picf_offset: picture.offset,
            shape: 75,
            raw_shape_flags: picture.raw_shape_flags,
            goal_twips: picture.goal,
            scale_per_mille: picture.scale,
            pib_flags: picture.pib_flags,
            client_anchor,
            client_anchor_options: picture.client_anchor_options,
        });
        // Existing Frame.extent validates range and performs exact stored
        // twip/per-mille scaling. 12700 EMU/pt is a unit conversion, not a
        // fitted Word marker scale. Never derive points from raster pixels.
        let display_box = picture.extent.map(|emu| emu as f64 / 12700.0);
        let projected =
            self.direct_picture_bullet(formatting, bullet, display_box, remaining_bytes)?;
        // Clone the selected resource identity, not an independently reconstructed
        // DOC path. Admission preflight precedes the owned copy; paragraph_metadata
        // charges its retained capacity once with the marker facts.
        remaining_bytes
            .checked_sub(projected.picture.resource_key.len())
            .ok_or("OUTPUT_TOO_LARGE")?;
        facts
            .resource_key
            .try_reserve_exact(projected.picture.resource_key.len())
            .map_err(|_| "OUTPUT_TOO_LARGE")?;
        facts.resource_key.push_str(&projected.picture.resource_key);
        Ok((projected, facts))
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
                self.omitted = true;
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
