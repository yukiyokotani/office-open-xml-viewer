//! Native, owning PowerPoint session boundary. Parsed offsets and admitted
//! media remain valid until this value is dropped; per-session budgets are
//! monotonic so requesting slides cannot multiply parser or model work limits.
use super::*;

const MAX_MODEL_BYTES: usize = 256 * 1024 * 1024;

pub(crate) struct DirectSession {
    document: Vec<u8>,
    pictures: Option<Vec<u8>>,
    presentation: persist::OwnedPresentation,
    media: media::SpanStore,
    work_budget: usize,
    text_budget: usize,
    model_budget: usize,
    poisoned: bool,
}

pub(crate) struct Resource<'a> {
    pub extension: &'static str,
    pub bytes: &'a [u8],
}

impl DirectSession {
    pub fn new(cfb: &CompoundFile<'_>) -> Result<Self, String> {
        if cfb.has_entry("EncryptedSummary") {
            return Err(unsupported(
                "encrypted PowerPoint binary documents are not supported",
            ));
        }
        let document = cfb.stream("PowerPoint Document").map_err(unsupported)?;
        let current_user = cfb.stream("Current User").map_err(unsupported)?;
        let mut work_budget = MAX_RECORDS;
        let current_edit = parse_current_user_atom(&current_user, &mut work_budget)?;
        let presentation = persist::resolve_owned(&document, current_edit, &mut work_budget)?;
        let pictures = (cfb.has_entry("Pictures") && !presentation.image_entries.is_empty())
            .then(|| cfb.stream("Pictures").map_err(unsupported))
            .transpose()?;
        Ok(Self::from_resolved(
            document,
            pictures,
            presentation,
            work_budget,
        ))
    }

    #[cfg(test)]
    fn from_streams(
        document: Vec<u8>,
        current_user: Vec<u8>,
        pictures: Option<Vec<u8>>,
    ) -> Result<Self, String> {
        let mut work_budget = MAX_RECORDS;
        let current_edit = parse_current_user_atom(&current_user, &mut work_budget)?;
        let presentation = persist::resolve_owned(&document, current_edit, &mut work_budget)?;
        Ok(Self::from_resolved(
            document,
            pictures,
            presentation,
            work_budget,
        ))
    }

    fn from_resolved(
        document: Vec<u8>,
        mut pictures: Option<Vec<u8>>,
        mut presentation: persist::OwnedPresentation,
        work_budget: usize,
    ) -> Self {
        let entries = std::mem::take(&mut presentation.image_entries);
        if entries.is_empty() {
            pictures = None;
        }
        Self {
            document,
            pictures,
            presentation,
            media: media::SpanStore::new(entries),
            work_budget,
            text_budget: MAX_TEXT_BYTES,
            model_budget: MAX_MODEL_BYTES,
            poisoned: false,
        }
    }

    pub fn slide(&mut self, index: usize) -> Result<pptx_model::Slide, String> {
        if self.poisoned {
            return Err(unsupported("PowerPoint direct session is poisoned"));
        }
        let result = (|| {
            let (record, _) = self
                .presentation
                .slides
                .get(index)
                .ok_or_else(|| unsupported("PowerPoint slide index out of range"))?;
            let record = record.view(&self.document)?;
            if contains_record(
                record.payload,
                DOCUMENT_ENCRYPTION_ATOM,
                0,
                &mut self.work_budget,
            )? {
                return Err(unsupported("encrypted PowerPoint slide"));
            }
            drawing::direct_model::slide(
                index,
                &self.presentation,
                &self.document,
                self.pictures.as_deref(),
                &mut self.media,
                &mut self.work_budget,
                &mut self.text_budget,
                &mut self.model_budget,
            )
        })();
        if result.is_err() {
            // A producer may have admitted media before a later shape fails.
            // Fail the session closed rather than exposing partial resources.
            self.poisoned = true;
        }
        result
    }

    pub fn slide_count(&self) -> usize {
        self.presentation.slides.len()
    }

    pub fn size(&self) -> (u32, u32) {
        self.presentation.size
    }

    pub fn resource(&self, key: &str) -> Result<Resource<'_>, String> {
        if self.poisoned {
            return Err(unsupported("PowerPoint direct session is poisoned"));
        }
        let index = resource_index(key)?;
        let (extension, bytes) = self
            .media
            .image(index, &self.document, self.pictures.as_deref())?
            .ok_or_else(|| unsupported("PowerPoint image is unsupported"))?;
        Ok(Resource { extension, bytes })
    }

    #[cfg(test)]
    fn budgets(&self) -> (usize, usize, usize) {
        (self.work_budget, self.text_budget, self.model_budget)
    }
}

fn resource_index(key: &str) -> Result<u32, String> {
    let digits = key
        .strip_prefix("legacy-ppt/image/")
        .filter(|digits| !digits.is_empty())
        .ok_or_else(|| unsupported("invalid PowerPoint resource key"))?;
    if (digits.len() > 1 && digits.starts_with('0'))
        || !digits.bytes().all(|byte| byte.is_ascii_digit())
    {
        return Err(unsupported("invalid PowerPoint resource key"));
    }
    digits
        .parse::<u32>()
        .ok()
        .filter(|index| *index != 0)
        .ok_or_else(|| unsupported("invalid PowerPoint resource key"))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(options: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
        [
            options.to_le_bytes().as_slice(),
            kind.to_le_bytes().as_slice(),
            (payload.len() as u32).to_le_bytes().as_slice(),
            payload,
        ]
        .concat()
    }

    fn png_blip() -> (Vec<u8>, Vec<u8>) {
        let mut png = vec![0; 33];
        png[..16].copy_from_slice(b"\x89PNG\r\n\x1a\n\0\0\0\x0dIHDR");
        png[19] = 1;
        png[23] = 1;
        png[24] = 8;
        png[25] = 6;
        let blip = record(0x6e00, 0xf01e, &[vec![0; 17], png.clone()].concat());
        (png, blip)
    }

    fn current_user(token: u32) -> Vec<u8> {
        [
            0u16.to_le_bytes().as_slice(),
            CURRENT_USER_ATOM.to_le_bytes().as_slice(),
            24u32.to_le_bytes().as_slice(),
            0x14u32.to_le_bytes().as_slice(),
            token.to_le_bytes().as_slice(),
        ]
        .concat()
    }

    fn slide_with_shape() -> Vec<u8> {
        let shape_flags = record(
            (1 << 4) | 2,
            0xf00a,
            &[42u32.to_le_bytes(), 0u32.to_le_bytes()].concat(),
        );
        let anchor = record(
            0,
            0xf010,
            &[0i16, 0, 576, 576]
                .into_iter()
                .flat_map(i16::to_le_bytes)
                .collect::<Vec<_>>(),
        );
        let shape = record(15, 0xf004, &[shape_flags, anchor].concat());
        let drawing = record(15, 1036, &record(15, 0xf002, &shape));
        record(15, SLIDE_CONTAINER, &drawing)
    }

    #[test]
    fn stream_constructor_rejects_encryption_and_malformed_current_user_first() {
        assert!(DirectSession::from_streams(
            Vec::new(),
            current_user(CURRENT_USER_ENCRYPTED),
            None,
        )
        .err()
        .expect("encrypted input must fail")
        .contains("encrypted"));
        assert!(DirectSession::from_streams(Vec::new(), vec![0; 7], None).is_err());
    }

    #[test]
    fn resource_keys_are_canonical_and_opaque() {
        for key in [
            "",
            "legacy-ppt/image/",
            "legacy-ppt/image/0",
            "legacy-ppt/image/01",
            "legacy-ppt/image/1.png",
            "../legacy-ppt/image/1",
            "https://example.invalid/legacy-ppt/image/1",
        ] {
            assert!(resource_index(key).is_err(), "{key}");
        }
        assert_eq!(resource_index("legacy-ppt/image/1").unwrap(), 1);
        assert_eq!(
            resource_index("legacy-ppt/image/4294967295").unwrap(),
            u32::MAX
        );
    }

    #[test]
    fn session_owns_sources_keeps_cumulative_budgets_and_poisoning_hides_admitted_media() {
        let (mut document, edit) = persist::tests::fixture();
        let mut work = MAX_RECORDS;
        let mut presentation = persist::resolve_owned(&document, edit, &mut work).unwrap();
        let (png, blip) = png_blip();
        let offset = document.len();
        document.extend_from_slice(&blip);
        let entry = record_span_with_end(&document, offset, &mut work, "PowerPoint")
            .unwrap()
            .0;
        presentation.image_entries = vec![entry];
        let slide_offset = document.len();
        document.extend_from_slice(&slide_with_shape());
        presentation.slides[0].0 =
            record_span_with_end(&document, slide_offset, &mut work, "PowerPoint")
                .unwrap()
                .0;
        let mut session = DirectSession::from_resolved(document, None, presentation, work);

        assert!(session
            .media
            .reference(1, &session.document, None, &mut session.work_budget)
            .unwrap());
        let first = session.resource("legacy-ppt/image/1").unwrap();
        assert_eq!(first.extension, "png");
        assert_eq!(first.bytes, png);
        let pointer = first.bytes.as_ptr();
        assert_eq!(
            session
                .resource("legacy-ppt/image/1")
                .unwrap()
                .bytes
                .as_ptr(),
            pointer
        );

        let before = session.budgets();
        let slide = session.slide(0).unwrap();
        assert_eq!(slide.elements.len(), 1);
        let after = session.budgets();
        assert!(after.0 < before.0);
        assert!(after.1 <= before.1);
        assert!(after.2 <= before.2);
        assert_eq!(session.resource("legacy-ppt/image/1").unwrap().bytes, png);

        assert!(session.slide(usize::MAX).is_err());
        assert!(session.resource("legacy-ppt/image/1").is_err());
        assert!(session.slide(0).unwrap_err().contains("poisoned"));
    }

    #[test]
    #[ignore = "requires LEGACY_DIRECT_PPT_INPUT private fixture"]
    fn private_native_session_smoke_resolves_every_model_image_key() {
        fn fill_key(fill: &pptx_model::Fill) -> Option<&str> {
            match fill {
                pptx_model::Fill::Image { image_path, .. } => Some(image_path),
                _ => None,
            }
        }

        let input = std::env::var("LEGACY_DIRECT_PPT_INPUT")
            .expect("LEGACY_DIRECT_PPT_INPUT must name a private PPT fixture");
        let bytes = std::fs::read(input).expect("private PPT fixture must be readable");
        let cfb = CompoundFile::open(&bytes).expect("private PPT fixture must be valid CFB");
        let mut session = DirectSession::new(&cfb).expect("direct PPT session must initialize");
        let slide_count = session.presentation.slides.len();
        let mut element_count = 0usize;
        let mut resource_count = 0usize;
        for index in 0..slide_count {
            let slide = session
                .slide(index)
                .unwrap_or_else(|error| panic!("direct slide {index} failed: {error}"));
            element_count = element_count
                .checked_add(slide.elements.len())
                .expect("element count overflow");
            let mut keys = slide
                .background
                .as_ref()
                .and_then(fill_key)
                .into_iter()
                .collect::<Vec<_>>();
            for element in &slide.elements {
                match element {
                    pptx_model::SlideElement::Picture(picture) => keys.push(&picture.image_path),
                    pptx_model::SlideElement::Shape(shape) => {
                        if let Some(key) = shape.fill.as_ref().and_then(fill_key) {
                            keys.push(key);
                        }
                    }
                    // The current native legacy producer admits only shapes
                    // and passive pictures. Fail the smoke test if that
                    // contract expands without adding resource traversal.
                    _ => panic!("direct legacy producer returned an unexpected element kind"),
                }
            }
            for key in keys {
                session.resource(key).unwrap_or_else(|error| {
                    panic!("direct slide {index} resource resolution failed: {error}")
                });
                resource_count += 1;
            }
        }
        eprintln!(
            "direct PPT smoke: {slide_count} slides, {element_count} elements, {resource_count} image references"
        );
    }
}
