//! Native, owning PowerPoint session boundary. Parsed offsets and admitted
//! media remain valid until this value is dropped; per-session budgets are
//! monotonic so requesting slides cannot multiply parser or model work limits.
use super::*;

const MAX_MODEL_BYTES: usize = 256 * 1024 * 1024;
/// Implementation resource policy for retained source streams, not an MS-PPT format limit.
const MAX_DIRECT_STREAM_BYTES: usize = 256 * 1024 * 1024;

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
    // Read by the inspection API.
    #[cfg_attr(not(any(test, feature = "inspection")), allow(dead_code))]
    pub extension: &'static str,
    pub bytes: &'a [u8],
}

impl DirectSession {
    pub fn new(cfb: &CompoundFile<'_>) -> Result<Self, String> {
        let streams = cfb.scoped_streams().map_err(unsupported)?;
        if streams
            .has_stream(&["EncryptedSummary"])
            .map_err(unsupported)?
        {
            return Err(unsupported(
                "encrypted PowerPoint binary documents are not supported",
            ));
        }
        let document = streams
            .stream(&["PowerPoint Document"], MAX_DIRECT_STREAM_BYTES)
            .map_err(unsupported)?;
        let current_user = streams
            .stream(&["Current User"], MAX_DIRECT_STREAM_BYTES)
            .map_err(unsupported)?;
        let mut work_budget = MAX_RECORDS;
        let current_edit = parse_current_user_atom(&current_user, &mut work_budget)?;
        let presentation = persist::resolve_owned(&document, current_edit, &mut work_budget)?;
        let pictures = if presentation.image_entries.is_empty() {
            None
        } else {
            streams
                .optional_stream(&["Pictures"], MAX_DIRECT_STREAM_BYTES)
                .map_err(unsupported)?
        };
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

/// The direct cursor's tests use this session fixture.
#[cfg(test)]
pub(super) fn cursor_fixture() -> (DirectSession, Vec<u8>) {
    tests::cursor_fixture()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cfb::test_support::{build_cfb, build_scoped_cfb};

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

    fn valid_current_user(edit: usize) -> Vec<u8> {
        let mut payload = vec![0; 24];
        payload[..4].copy_from_slice(&0x14u32.to_le_bytes());
        payload[4..8].copy_from_slice(&CURRENT_USER_NOT_ENCRYPTED.to_le_bytes());
        payload[8..12].copy_from_slice(&(edit as u32).to_le_bytes());
        payload[14..16].copy_from_slice(&0x03f4u16.to_le_bytes());
        payload[16..18].copy_from_slice(&[3, 0]);
        payload[20..24].copy_from_slice(&8u32.to_le_bytes());
        record(0, CURRENT_USER_ATOM, &payload)
    }

    fn directory_offset(streams: usize) -> usize {
        512 + streams * 8 * 512
    }

    fn directory_link(bytes: &mut [u8], streams: usize, id: usize, at: usize, target: u32) {
        let offset = directory_offset(streams) + id * 128 + at;
        bytes[offset..offset + 4].copy_from_slice(&target.to_le_bytes());
    }

    fn make_storage(bytes: &mut [u8], streams: usize, id: usize) {
        let offset = directory_offset(streams) + id * 128;
        bytes[offset + 66] = 1;
        bytes[offset + 116..offset + 120].copy_from_slice(&0xffff_fffeu32.to_le_bytes());
        bytes[offset + 120..offset + 128].fill(0);
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

    pub(super) fn cursor_fixture() -> (DirectSession, Vec<u8>) {
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
        (session, png)
    }

    /// A root-scoped CFB holding a minimal edit chain: DocumentContainer
    /// (DocumentAtom 5760 x 4320 master units, SlideListWithText with one
    /// SlidePersistAtom), one SlideContainer with `slide_children`, a
    /// PersistDirectoryAtom, a UserEditAtom and the Current User stream.
    fn minimal_ppt(slide_children: &[u8]) -> Vec<u8> {
        let mut document_atom = vec![0u8; 40];
        document_atom[..4].copy_from_slice(&5760u32.to_le_bytes());
        document_atom[4..8].copy_from_slice(&4320u32.to_le_bytes());
        let mut slide_ref = [0u8; 20];
        slide_ref[..4].copy_from_slice(&2u32.to_le_bytes());
        let document = record(
            15,
            1000,
            &[
                record(1, 1001, &document_atom),
                record(15, 4080, &record(0, 1011, &slide_ref)),
            ]
            .concat(),
        );
        let slide = record(15, SLIDE_CONTAINER, slide_children);
        let directory_offset = document.len() + slide.len();
        let directory = record(
            0,
            0x1772,
            &[
                0x00200001u32.to_le_bytes(),
                0u32.to_le_bytes(),
                (document.len() as u32).to_le_bytes(),
            ]
            .concat(),
        );
        let current_edit = directory_offset + directory.len();
        let mut user_payload = [0u8; 28];
        user_payload[12..16].copy_from_slice(&(directory_offset as u32).to_le_bytes());
        user_payload[16..20].copy_from_slice(&1u32.to_le_bytes());
        let user_edit = record(0, USER_EDIT_ATOM, &user_payload);
        build_scoped_cfb(&[
            (
                "PowerPoint Document",
                [document, slide, directory, user_edit].concat(),
            ),
            ("Current User", valid_current_user(current_edit)),
        ])
    }

    /// One text box whose ClientTextbox holds `text` as a TextCharsAtom and a
    /// TextRulerAtom giving level 0 explicit zero origins.
    fn text_drawing(text: &str) -> Vec<u8> {
        let units: Vec<u8> = text.encode_utf16().flat_map(u16::to_le_bytes).collect();
        let ruler = record(0, 4006, &[(8u32 | 256).to_le_bytes(), [0; 4]].concat());
        let textbox = record(
            15,
            0xf00d,
            &[record(0, TEXT_CHARS_ATOM, &units), ruler].concat(),
        );
        let flags = record(
            (202 << 4) | 2,
            0xf00a,
            &[42u32.to_le_bytes(), 0xa00u32.to_le_bytes()].concat(),
        );
        let anchor = record(
            0,
            0xf010,
            &[0i16, 0, 576, 576]
                .into_iter()
                .flat_map(i16::to_le_bytes)
                .collect::<Vec<_>>(),
        );
        let shape = record(15, 0xf004, &[flags, anchor, textbox].concat());
        record(15, 1036, &record(15, 0xf002, &shape))
    }

    fn session(bytes: &[u8]) -> Result<DirectSession, String> {
        DirectSession::new(&CompoundFile::open(bytes).unwrap())
    }

    #[test]
    fn projects_unicode_text_of_a_minimal_edit_chain() {
        let mut session = session(&minimal_ppt(&text_drawing("Legacy 日本語 slide"))).unwrap();
        assert_eq!(session.slide_count(), 1);
        assert_eq!(session.size(), (9_144_000, 6_858_000));
        let slide = session.slide(0).unwrap();
        assert!(!slide.hidden);
        let pptx_model::SlideElement::Shape(shape) = &slide.elements[0] else {
            panic!("expected the text box")
        };
        let text: String = shape.text_body.as_ref().unwrap().paragraphs[0]
            .runs
            .iter()
            .filter_map(|run| match run {
                pptx_model::TextRun::Text(run) => Some(run.text.as_str()),
                _ => None,
            })
            .collect();
        assert_eq!(text, "Legacy 日本語 slide");
    }

    #[test]
    fn projects_the_live_slide_hidden_flag() {
        // MS-PPT 2.5.1 / 2.6.6: SlideShowSlideInfoAtom.fHidden is bit 2 of
        // the flags word following the two effect bytes.
        let mut info = [0u8; 16];
        info[10] = 4;
        let children = [record(0, 0x03f9, &info), text_drawing("Hidden")].concat();
        let mut session = session(&minimal_ppt(&children)).unwrap();
        assert!(session.slide(0).unwrap().hidden);
    }

    #[test]
    fn rejects_an_invalid_user_edit_chain() {
        let mut user_edit_payload = vec![0; 24];
        user_edit_payload[8..12].copy_from_slice(&1u32.to_le_bytes());
        let ppt = build_scoped_cfb(&[
            (
                "PowerPoint Document",
                record(0, USER_EDIT_ATOM, &user_edit_payload),
            ),
            ("Current User", valid_current_user(0)),
        ]);
        assert!(session(&ppt)
            .err()
            .expect("malformed UserEditAtom must fail")
            .contains("invalid PowerPoint UserEditAtom"));
    }

    #[test]
    fn cfb_constructor_rejects_encrypted_current_user_and_encrypted_summary() {
        let encrypted_user = build_scoped_cfb(&[
            ("PowerPoint Document", vec![0; 16]),
            ("Current User", current_user(CURRENT_USER_ENCRYPTED)),
        ]);
        let summary = build_scoped_cfb(&[
            ("EncryptedSummary", vec![0; 16]),
            ("PowerPoint Document", vec![0; 16]),
            ("Current User", valid_current_user(0)),
        ]);
        for bytes in [encrypted_user, summary] {
            assert!(session(&bytes)
                .err()
                .expect("encrypted input must fail")
                .contains("encrypted"));
        }
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
    fn cfb_constructor_uses_only_root_presentation_streams() {
        let (document, edit) = persist::tests::fixture();
        let current = valid_current_user(edit);
        let streams = [
            ("ObjectPool", Vec::new()),
            ("PowerPoint Document", b"embedded decoy".to_vec()),
            ("PowerPoint Document", document),
            ("Current User", current),
        ];
        let mut bytes = build_cfb(&streams);
        make_storage(&mut bytes, streams.len(), 1);
        // Root owns ObjectPool, Current User, and the real document. The
        // duplicate document is owned only by ObjectPool.
        directory_link(&mut bytes, streams.len(), 0, 76, 1);
        directory_link(&mut bytes, streams.len(), 1, 68, 4);
        directory_link(&mut bytes, streams.len(), 1, 72, 3);
        directory_link(&mut bytes, streams.len(), 1, 76, 2);

        let cfb = CompoundFile::open(&bytes).unwrap();
        let session = DirectSession::new(&cfb).unwrap();
        assert_eq!(session.slide_count(), 2);
    }

    #[test]
    fn cfb_constructor_rejects_nested_only_and_wrong_kind_required_streams() {
        let streams = [
            ("ObjectPool", Vec::new()),
            ("PowerPoint Document", b"nested only".to_vec()),
        ];
        let mut nested = build_cfb(&streams);
        make_storage(&mut nested, streams.len(), 1);
        directory_link(&mut nested, streams.len(), 0, 76, 1);
        directory_link(&mut nested, streams.len(), 1, 76, 2);
        let error = DirectSession::new(&CompoundFile::open(&nested).unwrap())
            .err()
            .expect("nested-only document must fail");
        assert!(error.contains("missing CFB path entry: PowerPoint Document"));

        let streams = [("PowerPoint Document", Vec::new())];
        let mut wrong_kind = build_cfb(&streams);
        make_storage(&mut wrong_kind, streams.len(), 1);
        directory_link(&mut wrong_kind, streams.len(), 0, 76, 1);
        let error = DirectSession::new(&CompoundFile::open(&wrong_kind).unwrap())
            .err()
            .expect("wrong-kind document must fail");
        assert!(error.contains("does not end in a stream"));
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
        let (mut session, png) = cursor_fixture();
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
