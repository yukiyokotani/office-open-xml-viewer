//! Pull/acknowledge boundary for the owning native PowerPoint session.
//! Cancellation discards only an undelivered model buffer; the session retains
//! its cumulative budgets and admitted media until terminal close.
use super::direct_session::DirectSession;
use ooxml_common::json_measurement::measure_json;
use ooxml_common::pull::insufficient_credit_error;
use ooxml_common::resource::HARD_MAX_PPTX_SLIDE_JSON_BYTES;
use ooxml_common::resource::{HARD_MAX_PPTX_BOOTSTRAP_JSON_BYTES, HARD_MAX_PPTX_BOOTSTRAP_SLIDES};
use pptx_model::{BootstrapSlide, PresentationBootstrap};

struct PreparedSlide {
    index: usize,
    operation_id: u32,
    generation: u32,
    bytes: Option<Vec<u8>>,
    byte_length: usize,
}

pub(crate) struct DirectCursor {
    session: Option<DirectSession>,
    prepared: Option<PreparedSlide>,
    closed: bool,
}

impl DirectCursor {
    pub fn new(session: DirectSession) -> Self {
        Self {
            session: Some(session),
            prepared: None,
            closed: false,
        }
    }

    pub fn presentation_bootstrap(&self) -> Result<Vec<u8>, String> {
        let session = self.session()?;
        if self.prepared.is_some() {
            return Err("a slide unit is awaiting acknowledgement".to_string());
        }
        let (slide_width, slide_height) = session.size();
        let slide_count = session.slide_count();
        if slide_count as u64 > HARD_MAX_PPTX_BOOTSTRAP_SLIDES {
            return Err("PowerPoint bootstrap exceeds the PPTX slide ceiling".to_string());
        }
        let mut slides = Vec::new();
        slides
            .try_reserve_exact(slide_count)
            .map_err(|_| "PowerPoint bootstrap slide allocation failed".to_string())?;
        slides.extend((0..slide_count).map(|index| BootstrapSlide {
            index,
            part_name: None,
        }));
        let bootstrap = PresentationBootstrap {
            slide_count,
            slide_width: i64::from(slide_width),
            slide_height: i64::from(slide_height),
            default_text_color: None,
            major_font: None,
            minor_font: None,
            hlink_color: None,
            fol_hlink_color: None,
            embedded_fonts: Vec::new(),
            slides,
        };
        let measured = measure_json(&bootstrap)?;
        if measured.json_bytes > HARD_MAX_PPTX_BOOTSTRAP_JSON_BYTES {
            return Err("PowerPoint bootstrap exceeds the PPTX JSON ceiling".to_string());
        }
        let capacity = usize::try_from(measured.json_bytes)
            .map_err(|_| "PowerPoint bootstrap JSON size exceeds this platform".to_string())?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(capacity)
            .map_err(|_| "PowerPoint bootstrap JSON allocation failed".to_string())?;
        serde_json::to_writer(&mut bytes, &bootstrap)
            .map_err(|error| format!("serialize error: {error}"))?;
        debug_assert_eq!(measured.json_bytes, bytes.len() as u64);
        Ok(bytes)
    }

    pub fn pull_slide(
        &mut self,
        slide_index: usize,
        operation_id: u32,
        generation: u32,
        byte_credit: usize,
    ) -> Result<Vec<u8>, String> {
        self.session()?;
        if operation_id == 0 || generation == 0 || byte_credit == 0 {
            return Err("operation id, generation, and byte credit must be positive".to_string());
        }
        if self.prepared.is_some() {
            return self.pull_prepared(slide_index, operation_id, generation, byte_credit);
        }

        let slide = match self
            .session
            .as_mut()
            .expect("open session checked above")
            .slide(slide_index)
        {
            Ok(slide) => slide,
            Err(error) => {
                let _ = self.close_presentation_session();
                return Err(error);
            }
        };
        let measured = match measure_json(&slide) {
            Ok(measured) => measured,
            Err(error) => {
                let _ = self.close_presentation_session();
                return Err(error);
            }
        };
        if measured.json_bytes > HARD_MAX_PPTX_SLIDE_JSON_BYTES {
            let _ = self.close_presentation_session();
            return Err("PowerPoint slide exceeds the PPTX slide JSON ceiling".to_string());
        }
        let capacity = usize::try_from(measured.json_bytes).map_err(|_| {
            let _ = self.close_presentation_session();
            "PowerPoint slide JSON size exceeds this platform".to_string()
        })?;
        let mut bytes = Vec::new();
        bytes.try_reserve_exact(capacity).map_err(|_| {
            let _ = self.close_presentation_session();
            "PowerPoint slide JSON allocation failed".to_string()
        })?;
        serde_json::to_writer(&mut bytes, &slide).map_err(|error| {
            let _ = self.close_presentation_session();
            format!("serialize error: {error}")
        })?;
        let byte_length = bytes.len();
        debug_assert_eq!(measured.json_bytes, byte_length as u64);
        self.prepared = Some(PreparedSlide {
            index: slide_index,
            operation_id,
            generation,
            bytes: Some(bytes),
            byte_length,
        });
        self.pull_prepared(slide_index, operation_id, generation, byte_credit)
    }

    fn pull_prepared(
        &mut self,
        slide_index: usize,
        operation_id: u32,
        generation: u32,
        byte_credit: usize,
    ) -> Result<Vec<u8>, String> {
        let prepared = self
            .prepared
            .as_mut()
            .expect("prepared slide checked above");
        if (prepared.index, prepared.operation_id, prepared.generation)
            != (slide_index, operation_id, generation)
        {
            return Err("another slide unit is awaiting acknowledgement".to_string());
        }
        if prepared.bytes.is_none() {
            return Err("slide unit must be acknowledged before another pull".to_string());
        }
        if prepared.byte_length > byte_credit {
            return Err(insufficient_credit_error(prepared.byte_length, byte_credit));
        }
        Ok(prepared.bytes.take().expect("prepared bytes checked above"))
    }

    pub fn acknowledge_slide(&mut self, operation_id: u32, generation: u32) -> Result<(), String> {
        self.session()?;
        let prepared = self
            .prepared
            .as_ref()
            .ok_or_else(|| "no slide unit is awaiting acknowledgement".to_string())?;
        if (prepared.operation_id, prepared.generation) != (operation_id, generation) {
            return Err("slide acknowledgement identity is stale or invalid".to_string());
        }
        if prepared.bytes.is_some() {
            return Err("slide unit cannot be acknowledged before delivery".to_string());
        }
        self.prepared.take();
        Ok(())
    }

    pub fn cancel_slide(&mut self) -> Result<(), String> {
        self.session()?;
        self.prepared.take();
        Ok(())
    }

    pub fn close_presentation_session(&mut self) -> Result<(), String> {
        self.prepared.take();
        self.session.take();
        self.closed = true;
        Ok(())
    }

    pub fn extract_image(&self, path: &str) -> Result<Vec<u8>, String> {
        Ok(self.session()?.resource(path)?.bytes.to_vec())
    }

    pub fn assert_healthy(&self) -> Result<(), String> {
        self.session().map(|_| ())
    }

    fn session(&self) -> Result<&DirectSession, String> {
        if self.closed {
            return Err("PowerPoint direct cursor is closed".to_string());
        }
        self.session
            .as_ref()
            .ok_or_else(|| "PowerPoint direct cursor has no session".to_string())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cfb::CompoundFile;

    #[test]
    fn pull_retries_exact_buffer_and_acknowledges_exact_identity() {
        let (session, _) = super::super::direct_session::cursor_fixture();
        let mut cursor = DirectCursor::new(session);
        let bootstrap: serde_json::Value =
            serde_json::from_slice(&cursor.presentation_bootstrap().unwrap()).unwrap();
        assert_eq!(bootstrap["slideCount"], 2);
        assert_eq!(bootstrap["slideWidth"], 12_192_000);
        assert_eq!(bootstrap["slideHeight"], 6_858_000);
        for field in [
            "defaultTextColor",
            "majorFont",
            "minorFont",
            "hlinkColor",
            "folHlinkColor",
        ] {
            assert!(bootstrap[field].is_null(), "{field}");
        }
        assert_eq!(bootstrap["embeddedFonts"], serde_json::json!([]));
        assert_eq!(
            bootstrap["slides"],
            serde_json::json!([{ "index": 0 }, { "index": 1 }])
        );
        assert!(bootstrap["slides"][0].get("partName").is_none());
        for (operation, generation, credit) in [(0, 1, 1), (1, 0, 1), (1, 1, 0)] {
            assert!(cursor.pull_slide(0, operation, generation, credit).is_err());
            assert!(cursor.prepared.is_none());
        }

        assert!(cursor
            .pull_slide(0, 7, 3, 1)
            .unwrap_err()
            .contains("credit"));
        let prepared = cursor.prepared.as_ref().unwrap();
        let length = prepared.byte_length;
        let pointer = prepared.bytes.as_ref().unwrap().as_ptr();
        assert!(cursor
            .acknowledge_slide(7, 3)
            .unwrap_err()
            .contains("delivery"));
        assert!(cursor.pull_slide(1, 7, 3, length).is_err());
        assert!(cursor.pull_slide(0, 8, 3, length).is_err());
        assert!(cursor.pull_slide(0, 7, 3, length - 1).is_err());
        assert_eq!(
            cursor
                .prepared
                .as_ref()
                .unwrap()
                .bytes
                .as_ref()
                .unwrap()
                .as_ptr(),
            pointer
        );
        let bytes = cursor.pull_slide(0, 7, 3, length).unwrap();
        assert_eq!(bytes.as_ptr(), pointer);
        assert_eq!(bytes.len(), length);
        assert!(cursor
            .pull_slide(0, 7, 3, length)
            .unwrap_err()
            .contains("acknowledged"));
        assert!(cursor.acknowledge_slide(8, 3).is_err());
        assert!(cursor.acknowledge_slide(7, 4).is_err());
        cursor.acknowledge_slide(7, 3).unwrap();
        assert!(cursor
            .acknowledge_slide(7, 3)
            .unwrap_err()
            .contains("no slide"));
    }

    #[test]
    fn cancel_retains_admitted_images_and_close_is_terminal_and_idempotent() {
        let (session, png) = super::super::direct_session::cursor_fixture();
        let mut cursor = DirectCursor::new(session);
        assert_eq!(cursor.extract_image("legacy-ppt/image/1").unwrap(), png);

        assert!(cursor.pull_slide(0, 1, 1, 1).is_err());
        cursor.cancel_slide().unwrap();
        assert_eq!(cursor.extract_image("legacy-ppt/image/1").unwrap(), png);

        cursor.pull_slide(0, 2, 1, usize::MAX).unwrap();
        cursor.cancel_slide().unwrap();
        assert_eq!(cursor.extract_image("legacy-ppt/image/1").unwrap(), png);
        cursor.pull_slide(0, 3, 1, usize::MAX).unwrap();
        cursor.acknowledge_slide(3, 1).unwrap();

        cursor.close_presentation_session().unwrap();
        cursor.close_presentation_session().unwrap();
        assert!(cursor.assert_healthy().is_err());
        assert!(cursor.presentation_bootstrap().is_err());
        assert!(cursor.extract_image("legacy-ppt/image/1").is_err());
        assert!(cursor.cancel_slide().is_err());
    }

    #[test]
    fn producer_failure_closes_cursor_and_blocks_bootstrap_and_resources() {
        let (session, _) = super::super::direct_session::cursor_fixture();
        let mut cursor = DirectCursor::new(session);
        assert!(cursor.pull_slide(usize::MAX, 1, 1, usize::MAX).is_err());
        assert!(cursor.assert_healthy().is_err());
        assert!(cursor.presentation_bootstrap().is_err());
        assert!(cursor.extract_image("legacy-ppt/image/1").is_err());
    }

    #[test]
    #[ignore = "requires LEGACY_DIRECT_PPT_INPUT private fixture"]
    fn native_source_bootstrap_pull_ack_cancel_and_close() {
        let input = std::env::var("LEGACY_DIRECT_PPT_INPUT")
            .expect("LEGACY_DIRECT_PPT_INPUT must name a private PPT fixture");
        let bytes = std::fs::read(input).expect("private PPT fixture must be readable");
        let cfb = CompoundFile::open(&bytes).expect("private PPT fixture must be valid CFB");
        let session = DirectSession::new(&cfb).expect("direct PPT session must initialize");
        let mut cursor = DirectCursor::new(session);
        let bootstrap = cursor.presentation_bootstrap().unwrap();
        assert!(!bootstrap.is_empty());
        let slide = cursor.pull_slide(0, 1, 1, usize::MAX).unwrap();
        assert!(!slide.is_empty());
        cursor.acknowledge_slide(1, 1).unwrap();
        cursor.cancel_slide().unwrap();
        cursor.close_presentation_session().unwrap();
        assert!(cursor.presentation_bootstrap().is_err());
        assert!(cursor.extract_image("legacy-ppt/image/1").is_err());
    }
}
