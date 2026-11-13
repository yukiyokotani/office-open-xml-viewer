//! The PPTX model-source archive contract of the ODP reader: a bootstrap,
//! one serialized slide per pull (exact retry, acknowledge, cancel) and
//! image reads by admitted package path. Close is terminal.

use crate::odp::OdpDeck;
use crate::package::OdfPackage;
use ooxml_common::json_measurement::serialize_json_with_limit;
use ooxml_common::pull::insufficient_credit_error;
use ooxml_common::resource::{
    HardResourceLimitKind, HARD_MAX_PPTX_BOOTSTRAP_JSON_BYTES, HARD_MAX_PPTX_BOOTSTRAP_SLIDES,
};
use pptx_model::{BootstrapSlide, PresentationBootstrap};

struct Prepared {
    index: usize,
    operation_id: u32,
    generation: u32,
    bytes: Option<Vec<u8>>,
    length: usize,
}

pub struct OdpCursor {
    state: Option<(OdfPackage, OdpDeck)>,
    prepared: Option<Prepared>,
}

impl OdpCursor {
    pub fn new(package: OdfPackage, deck: OdpDeck) -> Self {
        Self {
            state: Some((package, deck)),
            prepared: None,
        }
    }

    fn open(&self) -> Result<&(OdfPackage, OdpDeck), String> {
        self.state
            .as_ref()
            .ok_or_else(|| "OpenDocument presentation cursor is closed".to_string())
    }

    pub fn presentation_bootstrap(&self) -> Result<Vec<u8>, String> {
        let (_, deck) = self.open()?;
        if self.prepared.is_some() {
            return Err("a slide unit is awaiting acknowledgement".to_string());
        }
        if deck.slides.len() as u64 > HARD_MAX_PPTX_BOOTSTRAP_SLIDES {
            return Err("OpenDocument bootstrap exceeds the PPTX slide ceiling".to_string());
        }
        let bootstrap = PresentationBootstrap {
            slide_count: deck.slides.len(),
            slide_width: deck.width,
            slide_height: deck.height,
            default_text_color: Some("000000".into()),
            major_font: None,
            minor_font: None,
            hlink_color: deck.hyperlink_colors.as_ref().map(|c| c.0.clone()),
            fol_hlink_color: deck.hyperlink_colors.as_ref().map(|c| c.1.clone()),
            embedded_fonts: Vec::new(),
            slides: deck
                .slides
                .iter()
                .map(|s| BootstrapSlide {
                    index: s.index,
                    part_name: s.part_name.clone(),
                })
                .collect(),
        };
        serialize_json_with_limit(
            &bootstrap,
            None,
            HardResourceLimitKind::PptxBootstrapJsonBytes,
            None,
            HARD_MAX_PPTX_BOOTSTRAP_JSON_BYTES,
            "OpenDocument bootstrap exceeds the PPTX JSON ceiling",
        )
    }

    pub fn pull_slide(
        &mut self,
        index: usize,
        operation_id: u32,
        generation: u32,
        credit: usize,
    ) -> Result<Vec<u8>, String> {
        self.open()?;
        if operation_id == 0 || generation == 0 || credit == 0 {
            return Err("operation id, generation, and byte credit must be positive".to_string());
        }
        if self.prepared.is_none() {
            let (_, deck) = self.open()?;
            // The slide was serialized (within the PPTX slide JSON ceiling)
            // when the deck was read.
            let bytes = deck
                .slides
                .get(index)
                .ok_or_else(|| "OpenDocument slide index out of range".to_string())?
                .json
                .clone();
            let length = bytes.len();
            self.prepared = Some(Prepared {
                index,
                operation_id,
                generation,
                bytes: Some(bytes),
                length,
            });
        }
        let prepared = self.prepared.as_mut().expect("prepared above");
        if (prepared.index, prepared.operation_id, prepared.generation)
            != (index, operation_id, generation)
        {
            return Err("another slide unit is awaiting acknowledgement".to_string());
        }
        if prepared.bytes.is_none() {
            return Err("slide unit must be acknowledged before another pull".to_string());
        }
        if prepared.length > credit {
            return Err(insufficient_credit_error(prepared.length, credit));
        }
        Ok(prepared.bytes.take().expect("checked above"))
    }

    pub fn acknowledge_slide(&mut self, operation_id: u32, generation: u32) -> Result<(), String> {
        self.open()?;
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
        self.prepared = None;
        Ok(())
    }

    pub fn cancel_slide(&mut self) -> Result<(), String> {
        self.open()?;
        self.prepared = None;
        Ok(())
    }

    pub fn close(&mut self) {
        self.prepared = None;
        self.state = None;
    }

    pub fn assert_healthy(&self) -> Result<(), String> {
        let (package, _) = self.open()?;
        package.session().assert_healthy()
    }

    /// The bytes of an admitted resource (a path the slides reference).
    pub fn extract_image(&self, path: &str) -> Result<Vec<u8>, String> {
        let (package, deck) = self.open()?;
        if !deck.resources.contains_key(path) {
            return Err("unknown OpenDocument resource".to_string());
        }
        if let Some(converted) = deck.converted.get(path) {
            return Ok(converted.clone());
        }
        package
            .session()
            .run_operation("odp-image", |operation| operation.read_bytes(path))
    }

    pub fn resource_mime(&self, path: &str) -> Result<&'static str, String> {
        let (_, deck) = self.open()?;
        deck.resources
            .get(path)
            .map(|r| r.mime)
            .ok_or_else(|| "unknown OpenDocument resource".to_string())
    }
}
