//! WASM boundary of the ODP reader: the PPTX model-source archive contract
//! (`presentation_bootstrap`, `pull_slide`, … `extract_image`).

use crate::odp;
use crate::odp_cursor::OdpCursor;
use crate::package::{OdfPackage, PackageLimits};
use ooxml_common::resource::OoxmlFormat;
use wasm_bindgen::prelude::*;

#[wasm_bindgen]
pub struct OdpPresentation {
    cursor: OdpCursor,
}

#[wasm_bindgen]
impl OdpPresentation {
    /// Open and fully admit the package; every rejection happens here.
    /// Limits are the caller's normalized ZIP policy in the shared
    /// governor's encoding (`undefined`: the standard limit; `0`: removed,
    /// hard ceilings still apply).
    #[wasm_bindgen(constructor)]
    pub fn new(
        data: js_sys::Uint8Array,
        max_archive_entry_bytes: Option<f64>,
        max_total_inflated_bytes: Option<f64>,
        max_archive_entries: Option<f64>,
    ) -> Result<OdpPresentation, JsValue> {
        console_error_panic_hook::set_once();
        let limits = PackageLimits {
            max_archive_entry_bytes: limit(max_archive_entry_bytes)?,
            max_total_inflated_bytes: limit(max_total_inflated_bytes)?,
            max_archive_entries: limit(max_archive_entries)?,
        };
        let package = OdfPackage::open(data.to_vec(), OoxmlFormat::Pptx, limits, &odp::MEDIA_TYPES)
            .map_err(|e| JsValue::from_str(&e))?;
        let deck = odp::read(&package).map_err(|e| JsValue::from_str(&e))?;
        Ok(Self {
            cursor: OdpCursor::new(package, deck),
        })
    }

    pub fn presentation_bootstrap(&self) -> Result<Vec<u8>, JsValue> {
        self.cursor.presentation_bootstrap().map_err(error)
    }

    pub fn pull_slide(
        &mut self,
        slide_index: u32,
        operation_id: u32,
        generation: u32,
        byte_credit: u32,
    ) -> Result<Vec<u8>, JsValue> {
        self.cursor
            .pull_slide(
                slide_index as usize,
                operation_id,
                generation,
                byte_credit as usize,
            )
            .map_err(error)
    }

    pub fn acknowledge_slide(&mut self, operation_id: u32, generation: u32) -> Result<(), JsValue> {
        self.cursor
            .acknowledge_slide(operation_id, generation)
            .map_err(error)
    }

    pub fn cancel_slide(&mut self) -> Result<(), JsValue> {
        self.cursor.cancel_slide().map_err(error)
    }

    pub fn close_presentation_session(&mut self) {
        self.cursor.close()
    }

    pub fn assert_healthy(&self) -> Result<(), JsValue> {
        self.cursor.assert_healthy().map_err(error)
    }

    pub fn extract_image(&self, path: &str) -> Result<Vec<u8>, JsValue> {
        self.cursor.extract_image(path).map_err(error)
    }

    pub fn image_mime_type(&self, path: &str) -> Result<String, JsValue> {
        self.cursor
            .resource_mime(path)
            .map(str::to_owned)
            .map_err(error)
    }

    /// The reader keeps no incremental usage snapshot.
    pub fn slide_cursor_resource_usage(&self) -> Result<Vec<u8>, JsValue> {
        self.cursor.assert_healthy().map_err(error)?;
        Err(JsValue::from_str("slide cursor usage is unavailable"))
    }
}

fn limit(value: Option<f64>) -> Result<Option<u64>, JsValue> {
    match value {
        None => Ok(None),
        Some(v)
            if v.is_finite() && v >= 0.0 && v.fract() == 0.0 && v <= 9_007_199_254_740_991.0 =>
        {
            Ok(Some(v as u64))
        }
        Some(_) => Err(JsValue::from_str("OpenDocument resource limit is invalid")),
    }
}

fn error(message: String) -> JsValue {
    JsValue::from_str(&message)
}
