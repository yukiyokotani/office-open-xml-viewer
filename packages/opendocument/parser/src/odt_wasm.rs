//! WASM boundary of the ODT reader: the DOCX model-source archive contract,
//! backed by the shared owned document cursor.

use crate::odt;
use crate::package::{OdfPackage, PackageLimits};
use docx_model::owned_cursor::OwnedDocumentCursor;
use ooxml_common::resource::OoxmlFormat;
use wasm_bindgen::prelude::*;

#[wasm_bindgen]
pub struct OdtDocument {
    cursor: OwnedDocumentCursor,
}

#[wasm_bindgen]
impl OdtDocument {
    /// Open and fully admit the package. Every rejection happens here, before
    /// the cursor publishes anything. Limits are the caller's normalized ZIP
    /// policy in the shared governor's encoding (`undefined`: the standard
    /// limit; `0`: the caller removed the limit, hard ceilings still apply).
    #[wasm_bindgen(constructor)]
    pub fn new(
        data: js_sys::Uint8Array,
        max_archive_entry_bytes: Option<f64>,
        max_total_inflated_bytes: Option<f64>,
        max_archive_entries: Option<f64>,
    ) -> Result<OdtDocument, JsValue> {
        console_error_panic_hook::set_once();
        let limits = PackageLimits {
            max_archive_entry_bytes: limit(max_archive_entry_bytes)?,
            max_total_inflated_bytes: limit(max_total_inflated_bytes)?,
            max_archive_entries: limit(max_archive_entries)?,
        };
        let package = OdfPackage::open(data.to_vec(), OoxmlFormat::Docx, limits, &odt::MEDIA_TYPES)
            .map_err(|e| JsValue::from_str(&e))?;
        let (document, images) = odt::read(&package).map_err(|e| JsValue::from_str(&e))?;
        let cursor =
            OwnedDocumentCursor::new(document, images).map_err(|e| JsValue::from_str(&e))?;
        Ok(Self { cursor })
    }

    pub fn open_document_cursor(
        &mut self,
        operation_id: u32,
        generation: u32,
    ) -> Result<(), JsValue> {
        self.cursor
            .open_document_cursor(operation_id, generation)
            .map_err(error)
    }
    pub fn pull_document_chunk(
        &mut self,
        sequence: u32,
        operation_id: u32,
        generation: u32,
        byte_credit: u32,
    ) -> Result<Vec<u8>, JsValue> {
        self.cursor
            .pull_document_chunk(sequence, operation_id, generation, byte_credit as usize)
            .map_err(error)
    }
    pub fn document_chunk_done(&self) -> Result<bool, JsValue> {
        self.cursor.document_chunk_done().map_err(error)
    }
    pub fn acknowledge_document_chunk(
        &mut self,
        sequence: u32,
        operation_id: u32,
        generation: u32,
    ) -> Result<(), JsValue> {
        self.cursor
            .acknowledge_document_chunk(sequence, operation_id, generation)
            .map_err(error)
    }
    pub fn cancel_document_cursor(&mut self) -> Result<(), JsValue> {
        self.cursor.cancel_document_cursor().map_err(error)
    }
    pub fn close_document_session(&mut self) {
        self.cursor.close_document_session()
    }
    pub fn assert_healthy(&self) -> Result<(), JsValue> {
        self.cursor.assert_healthy().map_err(error)
    }
    pub fn extract_image(&self, key: &str) -> Result<Vec<u8>, JsValue> {
        self.cursor.extract_image(key).map_err(error)
    }
    pub fn image_mime_type(&self, key: &str) -> Result<String, JsValue> {
        self.cursor
            .resource_mime_type(key)
            .map(str::to_owned)
            .map_err(error)
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
