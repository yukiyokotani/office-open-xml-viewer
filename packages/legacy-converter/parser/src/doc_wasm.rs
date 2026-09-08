//! Experimental direct `.doc` document-source boundary.

use crate::{
    cfb::CompoundFile,
    doc::{self, direct_cursor::DirectCursor},
};
use wasm_bindgen::prelude::*;

const MAX_DIRECT_DOC_SOURCE_BYTES: usize = 256 * 1024 * 1024;
/// Implementation heap-retention policy, not an MS-DOC format limit and not a
/// serialized-JSON size promise.
const MAX_DIRECT_DOC_MODEL_BYTES: usize = 256 * 1024 * 1024;

#[wasm_bindgen]
pub struct LegacyDocDocument {
    cursor: DirectCursor,
}

#[wasm_bindgen]
impl LegacyDocDocument {
    #[wasm_bindgen(constructor)]
    pub fn new(data: Vec<u8>, max_model_bytes: Option<u32>) -> Result<Self, JsValue> {
        console_error_panic_hook::set_once();
        if data.len() > MAX_DIRECT_DOC_SOURCE_BYTES {
            return Err(error("UNSUPPORTED:DOC direct source byte budget exceeded"));
        }
        let maximum = max_model_bytes
            .map(|n| n as usize)
            .unwrap_or(MAX_DIRECT_DOC_MODEL_BYTES);
        if maximum == 0 {
            return Err(error("OUTPUT_TOO_LARGE"));
        }
        if maximum > MAX_DIRECT_DOC_MODEL_BYTES {
            return Err(error(
                "UNSUPPORTED:DOC direct model byte budget exceeds policy",
            ));
        }
        let cfb = CompoundFile::open(&data).map_err(|e| error(&format!("UNSUPPORTED:{e}")))?;
        let result = doc::direct_model(&cfb, maximum).map_err(|e| error(&e))?;
        Ok(Self {
            cursor: DirectCursor::new(result).map_err(|e| error(&e))?,
        })
    }
    pub fn open_document_cursor(
        &mut self,
        operation_id: u32,
        generation: u32,
    ) -> Result<(), JsValue> {
        self.cursor
            .open_document_cursor(operation_id, generation)
            .map_err(string_error)
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
            .map_err(string_error)
    }
    pub fn document_chunk_done(&self) -> Result<bool, JsValue> {
        self.cursor.document_chunk_done().map_err(string_error)
    }
    pub fn acknowledge_document_chunk(
        &mut self,
        sequence: u32,
        operation_id: u32,
        generation: u32,
    ) -> Result<(), JsValue> {
        self.cursor
            .acknowledge_document_chunk(sequence, operation_id, generation)
            .map_err(string_error)
    }
    pub fn cancel_document_cursor(&mut self) -> Result<(), JsValue> {
        self.cursor.cancel_document_cursor().map_err(string_error)
    }
    pub fn close_document_session(&mut self) {
        self.cursor.close_document_session()
    }
    pub fn assert_healthy(&self) -> Result<(), JsValue> {
        self.cursor.assert_healthy().map_err(string_error)
    }
    pub fn extract_image(&self, key: &str) -> Result<Vec<u8>, JsValue> {
        self.cursor.extract_image(key).map_err(string_error)
    }
    pub fn image_mime_type(&self, key: &str) -> Result<String, JsValue> {
        self.cursor
            .resource_mime_type(key)
            .map(str::to_owned)
            .map_err(string_error)
    }
}

fn string_error(error: String) -> JsValue {
    JsValue::from_str(&error)
}
fn error(error: &str) -> JsValue {
    JsValue::from_str(error)
}
