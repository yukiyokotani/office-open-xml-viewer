//! Experimental direct `.doc` document-source boundary.

use crate::{
    cfb::CompoundFile,
    doc::{self, direct_cursor::DirectCursor},
};
use wasm_bindgen::prelude::*;

/// Implementation heap-retention policy, not an MS-DOC format limit and not a
/// serialized-JSON size promise.
const MAX_DIRECT_DOC_MODEL_BYTES: usize = 256 * 1024 * 1024;

#[wasm_bindgen]
pub struct LegacyDocDocument {
    cursor: DirectCursor,
    /// Retained source bytes for `to_markdown`: the pull cursor consumes the
    /// projected model while streaming, so the markdown projection re-parses
    /// from these bytes, the same way `DocxArchive::to_markdown` re-reads its
    /// retained package. One copy per open session, already bounded by the CFB
    /// acquisition limits at construction.
    bytes: Vec<u8>,
    max_model_bytes: usize,
}

#[wasm_bindgen]
impl LegacyDocDocument {
    #[wasm_bindgen(constructor)]
    pub fn new(data: js_sys::Uint8Array, max_model_bytes: Option<u32>) -> Result<Self, JsValue> {
        console_error_panic_hook::set_once();
        let data = crate::source_bytes::copy(data)?;
        let maximum = max_model_bytes
            .map(|n| n as usize)
            .unwrap_or(MAX_DIRECT_DOC_MODEL_BYTES);
        if maximum == 0 {
            return Err(error("OUTPUT_TOO_LARGE"));
        }
        if maximum > MAX_DIRECT_DOC_MODEL_BYTES {
            return Err(error(
                "LEGACY_RESOURCE_LIMIT:DOC direct model byte budget exceeds policy",
            ));
        }
        let cfb =
            CompoundFile::open(&data).map_err(|e| error(&crate::rejection::unsupported(e)))?;
        let result = doc::direct_model(&cfb, maximum).map_err(|e| error(&e))?;
        Ok(Self {
            cursor: DirectCursor::new(result).map_err(|e| error(&e))?,
            bytes: data,
            max_model_bytes: maximum,
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
    /// MS-DOC 2.7.2 DopBase fRMPrint: the DOC asks for its revision markup
    /// in print/PDF output. A host may use it to select the tracked-change
    /// view for this legacy input when the caller did not choose one: the
    /// display target for a legacy DOC is Word's PDF export, which follows
    /// fRMPrint (a corpus DOC with fRMView and fRMPrint set exports with
    /// insertions underlined and margin change bars), not the on-screen
    /// fRMView setting (MS-DOC 2.7.2 note <166>: the two can differ).
    pub fn revision_markup_in_print(&self) -> bool {
        self.cursor.revision_markup_in_print()
    }
    /// Whether the projected model carries any revision mark.
    pub fn has_revision_marks(&self) -> bool {
        self.cursor.has_revision_marks()
    }
    /// MS-DOC 2.7.2 DopBase fRMView: the DOC asks for its revision markup on
    /// screen.
    pub fn revision_markup_on_screen(&self) -> bool {
        self.cursor.revision_markup_on_screen()
    }
    pub fn image_mime_type(&self, key: &str) -> Result<String, JsValue> {
        self.cursor
            .resource_mime_type(key)
            .map(str::to_owned)
            .map_err(string_error)
    }
    /// GitHub-flavoured markdown projection of the document, independent of
    /// the pull cursor's streaming state. This is the `to_markdown` model
    /// source capability the DOCX worker checks for.
    pub fn to_markdown(&self) -> Result<String, JsValue> {
        doc::direct_markdown(&self.bytes, self.max_model_bytes).map_err(string_error)
    }
}

fn string_error(error: String) -> JsValue {
    JsValue::from_str(&error)
}
fn error(error: &str) -> JsValue {
    JsValue::from_str(error)
}
