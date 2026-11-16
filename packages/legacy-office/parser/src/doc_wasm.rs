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
    markdown: MarkdownSource,
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
            markdown: MarkdownSource::retain(data, maximum),
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
    /// GitHub-flavoured Markdown of the document, independent of the pull
    /// cursor's streaming state: the `to_markdown` model source capability
    /// the DOCX worker checks for.
    pub fn to_markdown(&self) -> Result<String, JsValue> {
        self.markdown.project().map_err(string_error)
    }
}

/// The source bytes `to_markdown` projects from. The pull cursor consumes the
/// projected model while streaming, so the Markdown projection reads the
/// source again, as `DocxArchive::to_markdown` re-reads its retained package.
///
/// The CFB and model budgets do not count the input's own length (a small
/// document can sit in a large compound file), so retention has its own
/// bound: the copy is kept only while it fits the model byte budget the
/// document was opened with. A longer source still loads and renders; only
/// its Markdown projection is a resource-limit rejection. A projection holds
/// a second model, under that same budget, for the duration of the call.
struct MarkdownSource {
    bytes: Option<Vec<u8>>,
    max_model_bytes: usize,
}

impl MarkdownSource {
    fn retain(data: Vec<u8>, max_model_bytes: usize) -> Self {
        Self {
            bytes: (data.len() <= max_model_bytes).then_some(data),
            max_model_bytes,
        }
    }

    fn project(&self) -> Result<String, String> {
        let bytes = self.bytes.as_deref().ok_or_else(|| {
            crate::rejection::resource("DOC source exceeds the Markdown retention budget")
        })?;
        doc::direct_markdown(bytes, self.max_model_bytes)
    }
}

fn string_error(error: String) -> JsValue {
    JsValue::from_str(&error)
}
fn error(error: &str) -> JsValue {
    JsValue::from_str(error)
}

#[cfg(test)]
mod tests {
    use super::MarkdownSource;

    #[test]
    fn markdown_source_is_retained_only_within_the_model_byte_budget() {
        // Within the budget the bytes are read (and here refused as no DOC).
        let kept = MarkdownSource::retain(vec![0; 16], 16);
        assert!(kept.project().unwrap_err().starts_with("UNSUPPORTED:"));
        // Beyond it nothing is kept, and the projection says why.
        let dropped = MarkdownSource::retain(vec![0; 16], 15);
        assert!(dropped.bytes.is_none());
        assert_eq!(
            dropped.project().unwrap_err(),
            "LEGACY_RESOURCE_LIMIT:DOC source exceeds the Markdown retention budget"
        );
    }
}
