use ooxml_common::json_measurement::{serialize_json_limited, LimitedJsonError};
use ooxml_common::package_session::PackageLimitReporter;
use ooxml_common::pull::insufficient_credit_error;
use ooxml_common::resource::{
    HardResourceLimitKind, ResourceUsage, HARD_MAX_DOCX_BODY_CHUNK_JSON_BYTES,
    HARD_MAX_DOCX_BOOTSTRAP_JSON_BYTES, HARD_MAX_DOCX_MARKDOWN_BYTES,
    HARD_MAX_DOCX_RETAINED_MODEL_JSON_BYTES,
};
use wasm_bindgen::prelude::*;

mod chart_compatibility;
mod document_projector;
mod drawing_compatibility;
mod markdown;
mod math;
mod numbering;
mod parser;
mod ref_bookmark_flow;
mod styles;
mod types;
mod xml_util;

#[cfg(test)]
thread_local! {
    static DOCX_RETAINED_MODEL_JSON_LIMIT_OVERRIDE: std::cell::Cell<Option<u64>> =
        const { std::cell::Cell::new(None) };
    static DOCX_MARKDOWN_LIMIT_OVERRIDE: std::cell::Cell<Option<u64>> =
        const { std::cell::Cell::new(None) };
}

fn docx_retained_model_json_limit() -> u64 {
    #[cfg(test)]
    if let Some(limit) = DOCX_RETAINED_MODEL_JSON_LIMIT_OVERRIDE.with(std::cell::Cell::get) {
        return limit;
    }
    HARD_MAX_DOCX_RETAINED_MODEL_JSON_BYTES
}

fn docx_markdown_limit() -> u64 {
    #[cfg(test)]
    if let Some(limit) = DOCX_MARKDOWN_LIMIT_OVERRIDE.with(std::cell::Cell::get) {
        return limit;
    }
    HARD_MAX_DOCX_MARKDOWN_BYTES
}

fn render_markdown_from_zip(zip: &mut parser::Zip) -> Result<String, String> {
    let doc = parser::parse(zip)?;
    let limit = docx_markdown_limit();
    let output = markdown::render_document_with_limit(&doc, limit);
    zip.operation()?.limit_reporter()?.observe_hard_limit(
        HardResourceLimitKind::DocxMarkdownBytes,
        None,
        limit,
        output.observed(),
    )?;
    Ok(output.into_string())
}

fn render_markdown_from_bytes_with_limits(
    data: &[u8],
    max_archive_entry_bytes: Option<u64>,
    max_total_inflated_bytes: Option<u64>,
) -> Result<String, String> {
    let mut zip = parser::open_document_package(
        data.to_vec(),
        max_archive_entry_bytes,
        max_total_inflated_bytes,
        None,
    )?;
    zip.run_operation("markdown", render_markdown_from_zip)
}

/// Parse a docx archive and return the model as UTF-8 JSON **bytes**.
///
/// Returning `Vec<u8>` (a fresh copy on the JS side) instead of `String` keeps
/// the model out of the JsString/UTF-16 representation: the worker forwards the
/// resulting `ArrayBuffer` to the main thread as a transferable and the main
/// thread does a single `TextDecoder.decode` + `JSON.parse`, collapsing three
/// serializations (Rust String → JsString → structured clone) into one decode.
#[wasm_bindgen]
pub fn parse_docx(
    data: &[u8],
    max_archive_entry_bytes: Option<u64>,
    max_total_inflated_bytes: Option<u64>,
) -> Result<Vec<u8>, JsValue> {
    console_error_panic_hook::set_once();
    let doc = parser::parse_from_bytes_streamed_with_limits(
        data,
        max_archive_entry_bytes,
        max_total_inflated_bytes,
        "parse",
    )
    .map_err(docx_parser_js_error)?;
    serde_json::to_vec(&doc).map_err(|e| JsValue::from_str(&format!("serialize error: {e}")))
}

/// WASM-callable markdown projection (mirrors `to_markdown_native`). Returns
/// GitHub-flavoured markdown of headings / paragraphs / tables / footnotes,
/// discarding positioning, section properties, fonts, and drawing shapes.
#[wasm_bindgen]
pub fn docx_to_markdown(
    data: &[u8],
    max_archive_entry_bytes: Option<u64>,
    max_total_inflated_bytes: Option<u64>,
) -> Result<String, JsValue> {
    console_error_panic_hook::set_once();
    render_markdown_from_bytes_with_limits(data, max_archive_entry_bytes, max_total_inflated_bytes)
        .map_err(docx_markdown_js_error)
}

/// Extract raw bytes for a single embedded image entry (e.g.
/// "word/media/image1.png") from a docx zip archive. Used by the main thread to
/// lazily materialize image blobs through one bounded package operation.
#[wasm_bindgen]
pub fn extract_image(
    data: &[u8],
    path: &str,
    max_archive_entry_bytes: Option<u64>,
    max_total_inflated_bytes: Option<u64>,
) -> Result<Vec<u8>, JsValue> {
    extract_entry_with_limits(
        data,
        path,
        max_archive_entry_bytes,
        max_total_inflated_bytes,
        "extract-image",
    )
    .map_err(|e| JsValue::from_str(&e))
}

/// A stateful handle over an opened docx archive.
///
/// The free functions above (`parse_docx` / `docx_to_markdown` / `extract_image`)
/// each re-copy the whole file into WASM and re-scan the ZIP central directory on
/// every call. A `DocxArchive` copies the bytes into WASM **once** (in `new`) and
/// keeps the opened [`parser::Zip`] session alive, so a `parse` followed by any number of
/// `extract_image` calls (the viewer's parse-then-lazily-load-media pattern)
/// pays the copy + open cost a single time. The session owns the source bytes,
/// validated central-directory index, resource governor, and first package-wide
/// poison error.
struct DocumentCursorState {
    operation_id: u32,
    generation: u32,
    next_sequence: u32,
    accepted_json_bytes: u64,
    cursor: Option<parser::DocxBodyCursor>,
    degraded_terminal: Option<types::Document>,
}

struct PreparedDocumentChunk {
    operation_id: u32,
    generation: u32,
    sequence: u32,
    bytes: Option<Vec<u8>>,
    byte_length: usize,
    done: bool,
    accepted_json_bytes_after: u64,
}

fn serialize_document_unit(
    unit: &parser::StreamedDocumentUnit,
    reporter: Option<&PackageLimitReporter>,
) -> Result<Vec<u8>, String> {
    let (kind, limit) = match unit {
        parser::StreamedDocumentUnit::Body { .. } => (
            HardResourceLimitKind::DocxBodyChunkJsonBytes,
            HARD_MAX_DOCX_BODY_CHUNK_JSON_BYTES,
        ),
        parser::StreamedDocumentUnit::Complete { .. } => (
            HardResourceLimitKind::DocxBootstrapJsonBytes,
            HARD_MAX_DOCX_BOOTSTRAP_JSON_BYTES,
        ),
    };
    let bytes = match serialize_json_limited(unit, limit) {
        Ok(bytes) => bytes,
        Err(LimitedJsonError::LimitExceeded { observed, .. }) => {
            if let Some(reporter) = reporter {
                reporter.observe_hard_limit(kind, Some("word/document.xml"), limit, observed)?;
            }
            return Err(format!(
                "document cursor JSON exceeds its hard ceiling: {observed} > {limit}"
            ));
        }
        Err(LimitedJsonError::Serialize(error)) => {
            return Err(format!("serialize measurement error: {error}"));
        }
    };
    if let Some(reporter) = reporter {
        reporter.observe_hard_limit(kind, Some("word/document.xml"), limit, bytes.len() as u64)?;
    }
    Ok(bytes)
}

#[wasm_bindgen]
pub struct DocxArchive {
    /// The admitted OPC package. Construction fails closed when the input is not
    /// a ZIP, not an OPC package, or lacks `word/document.xml`
    /// (`ooxml_common::opc`), so every method operates on a real document.
    archive: parser::Zip,
    document_cursor: Option<DocumentCursorState>,
    prepared_document_chunk: Option<PreparedDocumentChunk>,
    last_document_usage: Option<ResourceUsage>,
}

#[wasm_bindgen]
impl DocxArchive {
    /// Copy `data` into WASM once and open the ZIP central directory once.
    /// Resource limits are retained by the package session and applied to every
    /// subsequent logical operation.
    ///
    /// `data` is taken by value (`Vec<u8>`): wasm-bindgen copies the JS `Uint8Array`
    /// once into a WASM-owned buffer and hands that allocation to Rust as this
    /// `Vec`, which `Cursor` then takes by value — a single copy across the
    /// JS→WASM boundary. Taking `&[u8]` would force a second `to_vec()` copy so
    /// the `Cursor` could own its backing store, transiently doubling WASM
    /// linear memory to ~2x the file size during construction.
    #[wasm_bindgen(constructor)]
    pub fn new(
        data: Vec<u8>,
        max_archive_entry_bytes: Option<u64>,
        max_total_inflated_bytes: Option<u64>,
        max_archive_entries: Option<u64>,
    ) -> Result<DocxArchive, JsValue> {
        console_error_panic_hook::set_once();
        let archive = parser::open_document_package(
            data,
            max_archive_entry_bytes,
            max_total_inflated_bytes,
            max_archive_entries,
        )
        .map_err(|error| JsValue::from_str(&error))?;
        Ok(DocxArchive {
            archive,
            document_cursor: None,
            prepared_document_chunk: None,
            last_document_usage: None,
        })
    }

    /// Parse the retained archive and return the model as UTF-8 JSON bytes.
    /// Byte-for-byte identical to `parse_docx` on the same file — same parser,
    /// same serializer, same error strings.
    pub fn parse(&mut self) -> Result<Vec<u8>, JsValue> {
        if self.document_cursor.is_some() || self.prepared_document_chunk.is_some() {
            return Err(JsValue::from_str("a document cursor is active"));
        }
        let doc = self
            .archive
            .run_operation("parse", parser::parse_streamed_compatible)
            .map_err(docx_parser_js_error)?;
        serde_json::to_vec(&doc).map_err(|e| JsValue::from_str(&format!("serialize error: {e}")))
    }

    /// Begin one sequential document-body operation. The complete required part
    /// is preflighted before this returns; body units remain pull-driven.
    pub fn open_document_cursor(
        &mut self,
        operation_id: u32,
        generation: u32,
    ) -> Result<(), JsValue> {
        self.open_document_cursor_inner(operation_id, generation)
            .map_err(|error| JsValue::from_str(&error))
    }

    fn open_document_cursor_inner(
        &mut self,
        operation_id: u32,
        generation: u32,
    ) -> Result<(), String> {
        if operation_id == 0 || generation == 0 {
            return Err("operation id and generation must be positive".to_string());
        }
        if self.document_cursor.is_some() || self.prepared_document_chunk.is_some() {
            return Err("a document cursor is already active".to_string());
        }
        self.last_document_usage = None;
        let zip = &mut self.archive;
        zip.begin_operation("document-cursor")?;
        match parser::DocxBodyCursor::start(zip) {
            Ok(cursor) => {
                self.document_cursor = Some(DocumentCursorState {
                    operation_id,
                    generation,
                    next_sequence: 0,
                    accepted_json_bytes: 0,
                    cursor: Some(cursor),
                    degraded_terminal: None,
                });
                Ok(())
            }
            Err(failure) => {
                if let Err(resource_error) = zip.assert_healthy() {
                    zip.cancel_operation();
                    return Err(resource_error);
                }
                if let Some(error) = failure.word_ilvl_error() {
                    let error = error.to_string();
                    zip.cancel_operation();
                    return Err(error);
                }
                self.document_cursor = Some(DocumentCursorState {
                    operation_id,
                    generation,
                    next_sequence: 0,
                    accepted_json_bytes: 0,
                    cursor: None,
                    degraded_terminal: Some(failure.into_degraded_document()),
                });
                Ok(())
            }
        }
    }

    /// Prepare or replay the next indivisible body/terminal JSON unit. Credit
    /// rejection retains the exact bytes for retry; successful delivery must be
    /// acknowledged before the following sequence can be pulled.
    pub fn pull_document_chunk(
        &mut self,
        sequence: u32,
        operation_id: u32,
        generation: u32,
        byte_credit: u32,
    ) -> Result<Vec<u8>, JsValue> {
        self.pull_document_chunk_inner(sequence, operation_id, generation, byte_credit as usize)
            .map_err(|error| JsValue::from_str(&error))
    }

    fn pull_document_chunk_inner(
        &mut self,
        sequence: u32,
        operation_id: u32,
        generation: u32,
        byte_credit: usize,
    ) -> Result<Vec<u8>, String> {
        if operation_id == 0 || generation == 0 || byte_credit == 0 {
            return Err("operation id, generation, and byte credit must be positive".to_string());
        }
        if let Some(prepared) = self.prepared_document_chunk.as_mut() {
            if (
                prepared.sequence,
                prepared.operation_id,
                prepared.generation,
            ) != (sequence, operation_id, generation)
            {
                return Err("another document unit is awaiting acknowledgement".to_string());
            }
            if prepared.bytes.is_none() {
                return Err("document unit must be acknowledged before another pull".to_string());
            }
            if prepared.byte_length > byte_credit {
                return Err(insufficient_credit_error(prepared.byte_length, byte_credit));
            }
            return Ok(prepared
                .bytes
                .take()
                .expect("prepared document bytes checked above"));
        }

        let identity = self
            .document_cursor
            .as_ref()
            .map(|state| (state.next_sequence, state.operation_id, state.generation))
            .ok_or_else(|| "document cursor is not active".to_string())?;
        if identity != (sequence, operation_id, generation) {
            return Err("document cursor identity or sequence is stale".to_string());
        }
        self.archive.assert_healthy()?;

        let result = (|| -> Result<(Vec<u8>, bool), String> {
            let state = self
                .document_cursor
                .as_mut()
                .expect("document cursor checked above");
            let unit =
                if let Some(cursor) = state.cursor.as_mut() {
                    cursor.next_unit(&mut self.archive)?
                } else {
                    parser::StreamedDocumentUnit::Complete {
                        document: Box::new(state.degraded_terminal.take().ok_or_else(|| {
                            "degraded document terminal is unavailable".to_string()
                        })?),
                    }
                };
            let done = matches!(unit, parser::StreamedDocumentUnit::Complete { .. });
            let reporter = self.archive.operation()?.limit_reporter()?;
            let bytes = serialize_document_unit(&unit, Some(&reporter))?;
            Ok((bytes, done))
        })();
        let (bytes, done) = match result {
            Ok(prepared) => prepared,
            Err(error) => {
                self.cancel_document_cursor();
                return Err(self.archive.assert_healthy().err().unwrap_or(error));
            }
        };
        let byte_length = bytes.len();
        let accepted_json_bytes = self
            .document_cursor
            .as_ref()
            .expect("document cursor remains active while preparing a unit")
            .accepted_json_bytes;
        let accepted_json_bytes_after = accepted_json_bytes.saturating_add(byte_length as u64);
        let retained_model_limit = docx_retained_model_json_limit();
        let retained_limit = self
            .archive
            .operation()?
            .limit_reporter()?
            .observe_hard_limit(
                HardResourceLimitKind::DocxRetainedModelJsonBytes,
                Some("word/document.xml"),
                retained_model_limit,
                accepted_json_bytes_after,
            );
        if let Err(error) = retained_limit {
            self.cancel_document_cursor();
            return Err(self.archive.assert_healthy().err().unwrap_or(error));
        }
        self.prepared_document_chunk = Some(PreparedDocumentChunk {
            operation_id,
            generation,
            sequence,
            bytes: Some(bytes),
            byte_length,
            done,
            accepted_json_bytes_after,
        });
        self.pull_document_chunk_inner(sequence, operation_id, generation, byte_credit)
    }

    pub fn document_chunk_done(&self) -> Result<bool, JsValue> {
        self.prepared_document_chunk
            .as_ref()
            .map(|prepared| prepared.done)
            .ok_or_else(|| JsValue::from_str("no document unit is awaiting acknowledgement"))
    }

    pub fn acknowledge_document_chunk(
        &mut self,
        sequence: u32,
        operation_id: u32,
        generation: u32,
    ) -> Result<(), JsValue> {
        self.acknowledge_document_chunk_inner(sequence, operation_id, generation)
            .map_err(|error| JsValue::from_str(&error))
    }

    fn acknowledge_document_chunk_inner(
        &mut self,
        sequence: u32,
        operation_id: u32,
        generation: u32,
    ) -> Result<(), String> {
        let prepared = self
            .prepared_document_chunk
            .as_ref()
            .ok_or_else(|| "no document unit is awaiting acknowledgement".to_string())?;
        if (
            prepared.sequence,
            prepared.operation_id,
            prepared.generation,
        ) != (sequence, operation_id, generation)
        {
            return Err("document acknowledgement identity is stale or invalid".to_string());
        }
        if prepared.bytes.is_some() {
            return Err("document unit cannot be acknowledged before delivery".to_string());
        }
        self.archive.assert_healthy()?;
        let done = prepared.done;
        let accepted_json_bytes_after = prepared.accepted_json_bytes_after;
        self.prepared_document_chunk.take();
        if done {
            self.document_cursor.take();
            self.last_document_usage = self.archive.operation_usage();
            self.archive.finish_operation()?;
        } else if let Some(state) = self.document_cursor.as_mut() {
            state.next_sequence = state.next_sequence.saturating_add(1);
            state.accepted_json_bytes = accepted_json_bytes_after;
        }
        Ok(())
    }

    pub fn cancel_document_cursor(&mut self) {
        self.prepared_document_chunk.take();
        self.document_cursor.take();
        if let Some(usage) = self.archive.operation_usage() {
            self.last_document_usage = Some(usage);
        }
        self.archive.cancel_operation();
    }

    pub fn close_document_session(&mut self) {
        self.cancel_document_cursor();
    }

    /// Current or most recently completed document-cursor resource checkpoint.
    /// `None` when no checkpoint exists: a package that fails before its
    /// document cursor opens streams a placeholder document without one.
    pub fn document_cursor_resource_usage(&self) -> Result<Option<Vec<u8>>, JsValue> {
        let Some(usage) = self.archive.operation_usage().or(self.last_document_usage) else {
            return Ok(None);
        };
        serde_json::to_vec(&usage)
            .map(Some)
            .map_err(|error| JsValue::from_str(&format!("serialize error: {error}")))
    }

    /// Session-wide archive accounting after parsing or any later lazy part
    /// extraction. Diagnostic only: this is not an allocator-memory estimate.
    pub fn resource_usage(&self) -> Result<Vec<u8>, JsValue> {
        let usage = self.archive.usage();
        serde_json::to_vec(&usage)
            .map_err(|error| JsValue::from_str(&format!("serialize error: {error}")))
    }

    /// Fail cached worker operations after this package session was poisoned.
    pub fn assert_healthy(&self) -> Result<(), JsValue> {
        self.archive
            .assert_healthy()
            .map_err(|e| JsValue::from_str(&e))
    }

    /// Extract raw bytes for one embedded entry (e.g. "word/media/image1.png")
    /// from the retained archive. Twin of the free `extract_image`, but reads
    /// through the already-open archive instead of re-opening it.
    pub fn extract_image(&mut self, path: &str) -> Result<Vec<u8>, JsValue> {
        self.archive
            .run_operation("extract-image", |zip| parser::read_zip_bytes(zip, path))
            .map_err(|e| JsValue::from_str(&e))
    }

    /// GitHub-flavoured markdown projection of the retained archive. Mirrors the
    /// free `docx_to_markdown`.
    pub fn to_markdown(&mut self) -> Result<String, JsValue> {
        self.archive
            .run_operation("markdown", render_markdown_from_zip)
            .map_err(docx_markdown_js_error)
    }
}

/// Native equivalent of `parse_docx` for use from the MCP server.
#[cfg(not(target_arch = "wasm32"))]
pub fn parse_docx_native(data: &[u8]) -> Result<String, String> {
    parser::parse_from_bytes(data)
        .and_then(|doc| serde_json::to_string(&doc).map_err(|e| e.to_string()))
}

/// Parse a docx and project the result to GitHub-flavoured markdown:
/// headings (from outlineLevel), paragraphs with bullet/numbered lists,
/// tables, footnote references collated at the end, and rich-text
/// formatting (bold / italic / strikethrough / hyperlink). Designed for AI
/// agents that need to read content efficiently — discards positioning,
/// section properties, font metrics, drawing shapes.
#[cfg(not(target_arch = "wasm32"))]
pub fn to_markdown_native(data: &[u8]) -> Result<String, String> {
    render_markdown_from_bytes_with_limits(data, None, None)
}

fn docx_parser_js_error(error: String) -> JsValue {
    if error.starts_with("OOXML_RESOURCE_LIMIT:")
        || error.starts_with(numbering::WORD_ILVL_ERROR_PREFIX)
        || ooxml_common::opc::is_not_ooxml_error(&error)
    {
        JsValue::from_str(&error)
    } else {
        JsValue::from_str(&format!("docx-parser error: {error}"))
    }
}

fn docx_markdown_js_error(error: String) -> JsValue {
    JsValue::from_str(&error)
}

fn extract_entry_with_limits(
    data: &[u8],
    path: &str,
    max_archive_entry_bytes: Option<u64>,
    max_total_inflated_bytes: Option<u64>,
    operation: &str,
) -> Result<Vec<u8>, String> {
    let mut zip = parser::open_zip_with_limits(
        data.to_vec(),
        max_archive_entry_bytes,
        max_total_inflated_bytes,
    )?;
    zip.run_operation(operation, |zip| parser::read_zip_bytes(zip, path))
}

#[cfg(test)]
mod tests {
    use super::*;

    struct MarkdownLimitOverride(Option<u64>);

    impl MarkdownLimitOverride {
        fn set(limit: u64) -> Self {
            let previous = DOCX_MARKDOWN_LIMIT_OVERRIDE.with(|value| value.replace(Some(limit)));
            Self(previous)
        }
    }

    impl Drop for MarkdownLimitOverride {
        fn drop(&mut self) {
            DOCX_MARKDOWN_LIMIT_OVERRIDE.with(|value| value.set(self.0));
        }
    }

    fn zip_parts(parts: &[(&str, &[u8])]) -> Vec<u8> {
        use std::io::{Cursor, Write};
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            crate::parser::write_test_content_types(&mut writer);
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);
            for (path, body) in parts {
                writer.start_file(path, options).unwrap();
                writer.write_all(body).unwrap();
            }
            writer.finish().unwrap();
        }
        bytes
    }

    fn forge_declared_size(bytes: &mut [u8], target: &str, declared_size: u32) {
        let target = target.as_bytes();
        let mut cursor = 0;
        let mut local_found = false;
        while cursor + 30 <= bytes.len() {
            if bytes[cursor..cursor + 4] == 0x0403_4b50u32.to_le_bytes() {
                let name_len =
                    u16::from_le_bytes([bytes[cursor + 26], bytes[cursor + 27]]) as usize;
                let extra_len =
                    u16::from_le_bytes([bytes[cursor + 28], bytes[cursor + 29]]) as usize;
                let name_start = cursor + 30;
                let name_end = name_start + name_len;
                if name_end <= bytes.len() && &bytes[name_start..name_end] == target {
                    bytes[cursor + 22..cursor + 26].copy_from_slice(&declared_size.to_le_bytes());
                    local_found = true;
                    break;
                }
                cursor = name_end.saturating_add(extra_len);
            } else {
                cursor += 1;
            }
        }

        cursor = 0;
        let mut central_found = false;
        while cursor + 46 <= bytes.len() {
            if bytes[cursor..cursor + 4] == 0x0201_4b50u32.to_le_bytes() {
                let name_len =
                    u16::from_le_bytes([bytes[cursor + 28], bytes[cursor + 29]]) as usize;
                let extra_len =
                    u16::from_le_bytes([bytes[cursor + 30], bytes[cursor + 31]]) as usize;
                let comment_len =
                    u16::from_le_bytes([bytes[cursor + 32], bytes[cursor + 33]]) as usize;
                let name_start = cursor + 46;
                let name_end = name_start + name_len;
                if name_end <= bytes.len() && &bytes[name_start..name_end] == target {
                    bytes[cursor + 24..cursor + 28].copy_from_slice(&declared_size.to_le_bytes());
                    central_found = true;
                    break;
                }
                cursor = name_end
                    .saturating_add(extra_len)
                    .saturating_add(comment_len);
            } else {
                cursor += 1;
            }
        }
        assert!(local_found && central_found, "target part must be forged");
    }

    fn forged_document_package() -> Vec<u8> {
        let padding = "x".repeat(4096);
        let document = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>{padding}</w:t></w:r></w:p></w:body></w:document>"#
        );
        let rels = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#;
        let mut bytes = zip_parts(&[
            ("word/document.xml", document.as_bytes()),
            ("word/_rels/document.xml.rels", rels),
        ]);
        forge_declared_size(&mut bytes, "word/document.xml", 1);
        bytes
    }

    fn forged_optional_styles_package() -> Vec<u8> {
        let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>"#;
        let rels = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#;
        let padding = "x".repeat(4096);
        let styles = format!(
            r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><!--{padding}--></w:styles>"#
        );
        let mut bytes = zip_parts(&[
            ("word/document.xml", document),
            ("word/_rels/document.xml.rels", rels),
            ("word/styles.xml", styles.as_bytes()),
        ]);
        forge_declared_size(&mut bytes, "word/styles.xml", 1);
        bytes
    }

    fn cursor_test_package() -> Vec<u8> {
        let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
          <w:p><w:r><w:t>A</w:t></w:r></w:p>
          <w:p><w:r><w:t>B</w:t></w:r></w:p>
          <w:sectPr/>
        </w:body></w:document>"#;
        let rels = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#;
        zip_parts(&[
            ("word/document.xml", document),
            ("word/_rels/document.xml.rels", rels),
            ("word/media/later.bin", b"later diagnostic bytes"),
        ])
    }

    fn malformed_required_document_package() -> Vec<u8> {
        zip_parts(&[
            (
                "word/document.xml",
                br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>"#,
            ),
            (
                "word/_rels/document.xml.rels",
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>"#,
            ),
            (
                "word/theme/theme1.xml",
                br#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="Test"><a:majorFont><a:latin typeface="Diagnostic Major"/></a:majorFont><a:minorFont><a:latin typeface="Diagnostic Minor"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>"#,
            ),
        ])
    }

    fn cursor_archive(data: &[u8]) -> DocxArchive {
        DocxArchive {
            archive: parser::open_zip(data.to_vec()).unwrap(),
            document_cursor: None,
            prepared_document_chunk: None,
            last_document_usage: None,
        }
    }

    #[test]
    fn markdown_projection_limit_is_typed_attributed_and_inclusive() {
        let data = cursor_test_package();
        let expected = to_markdown_native(&data).unwrap();
        let exact = expected.len() as u64;

        {
            let _limit = MarkdownLimitOverride::set(exact);
            assert_eq!(to_markdown_native(&data).unwrap(), expected);
        }
        {
            let _limit = MarkdownLimitOverride::set(exact - 1);
            let error = to_markdown_native(&data).unwrap_err();
            assert!(error.starts_with("OOXML_RESOURCE_LIMIT:"), "{error}");
            assert!(error.contains(r#""operation":"markdown""#), "{error}");
            assert!(error.contains(r#""stage":"serialization""#), "{error}");
            assert!(error.contains(r#""resource":"docx-markdown""#), "{error}");
            assert!(error.contains(r#""metric":"bytes""#), "{error}");
            assert!(
                error.contains(&format!(r#""limit":{}"#, exact - 1)),
                "{error}"
            );
        }
    }

    #[test]
    fn document_cursor_replays_credit_and_materializes_the_compatibility_model() {
        let data = cursor_test_package();
        let expected = serde_json::to_value(
            parser::parse_from_bytes_streamed_with_limits(&data, None, None, "expected").unwrap(),
        )
        .unwrap();
        let mut archive = cursor_archive(&data);
        archive.open_document_cursor_inner(7, 3).unwrap();

        let too_small = archive.pull_document_chunk_inner(0, 7, 3, 1).unwrap_err();
        assert!(
            too_small.starts_with("OOXML_INSUFFICIENT_CREDIT:")
                && too_small.contains("\"code\":\"ooxml-insufficient-credit\"")
                && too_small.contains("\"offeredBytes\":1"),
            "{too_small}"
        );
        let exact = archive
            .prepared_document_chunk
            .as_ref()
            .unwrap()
            .byte_length;
        let first = archive.pull_document_chunk_inner(0, 7, 3, exact).unwrap();
        assert_eq!(first.len(), exact);
        assert!(!archive.document_chunk_done().unwrap());
        assert!(archive
            .pull_document_chunk_inner(0, 7, 3, exact)
            .unwrap_err()
            .contains("acknowledged"));
        archive.acknowledge_document_chunk_inner(0, 7, 3).unwrap();

        let mut body = serde_json::from_slice::<serde_json::Value>(&first).unwrap()["body"]
            .as_array()
            .unwrap()
            .clone();
        let mut sequence = 1;
        let terminal = loop {
            let bytes = archive
                .pull_document_chunk_inner(sequence, 7, 3, u32::MAX as usize)
                .unwrap();
            let done = archive.document_chunk_done().unwrap();
            let value: serde_json::Value = serde_json::from_slice(&bytes).unwrap();
            archive
                .acknowledge_document_chunk_inner(sequence, 7, 3)
                .unwrap();
            if done {
                break value["document"].clone();
            }
            body.extend(value["body"].as_array().unwrap().iter().cloned());
            sequence += 1;
        };
        let mut actual = terminal;
        actual["body"] = serde_json::Value::Array(body);
        assert_eq!(actual, expected);
        assert!(archive.document_cursor.is_none());
        assert!(archive.prepared_document_chunk.is_none());
        let usage: serde_json::Value =
            serde_json::from_slice(&archive.document_cursor_resource_usage().unwrap().unwrap())
                .unwrap();
        assert!(usage["operationInflatedBytes"].as_u64().unwrap() > 0);
    }

    #[test]
    fn missing_document_cursor_checkpoint_is_a_typed_absence() {
        // Before any document cursor operation has run there is no
        // checkpoint; the archive reports that as `None`, not as an error.
        // (Unreadable input never constructs an archive: see
        // `non_ooxml_input_is_rejected_at_every_public_boundary`.)
        let data = cursor_test_package();
        let archive = cursor_archive(&data);
        assert!(archive.document_cursor_resource_usage().unwrap().is_none());
    }

    #[test]
    fn cancel_document_cursor_releases_the_package_operation() {
        let data = cursor_test_package();
        let mut archive = cursor_archive(&data);
        archive.open_document_cursor_inner(1, 1).unwrap();
        archive
            .pull_document_chunk_inner(0, 1, 1, u32::MAX as usize)
            .unwrap();
        archive.cancel_document_cursor();
        let parsed = archive
            .archive
            .run_operation("after-cancel", parser::parse_streamed)
            .unwrap();
        assert_eq!(parsed.body.len(), 2);
    }

    #[test]
    fn required_part_failure_is_a_terminal_diagnostic_not_partial_body_success() {
        let data = malformed_required_document_package();
        let mut archive = cursor_archive(&data);
        archive.open_document_cursor_inner(9, 4).unwrap();
        let bytes = archive
            .pull_document_chunk_inner(0, 9, 4, u32::MAX as usize)
            .unwrap();
        assert!(archive.document_chunk_done().unwrap());
        let terminal: serde_json::Value = serde_json::from_slice(&bytes).unwrap();
        assert_eq!(terminal["kind"], "complete");
        assert!(terminal["document"]["body"].as_array().unwrap().is_empty());
        assert!(terminal["document"]["parseError"]
            .as_str()
            .unwrap()
            .starts_with("word/document.xml:"));
        assert_eq!(terminal["document"]["majorFont"], "Diagnostic Major");
        assert_eq!(terminal["document"]["minorFont"], "Diagnostic Minor");
        archive.acknowledge_document_chunk_inner(0, 9, 4).unwrap();

        let mut compatibility = cursor_archive(&data);
        let parsed: serde_json::Value =
            serde_json::from_slice(&compatibility.parse().unwrap()).unwrap();
        assert!(parsed["parseError"]
            .as_str()
            .unwrap()
            .starts_with("word/document.xml:"));
        assert_eq!(parsed["majorFont"], "Diagnostic Major");
        assert_eq!(parsed["minorFont"], "Diagnostic Minor");
    }

    #[test]
    fn session_usage_includes_lazy_parts_read_after_document_parsing() {
        let data = cursor_test_package();
        let mut archive = cursor_archive(&data);
        archive.parse().unwrap();
        let before: serde_json::Value =
            serde_json::from_slice(&archive.resource_usage().unwrap()).unwrap();
        archive.extract_image("word/media/later.bin").unwrap();
        let after: serde_json::Value =
            serde_json::from_slice(&archive.resource_usage().unwrap()).unwrap();
        assert!(
            after["distinctInflatedBytes"].as_u64().unwrap()
                > before["distinctInflatedBytes"].as_u64().unwrap()
        );
        assert!(after["operationInflatedBytes"].as_u64().unwrap() > 0);
    }

    #[test]
    fn retained_document_projection_limit_is_typed_and_poisons_before_transfer() {
        let data = cursor_test_package();
        let mut archive = cursor_archive(&data);
        archive.open_document_cursor_inner(1, 1).unwrap();
        let first = archive
            .pull_document_chunk_inner(0, 1, 1, u32::MAX as usize)
            .unwrap();
        archive.acknowledge_document_chunk_inner(0, 1, 1).unwrap();
        DOCX_RETAINED_MODEL_JSON_LIMIT_OVERRIDE.with(|limit| {
            limit.set(Some(first.len() as u64));
        });

        let error = archive
            .pull_document_chunk_inner(1, 1, 1, u32::MAX as usize)
            .unwrap_err();
        DOCX_RETAINED_MODEL_JSON_LIMIT_OVERRIDE.with(|limit| limit.set(None));
        assert!(error.starts_with("OOXML_RESOURCE_LIMIT:"), "{error}");
        assert!(error.contains("docx-retained-model-json"), "{error}");
        assert_eq!(archive.archive.assert_healthy(), Err(error));
    }

    #[test]
    fn extract_image_reads_entry() {
        use std::io::{Cursor, Write};
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(Cursor::new(&mut buf));
            crate::parser::write_test_content_types(&mut w);
            let o = zip::write::SimpleFileOptions::default();
            w.start_file("word/media/i.png", o).unwrap();
            w.write_all(b"X").unwrap();
            w.finish().unwrap();
        }
        assert_eq!(
            extract_image(&buf, "word/media/i.png", None, None).unwrap(),
            b"X"
        );
    }

    #[test]
    fn document_resource_overrun_is_typed_and_poisons_the_stateful_session() {
        const LIMIT: u64 = 2048;
        let mut zip =
            parser::open_zip_with_limits(forged_document_package(), Some(LIMIT), Some(64 * 1024))
                .expect("forged declaration passes package preflight");
        let first = zip
            .run_operation("parse", parser::parse)
            .expect_err("degraded document must not hide a resource violation");
        let details: serde_json::Value = serde_json::from_str(
            first
                .strip_prefix("OOXML_RESOURCE_LIMIT:")
                .expect("canonical typed resource-limit envelope"),
        )
        .unwrap();
        let violation = &details["details"]["violation"];
        assert_eq!(violation["operation"], "parse");
        assert_eq!(violation["part"], "word/document.xml");
        assert_eq!(violation["metric"], "actual-inflated-bytes");
        assert_eq!(violation["limit"], LIMIT);
        assert_eq!(violation["observed"], LIMIT + 1);

        let later = zip
            .run_operation("markdown", |_| Ok(()))
            .expect_err("poisoned session rejects later operations");
        assert_eq!(later, first);
    }

    #[test]
    fn free_parse_helper_attributes_resource_failure_to_markdown() {
        let error = parser::parse_from_bytes_with_limits(
            &forged_document_package(),
            Some(2048),
            Some(64 * 1024),
            "markdown",
        )
        .expect_err("free parse helper surfaces the package poison");
        let details: serde_json::Value = serde_json::from_str(
            error
                .strip_prefix("OOXML_RESOURCE_LIMIT:")
                .expect("canonical typed resource-limit envelope"),
        )
        .unwrap();
        assert_eq!(details["details"]["violation"]["operation"], "markdown");
        assert_eq!(details["details"]["violation"]["part"], "word/document.xml");
    }

    #[test]
    fn optional_styles_resource_overrun_survives_fallback_and_settles_typed() {
        const LIMIT: u64 = 2048;
        let error = parser::parse_from_bytes_with_limits(
            &forged_optional_styles_package(),
            Some(LIMIT),
            Some(64 * 1024),
            "parse",
        )
        .expect_err("optional styles fallback must not hide package poison");
        let details: serde_json::Value = serde_json::from_str(
            error
                .strip_prefix("OOXML_RESOURCE_LIMIT:")
                .expect("canonical typed resource-limit envelope"),
        )
        .unwrap();
        let violation = &details["details"]["violation"];
        assert_eq!(violation["operation"], "parse");
        assert_eq!(violation["part"], "word/styles.xml");
        assert_eq!(violation["metric"], "actual-inflated-bytes");
        assert_eq!(violation["limit"], LIMIT);
        assert_eq!(violation["observed"], LIMIT + 1);
    }

    #[test]
    fn large_image_stays_lazy_during_parse_but_extract_is_bounded() {
        const TOTAL_LIMIT: u64 = 4096;
        let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>"#;
        let rels = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/large.png"/></Relationships>"#;
        let image = vec![0x5a; 64 * 1024];
        let data = zip_parts(&[
            ("word/document.xml", document),
            ("word/_rels/document.xml.rels", rels),
            ("word/media/large.png", &image),
        ]);

        let doc = parser::parse_from_bytes_with_limits(&data, None, Some(TOTAL_LIMIT), "parse")
            .expect("media existence check must not inflate the image");
        assert!(doc.parse_error.is_none());

        let error = extract_entry_with_limits(
            &data,
            "word/media/large.png",
            None,
            Some(TOTAL_LIMIT),
            "extract-image",
        )
        .expect_err("materializing the large image exceeds the total budget");
        let details: serde_json::Value = serde_json::from_str(
            error
                .strip_prefix("OOXML_RESOURCE_LIMIT:")
                .expect("canonical typed resource-limit envelope"),
        )
        .unwrap();
        let violation = &details["details"]["violation"];
        assert_eq!(violation["operation"], "extract-image");
        assert_eq!(violation["part"], "word/media/large.png");
        assert_eq!(violation["metric"], "distinct-inflated-bytes");
        assert_eq!(violation["limit"], TOTAL_LIMIT);
        assert_eq!(violation["observed"], TOTAL_LIMIT + 1);
    }

    #[test]
    fn damaged_main_part_degrades_inside_an_admitted_package() {
        let malformed = zip_parts(&[
            ("word/document.xml", br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>"#),
            ("word/_rels/document.xml.rels", br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#),
        ]);
        let malformed_doc = parser::parse_from_bytes_with_limits(&malformed, None, None, "parse")
            .expect("malformed document.xml degrades");
        assert!(malformed_doc
            .parse_error
            .as_deref()
            .is_some_and(|error| error.starts_with("word/document.xml:")));
    }

    /// Input that is not a WordprocessingML package fails closed at every public
    /// Rust boundary instead of materializing a placeholder with a 0x0 section.
    #[test]
    fn non_ooxml_input_is_rejected_at_every_public_boundary() {
        use std::io::{Cursor, Write};
        let mut not_opc = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut not_opc));
            writer
                .start_file(
                    "word/document.xml",
                    zip::write::SimpleFileOptions::default(),
                )
                .unwrap();
            writer.write_all(b"<w:document/>").unwrap();
            writer.finish().unwrap();
        }
        let missing_main = zip_parts(&[(
            "_rels/.rels",
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#,
        )]);
        let cases: [(&str, &[u8], &str); 4] = [
            ("garbage", &[1, 2, 3], "not a readable ZIP package"),
            ("empty", &[], "not a readable ZIP package"),
            ("zip-not-opc", &not_opc, "[Content_Types].xml"),
            ("opc-missing-main-part", &missing_main, "word/document.xml"),
        ];
        for (name, bytes, detail) in cases {
            let assert_rejected = |boundary: &str, error: String| {
                assert!(
                    ooxml_common::opc::is_not_ooxml_error(&error) && error.contains(detail),
                    "{name} via {boundary}: {error}"
                );
            };
            assert_rejected(
                "archive",
                parser::open_document_package(bytes.to_vec(), None, None, None)
                    .err()
                    .expect("archive rejects"),
            );
            assert_rejected(
                "parse",
                parser::parse_from_bytes(bytes).expect_err("parse rejects"),
            );
            assert_rejected(
                "streamed parse",
                parser::parse_from_bytes_streamed_with_limits(bytes, None, None, "parse")
                    .expect_err("streamed parse rejects"),
            );
            assert_rejected(
                "markdown",
                render_markdown_from_bytes_with_limits(bytes, None, None)
                    .expect_err("markdown rejects"),
            );
        }
    }

    #[test]
    #[allow(clippy::type_complexity)] // Exact exported ABI shapes are the assertion.
    fn public_signatures_remain_stable() {
        let _: fn(&[u8], Option<u64>, Option<u64>) -> Result<Vec<u8>, JsValue> = parse_docx;
        let _: fn(&[u8], Option<u64>, Option<u64>) -> Result<String, JsValue> = docx_to_markdown;
        let _: fn(&[u8], &str, Option<u64>, Option<u64>) -> Result<Vec<u8>, JsValue> =
            extract_image;
        let _: fn(Vec<u8>, Option<u64>, Option<u64>, Option<u64>) -> Result<DocxArchive, JsValue> =
            DocxArchive::new;
        let _: fn(&mut DocxArchive) -> Result<Vec<u8>, JsValue> = DocxArchive::parse;
        let _: fn(&mut DocxArchive, &str) -> Result<Vec<u8>, JsValue> = DocxArchive::extract_image;
        let _: fn(&mut DocxArchive) -> Result<String, JsValue> = DocxArchive::to_markdown;
        let _: fn(&DocxArchive) -> Result<(), JsValue> = DocxArchive::assert_healthy;
        let _: fn(&[u8]) -> Result<String, String> = parse_docx_native;
        let _: fn(&[u8]) -> Result<String, String> = to_markdown_native;
    }
}
