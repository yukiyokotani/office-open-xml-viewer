use std::io::Cursor;
use wasm_bindgen::prelude::*;

mod drawing_compatibility;
mod markdown;
mod math;
mod numbering;
mod parser;
mod styles;
mod types;
mod xml_util;

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
    max_zip_entry_bytes: Option<u64>,
    max_zip_total_bytes: Option<u64>,
    max_zip_entries: Option<u64>,
) -> Result<Vec<u8>, JsValue> {
    console_error_panic_hook::set_once();
    let _guard =
        ooxml_common::zip::scoped_limits(max_zip_entry_bytes, max_zip_total_bytes, max_zip_entries);
    validate_zip_budget(data).map_err(|e| JsValue::from_str(&format!("docx-parser error: {e}")))?;
    let doc = parser::parse_from_bytes(data)
        .map_err(|e| JsValue::from_str(&format!("docx-parser error: {e}")))?;
    serde_json::to_vec(&doc).map_err(|e| JsValue::from_str(&format!("serialize error: {e}")))
}

/// WASM-callable markdown projection (mirrors `to_markdown_native`). Returns
/// GitHub-flavoured markdown of headings / paragraphs / tables / footnotes,
/// discarding positioning, section properties, fonts, and drawing shapes.
#[wasm_bindgen]
pub fn docx_to_markdown(
    data: &[u8],
    max_zip_entry_bytes: Option<u64>,
    max_zip_total_bytes: Option<u64>,
    max_zip_entries: Option<u64>,
) -> Result<String, JsValue> {
    console_error_panic_hook::set_once();
    let _guard =
        ooxml_common::zip::scoped_limits(max_zip_entry_bytes, max_zip_total_bytes, max_zip_entries);
    validate_zip_budget(data).map_err(|e| JsValue::from_str(&e))?;
    let doc = parser::parse_from_bytes(data).map_err(|e| JsValue::from_str(&e))?;
    Ok(markdown::render_document(&doc))
}

/// Extract raw bytes for a single embedded image entry (e.g.
/// "word/media/image1.png") from a docx zip archive. Thin `wasm_bindgen`
/// wrapper over the shared [`ooxml_common::zip::extract_zip_entry`] reader; used
/// by the main thread to lazily materialize image blobs on demand.
#[wasm_bindgen]
pub fn extract_image(
    data: &[u8],
    path: &str,
    max_zip_entry_bytes: Option<u64>,
    max_zip_total_bytes: Option<u64>,
    max_zip_entries: Option<u64>,
) -> Result<Vec<u8>, JsValue> {
    ooxml_common::zip::extract_zip_entry_with_limits(
        data,
        path,
        max_zip_entry_bytes,
        max_zip_total_bytes,
        max_zip_entries,
    )
    .map_err(|e| JsValue::from_str(&e))
}

/// A stateful handle over an opened docx archive.
///
/// The free functions above (`parse_docx` / `docx_to_markdown` / `extract_image`)
/// each re-copy the whole file into WASM and re-scan the ZIP central directory on
/// every call. A `DocxArchive` copies the bytes into WASM **once** (in `new`) and
/// keeps the opened [`parser::Zip`] alive, so a `parse` followed by any number of
/// `extract_image` calls (the viewer's parse-then-lazily-load-media pattern)
/// pays the copy + open cost a single time. `ZipArchive<Cursor<Vec<u8>>>` is
/// self-contained (it owns its bytes and holds no borrow into the input), which
/// is what lets it live in a `#[wasm_bindgen]` struct field.
///
/// The retained budget mirrors the per-call ZIP guard the free functions
/// install: every method re-installs it for its own scope so the zip-bomb entry
/// cap is honored identically whether callers use the handle or the free
/// functions.
#[wasm_bindgen]
pub struct DocxArchive {
    /// The opened archive, or the container-open error string when the ZIP itself
    /// was truncated / corrupt (RB7 MAJOR). Deferring the failure here — instead of
    /// erroring out of `new` — lets `parse()` return a degraded placeholder
    /// document (symmetric with a corrupt inner part) rather than the constructor
    /// throwing an opaque error the viewer can't turn into a placeholder page.
    archive: Result<parser::Zip, String>,
    budget: ooxml_common::zip::ZipBudget,
}

#[wasm_bindgen]
impl DocxArchive {
    /// Copy `data` into WASM once and open the ZIP central directory once.
    /// ZIP limits are retained and applied on every subsequent method
    /// call (identical semantics to the free functions' `scoped_max` guard).
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
        max_zip_entry_bytes: Option<u64>,
        max_zip_total_bytes: Option<u64>,
        max_zip_entries: Option<u64>,
    ) -> Result<DocxArchive, JsValue> {
        console_error_panic_hook::set_once();
        // RB7 (MAJOR): a truncated / corrupt CONTAINER is deferred, not thrown, so
        // `parse()` can degrade it to a placeholder document instead of the
        // constructor failing with an opaque error.
        let budget = ooxml_common::zip::ZipBudget::new(
            max_zip_entry_bytes,
            max_zip_total_bytes,
            max_zip_entries,
        );
        let _guard = ooxml_common::zip::scoped_budget(&budget);
        let archive = parser::open_zip(data).and_then(|mut archive| {
            ooxml_common::zip::validate_archive_limits(&mut archive)?;
            Ok(archive)
        });
        if let Err(error) = &archive {
            if error.starts_with("ZIP archive exceeds") {
                return Err(JsValue::from_str(error));
            }
        }
        Ok(DocxArchive { archive, budget })
    }

    /// Parse the retained archive and return the model as UTF-8 JSON bytes.
    /// Byte-for-byte identical to `parse_docx` on the same file — same parser,
    /// same serializer, same error strings. When the CONTAINER failed to open
    /// (RB7 MAJOR) the model is a degraded placeholder tagged with the container.
    pub fn parse(&mut self) -> Result<Vec<u8>, JsValue> {
        let _guard = ooxml_common::zip::scoped_budget(&self.budget);
        let doc = match self.archive.as_mut() {
            Ok(zip) => parser::parse(zip)
                .map_err(|e| JsValue::from_str(&format!("docx-parser error: {e}")))?,
            Err(e) => parser::degraded_container_document(e.clone()),
        };
        serde_json::to_vec(&doc).map_err(|e| JsValue::from_str(&format!("serialize error: {e}")))
    }

    /// Extract raw bytes for one embedded entry (e.g. "word/media/image1.png")
    /// from the retained archive. Twin of the free `extract_image`, but reads
    /// through the already-open archive instead of re-opening it. A corrupt
    /// container has no entries, so this surfaces the container-open error.
    pub fn extract_image(&mut self, path: &str) -> Result<Vec<u8>, JsValue> {
        let _guard = ooxml_common::zip::scoped_budget(&self.budget);
        let zip = self
            .archive
            .as_mut()
            .map_err(|e| JsValue::from_str(&format!("docx-parser error: {e}")))?;
        ooxml_common::zip::read_zip_bytes(zip, path).map_err(|e| JsValue::from_str(&e))
    }

    /// GitHub-flavoured markdown projection of the retained archive. Mirrors the
    /// free `docx_to_markdown`. A corrupt container degrades to an empty document.
    pub fn to_markdown(&mut self) -> Result<String, JsValue> {
        let _guard = ooxml_common::zip::scoped_budget(&self.budget);
        let doc = match self.archive.as_mut() {
            Ok(zip) => parser::parse(zip).map_err(|e| JsValue::from_str(&e))?,
            Err(e) => parser::degraded_container_document(e.clone()),
        };
        Ok(markdown::render_document(&doc))
    }
}

/// Native equivalent of `parse_docx` for use from the MCP server.
#[cfg(not(target_arch = "wasm32"))]
pub fn parse_docx_native(data: &[u8]) -> Result<String, String> {
    let _guard = ooxml_common::zip::scoped_limits(None, None, None);
    validate_zip_budget(data)?;
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
    let _guard = ooxml_common::zip::scoped_limits(None, None, None);
    validate_zip_budget(data)?;
    let doc = parser::parse_from_bytes(data)?;
    Ok(markdown::render_document(&doc))
}

fn validate_zip_budget(data: &[u8]) -> Result<(), String> {
    // Preserve the established degraded-container path for malformed ZIPs; a
    // readable central directory is required before a declared-budget violation
    // can be proven.
    let Ok(mut archive) = zip::ZipArchive::new(Cursor::new(data)) else {
        return Ok(());
    };
    ooxml_common::zip::validate_archive_limits(&mut archive)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn extract_image_reads_entry() {
        use std::io::{Cursor, Write};
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(Cursor::new(&mut buf));
            let o = zip::write::SimpleFileOptions::default();
            w.start_file("word/media/i.png", o).unwrap();
            w.write_all(b"X").unwrap();
            w.finish().unwrap();
        }
        assert_eq!(
            extract_image(&buf, "word/media/i.png", None, None, None).unwrap(),
            b"X"
        );
    }

    /// The docx JSON path never hand-assembles `{"error":"…"}` — a message with a
    /// `"` used to produce invalid JSON that made the TS-side `JSON.parse` throw a
    /// confusing SyntaxError. Since RB7 (MAJOR) a non-zip / corrupt CONTAINER no
    /// longer errors at all: `parse_from_bytes` degrades to a placeholder Document
    /// whose `parse_error` field is serialized by serde, so any quotes in the
    /// message are escaped by construction. This pins both facts: the input
    /// degrades (does not panic / error out) AND the placeholder serializes to
    /// valid JSON with the message intact.
    #[test]
    fn parse_non_zip_bytes_degrades_without_json_escaping_hazard() {
        // Not a zip archive — degrades to a placeholder, does not error or panic.
        let doc = parser::parse_from_bytes(&[1, 2, 3])
            .expect("non-zip bytes degrade to a placeholder, not an error");
        let err = doc
            .parse_error
            .as_deref()
            .expect("placeholder carries a container-tagged parse_error");
        assert!(
            err.contains("zip container"),
            "names the container; got {err:?}"
        );
        // serde escapes any quotes: the serialized model is valid JSON and the
        // message round-trips through it unharmed (the old hand-built JSON hazard).
        let json = serde_json::to_string(&doc).expect("serializes to valid JSON");
        let round: serde_json::Value = serde_json::from_str(&json).expect("valid JSON");
        assert_eq!(
            round["parseError"],
            serde_json::Value::String(err.to_string()),
            "parse_error round-trips through serde JSON intact"
        );
    }
}
