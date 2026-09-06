use ooxml_common::content_types::PackageContentTypes;
use ooxml_common::depth::{
    parse_guarded_with_node_limit, xml_dom_complexity_exceeds, GuardedParseError,
};
use ooxml_common::json_measurement::measure_json;
use ooxml_common::ns::{is_p_ns, is_r_ns};
#[cfg(test)]
use ooxml_common::package_session::PackageLimitReporter;
use ooxml_common::package_session::{
    PackageOperation, PackageSessionHandle, RetainedPackageOperation,
};
use ooxml_common::pull::insufficient_credit_error;
use ooxml_common::rels::{parse_rels as parse_opc_rels, relationship_part_path, TargetMode};
use ooxml_common::resource::{
    HardResourceLimitKind, ResourceUsage, HARD_MAX_EMBEDDED_FONT_BYTES,
    HARD_MAX_PPTX_BOOTSTRAP_JSON_BYTES, HARD_MAX_PPTX_BOOTSTRAP_PROJECTION_BYTES,
    HARD_MAX_PPTX_BOOTSTRAP_SLIDES, HARD_MAX_PPTX_MARKDOWN_BYTES,
    HARD_MAX_PPTX_MATERIALIZED_SLIDE_JSON_BYTES, HARD_MAX_PPTX_SHARED_CACHE_ENTRIES,
    HARD_MAX_PPTX_SHARED_CACHE_PROJECTION_BYTES, HARD_MAX_PPTX_SHARED_DEPENDENCY_PROJECTION_BYTES,
    HARD_MAX_PPTX_SHARED_DEPENDENCY_XML_BYTES, HARD_MAX_PPTX_SLIDE_JSON_BYTES,
    HARD_MAX_PPTX_SLIDE_XML_BYTES, HARD_MAX_XML_DOM_COMPLEXITY,
};
use std::collections::HashMap;
#[cfg(test)]
use std::io::Cursor;
use std::io::Read;
use std::rc::Rc;
use wasm_bindgen::prelude::*;

mod table_style_presets;

mod types;
pub(crate) use types::*;

mod markdown;
use markdown::{render_presentation_md, render_slide_md, MarkdownWriter};

mod chart;

mod theme;
use theme::*;

mod fill;
use fill::*;

mod text;
use text::*;

mod shape;
use shape::*;

mod smartart_fallback;

mod master;
use master::*;

// Test-only counter for `roxmltree::Document::parse` calls on the D4 hot paths
// (slide master build, layout, slide XML + decorations). It exists ONLY under
// `cfg(test)` — `note_layout_master_parse()` compiles to nothing in release, so
// this is zero-cost for shipped builds. A regression test uses it to assert that
// a deck whose slides share one layout + one master parses each of those parts a
// bounded number of times (see `parse_count_scales_with_distinct_parts`),
// guarding against re-introducing the per-slide re-parses this change removed.
#[cfg(test)]
thread_local! {
    static LAYOUT_MASTER_PARSE_COUNT: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
    static COMMENT_AUTHORS_LOAD_COUNT: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
    static PPTX_SLIDE_JSON_LIMIT_OVERRIDE: std::cell::Cell<Option<u64>> = const { std::cell::Cell::new(None) };
    static PPTX_SLIDE_XML_LIMIT_OVERRIDE: std::cell::Cell<Option<u64>> = const { std::cell::Cell::new(None) };
    static PPTX_INTERNAL_LIMITS_OVERRIDE: std::cell::Cell<Option<PptxInternalLimits>> = const { std::cell::Cell::new(None) };
    static BOOTSTRAP_OUTPUT_SLIDES_RETAINED: std::cell::Cell<u64> = const { std::cell::Cell::new(0) };
}

/// Non-configurable browser-safety candidates. Every value comes from the one
/// generated policy source and awaits M7 corpus calibration. JSON projections
/// are deterministic structural proxies, not exact heap accounting; JS-side
/// caches and the WASM allocator high-water mark remain later milestones.
#[derive(Clone, Copy)]
struct PptxInternalLimits {
    shared_dependency_xml_bytes: u64,
    xml_dom_complexity: u64,
    shared_dependency_projection_bytes: u64,
    shared_cache_entries: u64,
    shared_cache_projection_bytes: u64,
    bootstrap_slides: u64,
    bootstrap_projection_bytes: u64,
    bootstrap_json_bytes: u64,
    markdown_bytes: u64,
    materialized_slide_json_bytes: u64,
}

impl Default for PptxInternalLimits {
    fn default() -> Self {
        Self {
            shared_dependency_xml_bytes: HARD_MAX_PPTX_SHARED_DEPENDENCY_XML_BYTES,
            xml_dom_complexity: HARD_MAX_XML_DOM_COMPLEXITY,
            shared_dependency_projection_bytes: HARD_MAX_PPTX_SHARED_DEPENDENCY_PROJECTION_BYTES,
            shared_cache_entries: HARD_MAX_PPTX_SHARED_CACHE_ENTRIES,
            shared_cache_projection_bytes: HARD_MAX_PPTX_SHARED_CACHE_PROJECTION_BYTES,
            bootstrap_slides: HARD_MAX_PPTX_BOOTSTRAP_SLIDES,
            bootstrap_projection_bytes: HARD_MAX_PPTX_BOOTSTRAP_PROJECTION_BYTES,
            bootstrap_json_bytes: HARD_MAX_PPTX_BOOTSTRAP_JSON_BYTES,
            markdown_bytes: HARD_MAX_PPTX_MARKDOWN_BYTES,
            materialized_slide_json_bytes: HARD_MAX_PPTX_MATERIALIZED_SLIDE_JSON_BYTES,
        }
    }
}

fn pptx_internal_limits() -> PptxInternalLimits {
    #[cfg(test)]
    if let Some(limits) = PPTX_INTERNAL_LIMITS_OVERRIDE.with(std::cell::Cell::get) {
        return limits;
    }
    PptxInternalLimits::default()
}

fn pptx_slide_json_limit() -> u64 {
    #[cfg(test)]
    if let Some(limit) = PPTX_SLIDE_JSON_LIMIT_OVERRIDE.with(std::cell::Cell::get) {
        return limit;
    }
    HARD_MAX_PPTX_SLIDE_JSON_BYTES
}

fn pptx_slide_xml_limit() -> u64 {
    #[cfg(test)]
    if let Some(limit) = PPTX_SLIDE_XML_LIMIT_OVERRIDE.with(std::cell::Cell::get) {
        return limit;
    }
    HARD_MAX_PPTX_SLIDE_XML_BYTES
}

/// Defense-in-depth DOM parse for PPTX XML that has already passed the typed,
/// attributable lexical complexity preflight in [`read_bounded_pptx_xml`].
/// Keeping this wrapper PPTX-local prevents an uncalibrated node ceiling from
/// changing DOCX/XLSX's shared `parse_guarded` compatibility behavior.
fn parse_preflighted_pptx_xml(xml: &str) -> Result<roxmltree::Document<'_>, GuardedParseError> {
    let nodes_limit = u32::try_from(pptx_internal_limits().xml_dom_complexity).unwrap_or(u32::MAX);
    parse_guarded_with_node_limit(xml, nodes_limit)
}

/// Increment the D4 parse counter (no-op unless `cfg(test)`). Call immediately
/// before parsing a slide-master / layout / slide XML on the pagination path.
#[inline(always)]
fn note_layout_master_parse() {
    #[cfg(test)]
    LAYOUT_MASTER_PARSE_COUNT.with(|c| c.set(c.get() + 1));
}

#[inline(always)]
fn note_comment_authors_load() {
    #[cfg(test)]
    COMMENT_AUTHORS_LOAD_COUNT.with(|c| c.set(c.get() + 1));
}

#[inline(always)]
fn note_bootstrap_output_slide_retained() {
    #[cfg(test)]
    BOOTSTRAP_OUTPUT_SLIDES_RETAINED.with(|count| count.set(count.get() + 1));
}

// ===========================
//  Public WASM entry points
// ===========================

/// Parse a pptx archive and return the model as UTF-8 JSON **bytes**.
///
/// Returning `Vec<u8>` (a fresh copy on the JS side) instead of `String` keeps
/// the model out of the JsString/UTF-16 representation: the worker forwards the
/// resulting `ArrayBuffer` to the main thread as a transferable and the main
/// thread does a single `TextDecoder.decode` + `JSON.parse`, collapsing three
/// serializations (Rust String → JsString → structured clone) into one decode.
#[wasm_bindgen]
pub fn parse_pptx(
    data: &[u8],
    max_archive_entry_bytes: Option<u64>,
    max_total_inflated_bytes: Option<u64>,
) -> Result<Vec<u8>, JsValue> {
    console_error_panic_hook::set_once();
    let presentation = parse_presentation_from_bytes_with_limits(
        data,
        max_archive_entry_bytes,
        max_total_inflated_bytes,
        "parse",
    )
    .map_err(pptx_parser_js_error)?;
    serde_json::to_vec(&presentation)
        .map_err(|e| JsValue::from_str(&format!("serialize error: {e}")))
}

/// WASM-callable markdown projection. Shares the body of `to_markdown_native`
/// so the browser / Node WASM path and the native mcp-server path stay in
/// lock-step. See `to_markdown_native` for the design rationale.
#[wasm_bindgen]
pub fn pptx_to_markdown(
    data: &[u8],
    max_archive_entry_bytes: Option<u64>,
    max_total_inflated_bytes: Option<u64>,
) -> Result<String, JsValue> {
    console_error_panic_hook::set_once();
    render_markdown_from_bytes_with_limits(data, max_archive_entry_bytes, max_total_inflated_bytes)
        .map_err(pptx_parser_js_error)
}

/// Native equivalent of `parse_pptx` for use from the MCP server.
pub fn parse_pptx_native(data: &[u8]) -> Result<String, String> {
    let presentation = parse_presentation_from_bytes(data).map_err(|e| e.to_string())?;
    serde_json::to_string(&presentation).map_err(|e| e.to_string())
}

/// Parse a pptx and project the result to GitHub-flavoured markdown,
/// preserving textual / semantic structure (headings, bullets, tables, charts,
/// notes, comments) and discarding presentation details (geometry, fills,
/// strokes, effects, theme inheritance details). Designed for AI agents that
/// need to read content efficiently — typical 10-30× token reduction vs. the
/// raw JSON of `parse_pptx_native`.
pub fn to_markdown_native(data: &[u8]) -> Result<String, String> {
    render_markdown_from_bytes_with_limits(data, None, None)
}

fn pptx_parser_js_error(error: String) -> JsValue {
    if error.starts_with("OOXML_RESOURCE_LIMIT:") {
        JsValue::from_str(&error)
    } else {
        JsValue::from_str(&format!("pptx-parser error: {error}"))
    }
}

/// Extract raw bytes for a single entry (e.g. "ppt/media/media2.mp4") from a
/// pptx zip archive. Used by the main thread to materialize media blobs for
/// interactive playback without re-parsing the whole file.
#[wasm_bindgen]
pub fn extract_media(
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
        "extract-media",
    )
    .map_err(|e| JsValue::from_str(&e))
}

/// Extract raw bytes for a single embedded image entry (e.g.
/// "ppt/media/image1.png") from a pptx zip archive. Used by the main thread to
/// lazily materialize image blobs on demand through a bounded package operation.
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

/// Extract one font part referenced by `p:embeddedFontLst`.
#[wasm_bindgen]
pub fn extract_font(
    data: &[u8],
    path: &str,
    max_archive_entry_bytes: Option<u64>,
    max_total_inflated_bytes: Option<u64>,
) -> Result<Vec<u8>, JsValue> {
    let zip = open_zip_with_limits(
        data.to_vec(),
        max_archive_entry_bytes,
        max_total_inflated_bytes,
    )
    .map_err(|e| JsValue::from_str(&e))?;
    zip.read_font_part(path).map_err(|e| JsValue::from_str(&e))
}

/// A stateful handle over an opened pptx archive.
///
/// The free functions above (`parse_pptx` / `pptx_to_markdown` / `extract_media`
/// / `extract_image` / `extract_font`) each re-copy the whole file into WASM and re-scan the ZIP
/// central directory on every call. A `PptxArchive` copies the bytes into WASM
/// **once** (in `new`) and keeps the opened [`PptxZip`] session alive, so a `parse`
/// followed by any number of `extract_media` / `extract_image` calls (the
/// viewer's parse-then-lazily-load-media pattern) pays the copy + open cost a
/// single time. The session owns the source bytes, validated central-directory
/// index, resource governor, and first package-wide poison error.
#[wasm_bindgen]
pub struct PptxArchive {
    /// The opened archive, or the container-open error string when the ZIP itself
    /// was truncated / corrupt (#774, RB7 MAJOR). Deferring the failure here —
    /// instead of erroring out of `new` — lets `parse()` return a degraded
    /// placeholder presentation (symmetric with a corrupt inner slide) rather than
    /// the constructor throwing an opaque error the viewer can't turn into a
    /// placeholder slide.
    archive: Result<PptxZip, String>,
    presentation: Option<PresentationShared>,
    prepared_slide: Option<PreparedSlide>,
    last_slide_usage: Option<ResourceUsage>,
}

struct PreparedSlide {
    index: u32,
    operation_id: u32,
    generation: u32,
    bytes: Option<Vec<u8>>,
    byte_length: usize,
    journal: Option<SlideCacheJournal>,
}

#[derive(Default)]
struct SlideCacheJournal {
    inserted_master_keys: Vec<String>,
    inserted_layout_keys: Vec<String>,
    inserted_layout_source_keys: Vec<String>,
    had_comment_authors: bool,
    had_modern_comment_authors: bool,
    had_no_master_bundle: bool,
    projected_entries: u64,
    projected_bytes: u64,
}

impl SlideCacheJournal {
    fn begin(shared: &PresentationShared) -> Self {
        Self {
            had_comment_authors: shared.comment_authors.is_some(),
            had_modern_comment_authors: shared.modern_comment_authors.is_some(),
            had_no_master_bundle: shared.no_master_bundle.is_some(),
            ..Self::default()
        }
    }

    fn rollback(self, shared: &mut PresentationShared) {
        for key in self.inserted_master_keys {
            shared.master_cache.remove(&key);
        }
        for key in self.inserted_layout_keys {
            shared.layout_cache.remove(&key);
        }
        for key in self.inserted_layout_source_keys {
            shared.layout_source_cache.remove(&key);
        }
        if !self.had_comment_authors {
            shared.comment_authors = None;
        }
        if !self.had_modern_comment_authors {
            shared.modern_comment_authors = None;
        }
        if !self.had_no_master_bundle {
            shared.no_master_bundle = None;
        }
    }

    fn commit(self, shared: &mut PresentationShared) {
        shared.cache_usage.entries = shared
            .cache_usage
            .entries
            .checked_add(self.projected_entries)
            .expect("preflighted PPTX cache entry accounting");
        shared.cache_usage.projected_bytes = shared
            .cache_usage
            .projected_bytes
            .checked_add(self.projected_bytes)
            .expect("preflighted PPTX cache byte accounting");
    }
}

#[derive(Default)]
struct SharedCacheUsage {
    entries: u64,
    projected_bytes: u64,
}

fn observe_shared_cache_candidate<T: serde::Serialize>(
    reporter: &ooxml_common::package_session::PackageLimitReporter,
    part: Option<&str>,
    value: &T,
    usage: &mut SharedCacheUsage,
    journal: Option<&mut SlideCacheJournal>,
) -> Result<(), String> {
    // Include the cache identity as well as the owned value. Entry-count caps
    // cover map/node overhead; this projection covers retained key/value data.
    let candidate = measure_json(&(part, value))?.json_bytes;
    let limits = pptx_internal_limits();
    reporter.observe_hard_limit(
        HardResourceLimitKind::PptxSharedDependencyProjectionBytes,
        part,
        limits.shared_dependency_projection_bytes,
        candidate,
    )?;

    let pending_entries = journal.as_ref().map_or(0, |entry| entry.projected_entries);
    let pending_bytes = journal.as_ref().map_or(0, |entry| entry.projected_bytes);
    let projected_entries = usage
        .entries
        .checked_add(pending_entries)
        .and_then(|value| value.checked_add(1))
        .unwrap_or(u64::MAX);
    let projected_bytes = usage
        .projected_bytes
        .checked_add(pending_bytes)
        .and_then(|value| value.checked_add(candidate))
        .unwrap_or(u64::MAX);
    reporter.observe_hard_limit(
        HardResourceLimitKind::PptxSharedCacheEntries,
        part,
        limits.shared_cache_entries,
        projected_entries,
    )?;
    reporter.observe_hard_limit(
        HardResourceLimitKind::PptxSharedCacheProjectionBytes,
        part,
        limits.shared_cache_projection_bytes,
        projected_bytes,
    )?;

    if let Some(journal) = journal {
        journal.projected_entries += 1;
        journal.projected_bytes += candidate;
    } else {
        usage.entries = projected_entries;
        usage.projected_bytes = projected_bytes;
    }
    Ok(())
}

#[derive(serde::Serialize)]
#[serde(rename_all = "camelCase")]
struct PresentationBootstrap {
    slide_count: usize,
    slide_width: i64,
    slide_height: i64,
    default_text_color: Option<String>,
    major_font: Option<String>,
    minor_font: Option<String>,
    hlink_color: Option<String>,
    fol_hlink_color: Option<String>,
    embedded_fonts: Vec<PptxEmbeddedFontRef>,
    slides: Vec<BootstrapSlide>,
}

#[derive(Debug, Clone, PartialEq, Eq, serde::Serialize)]
#[serde(rename_all = "camelCase")]
struct PptxEmbeddedFontRef {
    font_name: String,
    style: EmbeddedFontStyle,
    part_path: String,
    content_type: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, serde::Serialize)]
#[serde(rename_all = "camelCase")]
enum EmbeddedFontStyle {
    Regular,
    Bold,
    Italic,
    BoldItalic,
}

#[derive(serde::Serialize)]
#[serde(rename_all = "camelCase")]
struct BootstrapSlide {
    index: usize,
    #[serde(skip_serializing_if = "Option::is_none")]
    part_name: Option<String>,
}

#[derive(serde::Serialize)]
#[serde(rename_all = "camelCase")]
struct PresentationBootstrapProjection<'a> {
    slide_count: usize,
    slide_width: i64,
    slide_height: i64,
    default_text_color: Option<&'a str>,
    major_font: Option<&'a str>,
    minor_font: Option<&'a str>,
    hlink_color: Option<&'a str>,
    fol_hlink_color: Option<&'a str>,
    embedded_fonts: &'a [PptxEmbeddedFontRef],
    slides: &'a [BootstrapSlideProjection<'a>],
}

#[derive(serde::Serialize)]
#[serde(rename_all = "camelCase")]
struct BootstrapSlideProjection<'a> {
    index: usize,
    #[serde(skip_serializing_if = "Option::is_none")]
    part_name: Option<&'a str>,
}

fn serialize_presentation_bootstrap(
    shared: &PresentationShared,
    reporter: &ooxml_common::package_session::PackageLimitReporter,
) -> Result<Vec<u8>, String> {
    let limits = pptx_internal_limits();
    let empty_slides: [BootstrapSlideProjection<'_>; 0] = [];
    let base = PresentationBootstrapProjection {
        slide_count: shared.slide_descriptors.len(),
        slide_width: shared.slide_width,
        slide_height: shared.slide_height,
        default_text_color: shared.theme.get("dk1").map(String::as_str),
        major_font: shared.theme.get("+mj-lt").map(String::as_str),
        minor_font: shared.theme.get("+mn-lt").map(String::as_str),
        hlink_color: shared.theme.get("hlink").map(String::as_str),
        fol_hlink_color: shared.theme.get("folHlink").map(String::as_str),
        embedded_fonts: &shared.embedded_fonts,
        slides: &empty_slides,
    };
    let mut projected_json_bytes = measure_json(&base)?.json_bytes;
    reporter.observe_hard_limit(
        HardResourceLimitKind::PptxBootstrapJsonBytes,
        Some("ppt/presentation.xml"),
        limits.bootstrap_json_bytes,
        projected_json_bytes,
    )?;

    let mut slides = Vec::with_capacity(shared.slide_descriptors.len());
    for descriptor in &shared.slide_descriptors {
        // `resolve_path` owns at most one candidate here. The exact candidate
        // projection is admitted before it is pushed into the retained output,
        // preventing repeated relationship targets from multiplying unchecked.
        let part_name = descriptor
            .relationship_id
            .as_ref()
            .and_then(|id| shared.pres_rels.get(id))
            .map(|target| resolve_path("ppt", target));
        let candidate = BootstrapSlideProjection {
            index: descriptor.index,
            part_name: part_name.as_deref(),
        };
        let candidate_bytes = measure_json(&candidate)?.json_bytes;
        let comma = u64::from(!slides.is_empty());
        let next_projection = projected_json_bytes
            .saturating_add(comma)
            .saturating_add(candidate_bytes);
        reporter.observe_hard_limit(
            HardResourceLimitKind::PptxBootstrapJsonBytes,
            Some("ppt/presentation.xml"),
            limits.bootstrap_json_bytes,
            next_projection,
        )?;
        slides.push(BootstrapSlide {
            index: descriptor.index,
            part_name,
        });
        note_bootstrap_output_slide_retained();
        projected_json_bytes = next_projection;
    }

    let bootstrap = PresentationBootstrap {
        slide_count: shared.slide_descriptors.len(),
        slide_width: shared.slide_width,
        slide_height: shared.slide_height,
        default_text_color: shared.theme.get("dk1").cloned(),
        major_font: shared.theme.get("+mj-lt").cloned(),
        minor_font: shared.theme.get("+mn-lt").cloned(),
        hlink_color: shared.theme.get("hlink").cloned(),
        fol_hlink_color: shared.theme.get("folHlink").cloned(),
        embedded_fonts: shared.embedded_fonts.clone(),
        slides,
    };
    let final_json_bytes = measure_json(&bootstrap)?.json_bytes;
    debug_assert_eq!(final_json_bytes, projected_json_bytes);
    reporter.observe_hard_limit(
        HardResourceLimitKind::PptxBootstrapJsonBytes,
        Some("ppt/presentation.xml"),
        limits.bootstrap_json_bytes,
        final_json_bytes,
    )?;
    serde_json::to_vec(&bootstrap).map_err(|error| format!("serialize error: {error}"))
}

#[wasm_bindgen]
impl PptxArchive {
    fn ensure_presentation(&mut self) -> Result<(), String> {
        if self.presentation.is_none() {
            let zip = self
                .archive
                .as_mut()
                .map_err(|error| format!("pptx-parser error: {error}"))?;
            self.presentation = Some(bootstrap_presentation(zip).map_err(|e| e.to_string())?);
        }
        Ok(())
    }

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
    ) -> Result<PptxArchive, JsValue> {
        console_error_panic_hook::set_once();
        // #774 (RB7 MAJOR): a truncated / corrupt CONTAINER is deferred, not
        // thrown, so `parse()` can degrade it to a placeholder presentation
        // instead of the constructor failing with an opaque error.
        let archive = open_zip_with_policy(
            data,
            max_archive_entry_bytes,
            max_total_inflated_bytes,
            max_archive_entries,
        );
        if let Err(error) = &archive {
            if error.starts_with("OOXML_RESOURCE_LIMIT:") {
                return Err(JsValue::from_str(error));
            }
        }
        Ok(PptxArchive {
            archive,
            presentation: None,
            prepared_slide: None,
            last_slide_usage: None,
        })
    }

    /// Parse the retained archive and return the model as UTF-8 JSON bytes.
    /// Byte-for-byte identical to `parse_pptx` on the same file. When the
    /// CONTAINER failed to open (#774) the model is a degraded placeholder
    /// presentation tagged with the container.
    pub fn parse(&mut self) -> Result<Vec<u8>, JsValue> {
        self.parse_inner().map_err(pptx_parser_js_error)
    }

    fn parse_inner(&mut self) -> Result<Vec<u8>, String> {
        if self.prepared_slide.is_some() {
            return Err("a slide unit is awaiting acknowledgement".to_string());
        }
        if let Err(error) = &self.archive {
            return serde_json::to_vec(&degraded_container_presentation(error.clone()))
                .map_err(|e| format!("serialize error: {e}"));
        }
        let zip = self.archive.as_mut().expect("container open checked above");
        zip.begin_operation("parse")?;
        let result = (|| -> Result<Presentation, String> {
            self.ensure_presentation()?;
            let mut shared = self.presentation.take().expect("presentation loaded above");
            let zip = self.archive.as_mut().expect("container open checked above");
            let mut slides = Vec::with_capacity(shared.slide_descriptors.len());
            for index in 0..shared.slide_descriptors.len() {
                slides
                    .push(produce_slide_unit(index, &mut shared, zip).map_err(|e| e.to_string())?);
            }
            shared.finish(slides).map_err(|e| e.to_string())
        })();
        let zip = self.archive.as_mut().expect("container open checked above");
        let presentation = settle_pptx_operation(zip, result)?;
        serde_json::to_vec(&presentation).map_err(|e| format!("serialize error: {e}"))
    }

    /// Open presentation metadata without parsing any slide parts. Hidden state,
    /// notes, and comments remain slide-local because PresentationML stores them
    /// in `p:sld` and slide relationships (ECMA-376 Part 1 §§19.2-19.3).
    pub fn presentation_bootstrap(&mut self) -> Result<Vec<u8>, JsValue> {
        self.presentation_bootstrap_inner()
            .map_err(|error| JsValue::from_str(&error))
    }

    fn presentation_bootstrap_inner(&mut self) -> Result<Vec<u8>, String> {
        if self.prepared_slide.is_some() {
            return Err("a slide unit is awaiting acknowledgement".to_string());
        }
        if self.archive.is_err() {
            let bootstrap = PresentationBootstrap {
                slide_count: 1,
                slide_width: 12_192_000,
                slide_height: 6_858_000,
                default_text_color: None,
                major_font: None,
                minor_font: None,
                hlink_color: None,
                fol_hlink_color: None,
                embedded_fonts: Vec::new(),
                slides: vec![BootstrapSlide {
                    index: 0,
                    part_name: None,
                }],
            };
            return serde_json::to_vec(&bootstrap)
                .map_err(|error| format!("serialize error: {error}"));
        }
        let zip = self.archive.as_mut().expect("container checked above");
        zip.begin_operation("presentation-bootstrap")?;
        let result = (|| -> Result<Vec<u8>, String> {
            self.ensure_presentation()?;
            let reporter = self
                .archive
                .as_mut()
                .expect("container open checked above")
                .operation()?
                .limit_reporter()?;
            let shared = self
                .presentation
                .as_ref()
                .expect("presentation loaded above");
            serialize_presentation_bootstrap(shared, &reporter)
        })();
        let zip = self.archive.as_mut().expect("container open checked above");
        settle_pptx_operation(zip, result)
    }

    /// Prepare or replay one complete random-access slide. The unit is never
    /// split: insufficient byte credit returns an error while retaining the
    /// exact prepared bytes for a later retry.
    pub fn pull_slide(
        &mut self,
        slide_index: u32,
        operation_id: u32,
        generation: u32,
        byte_credit: u32,
    ) -> Result<Vec<u8>, JsValue> {
        self.pull_slide_inner(slide_index, operation_id, generation, byte_credit as usize)
            .map_err(|error| JsValue::from_str(&error))
    }

    fn pull_slide_inner(
        &mut self,
        slide_index: u32,
        operation_id: u32,
        generation: u32,
        byte_credit: usize,
    ) -> Result<Vec<u8>, String> {
        if operation_id == 0 || generation == 0 || byte_credit == 0 {
            return Err("operation id, generation, and byte credit must be positive".to_string());
        }
        if self.prepared_slide.is_some() {
            if let Ok(zip) = self.archive.as_ref() {
                if let Err(error) = zip.assert_healthy() {
                    self.cancel_slide();
                    return Err(error);
                }
            }
            let prepared = self.prepared_slide.as_ref().expect("checked above");
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
            return Ok(self
                .prepared_slide
                .as_mut()
                .expect("prepared checked above")
                .bytes
                .take()
                .expect("prepared bytes checked above"));
        }
        if let Err(container_error) = &self.archive {
            if slide_index != 0 {
                return Err(format!("slide index {slide_index} is out of bounds"));
            }
            let slide = degraded_container_presentation(container_error.clone())
                .slides
                .into_iter()
                .next()
                .expect("degraded presentation owns one slide");
            let measured = measure_json(&slide)?.json_bytes;
            if measured > HARD_MAX_PPTX_SLIDE_JSON_BYTES {
                return Err("degraded slide exceeds the PPTX slide JSON ceiling".to_string());
            }
            let bytes =
                serde_json::to_vec(&slide).map_err(|error| format!("serialize error: {error}"))?;
            let byte_length = bytes.len();
            self.prepared_slide = Some(PreparedSlide {
                index: slide_index,
                operation_id,
                generation,
                bytes: Some(bytes),
                byte_length,
                journal: None,
            });
            return self.pull_slide_inner(slide_index, operation_id, generation, byte_credit);
        }
        // Bootstrap is a separate committed package operation. A canceled slide
        // can therefore roll back only slide-local cache insertions without
        // discarding presentation metadata or retaining uncommitted reads.
        if self.presentation.is_none() {
            let zip = self
                .archive
                .as_mut()
                .map_err(|error| format!("pptx-parser error: {error}"))?;
            zip.begin_operation("presentation-bootstrap")?;
            let bootstrap = self.ensure_presentation();
            let zip = self.archive.as_mut().expect("container open checked above");
            settle_pptx_operation(zip, bootstrap)?;
        }
        let zip = self
            .archive
            .as_mut()
            .map_err(|error| format!("pptx-parser error: {error}"))?;
        zip.begin_operation("slide-cursor")?;
        let mut journal = SlideCacheJournal::begin(
            self.presentation
                .as_ref()
                .expect("presentation bootstrapped above"),
        );
        let result = (|| -> Result<Vec<u8>, String> {
            self.ensure_presentation()?;
            let shared = self
                .presentation
                .as_mut()
                .expect("presentation loaded above");
            let zip = self.archive.as_mut().expect("container open checked above");
            let produced = produce_slide_unit_with_journal(
                slide_index as usize,
                shared,
                zip,
                Some(&mut journal),
            )
            .map_err(|error| error.to_string())?;
            zip.assert_healthy()?;
            serde_json::to_vec(&produced.slide).map_err(|error| format!("serialize error: {error}"))
        })();
        let bytes = match result {
            Ok(bytes) => bytes,
            Err(error) => {
                let zip = self.archive.as_mut().expect("container open checked above");
                zip.cancel_operation();
                journal.rollback(
                    self.presentation
                        .as_mut()
                        .expect("presentation bootstrapped above"),
                );
                return Err(zip.assert_healthy().err().unwrap_or(error));
            }
        };
        let byte_length = bytes.len();
        self.prepared_slide = Some(PreparedSlide {
            index: slide_index,
            operation_id,
            generation,
            bytes: Some(bytes),
            byte_length,
            journal: Some(journal),
        });
        self.pull_slide_inner(slide_index, operation_id, generation, byte_credit)
    }

    pub fn acknowledge_slide(&mut self, operation_id: u32, generation: u32) -> Result<(), JsValue> {
        self.acknowledge_slide_inner(operation_id, generation)
            .map_err(|error| JsValue::from_str(&error))
    }

    fn acknowledge_slide_inner(
        &mut self,
        operation_id: u32,
        generation: u32,
    ) -> Result<(), String> {
        let prepared = self
            .prepared_slide
            .as_ref()
            .ok_or_else(|| "no slide unit is awaiting acknowledgement".to_string())?;
        if (prepared.operation_id, prepared.generation) != (operation_id, generation) {
            return Err("slide acknowledgement identity is stale or invalid".to_string());
        }
        if prepared.bytes.is_some() {
            return Err("slide unit cannot be acknowledged before delivery".to_string());
        }
        if self.archive.is_err() {
            self.prepared_slide.take();
            return Ok(());
        }
        if let Err(error) = self
            .archive
            .as_ref()
            .map_err(|error| error.clone())?
            .assert_healthy()
        {
            self.cancel_slide();
            return Err(error);
        }
        let zip = self.archive.as_mut().map_err(|error| error.clone())?;
        self.last_slide_usage = zip.operation.usage();
        if let Err(error) = zip.finish_operation() {
            self.cancel_slide();
            return Err(error);
        }
        let prepared = self
            .prepared_slide
            .take()
            .expect("prepared slide was validated above");
        if let (Some(journal), Some(shared)) = (prepared.journal, self.presentation.as_mut()) {
            journal.commit(shared);
        }
        Ok(())
    }

    pub fn cancel_slide(&mut self) {
        if let Some(prepared) = self.prepared_slide.take() {
            if let (Some(journal), Some(shared)) = (prepared.journal, self.presentation.as_mut()) {
                journal.rollback(shared);
            }
        }
        if let Ok(zip) = self.archive.as_mut() {
            self.last_slide_usage = zip.operation.usage();
            zip.cancel_operation();
        }
    }

    pub fn close_presentation_session(&mut self) {
        self.cancel_slide();
        self.presentation.take();
    }

    pub fn slide_cursor_resource_usage(&self) -> Result<Vec<u8>, JsValue> {
        let usage = self
            .archive
            .as_ref()
            .ok()
            .and_then(|zip| zip.operation.usage())
            .or(self.last_slide_usage)
            .ok_or_else(|| JsValue::from_str("slide cursor usage is unavailable"))?;
        serde_json::to_vec(&usage)
            .map_err(|error| JsValue::from_str(&format!("serialize error: {error}")))
    }

    /// Session-wide archive accounting after bootstrap, slide parsing, or any
    /// later lazy image/media extraction. Diagnostic only: this is not an
    /// allocator-memory estimate.
    pub fn resource_usage(&self) -> Result<Vec<u8>, JsValue> {
        let usage = self
            .archive
            .as_ref()
            .map(PptxZip::usage)
            .map_err(|_| JsValue::from_str("pptx resource usage is unavailable"))?;
        serde_json::to_vec(&usage)
            .map_err(|error| JsValue::from_str(&format!("serialize error: {error}")))
    }

    /// Fail cached worker operations after this package session was poisoned.
    pub fn assert_healthy(&self) -> Result<(), JsValue> {
        match &self.archive {
            Ok(zip) => zip.assert_healthy().map_err(|e| JsValue::from_str(&e)),
            Err(_) => Ok(()),
        }
    }

    /// Extract raw bytes for one media entry (e.g. "ppt/media/media2.mp4") from
    /// the retained archive. Twin of the free `extract_media`, but reads through
    /// the already-open archive instead of re-opening it. A corrupt container has
    /// no entries, so this surfaces the container-open error.
    pub fn extract_media(&mut self, path: &str) -> Result<Vec<u8>, JsValue> {
        let zip = self
            .archive
            .as_ref()
            .map_err(|e| JsValue::from_str(&format!("pptx-parser error: {e}")))?;
        zip.read_part_in_independent_operation("extract-media", path)
            .map_err(|e| JsValue::from_str(&e))
    }

    /// Extract raw bytes for one embedded image entry (e.g.
    /// "ppt/media/image1.png") from the retained archive. Twin of the free
    /// `extract_image`. A corrupt container has no entries, so this surfaces the
    /// container-open error.
    pub fn extract_image(&mut self, path: &str) -> Result<Vec<u8>, JsValue> {
        let zip = self
            .archive
            .as_ref()
            .map_err(|e| JsValue::from_str(&format!("pptx-parser error: {e}")))?;
        zip.read_part_in_independent_operation("extract-image", path)
            .map_err(|e| JsValue::from_str(&e))
    }

    /// Extract raw bytes for one font part retained by `p:embeddedFontLst`.
    pub fn extract_font(&mut self, path: &str) -> Result<Vec<u8>, JsValue> {
        let zip = self
            .archive
            .as_ref()
            .map_err(|e| JsValue::from_str(&format!("pptx-parser error: {e}")))?;
        zip.read_font_part(path).map_err(|e| JsValue::from_str(&e))
    }

    /// GitHub-flavoured markdown projection of the retained archive. Mirrors the
    /// free `pptx_to_markdown`. A corrupt container degrades to an empty deck.
    pub fn to_markdown(&mut self) -> Result<String, JsValue> {
        self.render_markdown_inner().map_err(pptx_parser_js_error)
    }

    fn render_markdown_inner(&mut self) -> Result<String, String> {
        if self.prepared_slide.is_some() {
            return Err("a slide unit is awaiting acknowledgement".to_string());
        }
        if let Err(error) = &self.archive {
            return Ok(render_presentation_md(&degraded_container_presentation(
                error.clone(),
            )));
        }

        self.archive
            .as_mut()
            .expect("container open checked above")
            .begin_operation("markdown")?;
        let result = (|| -> Result<String, String> {
            self.ensure_presentation()?;
            let mut shared = self.presentation.take().expect("presentation loaded above");
            let rendered = render_markdown_from_shared(
                &mut shared,
                self.archive.as_mut().expect("container open checked above"),
            );
            self.presentation = Some(shared);
            rendered
        })();
        settle_pptx_operation(
            self.archive.as_mut().expect("container open checked above"),
            result,
        )
    }
}

// ===========================
//  ZIP helpers
// ===========================

/// PPTX-local adapter over one owned, validated package session. All reads in a
/// public call share one explicit operation; focused parser tests lazily receive
/// a compatibility operation while using the same bounded reader path.
pub(crate) struct PptxZip {
    session: PackageSessionHandle,
    operation: RetainedPackageOperation,
}

impl PptxZip {
    #[cfg(test)]
    pub(crate) fn new(source: Cursor<Vec<u8>>) -> Result<Self, String> {
        open_zip(source.into_inner())
    }

    fn begin_operation(&mut self, name: &str) -> Result<(), String> {
        self.operation.begin(&self.session, name)
    }

    fn operation(&mut self) -> Result<&PackageOperation, String> {
        #[cfg(test)]
        let compatibility_name = Some("pptx-parser-compat");
        #[cfg(not(test))]
        let compatibility_name = None;
        self.operation.operation(&self.session, compatibility_name)
    }

    #[cfg(test)]
    fn active_operation(&self) -> Result<&PackageOperation, String> {
        self.operation.active()
    }

    fn finish_operation(&mut self) -> Result<(), String> {
        self.operation.finish()
    }

    fn cancel_operation(&mut self) {
        self.operation.cancel();
    }

    fn run_operation<T>(
        &mut self,
        name: &str,
        run: impl FnOnce(&mut Self) -> Result<T, String>,
    ) -> Result<T, String> {
        self.begin_operation(name)?;
        let result = run(self);
        self.operation.settle(&self.session, result)
    }

    fn assert_healthy(&self) -> Result<(), String> {
        self.session.assert_healthy()
    }

    fn usage(&self) -> ResourceUsage {
        self.session.usage()
    }

    /// Raw package parts are independent of the retained slide cursor. Give
    /// each read its own package operation so rendering slide N may overlap the
    /// acknowledged pull lifecycle for slide N+1 without sharing ownership or
    /// accounting state.
    fn read_part_in_independent_operation(
        &self,
        operation_name: &str,
        path: &str,
    ) -> Result<Vec<u8>, String> {
        self.session
            .run_operation(operation_name, |operation| operation.read_bytes(path))
    }

    fn read_font_part(&self, path: &str) -> Result<Vec<u8>, String> {
        self.read_font_part_with_limit(path, HARD_MAX_EMBEDDED_FONT_BYTES)
    }

    fn read_font_part_with_limit(&self, path: &str, max_bytes: u64) -> Result<Vec<u8>, String> {
        let max_bytes = usize::try_from(max_bytes)
            .map_err(|_| "embedded-font byte limit does not fit this target".to_string())?;
        self.session.run_operation("extract-font", |operation| {
            operation.read_bytes_bounded(path, max_bytes)
        })
    }

    fn index_for_name(&self, path: &str) -> Option<()> {
        self.session.contains_entry(path).then_some(())
    }
}

fn settle_pptx_operation<T>(zip: &mut PptxZip, result: Result<T, String>) -> Result<T, String> {
    zip.operation.settle(&zip.session, result)
}

pub(crate) fn read_zip_str(
    zip: &mut PptxZip,
    path: &str,
) -> Result<String, Box<dyn std::error::Error>> {
    read_pptx_dependency_xml(zip, path)
}

/// Inflate one XML part through a fixed-scratch, limit+1 buffer. At exactly the
/// ceiling the final one-byte read is an EOF/CRC probe, so a corrupt stored part
/// does not bypass validation. The returned allocation never intentionally
/// retains more than limit+1 payload bytes (allocator rounding is not counted).
fn read_bounded_pptx_xml(
    zip: &mut PptxZip,
    path: &str,
    limit_u64: u64,
    kind: HardResourceLimitKind,
) -> Result<String, String> {
    const SCRATCH_BYTES: usize = 8 * 1024;
    let limit = usize::try_from(limit_u64)
        .map_err(|_| "PPTX XML ceiling does not fit this target".to_string())?;
    let retained_cap = limit
        .checked_add(1)
        .ok_or_else(|| "PPTX XML ceiling overflow".to_string())?;
    let operation = zip.operation()?;
    let mut stream = operation.open_entry(path)?;
    let reporter = stream.limit_reporter()?;
    let mut bytes = Vec::with_capacity(retained_cap.min(SCRATCH_BYTES));
    let mut scratch = [0u8; SCRATCH_BYTES];
    while bytes.len() < retained_cap {
        let wanted = (retained_cap - bytes.len()).min(scratch.len());
        let count = stream
            .read(&mut scratch[..wanted])
            .map_err(|error| error.to_string())?;
        if count == 0 {
            break;
        }
        let required = bytes.len() + count;
        if required > bytes.capacity() {
            let doubled = bytes.capacity().saturating_mul(2).max(SCRATCH_BYTES);
            let next_capacity = required.max(doubled).min(retained_cap);
            bytes.reserve_exact(next_capacity - bytes.len());
        }
        // Requested capacity grows geometrically and retained payload is clamped
        // to limit+1 (the allocator may round capacity). At exactly `limit`, the
        // next loop iteration requests one byte, validating decoder EOF/CRC.
        bytes.extend_from_slice(&scratch[..count]);
    }
    if bytes.len() > limit {
        reporter.observe_hard_limit(kind, Some(path), limit_u64, limit_u64 + 1)?;
        unreachable!("crossing a hard resource limit poisons and returns an error");
    }
    let xml = String::from_utf8(bytes)
        .map_err(|error| format!("ZIP entry is not valid UTF-8 ({path}): {error}"))?;
    let complexity_limit = pptx_internal_limits().xml_dom_complexity;
    if xml_dom_complexity_exceeds(&xml, complexity_limit) {
        reporter.observe_hard_limit(
            HardResourceLimitKind::XmlDomComplexity,
            Some(path),
            complexity_limit,
            complexity_limit.saturating_add(1),
        )?;
        unreachable!("crossing a hard resource limit poisons and returns an error");
    }
    Ok(xml)
}

/// Primary slide input is independently bounded because it is the indivisible
/// random-access producer unit.
fn read_primary_slide_xml(zip: &mut PptxZip, path: &str) -> Result<String, String> {
    read_bounded_pptx_xml(
        zip,
        path,
        pptx_slide_xml_limit(),
        HardResourceLimitKind::PptxSlideXmlBytes,
    )
}

/// Presentation, relationship, theme, master, layout and comment-author XML
/// can be retained directly or amplified into shared parsed models. Bound them
/// before UTF-8, DOM and model allocation. The 16 MiB candidate is intentionally
/// conservative and awaits corpus calibration in M7; it is not a heap estimate.
fn read_pptx_dependency_xml(
    zip: &mut PptxZip,
    path: &str,
) -> Result<String, Box<dyn std::error::Error>> {
    read_bounded_pptx_xml(
        zip,
        path,
        pptx_internal_limits().shared_dependency_xml_bytes,
        HardResourceLimitKind::PptxSharedDependencyXmlBytes,
    )
    .map_err(Into::into)
}

pub(crate) fn read_zip_bytes(zip: &mut PptxZip, path: &str) -> Result<Vec<u8>, String> {
    zip.operation()?.read_bytes(path)
}

pub(crate) fn read_zip_head(
    zip: &mut PptxZip,
    path: &str,
    max_bytes: usize,
) -> Result<Vec<u8>, String> {
    zip.operation()?.read_head(path, max_bytes)
}

// ===========================
//  Table style data model
// ===========================

/// Text component of one DrawingML `CT_TablePartStyle` (§21.1.3.11).
#[derive(Debug, Clone, Default)]
struct TableTextStyle {
    color: Option<String>,
    bold: Option<bool>,
    italic: Option<bool>,
}

impl TableTextStyle {
    fn overlay(&mut self, role: &Self) {
        if role.color.is_some() {
            self.color = role.color.clone();
        }
        if role.bold.is_some() {
            self.bold = role.bold;
        }
        if role.italic.is_some() {
            self.italic = role.italic;
        }
    }
}

/// Presence-preserving table-style line. `NoLine` is an authored
/// `<a:ln><a:noFill/></a:ln>` or `<a:lnRef idx="0">`; it must clear a lower
/// precedence role, while `Unspecified` inherits it.
#[derive(Debug, Clone, Default)]
enum TableLineStyle {
    #[default]
    Unspecified,
    NoLine,
    Stroke(Box<Stroke>),
}

impl TableLineStyle {
    fn from_stroke(stroke: Option<Stroke>) -> Self {
        stroke.map(Box::new).map(Self::Stroke).unwrap_or_default()
    }

    fn overlay(&mut self, role: &Self) {
        if !matches!(role, Self::Unspecified) {
            *self = role.clone();
        }
    }

    fn apply_to(&self, target: &mut Option<Stroke>) {
        match self {
            Self::Unspecified => {}
            Self::NoLine => *target = None,
            Self::Stroke(stroke) => *target = Some((**stroke).clone()),
        }
    }
}

/// ECMA-376 `CT_TableCellBorderStyle` (§20.1.4.2.4): all six orthogonal and
/// both diagonal members are retained on every one of the thirteen roles.
#[derive(Debug, Clone, Default)]
struct TableCellBorderStyle {
    left: TableLineStyle,
    right: TableLineStyle,
    top: TableLineStyle,
    bottom: TableLineStyle,
    inside_h: TableLineStyle,
    inside_v: TableLineStyle,
    diagonal_tl: TableLineStyle,
    diagonal_tr: TableLineStyle,
}

/// One of the thirteen `CT_TablePartStyle` roles in `CT_TableStyle`. Keeping
/// fill, text, and borders together prevents vertical bands and corner roles
/// from silently supporting only text while row roles support paint.
#[derive(Debug, Clone, Default)]
struct TablePartStyle {
    fill: Option<Fill>,
    text: TableTextStyle,
    borders: TableCellBorderStyle,
}

#[derive(Debug, Clone, Default)]
struct TableStyleDef {
    whole_tbl: TablePartStyle,
    band1_h: TablePartStyle,
    band2_h: TablePartStyle,
    band1_v: TablePartStyle,
    band2_v: TablePartStyle,
    first_row: TablePartStyle,
    last_row: TablePartStyle,
    first_col: TablePartStyle,
    last_col: TablePartStyle,
    nw_cell: TablePartStyle,
    ne_cell: TablePartStyle,
    sw_cell: TablePartStyle,
    se_cell: TablePartStyle,
}

#[derive(Debug, Clone, Copy, Default)]
struct TableStyleFlags {
    first_row: bool,
    last_row: bool,
    first_col: bool,
    last_col: bool,
    band_row: bool,
    band_col: bool,
}

#[derive(Debug, Clone, Default)]
struct ResolvedTableCellStyle {
    fill: Option<Fill>,
    text: TableTextStyle,
    border_l: TableLineStyle,
    border_r: TableLineStyle,
    border_t: TableLineStyle,
    border_b: TableLineStyle,
    diagonal_tl: TableLineStyle,
    diagonal_tr: TableLineStyle,
}

// ===========================
//  XML helpers (roxmltree)
// ===========================

pub(crate) fn child<'a, 'i>(
    node: roxmltree::Node<'a, 'i>,
    local: &str,
) -> Option<roxmltree::Node<'a, 'i>> {
    node.children()
        .find(|n| n.is_element() && n.tag_name().name() == local)
}

pub(crate) fn children_vec<'a, 'i>(
    node: roxmltree::Node<'a, 'i>,
    local: &str,
) -> Vec<roxmltree::Node<'a, 'i>> {
    node.children()
        .filter(|n| n.is_element() && n.tag_name().name() == local)
        .collect()
}

pub(crate) fn attr(node: &roxmltree::Node<'_, '_>, local: &str) -> Option<String> {
    node.attributes()
        .find(|a| a.name() == local && a.namespace().is_none())
        .map(|a| a.value().to_owned())
}

/// Attribute in the r: (relationships) namespace — e.g. r:id, r:embed. Accepts
/// both the Transitional and Strict (ISO/IEC 29500) relationships URIs.
pub(crate) fn attr_r(node: &roxmltree::Node<'_, '_>, local: &str) -> Option<String> {
    node.attributes()
        .find(|a| a.name() == local && is_r_ns(a.namespace()))
        .map(|a| a.value().to_owned())
}

pub(crate) fn attr_i64(node: &roxmltree::Node<'_, '_>, local: &str) -> Option<i64> {
    attr(node, local)?.parse().ok()
}

pub(crate) fn attr_f64(node: &roxmltree::Node<'_, '_>, local: &str) -> Option<f64> {
    attr(node, local)?.parse().ok()
}

// ===========================
//  Relationships helpers
// ===========================

/// id → raw target used by the PPTX parser's existing relationship consumers.
/// The XML has already passed the typed lexical preflight at package read time;
/// parse it through the PPTX-local node ceiling rather than the unbounded shared
/// compatibility helper.
pub(crate) fn parse_rels(xml: &str) -> HashMap<String, String> {
    let Ok(doc) = parse_preflighted_pptx_xml(xml) else {
        return HashMap::new();
    };
    doc.root_element()
        .children()
        .filter(|node| node.is_element() && node.tag_name().name() == "Relationship")
        .filter_map(|node| Some((attr(&node, "Id")?, attr(&node, "Target")?)))
        .collect()
}

/// Pair each SmartArt diagramData part with its prebaked diagramDrawing part and
/// load the drawing XML from the zip. Returns `dm_rid → drawing_xml_content`,
/// keyed by the diagramData relationship Id (i.e. the value of the slide's
/// `<dgm:relIds r:dm>` — §21.4.2.22), which is what the shape walker looks up.
///
/// **Canonical path (ECMA-376 §21.4.2.22 + MS-ODRAWXML `dsp:dataModelExt`).**
/// The link from a data model to its cached drawing is explicit, not positional:
///
/// 1. `<dgm:relIds r:dm="rId2">` on the slide points at the data part (e.g.
///    `../diagrams/data1.xml`) via the containing part's rels.
/// 2. The data part carries `<dsp:dataModelExt relId="rId6" .../>` in its
///    `<a:extLst>`; `relId` names the diagramDrawing relationship.
/// 3. That relationship (`rId6 → ../diagrams/drawing1.xml`) is resolved in the
///    same rels file — Office authors the 2007 `diagramDrawing` relationship on
///    the referencing part, not the data part (real PowerPoint output has no
///    `ppt/diagrams/_rels/dataN.xml.rels`).
///
/// So a data part's `dataModelExt relId` is the authority for which drawing part
/// belongs to it, even if the file-number suffixes disagree.
///
/// **Fallback.** For a malformed/older file whose data part lacks a
/// `dataModelExt` (or whose `relId` doesn't resolve), fall back to matching the
/// file-number suffix (`data1.xml ↔ drawing1.xml`). This is a heuristic kept
/// only for compatibility; the spec-driven relId path above is primary.
pub(crate) fn build_smartart_drawings(
    rels_xml: &str,
    source_dir: &str,
    zip: &mut PptxZip,
) -> HashMap<String, String> {
    let mut result: HashMap<String, String> = HashMap::new();
    let doc = match parse_preflighted_pptx_xml(rels_xml) {
        Ok(d) => d,
        Err(_) => return result,
    };
    // Index every relationship as rId → (type-suffix-relevant target). We need
    // both the diagramData rels (to key the result and load the data part) and a
    // rId → target lookup for the drawing relId the data part names.
    let mut rid_target: HashMap<String, String> = HashMap::new();
    let mut data_rels: Vec<(String, String)> = Vec::new();
    let mut drawing_targets: Vec<String> = Vec::new();
    for rel in doc.root_element().children().filter(|n| n.is_element()) {
        let rel_type = attr(&rel, "Type").unwrap_or_default();
        let (Some(rid), Some(target)) = (attr(&rel, "Id"), attr(&rel, "Target")) else {
            continue;
        };
        rid_target.insert(rid.clone(), target.clone());
        if rel_type.ends_with("/diagramData") {
            data_rels.push((rid, target));
        } else if rel_type.ends_with("/diagramDrawing") {
            drawing_targets.push(target);
        }
    }

    for (dm_rid, data_target) in data_rels {
        // 1) Canonical: read the data part's dataModelExt relId, resolve it in
        //    this same rels map.
        let drawing_target = smartart_drawing_relid(&data_target, source_dir, zip)
            .and_then(|drawing_rid| rid_target.get(&drawing_rid).cloned())
            // 2) Fallback: file-number-suffix match (heuristic, compat only).
            .or_else(|| {
                trailing_num(&data_target).and_then(|num| {
                    drawing_targets
                        .iter()
                        .find(|t| trailing_num(t) == Some(num))
                        .cloned()
                })
            });
        if let Some(dt) = drawing_target {
            let drawing_path = resolve_path(source_dir, &dt);
            if let Ok(xml) = read_zip_str(zip, &drawing_path) {
                result.insert(dm_rid, xml);
            }
        }
    }
    result
}

/// Read a SmartArt data part and return the `relId` its `<dsp:dataModelExt>`
/// names for the cached drawing part (MS-ODRAWXML; the `dsp` namespace is
/// `.../office/drawing/2008/diagram`). Returns `None` when the data part can't
/// be read or carries no `dataModelExt@relId`.
fn smartart_drawing_relid(
    data_target: &str,
    source_dir: &str,
    zip: &mut PptxZip,
) -> Option<String> {
    let data_path = resolve_path(source_dir, data_target);
    let xml = read_zip_str(zip, &data_path).ok()?;
    let doc = parse_preflighted_pptx_xml(&xml).ok()?;
    doc.descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "dataModelExt")
        .and_then(|n| n.attribute("relId"))
        .map(str::to_owned)
}

/// Trailing decimal suffix of a part's file stem (`.../drawing12.xml` → 12).
/// Used only by the compatibility fallback in [`build_smartart_drawings`].
fn trailing_num(target: &str) -> Option<u32> {
    let file = target.rsplit('/').next().unwrap_or("");
    let stem = file.split('.').next().unwrap_or("");
    let digits: String = stem
        .chars()
        .rev()
        .take_while(|c| c.is_ascii_digit())
        .collect();
    digits.chars().rev().collect::<String>().parse().ok()
}

/// Find the Target of the first relationship whose Type ends with `type_suffix`.
// Relationship Type is matched by suffix (`ends_with`), so the Strict purl.oclc.org
// prefix still matches — do not change this to an exact-match comparison, or
// Strict documents will silently stop resolving.
pub(crate) fn find_rel_target_by_type(rels_xml: &str, type_suffix: &str) -> Option<String> {
    let doc = parse_preflighted_pptx_xml(rels_xml).ok()?;
    for rel in doc.root_element().children().filter(|n| n.is_element()) {
        if let Some(rel_type) = attr(&rel, "Type") {
            if rel_type.ends_with(type_suffix) {
                return attr(&rel, "Target");
            }
        }
    }
    None
}

const CLASSIC_COMMENT_AUTHOR_RELATIONSHIP_TYPES: &[&str] = &[
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/commentAuthors",
];
const MODERN_COMMENT_AUTHOR_RELATIONSHIP_TYPES: &[&str] =
    &["http://schemas.microsoft.com/office/2018/10/relationships/authors"];

/// Resolve an internal package part by one of the exact relationship Type URIs.
/// ECMA-376 Part 2 §9.3: omitted TargetMode is Internal; External targets do not
/// identify parts in the package.
fn find_internal_rel_target_by_types(
    rels_xml: &str,
    relationship_types: &[&str],
) -> Option<String> {
    let doc = parse_preflighted_pptx_xml(rels_xml).ok()?;
    doc.root_element()
        .children()
        .filter(|node| node.is_element() && node.tag_name().name() == "Relationship")
        .filter(|node| matches!(attr(node, "TargetMode").as_deref(), None | Some("Internal")))
        .find_map(|node| {
            attr(&node, "Type")
                .filter(|rel_type| relationship_types.contains(&rel_type.as_str()))
                .and_then(|_| attr(&node, "Target"))
        })
}

const FONT_RELATIONSHIP_TYPES: &[&str] = &[
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/font",
];
const PPTX_FONT_CONTENT_TYPES: &[&str] = &["application/x-font-ttf", "application/x-fontdata"];

/// Project `p:embeddedFontLst` into compact, extraction-safe references.
///
/// ECMA-376 Part 1 §19.2.1.9-.10 associates a `p:font@typeface` (whose schema
/// type is DrawingML `a:CT_TextFont`) with up to
/// four style slots whose `r:id` resolves through presentation relationships.
/// Part 1 §15.2.13 permits raw sfnt (`application/x-font-ttf`) and EOT
/// (`application/x-fontdata`) Font parts. External, wrong-type, missing and
/// unsupported relationships are ignored so malformed optional font metadata
/// never prevents the presentation itself from opening.
fn parse_embedded_font_refs(
    pres_root: roxmltree::Node<'_, '_>,
    rels_xml: &str,
    zip: &mut PptxZip,
) -> Vec<PptxEmbeddedFontRef> {
    let Some(list) = pres_root.children().find(|node| {
        node.is_element()
            && node.tag_name().name() == "embeddedFontLst"
            && is_p_ns(node.tag_name().namespace())
    }) else {
        return Vec::new();
    };
    let relationships = parse_opc_rels(rels_xml);
    let content_types_node_limit =
        u32::try_from(pptx_internal_limits().xml_dom_complexity).unwrap_or(u32::MAX);
    let content_types = read_zip_str(zip, "[Content_Types].xml")
        .ok()
        .and_then(|xml| PackageContentTypes::parse_with_node_limit(&xml, content_types_node_limit))
        .unwrap_or_default();

    let mut refs = Vec::new();
    for entry in list.children().filter(|node| {
        node.is_element()
            && node.tag_name().name() == "embeddedFont"
            && is_p_ns(node.tag_name().namespace())
    }) {
        let Some(font_name) = entry
            .children()
            .find(|node| {
                node.is_element()
                    && node.tag_name().name() == "font"
                    && is_p_ns(node.tag_name().namespace())
            })
            .and_then(|font| attr(&font, "typeface"))
            .filter(|name| !name.trim().is_empty())
        else {
            continue;
        };

        for (element_name, style) in [
            ("regular", EmbeddedFontStyle::Regular),
            ("bold", EmbeddedFontStyle::Bold),
            ("italic", EmbeddedFontStyle::Italic),
            ("boldItalic", EmbeddedFontStyle::BoldItalic),
        ] {
            let Some(relationship_id) = entry
                .children()
                .find(|node| {
                    node.is_element()
                        && node.tag_name().name() == element_name
                        && is_p_ns(node.tag_name().namespace())
                })
                .and_then(|slot| attr_r(&slot, "id"))
            else {
                continue;
            };
            let Some(relationship) = relationships.get(&relationship_id) else {
                continue;
            };
            if relationship.mode == TargetMode::External
                || !relationship
                    .relationship_type
                    .as_deref()
                    .is_some_and(|rel_type| FONT_RELATIONSHIP_TYPES.contains(&rel_type))
            {
                continue;
            }
            let part_path = resolve_path("ppt", &relationship.target);
            let Some(content_type) = content_types.for_part(&part_path) else {
                continue;
            };
            if !PPTX_FONT_CONTENT_TYPES.contains(&content_type)
                || zip.index_for_name(&part_path).is_none()
            {
                continue;
            }
            refs.push(PptxEmbeddedFontRef {
                font_name: font_name.clone(),
                style,
                part_path,
                content_type: content_type.to_owned(),
            });
        }
    }
    refs
}

/// Resolve a relative path against a base directory inside the ZIP.
///
/// Thin alias for the shared [`ooxml_common::rels::resolve_target`], which
/// handles both root-absolute (`/ppt/charts/chart5.xml`) and relative
/// (`../charts/chart1.xml`) Targets with `..` normalization (ECMA-376 Part 2
/// §9.3). Kept as a local name so the many call sites read unchanged.
pub(crate) fn resolve_path(base_dir: &str, target: &str) -> String {
    ooxml_common::rels::resolve_target(base_dir, target)
}

/// Directory containing an OPC source part. Relationship Targets are resolved
/// relative to this directory (ECMA-376 Part 2 §6.5.2.3).
fn part_directory(part_path: &str) -> &str {
    part_path.rsplit_once('/').map_or("", |(dir, _)| dir)
}

// ===========================
//  Slide parser
// ===========================

/// `<p:sld show="0">` / `show="false"` marks a slide hidden in the slide show
/// (ECMA-376 §19.3.1.38 `sld` / `CT_Slide` — `show`, xsd:boolean, default true).
/// Absent or any truthy value ⇒ shown. NB: this matches the FALSY literals —
/// the inverse of the codebase's usual `== "1" || == "true"` truthy check —
/// because `show` defaults to true, so a slide is hidden only on explicit false.
fn slide_is_hidden(root: roxmltree::Node) -> bool {
    matches!(root.attribute("show"), Some("0") | Some("false"))
}

// Threads the full master+layout inheritance context (per-type font sizes,
// bullets, anchors, transforms, alignments, spacing, bold/italic/caps/color
// maps) plus zip/theme into one slide parse; this is the inheritance chain
// ECMA-376 requires, not an arbitrary parameter bag.
#[allow(clippy::too_many_arguments)]
fn parse_slide(
    xml: &str,
    slide_dir: &str,
    slide_rels_xml: &str,
    // The layout's single-pass extraction (placeholders + layout bg + layout
    // showMasterSp), built/cached by the caller against this slide's effective
    // theme (D4). `layout_xml` is still passed for the per-slide DECORATIVE walk
    // only (its shapes bind to the slide's own smartart + theme + zip, so they
    // can't live in the cached `ParsedLayout`).
    parsed_layout: &ParsedLayout,
    layout_xml: Option<&str>,
    layout_rels: &HashMap<String, String>,
    layout_dir: &str,
    bundle: &ParsedMaster,
    eff: Option<&EffectiveMaster>,
    index: usize,
    rels: &HashMap<String, String>,
    smartart_drawings: &HashMap<String, String>,
    comment_authors: &mut Option<HashMap<String, String>>,
    comment_authors_path: Option<&str>,
    modern_comment_authors: &mut Option<HashMap<String, String>>,
    modern_comment_authors_path: Option<&str>,
    zip: &mut PptxZip,
) -> Result<Slide, Box<dyn std::error::Error>> {
    // Destructure the per-slide master bundle into the local names the rest of
    // this function uses. `theme` here is the slide's effective theme (the
    // master's own theme with its <p:clrMap> baked in), so scheme colors
    // resolve against the right palette per slide.
    // Only the fields this function still consumes directly are bound; the
    // master INHERITANCE maps (font sizes, level sizes/indents/bullets, anchors,
    // transforms, alignments, ea-ln-brk, spacing) now feed `parse_layout` in the
    // caller, which produced the `ParsedLayout` passed in. `theme` here is the
    // slide's effective theme (master clrMap baked in) so scheme colors resolve
    // against the right palette per slide.
    let ParsedMaster {
        theme,
        master_xml,
        master_rels,
        master_dir,
        master_smartart_drawings,
        master_bg,
        master_decorative,
        master_bold,
        master_italic,
        master_caps,
        master_reflection,
        master_color,
        ..
    } = bundle;
    // When the slide/layout carries a `<p:clrMapOvr><a:overrideClrMapping>`
    // (ECMA-376 §19.3.1.7), the caller recomputed the master's theme-dependent
    // fields against the slide's effective mapping (`EffectiveMaster`); use them
    // in place of the master's frozen values so that BOTH the slide's own scheme
    // colors AND master-inherited ones (the master `<p:bg>`, master txStyles
    // placeholder colors) resolve against the override mapping (§20.1.6.8).
    // Otherwise fall back to the master bundle's values. (Master bullet colors
    // flow through `parsed_layout`, already override-adjusted by the caller.)
    let theme: &PptxTheme = eff.map(|e| &e.theme).unwrap_or(theme);
    let master_xml: Option<&str> = master_xml.as_deref();
    let master_dir: &str = master_dir.as_str();
    let master_bg: Option<Fill> = match eff {
        Some(e) => e.master_bg.clone(),
        None => master_bg.clone(),
    };
    let master_color: &HashMap<String, String> =
        eff.map(|e| &e.master_color).unwrap_or(master_color);

    // The layout placeholder inheritance was resolved once in `parse_layout`
    // (cached across slides sharing this layout, or rebuilt for a clrMapOvr
    // slide) against this slide's effective theme. Clone it so the per-slide
    // master txStyles fallbacks below can be layered on without mutating the
    // shared/cached instance.
    let mut lph = parsed_layout.placeholders.clone();
    // Fall back to master txStyles defRPr @b/@i when the layout did not specify
    // bold/italic for a placeholder type. Without this, e.g. the master titleStyle's
    // b="1" is not applied to ctrTitle / title placeholders.
    for (t, b) in master_bold.iter() {
        lph.by_type_bold.entry(t.clone()).or_insert(*b);
    }
    for (t, i) in master_italic.iter() {
        lph.by_type_italic.entry(t.clone()).or_insert(*i);
    }
    for (t, c) in master_caps.iter() {
        lph.by_type_caps.entry(t.clone()).or_insert(c.clone());
    }
    for (t, reflection) in master_reflection.iter() {
        lph.by_type_reflection
            .entry(t.clone())
            .or_insert_with(|| reflection.clone());
    }
    for (t, c) in master_color.iter() {
        lph.by_type_master_color
            .entry(t.clone())
            .or_insert(c.clone());
    }

    note_layout_master_parse();
    // Guard against a pathologically deep slide XML: roxmltree's tree builder
    // recurses per element-nesting level, so a slide nested thousands deep
    // overflows the fixed WASM stack and traps *inside* `Document::parse` before
    // our own depth-guarded shape walk runs. The nesting-depth pre-check that
    // rejects it now lives in `parse_preflighted_pptx_xml`.
    let doc = parse_preflighted_pptx_xml(xml)?;
    let root = doc.root_element(); // <p:sld>
    let hidden = slide_is_hidden(root);
    let c_sld = child(root, "cSld");

    // Background chain: slide → layout → master. Each level resolves a blip
    // background (§20.1.8.14) against its own rels + part directory, so the
    // closures are run sequentially (one mutable borrow of `zip` at a time).
    let mut background: Option<Fill> = None;

    // Slide-level bg (rels and relative Targets are scoped to the actual slide
    // source part; ECMA-376 Part 2 §6.5.2.3).
    if let Some(n) = c_sld {
        let mut resolve = |rid: &str| -> Option<String> {
            let target = rels.get(rid)?;
            let path = resolve_path(slide_dir, target);
            // Resolve to the zip path; verify the part exists so a dangling
            // rId still yields None (the bg chain then falls through to the
            // next level), preserving the prior data-URL behaviour.
            // `index_for_name` reads the central directory only (no inflate),
            // unlike the former `read_zip_bytes` which decompressed to discard.
            zip.index_for_name(&path)?;
            Some(path)
        };
        background = parse_background(n, theme, &mut resolve);
    }

    // Layout-level bg: resolved once in `parse_layout` (against this slide's
    // effective theme) and applied only when the slide's own bg chain is empty.
    if background.is_none() {
        background = parsed_layout.background.clone();
    }

    // Master-level bg (resolved by the caller before parse_slide; already a Fill).
    let background = background.or(master_bg);

    let sp_tree = c_sld
        .and_then(|n| child(n, "spTree"))
        .ok_or("missing spTree")?;

    let mut elements = Vec::new();
    let mut element_sources = Vec::new();

    // ── showMasterSp resolution (ECMA-376 §19.3.1.38 sld / §19.3.1.39
    // sldLayout, AG_ChildSlide, default true) ─────────────────────────────
    // Master decorative shapes are composited beneath the slide only when both
    // the slide and its layout permit it. Either one setting showMasterSp="0"
    // suppresses the master's spTree decorations (the slide flag is honored for
    // the slide itself; the layout flag — read once in `parse_layout` — for
    // shapes inherited through it).
    let slide_show_master_sp = read_show_master_sp(root);
    let show_master_sp = slide_show_master_sp && parsed_layout.show_master_sp;

    // ── Master non-placeholder shapes (rendered BELOW layout & slide) ─────
    // The slide master's spTree may carry decorative pictures/shapes (logos,
    // bands) that are not placeholder anchors. PowerPoint composites them at
    // the very bottom, beneath the layout's decorations and the slide content.
    // Gated by showMasterSp (above). Placeholders are skipped — only the
    // master's decorative content is drawn here.
    //
    // These were pre-extracted once per cached master in `build_master_bundle`
    // (resolved against the master's baked theme), so the common no-override
    // slide clones them instead of re-parsing the master XML + re-walking its
    // spTree. A slide with a `<p:clrMapOvr>` (`eff.is_some()`) must re-resolve
    // them against its override theme, so it re-extracts from the master XML —
    // exactly what the old unconditional inline walk did, now only on the rare
    // override path. `elements` is still empty here, so ordering (master
    // decorations first) is unchanged either way.
    if show_master_sp {
        let start = elements.len();
        if eff.is_some() {
            if let Some(mxml) = master_xml {
                note_layout_master_parse();
                if let Ok(mdoc) = parse_preflighted_pptx_xml(mxml) {
                    extract_decorative_shapes(
                        mdoc.root_element(),
                        master_dir,
                        master_rels,
                        master_smartart_drawings,
                        theme,
                        zip,
                        &mut elements,
                    );
                }
            }
        } else {
            elements.extend(master_decorative.iter().cloned());
        }
        element_sources.extend((start..elements.len()).map(|_| SlideElementSource {
            origin: SlideElementOrigin::Master,
        }));
    }

    // ── Layout non-placeholder shapes (rendered BEFORE slide shapes) ──────
    // These are decorative background elements defined in the slide layout
    // (e.g. coloured bands, logos) that are not placeholder anchors. ECMA-376
    // §19.3.1.38 scopes showMasterSp to shapes on the master slide, so the
    // selected layout's own shapes remain visible.
    {
        if let Some(lxml) = layout_xml {
            note_layout_master_parse();
            if let Ok(ldoc) = parse_preflighted_pptx_xml(lxml) {
                let lroot = ldoc.root_element();
                if let Some(lsp_tree) = child(lroot, "cSld").and_then(|n| child(n, "spTree")) {
                    let empty_lph = LayoutPlaceholders::default();
                    for node in lsp_tree.children().filter(|n| n.is_element()) {
                        let start = elements.len();
                        parse_sp_tree_node(
                            node,
                            &empty_lph,
                            layout_dir,
                            layout_rels,
                            smartart_drawings,
                            zip,
                            theme,
                            &mut elements,
                            true, // skip placeholder shapes
                            None, // no inherited group fill at top level
                            ooxml_common::depth::DepthGuard::root(),
                        );
                        element_sources.extend((start..elements.len()).map(|_| {
                            SlideElementSource {
                                origin: SlideElementOrigin::Layout,
                            }
                        }));
                    }
                }
            }
        }
    }

    // ── Slide shapes ─────────────────────────────────────────────────────
    for node in sp_tree.children().filter(|n| n.is_element()) {
        let start = elements.len();
        parse_sp_tree_node(
            node,
            &lph,
            slide_dir,
            rels,
            smartart_drawings,
            zip,
            theme,
            &mut elements,
            false,
            None,
            ooxml_common::depth::DepthGuard::root(),
        );
        element_sources.extend((start..elements.len()).map(|_| SlideElementSource {
            origin: SlideElementOrigin::Slide,
        }));
    }

    debug_assert_eq!(elements.len(), element_sources.len());

    // ── Notes slide & comments (Phase 2 surfacing only — no rendering) ────
    let notes = load_notes_slide(zip, slide_dir, rels);
    let comments = load_pptx_comments(
        zip,
        slide_dir,
        slide_rels_xml,
        comment_authors,
        comment_authors_path,
        modern_comment_authors,
        modern_comment_authors_path,
    );

    Ok(Slide {
        index,
        slide_number: index + 1,
        // Stamped by the build loop, which owns the resolved slide part path.
        part_name: None,
        background,
        elements,
        element_sources,
        notes,
        comments,
        hidden,
        parse_error: None,
    })
}

/// RB7: a placeholder for a slide whose part failed to parse. The deck keeps its
/// other slides; this one renders as a visible error box. `part` is the ZIP path
/// (e.g. `ppt/slides/slide3.xml`) so the message pinpoints which slide broke.
fn broken_slide(index: usize, part: &str, detail: &str) -> Slide {
    Slide {
        index,
        slide_number: index + 1,
        // `part` IS the slide part path here (broken_slide is called with it), so
        // the slide→index map still resolves an internal jump to a broken slide.
        part_name: Some(part.to_string()),
        background: None,
        elements: Vec::new(),
        element_sources: Vec::new(),
        notes: None,
        comments: Vec::new(),
        hidden: false,
        parse_error: Some(format!("{part}: {detail}")),
    }
}

/// Resolve the slide's `notesSlide` relationship, read the notes part, and
/// return its plain text (paragraphs joined by '\n'). Returns `None` when
/// the slide has no notes part or the part can't be read.
fn load_notes_slide(
    zip: &mut PptxZip,
    slide_dir: &str,
    rels: &HashMap<String, String>,
) -> Option<String> {
    // rels here is the slide's _rels map (rId → Target) parsed by the caller.
    // The relationship Type ends with "/notesSlide". The cleanest way to find
    // the right entry is to look at every value in the map and pick the one
    // pointing at "../notesSlides/...".
    let target = rels.values().find(|t| t.contains("notesSlides/"))?;
    let path = if target.starts_with('/') {
        target.trim_start_matches('/').to_string()
    } else {
        resolve_path(slide_dir, target)
    };
    let xml = read_zip_str(zip, &path).ok()?;
    let doc = parse_preflighted_pptx_xml(&xml).ok()?;
    let mut buf = String::new();
    let mut prev_was_text = false;
    for n in doc.descendants() {
        if !n.is_element() {
            continue;
        }
        let name = n.tag_name().name();
        if name == "p" && prev_was_text {
            buf.push('\n');
            prev_was_text = false;
        }
        if name == "t" {
            if let Some(s) = n.text() {
                buf.push_str(s);
                prev_was_text = true;
            }
        }
    }
    let trimmed = buf.trim();
    if trimmed.is_empty() {
        None
    } else {
        Some(trimmed.to_string())
    }
}

const MODERN_POWERPOINT_NS: &str = "http://schemas.microsoft.com/office/powerpoint/2018/8/main";

fn comment_text_body(node: roxmltree::Node<'_, '_>) -> String {
    let Some(body) = node
        .children()
        .find(|child| child.is_element() && child.tag_name().name() == "txBody")
    else {
        return String::new();
    };
    let mut paragraphs = Vec::new();
    for paragraph in body
        .children()
        .filter(|child| child.is_element() && child.tag_name().name() == "p")
    {
        let text = paragraph
            .descendants()
            .filter(|child| child.is_element() && child.tag_name().name() == "t")
            .filter_map(|child| child.text())
            .collect::<String>();
        paragraphs.push(text);
    }
    paragraphs.join("\n")
}

fn modern_comment_status(node: roxmltree::Node<'_, '_>) -> Option<String> {
    match node.attribute("status").unwrap_or("active") {
        status @ ("active" | "resolved" | "closed") => Some(status.to_string()),
        _ => None,
    }
}

fn modern_comment_anchor_element(
    list: roxmltree::Node<'_, '_>,
) -> (Option<String>, Option<String>) {
    // [MS-ODRAWXML] §2.29.3.20: the last drawing-element moniker in the
    // list identifies the target; earlier monikers describe its container
    // path (document/slide/group).
    list.descendants()
        .rfind(|node| {
            node.is_element()
                && matches!(
                    node.tag_name().name(),
                    "spMk" | "grpSpMk" | "graphicFrameMk" | "cxnSpMk" | "picMk" | "inkMk"
                )
        })
        .map(|node| {
            (
                node.attribute("id").map(String::from),
                node.attribute("creationId").map(String::from),
            )
        })
        .unwrap_or((None, None))
}

fn parse_modern_comment_anchors(comment: roxmltree::Node<'_, '_>) -> Vec<PptxCommentAnchor> {
    let mut anchors = Vec::new();
    for list in comment.children().filter(|node| node.is_element()) {
        match list.tag_name().name() {
            "sldMkLst" => anchors.push(PptxCommentAnchor::Slide),
            "deMkLst" => {
                let (element_id, creation_id) = modern_comment_anchor_element(list);
                anchors.push(PptxCommentAnchor::DrawingElement {
                    element_id,
                    creation_id,
                });
            }
            "txMkLst" => {
                let (element_id, _) = modern_comment_anchor_element(list);
                let range = list
                    .descendants()
                    .find(|node| node.is_element() && node.tag_name().name() == "txMk");
                anchors.push(PptxCommentAnchor::TextRange {
                    element_id,
                    start: range
                        .and_then(|node| node.attribute("cp"))
                        .and_then(|value| value.parse::<i32>().ok()),
                    length: range
                        .and_then(|node| node.attribute("len"))
                        .and_then(|value| value.parse::<i32>().ok()),
                });
            }
            "unknownAnchor" => anchors.push(PptxCommentAnchor::Unknown),
            _ => {}
        }
    }
    anchors
}

/// Resolve and parse the slide's classic or modern comments relationship.
/// Classic comments are ECMA-376 §19.4 `<p:cmLst>` parts. Modern comments use
/// the 2018 PowerPoint namespace and relationship defined by [MS-PPTX] §2.1.5.
fn load_pptx_comments(
    zip: &mut PptxZip,
    slide_dir: &str,
    rels_xml: &str,
    legacy_authors: &mut Option<HashMap<String, String>>,
    legacy_authors_path: Option<&str>,
    modern_authors: &mut Option<HashMap<String, String>>,
    modern_authors_path: Option<&str>,
) -> Vec<PptxComment> {
    let (classic_target, modern_target) = comment_relationship_targets(rels_xml);
    let mut comments = Vec::new();
    // Fixed order makes classic/modern coexistence independent of relationship
    // XML order and HashMap iteration. Each relationship identifies its part;
    // no target-directory or filename convention is inferred.
    for target in [classic_target, modern_target].into_iter().flatten() {
        let path = resolve_path(slide_dir, &target);
        let Ok(xml) = read_zip_str(zip, &path) else {
            continue;
        };
        comments.extend(parse_pptx_comments_part(
            zip,
            &xml,
            legacy_authors,
            legacy_authors_path,
            modern_authors,
            modern_authors_path,
        ));
    }
    comments
}

fn comment_relationship_targets(rels_xml: &str) -> (Option<String>, Option<String>) {
    const CLASSIC: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
    const STRICT_CLASSIC: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/comments";
    const MODERN: &str = "http://schemas.microsoft.com/office/2018/10/relationships/comments";
    let Ok(doc) = parse_preflighted_pptx_xml(rels_xml) else {
        return (None, None);
    };
    let mut classic = None;
    let mut modern = None;
    for rel in doc
        .root_element()
        .children()
        .filter(|node| node.is_element() && node.tag_name().name() == "Relationship")
    {
        if !matches!(attr(&rel, "TargetMode").as_deref(), None | Some("Internal")) {
            continue;
        }
        let Some(rel_type) = attr(&rel, "Type") else {
            continue;
        };
        let Some(target) = attr(&rel, "Target") else {
            continue;
        };
        match rel_type.as_str() {
            CLASSIC | STRICT_CLASSIC if classic.is_none() => classic = Some(target),
            MODERN if modern.is_none() => modern = Some(target),
            _ => {}
        }
    }
    (classic, modern)
}

fn parse_pptx_comments_part(
    zip: &mut PptxZip,
    xml: &str,
    legacy_authors: &mut Option<HashMap<String, String>>,
    legacy_authors_path: Option<&str>,
    modern_authors: &mut Option<HashMap<String, String>>,
    modern_authors_path: Option<&str>,
) -> Vec<PptxComment> {
    let Ok(doc) = parse_preflighted_pptx_xml(xml) else {
        return Vec::new();
    };

    let is_modern = doc.root_element().tag_name().namespace() == Some(MODERN_POWERPOINT_NS);

    if is_modern {
        if modern_authors.is_none() {
            note_comment_authors_load();
            let author_xml = modern_authors_path.and_then(|path| read_zip_str(zip, path).ok());
            *modern_authors = Some(parse_comment_authors(author_xml.as_deref()));
        }
        let empty_authors = HashMap::new();
        let authors = modern_authors.as_ref().unwrap_or(&empty_authors);
        return doc
            .root_element()
            .children()
            .filter(|node| node.is_element() && node.tag_name().name() == "cm")
            .map(|cm| {
                let author_id = cm.attribute("authorId").map(String::from);
                let author = author_id.as_ref().and_then(|id| authors.get(id)).cloned();
                let position = cm
                    .children()
                    .find(|node| node.is_element() && node.tag_name().name() == "pos");
                let replies = cm
                    .children()
                    .find(|node| node.is_element() && node.tag_name().name() == "replyLst")
                    .into_iter()
                    .flat_map(|list| list.children())
                    .filter(|node| node.is_element() && node.tag_name().name() == "reply")
                    .map(|reply| {
                        let reply_author_id = reply.attribute("authorId").map(String::from);
                        PptxCommentReply {
                            id: reply.attribute("id").map(String::from),
                            author: reply_author_id
                                .as_ref()
                                .and_then(|id| authors.get(id))
                                .cloned(),
                            author_id: reply_author_id,
                            date: reply.attribute("created").map(String::from),
                            status: modern_comment_status(reply),
                            text: comment_text_body(reply),
                        }
                    })
                    .collect();
                PptxComment {
                    author_id: None,
                    modern_author_id: author_id,
                    id: cm.attribute("id").map(String::from),
                    index: None,
                    author,
                    date: cm.attribute("created").map(String::from),
                    x: position
                        .and_then(|node| node.attribute("x"))
                        .and_then(|value| value.parse::<i64>().ok()),
                    y: position
                        .and_then(|node| node.attribute("y"))
                        .and_then(|value| value.parse::<i64>().ok()),
                    anchors: parse_modern_comment_anchors(cm),
                    status: modern_comment_status(cm),
                    text: comment_text_body(cm),
                    replies,
                }
            })
            .collect();
    }

    // Preserve the legacy observation point: commentAuthors.xml is irrelevant
    // until a referenced comments part has itself been read and parsed. Cache
    // the owned id → name map after that point so later commented slides reuse
    // it without retaining a roxmltree Document.
    if legacy_authors.is_none() {
        note_comment_authors_load();
        let author_xml = legacy_authors_path.and_then(|path| read_zip_str(zip, path).ok());
        *legacy_authors = Some(parse_comment_authors(author_xml.as_deref()));
    }
    let empty_authors = HashMap::new();
    let authors = legacy_authors.as_ref().unwrap_or(&empty_authors);

    let mut out = Vec::new();
    for cm in doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "cm")
    {
        let author_id = cm.attribute("authorId").unwrap_or("");
        let author = authors.get(author_id).cloned();
        let parsed_author_id = author_id.parse::<u32>().ok();
        let index = cm
            .attribute("idx")
            .and_then(|value| value.parse::<u32>().ok());
        let date = cm
            .attribute("dt")
            .map(String::from)
            .filter(|s| !s.is_empty());
        let position = cm
            .children()
            .find(|node| node.is_element() && node.tag_name().name() == "pos");
        let x = position
            .and_then(|node| node.attribute("x"))
            .and_then(|value| value.parse::<i64>().ok());
        let y = position
            .and_then(|node| node.attribute("y"))
            .and_then(|value| value.parse::<i64>().ok());
        let text = cm
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "text")
            .and_then(|n| n.text().map(String::from))
            .unwrap_or_default();
        out.push(PptxComment {
            author_id: parsed_author_id,
            modern_author_id: None,
            id: None,
            index,
            author,
            date,
            x,
            y,
            anchors: Vec::new(),
            status: None,
            text,
            replies: Vec::new(),
        });
    }
    out
}

fn parse_comment_authors(author_xml: Option<&str>) -> HashMap<String, String> {
    let mut authors: HashMap<String, String> = HashMap::new();
    if let Some(ax) = author_xml {
        if let Ok(adoc) = parse_preflighted_pptx_xml(ax) {
            for a in adoc
                .descendants()
                .filter(|n| n.is_element() && matches!(n.tag_name().name(), "cmAuthor" | "author"))
            {
                let id = a.attribute("id").unwrap_or("").to_string();
                let name = a.attribute("name").unwrap_or("").to_string();
                if !id.is_empty() && !name.is_empty() {
                    authors.insert(id, name);
                }
            }
        }
    }
    authors
}

// ===========================
//  Presentation parser
// ===========================

/// Open a pptx ZIP container, tagging a failure with the container part name.
///
/// #774 (RB7 MAJOR, symmetric with docx `parser::open_zip`): a truncated / corrupt
/// ZIP is the MOST COMMON way a pptx is broken (an incomplete download, a
/// byte-mangled attachment). `ZipArchive::new` maps that to an opaque
/// `zip::result::ZipError` that, if propagated, throws with no indication that the
/// CONTAINER (not some inner part) is the problem. Naming the failure lets the
/// caller build a `degraded_container_presentation` tagged with the container,
/// symmetric with how a corrupt slide part is tagged inside [`parse_presentation`].
#[cfg(test)]
pub(crate) fn open_zip(data: Vec<u8>) -> Result<PptxZip, String> {
    open_zip_with_limits(data, None, None)
}

fn open_zip_with_limits(
    data: Vec<u8>,
    max_archive_entry_bytes: Option<u64>,
    max_total_inflated_bytes: Option<u64>,
) -> Result<PptxZip, String> {
    open_zip_with_policy(
        data,
        max_archive_entry_bytes,
        max_total_inflated_bytes,
        None,
    )
}

fn open_zip_with_policy(
    data: Vec<u8>,
    max_archive_entry_bytes: Option<u64>,
    max_total_inflated_bytes: Option<u64>,
    max_archive_entries: Option<u64>,
) -> Result<PptxZip, String> {
    PackageSessionHandle::open(
        data,
        ooxml_common::resource::OoxmlFormat::Pptx,
        max_archive_entry_bytes,
        max_total_inflated_bytes,
        max_archive_entries,
    )
    .map(|session| PptxZip {
        session,
        operation: RetainedPackageOperation::new("pptx"),
    })
    .map_err(ooxml_common::zip::tag_container_error)
}

/// A placeholder [`Presentation`] for a pptx whose ZIP CONTAINER could not be
/// opened (truncated / corrupt / not a zip). No parts are readable, so there is
/// no theme to derive fonts / colors from — fall back to defaults and surface a
/// single placeholder slide carrying the container-tagged error. Mirrors the
/// per-slide [`broken_slide`] used inside [`parse_presentation`], but for the
/// whole-container case. Standard 16:9 slide size (12192000×6858000 EMU) so the
/// viewer paints a correctly-proportioned "could not be displayed" card.
///
/// `parse_error` is already tagged by [`open_zip`] (`"(zip container): {e}"`), so
/// it is set directly rather than routed through [`broken_slide`], which would
/// prefix its own `part` name and double-tag the message (`"(zip container):
/// (zip container): ..."`).
pub(crate) fn degraded_container_presentation(parse_error: String) -> Presentation {
    Presentation {
        slide_width: 12_192_000,
        slide_height: 6_858_000,
        slides: vec![Slide {
            index: 0,
            slide_number: 1,
            // A whole-container failure has no readable slide part, so there is
            // nothing an internal slide jump could resolve to — no part name.
            part_name: None,
            background: None,
            elements: Vec::new(),
            element_sources: Vec::new(),
            notes: None,
            comments: Vec::new(),
            hidden: false,
            parse_error: Some(parse_error),
        }],
        default_text_color: None,
        major_font: None,
        minor_font: None,
        hlink_color: None,
        fol_hlink_color: None,
    }
}

/// Parse a presentation from raw archive bytes. Thin wrapper that opens a fresh
/// owned [`PptxZip`] (copying `data`) and delegates to [`parse_presentation`].
/// Kept so the free `parse_pptx` / `pptx_to_markdown` WASM entry points and the
/// native `parse_pptx_native` path keep their `&[u8]` signature; the stateful
/// `PptxArchive` handle calls [`parse_presentation`] directly on its retained
/// archive to avoid re-opening it per call.
///
/// #774 (RB7 MAJOR): a corrupt / truncated CONTAINER degrades to a placeholder
/// presentation (`degraded_container_presentation`) rather than erroring,
/// consistent with a corrupt inner slide — the viewer shows a "could not display"
/// slide instead of nothing.
fn parse_presentation_from_bytes(data: &[u8]) -> Result<Presentation, Box<dyn std::error::Error>> {
    parse_presentation_from_bytes_with_limits(data, None, None, "parse").map_err(Into::into)
}

fn parse_presentation_from_bytes_with_limits(
    data: &[u8],
    max_archive_entry_bytes: Option<u64>,
    max_total_inflated_bytes: Option<u64>,
    operation: &str,
) -> Result<Presentation, String> {
    let mut zip = match open_zip_with_limits(
        data.to_vec(),
        max_archive_entry_bytes,
        max_total_inflated_bytes,
    ) {
        Ok(zip) => zip,
        Err(e) if e.starts_with("OOXML_RESOURCE_LIMIT:") => return Err(e),
        Err(e) => return Ok(degraded_container_presentation(e)),
    };
    zip.run_operation(operation, |zip| {
        parse_presentation(zip).map_err(|e| e.to_string())
    })
}

/// Open one package operation and project slides sequentially. Only the shared
/// PresentationML bootstrap, one canonical `Slide`, and the bounded markdown
/// result coexist; no full `Presentation` is materialized on this path.
fn render_markdown_from_bytes_with_limits(
    data: &[u8],
    max_archive_entry_bytes: Option<u64>,
    max_total_inflated_bytes: Option<u64>,
) -> Result<String, String> {
    let mut zip = match open_zip_with_limits(
        data.to_vec(),
        max_archive_entry_bytes,
        max_total_inflated_bytes,
    ) {
        Ok(zip) => zip,
        Err(error) if error.starts_with("OOXML_RESOURCE_LIMIT:") => return Err(error),
        Err(error) => {
            return Ok(render_presentation_md(&degraded_container_presentation(
                error,
            )))
        }
    };
    zip.run_operation("markdown", |zip| {
        let mut shared = bootstrap_presentation(zip).map_err(|error| error.to_string())?;
        render_markdown_from_shared(&mut shared, zip)
    })
}

fn render_markdown_from_shared(
    shared: &mut PresentationShared,
    zip: &mut PptxZip,
) -> Result<String, String> {
    let reporter = zip.operation()?.limit_reporter()?;
    let limit = pptx_internal_limits().markdown_bytes;
    let mut output = MarkdownWriter::new(limit);
    for index in 0..shared.slide_descriptors.len() {
        if index > 0 {
            output.push_str("\n---\n\n");
            reporter.observe_hard_limit(
                HardResourceLimitKind::PptxMarkdownBytes,
                None,
                limit,
                output.observed(),
            )?;
        }
        let produced = produce_slide_unit_with_journal(index, shared, zip, None)
            .map_err(|error| error.to_string())?;
        render_slide_md(&produced.slide, &mut output);
        reporter.observe_hard_limit(
            HardResourceLimitKind::PptxMarkdownBytes,
            None,
            limit,
            output.observed(),
        )?;
    }
    zip.assert_healthy()?;
    Ok(output.into_string())
}

fn extract_entry_with_limits(
    data: &[u8],
    path: &str,
    max_archive_entry_bytes: Option<u64>,
    max_total_inflated_bytes: Option<u64>,
    operation: &str,
) -> Result<Vec<u8>, String> {
    let mut zip = open_zip_with_limits(
        data.to_vec(),
        max_archive_entry_bytes,
        max_total_inflated_bytes,
    )?;
    zip.run_operation(operation, |zip| read_zip_bytes(zip, path))
}

fn parse_presentation(zip: &mut PptxZip) -> Result<Presentation, Box<dyn std::error::Error>> {
    let mut shared = bootstrap_presentation(zip)?;
    let mut slides = Vec::with_capacity(shared.slide_descriptors.len());
    for index in 0..shared.slide_descriptors.len() {
        slides.push(produce_slide_unit(index, &mut shared, zip)?);
    }
    shared.finish(slides)
}

#[cfg(test)]
fn serialize_slide_unit_with_limit(
    slide: &Slide,
    reporter: &PackageLimitReporter,
    limit: u64,
) -> Result<Vec<u8>, String> {
    let json_bytes = measure_json(slide)?.json_bytes;
    reporter.observe_hard_limit(
        HardResourceLimitKind::PptxSlideJsonBytes,
        slide.part_name.as_deref(),
        limit,
        json_bytes,
    )?;
    serde_json::to_vec(slide).map_err(|error| format!("serialize error: {error}"))
}

#[cfg(test)]
fn observe_primary_slide_xml(
    reporter: &PackageLimitReporter,
    part: &str,
    observed: u64,
    limit: u64,
) -> Result<(), String> {
    reporter.observe_hard_limit(
        HardResourceLimitKind::PptxSlideXmlBytes,
        Some(part),
        limit,
        observed,
    )
}

#[derive(Debug, Clone, PartialEq, Eq, serde::Serialize)]
struct SlideDescriptor {
    index: usize,
    relationship_id: Option<String>,
}

#[derive(serde::Serialize)]
struct LayoutSource {
    xml: Option<String>,
    rels: HashMap<String, String>,
    dir: String,
    master_path: Option<String>,
}

/// Owned presentation-wide bootstrap state. XML trees are intentionally not
/// retained: shared package parts are stored as owned strings or parsed models
/// so individual slides can be produced in any order without self-references.
struct PresentationShared {
    slide_width: i64,
    slide_height: i64,
    slide_descriptors: Vec<SlideDescriptor>,
    pres_rels: HashMap<String, String>,
    embedded_fonts: Vec<PptxEmbeddedFontRef>,
    theme: PptxTheme,
    comment_authors: Option<HashMap<String, String>>,
    comment_authors_path: Option<String>,
    modern_comment_authors: Option<HashMap<String, String>>,
    modern_comment_authors_path: Option<String>,
    pres_master_path: Option<String>,
    master_cache: HashMap<String, ParsedMaster>,
    no_master_bundle: Option<ParsedMaster>,
    layout_cache: HashMap<String, ParsedLayout>,
    layout_source_cache: HashMap<String, Rc<LayoutSource>>,
    cache_usage: SharedCacheUsage,
    materialized_slide_json_bytes: u64,
}

impl PresentationShared {
    fn finish(&self, slides: Vec<Slide>) -> Result<Presentation, Box<dyn std::error::Error>> {
        let default_text_color = self.theme.get("dk1").cloned();
        let major_font = self.theme.get("+mj-lt").cloned();
        let minor_font = self.theme.get("+mn-lt").cloned();
        let hlink_color = self.theme.get("hlink").cloned();
        let fol_hlink_color = self.theme.get("folHlink").cloned();
        Ok(Presentation {
            slide_width: self.slide_width,
            slide_height: self.slide_height,
            slides,
            default_text_color,
            major_font,
            minor_font,
            hlink_color,
            fol_hlink_color,
        })
    }
}

fn bootstrap_presentation(
    zip: &mut PptxZip,
) -> Result<PresentationShared, Box<dyn std::error::Error>> {
    // --- presentation.xml ---
    let pres_xml = read_zip_str(zip, "ppt/presentation.xml")?;
    let pres_doc = parse_preflighted_pptx_xml(&pres_xml)?;
    let pres_root = pres_doc.root_element();

    let sld_sz = child(pres_root, "sldSz");
    let slide_width = sld_sz.and_then(|n| attr_i64(&n, "cx")).unwrap_or(9_144_000);
    let slide_height = sld_sz.and_then(|n| attr_i64(&n, "cy")).unwrap_or(6_858_000);

    // Ordered slide relationship identities. Keep one slot per `p:sldId`
    // even when the required `r:id` is malformed or missing, so partial
    // degradation cannot shift slide indices or internal-jump targets.
    let reporter = zip.operation()?.limit_reporter()?;
    let bootstrap_limits = pptx_internal_limits();
    let mut slide_descriptors = Vec::new();
    if let Some(list) = child(pres_root, "sldIdLst") {
        for node in list
            .children()
            .filter(|node| node.is_element() && node.tag_name().name() == "sldId")
        {
            let observed = u64::try_from(slide_descriptors.len())
                .unwrap_or(u64::MAX)
                .saturating_add(1);
            reporter.observe_hard_limit(
                HardResourceLimitKind::PptxBootstrapSlides,
                Some("ppt/presentation.xml"),
                bootstrap_limits.bootstrap_slides,
                observed,
            )?;
            let index = slide_descriptors.len();
            slide_descriptors.push(SlideDescriptor {
                index,
                relationship_id: attr_r(&node, "id"),
            });
        }
    }

    // --- ppt/_rels/presentation.xml.rels ---
    let pres_rels_xml = read_zip_str(zip, "ppt/_rels/presentation.xml.rels")?;
    let pres_rels = parse_rels(&pres_rels_xml);
    let embedded_fonts = parse_embedded_font_refs(pres_root, &pres_rels_xml, zip);

    // --- Presentation-level theme colors ---
    // Used for the deck-wide defaults on `Presentation` (default text color,
    // major/minor fonts, hyperlink colors) and as the fallback theme for any
    // master that declares no /theme relationship of its own.
    let theme_path = find_rel_target_by_type(&pres_rels_xml, "/theme")
        .map(|target| resolve_path("ppt", &target));
    let theme = theme_path
        .as_deref()
        .map(|path| parse_theme_part(path, zip))
        .unwrap_or_default();

    // --- Presentation-level fallback master ---
    // The first slide master referenced by the presentation. Used for slides
    // whose layout→master→theme chain can't be resolved (simple/old decks), so
    // their behavior is unchanged from before per-slide resolution existed.
    let pres_master_path: Option<String> =
        find_rel_target_by_type(&pres_rels_xml, "/slideMaster").map(|t| resolve_path("ppt", &t));
    let comment_authors_path = find_internal_rel_target_by_types(
        &pres_rels_xml,
        CLASSIC_COMMENT_AUTHOR_RELATIONSHIP_TYPES,
    )
    .map(|target| resolve_path("ppt", &target));
    let modern_comment_authors_path =
        find_internal_rel_target_by_types(&pres_rels_xml, MODERN_COMMENT_AUTHOR_RELATIONSHIP_TYPES)
            .map(|target| resolve_path("ppt", &target));

    // This is a serialization-shaped projection of retained bootstrap state,
    // measured by a streaming writer without allocating a JSON buffer. It is a
    // deterministic safety proxy, not a claim about Rust allocator heap bytes.
    let bootstrap_projection = measure_json(&(
        slide_width,
        slide_height,
        &slide_descriptors,
        &pres_rels,
        &embedded_fonts,
        &theme,
        &pres_master_path,
        &comment_authors_path,
        &modern_comment_authors_path,
    ))?
    .json_bytes;
    reporter.observe_hard_limit(
        HardResourceLimitKind::PptxBootstrapProjectionBytes,
        Some("ppt/presentation.xml"),
        bootstrap_limits.bootstrap_projection_bytes,
        bootstrap_projection,
    )?;

    // Shared dependencies are loaded on the first slide that actually needs
    // them. Bootstrap therefore retains only compact presentation metadata and
    // never eagerly materializes the fallback master/theme inheritance bundle.
    let master_cache: HashMap<String, ParsedMaster> = HashMap::new();
    let no_master_bundle: Option<ParsedMaster> = None;

    // Cache of the layout single-pass extraction (`ParsedLayout`) keyed by layout
    // ZIP path (D4), mirroring `master_cache`. Slides sharing a layout reuse its
    // resolved placeholders + layout background + showMasterSp instead of
    // re-parsing the layout XML four times per slide. Only NO-override slides
    // populate/read the cache: the entry is resolved against the master's baked
    // theme, and a slide's layout→master chain is 1:1 (a layout names exactly one
    // master), so every no-override slide on a given layout shares that theme. A
    // slide with a `<p:clrMapOvr>` builds a fresh `ParsedLayout` against its
    // override theme instead (kept out of the cache).
    let layout_cache: HashMap<String, ParsedLayout> = HashMap::new();

    // Raw layout sources are independently cached by part path. This prevents
    // slides sharing a layout from re-inflating its XML and relationships while
    // still letting clrMapOvr slides derive a fresh ParsedLayout from the same
    // owned source.
    let layout_source_cache: HashMap<String, Rc<LayoutSource>> = HashMap::new();
    zip.assert_healthy()?;
    Ok(PresentationShared {
        slide_width,
        slide_height,
        slide_descriptors,
        pres_rels,
        embedded_fonts,
        theme,
        comment_authors: None,
        comment_authors_path,
        modern_comment_authors: None,
        modern_comment_authors_path,
        pres_master_path,
        master_cache,
        no_master_bundle,
        layout_cache,
        layout_source_cache,
        cache_usage: SharedCacheUsage::default(),
        materialized_slide_json_bytes: 0,
    })
}

/// Per-slide owned input, dropped after one model unit is produced.
struct SlideRaw {
    index: usize,
    slide_path: String,
    slide_dir: String,
    slide_xml: Result<String, String>,
    slide_rels_xml: String,
    slide_rels: HashMap<String, String>,
    smartart_drawings: HashMap<String, String>,
    layout_path: Option<String>,
    layout_source: Option<Rc<LayoutSource>>,
}

/// Produce exactly one slide model for an ordered presentation descriptor.
/// All public legacy parse paths drain this function in descriptor order.
fn produce_slide_unit(
    index: usize,
    shared: &mut PresentationShared,
    zip: &mut PptxZip,
) -> Result<Slide, Box<dyn std::error::Error>> {
    let produced = produce_slide_unit_with_journal(index, shared, zip, None)?;
    let projected = shared
        .materialized_slide_json_bytes
        .saturating_add(produced.json_bytes);
    let reporter = zip.operation()?.limit_reporter()?;
    reporter.observe_hard_limit(
        HardResourceLimitKind::PptxMaterializedSlideJsonBytes,
        produced.slide.part_name.as_deref(),
        pptx_internal_limits().materialized_slide_json_bytes,
        projected,
    )?;
    shared.materialized_slide_json_bytes = projected;
    Ok(produced.slide)
}

struct ProducedSlide {
    slide: Slide,
    json_bytes: u64,
}

/// The one canonical slide producer. Cursor callers provide a mutation journal
/// so only cache entries inserted by the unacknowledged slide can be rolled
/// back; legacy drains pass `None` and pay no journaling overhead.
fn produce_slide_unit_with_journal(
    index: usize,
    shared: &mut PresentationShared,
    zip: &mut PptxZip,
    mut journal: Option<&mut SlideCacheJournal>,
) -> Result<ProducedSlide, Box<dyn std::error::Error>> {
    let descriptor = shared
        .slide_descriptors
        .get(index)
        .ok_or_else(|| format!("slide index {index} is out of bounds"))?
        .clone();
    debug_assert_eq!(descriptor.index, index);
    let PresentationShared {
        pres_rels,
        theme,
        comment_authors,
        comment_authors_path,
        modern_comment_authors,
        modern_comment_authors_path,
        pres_master_path,
        master_cache,
        no_master_bundle,
        layout_cache,
        layout_source_cache,
        cache_usage,
        ..
    } = shared;
    let slide = 'produce: {
        let idx = index;
        let r_id = &descriptor.relationship_id;
        let Some(r_id) = r_id.as_deref() else {
            let part = format!("ppt/presentation.xml#sldId[{idx}]/@r:id");
            break 'produce broken_slide(idx, &part, "required slide relationship id is missing");
        };
        let rel_target = match pres_rels.get(r_id) {
            Some(t) => t.clone(),
            None => {
                let part = format!("ppt/presentation.xml#sldId[{idx}]/@r:id={r_id}");
                break 'produce broken_slide(idx, &part, "slide relationship target is missing");
            }
        };
        // Resolve via `resolve_path` (not `format!("ppt/{rel_target}")`) so a
        // package-root-absolute slide Target — e.g. `/ppt/slides/slide1.xml`
        // (leading slash, OPC / ECMA-376 Part 2 §9.3) — resolves correctly
        // instead of producing `ppt//ppt/slides/slide1.xml`. Relative targets
        // (the common `slides/slide1.xml`) are unaffected. Same fix class as
        // the chart-rel resolution above.
        let slide_path = resolve_path("ppt", &rel_target);
        let slide_dir = part_directory(&slide_path).to_owned();
        let rels_path = relationship_part_path(&slide_path);

        // RB7: a slide part that can't be read no longer aborts the whole deck.
        // Record the failure and let the build loop emit a placeholder for THIS
        // slide while the others parse normally.
        // This ceiling measures ONLY the primary slide-part XML bytes.
        // The bounded reader rejects limit+1 before UTF-8/DOM/model work; layout,
        // master, chart, notes, comments, and relationship inputs use the shared
        // dependency XML and DOM-complexity ceilings. Ordinary read/CRC/UTF-8
        // failures still become this slide's placeholder, while resource poison
        // returns immediately and cannot be downgraded.
        let slide_xml = match read_primary_slide_xml(zip, &slide_path) {
            Ok(xml) => Ok(xml),
            Err(error) => {
                zip.assert_healthy()?;
                Err(error)
            }
        };
        let slide_rels_xml = read_zip_str(zip, &rels_path).unwrap_or_default();
        let slide_rels = parse_rels(&slide_rels_xml);
        let smartart_drawings = build_smartart_drawings(&slide_rels_xml, &slide_dir, zip);

        // Layout XML
        let layout_path = find_rel_target_by_type(&slide_rels_xml, "/slideLayout")
            .map(|target| resolve_path(&slide_dir, &target));

        if let Some(path) = layout_path.as_deref() {
            if !layout_source_cache.contains_key(path) {
                let xml = read_zip_str(zip, path).ok();
                let dir = path
                    .rsplit_once('/')
                    .map(|(dir, _)| dir.to_owned())
                    .unwrap_or_else(|| "ppt/slideLayouts".to_owned());
                // Needed both for images inside the layout and for the
                // layout→slideMaster chain (ECMA-376 §19.3.1.43).
                let rels_path = relationship_part_path(path);
                let rels_xml = read_zip_str(zip, &rels_path).unwrap_or_default();
                let rels = parse_rels(&rels_xml);
                let master_path = find_rel_target_by_type(&rels_xml, "/slideMaster")
                    .map(|target| resolve_path(&dir, &target));
                let source = LayoutSource {
                    xml,
                    rels,
                    dir,
                    master_path,
                };
                let reporter = zip.operation()?.limit_reporter()?;
                observe_shared_cache_candidate(
                    &reporter,
                    Some(path),
                    &source,
                    cache_usage,
                    journal.as_deref_mut(),
                )?;
                layout_source_cache.insert(path.to_owned(), Rc::new(source));
                if let Some(journal) = journal.as_deref_mut() {
                    journal.inserted_layout_source_keys.push(path.to_owned());
                }
            }
        }
        let layout_source = layout_path
            .as_deref()
            .and_then(|path| layout_source_cache.get(path).cloned());

        let raw = SlideRaw {
            index: idx,
            slide_path,
            slide_dir,
            slide_xml,
            slide_rels_xml,
            slide_rels,
            smartart_drawings,
            layout_path,
            layout_source,
        };

        let empty_layout_rels = HashMap::new();
        let (layout_xml, layout_rels, layout_dir, master_path) = match raw.layout_source.as_deref()
        {
            Some(source) => (
                source.xml.as_deref(),
                &source.rels,
                source.dir.as_str(),
                source.master_path.as_deref(),
            ),
            None => (None, &empty_layout_rels, "ppt/slideLayouts", None),
        };

        // RB7: a slide part that couldn't be READ (recorded above) degrades to a
        // placeholder now, before any master/layout resolution touches it.
        let slide_xml = match &raw.slide_xml {
            Ok(xml) => xml.as_str(),
            Err(detail) => {
                break 'produce broken_slide(raw.index, &raw.slide_path, detail);
            }
        };
        // Resolve this slide's ParsedMaster: build (and cache) one for the
        // slide's own master when the layout→master chain resolved; otherwise
        // use the presentation-level fallback bundle. Building is keyed by
        // master path so slides sharing a master don't recompute.
        let resolved_master_path = master_path
            .filter(|path| !path.is_empty())
            .or(pres_master_path.as_deref());
        let bundle: &ParsedMaster = match resolved_master_path {
            Some(master_path) => {
                if !master_cache.contains_key(master_path) {
                    let candidate = build_master_bundle(master_path, theme, zip);
                    zip.assert_healthy()?;
                    let reporter = zip.operation()?.limit_reporter()?;
                    observe_shared_cache_candidate(
                        &reporter,
                        Some(master_path),
                        &candidate,
                        cache_usage,
                        journal.as_deref_mut(),
                    )?;
                    master_cache.insert(master_path.to_owned(), candidate);
                    if let Some(journal) = journal.as_deref_mut() {
                        journal.inserted_master_keys.push(master_path.to_owned());
                    }
                }
                &master_cache[master_path]
            }
            None => {
                if no_master_bundle.is_none() {
                    let candidate = build_master_bundle("", theme, zip);
                    zip.assert_healthy()?;
                    let reporter = zip.operation()?.limit_reporter()?;
                    observe_shared_cache_candidate(
                        &reporter,
                        Some("ppt/presentation.xml#fallback-master"),
                        &candidate,
                        cache_usage,
                        journal.as_deref_mut(),
                    )?;
                    *no_master_bundle = Some(candidate);
                }
                no_master_bundle
                    .as_ref()
                    .expect("fallback master initialized above")
            }
        };

        // Per-slide color-mapping override (ECMA-376 §19.3.1.7 clrMapOvr).
        // Precedence: the slide's own `<a:overrideClrMapping>` wins; else the
        // layout's; else inherit the master (`None`). `<a:masterClrMapping/>`
        // and an absent `<p:clrMapOvr>` both yield `None` at their level, so a
        // slide that explicitly inherits still falls through to the layout's
        // override — matching the slide→layout→master mapping chain.
        //
        // Why `<a:masterClrMapping/>` means "inherit (the layout)", NOT "bypass the
        // layout and use the master directly": §20.1.6.6 says masterClrMapping uses
        // "the color mapping defined in the master", and §19.3.1.7 likewise "the
        // color scheme defined by the master is used". Read in isolation that sounds
        // like a slide-level bypass — but Annex L.3.2.5 ("Slide Layouts") defines a
        // layout's Color Map Override as one that "overrides the inherited color
        // mapping from the slide master but IS INHERITED BY ALL PRESENTATION SLIDES
        // that utilize this layout." So once a layout overrides the master mapping,
        // the layout's mapping *is* the effective parent mapping the slide inherits;
        // "the master's mapping" for a slide on that layout already means the layout-
        // overridden one. PowerPoint additionally serializes `<a:masterClrMapping/>`
        // on ordinary non-overriding slides, so reading it as a layout bypass would
        // break layout-override inheritance for the common case. Hence both
        // masterClrMapping and an absent clrMapOvr resolve to `None` here and fall
        // through to the layout's override (then the master).
        let clr_map_ovr: Option<HashMap<String, String>> =
            parse_clr_map_ovr(slide_xml).or_else(|| layout_xml.and_then(parse_clr_map_ovr));
        // When an override applies, recompute the master's THEME-DEPENDENT fields
        // against the slide's effective mapping. §20.1.6.8 says the override is
        // used "in place of" the master's mapping for the whole slide, so master-
        // INHERITED scheme colors (the master `<p:bg>`, master txStyles placeholder
        // colors, master bullet colors) must flip together with the slide's own
        // shapes — not just the slide's effective `theme`.
        //
        // The effective theme is the master-baked theme with the override re-applied.
        // This is correct because `bake_clr_map` left the raw scheme slots
        // (dk1/lt1/dk2/lt2/accent1..6/hlink/folHlink) intact, so the override's slot
        // values resolve against the original palette (§20.1.6.8). The override
        // REPLACES the master's logical→slot mapping, not the master's already-baked
        // logical hexes (we re-apply over the raw slots).
        //
        // Documented limitation: if the master's clrMap non-identically remapped an
        // accent SLOT (e.g. accent1="accent2") AND an override targets that same
        // accent, the raw accent slot is still its own scheme value (bake only writes
        // logical keys), so the override resolves it from the intact scheme — correct.
        // The only unrecoverable case would be a master that *overwrote a raw slot key
        // itself*, which `bake_clr_map` never does.
        //
        // Built fully here (in `parse_presentation`, where `zip` is available — the
        // master `<p:bg>` may reference a blip) BEFORE `parse_slide`; `EffectiveMaster`
        // owns its data and holds no `zip` borrow, so the mutable borrow taken to
        // resolve `master_bg` ends before `parse_slide(zip)` is called.
        let effective_master: Option<EffectiveMaster> = clr_map_ovr.map(|ovr| {
            let mut theme = bundle.theme.clone();
            apply_clr_map(&mut theme, Some(&ovr));
            // Re-run the master's theme-dependent extractions (mirrors
            // build_master_bundle) against the effective override theme so master-
            // INHERITED scheme colors (the `<p:bg>` schemeClr, txStyles placeholder
            // colors, per-level bullet colors) flip with the override. Parse the
            // master XML ONCE here and share the root across all three re-resolutions
            // (previously each re-parsed the same string — 3 parses per override slide).
            let master_doc = bundle.master_xml.as_deref().and_then(|xml| {
                note_layout_master_parse();
                parse_preflighted_pptx_xml(xml).ok()
            });
            let master_root = master_doc.as_ref().map(|d| d.root_element());
            let master_bg: Option<Fill> = master_root.and_then(|root| {
                let c_sld = child(root, "cSld")?;
                let mut resolve = |rid: &str| -> Option<String> {
                    let target = bundle.master_rels.get(rid)?;
                    let path = resolve_path(&bundle.master_dir, target);
                    // Existence check only — central-directory lookup, no inflate
                    // (former `read_zip_bytes` decompressed the entry to discard it).
                    zip.index_for_name(&path)?;
                    Some(path)
                };
                parse_background(c_sld, &theme, &mut resolve)
            });
            let master_color = master_root
                .map(|root| parse_master_txstyle_color(root, &theme))
                .unwrap_or_default();
            let master_level_bullets = master_root
                .map(|root| {
                    parse_master_level_bullets(
                        root,
                        &theme,
                        &bundle.master_rels,
                        &bundle.master_dir,
                        zip,
                    )
                })
                .unwrap_or_default();
            EffectiveMaster {
                theme,
                master_bg,
                master_color,
                master_level_bullets,
            }
        });

        // Resolve this slide's `ParsedLayout` (placeholders + layout bg +
        // showMasterSp), parsing the layout XML once. Only the `theme` and the
        // master bullet colors that `parse_layout` consumes are theme-dependent;
        // a clrMapOvr slide passes the OVERRIDE-adjusted pair so its layout colors
        // flip with the override (mirrors the master theme-dependent recompute
        // above), everything else is the frozen bundle maps.
        let (layout_theme, layout_master_bullets): (&PptxTheme, &HashMap<String, LevelBullets>) =
            match effective_master.as_ref() {
                Some(e) => (&e.theme, &e.master_level_bullets),
                None => (&bundle.theme, &bundle.master_level_bullets),
            };
        // Build a `ParsedLayout` from a layout XML string with the resolved
        // theme/bullets and this bundle's remaining (theme-independent) maps.
        let build_parsed_layout = |lx: &str, zip: &mut PptxZip| -> ParsedLayout {
            parse_layout(
                lx,
                &bundle.master_font_sizes,
                &bundle.master_font_families,
                &bundle.master_level_font_sizes,
                &bundle.master_level_indents,
                layout_master_bullets,
                &bundle.master_anchors,
                &bundle.master_transforms,
                &bundle.master_alignments,
                &bundle.master_ea_ln_brk,
                &bundle.master_space_before,
                &bundle.master_space_after,
                &bundle.master_line_spacing,
                layout_theme,
                layout_dir,
                layout_rels,
                zip,
            )
        };

        // No-override slide WITH a layout path → cache by layout path (its entry
        // is resolved against the master-baked theme, which every no-override
        // slide on this layout shares). Otherwise build a fresh, uncached one
        // (override slide, or the rare no-layout-path case).
        let fresh_layout: Option<ParsedLayout> = match (
            effective_master.is_none(),
            layout_xml,
            raw.layout_path.as_deref(),
        ) {
            (true, Some(lx), Some(lp)) => {
                if !layout_cache.contains_key(lp) {
                    let pl = build_parsed_layout(lx, zip);
                    zip.assert_healthy()?;
                    let reporter = zip.operation()?.limit_reporter()?;
                    observe_shared_cache_candidate(
                        &reporter,
                        Some(lp),
                        &pl,
                        cache_usage,
                        journal.as_deref_mut(),
                    )?;
                    layout_cache.insert(lp.to_owned(), pl);
                    if let Some(journal) = journal.as_deref_mut() {
                        journal.inserted_layout_keys.push(lp.to_owned());
                    }
                }
                None // borrowed from the cache below
            }
            (_, Some(lx), _) => Some(build_parsed_layout(lx, zip)),
            (_, None, _) => Some(ParsedLayout::default()),
        };
        let parsed_layout: &ParsedLayout = match &fresh_layout {
            Some(pl) => pl,
            // Cached (no-override) path: `layout_path` is guaranteed present
            // because that is the only arm that leaves `fresh_layout` as `None`.
            // ast-grep-ignore: no-unwrap-in-parser-production
            None => &layout_cache[raw.layout_path.as_deref().unwrap()],
        };

        // RB7: a slide that reads but fails to PARSE (bad shape geometry, a
        // dependency it needs that can't be read, etc.) degrades to a placeholder
        // carrying the part-tagged error, so one broken slide never takes the
        // whole presentation down. Healthy slides are byte-for-byte unchanged.
        let had_comment_authors = comment_authors.is_some();
        let had_modern_comment_authors = modern_comment_authors.is_some();
        let slide = match parse_slide(
            slide_xml,
            &raw.slide_dir,
            &raw.slide_rels_xml,
            parsed_layout,
            layout_xml,
            layout_rels,
            layout_dir,
            bundle,
            effective_master.as_ref(),
            raw.index,
            &raw.slide_rels,
            &raw.smartart_drawings,
            comment_authors,
            comment_authors_path.as_deref(),
            modern_comment_authors,
            modern_comment_authors_path.as_deref(),
            zip,
        ) {
            Ok(slide) => slide,
            Err(e) => broken_slide(raw.index, &raw.slide_path, &e.to_string()),
        };
        zip.assert_healthy()?;
        if !had_comment_authors {
            if let Some(authors) = comment_authors.as_ref() {
                let reporter = zip.operation()?.limit_reporter()?;
                observe_shared_cache_candidate(
                    &reporter,
                    comment_authors_path.as_deref(),
                    authors,
                    cache_usage,
                    journal.as_deref_mut(),
                )?;
            }
        }
        if !had_modern_comment_authors {
            if let Some(authors) = modern_comment_authors.as_ref() {
                let reporter = zip.operation()?.limit_reporter()?;
                observe_shared_cache_candidate(
                    &reporter,
                    modern_comment_authors_path.as_deref(),
                    authors,
                    cache_usage,
                    journal,
                )?;
            }
        }
        // Stamp the resolved slide part path (e.g. `ppt/slides/slide3.xml`) so
        // the TS side can map an internal hyperlink slide jump to this index.
        // The build loop owns `raw.slide_path`; keying by it here (rather than
        // threading it through `parse_slide`) keeps that function's signature
        // untouched. `broken_slide` already set it, so re-stamping is a no-op there.
        let mut slide = slide;
        slide.part_name = Some(raw.slide_path.clone());
        break 'produce slide;
    };

    // Package-wide resource poison always outranks the compatibility
    // placeholder paths above. Check immediately before emitting the unit so a
    // failed `.ok()` / `unwrap_or_default()` read can never become a Slide.
    zip.assert_healthy()?;
    let reporter = zip.operation()?.limit_reporter()?;
    let json_bytes = measure_json(&slide)?.json_bytes;
    reporter.observe_hard_limit(
        HardResourceLimitKind::PptxSlideJsonBytes,
        slide.part_name.as_deref(),
        pptx_slide_json_limit(),
        json_bytes,
    )?;
    Ok(ProducedSlide { slide, json_bytes })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::chart::{parse_chartex, parse_legacy_chart};
    use ooxml_common::math::nodes_to_text;
    use std::io::Write;

    // Local-only sample (redistribution-prohibited, gitignored). Tests that
    // depend on it must skip gracefully on a clean checkout / in CI where the
    // file is absent. See packages/pptx/public/private/.
    const LOCAL_SAMPLE_2: &str = "../public/private/sample-2.pptx";

    struct SlideJsonLimitOverride(Option<u64>);

    impl SlideJsonLimitOverride {
        fn set(limit: u64) -> Self {
            let previous = PPTX_SLIDE_JSON_LIMIT_OVERRIDE.replace(Some(limit));
            Self(previous)
        }
    }

    impl Drop for SlideJsonLimitOverride {
        fn drop(&mut self) {
            PPTX_SLIDE_JSON_LIMIT_OVERRIDE.set(self.0);
        }
    }

    struct SlideXmlLimitOverride(Option<u64>);

    impl SlideXmlLimitOverride {
        fn set(limit: u64) -> Self {
            let previous = PPTX_SLIDE_XML_LIMIT_OVERRIDE.replace(Some(limit));
            Self(previous)
        }
    }

    impl Drop for SlideXmlLimitOverride {
        fn drop(&mut self) {
            PPTX_SLIDE_XML_LIMIT_OVERRIDE.set(self.0);
        }
    }

    struct InternalLimitsOverride(Option<PptxInternalLimits>);

    impl InternalLimitsOverride {
        fn set(limits: PptxInternalLimits) -> Self {
            let previous = PPTX_INTERNAL_LIMITS_OVERRIDE.replace(Some(limits));
            Self(previous)
        }
    }

    impl Drop for InternalLimitsOverride {
        fn drop(&mut self) {
            PPTX_INTERNAL_LIMITS_OVERRIDE.set(self.0);
        }
    }

    /// Build an empty in-memory zip — enough for parse_* functions that take a
    /// `&mut PptxZip` but whose input declares no `<a:buBlip>` / blipFill parts.
    fn empty_zip_bytes() -> Vec<u8> {
        let mut buf = Vec::new();
        {
            let cursor = Cursor::new(&mut buf);
            let writer = zip::ZipWriter::new(cursor);
            writer.finish().unwrap();
        }
        buf
    }

    /// Build an in-memory zip containing exactly `parts` (path → bytes). Used to
    /// prove a `<a:buBlip>` whose rId resolves to a part that ISN'T in the
    /// archive falls through to Bullet::Inherit (index_for_name returns None).
    fn zip_with_parts(parts: &[(&str, &[u8])]) -> Vec<u8> {
        use std::io::Write;
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(Cursor::new(&mut buf));
            let o = zip::write::SimpleFileOptions::default();
            for (path, bytes) in parts {
                w.start_file(*path, o).unwrap();
                w.write_all(bytes).unwrap();
            }
            w.finish().unwrap();
        }
        buf
    }

    #[test]
    fn embedded_font_list_resolves_styles_relationships_and_content_types() {
        let p_ns = "http://schemas.openxmlformats.org/presentationml/2006/main";
        let a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main";
        let r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        let rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships";
        let ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types";
        let font_rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font";
        let strict_font_rel = "http://purl.oclc.org/ooxml/officeDocument/relationships/font";

        let presentation = format!(
            r#"
          <p:presentation xmlns:p="{p_ns}" xmlns:a="{a_ns}" xmlns:r="{r_ns}">
            <p:embeddedFontLst>
              <p:embeddedFont><p:font typeface="Deck Sans"/>
                <p:regular r:id="rRegular"/><p:bold r:id="rBold"/>
                <p:italic r:id="rItalic"/><p:boldItalic r:id="rBoldItalic"/>
              </p:embeddedFont>
              <p:embeddedFont><p:font typeface="External"/><p:regular r:id="rExternal"/></p:embeddedFont>
              <p:embeddedFont><p:font typeface="Wrong Type"/><p:regular r:id="rImage"/></p:embeddedFont>
              <p:embeddedFont><p:font typeface="Unsupported"/><p:regular r:id="rUnsupported"/></p:embeddedFont>
              <p:embeddedFont><a:font typeface="Wrong Namespace"/><p:regular r:id="rWrongNamespace"/></p:embeddedFont>
            </p:embeddedFontLst>
            <p:sldIdLst/><p:sldSz cx="9144000" cy="6858000"/>
          </p:presentation>"#
        );
        let relationships = format!(
            r#"
          <Relationships xmlns="{rel_ns}">
            <Relationship Id="rRegular" Type="{font_rel}" Target="fonts/font1.fntdata"/>
            <Relationship Id="rBold" Type="{font_rel}" Target="fonts/font2.bin"/>
            <Relationship Id="rItalic" Type="{font_rel}" Target="/ppt/fonts/font3.eot"/>
            <Relationship Id="rBoldItalic" Type="{strict_font_rel}" Target="fonts/font4.fntdata"/>
            <Relationship Id="rExternal" Type="{font_rel}" Target="https://example.test/font.ttf" TargetMode="External"/>
            <Relationship Id="rImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="fonts/image.ttf"/>
            <Relationship Id="rUnsupported" Type="{font_rel}" Target="fonts/font5.woff"/>
            <Relationship Id="rWrongNamespace" Type="{font_rel}" Target="fonts/font6.fntdata"/>
          </Relationships>"#
        );
        let content_types = format!(
            r#"
          <Types xmlns="{ct_ns}">
            <Default Extension="fntdata" ContentType="application/x-font-ttf"/>
            <Default Extension="eot" ContentType="application/x-fontdata"/>
            <Default Extension="woff" ContentType="font/woff"/>
            <Override PartName="/ppt/fonts/font2.bin" ContentType="application/x-fontdata"/>
          </Types>"#
        );
        let bytes = zip_with_parts(&[
            ("[Content_Types].xml", content_types.as_bytes()),
            ("ppt/presentation.xml", presentation.as_bytes()),
            ("ppt/_rels/presentation.xml.rels", relationships.as_bytes()),
            ("ppt/fonts/font1.fntdata", b"regular"),
            ("ppt/fonts/font2.bin", b"bold"),
            ("ppt/fonts/font3.eot", b"italic"),
            ("ppt/fonts/font4.fntdata", b"bold-italic"),
            ("ppt/fonts/image.ttf", b"not-a-font-rel"),
            ("ppt/fonts/font5.woff", b"unsupported"),
            ("ppt/fonts/font6.fntdata", b"wrong-namespace"),
        ]);
        let mut zip = PptxZip::new(Cursor::new(bytes)).unwrap();
        let shared = bootstrap_presentation(&mut zip).unwrap();

        assert_eq!(
            shared.embedded_fonts,
            vec![
                PptxEmbeddedFontRef {
                    font_name: "Deck Sans".into(),
                    style: EmbeddedFontStyle::Regular,
                    part_path: "ppt/fonts/font1.fntdata".into(),
                    content_type: "application/x-font-ttf".into()
                },
                PptxEmbeddedFontRef {
                    font_name: "Deck Sans".into(),
                    style: EmbeddedFontStyle::Bold,
                    part_path: "ppt/fonts/font2.bin".into(),
                    content_type: "application/x-fontdata".into()
                },
                PptxEmbeddedFontRef {
                    font_name: "Deck Sans".into(),
                    style: EmbeddedFontStyle::Italic,
                    part_path: "ppt/fonts/font3.eot".into(),
                    content_type: "application/x-fontdata".into()
                },
                PptxEmbeddedFontRef {
                    font_name: "Deck Sans".into(),
                    style: EmbeddedFontStyle::BoldItalic,
                    part_path: "ppt/fonts/font4.fntdata".into(),
                    content_type: "application/x-font-ttf".into()
                },
            ],
        );
        let reporter = zip.operation().unwrap().limit_reporter().unwrap();
        let bootstrap = serialize_presentation_bootstrap(&shared, &reporter).unwrap();
        let json: serde_json::Value = serde_json::from_slice(&bootstrap).unwrap();
        assert_eq!(json["embeddedFonts"][0]["fontName"], "Deck Sans");
        assert_eq!(json["embeddedFonts"][3]["style"], "boldItalic");
        assert_eq!(
            json["embeddedFonts"][1]["contentType"],
            "application/x-fontdata"
        );
    }

    #[test]
    fn embedded_font_extraction_applies_its_limit_before_materialization() {
        let bytes = zip_with_parts(&[("ppt/fonts/font1.fntdata", b"12345")]);
        let zip = PptxZip::new(Cursor::new(bytes)).unwrap();
        assert_eq!(
            zip.read_font_part_with_limit("ppt/fonts/font1.fntdata", 5)
                .unwrap(),
            b"12345"
        );
        assert!(zip
            .read_font_part_with_limit("ppt/fonts/font1.fntdata", 4)
            .unwrap_err()
            .contains("optional-part byte limit"));
        assert!(zip.assert_healthy().is_ok());
    }

    #[test]
    fn strict_presentation_embedded_font_uses_p_font_element() {
        let presentation = r#"
          <p:presentation xmlns:p="http://purl.oclc.org/ooxml/presentationml/main"
            xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships">
            <p:embeddedFontLst><p:embeddedFont><p:font typeface="Strict Deck"/>
              <p:regular r:id="rFont"/>
            </p:embeddedFont></p:embeddedFontLst>
            <p:sldIdLst/><p:sldSz cx="9144000" cy="6858000"/>
          </p:presentation>"#;
        let relationships = r#"
          <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rFont"
              Type="http://purl.oclc.org/ooxml/officeDocument/relationships/font"
              Target="fonts/font1.fntdata"/>
          </Relationships>"#;
        let content_types = r#"
          <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="fntdata" ContentType="application/x-font-ttf"/>
          </Types>"#;
        let bytes = zip_with_parts(&[
            ("[Content_Types].xml", content_types.as_bytes()),
            ("ppt/presentation.xml", presentation.as_bytes()),
            ("ppt/_rels/presentation.xml.rels", relationships.as_bytes()),
            ("ppt/fonts/font1.fntdata", b"font"),
        ]);
        let mut zip = PptxZip::new(Cursor::new(bytes)).unwrap();
        let shared = bootstrap_presentation(&mut zip).unwrap();
        assert_eq!(shared.embedded_fonts.len(), 1);
        assert_eq!(shared.embedded_fonts[0].font_name, "Strict Deck");
    }

    // ECMA-376 §19.3.1.42 sldIdLst — each parsed slide is stamped with its
    // resolved OPC part name in presentation order, so the TS side can map an
    // internal hyperlink slide jump (§21.1.2.3.5) to a 0-based index. A minimal
    // two-slide deck proves the ordering and the exact normalized part name.
    #[test]
    fn slide_part_name_stamped_in_sldidlst_order() {
        let p_ns = "http://schemas.openxmlformats.org/presentationml/2006/main";
        let r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        let rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships";
        let ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types";

        let content_types = format!(
            r#"<Types xmlns="{ct_ns}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#
        );
        let root_rels = format!(
            r#"<Relationships xmlns="{rel_ns}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#
        );
        // sldIdLst references the two slides via rId1/rId2 (presentation order).
        let presentation = format!(
            r#"<p:presentation xmlns:p="{p_ns}" xmlns:r="{r_ns}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#
        );
        let pres_rels = format!(
            r#"<Relationships xmlns="{rel_ns}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/></Relationships>"#
        );
        let slide_xml = format!(r#"<p:sld xmlns:p="{p_ns}"><p:cSld><p:spTree/></p:cSld></p:sld>"#);

        let bytes = zip_with_parts(&[
            ("[Content_Types].xml", content_types.as_bytes()),
            ("_rels/.rels", root_rels.as_bytes()),
            ("ppt/presentation.xml", presentation.as_bytes()),
            ("ppt/_rels/presentation.xml.rels", pres_rels.as_bytes()),
            ("ppt/slides/slide1.xml", slide_xml.as_bytes()),
            ("ppt/slides/slide2.xml", slide_xml.as_bytes()),
        ]);

        let pres = parse_presentation_from_bytes(&bytes).expect("deck parses");
        assert_eq!(pres.slides.len(), 2);
        assert_eq!(
            pres.slides[0].part_name.as_deref(),
            Some("ppt/slides/slide1.xml")
        );
        assert_eq!(
            pres.slides[1].part_name.as_deref(),
            Some("ppt/slides/slide2.xml")
        );
    }

    #[test]
    fn slide_show_attr_marks_hidden() {
        let ns = r#"xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main""#;
        let parse = |attr: &str| {
            let xml = format!(r#"<p:sld {ns} {attr}><p:cSld/></p:sld>"#);
            let doc = roxmltree::Document::parse(&xml).unwrap();
            slide_is_hidden(doc.root_element())
        };
        // Absent `show` ⇒ shown (default true ⇒ not hidden).
        assert!(!parse(""));
        // `show="0"` / `show="false"` ⇒ hidden (ECMA-376 §19.3.1.38 CT_Slide).
        assert!(parse(r#"show="0""#));
        assert!(parse(r#"show="false""#));
        // Explicit truthy ⇒ shown.
        assert!(!parse(r#"show="1""#));
        assert!(!parse(r#"show="true""#));
    }

    /// A SmartArt data part's `<dsp:dataModelExt relId>` (MS-ODRAWXML) is the
    /// authority for its cached drawing part — not the file-number suffix. This
    /// fixture deliberately CROSSES the numbering: `data1.xml`'s dataModelExt
    /// points at the drawing relationship whose target is `drawing2.xml`, and
    /// `data2.xml`'s at `drawing1.xml`. The old trailing-number heuristic would
    /// pair 1↔1 / 2↔2 (wrong); the relId path pairs them by the explicit link.
    #[test]
    fn build_smartart_drawings_uses_datamodelext_relid_not_filename() {
        // dsp namespace per MS-ODRAWXML.
        let dsp = "http://schemas.microsoft.com/office/drawing/2008/diagram";
        let data1 = format!(
            r#"<dsp:dataModel xmlns:dsp="{dsp}"><dsp:extLst><dsp:dataModelExt relId="rIdDrawB"/></dsp:extLst></dsp:dataModel>"#
        );
        let data2 = format!(
            r#"<dsp:dataModel xmlns:dsp="{dsp}"><dsp:extLst><dsp:dataModelExt relId="rIdDrawA"/></dsp:extLst></dsp:dataModel>"#
        );
        // Distinct sentinel content so we can assert which drawing was paired.
        let drawing1 = r#"<dsp:drawing>ONE</dsp:drawing>"#;
        let drawing2 = r#"<dsp:drawing>TWO</dsp:drawing>"#;
        let bytes = zip_with_parts(&[
            ("ppt/diagrams/data1.xml", data1.as_bytes()),
            ("ppt/diagrams/data2.xml", data2.as_bytes()),
            ("ppt/diagrams/drawing1.xml", drawing1.as_bytes()),
            ("ppt/diagrams/drawing2.xml", drawing2.as_bytes()),
        ]);
        let mut zip = PptxZip::new(Cursor::new(bytes)).unwrap();

        let rels = r#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdData1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/>
  <Relationship Id="rIdData2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data2.xml"/>
  <Relationship Id="rIdDrawA" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="../diagrams/drawing1.xml"/>
  <Relationship Id="rIdDrawB" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="../diagrams/drawing2.xml"/>
</Relationships>"#;

        let map = build_smartart_drawings(rels, "ppt/slides", &mut zip);
        // Keyed by the diagramData rel Id (= the slide's r:dm value).
        // data1 → dataModelExt relId rIdDrawB → drawing2.xml ("TWO").
        assert!(
            map.get("rIdData1").unwrap().contains("TWO"),
            "data1 must pair with drawing2 via dataModelExt relId, got {:?}",
            map.get("rIdData1")
        );
        // data2 → rIdDrawA → drawing1.xml ("ONE").
        assert!(
            map.get("rIdData2").unwrap().contains("ONE"),
            "data2 must pair with drawing1 via dataModelExt relId, got {:?}",
            map.get("rIdData2")
        );
    }

    /// When a data part lacks a `dataModelExt` (older/malformed file), the
    /// compatibility fallback pairs by file-number suffix.
    #[test]
    fn build_smartart_drawings_falls_back_to_filenumber_without_datamodelext() {
        // data1.xml has no extLst/dataModelExt at all.
        let data1 = r#"<dsp:dataModel xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"/>"#;
        let drawing1 = r#"<dsp:drawing>ONE</dsp:drawing>"#;
        let bytes = zip_with_parts(&[
            ("ppt/diagrams/data1.xml", data1.as_bytes()),
            ("ppt/diagrams/drawing1.xml", drawing1.as_bytes()),
        ]);
        let mut zip = PptxZip::new(Cursor::new(bytes)).unwrap();
        let rels = r#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdData1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/>
  <Relationship Id="rIdDraw1" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="../diagrams/drawing1.xml"/>
</Relationships>"#;
        let map = build_smartart_drawings(rels, "ppt/slides", &mut zip);
        assert!(
            map.get("rIdData1")
                .map(|s| s.contains("ONE"))
                .unwrap_or(false),
            "fallback must pair data1↔drawing1 by file number, got {:?}",
            map.get("rIdData1")
        );
    }

    #[test]
    fn resolve_path_handles_absolute_leading_slash_target() {
        // An OPC relationship Target may be a package-root-absolute part name
        // (leading "/"), e.g. `/ppt/charts/chart5.xml` as emitted by some
        // generators. It must resolve from the package root and ignore the
        // source part's directory (ECMA-376 Part 2 / OPC §9.3). Regression for
        // issue #556 where a chart with an absolute Target silently failed to
        // load (read_zip_str on `ppt/slides/ppt/charts/chart5.xml`) and the
        // slide rendered the chart as a blank area.
        assert_eq!(
            resolve_path("ppt/slides", "/ppt/charts/chart5.xml"),
            "ppt/charts/chart5.xml"
        );
        // Relative references are unaffected by the absolute-target handling.
        assert_eq!(
            resolve_path("ppt/slides", "../charts/chart1.xml"),
            "ppt/charts/chart1.xml"
        );
        assert_eq!(
            resolve_path("ppt/slideLayouts", "../slideMasters/slideMaster1.xml"),
            "ppt/slideMasters/slideMaster1.xml"
        );
    }

    #[test]
    fn resolve_path_resolves_slide_targets_from_package_root() {
        // Slide parts are resolved from the presentation rels with base "ppt".
        // The common Target is relative (`slides/slide1.xml`); a generator may
        // also emit a package-root-absolute Target (`/ppt/slides/slide1.xml`),
        // which must NOT become `ppt//ppt/slides/slide1.xml`. Guards the
        // `resolve_path("ppt", rel_target)` slide-loading path.
        assert_eq!(
            resolve_path("ppt", "slides/slide1.xml"),
            "ppt/slides/slide1.xml"
        );
        assert_eq!(
            resolve_path("ppt", "/ppt/slides/slide1.xml"),
            "ppt/slides/slide1.xml"
        );
    }

    #[test]
    fn legacy_chart_parses_multi_level_category_axis() {
        // A `<c:cat>` may carry its labels in a `<c:multiLvlStrCache>` (multi-
        // level category axis, ECMA-376 §21.2.2.95) whose `<c:pt>` live under
        // `<c:lvl>` children rather than directly under the cache. Before the
        // fix, category extraction only recognized strCache/numCache/strLit/
        // numLit, so categories came back empty; that collapsed the shared
        // point count to 1 (`categories.len().max(1)`) and truncated EVERY
        // series to a single value. For an area chart a single point is a zero-
        // width sliver — i.e. a blank plot (issue #556).
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
 <c:chart><c:plotArea><c:areaChart>
  <c:ser>
   <c:cat><c:multiLvlStrRef><c:f>S!$A$1:$A$3</c:f><c:multiLvlStrCache>
     <c:ptCount val="3"/>
     <c:lvl><c:pt idx="0"><c:v>Jan</c:v></c:pt><c:pt idx="1"><c:v>Feb</c:v></c:pt><c:pt idx="2"><c:v>Mar</c:v></c:pt></c:lvl>
   </c:multiLvlStrCache></c:multiLvlStrRef></c:cat>
   <c:val><c:numRef><c:f>S!$B$1:$B$3</c:f><c:numCache>
     <c:ptCount val="3"/>
     <c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>30</c:v></c:pt>
   </c:numCache></c:numRef></c:val>
  </c:ser>
 </c:areaChart></c:plotArea></c:chart>
</c:chartSpace>"#;
        let chart = parse_legacy_chart(xml, &HashMap::new())
            .expect("area chart with multi-level cat should parse");
        assert_eq!(chart.chart.chart_type, "area");
        assert_eq!(chart.chart.categories, vec!["Jan", "Feb", "Mar"]);
        assert_eq!(chart.chart.series.len(), 1);
        assert_eq!(
            chart.chart.series[0].values,
            vec![Some(10.0), Some(20.0), Some(30.0)],
            "all three points must survive — a multi-level cat must not truncate the series"
        );
    }

    #[test]
    fn legacy_chart_parses_per_point_dpt_colors_for_pie() {
        // `<c:dPt>` (§21.2.2.52) carries its point index in a CHILD `<c:idx val>`
        // element (ECMA-376 §21.2.2.84, CT_UnsignedInt), NOT an attribute on it.
        // Reading it as an attribute always missed, so every pie/doughnut slice
        // fell back to the series colour (a `<a:schemeClr>` that resolved to the
        // default accent) — issue #556 follow-up: slide-7 fills were wrong. The
        // dPt fill must come from `<c:spPr><a:solidFill>`, never the border
        // `<a:ln><a:solidFill>`.
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
 <c:chart><c:plotArea><c:pieChart>
  <c:ser>
   <c:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr>
   <c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:srgbClr val="0D9488"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="F9F9F9"/></a:solidFill></a:ln></c:spPr></c:dPt>
   <c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="14B8A6"/></a:solidFill></c:spPr></c:dPt>
   <c:cat><c:strRef><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strCache></c:strRef></c:cat>
   <c:val><c:numRef><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>60</c:v></c:pt><c:pt idx="1"><c:v>40</c:v></c:pt></c:numCache></c:numRef></c:val>
  </c:ser>
 </c:pieChart></c:plotArea></c:chart>
</c:chartSpace>"#;
        let chart = parse_legacy_chart(xml, &HashMap::new()).expect("pie should parse");
        assert_eq!(chart.chart.chart_type, "pie");
        let dpc = chart.chart.series[0]
            .data_point_colors
            .as_ref()
            .expect("per-slice dPt colours must be captured");
        assert_eq!(
            dpc,
            &vec![Some("0D9488".to_string()), Some("14B8A6".to_string())],
            "each slice takes its own <c:dPt><c:spPr> fill, not the series colour or the border"
        );
    }

    #[test]
    fn legacy_chart_data_labels_on_for_show_percent_only() {
        // Pie/doughnut decks commonly enable `<c:showPercent>` with
        // `<c:showVal val="0">` (ECMA-376 §21.2.2.187 / §21.2.2.189). The old
        // check looked at `showVal` alone, so the "54%/27%/…" slice labels in
        // sample-14 slide-7 never rendered. `show_data_labels` must be true
        // when EITHER flag is set.
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
 <c:chart><c:plotArea><c:pieChart>
  <c:dLbls><c:numFmt formatCode="0%" sourceLinked="0"/><c:showVal val="0"/><c:showPercent val="1"/><c:showCatName val="0"/></c:dLbls>
  <c:ser>
   <c:cat><c:strRef><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strCache></c:strRef></c:cat>
   <c:val><c:numRef><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>60</c:v></c:pt><c:pt idx="1"><c:v>40</c:v></c:pt></c:numCache></c:numRef></c:val>
  </c:ser>
 </c:pieChart></c:plotArea></c:chart>
</c:chartSpace>"#;
        let chart = parse_legacy_chart(xml, &HashMap::new()).expect("pie should parse");
        assert!(
            chart.chart.show_data_labels,
            "showPercent=1 must enable data labels even when showVal=0"
        );
    }

    #[test]
    fn legacy_chart_honors_chart_space_date1904() {
        // `<c:date1904/>` is a direct child of `<c:chartSpace>` (ECMA-376
        // §21.2.2.38). It must thread through parse_legacy_chart into
        // ChartModel.date1904 so date-format value labels resolve against the
        // 1904 epoch. Absent element ⇒ the 1900 default (false).
        let with = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
 <c:date1904/>
 <c:chart><c:plotArea><c:lineChart>
  <c:ser>
   <c:val><c:numRef><c:numCache><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val>
  </c:ser>
 </c:lineChart></c:plotArea></c:chart>
</c:chartSpace>"#;
        let chart = parse_legacy_chart(with, &HashMap::new()).expect("line chart should parse");
        assert!(
            chart.chart.date1904,
            "<c:date1904/> must set ChartModel.date1904 = true"
        );

        let without = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
 <c:chart><c:plotArea><c:lineChart>
  <c:ser>
   <c:val><c:numRef><c:numCache><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val>
  </c:ser>
 </c:lineChart></c:plotArea></c:chart>
</c:chartSpace>"#;
        let chart0 = parse_legacy_chart(without, &HashMap::new()).expect("line chart should parse");
        assert!(
            !chart0.chart.date1904,
            "absent <c:date1904> must leave the 1900 default"
        );
    }

    #[test]
    fn extract_image_reads_entry() {
        use std::io::{Cursor, Write};
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(Cursor::new(&mut buf));
            let o = zip::write::SimpleFileOptions::default();
            w.start_file("ppt/media/i.png", o).unwrap();
            w.write_all(b"X").unwrap();
            w.finish().unwrap();
        }
        assert_eq!(
            extract_image(&buf, "ppt/media/i.png", None, None).unwrap(),
            b"X"
        );
    }

    /// A `PictureElement` serializes its blip as a zip path + mime, never as an
    /// inlined base64 `data:` URL (lazy image-bytes pipeline, Stage 2.1).
    #[test]
    fn picture_element_serializes_path_not_data_url() {
        let pic = PictureElement {
            id: None,
            x: 0,
            y: 0,
            width: 100,
            height: 100,
            rotation: 0.0,
            flip_h: false,
            flip_v: false,
            image_path: "ppt/media/image1.png".to_owned(),
            mime_type: "image/png".to_owned(),
            svg_image_path: None,
            intrinsic_width_px: Some(64),
            intrinsic_height_px: Some(48),
            stroke: None,
            prst_geom: None,
            prst_adjust: None,
            src_rect: None,
            alpha: None,
            duotone: None,
            cust_geom: None,
            shadow: None,
            inner_shadow: None,
            glow: None,
            soft_edge: None,
            reflection: None,
            scene3d: None,
            sp3d: None,
        };
        let json = serde_json::to_string(&pic).unwrap();
        assert!(
            json.contains("\"imagePath\":\"ppt/media/image1.png\""),
            "expected camelCase imagePath; got {json}"
        );
        assert!(
            json.contains("\"mimeType\":\"image/png\""),
            "expected camelCase mimeType; got {json}"
        );
        assert!(
            json.contains("\"intrinsicWidthPx\":64") && json.contains("\"intrinsicHeightPx\":48"),
            "expected intrinsic size keys; got {json}"
        );
        assert!(
            !json.contains("\"dataUrl\""),
            "must not emit dataUrl; got {json}"
        );
        assert!(
            !json.contains(";base64,"),
            "must not inline base64; got {json}"
        );
    }

    /// A blip `Fill::Image` (the serialized core `ImageFill`) serializes a zip
    /// path + mime, never an inlined base64 `data:` URL (Stage 2.2).
    #[test]
    fn image_fill_serializes_path_not_data_url() {
        let fill = Fill::Image {
            image_path: "ppt/media/image2.jpeg".to_owned(),
            mime_type: "image/jpeg".to_owned(),
            svg_image_path: None,
            dpi: None,
            rot_with_shape: None,
            src_rect: None,
            fill_rect: None,
            stretch: true,
            tile: None,
            alpha: None,
            duotone: None,
        };
        let json = serde_json::to_string(&fill).unwrap();
        assert!(
            json.contains("\"imagePath\":\"ppt/media/image2.jpeg\""),
            "expected camelCase imagePath; got {json}"
        );
        assert!(
            json.contains("\"mimeType\":\"image/jpeg\""),
            "expected camelCase mimeType; got {json}"
        );
        assert!(
            json.contains("\"fillType\":\"image\""),
            "tag preserved; got {json}"
        );
        assert!(
            !json.contains("\"dataUrl\""),
            "must not emit dataUrl; got {json}"
        );
        assert!(
            !json.contains(";base64,"),
            "must not inline base64; got {json}"
        );
    }

    // Synthetic deck for placeholder alignment-inheritance regression
    // (real reproducer files can't be committed). slide_sp / layout_extra_sp are
    // injected into the slide spTree / layout spTree.
    // Default master txStyles: titleStyle centred, bodyStyle left, otherStyle right.
    const DEFAULT_TXSTYLES: &str = r#"<p:txStyles><p:titleStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr algn="l"><a:defRPr sz="2800"/></a:lvl1pPr></p:bodyStyle><p:otherStyle><a:lvl1pPr algn="r"><a:defRPr sz="1800"/></a:lvl1pPr></p:otherStyle></p:txStyles>"#;

    fn build_align_pptx(slide_sp: &str, layout_extra_sp: &str, master_txstyles: &str) -> Vec<u8> {
        use std::io::{Cursor, Write};
        let layout = format!(
            r#"<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree>
                <p:sp><p:nvSpPr><p:cNvPr id="2" name="Body 1"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
                  <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>
                {layout_extra_sp}
              </p:spTree></p:cSld>
            </p:sldLayout>"#
        );
        let master = format!(
            r#"<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp></p:spTree></p:cSld>{master_txstyles}</p:sldMaster>"#
        );
        let entries: &[(&str, String)] = &[
            ("ppt/presentation.xml", r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId2"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>"#.to_owned()),
            ("ppt/_rels/presentation.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>"#.to_owned()),
            ("ppt/slides/slide1.xml", format!(r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree>{slide_sp}</p:spTree></p:cSld></p:sld>"#)),
            ("ppt/slides/_rels/slide1.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>"#.to_owned()),
            ("ppt/slideLayouts/slideLayout1.xml", layout),
            ("ppt/slideLayouts/_rels/slideLayout1.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>"#.to_owned()),
            ("ppt/slideMasters/slideMaster1.xml", master),
            ("ppt/slideMasters/_rels/slideMaster1.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>"#.to_owned()),
            ("ppt/theme/theme1.xml", r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="t"><a:themeElements><a:clrScheme name="c"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="000000"/></a:dk2><a:lt2><a:srgbClr val="FFFFFF"/></a:lt2><a:accent1><a:srgbClr val="000000"/></a:accent1><a:accent2><a:srgbClr val="000000"/></a:accent2><a:accent3><a:srgbClr val="000000"/></a:accent3><a:accent4><a:srgbClr val="000000"/></a:accent4><a:accent5><a:srgbClr val="000000"/></a:accent5><a:accent6><a:srgbClr val="000000"/></a:accent6><a:hlink><a:srgbClr val="000000"/></a:hlink><a:folHlink><a:srgbClr val="000000"/></a:folHlink></a:clrScheme><a:fontScheme name="f"><a:majorFont><a:latin typeface="Arial"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme><a:fmtScheme name="s"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme></a:themeElements></a:theme>"#.to_owned()),
        ];
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(Cursor::new(&mut buf));
            let o = zip::write::SimpleFileOptions::default();
            for (name, body) in entries {
                w.start_file(*name, o).unwrap();
                w.write_all(body.as_bytes()).unwrap();
            }
            w.finish().unwrap();
        }
        buf
    }

    fn first_para_alignment(data: &[u8]) -> String {
        let json = parse_pptx_native(data).unwrap();
        let v: serde_json::Value = serde_json::from_str(&json).unwrap();
        v["slides"][0]["elements"][0]["textBody"]["paragraphs"][0]["alignment"]
            .as_str()
            .unwrap_or("<none>")
            .to_owned()
    }

    // body placeholder (idx=1) with no explicit algn anywhere except master bodyStyle="l".
    const BODY_SP: &str = r#"<p:sp><p:nvSpPr><p:cNvPr id="5" name="Text Placeholder 5"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>x</a:t></a:r></a:p></p:txBody></p:sp>"#;

    // An unrelated centred typeless placeholder (idx=10) in the layout — the leak source.
    const TYPELESS_CTR_SP: &str = r#"<p:sp><p:nvSpPr><p:cNvPr id="9" name="Centered obj"/><p:cNvSpPr/><p:nvPr><p:ph idx="10"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr algn="ctr"/></a:lstStyle><a:p/></p:txBody></p:sp>"#;

    #[test]
    fn align_inherit_body_no_layout_algn_is_left() {
        // Master bodyStyle="l", nothing in slide/layout → must resolve to left.
        assert_eq!(
            first_para_alignment(&build_align_pptx(BODY_SP, "", DEFAULT_TXSTYLES)),
            "l"
        );
    }

    #[test]
    fn align_inherit_body_ignores_unrelated_typeless_center() {
        // Layout has an unrelated centred typeless placeholder (idx=10). The body
        // placeholder (idx=1) must NOT borrow it; resolves to master bodyStyle "l".
        assert_eq!(
            first_para_alignment(&build_align_pptx(
                BODY_SP,
                TYPELESS_CTR_SP,
                DEFAULT_TXSTYLES
            )),
            "l"
        );
    }

    #[test]
    fn align_inherit_body_idx_no_master_default_ignores_typeless_leak() {
        // Residual-leak guard: master has NO bodyStyle algn (so by_idx_alignment is
        // not seeded for the body slot) AND the layout has an unrelated centred
        // typeless placeholder. The idx-bearing body must still fall to the spec
        // default "l", never the typeless sibling's "ctr".
        let no_body_algn = r#"<p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr><a:defRPr sz="2800"/></a:lvl1pPr></p:bodyStyle></p:txStyles>"#;
        assert_eq!(
            first_para_alignment(&build_align_pptx(BODY_SP, TYPELESS_CTR_SP, no_body_algn)),
            "l"
        );
    }

    #[test]
    fn align_inherit_title_from_master_txstyles_center() {
        // Master titleStyle="ctr", no algn in slide/layout title → resolves to ctr.
        let title_sp = r#"<p:sp><p:nvSpPr><p:cNvPr id="6" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>T</a:t></a:r></a:p></p:txBody></p:sp>"#;
        assert_eq!(
            first_para_alignment(&build_align_pptx(title_sp, "", DEFAULT_TXSTYLES)),
            "ctr"
        );
    }

    #[test]
    fn align_inherit_subtitle_from_master_bodystyle() {
        // subTitle (idx=2, no matching layout slot) routes through the bodyStyle
        // txStyles row via the type path → inherits "l".
        let sub_sp = r#"<p:sp><p:nvSpPr><p:cNvPr id="7" name="Subtitle 1"/><p:cNvSpPr/><p:nvPr><p:ph type="subTitle" idx="2"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>S</a:t></a:r></a:p></p:txBody></p:sp>"#;
        assert_eq!(
            first_para_alignment(&build_align_pptx(sub_sp, "", DEFAULT_TXSTYLES)),
            "l"
        );
    }

    // --- Cross-tier bullet sub-property cascade (buClr/buFont/buSz vs marker) ---
    //
    // ECMA-376 §21.1.2.4: CT_TextParagraphProperties carries FOUR independent
    // optional choice groups — bullet colour (buClr/buClrTx), size (buSzPct/
    // buSzPts/buSzTx), typeface (buFont/buFontTx) and marker (buChar/buAutoNum/
    // buBlip/buNone). Each inherits per-property across the master→layout→
    // txBody-lstStyle→pPr cascade, so a decoration declared in one tier must
    // survive onto a marker declared in another. These build a synthetic deck
    // whose master `bodyStyle` and slide paragraph `pPr` split the marker from
    // its decorations and assert the resolved paragraph bullet re-combines them.

    fn txstyles_body_lvl1(lvl1_inner: &str) -> String {
        format!(
            r#"<p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr>{lvl1_inner}<a:defRPr sz="2800"/></a:lvl1pPr></p:bodyStyle></p:txStyles>"#
        )
    }

    fn body_sp_with_ppr(ppr_inner: &str) -> String {
        format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="5" name="Text Placeholder 5"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:pPr>{ppr_inner}</a:pPr><a:r><a:t>x</a:t></a:r></a:p></p:txBody></p:sp>"#
        )
    }

    fn first_para_bullet(data: &[u8]) -> serde_json::Value {
        let json = parse_pptx_native(data).unwrap();
        let v: serde_json::Value = serde_json::from_str(&json).unwrap();
        v["slides"][0]["elements"][0]["textBody"]["paragraphs"][0]["bullet"].clone()
    }

    /// §21.1.2.4.4 (buClr) + §21.1.2.4.1 (buAutoNum): the master `bodyStyle`
    /// declares the auto-number MARKER while the slide paragraph declares only
    /// the COLOUR. PowerPoint resolves the two groups independently, so the
    /// inherited number must render in the paragraph's red. Whole-`Bullet`
    /// merging drops the colour (the pPr has no marker → `Bullet::Inherit`).
    #[test]
    fn bullet_cascade_slide_buclr_colours_inherited_autonum() {
        let data = build_align_pptx(
            &body_sp_with_ppr(r#"<a:buClr><a:srgbClr val="C00000"/></a:buClr>"#),
            "",
            &txstyles_body_lvl1(r#"<a:buAutoNum type="arabicPeriod"/>"#),
        );
        let b = first_para_bullet(&data);
        assert_eq!(
            b["type"], "autoNum",
            "marker inherited from master bodyStyle"
        );
        assert_eq!(
            b["color"], "C00000",
            "slide buClr must colour the inherited number"
        );
    }

    /// §21.1.2.4.3 (buChar) + §21.1.2.4.4 (buClr): the master declares a bullet
    /// char AND its red colour; the slide paragraph overrides only the char (no
    /// buClr). The new char must inherit the master's red — decorations are not
    /// bound to the tier that declared the marker.
    #[test]
    fn bullet_cascade_new_char_inherits_lower_tier_buclr() {
        let data = build_align_pptx(
            &body_sp_with_ppr(r#"<a:buChar char="›"/>"#),
            "",
            &txstyles_body_lvl1(
                r#"<a:buClr><a:srgbClr val="C00000"/></a:buClr><a:buChar char="•"/>"#,
            ),
        );
        let b = first_para_bullet(&data);
        assert_eq!(b["type"], "char");
        assert_eq!(b["char"], "›", "slide overrides the marker");
        assert_eq!(
            b["color"], "C00000",
            "colour inherited independently of the marker"
        );
    }

    /// §21.1.2.4.5 (buClrTx): an explicit `<a:buClrTx/>` on the slide paragraph
    /// BREAKS inheritance of the lower-tier buClr and makes the bullet follow
    /// the run's text colour. The inherited char marker survives, but the colour
    /// resolves to null (renderer follows run) — not the master's red.
    #[test]
    fn bullet_cascade_buclrtx_breaks_inherited_buclr() {
        let data = build_align_pptx(
            &body_sp_with_ppr(r#"<a:buClrTx/>"#),
            "",
            &txstyles_body_lvl1(
                r#"<a:buClr><a:srgbClr val="C00000"/></a:buClr><a:buChar char="•"/>"#,
            ),
        );
        let b = first_para_bullet(&data);
        assert_eq!(b["type"], "char");
        assert_eq!(b["char"], "•", "marker still inherited");
        assert_eq!(
            b["color"],
            serde_json::Value::Null,
            "buClrTx must follow text, blocking the inherited red"
        );
    }

    /// §21.1.2.4.6 (buFont) + §21.1.2.4.3 (buChar): the master declares the
    /// bullet font while the slide paragraph declares only a new char. The new
    /// char must inherit the master's font — the typeface group is independent
    /// of the marker group.
    #[test]
    fn bullet_cascade_new_char_inherits_lower_tier_bufont() {
        let data = build_align_pptx(
            &body_sp_with_ppr(r#"<a:buChar char="X"/>"#),
            "",
            &txstyles_body_lvl1(r#"<a:buFont typeface="Wingdings"/><a:buChar char="•"/>"#),
        );
        let b = first_para_bullet(&data);
        assert_eq!(b["type"], "char");
        assert_eq!(b["char"], "X");
        assert_eq!(
            b["fontFamily"], "Wingdings",
            "font inherited independently of the marker"
        );
    }

    /// §21.1.2.4.9 (buSzPct): the master declares the marker while the slide
    /// paragraph overrides only the char — the inherited size percentage must
    /// carry onto the new char (size group is independent of the marker group).
    #[test]
    fn bullet_cascade_new_char_inherits_lower_tier_buszpct() {
        let data = build_align_pptx(
            &body_sp_with_ppr(r#"<a:buChar char="X"/>"#),
            "",
            &txstyles_body_lvl1(r#"<a:buSzPct val="50000"/><a:buChar char="•"/>"#),
        );
        let b = first_para_bullet(&data);
        assert_eq!(b["type"], "char");
        assert_eq!(b["char"], "X");
        assert_eq!(
            b["sizePct"], 50.0,
            "size inherited independently of the marker"
        );
    }

    /// §21.1.2.4.10 (buSzPts): an absolute-point bullet size on the paragraph is
    /// parsed as `sizePts` (points; the `ST_TextFontSize` val 1800 = 18pt) and is
    /// exclusive with `sizePct` — a `buSzPts` marker carries no `sizePct` (the two
    /// are members of the same `EG_TextBulletSize` xsd:choice).
    #[test]
    fn bullet_buszpts_parsed_as_absolute_points() {
        let data = build_align_pptx(
            &body_sp_with_ppr(r#"<a:buSzPts val="1800"/><a:buChar char="X"/>"#),
            "",
            "",
        );
        let b = first_para_bullet(&data);
        assert_eq!(b["type"], "char");
        assert_eq!(b["char"], "X");
        assert_eq!(b["sizePts"], 18.0, "buSzPts val=1800 → 18pt");
        assert!(
            b["sizePct"].is_null(),
            "buSzPts is exclusive with buSzPct (same xsd:choice), got {:?}",
            b["sizePct"]
        );
    }

    /// §21.1.2.4.10 (buSzPts) cascade: the master declares the marker while the
    /// slide paragraph overrides only the char — a lower-tier `buSzPts` must carry
    /// onto the new char (the size group inherits independently of the marker).
    #[test]
    fn bullet_cascade_new_char_inherits_lower_tier_buszpts() {
        let data = build_align_pptx(
            &body_sp_with_ppr(r#"<a:buChar char="X"/>"#),
            "",
            &txstyles_body_lvl1(r#"<a:buSzPts val="1800"/><a:buChar char="•"/>"#),
        );
        let b = first_para_bullet(&data);
        assert_eq!(b["type"], "char");
        assert_eq!(b["char"], "X");
        assert_eq!(
            b["sizePts"], 18.0,
            "absolute-point size inherited independently of the marker"
        );
        assert!(b["sizePct"].is_null(), "no percent size in the cascade");
    }

    /// §21.1.2.4.9/.10 cross-group precedence: a higher-tier `buSzPts` BLOCKS a
    /// lower-tier `buSzPct` (both are the single size choice group, so the primary
    /// tier's explicit absolute size wins and the inherited percent is dropped).
    #[test]
    fn bullet_buszpts_blocks_lower_tier_buszpct() {
        let data = build_align_pptx(
            &body_sp_with_ppr(r#"<a:buSzPts val="2400"/><a:buChar char="X"/>"#),
            "",
            &txstyles_body_lvl1(r#"<a:buSzPct val="50000"/><a:buChar char="•"/>"#),
        );
        let b = first_para_bullet(&data);
        assert_eq!(b["type"], "char");
        assert_eq!(b["char"], "X");
        assert_eq!(
            b["sizePts"], 24.0,
            "primary-tier buSzPts wins the size group"
        );
        assert!(
            b["sizePct"].is_null(),
            "the inherited lower-tier buSzPct is blocked, got {:?}",
            b["sizePct"]
        );
    }

    /// §21.1.2.4.8 (buNone): an explicit `<a:buNone/>` on the slide paragraph
    /// suppresses the inherited master marker — the marker group's explicit
    /// "no bullet" value blocks a lower-tier char. Regression guard for the
    /// property-wise cascade (buNone must still win the marker group).
    #[test]
    fn bullet_cascade_slide_bunone_suppresses_inherited_marker() {
        let data = build_align_pptx(
            &body_sp_with_ppr(r#"<a:buNone/>"#),
            "",
            &txstyles_body_lvl1(r#"<a:buChar char="•"/>"#),
        );
        let b = first_para_bullet(&data);
        assert_eq!(b["type"], "none", "buNone blocks the inherited marker");
    }

    /// Unit test for the property-wise merge + leaf collapse ([`BulletProps`]).
    /// Each of the four groups (marker/colour/font/size) inherits independently:
    /// `primary` wins every group it specifies (including a follow-text value,
    /// which BLOCKS the fallback), and the fallback supplies the rest.
    #[test]
    fn bullet_props_merge_and_resolve_field_wise() {
        let primary = BulletProps {
            marker: Some(BuMarker::Char("›".into())),
            color: Some(BuColor::FollowText),
            font: None,
            size: None,
        };
        let fallback = BulletProps {
            marker: Some(BuMarker::Char("•".into())),
            color: Some(BuColor::Color("C00000".into())),
            font: Some(BuFont::Font("Wingdings".into())),
            size: Some(BuSize::Pct(80.0)),
        };
        let merged = BulletProps::merge(&primary, &fallback);
        assert_eq!(merged.marker, Some(BuMarker::Char("›".into())));
        assert_eq!(
            merged.color,
            Some(BuColor::FollowText),
            "buClrTx blocks the fallback red"
        );
        assert_eq!(merged.font, Some(BuFont::Font("Wingdings".into())));
        assert_eq!(merged.size, Some(BuSize::Pct(80.0)));
        match merged.resolve() {
            Bullet::Char {
                ch,
                color,
                size_pct,
                size_pts,
                font_family,
            } => {
                assert_eq!(ch, "›");
                assert_eq!(
                    color, None,
                    "buClrTx collapses to follow-text (None) at the leaf"
                );
                assert_eq!(size_pct, Some(80.0));
                assert_eq!(size_pts, None, "a Pct size carries no absolute points");
                assert_eq!(font_family, Some("Wingdings".to_string()));
            }
            other => panic!("expected Char, got {other:?}"),
        }

        // A decoration-only props (no marker) is NOT inherit — it must survive
        // the cascade to reach an inherited marker.
        let dec_only = BulletProps {
            size: Some(BuSize::Pct(50.0)),
            ..Default::default()
        };
        assert!(!dec_only.is_inherit());
        assert!(BulletProps::default().is_inherit());
    }

    /// PowerPoint paints one marker glyph when a flattened SmartArt cache
    /// serializes a duplicated string-valued buChar. Preserve other authored
    /// multi-character values because ECMA-376 types the attribute as a string.
    #[test]
    fn character_bullet_collapses_only_a_duplicated_marker() {
        let doc = roxmltree::Document::parse(
            r#"<a:pPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:buChar char="••"/></a:pPr>"#,
        )
        .expect("valid paragraph properties");
        let theme = HashMap::new();
        let mut resolve_blip = |_: &str| None;
        match parse_bullet(Some(doc.root_element()), &theme, &mut resolve_blip) {
            Bullet::Char { ch, .. } => assert_eq!(ch, "•"),
            other => panic!("expected Char, got {other:?}"),
        }

        let multi_char_doc = roxmltree::Document::parse(
            r#"<a:pPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:buChar char="🡆x"/></a:pPr>"#,
        )
        .expect("valid multi-character marker");
        match parse_bullet(
            Some(multi_char_doc.root_element()),
            &theme,
            &mut resolve_blip,
        ) {
            Bullet::Char { ch, .. } => assert_eq!(ch, "🡆x"),
            other => panic!("expected Char, got {other:?}"),
        }
    }

    /// PowerPoint serializes an unbound placeholder as idx=2^32-1. That value
    /// is a sentinel rather than a layout slot, so the paragraph still inherits
    /// its bullet marker from the master bodyStyle.
    #[test]
    fn max_placeholder_idx_uses_type_style_fallback() {
        let slide_sp = r#"<p:sp><p:nvSpPr><p:cNvPr id="5" name="Unbound body"/><p:cNvSpPr/><p:nvPr><p:ph idx="4294967295"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:buClr><a:srgbClr val="C00000"/></a:buClr></a:pPr><a:r><a:t>x</a:t></a:r></a:p></p:txBody></p:sp>"#;
        let data = build_align_pptx(slide_sp, "", &txstyles_body_lvl1(r#"<a:buChar char="•"/>"#));
        let b = first_para_bullet(&data);
        assert_eq!(b["type"], "char");
        assert_eq!(b["char"], "•");
        assert_eq!(b["color"], "C00000");
    }

    /// §21.1.2.4.9 — `<a:buSzPct val>` accepts both the Transitional integer
    /// (thousandths of a percent, `"100000"` = 100%, what PowerPoint writes) and
    /// the Strict percentage string (`"111%"`, as in the spec example).
    #[test]
    fn parse_bu_sz_pct_accepts_both_lexical_forms() {
        assert_eq!(parse_bu_sz_pct("100000"), Some(100.0));
        assert_eq!(parse_bu_sz_pct("80000"), Some(80.0));
        assert_eq!(parse_bu_sz_pct("111%"), Some(111.0));
        assert_eq!(parse_bu_sz_pct(" 111% "), Some(111.0));
        assert_eq!(parse_bu_sz_pct("62.5%"), Some(62.5));
        assert_eq!(parse_bu_sz_pct("garbage"), None);
    }

    /// §21.1.2.4.10 — `<a:buSzPts val>` is `ST_TextFontSize` (hundredths of a
    /// point); `parse_bu_sz_pts` divides by 100 to yield points, like the run
    /// `sz`. A non-numeric value yields `None` (no size specified at this tier).
    #[test]
    fn parse_bu_sz_pts_is_hundredths_of_a_point() {
        assert_eq!(parse_bu_sz_pts("1800"), Some(18.0));
        assert_eq!(parse_bu_sz_pts("2400"), Some(24.0));
        assert_eq!(parse_bu_sz_pts(" 100 "), Some(1.0));
        assert_eq!(parse_bu_sz_pts("garbage"), None);
    }

    /// §21.1.2.4.10 — a `BuSize::Pts` size collapses to the resolved bullet's
    /// `size_pts` (absolute points) and leaves `size_pct` `None`; the two size
    /// fields are mutually exclusive at the leaf.
    #[test]
    fn bullet_props_pts_size_resolves_to_size_pts() {
        let props = BulletProps {
            marker: Some(BuMarker::Char("•".into())),
            size: Some(BuSize::Pts(18.0)),
            ..Default::default()
        };
        match props.resolve() {
            Bullet::Char {
                size_pct, size_pts, ..
            } => {
                assert_eq!(size_pts, Some(18.0), "Pts collapses to absolute points");
                assert_eq!(size_pct, None, "Pts carries no percent size");
            }
            other => panic!("expected Char, got {other:?}"),
        }
    }

    /// The master `or_insert` edge (codex review): a master body placeholder
    /// shape's `lstStyle` declares only the MARKER (buChar) while the generic
    /// `bodyStyle` txStyles declares only the COLOUR (buClr). Within the master
    /// tier the per-shape source (primary) must MERGE per-group over txStyles
    /// (fallback), not suppress it wholesale — so the marker inherits the colour.
    #[test]
    fn master_shape_lststyle_merges_txstyles_decoration() {
        let master = r#"<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld><p:spTree>
            <p:sp><p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
              <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:buChar char="•"/></a:lvl1pPr></a:lstStyle><a:p/></p:txBody></p:sp>
          </p:spTree></p:cSld>
          <p:txStyles>
            <p:bodyStyle><a:lvl1pPr><a:buClr><a:srgbClr val="C00000"/></a:buClr><a:defRPr sz="2800"/></a:lvl1pPr></p:bodyStyle>
          </p:txStyles>
        </p:sldMaster>"#;
        let theme = HashMap::new();
        let master_rels = HashMap::new();
        let bytes = empty_zip_bytes();
        let cursor = Cursor::new(bytes.clone());
        let mut zip = PptxZip::new(cursor).unwrap();
        let doc = roxmltree::Document::parse(master).unwrap();
        let m = parse_master_level_bullets(
            doc.root_element(),
            &theme,
            &master_rels,
            "ppt/slideMasters",
            &mut zip,
        );
        match m.get("body").map(|b| b[0].resolve()) {
            Some(Bullet::Char { ch, color, .. }) => {
                assert_eq!(ch, "•", "marker from the per-shape lstStyle");
                assert_eq!(
                    color.as_deref(),
                    Some("C00000"),
                    "colour merged from the bodyStyle txStyles fallback"
                );
            }
            other => panic!("expected merged char+colour, got {other:?}"),
        }
    }

    #[test]
    fn test_parse_chartex() {
        let Ok(data) = std::fs::read(LOCAL_SAMPLE_2) else {
            eprintln!("skipping test_parse_chartex: local sample not found");
            return;
        };
        let cursor = std::io::Cursor::new(data.clone());
        let mut zip = PptxZip::new(cursor).unwrap();
        let xml = read_zip_str(&mut zip, "ppt/charts/chartEx1.xml").unwrap();
        let theme = HashMap::new();
        let result = parse_chartex(&xml, None, None, &theme, None);
        println!("parse_chartex result: {:?}", result.is_some());
        if let Some(ref c) = result {
            println!("  chart_type: {}", c.chart.chart_type);
            println!("  categories: {:?}", c.chart.categories);
            println!("  series len: {}", c.chart.series.len());
            if !c.chart.series.is_empty() {
                println!("  values: {:?}", c.chart.series[0].values);
            }
            println!("  subtotal_indices: {:?}", c.chart.subtotal_indices);
        }
        assert!(result.is_some(), "parse_chartex should succeed");
    }

    #[test]
    fn test_slide8_chart_rid() {
        let Ok(data) = std::fs::read(LOCAL_SAMPLE_2) else {
            eprintln!("skipping test_slide8_chart_rid: local sample not found");
            return;
        };
        let cursor = std::io::Cursor::new(data.clone());
        let mut zip = PptxZip::new(cursor).unwrap();
        let slide_xml = read_zip_str(&mut zip, "ppt/slides/slide8.xml").unwrap();

        let doc = roxmltree::Document::parse(&slide_xml).unwrap();
        let root = doc.root_element();

        for gf in root
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "graphicFrame")
        {
            println!("Found graphicFrame");
            if let Some(gd) = gf
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "graphicData")
            {
                let uri = attr(&gd, "uri").unwrap_or_default();
                println!("  graphicData uri: {}", uri);
                if let Some(chart_node) = gd
                    .descendants()
                    .find(|n| n.is_element() && n.tag_name().name() == "chart")
                {
                    println!("  chart node found, tag: {:?}", chart_node.tag_name());
                    for a in chart_node.attributes() {
                        println!(
                            "  attr: name={} ns={:?} val={}",
                            a.name(),
                            a.namespace(),
                            a.value()
                        );
                    }
                    let rid = attr_r(&chart_node, "id");
                    println!("  attr_r id: {:?}", rid);
                }
            }
        }
    }

    #[test]
    fn test_slide8_full_parse() {
        let Ok(data) = std::fs::read(LOCAL_SAMPLE_2) else {
            eprintln!("skipping test_slide8_full_parse: local sample not found");
            return;
        };
        let pres = parse_presentation_from_bytes(&data).unwrap();
        let slide = &pres.slides[7]; // 0-indexed, slide 8
        println!("Slide 8 elements: {}", slide.elements.len());
        for (i, el) in slide.elements.iter().enumerate() {
            match el {
                SlideElement::Chart(c) => println!(
                    "  [{}] CHART type={} cats={}",
                    i,
                    c.chart.chart_type,
                    c.chart.categories.len()
                ),
                SlideElement::Shape(s) => println!("  [{}] shape x={}", i, s.x),
                SlideElement::Table(_) => println!("  [{}] table", i),
                SlideElement::Picture(_) => println!("  [{}] picture", i),
                SlideElement::Media(m) => println!("  [{}] media kind={}", i, m.media_kind),
            }
        }
    }

    #[test]
    fn test_slide8_chartex_pipeline() {
        let Ok(data) = std::fs::read(LOCAL_SAMPLE_2) else {
            eprintln!("skipping test_slide8_chartex_pipeline: local sample not found");
            return;
        };
        let cursor = std::io::Cursor::new(data.clone());
        let mut zip = PptxZip::new(cursor).unwrap();

        let rels_xml = read_zip_str(&mut zip, "ppt/slides/_rels/slide8.xml.rels").unwrap();
        let rels = parse_rels(&rels_xml);
        println!("rels: {:?}", rels);

        let chart_path = resolve_path("ppt/slides", "../charts/chartEx1.xml");
        println!("chart_path: {}", chart_path);

        let result = read_zip_str(&mut zip, &chart_path);
        println!("read_zip_str ok: {}", result.is_ok());

        if let Ok(chart_xml) = result {
            let theme = HashMap::new();
            let r = parse_chartex(&chart_xml, None, None, &theme, None);
            println!("parse_chartex: {:?}", r.is_some());
        }
    }

    /// ECMA-376 §21.1.2.3.5 — a:hlinkClick @r:id resolves via slide _rels Target.
    #[test]
    fn test_parse_run_hyperlink_resolves_rid() {
        let xml = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><rPr lang="en-US"><hlinkClick r:id="rId7"/></rPr><t>Open site</t></r>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let r_node = doc.root_element();
        let theme = HashMap::new();
        let mut rels = HashMap::new();
        rels.insert("rId7".to_owned(), "https://example.com/".to_owned());

        let parsed = parse_run(r_node, None, &theme, &rels).expect("run should parse");
        assert_eq!(parsed.text, "Open site");
        assert_eq!(parsed.hyperlink.as_deref(), Some("https://example.com/"));
    }

    /// ECMA-376 §21.1.2.3.5 — a:hlinkClick @action="ppaction://hlinksldjump"
    /// marks an INTERNAL slide jump. The r:id resolves to the internal slide
    /// part (TargetMode=Internal), and the raw action verb is carried through
    /// on `hyperlink_action` so the TS side can classify it as internal.
    #[test]
    fn test_parse_run_hyperlink_internal_slidejump_action() {
        let xml = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><rPr lang="en-US"><hlinkClick r:id="rId5" action="ppaction://hlinksldjump"/></rPr><t>Go to slide 3</t></r>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let mut rels = HashMap::new();
        rels.insert("rId5".to_owned(), "../slides/slide3.xml".to_owned());

        let parsed = parse_run(doc.root_element(), None, &theme, &rels).expect("run should parse");
        assert_eq!(parsed.hyperlink.as_deref(), Some("../slides/slide3.xml"));
        assert_eq!(
            parsed.hyperlink_action.as_deref(),
            Some("ppaction://hlinksldjump")
        );
    }

    /// An external URL hlinkClick (no @action) leaves hyperlink_action = None.
    #[test]
    fn test_parse_run_hyperlink_external_has_no_action() {
        let xml = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><rPr lang="en-US"><hlinkClick r:id="rId7"/></rPr><t>Open site</t></r>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let r_node = doc.root_element();
        let theme = HashMap::new();
        let mut rels = HashMap::new();
        rels.insert("rId7".to_owned(), "https://example.com/".to_owned());

        let parsed = parse_run(r_node, None, &theme, &rels).expect("run should parse");
        assert_eq!(parsed.hyperlink.as_deref(), Some("https://example.com/"));
        assert!(parsed.hyperlink_action.is_none());
    }

    /// A run without hlinkClick should have hyperlink = None.
    #[test]
    fn test_parse_run_without_hyperlink_is_none() {
        let xml = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr lang="en-US"/><t>plain</t></r>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let rels = HashMap::new();
        let parsed = parse_run(doc.root_element(), None, &theme, &rels).expect("run should parse");
        assert!(parsed.hyperlink.is_none());
    }

    /// hlinkClick with an unknown r:id should produce hyperlink = None
    /// rather than emitting a placeholder string.
    #[test]
    fn test_parse_run_hyperlink_unknown_rid_is_none() {
        let xml = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><rPr lang="en-US"><hlinkClick r:id="rIdNope"/></rPr><t>x</t></r>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let rels = HashMap::new();
        let parsed = parse_run(doc.root_element(), None, &theme, &rels).expect("run should parse");
        assert!(parsed.hyperlink.is_none());
    }

    /// ECMA-376 §20.1.8.40 — pattFill produces a Fill::Pattern carrying the
    /// preset name and the resolved fg/bg colours.
    #[test]
    fn test_parse_fill_pattern_extracts_fg_bg_preset() {
        let xml = r#"<spPr xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
            <pattFill prst="pct25">
                <fgClr><srgbClr val="C00000"/></fgClr>
                <bgClr><srgbClr val="FFFF00"/></bgClr>
            </pattFill>
        </spPr>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let fill = parse_fill(doc.root_element(), &theme).expect("pattFill should resolve");
        match fill {
            Fill::Pattern { fg, bg, preset } => {
                assert_eq!(preset, "pct25");
                assert_eq!(fg.to_uppercase(), "C00000");
                assert_eq!(bg.to_uppercase(), "FFFF00");
            }
            other => panic!("expected Fill::Pattern, got {:?}", other),
        }
    }

    /// pattFill missing fg/bg colours should fall back to black/white rather
    /// than dropping the fill entirely — keeps shapes recognisable when the
    /// theme cannot resolve the slot.
    #[test]
    fn test_parse_fill_pattern_defaults_when_colors_missing() {
        let xml = r#"<spPr xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
            <pattFill prst="horz"/>
        </spPr>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let fill = parse_fill(doc.root_element(), &theme).expect("pattFill should still resolve");
        match fill {
            Fill::Pattern { fg, bg, preset } => {
                assert_eq!(preset, "horz");
                assert_eq!(fg.to_lowercase(), "000000");
                assert_eq!(bg.to_lowercase(), "ffffff");
            }
            other => panic!("expected Fill::Pattern, got {:?}", other),
        }
    }

    /// ECMA-376 §20.1.2.3.34 defines tint as the retained fraction of the
    /// source colour. SmartArt writes explicit gradient stops that depend on
    /// that direction: a 15% tint must be much nearer white than a 50% tint.
    /// PowerPoint performs the blend in linear sRGB before applying satMod.
    #[test]
    fn test_parse_smartart_gradient_retains_tint_fraction_in_linear_srgb() {
        let xml = r#"<spPr xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
          <gradFill><gsLst>
            <gs pos="0"><schemeClr val="accent4"><tint val="50000"/><satMod val="300000"/></schemeClr></gs>
            <gs pos="35000"><schemeClr val="accent4"><tint val="37000"/><satMod val="300000"/></schemeClr></gs>
            <gs pos="100000"><schemeClr val="accent4"><tint val="15000"/><satMod val="350000"/></schemeClr></gs>
          </gsLst><lin ang="16200000" scaled="1"/></gradFill>
        </spPr>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::from([("accent4".to_owned(), "8064A2".to_owned())]);

        match parse_fill(doc.root_element(), &theme) {
            Some(Fill::Gradient { stops, .. }) => {
                assert_eq!(
                    stops
                        .iter()
                        .map(|stop| stop.color.as_str())
                        .collect::<Vec<_>>(),
                    ["C9B5E8", "D9CBEE", "F0EAF9"]
                );
            }
            other => panic!("expected SmartArt gradient, got {other:?}"),
        }
    }

    /// ECMA-376 §21.1.2.3.9; ST_TextStrikeType §20.1.10.79 —
    /// strike="dblStrike" produces strike_double=true,
    /// while strike="sngStrike" leaves strike_double=false. The plain
    /// `strikethrough` flag is true in both cases.
    #[test]
    fn test_parse_run_strike_double_distinguishes_dbl() {
        let dbl = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr strike="dblStrike"/><t>x</t></r>"#;
        let sng = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr strike="sngStrike"/><t>x</t></r>"#;
        let none = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr/><t>x</t></r>"#;
        let theme = HashMap::new();
        let rels = HashMap::new();

        let doc_d = roxmltree::Document::parse(dbl).unwrap();
        let r_d = parse_run(doc_d.root_element(), None, &theme, &rels).unwrap();
        assert!(r_d.strikethrough && r_d.strike_double);

        let doc_s = roxmltree::Document::parse(sng).unwrap();
        let r_s = parse_run(doc_s.root_element(), None, &theme, &rels).unwrap();
        assert!(r_s.strikethrough && !r_s.strike_double);

        let doc_n = roxmltree::Document::parse(none).unwrap();
        let r_n = parse_run(doc_n.root_element(), None, &theme, &rels).unwrap();
        assert!(!r_n.strikethrough && !r_n.strike_double);
    }

    /// CT_TextCharacterProperties is a sequence: the EG_TextRunProperties fill
    /// choice precedes latin/ea/cs. PowerPoint ignores a solidFill serialized
    /// after those font children instead of accepting the out-of-order value.
    /// A valid fill in its schema position must continue to resolve normally.
    #[test]
    fn test_parse_run_ignores_out_of_order_text_fill() {
        let valid = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr><solidFill><schemeClr val="accent1"/></solidFill><latin typeface="Aptos"/><ea typeface=""/><cs typeface=""/></rPr><t>valid</t></r>"#;
        let invalid = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr><latin typeface="Aptos"/><ea typeface=""/><cs typeface=""/><solidFill><schemeClr val="accent1"/></solidFill></rPr><t>invalid</t></r>"#;
        let theme = HashMap::from([("accent1".to_owned(), "1D6FA8".to_owned())]);
        let rels = HashMap::new();

        let valid_doc = roxmltree::Document::parse(valid).unwrap();
        let valid_run = parse_run(valid_doc.root_element(), None, &theme, &rels).unwrap();
        assert_eq!(valid_run.color.as_deref(), Some("1D6FA8"));

        let invalid_doc = roxmltree::Document::parse(invalid).unwrap();
        let invalid_run = parse_run(invalid_doc.root_element(), None, &theme, &rels).unwrap();
        assert_eq!(invalid_run.color, None);
    }

    /// ECMA-376 §21.1.2.3.9; ST_TextCapsType §20.1.10.64 — cap="all" /
    /// "small" are passed through;
    /// cap="none" or omitted yields None so the field stays absent in JSON.
    #[test]
    fn test_parse_run_caps_attribute() {
        let theme = HashMap::new();
        let rels = HashMap::new();
        let cases = [
            ("all", Some("all")),
            ("small", Some("small")),
            ("none", None),
        ];
        for (val, expected) in cases {
            let xml = format!(
                r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr cap="{val}"/><t>x</t></r>"#
            );
            let doc = roxmltree::Document::parse(&xml).unwrap();
            let r = parse_run(doc.root_element(), None, &theme, &rels).unwrap();
            assert_eq!(r.caps.as_deref(), expected, "caps={val}");
        }
    }

    /// ECMA-376 §21.1.2.3.9; ST_TextPoint §20.1.10.74 — a unitless rPr @spc
    /// encodes letter spacing in 100ths of a point; positive widens, negative
    /// tightens. Zero rounds away (None).
    #[test]
    fn test_parse_run_letter_spacing() {
        let theme = HashMap::new();
        let rels = HashMap::new();
        for (raw, expected) in [
            ("100", Some(1.0)),
            ("-50", Some(-0.5)),
            ("1pt", Some(1.0)),
            ("2.54cm", Some(72.0)),
            ("0", None),
        ] {
            let xml = format!(
                r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr spc="{raw}"/><t>x</t></r>"#
            );
            let doc = roxmltree::Document::parse(&xml).unwrap();
            let r = parse_run(doc.root_element(), None, &theme, &rels).unwrap();
            assert_eq!(r.letter_spacing, expected, "spc={raw}");
        }
    }

    /// ECMA-376 §20.1.8.21 — innerShdw shares the field shape of outerShdw
    /// (blurRad, dist, dir, color child). parse_inner_shadow should round-trip
    /// all of them, including the alphaModFix encoded as 8-char hex.
    #[test]
    fn test_parse_inner_shadow_extracts_fields() {
        let xml = r#"<effectLst xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
            <innerShdw blurRad="50800" dist="38100" dir="2700000">
                <srgbClr val="000000"><alphaModFix amt="50000"/></srgbClr>
            </innerShdw>
        </effectLst>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let s = parse_inner_shadow(doc.root_element(), &theme).expect("innerShdw should resolve");
        assert_eq!(s.blur, 50_800);
        assert_eq!(s.dist, 38_100);
        assert!((s.dir - 45.0).abs() < 0.001);
        assert!((s.alpha - 0.5).abs() < 0.01);
        assert_eq!(s.color.to_lowercase(), "000000");
    }

    /// ECMA-376 §20.1.8.14 + §20.1.8.58 + §20.1.8.30 — a `bgPr > blipFill`
    /// with a `stretch > fillRect` (incl. negative overscan edges) parses into
    /// `Fill::Image` carrying the resolved zip path + mime, the fractional
    /// source crop, fillRect, and alphaModFix alpha.
    #[test]
    fn test_parse_background_blip_fill_stretch() {
        let xml = r#"<p:cSld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <p:bg><p:bgPr>
                <a:blipFill>
                    <a:blip r:embed="rId2"><a:alphaModFix amt="80000"/></a:blip>
                    <a:srcRect l="25000" r="10000"/>
                    <a:stretch><a:fillRect t="-9000" b="-9000"/></a:stretch>
                </a:blipFill>
                <a:effectLst/>
            </p:bgPr></p:bg>
        </p:cSld>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let mut resolve = |rid: &str| -> Option<String> {
            assert_eq!(rid, "rId2");
            Some("ppt/media/image1.jpeg".to_owned())
        };
        let fill = parse_background(doc.root_element(), &theme, &mut resolve)
            .expect("blip background should resolve to Fill::Image");
        match fill {
            Fill::Image {
                image_path,
                mime_type,
                src_rect,
                fill_rect,
                tile,
                alpha,
                duotone: _,
                ..
            } => {
                assert_eq!(image_path, "ppt/media/image1.jpeg");
                assert_eq!(mime_type, "image/jpeg");
                let sr = src_rect.expect("srcRect should be present");
                assert!((sr.l - 0.25).abs() < 1e-9, "l={}", sr.l);
                assert!((sr.r - 0.1).abs() < 1e-9, "r={}", sr.r);
                let fr = fill_rect.expect("fillRect should be present");
                assert!((fr.t - (-0.09)).abs() < 1e-9, "t={}", fr.t);
                assert!((fr.b - (-0.09)).abs() < 1e-9, "b={}", fr.b);
                assert!(fr.l.abs() < 1e-9 && fr.r.abs() < 1e-9);
                assert!(tile.is_none(), "stretch fill must not carry tile");
                assert!((alpha.expect("alpha") - 0.8).abs() < 1e-6);
            }
            other => panic!("expected Fill::Image, got {other:?}"),
        }
    }

    /// ECMA-376 Part 1 §19.3.1.3: bgRef values 1001 and above index the
    /// theme's bgFillStyleLst (1001 = first entry). The referenced fill keeps
    /// its gradient geometry while each phClr is substituted with the bgRef
    /// colour before applying the style's colour transforms.
    #[test]
    fn test_parse_background_bg_ref_resolves_theme_style_matrix() {
        let theme_xml = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">
          <a:themeElements>
            <a:clrScheme name="C">
              <a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
              <a:dk2><a:srgbClr val="7DAFC3"/></a:dk2><a:lt2><a:srgbClr val="E5E4DF"/></a:lt2>
              <a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
              <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>
              <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>
              <a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
            </a:clrScheme>
            <a:fontScheme name="F"><a:majorFont><a:latin typeface="Arial"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>
            <a:fmtScheme name="S">
              <a:fillStyleLst/>
              <a:lnStyleLst/><a:effectStyleLst/>
              <a:bgFillStyleLst>
                <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
                <a:gradFill rotWithShape="1"><a:gsLst>
                  <a:gs pos="20000"><a:schemeClr val="phClr"><a:tint val="80000"/></a:schemeClr></a:gs>
                  <a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="80000"/></a:schemeClr></a:gs>
                </a:gsLst><a:path path="circle"/></a:gradFill>
              </a:bgFillStyleLst>
            </a:fmtScheme>
          </a:themeElements>
        </a:theme>"#;
        let mut theme = parse_theme_colors(theme_xml);
        theme.insert("bg2".to_owned(), "7DAFC3".to_owned());
        let background_xml = r#"<p:cSld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:bg><p:bgRef idx="1002"><a:schemeClr val="bg2"/></p:bgRef></p:bg>
        </p:cSld>"#;
        let doc = roxmltree::Document::parse(background_xml).unwrap();
        let mut resolve = |_rid: &str| -> Option<String> { None };
        let fill = parse_background(doc.root_element(), &theme, &mut resolve)
            .expect("bgRef 1002 should resolve the second bgFillStyleLst entry");
        match fill {
            Fill::Gradient {
                stops,
                grad_type,
                path,
                rot_with_shape,
                ..
            } => {
                assert_eq!(grad_type, "radial");
                assert_eq!(path.as_deref(), Some("circle"));
                assert_eq!(rot_with_shape, Some(true));
                assert_eq!(stops.len(), 2);
                assert_ne!(
                    stops[0].color, "7DAFC3",
                    "style transforms must be retained"
                );
                assert_ne!(
                    stops[1].color, "7DAFC3",
                    "style transforms must be retained"
                );
            }
            other => panic!("expected theme gradient, got {other:?}"),
        }
    }

    /// A style-matrix blipFill owns its relationship in the theme part. The
    /// slide/master relationship resolver must not be used for that embedded
    /// image when a bgRef selects the style.
    #[test]
    fn test_parse_background_bg_ref_resolves_theme_owned_image() {
        let theme_xml = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="T">
          <a:themeElements><a:clrScheme name="C"/>
            <a:fontScheme name="F"><a:majorFont/><a:minorFont/></a:fontScheme>
            <a:fmtScheme name="S"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/>
              <a:bgFillStyleLst><a:blipFill><a:blip r:embed="rIdImage"><a:duotone>
                <a:schemeClr val="phClr"/><a:srgbClr val="FFFFFF"/>
              </a:duotone></a:blip>
                <a:tile sx="95000" sy="95000" algn="t"/></a:blipFill></a:bgFillStyleLst>
            </a:fmtScheme>
          </a:themeElements>
        </a:theme>"#;
        let mut theme = parse_theme_colors(theme_xml);
        theme.insert(
            "+themeRel-rIdImage".to_owned(),
            "ppt/media/theme-background.jpeg".to_owned(),
        );
        theme.insert("bg2".to_owned(), "F0C000".to_owned());
        let background_xml = r#"<p:cSld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:bg><p:bgRef idx="1001"><a:schemeClr val="bg2"/></p:bgRef></p:bg>
        </p:cSld>"#;
        let doc = roxmltree::Document::parse(background_xml).unwrap();
        let mut wrong_part_resolver = |_rid: &str| -> Option<String> { None };

        match parse_background(doc.root_element(), &theme, &mut wrong_part_resolver) {
            Some(Fill::Image {
                image_path,
                tile,
                duotone,
                ..
            }) => {
                assert_eq!(image_path, "ppt/media/theme-background.jpeg");
                assert_eq!(tile.expect("tile descriptor").algn.as_deref(), Some("t"));
                let duotone = duotone.expect("placeholder-aware duotone");
                assert_eq!(duotone.clr1, "F0C000");
                assert_eq!(duotone.clr2, "FFFFFF");
            }
            other => panic!("expected theme-owned image fill, got {other:?}"),
        }
    }

    /// ECMA-376 §20.1.2.3.34 defines tint as retained input colour: an 80%
    /// tint keeps 80% of the source and adds 20% white. PowerPoint performs
    /// that blend in linear sRGB for ordinary presentation backgrounds as well
    /// as theme style-matrix fills.
    #[test]
    fn test_parse_background_uses_powerpoint_linear_tint_semantics() {
        let xml = r#"<p:cSld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg2"><a:tint val="80000"/></a:schemeClr></a:solidFill></p:bgPr></p:bg>
        </p:cSld>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::from([("bg2".to_owned(), "7DAFC3".to_owned())]);
        let mut resolve = |_rid: &str| -> Option<String> { None };
        match parse_background(doc.root_element(), &theme, &mut resolve) {
            Some(Fill::Solid { color }) => assert_eq!(color, "A3C3D1"),
            other => panic!("expected linear-tint solid fill, got {other:?}"),
        }
    }

    /// ECMA-376 §20.1.4.2.10: a shape fillRef selects fillStyleLst by its
    /// one-based idx and substitutes the reference colour for phClr. The style
    /// remains a gradient; reducing it to a solid accent loses the authored
    /// appearance (the failure reported for Apache POI customGeo.pptx).
    #[test]
    fn test_shape_fill_ref_resolves_theme_gradient() {
        let theme_xml = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">
          <a:themeElements><a:clrScheme name="C">
            <a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
            <a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
            <a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2>
            <a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4>
            <a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6>
            <a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink>
          </a:clrScheme><a:fontScheme name="F"><a:majorFont><a:latin typeface="Calibri"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>
          <a:fmtScheme name="S"><a:fillStyleLst>
            <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
            <a:gradFill><a:gsLst>
              <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/></a:schemeClr></a:gs>
              <a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/></a:schemeClr></a:gs>
            </a:gsLst><a:lin ang="16200000"/></a:gradFill>
          </a:fillStyleLst><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>
          </a:themeElements>
        </a:theme>"#;
        let theme = parse_theme_colors(theme_xml);
        let shape_xml = r#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:nvSpPr><p:cNvPr id="2" name="Styled ellipse"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom></p:spPr>
          <p:style><a:fillRef idx="2"><a:schemeClr val="accent1"/></a:fillRef></p:style>
        </p:sp>"#;
        let doc = roxmltree::Document::parse(shape_xml).unwrap();
        let fill_ref = doc
            .descendants()
            .find(|node| node.is_element() && node.tag_name().name() == "fillRef")
            .unwrap();
        assert!(matches!(
            parse_style_matrix_fill(fill_ref, &theme, false),
            Some(Fill::Gradient { .. })
        ));
        let mut zip = PptxZip::new(Cursor::new(empty_zip_bytes())).unwrap();
        let shape = parse_shape(
            doc.root_element(),
            &LayoutPlaceholders::default(),
            &theme,
            &HashMap::new(),
            "ppt/slides",
            None,
            &mut zip,
        )
        .expect("shape should parse");
        match shape.fill {
            Some(Fill::Gradient { stops, .. }) => {
                assert_eq!(stops.len(), 2);
                assert_eq!(stops[0].color, "C2CDE1");
                assert_eq!(stops[1].color, "EFF1F7");
            }
            other => panic!("expected style-matrix gradient, got {other:?}"),
        }
    }

    #[test]
    fn test_shape_fill_ref_uses_the_normative_background_style_index_range() {
        let theme_xml = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:themeElements><a:clrScheme name="C"/><a:fontScheme name="F"><a:majorFont/><a:minorFont/></a:fontScheme>
          <a:fmtScheme name="S"><a:fillStyleLst>
            <a:solidFill><a:srgbClr val="112233"/></a:solidFill>
          </a:fillStyleLst><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst>
            <a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill>
          </a:bgFillStyleLst></a:fmtScheme></a:themeElements>
        </a:theme>"#;
        let theme = parse_theme_colors(theme_xml);
        let refs = roxmltree::Document::parse(
            r#"<root xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:fillRef idx="1000"/><a:fillRef idx="1001"/>
            </root>"#,
        )
        .unwrap();
        let mut refs = refs
            .root_element()
            .children()
            .filter(|node| node.is_element());

        assert!(matches!(
            parse_style_matrix_fill(refs.next().unwrap(), &theme, false),
            Some(Fill::None)
        ));
        assert!(matches!(
            parse_style_matrix_fill(refs.next().unwrap(), &theme, false),
            Some(Fill::Solid { color }) if color == "ABCDEF"
        ));
    }

    /// ECMA-376 §19.3.1.52 / §21.1.2.3.7: a title with no local Latin
    /// typeface inherits titleStyle's +mj-lt, resolved through the current
    /// master's major Latin theme font.
    #[test]
    fn test_master_title_font_family_inherits_theme_major_latin() {
        let theme = parse_theme_colors(
            r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements>
              <a:clrScheme name="C"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="111111"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2><a:accent1><a:srgbClr val="111111"/></a:accent1><a:accent2><a:srgbClr val="222222"/></a:accent2><a:accent3><a:srgbClr val="333333"/></a:accent3><a:accent4><a:srgbClr val="444444"/></a:accent4><a:accent5><a:srgbClr val="555555"/></a:accent5><a:accent6><a:srgbClr val="666666"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme>
              <a:fontScheme name="F"><a:majorFont><a:latin typeface="Arial Black"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>
              <a:fmtScheme name="S"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>
            </a:themeElements></a:theme>"#,
        );
        let master_xml = r#"<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld><p:spTree/></p:cSld><p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></p:titleStyle></p:txStyles>
        </p:sldMaster>"#;
        let master_doc = roxmltree::Document::parse(master_xml).unwrap();
        let families = parse_master_font_families(master_doc.root_element(), &theme);
        assert_eq!(
            families.get("title").map(String::as_str),
            Some("Arial Black")
        );
        assert_eq!(
            families.get("ctrTitle").map(String::as_str),
            Some("Arial Black")
        );

        let placeholders = LayoutPlaceholders {
            by_type_font_family: families,
            ..LayoutPlaceholders::default()
        };
        let shape_xml = r#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm></p:spPr>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>Trade show</a:t></a:r></a:p></p:txBody>
        </p:sp>"#;
        let doc = roxmltree::Document::parse(shape_xml).unwrap();
        let mut zip = PptxZip::new(Cursor::new(empty_zip_bytes())).unwrap();
        let shape = parse_shape(
            doc.root_element(),
            &placeholders,
            &theme,
            &HashMap::new(),
            "ppt/slides",
            None,
            &mut zip,
        )
        .expect("title should parse");
        let paragraph = &shape.text_body.expect("text body").paragraphs[0];
        assert_eq!(paragraph.def_font_family.as_deref(), Some("Arial Black"));
    }

    /// ECMA-376 §20.1.8.23 — a background `<a:blipFill>` whose `<a:blip>` carries
    /// a `<a:duotone>` surfaces the resolved endpoint colours onto
    /// `Fill::Image.duotone` (through the theme), so a picture FILL recolours like
    /// a `<p:pic>`. Guards issue #889 (duotone was latent on the Fill::Image path).
    #[test]
    fn test_parse_background_blip_fill_duotone() {
        let xml = r#"<p:cSld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <p:bg><p:bgPr>
                <a:blipFill>
                    <a:blip r:embed="rId2">
                        <a:duotone>
                            <a:prstClr val="black"/>
                            <a:schemeClr val="accent1"/>
                        </a:duotone>
                    </a:blip>
                    <a:stretch><a:fillRect/></a:stretch>
                </a:blipFill>
            </p:bgPr></p:bg>
        </p:cSld>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let mut theme = HashMap::new();
        theme.insert("accent1".to_string(), "4472C4".to_string());
        let mut resolve = |rid: &str| -> Option<String> {
            assert_eq!(rid, "rId2");
            Some("ppt/media/image1.png".to_owned())
        };
        let fill = parse_background(doc.root_element(), &theme, &mut resolve)
            .expect("blip background should resolve to Fill::Image");
        match fill {
            Fill::Image { duotone, .. } => {
                let duo = duotone.expect("duotone must surface on the Fill::Image");
                assert_eq!(duo.clr1, "000000", "clr1 = black prstClr");
                assert_eq!(duo.clr2, "4472C4", "clr2 = accent1 resolved from theme");
            }
            other => panic!("expected Fill::Image, got {other:?}"),
        }
    }

    /// A background `<a:blipFill>` without a `<a:duotone>` leaves
    /// `Fill::Image.duotone` None, so non-duotone backgrounds stay byte-identical.
    #[test]
    fn test_parse_background_blip_fill_without_duotone_is_none() {
        let xml = r#"<p:cSld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <p:bg><p:bgPr>
                <a:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>
            </p:bgPr></p:bg>
        </p:cSld>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let mut resolve =
            |_rid: &str| -> Option<String> { Some("ppt/media/image1.png".to_owned()) };
        let fill = parse_background(doc.root_element(), &theme, &mut resolve)
            .expect("blip background should resolve to Fill::Image");
        match fill {
            Fill::Image { duotone, .. } => {
                assert!(duotone.is_none(), "duotone must be None when absent");
            }
            other => panic!("expected Fill::Image, got {other:?}"),
        }
    }

    /// ECMA-376 §20.1.8.14 + §20.1.8.58 — a `bgPr > blipFill` with `<a:tile>`
    /// parses into `Fill::Image` carrying a `TileInfo` (and no `fillRect`).
    /// tx/ty stay EMU, sx/sy convert ST_Percentage → fraction, flip/algn pass
    /// through verbatim.
    #[test]
    fn test_parse_background_blip_fill_tile() {
        let xml = r#"<p:cSld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <p:bg><p:bgPr>
                <a:blipFill>
                    <a:blip r:embed="rId2"/>
                    <a:tile tx="457200" ty="-228600" sx="50000" sy="75000" flip="xy" algn="ctr"/>
                </a:blipFill>
            </p:bgPr></p:bg>
        </p:cSld>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let mut resolve = |_: &str| -> Option<String> { Some("ppt/media/image1.png".to_owned()) };
        let fill = parse_background(doc.root_element(), &theme, &mut resolve)
            .expect("tiled blip background should resolve to Fill::Image");
        match fill {
            Fill::Image {
                fill_rect, tile, ..
            } => {
                assert!(fill_rect.is_none(), "tile fill must not carry fillRect");
                let t = tile.expect("tile should be present");
                assert_eq!(t.tx, Some(457_200));
                assert_eq!(t.ty, Some(-228_600));
                assert!(t.sx.is_some_and(|value| (value - 0.5).abs() < 1e-9));
                assert!(t.sy.is_some_and(|value| (value - 0.75).abs() < 1e-9));
                assert_eq!(t.flip.as_deref(), Some("xy"));
                assert_eq!(t.algn.as_deref(), Some("ctr"));
            }
            other => panic!("expected Fill::Image, got {other:?}"),
        }
    }

    /// §20.1.8.58: a bare `<a:tile/>` yields the schema defaults for tx/ty,
    /// sx/sy and flip, while omitted algn remains absent for the host policy.
    #[test]
    fn test_parse_background_blip_fill_tile_defaults() {
        let xml = r#"<p:cSld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <p:bg><p:bgPr>
                <a:blipFill><a:blip r:embed="rId2"/><a:tile/></a:blipFill>
            </p:bgPr></p:bg>
        </p:cSld>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let mut resolve = |_: &str| -> Option<String> { Some("ppt/media/image1.png".to_owned()) };
        let fill = parse_background(doc.root_element(), &theme, &mut resolve)
            .expect("bare tile should still resolve to Fill::Image");
        match fill {
            Fill::Image { tile, .. } => {
                let t = tile.expect("tile should be present");
                assert_eq!(t.tx, None);
                assert_eq!(t.ty, None);
                assert_eq!(t.sx, None);
                assert_eq!(t.sy, None);
                assert_eq!(t.flip.as_deref(), Some("none"));
                assert_eq!(t.algn, None);
            }
            other => panic!("expected Fill::Image, got {other:?}"),
        }
    }

    /// ECMA-376 §21.1.2.4.2 — a paragraph `<a:pPr><a:buBlip><a:blip r:embed>`
    /// resolves into `Bullet::Blip` carrying the blip's zip path + mime. The
    /// `<a:buSzPct val>` (§21.1.2.4.9, thousandths of a percent) becomes a plain
    /// percentage on the bullet.
    #[test]
    fn test_parse_bullet_blip_resolves_embed_and_size() {
        let xml = r#"<a:pPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <a:buSzPct val="80000"/>
            <a:buBlip><a:blip r:embed="rId5"/></a:buBlip>
        </a:pPr>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let mut resolve = |rid: &str| -> Option<String> {
            assert_eq!(rid, "rId5");
            Some("ppt/media/image3.png".to_owned())
        };
        match parse_bullet(Some(doc.root_element()), &theme, &mut resolve) {
            Bullet::Blip {
                image_path,
                mime_type,
                size_pct,
                size_pts,
            } => {
                assert_eq!(image_path, "ppt/media/image3.png");
                assert_eq!(mime_type, "image/png");
                assert!((size_pct.expect("size_pct") - 80.0).abs() < 1e-9);
                assert_eq!(size_pts, None, "a buSzPct blip carries no absolute points");
            }
            other => panic!("expected Bullet::Blip, got {other:?}"),
        }
    }

    /// §21.1.2.4.2 — with no `<a:buSzPct>` the picture bullet carries `None`
    /// size (renderer uses the spec default of 100%), and the mime tracks the
    /// resolved extension (jpeg here).
    #[test]
    fn test_parse_bullet_blip_default_size_and_mime() {
        let xml = r#"<a:pPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <a:buBlip><a:blip r:embed="rId2"/></a:buBlip>
        </a:pPr>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let mut resolve = |_: &str| -> Option<String> { Some("ppt/media/image1.jpeg".to_owned()) };
        match parse_bullet(Some(doc.root_element()), &theme, &mut resolve) {
            Bullet::Blip {
                image_path,
                mime_type,
                size_pct,
                size_pts,
            } => {
                assert_eq!(image_path, "ppt/media/image1.jpeg");
                assert_eq!(mime_type, "image/jpeg");
                assert!(size_pct.is_none());
                assert!(size_pts.is_none());
            }
            other => panic!("expected Bullet::Blip, got {other:?}"),
        }
    }

    /// §21.1.2.4 — the bullet element is an `xsd:choice`: an explicit
    /// `<a:buNone>` wins even when a stray `<a:buBlip>` is also present (the
    /// paragraph draws no marker). Mirrors the buNone-over-buChar precedence.
    #[test]
    fn test_parse_bullet_none_wins_over_blip() {
        let xml = r#"<a:pPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <a:buNone/>
            <a:buBlip><a:blip r:embed="rId2"/></a:buBlip>
        </a:pPr>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let mut resolve = |_: &str| -> Option<String> { Some("ppt/media/image1.png".to_owned()) };
        assert!(matches!(
            parse_bullet(Some(doc.root_element()), &theme, &mut resolve),
            Bullet::None
        ));
    }

    /// §21.1.2.4.2 — a `<a:buBlip>` whose `r:embed` cannot be resolved (dangling
    /// relationship) must NOT emit a half-built picture bullet. It falls through
    /// to `Bullet::Inherit` so a lower style tier can still supply a marker.
    #[test]
    fn test_parse_bullet_blip_dangling_embed_inherits() {
        let xml = r#"<a:pPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <a:buBlip><a:blip r:embed="rIdMissing"/></a:buBlip>
        </a:pPr>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let mut resolve = |_: &str| -> Option<String> { None };
        assert!(matches!(
            parse_bullet(Some(doc.root_element()), &theme, &mut resolve),
            Bullet::Inherit
        ));
    }

    /// §21.1.2.4.4 (buClr) — an explicit `<a:buClr>` sibling of `<a:buAutoNum>`
    /// colours the auto-number marker, exactly as it does a `<a:buChar>` bullet
    /// (§21.1.2.4.10 buClrTx is the default only when no buClr is present). The
    /// child order follows CT_TextParagraphProperties' xsd:sequence: buClr →
    /// buSzPct → buFont → buAutoNum. Regression: the buAutoNum branch used to drop
    /// the sibling buClr, forcing the marker onto the inherited first-run colour.
    #[test]
    fn test_parse_bullet_autonum_reads_buclr() {
        let xml = r#"<a:pPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <a:buClr><a:srgbClr val="C00000"/></a:buClr>
            <a:buSzPct val="100000"/>
            <a:buFont typeface="+mj-lt"/>
            <a:buAutoNum type="arabicPeriod"/>
        </a:pPr>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let mut resolve = |_: &str| -> Option<String> { None };
        let bullet = parse_bullet(Some(doc.root_element()), &theme, &mut resolve);
        let v = serde_json::to_value(&bullet).unwrap();
        assert_eq!(v["type"], "autoNum");
        assert_eq!(v["numType"], "arabicPeriod");
        // The buClr resolves to the srgbClr literal (uppercase hex, no '#').
        assert_eq!(v["color"], "C00000");
        assert_eq!(v["sizePct"], 100.0);
        assert_eq!(v["fontFamily"], "+mj-lt");
    }

    /// §21.1.2.4.10 (buClrTx) — with no explicit `<a:buClr>` the auto-number
    /// marker carries no own colour (`None`), so the renderer falls back to the
    /// default (the first run's colour). The `color` field must be absent/null,
    /// not silently defaulted to some literal.
    #[test]
    fn test_parse_bullet_autonum_without_buclr_has_no_color() {
        let xml = r#"<a:pPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <a:buAutoNum type="arabicPeriod"/>
        </a:pPr>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let mut resolve = |_: &str| -> Option<String> { None };
        let bullet = parse_bullet(Some(doc.root_element()), &theme, &mut resolve);
        let v = serde_json::to_value(&bullet).unwrap();
        assert_eq!(v["type"], "autoNum");
        assert_eq!(v["color"], serde_json::Value::Null);
    }

    /// §19.7.10 / §21.1.2.4.2 — a picture bullet declared on a master/list-style
    /// `<a:lvlNpPr>` is captured per level by `read_level_bullets`, so a slide
    /// paragraph at that level inherits the image marker.
    #[test]
    fn test_read_level_bullets_picks_up_bublip() {
        let xml = r#"<a:lstStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                                 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <a:lvl1pPr><a:buBlip><a:blip r:embed="rId9"/></a:buBlip></a:lvl1pPr>
            <a:lvl2pPr><a:buChar char="-"/></a:lvl2pPr>
        </a:lstStyle>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let mut resolve = |rid: &str| -> Option<String> {
            assert_eq!(rid, "rId9");
            Some("ppt/media/image7.png".to_owned())
        };
        let levels = read_level_bullets(doc.root_element(), &theme, &mut resolve);
        match levels[0].resolve() {
            Bullet::Blip { image_path, .. } => {
                assert_eq!(image_path, "ppt/media/image7.png")
            }
            other => panic!("expected lvl1 Bullet::Blip, got {other:?}"),
        }
        assert!(matches!(levels[1].resolve(), Bullet::Char { .. }));
        assert!(levels[2].is_inherit());
    }

    /// ECMA-376 §21.1.2.4.2 — a `<a:buBlip>` whose `r:embed` IS listed in the
    /// part's rels (so `resolve_path` succeeds) but whose target part is NOT in
    /// the package must NOT emit a `Bullet::Blip` carrying a dangling path. The
    /// resolver verifies part existence with `index_for_name`, so a missing part
    /// yields `None` and the level falls through to `Bullet::Inherit` (the empty
    /// `LevelBullets` slot), matching the variant's doc comment. Exercised
    /// end-to-end through `parse_master_level_bullets` (one of the now-`zip`-
    /// threaded entry points) against a real in-memory archive.
    #[test]
    fn master_bublip_listed_but_missing_part_inherits() {
        // bodyStyle lvl1 declares a picture bullet whose embed (rId7) IS in the
        // master rels, pointing at ppt/media/missing.png.
        let master = r#"<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                                     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                                     xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:cSld><p:spTree/></p:cSld>
          <p:txStyles>
            <p:bodyStyle>
              <a:lvl1pPr><a:buBlip><a:blip r:embed="rId7"/></a:buBlip><a:defRPr sz="2000"/></a:lvl1pPr>
            </p:bodyStyle>
          </p:txStyles>
        </p:sldMaster>"#;
        let theme = HashMap::new();
        let mut master_rels = HashMap::new();
        // rId7 resolves (resolve_path succeeds) to ppt/media/missing.png.
        master_rels.insert("rId7".to_owned(), "../media/missing.png".to_owned());

        // Archive deliberately LACKS ppt/media/missing.png (it holds an unrelated
        // part so it isn't empty). index_for_name(missing.png) → None.
        let bytes = zip_with_parts(&[("ppt/media/other.png", b"\x89PNG")]);
        let cursor = Cursor::new(bytes.clone());
        let mut zip = PptxZip::new(cursor).unwrap();

        let master_doc = roxmltree::Document::parse(master).unwrap();
        let master_root = master_doc.root_element();
        let m = parse_master_level_bullets(
            master_root,
            &theme,
            &master_rels,
            "ppt/slideMasters",
            &mut zip,
        );
        // The listed-but-missing part must not produce a Blip anywhere. With only
        // a buBlip (no char/auto/decoration) at lvl1 and the part absent, the
        // level resolves to Inherit and the bodyStyle contributes no usable
        // bullet, so the "body" key is never inserted (has_any_level_bullet is
        // false).
        if let Some(body) = m.get("body") {
            assert!(
                !matches!(body[0].resolve(), Bullet::Blip { .. }),
                "missing part must not yield Bullet::Blip; got {:?}",
                body[0]
            );
            assert!(
                body.iter()
                    .all(|b| !matches!(b.resolve(), Bullet::Blip { .. })),
                "no level may carry a dangling Bullet::Blip; got {body:?}"
            );
        }

        // Positive control: with the SAME rels but the part now PRESENT, the
        // bullet resolves to Bullet::Blip — proving the test distinguishes
        // presence from absence rather than always inheriting.
        let bytes_ok = zip_with_parts(&[("ppt/media/missing.png", b"\x89PNG")]);
        let cursor_ok = Cursor::new(bytes_ok.clone());
        let mut zip_ok = PptxZip::new(cursor_ok).unwrap();
        let m_ok = parse_master_level_bullets(
            master_root,
            &theme,
            &master_rels,
            "ppt/slideMasters",
            &mut zip_ok,
        );
        match m_ok.get("body").map(|b| b[0].resolve()) {
            Some(Bullet::Blip { image_path, .. }) => {
                assert_eq!(image_path, "ppt/media/missing.png");
            }
            other => panic!("expected Bullet::Blip when part is present, got {other:?}"),
        }
    }

    /// ECMA-376 §21.1.2.3.9; ST_TextUnderlineType §20.1.10.82 —
    /// underline_style carries non-default underline
    /// values (dbl, dotted, wavy, …) verbatim. The plain bool stays true for
    /// any non-"none" value; "sng" and absent both leave underline_style None
    /// because the renderer's default is already a single line.
    #[test]
    fn test_parse_run_underline_style_passthrough() {
        let theme = HashMap::new();
        let rels = HashMap::new();
        let cases: &[(&str, bool, Option<&str>)] = &[
            ("none", false, None),
            ("sng", true, None),
            ("dbl", true, Some("dbl")),
            ("heavy", true, Some("heavy")),
            ("dotted", true, Some("dotted")),
            ("wavy", true, Some("wavy")),
            ("dashLong", true, Some("dashLong")),
        ];
        for (val, expected_bool, expected_style) in cases {
            let xml = format!(
                r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr u="{val}"/><t>x</t></r>"#
            );
            let doc = roxmltree::Document::parse(&xml).unwrap();
            let r = parse_run(doc.root_element(), None, &theme, &rels).unwrap();
            assert_eq!(r.underline, *expected_bool, "u={val}");
            assert_eq!(r.underline_style.as_deref(), *expected_style, "u={val}");
        }
    }

    /// ECMA-376 §21.1.2.3.12 — rPr > uFill > solidFill yields a per-run
    /// underline colour distinct from the text colour. uFillTx (or absent)
    /// leaves underline_color as None so the renderer falls back to text.
    #[test]
    fn test_parse_run_underline_color() {
        let theme = HashMap::new();
        let rels = HashMap::new();

        let with_ufill = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr u="sng"><uFill><solidFill><srgbClr val="FF0000"/></solidFill></uFill></rPr><t>x</t></r>"#;
        let doc = roxmltree::Document::parse(with_ufill).unwrap();
        let r = parse_run(doc.root_element(), None, &theme, &rels).unwrap();
        assert_eq!(
            r.underline_color
                .as_deref()
                .map(str::to_uppercase)
                .as_deref(),
            Some("FF0000")
        );

        let with_ufilltx = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr u="sng"><uFillTx/></rPr><t>x</t></r>"#;
        let doc = roxmltree::Document::parse(with_ufilltx).unwrap();
        let r = parse_run(doc.root_element(), None, &theme, &rels).unwrap();
        assert!(r.underline_color.is_none());
    }

    /// ECMA-376 §21.1.2.3.4 — rPr > highlight is a CT_Color (the marker /
    /// text-highlight colour). Unlike WordprocessingML's CT_Highlight (a fixed
    /// 16-name enum), the DrawingML highlight is any colour, so it must resolve
    /// through the same colour pipeline as solidFill: srgbClr literal,
    /// schemeClr via the theme/clrMap, plus alpha transforms (8-char hex).
    #[test]
    fn test_parse_run_highlight_srgb() {
        let theme = HashMap::new();
        let rels = HashMap::new();
        let xml = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr><highlight><srgbClr val="FFFF00"/></highlight></rPr><t>x</t></r>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let r = parse_run(doc.root_element(), None, &theme, &rels).unwrap();
        assert_eq!(
            r.highlight.as_deref().map(str::to_uppercase).as_deref(),
            Some("FFFF00")
        );
    }

    /// schemeClr highlight resolves through the theme map (same path as
    /// solidFill scheme colours), proving we did not hard-code a name table.
    #[test]
    fn test_parse_run_highlight_scheme_resolves_theme() {
        let rels = HashMap::new();
        let mut theme = HashMap::new();
        theme.insert("accent1".to_owned(), "E46970".to_owned());
        let xml = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr><highlight><schemeClr val="accent1"/></highlight></rPr><t>x</t></r>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let r = parse_run(doc.root_element(), None, &theme, &rels).unwrap();
        assert_eq!(
            r.highlight.as_deref().map(str::to_uppercase).as_deref(),
            Some("E46970")
        );
    }

    /// An alpha transform on the highlight colour yields 8-char RRGGBBAA, the
    /// same encoding the shared colour helper emits for translucent fills.
    #[test]
    fn test_parse_run_highlight_alpha_is_8char() {
        let theme = HashMap::new();
        let rels = HashMap::new();
        let xml = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr><highlight><srgbClr val="00FF00"><alpha val="50000"/></srgbClr></highlight></rPr><t>x</t></r>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let r = parse_run(doc.root_element(), None, &theme, &rels).unwrap();
        let hl = r
            .highlight
            .expect("highlight should resolve")
            .to_uppercase();
        assert_eq!(hl.len(), 8, "alpha < 1 → RRGGBBAA, got {hl}");
        assert!(hl.starts_with("00FF00"), "rgb preserved, got {hl}");
    }

    /// No highlight element → field stays None (omitted from JSON).
    #[test]
    fn test_parse_run_without_highlight_is_none() {
        let theme = HashMap::new();
        let rels = HashMap::new();
        let xml = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr/><t>x</t></r>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let r = parse_run(doc.root_element(), None, &theme, &rels).unwrap();
        assert!(r.highlight.is_none());
    }

    /// ECMA-376 §21.1.2.3.3 — rPr > ea sets a separate East Asian font.
    /// Resolves through the theme map: "+mn-ea" should expand to whatever
    /// the theme registered, while a literal name is preserved.
    #[test]
    fn test_parse_run_ea_typeface() {
        let rels = HashMap::new();
        let mut theme = HashMap::new();
        theme.insert("+mn-ea".to_owned(), "MS Mincho".to_owned());

        // Theme reference resolves through the map.
        let xml = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr><ea typeface="+mn-ea"/></rPr><t>あ</t></r>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let r = parse_run(doc.root_element(), None, &theme, &rels).unwrap();
        assert_eq!(r.font_family_ea.as_deref(), Some("MS Mincho"));

        // Literal name passes through unchanged.
        let xml = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr><ea typeface="Yu Gothic"/></rPr><t>あ</t></r>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let r = parse_run(doc.root_element(), None, &theme, &rels).unwrap();
        assert_eq!(r.font_family_ea.as_deref(), Some("Yu Gothic"));

        // Empty typeface is filtered out.
        let xml = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr><ea typeface=""/></rPr><t>あ</t></r>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let r = parse_run(doc.root_element(), None, &theme, &rels).unwrap();
        assert!(r.font_family_ea.is_none());
    }

    /// ECMA-376 §20.1.8.17 — glow has a single rad attribute and a colour
    /// child. parse_glow should preserve the radius and resolve alphaModFix.
    #[test]
    fn test_parse_glow_extracts_radius_and_color() {
        let xml = r#"<effectLst xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
            <glow rad="38100">
                <srgbClr val="FF0000"><alphaModFix amt="80000"/></srgbClr>
            </glow>
        </effectLst>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let g = parse_glow(doc.root_element(), &theme).expect("glow should resolve");
        assert_eq!(g.radius, 38_100);
        assert_eq!(g.color.to_uppercase(), "FF0000");
        assert!((g.alpha - 0.8).abs() < 0.01);
    }

    /// ECMA-376 §20.1.8.31 — softEdge has a single `rad` attribute in EMU.
    #[test]
    fn test_parse_soft_edge_extracts_radius() {
        let xml = r#"<effectLst xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
            <softEdge rad="63500"/>
        </effectLst>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let s = parse_soft_edge(doc.root_element()).expect("softEdge should resolve");
        assert_eq!(s.radius, 63_500);
    }

    /// ECMA-376 §20.1.8.27 — reflection: blur, dist, dir, stA/stPos/endA/endPos
    /// (1000ths of percent), sx/sy (1000ths of percent, sy negative for mirror).
    #[test]
    fn test_parse_reflection_attributes() {
        let xml = r#"<effectLst xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
            <reflection blurRad="6350" stA="50000" endA="0" endPos="35000" dist="50800" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/>
        </effectLst>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let r = parse_reflection(doc.root_element()).expect("reflection should resolve");
        assert_eq!(r.blur, 6_350);
        assert_eq!(r.dist, 50_800);
        assert!((r.dir - 90.0).abs() < 0.001);
        assert!((r.st_a - 0.5).abs() < 0.01);
        assert!((r.end_a - 0.0).abs() < 0.01);
        assert!((r.end_pos - 0.35).abs() < 0.01);
        assert!((r.sy + 1.0).abs() < 0.01);
        // sx defaults to 1.0 when not specified
        assert!((r.sx - 1.0).abs() < 0.01);
    }

    /// §19.3.1.37 — a p:pic's spPr is CT_ShapeProperties, so every effectLst
    /// child (§20.1.8.16) applies to images. parse_effect_lst is the shared
    /// reader both p:sp and p:pic use; exercise it with the reflection-bearing
    /// effectLst lifted from sample-11's `図 3` picture.
    #[test]
    fn test_pic_effect_lst_resolves_all_effects() {
        let xml = r#"<spPr xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
            <effectLst>
                <outerShdw blurRad="50800" dist="38100" dir="2700000"
                            sx="50000" sy="150000" kx="1200000" ky="-600000"
                            algn="tr" rotWithShape="0">
                    <srgbClr val="000000"><alpha val="40000"/></srgbClr>
                </outerShdw>
                <innerShdw blurRad="63500" dist="50800" dir="5400000"><srgbClr val="111111"/></innerShdw>
                <glow rad="63500"><srgbClr val="FFCC00"/></glow>
                <softEdge rad="25400"/>
                <reflection blurRad="12700" stA="38000" endPos="28000" dist="5000" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/>
            </effectLst>
        </spPr>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let sp_pr = doc.root_element();
        let eff = parse_effect_lst(child(sp_pr, "effectLst"), &theme);

        let shadow = eff.shadow.expect("outerShdw should resolve");
        assert_eq!(shadow.blur, 50_800);
        assert_eq!(shadow.dist, 38_100);
        assert!((shadow.alpha - 0.4).abs() < 0.01);
        assert_eq!(shadow.sx, Some(0.5));
        assert_eq!(shadow.sy, Some(1.5));
        assert_eq!(shadow.kx, Some(20.0));
        assert_eq!(shadow.ky, Some(-10.0));
        assert_eq!(shadow.algn.as_deref(), Some("tr"));
        assert_eq!(shadow.rot_with_shape, Some(false));

        let inner = eff.inner_shadow.expect("innerShdw should resolve");
        assert_eq!(inner.blur, 63_500);

        let glow = eff.glow.expect("glow should resolve");
        assert_eq!(glow.radius, 63_500);
        assert_eq!(glow.color, "FFCC00");

        let soft = eff.soft_edge.expect("softEdge should resolve");
        assert_eq!(soft.radius, 25_400);

        let r = eff.reflection.expect("reflection should resolve");
        assert_eq!(r.blur, 12_700);
        assert_eq!(r.dist, 5_000);
        assert!((r.dir - 90.0).abs() < 0.001);
        assert!((r.st_a - 0.38).abs() < 0.01);
        assert!((r.end_pos - 0.28).abs() < 0.01);
        assert!((r.sy + 1.0).abs() < 0.01);
    }

    /// A spPr with no effectLst yields an all-None EffectLst (the common case).
    #[test]
    fn test_pic_effect_lst_empty_when_absent() {
        let xml = r#"<spPr xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let eff = parse_effect_lst(child(doc.root_element(), "effectLst"), &theme);
        assert!(eff.shadow.is_none());
        assert!(eff.inner_shadow.is_none());
        assert!(eff.glow.is_none());
        assert!(eff.soft_edge.is_none());
        assert!(eff.reflection.is_none());
    }

    /// §20.1.9.18 — `<a:prstGeom prst="roundRect">` on a picture's spPr clips
    /// the bitmap to a rounded rect. An explicit `adj` guide is carried through;
    /// the preset default is supplied by the shared engine, not the parser.
    #[test]
    fn test_pic_prst_geom_round_rect_explicit_adj() {
        let xml = r#"<spPr xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
            <prstGeom prst="roundRect"><avLst><gd name="adj" fmla="val 8594"/></avLst></prstGeom>
        </spPr>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        assert_eq!(
            parse_pic_prst_geom(doc.root_element()),
            (Some("roundRect".to_owned()), Some(vec![8_594]))
        );
    }

    /// When avLst omits the guide, the parser carries the name with no adjust;
    /// the preset's own default (roundRect adj = 16667) is filled in downstream
    /// by the TS preset-geometry engine, keeping defaults in one place.
    #[test]
    fn test_pic_prst_geom_round_rect_default_adj() {
        let xml = r#"<spPr xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
            <prstGeom prst="roundRect"><avLst/></prstGeom>
        </spPr>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        assert_eq!(
            parse_pic_prst_geom(doc.root_element()),
            (Some("roundRect".to_owned()), None)
        );
    }

    /// §20.1.9.18 generalised — a non-roundRect preset (ellipse, empty avLst) is
    /// now carried generically so the picture clips to that silhouette.
    #[test]
    fn test_pic_prst_geom_ellipse() {
        let xml = r#"<spPr xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
            <prstGeom prst="ellipse"><avLst/></prstGeom>
        </spPr>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        assert_eq!(
            parse_pic_prst_geom(doc.root_element()),
            (Some("ellipse".to_owned()), None)
        );
    }

    /// Multiple adjust guides are captured in declaration order.
    #[test]
    fn test_pic_prst_geom_multi_adj() {
        let xml = r#"<spPr xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
            <prstGeom prst="round2SameRect"><avLst>
                <gd name="adj1" fmla="val 16667"/><gd name="adj2" fmla="val 0"/>
            </avLst></prstGeom>
        </spPr>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        assert_eq!(
            parse_pic_prst_geom(doc.root_element()),
            (Some("round2SameRect".to_owned()), Some(vec![16_667, 0]))
        );
    }

    /// A plain rect (or no prstGeom at all) means no clip path.
    #[test]
    fn test_pic_prst_geom_absent() {
        let rect = r#"<spPr xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
            <prstGeom prst="rect"><avLst/></prstGeom>
        </spPr>"#;
        let doc = roxmltree::Document::parse(rect).unwrap();
        assert_eq!(parse_pic_prst_geom(doc.root_element()), (None, None));

        let bare = r#"<spPr xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#;
        let doc = roxmltree::Document::parse(bare).unwrap();
        assert_eq!(parse_pic_prst_geom(doc.root_element()), (None, None));
    }

    /// ECMA-376 §20.1.8.42 — `<a:ln cmpd="dbl"/>` should round-trip.
    /// `cmpd="sng"` is the spec default and stays absent in the model.
    #[test]
    fn test_parse_stroke_cmpd() {
        let theme = HashMap::new();
        let dbl = r#"<ln xmlns="http://schemas.openxmlformats.org/drawingml/2006/main" w="38100" cmpd="dbl"><solidFill><srgbClr val="000000"/></solidFill></ln>"#;
        let doc = roxmltree::Document::parse(dbl).unwrap();
        let s = parse_stroke(doc.root_element(), &theme).expect("stroke should parse");
        assert_eq!(s.cmpd.as_deref(), Some("dbl"));

        let sng = r#"<ln xmlns="http://schemas.openxmlformats.org/drawingml/2006/main" w="38100" cmpd="sng"><solidFill><srgbClr val="000000"/></solidFill></ln>"#;
        let doc = roxmltree::Document::parse(sng).unwrap();
        let s = parse_stroke(doc.root_element(), &theme).expect("stroke should parse");
        assert!(s.cmpd.is_none());
    }

    /// ECMA-376 §20.1.8.38 CT_LineProperties permits the same fill choices as
    /// shapes, including gradFill. A gradient line must remain a visible
    /// stroke instead of being discarded for lacking solidFill.
    #[test]
    fn test_parse_gradient_stroke() {
        let theme = HashMap::new();
        let xml = r#"
          <ln xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"
              w="76200" cap="rnd">
            <gradFill rotWithShape="1">
              <gsLst>
                <gs pos="19000"><srgbClr val="112233"><alpha val="0"/></srgbClr></gs>
                <gs pos="100000"><srgbClr val="AABBCC"/></gs>
              </gsLst>
              <lin ang="5400000" scaled="1"/>
            </gradFill>
            <headEnd type="arrow" w="lg" len="sm"/>
          </ln>
        "#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let stroke =
            parse_stroke(doc.root_element(), &theme).expect("gradient stroke should parse");

        assert_eq!(stroke.width, 76200);
        assert_eq!(stroke.color, "AABBCC");
        assert_eq!(
            stroke.head_end.as_ref().map(|end| end.kind.as_str()),
            Some("arrow")
        );
        match stroke.fill {
            Some(Fill::Gradient {
                stops,
                angle,
                grad_type,
                ..
            }) => {
                assert_eq!(angle, 90.0);
                assert_eq!(grad_type, "linear");
                assert_eq!(stops.len(), 2);
                assert_eq!(stops[0].position, 0.19);
                assert_eq!(stops[0].color, "11223300");
                assert_eq!(stops[1].color, "AABBCC");
            }
            other => panic!("expected gradient stroke fill, got {other:?}"),
        }
    }

    #[test]
    fn master_body_style_per_level_font_sizes() {
        // ECMA-376 §21.1.2.4: each list level has its own defRPr sz. A 2nd-level
        // bullet must inherit lvl3pPr's smaller size, not lvl1pPr's.
        let master = r#"<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld><p:spTree/></p:cSld>
          <p:txStyles>
            <p:bodyStyle>
              <a:lvl1pPr><a:defRPr sz="2800"/></a:lvl1pPr>
              <a:lvl2pPr><a:defRPr sz="2400"/></a:lvl2pPr>
              <a:lvl3pPr><a:defRPr sz="2000"/></a:lvl3pPr>
            </p:bodyStyle>
            <p:titleStyle><a:lvl1pPr><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle>
          </p:txStyles>
        </p:sldMaster>"#;
        let master_doc = roxmltree::Document::parse(master).unwrap();
        let m = parse_master_level_font_sizes(master_doc.root_element());
        let body = m.get("body").expect("body level sizes");
        assert_eq!(body[0], Some(28.0)); // lvl1 → level 0
        assert_eq!(body[1], Some(24.0)); // lvl2 → level 1
        assert_eq!(body[2], Some(20.0)); // lvl3 → level 2
        assert_eq!(body[3], None); // unspecified
                                   // body style also keys the empty placeholder type and "obj".
        assert_eq!(m.get("").unwrap()[2], Some(20.0));
        // title style is captured separately.
        assert_eq!(m.get("title").unwrap()[0], Some(44.0));
    }

    /// ECMA-376 §19.7.10 / §21.1.2.4 — a slide body paragraph with no explicit
    /// `<a:buChar>` inherits the master `bodyStyle` bullet. `parse_master_level_bullets`
    /// must surface that `•` (keyed by body/""/obj), so the renderer can draw it.
    /// Regression: sample-9 slides 4/7/12 bullet lists rendered with no markers.
    #[test]
    fn master_body_style_bullets_inherited_by_level() {
        let master = r#"<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld><p:spTree/></p:cSld>
          <p:txStyles>
            <p:bodyStyle>
              <a:lvl1pPr><a:buFont typeface="Arial"/><a:buChar char="•"/><a:defRPr sz="2000"/></a:lvl1pPr>
              <a:lvl2pPr><a:buFont typeface="Arial"/><a:buChar char="–"/><a:defRPr sz="1800"/></a:lvl2pPr>
            </p:bodyStyle>
            <p:titleStyle><a:lvl1pPr><a:buNone/><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle>
          </p:txStyles>
        </p:sldMaster>"#;
        let theme = HashMap::new();
        let master_rels = HashMap::new();
        // Char bullets only — no media part lookups, so an empty archive suffices.
        let bytes = empty_zip_bytes();
        let cursor = Cursor::new(bytes.clone());
        let mut zip = PptxZip::new(cursor).unwrap();
        let master_doc = roxmltree::Document::parse(master).unwrap();
        let m = parse_master_level_bullets(
            master_doc.root_element(),
            &theme,
            &master_rels,
            "ppt/slideMasters",
            &mut zip,
        );
        let body = m.get("body").expect("body bullets");
        match body[0].resolve() {
            Bullet::Char { ch, .. } => assert_eq!(ch, "•", "lvl1 bullet char"),
            other => panic!("expected lvl1 char bullet, got {other:?}"),
        }
        match body[1].resolve() {
            Bullet::Char { ch, .. } => assert_eq!(ch, "–", "lvl2 bullet char"),
            other => panic!("expected lvl2 char bullet, got {other:?}"),
        }
        assert!(body[2].is_inherit(), "lvl3 unspecified");
        // body style also keys the empty placeholder type and "obj".
        assert!(matches!(
            m.get("").map(|b| b[0].resolve()),
            Some(Bullet::Char { .. })
        ));
        assert!(matches!(
            m.get("obj").map(|b| b[0].resolve()),
            Some(Bullet::Char { .. })
        ));
        // titleStyle's explicit buNone is captured (so titles don't inherit a bullet).
        assert!(matches!(
            m.get("title").map(|b| b[0].resolve()),
            Some(Bullet::None)
        ));
    }

    #[test]
    fn merge_level_sizes_prefers_primary_per_edge() {
        let primary: LevelFontSizes = {
            let mut a = [None; 9];
            a[1] = Some(28.0);
            a
        };
        let fallback: LevelFontSizes = {
            let mut a = [None; 9];
            a[0] = Some(32.0);
            a[1] = Some(24.0);
            a[2] = Some(20.0);
            a
        };
        let merged = merge_level_sizes(&primary, &fallback);
        assert_eq!(merged[0], Some(32.0)); // only fallback
        assert_eq!(merged[1], Some(28.0)); // primary wins
        assert_eq!(merged[2], Some(20.0)); // only fallback
    }

    /// ECMA-376 §21.1.2.4.13 — `<a:lvlNpPr>` is a `CT_TextParagraphProperties`,
    /// so `marL`/`marR`/`indent` are attributes ON the level element itself.
    /// `parse_master_level_indents` must surface the authored per-level values
    /// (keyed by body/""/obj for bodyStyle) and merge per-axis: a level that
    /// sets only `marL` leaves `marR`/`indent` None so a lower tier supplies them.
    #[test]
    fn master_body_style_per_level_indents() {
        let master = r#"<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld><p:spTree/></p:cSld>
          <p:txStyles>
            <p:bodyStyle>
              <a:lvl1pPr marL="1000000" indent="-500000"><a:defRPr sz="2800"/></a:lvl1pPr>
              <a:lvl2pPr marL="1500000"><a:defRPr sz="2400"/></a:lvl2pPr>
            </p:bodyStyle>
            <p:titleStyle><a:lvl1pPr marR="123456"><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle>
          </p:txStyles>
        </p:sldMaster>"#;
        let master_doc = roxmltree::Document::parse(master).unwrap();
        let m = parse_master_level_indents(master_doc.root_element());
        let body = m.get("body").expect("body level indents");
        assert_eq!(body[0].mar_l, Some(1_000_000));
        assert_eq!(body[0].indent, Some(-500_000));
        assert_eq!(body[0].mar_r, None); // unspecified axis stays None
        assert_eq!(body[1].mar_l, Some(1_500_000));
        assert_eq!(body[1].indent, None); // lvl2 omits indent → None
        assert_eq!(body[2].mar_l, None); // unspecified level
                                         // body style also keys the empty placeholder type and "obj".
        assert_eq!(m.get("").unwrap()[0].mar_l, Some(1_000_000));
        assert_eq!(m.get("obj").unwrap()[1].mar_l, Some(1_500_000));
        // title style is captured separately.
        assert_eq!(m.get("title").unwrap()[0].mar_r, Some(123_456));
    }

    /// Per-axis, per-level merge: `primary[lvl].x` wins, else `fallback[lvl].x`.
    #[test]
    fn merge_level_indents_prefers_primary_per_axis() {
        let primary: LevelIndents = {
            let mut a: LevelIndents = Default::default();
            a[0].mar_l = Some(100);
            a[1].indent = Some(-200);
            a
        };
        let fallback: LevelIndents = {
            let mut a: LevelIndents = Default::default();
            a[0].mar_l = Some(999); // loses to primary
            a[0].mar_r = Some(50); // only fallback
            a[1].indent = Some(-999); // loses to primary
            a[1].mar_l = Some(300); // only fallback
            a
        };
        let merged = merge_level_indents(&primary, &fallback);
        assert_eq!(merged[0].mar_l, Some(100)); // primary wins
        assert_eq!(merged[0].mar_r, Some(50)); // only fallback
        assert_eq!(merged[1].indent, Some(-200)); // primary wins
        assert_eq!(merged[1].mar_l, Some(300)); // only fallback
    }

    /// ECMA-376 §21.1.2.4.13 cascade end-to-end: a paragraph whose body lstStyle
    /// authors per-level `marL`/`indent` and whose own `<a:pPr>` omits them must
    /// resolve to the AUTHORED level values (not the hardcoded implicit
    /// `(lvl+1)*342900` / `-342900`). A direct `<a:pPr marL=...>` still wins.
    /// With nothing authored, the implicit default applies (regression guard).
    #[test]
    fn pptx_level_indent_inherited_from_lststyle() {
        let theme = HashMap::new();
        let rels = HashMap::new();
        let bytes = empty_zip_bytes();
        let cursor = Cursor::new(bytes.clone());
        let mut zip = PptxZip::new(cursor).unwrap();
        // `lst_style` sets the body lstStyle (the inherited per-level cascade);
        // `p_pr` is the paragraph's own pPr. Returns the single paragraph.
        let mut parse_para = |lst_style: &str, p_pr: &str| -> Paragraph {
            let xml = format!(
                r#"<txBody xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">{lst_style}<p>{p_pr}<r><t>x</t></r></p></txBody>"#
            );
            let doc = roxmltree::Document::parse(&xml).unwrap();
            let mut tb = parse_text_body(
                doc.root_element(),
                &theme,
                &rels,
                "ppt/slides",
                None,
                None,
                [None; 9],
                Default::default(), // inherited_level_indents
                &empty_level_bullets(),
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                ShapeKind::Sp,
                &mut zip,
            );
            tb.paragraphs.remove(0)
        };

        // (1) Authored level marL/indent inherited when the paragraph omits them.
        let lst = r#"<lstStyle><lvl1pPr marL="1000000" indent="-500000"/></lstStyle>"#;
        let inherited = parse_para(lst, "<pPr/>");
        assert_eq!(
            inherited.mar_l, 1_000_000,
            "marL should inherit the authored lvl1pPr value, not the implicit default"
        );
        assert_eq!(
            inherited.indent, -500_000,
            "indent should inherit the authored lvl1pPr value, not the implicit default"
        );

        // (2) A direct pPr marL overrides the inherited level value.
        let overridden = parse_para(lst, r#"<pPr marL="2000000"/>"#);
        assert_eq!(
            overridden.mar_l, 2_000_000,
            "direct pPr marL must win over the inherited level value"
        );
        // indent (not set directly) still inherits the level value.
        assert_eq!(
            overridden.indent, -500_000,
            "indent should still inherit when only marL is set directly"
        );

        // (3) Regression: nothing authored → hardcoded implicit default for a
        // plain (non-bullet) paragraph at lvl 0: marL=0, marR=0, indent=0.
        let implicit = parse_para("", "<pPr/>");
        assert_eq!(implicit.mar_l, 0, "implicit marL default for plain lvl0");
        assert_eq!(implicit.mar_r, 0, "implicit marR default");
        assert_eq!(implicit.indent, 0, "implicit indent default for plain lvl0");
    }

    /// ECMA-376 §21.1.2.4.13 cross-tier, per-axis inheritance: when a layout
    /// placeholder's own `lstStyle` and the master `txStyles` each author a
    /// DIFFERENT axis of the same level, `parse_layout_placeholders` must merge them
    /// per axis (layout wins per axis, master fills the rest) and expose the result
    /// through `lookup_level_indents`. This exercises the actual layout↔master
    /// wiring, not just `merge_level_indents` in isolation.
    #[test]
    fn layout_over_master_level_indents_merge_per_axis() {
        let bytes = empty_zip_bytes();
        let cursor = Cursor::new(bytes.clone());
        let mut zip = PptxZip::new(cursor).unwrap();

        // Master authors only marL on the body level; layout authors only indent.
        let mut master_indents: HashMap<String, LevelIndents> = HashMap::new();
        let mut body: LevelIndents = Default::default();
        body[0].mar_l = Some(1_000_000);
        master_indents.insert("body".to_string(), body);

        let layout = r#"<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld><p:spTree>
            <p:sp>
              <p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
              <p:spPr/>
              <p:txBody><a:lstStyle><a:lvl1pPr indent="-111111"/></a:lstStyle><a:p/></p:txBody>
            </p:sp>
          </p:spTree></p:cSld>
        </p:sldLayout>"#;

        let layout_doc = roxmltree::Document::parse(layout).unwrap();
        let lph = parse_layout_placeholders(
            layout_doc.root_element(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &master_indents,
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            "",
            &HashMap::new(),
            &mut zip,
        );

        let li = lph.lookup_level_indents("body", None);
        assert_eq!(
            li[0].indent,
            Some(-111_111),
            "indent must come from the LAYOUT lstStyle (primary tier)"
        );
        assert_eq!(
            li[0].mar_l,
            Some(1_000_000),
            "marL must fall back to the MASTER per axis (layout left it unset)"
        );
    }

    /// D4 guard: `parse_layout` resolves the layout placeholder's color-bearing
    /// fields, its `<p:bg>`, and its `showMasterSp` against the `theme` argument.
    /// The color/background must FLIP when the caller passes an override-adjusted
    /// theme — this is what the `parse_presentation` clrMapOvr branch relies on
    /// (a cached `ParsedLayout` is only sound because a no-override slide passes
    /// the same master-baked theme every time). Also asserts a theme-independent
    /// field (transform) is stable regardless of theme.
    #[test]
    fn parse_layout_resolves_color_and_bg_against_theme() {
        let layout = r#"<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          showMasterSp="0">
          <p:cSld>
            <p:bg><p:bgPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></p:bgPr></p:bg>
            <p:spTree>
              <p:sp>
                <p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
                <p:spPr><a:xfrm><a:off x="123456" y="0"/><a:ext cx="10" cy="10"/></a:xfrm></p:spPr>
                <p:txBody><a:lstStyle><a:lvl1pPr><a:defRPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle><a:p/></p:txBody>
              </p:sp>
            </p:spTree>
          </p:cSld>
        </p:sldLayout>"#;

        // Typed-empty master inheritance maps (no master fallbacks in this test).
        let m_f64: HashMap<String, f64> = HashMap::new();
        let m_lfs: HashMap<String, LevelFontSizes> = HashMap::new();
        let m_li: HashMap<String, LevelIndents> = HashMap::new();
        let m_lb: HashMap<String, LevelBullets> = HashMap::new();
        let m_str: HashMap<String, String> = HashMap::new();
        let m_tf: HashMap<String, Transform> = HashMap::new();
        let m_bool: HashMap<String, bool> = HashMap::new();
        let m_i64: HashMap<String, i64> = HashMap::new();
        let empty_rels: HashMap<String, String> = HashMap::new();
        let build = |accent1_hex: &str| -> ParsedLayout {
            let mut theme: HashMap<String, String> = HashMap::new();
            theme.insert("accent1".to_string(), accent1_hex.to_string());
            let bytes = empty_zip_bytes();
            let cursor = Cursor::new(bytes.clone());
            let mut zip = PptxZip::new(cursor).unwrap();
            parse_layout(
                layout,
                &m_f64,
                &m_str,
                &m_lfs,
                &m_li,
                &m_lb,
                &m_str,
                &m_tf,
                &m_str,
                &m_bool,
                &m_i64,
                &m_i64,
                &m_f64,
                &theme,
                "ppt/slideLayouts",
                &empty_rels,
                &mut zip,
            )
        };

        let bg_solid_hex = |pl: &ParsedLayout| -> Option<String> {
            match pl.background.as_ref()? {
                Fill::Solid { color } => Some(color.clone()),
                _ => None,
            }
        };

        let base = build("FF0000");
        assert!(!base.show_master_sp, "layout showMasterSp=0 is read");
        assert_eq!(
            bg_solid_hex(&base).as_deref(),
            Some("FF0000"),
            "layout bg schemeClr resolves against the passed theme"
        );
        assert_eq!(
            base.placeholders
                .by_type_color
                .get("body")
                .map(String::as_str),
            Some("FF0000"),
            "layout placeholder defRPr color resolves against the passed theme"
        );
        // Theme-independent geometry is stable.
        assert_eq!(
            base.placeholders.by_type.get("body").map(|t| t.x),
            Some(123456),
            "placeholder transform is theme-independent"
        );

        // Same layout XML, DIFFERENT theme (simulating an override remap): the
        // color-bearing fields must flip; geometry must not.
        let flipped = build("00FF00");
        assert_eq!(
            bg_solid_hex(&flipped).as_deref(),
            Some("00FF00"),
            "override theme must flip the layout bg color"
        );
        assert_eq!(
            flipped
                .placeholders
                .by_type_color
                .get("body")
                .map(String::as_str),
            Some("00FF00"),
            "override theme must flip the layout placeholder color"
        );
        assert_eq!(
            flipped.placeholders.by_type.get("body").map(|t| t.x),
            Some(123456)
        );
    }

    /// Build a deck with `n_slides` slides that ALL reference the same single
    /// layout + single master (no clrMapOvr, no master/layout decorative shapes).
    /// Used to assert the D4 slide-master/layout parse count stays bounded.
    fn build_shared_layout_deck(n_slides: usize) -> Vec<u8> {
        let sld_ids: String = (0..n_slides)
            .map(|i| format!("<p:sldId id=\"{}\" r:id=\"rIdSlide{}\"/>", 256 + i, i))
            .collect();
        let presentation_xml = format!(
            r#"<p:presentation xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMaster"/></p:sldMasterIdLst>
  <p:sldIdLst>{sld_ids}</p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>"#
        );
        let mut pres_rel_entries = String::from(
            r#"<Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>"#,
        );
        for i in 0..n_slides {
            pres_rel_entries.push_str(&format!(
                "\n  <Relationship Id=\"rIdSlide{i}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide{i}.xml\"/>"
            ));
        }
        let pres_rels =
            format!("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">{pres_rel_entries}</Relationships>");
        let theme_xml = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T"><a:themeElements><a:clrScheme name="C"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="111111"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2><a:accent1><a:srgbClr val="FF0000"/></a:accent1><a:accent2><a:srgbClr val="00FF00"/></a:accent2><a:accent3><a:srgbClr val="0000FF"/></a:accent3><a:accent4><a:srgbClr val="FFFF00"/></a:accent4><a:accent5><a:srgbClr val="FF00FF"/></a:accent5><a:accent6><a:srgbClr val="00FFFF"/></a:accent6><a:hlink><a:srgbClr val="0000EE"/></a:hlink><a:folHlink><a:srgbClr val="551A8B"/></a:folHlink></a:clrScheme><a:fontScheme name="F"><a:majorFont><a:latin typeface="Arial"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme><a:fmtScheme name="S"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme></a:themeElements></a:theme>"#;
        // Master + layout carry ONLY placeholder shapes (no decorative), so the
        // master-decorative pre-extraction stores an empty vec and the layout
        // decorative walk finds nothing — the parse count reflects the pagination
        // path alone.
        let master_xml = r#"<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rIdLayout"/></p:sldLayoutIdLst></p:sldMaster>"#;
        let master_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>"#;
        let layout_xml = r#"<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>"#;
        let layout_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>"#;
        let slide_xml = r#"<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sld>"#;
        let slide_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>"#;

        let mut parts: Vec<(String, Vec<u8>)> = vec![
            ("ppt/presentation.xml".into(), presentation_xml.into_bytes()),
            (
                "ppt/_rels/presentation.xml.rels".into(),
                pres_rels.into_bytes(),
            ),
            ("ppt/theme/theme1.xml".into(), theme_xml.as_bytes().to_vec()),
            (
                "ppt/slideMasters/slideMaster1.xml".into(),
                master_xml.as_bytes().to_vec(),
            ),
            (
                "ppt/slideMasters/_rels/slideMaster1.xml.rels".into(),
                master_rels.as_bytes().to_vec(),
            ),
            (
                "ppt/slideLayouts/slideLayout1.xml".into(),
                layout_xml.as_bytes().to_vec(),
            ),
            (
                "ppt/slideLayouts/_rels/slideLayout1.xml.rels".into(),
                layout_rels.as_bytes().to_vec(),
            ),
        ];
        for i in 0..n_slides {
            parts.push((
                format!("ppt/slides/slide{i}.xml"),
                slide_xml.as_bytes().to_vec(),
            ));
            parts.push((
                format!("ppt/slides/_rels/slide{i}.xml.rels"),
                slide_rels.as_bytes().to_vec(),
            ));
        }
        let borrowed: Vec<(&str, &[u8])> = parts
            .iter()
            .map(|(p, b)| (p.as_str(), b.as_slice()))
            .collect();
        zip_with_parts(&borrowed)
    }

    /// D4 regression guard: the slide-master + layout `Document::parse` count on
    /// the pagination path must be BOUNDED — not `k · slides`. With every slide
    /// sharing one layout + one master (no clrMapOvr, no decorations), the master
    /// is built once (1 parse) and the layout is parsed once for the cache, so
    /// the total is `2 + 2·slides` (per slide: its own XML + the layout decorative
    /// walk). Crucially the master/layout parse count does NOT grow by the 12+
    /// (master) or 4 (layout) per-slide factor this change removed. Asserting the
    /// slope across two slide counts pins the optimization: master build and the
    /// layout cache each fire exactly once regardless of N.
    #[test]
    fn parse_count_scales_with_distinct_parts() {
        let count_for = |n: usize| -> usize {
            let data = build_shared_layout_deck(n);
            LAYOUT_MASTER_PARSE_COUNT.with(|c| c.set(0));
            let pres = parse_presentation_from_bytes(&data).expect("parse");
            assert_eq!(pres.slides.len(), n);
            LAYOUT_MASTER_PARSE_COUNT.with(|c| c.get())
        };
        let c3 = count_for(3);
        let c7 = count_for(7);
        // Exact model: 1 (master build) + 1 (layout cache build) + 2·N
        // (per-slide: slide XML + layout decorative walk).
        assert_eq!(c3, 2 + 2 * 3, "3-slide deck D4 parse count");
        assert_eq!(c7, 2 + 2 * 7, "7-slide deck D4 parse count");
        // Slope check: exactly 2 extra parses per added slide (NOT 12+ or 4·k),
        // i.e. the master build + layout parse are amortized to O(1), not O(N).
        assert_eq!(
            (c7 - c3) / (7 - 3),
            2,
            "per-slide D4 parse slope must be 2 (slide + layout-decorative), \
             proving master/layout parses are cached, not per-slide"
        );
    }

    #[test]
    fn out_of_order_single_slide_producer_matches_legacy_drain() {
        let data = build_shared_layout_deck(4);
        let legacy = parse_presentation_from_bytes(&data).expect("legacy drain parses");
        let mut zip = PptxZip::new(Cursor::new(data)).expect("archive opens");
        let produced = zip
            .run_operation("single-slide-test", |zip| {
                let mut shared = bootstrap_presentation(zip).map_err(|e| e.to_string())?;
                // Produce slide 3 first: no earlier slide may be required to warm
                // the master/layout caches or establish its stable index.
                produce_slide_unit(3, &mut shared, zip).map_err(|e| e.to_string())
            })
            .expect("out-of-order slide parses");
        assert_eq!(
            serde_json::to_value(&produced).unwrap(),
            serde_json::to_value(&legacy.slides[3]).unwrap()
        );
    }

    #[test]
    fn explicit_all_slide_producer_drain_assembles_ordered_presentation() {
        let data = build_shared_layout_deck(5);
        let mut zip = PptxZip::new(Cursor::new(data)).expect("archive opens");
        let produced = zip
            .run_operation("explicit-drain-test", |zip| {
                let mut shared = bootstrap_presentation(zip).map_err(|e| e.to_string())?;
                let mut slides = Vec::with_capacity(shared.slide_descriptors.len());
                for index in 0..shared.slide_descriptors.len() {
                    slides.push(
                        produce_slide_unit(index, &mut shared, zip).map_err(|e| e.to_string())?,
                    );
                }
                shared.finish(slides).map_err(|e| e.to_string())
            })
            .expect("explicit drain parses");
        assert_eq!(
            (produced.slide_width, produced.slide_height),
            (9_144_000, 6_858_000)
        );
        assert_eq!(produced.default_text_color.as_deref(), Some("000000"));
        assert_eq!(produced.major_font.as_deref(), Some("Arial"));
        assert_eq!(produced.minor_font.as_deref(), Some("Arial"));
        assert_eq!(produced.slides.len(), 5);
        for (index, slide) in produced.slides.iter().enumerate() {
            assert_eq!(slide.index, index);
            assert_eq!(slide.slide_number, index + 1);
            assert_eq!(
                slide.part_name.as_deref(),
                Some(format!("ppt/slides/slide{index}.xml").as_str())
            );
            assert!(slide.parse_error.is_none());
        }
    }

    #[test]
    fn out_of_order_production_reuses_shared_master_and_layout_sources() {
        let data = build_shared_layout_deck(4);
        let mut zip = PptxZip::new(Cursor::new(data)).expect("archive opens");
        LAYOUT_MASTER_PARSE_COUNT.with(|c| c.set(0));
        zip.run_operation("cache-test", |zip| {
            let mut shared = bootstrap_presentation(zip).map_err(|e| e.to_string())?;
            for index in [3, 0, 2, 1] {
                produce_slide_unit(index, &mut shared, zip).map_err(|e| e.to_string())?;
            }
            assert_eq!(shared.master_cache.len(), 1);
            assert_eq!(shared.layout_source_cache.len(), 1);
            assert_eq!(shared.layout_cache.len(), 1);
            Ok(())
        })
        .expect("out-of-order production succeeds");
        let count = LAYOUT_MASTER_PARSE_COUNT.with(|c| c.get());
        assert_eq!(count, 2 + 2 * 4, "shared master/layout parse count");
    }

    /// PowerPoint stores equations as `a14:m` inside `mc:AlternateContent`
    /// (ECMA-376 §22.1 OMML + 2010 drawing ext). The Choice branch holds the
    /// live `m:oMathPara`; the Fallback (a rasterized picture/text) must be
    /// ignored so the equation isn't double-rendered.
    #[test]
    fn extracts_math_from_alternatecontent_a14m() {
        let xml = r#"<p
            xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
          <mc:AlternateContent>
            <mc:Choice Requires="a14">
              <a14:m>
                <m:oMathPara><m:oMath>
                  <m:r><m:t>x</m:t></m:r>
                </m:oMath></m:oMathPara>
              </a14:m>
            </mc:Choice>
            <mc:Fallback><r><t>fallback</t></r></mc:Fallback>
          </mc:AlternateContent>
        </p>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let ac = doc
            .root_element()
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "AlternateContent")
            .unwrap();
        let theme = HashMap::new();
        let mut runs = Vec::new();
        push_math_runs(ac, Some(18.0), &theme, &mut runs);
        assert_eq!(runs.len(), 1, "exactly one math run, fallback ignored");
        match &runs[0] {
            TextRun::Math {
                display,
                nodes,
                font_size,
                ..
            } => {
                assert!(*display, "oMathPara → display math");
                assert_eq!(*font_size, Some(18.0));
                assert_eq!(nodes_to_text(nodes), "x");
            }
            other => panic!("expected math run, got {other:?}"),
        }
    }

    /// PowerPoint also stores INLINE math as a bare `a14:m` (local name "m")
    /// directly under `a:p` — not wrapped in AlternateContent — holding an
    /// `m:oMath` (not oMathPara). It must extract as inline (display:false) and
    /// pick up its run size from the math run's rPr `sz` (hundredths of a pt).
    #[test]
    fn extracts_inline_bare_a14m_with_run_size() {
        let xml = r#"<m
            xmlns="http://schemas.microsoft.com/office/drawing/2010/main"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
          <m:oMath><m:r>
            <a:rPr sz="2800" i="1"><a:solidFill><a:srgbClr val="7030A0"/></a:solidFill></a:rPr>
            <m:t>n</m:t>
          </m:r></m:oMath>
        </m>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let mut runs = Vec::new();
        push_math_runs(doc.root_element(), None, &theme, &mut runs);
        assert_eq!(runs.len(), 1);
        match &runs[0] {
            TextRun::Math {
                display,
                font_size,
                nodes,
                color,
            } => {
                assert!(!*display, "bare a14:m with m:oMath → inline");
                assert_eq!(*font_size, Some(28.0), "size read from math run rPr sz");
                assert_eq!(
                    color.as_deref(),
                    Some("7030A0"),
                    "colour read from math run rPr solidFill"
                );
                assert_eq!(nodes_to_text(nodes), "n");
            }
            other => panic!("expected math run, got {other:?}"),
        }
    }

    /// ECMA-376 §21.1.2.4 / §19.3.1 — a slide body placeholder bound by `idx`
    /// whose layout shape sets size-but-not-colour must still inherit the master
    /// `txStyles` bodyStyle colour (keyed by placeholder *type*). The idx-strict
    /// rule only blocks a sibling *layout* placeholder from leaking its colour; it
    /// must NOT block the master's type-keyed document default.
    ///
    /// Regression: sample-9 slide 2+ body text rendered black instead of the
    /// master bodyStyle's `schemeClr val="bg1"` (→ lt1 → white on a dark theme),
    /// because `lookup_color` returned early on a missing `by_idx_color` entry.
    #[test]
    fn idx_placeholder_inherits_master_txstyle_color() {
        let mut lph = LayoutPlaceholders::default();
        // Master bodyStyle resolves to white and is keyed by type (incl. "" and "body").
        lph.by_type_master_color
            .insert("body".to_string(), "FFFFFF".to_string());
        lph.by_type_master_color
            .insert("".to_string(), "FFFFFF".to_string());

        // Layout idx=35 placeholder declared size only → no by_idx_color entry.
        assert_eq!(
            lph.lookup_color("body", Some(35)),
            Some("FFFFFF".to_string()),
            "idx-bound body placeholder must fall through to the master bodyStyle colour"
        );

        // The layout idx colour still wins when present (idx-strict for the layout tier).
        lph.by_idx_color.insert(35, "112233".to_string());
        assert_eq!(
            lph.lookup_color("body", Some(35)),
            Some("112233".to_string()),
            "an explicit layout idx colour takes priority over the master default"
        );
    }

    /// ECMA-376 §20.1.4.2.27 (`CT_TableStyleCellStyle`) — a cell style's fill is
    /// wrapped in `<a:fill>` and its text colour lives in `<a:tcTxStyle>`. Both the
    /// `firstRow` (header) and `wholeTbl` roles must resolve. Regression: sample-9
    /// slides 9–10 — the orange header fill / pink banding never rendered (fill was
    /// parsed off `<a:tcStyle>` directly, missing the `<a:fill>` wrapper) and the
    /// white header text was ignored (tcTxStyle was never read).
    #[test]
    fn table_style_resolves_fill_wrapper_and_tctxstyle_colour() {
        let theme: HashMap<String, String> =
            [("dk1", "000000"), ("lt1", "FFFFFF"), ("accent2", "B83903")]
                .iter()
                .map(|(k, v)| (k.to_string(), v.to_string()))
                .collect();

        let xml = r#"<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:tblStyle styleId="{TEST}" styleName="Medium Style 1 - Accent 2">
            <a:wholeTbl>
              <a:tcTxStyle b="def" i="on">
                <a:fontRef idx="minor"><a:scrgbClr r="0" g="0" b="0"/></a:fontRef>
                <a:schemeClr val="dk1"/>
              </a:tcTxStyle>
              <a:tcStyle>
                <a:tcBdr>
                  <a:insideH><a:ln w="12700"><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></a:ln></a:insideH>
                </a:tcBdr>
                <a:fill><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:fill>
              </a:tcStyle>
            </a:wholeTbl>
            <a:band1H>
              <a:tcTxStyle b="on"><a:fontRef idx="minor"/><a:schemeClr val="accent2"/></a:tcTxStyle>
              <a:tcStyle>
                <a:tcBdr/>
                <a:fill><a:solidFill><a:schemeClr val="accent2"><a:tint val="20000"/></a:schemeClr></a:solidFill></a:fill>
              </a:tcStyle>
            </a:band1H>
            <a:firstRow>
              <a:tcTxStyle b="on">
                <a:fontRef idx="minor"><a:scrgbClr r="0" g="0" b="0"/></a:fontRef>
                <a:schemeClr val="lt1"/>
              </a:tcTxStyle>
              <a:tcStyle>
                <a:tcBdr/>
                <a:fill><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></a:fill>
              </a:tcStyle>
            </a:firstRow>
            <a:nwCell><a:tcTxStyle b="off"><a:fontRef idx="minor"/><a:schemeClr val="dk1"/></a:tcTxStyle></a:nwCell>
          </a:tblStyle>
        </a:tblStyleLst>"#;

        let map = parse_table_styles_xml(xml, &theme);
        let def = map.get("{TEST}").expect("style parsed");

        // Fills (wrapped in <a:fill>) must resolve.
        let solid = |f: &Option<Fill>| match f {
            Some(Fill::Solid { color }) => Some(color.clone()),
            _ => None,
        };
        assert_eq!(
            solid(&def.whole_tbl.fill).as_deref(),
            Some("FFFFFF"),
            "wholeTbl fill should be lt1 white"
        );
        assert_eq!(
            solid(&def.first_row.fill).as_deref(),
            Some("B83903"),
            "firstRow header fill should be accent2 orange"
        );
        // band1H = accent2 + `<a:tint val="20000">`. Table styles use the literal
        // ECMA-376 tint (val·input + (1-val)·white), giving a near-white wash —
        // NOT the saturated linear-lerp. 0.2·B83903 + 0.8·white = F1D7CD.
        assert_eq!(
            solid(&def.band1_h.fill).as_deref(),
            Some("F1D7CD"),
            "band1H tint should be the literal near-white wash, not a saturated lerp"
        );

        // Text colours from tcTxStyle.
        assert_eq!(
            def.whole_tbl.text.color.as_deref(),
            Some("000000"),
            "wholeTbl text colour should be dk1 black"
        );
        assert_eq!(
            def.first_row.text.color.as_deref(),
            Some("FFFFFF"),
            "firstRow header text colour should be lt1 white"
        );

        // firstRow `<a:tcTxStyle b="on">` → bold header.
        assert_eq!(
            def.first_row.text.bold,
            Some(true),
            "firstRow header should be bold from tcTxStyle b=on"
        );
        assert_eq!(
            def.whole_tbl.text.bold, None,
            "b=def inherits instead of forcing off"
        );
        assert_eq!(def.whole_tbl.text.italic, Some(true));
        assert_eq!(def.band1_h.text.color.as_deref(), Some("B83903"));
        assert_eq!(def.band1_h.text.bold, Some(true));
        assert_eq!(def.nw_cell.text.bold, Some(false));
    }

    #[test]
    fn table_style_preserves_vertical_bands_corners_and_all_eight_border_roles() {
        let theme = HashMap::new();
        let xml = r#"<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:tblStyle styleId="{FULL}" styleName="Full">
            <a:wholeTbl><a:tcStyle><a:tcBdr>
              <a:left><a:ln w="100"><a:solidFill><a:srgbClr val="110000"/></a:solidFill></a:ln></a:left>
              <a:right><a:ln w="200"><a:solidFill><a:srgbClr val="220000"/></a:solidFill></a:ln></a:right>
              <a:top><a:ln w="300"><a:solidFill><a:srgbClr val="330000"/></a:solidFill></a:ln></a:top>
              <a:bottom><a:ln w="400"><a:solidFill><a:srgbClr val="440000"/></a:solidFill></a:ln></a:bottom>
              <a:insideH><a:ln w="500"><a:solidFill><a:srgbClr val="550000"/></a:solidFill></a:ln></a:insideH>
              <a:insideV><a:ln w="600"><a:solidFill><a:srgbClr val="660000"/></a:solidFill></a:ln></a:insideV>
              <a:tl2br><a:ln w="700"><a:solidFill><a:srgbClr val="770000"/></a:solidFill></a:ln></a:tl2br>
              <a:tr2bl><a:ln w="800"><a:solidFill><a:srgbClr val="880000"/></a:solidFill></a:ln></a:tr2bl>
            </a:tcBdr></a:tcStyle></a:wholeTbl>
            <a:band1V><a:tcStyle><a:fill><a:solidFill><a:srgbClr val="00AA00"/></a:solidFill></a:fill></a:tcStyle></a:band1V>
            <a:band2V><a:tcStyle><a:fill><a:solidFill><a:srgbClr val="00BB00"/></a:solidFill></a:fill></a:tcStyle></a:band2V>
            <a:nwCell><a:tcStyle><a:tcBdr><a:bottom><a:ln><a:noFill/></a:ln></a:bottom></a:tcBdr>
              <a:fill><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></a:fill></a:tcStyle></a:nwCell>
          </a:tblStyle>
        </a:tblStyleLst>"#;
        let style = parse_table_styles_xml(xml, &theme)
            .remove("{FULL}")
            .expect("style parsed");
        let solid_color = |fill: &Option<Fill>| match fill {
            Some(Fill::Solid { color }) => Some(color.clone()),
            _ => None,
        };
        assert_eq!(solid_color(&style.band1_v.fill).as_deref(), Some("00AA00"));
        assert_eq!(solid_color(&style.band2_v.fill).as_deref(), Some("00BB00"));
        assert_eq!(solid_color(&style.nw_cell.fill).as_deref(), Some("ABCDEF"));
        assert!(matches!(
            style.nw_cell.borders.bottom,
            TableLineStyle::NoLine
        ));
        for border in [
            &style.whole_tbl.borders.left,
            &style.whole_tbl.borders.right,
            &style.whole_tbl.borders.top,
            &style.whole_tbl.borders.bottom,
            &style.whole_tbl.borders.inside_h,
            &style.whole_tbl.borders.inside_v,
            &style.whole_tbl.borders.diagonal_tl,
            &style.whole_tbl.borders.diagonal_tr,
        ] {
            assert!(matches!(border, TableLineStyle::Stroke(_)));
        }
    }

    #[test]
    fn table_style_refs_use_complete_theme_fill_and_line_recipes() {
        let theme = crate::theme::PptxTheme::from_xml(
            r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:themeElements>
                <a:clrScheme name="table"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:accent1><a:srgbClr val="112233"/></a:accent1></a:clrScheme>
                <a:fontScheme name="table"><a:majorFont/><a:minorFont/></a:fontScheme>
                <a:fmtScheme name="table">
                  <a:fillStyleLst><a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"/></a:gs><a:gs pos="100000"><a:srgbClr val="ABCDEF"/></a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill></a:fillStyleLst>
                  <a:lnStyleLst><a:ln w="25400" cap="rnd"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="dash"/><a:round/></a:ln><a:ln><a:noFill/></a:ln></a:lnStyleLst>
                  <a:effectStyleLst/><a:bgFillStyleLst/>
                </a:fmtScheme>
              </a:themeElements>
            </a:theme>"#,
        );
        let styles = parse_table_styles_xml(
            r#"<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:tblStyle styleId="{REF}"><a:wholeTbl><a:tcStyle><a:tcBdr><a:left><a:lnRef idx="1"><a:srgbClr val="445566"/></a:lnRef></a:left></a:tcBdr><a:fillRef idx="1"><a:srgbClr val="778899"/></a:fillRef></a:tcStyle></a:wholeTbl><a:firstRow><a:tcStyle><a:tcBdr><a:left><a:lnRef idx="2"/></a:left></a:tcBdr></a:tcStyle></a:firstRow></a:tblStyle></a:tblStyleLst>"#,
            &theme,
        );
        let style = styles.get("{REF}").expect("table style");
        assert!(
            matches!(style.whole_tbl.fill, Some(Fill::Gradient { ref stops, angle, .. }) if stops[0].color == "778899" && angle == 90.0)
        );
        assert!(
            matches!(style.whole_tbl.borders.left, TableLineStyle::Stroke(ref stroke) if stroke.color == "445566" && stroke.width == 25400 && stroke.dash_style.as_deref() == Some("dash") && stroke.line_cap.as_deref() == Some("round"))
        );
        assert!(matches!(
            style.first_row.borders.left,
            TableLineStyle::NoLine
        ));
        let resolved = resolve_table_cell_style(
            style,
            TableStyleFlags {
                first_row: true,
                ..Default::default()
            },
            0,
            0,
            2,
            2,
        );
        assert!(matches!(resolved.border_l, TableLineStyle::NoLine));
    }

    #[test]
    fn table_style_cascade_orders_bands_conditionals_and_corner_and_keeps_asymmetric_edges() {
        let mut style = TableStyleDef::default();
        let solid_color = |fill: &Option<Fill>| match fill {
            Some(Fill::Solid { color }) => Some(color.clone()),
            _ => None,
        };
        let fill = |color: &str| {
            Some(Fill::Solid {
                color: color.to_owned(),
            })
        };
        let line = |color: &str| {
            TableLineStyle::Stroke(Box::new(Stroke {
                color: color.to_owned(),
                width: 100,
                fill: None,
                dash_style: None,
                custom_dash: Vec::new(),
                line_cap: None,
                line_join: None,
                miter_limit: None,
                alignment: None,
                head_end: None,
                tail_end: None,
                cmpd: None,
            }))
        };
        style.whole_tbl.fill = fill("WHOLE");
        style.whole_tbl.borders.left = line("LEFT");
        style.whole_tbl.borders.right = line("RIGHT");
        style.whole_tbl.borders.top = line("TOP");
        style.whole_tbl.borders.bottom = line("BOTTOM");
        style.whole_tbl.borders.inside_h = line("INSIDE-H");
        style.whole_tbl.borders.inside_v = line("INSIDE-V");
        style.band1_h.fill = fill("BAND-ROW");
        style.band1_v.fill = fill("BAND-COL");
        style.first_row.fill = fill("FIRST-ROW");
        style.first_col.fill = fill("FIRST-COL");
        style.nw_cell.fill = fill("CORNER");
        style.nw_cell.borders.bottom = TableLineStyle::NoLine;

        let resolved = resolve_table_cell_style(
            &style,
            TableStyleFlags {
                first_row: true,
                first_col: true,
                band_row: true,
                band_col: true,
                ..Default::default()
            },
            0,
            0,
            3,
            3,
        );
        assert_eq!(solid_color(&resolved.fill).as_deref(), Some("CORNER"));
        assert!(matches!(resolved.border_b, TableLineStyle::NoLine));
        assert!(
            matches!(resolved.border_l, TableLineStyle::Stroke(ref stroke) if stroke.color == "LEFT")
        );
        assert!(
            matches!(resolved.border_t, TableLineStyle::Stroke(ref stroke) if stroke.color == "TOP")
        );

        let bottom_right = resolve_table_cell_style(&style, TableStyleFlags::default(), 2, 2, 3, 3);
        assert!(
            matches!(bottom_right.border_r, TableLineStyle::Stroke(ref stroke) if stroke.color == "RIGHT")
        );
        assert!(
            matches!(bottom_right.border_b, TableLineStyle::Stroke(ref stroke) if stroke.color == "BOTTOM")
        );

        let first_vertical_band = resolve_table_cell_style(
            &style,
            TableStyleFlags {
                band_col: true,
                ..Default::default()
            },
            1,
            0,
            3,
            3,
        );
        assert_eq!(
            solid_color(&first_vertical_band.fill).as_deref(),
            Some("BAND-COL")
        );
        let second_vertical_band = resolve_table_cell_style(
            &style,
            TableStyleFlags {
                band_col: true,
                ..Default::default()
            },
            1,
            1,
            3,
            3,
        );
        assert_eq!(
            solid_color(&second_vertical_band.fill).as_deref(),
            Some("WHOLE"),
            "an unspecified band2V inherits wholeTbl"
        );
    }

    /// ECMA-376 §21.1.3.17 (`CT_TableCellProperties`) — direct cell fill and
    /// line choices are the final formatting tier. Explicit noFill/no-line must
    /// suppress, rather than inherit, a lower-precedence table-style value.
    #[test]
    fn table_cell_direct_formatting_overrides_style_including_explicit_no_line() {
        let xml = r#"<a:tc xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:tcPr>
            <a:lnL><a:noFill/></a:lnL>
            <a:lnR w="250"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:lnR>
            <a:lnTlToBr><a:noFill/></a:lnTlToBr>
            <a:noFill/>
          </a:tcPr>
        </a:tc>"#;
        let doc = roxmltree::Document::parse(xml).expect("valid table cell");
        let theme = HashMap::new();
        let rels = HashMap::new();
        let mut zip = PptxZip::new(Cursor::new(empty_zip_bytes())).expect("empty OOXML zip");
        let mut cell = parse_table_cell(doc.root_element(), &theme, &rels, "ppt/slides", &mut zip);

        assert!(cell.has_direct_fill);
        assert!(cell.has_direct_border_l);
        assert!(cell.has_direct_border_r);
        assert!(cell.has_direct_diagonal_tl);
        assert!(matches!(cell.fill, Some(Fill::None)));
        assert!(cell.border_l.is_none(), "direct no-line paints no stroke");
        assert_eq!(
            cell.border_r.as_ref().map(|line| line.color.as_str()),
            Some("0000FF")
        );

        let style_line = |color: &str| {
            TableLineStyle::Stroke(Box::new(Stroke {
                color: color.to_owned(),
                width: 100,
                fill: None,
                dash_style: None,
                custom_dash: Vec::new(),
                line_cap: None,
                line_join: None,
                miter_limit: None,
                alignment: None,
                head_end: None,
                tail_end: None,
                cmpd: None,
            }))
        };
        let effective = ResolvedTableCellStyle {
            fill: Some(Fill::Solid {
                color: "FF0000".to_owned(),
            }),
            border_l: style_line("STYLE-LEFT"),
            border_r: style_line("STYLE-RIGHT"),
            border_t: style_line("STYLE-TOP"),
            border_b: style_line("STYLE-BOTTOM"),
            diagonal_tl: style_line("STYLE-DIAGONAL"),
            ..Default::default()
        };
        apply_resolved_table_cell_style(&mut cell, effective);

        assert!(matches!(cell.fill, Some(Fill::None)));
        assert!(
            cell.border_l.is_none(),
            "direct no-line suppresses style line"
        );
        assert_eq!(
            cell.border_r.as_ref().map(|line| line.color.as_str()),
            Some("0000FF")
        );
        assert!(
            cell.diagonal_tl.is_none(),
            "direct diagonal no-line suppresses style line"
        );
        assert_eq!(
            cell.border_t.as_ref().map(|line| line.color.as_str()),
            Some("STYLE-TOP")
        );
        assert_eq!(
            cell.border_b.as_ref().map(|line| line.color.as_str()),
            Some("STYLE-BOTTOM")
        );
    }

    /// ECMA-376 §21.1.2.1.1 — `<a:bodyPr rtlCol="1">` lays out a multi-column
    /// text body's columns right-to-left. parse_text_body should surface it as
    /// rtl_col=true; an absent attribute yields false (and is omitted from JSON
    /// via skip_serializing_if).
    #[test]
    fn test_parse_text_body_rtl_col() {
        let theme = HashMap::new();
        let rels = HashMap::new();
        // parse_text_body now takes a &mut PptxZip (to verify buBlip parts). This
        // body declares no picture bullets, so an empty archive is sufficient.
        let bytes = empty_zip_bytes();
        let cursor = Cursor::new(bytes.clone());
        let mut zip = PptxZip::new(cursor).unwrap();
        let mut parse = |body_pr: &str| -> TextBody {
            let xml = format!(
                r#"<txBody xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">{body_pr}<p><r><t>x</t></r></p></txBody>"#
            );
            let doc = roxmltree::Document::parse(&xml).unwrap();
            parse_text_body(
                doc.root_element(),
                &theme,
                &rels,
                "ppt/slides",
                None,               // inherited_font_size
                None,               // inherited_font_family
                [None; 9],          // inherited_level_font_sizes
                Default::default(), // inherited_level_indents
                &empty_level_bullets(),
                None, // inherited_bold
                None, // inherited_italic
                None, // inherited_caps
                None, // inherited_reflection
                None, // inherited_anchor
                None, // inherited_text_insets
                None, // inherited_alignment
                None, // inherited_ea_ln_brk
                None, // inherited_space_before
                None, // inherited_space_after
                None, // inherited_line_spacing
                ShapeKind::Sp,
                &mut zip,
            )
        };

        // rtlCol="1" → true.
        let tb = parse(r#"<bodyPr numCol="2" rtlCol="1"/>"#);
        assert!(tb.rtl_col, "rtlCol=\"1\" should yield rtl_col=true");

        // rtlCol="true" is also accepted (xsd:boolean lexical form).
        let tb_true = parse(r#"<bodyPr numCol="2" rtlCol="true"/>"#);
        assert!(tb_true.rtl_col, "rtlCol=\"true\" should yield rtl_col=true");

        // Absent attribute → false (spec default).
        let tb_absent = parse(r#"<bodyPr numCol="2"/>"#);
        assert!(
            !tb_absent.rtl_col,
            "absent rtlCol should yield rtl_col=false"
        );

        // false is omitted from the serialized JSON.
        let json = serde_json::to_string(&tb_absent).unwrap();
        assert!(
            !json.contains("rtlCol"),
            "rtl_col=false must be omitted from JSON; got {json}"
        );

        // rtlCol="1" appears under the camelCase key "rtlCol".
        let json_true = serde_json::to_string(&tb).unwrap();
        assert!(
            json_true.contains("\"rtlCol\":true"),
            "expected rtlCol:true in JSON; got {json_true}"
        );
    }

    /// ECMA-376 §20.1.9.19 — `<a:bodyPr><a:prstTxWarp prst="…">` (WordArt text
    /// warp). parse_text_body should surface the preset name and its `<a:avLst>`
    /// adjust values; an absent element (or `prst="textNoShape"`) yields None,
    /// which skip_serializing_if omits from the JSON so unwarped bodies are
    /// byte-identical.
    #[test]
    fn test_parse_text_body_prst_tx_warp() {
        let theme = HashMap::new();
        let rels = HashMap::new();
        let bytes = empty_zip_bytes();
        let cursor = Cursor::new(bytes.clone());
        let mut zip = PptxZip::new(cursor).unwrap();
        let mut parse = |body_pr: &str| -> TextBody {
            let xml = format!(
                r#"<txBody xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">{body_pr}<p><r><t>x</t></r></p></txBody>"#
            );
            let doc = roxmltree::Document::parse(&xml).unwrap();
            parse_text_body(
                doc.root_element(),
                &theme,
                &rels,
                "ppt/slides",
                None,
                None,
                [None; 9],
                Default::default(),
                &empty_level_bullets(),
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                ShapeKind::Sp,
                &mut zip,
            )
        };

        // A warp with two avLst adjust values, in the real Office child order:
        // CT_TextBodyProperties is an xsd:sequence with prstTxWarp FIRST, then
        // the EG_TextAutofit group (here <spAutoFit/>) — PowerPoint emits (and
        // only honours) this order, so the fixture mimics it.
        let tb = parse(
            r#"<bodyPr wrap="none"><prstTxWarp prst="textArchUp"><avLst><gd name="adj1" fmla="val 10800000"/><gd name="adj2" fmla="val 25000"/></avLst></prstTxWarp><spAutoFit/></bodyPr>"#,
        );
        let warp = tb.text_warp.as_ref().expect("textArchUp warp present");
        assert_eq!(warp.preset, "textArchUp");
        assert_eq!(warp.adj, vec![10_800_000, 25_000]);

        // A warp with an empty avLst → preset defaults (empty adj vec).
        let tb_empty =
            parse(r#"<bodyPr><prstTxWarp prst="textWave1"><avLst/></prstTxWarp></bodyPr>"#);
        let warp_empty = tb_empty.text_warp.as_ref().expect("textWave1 warp present");
        assert_eq!(warp_empty.preset, "textWave1");
        assert!(warp_empty.adj.is_empty());
        // Empty adj is omitted from JSON.
        let json_empty = serde_json::to_string(&tb_empty).unwrap();
        assert!(
            json_empty.contains(r#""textWarp":{"preset":"textWave1"}"#),
            "empty-adj warp should omit the adj key; got {json_empty}"
        );

        // prst="textNoShape" is treated as no warp.
        let tb_none = parse(r#"<bodyPr><prstTxWarp prst="textNoShape"/></bodyPr>"#);
        assert!(tb_none.text_warp.is_none(), "textNoShape → no warp");

        // No prstTxWarp at all → None, and omitted from JSON.
        let tb_absent = parse(r#"<bodyPr/>"#);
        assert!(tb_absent.text_warp.is_none());
        let json_absent = serde_json::to_string(&tb_absent).unwrap();
        assert!(
            !json_absent.contains("textWarp"),
            "absent warp must be omitted from JSON; got {json_absent}"
        );
    }

    /// ECMA-376 §21.1.2.2.7 — `<a:pPr eaLnBrk>` (xsd:boolean, default true)
    /// controls whether East Asian words may break at a line wrap. The parser
    /// must surface the paragraph's own value, fall back to the body lstStyle
    /// lvl1pPr default, and default to true when nothing specifies it. Mirrors
    /// the `alignment` inheritance shape.
    #[test]
    fn test_parse_paragraph_ea_ln_brk() {
        let theme = HashMap::new();
        let rels = HashMap::new();
        // parse_text_body now takes a &mut PptxZip (to verify buBlip parts). No
        // picture bullets here, so an empty archive is sufficient.
        let bytes = empty_zip_bytes();
        let cursor = Cursor::new(bytes.clone());
        let mut zip = PptxZip::new(cursor).unwrap();
        // `lst_style` lets a test set the body lvl1pPr default; `p_pr` is the
        // paragraph's own pPr. Returns the single paragraph's ea_ln_brk.
        let mut parse_para = |lst_style: &str, p_pr: &str| -> Paragraph {
            let xml = format!(
                r#"<txBody xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">{lst_style}<p>{p_pr}<r><t>東</t></r></p></txBody>"#
            );
            let doc = roxmltree::Document::parse(&xml).unwrap();
            let mut tb = parse_text_body(
                doc.root_element(),
                &theme,
                &rels,
                "ppt/slides",
                None,
                None,
                [None; 9],
                Default::default(), // inherited_level_indents
                &empty_level_bullets(),
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                ShapeKind::Sp,
                &mut zip,
            );
            tb.paragraphs.remove(0)
        };

        // eaLnBrk="0" on the paragraph → false.
        assert!(
            !parse_para("", r#"<pPr eaLnBrk="0"/>"#).ea_ln_brk,
            "eaLnBrk=\"0\" should yield ea_ln_brk=false"
        );
        // eaLnBrk="false" (xsd:boolean lexical form) → false.
        assert!(
            !parse_para("", r#"<pPr eaLnBrk="false"/>"#).ea_ln_brk,
            "eaLnBrk=\"false\" should yield ea_ln_brk=false"
        );
        // Omitted everywhere → true (spec default).
        assert!(
            parse_para("", "").ea_ln_brk,
            "omitted eaLnBrk should default to ea_ln_brk=true"
        );
        // eaLnBrk="1" on the paragraph → true.
        assert!(
            parse_para("", r#"<pPr eaLnBrk="1"/>"#).ea_ln_brk,
            "eaLnBrk=\"1\" should yield ea_ln_brk=true"
        );

        // Inheritance: body lstStyle lvl1pPr eaLnBrk="0" propagates to a
        // paragraph that declares no eaLnBrk of its own.
        let inherited = parse_para(r#"<lstStyle><lvl1pPr eaLnBrk="0"/></lstStyle>"#, "");
        assert!(
            !inherited.ea_ln_brk,
            "paragraph should inherit eaLnBrk=false from body lvl1pPr"
        );
        // The paragraph's own value still wins over the inherited body default.
        let overridden = parse_para(
            r#"<lstStyle><lvl1pPr eaLnBrk="0"/></lstStyle>"#,
            r#"<pPr eaLnBrk="1"/>"#,
        );
        assert!(
            overridden.ea_ln_brk,
            "paragraph's own eaLnBrk=\"1\" should override inherited false"
        );

        // ea_ln_brk is serialized under the camelCase key "eaLnBrk".
        let json = serde_json::to_string(&parse_para("", r#"<pPr eaLnBrk="0"/>"#)).unwrap();
        assert!(
            json.contains("\"eaLnBrk\":false"),
            "expected eaLnBrk:false in JSON; got {json}"
        );
    }

    /// ECMA-376 §21.1.2.2.7 (`a:pPr@defTabSz`) — the default tab interval used by
    /// the renderer's wrap-aware tab grid (issue #1006). An explicit value is
    /// carried through as `def_tab_sz` (EMU); absent yields None (the renderer
    /// then applies the 1-inch default).
    #[test]
    fn test_parse_paragraph_def_tab_sz() {
        let theme = HashMap::new();
        let rels = HashMap::new();
        let bytes = empty_zip_bytes();
        let cursor = Cursor::new(bytes.clone());
        let mut zip = PptxZip::new(cursor).unwrap();
        let mut parse_para = |p_pr: &str| -> Paragraph {
            let xml = format!(
                r#"<txBody xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><p>{p_pr}<r><t>a</t></r></p></txBody>"#
            );
            let doc = roxmltree::Document::parse(&xml).unwrap();
            let mut tb = parse_text_body(
                doc.root_element(),
                &theme,
                &rels,
                "ppt/slides",
                None,
                None,
                [None; 9],
                Default::default(),
                &empty_level_bullets(),
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                ShapeKind::Sp,
                &mut zip,
            );
            tb.paragraphs.remove(0)
        };

        // Explicit defTabSz="914400" (1 inch) → Some(914400).
        assert_eq!(
            parse_para(r#"<pPr algn="l" defTabSz="914400"/>"#).def_tab_sz,
            Some(914400),
            "defTabSz=\"914400\" should parse to Some(914400)"
        );
        // Absent → None (renderer supplies the 1-inch default).
        assert_eq!(
            parse_para("").def_tab_sz,
            None,
            "omitted defTabSz should yield None"
        );
        // A non-positive value is ignored (treated as absent).
        assert_eq!(
            parse_para(r#"<pPr defTabSz="0"/>"#).def_tab_sz,
            None,
            "defTabSz=\"0\" should be ignored"
        );
        // Serialized under the camelCase key "defTabSz"; omitted when None.
        let json = serde_json::to_string(&parse_para(r#"<pPr defTabSz="457200"/>"#)).unwrap();
        assert!(
            json.contains("\"defTabSz\":457200"),
            "expected defTabSz:457200 in JSON; got {json}"
        );
        let json_absent = serde_json::to_string(&parse_para("")).unwrap();
        assert!(
            !json_absent.contains("defTabSz"),
            "absent defTabSz must be omitted from JSON; got {json_absent}"
        );
    }

    /// ECMA-376 §21.1.3.13 (`a:tblPr@rtl`): a right-to-left table sets `rtl=true`
    /// so the renderer can place column 0 at the right edge. Absent/false must be
    /// omitted from the serialized JSON (TableElement.rtl is optional in TS).
    #[test]
    fn table_rtl_attribute_parses() {
        // An empty in-memory zip is enough: parse_table only reads
        // ppt/tableStyles.xml (absent → no style cascade) and the tbl node.
        fn parse_tbl(tbl_xml: &str) -> TableElement {
            let xml = format!(
                r#"<root xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">{tbl_xml}</root>"#
            );
            let doc = roxmltree::Document::parse(&xml).unwrap();
            let tbl = doc
                .root_element()
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "tbl")
                .unwrap();
            let t = Transform {
                x: 0,
                y: 0,
                cx: 100,
                cy: 100,
                rot: 0.0,
                flip_h: false,
                flip_v: false,
            };
            let theme: HashMap<String, String> = HashMap::new();
            let rels: HashMap<String, String> = HashMap::new();
            let bytes = empty_zip_bytes();
            let cursor = Cursor::new(bytes.clone());
            let mut zip = PptxZip::new(cursor).unwrap();
            parse_table(tbl, &t, &theme, &rels, "ppt/slides", &mut zip).unwrap()
        }

        // rtl="1" → rtl=true, serialized.
        let t_rtl = parse_tbl(
            r#"<a:tbl><a:tblPr rtl="1"/><a:tblGrid><a:gridCol w="100"/></a:tblGrid>
               <a:tr h="0"><a:tc><a:txBody/></a:tc></a:tr></a:tbl>"#,
        );
        assert!(t_rtl.rtl, "rtl=\"1\" should yield rtl=true");
        let json = serde_json::to_string(&t_rtl).unwrap();
        assert!(
            json.contains("\"rtl\":true"),
            "expected rtl:true in JSON; got {json}"
        );

        // Absent tblPr@rtl → false, omitted from JSON.
        let t_ltr = parse_tbl(
            r#"<a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="100"/></a:tblGrid>
               <a:tr h="0"><a:tc><a:txBody/></a:tc></a:tr></a:tbl>"#,
        );
        assert!(!t_ltr.rtl, "absent rtl should yield rtl=false");
        let json_ltr = serde_json::to_string(&t_ltr).unwrap();
        assert!(
            !json_ltr.contains("\"rtl\""),
            "rtl=false must be omitted; got {json_ltr}"
        );
    }

    // ===== scene3d / sp3d parsing (ECMA-376 §20.1.5.5 / §20.1.5.12) =====

    /// Wrap a `<p:spPr>` fragment with the `a:`/`p:` namespaces and return the
    /// spPr node so parse_scene3d / parse_sp3d can run against it.
    fn parse_sppr_frag<'a>(doc: &'a roxmltree::Document<'a>) -> roxmltree::Node<'a, 'a> {
        doc.root_element()
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "spPr")
            .unwrap()
    }

    #[test]
    fn test_parse_scene3d_slide3_fragment() {
        // The exact scene3d/sp3d from sample-11 slide 3, "図 3".
        let xml = r#"<root
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:spPr>
            <a:scene3d>
              <a:camera prst="perspectiveRelaxed">
                <a:rot lat="19800000" lon="1200000" rev="20820000"/>
              </a:camera>
              <a:lightRig rig="threePt" dir="t"/>
            </a:scene3d>
            <a:sp3d contourW="6350" prstMaterial="matte">
              <a:bevelT w="101600" h="101600"/>
            </a:sp3d>
          </p:spPr>
        </root>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let sppr = parse_sppr_frag(&doc);

        let scene = parse_scene3d(sppr).expect("scene3d should parse");
        assert_eq!(scene.camera.prst, "perspectiveRelaxed");
        let rot = scene.camera.rot.expect("rot present");
        // 60000ths of a degree → degrees.
        assert!((rot.lat - 330.0).abs() < 1e-9, "lat = {}", rot.lat);
        assert!((rot.lon - 20.0).abs() < 1e-9, "lon = {}", rot.lon);
        assert!((rot.rev - 347.0).abs() < 1e-9, "rev = {}", rot.rev);
        // No fov/zoom in this file → None.
        assert!(scene.camera.fov.is_none());
        assert!(scene.camera.zoom.is_none());
        let lr = scene.light_rig.as_ref().expect("lightRig present");
        assert_eq!(lr.rig, "threePt");
        assert_eq!(lr.dir, "t");

        let sp3d = parse_sp3d(sppr, &HashMap::new()).expect("sp3d should parse");
        assert_eq!(sp3d.contour_w, 6350);
        assert_eq!(sp3d.prst_material, "matte");
        assert_eq!(sp3d.z, 0); // default
        assert_eq!(sp3d.extrusion_h, 0); // default
        let bt = sp3d.bevel_t.expect("bevelT present");
        assert_eq!(bt.w, 101600);
        assert_eq!(bt.h, 101600);
        assert_eq!(bt.prst, "circle"); // schema default
        assert!(sp3d.bevel_b.is_none());

        // camelCase JSON round-trip surfaces the right keys.
        let json = serde_json::to_string(&scene).unwrap();
        assert!(json.contains("\"prst\":\"perspectiveRelaxed\""), "{json}");
        assert!(json.contains("\"lat\":330.0"), "{json}");
        assert!(json.contains("\"lightRig\""), "{json}");
    }

    #[test]
    fn test_parse_camera_fov_zoom_and_defaults() {
        // fov + zoom present; sp3d with all attributes omitted → schema defaults.
        let xml = r#"<root
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:spPr>
            <a:scene3d>
              <a:camera prst="perspectiveContrastingRightFacing" fov="6900000" zoom="200000"/>
              <a:lightRig rig="threePt" dir="t"/>
            </a:scene3d>
            <a:sp3d/>
          </p:spPr>
        </root>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let sppr = parse_sppr_frag(&doc);

        let scene = parse_scene3d(sppr).unwrap();
        // fov: 6900000 / 60000 = 115 degrees.
        assert!((scene.camera.fov.unwrap() - 115.0).abs() < 1e-9);
        // zoom: 200000 / 100000 = 2.0 (200%).
        assert!((scene.camera.zoom.unwrap() - 2.0).abs() < 1e-9);
        // No <a:rot> → None (renderer uses the preset base orientation).
        assert!(scene.camera.rot.is_none());

        let sp3d = parse_sp3d(sppr, &HashMap::new()).unwrap();
        assert_eq!(sp3d.z, 0);
        assert_eq!(sp3d.extrusion_h, 0);
        assert_eq!(sp3d.contour_w, 0);
        assert_eq!(sp3d.prst_material, "warmMatte"); // schema default
        assert!(sp3d.bevel_t.is_none());
        assert!(sp3d.bevel_b.is_none());
    }

    #[test]
    fn test_parse_scene3d_absent_is_none() {
        let xml = r#"<root
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:spPr><a:prstGeom prst="rect"/></p:spPr>
        </root>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let sppr = parse_sppr_frag(&doc);
        assert!(parse_scene3d(sppr).is_none());
        assert!(parse_sp3d(sppr, &HashMap::new()).is_none());
    }

    // ===== sp3d contour colour (ECMA-376 §20.1.5.12 contourClr) =====

    #[test]
    fn test_parse_sp3d_contour_clr_slide3() {
        // The exact sp3d from sample-11 slide 3: contourW + grey contourClr.
        let xml = r#"<root
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:spPr>
            <a:sp3d contourW="6350" extrusionH="12700" prstMaterial="matte">
              <a:bevelT w="101600" h="101600"/>
              <a:contourClr><a:schemeClr val="accent1"/></a:contourClr>
              <a:extrusionClr><a:schemeClr val="accent2"/></a:extrusionClr>
            </a:sp3d>
          </p:spPr>
        </root>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let sppr = parse_sppr_frag(&doc);
        let theme = HashMap::from([
            ("accent1".to_owned(), "969696".to_owned()),
            ("accent2".to_owned(), "4472C4".to_owned()),
        ]);
        let sp3d = parse_sp3d(sppr, &theme).expect("sp3d should parse");
        assert_eq!(sp3d.contour_w, 6350);
        assert_eq!(sp3d.contour_clr.as_deref(), Some("969696"));
        assert_eq!(sp3d.extrusion_clr.as_deref(), Some("4472C4"));
        let json = serde_json::to_string(&sp3d).unwrap();
        assert!(json.contains("\"contourClr\":\"969696\""), "{json}");
        assert!(json.contains("\"extrusionClr\":\"4472C4\""), "{json}");
    }

    #[test]
    fn test_parse_sp3d_contour_clr_absent() {
        let xml = r#"<root
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:spPr><a:sp3d contourW="6350"/></p:spPr>
        </root>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let sppr = parse_sppr_frag(&doc);
        let sp3d = parse_sp3d(sppr, &HashMap::new()).unwrap();
        assert!(sp3d.contour_clr.is_none());
        // Omitted from JSON when absent.
        let json = serde_json::to_string(&sp3d).unwrap();
        assert!(!json.contains("contourClr"), "{json}");
    }

    // ===== picture a:ln stroke (ECMA-376 §20.1.2.2.24, §19.3.1.37) =====

    #[test]
    fn test_parse_pic_stroke_solid_fill() {
        // <p:pic>'s spPr > ln with a solidFill → a visible border.
        let xml = r#"<root
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:spPr>
            <a:ln w="38100"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>
          </p:spPr>
        </root>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let sppr = parse_sppr_frag(&doc);
        let theme: HashMap<String, String> = HashMap::new();
        let stroke = child(sppr, "ln")
            .and_then(|n| parse_stroke(n, &theme))
            .expect("pic stroke should parse");
        assert_eq!(stroke.color, "FFFFFF");
        assert_eq!(stroke.width, 38100);
    }

    #[test]
    fn test_parse_pic_stroke_no_fill_is_none() {
        // sample-11's pic borders are <a:ln><a:noFill/></a:ln> → no border.
        let xml = r#"<root
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:spPr><a:ln><a:noFill/></a:ln></p:spPr>
        </root>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let sppr = parse_sppr_frag(&doc);
        let theme: HashMap<String, String> = HashMap::new();
        let stroke = child(sppr, "ln").and_then(|n| parse_stroke(n, &theme));
        assert!(stroke.is_none());
    }

    // ===== p14:media-only embeds (ECMA-376 §19.3.1.17/18; the p14 extension
    // carries no audio/video tag, so media_kind is decided from the MIME of the
    // referenced part). A `<p:pic>` with no `a:videoFile`/`a:audioFile`, just a
    // `<p14:media r:embed>`, must still parse as a MediaElement — not fall
    // through to a poster-only Picture. =====

    /// `<p:pic>` whose only media marker is `<p14:media r:embed>` pointing at a
    /// `.m4v` (a MIME the table must recognise) parses as a video MediaElement.
    /// rId1 → media/clip.m4v, with a poster blip so the renderer has a thumbnail.
    #[test]
    fn test_parse_media_p14_only_m4v_is_video() {
        let xml = r#"<p:pic
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
            xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:nvPicPr>
            <p:cNvPr id="5" name="Media"/>
            <p:nvPr>
              <p:extLst>
                <p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">
                  <p14:media r:embed="rId1"/>
                </p:ext>
              </p:extLst>
            </p:nvPr>
          </p:nvPicPr>
          <p:blipFill>
            <a:blip r:embed="rId2"/>
          </p:blipFill>
          <p:spPr>
            <a:xfrm>
              <a:off x="100" y="200"/>
              <a:ext cx="3000" cy="4000"/>
            </a:xfrm>
          </p:spPr>
        </p:pic>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let pic = doc.root_element();
        let mut rels: HashMap<String, String> = HashMap::new();
        rels.insert("rId1".to_string(), "../media/clip.m4v".to_string());
        rels.insert("rId2".to_string(), "../media/image1.png".to_string());

        let media = parse_media(pic, "ppt/slides", &rels)
            .expect("p14:media-only .m4v should parse as a MediaElement");
        assert_eq!(media.id.as_deref(), Some("5"));
        assert_eq!(media.media_kind, "video");
        assert_eq!(media.mime_type, "video/mp4");
        assert_eq!(media.media_path, "ppt/media/clip.m4v");
        assert_eq!(media.poster_path, "ppt/media/image1.png");
    }

    /// Same shape but the embed targets a `.wav` → audio MediaElement.
    #[test]
    fn test_parse_media_p14_only_wav_is_audio() {
        let xml = r#"<p:pic
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
            xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:nvPicPr>
            <p:cNvPr id="6" name="Audio"/>
            <p:nvPr>
              <p:extLst>
                <p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">
                  <p14:media r:embed="rId1"/>
                </p:ext>
              </p:extLst>
            </p:nvPr>
          </p:nvPicPr>
          <p:spPr>
            <a:xfrm>
              <a:off x="0" y="0"/>
              <a:ext cx="800" cy="800"/>
            </a:xfrm>
          </p:spPr>
        </p:pic>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let pic = doc.root_element();
        let mut rels: HashMap<String, String> = HashMap::new();
        rels.insert("rId1".to_string(), "../media/sound.wav".to_string());

        let media = parse_media(pic, "ppt/slides", &rels)
            .expect("p14:media-only .wav should parse as a MediaElement");
        assert_eq!(media.media_kind, "audio");
        assert_eq!(media.mime_type, "audio/wav");
    }

    /// A `<p:pic>` whose legacy `<a:videoFile r:link>` is broken — here modeled
    /// as a missing rId (`rIdBroken` is absent from rels, so `rels.get` is None)
    /// — but whose `<p14:media r:embed>` points at the real embedded clip must
    /// still parse as a video: the good embed must not be shadowed by the broken
    /// link. This exercises the embed-before-link ordering, not the empty-Target
    /// guard (a real External link would instead carry a non-empty URL).
    #[test]
    fn test_parse_media_prefers_p14_embed_over_broken_videofile_link() {
        let xml = r#"<p:pic
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
            xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:nvPicPr>
            <p:cNvPr id="7" name="Video"/>
            <p:nvPr>
              <a:videoFile r:link="rIdBroken"/>
              <p:extLst>
                <p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">
                  <p14:media r:embed="rIdGood"/>
                </p:ext>
              </p:extLst>
            </p:nvPr>
          </p:nvPicPr>
          <p:blipFill><a:blip r:embed="rIdPoster"/></p:blipFill>
          <p:spPr>
            <a:xfrm>
              <a:off x="0" y="0"/>
              <a:ext cx="1280" cy="720"/>
            </a:xfrm>
          </p:spPr>
        </p:pic>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let pic = doc.root_element();
        let mut rels: HashMap<String, String> = HashMap::new();
        // rIdBroken intentionally absent — the link's rId does not resolve
        // (`rels.get` is None). Only the embedded p14:media resolves.
        rels.insert("rIdGood".to_string(), "../media/clip.mp4".to_string());
        rels.insert("rIdPoster".to_string(), "../media/image1.png".to_string());

        let media = parse_media(pic, "ppt/slides", &rels)
            .expect("a broken videoFile link must not shadow the good p14:media embed");
        assert_eq!(media.media_kind, "video");
        assert_eq!(media.media_path, "ppt/media/clip.mp4");
        assert_eq!(media.mime_type, "video/mp4");
    }

    // ===== Master spTree decorative shapes (ECMA-376 §19.3.1.38 sld /
    // §19.3.1.39 sldLayout, showMasterSp) =====

    /// Build a minimal in-memory .pptx whose slide master spTree carries a
    /// decorative picture (image1.png at a non-centred position) plus a
    /// solid-fill rectangle. `layout_show_master_sp` controls the layout's
    /// `showMasterSp` attribute so the test can exercise the suppression path.
    fn build_master_sp_pptx(
        layout_show_master_sp: Option<bool>,
        slide_show_master_sp: Option<bool>,
        include_layout_and_slide_shapes: bool,
    ) -> Vec<u8> {
        use zip::write::SimpleFileOptions;

        // 1×1 transparent PNG (smallest valid PNG).
        const PNG_1X1: &[u8] = &[
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48,
            0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00,
            0x00, 0x1F, 0x15, 0xC4, 0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x78,
            0x9C, 0x62, 0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
            0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
        ];

        let layout_attr = match layout_show_master_sp {
            Some(true) => r#" showMasterSp="1""#.to_string(),
            Some(false) => r#" showMasterSp="0""#.to_string(),
            None => String::new(),
        };
        let slide_attr = match slide_show_master_sp {
            Some(true) => r#" showMasterSp="1""#,
            Some(false) => r#" showMasterSp="0""#,
            None => "",
        };
        let layout_shape = if include_layout_and_slide_shapes {
            r#"<p:sp><p:nvSpPr><p:cNvPr id="20" name="LayoutBand"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="300000"/><a:ext cx="1000000" cy="100000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>"#
        } else {
            ""
        };
        let slide_shape = if include_layout_and_slide_shapes {
            r#"<p:sp><p:nvSpPr><p:cNvPr id="30" name="SlideShape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="500000"/><a:ext cx="1000000" cy="100000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>"#
        } else {
            ""
        };

        let presentation_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMaster"/></p:sldMasterIdLst>
  <p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>"#;

        let pres_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>"#;

        let theme_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">
  <a:themeElements><a:clrScheme name="C">
    <a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
    <a:dk2><a:srgbClr val="111111"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
    <a:accent1><a:srgbClr val="FF0000"/></a:accent1><a:accent2><a:srgbClr val="00FF00"/></a:accent2>
    <a:accent3><a:srgbClr val="0000FF"/></a:accent3><a:accent4><a:srgbClr val="FFFF00"/></a:accent4>
    <a:accent5><a:srgbClr val="FF00FF"/></a:accent5><a:accent6><a:srgbClr val="00FFFF"/></a:accent6>
    <a:hlink><a:srgbClr val="0000EE"/></a:hlink><a:folHlink><a:srgbClr val="551A8B"/></a:folHlink>
  </a:clrScheme>
  <a:fontScheme name="F"><a:majorFont><a:latin typeface="Arial"/></a:majorFont>
    <a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>
  <a:fmtScheme name="S"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>
  </a:themeElements>
</a:theme>"#;

        // Master spTree: a decorative pic (image1.png at x=600000,y=400000) and a
        // solid-fill rectangle. No placeholder, so both are decorative.
        let master_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:pic>
      <p:nvPicPr><p:cNvPr id="10" name="MasterLogo"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
      <p:blipFill><a:blip r:embed="rIdImg1"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
      <p:spPr><a:xfrm><a:off x="600000" y="400000"/><a:ext cx="800000" cy="800000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
    </p:pic>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="11" name="MasterBand"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="200000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="123456"/></a:solidFill></p:spPr>
    </p:sp>
  </p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"
    accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rIdLayout"/></p:sldLayoutIdLst>
</p:sldMaster>"#;

        let master_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rIdImg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;

        let layout_xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"{layout_attr} type="blank">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    {layout_shape}
  </p:spTree></p:cSld>
</p:sldLayout>"#
        );

        let layout_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"#;

        let slide_xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"{slide_attr}>
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    {slide_shape}
  </p:spTree></p:cSld>
</p:sld>"#
        );

        let slide_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"#;

        let mut buf = Vec::new();
        {
            let cursor = Cursor::new(&mut buf);
            let mut zw = zip::ZipWriter::new(cursor);
            let opts = SimpleFileOptions::default();
            let mut put = |path: &str, bytes: &[u8]| {
                zw.start_file(path, opts).unwrap();
                use std::io::Write;
                zw.write_all(bytes).unwrap();
            };
            put("ppt/presentation.xml", presentation_xml.as_bytes());
            put("ppt/_rels/presentation.xml.rels", pres_rels.as_bytes());
            put("ppt/theme/theme1.xml", theme_xml.as_bytes());
            put("ppt/slideMasters/slideMaster1.xml", master_xml.as_bytes());
            put(
                "ppt/slideMasters/_rels/slideMaster1.xml.rels",
                master_rels.as_bytes(),
            );
            put("ppt/slideLayouts/slideLayout1.xml", layout_xml.as_bytes());
            put(
                "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
                layout_rels.as_bytes(),
            );
            put("ppt/slides/slide1.xml", slide_xml.as_bytes());
            put("ppt/slides/_rels/slide1.xml.rels", slide_rels.as_bytes());
            put("ppt/media/image1.png", PNG_1X1);
            zw.finish().unwrap();
        }
        buf
    }

    /// §19.3.1.38/§19.3.1.39: a master spTree picture (non-placeholder) is
    /// composited onto the slide. Without the fix the master spTree is dropped
    /// and the slide has no elements.
    #[test]
    fn master_sptree_pic_appears_on_slide() {
        let data = build_master_sp_pptx(None, None, false);
        let pres = parse_presentation_from_bytes(&data).expect("parse");
        let slide = &pres.slides[0];

        let pic = slide.elements.iter().find_map(|e| match e {
            SlideElement::Picture(p) => Some(p),
            _ => None,
        });
        let pic = pic.expect("master decorative picture should be rendered on the slide");
        // Non-centred position from the master xfrm is preserved.
        assert_eq!(pic.x, 600000, "master pic x");
        assert_eq!(pic.y, 400000, "master pic y");
        assert!(
            pic.image_path.ends_with("media/image1.png"),
            "master pic should resolve image1.png via master rels; got {}",
            pic.image_path
        );
        assert_eq!(pic.mime_type, "image/png", "master pic mime");

        // The decorative rectangle also shows up.
        let has_band = slide
            .elements
            .iter()
            .any(|e| matches!(e, SlideElement::Shape(_)));
        assert!(has_band, "master decorative shape should be rendered");
        assert_eq!(slide.element_sources.len(), slide.elements.len());
        assert!(
            slide
                .element_sources
                .iter()
                .all(|source| { source.origin == SlideElementOrigin::Master }),
            "every inherited master decoration must retain master provenance"
        );
    }

    #[test]
    fn element_sources_distinguish_composite_paint_origins() {
        let data = build_master_sp_pptx(None, None, true);
        let pres = parse_presentation_from_bytes(&data).expect("parse");
        let slide = &pres.slides[0];

        assert_eq!(slide.element_sources.len(), slide.elements.len());
        assert_eq!(
            slide
                .element_sources
                .iter()
                .map(|source| source.origin)
                .collect::<Vec<_>>(),
            vec![
                SlideElementOrigin::Master,
                SlideElementOrigin::Master,
                SlideElementOrigin::Layout,
                SlideElementOrigin::Slide,
            ],
        );
    }

    /// §19.3.1.39: a layout with showMasterSp="0" suppresses the master's
    /// decorative shapes for slides using that layout.
    #[test]
    fn master_sptree_hidden_when_layout_show_master_sp_false() {
        let data = build_master_sp_pptx(Some(false), None, false);
        let pres = parse_presentation_from_bytes(&data).expect("parse");
        let slide = &pres.slides[0];

        let has_master_pic = slide
            .elements
            .iter()
            .any(|e| matches!(e, SlideElement::Picture(_)));
        assert!(
            !has_master_pic,
            "showMasterSp=\"0\" on the layout must suppress master decorations"
        );
        assert!(
            slide.elements.is_empty(),
            "no master decorations expected; got {} elements",
            slide.elements.len()
        );
    }

    /// ECMA-376 §19.3.1.38 scopes showMasterSp to shapes on the master slide.
    /// Layout-local decorations remain part of the selected layout.
    #[test]
    fn slide_show_master_sp_false_hides_master_but_keeps_layout_decorations() {
        let data = build_master_sp_pptx(None, Some(false), true);
        let pres = parse_presentation_from_bytes(&data).expect("parse");
        let slide = &pres.slides[0];

        assert_eq!(slide.elements.len(), 2);
        assert_eq!(slide.element_sources.len(), 2);
        assert_eq!(slide.element_sources[0].origin, SlideElementOrigin::Layout);
        assert_eq!(slide.element_sources[1].origin, SlideElementOrigin::Slide);
    }

    /// showMasterSp="1" (explicit true) on the layout keeps master shapes —
    /// guards against an inverted boolean parse.
    #[test]
    fn master_sptree_shown_when_layout_show_master_sp_true() {
        let data = build_master_sp_pptx(Some(true), None, false);
        let pres = parse_presentation_from_bytes(&data).expect("parse");
        let slide = &pres.slides[0];
        assert!(
            slide
                .elements
                .iter()
                .any(|e| matches!(e, SlideElement::Picture(_))),
            "showMasterSp=\"1\" must keep master decorations"
        );
    }

    /// Build a minimal in-memory .pptx whose master carries a decorative
    /// (non-placeholder) rectangle filled with a `schemeClr` (accent1). When
    /// `slide_clr_map_ovr` is set, the slide gets a `<p:clrMapOvr>` remapping
    /// accent1→accent2. Exercises the master-decorative pre-extraction (D4): the
    /// no-override slide must reuse the pre-extracted element (accent1's hex),
    /// while an override slide must RE-RESOLVE the decorative fill against its
    /// override theme (accent2's hex) rather than serving the frozen bundle copy.
    fn build_master_scheme_decoration_pptx(remap_accent1_to_accent2: bool) -> Vec<u8> {
        use zip::write::SimpleFileOptions;

        let presentation_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMaster"/></p:sldMasterIdLst>
  <p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>"#;
        let pres_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rIdSlide" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>"#;
        let theme_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">
  <a:themeElements><a:clrScheme name="C">
    <a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
    <a:dk2><a:srgbClr val="111111"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
    <a:accent1><a:srgbClr val="FF0000"/></a:accent1><a:accent2><a:srgbClr val="00FF00"/></a:accent2>
    <a:accent3><a:srgbClr val="0000FF"/></a:accent3><a:accent4><a:srgbClr val="FFFF00"/></a:accent4>
    <a:accent5><a:srgbClr val="FF00FF"/></a:accent5><a:accent6><a:srgbClr val="00FFFF"/></a:accent6>
    <a:hlink><a:srgbClr val="0000EE"/></a:hlink><a:folHlink><a:srgbClr val="551A8B"/></a:folHlink>
  </a:clrScheme>
  <a:fontScheme name="F"><a:majorFont><a:latin typeface="Arial"/></a:majorFont>
    <a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>
  <a:fmtScheme name="S"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>
  </a:themeElements>
</a:theme>"#;
        // Master decorative rectangle filled with schemeClr accent1 (no placeholder).
        let master_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="11" name="MasterBand"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="200000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:schemeClr val="accent1"/></a:solidFill></p:spPr>
    </p:sp>
  </p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"
    accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rIdLayout"/></p:sldLayoutIdLst>
</p:sldMaster>"#;
        let master_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"#;
        let layout_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
</p:sldLayout>"#;
        let layout_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"#;
        // Optional slide-level clrMapOvr that remaps accent1 → accent2.
        let clr_map_ovr = if remap_accent1_to_accent2 {
            r#"<p:clrMapOvr><a:overrideClrMapping bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"
      accent1="accent2" accent2="accent2" accent3="accent3" accent4="accent4"
      accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/></p:clrMapOvr>"#
        } else {
            ""
        };
        let slide_xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  {clr_map_ovr}
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
</p:sld>"#
        );
        let slide_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"#;

        let mut buf = Vec::new();
        {
            let cursor = Cursor::new(&mut buf);
            let mut zw = zip::ZipWriter::new(cursor);
            let opts = SimpleFileOptions::default();
            let mut put = |path: &str, bytes: &[u8]| {
                zw.start_file(path, opts).unwrap();
                use std::io::Write;
                zw.write_all(bytes).unwrap();
            };
            put("ppt/presentation.xml", presentation_xml.as_bytes());
            put("ppt/_rels/presentation.xml.rels", pres_rels.as_bytes());
            put("ppt/theme/theme1.xml", theme_xml.as_bytes());
            put("ppt/slideMasters/slideMaster1.xml", master_xml.as_bytes());
            put(
                "ppt/slideMasters/_rels/slideMaster1.xml.rels",
                master_rels.as_bytes(),
            );
            put("ppt/slideLayouts/slideLayout1.xml", layout_xml.as_bytes());
            put(
                "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
                layout_rels.as_bytes(),
            );
            put("ppt/slides/slide1.xml", slide_xml.as_bytes());
            put("ppt/slides/_rels/slide1.xml.rels", slide_rels.as_bytes());
            zw.finish().unwrap();
        }
        buf
    }

    fn master_band_fill_hex(data: &[u8]) -> String {
        let pres = parse_presentation_from_bytes(data).expect("parse");
        let slide = &pres.slides[0];
        let shape = slide
            .elements
            .iter()
            .find_map(|e| match e {
                SlideElement::Shape(s) => Some(s),
                _ => None,
            })
            .expect("master decorative shape present on slide");
        match shape.fill.as_ref().expect("shape has fill") {
            Fill::Solid { color } => color.clone(),
            other => panic!("expected solid fill, got {other:?}"),
        }
    }

    /// D4 guard: the pre-extracted master decorative shape (no override) resolves
    /// its `schemeClr accent1` against the master's own theme — accent1 = FF0000.
    #[test]
    fn master_decoration_scheme_fill_no_override_uses_master_theme() {
        let hex = master_band_fill_hex(&build_master_scheme_decoration_pptx(false));
        assert_eq!(
            hex.to_uppercase(),
            "FF0000",
            "no-override slide must resolve accent1 to its master-theme hex"
        );
    }

    /// D4 guard: a slide with `<p:clrMapOvr>` remapping accent1→accent2 must
    /// RE-RESOLVE the master decorative shape against its override theme
    /// (accent2 = 00FF00), NOT serve the frozen pre-extracted copy (FF0000).
    /// This is the `eff.is_some()` re-extraction branch in `parse_slide`.
    #[test]
    fn master_decoration_scheme_fill_override_reresolves_against_override_theme() {
        let hex = master_band_fill_hex(&build_master_scheme_decoration_pptx(true));
        assert_eq!(
            hex.to_uppercase(),
            "00FF00",
            "override slide must flip the master decorative accent1→accent2 hex"
        );
    }

    /// Build a minimal in-memory .pptx whose **master** carries a
    /// `<p:bg><p:bgPr><a:blipFill>` image background and whose slide + layout
    /// have NO `<p:bg>`. Exercises the slide → layout → master background
    /// inheritance (ECMA-376 §19.3.1.1 / §20.1.8.14) end-to-end so the slide's
    /// resolved `background` should be the master's image fill.
    fn build_master_bg_blip_pptx() -> Vec<u8> {
        use zip::write::SimpleFileOptions;
        const PNG_1X1: &[u8] = &[
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48,
            0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00,
            0x00, 0x1F, 0x15, 0xC4, 0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x78,
            0x9C, 0x62, 0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
            0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
        ];

        let presentation_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMaster"/></p:sldMasterIdLst>
  <p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>"#;
        let pres_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>"#;
        let theme_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">
  <a:themeElements><a:clrScheme name="C">
    <a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
    <a:dk2><a:srgbClr val="111111"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
    <a:accent1><a:srgbClr val="FF0000"/></a:accent1><a:accent2><a:srgbClr val="00FF00"/></a:accent2>
    <a:accent3><a:srgbClr val="0000FF"/></a:accent3><a:accent4><a:srgbClr val="FFFF00"/></a:accent4>
    <a:accent5><a:srgbClr val="FF00FF"/></a:accent5><a:accent6><a:srgbClr val="00FFFF"/></a:accent6>
    <a:hlink><a:srgbClr val="0000EE"/></a:hlink><a:folHlink><a:srgbClr val="551A8B"/></a:folHlink>
  </a:clrScheme>
  <a:fontScheme name="F"><a:majorFont><a:latin typeface="Arial"/></a:majorFont>
    <a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>
  <a:fmtScheme name="S"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>
  </a:themeElements>
</a:theme>"#;
        // Master defines a blipFill background; slide + layout do not.
        let master_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg><p:bgPr><a:blipFill><a:blip r:embed="rIdImg1"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"
    accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rIdLayout"/></p:sldLayoutIdLst>
</p:sldMaster>"#;
        let master_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rIdImg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;
        let layout_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
</p:sldLayout>"#;
        let layout_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"#;
        let slide_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
</p:sld>"#;
        let slide_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"#;

        let mut buf = Vec::new();
        {
            let cursor = Cursor::new(&mut buf);
            let mut zw = zip::ZipWriter::new(cursor);
            let opts = SimpleFileOptions::default();
            let mut put = |path: &str, bytes: &[u8]| {
                zw.start_file(path, opts).unwrap();
                use std::io::Write;
                zw.write_all(bytes).unwrap();
            };
            put("ppt/presentation.xml", presentation_xml.as_bytes());
            put("ppt/_rels/presentation.xml.rels", pres_rels.as_bytes());
            put("ppt/theme/theme1.xml", theme_xml.as_bytes());
            put("ppt/slideMasters/slideMaster1.xml", master_xml.as_bytes());
            put(
                "ppt/slideMasters/_rels/slideMaster1.xml.rels",
                master_rels.as_bytes(),
            );
            put("ppt/slideLayouts/slideLayout1.xml", layout_xml.as_bytes());
            put(
                "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
                layout_rels.as_bytes(),
            );
            put("ppt/slides/slide1.xml", slide_xml.as_bytes());
            put("ppt/slides/_rels/slide1.xml.rels", slide_rels.as_bytes());
            put("ppt/media/image1.png", PNG_1X1);
            zw.finish().unwrap();
        }
        buf
    }

    /// ECMA-376 §19.3.1.1 + §20.1.8.14 — a slide with no `<p:bg>` inherits the
    /// master's `<p:bg><p:bgPr><a:blipFill>` image background through the
    /// slide → layout → master chain. The resolved `slide.background` must be a
    /// `Fill::Image` carrying the master-rels-resolved zip path.
    #[test]
    fn slide_inherits_master_blip_background() {
        let data = build_master_bg_blip_pptx();
        let pres = parse_presentation_from_bytes(&data).expect("parse");
        let bg = pres.slides[0]
            .background
            .as_ref()
            .expect("slide should inherit a background from the master");
        match bg {
            Fill::Image {
                image_path,
                mime_type,
                ..
            } => {
                assert!(
                    image_path.ends_with("media/image1.png"),
                    "master bg should resolve image1.png via master rels; got {image_path}"
                );
                assert_eq!(mime_type, "image/png", "master bg mime");
            }
            other => panic!("expected inherited Fill::Image background, got {other:?}"),
        }
    }

    // ── Embedded SVG images (Microsoft asvg:svgBlip extension) ────────────
    //
    // PowerPoint stores an SVG picture as a `<p:pic>` whose `<a:blip>` points
    // at a PNG *fallback* (r:embed) and carries the real .svg part inside an
    // `<a:extLst><a:ext uri="{96DAC541-…}"><asvg:svgBlip r:embed="…"/>`
    // extension (Microsoft 2016 SVG extension; the core blip fill is
    // ECMA-376 §20.1.8.14). The parser must keep emitting the PNG fallback's
    // zip path as `image_path` (regression-safe) while additionally surfacing
    // the SVG part's path on `svg_image_path` so the renderer can prefer the
    // vector original.

    /// Build a tiny zip containing only the two media parts a `<p:pic>` blip
    /// references (a PNG fallback and an SVG body), so `parse_picture` can be
    /// driven directly with a hand-rolled rels map.
    fn build_blip_media_zip(png: &[u8], svg: &[u8]) -> Vec<u8> {
        use zip::write::SimpleFileOptions;
        let mut buf = Vec::new();
        {
            let cursor = Cursor::new(&mut buf);
            let mut zw = zip::ZipWriter::new(cursor);
            let opts = SimpleFileOptions::default();
            zw.start_file("ppt/media/image1.png", opts).unwrap();
            {
                use std::io::Write;
                zw.write_all(png).unwrap();
            }
            zw.start_file("ppt/media/image2.svg", opts).unwrap();
            {
                use std::io::Write;
                zw.write_all(svg).unwrap();
            }
            zw.finish().unwrap();
        }
        buf
    }

    /// A `<p:pic>` carrying a PNG fallback blip plus an `asvg:svgBlip`
    /// extension must yield both the PNG `image_path` and the SVG
    /// `svg_image_path` (with mimes), never inlined base64.
    #[test]
    fn picture_with_svg_blip_extension_emits_both_urls() {
        // 1×1 transparent PNG (smallest valid PNG).
        const PNG_1X1: &[u8] = &[
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48,
            0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00,
            0x00, 0x1F, 0x15, 0xC4, 0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x78,
            0x9C, 0x62, 0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
            0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
        ];
        const SVG: &[u8] =
            br##"<svg xmlns="http://www.w3.org/2000/svg" width="2" height="2"><rect width="2" height="2" fill="#0a0"/></svg>"##;

        // The svgBlip uses a different prefix (asvg:) on purpose — matching is by
        // namespace-local name, so the prefix must not matter.
        let pic_xml = r#"<p:pic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main">
  <p:nvPicPr><p:cNvPr id="5" name="SvgPic"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="rIdPng">
      <a:extLst>
        <a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">
          <asvg:svgBlip r:embed="rIdSvg"/>
        </a:ext>
      </a:extLst>
    </a:blip>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300000" cy="300000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
</p:pic>"#;

        let doc = roxmltree::Document::parse(pic_xml).unwrap();
        let pic_node = doc.root_element();

        let mut rels = HashMap::new();
        rels.insert("rIdPng".to_string(), "../media/image1.png".to_string());
        rels.insert("rIdSvg".to_string(), "../media/image2.svg".to_string());

        let theme = HashMap::new();
        let data = build_blip_media_zip(PNG_1X1, SVG);
        let cursor = Cursor::new(data.clone());
        let mut zip = PptxZip::new(cursor).unwrap();

        let pic = parse_picture(pic_node, "ppt/slides", &rels, &theme, &mut zip)
            .expect("parse_picture should succeed for an SVG-blip picture");

        // PNG fallback is preserved as the raster image_path (regression-safe);
        // never an inlined data URL.
        assert_eq!(pic.image_path, "ppt/media/image1.png", "raster path");
        assert_eq!(pic.mime_type, "image/png", "raster mime");
        assert!(
            !pic.image_path.contains(";base64,"),
            "image_path must not inline base64; got {}",
            pic.image_path
        );

        // The SVG original is surfaced separately as a zip path.
        assert_eq!(
            pic.svg_image_path.as_deref(),
            Some("ppt/media/image2.svg"),
            "svg_image_path must point at the .svg part",
        );
        // And the resolved path must hold the original SVG bytes.
        let svg_bytes = extract_image(&data, "ppt/media/image2.svg", None, None)
            .expect("svg part must be readable by its resolved path");
        assert_eq!(
            svg_bytes, SVG,
            "bytes at svg_image_path must equal the .svg part"
        );
    }

    /// A plain `<p:pic>` with no svgBlip extension must leave `svg_image_path`
    /// as None (and still emit the PNG `image_path` + intrinsic size) — guards
    /// against the new branch firing spuriously.
    #[test]
    fn picture_without_svg_blip_has_no_svg_url() {
        const PNG_1X1: &[u8] = &[
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48,
            0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00,
            0x00, 0x1F, 0x15, 0xC4, 0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x78,
            0x9C, 0x62, 0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
            0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
        ];
        let pic_xml = r#"<p:pic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:nvPicPr><p:cNvPr id="5" name="PngPic"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
  <p:blipFill><a:blip r:embed="rIdPng"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
  <p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300000" cy="300000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
</p:pic>"#;
        let doc = roxmltree::Document::parse(pic_xml).unwrap();
        let pic_node = doc.root_element();
        let mut rels = HashMap::new();
        rels.insert("rIdPng".to_string(), "../media/image1.png".to_string());
        let theme = HashMap::new();
        let data = build_blip_media_zip(PNG_1X1, b"<svg/>");
        let cursor = Cursor::new(data.clone());
        let mut zip = PptxZip::new(cursor).unwrap();
        let pic = parse_picture(pic_node, "ppt/slides", &rels, &theme, &mut zip)
            .expect("parse_picture should succeed");
        assert_eq!(pic.image_path, "ppt/media/image1.png");
        assert_eq!(pic.mime_type, "image/png");
        // 1×1 PNG → intrinsic size read from the IHDR.
        assert_eq!(pic.intrinsic_width_px, Some(1), "intrinsic width");
        assert_eq!(pic.intrinsic_height_px, Some(1), "intrinsic height");
        assert!(
            pic.svg_image_path.is_none(),
            "svg_image_path must be None without an svgBlip extension"
        );
    }

    /// p:pic carries the same CT_ShapeStyle references as p:sp. A picture with
    /// no local line/effect component must therefore inherit lnRef/effectRef,
    /// including effect-style 3D colours resolved through phClr.
    #[test]
    fn picture_inherits_line_effect_and_3d_from_style_matrix() {
        let pic_xml = r#"<p:pic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:nvPicPr><p:cNvPr id="5" name="StyledPic"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
  <p:blipFill><a:blip r:embed="rIdPng"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
  <p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300000" cy="300000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
  <p:style>
    <a:lnRef idx="2"><a:schemeClr val="accent2"/></a:lnRef>
    <a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>
    <a:effectRef idx="1"><a:schemeClr val="accent1"/></a:effectRef>
    <a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef>
  </p:style>
</p:pic>"#;
        let doc = roxmltree::Document::parse(pic_xml).unwrap();
        let mut rels = HashMap::new();
        rels.insert("rIdPng".to_owned(), "../media/image1.png".to_owned());
        let theme = PptxTheme::from_xml(
            r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:themeElements>
                <a:clrScheme name="Test">
                  <a:accent1><a:srgbClr val="112233"/></a:accent1>
                  <a:accent2><a:srgbClr val="445566"/></a:accent2>
                </a:clrScheme>
                <a:fmtScheme name="Test">
                  <a:fillStyleLst/>
                  <a:lnStyleLst>
                    <a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
                    <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
                  </a:lnStyleLst>
                  <a:effectStyleLst><a:effectStyle>
                    <a:effectLst><a:outerShdw blurRad="12700"><a:schemeClr val="phClr"/></a:outerShdw></a:effectLst>
                    <a:scene3d><a:camera prst="orthographicFront"/></a:scene3d>
                    <a:sp3d extrusionH="25400"><a:extrusionClr><a:schemeClr val="phClr"/></a:extrusionClr></a:sp3d>
                  </a:effectStyle></a:effectStyleLst>
                  <a:bgFillStyleLst/>
                </a:fmtScheme>
              </a:themeElements>
            </a:theme>"#,
        );
        let data = build_blip_media_zip(b"png", b"<svg/>");
        let mut zip = PptxZip::new(Cursor::new(data)).unwrap();

        let pic = parse_picture(doc.root_element(), "ppt/slides", &rels, &theme, &mut zip)
            .expect("styled picture should parse");

        let stroke = pic.stroke.expect("lnRef should supply a picture border");
        assert_eq!(stroke.width, 19_050);
        assert_eq!(stroke.color, "445566");
        assert_eq!(pic.shadow.expect("effectRef shadow").color, "112233");
        assert_eq!(
            pic.scene3d.expect("effectRef scene3d").camera.prst,
            "orthographicFront"
        );
        assert_eq!(
            pic.sp3d.expect("effectRef sp3d").extrusion_clr.as_deref(),
            Some("112233")
        );
    }

    /// ECMA-376 §20.1.8.23 — a `<p:pic>` whose `<a:blip>` carries a
    /// `<a:duotone>` (a CT_Blip effect child, per the XSD sequence) parses its
    /// two `EG_ColorChoice` endpoints through the slide theme, resolving a
    /// `<a:schemeClr>` against the theme palette. `clr1` is the dark endpoint,
    /// `clr2` the light endpoint.
    #[test]
    fn picture_duotone_resolves_two_colours_through_theme() {
        const PNG_1X1: &[u8] = &[
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48,
            0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00,
            0x00, 0x1F, 0x15, 0xC4, 0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x78,
            0x9C, 0x62, 0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
            0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
        ];
        // <a:blip> holds the duotone (CT_Blip effect); clr1 = black prstClr,
        // clr2 = accent1 schemeClr (resolved from the theme map).
        let pic_xml = r#"<p:pic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:nvPicPr><p:cNvPr id="5" name="DuoPic"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="rIdPng">
      <a:duotone>
        <a:prstClr val="black"/>
        <a:schemeClr val="accent1"/>
      </a:duotone>
    </a:blip>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300000" cy="300000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
</p:pic>"#;
        let doc = roxmltree::Document::parse(pic_xml).unwrap();
        let pic_node = doc.root_element();
        let mut rels = HashMap::new();
        rels.insert("rIdPng".to_string(), "../media/image1.png".to_string());
        let mut theme = HashMap::new();
        theme.insert("accent1".to_string(), "4472C4".to_string());
        let data = build_blip_media_zip(PNG_1X1, b"<svg/>");
        let cursor = Cursor::new(data.clone());
        let mut zip = PptxZip::new(cursor).unwrap();
        let pic = parse_picture(pic_node, "ppt/slides", &rels, &theme, &mut zip)
            .expect("parse_picture should succeed for a duotone picture");
        let duo = pic.duotone.expect("duotone must be surfaced");
        assert_eq!(duo.clr1, "000000", "clr1 = black prstClr");
        assert_eq!(duo.clr2, "4472C4", "clr2 = accent1 resolved from theme");
    }

    /// A `<p:pic>` without a `<a:duotone>` leaves `duotone` None — guards the new
    /// branch from firing spuriously, so non-duotone pictures stay byte-identical.
    #[test]
    fn picture_without_duotone_is_none() {
        const PNG_1X1: &[u8] = &[
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48,
            0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00,
            0x00, 0x1F, 0x15, 0xC4, 0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x78,
            0x9C, 0x62, 0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
            0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
        ];
        let pic_xml = r#"<p:pic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:nvPicPr><p:cNvPr id="5" name="PngPic"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
  <p:blipFill><a:blip r:embed="rIdPng"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
  <p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300000" cy="300000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
</p:pic>"#;
        let doc = roxmltree::Document::parse(pic_xml).unwrap();
        let pic_node = doc.root_element();
        let mut rels = HashMap::new();
        rels.insert("rIdPng".to_string(), "../media/image1.png".to_string());
        let theme = HashMap::new();
        let data = build_blip_media_zip(PNG_1X1, b"<svg/>");
        let cursor = Cursor::new(data.clone());
        let mut zip = PptxZip::new(cursor).unwrap();
        let pic = parse_picture(pic_node, "ppt/slides", &rels, &theme, &mut zip)
            .expect("parse_picture should succeed");
        assert!(pic.duotone.is_none(), "duotone must be None when absent");
    }

    /// A `<p:pic>` whose `<a:blip>` carries ONLY the `asvg:svgBlip` extension —
    /// no raster `r:embed` fallback at all (an icon inserted as a pure SVG, as
    /// in sample-12) — must still parse. Previously the mandatory raster embed
    /// (`attr_r(&blip, "embed")?`) made `parse_picture` return None, so the whole
    /// picture was silently dropped and the SVG never rendered.
    #[test]
    fn picture_with_only_svg_blip_and_no_raster_embed_still_parses() {
        const SVG: &[u8] = br##"<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24"><path d="M0 0h24v24H0z" fill="#0a0"/></svg>"##;

        // The `<a:blip>` has NO r:embed attribute — the image is referenced only
        // through the svgBlip extension.
        let pic_xml = r#"<p:pic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main">
  <p:nvPicPr><p:cNvPr id="4" name="SvgOnly"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
  <p:blipFill>
    <a:blip>
      <a:extLst>
        <a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">
          <asvg:svgBlip r:embed="rIdSvg"/>
        </a:ext>
      </a:extLst>
    </a:blip>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300000" cy="300000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
</p:pic>"#;

        let doc = roxmltree::Document::parse(pic_xml).unwrap();
        let pic_node = doc.root_element();

        let mut rels = HashMap::new();
        rels.insert("rIdSvg".to_string(), "../media/image2.svg".to_string());

        let theme = HashMap::new();
        // Only the .svg part is referenced; the PNG arg is unused here.
        let data = build_blip_media_zip(b"", SVG);
        let cursor = Cursor::new(data.clone());
        let mut zip = PptxZip::new(cursor).unwrap();

        let pic = parse_picture(pic_node, "ppt/slides", &rels, &theme, &mut zip)
            .expect("parse_picture must succeed for an svgBlip-only picture (sample-12 case)");

        // The SVG original is surfaced on svg_image_path so the renderer prefers it.
        assert_eq!(
            pic.svg_image_path.as_deref(),
            Some("ppt/media/image2.svg"),
            "svg_image_path must point at the .svg part",
        );

        // With no raster blip, image_path falls back to the SVG part itself so
        // the element is always drawable (rather than being dropped or empty);
        // its mime is image/svg+xml and no PNG intrinsic size is recorded.
        assert_eq!(
            pic.image_path, "ppt/media/image2.svg",
            "image_path must fall back to the SVG when no raster blip is embedded",
        );
        assert_eq!(pic.mime_type, "image/svg+xml");
        assert_eq!(pic.intrinsic_width_px, None, "no PNG intrinsic for SVG");
        assert_eq!(pic.intrinsic_height_px, None);
        // The resolved SVG path must hold the original SVG bytes.
        let svg_bytes = extract_image(&data, "ppt/media/image2.svg", None, None)
            .expect("svg part must be readable by its resolved path");
        assert_eq!(
            svg_bytes, SVG,
            "bytes at svg_image_path must equal the .svg part"
        );
    }

    // ── Per-slide theme/master resolution (slide→layout→master→theme) ─────
    //
    // A deck with TWO masters, each carrying a DIFFERENT theme (different
    // accent1). Two layouts (layoutA→masterA, layoutB→masterB) and two slides
    // (slide1→layoutA, slide2→layoutB). Each slide has a shape whose fill comes
    // from `<p:style><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>`
    // with no explicit spPr fill. Before the fix the parser loaded the
    // presentation's first theme/master once and applied it to every slide, so
    // both shapes resolved to masterA's accent1. After the fix each slide must
    // resolve accent1 from its own master's theme.
    //
    // `clr_map_a` lets the test optionally give masterA a non-default
    // `<p:clrMap>` (e.g. bg1/tx1 swapped) so the clrMap-honoring assertion can
    // reuse the same builder.
    fn build_two_master_pptx(clr_map_a: &str) -> Vec<u8> {
        use zip::write::SimpleFileOptions;

        let presentation_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rIdMasterA"/>
    <p:sldMasterId id="2147483649" r:id="rIdMasterB"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rIdSlide1"/>
    <p:sldId id="257" r:id="rIdSlide2"/>
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>"#;

        // presentation rels intentionally lists masterA FIRST so the legacy
        // "first master / first theme" path would pick masterA's accent1.
        let pres_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMasterA" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rIdMasterB" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster2.xml"/>
  <Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rIdSlide2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
  <Relationship Id="rIdThemeA" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>"#;

        // Two themes that differ only in accent1 (and tx1/bg1 hex so the clrMap
        // swap is observable).
        let theme_a = |accent1: &str| {
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">
  <a:themeElements><a:clrScheme name="C">
    <a:dk1><a:srgbClr val="222222"/></a:dk1><a:lt1><a:srgbClr val="FAFAFA"/></a:lt1>
    <a:dk2><a:srgbClr val="111111"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
    <a:accent1><a:srgbClr val="{accent1}"/></a:accent1><a:accent2><a:srgbClr val="00FF00"/></a:accent2>
    <a:accent3><a:srgbClr val="0000FF"/></a:accent3><a:accent4><a:srgbClr val="FFFF00"/></a:accent4>
    <a:accent5><a:srgbClr val="FF00FF"/></a:accent5><a:accent6><a:srgbClr val="00FFFF"/></a:accent6>
    <a:hlink><a:srgbClr val="0000EE"/></a:hlink><a:folHlink><a:srgbClr val="551A8B"/></a:folHlink>
  </a:clrScheme>
  <a:fontScheme name="F"><a:majorFont><a:latin typeface="Arial"/></a:majorFont>
    <a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>
  <a:fmtScheme name="S"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>
  </a:themeElements>
</a:theme>"#
            )
        };
        let theme1_xml = theme_a("72A376"); // masterA accent1
        let theme2_xml = theme_a("4F81BD"); // masterB accent1

        let master = |clr_map: &str| {
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
  {clr_map}
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483650" r:id="rIdLayout"/></p:sldLayoutIdLst>
</p:sldMaster>"#
            )
        };
        let default_clr_map = r#"<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>"#;
        let master1_xml = master(clr_map_a);
        let master2_xml = master(default_clr_map);

        // Each master's rels points at its OWN theme and its OWN layout.
        let master1_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>"#;
        let master2_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"/>
  <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme2.xml"/>
</Relationships>"#;

        let layout = || {
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
</p:sldLayout>"#
                .to_string()
        };
        let layout1_xml = layout();
        let layout2_xml = layout();

        // layoutA→masterA, layoutB→masterB.
        let layout1_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"#;
        let layout2_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster2.xml"/>
</Relationships>"#;

        // Each slide: one rect with NO explicit fill, fill comes from
        // `<p:style><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>`.
        // slide2 additionally references tx1 on a second shape so the clrMap
        // swap (tx1→lt1) is observable.
        let slide = |extra_shape: &str| {
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="StyledRect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="500000" cy="500000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
      <p:style><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef></p:style>
    </p:sp>
    {extra_shape}
  </p:spTree></p:cSld>
</p:sld>"#
            )
        };
        let tx1_shape = r#"<p:sp>
      <p:nvSpPr><p:cNvPr id="3" name="Tx1Rect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="700000" y="100000"/><a:ext cx="500000" cy="500000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
      <p:style><a:fillRef idx="1"><a:schemeClr val="tx1"/></a:fillRef></p:style>
    </p:sp>"#;
        let slide1_xml = slide("");
        let slide2_xml = slide(tx1_shape);

        let slide1_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"#;
        let slide2_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"/>
</Relationships>"#;

        let mut buf = Vec::new();
        {
            let cursor = Cursor::new(&mut buf);
            let mut zw = zip::ZipWriter::new(cursor);
            let opts = SimpleFileOptions::default();
            let mut put = |path: &str, bytes: &[u8]| {
                zw.start_file(path, opts).unwrap();
                use std::io::Write;
                zw.write_all(bytes).unwrap();
            };
            put("ppt/presentation.xml", presentation_xml.as_bytes());
            put("ppt/_rels/presentation.xml.rels", pres_rels.as_bytes());
            put("ppt/theme/theme1.xml", theme1_xml.as_bytes());
            put("ppt/theme/theme2.xml", theme2_xml.as_bytes());
            put("ppt/slideMasters/slideMaster1.xml", master1_xml.as_bytes());
            put("ppt/slideMasters/slideMaster2.xml", master2_xml.as_bytes());
            put(
                "ppt/slideMasters/_rels/slideMaster1.xml.rels",
                master1_rels.as_bytes(),
            );
            put(
                "ppt/slideMasters/_rels/slideMaster2.xml.rels",
                master2_rels.as_bytes(),
            );
            put("ppt/slideLayouts/slideLayout1.xml", layout1_xml.as_bytes());
            put("ppt/slideLayouts/slideLayout2.xml", layout2_xml.as_bytes());
            put(
                "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
                layout1_rels.as_bytes(),
            );
            put(
                "ppt/slideLayouts/_rels/slideLayout2.xml.rels",
                layout2_rels.as_bytes(),
            );
            put("ppt/slides/slide1.xml", slide1_xml.as_bytes());
            put("ppt/slides/slide2.xml", slide2_xml.as_bytes());
            put("ppt/slides/_rels/slide1.xml.rels", slide1_rels.as_bytes());
            put("ppt/slides/_rels/slide2.xml.rels", slide2_rels.as_bytes());
            zw.finish().unwrap();
        }
        buf
    }

    fn first_shape_fill_color(slide: &Slide) -> Option<String> {
        slide.elements.iter().find_map(|e| match e {
            SlideElement::Shape(s) => match &s.fill {
                Some(Fill::Solid { color }) => Some(color.clone()),
                _ => None,
            },
            _ => None,
        })
    }

    fn shape_fill_color_by_name(slide: &Slide, name: &str) -> Option<String> {
        slide.elements.iter().find_map(|e| match e {
            SlideElement::Shape(s) if s.name.as_deref() == Some(name) => match &s.fill {
                Some(Fill::Solid { color }) => Some(color.clone()),
                _ => None,
            },
            _ => None,
        })
    }

    /// Core regression: each slide must resolve scheme colors against its OWN
    /// master's theme (slide→layout→master→theme), not the presentation's first
    /// theme. slide1's accent1 = masterA theme (#72A376); slide2's accent1 =
    /// masterB theme (#4F81BD).
    #[test]
    fn theme_resolved_per_slide_via_layout_master_chain() {
        let default_clr_map = r#"<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>"#;
        let data = build_two_master_pptx(default_clr_map);
        let pres = parse_presentation_from_bytes(&data).expect("parse");
        assert_eq!(pres.slides.len(), 2, "expected two slides");

        let s1 = first_shape_fill_color(&pres.slides[0]);
        let s2 = first_shape_fill_color(&pres.slides[1]);
        assert_eq!(
            s1.as_deref(),
            Some("72A376"),
            "slide1 accent1 must resolve from masterA theme"
        );
        assert_eq!(
            s2.as_deref(),
            Some("4F81BD"),
            "slide2 accent1 must resolve from masterB theme"
        );
    }

    /// §19.3.1.6 clrMap: a master with `bg1`/`tx1` swapped (bg1="dk1",
    /// tx1="lt1") must remap logical scheme names. `<a:schemeClr val="tx1">`
    /// then resolves to lt1's hex (#FAFAFA), not dk1's. masterB keeps the
    /// default clrMap, so its tx1 stays dk1 (#222222).
    #[test]
    fn clr_map_remaps_logical_scheme_names() {
        // Swap bg1<->tx1 on masterA only.
        let swapped = r#"<p:clrMap bg1="dk1" tx1="lt1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>"#;
        let data = build_two_master_pptx(swapped);
        let pres = parse_presentation_from_bytes(&data).expect("parse");

        // slide2 (masterB, default clrMap) has the Tx1Rect: tx1 -> dk1 (#222222).
        let tx1_default = shape_fill_color_by_name(&pres.slides[1], "Tx1Rect");
        assert_eq!(
            tx1_default.as_deref(),
            Some("222222"),
            "default clrMap: tx1 must resolve to dk1"
        );

        // To observe the swap on masterA, place the same tx1 shape via a
        // dedicated parse against masterA's theme. We reuse slide1 which uses
        // masterA; assert that accent1 still resolves correctly under the swap
        // (accent slots are identity-mapped) and that tx1 on a masterA slide
        // would map to lt1. slide1 has no tx1 shape, so we assert via the
        // builder variant below.
        let s1_accent = first_shape_fill_color(&pres.slides[0]);
        assert_eq!(
            s1_accent.as_deref(),
            Some("72A376"),
            "accent1 is identity-mapped and unaffected by bg1/tx1 swap"
        );
    }

    /// Dedicated clrMap assertion on the swapped master: a slide on masterA
    /// (bg1<->tx1 swapped) resolves `<a:schemeClr val="tx1">` to lt1 (#FAFAFA).
    #[test]
    fn clr_map_tx1_resolves_to_lt1_on_swapped_master() {
        let swapped = r#"<p:clrMap bg1="dk1" tx1="lt1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>"#;
        let data = build_two_master_pptx_with_tx1_on_a(swapped);
        let pres = parse_presentation_from_bytes(&data).expect("parse");
        // slide1 (masterA, swapped) has the Tx1Rect: tx1 -> lt1 (#FAFAFA).
        let tx1_swapped = shape_fill_color_by_name(&pres.slides[0], "Tx1Rect");
        assert_eq!(
            tx1_swapped.as_deref(),
            Some("FAFAFA"),
            "swapped clrMap: tx1 must resolve to lt1's hex"
        );
    }

    // Variant of build_two_master_pptx where slide1 (masterA) carries the tx1
    // shape, so the clrMap swap on masterA is directly observable.
    fn build_two_master_pptx_with_tx1_on_a(clr_map_a: &str) -> Vec<u8> {
        // Reuse the standard builder, then patch slide1 to include the tx1
        // shape by rebuilding with the tx1 shape on slide1. Simplest: build a
        // fresh deck inline mirroring build_two_master_pptx but swapping which
        // slide gets the tx1 shape. To avoid duplication we shell out to the
        // generic builder and post-process is not feasible on a zip, so we
        // construct directly here with the minimum needed parts.
        use zip::write::SimpleFileOptions;

        let presentation_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMasterA"/></p:sldMasterIdLst>
  <p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>"#;
        let pres_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMasterA" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rIdThemeA" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>"#;
        let theme1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">
  <a:themeElements><a:clrScheme name="C">
    <a:dk1><a:srgbClr val="222222"/></a:dk1><a:lt1><a:srgbClr val="FAFAFA"/></a:lt1>
    <a:dk2><a:srgbClr val="111111"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
    <a:accent1><a:srgbClr val="72A376"/></a:accent1><a:accent2><a:srgbClr val="00FF00"/></a:accent2>
    <a:accent3><a:srgbClr val="0000FF"/></a:accent3><a:accent4><a:srgbClr val="FFFF00"/></a:accent4>
    <a:accent5><a:srgbClr val="FF00FF"/></a:accent5><a:accent6><a:srgbClr val="00FFFF"/></a:accent6>
    <a:hlink><a:srgbClr val="0000EE"/></a:hlink><a:folHlink><a:srgbClr val="551A8B"/></a:folHlink>
  </a:clrScheme>
  <a:fontScheme name="F"><a:majorFont><a:latin typeface="Arial"/></a:majorFont>
    <a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>
  <a:fmtScheme name="S"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>
  </a:themeElements>
</a:theme>"#;
        let master1_xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
  {clr_map_a}
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483650" r:id="rIdLayout"/></p:sldLayoutIdLst>
</p:sldMaster>"#
        );
        let master1_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>"#;
        let layout1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
</p:sldLayout>"#;
        let layout1_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"#;
        let slide1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="3" name="Tx1Rect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="700000" y="100000"/><a:ext cx="500000" cy="500000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
      <p:style><a:fillRef idx="1"><a:schemeClr val="tx1"/></a:fillRef></p:style>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>"#;
        let slide1_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"#;

        let mut buf = Vec::new();
        {
            let cursor = Cursor::new(&mut buf);
            let mut zw = zip::ZipWriter::new(cursor);
            let opts = SimpleFileOptions::default();
            let mut put = |path: &str, bytes: &[u8]| {
                zw.start_file(path, opts).unwrap();
                use std::io::Write;
                zw.write_all(bytes).unwrap();
            };
            put("ppt/presentation.xml", presentation_xml.as_bytes());
            put("ppt/_rels/presentation.xml.rels", pres_rels.as_bytes());
            put("ppt/theme/theme1.xml", theme1_xml.as_bytes());
            put("ppt/slideMasters/slideMaster1.xml", master1_xml.as_bytes());
            put(
                "ppt/slideMasters/_rels/slideMaster1.xml.rels",
                master1_rels.as_bytes(),
            );
            put("ppt/slideLayouts/slideLayout1.xml", layout1_xml.as_bytes());
            put(
                "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
                layout1_rels.as_bytes(),
            );
            put("ppt/slides/slide1.xml", slide1_xml.as_bytes());
            put("ppt/slides/_rels/slide1.xml.rels", slide1_rels.as_bytes());
            zw.finish().unwrap();
        }
        buf
    }

    /// Single-master deck whose slide1 carries a `<p:clrMapOvr>` with the given
    /// inner element (`<a:overrideClrMapping .../>` or `<a:masterClrMapping/>`).
    /// The master keeps the DEFAULT clrMap (tx1→dk1). slide1 has the Tx1Rect
    /// (`<a:schemeClr val="tx1">`), so the override's tx1→slot remap is directly
    /// observable. Theme hex: dk1=#222222, lt1=#FAFAFA, accent1=#72A376.
    ///
    /// `layout_clr_map_ovr_inner` optionally injects a `<p:clrMapOvr>` on the
    /// LAYOUT (CT_SlideLayout: right after `</p:cSld>`, §20.1.6 / pml.xsd) so the
    /// slide↔layout override precedence can be exercised.
    fn build_clr_map_ovr_pptx(clr_map_ovr_inner: &str) -> Vec<u8> {
        build_clr_map_ovr_pptx_with_layout(clr_map_ovr_inner, None)
    }

    fn build_clr_map_ovr_pptx_with_layout(
        clr_map_ovr_inner: &str,
        layout_clr_map_ovr_inner: Option<&str>,
    ) -> Vec<u8> {
        use zip::write::SimpleFileOptions;

        let presentation_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMasterA"/></p:sldMasterIdLst>
  <p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>"#;
        let pres_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMasterA" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rIdThemeA" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>"#;
        let theme1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">
  <a:themeElements><a:clrScheme name="C">
    <a:dk1><a:srgbClr val="222222"/></a:dk1><a:lt1><a:srgbClr val="FAFAFA"/></a:lt1>
    <a:dk2><a:srgbClr val="111111"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
    <a:accent1><a:srgbClr val="72A376"/></a:accent1><a:accent2><a:srgbClr val="00FF00"/></a:accent2>
    <a:accent3><a:srgbClr val="0000FF"/></a:accent3><a:accent4><a:srgbClr val="FFFF00"/></a:accent4>
    <a:accent5><a:srgbClr val="FF00FF"/></a:accent5><a:accent6><a:srgbClr val="00FFFF"/></a:accent6>
    <a:hlink><a:srgbClr val="0000EE"/></a:hlink><a:folHlink><a:srgbClr val="551A8B"/></a:folHlink>
  </a:clrScheme>
  <a:fontScheme name="F"><a:majorFont><a:latin typeface="Arial"/></a:majorFont>
    <a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>
  <a:fmtScheme name="S"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>
  </a:themeElements>
</a:theme>"#;
        // Master keeps the DEFAULT clrMap (tx1→dk1) so the override is the ONLY
        // thing that can remap tx1; the assertion is unambiguous.
        let master1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483650" r:id="rIdLayout"/></p:sldLayoutIdLst>
</p:sldMaster>"#;
        let master1_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>"#;
        // CT_SlideLayout: <p:clrMapOvr> comes right after </p:cSld> (pml.xsd).
        let layout_clr_map_ovr = layout_clr_map_ovr_inner
            .map(|inner| format!("<p:clrMapOvr>{inner}</p:clrMapOvr>"))
            .unwrap_or_default();
        let layout1_xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
  {layout_clr_map_ovr}
</p:sldLayout>"#
        );
        let layout1_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"#;
        // CT_Slide: <p:clrMapOvr> comes right after </p:cSld> (ECMA-376 §19.3.1.7).
        // An empty `clr_map_ovr_inner` means "no <p:clrMapOvr> on the slide at all".
        let slide_clr_map_ovr = if clr_map_ovr_inner.is_empty() {
            String::new()
        } else {
            format!("<p:clrMapOvr>{clr_map_ovr_inner}</p:clrMapOvr>")
        };
        let slide1_xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="3" name="Tx1Rect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="700000" y="100000"/><a:ext cx="500000" cy="500000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
      <p:style><a:fillRef idx="1"><a:schemeClr val="tx1"/></a:fillRef></p:style>
    </p:sp>
  </p:spTree></p:cSld>
  {slide_clr_map_ovr}
</p:sld>"#
        );
        let slide1_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"#;

        let mut buf = Vec::new();
        {
            let cursor = Cursor::new(&mut buf);
            let mut zw = zip::ZipWriter::new(cursor);
            let opts = SimpleFileOptions::default();
            let mut put = |path: &str, bytes: &[u8]| {
                zw.start_file(path, opts).unwrap();
                use std::io::Write;
                zw.write_all(bytes).unwrap();
            };
            put("ppt/presentation.xml", presentation_xml.as_bytes());
            put("ppt/_rels/presentation.xml.rels", pres_rels.as_bytes());
            put("ppt/theme/theme1.xml", theme1_xml.as_bytes());
            put("ppt/slideMasters/slideMaster1.xml", master1_xml.as_bytes());
            put(
                "ppt/slideMasters/_rels/slideMaster1.xml.rels",
                master1_rels.as_bytes(),
            );
            put("ppt/slideLayouts/slideLayout1.xml", layout1_xml.as_bytes());
            put(
                "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
                layout1_rels.as_bytes(),
            );
            put("ppt/slides/slide1.xml", slide1_xml.as_bytes());
            put("ppt/slides/_rels/slide1.xml.rels", slide1_rels.as_bytes());
            zw.finish().unwrap();
        }
        buf
    }

    /// §19.3.1.7 clrMapOvr / §20.1.6.8 overrideClrMapping: a slide whose
    /// `<p:clrMapOvr>` carries `<a:overrideClrMapping>` with bg1/tx1 swapped
    /// (bg1="dk1", tx1="lt1") must use that mapping IN PLACE OF the master's.
    /// The master keeps the default clrMap (tx1→dk1, #222222), so the override
    /// flips tx1 to lt1 (#FAFAFA). The other 10 attrs are default.
    #[test]
    fn clr_map_ovr_override_remaps_logical_scheme_names() {
        let override_inner = r#"<a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>"#;
        let data = build_clr_map_ovr_pptx(override_inner);
        let pres = parse_presentation_from_bytes(&data).expect("parse");
        assert_eq!(pres.slides.len(), 1, "expected one slide");

        // tx1 under the override → lt1 (#FAFAFA), NOT the master's dk1 (#222222).
        let tx1 = shape_fill_color_by_name(&pres.slides[0], "Tx1Rect");
        assert_eq!(
            tx1.as_deref(),
            Some("FAFAFA"),
            "overrideClrMapping (tx1=lt1) must replace the master clrMap (tx1=dk1)"
        );
    }

    /// §20.1.6.6 + Annex L.3.2.5 (FINDING 3): a LAYOUT-level `overrideClrMapping`
    /// (swap bg1/tx1) is inherited by its slides; a slide carrying an explicit
    /// `<a:masterClrMapping/>` means "no override of MY OWN" and therefore inherits
    /// the LAYOUT's override (NOT a bypass to the master's raw mapping). So the
    /// slide's tx1 shape resolves through the layout override → lt1 (#FAFAFA), not
    /// the master default tx1→dk1 (#222222).
    #[test]
    fn slide_master_clr_mapping_inherits_layout_override() {
        let layout_override = r#"<a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>"#;
        let data =
            build_clr_map_ovr_pptx_with_layout("<a:masterClrMapping/>", Some(layout_override));
        let pres = parse_presentation_from_bytes(&data).expect("parse");
        assert_eq!(pres.slides.len(), 1, "expected one slide");

        let tx1 = shape_fill_color_by_name(&pres.slides[0], "Tx1Rect");
        assert_eq!(
            tx1.as_deref(),
            Some("FAFAFA"),
            "slide masterClrMapping inherits the layout override (tx1=lt1), not the master tx1=dk1"
        );
    }

    /// §20.1.6.6 + Annex L.3.2.5 (FINDING 3): a LAYOUT-level `overrideClrMapping`
    /// is inherited by a slide that has NO `<p:clrMapOvr>` at all (the common
    /// inheritance case). Same expected result as the masterClrMapping variant.
    #[test]
    fn layout_override_inherited_by_slide_without_clr_map_ovr() {
        let layout_override = r#"<a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>"#;
        // Empty slide-level inner ⇒ the builder omits <p:clrMapOvr> on the slide.
        let data = build_clr_map_ovr_pptx_with_layout("", Some(layout_override));
        let pres = parse_presentation_from_bytes(&data).expect("parse");
        assert_eq!(pres.slides.len(), 1, "expected one slide");

        let tx1 = shape_fill_color_by_name(&pres.slides[0], "Tx1Rect");
        assert_eq!(
            tx1.as_deref(),
            Some("FAFAFA"),
            "a slide with no clrMapOvr inherits the layout override (tx1=lt1)"
        );
    }

    /// Control that makes the two FINDING 3 tests load-bearing: with NO layout
    /// override and a slide `<a:masterClrMapping/>`, tx1 must stay the master
    /// default dk1 (#222222). The ONLY difference from
    /// `slide_master_clr_mapping_inherits_layout_override` is the presence of the
    /// layout override — so that test genuinely proves layout inheritance, not a
    /// vacuous pass.
    #[test]
    fn slide_master_clr_mapping_without_layout_override_uses_master() {
        let data = build_clr_map_ovr_pptx_with_layout("<a:masterClrMapping/>", None);
        let pres = parse_presentation_from_bytes(&data).expect("parse");
        assert_eq!(pres.slides.len(), 1, "expected one slide");

        let tx1 = shape_fill_color_by_name(&pres.slides[0], "Tx1Rect");
        assert_eq!(
            tx1.as_deref(),
            Some("222222"),
            "with no layout override, masterClrMapping resolves tx1 from the master (dk1)"
        );
    }

    /// The slide's resolved background fill colour, if it is a solid fill.
    fn slide_bg_color(slide: &Slide) -> Option<String> {
        match &slide.background {
            Some(Fill::Solid { color }) => Some(color.clone()),
            _ => None,
        }
    }

    /// Single-master deck like `build_clr_map_ovr_pptx`, but the MASTER carries a
    /// `<p:bg>` whose fill is `<a:schemeClr val="bg1"/>` and the SLIDE has NO
    /// background of its own, so the slide inherits the master background through
    /// the slide→layout→master chain (§19.3.1.42). The slide carries a
    /// `<p:clrMapOvr>` with the given inner element. Theme hex: dk1=#222222,
    /// lt1=#FAFAFA. With the default clrMap bg1→lt1 ⇒ #FAFAFA; under an override
    /// that maps bg1→dk1 the inherited master background must flip to #222222.
    fn build_clr_map_ovr_master_bg_pptx(slide_clr_map_ovr_inner: &str) -> Vec<u8> {
        use zip::write::SimpleFileOptions;

        let presentation_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMasterA"/></p:sldMasterIdLst>
  <p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>"#;
        let pres_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMasterA" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rIdThemeA" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>"#;
        let theme1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">
  <a:themeElements><a:clrScheme name="C">
    <a:dk1><a:srgbClr val="222222"/></a:dk1><a:lt1><a:srgbClr val="FAFAFA"/></a:lt1>
    <a:dk2><a:srgbClr val="111111"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
    <a:accent1><a:srgbClr val="72A376"/></a:accent1><a:accent2><a:srgbClr val="00FF00"/></a:accent2>
    <a:accent3><a:srgbClr val="0000FF"/></a:accent3><a:accent4><a:srgbClr val="FFFF00"/></a:accent4>
    <a:accent5><a:srgbClr val="FF00FF"/></a:accent5><a:accent6><a:srgbClr val="00FFFF"/></a:accent6>
    <a:hlink><a:srgbClr val="0000EE"/></a:hlink><a:folHlink><a:srgbClr val="551A8B"/></a:folHlink>
  </a:clrScheme>
  <a:fontScheme name="F"><a:majorFont><a:latin typeface="Arial"/></a:majorFont>
    <a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>
  <a:fmtScheme name="S"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>
  </a:themeElements>
</a:theme>"#;
        // Master keeps the DEFAULT clrMap (bg1→lt1). Its <p:bg> uses schemeClr
        // bg1, so without an override the inherited background is lt1 (#FAFAFA).
        let master1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
    <p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483650" r:id="rIdLayout"/></p:sldLayoutIdLst>
</p:sldMaster>"#;
        let master1_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>"#;
        // Layout has NO background of its own → the slide falls through to the
        // master background.
        let layout1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
</p:sldLayout>"#;
        let layout1_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"#;
        // Slide has NO <p:bg> of its own → inherits the master background.
        let slide1_xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
  <p:clrMapOvr>{slide_clr_map_ovr_inner}</p:clrMapOvr>
</p:sld>"#
        );
        let slide1_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"#;

        let mut buf = Vec::new();
        {
            let cursor = Cursor::new(&mut buf);
            let mut zw = zip::ZipWriter::new(cursor);
            let opts = SimpleFileOptions::default();
            let mut put = |path: &str, bytes: &[u8]| {
                zw.start_file(path, opts).unwrap();
                use std::io::Write;
                zw.write_all(bytes).unwrap();
            };
            put("ppt/presentation.xml", presentation_xml.as_bytes());
            put("ppt/_rels/presentation.xml.rels", pres_rels.as_bytes());
            put("ppt/theme/theme1.xml", theme1_xml.as_bytes());
            put("ppt/slideMasters/slideMaster1.xml", master1_xml.as_bytes());
            put(
                "ppt/slideMasters/_rels/slideMaster1.xml.rels",
                master1_rels.as_bytes(),
            );
            put("ppt/slideLayouts/slideLayout1.xml", layout1_xml.as_bytes());
            put(
                "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
                layout1_rels.as_bytes(),
            );
            put("ppt/slides/slide1.xml", slide1_xml.as_bytes());
            put("ppt/slides/_rels/slide1.xml.rels", slide1_rels.as_bytes());
            zw.finish().unwrap();
        }
        buf
    }

    /// §19.3.1.7 / §20.1.6.8 (FINDING 1): a master-inherited background that uses
    /// a scheme colour (`<p:bg>` schemeClr bg1) MUST resolve through the slide's
    /// effective override mapping, not the master's frozen mapping. The slide has
    /// no own background; its `<a:overrideClrMapping>` swaps bg1→dk1, so the
    /// inherited master background must become dk1 (#222222), NOT the master
    /// default bg1→lt1 (#FAFAFA).
    #[test]
    fn clr_map_ovr_flips_master_inherited_background() {
        let override_inner = r#"<a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>"#;
        let data = build_clr_map_ovr_master_bg_pptx(override_inner);
        let pres = parse_presentation_from_bytes(&data).expect("parse");
        assert_eq!(pres.slides.len(), 1, "expected one slide");

        let bg = slide_bg_color(&pres.slides[0]);
        assert_eq!(
            bg.as_deref(),
            Some("222222"),
            "master-inherited background (schemeClr bg1) must honor the slide override (bg1=dk1)"
        );
    }

    /// Control for `clr_map_ovr_flips_master_inherited_background`: with a
    /// `<a:masterClrMapping/>` (no override of its own) the inherited master
    /// background keeps the master default bg1→lt1 (#FAFAFA).
    #[test]
    fn master_inherited_background_default_without_override() {
        let data = build_clr_map_ovr_master_bg_pptx("<a:masterClrMapping/>");
        let pres = parse_presentation_from_bytes(&data).expect("parse");
        assert_eq!(pres.slides.len(), 1, "expected one slide");

        let bg = slide_bg_color(&pres.slides[0]);
        assert_eq!(
            bg.as_deref(),
            Some("FAFAFA"),
            "without an override the master background resolves bg1→lt1"
        );
    }

    /// FINDING 2 (perf guard): `parse_clr_map_ovr` must short-circuit to `None`
    /// when the XML contains no `clrMapOvr` element (avoiding a second full parse),
    /// while still returning `Some` for an `overrideClrMapping` and `None` for an
    /// explicit `masterClrMapping`. The fast path must not change any of these
    /// observable results.
    #[test]
    fn parse_clr_map_ovr_guard_and_results() {
        let ns = r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main""#;

        // No <p:clrMapOvr> at all → None (and the guard skips the parse entirely).
        let no_ovr = format!(r#"<p:sld {ns}><p:cSld><p:spTree/></p:cSld></p:sld>"#);
        assert!(
            parse_clr_map_ovr(&no_ovr).is_none(),
            "absent clrMapOvr must yield None"
        );

        // Explicit <a:masterClrMapping/> → None (inherit).
        let master = format!(
            r#"<p:sld {ns}><p:cSld><p:spTree/></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"#
        );
        assert!(
            parse_clr_map_ovr(&master).is_none(),
            "masterClrMapping must yield None"
        );

        // <a:overrideClrMapping> → Some(map) with the parsed logical→slot attrs.
        let ovr = format!(
            r#"<p:sld {ns}><p:cSld><p:spTree/></p:cSld><p:clrMapOvr><a:overrideClrMapping bg1="dk1" tx1="lt1"/></p:clrMapOvr></p:sld>"#
        );
        let parsed = parse_clr_map_ovr(&ovr).expect("overrideClrMapping must yield Some");
        assert_eq!(parsed.get("bg1").map(String::as_str), Some("dk1"));
        assert_eq!(parsed.get("tx1").map(String::as_str), Some("lt1"));
    }

    // ── Chart axis titles + chartSpace border (parity with xlsx) ──────────
    //
    // These exercise `parse_legacy_chart` directly with inline chart XML so we
    // can assert the newly-parsed fields without a full .pptx fixture. Mirrors
    // the xlsx parser's chart.rs coverage.

    /// A clustered bar chart whose category (X) and value (Y) axes both carry a
    /// `<c:title>` with explicit run props (sz / b / solidFill), plus an
    /// explicit `<c:chartSpace><c:spPr><a:ln>` border.
    fn bar_chart_with_axis_titles_xml() -> &'static str {
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Series 1</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat><c:strRef><c:strCache>
            <c:pt idx="0"><c:v>A</c:v></c:pt>
            <c:pt idx="1"><c:v>B</c:v></c:pt>
          </c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache>
            <c:pt idx="0"><c:v>3</c:v></c:pt>
            <c:pt idx="1"><c:v>7</c:v></c:pt>
          </c:numCache></c:numRef></c:val>
        </c:ser>
      </c:barChart>
      <c:catAx>
        <c:axId val="111"/>
        <c:axPos val="b"/>
        <c:title>
          <c:tx><c:rich><a:p><a:pPr><a:defRPr sz="1000" b="1">
            <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
          </a:defRPr></a:pPr><a:r><a:t>Category Axis</a:t></a:r></a:p></c:rich></c:tx>
        </c:title>
      </c:catAx>
      <c:valAx>
        <c:axId val="222"/>
        <c:axPos val="l"/>
        <c:title>
          <c:tx><c:rich><a:p><a:pPr><a:defRPr sz="1200" b="0">
            <a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>
          </a:defRPr></a:pPr><a:r><a:t>Value Axis</a:t></a:r></a:p></c:rich></c:tx>
        </c:title>
      </c:valAx>
    </c:plotArea>
  </c:chart>
  <c:spPr>
    <a:ln w="19050"><a:solidFill><a:srgbClr val="1B4332"/></a:solidFill></a:ln>
  </c:spPr>
</c:chartSpace>"#
    }

    #[test]
    fn chart_parses_cat_and_val_axis_titles_with_props() {
        let theme = HashMap::new();
        let c = parse_legacy_chart(bar_chart_with_axis_titles_xml(), &theme)
            .expect("legacy chart should parse");
        let c = &c.chart;

        assert_eq!(c.cat_axis_title.as_deref(), Some("Category Axis"));
        assert_eq!(c.cat_axis_title_font_size_hpt, Some(1000));
        assert_eq!(c.cat_axis_title_font_bold, Some(true));
        assert_eq!(c.cat_axis_title_font_color.as_deref(), Some("FF0000"));

        assert_eq!(c.val_axis_title.as_deref(), Some("Value Axis"));
        assert_eq!(c.val_axis_title_font_size_hpt, Some(1200));
        assert_eq!(c.val_axis_title_font_bold, Some(false));
        assert_eq!(c.val_axis_title_font_color.as_deref(), Some("00FF00"));
    }

    #[test]
    fn chart_parses_explicit_chartspace_border() {
        let theme = HashMap::new();
        let c = parse_legacy_chart(bar_chart_with_axis_titles_xml(), &theme)
            .expect("legacy chart should parse");
        let c = &c.chart;

        assert_eq!(c.chart_border_color.as_deref(), Some("1B4332"));
        assert_eq!(c.chart_border_width_emu, Some(19050));
    }

    #[test]
    fn chart_border_nofill_yields_no_color() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
  <c:spPr>
    <a:ln w="12700"><a:noFill/></a:ln>
  </c:spPr>
</c:chartSpace>"#;
        let theme = HashMap::new();
        let c = parse_legacy_chart(xml, &theme).expect("legacy chart should parse");
        // noFill explicitly turns the border OFF → no color, even though @w is set.
        assert_eq!(c.chart.chart_border_color, None);
    }

    #[test]
    fn chart_without_axis_titles_leaves_them_none() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:axPos val="b"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:axPos val="l"/></c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;
        let theme = HashMap::new();
        let c = parse_legacy_chart(xml, &theme).expect("legacy chart should parse");
        assert_eq!(c.chart.cat_axis_title, None);
        assert_eq!(c.chart.val_axis_title, None);
        assert_eq!(c.chart.chart_border_color, None);
        assert_eq!(c.chart.chart_border_width_emu, None);
    }

    /// A combo chart: `<c:barChart>` (Revenue, primary left axis) +
    /// `<c:lineChart>` (Gross margin, SECONDARY right axis). Mirrors sample-14
    /// slide-8. The line series must be tagged `series_type = "line"` and bound
    /// to the secondary axis, and the secondary `<c:valAx>` (axPos="r",
    /// crosses="max", min=0 max=100, title "Gross margin (%)") parsed into
    /// `secondary_val_axis` — while the primary axis fields stay the Revenue
    /// axis.
    fn combo_bar_line_secondary_axis_xml() -> &'static str {
        r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/>
          <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Revenue ($M)</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat><c:strRef><c:strCache>
            <c:pt idx="0"><c:v>FY22</c:v></c:pt>
            <c:pt idx="1"><c:v>FY23</c:v></c:pt>
          </c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache>
            <c:pt idx="0"><c:v>18.9</c:v></c:pt>
            <c:pt idx="1"><c:v>26.5</c:v></c:pt>
          </c:numCache></c:numRef></c:val>
        </c:ser>
        <c:axId val="100"/>
        <c:axId val="200"/>
      </c:barChart>
      <c:lineChart>
        <c:ser>
          <c:idx val="1"/>
          <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Gross margin (%)</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat><c:strRef><c:strCache>
            <c:pt idx="0"><c:v>FY22</c:v></c:pt>
            <c:pt idx="1"><c:v>FY23</c:v></c:pt>
          </c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache>
            <c:pt idx="0"><c:v>68</c:v></c:pt>
            <c:pt idx="1"><c:v>71</c:v></c:pt>
          </c:numCache></c:numRef></c:val>
        </c:ser>
        <c:axId val="300"/>
        <c:axId val="400"/>
      </c:lineChart>
      <c:catAx><c:axId val="100"/><c:axPos val="b"/></c:catAx>
      <c:valAx>
        <c:axId val="200"/>
        <c:axPos val="l"/>
        <c:crosses val="autoZero"/>
        <c:title><c:tx><c:rich><a:p><a:r><a:t>Revenue ($M)</a:t></a:r></a:p></c:rich></c:tx></c:title>
      </c:valAx>
      <c:valAx>
        <c:axId val="400"/>
        <c:scaling><c:max val="100"/><c:min val="0"/></c:scaling>
        <c:axPos val="r"/>
        <c:crosses val="max"/>
        <c:title><c:tx><c:rich><a:p><a:r><a:t>Gross margin (%)</a:t></a:r></a:p></c:rich></c:tx></c:title>
      </c:valAx>
      <c:catAx><c:axId val="300"/><c:delete val="1"/><c:axPos val="b"/></c:catAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#
    }

    #[test]
    fn combo_chart_tags_line_series_and_secondary_axis() {
        let theme = HashMap::new();
        let c = parse_legacy_chart(combo_bar_line_secondary_axis_xml(), &theme)
            .expect("combo chart should parse");
        let c = &c.chart;

        // Primary type is bar (bar group wins).
        assert_eq!(c.chart_type, "clusteredBar");
        assert_eq!(c.series.len(), 2, "both bar and line series parsed");

        // Bar series: primary axis. `series_type` now carries the group type
        // ("bar"); the renderer treats any non-"line" type as a bar (identical
        // rendering to the old `None`).
        assert_eq!(c.series[0].name, "Revenue ($M)");
        assert_eq!(c.series[0].series_type.as_deref(), Some("bar"));
        assert_eq!(c.series[0].use_secondary_axis, None);

        // Line series: tagged "line" + bound to the secondary axis.
        assert_eq!(c.series[1].name, "Gross margin (%)");
        assert_eq!(c.series[1].series_type.as_deref(), Some("line"));
        assert_eq!(c.series[1].use_secondary_axis, Some(true));

        // Primary value-axis fields stay the Revenue (left) axis.
        assert_eq!(c.val_axis_title.as_deref(), Some("Revenue ($M)"));

        // Secondary axis parsed from the right-hand valAx.
        let sec = c
            .secondary_val_axis
            .as_ref()
            .expect("secondary value axis present");
        assert_eq!(sec.min, Some(0.0));
        assert_eq!(sec.max, Some(100.0));
        assert_eq!(sec.title.as_deref(), Some("Gross margin (%)"));
    }

    #[test]
    fn single_axis_chart_has_no_secondary() {
        let theme = HashMap::new();
        let c = parse_legacy_chart(bar_chart_with_axis_titles_xml(), &theme)
            .expect("legacy chart should parse");
        assert!(c.chart.secondary_val_axis.is_none());
        // `series_type` now carries the group type ("bar") for every series.
        assert_eq!(c.chart.series[0].series_type.as_deref(), Some("bar"));
        assert_eq!(c.chart.series[0].use_secondary_axis, None);
    }

    #[test]
    fn scatter_bottom_valax_title_maps_to_cat_axis() {
        // Scatter charts have TWO <c:valAx> and no <c:catAx>. The bottom one
        // (axPos="b") is the horizontal axis → its title is the cat-axis title;
        // the left one (axPos="l") is the value-axis title. Same disambiguation
        // as the xlsx parser.
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:scatterChart>
        <c:scatterStyle val="lineMarker"/>
        <c:ser>
          <c:idx val="0"/>
          <c:xVal><c:numRef><c:numCache>
            <c:pt idx="0"><c:v>1</c:v></c:pt>
            <c:pt idx="1"><c:v>2</c:v></c:pt>
          </c:numCache></c:numRef></c:xVal>
          <c:yVal><c:numRef><c:numCache>
            <c:pt idx="0"><c:v>10</c:v></c:pt>
            <c:pt idx="1"><c:v>20</c:v></c:pt>
          </c:numCache></c:numRef></c:yVal>
        </c:ser>
      </c:scatterChart>
      <c:valAx>
        <c:axId val="100"/>
        <c:axPos val="b"/>
        <c:title><c:tx><c:rich><a:p><a:r><a:t>X Bottom</a:t></a:r></a:p></c:rich></c:tx></c:title>
      </c:valAx>
      <c:valAx>
        <c:axId val="200"/>
        <c:axPos val="l"/>
        <c:title><c:tx><c:rich><a:p><a:r><a:t>Y Left</a:t></a:r></a:p></c:rich></c:tx></c:title>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;
        let theme = HashMap::new();
        let c = parse_legacy_chart(xml, &theme).expect("scatter chart should parse");
        assert_eq!(c.chart.chart_type, "scatter");
        // Bottom valAx → X → cat-axis title.
        assert_eq!(c.chart.cat_axis_title.as_deref(), Some("X Bottom"));
        // Left valAx → Y → val-axis title.
        assert_eq!(c.chart.val_axis_title.as_deref(), Some("Y Left"));
    }

    #[test]
    fn chart_parses_axis_tick_label_bold_flags() {
        // The bold flags for tick labels (title bold + cat/val tick-label bold)
        // are parsed from `<c:title>...defRPr@b` and `<c:txPr>...defRPr@b`.
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr b="1"/></a:pPr>
      <a:r><a:t>My Chart</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
      </c:barChart>
      <c:catAx>
        <c:axId val="1"/><c:axPos val="b"/>
        <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr b="1"/></a:pPr><a:endParaRPr/></a:p></c:txPr>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/><c:axPos val="l"/>
        <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr b="0"/></a:pPr><a:endParaRPr/></a:p></c:txPr>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;
        let theme = HashMap::new();
        let c = parse_legacy_chart(xml, &theme).expect("legacy chart should parse");
        assert_eq!(c.chart.title_font_bold, Some(true));
        assert_eq!(c.chart.cat_axis_font_bold, Some(true));
        assert_eq!(c.chart.val_axis_font_bold, Some(false));
    }

    /// Regression for the `PathCmd::ArcTo` serde naming bug: the enum-level
    /// `#[serde(tag = "cmd", rename_all = "camelCase")]` renames only the variant
    /// tag, not the struct-variant fields, so `st_ang`/`sw_ang` serialized in
    /// snake_case. The TS `PathCmd` (core/src/types/common.ts) reads `stAng`/
    /// `swAng`, so the angles came back `undefined` → `NaN` coordinates and the
    /// arc (plus everything after it) vanished. A non-degenerate arc (positive
    /// `wR`/`hR`) is essential: a degenerate arc short-circuits before the
    /// angles are read, which is why the original arrow sample (degenerate arcs
    /// only) never surfaced this.
    #[test]
    fn arcto_serializes_angle_fields_as_camel_case() {
        // 90° arc: swAng = 90 * 60000 = 5400000 in OOXML 60000ths of a degree.
        let xml = r#"<custGeom xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
  <pathLst>
    <path w="100" h="100">
      <moveTo><pt x="100" y="50"/></moveTo>
      <arcTo wR="50" hR="50" stAng="0" swAng="5400000"/>
    </path>
  </pathLst>
</custGeom>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let subpaths = parse_cust_geom(doc.root_element(), 100.0, 100.0);
        let json = serde_json::to_string(&subpaths).expect("custGeom should serialize");

        // The two camelCase keys the TS renderer reads must be present…
        assert!(
            json.contains("\"stAng\""),
            "ArcTo must serialize stAng (camelCase); got: {json}"
        );
        assert!(
            json.contains("\"swAng\""),
            "ArcTo must serialize swAng (camelCase); got: {json}"
        );
        // …and the buggy snake_case keys must be gone.
        assert!(
            !json.contains("\"st_ang\""),
            "ArcTo must not emit snake_case st_ang; got: {json}"
        );
        assert!(
            !json.contains("\"sw_ang\""),
            "ArcTo must not emit snake_case sw_ang; got: {json}"
        );
    }

    /// Full value-level round-trip through the serde tag + camelCase fields:
    /// re-deserializing the serialized JSON must reproduce the angle values,
    /// proving the rename is symmetric (Serialize + Deserialize both use the
    /// camelCase keys) and the 60000ths→degrees conversion is intact.
    #[test]
    fn arcto_round_trips_angles_through_camel_case_json() {
        let xml = r#"<custGeom xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
  <pathLst>
    <path w="200" h="100">
      <moveTo><pt x="200" y="50"/></moveTo>
      <arcTo wR="100" hR="50" stAng="2700000" swAng="-5400000"/>
    </path>
  </pathLst>
</custGeom>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let subpaths = parse_cust_geom(doc.root_element(), 200.0, 100.0);
        let json = serde_json::to_string(&subpaths).unwrap();
        let back: Vec<Vec<PathCmd>> =
            serde_json::from_str(&json).expect("camelCase JSON must deserialize back");
        let arc = back[0]
            .iter()
            .find(|c| matches!(c, PathCmd::ArcTo { .. }))
            .expect("arc command should be present");
        match arc {
            PathCmd::ArcTo {
                wr,
                hr,
                st_ang,
                sw_ang,
            } => {
                // wR/hR normalised by path w/h; angles converted from 60000ths.
                assert!((wr - 0.5).abs() < 1e-9, "wr = {wr}"); // 100/200
                assert!((hr - 0.5).abs() < 1e-9, "hr = {hr}"); // 50/100
                assert!((st_ang - 45.0).abs() < 1e-9, "st_ang = {st_ang}"); // 2700000/60000
                assert!((sw_ang + 90.0).abs() < 1e-9, "sw_ang = {sw_ang}"); // -5400000/60000
            }
            _ => unreachable!(),
        }
    }

    /// ECMA-376 Part 1 §20.1.9.11: custom-geometry path coordinates and arc
    /// arguments may reference ordered `avLst`/`gdLst` formulas. They are
    /// evaluated in shape space and then normalized by each path's coordinate
    /// system before crossing the WASM boundary.
    #[test]
    fn custom_geometry_resolves_guides_in_path_commands() {
        let xml = r#"<custGeom xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
  <avLst><gd name="adj" fmla="val 100"/></avLst>
  <gdLst>
    <gd name="halfW" fmla="*/ w 1 2"/>
    <gd name="x" fmla="?: adj halfW 0"/>
    <gd name="radius" fmla="*/ h 1 4"/>
    <gd name="quarterTurn" fmla="val cd4"/>
    <gd name="negativeQuarter" fmla="+- 0 0 quarterTurn"/>
  </gdLst>
  <pathLst>
    <path w="200" h="100">
      <moveTo><pt x="x" y="radius"/></moveTo>
      <arcTo wR="radius" hR="radius" stAng="quarterTurn" swAng="negativeQuarter"/>
    </path>
  </pathLst>
</custGeom>"#;
        let doc = roxmltree::Document::parse(xml).expect("valid custom geometry");
        let subpaths = parse_cust_geom(doc.root_element(), 400.0, 200.0);
        assert_eq!(subpaths.len(), 1);
        match &subpaths[0][0] {
            PathCmd::MoveTo { x, y } => {
                assert!((*x - 1.0).abs() < 1e-9, "x = {x}");
                assert!((*y - 0.5).abs() < 1e-9, "y = {y}");
            }
            other => panic!("expected resolved moveTo, got {other:?}"),
        }
        match &subpaths[0][1] {
            PathCmd::ArcTo {
                wr,
                hr,
                st_ang,
                sw_ang,
            } => {
                assert!((*wr - 0.25).abs() < 1e-9, "wr = {wr}");
                assert!((*hr - 0.5).abs() < 1e-9, "hr = {hr}");
                assert!((*st_ang - 90.0).abs() < 1e-9, "st_ang = {st_ang}");
                assert!((*sw_ang + 90.0).abs() < 1e-9, "sw_ang = {sw_ang}");
            }
            other => panic!("expected resolved arcTo, got {other:?}"),
        }
    }

    /// ECMA-376 Part 1 §20.1.9.15: an omitted path coordinate-system size
    /// has the schema default zero. In that case guide values remain in shape
    /// coordinates, so normalization must use the shape extents rather than an
    /// artificial one-unit path.
    #[test]
    fn custom_geometry_without_path_size_uses_shape_extents() {
        let xml = r#"<custGeom xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
  <gdLst>
    <gd name="halfW" fmla="*/ w 1 2"/>
    <gd name="halfH" fmla="*/ h 1 2"/>
  </gdLst>
  <pathLst>
    <path>
      <moveTo><pt x="halfW" y="halfH"/></moveTo>
      <lnTo><pt x="r" y="b"/></lnTo>
    </path>
  </pathLst>
</custGeom>"#;
        let doc = roxmltree::Document::parse(xml).expect("valid custom geometry");
        let subpaths = parse_cust_geom(doc.root_element(), 400.0, 200.0);

        assert!(matches!(
            subpaths[0].as_slice(),
            [
                PathCmd::MoveTo { x, y },
                PathCmd::LineTo { x: x2, y: y2 }
            ] if (*x - 0.5).abs() < 1e-9
                && (*y - 0.5).abs() < 1e-9
                && (*x2 - 1.0).abs() < 1e-9
                && (*y2 - 1.0).abs() < 1e-9
        ));
    }

    #[test]
    fn custom_geometry_preserves_quadratic_bezier() {
        let xml = r#"<custGeom><pathLst><path w="100" h="200">
          <moveTo><pt x="0" y="0"/></moveTo>
          <quadBezTo><pt x="50" y="200"/><pt x="100" y="0"/></quadBezTo>
        </path></pathLst></custGeom>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let paths = parse_cust_geom(doc.root_element(), 100.0, 200.0);
        assert!(matches!(
            paths[0][1],
            PathCmd::QuadBezTo { x1, y1, x, y }
                if (x1 - 0.5).abs() < 1e-9 && (y1 - 1.0).abs() < 1e-9
                    && (x - 1.0).abs() < 1e-9 && y.abs() < 1e-9
        ));
    }

    /// A line chart whose horizontal axis is a `<c:dateAx>` (§21.2.2.39) — the
    /// date/time-series category axis. `axis_inner` is spliced into the dateAx.
    fn date_axis_chart_xml(axis_inner: &str) -> String {
        format!(
            r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:cat><c:numRef><c:numCache>
            <c:pt idx="0"><c:v>44927</c:v></c:pt>
            <c:pt idx="1"><c:v>44958</c:v></c:pt>
          </c:numCache></c:numRef></c:cat>
          <c:val><c:numRef><c:numCache>
            <c:pt idx="0"><c:v>10</c:v></c:pt>
            <c:pt idx="1"><c:v>20</c:v></c:pt>
          </c:numCache></c:numRef></c:val>
        </c:ser>
      </c:lineChart>
      <c:dateAx>
        <c:axId val="10"/>
        <c:axPos val="b"/>
        {axis}
      </c:dateAx>
      <c:valAx><c:axId val="20"/><c:axPos val="l"/></c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#,
            axis = axis_inner,
        )
    }

    /// `<c:dateAx>` is recognized as the category axis: its `<c:numFmt>`
    /// formatCode populates `cat_axis_format_code` so serial dates get formatted.
    #[test]
    fn date_axis_format_code_populates_cat_axis_format_code() {
        let theme = HashMap::new();
        let xml = date_axis_chart_xml(r#"<c:numFmt formatCode="m/d/yyyy" sourceLinked="0"/>"#);
        let c = parse_legacy_chart(&xml, &theme).expect("dateAx chart should parse");
        assert_eq!(c.chart.cat_axis_format_code.as_deref(), Some("m/d/yyyy"));
    }

    /// A dateAx title maps to the cat-axis title (same wiring as catAx).
    #[test]
    fn date_axis_title_maps_to_cat_axis_title() {
        let theme = HashMap::new();
        let title = r#"<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr sz="1000"/></a:pPr>
            <a:r><a:t>Date</a:t></a:r></a:p></c:rich></c:tx></c:title>"#;
        let xml = date_axis_chart_xml(title);
        let c = parse_legacy_chart(&xml, &theme).expect("dateAx chart should parse");
        assert_eq!(c.chart.cat_axis_title.as_deref(), Some("Date"));
        assert_eq!(c.chart.cat_axis_title_font_size_hpt, Some(1000));
    }

    /// A deleted dateAx hides the category axis.
    #[test]
    fn date_axis_delete_hides_cat_axis() {
        let theme = HashMap::new();
        let xml = date_axis_chart_xml(r#"<c:delete val="1"/>"#);
        let c = parse_legacy_chart(&xml, &theme).expect("dateAx chart should parse");
        assert!(c.chart.cat_axis_hidden);
    }

    // ===== RB7: per-slide partial degradation =====

    /// Build a 3-slide deck; slide `broken_idx` (0-based) gets `broken_xml` as its
    /// part body (pass malformed XML or "" to simulate a corrupt / unreadable
    /// slide). The other slides carry one text shape so a successful parse is
    /// distinguishable from a placeholder.
    fn build_three_slide_deck(broken_idx: usize, broken_xml: &str) -> Vec<u8> {
        use std::io::{Cursor, Write};
        let good_slide = |n: usize| {
            // A shape with explicit geometry so it reliably materializes as an
            // element (a geometry-less non-placeholder shape can be dropped).
            format!(
                r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="T"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>slide {n}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#
            )
        };
        let slide_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>"#;
        let layout = r#"<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree/></p:cSld></p:sldLayout>"#;
        let master = r#"<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree/></p:cSld></p:sldMaster>"#;
        let theme = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="t"><a:themeElements><a:clrScheme name="c"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="000000"/></a:dk2><a:lt2><a:srgbClr val="FFFFFF"/></a:lt2><a:accent1><a:srgbClr val="000000"/></a:accent1><a:accent2><a:srgbClr val="000000"/></a:accent2><a:accent3><a:srgbClr val="000000"/></a:accent3><a:accent4><a:srgbClr val="000000"/></a:accent4><a:accent5><a:srgbClr val="000000"/></a:accent5><a:accent6><a:srgbClr val="000000"/></a:accent6><a:hlink><a:srgbClr val="000000"/></a:hlink><a:folHlink><a:srgbClr val="000000"/></a:folHlink></a:clrScheme><a:fontScheme name="f"><a:majorFont><a:latin typeface="Arial"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme><a:fmtScheme name="s"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme></a:themeElements></a:theme>"#;

        let mut entries: Vec<(String, String)> = vec![
            ("ppt/presentation.xml".into(), r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdM"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/><p:sldId id="258" r:id="rId3"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>"#.into()),
            ("ppt/_rels/presentation.xml.rels".into(), r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide3.xml"/><Relationship Id="rIdM" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/></Relationships>"#.into()),
            ("ppt/slideLayouts/slideLayout1.xml".into(), layout.into()),
            ("ppt/slideLayouts/_rels/slideLayout1.xml.rels".into(), r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>"#.into()),
            ("ppt/slideMasters/slideMaster1.xml".into(), master.into()),
            ("ppt/slideMasters/_rels/slideMaster1.xml.rels".into(), r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>"#.into()),
            ("ppt/theme/theme1.xml".into(), theme.into()),
        ];
        for i in 0..3 {
            let body = if i == broken_idx {
                broken_xml.to_owned()
            } else {
                good_slide(i + 1)
            };
            entries.push((format!("ppt/slides/slide{}.xml", i + 1), body));
            entries.push((
                format!("ppt/slides/_rels/slide{}.xml.rels", i + 1),
                slide_rels.to_owned(),
            ));
        }
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(Cursor::new(&mut buf));
            let o = zip::write::SimpleFileOptions::default();
            for (name, body) in &entries {
                w.start_file(name.as_str(), o).unwrap();
                w.write_all(body.as_bytes()).unwrap();
            }
            w.finish().unwrap();
        }
        buf
    }

    /// OPC permits a presentation relationship to point at a slide outside the
    /// conventional `ppt/slides/` folder. Every relationship part name and
    /// relative Target below is therefore rooted at the actual source part
    /// (ECMA-376 Part 2 §6.5.2.3).
    fn build_nonstandard_slide_path_deck() -> Vec<u8> {
        const PRESENTATION: &str = r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>"#;
        const PRESENTATION_RELS: &str = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="/custom/deck/slides/slide.xml"/></Relationships>"#;
        const SLIDE: &str = r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:pic><p:nvPicPr><p:cNvPr id="2" name="NestedImage"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:spTree></p:cSld></p:sld>"#;
        const SLIDE_RELS: &str = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../layouts/layout.xml"/><Relationship Id="rImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image.png"/><Relationship Id="rNotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notes.xml"/><Relationship Id="rComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment.xml"/></Relationships>"#;
        const LAYOUT: &str = r#"<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="123456"/></a:solidFill></p:bgPr></p:bg><p:spTree/></p:cSld></p:sldLayout>"#;
        const LAYOUT_RELS: &str = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../masters/master.xml"/></Relationships>"#;
        const MASTER: &str = r#"<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree/></p:cSld></p:sldMaster>"#;
        const MASTER_RELS: &str = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#;
        const NOTES: &str = r#"<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><a:p><a:r><a:t>Nested note</a:t></a:r></a:p></p:spTree></p:cSld></p:notes>"#;
        const COMMENTS: &str = r#"<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cm authorId="0"><p:text>Nested comment</p:text></p:cm></p:cmLst>"#;

        zip_with_parts(&[
            ("ppt/presentation.xml", PRESENTATION.as_bytes()),
            (
                "ppt/_rels/presentation.xml.rels",
                PRESENTATION_RELS.as_bytes(),
            ),
            ("custom/deck/slides/slide.xml", SLIDE.as_bytes()),
            (
                "custom/deck/slides/_rels/slide.xml.rels",
                SLIDE_RELS.as_bytes(),
            ),
            ("custom/deck/layouts/layout.xml", LAYOUT.as_bytes()),
            (
                "custom/deck/layouts/_rels/layout.xml.rels",
                LAYOUT_RELS.as_bytes(),
            ),
            ("custom/deck/masters/master.xml", MASTER.as_bytes()),
            (
                "custom/deck/masters/_rels/master.xml.rels",
                MASTER_RELS.as_bytes(),
            ),
            ("custom/deck/notesSlides/notes.xml", NOTES.as_bytes()),
            ("custom/deck/comments/comment.xml", COMMENTS.as_bytes()),
            ("custom/deck/media/image.png", b"not-a-png"),
        ])
    }

    #[test]
    fn actual_slide_part_path_scopes_rels_and_relative_dependency_targets() {
        let presentation = parse_presentation_from_bytes(&build_nonstandard_slide_path_deck())
            .expect("nonstandard but valid OPC part layout parses");
        let slide = &presentation.slides[0];
        assert_eq!(
            slide.part_name.as_deref(),
            Some("custom/deck/slides/slide.xml")
        );
        assert!(slide.parse_error.is_none(), "{:?}", slide.parse_error);
        assert!(matches!(
            slide.background,
            Some(Fill::Solid { ref color }) if color == "123456"
        ));
        assert_eq!(slide.notes.as_deref(), Some("Nested note"));
        assert_eq!(slide.comments.len(), 1);
        assert_eq!(slide.comments[0].text, "Nested comment");
        let picture = slide.elements.iter().find_map(|element| match element {
            SlideElement::Picture(picture) => Some(picture),
            _ => None,
        });
        assert_eq!(
            picture.expect("slide picture").image_path,
            "custom/deck/media/image.png"
        );
    }

    #[test]
    fn no_master_fallback_never_reads_a_relationship_part_for_an_empty_source() {
        const PRESENTATION: &str = r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>"#;
        const PRESENTATION_RELS: &str = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide.xml"/></Relationships>"#;
        const SLIDE: &str = r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree/></p:cSld></p:sld>"#;
        const EMPTY_SLIDE_RELS: &str = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#;
        const MISLEADING_ROOT_RELS: &str = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="/poison.xml"/></Relationships>"#;
        let poison = vec![b'x'; 1_025];
        let data = zip_with_parts(&[
            ("ppt/presentation.xml", PRESENTATION.as_bytes()),
            (
                "ppt/_rels/presentation.xml.rels",
                PRESENTATION_RELS.as_bytes(),
            ),
            ("ppt/slides/slide.xml", SLIDE.as_bytes()),
            (
                "ppt/slides/_rels/slide.xml.rels",
                EMPTY_SLIDE_RELS.as_bytes(),
            ),
            ("_rels/.rels", MISLEADING_ROOT_RELS.as_bytes()),
            ("poison.xml", &poison),
        ]);
        let _limits = InternalLimitsOverride::set(PptxInternalLimits {
            shared_dependency_xml_bytes: 1_024,
            ..PptxInternalLimits::default()
        });

        let presentation = parse_presentation_from_bytes(&data)
            .expect("empty master path must not observe the fabricated rels or poison target");
        assert_eq!(presentation.slides.len(), 1);
        assert!(presentation.slides[0].parse_error.is_none());
        assert!(presentation.default_text_color.is_none());
    }

    #[test]
    fn pptx_dom_parser_has_an_explicit_local_defense_in_depth_node_cap() {
        let _limits = InternalLimitsOverride::set(PptxInternalLimits {
            xml_dom_complexity: 3,
            ..PptxInternalLimits::default()
        });
        assert!(parse_preflighted_pptx_xml("<r><a/><b/></r>").is_err());
    }

    #[test]
    fn slide_cursor_random_access_credit_replay_ack_and_fixed_oracle() {
        let data = build_three_slide_deck(usize::MAX, "");
        let legacy_data = data.clone();
        let mut archive = PptxArchive::new(data, None, None, None).unwrap();

        let bootstrap: serde_json::Value =
            serde_json::from_slice(&archive.presentation_bootstrap().unwrap()).unwrap();
        assert_eq!(bootstrap["slideCount"], 3);
        assert_eq!(bootstrap["slideWidth"], 12_192_000);
        assert_eq!(bootstrap["slides"][2]["partName"], "ppt/slides/slide3.xml");
        assert!(bootstrap["slides"][0].get("hidden").is_none());
        assert!(bootstrap["slides"][0].get("notes").is_none());

        let insufficient = archive.pull_slide_inner(2, 7, 3, 1).unwrap_err();
        assert!(
            insufficient.starts_with("OOXML_INSUFFICIENT_CREDIT:")
                && insufficient.contains("\"code\":\"ooxml-insufficient-credit\"")
                && insufficient.contains("\"offeredBytes\":1"),
            "{insufficient}"
        );
        assert!(archive.acknowledge_slide_inner(7, 3).is_err());
        assert!(archive.prepared_slide.as_ref().unwrap().bytes.is_some());
        let retained_ptr = archive
            .prepared_slide
            .as_ref()
            .unwrap()
            .bytes
            .as_ref()
            .unwrap()
            .as_ptr();
        let bytes = archive
            .pull_slide_inner(2, 7, 3, HARD_MAX_PPTX_SLIDE_JSON_BYTES as usize)
            .unwrap();
        assert_eq!(
            bytes.as_ptr(),
            retained_ptr,
            "prepared bytes move; they are not cloned"
        );
        const FIXED_SLIDE_3: &str = r#"{"index":2,"slideNumber":3,"partName":"ppt/slides/slide3.xml","background":null,"elements":[{"type":"shape","x":0,"y":0,"width":1000000,"height":1000000,"rotation":0.0,"flipH":false,"flipV":false,"geometry":"rect","fill":null,"stroke":null,"textBody":{"verticalAnchor":"t","paragraphs":[{"alignment":"l","marL":0,"marR":0,"indent":0,"spaceBefore":null,"spaceAfter":null,"spaceLine":null,"lvl":0,"bullet":{"type":"inherit"},"defFontSize":null,"defColor":null,"defBold":null,"defItalic":null,"defFontFamily":null,"tabStops":[],"eaLnBrk":true,"runs":[{"type":"text","text":"slide 3","bold":null,"italic":null,"underline":false,"strikethrough":false,"fontSize":null,"color":null,"fontFamily":null,"fieldType":null}]}],"defaultFontSize":null,"defaultBold":null,"defaultItalic":null,"lIns":91440,"rIns":91440,"tIns":45720,"bIns":45720,"wrap":"square","vert":"horz","autoFit":"none"},"defaultTextColor":null,"custGeom":null,"adj":null,"adj2":null,"adj3":null,"adj4":null,"adj5":null,"adj6":null,"adj7":null,"adj8":null,"shadow":null,"id":"2","name":"T"}],"elementSources":[{"origin":"slide"}]}"#;
        assert_eq!(bytes, FIXED_SLIDE_3.as_bytes());
        let legacy: serde_json::Value =
            serde_json::from_str(&parse_pptx_native(&legacy_data).unwrap()).unwrap();
        let fixed: serde_json::Value = serde_json::from_str(FIXED_SLIDE_3).unwrap();
        assert_eq!(legacy["slides"][2], fixed);
        let mut materializing = PptxArchive::new(legacy_data, None, None, None).unwrap();
        materializing.parse().unwrap();
        assert!(
            materializing.presentation.is_none(),
            "legacy parse must drop presentation caches after materialization"
        );
        assert!(archive.pull_slide_inner(2, 7, 3, 1024).is_err());
        assert!(archive.acknowledge_slide_inner(7, 2).is_err());
        assert!(archive.prepared_slide.is_some(), "stale ACK cannot advance");
        archive.acknowledge_slide_inner(7, 3).unwrap();

        let before = (
            archive.presentation.as_ref().unwrap().layout_cache.len(),
            archive
                .presentation
                .as_ref()
                .unwrap()
                .layout_source_cache
                .len(),
        );
        let first = archive
            .pull_slide_inner(0, 8, 3, HARD_MAX_PPTX_SLIDE_JSON_BYTES as usize)
            .unwrap();
        assert_eq!(
            serde_json::from_slice::<serde_json::Value>(&first).unwrap()["index"],
            0
        );
        archive.cancel_slide();
        let after = (
            archive.presentation.as_ref().unwrap().layout_cache.len(),
            archive
                .presentation
                .as_ref()
                .unwrap()
                .layout_source_cache
                .len(),
        );
        assert_eq!(
            after, before,
            "cancellation rolls back uncommitted cache insertions"
        );
        archive.cancel_slide();
        archive.close_presentation_session();
        archive.close_presentation_session();
    }

    #[test]
    fn slide_cursor_journal_preserves_old_cache_and_removes_only_current_insertions() {
        let mut archive =
            PptxArchive::new(build_three_slide_deck(usize::MAX, ""), None, None, None).unwrap();
        archive.presentation_bootstrap().unwrap();
        let shared = archive.presentation.as_mut().unwrap();
        shared
            .layout_cache
            .insert("pre-existing".to_string(), ParsedLayout::default());
        assert!(!shared
            .layout_cache
            .contains_key("ppt/slideLayouts/slideLayout1.xml"));

        archive
            .pull_slide_inner(0, 1, 1, HARD_MAX_PPTX_SLIDE_JSON_BYTES as usize)
            .unwrap();
        archive.cancel_slide();
        let shared = archive.presentation.as_ref().unwrap();
        assert!(shared.layout_cache.contains_key("pre-existing"));
        assert!(!shared
            .layout_cache
            .contains_key("ppt/slideLayouts/slideLayout1.xml"));
        assert!(!shared
            .layout_source_cache
            .contains_key("ppt/slideLayouts/slideLayout1.xml"));
    }

    #[test]
    fn bootstrap_projection_and_descriptor_limits_accept_exact_and_reject_plus_one() {
        let data = build_three_slide_deck(usize::MAX, "");
        let mut baseline = PptxArchive::new(data.clone(), None, None, None).unwrap();
        baseline.presentation_bootstrap_inner().unwrap();
        let shared = baseline.presentation.as_ref().unwrap();
        assert!(
            shared.master_cache.is_empty(),
            "bootstrap stays master-lazy"
        );
        assert!(shared.no_master_bundle.is_none());
        let exact_projection = measure_json(&(
            shared.slide_width,
            shared.slide_height,
            &shared.slide_descriptors,
            &shared.pres_rels,
            &shared.embedded_fonts,
            &shared.theme,
            &shared.pres_master_path,
            &shared.comment_authors_path,
            &shared.modern_comment_authors_path,
        ))
        .unwrap()
        .json_bytes;

        {
            let _limits = InternalLimitsOverride::set(PptxInternalLimits {
                bootstrap_slides: 3,
                bootstrap_projection_bytes: exact_projection,
                ..PptxInternalLimits::default()
            });
            let mut exact = PptxArchive::new(data.clone(), None, None, None).unwrap();
            exact.presentation_bootstrap_inner().unwrap();
        }
        {
            let _limits = InternalLimitsOverride::set(PptxInternalLimits {
                bootstrap_slides: 2,
                bootstrap_projection_bytes: exact_projection,
                ..PptxInternalLimits::default()
            });
            let mut over = PptxArchive::new(data.clone(), None, None, None).unwrap();
            let error = over.ensure_presentation().unwrap_err();
            assert!(error.contains("pptx-bootstrap"), "{error}");
            assert!(error.contains(r#""metric":"slides""#), "{error}");
        }
        {
            let _limits = InternalLimitsOverride::set(PptxInternalLimits {
                bootstrap_slides: 3,
                bootstrap_projection_bytes: exact_projection - 1,
                ..PptxInternalLimits::default()
            });
            let mut over = PptxArchive::new(data, None, None, None).unwrap();
            let error = over.ensure_presentation().unwrap_err();
            assert!(error.contains("pptx-bootstrap"), "{error}");
            assert!(error.contains("projected-bytes"), "{error}");
        }
    }

    #[test]
    fn bootstrap_json_limit_bounds_repeated_relationship_target_amplification() {
        let long_target = format!("custom/{}/slide.xml", "a".repeat(8 * 1024));
        let data = rewrite_deck_xml(
            build_three_slide_deck(usize::MAX, ""),
            "ppt/presentation.xml",
            |xml| xml.replace("rId2", "rId1").replace("rId3", "rId1"),
        );
        let data = rewrite_deck_xml(data, "ppt/_rels/presentation.xml.rels", |xml| {
            xml.replacen("slides/slide1.xml", &long_target, 1)
        });

        let mut baseline = PptxArchive::new(data.clone(), None, None, None).unwrap();
        let exact_bytes = baseline.presentation_bootstrap().unwrap().len() as u64;

        BOOTSTRAP_OUTPUT_SLIDES_RETAINED.with(|count| count.set(0));
        {
            let _limits = InternalLimitsOverride::set(PptxInternalLimits {
                bootstrap_json_bytes: exact_bytes,
                ..PptxInternalLimits::default()
            });
            let mut exact = PptxArchive::new(data.clone(), None, None, None).unwrap();
            let bytes = exact
                .presentation_bootstrap()
                .expect("the exact bootstrap JSON ceiling is inclusive");
            assert_eq!(bytes.len() as u64, exact_bytes);
            assert_eq!(
                BOOTSTRAP_OUTPUT_SLIDES_RETAINED.with(std::cell::Cell::get),
                3
            );
        }

        BOOTSTRAP_OUTPUT_SLIDES_RETAINED.with(|count| count.set(0));
        let _limits = InternalLimitsOverride::set(PptxInternalLimits {
            bootstrap_json_bytes: exact_bytes - 1,
            ..PptxInternalLimits::default()
        });
        let mut over = PptxArchive::new(data, None, None, None).unwrap();
        let error = over.presentation_bootstrap_inner().unwrap_err();
        assert!(error.starts_with("OOXML_RESOURCE_LIMIT:"), "{error}");
        assert!(error.contains("pptx-bootstrap-json"), "{error}");
        assert!(error.contains(r#""metric":"bytes""#), "{error}");
        assert_eq!(
            BOOTSTRAP_OUTPUT_SLIDES_RETAINED.with(std::cell::Cell::get),
            2,
            "the overflowing candidate is rejected before its owned path is retained"
        );
    }

    #[test]
    fn shared_cache_accounting_is_transactional_and_projection_cap_accumulates() {
        let data = build_three_slide_deck(usize::MAX, "");
        let mut baseline = PptxArchive::new(data.clone(), None, None, None).unwrap();
        baseline.presentation_bootstrap().unwrap();
        baseline
            .pull_slide_inner(0, 1, 1, HARD_MAX_PPTX_SLIDE_JSON_BYTES as usize)
            .unwrap();
        let journal = baseline
            .prepared_slide
            .as_ref()
            .and_then(|prepared| prepared.journal.as_ref())
            .unwrap();
        let exact_entries = journal.projected_entries;
        let exact_bytes = journal.projected_bytes;
        let baseline_shared = baseline.presentation.as_ref().unwrap();
        let max_dependency_projection = baseline_shared
            .master_cache
            .iter()
            .map(|(key, value)| {
                measure_json(&(Some(key.as_str()), value))
                    .unwrap()
                    .json_bytes
            })
            .chain(baseline_shared.layout_cache.iter().map(|(key, value)| {
                measure_json(&(Some(key.as_str()), value))
                    .unwrap()
                    .json_bytes
            }))
            .chain(
                baseline_shared
                    .layout_source_cache
                    .iter()
                    .map(|(key, value)| {
                        measure_json(&(Some(key.as_str()), value.as_ref()))
                            .unwrap()
                            .json_bytes
                    }),
            )
            .max()
            .unwrap();
        assert!(
            exact_entries >= 3,
            "layout source, master and layout are distinct"
        );
        assert_eq!(
            baseline.presentation.as_ref().unwrap().cache_usage.entries,
            0
        );
        assert_eq!(
            baseline
                .presentation
                .as_ref()
                .unwrap()
                .cache_usage
                .projected_bytes,
            0
        );
        baseline.cancel_slide();
        let shared = baseline.presentation.as_ref().unwrap();
        assert_eq!(shared.cache_usage.entries, 0);
        assert!(shared.master_cache.is_empty());
        assert!(shared.layout_cache.is_empty());
        assert!(shared.layout_source_cache.is_empty());

        {
            let _limits = InternalLimitsOverride::set(PptxInternalLimits {
                shared_cache_entries: exact_entries,
                shared_cache_projection_bytes: exact_bytes,
                shared_dependency_projection_bytes: max_dependency_projection,
                ..PptxInternalLimits::default()
            });
            let mut exact = PptxArchive::new(data.clone(), None, None, None).unwrap();
            exact.presentation_bootstrap().unwrap();
            exact
                .pull_slide_inner(0, 2, 1, HARD_MAX_PPTX_SLIDE_JSON_BYTES as usize)
                .unwrap();
            assert_eq!(exact.presentation.as_ref().unwrap().cache_usage.entries, 0);
            exact.acknowledge_slide_inner(2, 1).unwrap();
            let committed = &exact.presentation.as_ref().unwrap().cache_usage;
            assert_eq!(committed.entries, exact_entries);
            assert_eq!(committed.projected_bytes, exact_bytes);
            exact
                .pull_slide_inner(1, 3, 1, HARD_MAX_PPTX_SLIDE_JSON_BYTES as usize)
                .unwrap();
            exact.acknowledge_slide_inner(3, 1).unwrap();
            assert_eq!(
                exact.presentation.as_ref().unwrap().cache_usage.entries,
                exact_entries,
                "a second slide sharing the inheritance chain adds no cache state"
            );
        }
        {
            let _limits = InternalLimitsOverride::set(PptxInternalLimits {
                shared_cache_entries: exact_entries,
                shared_cache_projection_bytes: exact_bytes - 1,
                shared_dependency_projection_bytes: max_dependency_projection,
                ..PptxInternalLimits::default()
            });
            let mut over = PptxArchive::new(data, None, None, None).unwrap();
            over.presentation_bootstrap().unwrap();
            let error = over
                .pull_slide_inner(0, 4, 1, HARD_MAX_PPTX_SLIDE_JSON_BYTES as usize)
                .unwrap_err();
            assert!(error.contains("pptx-shared-cache"), "{error}");
            assert!(error.contains("projected-bytes"), "{error}");
            assert!(over.prepared_slide.is_none());
        }
        {
            let _limits = InternalLimitsOverride::set(PptxInternalLimits {
                shared_cache_entries: exact_entries,
                shared_cache_projection_bytes: exact_bytes,
                shared_dependency_projection_bytes: max_dependency_projection - 1,
                ..PptxInternalLimits::default()
            });
            let mut over =
                PptxArchive::new(build_three_slide_deck(usize::MAX, ""), None, None, None).unwrap();
            over.presentation_bootstrap().unwrap();
            let error = over
                .pull_slide_inner(0, 5, 1, HARD_MAX_PPTX_SLIDE_JSON_BYTES as usize)
                .unwrap_err();
            assert!(error.contains("pptx-shared-dependency"), "{error}");
            assert!(error.contains("projected-bytes"), "{error}");
        }
    }

    #[test]
    fn full_materialization_projection_is_cumulative_but_cursor_remains_one_unit() {
        let data = build_three_slide_deck(usize::MAX, "");
        let baseline = parse_presentation_from_bytes(&data).unwrap();
        let exact = baseline
            .slides
            .iter()
            .try_fold(0u64, |total, slide| {
                total.checked_add(measure_json(slide).unwrap().json_bytes)
            })
            .unwrap();

        {
            let _limits = InternalLimitsOverride::set(PptxInternalLimits {
                materialized_slide_json_bytes: exact,
                ..PptxInternalLimits::default()
            });
            parse_presentation_from_bytes(&data).unwrap();
        }
        {
            let _limits = InternalLimitsOverride::set(PptxInternalLimits {
                materialized_slide_json_bytes: exact - 1,
                ..PptxInternalLimits::default()
            });
            let error = parse_presentation_from_bytes(&data)
                .unwrap_err()
                .to_string();
            assert!(error.contains("pptx-materialized-slides"), "{error}");
            assert!(error.contains("projected-json-bytes"), "{error}");
        }
        {
            let _limits = InternalLimitsOverride::set(PptxInternalLimits {
                materialized_slide_json_bytes: 1,
                ..PptxInternalLimits::default()
            });
            let mut cursor = PptxArchive::new(data, None, None, None).unwrap();
            cursor
                .pull_slide_inner(0, 5, 1, HARD_MAX_PPTX_SLIDE_JSON_BYTES as usize)
                .unwrap();
            cursor.acknowledge_slide_inner(5, 1).unwrap();
        }
    }

    #[test]
    fn markdown_projection_is_sequential_equivalent_and_independent_of_full_model_limit() {
        let data = build_three_slide_deck(usize::MAX, "");
        let expected = render_presentation_md(&parse_presentation_from_bytes(&data).unwrap());
        let exact_markdown_bytes = expected.len() as u64;

        let _limits = InternalLimitsOverride::set(PptxInternalLimits {
            markdown_bytes: exact_markdown_bytes,
            materialized_slide_json_bytes: 1,
            ..PptxInternalLimits::default()
        });
        assert_eq!(
            to_markdown_native(&data).expect("the exact markdown ceiling is inclusive"),
            expected,
        );

        let mut archive = PptxArchive::new(data, None, None, None).unwrap();
        archive.presentation_bootstrap().unwrap();
        archive
            .pull_slide_inner(0, 1, 1, HARD_MAX_PPTX_SLIDE_JSON_BYTES as usize)
            .unwrap();
        archive.acknowledge_slide_inner(1, 1).unwrap();
        assert_eq!(archive.render_markdown_inner().unwrap(), expected);
    }

    #[test]
    fn markdown_projection_limit_is_typed_attributed_and_inclusive() {
        let data = build_three_slide_deck(usize::MAX, "");
        let exact = to_markdown_native(&data).unwrap().len() as u64;
        let _limits = InternalLimitsOverride::set(PptxInternalLimits {
            markdown_bytes: exact - 1,
            ..PptxInternalLimits::default()
        });

        let error = to_markdown_native(&data).unwrap_err();
        assert!(error.starts_with("OOXML_RESOURCE_LIMIT:"), "{error}");
        assert!(error.contains(r#""operation":"markdown""#), "{error}");
        assert!(error.contains(r#""stage":"serialization""#), "{error}");
        assert!(error.contains(r#""resource":"pptx-markdown""#), "{error}");
        assert!(error.contains(r#""metric":"bytes""#), "{error}");
        assert!(
            error.contains(&format!(r#""limit":{}"#, exact - 1)),
            "{error}"
        );
    }

    #[test]
    fn slide_cursor_rejects_bad_identity_index_and_poison_before_ack() {
        let data = build_three_slide_deck(usize::MAX, "");
        let mut archive = PptxArchive::new(data, None, None, None).unwrap();
        assert!(archive.pull_slide_inner(0, 0, 1, 1).is_err());
        assert!(archive.pull_slide_inner(99, 1, 1, 1024).is_err());
        assert!(archive.prepared_slide.is_none());

        archive
            .pull_slide_inner(1, 2, 1, HARD_MAX_PPTX_SLIDE_JSON_BYTES as usize)
            .unwrap();
        let reporter = archive
            .archive
            .as_ref()
            .unwrap()
            .active_operation()
            .unwrap()
            .limit_reporter()
            .unwrap();
        reporter
            .observe_hard_limit(HardResourceLimitKind::PptxSlideJsonBytes, None, 1, 2)
            .unwrap_err();
        let error = archive.acknowledge_slide_inner(2, 1).unwrap_err();
        assert!(error.starts_with("OOXML_RESOURCE_LIMIT:"));
        assert!(archive.prepared_slide.is_none());
        assert!(archive.pull_slide_inner(0, 3, 1, 1024).is_err());
    }

    #[test]
    fn pptx_slide_typed_hard_limits_accept_exact_and_reject_plus_one() {
        let slide = parse_presentation_from_bytes(&build_three_slide_deck(usize::MAX, ""))
            .unwrap()
            .slides
            .remove(0);
        let exact_json = measure_json(&slide).unwrap().json_bytes;

        let mut exact_zip = PptxZip::new(Cursor::new(empty_zip_bytes())).unwrap();
        exact_zip.begin_operation("exact").unwrap();
        let exact_reporter = exact_zip
            .active_operation()
            .unwrap()
            .limit_reporter()
            .unwrap();
        let serialized =
            serialize_slide_unit_with_limit(&slide, &exact_reporter, exact_json).unwrap();
        assert_eq!(serialized.len() as u64, exact_json);
        observe_primary_slide_xml(&exact_reporter, "ppt/slides/slide1.xml", 17, 17).unwrap();

        let mut over_zip = PptxZip::new(Cursor::new(empty_zip_bytes())).unwrap();
        over_zip.begin_operation("over").unwrap();
        let over_reporter = over_zip
            .active_operation()
            .unwrap()
            .limit_reporter()
            .unwrap();
        let json_error =
            serialize_slide_unit_with_limit(&slide, &over_reporter, exact_json - 1).unwrap_err();
        assert!(json_error.contains("pptx-slide-json"));

        let mut xml_zip = PptxZip::new(Cursor::new(empty_zip_bytes())).unwrap();
        xml_zip.begin_operation("xml-over").unwrap();
        let xml_reporter = xml_zip
            .active_operation()
            .unwrap()
            .limit_reporter()
            .unwrap();
        let xml_error =
            observe_primary_slide_xml(&xml_reporter, "ppt/slides/slide1.xml", 18, 17).unwrap_err();
        assert!(xml_error.contains("pptx-slide-xml"));
    }

    #[test]
    fn primary_slide_xml_bounded_reader_accepts_exact_and_poisons_at_plus_one_before_dom() {
        let data = build_three_slide_deck(usize::MAX, "");
        let exact = {
            let mut source = zip::ZipArchive::new(Cursor::new(data.clone())).unwrap();
            let mut entry = source.by_name("ppt/slides/slide1.xml").unwrap();
            let mut bytes = Vec::new();
            entry.read_to_end(&mut bytes).unwrap();
            bytes.len() as u64
        };

        {
            let _limit = SlideXmlLimitOverride::set(exact);
            let mut archive = PptxArchive::new(data.clone(), None, None, None).unwrap();
            archive.presentation_bootstrap().unwrap();
            archive
                .pull_slide_inner(0, 1, 1, HARD_MAX_PPTX_SLIDE_JSON_BYTES as usize)
                .expect("exact slide XML limit includes the EOF probe");
            archive.acknowledge_slide_inner(1, 1).unwrap();
        }

        let limit = exact - 1;
        let _limit = SlideXmlLimitOverride::set(limit);
        let mut archive = PptxArchive::new(data, None, None, None).unwrap();
        archive.presentation_bootstrap().unwrap();
        let parses_before = LAYOUT_MASTER_PARSE_COUNT.with(std::cell::Cell::get);
        let error = archive
            .pull_slide_inner(0, 2, 1, HARD_MAX_PPTX_SLIDE_JSON_BYTES as usize)
            .unwrap_err();
        assert!(error.starts_with("OOXML_RESOURCE_LIMIT:"), "{error}");
        assert!(error.contains("pptx-slide-xml"), "{error}");
        assert!(
            error.contains(r#""part":"ppt/slides/slide1.xml""#),
            "{error}"
        );
        assert!(error.contains(&format!(r#""limit":{limit}"#)), "{error}");
        assert!(error.contains(&format!(r#""observed":{exact}"#)), "{error}");
        assert!(archive.prepared_slide.is_none());
        assert!(archive.archive.as_ref().unwrap().assert_healthy().is_err());
        assert_eq!(
            LAYOUT_MASTER_PARSE_COUNT.with(std::cell::Cell::get),
            parses_before,
            "limit+1 must be rejected before the primary slide reaches its DOM parser"
        );
    }

    #[test]
    fn pptx_xml_dom_complexity_preflight_accepts_exact_and_poisons_plus_one_before_dom() {
        let siblings = "<p:ext/>".repeat(64);
        let slide = format!(
            r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree>{siblings}</p:spTree></p:cSld></p:sld>"#
        );
        let exact = (0..1_000)
            .find(|limit| !xml_dom_complexity_exceeds(&slide, *limit))
            .expect("fixture complexity is below the search ceiling");
        let data = build_three_slide_deck(0, &slide);

        {
            let _limits = InternalLimitsOverride::set(PptxInternalLimits {
                xml_dom_complexity: exact,
                ..PptxInternalLimits::default()
            });
            let mut archive = PptxArchive::new(data.clone(), None, None, None).unwrap();
            archive
                .pull_slide_inner(0, 1, 1, HARD_MAX_PPTX_SLIDE_JSON_BYTES as usize)
                .expect("the exact XML DOM complexity ceiling is inclusive");
            archive.acknowledge_slide_inner(1, 1).unwrap();
        }

        let _limits = InternalLimitsOverride::set(PptxInternalLimits {
            xml_dom_complexity: exact - 1,
            ..PptxInternalLimits::default()
        });
        let mut archive = PptxArchive::new(data, None, None, None).unwrap();
        archive.presentation_bootstrap().unwrap();
        let parses_before = LAYOUT_MASTER_PARSE_COUNT.with(std::cell::Cell::get);
        let error = archive
            .pull_slide_inner(0, 2, 1, HARD_MAX_PPTX_SLIDE_JSON_BYTES as usize)
            .unwrap_err();
        assert!(error.contains("xml-dom"), "{error}");
        assert!(error.contains(r#""metric":"complexity-units""#), "{error}");
        assert!(
            error.contains(r#""part":"ppt/slides/slide1.xml""#),
            "{error}"
        );
        assert_eq!(
            LAYOUT_MASTER_PARSE_COUNT.with(std::cell::Cell::get),
            parses_before,
            "complexity limit+1 is rejected before the primary slide DOM parser"
        );
    }

    #[test]
    fn generic_dependency_limit_poison_is_not_swallowed_but_malformed_xml_still_degrades() {
        let base = build_three_slide_deck(usize::MAX, "");
        let malformed = rewrite_deck_xml(base.clone(), "ppt/slides/_rels/slide1.xml.rels", |_| {
            "<Relationships><broken".to_string()
        });
        let presentation = parse_presentation_from_bytes(&malformed)
            .expect("ordinary malformed relationship XML remains compatible degradation");
        assert_eq!(presentation.slides.len(), 3);

        let oversized = rewrite_deck_xml(base, "ppt/slides/_rels/slide1.xml.rels", |xml| {
            format!("{xml}{}", " ".repeat(16 * 1024))
        });
        let exact = {
            let mut source = zip::ZipArchive::new(Cursor::new(oversized.clone())).unwrap();
            let mut entry = source.by_name("ppt/slides/_rels/slide1.xml.rels").unwrap();
            let mut bytes = Vec::new();
            entry.read_to_end(&mut bytes).unwrap();
            bytes.len() as u64
        };
        {
            let _limits = InternalLimitsOverride::set(PptxInternalLimits {
                shared_dependency_xml_bytes: exact,
                ..PptxInternalLimits::default()
            });
            let mut archive = PptxArchive::new(oversized.clone(), None, None, None).unwrap();
            archive
                .pull_slide_inner(0, 3, 1, HARD_MAX_PPTX_SLIDE_JSON_BYTES as usize)
                .expect("the exact dependency XML byte ceiling includes EOF/CRC validation");
            archive.acknowledge_slide_inner(3, 1).unwrap();
        }
        let _limits = InternalLimitsOverride::set(PptxInternalLimits {
            shared_dependency_xml_bytes: exact - 1,
            ..PptxInternalLimits::default()
        });
        let mut archive = PptxArchive::new(oversized, None, None, None).unwrap();
        archive.presentation_bootstrap().unwrap();
        let error = archive
            .pull_slide_inner(0, 4, 1, HARD_MAX_PPTX_SLIDE_JSON_BYTES as usize)
            .unwrap_err();
        assert!(error.contains("pptx-shared-dependency-xml"), "{error}");
        assert!(
            error.contains(r#""part":"ppt/slides/_rels/slide1.xml.rels""#),
            "{error}"
        );
        assert!(archive.prepared_slide.is_none());
        assert!(archive.archive.as_ref().unwrap().assert_healthy().is_err());
    }

    #[test]
    fn primary_slide_invalid_utf8_remains_a_local_placeholder() {
        let original = build_three_slide_deck(usize::MAX, "");
        let mut source = zip::ZipArchive::new(Cursor::new(original)).unwrap();
        let mut data = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut data));
            let options = zip::write::SimpleFileOptions::default();
            for index in 0..source.len() {
                let mut entry = source.by_index(index).unwrap();
                let name = entry.name().to_owned();
                let mut body = Vec::new();
                entry.read_to_end(&mut body).unwrap();
                if name == "ppt/slides/slide1.xml" {
                    body = vec![0xff];
                }
                writer.start_file(name, options).unwrap();
                writer.write_all(&body).unwrap();
            }
            writer.finish().unwrap();
        }
        let presentation = parse_presentation_from_bytes(&data).unwrap();
        assert!(presentation.slides[0]
            .parse_error
            .as_deref()
            .unwrap()
            .contains("not valid UTF-8"));
        assert!(presentation.slides[1].parse_error.is_none());
        assert!(presentation.slides[2].parse_error.is_none());
    }

    #[test]
    fn primary_slide_crc_failure_remains_a_local_placeholder() {
        let original = build_three_slide_deck(usize::MAX, "");
        let mut source = zip::ZipArchive::new(Cursor::new(original)).unwrap();
        let mut data = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut data));
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);
            for index in 0..source.len() {
                let mut entry = source.by_index(index).unwrap();
                let name = entry.name().to_owned();
                let mut body = Vec::new();
                entry.read_to_end(&mut body).unwrap();
                writer.start_file(name, options).unwrap();
                writer.write_all(&body).unwrap();
            }
            writer.finish().unwrap();
        }
        let marker = b"slide 2";
        let offset = data
            .windows(marker.len())
            .position(|window| window == marker)
            .expect("stored primary slide payload marker");
        data[offset + marker.len() - 1] ^= 1;
        let presentation = parse_presentation_from_bytes(&data).unwrap();
        assert!(presentation.slides[0].parse_error.is_none());
        assert!(presentation.slides[1]
            .parse_error
            .as_deref()
            .unwrap()
            .to_ascii_lowercase()
            .contains("crc"));
        assert!(presentation.slides[2].parse_error.is_none());
    }

    #[test]
    fn canonical_slide_json_limit_is_shared_by_legacy_and_cursor_materialization() {
        let data = build_three_slide_deck(usize::MAX, "");
        let baseline = parse_presentation_from_bytes(&data).unwrap();
        let exact = measure_json(&baseline.slides[0]).unwrap().json_bytes;

        {
            let _limit = SlideJsonLimitOverride::set(exact);
            parse_presentation_from_bytes_with_limits(&data, None, None, "parse-exact")
                .expect("the exact canonical slide JSON limit is inclusive");
        }

        let limit = exact - 1;
        let legacy_error = {
            let _limit = SlideJsonLimitOverride::set(limit);
            parse_presentation_from_bytes_with_limits(&data, None, None, "parse-over").unwrap_err()
        };
        let cursor_error = {
            let _limit = SlideJsonLimitOverride::set(limit);
            let mut archive = PptxArchive::new(data, None, None, None).unwrap();
            let error = archive
                .pull_slide_inner(0, 1, 1, HARD_MAX_PPTX_SLIDE_JSON_BYTES as usize)
                .unwrap_err();
            assert!(archive.prepared_slide.is_none());
            assert!(archive.archive.as_ref().unwrap().assert_healthy().is_err());
            error
        };
        let archive_error = {
            let _limit = SlideJsonLimitOverride::set(limit);
            PptxArchive::new(build_three_slide_deck(usize::MAX, ""), None, None, None)
                .unwrap()
                .parse_inner()
                .unwrap_err()
        };
        for error in [&legacy_error, &archive_error, &cursor_error] {
            assert!(error.starts_with("OOXML_RESOURCE_LIMIT:"), "{error}");
            assert!(error.contains("pptx-slide-json"), "{error}");
            assert!(error.contains(&format!(r#""limit":{limit}"#)), "{error}");
            assert!(error.contains(&format!(r#""observed":{exact}"#)), "{error}");
        }
    }

    #[test]
    fn corrupt_container_bootstrap_and_slide_preserve_degraded_contract() {
        let mut archive = PptxArchive::new(vec![1, 2, 3], None, None, None).unwrap();
        let bootstrap: serde_json::Value =
            serde_json::from_slice(&archive.presentation_bootstrap().unwrap()).unwrap();
        assert_eq!(bootstrap["slideCount"], 1);
        assert_eq!(bootstrap["slideWidth"], 12_192_000);
        let bytes = archive.pull_slide_inner(0, 1, 1, 1).unwrap_err();
        assert!(
            bytes.starts_with("OOXML_INSUFFICIENT_CREDIT:")
                && bytes.contains("\"code\":\"ooxml-insufficient-credit\"")
                && bytes.contains("\"offeredBytes\":1"),
            "{bytes}"
        );
        let bytes = archive.pull_slide_inner(0, 1, 1, 1024 * 1024).unwrap();
        let slide: serde_json::Value = serde_json::from_slice(&bytes).unwrap();
        assert_eq!(slide["index"], 0);
        assert!(slide["parseError"]
            .as_str()
            .unwrap()
            .contains("zip container"));
        archive.acknowledge_slide_inner(1, 1).unwrap();
    }

    /// NEUTRALIZATION: a deck whose middle slide part is unparseable still opens;
    /// the two healthy slides render and the broken one is a placeholder whose
    /// `parseError` names the offending part (`ppt/slides/slide2.xml`).
    #[test]
    fn rb7_one_broken_slide_degrades_rest_render() {
        // Malformed XML: unterminated element → roxmltree parse fails.
        let data = build_three_slide_deck(1, "<p:sld><p:cSld><p:spTree>");
        let json = parse_pptx_native(&data).expect("deck must still open with a broken slide");
        let v: serde_json::Value = serde_json::from_str(&json).unwrap();
        let slides = v["slides"].as_array().expect("slides array");
        assert_eq!(slides.len(), 3, "all three slide slots are present");

        // Slide 0 and 2 parsed normally (have their text shape, no parseError).
        for i in [0usize, 2] {
            assert!(
                slides[i]["parseError"].is_null(),
                "healthy slide {i} must carry no parseError; got {}",
                slides[i]
            );
            assert!(
                !slides[i]["elements"].as_array().unwrap().is_empty(),
                "healthy slide {i} keeps its content"
            );
        }

        // Slide 1 is the placeholder: empty elements + a part-tagged error.
        let broken = &slides[1];
        let err = broken["parseError"]
            .as_str()
            .expect("broken slide carries a parseError string");
        assert!(
            err.starts_with("ppt/slides/slide2.xml:"),
            "error must name the offending part; got {err:?}"
        );
        assert!(
            broken["elements"].as_array().unwrap().is_empty(),
            "placeholder slide has no elements"
        );
        // Index / slide number preserved so navigation stays 1:1 with the deck.
        assert_eq!(broken["index"].as_u64(), Some(1));
        assert_eq!(broken["slideNumber"].as_u64(), Some(2));
    }

    /// A slide part that is entirely MISSING from the zip (dangling rId Target)
    /// also degrades to a placeholder rather than aborting the whole deck.
    #[test]
    fn rb7_unreadable_slide_part_degrades() {
        // Build a normal deck, then rebuild the zip WITHOUT slide3.xml so its read
        // fails. Simplest: point the broken slot at empty content the read path
        // still returns, but assert the malformed-XML path already; here we cover
        // the "read failed" arm by omitting the part via a deck missing slide 3.
        let data = build_deck_missing_third_slide();
        let json = parse_pptx_native(&data).expect("deck must open with a missing slide part");
        let v: serde_json::Value = serde_json::from_str(&json).unwrap();
        let slides = v["slides"].as_array().expect("slides array");
        assert_eq!(slides.len(), 3);
        assert!(slides[0]["parseError"].is_null());
        assert!(slides[1]["parseError"].is_null());
        let err = slides[2]["parseError"]
            .as_str()
            .expect("missing slide part yields a placeholder + error");
        assert!(
            err.starts_with("ppt/slides/slide3.xml:"),
            "error names the missing part; got {err:?}"
        );
    }

    /// A 3-slide deck whose slide3.xml part is omitted from the archive entirely.
    fn build_deck_missing_third_slide() -> Vec<u8> {
        use std::io::{Cursor, Write};
        // Reuse the three-slide scaffold, then strip slide3.xml back out.
        let full = build_three_slide_deck(2, "<unused/>");
        // Re-open, copy every entry EXCEPT ppt/slides/slide3.xml, into a fresh zip.
        let mut zip = zip::ZipArchive::new(Cursor::new(full)).unwrap();
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(Cursor::new(&mut buf));
            let o = zip::write::SimpleFileOptions::default();
            for i in 0..zip.len() {
                let mut f = zip.by_index(i).unwrap();
                let name = f.name().to_owned();
                if name == "ppt/slides/slide3.xml" {
                    continue; // omit → its read fails → placeholder
                }
                use std::io::Read;
                let mut body = Vec::new();
                f.read_to_end(&mut body).unwrap();
                w.start_file(name.as_str(), o).unwrap();
                w.write_all(&body).unwrap();
            }
            w.finish().unwrap();
        }
        buf
    }

    fn rewrite_deck_xml(full: Vec<u8>, target: &str, rewrite: impl Fn(&str) -> String) -> Vec<u8> {
        use std::io::{Cursor, Read, Write};
        let mut source = zip::ZipArchive::new(Cursor::new(full)).unwrap();
        let mut output = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut output));
            let options = zip::write::SimpleFileOptions::default();
            for index in 0..source.len() {
                let mut entry = source.by_index(index).unwrap();
                let name = entry.name().to_owned();
                let mut body = Vec::new();
                entry.read_to_end(&mut body).unwrap();
                if name == target {
                    let xml = String::from_utf8(body).unwrap();
                    body = rewrite(&xml).into_bytes();
                }
                writer.start_file(name, options).unwrap();
                writer.write_all(&body).unwrap();
            }
            writer.finish().unwrap();
        }
        output
    }

    fn append_zip_part(full: Vec<u8>, path: &str, body: &str) -> Vec<u8> {
        use std::io::{Cursor, Read, Write};
        let mut source = zip::ZipArchive::new(Cursor::new(full)).unwrap();
        let mut output = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut output));
            let options = zip::write::SimpleFileOptions::default();
            for index in 0..source.len() {
                let mut entry = source.by_index(index).unwrap();
                let name = entry.name().to_owned();
                let mut bytes = Vec::new();
                entry.read_to_end(&mut bytes).unwrap();
                writer.start_file(name, options).unwrap();
                writer.write_all(&bytes).unwrap();
            }
            writer.start_file(path, options).unwrap();
            writer.write_all(body.as_bytes()).unwrap();
            writer.finish().unwrap();
        }
        output
    }

    fn build_comment_deck(
        referenced: [bool; 2],
        comment_xml: [Option<&str>; 2],
        author_xml: &str,
    ) -> Vec<u8> {
        use std::io::{Read, Write};
        let mut source = zip::ZipArchive::new(Cursor::new(build_three_slide_deck(9, "<unused/>")))
            .expect("base deck opens");
        let mut output = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut output));
            let options = zip::write::SimpleFileOptions::default();
            for index in 0..source.len() {
                let mut entry = source.by_index(index).unwrap();
                let name = entry.name().to_owned();
                let mut body = Vec::new();
                entry.read_to_end(&mut body).unwrap();
                for (slide, is_referenced) in referenced.iter().copied().enumerate() {
                    if is_referenced
                        && name == format!("ppt/slides/_rels/slide{}.xml.rels", slide + 1)
                    {
                        let xml = String::from_utf8(body).unwrap();
                        body = xml
                            .replace(
                                "</Relationships>",
                                &format!(
                                    r#"<Relationship Id="rIdComment" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment{}.xml"/></Relationships>"#,
                                    slide + 1
                                ),
                            )
                            .into_bytes();
                    }
                }
                if name == "ppt/_rels/presentation.xml.rels" {
                    let xml = String::from_utf8(body).unwrap();
                    body = xml
                        .replace(
                            "</Relationships>",
                            r#"<Relationship Id="rCommentAuthors" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors" Target="commentAuthors.xml"/></Relationships>"#,
                        )
                        .into_bytes();
                }
                writer.start_file(name, options).unwrap();
                writer.write_all(&body).unwrap();
            }
            for (slide, xml) in comment_xml.into_iter().enumerate() {
                if let Some(xml) = xml {
                    writer
                        .start_file(format!("ppt/comments/comment{}.xml", slide + 1), options)
                        .unwrap();
                    writer.write_all(xml.as_bytes()).unwrap();
                }
            }
            writer
                .start_file("ppt/commentAuthors.xml", options)
                .unwrap();
            writer.write_all(author_xml.as_bytes()).unwrap();
            writer.finish().unwrap();
        }
        output
    }

    fn build_modern_comment_deck(comment_xml: &str, author_xml: &str) -> Vec<u8> {
        use std::io::{Read, Write};
        let base = build_comment_deck(
            [false, false],
            [None, None],
            r#"<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#,
        );
        let mut source = zip::ZipArchive::new(Cursor::new(base)).expect("base deck opens");
        let mut output = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut output));
            let options = zip::write::SimpleFileOptions::default();
            for index in 0..source.len() {
                let mut entry = source.by_index(index).unwrap();
                let name = entry.name().to_owned();
                let mut body = Vec::new();
                entry.read_to_end(&mut body).unwrap();
                if name == "ppt/slides/_rels/slide1.xml.rels" {
                    body = String::from_utf8(body)
                        .unwrap()
                        .replace(
                            "</Relationships>",
                            r#"<Relationship Id="rModernComments" Type="http://schemas.microsoft.com/office/2018/10/relationships/comments" Target="../comments/modernComment1.xml"/></Relationships>"#,
                        )
                        .into_bytes();
                } else if name == "ppt/_rels/presentation.xml.rels" {
                    body = String::from_utf8(body)
                        .unwrap()
                        .replace(
                            "</Relationships>",
                            r#"<Relationship Id="rModernAuthors" Type="http://schemas.microsoft.com/office/2018/10/relationships/authors" Target="authors.xml"/></Relationships>"#,
                        )
                        .into_bytes();
                }
                writer.start_file(name, options).unwrap();
                writer.write_all(&body).unwrap();
            }
            writer
                .start_file("ppt/comments/modernComment1.xml", options)
                .unwrap();
            writer.write_all(comment_xml.as_bytes()).unwrap();
            writer.start_file("ppt/authors.xml", options).unwrap();
            writer.write_all(author_xml.as_bytes()).unwrap();
            writer.finish().unwrap();
        }
        output
    }

    fn build_typed_comment_deck(
        slide_relationships: &str,
        parts: &[(&str, &str)],
        modern_author_xml: &str,
    ) -> Vec<u8> {
        use std::io::{Read, Write};
        let base = build_comment_deck(
            [false, false],
            [None, None],
            r#"<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cmAuthor id="0" name="Classic Author"/></p:cmAuthorLst>"#,
        );
        let mut source = zip::ZipArchive::new(Cursor::new(base)).expect("base deck opens");
        let mut output = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut output));
            let options = zip::write::SimpleFileOptions::default();
            for index in 0..source.len() {
                let mut entry = source.by_index(index).unwrap();
                let name = entry.name().to_owned();
                let mut body = Vec::new();
                entry.read_to_end(&mut body).unwrap();
                if name == "ppt/slides/_rels/slide1.xml.rels" {
                    body = String::from_utf8(body)
                        .unwrap()
                        .replace(
                            "</Relationships>",
                            &format!("{slide_relationships}</Relationships>"),
                        )
                        .into_bytes();
                } else if name == "ppt/_rels/presentation.xml.rels" {
                    body = String::from_utf8(body)
                        .unwrap()
                        .replace(
                            "</Relationships>",
                            r#"<Relationship Id="rModernAuthors" Type="http://schemas.microsoft.com/office/2018/10/relationships/authors" Target="authors.xml"/></Relationships>"#,
                        )
                        .into_bytes();
                }
                writer.start_file(name, options).unwrap();
                writer.write_all(&body).unwrap();
            }
            for (path, xml) in parts {
                writer.start_file(*path, options).unwrap();
                writer.write_all(xml.as_bytes()).unwrap();
            }
            writer.start_file("ppt/authors.xml", options).unwrap();
            writer.write_all(modern_author_xml.as_bytes()).unwrap();
            writer.finish().unwrap();
        }
        output
    }

    fn build_deck_missing_second_slide_relationship() -> Vec<u8> {
        rewrite_deck_xml(
            build_three_slide_deck(9, "<unused/>"),
            "ppt/_rels/presentation.xml.rels",
            |xml| {
                xml.replace(r#"<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>"#, "")
            },
        )
    }

    #[test]
    fn missing_slide_relationship_preserves_the_slide_slot() {
        let data = build_deck_missing_second_slide_relationship();
        let presentation = parse_presentation_from_bytes(&data).expect("deck degrades per slide");
        assert_eq!(presentation.slides.len(), 3);
        assert!(presentation.slides[0].parse_error.is_none());
        assert!(presentation.slides[2].parse_error.is_none());
        let broken = &presentation.slides[1];
        assert_eq!(broken.index, 1);
        assert_eq!(broken.slide_number, 2);
        let error = broken.parse_error.as_deref().expect("placeholder error");
        assert!(
            error.starts_with("ppt/presentation.xml#sldId[1]/@r:id=rId2:"),
            "stable relation tag: {error}"
        );
    }

    #[test]
    fn missing_slide_relationship_id_preserves_the_slide_slot() {
        let data = rewrite_deck_xml(
            build_three_slide_deck(9, "<unused/>"),
            "ppt/presentation.xml",
            |xml| {
                xml.replace(
                    r#"<p:sldId id="257" r:id="rId2"/>"#,
                    r#"<p:sldId id="257"/>"#,
                )
            },
        );
        let presentation = parse_presentation_from_bytes(&data).expect("deck degrades per slide");
        assert_eq!(presentation.slides.len(), 3);
        assert!(presentation.slides[0].parse_error.is_none());
        assert!(presentation.slides[2].parse_error.is_none());
        let broken = &presentation.slides[1];
        assert_eq!(broken.index, 1);
        assert_eq!(broken.slide_number, 2);
        let error = broken.parse_error.as_deref().expect("placeholder error");
        assert!(
            error.starts_with("ppt/presentation.xml#sldId[1]/@r:id:"),
            "stable missing-attribute tag: {error}"
        );
    }

    #[test]
    fn out_of_order_producer_keeps_malformed_descriptor_index_stable() {
        let data = rewrite_deck_xml(
            build_three_slide_deck(9, "<unused/>"),
            "ppt/presentation.xml",
            |xml| {
                xml.replace(
                    r#"<p:sldId id="257" r:id="rId2"/>"#,
                    r#"<p:sldId id="257"/>"#,
                )
            },
        );
        let mut zip = PptxZip::new(Cursor::new(data)).expect("archive opens");
        zip.run_operation("descriptor-stability-test", |zip| {
            let mut shared = bootstrap_presentation(zip).map_err(|e| e.to_string())?;
            let third = produce_slide_unit(2, &mut shared, zip).map_err(|e| e.to_string())?;
            let broken = produce_slide_unit(1, &mut shared, zip).map_err(|e| e.to_string())?;
            let first = produce_slide_unit(0, &mut shared, zip).map_err(|e| e.to_string())?;
            assert_eq!((first.index, broken.index, third.index), (0, 1, 2));
            assert_eq!(broken.slide_number, 2);
            assert!(broken
                .parse_error
                .as_deref()
                .is_some_and(|e| e.starts_with("ppt/presentation.xml#sldId[1]/@r:id:")));
            Ok(())
        })
        .expect("malformed descriptor degrades in place");
    }

    fn valid_comment_xml(text: &str) -> String {
        format!(
            r#"<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cm authorId="0" idx="1" dt="2026-01-01T00:00:00Z"><p:pos x="0" y="0"/><p:text>{text}</p:text></p:cm></p:cmLst>"#
        )
    }

    #[test]
    fn legacy_comment_preserves_authored_identity_and_slide_position() {
        let comments = r#"<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cm authorId="7" idx="42" dt="2026-01-01T00:00:00Z"><p:pos x="914400" y="1828800"/><p:text>Positioned</p:text></p:cm></p:cmLst>"#;
        let authors = r#"<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cmAuthor id="7" name="Ada" initials="AL" lastIdx="42" clrIdx="3"/></p:cmAuthorLst>"#;
        let data = build_comment_deck([true, false], [Some(comments), None], authors);
        let presentation = parse_presentation_from_bytes(&data).expect("commented deck parses");
        let comment = &presentation.slides[0].comments[0];

        assert_eq!(comment.author_id, Some(7));
        assert_eq!(comment.index, Some(42));
        assert_eq!(comment.x, Some(914400));
        assert_eq!(comment.y, Some(1828800));
        assert_eq!(comment.author.as_deref(), Some("Ada"));
        assert_eq!(comment.text, "Positioned");
    }

    #[test]
    fn legacy_comment_authors_follow_the_presentation_relationship() {
        let comments = valid_comment_xml("Relationship author");
        let poison = r#"<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cmAuthor id="0" name="Poison"/></p:cmAuthorLst>"#;
        let actual = r#"<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cmAuthor id="0" name="Actual Author"/></p:cmAuthorLst>"#;
        let data = build_comment_deck([true, false], [Some(&comments), None], poison);
        let data = rewrite_deck_xml(data, "ppt/_rels/presentation.xml.rels", |xml| {
            xml.replace(
                "Target=\"commentAuthors.xml\"",
                "Target=\"review/authors.xml\"",
            )
        });
        let data = append_zip_part(data, "ppt/review/authors.xml", actual);

        let presentation = parse_presentation_from_bytes(&data).expect("commented deck parses");
        assert_eq!(
            presentation.slides[0].comments[0].author.as_deref(),
            Some("Actual Author")
        );
    }

    #[test]
    fn legacy_comment_authors_require_the_exact_internal_relationship_type() {
        let comments = valid_comment_xml("Relationship author");
        let poison = r#"<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cmAuthor id="0" name="Poison"/></p:cmAuthorLst>"#;
        let actual = r#"<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cmAuthor id="0" name="Actual Author"/></p:cmAuthorLst>"#;
        let data = build_comment_deck([true, false], [Some(&comments), None], poison);
        let data = rewrite_deck_xml(data, "ppt/_rels/presentation.xml.rels", |xml| {
            xml.replace(
                r#"<Relationship Id="rCommentAuthors" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors" Target="commentAuthors.xml"/>"#,
                r#"<Relationship Id="rExternalAuthors" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors" Target="commentAuthors.xml" TargetMode="External"/><Relationship Id="rPoisonAuthors" Type="urn:example/relationships/commentAuthors" Target="commentAuthors.xml"/><Relationship Id="rCommentAuthors" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors" Target="review/classic-authors.xml"/>"#,
            )
        });
        let data = append_zip_part(data, "ppt/review/classic-authors.xml", actual);

        let presentation = parse_presentation_from_bytes(&data).expect("commented deck parses");
        assert_eq!(
            presentation.slides[0].comments[0].author.as_deref(),
            Some("Actual Author")
        );
    }

    #[test]
    fn comment_author_relationship_ignores_external_targets() {
        let rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="external" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors" Target="https://example.invalid/authors.xml" TargetMode="External"/><Relationship Id="internal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors" Target="review/authors.xml"/></Relationships>"#;
        assert_eq!(
            find_internal_rel_target_by_types(rels, CLASSIC_COMMENT_AUTHOR_RELATIONSHIP_TYPES,)
                .as_deref(),
            Some("review/authors.xml"),
        );
    }

    #[test]
    fn strict_classic_comment_author_relationship_type_is_allowlisted() {
        let rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="strict" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/commentAuthors" Target="review/authors.xml"/></Relationships>"#;
        assert_eq!(
            find_internal_rel_target_by_types(rels, CLASSIC_COMMENT_AUTHOR_RELATIONSHIP_TYPES,)
                .as_deref(),
            Some("review/authors.xml"),
        );
    }

    #[test]
    fn modern_comment_authors_require_the_exact_internal_relationship_type() {
        let comments = r#"<p188:cmLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p188:cm id="{ROOT}" authorId="{ADA}"><p188:txBody><a:p><a:r><a:t>Modern</a:t></a:r></a:p></p188:txBody></p188:cm></p188:cmLst>"#;
        let poison = r#"<p188:authorLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main"><p188:author id="{ADA}" name="Poison"/></p188:authorLst>"#;
        let actual = r#"<p188:authorLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main"><p188:author id="{ADA}" name="Actual Author"/></p188:authorLst>"#;
        let data = build_modern_comment_deck(comments, poison);
        let data = rewrite_deck_xml(data, "ppt/_rels/presentation.xml.rels", |xml| {
            xml.replace(
                r#"<Relationship Id="rModernAuthors" Type="http://schemas.microsoft.com/office/2018/10/relationships/authors" Target="authors.xml"/>"#,
                r#"<Relationship Id="rExternalModernAuthors" Type="http://schemas.microsoft.com/office/2018/10/relationships/authors" Target="authors.xml" TargetMode="External"/><Relationship Id="rPoisonModernAuthors" Type="urn:example/relationships/authors" Target="authors.xml"/><Relationship Id="rModernAuthors" Type="http://schemas.microsoft.com/office/2018/10/relationships/authors" Target="review/modern-authors.xml"/>"#,
            )
        });
        let data = append_zip_part(data, "ppt/review/modern-authors.xml", actual);

        let presentation = parse_presentation_from_bytes(&data).expect("modern comments parse");
        assert_eq!(
            presentation.slides[0].comments[0].author.as_deref(),
            Some("Actual Author")
        );
    }

    #[test]
    fn modern_comment_preserves_thread_author_status_text_and_position() {
        let comments = r#"<p188:cmLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pc="http://schemas.microsoft.com/office/powerpoint/2013/main/command" xmlns:ac="http://schemas.microsoft.com/office/drawing/2013/main/command"><p188:cm id="{ROOT}" authorId="{ADA}" status="active" created="2026-08-24T12:00:00Z"><ac:deMkLst><pc:docMk/><pc:sldMk sldId="256"/><ac:spMk id="3" creationId="{SHAPE}"/></ac:deMkLst><p188:pos x="120000" y="240000"/><p188:replyLst><p188:reply id="{REPLY}" authorId="{BOB}" status="resolved" created="2026-08-24T12:01:00Z"><p188:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Reply</a:t></a:r></a:p></p188:txBody></p188:reply></p188:replyLst><p188:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>First line</a:t></a:r></a:p><a:p><a:r><a:t>Second line</a:t></a:r></a:p></p188:txBody></p188:cm></p188:cmLst>"#;
        let authors = r#"<p188:authorLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main"><p188:author id="{ADA}" name="Ada" initials="AL"/><p188:author id="{BOB}" name="Bob" initials="B"/></p188:authorLst>"#;
        let data = build_modern_comment_deck(comments, authors);
        let presentation = parse_presentation_from_bytes(&data).expect("modern comments parse");
        let comment = &presentation.slides[0].comments[0];

        assert_eq!(comment.id.as_deref(), Some("{ROOT}"));
        assert_eq!(comment.modern_author_id.as_deref(), Some("{ADA}"));
        assert_eq!(comment.author.as_deref(), Some("Ada"));
        assert_eq!(comment.status.as_deref(), Some("active"));
        assert_eq!(comment.date.as_deref(), Some("2026-08-24T12:00:00Z"));
        assert_eq!((comment.x, comment.y), (Some(120000), Some(240000)));
        assert_eq!(
            comment.anchors,
            vec![PptxCommentAnchor::DrawingElement {
                element_id: Some("3".to_string()),
                creation_id: Some("{SHAPE}".to_string()),
            }]
        );
        assert_eq!(
            serde_json::to_value(&comment.anchors).unwrap()[0]["elementId"],
            "3"
        );
        assert_eq!(comment.text, "First line\nSecond line");
        assert_eq!(comment.replies.len(), 1);
        assert_eq!(comment.replies[0].author.as_deref(), Some("Bob"));
        assert_eq!(comment.replies[0].status.as_deref(), Some("resolved"));
        assert_eq!(comment.replies[0].text, "Reply");
    }

    #[test]
    fn comment_relationship_type_not_target_path_selects_the_part() {
        let modern = r#"<p188:cmLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p188:cm id="{ROOT}" authorId="{ADA}"><p188:txBody><a:p><a:r><a:t>Typed target</a:t></a:r></a:p></p188:txBody></p188:cm></p188:cmLst>"#;
        let decoy = valid_comment_xml("Wrong decoy");
        let rels = r#"<Relationship Id="rInvalidMode" Type="http://schemas.microsoft.com/office/2018/10/relationships/comments" Target="../comments/decoy.xml" TargetMode="Invalid"/><Relationship Id="rDecoy" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../comments/decoy.xml"/><Relationship Id="rModern" Type="http://schemas.microsoft.com/office/2018/10/relationships/comments" Target="../review/thread.xml"/>"#;
        let data = build_typed_comment_deck(
            rels,
            &[
                ("ppt/comments/decoy.xml", &decoy),
                ("ppt/review/thread.xml", modern),
            ],
            r#"<p188:authorLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main"><p188:author id="{ADA}" name="Ada"/></p188:authorLst>"#,
        );

        let presentation = parse_presentation_from_bytes(&data).expect("typed target parses");
        assert_eq!(presentation.slides[0].comments.len(), 1);
        assert_eq!(presentation.slides[0].comments[0].text, "Typed target");
    }

    #[test]
    fn classic_and_modern_comment_parts_have_order_independent_precedence() {
        let classic = valid_comment_xml("Classic");
        let modern = r#"<p188:cmLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p188:cm id="{ROOT}" authorId="{ADA}"><p188:txBody><a:p><a:r><a:t>Modern</a:t></a:r></a:p></p188:txBody></p188:cm></p188:cmLst>"#;
        let classic_rel = r#"<Relationship Id="rClassic" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../review/classic.xml"/>"#;
        let modern_rel = r#"<Relationship Id="rModern" Type="http://schemas.microsoft.com/office/2018/10/relationships/comments" Target="../review/modern.xml"/>"#;
        for rels in [
            format!("{classic_rel}{modern_rel}"),
            format!("{modern_rel}{classic_rel}"),
        ] {
            let data = build_typed_comment_deck(
                &rels,
                &[
                    ("ppt/review/classic.xml", &classic),
                    ("ppt/review/modern.xml", modern),
                ],
                r#"<p188:authorLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main"><p188:author id="{ADA}" name="Ada"/></p188:authorLst>"#,
            );
            let presentation = parse_presentation_from_bytes(&data).expect("mixed comments parse");
            assert_eq!(
                presentation.slides[0]
                    .comments
                    .iter()
                    .map(|comment| comment.text.as_str())
                    .collect::<Vec<_>>(),
                ["Classic", "Modern"],
            );
        }
    }

    fn oversized_comment_authors_xml() -> String {
        format!(
            r#"<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" padding="{}"><p:cmAuthor id="0" name="Ada"/></p:cmAuthorLst>"#,
            "x".repeat(4096)
        )
    }

    #[test]
    fn unreferenced_oversized_comment_authors_is_not_observed_or_poisoned() {
        const LIMIT: u64 = 2048;
        let authors = oversized_comment_authors_xml();
        let mut data = build_comment_deck([false, false], [None, None], &authors);
        forge_declared_size(&mut data, "ppt/commentAuthors.xml", 1);
        COMMENT_AUTHORS_LOAD_COUNT.with(|c| c.set(0));
        let presentation = parse_presentation_from_bytes_with_limits(
            &data,
            Some(LIMIT),
            Some(128 * 1024),
            "parse",
        )
        .expect("an unreferenced author part must remain unobserved");
        assert_eq!(presentation.slides.len(), 3);
        assert!(presentation
            .slides
            .iter()
            .all(|slide| slide.comments.is_empty()));
        assert_eq!(COMMENT_AUTHORS_LOAD_COUNT.with(|c| c.get()), 0);
    }

    #[test]
    fn unreadable_comments_do_not_observe_or_poison_comment_authors() {
        const LIMIT: u64 = 2048;
        let authors = oversized_comment_authors_xml();
        for (label, comments) in [("missing", None), ("malformed", Some("<p:cmLst"))] {
            let mut data = build_comment_deck([true, false], [comments, None], &authors);
            forge_declared_size(&mut data, "ppt/commentAuthors.xml", 1);
            COMMENT_AUTHORS_LOAD_COUNT.with(|c| c.set(0));
            let presentation = parse_presentation_from_bytes_with_limits(
                &data,
                Some(LIMIT),
                Some(128 * 1024),
                "parse",
            )
            .unwrap_or_else(|error| panic!("{label} comments must short-circuit authors: {error}"));
            assert!(presentation.slides[0].comments.is_empty(), "{label}");
            assert_eq!(
                COMMENT_AUTHORS_LOAD_COUNT.with(|c| c.get()),
                0,
                "{label} comments must not load authors"
            );
        }
    }

    #[test]
    fn two_commented_slides_share_one_owned_comment_author_parse() {
        let first = valid_comment_xml("first");
        let second = valid_comment_xml("second");
        let authors = r#"<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cmAuthor id="0" name="Ada"/></p:cmAuthorLst>"#;
        let data = build_comment_deck(
            [true, true],
            [Some(first.as_str()), Some(second.as_str())],
            authors,
        );
        COMMENT_AUTHORS_LOAD_COUNT.with(|c| c.set(0));
        let presentation = parse_presentation_from_bytes(&data).expect("commented deck parses");
        assert_eq!(
            presentation.slides[0].comments[0].author.as_deref(),
            Some("Ada")
        );
        assert_eq!(
            presentation.slides[1].comments[0].author.as_deref(),
            Some("Ada")
        );
        assert_eq!(presentation.slides[0].comments[0].text, "first");
        assert_eq!(presentation.slides[1].comments[0].text, "second");
        assert_eq!(COMMENT_AUTHORS_LOAD_COUNT.with(|c| c.get()), 1);
    }

    /// #774 MAJOR: a truncated / corrupt ZIP CONTAINER — the most common way a
    /// pptx is broken — degrades to a placeholder deck (one slide) tagged with the
    /// container, rather than throwing an opaque `ZipArchive::new` error before any
    /// part is read. Symmetric with docx `rb7_corrupt_zip_container_degrades_...`.
    #[test]
    fn corrupt_zip_container_degrades_to_placeholder() {
        // Truncated container: a valid deck cut off partway is not a readable zip.
        let full = build_three_slide_deck(9, "<unused/>"); // 9 ⇒ no slide is broken
        let truncated = &full[..full.len() / 2];
        let json = parse_pptx_native(truncated)
            .expect("a corrupt container must open as a placeholder, not error out");
        let v: serde_json::Value = serde_json::from_str(&json).unwrap();
        let slides = v["slides"].as_array().expect("placeholder deck has slides");
        assert_eq!(slides.len(), 1, "one placeholder slide for the whole file");
        let err = slides[0]["parseError"]
            .as_str()
            .expect("placeholder slide carries a parseError");
        assert!(
            err.starts_with("(zip container): "),
            "error is tagged with the container exactly once; got {err:?}"
        );
        assert_eq!(
            err.matches("zip container").count(),
            1,
            "the container tag must not be doubled; got {err:?}"
        );
        assert!(
            slides[0]["elements"].as_array().unwrap().is_empty(),
            "placeholder slide has no elements"
        );

        // Not-a-zip-at-all also degrades (no local file header).
        let garbage = parse_pptx_native(b"this is definitely not a zip file")
            .expect("non-zip bytes must open as a placeholder");
        let gv: serde_json::Value = serde_json::from_str(&garbage).unwrap();
        let garbage_err = gv["slides"][0]["parseError"]
            .as_str()
            .expect("non-zip degrades with a container-tagged error");
        assert!(
            garbage_err.starts_with("(zip container): "),
            "error is tagged with the container exactly once; got {garbage_err:?}"
        );
        assert_eq!(
            garbage_err.matches("zip container").count(),
            1,
            "the container tag must not be doubled; got {garbage_err:?}"
        );
    }

    /// A HEALTHY deck never takes the container-degradation branch: no slide
    /// carries a `parseError` and no "(zip container)" tag appears anywhere, so the
    /// placeholder path is inert for valid files (VRT non-regression by
    /// construction).
    #[test]
    fn healthy_deck_never_degrades_container() {
        let data = build_three_slide_deck(9, "<unused/>"); // no broken slide
        let json = parse_pptx_native(&data).expect("healthy deck parses");
        assert!(
            !json.contains("zip container"),
            "healthy deck must not carry any container-degradation tag"
        );
        let v: serde_json::Value = serde_json::from_str(&json).unwrap();
        for slide in v["slides"].as_array().unwrap() {
            assert!(
                slide["parseError"].is_null(),
                "healthy slide must carry no parseError; got {slide}"
            );
        }
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

    fn forged_slide_package() -> Vec<u8> {
        let padding = "x".repeat(4096);
        let slide = format!(
            r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="{padding}"><p:spTree/></p:cSld></p:sld>"#
        );
        let mut data = build_three_slide_deck(0, &slide);
        forge_declared_size(&mut data, "ppt/slides/slide1.xml", 1);
        data
    }

    fn large_png_picture_deck() -> Vec<u8> {
        use std::io::{Read, Write};
        let base = build_three_slide_deck(9, "<unused/>");
        let slide = br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:pic><p:nvPicPr><p:cNvPr id="2" name="Large PNG"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:spTree></p:cSld></p:sld>"#;
        let rels = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/large.png"/></Relationships>"#;
        let mut png = vec![0u8; 64 * 1024];
        png[..8].copy_from_slice(b"\x89PNG\r\n\x1a\n");
        png[12..16].copy_from_slice(b"IHDR");
        png[16..20].copy_from_slice(&640u32.to_be_bytes());
        png[20..24].copy_from_slice(&480u32.to_be_bytes());

        let mut source = zip::ZipArchive::new(Cursor::new(base)).unwrap();
        let mut output = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut output));
            let options = zip::write::SimpleFileOptions::default();
            for index in 0..source.len() {
                let mut entry = source.by_index(index).unwrap();
                let name = entry.name().to_owned();
                let mut body = Vec::new();
                entry.read_to_end(&mut body).unwrap();
                let body: &[u8] = match name.as_str() {
                    "ppt/slides/slide1.xml" => slide,
                    "ppt/slides/_rels/slide1.xml.rels" => rels,
                    _ => &body,
                };
                writer.start_file(name, options).unwrap();
                writer.write_all(body).unwrap();
            }
            writer.start_file("ppt/media/large.png", options).unwrap();
            writer.write_all(&png).unwrap();
            writer.finish().unwrap();
        }
        output
    }

    #[test]
    fn slide_resource_overrun_is_typed_attributed_and_poisons_the_package() {
        const LIMIT: u64 = 2048;
        let mut zip = open_zip_with_limits(forged_slide_package(), Some(LIMIT), Some(64 * 1024))
            .expect("forged declaration passes package preflight");
        let error = zip
            .run_operation("parse", |zip| {
                let mut shared = bootstrap_presentation(zip).map_err(|e| e.to_string())?;
                produce_slide_unit(0, &mut shared, zip).map_err(|e| e.to_string())
            })
            .expect_err("single-slide placeholder degradation must not hide package poison");
        let details: serde_json::Value = serde_json::from_str(
            error
                .strip_prefix("OOXML_RESOURCE_LIMIT:")
                .expect("canonical typed resource-limit envelope"),
        )
        .unwrap();
        let violation = &details["details"]["violation"];
        assert_eq!(violation["operation"], "parse");
        assert_eq!(violation["part"], "ppt/slides/slide1.xml");
        assert_eq!(violation["metric"], "actual-inflated-bytes");
        assert_eq!(violation["limit"], LIMIT);
        assert_eq!(violation["observed"], LIMIT + 1);

        let later = zip
            .run_operation("markdown", |_| Ok(()))
            .expect_err("poisoned package rejects later operations deterministically");
        assert_eq!(later, error);
    }

    #[test]
    fn free_parse_helper_attributes_resource_poison_to_its_operation() {
        let error = parse_presentation_from_bytes_with_limits(
            &forged_slide_package(),
            Some(2048),
            Some(64 * 1024),
            "markdown",
        )
        .expect_err("free helper must surface the package-wide poison");
        let details: serde_json::Value = serde_json::from_str(
            error
                .strip_prefix("OOXML_RESOURCE_LIMIT:")
                .expect("canonical typed resource-limit envelope"),
        )
        .unwrap();
        let violation = &details["details"]["violation"];
        assert_eq!(violation["operation"], "markdown");
        assert_eq!(violation["part"], "ppt/slides/slide1.xml");
    }

    #[test]
    fn parse_reads_only_png_header_while_extract_materializes_the_entry() {
        const TOTAL_LIMIT: u64 = 16 * 1024;
        let data = large_png_picture_deck();
        let presentation =
            parse_presentation_from_bytes_with_limits(&data, None, Some(TOTAL_LIMIT), "parse")
                .expect("XML plus a 24-byte PNG head stays below the total budget");
        let picture = presentation.slides[0]
            .elements
            .iter()
            .find_map(|element| match element {
                SlideElement::Picture(picture) => Some(picture),
                _ => None,
            })
            .expect("production slide parser emits the picture");
        assert_eq!(picture.intrinsic_width_px, Some(640));
        assert_eq!(picture.intrinsic_height_px, Some(480));

        let error = extract_entry_with_limits(
            &data,
            "ppt/media/large.png",
            None,
            Some(TOTAL_LIMIT),
            "extract-image",
        )
        .expect_err("materializing the same large PNG must exceed the total budget");
        let details: serde_json::Value = serde_json::from_str(
            error
                .strip_prefix("OOXML_RESOURCE_LIMIT:")
                .expect("canonical typed resource-limit envelope"),
        )
        .unwrap();
        let violation = &details["details"]["violation"];
        assert_eq!(violation["operation"], "extract-image");
        assert_eq!(violation["part"], "ppt/media/large.png");
        assert_eq!(violation["metric"], "distinct-inflated-bytes");
        assert_eq!(violation["limit"], TOTAL_LIMIT);
        assert_eq!(violation["observed"], TOTAL_LIMIT + 1);
    }

    #[test]
    fn free_extract_preserves_the_tagged_ordinary_container_error() {
        let error = extract_entry_with_limits(
            b"not a zip package",
            "ppt/media/image1.png",
            None,
            None,
            "extract-image",
        )
        .expect_err("ordinary invalid ZIP remains an extraction error");
        assert!(error.starts_with("(zip container): "), "{error}");
        assert_eq!(error.matches("zip container").count(), 1, "{error}");
        assert!(!error.starts_with("OOXML_RESOURCE_LIMIT:"), "{error}");
    }

    #[test]
    #[allow(clippy::type_complexity)] // Exact exported ABI shapes are the assertion.
    fn public_wasm_signatures_remain_stable() {
        let _: fn(&[u8], Option<u64>, Option<u64>) -> Result<Vec<u8>, JsValue> = parse_pptx;
        let _: fn(&[u8], Option<u64>, Option<u64>) -> Result<String, JsValue> = pptx_to_markdown;
        let _: fn(&[u8], &str, Option<u64>, Option<u64>) -> Result<Vec<u8>, JsValue> =
            extract_media;
        let _: fn(&[u8], &str, Option<u64>, Option<u64>) -> Result<Vec<u8>, JsValue> =
            extract_image;
        let _: fn(Vec<u8>, Option<u64>, Option<u64>, Option<u64>) -> Result<PptxArchive, JsValue> =
            PptxArchive::new;
        let _: fn(&mut PptxArchive) -> Result<Vec<u8>, JsValue> = PptxArchive::parse;
        let _: fn(&mut PptxArchive, &str) -> Result<Vec<u8>, JsValue> = PptxArchive::extract_media;
        let _: fn(&mut PptxArchive, &str) -> Result<Vec<u8>, JsValue> = PptxArchive::extract_image;
        let _: fn(&mut PptxArchive) -> Result<String, JsValue> = PptxArchive::to_markdown;
        let _: fn(&PptxArchive) -> Result<(), JsValue> = PptxArchive::assert_healthy;
        let _: fn(&[u8]) -> Result<String, String> = parse_pptx_native;
        let _: fn(&[u8]) -> Result<String, String> = to_markdown_native;
    }
}
