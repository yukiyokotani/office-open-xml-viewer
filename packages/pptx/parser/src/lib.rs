use ooxml_common::depth::parse_guarded;
use ooxml_common::ns::is_r_ns;
use std::collections::HashMap;
use std::io::Cursor;
use wasm_bindgen::prelude::*;

mod table_style_presets;

mod types;
pub(crate) use types::*;

mod markdown;
use markdown::render_presentation_md;

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
}

/// Increment the D4 parse counter (no-op unless `cfg(test)`). Call immediately
/// before parsing a slide-master / layout / slide XML on the pagination path.
#[inline(always)]
fn note_layout_master_parse() {
    #[cfg(test)]
    LAYOUT_MASTER_PARSE_COUNT.with(|c| c.set(c.get() + 1));
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
    max_zip_entry_bytes: Option<u64>,
    max_zip_total_bytes: Option<u64>,
    max_zip_entries: Option<u64>,
) -> Result<Vec<u8>, JsValue> {
    console_error_panic_hook::set_once();
    let _guard =
        ooxml_common::zip::scoped_limits(max_zip_entry_bytes, max_zip_total_bytes, max_zip_entries);
    validate_zip_budget(data).map_err(|e| JsValue::from_str(&format!("pptx-parser error: {e}")))?;
    let presentation = parse_presentation_from_bytes(data)
        .map_err(|e| JsValue::from_str(&format!("pptx-parser error: {e}")))?;
    serde_json::to_vec(&presentation)
        .map_err(|e| JsValue::from_str(&format!("serialize error: {e}")))
}

/// WASM-callable markdown projection. Shares the body of `to_markdown_native`
/// so the browser / Node WASM path and the native mcp-server path stay in
/// lock-step. See `to_markdown_native` for the design rationale.
#[wasm_bindgen]
pub fn pptx_to_markdown(
    data: &[u8],
    max_zip_entry_bytes: Option<u64>,
    max_zip_total_bytes: Option<u64>,
    max_zip_entries: Option<u64>,
) -> Result<String, JsValue> {
    console_error_panic_hook::set_once();
    let _guard =
        ooxml_common::zip::scoped_limits(max_zip_entry_bytes, max_zip_total_bytes, max_zip_entries);
    validate_zip_budget(data).map_err(|e| JsValue::from_str(&format!("pptx-parser error: {e}")))?;
    let pres = parse_presentation_from_bytes(data)
        .map_err(|e| JsValue::from_str(&format!("pptx-parser error: {e}")))?;
    Ok(render_presentation_md(&pres))
}

/// Native equivalent of `parse_pptx` for use from the MCP server.
pub fn parse_pptx_native(data: &[u8]) -> Result<String, String> {
    let _guard = ooxml_common::zip::scoped_limits(None, None, None);
    validate_zip_budget(data)?;
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
    let _guard = ooxml_common::zip::scoped_limits(None, None, None);
    validate_zip_budget(data)?;
    let pres = parse_presentation_from_bytes(data).map_err(|e| e.to_string())?;
    Ok(render_presentation_md(&pres))
}

/// Extract raw bytes for a single entry (e.g. "ppt/media/media2.mp4") from a
/// pptx zip archive. Used by the main thread to materialize media blobs for
/// interactive playback without re-parsing the whole file.
#[wasm_bindgen]
pub fn extract_media(
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

/// Extract raw bytes for a single embedded image entry (e.g.
/// "ppt/media/image1.png") from a pptx zip archive. Thin `wasm_bindgen` wrapper
/// over the shared [`ooxml_common::zip::extract_zip_entry`] reader; used by the
/// main thread to lazily materialize image blobs on demand.
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

/// A stateful handle over an opened pptx archive.
///
/// The free functions above (`parse_pptx` / `pptx_to_markdown` / `extract_media`
/// / `extract_image`) each re-copy the whole file into WASM and re-scan the ZIP
/// central directory on every call. A `PptxArchive` copies the bytes into WASM
/// **once** (in `new`) and keeps the opened [`PptxZip`] alive, so a `parse`
/// followed by any number of `extract_media` / `extract_image` calls (the
/// viewer's parse-then-lazily-load-media pattern) pays the copy + open cost a
/// single time. `ZipArchive<Cursor<Vec<u8>>>` is self-contained (it owns its
/// bytes and holds no borrow into the input), which is what lets it live in a
/// `#[wasm_bindgen]` struct field.
///
/// The retained `max` mirrors the per-call `scoped_max` guard the free functions
/// install: every method re-installs it for its own scope so the zip-bomb entry
/// cap is honored identically whether callers use the handle or the free
/// functions.
#[wasm_bindgen]
pub struct PptxArchive {
    /// The opened archive, or the container-open error string when the ZIP itself
    /// was truncated / corrupt (#774, RB7 MAJOR). Deferring the failure here —
    /// instead of erroring out of `new` — lets `parse()` return a degraded
    /// placeholder presentation (symmetric with a corrupt inner slide) rather than
    /// the constructor throwing an opaque error the viewer can't turn into a
    /// placeholder slide.
    archive: Result<PptxZip, String>,
    budget: ooxml_common::zip::ZipBudget,
}

#[wasm_bindgen]
impl PptxArchive {
    /// Copy `data` into WASM once and open the ZIP central directory once.
    /// `max_zip_entry_bytes` is retained and applied on every subsequent method
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
    ) -> Result<PptxArchive, JsValue> {
        console_error_panic_hook::set_once();
        // #774 (RB7 MAJOR): a truncated / corrupt CONTAINER is deferred, not
        // thrown, so `parse()` can degrade it to a placeholder presentation
        // instead of the constructor failing with an opaque error.
        let budget = ooxml_common::zip::ZipBudget::new(
            max_zip_entry_bytes,
            max_zip_total_bytes,
            max_zip_entries,
        );
        let _guard = ooxml_common::zip::scoped_budget(&budget);
        let archive = match open_zip(data) {
            Ok(archive) => Ok(archive),
            Err(error) if error.starts_with("ZIP archive exceeds") => {
                return Err(JsValue::from_str(&error));
            }
            Err(error) => Err(error),
        };
        Ok(PptxArchive { archive, budget })
    }

    /// Parse the retained archive and return the model as UTF-8 JSON bytes.
    /// Byte-for-byte identical to `parse_pptx` on the same file. When the
    /// CONTAINER failed to open (#774) the model is a degraded placeholder
    /// presentation tagged with the container.
    pub fn parse(&mut self) -> Result<Vec<u8>, JsValue> {
        let _guard = ooxml_common::zip::scoped_budget(&self.budget);
        let presentation = match self.archive.as_mut() {
            Ok(zip) => parse_presentation(zip)
                .map_err(|e| JsValue::from_str(&format!("pptx-parser error: {e}")))?,
            Err(e) => degraded_container_presentation(e.clone()),
        };
        serde_json::to_vec(&presentation)
            .map_err(|e| JsValue::from_str(&format!("serialize error: {e}")))
    }

    /// Extract raw bytes for one media entry (e.g. "ppt/media/media2.mp4") from
    /// the retained archive. Twin of the free `extract_media`, but reads through
    /// the already-open archive instead of re-opening it. A corrupt container has
    /// no entries, so this surfaces the container-open error.
    pub fn extract_media(&mut self, path: &str) -> Result<Vec<u8>, JsValue> {
        let _guard = ooxml_common::zip::scoped_budget(&self.budget);
        let zip = self
            .archive
            .as_mut()
            .map_err(|e| JsValue::from_str(&format!("pptx-parser error: {e}")))?;
        ooxml_common::zip::read_zip_bytes(zip, path).map_err(|e| JsValue::from_str(&e))
    }

    /// Extract raw bytes for one embedded image entry (e.g.
    /// "ppt/media/image1.png") from the retained archive. Twin of the free
    /// `extract_image`. A corrupt container has no entries, so this surfaces the
    /// container-open error.
    pub fn extract_image(&mut self, path: &str) -> Result<Vec<u8>, JsValue> {
        let _guard = ooxml_common::zip::scoped_budget(&self.budget);
        let zip = self
            .archive
            .as_mut()
            .map_err(|e| JsValue::from_str(&format!("pptx-parser error: {e}")))?;
        ooxml_common::zip::read_zip_bytes(zip, path).map_err(|e| JsValue::from_str(&e))
    }

    /// GitHub-flavoured markdown projection of the retained archive. Mirrors the
    /// free `pptx_to_markdown`. A corrupt container degrades to an empty deck.
    pub fn to_markdown(&mut self) -> Result<String, JsValue> {
        let _guard = ooxml_common::zip::scoped_budget(&self.budget);
        let pres = match self.archive.as_mut() {
            Ok(zip) => parse_presentation(zip)
                .map_err(|e| JsValue::from_str(&format!("pptx-parser error: {e}")))?,
            Err(e) => degraded_container_presentation(e.clone()),
        };
        Ok(render_presentation_md(&pres))
    }
}

// ===========================
//  ZIP helpers
// ===========================

/// The parser's ZIP archive type. Owns its backing bytes (`Cursor<Vec<u8>>`)
/// rather than borrowing them, so a `PptxArchive` handle can keep a single
/// opened archive alive across `parse` / `extract_media` / `extract_image` /
/// `to_markdown` calls — the central directory is scanned once and the bytes are
/// copied into WASM once. `ZipArchive<Cursor<Vec<u8>>>` is fully self-contained
/// (no borrow into the input), which is what lets a `#[wasm_bindgen]` handle
/// store it as a field.
pub(crate) type PptxZip = zip::ZipArchive<Cursor<Vec<u8>>>;

pub(crate) fn read_zip_str(
    zip: &mut PptxZip,
    path: &str,
) -> Result<String, Box<dyn std::error::Error>> {
    ooxml_common::zip::read_zip_string(zip, path).map_err(Into::into)
}

// ===========================
//  Table style data model
// ===========================

/// Resolved fills and borders extracted from a single <a:tblStyle> definition.
#[derive(Debug, Clone, Default)]
struct TableStyleDef {
    whole_fill: Option<Fill>,
    whole_inside_h: Option<Stroke>,
    whole_inside_v: Option<Stroke>,
    /// Outer top/bottom edge border (from wholeTbl tcBdr top/bottom)
    whole_outer_h: Option<Stroke>,
    /// Outer left/right edge border (from wholeTbl tcBdr left/right)
    whole_outer_v: Option<Stroke>,
    band1h_fill: Option<Fill>,
    band2h_fill: Option<Fill>,
    first_row_fill: Option<Fill>,
    first_row_border_b: Option<Stroke>,
    last_row_fill: Option<Fill>,
    first_col_fill: Option<Fill>,
    last_col_fill: Option<Fill>,
    /// Default text colour per role, from `<a:tcTxStyle>` (schemeClr/srgbClr).
    /// e.g. wholeTbl → dk1, firstRow header → lt1 (white). Hex, no `#`.
    whole_text_color: Option<String>,
    first_row_text_color: Option<String>,
    last_row_text_color: Option<String>,
    first_col_text_color: Option<String>,
    last_col_text_color: Option<String>,
    /// Default bold per role, from `<a:tcTxStyle b="on">` (ECMA-376 §20.1.4.2.28).
    /// e.g. a firstRow header is commonly bold.
    first_row_bold: Option<bool>,
    last_row_bold: Option<bool>,
    first_col_bold: Option<bool>,
    last_col_bold: Option<bool>,
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

/// id → target  (used for image/slide lookups by rId). Thin adapter over
/// [`ooxml_common::rels::parse_rels`] that flattens each `RelTarget` back to its
/// raw target string (both Internal part names and External URLs are kept
/// verbatim — resolution to a zip part happens later via [`resolve_path`]),
/// preserving this parser's long-standing `HashMap<rId, Target>` shape.
pub(crate) fn parse_rels(xml: &str) -> HashMap<String, String> {
    ooxml_common::rels::parse_rels(xml)
        .into_iter()
        .map(|(id, rel)| (id, rel.target))
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
    zip: &mut PptxZip,
) -> HashMap<String, String> {
    let mut result: HashMap<String, String> = HashMap::new();
    let doc = match parse_guarded(rels_xml) {
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
        let drawing_target = smartart_drawing_relid(&data_target, zip)
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
            // Diagram parts live at ppt/diagrams/; every referencing part
            // (slide / master) is one level under ppt/, so a "../diagrams/…"
            // target resolves the same from any ppt/*/ base.
            let drawing_path = resolve_path("ppt/slides", &dt);
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
fn smartart_drawing_relid(data_target: &str, zip: &mut PptxZip) -> Option<String> {
    let data_path = resolve_path("ppt/slides", data_target);
    let xml = read_zip_str(zip, &data_path).ok()?;
    let doc = parse_guarded(&xml).ok()?;
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
    let doc = parse_guarded(rels_xml).ok()?;
    for rel in doc.root_element().children().filter(|n| n.is_element()) {
        if let Some(rel_type) = attr(&rel, "Type") {
            if rel_type.ends_with(type_suffix) {
                return attr(&rel, "Target");
            }
        }
    }
    None
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
    let theme: &HashMap<String, String> = eff.map(|e| &e.theme).unwrap_or(theme);
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
    // rejects it now lives in `parse_guarded`.
    let doc = parse_guarded(xml)?;
    let root = doc.root_element(); // <p:sld>
    let hidden = slide_is_hidden(root);
    let c_sld = child(root, "cSld");

    // Background chain: slide → layout → master. Each level resolves a blip
    // background (§20.1.8.14) against its own rels + part directory, so the
    // closures are run sequentially (one mutable borrow of `zip` at a time).
    let mut background: Option<Fill> = None;

    // Slide-level bg (rels = slide rels, part dir = ppt/slides).
    if let Some(n) = c_sld {
        let mut resolve = |rid: &str| -> Option<String> {
            let target = rels.get(rid)?;
            let path = resolve_path("ppt/slides", target);
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

    let slide_dir = "ppt/slides";
    let mut elements = Vec::new();

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
        if eff.is_some() {
            if let Some(mxml) = master_xml {
                note_layout_master_parse();
                if let Ok(mdoc) = parse_guarded(mxml) {
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
    }

    // ── Layout non-placeholder shapes (rendered BEFORE slide shapes) ──────
    // These are decorative background elements defined in the slide layout
    // (e.g. coloured bands, logos) that are not placeholder anchors.
    if let Some(lxml) = layout_xml {
        note_layout_master_parse();
        if let Ok(ldoc) = parse_guarded(lxml) {
            let lroot = ldoc.root_element();
            if let Some(lsp_tree) = child(lroot, "cSld").and_then(|n| child(n, "spTree")) {
                let empty_lph = LayoutPlaceholders::default();
                for node in lsp_tree.children().filter(|n| n.is_element()) {
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
                }
            }
        }
    }

    // ── Slide shapes ─────────────────────────────────────────────────────
    for node in sp_tree.children().filter(|n| n.is_element()) {
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
    }

    // ── Notes slide & comments (Phase 2 surfacing only — no rendering) ────
    let notes = load_notes_slide(zip, rels);
    let comments = load_pptx_comments(zip, rels);

    Ok(Slide {
        index,
        slide_number: index + 1,
        // Stamped by the build loop, which owns the resolved slide part path.
        part_name: None,
        background,
        elements,
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
        notes: None,
        comments: Vec::new(),
        hidden: false,
        parse_error: Some(format!("{part}: {detail}")),
    }
}

/// Resolve the slide's `notesSlide` relationship, read the notes part, and
/// return its plain text (paragraphs joined by '\n'). Returns `None` when
/// the slide has no notes part or the part can't be read.
fn load_notes_slide(zip: &mut PptxZip, rels: &HashMap<String, String>) -> Option<String> {
    // rels here is the slide's _rels map (rId → Target) parsed by the caller.
    // The relationship Type ends with "/notesSlide". The cleanest way to find
    // the right entry is to look at every value in the map and pick the one
    // pointing at "../notesSlides/...".
    let target = rels.values().find(|t| t.contains("notesSlides/"))?;
    let path = if target.starts_with('/') {
        target.trim_start_matches('/').to_string()
    } else {
        // Relative to ppt/slides/ — resolve "../notesSlides/notesSlide1.xml".
        resolve_path("ppt/slides", target)
    };
    let xml = read_zip_str(zip, &path).ok()?;
    let doc = parse_guarded(&xml).ok()?;
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

/// Resolve and parse the slide's `comments` relationship (legacy
/// `<p:cmLst>` format). Modern threaded comments live in a different
/// namespace and are not yet supported.
fn load_pptx_comments(zip: &mut PptxZip, rels: &HashMap<String, String>) -> Vec<PptxComment> {
    let Some(target) = rels.values().find(|t| t.contains("comments/")) else {
        return Vec::new();
    };
    let path = if target.starts_with('/') {
        target.trim_start_matches('/').to_string()
    } else {
        resolve_path("ppt/slides", target)
    };
    let Ok(xml) = read_zip_str(zip, &path) else {
        return Vec::new();
    };
    let Ok(doc) = parse_guarded(&xml) else {
        return Vec::new();
    };

    // commentAuthors.xml is a top-level part — look it up directly.
    let author_xml = read_zip_str(zip, "ppt/commentAuthors.xml").ok();
    let mut authors: HashMap<String, String> = HashMap::new();
    if let Some(ax) = author_xml {
        if let Ok(adoc) = parse_guarded(&ax) {
            for a in adoc
                .descendants()
                .filter(|n| n.is_element() && n.tag_name().name() == "cmAuthor")
            {
                let id = a.attribute("id").unwrap_or("").to_string();
                let name = a.attribute("name").unwrap_or("").to_string();
                if !id.is_empty() && !name.is_empty() {
                    authors.insert(id, name);
                }
            }
        }
    }

    let mut out = Vec::new();
    for cm in doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "cm")
    {
        let author_id = cm.attribute("authorId").unwrap_or("");
        let author = authors.get(author_id).cloned();
        let date = cm
            .attribute("dt")
            .map(String::from)
            .filter(|s| !s.is_empty());
        let text = cm
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "text")
            .and_then(|n| n.text().map(String::from))
            .unwrap_or_default();
        out.push(PptxComment { author, date, text });
    }
    out
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
pub(crate) fn open_zip(data: Vec<u8>) -> Result<PptxZip, String> {
    let mut archive =
        zip::ZipArchive::new(Cursor::new(data)).map_err(|e| format!("(zip container): {e}"))?;
    ooxml_common::zip::validate_archive_limits(&mut archive)?;
    Ok(archive)
}

fn validate_zip_budget(data: &[u8]) -> Result<(), String> {
    // Keep malformed-container degradation intact; reject only limits proven by
    // a readable central directory before parsing/inflation starts.
    let Ok(mut archive) = zip::ZipArchive::new(Cursor::new(data)) else {
        return Ok(());
    };
    ooxml_common::zip::validate_archive_limits(&mut archive)
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
    let mut zip = match open_zip(data.to_vec()) {
        Ok(zip) => zip,
        Err(e) => return Ok(degraded_container_presentation(e)),
    };
    parse_presentation(&mut zip)
}

fn parse_presentation(zip: &mut PptxZip) -> Result<Presentation, Box<dyn std::error::Error>> {
    // --- presentation.xml ---
    let pres_xml = read_zip_str(zip, "ppt/presentation.xml")?;
    let pres_doc = parse_guarded(&pres_xml)?;
    let pres_root = pres_doc.root_element();

    let sld_sz = child(pres_root, "sldSz");
    let slide_width = sld_sz.and_then(|n| attr_i64(&n, "cx")).unwrap_or(9_144_000);
    let slide_height = sld_sz.and_then(|n| attr_i64(&n, "cy")).unwrap_or(6_858_000);

    // Ordered slide rIds
    let slide_rids: Vec<String> = child(pres_root, "sldIdLst")
        .map(|lst| {
            children_vec(lst, "sldId")
                .into_iter()
                .filter_map(|n| attr_r(&n, "id"))
                .collect()
        })
        .unwrap_or_default();

    // --- ppt/_rels/presentation.xml.rels ---
    let pres_rels_xml = read_zip_str(zip, "ppt/_rels/presentation.xml.rels")?;
    let pres_rels = parse_rels(&pres_rels_xml);

    // --- Presentation-level theme colors ---
    // Used for the deck-wide defaults on `Presentation` (default text color,
    // major/minor fonts, hyperlink colors) and as the fallback theme for any
    // master that declares no /theme relationship of its own.
    let theme_xml = find_rel_target_by_type(&pres_rels_xml, "/theme")
        .map(|t| resolve_path("ppt", &t))
        .and_then(|path| read_zip_str(zip, &path).ok())
        .unwrap_or_default();
    let theme = parse_theme_colors(&theme_xml);

    // --- Presentation-level fallback master ---
    // The first slide master referenced by the presentation. Used for slides
    // whose layout→master→theme chain can't be resolved (simple/old decks), so
    // their behavior is unchanged from before per-slide resolution existed.
    let pres_master_path: Option<String> =
        find_rel_target_by_type(&pres_rels_xml, "/slideMaster").map(|t| resolve_path("ppt", &t));

    // Cache of ParsedMaster keyed by master ZIP path. Slides sharing a master
    // reuse the bundle instead of recomputing every master-derived map. Seed it
    // with the presentation master so the slide loop reuses the fallback build
    // instead of rebuilding it — for a single-master deck every slide resolves
    // to this same master, and without seeding the fallback build and the first
    // slide's cache-miss build would each compute the identical bundle twice.
    let mut master_cache: HashMap<String, ParsedMaster> = HashMap::new();
    // Bundle for the truly no-master / empty-path case (no /slideMaster rel on
    // the presentation). Only built then; otherwise the fallback is the cached
    // presentation-master bundle.
    let no_master_bundle: Option<ParsedMaster> = match pres_master_path.as_deref() {
        Some(p) => {
            master_cache.insert(p.to_owned(), build_master_bundle(p, &theme, zip));
            None
        }
        None => Some(build_master_bundle("", &theme, zip)),
    };

    // Cache of the layout single-pass extraction (`ParsedLayout`) keyed by layout
    // ZIP path (D4), mirroring `master_cache`. Slides sharing a layout reuse its
    // resolved placeholders + layout background + showMasterSp instead of
    // re-parsing the layout XML four times per slide. Only NO-override slides
    // populate/read the cache: the entry is resolved against the master's baked
    // theme, and a slide's layout→master chain is 1:1 (a layout names exactly one
    // master), so every no-override slide on a given layout shares that theme. A
    // slide with a `<p:clrMapOvr>` builds a fresh `ParsedLayout` against its
    // override theme instead (kept out of the cache).
    let mut layout_cache: HashMap<String, ParsedLayout> = HashMap::new();

    // Pre-collect slide XMLs, their rels, the layout XML, and layout rels
    struct SlideRaw {
        index: usize,
        /// ZIP path of the slide part (e.g. `ppt/slides/slide3.xml`). Carried so
        /// a parse failure in the build loop can name the offending part (RB7).
        slide_path: String,
        /// The slide XML, or `Err(detail)` when the part itself could not be
        /// read (RB7 partial degradation) — the build loop turns that into a
        /// placeholder slide instead of aborting the whole deck.
        slide_xml: Result<String, String>,
        slide_rels: HashMap<String, String>,
        smartart_drawings: HashMap<String, String>,
        /// ZIP path of the slide's layout (e.g. `ppt/slideLayouts/slideLayout3.xml`).
        /// Used as the `layout_cache` key so slides sharing a layout reuse its
        /// single-pass `ParsedLayout` (D4). `None` when the slide has no layout.
        layout_path: Option<String>,
        layout_xml: Option<String>,
        layout_rels: HashMap<String, String>,
        layout_dir: String,
        /// ZIP path of the slide's effective master, resolved through the
        /// slide→slideLayout→slideMaster rels chain. `None` when the chain
        /// can't be followed (no layout, or the layout has no /slideMaster
        /// relationship); such slides fall back to the presentation master.
        master_path: Option<String>,
    }

    let mut raw_slides: Vec<SlideRaw> = Vec::new();

    for (idx, r_id) in slide_rids.iter().enumerate() {
        let rel_target = match pres_rels.get(r_id) {
            Some(t) => t.clone(),
            None => continue,
        };
        // Resolve via `resolve_path` (not `format!("ppt/{rel_target}")`) so a
        // package-root-absolute slide Target — e.g. `/ppt/slides/slide1.xml`
        // (leading slash, OPC / ECMA-376 Part 2 §9.3) — resolves correctly
        // instead of producing `ppt//ppt/slides/slide1.xml`. Relative targets
        // (the common `slides/slide1.xml`) are unaffected. Same fix class as
        // the chart-rel resolution above.
        let slide_path = resolve_path("ppt", &rel_target);
        let slide_file = rel_target
            .split('/')
            .next_back()
            .unwrap_or("slide.xml")
            .to_owned();
        let rels_path = format!("ppt/slides/_rels/{slide_file}.rels");

        // RB7: a slide part that can't be read no longer aborts the whole deck.
        // Record the failure and let the build loop emit a placeholder for THIS
        // slide while the others parse normally.
        let slide_xml = read_zip_str(zip, &slide_path).map_err(|e| e.to_string());
        let slide_rels_xml = read_zip_str(zip, &rels_path).unwrap_or_default();
        let slide_rels = parse_rels(&slide_rels_xml);
        let smartart_drawings = build_smartart_drawings(&slide_rels_xml, zip);

        // Layout XML
        let layout_path = find_rel_target_by_type(&slide_rels_xml, "/slideLayout")
            .map(|target| resolve_path("ppt/slides", &target));

        let layout_xml = layout_path
            .as_deref()
            .and_then(|path| read_zip_str(zip, path).ok());

        let layout_dir = layout_path
            .as_deref()
            .and_then(|p| p.rsplit_once('/').map(|(dir, _)| dir.to_owned()))
            .unwrap_or_else(|| "ppt/slideLayouts".to_owned());

        // Layout rels XML — needed both to resolve images inside the layout and
        // to follow the layout→slideMaster relationship for per-slide master
        // resolution (ECMA-376 §19.3.1.43).
        let layout_rels_xml: String = layout_path
            .as_deref()
            .and_then(|path| {
                let file = path.split('/').next_back().unwrap_or("layout.xml");
                let rels_p = format!("ppt/slideLayouts/_rels/{file}.rels");
                read_zip_str(zip, &rels_p).ok()
            })
            .unwrap_or_default();
        let layout_rels = parse_rels(&layout_rels_xml);

        // Resolve this slide's master via the layout's /slideMaster rel.
        let master_path: Option<String> = find_rel_target_by_type(&layout_rels_xml, "/slideMaster")
            .map(|t| resolve_path(&layout_dir, &t));

        raw_slides.push(SlideRaw {
            index: idx,
            slide_path,
            slide_xml,
            slide_rels,
            smartart_drawings,
            layout_path,
            layout_xml,
            layout_rels,
            layout_dir,
            master_path,
        });
    }

    let mut slides = Vec::new();
    for raw in &raw_slides {
        // RB7: a slide part that couldn't be READ (recorded above) degrades to a
        // placeholder now, before any master/layout resolution touches it.
        let slide_xml = match &raw.slide_xml {
            Ok(xml) => xml.as_str(),
            Err(detail) => {
                slides.push(broken_slide(raw.index, &raw.slide_path, detail));
                continue;
            }
        };
        // Resolve this slide's ParsedMaster: build (and cache) one for the
        // slide's own master when the layout→master chain resolved; otherwise
        // use the presentation-level fallback bundle. Building is keyed by
        // master path so slides sharing a master don't recompute.
        let bundle: &ParsedMaster = match raw.master_path.as_deref() {
            Some(mp) if !mp.is_empty() => {
                if !master_cache.contains_key(mp) {
                    let b = build_master_bundle(mp, &theme, zip);
                    master_cache.insert(mp.to_owned(), b);
                }
                &master_cache[mp]
            }
            // Unresolved chain → presentation-level fallback. The presentation
            // master (when present) was seeded into the cache above, so reuse
            // that entry; only a deck with no /slideMaster rel falls through to
            // `no_master_bundle`.
            _ => pres_master_path
                .as_deref()
                .map(|p| &master_cache[p])
                // ast-grep-ignore: no-unwrap-in-parser-production
                .unwrap_or_else(|| no_master_bundle.as_ref().unwrap()),
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
        let clr_map_ovr: Option<HashMap<String, String>> = parse_clr_map_ovr(slide_xml)
            .or_else(|| raw.layout_xml.as_deref().and_then(parse_clr_map_ovr));
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
                parse_guarded(xml).ok()
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
        let (layout_theme, layout_master_bullets): (
            &HashMap<String, String>,
            &HashMap<String, LevelBullets>,
        ) = match effective_master.as_ref() {
            Some(e) => (&e.theme, &e.master_level_bullets),
            None => (&bundle.theme, &bundle.master_level_bullets),
        };
        // Build a `ParsedLayout` from a layout XML string with the resolved
        // theme/bullets and this bundle's remaining (theme-independent) maps.
        let build_parsed_layout = |lx: &str, zip: &mut PptxZip| -> ParsedLayout {
            parse_layout(
                lx,
                &bundle.master_font_sizes,
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
                &raw.layout_dir,
                &raw.layout_rels,
                zip,
            )
        };

        // No-override slide WITH a layout path → cache by layout path (its entry
        // is resolved against the master-baked theme, which every no-override
        // slide on this layout shares). Otherwise build a fresh, uncached one
        // (override slide, or the rare no-layout-path case).
        let fresh_layout: Option<ParsedLayout> = match (
            effective_master.is_none(),
            raw.layout_xml.as_deref(),
            raw.layout_path.as_deref(),
        ) {
            (true, Some(lx), Some(lp)) => {
                if !layout_cache.contains_key(lp) {
                    let pl = build_parsed_layout(lx, zip);
                    layout_cache.insert(lp.to_owned(), pl);
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
        let slide = match parse_slide(
            slide_xml,
            parsed_layout,
            raw.layout_xml.as_deref(),
            &raw.layout_rels,
            &raw.layout_dir,
            bundle,
            effective_master.as_ref(),
            raw.index,
            &raw.slide_rels,
            &raw.smartart_drawings,
            zip,
        ) {
            Ok(slide) => slide,
            Err(e) => broken_slide(raw.index, &raw.slide_path, &e.to_string()),
        };
        // Stamp the resolved slide part path (e.g. `ppt/slides/slide3.xml`) so
        // the TS side can map an internal hyperlink slide jump to this index.
        // The build loop owns `raw.slide_path`; keying by it here (rather than
        // threading it through `parse_slide`) keeps that function's signature
        // untouched. `broken_slide` already set it, so re-stamping is a no-op there.
        let mut slide = slide;
        slide.part_name = Some(raw.slide_path.clone());
        slides.push(slide);
    }

    let default_text_color = theme.get("dk1").cloned();
    let major_font = theme.get("+mj-lt").cloned();
    let minor_font = theme.get("+mn-lt").cloned();
    let hlink_color = theme.get("hlink").cloned();
    let fol_hlink_color = theme.get("folHlink").cloned();
    Ok(Presentation {
        slide_width,
        slide_height,
        slides,
        default_text_color,
        major_font,
        minor_font,
        hlink_color,
        fol_hlink_color,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Read;
    use crate::chart::{parse_chartex, parse_legacy_chart};
    use ooxml_common::math::nodes_to_text;

    // Local-only sample (redistribution-prohibited, gitignored). Tests that
    // depend on it must skip gracefully on a clean checkout / in CI where the
    // file is absent. See packages/pptx/public/private/.
    const LOCAL_SAMPLE_2: &str = "../public/private/sample-2.pptx";

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
        let mut zip = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();

        let rels = r#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdData1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/>
  <Relationship Id="rIdData2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data2.xml"/>
  <Relationship Id="rIdDrawA" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="../diagrams/drawing1.xml"/>
  <Relationship Id="rIdDrawB" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="../diagrams/drawing2.xml"/>
</Relationships>"#;

        let map = build_smartart_drawings(rels, &mut zip);
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
        let mut zip = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
        let rels = r#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdData1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/>
  <Relationship Id="rIdDraw1" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="../diagrams/drawing1.xml"/>
</Relationships>"#;
        let map = build_smartart_drawings(rels, &mut zip);
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
            extract_image(&buf, "ppt/media/i.png", None, None, None).unwrap(),
            b"X"
        );
    }

    /// A `PictureElement` serializes its blip as a zip path + mime, never as an
    /// inlined base64 `data:` URL (lazy image-bytes pipeline, Stage 2.1).
    #[test]
    fn picture_element_serializes_path_not_data_url() {
        let pic = PictureElement {
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
            fill_rect: None,
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
        let mut zip = zip::ZipArchive::new(cursor).unwrap();
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
        let mut zip = zip::ZipArchive::new(cursor).unwrap();
        let mut xml = String::new();
        zip.by_name("ppt/charts/chartEx1.xml")
            .unwrap()
            .read_to_string(&mut xml)
            .unwrap();
        let theme = HashMap::new();
        let result = parse_chartex(&xml, None, &theme);
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
        let mut zip = zip::ZipArchive::new(cursor).unwrap();
        let mut slide_xml = String::new();
        zip.by_name("ppt/slides/slide8.xml")
            .unwrap()
            .read_to_string(&mut slide_xml)
            .unwrap();

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
        let mut zip = zip::ZipArchive::new(cursor).unwrap();

        let mut rels_xml = String::new();
        zip.by_name("ppt/slides/_rels/slide8.xml.rels")
            .unwrap()
            .read_to_string(&mut rels_xml)
            .unwrap();
        let rels = parse_rels(&rels_xml);
        println!("rels: {:?}", rels);

        let chart_path = resolve_path("ppt/slides", "../charts/chartEx1.xml");
        println!("chart_path: {}", chart_path);

        let result = read_zip_str(&mut zip, &chart_path);
        println!("read_zip_str ok: {}", result.is_ok());

        if let Ok(chart_xml) = result {
            let theme = HashMap::new();
            let r = parse_chartex(&chart_xml, None, &theme);
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

    /// ECMA-376 §21.1.2.3.10 — strike="dblStrike" produces strike_double=true,
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

    /// ECMA-376 §21.1.2.3.13 — cap="all" / "small" are passed through;
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

    /// ECMA-376 §21.1.2.3.5 — rPr @spc encodes letter spacing in 100ths of a
    /// point; positive widens, negative tightens. Zero rounds away (None).
    #[test]
    fn test_parse_run_letter_spacing() {
        let theme = HashMap::new();
        let rels = HashMap::new();
        for (raw, expected) in [("100", Some(1.0)), ("-50", Some(-0.5)), ("0", None)] {
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
    /// fillRect, and the alphaModFix alpha. Mirrors sample-12 slide1's background.
    #[test]
    fn test_parse_background_blip_fill_stretch() {
        let xml = r#"<p:cSld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <p:bg><p:bgPr>
                <a:blipFill>
                    <a:blip r:embed="rId2"><a:alphaModFix amt="80000"/></a:blip>
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
                fill_rect,
                tile,
                alpha,
                duotone: _,
            } => {
                assert_eq!(image_path, "ppt/media/image1.jpeg");
                assert_eq!(mime_type, "image/jpeg");
                let fr = fill_rect.expect("fillRect should be present");
                assert!((fr.t - (-0.09)).abs() < 1e-9, "t={}", fr.t);
                assert!((fr.b - (-0.09)).abs() < 1e-9, "b={}", fr.b);
                assert!(is_zero_f64(&fr.l) && is_zero_f64(&fr.r));
                assert!(tile.is_none(), "stretch fill must not carry tile");
                assert!((alpha.expect("alpha") - 0.8).abs() < 1e-6);
            }
            other => panic!("expected Fill::Image, got {other:?}"),
        }
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
                assert_eq!(t.tx, 457_200);
                assert_eq!(t.ty, -228_600);
                assert!((t.sx - 0.5).abs() < 1e-9, "sx={}", t.sx);
                assert!((t.sy - 0.75).abs() < 1e-9, "sy={}", t.sy);
                assert_eq!(t.flip, "xy");
                assert_eq!(t.algn, "ctr");
            }
            other => panic!("expected Fill::Image, got {other:?}"),
        }
    }

    /// §20.1.8.58 defaults: a bare `<a:tile/>` yields tx/ty=0, sx/sy=1.0
    /// (100% native size), flip="none", algn="tl".
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
                assert_eq!(t.tx, 0);
                assert_eq!(t.ty, 0);
                assert!((t.sx - 1.0).abs() < 1e-9);
                assert!((t.sy - 1.0).abs() < 1e-9);
                assert_eq!(t.flip, "none");
                assert_eq!(t.algn, "tl");
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
        let mut zip = zip::ZipArchive::new(cursor).unwrap();

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
        let mut zip_ok = zip::ZipArchive::new(cursor_ok).unwrap();
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

    /// ECMA-376 §21.1.2.3.16 — underline_style carries non-default underline
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

    /// ECMA-376 §21.1.2.3.20 — rPr > uFill > solidFill yields a per-run
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

    /// ECMA-376 §21.1.2.3.7 — rPr > ea sets a separate East Asian font.
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
                <outerShdw blurRad="50800" dist="38100" dir="2700000"><srgbClr val="000000"><alpha val="40000"/></srgbClr></outerShdw>
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
        let mut zip = zip::ZipArchive::new(cursor).unwrap();
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
        let mut zip = zip::ZipArchive::new(cursor).unwrap();
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
        let mut zip = zip::ZipArchive::new(cursor).unwrap();

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
            let mut zip = zip::ZipArchive::new(cursor).unwrap();
            parse_layout(
                layout,
                &m_f64,
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
              <a:tcTxStyle>
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
            solid(&def.whole_fill).as_deref(),
            Some("FFFFFF"),
            "wholeTbl fill should be lt1 white"
        );
        assert_eq!(
            solid(&def.first_row_fill).as_deref(),
            Some("B83903"),
            "firstRow header fill should be accent2 orange"
        );
        // band1H = accent2 + `<a:tint val="20000">`. Table styles use the literal
        // ECMA-376 tint (val·input + (1-val)·white), giving a near-white wash —
        // NOT the saturated linear-lerp. 0.2·B83903 + 0.8·white = F1D7CD.
        assert_eq!(
            solid(&def.band1h_fill).as_deref(),
            Some("F1D7CD"),
            "band1H tint should be the literal near-white wash, not a saturated lerp"
        );

        // Text colours from tcTxStyle.
        assert_eq!(
            def.whole_text_color.as_deref(),
            Some("000000"),
            "wholeTbl text colour should be dk1 black"
        );
        assert_eq!(
            def.first_row_text_color.as_deref(),
            Some("FFFFFF"),
            "firstRow header text colour should be lt1 white"
        );

        // firstRow `<a:tcTxStyle b="on">` → bold header.
        assert_eq!(
            def.first_row_bold,
            Some(true),
            "firstRow header should be bold from tcTxStyle b=on"
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
        let mut zip = zip::ZipArchive::new(cursor).unwrap();
        let mut parse = |body_pr: &str| -> TextBody {
            let xml = format!(
                r#"<txBody xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">{body_pr}<p><r><t>x</t></r></p></txBody>"#
            );
            let doc = roxmltree::Document::parse(&xml).unwrap();
            parse_text_body(
                doc.root_element(),
                &theme,
                &rels,
                None,               // inherited_font_size
                [None; 9],          // inherited_level_font_sizes
                Default::default(), // inherited_level_indents
                &empty_level_bullets(),
                None, // inherited_bold
                None, // inherited_italic
                None, // inherited_caps
                None, // inherited_anchor
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
        let mut zip = zip::ZipArchive::new(cursor).unwrap();
        let mut parse = |body_pr: &str| -> TextBody {
            let xml = format!(
                r#"<txBody xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">{body_pr}<p><r><t>x</t></r></p></txBody>"#
            );
            let doc = roxmltree::Document::parse(&xml).unwrap();
            parse_text_body(
                doc.root_element(),
                &theme,
                &rels,
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
        let mut zip = zip::ZipArchive::new(cursor).unwrap();
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
        let mut zip = zip::ZipArchive::new(cursor).unwrap();
        let mut parse_para = |p_pr: &str| -> Paragraph {
            let xml = format!(
                r#"<txBody xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><p>{p_pr}<r><t>a</t></r></p></txBody>"#
            );
            let doc = roxmltree::Document::parse(&xml).unwrap();
            let mut tb = parse_text_body(
                doc.root_element(),
                &theme,
                &rels,
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
            let mut zip = zip::ZipArchive::new(cursor).unwrap();
            parse_table(tbl, &t, &theme, &rels, &mut zip).unwrap()
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

        let sp3d = parse_sp3d(sppr).expect("sp3d should parse");
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

        let sp3d = parse_sp3d(sppr).unwrap();
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
        assert!(parse_sp3d(sppr).is_none());
    }

    // ===== sp3d contour colour (ECMA-376 §20.1.5.12 contourClr) =====

    #[test]
    fn test_parse_sp3d_contour_clr_slide3() {
        // The exact sp3d from sample-11 slide 3: contourW + grey contourClr.
        let xml = r#"<root
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:spPr>
            <a:sp3d contourW="6350" prstMaterial="matte">
              <a:bevelT w="101600" h="101600"/>
              <a:contourClr><a:srgbClr val="969696"/></a:contourClr>
            </a:sp3d>
          </p:spPr>
        </root>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let sppr = parse_sppr_frag(&doc);
        let sp3d = parse_sp3d(sppr).expect("sp3d should parse");
        assert_eq!(sp3d.contour_w, 6350);
        assert_eq!(sp3d.contour_clr.as_deref(), Some("969696"));
        let json = serde_json::to_string(&sp3d).unwrap();
        assert!(json.contains("\"contourClr\":\"969696\""), "{json}");
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
        let sp3d = parse_sp3d(sppr).unwrap();
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
    fn build_master_sp_pptx(layout_show_master_sp: Option<bool>) -> Vec<u8> {
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
  </p:spTree></p:cSld>
</p:sldLayout>"#
        );

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

    /// §19.3.1.38/§19.3.1.39: a master spTree picture (non-placeholder) is
    /// composited onto the slide. Without the fix the master spTree is dropped
    /// and the slide has no elements.
    #[test]
    fn master_sptree_pic_appears_on_slide() {
        let data = build_master_sp_pptx(None);
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
    }

    /// §19.3.1.39: a layout with showMasterSp="0" suppresses the master's
    /// decorative shapes for slides using that layout.
    #[test]
    fn master_sptree_hidden_when_layout_show_master_sp_false() {
        let data = build_master_sp_pptx(Some(false));
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

    /// showMasterSp="1" (explicit true) on the layout keeps master shapes —
    /// guards against an inverted boolean parse.
    #[test]
    fn master_sptree_shown_when_layout_show_master_sp_true() {
        let data = build_master_sp_pptx(Some(true));
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
        let mut zip = zip::ZipArchive::new(cursor).unwrap();

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
        let svg_bytes = extract_image(&data, "ppt/media/image2.svg", None, None, None)
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
        let mut zip = zip::ZipArchive::new(cursor).unwrap();
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
        let mut zip = zip::ZipArchive::new(cursor).unwrap();
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
        let mut zip = zip::ZipArchive::new(cursor).unwrap();
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
        let mut zip = zip::ZipArchive::new(cursor).unwrap();

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
        let svg_bytes = extract_image(&data, "ppt/media/image2.svg", None, None, None)
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
        let subpaths = parse_cust_geom(doc.root_element());
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
        let subpaths = parse_cust_geom(doc.root_element());
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
}
