use rmcp::{handler::server::wrapper::Parameters, tool};
use schemars::JsonSchema;
use serde::Deserialize;
use serde_json::Value;
use std::fs;

// ─── Parameter types ─────────────────────────────────────────────────────────

#[derive(Debug, Deserialize, JsonSchema)]
pub struct PptxPathParam {
    /// Absolute path to the PPTX file
    pub path: String,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct PptxSearchParam {
    /// Absolute path to the PPTX file
    pub path: String,
    /// Case-insensitive substring to search for across all slide text
    pub query: String,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct PptxSlideParam {
    /// Absolute path to the PPTX file
    pub path: String,
    /// 0-based slide index
    pub slide_index: usize,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct PptxTextParam {
    /// Absolute path to the PPTX file
    pub path: String,
    /// 0-based slide index; omit to extract text from all slides
    pub slide_index: Option<usize>,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct PptxElementParam {
    /// Absolute path to the PPTX file
    pub path: String,
    /// 0-based slide index
    pub slide_index: usize,
    /// 0-based index into the slide's elements array (matches `pptx_get_slide_structure`)
    pub element_index: usize,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct PptxOptSlideParam {
    /// Absolute path to the PPTX file
    pub path: String,
    /// 0-based slide index; omit to scan every slide
    pub slide_index: Option<usize>,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct PptxPicturesParam {
    /// Absolute path to the PPTX file
    pub path: String,
    /// 0-based slide index; omit to scan every slide
    pub slide_index: Option<usize>,
    /// When true include the base64 `dataUrl` for each picture. Defaults to
    /// false because picture bytes can be large.
    #[serde(default)]
    pub include_data_url: bool,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct PptxRelationsParam {
    /// Absolute path to the PPTX file
    pub path: String,
    /// 0-based slide index
    pub slide_index: usize,
    /// Geometry tolerance in EMU for endpoint / alignment matching. EMU is the
    /// PPTX coordinate unit (914400 EMU = 1 inch). Default ≈ 50 000 EMU
    /// (~5.5 pt) — generous enough to absorb floating-point drift on snapped
    /// connectors, tight enough to avoid false-positive alignments. Increase
    /// for hand-drawn slides, decrease for tightly-snapped templates.
    #[serde(default)]
    pub tolerance_emu: Option<i64>,
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

fn read_file(path: &str) -> Result<Vec<u8>, String> {
    fs::read(path).map_err(|e| format!("Cannot read '{}': {}", path, e))
}

fn extract_text_runs(node: &Value, out: &mut String) {
    // pptx-parser serializes TextRun as a tagged enum with `rename_all =
    // "camelCase"` — variants become "text" (TextRun::Text) and "break"
    // (TextRun::Break). Earlier revisions matched "textRun" / "run", which
    // never fire and silently produced empty extractions.
    match node["type"].as_str().unwrap_or("") {
        "text" => {
            if let Some(t) = node["text"].as_str() {
                out.push_str(t);
            }
        }
        "break" => {
            // ECMA-376 §21.1.2.2.1 — intra-paragraph <a:br/>. Map to a newline
            // so multi-line shape text isn't collapsed into a single line.
            out.push('\n');
        }
        _ => {}
    }
    // Recurse into common container fields (paragraph.runs, etc.).
    for key in &["runs", "paragraphs", "elements", "children"] {
        if let Some(arr) = node[key].as_array() {
            for child in arr {
                extract_text_runs(child, out);
            }
        }
    }
}

/// Title extraction backed by the parser-emitted `placeholderType`. Looks for
/// shapes whose `placeholderType` is "title" or "ctrTitle"
/// (ECMA-376 §19.7.10), falling back to the first shape with non-empty text
/// for slides that don't carry an explicit title placeholder (e.g. blank
/// layout, decorative slides).
fn slide_title(slide: &Value) -> Option<String> {
    let elements = slide["elements"].as_array()?;

    let read_text = |el: &Value| -> String {
        let mut text = String::new();
        if let Some(tb) = el.get("textBody") {
            if let Some(paras) = tb["paragraphs"].as_array() {
                for para in paras {
                    extract_text_runs(para, &mut text);
                    text.push(' ');
                }
            }
        }
        text.trim().to_string()
    };

    // First pass: prefer the explicit title placeholder.
    for el in elements {
        if el["type"].as_str() != Some("shape") {
            continue;
        }
        let Some(ph) = el["placeholderType"].as_str() else {
            continue;
        };
        if ph == "title" || ph == "ctrTitle" {
            let trimmed = read_text(el);
            if !trimmed.is_empty() {
                return Some(trimmed);
            }
        }
    }

    // Fallback: first non-empty shape text. Same heuristic the previous
    // implementation used; kept for slides without a title placeholder.
    for el in elements {
        if el["type"].as_str() != Some("shape") {
            continue;
        }
        let trimmed = read_text(el);
        if !trimmed.is_empty() {
            return Some(trimmed);
        }
    }
    None
}

fn extract_slide_text(slide: &Value) -> String {
    let mut out = String::new();
    if let Some(elements) = slide["elements"].as_array() {
        for el in elements {
            if let Some(tb) = el.get("textBody") {
                if let Some(paras) = tb["paragraphs"].as_array() {
                    for para in paras {
                        extract_text_runs(para, &mut out);
                        out.push('\n');
                    }
                }
            }
            // Table elements. TableCell holds its paragraphs under
            // `textBody.paragraphs`, not at the top level — the previous code
            // looked at `c["paragraphs"]` which always came back empty.
            if el["type"].as_str() == Some("table") {
                if let Some(rows) = el["rows"].as_array() {
                    for row in rows {
                        if let Some(cells) = row["cells"].as_array() {
                            let cell_texts: Vec<String> = cells
                                .iter()
                                .map(|c| {
                                    let mut t = String::new();
                                    if let Some(tb) = c.get("textBody") {
                                        if let Some(paras) = tb["paragraphs"].as_array() {
                                            for para in paras {
                                                extract_text_runs(para, &mut t);
                                            }
                                        }
                                    }
                                    t
                                })
                                .collect();
                            out.push_str(&cell_texts.join("\t"));
                            out.push('\n');
                        }
                    }
                }
            }
        }
    }
    out
}

fn slide_structure(slide: &Value) -> Value {
    let elements: Vec<Value> = slide["elements"]
        .as_array()
        .map(|els| {
            els.iter()
                .map(|el| {
                    let mut text = String::new();
                    if let Some(tb) = el.get("textBody") {
                        if let Some(paras) = tb["paragraphs"].as_array() {
                            for para in paras {
                                extract_text_runs(para, &mut text);
                            }
                        }
                    }
                    serde_json::json!({
                        "type": el["type"],
                        "placeholderType": el["placeholderType"],
                        "x": el["x"], "y": el["y"],
                        "width": el["width"], "height": el["height"],
                        "text": text.trim().to_string(),
                    })
                })
                .collect()
        })
        .unwrap_or_default();

    serde_json::json!({
        "index": slide["index"],
        "slideNumber": slide["slideNumber"],
        "elements": elements,
    })
}

// ─── Spatial geometry helpers (used by `pptx_get_shape_relations`) ───────────

/// A named alignment axis: its label plus an accessor extracting the edge/center
/// coordinate (EMU) used to group shapes that share that axis.
type AlignmentAxis = (&'static str, fn(&Bbox) -> i64);

/// EMU values are i64 in the parser output. Bounding box in EMU.
#[derive(Debug, Clone, Copy)]
struct Bbox {
    x: i64,
    y: i64,
    w: i64,
    h: i64,
}

impl Bbox {
    fn from_element(elem: &Value) -> Option<Self> {
        let x = elem["x"].as_i64()?;
        let y = elem["y"].as_i64()?;
        let w = elem["width"].as_i64()?;
        let h = elem["height"].as_i64()?;
        Some(Bbox { x, y, w, h })
    }

    fn right(&self) -> i64 {
        self.x + self.w
    }
    fn bottom(&self) -> i64 {
        self.y + self.h
    }
    fn cx(&self) -> i64 {
        self.x + self.w / 2
    }
    fn cy(&self) -> i64 {
        self.y + self.h / 2
    }

    /// Intersection-over-union with another bbox. Returns 0.0 when either box
    /// has zero area or they don't overlap.
    fn iou(&self, other: &Bbox) -> f64 {
        let ix1 = self.x.max(other.x);
        let iy1 = self.y.max(other.y);
        let ix2 = self.right().min(other.right());
        let iy2 = self.bottom().min(other.bottom());
        if ix2 <= ix1 || iy2 <= iy1 {
            return 0.0;
        }
        let inter = (ix2 - ix1) as f64 * (iy2 - iy1) as f64;
        let a = (self.w as f64) * (self.h as f64);
        let b = (other.w as f64) * (other.h as f64);
        let union = a + b - inter;
        if union <= 0.0 {
            0.0
        } else {
            inter / union
        }
    }

    /// True when `self` fully contains `other`, allowing `tol` EMU of slack
    /// on each edge.
    fn contains(&self, other: &Bbox, tol: i64) -> bool {
        other.x >= self.x - tol
            && other.y >= self.y - tol
            && other.right() <= self.right() + tol
            && other.bottom() <= self.bottom() + tol
            // Reject the trivial "same rectangle" case so identical bboxes
            // don't both contain each other.
            && (other.x > self.x - tol
                || other.y > self.y - tol
                || other.right() < self.right() + tol
                || other.bottom() < self.bottom() + tol)
    }
}

/// Returns true when the element looks like a line/connector. Recognises the
/// preset names PowerPoint emits for `<p:cxnSp>` and bare `<p:sp>` with line
/// preset (ECMA-376 §20.1.9.18 + connector geometries).
fn is_connector(geometry: &str) -> bool {
    matches!(
        geometry,
        "line"
            | "straightConnector1"
            | "bentConnector2"
            | "bentConnector3"
            | "bentConnector4"
            | "bentConnector5"
            | "curvedConnector2"
            | "curvedConnector3"
            | "curvedConnector4"
            | "curvedConnector5"
    )
}

/// Returns true when the arrow descriptor (`headEnd` / `tailEnd`) terminates
/// the line in something visible (i.e. not "none"). Matches ECMA-376 §20.1.10.46
/// `ST_LineEndType` values that draw a glyph.
fn arrow_is_directional(arrow: &Value) -> bool {
    matches!(
        arrow["type"].as_str(),
        Some("triangle") | Some("stealth") | Some("arrow") | Some("diamond") | Some("oval")
    )
}

/// Where on a shape's bbox a point sits. Uses `tol` EMU to absorb sub-pixel
/// snapping drift. Returns the closest side label, or `None` if the point is
/// further than `tol` from every side.
fn point_on_shape(point: (i64, i64), bbox: &Bbox, tol: i64) -> Option<&'static str> {
    let (px, py) = point;
    if px < bbox.x - tol || px > bbox.right() + tol || py < bbox.y - tol || py > bbox.bottom() + tol
    {
        return None;
    }
    let candidates: [(&'static str, i64); 8] = [
        ("topLeft", (px - bbox.x).abs() + (py - bbox.y).abs()),
        ("topRight", (px - bbox.right()).abs() + (py - bbox.y).abs()),
        (
            "bottomLeft",
            (px - bbox.x).abs() + (py - bbox.bottom()).abs(),
        ),
        (
            "bottomRight",
            (px - bbox.right()).abs() + (py - bbox.bottom()).abs(),
        ),
        ("top", (py - bbox.y).abs() + (px - bbox.cx()).abs() / 4),
        (
            "bottom",
            (py - bbox.bottom()).abs() + (px - bbox.cx()).abs() / 4,
        ),
        ("left", (px - bbox.x).abs() + (py - bbox.cy()).abs() / 4),
        (
            "right",
            (px - bbox.right()).abs() + (py - bbox.cy()).abs() / 4,
        ),
    ];
    let best = candidates.iter().min_by_key(|(_, d)| *d)?;
    Some(best.0)
}

/// Pick the bbox endpoint closest to `point` (only top-left corner or
/// bottom-right corner — the two diagonal endpoints of an unrotated line).
/// Returns "head" for the (x, y) endpoint and "tail" for the (x+w, y+h)
/// endpoint, which is the convention `pptx_get_shape_relations` uses to
/// orient `headEnd` / `tailEnd` arrows.
fn line_endpoints(bbox: &Bbox) -> [(&'static str, (i64, i64)); 2] {
    [
        ("head", (bbox.x, bbox.y)),
        ("tail", (bbox.right(), bbox.bottom())),
    ]
}

/// Pick the closest non-connector shape to `point` within `tol` EMU. Returns
/// `(shape_index, side_label)` when exactly one shape is in range; `None`
/// otherwise. Connectors themselves are skipped — we only resolve endpoints to
/// "real" shapes.
fn nearest_shape_to_point(
    point: (i64, i64),
    bboxes: &[(usize, Bbox, bool)],
    skip_index: usize,
    tol: i64,
) -> Option<(usize, &'static str)> {
    let mut best: Option<(usize, &'static str, i64)> = None;
    for (idx, bbox, is_conn) in bboxes {
        if *idx == skip_index || *is_conn {
            continue;
        }
        if let Some(side) = point_on_shape(point, bbox, tol) {
            // Distance from the actual side anchor for tie-breaking.
            let anchor_x = match side {
                "topLeft" | "left" | "bottomLeft" => bbox.x,
                "topRight" | "right" | "bottomRight" => bbox.right(),
                _ => bbox.cx(),
            };
            let anchor_y = match side {
                "topLeft" | "top" | "topRight" => bbox.y,
                "bottomLeft" | "bottom" | "bottomRight" => bbox.bottom(),
                _ => bbox.cy(),
            };
            let dist = (point.0 - anchor_x).abs() + (point.1 - anchor_y).abs();
            match best {
                Some((_, _, prev_dist)) if dist >= prev_dist => {}
                _ => best = Some((*idx, side, dist)),
            }
        }
    }
    best.map(|(i, s, _)| (i, s))
}

#[cfg(test)]
mod relations_tests {
    use super::*;

    fn b(x: i64, y: i64, w: i64, h: i64) -> Bbox {
        Bbox { x, y, w, h }
    }

    #[test]
    fn iou_disjoint_is_zero() {
        let a = b(0, 0, 100, 100);
        let other = b(200, 200, 50, 50);
        assert_eq!(a.iou(&other), 0.0);
    }

    #[test]
    fn iou_full_overlap_is_one() {
        let a = b(0, 0, 100, 100);
        assert!((a.iou(&a) - 1.0).abs() < 1e-9);
    }

    #[test]
    fn iou_half_overlap() {
        let a = b(0, 0, 100, 100);
        let other = b(50, 0, 100, 100);
        // intersection = 50*100 = 5000, union = 100*100 + 100*100 - 5000 = 15000
        let iou = a.iou(&other);
        assert!((iou - (5000.0 / 15000.0)).abs() < 1e-9);
    }

    #[test]
    fn contains_strict() {
        let outer = b(0, 0, 100, 100);
        let inner = b(10, 10, 50, 50);
        assert!(outer.contains(&inner, 0));
        assert!(!inner.contains(&outer, 0));
    }

    #[test]
    fn contains_rejects_identical() {
        let a = b(0, 0, 100, 100);
        assert!(!a.contains(&a, 0));
    }

    #[test]
    fn point_on_shape_corner() {
        let bb = b(100, 100, 200, 100);
        assert_eq!(point_on_shape((100, 100), &bb, 5), Some("topLeft"));
        assert_eq!(point_on_shape((300, 200), &bb, 5), Some("bottomRight"));
        // Outside tolerance
        assert_eq!(point_on_shape((50, 50), &bb, 5), None);
    }

    #[test]
    fn nearest_shape_skips_self_and_connectors() {
        let shapes = vec![
            (0usize, b(0, 0, 100, 100), false),
            (1, b(100, 0, 50, 50), true), // connector
            (2, b(150, 0, 100, 100), false),
        ];
        // Point at the right edge of shape 0
        let res = nearest_shape_to_point((100, 50), &shapes, 99, 10);
        assert_eq!(res, Some((0, "right")));
        // Point near shape 2's left side, but skipping shape 0
        let res = nearest_shape_to_point((150, 50), &shapes, 0, 10);
        assert_eq!(res, Some((2, "left")));
    }
}

// ─── Tool implementations ─────────────────────────────────────────────────────

pub struct PptxTools;

impl PptxTools {
    #[tool(description = "Return the number of slides and each slide's title from a PPTX file")]
    pub fn pptx_get_slides(Parameters(p): Parameters<PptxPathParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        let pres_json = match pptx_parser::parse_pptx_native(&data) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let pres: Value = match serde_json::from_str(&pres_json) {
            Ok(v) => v,
            Err(e) => return format!("Error: {}", e),
        };
        let slides = pres["slides"]
            .as_array()
            .map(|s| s.as_slice())
            .unwrap_or(&[]);
        let summary: Vec<Value> = slides
            .iter()
            .map(|s| {
                serde_json::json!({
                    "index": s["index"],
                    "slideNumber": s["slideNumber"],
                    "title": slide_title(s),
                })
            })
            .collect();
        serde_json::json!({
            "slideCount": slides.len(),
            "slides": summary,
        })
        .to_string()
    }

    #[tool(
        description = "Extract plain text from a PPTX file; optionally filter to a single slide by 0-based index"
    )]
    pub fn pptx_extract_text(Parameters(p): Parameters<PptxTextParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        let pres_json = match pptx_parser::parse_pptx_native(&data) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let pres: Value = match serde_json::from_str(&pres_json) {
            Ok(v) => v,
            Err(e) => return format!("Error: {}", e),
        };
        let slides = pres["slides"]
            .as_array()
            .map(|s| s.as_slice())
            .unwrap_or(&[]);

        if let Some(idx) = p.slide_index {
            let slide = match slides.get(idx) {
                Some(s) => s,
                None => {
                    return format!(
                        "Error: slide index {} out of range (total: {})",
                        idx,
                        slides.len()
                    )
                }
            };
            return extract_slide_text(slide);
        }

        let mut out = String::new();
        for (i, slide) in slides.iter().enumerate() {
            out.push_str(&format!("=== Slide {} ===\n", i + 1));
            out.push_str(&extract_slide_text(slide));
            out.push('\n');
        }
        out
    }

    #[tool(
        description = "Return the structure (elements with position, size, text) of a single slide"
    )]
    pub fn pptx_get_slide_structure(Parameters(p): Parameters<PptxSlideParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        let pres_json = match pptx_parser::parse_pptx_native(&data) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let pres: Value = match serde_json::from_str(&pres_json) {
            Ok(v) => v,
            Err(e) => return format!("Error: {}", e),
        };
        let slides = pres["slides"]
            .as_array()
            .map(|s| s.as_slice())
            .unwrap_or(&[]);
        let slide = match slides.get(p.slide_index) {
            Some(s) => s,
            None => {
                return format!(
                    "Error: slide index {} out of range (total: {})",
                    p.slide_index,
                    slides.len()
                )
            }
        };
        serde_json::to_string(&slide_structure(slide)).unwrap_or_else(|e| format!("Error: {}", e))
    }

    #[tool(
        description = "Search for a substring across all text in a PPTX file; returns matching slide numbers and the text snippets that matched"
    )]
    pub fn pptx_search_text(Parameters(p): Parameters<PptxSearchParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        let pres_json = match pptx_parser::parse_pptx_native(&data) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let pres: Value = match serde_json::from_str(&pres_json) {
            Ok(v) => v,
            Err(e) => return format!("Error: {}", e),
        };
        let slides = pres["slides"]
            .as_array()
            .map(|s| s.as_slice())
            .unwrap_or(&[]);
        let query_lower = p.query.to_lowercase();
        let mut matches: Vec<Value> = Vec::new();

        for slide in slides {
            let slide_index = slide["index"].as_u64().unwrap_or(0) as usize;
            let slide_number = slide["slideNumber"].as_u64().unwrap_or(0) as usize;

            if let Some(elements) = slide["elements"].as_array() {
                for el in elements {
                    // Collect all text from this element
                    let mut element_text = String::new();
                    if let Some(tb) = el.get("textBody") {
                        if let Some(paras) = tb["paragraphs"].as_array() {
                            for para in paras {
                                extract_text_runs(para, &mut element_text);
                                element_text.push('\n');
                            }
                        }
                    }
                    // Table cells. Same fix as `extract_slide_text`: paragraphs
                    // live under `cell.textBody.paragraphs`, not the top level.
                    if el["type"].as_str() == Some("table") {
                        if let Some(rows) = el["rows"].as_array() {
                            for row in rows {
                                if let Some(cells) = row["cells"].as_array() {
                                    for cell in cells {
                                        if let Some(tb) = cell.get("textBody") {
                                            if let Some(paras) = tb["paragraphs"].as_array() {
                                                for para in paras {
                                                    extract_text_runs(para, &mut element_text);
                                                }
                                            }
                                        }
                                        element_text.push('\t');
                                    }
                                }
                                element_text.push('\n');
                            }
                        }
                    }

                    if element_text.to_lowercase().contains(&query_lower) {
                        matches.push(serde_json::json!({
                            "slideIndex": slide_index,
                            "slideNumber": slide_number,
                            "elementType": el["type"],
                            "placeholderType": el["placeholderType"],
                            "text": element_text.trim(),
                        }));
                    }
                }
            }
        }

        serde_json::json!({
            "query": p.query,
            "matchCount": matches.len(),
            "matches": matches,
        })
        .to_string()
    }

    #[tool(
        description = "Return one slide element's full detail by slide and element index. `elementIndex` matches the element index in `pptx_get_slide_structure`. Includes shapes, pictures, charts, tables, geometry, position/size, rotation/flip, fill, stroke, effects, and the text body when present"
    )]
    pub fn pptx_get_element(Parameters(p): Parameters<PptxElementParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        let pres_json = match pptx_parser::parse_pptx_native(&data) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let pres: Value = match serde_json::from_str(&pres_json) {
            Ok(v) => v,
            Err(e) => return format!("Error: {}", e),
        };
        let slides = pres["slides"]
            .as_array()
            .map(|s| s.as_slice())
            .unwrap_or(&[]);
        let slide = match slides.get(p.slide_index) {
            Some(s) => s,
            None => {
                return format!(
                    "Error: slide index {} out of range (total: {})",
                    p.slide_index,
                    slides.len()
                )
            }
        };
        let elements = slide["elements"]
            .as_array()
            .map(|s| s.as_slice())
            .unwrap_or(&[]);
        let element = match elements.get(p.element_index) {
            Some(e) => e,
            None => {
                return format!(
                    "Error: element index {} out of range (slide elements: {})",
                    p.element_index,
                    elements.len()
                )
            }
        };
        let mut out = element.clone();
        if let Some(obj) = out.as_object_mut() {
            obj.insert("slideIndex".into(), Value::from(p.slide_index));
            obj.insert("elementIndex".into(), Value::from(p.element_index));
        }
        out.to_string()
    }

    #[tool(
        description = "List all charts on a slide (or every slide when `slide_index` is omitted). Each entry exposes type, position, title, categories, and series (with values)"
    )]
    pub fn pptx_get_charts(Parameters(p): Parameters<PptxOptSlideParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        let pres_json = match pptx_parser::parse_pptx_native(&data) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let pres: Value = match serde_json::from_str(&pres_json) {
            Ok(v) => v,
            Err(e) => return format!("Error: {}", e),
        };
        let slides = pres["slides"]
            .as_array()
            .map(|s| s.as_slice())
            .unwrap_or(&[]);
        let mut charts: Vec<Value> = Vec::new();
        for (slide_idx, slide) in slides.iter().enumerate() {
            if let Some(filter) = p.slide_index {
                if slide_idx != filter {
                    continue;
                }
            }
            let Some(elements) = slide["elements"].as_array() else {
                continue;
            };
            for (shape_idx, el) in elements.iter().enumerate() {
                if el["type"].as_str() != Some("chart") {
                    continue;
                }
                charts.push(serde_json::json!({
                    "slideIndex": slide_idx,
                    "shapeIndex": shape_idx,
                    "type": el["chartType"],
                    "title": el["title"],
                    "position": {
                        "x": el["x"], "y": el["y"],
                        "width": el["width"], "height": el["height"],
                    },
                    "categories": el["categories"],
                    "series": el["series"],
                    "showLegend": el["showLegend"],
                    "legendPos": el["legendPos"],
                    "showDataLabels": el["showDataLabels"],
                }));
            }
        }
        serde_json::json!({ "charts": charts }).to_string()
    }

    #[tool(
        description = "List all tables on a slide (or every slide when `slide_index` is omitted). Each entry includes column widths, row heights, and per-cell content (textBody) plus colSpan/rowSpan/merge information"
    )]
    pub fn pptx_get_tables(Parameters(p): Parameters<PptxOptSlideParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        let pres_json = match pptx_parser::parse_pptx_native(&data) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let pres: Value = match serde_json::from_str(&pres_json) {
            Ok(v) => v,
            Err(e) => return format!("Error: {}", e),
        };
        let slides = pres["slides"]
            .as_array()
            .map(|s| s.as_slice())
            .unwrap_or(&[]);
        let mut tables: Vec<Value> = Vec::new();
        for (slide_idx, slide) in slides.iter().enumerate() {
            if let Some(filter) = p.slide_index {
                if slide_idx != filter {
                    continue;
                }
            }
            let Some(elements) = slide["elements"].as_array() else {
                continue;
            };
            for (shape_idx, el) in elements.iter().enumerate() {
                if el["type"].as_str() != Some("table") {
                    continue;
                }
                tables.push(serde_json::json!({
                    "slideIndex": slide_idx,
                    "shapeIndex": shape_idx,
                    "position": {
                        "x": el["x"], "y": el["y"],
                        "width": el["width"], "height": el["height"],
                    },
                    "cols": el["cols"],
                    "rows": el["rows"],
                }));
            }
        }
        serde_json::json!({ "tables": tables }).to_string()
    }

    #[tool(
        description = "List all picture elements on a slide (or every slide when `slide_index` is omitted). Returns metadata only by default; pass `include_data_url=true` to include the inline base64 bytes"
    )]
    pub fn pptx_get_pictures(Parameters(p): Parameters<PptxPicturesParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        let pres_json = match pptx_parser::parse_pptx_native(&data) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let pres: Value = match serde_json::from_str(&pres_json) {
            Ok(v) => v,
            Err(e) => return format!("Error: {}", e),
        };
        let slides = pres["slides"]
            .as_array()
            .map(|s| s.as_slice())
            .unwrap_or(&[]);
        let mut pictures: Vec<Value> = Vec::new();
        for (slide_idx, slide) in slides.iter().enumerate() {
            if let Some(filter) = p.slide_index {
                if slide_idx != filter {
                    continue;
                }
            }
            let Some(elements) = slide["elements"].as_array() else {
                continue;
            };
            for (shape_idx, el) in elements.iter().enumerate() {
                if el["type"].as_str() != Some("picture") {
                    continue;
                }
                let mut entry = serde_json::json!({
                    "slideIndex": slide_idx,
                    "shapeIndex": shape_idx,
                    "x": el["x"], "y": el["y"],
                    "width": el["width"], "height": el["height"],
                    "rotation": el["rotation"],
                    "flipH": el["flipH"], "flipV": el["flipV"],
                    "srcRect": el["srcRect"],
                    "alpha": el["alpha"],
                    "clipAdjust": el["clipAdjust"],
                });
                if p.include_data_url {
                    if let Some(obj) = entry.as_object_mut() {
                        obj.insert("dataUrl".into(), el["dataUrl"].clone());
                    }
                }
                pictures.push(entry);
            }
        }
        serde_json::json!({ "pictures": pictures }).to_string()
    }

    #[tool(
        description = "Return presentation-level metadata: slide width/height (EMU), default text color, theme major/minor fonts, and hyperlink colors"
    )]
    pub fn pptx_get_presentation_meta(Parameters(p): Parameters<PptxPathParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        let pres_json = match pptx_parser::parse_pptx_native(&data) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let pres: Value = match serde_json::from_str(&pres_json) {
            Ok(v) => v,
            Err(e) => return format!("Error: {}", e),
        };
        let slide_count = pres["slides"].as_array().map(|s| s.len()).unwrap_or(0);
        serde_json::json!({
            "slideWidth": pres["slideWidth"],
            "slideHeight": pres["slideHeight"],
            "slideCount": slide_count,
            "defaultTextColor": pres["defaultTextColor"],
            "majorFont": pres["majorFont"],
            "minorFont": pres["minorFont"],
            "hlinkColor": pres["hlinkColor"],
            "folHlinkColor": pres["folHlinkColor"],
        })
        .to_string()
    }

    #[tool(
        description = "Convert a PPTX file to GitHub-flavoured markdown. Preserves textual structure (titles, bullets at correct nesting, tables, chart summaries, speaker notes, comments) and discards presentation details (geometry, fills, strokes, theme inheritance, positions). Designed for agents that need to *read* a deck efficiently — typical 10-30× token reduction vs. `pptx_get_slides` / `pptx_extract_text`. Lossy by design: when you need precise layout or styling, fall back to the structured tools (`pptx_get_element`, `pptx_get_slide_structure`, etc.)"
    )]
    pub fn pptx_to_markdown(Parameters(p): Parameters<PptxPathParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        match pptx_parser::to_markdown_native(&data) {
            Ok(md) => md,
            Err(e) => format!("Error: {}", e),
        }
    }

    #[tool(
        description = "Return speaker-notes text for one or all slides. Each entry: { slideIndex, slideNumber, notes }. Slides without a notesSlide part are omitted"
    )]
    pub fn pptx_get_notes(Parameters(p): Parameters<PptxOptSlideParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        let pres_json = match pptx_parser::parse_pptx_native(&data) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let pres: Value = match serde_json::from_str(&pres_json) {
            Ok(v) => v,
            Err(e) => return format!("Error: {}", e),
        };
        let slides = pres["slides"]
            .as_array()
            .map(|s| s.as_slice())
            .unwrap_or(&[]);
        let mut notes: Vec<Value> = Vec::new();
        for (slide_idx, slide) in slides.iter().enumerate() {
            if let Some(filter) = p.slide_index {
                if slide_idx != filter {
                    continue;
                }
            }
            let Some(text) = slide["notes"].as_str() else {
                continue;
            };
            notes.push(serde_json::json!({
                "slideIndex": slide_idx,
                "slideNumber": slide["slideNumber"],
                "notes": text,
            }));
        }
        serde_json::json!({ "notes": notes }).to_string()
    }

    #[tool(
        description = "Return legacy (non-threaded) slide comments. Each entry: { slideIndex, slideNumber, author?, date?, text }. Office365 modern threaded comments are not yet supported"
    )]
    pub fn pptx_get_comments(Parameters(p): Parameters<PptxOptSlideParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        let pres_json = match pptx_parser::parse_pptx_native(&data) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let pres: Value = match serde_json::from_str(&pres_json) {
            Ok(v) => v,
            Err(e) => return format!("Error: {}", e),
        };
        let slides = pres["slides"]
            .as_array()
            .map(|s| s.as_slice())
            .unwrap_or(&[]);
        let mut comments: Vec<Value> = Vec::new();
        for (slide_idx, slide) in slides.iter().enumerate() {
            if let Some(filter) = p.slide_index {
                if slide_idx != filter {
                    continue;
                }
            }
            let Some(arr) = slide["comments"].as_array() else {
                continue;
            };
            for c in arr {
                comments.push(serde_json::json!({
                    "slideIndex": slide_idx,
                    "slideNumber": slide["slideNumber"],
                    "author": c["author"],
                    "date": c["date"],
                    "text": c["text"],
                }));
            }
        }
        serde_json::json!({ "comments": comments }).to_string()
    }

    #[tool(
        description = "Infer geometric relations between shapes on a slide: connector hookups (with arrow direction when stroke ends are arrows), containment, overlap, axis-aligned alignment groups, and equal distribution. Detection is purely spatial — `confidence: \"inferred\"` flags this — until the parser exposes ECMA-376 §20.5.2.2 stCxn/endCxn references"
    )]
    pub fn pptx_get_shape_relations(Parameters(p): Parameters<PptxRelationsParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        let pres_json = match pptx_parser::parse_pptx_native(&data) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let pres: Value = match serde_json::from_str(&pres_json) {
            Ok(v) => v,
            Err(e) => return format!("Error: {}", e),
        };
        let slides = pres["slides"]
            .as_array()
            .map(|s| s.as_slice())
            .unwrap_or(&[]);
        let slide = match slides.get(p.slide_index) {
            Some(s) => s,
            None => {
                return format!(
                    "Error: slide index {} out of range (total: {})",
                    p.slide_index,
                    slides.len()
                )
            }
        };
        let elements = slide["elements"]
            .as_array()
            .map(|s| s.as_slice())
            .unwrap_or(&[]);

        let tol: i64 = p.tolerance_emu.unwrap_or(50_000).max(0);

        // Collect bbox + connector flag for each element. Skip elements with no
        // geometry (parser shouldn't emit any, but be defensive).
        let mut shape_summaries: Vec<Value> = Vec::with_capacity(elements.len());
        let mut bboxes: Vec<(usize, Bbox, bool)> = Vec::with_capacity(elements.len());
        for (idx, el) in elements.iter().enumerate() {
            let Some(bbox) = Bbox::from_element(el) else {
                continue;
            };
            let geometry = el["geometry"].as_str().unwrap_or("");
            let is_conn = el["type"].as_str() == Some("shape") && is_connector(geometry);
            // Pull a one-line text snippet for reader convenience.
            let mut text_snippet = String::new();
            if let Some(tb) = el.get("textBody") {
                if let Some(paras) = tb["paragraphs"].as_array() {
                    for para in paras.iter().take(3) {
                        extract_text_runs(para, &mut text_snippet);
                        text_snippet.push(' ');
                    }
                }
            }
            let trimmed = text_snippet.trim();
            shape_summaries.push(serde_json::json!({
                "shapeIndex": idx,
                "type": el["type"],
                "geometry": el["geometry"],
                "isConnector": is_conn,
                "bbox": {
                    "x": bbox.x, "y": bbox.y, "w": bbox.w, "h": bbox.h,
                },
                "text": if trimmed.is_empty() { Value::Null } else { Value::String(trimmed.to_string()) },
            }));
            bboxes.push((idx, bbox, is_conn));
        }

        let mut relations: Vec<Value> = Vec::new();

        // ── Connections (connector → resolved endpoints) ──────────────────
        for (idx, bbox, is_conn) in &bboxes {
            if !*is_conn {
                continue;
            }
            let element = &elements[*idx];
            let stroke = &element["stroke"];
            let head_arrow = stroke
                .get("headEnd")
                .map(arrow_is_directional)
                .unwrap_or(false);
            let tail_arrow = stroke
                .get("tailEnd")
                .map(arrow_is_directional)
                .unwrap_or(false);

            let endpoints = line_endpoints(bbox);
            let head_target = nearest_shape_to_point(endpoints[0].1, &bboxes, *idx, tol);
            let tail_target = nearest_shape_to_point(endpoints[1].1, &bboxes, *idx, tol);

            // Direction:
            //   tailEnd arrow only      → "headShape -> tailShape"
            //   headEnd arrow only      → "tailShape -> headShape"
            //   both                    → "bidirectional"
            //   neither                 → undirected
            let direction = match (head_arrow, tail_arrow) {
                (false, true) => Some("forward"),
                (true, false) => Some("reverse"),
                (true, true) => Some("bidirectional"),
                (false, false) => None,
            };

            // Skip when neither endpoint resolved — pure floating connector.
            if head_target.is_none() && tail_target.is_none() {
                continue;
            }

            relations.push(serde_json::json!({
                "kind": "connection",
                "connector": {
                    "shapeIndex": *idx,
                    "geometry": element["geometry"],
                    "headEnd": stroke.get("headEnd").cloned().unwrap_or(Value::Null),
                    "tailEnd": stroke.get("tailEnd").cloned().unwrap_or(Value::Null),
                },
                "head": head_target.map(|(i, side)| serde_json::json!({
                    "shapeIndex": i, "side": side,
                })).unwrap_or(Value::Null),
                "tail": tail_target.map(|(i, side)| serde_json::json!({
                    "shapeIndex": i, "side": side,
                })).unwrap_or(Value::Null),
                "direction": direction,
                "confidence": "inferred",
            }));
        }

        // ── Contains (real shapes only — connectors don't "contain") ──────
        for i in 0..bboxes.len() {
            let (idx_outer, outer_bb, outer_conn) = bboxes[i];
            if outer_conn {
                continue;
            }
            for (j, &(idx_inner, inner_bb, inner_conn)) in bboxes.iter().enumerate() {
                if i == j {
                    continue;
                }
                if inner_conn {
                    continue;
                }
                if outer_bb.contains(&inner_bb, tol) {
                    relations.push(serde_json::json!({
                        "kind": "contains",
                        "outer": idx_outer,
                        "inner": idx_inner,
                    }));
                }
            }
        }

        // ── Overlap (excluding contains) ──────────────────────────────────
        for i in 0..bboxes.len() {
            let (idx_a, a_bb, a_conn) = bboxes[i];
            if a_conn {
                continue;
            }
            for &(idx_b, b_bb, b_conn) in bboxes.iter().skip(i + 1) {
                if b_conn {
                    continue;
                }
                if a_bb.contains(&b_bb, tol) || b_bb.contains(&a_bb, tol) {
                    continue;
                }
                let iou = a_bb.iou(&b_bb);
                if iou > 0.0 {
                    relations.push(serde_json::json!({
                        "kind": "overlap",
                        "a": idx_a,
                        "b": idx_b,
                        "iou": iou,
                    }));
                }
            }
        }

        // ── Axis-aligned alignment groups (3+ shapes sharing an edge) ─────
        let real: Vec<(usize, Bbox)> = bboxes
            .iter()
            .filter(|(_, _, c)| !*c)
            .map(|(i, b, _)| (*i, *b))
            .collect();

        let alignment_axes: [AlignmentAxis; 3] = [
            ("top", |b| b.y),
            ("center", |b| b.cy()),
            ("bottom", |b| b.bottom()),
        ];
        for (axis, get_y) in alignment_axes {
            push_alignment_groups(&real, "alignH", axis, get_y, tol, &mut relations);
        }
        let alignment_axes_v: [AlignmentAxis; 3] = [
            ("left", |b| b.x),
            ("center", |b| b.cx()),
            ("right", |b| b.right()),
        ];
        for (axis, get_x) in alignment_axes_v {
            push_alignment_groups(&real, "alignV", axis, get_x, tol, &mut relations);
        }

        serde_json::json!({
            "slideIndex": p.slide_index,
            "toleranceEmu": tol,
            "shapes": shape_summaries,
            "relations": relations,
        })
        .to_string()
    }
}

/// Group `shapes` whose value under `key` falls within `tol` of each other,
/// emitting one `kind` relation per group of ≥ 3 shapes.
fn push_alignment_groups(
    shapes: &[(usize, Bbox)],
    kind: &str,
    axis: &str,
    key: fn(&Bbox) -> i64,
    tol: i64,
    out: &mut Vec<Value>,
) {
    let mut entries: Vec<(usize, i64)> = shapes.iter().map(|(i, b)| (*i, key(b))).collect();
    entries.sort_by_key(|(_, v)| *v);
    let mut i = 0;
    while i < entries.len() {
        let mut group: Vec<usize> = vec![entries[i].0];
        let pivot = entries[i].1;
        let mut j = i + 1;
        while j < entries.len() && (entries[j].1 - pivot).abs() <= tol {
            group.push(entries[j].0);
            j += 1;
        }
        if group.len() >= 3 {
            out.push(serde_json::json!({
                "kind": kind,
                "axis": axis,
                "shapes": group,
                "value": pivot,
            }));
        }
        i = j;
    }
}

#[cfg(test)]
mod sample_tests {
    use super::*;

    fn sample_path() -> String {
        format!(
            "{}/../pptx/public/demo/sample-1.pptx",
            env!("CARGO_MANIFEST_DIR")
        )
    }

    fn pp(path: &str) -> Parameters<PptxPathParam> {
        Parameters(PptxPathParam { path: path.into() })
    }

    #[test]
    fn pptx_get_slides_sample() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = PptxTools::pptx_get_slides(pp(&path));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        let count = v["slideCount"].as_u64().unwrap_or(0);
        assert!(count >= 1, "sample-1.pptx should have ≥1 slide");
    }

    #[test]
    fn pptx_get_presentation_meta_sample() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = PptxTools::pptx_get_presentation_meta(pp(&path));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        assert!(
            v["slideWidth"].as_i64().unwrap_or(0) > 0,
            "slideWidth should be > 0"
        );
        assert!(
            v["slideHeight"].as_i64().unwrap_or(0) > 0,
            "slideHeight should be > 0"
        );
    }

    #[test]
    fn pptx_get_shape_relations_sample() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = PptxTools::pptx_get_shape_relations(Parameters(PptxRelationsParam {
            path,
            slide_index: 0,
            tolerance_emu: None,
        }));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        assert!(
            v["shapes"].as_array().is_some(),
            "missing 'shapes' array: {out}"
        );
        assert!(
            v["relations"].as_array().is_some(),
            "missing 'relations' array: {out}"
        );
        assert_eq!(v["slideIndex"].as_u64(), Some(0));
    }

    #[test]
    fn pptx_get_charts_returns_array() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = PptxTools::pptx_get_charts(Parameters(PptxOptSlideParam {
            path,
            slide_index: None,
        }));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        assert!(v["charts"].as_array().is_some(), "missing 'charts'");
    }

    #[test]
    fn pptx_get_pictures_excludes_data_url_by_default() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = PptxTools::pptx_get_pictures(Parameters(PptxPicturesParam {
            path,
            slide_index: None,
            include_data_url: false,
        }));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        let pics = v["pictures"].as_array().expect("missing 'pictures'");
        for p in pics {
            assert!(
                p.get("dataUrl").is_none(),
                "dataUrl should be omitted by default"
            );
        }
    }

    #[test]
    fn pptx_invalid_path_returns_error_string() {
        let out = PptxTools::pptx_get_slides(pp("/nonexistent/x.pptx"));
        assert!(out.starts_with("Error:"), "expected error, got: {out}");
    }

    /// Regression for the run-tag bug where pptx-parser emits `type: "text"`
    /// but the helper used to match `"textRun" | "run"` and silently extracted
    /// nothing. Hard-asserts that sample-1.pptx slide-1 yields the visible
    /// title text, so anyone who loosens the matcher again will fail this.
    #[test]
    fn pptx_extract_text_returns_non_empty_for_sample() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = PptxTools::pptx_extract_text(Parameters(PptxTextParam {
            path: path.clone(),
            slide_index: Some(0),
        }));
        assert!(
            !out.starts_with("Error:"),
            "pptx_extract_text errored: {out}"
        );
        assert!(
            !out.trim().is_empty(),
            "pptx_extract_text returned empty for slide 0 — extract_text_runs is likely matching the wrong run.type tag again"
        );
    }

    #[test]
    fn pptx_get_slides_returns_titles() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = PptxTools::pptx_get_slides(pp(&path));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        let slides = v["slides"].as_array().expect("must have 'slides'");
        let any_title = slides
            .iter()
            .any(|s| s["title"].as_str().map(|t| !t.is_empty()).unwrap_or(false));
        assert!(
            any_title,
            "no slide reported a title — slide_title heuristic is broken"
        );
    }

    /// sample-1.pptx slide 8 is a pull-quote layout — its first shape is the
    /// decorative "“" glyph. Before placeholder_type was exposed by the parser
    /// the heuristic returned that quote. With placeholder_type the title
    /// resolver now skips non-title placeholders and either finds the real
    /// title placeholder or falls through to the first non-empty shape text.
    /// Either way the result must NOT be a single decorative character.
    #[test]
    fn pptx_get_slides_skips_decorative_quote_for_title() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = PptxTools::pptx_get_slides(pp(&path));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        let slides = v["slides"].as_array().expect("must have 'slides'");
        if let Some(s8) = slides.iter().find(|s| s["slideNumber"].as_u64() == Some(8)) {
            let title = s8["title"].as_str().unwrap_or("");
            assert!(
                title.chars().count() > 1,
                "slide 8 title was '{title}' — placeholder_type filter likely not applied"
            );
        }
    }

    #[test]
    fn pptx_get_notes_smoke() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = PptxTools::pptx_get_notes(Parameters(PptxOptSlideParam {
            path,
            slide_index: None,
        }));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        assert!(v["notes"].as_array().is_some(), "missing 'notes' array");
    }

    #[test]
    fn pptx_to_markdown_sample() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = PptxTools::pptx_to_markdown(pp(&path));
        assert!(!out.starts_with("Error:"), "errored: {out}");
        // Slide titles should appear as level-1 headings.
        assert!(
            out.contains("# STATE OF THE FOREST"),
            "missing slide-1 heading: {}",
            &out[..200.min(out.len())]
        );
        // Slide separator between slides.
        assert!(out.contains("\n---\n"), "missing slide separator");
        // Bold/italic markers from rich-text runs should be preserved.
        assert!(
            out.contains("**3.4%**") || out.contains("**+3.4%**"),
            "missing bold marker"
        );
        // Tables → pipe rows.
        assert!(out.contains("| Taxon |"), "biodiversity table missing");
        // Chart → markdown summary.
        assert!(out.contains("**Chart (line):"), "chart summary missing");
        // Compared with pptx_extract_text, markdown adds structure but keeps
        // size in the same order of magnitude. Sanity-check the bound so a
        // future bug that explodes the output (e.g. accidentally serializing
        // the full presentation) trips the test.
        let plain = PptxTools::pptx_extract_text(Parameters(PptxTextParam {
            path: path.clone(),
            slide_index: None,
        }));
        assert!(
            out.len() < plain.len() * 3,
            "markdown should be within 3× of plain text — got {} vs {}",
            out.len(),
            plain.len()
        );
    }

    #[test]
    fn pptx_get_comments_smoke() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = PptxTools::pptx_get_comments(Parameters(PptxOptSlideParam {
            path,
            slide_index: None,
        }));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        assert!(
            v["comments"].as_array().is_some(),
            "missing 'comments' array"
        );
    }

    #[test]
    fn pptx_element_index_out_of_range_errors() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = PptxTools::pptx_get_element(Parameters(PptxElementParam {
            path,
            slide_index: 0,
            element_index: 999_999,
        }));
        assert!(
            out.starts_with("Error:"),
            "expected out-of-range error, got: {out}"
        );
    }
}
