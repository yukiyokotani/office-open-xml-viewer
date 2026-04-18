use base64::{engine::general_purpose::STANDARD as B64, Engine as _};
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::io::{Cursor, Read};
use wasm_bindgen::prelude::*;

// ===========================
//  Public WASM entry points
// ===========================

#[wasm_bindgen]
pub fn parse_pptx(data: &[u8]) -> Result<String, JsValue> {
    console_error_panic_hook::set_once();
    let presentation = parse_presentation(data)
        .map_err(|e| JsValue::from_str(&format!("pptx-parser error: {e}")))?;
    serde_json::to_string(&presentation)
        .map_err(|e| JsValue::from_str(&format!("serialize error: {e}")))
}

// ===========================
//  Data types  (camelCase JSON → TypeScript)
// ===========================

#[derive(Serialize, Deserialize, Debug)]
#[serde(rename_all = "camelCase")]
struct Presentation {
    slide_width: i64,
    slide_height: i64,
    slides: Vec<Slide>,
    /// Default text color from theme dk1 (hex 6 chars, e.g. "383838").
    default_text_color: Option<String>,
}

#[derive(Serialize, Deserialize, Debug)]
#[serde(rename_all = "camelCase")]
struct Slide {
    index: usize,
    /// 1-based slide number (index + 1); used for slidenum field rendering
    slide_number: usize,
    background: Option<Fill>,
    elements: Vec<SlideElement>,
}

#[derive(Serialize, Deserialize, Debug)]
#[serde(tag = "type", rename_all = "camelCase")]
enum SlideElement {
    Shape(ShapeElement),
    Picture(PictureElement),
    Table(TableElement),
    Chart(ChartElement),
}

#[derive(Serialize, Deserialize, Debug)]
#[serde(rename_all = "camelCase")]
struct ChartSeriesData {
    name: String,
    values: Vec<Option<f64>>,
    color: Option<String>,
    /// Per-data-point colors (used for pie/doughnut charts). None if all points use series color.
    #[serde(skip_serializing_if = "Option::is_none")]
    data_point_colors: Option<Vec<Option<String>>>,
}

#[derive(Serialize, Deserialize, Debug)]
#[serde(rename_all = "camelCase")]
struct ChartElement {
    x: i64, y: i64, width: i64, height: i64,
    chart_type: String,
    title: Option<String>,
    categories: Vec<String>,
    series: Vec<ChartSeriesData>,
    val_max: Option<f64>,
    subtotal_indices: Vec<u32>,
    /// Whether to render data value labels on bars/segments
    show_data_labels: bool,
}

// ===== Table data model =====

#[derive(Serialize, Deserialize, Debug)]
#[serde(rename_all = "camelCase")]
struct TableElement {
    x: i64, y: i64, width: i64, height: i64,
    /// Column widths in EMU
    cols: Vec<i64>,
    rows: Vec<TableRow>,
}

#[derive(Serialize, Deserialize, Debug)]
#[serde(rename_all = "camelCase")]
struct TableRow {
    /// Row height in EMU
    height: i64,
    cells: Vec<TableCell>,
}

#[derive(Serialize, Deserialize, Debug)]
#[serde(rename_all = "camelCase")]
struct TableCell {
    text_body: Option<TextBody>,
    fill: Option<Fill>,
    border_l: Option<Stroke>,
    border_r: Option<Stroke>,
    border_t: Option<Stroke>,
    border_b: Option<Stroke>,
    /// Column span (gridSpan attribute)
    grid_span: u32,
    /// Row span
    row_span: u32,
    /// Horizontal merge continuation (cell has no content, covered by left neighbour)
    h_merge: bool,
    /// Vertical merge continuation
    v_merge: bool,
}

#[derive(Serialize, Deserialize, Debug)]
#[serde(rename_all = "camelCase")]
struct ShapeElement {
    x: i64,
    y: i64,
    width: i64,
    height: i64,
    rotation: f64,
    flip_h: bool,
    flip_v: bool,
    /// OOXML preset name (e.g. "rect", "ellipse") or "custGeom" when custom paths are used.
    geometry: String,
    fill: Option<Fill>,
    stroke: Option<Stroke>,
    text_body: Option<TextBody>,
    /// Default text color from p:style > fontRef (hex). Overrides renderer default
    /// when present; individual run colors still take precedence.
    default_text_color: Option<String>,
    /// Custom geometry paths (only set when geometry == "custGeom").
    /// Outer vec: one entry per <a:path>; inner vec: path commands with coords in [0,1].
    cust_geom: Option<Vec<Vec<PathCmd>>>,
    /// First adjustment value from prstGeom avLst (e.g. trapezoid inset).
    /// Value is in OOXML units (0–100000 range).
    adj: Option<f64>,
    /// Second adjustment value from prstGeom avLst (e.g. arrow-head width).
    adj2: Option<f64>,
    /// Third adjustment value from prstGeom avLst (e.g. callout tip x).
    adj3: Option<f64>,
    /// Fourth adjustment value from prstGeom avLst (e.g. callout tip y).
    adj4: Option<f64>,
    /// Drop shadow from spPr > effectLst > outerShdw (None if not present).
    shadow: Option<Shadow>,
}

#[derive(Serialize, Deserialize, Debug)]
#[serde(rename_all = "camelCase")]
struct PictureElement {
    x: i64,
    y: i64,
    width: i64,
    height: i64,
    rotation: f64,
    flip_h: bool,
    flip_v: bool,
    data_url: String,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
struct GradStop {
    /// 0.0–1.0
    position: f64,
    /// hex color (6 chars = opaque, 8 chars = RRGGBBAA with alpha)
    color: String,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
struct Shadow {
    /// hex color (6 chars)
    color: String,
    /// opacity 0.0–1.0
    alpha: f64,
    /// blur radius in EMU
    blur: i64,
    /// distance from shape in EMU
    dist: i64,
    /// direction in degrees, clockwise from East
    dir: f64,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(tag = "fillType", rename_all = "camelCase")]
enum Fill {
    Solid { color: String },
    None,
    #[serde(rename_all = "camelCase")]
    Gradient {
        stops: Vec<GradStop>,
        /// degrees, 0 = left→right, 90 = top→bottom
        angle: f64,
        /// "linear" | "radial"
        grad_type: String,
    },
}

#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
struct Stroke {
    color: String,
    width: i64,
    /// OOXML prstDash value: "dash", "dot", "dashDot", "lgDash", "lgDashDot", "sysDash", "sysDot", etc.
    #[serde(skip_serializing_if = "Option::is_none")]
    dash_style: Option<String>,
}

/// A single path command inside a custGeom pathLst.
/// Coordinates are normalised to [0.0, 1.0] relative to the path's w/h,
/// so the renderer can map them directly to shape-local pixel coordinates.
#[derive(Serialize, Deserialize, Debug)]
#[serde(tag = "cmd", rename_all = "camelCase")]
enum PathCmd {
    MoveTo { x: f64, y: f64 },
    LineTo { x: f64, y: f64 },
    /// Cubic Bézier: two control points + endpoint
    CubicBezTo { x1: f64, y1: f64, x2: f64, y2: f64, x: f64, y: f64 },
    /// Elliptical arc (all angles in degrees)
    ArcTo { wr: f64, hr: f64, st_ang: f64, sw_ang: f64 },
    Close,
}

#[derive(Serialize, Deserialize, Debug)]
#[serde(rename_all = "camelCase")]
struct TextBody {
    vertical_anchor: String,
    paragraphs: Vec<Paragraph>,
    default_font_size: Option<f64>,
    /// Inherited bold from layout/master lstStyle defRPr (None = not inherited)
    default_bold: Option<bool>,
    /// Inherited italic from layout/master lstStyle defRPr (None = not inherited)
    default_italic: Option<bool>,
    /// Text insets in EMU. Defaults: lIns=rIns=91440, tIns=bIns=45720
    l_ins: i64,
    r_ins: i64,
    t_ins: i64,
    b_ins: i64,
    /// Whether text wraps inside the bounding box ("square") or not ("none")
    wrap: String,
    /// Text direction from bodyPr vert attribute: "horz" | "vert" | "vert270" | "eaVert" etc.
    vert: String,
    /// Auto-fit mode from bodyPr: "sp" = spAutoFit (shape grows), "norm" = normAutoFit (font shrinks), "none" = noAutofit
    auto_fit: String,
}

/// Line spacing specification
#[derive(Serialize, Deserialize, Debug)]
#[serde(tag = "type", rename_all = "camelCase")]
enum SpaceLine {
    /// Percentage of the font height (val: e.g. 100000 = 100%, 150000 = 150%)
    Pct { val: f64 },
    /// Fixed points (val in pt)
    Pts { val: f64 },
}

/// Bullet / list-item marker for a paragraph
#[derive(Serialize, Deserialize, Debug)]
#[serde(tag = "type", rename_all = "camelCase")]
enum Bullet {
    /// Explicitly no bullet (buNone)
    None,
    /// No bullet element present – inherit from layout/master
    Inherit,
    /// Character bullet (buChar)
    #[serde(rename_all = "camelCase")]
    Char {
        #[serde(rename = "char")]
        ch: String,
        color: Option<String>,
        /// Size as % of text size (100.0 = same size)
        size_pct: Option<f64>,
        font_family: Option<String>,
    },
    /// Auto-numbered bullet (buAutoNum)
    #[serde(rename_all = "camelCase")]
    AutoNum {
        num_type: String,
        start_at: Option<u32>,
    },
}

/// A tab stop defined in a paragraph's pPr > tabLst.
#[derive(Serialize, Deserialize, Debug)]
#[serde(rename_all = "camelCase")]
struct TabStop {
    /// Position in EMU from the left edge of the text area (after lIns)
    pos: i64,
    /// Alignment: "l" | "r" | "ctr" | "dec"
    algn: String,
}

#[derive(Serialize, Deserialize, Debug)]
#[serde(rename_all = "camelCase")]
struct Paragraph {
    alignment: String,
    /// Left margin in EMU
    mar_l: i64,
    /// Right margin in EMU
    mar_r: i64,
    /// First-line indent in EMU (negative = hanging indent for bullets)
    indent: i64,
    space_before: Option<i64>,
    space_after: Option<i64>,
    space_line: Option<SpaceLine>,
    /// List nesting level (0–8)
    lvl: u32,
    bullet: Bullet,
    /// Paragraph-level default run properties (from pPr > defRPr)
    def_font_size: Option<f64>,
    def_color: Option<String>,
    def_bold: Option<bool>,
    def_italic: Option<bool>,
    def_font_family: Option<String>,
    /// Tab stops from pPr > tabLst
    tab_stops: Vec<TabStop>,
    runs: Vec<TextRun>,
}

#[derive(Serialize, Deserialize, Debug)]
#[serde(tag = "type", rename_all = "camelCase")]
enum TextRun {
    Text(TextRunData),
    Break,
}

#[derive(Serialize, Deserialize, Debug)]
#[serde(rename_all = "camelCase")]
struct TextRunData {
    text: String,
    /// None = not set (inherit from paragraph/body/layout defaults); Some(true/false) = explicit
    bold: Option<bool>,
    /// None = not set; Some(true/false) = explicit
    italic: Option<bool>,
    underline: bool,
    /// true when strike == "sngStrike" or "dblStrike"
    strikethrough: bool,
    font_size: Option<f64>,
    color: Option<String>,
    font_family: Option<String>,
    /// Baseline shift in thousandths of a point. Positive = superscript, negative = subscript.
    #[serde(skip_serializing_if = "Option::is_none")]
    baseline: Option<i32>,
    /// Set for OOXML field elements (e.g. "slidenum" for slide number fields)
    field_type: Option<String>,
}

// ===========================
//  ZIP helpers
// ===========================

type PptxZip<'a> = zip::ZipArchive<Cursor<&'a [u8]>>;

fn read_zip_str(zip: &mut PptxZip<'_>, path: &str) -> Result<String, Box<dyn std::error::Error>> {
    let mut file = zip
        .by_name(path)
        .map_err(|_| format!("missing ZIP entry: {path}"))?;
    let mut buf = String::new();
    file.read_to_string(&mut buf)?;
    Ok(buf)
}

fn read_zip_bytes(zip: &mut PptxZip<'_>, path: &str) -> Option<Vec<u8>> {
    let mut file = zip.by_name(path).ok()?;
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).ok()?;
    Some(buf)
}

// ===========================
//  Table style data model
// ===========================

/// Resolved fills and borders extracted from a single <a:tblStyle> definition.
#[derive(Debug, Clone, Default)]
struct TableStyleDef {
    whole_fill:        Option<Fill>,
    whole_inside_h:    Option<Stroke>,
    whole_inside_v:    Option<Stroke>,
    /// Outer top/bottom edge border (from wholeTbl tcBdr top/bottom)
    whole_outer_h:     Option<Stroke>,
    /// Outer left/right edge border (from wholeTbl tcBdr left/right)
    whole_outer_v:     Option<Stroke>,
    band1h_fill:       Option<Fill>,
    band2h_fill:       Option<Fill>,
    first_row_fill:    Option<Fill>,
    first_row_border_b: Option<Stroke>,
    last_row_fill:     Option<Fill>,
    first_col_fill:    Option<Fill>,
    last_col_fill:     Option<Fill>,
}

// ===========================
//  XML helpers (roxmltree)
// ===========================

fn child<'a, 'i>(
    node: roxmltree::Node<'a, 'i>,
    local: &str,
) -> Option<roxmltree::Node<'a, 'i>> {
    node.children()
        .find(|n| n.is_element() && n.tag_name().name() == local)
}

fn children_vec<'a, 'i>(
    node: roxmltree::Node<'a, 'i>,
    local: &str,
) -> Vec<roxmltree::Node<'a, 'i>> {
    node.children()
        .filter(|n| n.is_element() && n.tag_name().name() == local)
        .collect()
}

const R_NS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn attr(node: &roxmltree::Node<'_, '_>, local: &str) -> Option<String> {
    node.attributes()
        .find(|a| a.name() == local && a.namespace().is_none())
        .map(|a| a.value().to_owned())
}

/// Attribute in the r: (relationships) namespace — e.g. r:id, r:embed
fn attr_r(node: &roxmltree::Node<'_, '_>, local: &str) -> Option<String> {
    node.attributes()
        .find(|a| a.name() == local && a.namespace() == Some(R_NS))
        .map(|a| a.value().to_owned())
}

fn attr_i64(node: &roxmltree::Node<'_, '_>, local: &str) -> Option<i64> {
    attr(node, local)?.parse().ok()
}

fn attr_f64(node: &roxmltree::Node<'_, '_>, local: &str) -> Option<f64> {
    attr(node, local)?.parse().ok()
}

// ===========================
//  Relationships helpers
// ===========================

/// id → target  (used for image/slide lookups by rId)
fn parse_rels(xml: &str) -> HashMap<String, String> {
    let doc = match roxmltree::Document::parse(xml) {
        Ok(d) => d,
        Err(_) => return HashMap::new(),
    };
    let mut map = HashMap::new();
    for rel in doc.root_element().children().filter(|n| n.is_element()) {
        if let (Some(id), Some(target)) = (attr(&rel, "Id"), attr(&rel, "Target")) {
            map.insert(id, target);
        }
    }
    map
}

/// Find the Target of the first relationship whose Type ends with `type_suffix`.
fn find_rel_target_by_type(rels_xml: &str, type_suffix: &str) -> Option<String> {
    let doc = roxmltree::Document::parse(rels_xml).ok()?;
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
fn resolve_path(base_dir: &str, target: &str) -> String {
    let mut parts: Vec<&str> = base_dir.split('/').collect();
    for seg in target.split('/') {
        match seg {
            ".." => { parts.pop(); }
            "." | "" => {}
            s => parts.push(s),
        }
    }
    parts.join("/")
}

// ===========================
//  Color parsing
// ===========================

/// Parse the color scheme from a theme XML file.
/// Returns a map: scheme slot name (e.g. "dk1", "lt1", "acc1") → hex string.
fn parse_theme_colors(xml: &str) -> HashMap<String, String> {
    let mut map = HashMap::new();
    let doc = match roxmltree::Document::parse(xml) {
        Ok(d) => d,
        Err(_) => return map,
    };
    let root = doc.root_element();

    let clr_scheme = match root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "clrScheme")
    {
        Some(n) => n,
        None => return map,
    };
    // Each child of clrScheme is a slot: dk1, lt1, dk2, lt2, acc1–acc6, hlink, folHlink
    for slot in clr_scheme.children().filter(|n| n.is_element()) {
        let name = slot.tag_name().name().to_owned();
        // The slot contains exactly one color child; parse it without theme context
        for c in slot.children().filter(|n| n.is_element()) {
            let hex = match c.tag_name().name() {
                "srgbClr" => attr(&c, "val"),
                "sysClr"  => attr(&c, "lastClr"),
                "prstClr" => preset_color(attr(&c, "val").unwrap_or_default().as_str()),
                _         => None,
            };
            if let Some(h) = hex {
                map.insert(name, h);
                break;
            }
        }
    }

    // Parse font scheme: majorFont (+mj-lt, +mj-ea, +mj-cs) and minorFont (+mn-lt, +mn-ea, +mn-cs)
    // Store as special keys in the theme map so the renderer can resolve +mj-lt → actual typeface.
    if let Some(font_scheme) = root.descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "fontScheme")
    {
        let pairs: &[(&str, &[&str])] = &[
            ("majorFont", &["+mj-lt", "+mj-ea", "+mj-cs"]),
            ("minorFont", &["+mn-lt", "+mn-ea", "+mn-cs"]),
        ];
        let scripts = ["latin", "ea", "cs"];
        for (element_name, keys) in pairs {
            if let Some(font_node) = child(font_scheme, element_name) {
                for (script, key) in scripts.iter().zip(keys.iter()) {
                    if let Some(typeface) = child(font_node, script)
                        .and_then(|n| attr(&n, "typeface"))
                    {
                        if !typeface.is_empty() {
                            map.insert(key.to_string(), typeface.to_string());
                        }
                    }
                }
            }
        }
    }

    map
}

/// Resolve a theme typeface reference (e.g. "+mj-lt") to the actual font family name.
/// If the typeface starts with '+' and has a matching entry in the theme map (added by
/// parse_theme_colors from the fontScheme), returns the resolved name; otherwise returns
/// the original string unchanged.
fn resolve_theme_typeface(typeface: &str, theme: &HashMap<String, String>) -> String {
    if typeface.starts_with('+') {
        if let Some(resolved) = theme.get(typeface) {
            return resolved.clone();
        }
    }
    typeface.to_string()
}

/// Resolve a color node (solidFill child / run rPr child) to a hex string.
/// Handles srgbClr, sysClr, prstClr, and schemeClr (with transform support).
fn parse_color_node(
    node: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<String> {
    for c in node.children().filter(|n| n.is_element()) {
        match c.tag_name().name() {
            "srgbClr" => {
                let hex = attr(&c, "val")?;
                return Some(apply_color_transforms(&hex, c));
            }
            "sysClr"  => {
                let hex = attr(&c, "lastClr")?;
                return Some(apply_color_transforms(&hex, c));
            }
            "prstClr" => return preset_color(attr(&c, "val")?.as_str()),
            "schemeClr" => {
                let scheme_name = attr(&c, "val")?;
                // OOXML semantic aliases → canonical theme slot names
                let canonical: &str = match scheme_name.as_str() {
                    "tx1" | "dk1"   => "dk1",
                    "tx2" | "dk2"   => "dk2",
                    "bg1" | "lt1"   => "lt1",
                    "bg2" | "lt2"   => "lt2",
                    // phClr = "placeholder color" (inherits from layout).
                    // Approximate as the primary dark text color.
                    "phClr"         => "dk1",
                    other           => other,
                };
                let base_hex = theme.get(canonical)?.clone();
                return Some(apply_color_transforms(&base_hex, c));
            }
            _ => {}
        }
    }
    None
}

/// Apply OOXML color transforms (lumMod, lumOff, shade, tint, alpha) to a base hex color.
/// Returns 6-char hex when fully opaque, or 8-char hex (RRGGBBAA) when alpha < 1.
fn apply_color_transforms(hex: &str, node: roxmltree::Node<'_, '_>) -> String {
    if hex.len() < 6 {
        return hex.to_owned();
    }
    let r = u8::from_str_radix(&hex[0..2], 16).unwrap_or(0);
    let g = u8::from_str_radix(&hex[2..4], 16).unwrap_or(0);
    let b = u8::from_str_radix(&hex[4..6], 16).unwrap_or(0);

    let mut rf = r as f64 / 255.0;
    let mut gf = g as f64 / 255.0;
    let mut bf = b as f64 / 255.0;
    let mut alpha = 1.0_f64;

    for t in node.children().filter(|n| n.is_element()) {
        match t.tag_name().name() {
            "lumMod" => {
                let val = attr_f64(&t, "val").unwrap_or(100_000.0) / 100_000.0;
                let (h, l, s) = rgb_to_hls(rf, gf, bf);
                let (nr, ng, nb) = hls_to_rgb(h, (l * val).min(1.0), s);
                rf = nr; gf = ng; bf = nb;
            }
            "lumOff" => {
                let val = attr_f64(&t, "val").unwrap_or(0.0) / 100_000.0;
                let (h, l, s) = rgb_to_hls(rf, gf, bf);
                let (nr, ng, nb) = hls_to_rgb(h, (l + val).clamp(0.0, 1.0), s);
                rf = nr; gf = ng; bf = nb;
            }
            "shade" => {
                // shade=100000 → no change; shade=0 → black
                let val = attr_f64(&t, "val").unwrap_or(100_000.0) / 100_000.0;
                rf *= val; gf *= val; bf *= val;
            }
            "tint" => {
                // tint=100000 → no change; tint=0 → white
                let val = attr_f64(&t, "val").unwrap_or(100_000.0) / 100_000.0;
                rf = rf * val + (1.0 - val);
                gf = gf * val + (1.0 - val);
                bf = bf * val + (1.0 - val);
            }
            "alpha" => {
                // alpha=100000 → fully opaque, alpha=0 → fully transparent
                alpha = attr_f64(&t, "val").unwrap_or(100_000.0) / 100_000.0;
            }
            _ => {}
        }
    }

    let r = (rf.clamp(0.0, 1.0) * 255.0).round() as u8;
    let g = (gf.clamp(0.0, 1.0) * 255.0).round() as u8;
    let b = (bf.clamp(0.0, 1.0) * 255.0).round() as u8;
    if (alpha - 1.0).abs() < 0.004 {
        format!("{:02X}{:02X}{:02X}", r, g, b)
    } else {
        let a = (alpha.clamp(0.0, 1.0) * 255.0).round() as u8;
        format!("{:02X}{:02X}{:02X}{:02X}", r, g, b, a)
    }
}

fn rgb_to_hls(r: f64, g: f64, b: f64) -> (f64, f64, f64) {
    let max = r.max(g).max(b);
    let min = r.min(g).min(b);
    let l = (max + min) / 2.0;
    let d = max - min;
    if d < 1e-10 {
        return (0.0, l, 0.0);
    }
    let s = if l > 0.5 { d / (2.0 - max - min) } else { d / (max + min) };
    let h = if (max - r).abs() < 1e-10 {
        (g - b) / d + if g < b { 6.0 } else { 0.0 }
    } else if (max - g).abs() < 1e-10 {
        (b - r) / d + 2.0
    } else {
        (r - g) / d + 4.0
    };
    (h / 6.0, l, s)
}

fn hls_to_rgb(h: f64, l: f64, s: f64) -> (f64, f64, f64) {
    if s < 1e-10 {
        return (l, l, l);
    }
    fn hue2rgb(p: f64, q: f64, mut t: f64) -> f64 {
        if t < 0.0 { t += 1.0; }
        if t > 1.0 { t -= 1.0; }
        if t < 1.0 / 6.0 { return p + (q - p) * 6.0 * t; }
        if t < 0.5        { return q; }
        if t < 2.0 / 3.0  { return p + (q - p) * (2.0 / 3.0 - t) * 6.0; }
        p
    }
    let q = if l < 0.5 { l * (1.0 + s) } else { l + s - l * s };
    let p = 2.0 * l - q;
    (
        hue2rgb(p, q, h + 1.0 / 3.0),
        hue2rgb(p, q, h),
        hue2rgb(p, q, h - 1.0 / 3.0),
    )
}

fn preset_color(name: &str) -> Option<String> {
    let hex = match name {
        "black"                  => "000000",
        "white"                  => "FFFFFF",
        "red"                    => "FF0000",
        "green"                  => "008000",
        "blue"                   => "0000FF",
        "yellow"                 => "FFFF00",
        "cyan"                   => "00FFFF",
        "magenta"                => "FF00FF",
        "orange"                 => "FFA500",
        "gray"   | "grey"        => "808080",
        "darkGray"  | "darkGrey"  => "404040",
        "lightGray" | "lightGrey" => "D3D3D3",
        _ => return None,
    };
    Some(hex.to_owned())
}

// ===========================
//  Fill / Stroke parsing
// ===========================

fn parse_fill(
    node: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<Fill> {
    for c in node.children().filter(|n| n.is_element()) {
        match c.tag_name().name() {
            "solidFill" => {
                // If the color resolves, use it. If not (e.g. phClr with no theme slot),
                // return None so the caller can fall back to the shape style color.
                if let Some(color) = parse_color_node(c, theme) {
                    return Some(Fill::Solid { color });
                }
                // Unresolvable → don't default to black; let fallback logic handle it
            }
            "noFill" => return Some(Fill::None),
            "gradFill" => {
                let mut stops: Vec<GradStop> = child(c, "gsLst")
                    .map(|gs_lst| {
                        gs_lst
                            .children()
                            .filter(|n| n.is_element() && n.tag_name().name() == "gs")
                            .filter_map(|gs| {
                                let position = attr_f64(&gs, "pos").unwrap_or(0.0) / 100_000.0;
                                let color = parse_color_node(gs, theme)?;
                                Some(GradStop { position, color })
                            })
                            .collect()
                    })
                    .unwrap_or_default();

                if stops.is_empty() {
                    // No valid stops — continue scanning other fill elements
                } else {
                    stops.sort_by(|a, b| {
                        a.position.partial_cmp(&b.position).unwrap_or(std::cmp::Ordering::Equal)
                    });
                    let (grad_type, angle) = if let Some(lin) = child(c, "lin") {
                        // OOXML ang: 60000ths of degree, 0 = left→right, 5400000 = top→bottom
                        let ang = attr_f64(&lin, "ang").unwrap_or(0.0) / 60_000.0;
                        ("linear".to_owned(), ang)
                    } else if child(c, "path").is_some() {
                        ("radial".to_owned(), 0.0)
                    } else {
                        ("linear".to_owned(), 0.0)
                    };
                    return Some(Fill::Gradient { stops, angle, grad_type });
                }
            }
            _ => {}
        }
    }
    None
}

fn parse_stroke(
    ln_node: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<Stroke> {
    if child(ln_node, "noFill").is_some() {
        return None;
    }
    let width = attr_i64(&ln_node, "w").unwrap_or(9525);
    let color = child(ln_node, "solidFill")
        .and_then(|n| parse_color_node(n, theme))?;
    let dash_style = child(ln_node, "prstDash")
        .and_then(|n| attr(&n, "val"))
        .filter(|v| v != "solid");
    Some(Stroke { color, width, dash_style })
}

// ===========================
//  Shadow parsing
// ===========================

/// Parse spPr > effectLst > outerShdw into a Shadow.
fn parse_shadow(
    effect_lst: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<Shadow> {
    let outer_shdw = child(effect_lst, "outerShdw")?;
    let blur = attr_i64(&outer_shdw, "blurRad").unwrap_or(0);
    let dist = attr_i64(&outer_shdw, "dist").unwrap_or(0);
    // dir: 60000ths of a degree, clockwise from East (positive x-axis)
    let dir = attr_f64(&outer_shdw, "dir").unwrap_or(0.0) / 60_000.0;

    // Color with optional alpha (8-char hex when alpha != 1)
    let color_str = parse_color_node(outer_shdw, theme).unwrap_or_else(|| "000000".to_owned());
    let (color, alpha) = if color_str.len() >= 8 {
        let a = u8::from_str_radix(&color_str[6..8], 16).unwrap_or(255) as f64 / 255.0;
        (color_str[..6].to_owned(), a)
    } else {
        (color_str, 1.0)
    };

    Some(Shadow { color, alpha, blur, dist, dir })
}

// ===========================
//  Custom geometry parsing
// ===========================

/// Parse a single path command node; coordinates are normalised to [0,1].
fn parse_path_cmd(
    cmd_node: roxmltree::Node<'_, '_>,
    path_w: f64,
    path_h: f64,
) -> Option<PathCmd> {
    match cmd_node.tag_name().name() {
        "moveTo" => {
            let pt = child(cmd_node, "pt")?;
            let x = attr_f64(&pt, "x")? / path_w;
            let y = attr_f64(&pt, "y")? / path_h;
            Some(PathCmd::MoveTo { x, y })
        }
        "lnTo" => {
            let pt = child(cmd_node, "pt")?;
            let x = attr_f64(&pt, "x")? / path_w;
            let y = attr_f64(&pt, "y")? / path_h;
            Some(PathCmd::LineTo { x, y })
        }
        "cubicBezTo" => {
            let pts: Vec<_> = children_vec(cmd_node, "pt");
            if pts.len() < 3 { return None; }
            let x1 = attr_f64(&pts[0], "x")? / path_w;
            let y1 = attr_f64(&pts[0], "y")? / path_h;
            let x2 = attr_f64(&pts[1], "x")? / path_w;
            let y2 = attr_f64(&pts[1], "y")? / path_h;
            let x  = attr_f64(&pts[2], "x")? / path_w;
            let y  = attr_f64(&pts[2], "y")? / path_h;
            Some(PathCmd::CubicBezTo { x1, y1, x2, y2, x, y })
        }
        "arcTo" => {
            // wR/hR are radii in path-local units; stAng/swAng in 60000ths of a degree
            let wr     = attr_f64(&cmd_node, "wR").unwrap_or(0.0) / path_w;
            let hr     = attr_f64(&cmd_node, "hR").unwrap_or(0.0) / path_h;
            let st_ang = attr_f64(&cmd_node, "stAng").unwrap_or(0.0) / 60000.0;
            let sw_ang = attr_f64(&cmd_node, "swAng").unwrap_or(0.0) / 60000.0;
            Some(PathCmd::ArcTo { wr, hr, st_ang, sw_ang })
        }
        "close" => Some(PathCmd::Close),
        _ => None,
    }
}

/// Parse custGeom > pathLst into a list of sub-paths (one per <a:path> element).
fn parse_cust_geom(cust_geom: roxmltree::Node<'_, '_>) -> Vec<Vec<PathCmd>> {
    let path_lst = match child(cust_geom, "pathLst") {
        Some(n) => n,
        None => return vec![],
    };

    path_lst
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "path")
        .map(|path_node| {
            let path_w = attr_f64(&path_node, "w").unwrap_or(1.0).max(1.0);
            let path_h = attr_f64(&path_node, "h").unwrap_or(1.0).max(1.0);
            path_node
                .children()
                .filter(|n| n.is_element())
                .filter_map(|cmd| parse_path_cmd(cmd, path_w, path_h))
                .collect()
        })
        .collect()
}

// ===========================
//  Transform (a:xfrm)
// ===========================

#[derive(Clone, Debug, Default)]
struct Transform {
    x: i64,
    y: i64,
    cx: i64,
    cy: i64,
    /// Degrees, clockwise
    rot: f64,
    flip_h: bool,
    flip_v: bool,
}

fn parse_xfrm(xfrm: roxmltree::Node<'_, '_>) -> Transform {
    let rot    = attr_f64(&xfrm, "rot").unwrap_or(0.0) / 60000.0;
    let flip_h = attr(&xfrm, "flipH").map(|v| v == "1" || v == "true").unwrap_or(false);
    let flip_v = attr(&xfrm, "flipV").map(|v| v == "1" || v == "true").unwrap_or(false);
    let off = child(xfrm, "off");
    let ext = child(xfrm, "ext");
    Transform {
        x:  off.and_then(|n| attr_i64(&n, "x")).unwrap_or(0),
        y:  off.and_then(|n| attr_i64(&n, "y")).unwrap_or(0),
        cx: ext.and_then(|n| attr_i64(&n, "cx")).unwrap_or(0),
        cy: ext.and_then(|n| attr_i64(&n, "cy")).unwrap_or(0),
        rot, flip_h, flip_v,
    }
}

// ===========================
//  Group transform
// ===========================

#[derive(Clone, Debug, Default)]
struct GroupTransform {
    x: i64, y: i64,
    cx: i64, cy: i64,
    ch_x: i64, ch_y: i64,
    ch_cx: i64, ch_cy: i64,
    flip_h: bool,
    flip_v: bool,
    /// Group rotation in degrees, clockwise
    rot: f64,
}

impl GroupTransform {
    fn apply_to_transform(&self, t: Transform) -> Transform {
        let sx = if self.ch_cx != 0 { self.cx as f64 / self.ch_cx as f64 } else { 1.0 };
        let sy = if self.ch_cy != 0 { self.cy as f64 / self.ch_cy as f64 } else { 1.0 };
        // If the group is flipped, mirror child positions in child coordinate space
        // before applying the normal scale+translate.
        // Mirror formula: new_left = (ch_x + ch_cx) - (t.x - ch_x) - t.cx
        //                          = 2*ch_x + ch_cx - t.x - t.cx
        let child_x = if self.flip_h {
            2 * self.ch_x + self.ch_cx - t.x - t.cx
        } else {
            t.x
        };
        let child_y = if self.flip_v {
            2 * self.ch_y + self.ch_cy - t.y - t.cy
        } else {
            t.y
        };

        // Child position and size in parent space (before group rotation)
        let new_x = (child_x - self.ch_x) as f64 * sx + self.x as f64;
        let new_y = (child_y - self.ch_y) as f64 * sy + self.y as f64;
        let new_cx = (t.cx as f64 * sx).round() as i64;
        let new_cy = (t.cy as f64 * sy).round() as i64;

        // Apply group rotation: rotate child center around group center (clockwise, screen coords)
        let (final_x, final_y) = if self.rot != 0.0 {
            let rot_rad = self.rot.to_radians();
            let cos_r = rot_rad.cos();
            let sin_r = rot_rad.sin();
            let group_cx = self.x as f64 + self.cx as f64 / 2.0;
            let group_cy = self.y as f64 + self.cy as f64 / 2.0;
            let child_cx = new_x + new_cx as f64 / 2.0;
            let child_cy = new_y + new_cy as f64 / 2.0;
            let dx = child_cx - group_cx;
            let dy = child_cy - group_cy;
            // Clockwise rotation in screen coords (y-axis down): x' = x*cos - y*sin, y' = x*sin + y*cos
            let dx_new = dx * cos_r - dy * sin_r;
            let dy_new = dx * sin_r + dy * cos_r;
            (group_cx + dx_new - new_cx as f64 / 2.0, group_cy + dy_new - new_cy as f64 / 2.0)
        } else {
            (new_x, new_y)
        };

        // When the group has a net flip, the child's own rotation direction is negated
        // before the group rotation is added (scale→flip→rotate OOXML order).
        // GF (group net flip) = flip_h XOR flip_v.
        let gf = self.flip_h ^ self.flip_v;
        Transform {
            x:  final_x.round() as i64,
            y:  final_y.round() as i64,
            cx: new_cx,
            cy: new_cy,
            rot: self.rot + if gf { -t.rot } else { t.rot },
            // Propagate group flip to child element flip flags
            flip_h: t.flip_h ^ self.flip_h,
            flip_v: t.flip_v ^ self.flip_v,
        }
    }
}

fn apply_group_transform_to_element(el: &mut SlideElement, gt: &GroupTransform) {
    match el {
        SlideElement::Shape(s) => {
            let t = Transform { x: s.x, y: s.y, cx: s.width, cy: s.height, rot: s.rotation, flip_h: s.flip_h, flip_v: s.flip_v };
            let nt = gt.apply_to_transform(t);
            s.x = nt.x; s.y = nt.y; s.width = nt.cx; s.height = nt.cy;
            s.rotation = nt.rot; s.flip_h = nt.flip_h; s.flip_v = nt.flip_v;
        }
        SlideElement::Picture(p) => {
            let t = Transform { x: p.x, y: p.y, cx: p.width, cy: p.height, rot: p.rotation, flip_h: p.flip_h, flip_v: p.flip_v };
            let nt = gt.apply_to_transform(t);
            p.x = nt.x; p.y = nt.y; p.width = nt.cx; p.height = nt.cy;
            p.rotation = nt.rot; p.flip_h = nt.flip_h; p.flip_v = nt.flip_v;
        }
        SlideElement::Table(tbl) => {
            // If the table has no xfrm (zero dimensions), it fills the group's child space.
            let (ex, ey, ecx, ecy) = if tbl.width == 0 && tbl.height == 0 {
                (gt.ch_x, gt.ch_y, gt.ch_cx, gt.ch_cy)
            } else {
                (tbl.x, tbl.y, tbl.width, tbl.height)
            };
            let t = Transform { x: ex, y: ey, cx: ecx, cy: ecy, rot: 0.0, flip_h: false, flip_v: false };
            let nt = gt.apply_to_transform(t);
            tbl.x = nt.x; tbl.y = nt.y; tbl.width = nt.cx; tbl.height = nt.cy;
        }
        SlideElement::Chart(chart) => {
            // If the chart graphicFrame has no xfrm (zero dimensions), it fills the group's child space.
            let (ex, ey, ecx, ecy) = if chart.width == 0 && chart.height == 0 {
                (gt.ch_x, gt.ch_y, gt.ch_cx, gt.ch_cy)
            } else {
                (chart.x, chart.y, chart.width, chart.height)
            };
            let t = Transform { x: ex, y: ey, cx: ecx, cy: ecy, rot: 0.0, flip_h: false, flip_v: false };
            let nt = gt.apply_to_transform(t);
            chart.x = nt.x; chart.y = nt.y; chart.width = nt.cx; chart.height = nt.cy;
        }
    }
}

// ===========================
//  Layout placeholder map
// ===========================

/// Keyed first by idx (integer), then by type string.
#[derive(Default)]
struct LayoutPlaceholders {
    by_idx:  HashMap<u32, Transform>,
    by_type: HashMap<String, Transform>,
    /// Fallback transforms from slide master (by ph_type), used when layout has no xfrm
    master_by_type: HashMap<String, Transform>,
    /// Default font size (pt) per placeholder idx, from layout/master lstStyle
    by_idx_font_size:  HashMap<u32, f64>,
    /// Default font size (pt) per placeholder type, from layout/master lstStyle
    by_type_font_size: HashMap<String, f64>,
    /// Default bold per placeholder type, from layout lstStyle defRPr b attribute
    by_type_bold: HashMap<String, bool>,
    /// Default italic per placeholder type, from layout lstStyle defRPr i attribute
    by_type_italic: HashMap<String, bool>,
    /// Vertical anchor ("t"/"ctr"/"b") per placeholder type, from layout/master bodyPr
    by_type_anchor: HashMap<String, String>,
    /// Default paragraph alignment per placeholder type, from layout/master lstStyle
    by_type_alignment: HashMap<String, String>,
    /// Default space-before (hundredths of pt) per placeholder type, from layout lstStyle
    by_type_space_before: HashMap<String, i64>,
    /// Default space-after (hundredths of pt) per placeholder type, from layout lstStyle
    by_type_space_after: HashMap<String, i64>,
    /// Default space-before from master txStyles (fallback when layout has none)
    by_type_master_space_before: HashMap<String, i64>,
    /// Default space-after from master txStyles (fallback when layout has none)
    by_type_master_space_after: HashMap<String, i64>,
    /// Stroke per placeholder type from layout spPr > ln
    by_type_stroke: HashMap<String, Stroke>,
    /// Stroke per placeholder idx from layout spPr > ln
    by_idx_stroke: HashMap<u32, Stroke>,
    /// Default line spacing (spcPct val, e.g. 90000 = 90%) per placeholder idx, from layout lstStyle
    by_idx_line_spacing: HashMap<u32, f64>,
    /// Default line spacing (spcPct val) per placeholder type, from layout lstStyle
    by_type_line_spacing: HashMap<String, f64>,
    /// Paragraph alignment per placeholder type from master lstStyle > lvl1pPr algn (fallback)
    by_type_master_alignment: HashMap<String, String>,
    /// Default line spacing from master txStyles (fallback when layout has none)
    by_type_master_line_spacing: HashMap<String, f64>,
}

impl LayoutPlaceholders {
    fn lookup(&self, ph_type: &str, ph_idx: Option<u32>) -> Option<&Transform> {
        ph_idx
            .and_then(|i| self.by_idx.get(&i))
            .or_else(|| self.by_type.get(ph_type))
            .or_else(|| {
                if ph_type == "body" { self.by_type.get("") } else { None }
            })
            .or_else(|| self.master_by_type.get(ph_type))
    }

    /// Look up the inherited default font size for a placeholder (layout then master fallback).
    fn lookup_font_size(&self, ph_type: &str, ph_idx: Option<u32>) -> Option<f64> {
        ph_idx
            .and_then(|i| self.by_idx_font_size.get(&i).copied())
            .or_else(|| self.by_type_font_size.get(ph_type).copied())
            .or_else(|| {
                if ph_type == "body" { self.by_type_font_size.get("").copied() } else { None }
            })
    }

    /// Look up inherited bold for this placeholder type.
    fn lookup_bold(&self, ph_type: &str) -> Option<bool> {
        self.by_type_bold.get(ph_type).copied()
            .or_else(|| if ph_type == "body" { self.by_type_bold.get("").copied() } else { None })
    }

    /// Look up inherited italic for this placeholder type.
    fn lookup_italic(&self, ph_type: &str) -> Option<bool> {
        self.by_type_italic.get(ph_type).copied()
            .or_else(|| if ph_type == "body" { self.by_type_italic.get("").copied() } else { None })
    }

    /// Look up inherited vertical anchor for this placeholder type.
    fn lookup_anchor(&self, ph_type: &str) -> Option<String> {
        self.by_type_anchor.get(ph_type).cloned()
            .or_else(|| if ph_type == "body" { self.by_type_anchor.get("").cloned() } else { None })
    }

    /// Look up inherited paragraph alignment for this placeholder type.
    fn lookup_alignment(&self, ph_type: &str) -> Option<String> {
        self.by_type_alignment.get(ph_type).cloned()
            .or_else(|| if ph_type == "body" { self.by_type_alignment.get("").cloned() } else { None })
            .or_else(|| self.by_type_master_alignment.get(ph_type).cloned())
            .or_else(|| if ph_type == "body" { self.by_type_master_alignment.get("").cloned() } else { None })
    }

    fn lookup_space_before(&self, ph_type: &str) -> Option<i64> {
        self.by_type_space_before.get(ph_type).copied()
            .or_else(|| if ph_type == "body" { self.by_type_space_before.get("").copied() } else { None })
            .or_else(|| self.by_type_master_space_before.get(ph_type).copied())
            .or_else(|| if ph_type == "body" { self.by_type_master_space_before.get("").copied() } else { None })
    }

    fn lookup_space_after(&self, ph_type: &str) -> Option<i64> {
        self.by_type_space_after.get(ph_type).copied()
            .or_else(|| if ph_type == "body" { self.by_type_space_after.get("").copied() } else { None })
            .or_else(|| self.by_type_master_space_after.get(ph_type).copied())
            .or_else(|| if ph_type == "body" { self.by_type_master_space_after.get("").copied() } else { None })
    }

    /// Look up inherited stroke from the layout placeholder spPr > ln.
    fn lookup_stroke(&self, ph_type: &str, ph_idx: Option<u32>) -> Option<Stroke> {
        ph_idx
            .and_then(|i| self.by_idx_stroke.get(&i).cloned())
            .or_else(|| self.by_type_stroke.get(ph_type).cloned())
            .or_else(|| if ph_type == "body" { self.by_type_stroke.get("").cloned() } else { None })
    }

    /// Look up inherited line spacing (spcPct val, e.g. 90000 = 90%) for this placeholder.
    fn lookup_line_spacing(&self, ph_type: &str, ph_idx: Option<u32>) -> Option<f64> {
        ph_idx
            .and_then(|i| self.by_idx_line_spacing.get(&i).copied())
            .or_else(|| self.by_type_line_spacing.get(ph_type).copied())
            .or_else(|| if ph_type == "body" { self.by_type_line_spacing.get("").copied() } else { None })
            .or_else(|| self.by_type_master_line_spacing.get(ph_type).copied())
            .or_else(|| if ph_type == "body" { self.by_type_master_line_spacing.get("").copied() } else { None })
    }
}

/// Extract the lvl1pPr defRPr font size from a txBody node.
fn extract_lvl1_font_size(tx_body: roxmltree::Node<'_, '_>) -> Option<f64> {
    child(tx_body, "lstStyle")
        .and_then(|ls| child(ls, "lvl1pPr"))
        .and_then(|lp| child(lp, "defRPr"))
        .and_then(|rp| attr_f64(&rp, "sz"))
        .map(|v| v / 100.0)
}

/// Parse bodyPr anchor ("t"/"ctr"/"b") from master placeholder shapes.
fn parse_master_anchors(master_xml: &str) -> HashMap<String, String> {
    let mut map = HashMap::new();
    let doc = match roxmltree::Document::parse(master_xml) {
        Ok(d) => d,
        Err(_) => return map,
    };
    let root = doc.root_element();
    if let Some(sp_tree) = child(root, "cSld").and_then(|n| child(n, "spTree")) {
        for sp in sp_tree.children().filter(|n| n.is_element() && n.tag_name().name() == "sp") {
            let ph_node = sp.descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "ph");
            if let Some(ph) = ph_node {
                let ph_type = attr(&ph, "type").unwrap_or_default();
                if let Some(anchor) = child(sp, "txBody")
                    .and_then(|tb| child(tb, "bodyPr"))
                    .and_then(|bp| attr(&bp, "anchor"))
                {
                    map.entry(ph_type).or_insert(anchor.to_string());
                }
            }
        }
    }
    map
}

/// Parse paragraph alignment from master placeholder shapes' lstStyle > lvl1pPr algn attribute.
fn parse_master_alignments(master_xml: &str) -> HashMap<String, String> {
    let mut map = HashMap::new();
    let doc = match roxmltree::Document::parse(master_xml) {
        Ok(d) => d,
        Err(_) => return map,
    };
    let root = doc.root_element();
    if let Some(sp_tree) = child(root, "cSld").and_then(|n| child(n, "spTree")) {
        for sp in sp_tree.children().filter(|n| n.is_element() && n.tag_name().name() == "sp") {
            let ph_node = sp.descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "ph");
            if let Some(ph) = ph_node {
                let ph_type = attr(&ph, "type").unwrap_or_default();
                if let Some(algn) = child(sp, "txBody")
                    .and_then(|tb| child(tb, "lstStyle"))
                    .and_then(|ls| child(ls, "lvl1pPr"))
                    .and_then(|lp| attr(&lp, "algn"))
                {
                    map.entry(ph_type).or_insert(algn.to_string());
                }
            }
        }
    }
    map
}

/// Parse master-level default font sizes from txStyles (titleStyle / bodyStyle / otherStyle)
/// and from individual placeholder shapes in the master spTree.
/// Individual shape lstStyle takes priority over txStyles generic defaults.
fn parse_master_font_sizes(master_xml: &str) -> HashMap<String, f64> {
    let mut map = HashMap::new();
    let doc = match roxmltree::Document::parse(master_xml) {
        Ok(d) => d,
        Err(_) => return map,
    };
    let root = doc.root_element();

    // Scan master spTree placeholder shapes first — per-shape lstStyle is more specific
    if let Some(sp_tree) = child(root, "cSld").and_then(|n| child(n, "spTree")) {
        for sp in sp_tree.children().filter(|n| n.is_element() && n.tag_name().name() == "sp") {
            let ph_node = sp.descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "ph");
            if let Some(ph) = ph_node {
                let ph_type = attr(&ph, "type").unwrap_or_default();
                if let Some(tx_body) = child(sp, "txBody") {
                    if let Some(sz) = extract_lvl1_font_size(tx_body) {
                        map.entry(ph_type).or_insert(sz);
                    }
                }
            }
        }
    }

    // p:txStyles > a:titleStyle / a:bodyStyle / a:otherStyle as fallback
    if let Some(tx_styles) = child(root, "txStyles") {
        let style_ph_map: &[(&str, &[&str])] = &[
            ("titleStyle",  &["title", "ctrTitle"]),
            ("bodyStyle",   &["body", "subTitle", "obj", ""]),
            ("otherStyle",  &["dt", "ftr", "sldNum"]),
        ];
        for (style_name, ph_types) in style_ph_map {
            let sz = child(tx_styles, style_name)
                .and_then(|sn| child(sn, "lvl1pPr"))
                .and_then(|lp| child(lp, "defRPr"))
                .and_then(|rp| attr_f64(&rp, "sz"))
                .map(|v| v / 100.0);
            if let Some(fs) = sz {
                for ph_type in *ph_types {
                    map.entry(ph_type.to_string()).or_insert(fs);
                }
            }
        }
    }

    map
}

/// Parse default paragraph spacing from master txStyles.
/// Returns (space_before_map, space_after_map, line_spacing_map) keyed by ph_type string.
/// space_before/after values are in hundredths of a point (same as Paragraph.space_before/after).
/// Note: line_spacing_map is intentionally NOT populated. Inheriting txStyles lnSpc hurts VRT
/// scores because our font substitutes (sans-serif) have different em-square metrics than the
/// original Aptos font, so applying the master's 120% line spacing over-expands text layout.
fn parse_master_txstyle_spacing(master_xml: &str) -> (HashMap<String, i64>, HashMap<String, i64>, HashMap<String, f64>) {
    let mut before_map: HashMap<String, i64> = HashMap::new();
    let mut after_map:  HashMap<String, i64> = HashMap::new();
    let line_map:       HashMap<String, f64> = HashMap::new(); // intentionally not populated
    let doc = match roxmltree::Document::parse(master_xml) {
        Ok(d) => d,
        Err(_) => return (before_map, after_map, line_map),
    };
    let root = doc.root_element();
    let tx_styles = match child(root, "txStyles") {
        Some(n) => n,
        None => return (before_map, after_map, line_map),
    };
    let style_ph_map: &[(&str, &[&str])] = &[
        ("titleStyle",  &["title", "ctrTitle"]),
        ("bodyStyle",   &["body", "subTitle", "obj", ""]),
        ("otherStyle",  &["dt", "ftr", "sldNum"]),
    ];
    for (style_name, ph_types) in style_ph_map {
        let lvl1 = child(tx_styles, style_name).and_then(|sn| child(sn, "lvl1pPr"));
        let spc_before = lvl1.and_then(|lp| child(lp, "spcBef"))
            .and_then(|s| child(s, "spcPts").and_then(|n| attr_i64(&n, "val")));
        let spc_after = lvl1.and_then(|lp| child(lp, "spcAft"))
            .and_then(|s| child(s, "spcPts").and_then(|n| attr_i64(&n, "val")));
        if let Some(v) = spc_before {
            for ph_type in *ph_types {
                before_map.entry(ph_type.to_string()).or_insert(v);
            }
        }
        if let Some(v) = spc_after {
            for ph_type in *ph_types {
                after_map.entry(ph_type.to_string()).or_insert(v);
            }
        }
    }
    (before_map, after_map, line_map)
}

fn parse_master_transforms(master_xml: &str) -> HashMap<String, Transform> {
    let mut map = HashMap::new();
    let doc = match roxmltree::Document::parse(master_xml) {
        Ok(d) => d,
        Err(_) => return map,
    };
    let root = doc.root_element();
    if let Some(sp_tree) = child(root, "cSld").and_then(|n| child(n, "spTree")) {
        for sp in sp_tree.children().filter(|n| n.is_element() && n.tag_name().name() == "sp") {
            let ph_node = sp.descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "ph");
            if let Some(ph) = ph_node {
                let ph_type = attr(&ph, "type").unwrap_or_default();
                if let Some(xfrm) = child(sp, "spPr").and_then(|p| child(p, "xfrm")) {
                    map.entry(ph_type).or_insert_with(|| parse_xfrm(xfrm));
                }
            }
        }
    }
    map
}

fn parse_layout_placeholders(layout_xml: &str, master_font_sizes: &HashMap<String, f64>, master_anchors: &HashMap<String, String>, master_transforms: &HashMap<String, Transform>, master_alignments: &HashMap<String, String>, master_space_before: &HashMap<String, i64>, master_space_after: &HashMap<String, i64>, master_line_spacing: &HashMap<String, f64>, theme: &HashMap<String, String>) -> LayoutPlaceholders {
    let mut lph = LayoutPlaceholders::default();
    lph.master_by_type = master_transforms.clone();
    lph.by_type_master_alignment = master_alignments.clone();
    lph.by_type_master_space_before = master_space_before.clone();
    lph.by_type_master_space_after = master_space_after.clone();
    lph.by_type_master_line_spacing = master_line_spacing.clone();
    let doc = match roxmltree::Document::parse(layout_xml) {
        Ok(d) => d,
        Err(_) => return lph,
    };
    let root = doc.root_element();

    let sp_tree = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "spTree");
    let sp_tree = match sp_tree {
        Some(n) => n,
        None => return lph,
    };

    for sp in sp_tree
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "sp")
    {
        let ph_node = sp
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "ph");
        let sp_pr = match child(sp, "spPr") {
            Some(n) => n,
            None => continue,
        };
        // xfrm may be absent (placeholder inherits transform from master); parse if present
        let t_opt: Option<Transform> = child(sp_pr, "xfrm").map(parse_xfrm);

        // Extract layout-level defaults from the placeholder's txBody > lstStyle > lvl1pPr
        let layout_lvl1_ppr: Option<roxmltree::Node<'_, '_>> = child(sp, "txBody")
            .and_then(|tb| child(tb, "lstStyle"))
            .and_then(|ls| child(ls, "lvl1pPr"));
        let layout_def_rpr: Option<roxmltree::Node<'_, '_>> = layout_lvl1_ppr
            .and_then(|lp| child(lp, "defRPr"));
        let layout_font_size = layout_def_rpr.and_then(|rp| attr_f64(&rp, "sz")).map(|v| v / 100.0);
        let layout_bold   = layout_def_rpr.and_then(|rp| attr(&rp, "b")).map(|v| v == "1" || v == "true");
        let layout_italic = layout_def_rpr.and_then(|rp| attr(&rp, "i")).map(|v| v == "1" || v == "true");
        let layout_alignment: Option<String> = layout_lvl1_ppr
            .and_then(|lp| attr(&lp, "algn"))
            .map(|a| a.to_string());
        let layout_space_before: Option<i64> = layout_lvl1_ppr
            .and_then(|lp| child(lp, "spcBef"))
            .and_then(|s| child(s, "spcPts"))
            .and_then(|s| attr_i64(&s, "val"));
        let layout_space_after: Option<i64> = layout_lvl1_ppr
            .and_then(|lp| child(lp, "spcAft"))
            .and_then(|s| child(s, "spcPts"))
            .and_then(|s| attr_i64(&s, "val"));
        // lnSpc > spcPct val (e.g. 90000 = 90%)
        let layout_line_spacing: Option<f64> = layout_lvl1_ppr
            .and_then(|lp| child(lp, "lnSpc"))
            .and_then(|ls| child(ls, "spcPct"))
            .and_then(|s| attr_f64(&s, "val"));

        // Layout bodyPr anchor; fall back to master anchor map
        let layout_anchor: Option<String> = child(sp, "txBody")
            .and_then(|tb| child(tb, "bodyPr"))
            .and_then(|bp| attr(&bp, "anchor"))
            .map(|a| a.to_string());

        // Layout spPr > ln stroke (real visible border, not edit-mode indicator when solidFill is present)
        let layout_stroke: Option<Stroke> = child(sp_pr, "ln")
            .and_then(|n| parse_stroke(n, theme));

        if let Some(ph) = ph_node {
            let ph_type = attr(&ph, "type").unwrap_or_default();
            let ph_idx: Option<u32> = attr(&ph, "idx").and_then(|v| v.parse().ok());

            if let Some(idx) = ph_idx {
                if let Some(ref t) = t_opt {
                    lph.by_idx.entry(idx).or_insert_with(|| t.clone());
                }
                // Prefer layout font size; fall back to master
                let fs = layout_font_size
                    .or_else(|| master_font_sizes.get(&ph_type).copied());
                if let Some(fs) = fs {
                    lph.by_idx_font_size.entry(idx).or_insert(fs);
                }
                if let Some(ref s) = layout_stroke {
                    lph.by_idx_stroke.entry(idx).or_insert(s.clone());
                }
                if let Some(ls) = layout_line_spacing {
                    lph.by_idx_line_spacing.entry(idx).or_insert(ls);
                }
            }
            let effective_fs = layout_font_size
                .or_else(|| master_font_sizes.get(&ph_type).copied());
            if let Some(fs) = effective_fs {
                lph.by_type_font_size.entry(ph_type.clone()).or_insert(fs);
            }
            if let Some(b) = layout_bold {
                lph.by_type_bold.entry(ph_type.clone()).or_insert(b);
            }
            if let Some(i) = layout_italic {
                lph.by_type_italic.entry(ph_type.clone()).or_insert(i);
            }
            if let Some(a) = layout_alignment {
                lph.by_type_alignment.entry(ph_type.clone()).or_insert(a);
            }
            if let Some(v) = layout_space_before {
                lph.by_type_space_before.entry(ph_type.clone()).or_insert(v);
            }
            if let Some(v) = layout_space_after {
                lph.by_type_space_after.entry(ph_type.clone()).or_insert(v);
            }
            if let Some(ls) = layout_line_spacing {
                lph.by_type_line_spacing.entry(ph_type.clone()).or_insert(ls);
            }
            // Anchor: layout bodyPr > fall back to master anchor map
            let effective_anchor = layout_anchor.clone()
                .or_else(|| master_anchors.get(&ph_type).cloned());
            if let Some(a) = effective_anchor {
                lph.by_type_anchor.entry(ph_type.clone()).or_insert(a);
            }
            if let Some(s) = layout_stroke {
                lph.by_type_stroke.entry(ph_type.clone()).or_insert(s);
            }
            if let Some(t) = t_opt {
                lph.by_type.entry(ph_type).or_insert(t);
            }
        }
    }
    lph
}

// ===========================
//  Text body parsing
// ===========================

fn parse_text_body(
    tx_body: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    inherited_font_size: Option<f64>,
    inherited_bold: Option<bool>,
    inherited_italic: Option<bool>,
    inherited_anchor: Option<String>,
    inherited_alignment: Option<String>,
    inherited_space_before: Option<i64>,
    inherited_space_after: Option<i64>,
    inherited_line_spacing: Option<f64>,
) -> TextBody {
    let body_pr = child(tx_body, "bodyPr");
    let vertical_anchor = body_pr
        .and_then(|n| attr(&n, "anchor"))
        .map(|a| a.to_string())
        .or(inherited_anchor)
        .unwrap_or_else(|| "t".into());
    // Text insets (EMU). OOXML defaults: lIns=rIns=91440, tIns=bIns=45720
    let l_ins = body_pr.and_then(|n| attr_i64(&n, "lIns")).unwrap_or(91_440);
    let r_ins = body_pr.and_then(|n| attr_i64(&n, "rIns")).unwrap_or(91_440);
    let t_ins = body_pr.and_then(|n| attr_i64(&n, "tIns")).unwrap_or(45_720);
    let b_ins = body_pr.and_then(|n| attr_i64(&n, "bIns")).unwrap_or(45_720);
    let wrap = body_pr.and_then(|n| attr(&n, "wrap")).unwrap_or_else(|| "square".into());
    let vert = body_pr.and_then(|n| attr(&n, "vert")).unwrap_or_else(|| "horz".into());
    let auto_fit = body_pr.map(|n| {
        if child(n, "spAutoFit").is_some() { "sp".to_owned() }
        else if child(n, "normAutoFit").is_some() { "norm".to_owned() }
        else { "none".to_owned() }
    }).unwrap_or_else(|| "none".to_owned());

    // Own lstStyle > lvl1pPr, then fall back to layout/master inherited values
    let own_lvl1_ppr = child(tx_body, "lstStyle")
        .and_then(|ls| child(ls, "lvl1pPr"));
    let own_def_rpr = own_lvl1_ppr.and_then(|lp| child(lp, "defRPr"));
    let default_font_size = own_def_rpr.and_then(|rp| attr_f64(&rp, "sz"))
        .map(|v| v / 100.0)
        .or(inherited_font_size);
    let default_bold = own_def_rpr
        .and_then(|rp| attr(&rp, "b")).map(|v| v == "1" || v == "true")
        .or(inherited_bold);
    let default_italic = own_def_rpr
        .and_then(|rp| attr(&rp, "i")).map(|v| v == "1" || v == "true")
        .or(inherited_italic);
    // Own lstStyle > lvl1pPr > algn overrides inherited alignment
    let body_default_alignment = own_lvl1_ppr
        .and_then(|lp| attr(&lp, "algn"))
        .map(|a| a.to_string())
        .or(inherited_alignment);

    // Own lstStyle > lvl1pPr spacing overrides inherited
    let own_lvl1_spcbef: Option<i64> = own_lvl1_ppr
        .and_then(|lp| child(lp, "spcBef"))
        .and_then(|s| child(s, "spcPts"))
        .and_then(|s| attr_i64(&s, "val"));
    let own_lvl1_spcaft: Option<i64> = own_lvl1_ppr
        .and_then(|lp| child(lp, "spcAft"))
        .and_then(|s| child(s, "spcPts"))
        .and_then(|s| attr_i64(&s, "val"));
    let body_default_space_before = own_lvl1_spcbef.or(inherited_space_before);
    let body_default_space_after  = own_lvl1_spcaft.or(inherited_space_after);

    // Own lstStyle > lvl1pPr > lnSpc overrides inherited line spacing
    let own_lvl1_line_spacing: Option<f64> = own_lvl1_ppr
        .and_then(|lp| child(lp, "lnSpc"))
        .and_then(|ls| child(ls, "spcPct"))
        .and_then(|s| attr_f64(&s, "val"));
    let body_default_line_spacing = own_lvl1_line_spacing.or(inherited_line_spacing);

    let paragraphs = children_vec(tx_body, "p")
        .into_iter()
        .map(|p| parse_paragraph(p, theme, body_default_alignment.as_deref(), body_default_space_before, body_default_space_after, body_default_line_spacing))
        .collect();

    TextBody { vertical_anchor, paragraphs, default_font_size, default_bold, default_italic, l_ins, r_ins, t_ins, b_ins, wrap, vert, auto_fit }
}

fn parse_paragraph(
    p_node: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    body_default_alignment: Option<&str>,
    body_default_space_before: Option<i64>,
    body_default_space_after: Option<i64>,
    body_default_line_spacing: Option<f64>,
) -> Paragraph {
    let p_pr = child(p_node, "pPr");

    // Paragraph's own algn → body/layout/master default → "l"
    let alignment = p_pr.and_then(|n| attr(&n, "algn"))
        .map(|a| a.to_string())
        .or_else(|| body_default_alignment.map(|a| a.to_string()))
        .unwrap_or_else(|| "l".into());
    let lvl: u32   = p_pr.and_then(|n| attr(&n, "lvl")).and_then(|v| v.parse().ok()).unwrap_or(0);

    // Detect whether the paragraph has an explicit bullet character (buChar/buAutoNum).
    // Used to choose between bullet-list defaults and plain-text defaults.
    let has_explicit_bullet = p_pr.map(|n| {
        child(n, "buChar").is_some() || child(n, "buAutoNum").is_some()
    }).unwrap_or(false);

    // marL / indent defaults follow PowerPoint's implicit list style:
    //   Bullet paragraphs:  marL = (lvl+1)*342900, indent = -342900 (hanging)
    //   Plain paragraphs:   marL = lvl*457200 (matches presentation.xml defaultTextStyle)
    let mar_l = p_pr.and_then(|n| attr_i64(&n, "marL")).unwrap_or_else(|| {
        if has_explicit_bullet {
            (lvl as i64 + 1) * 342900
        } else {
            lvl as i64 * 457200
        }
    });
    let mar_r  = p_pr.and_then(|n| attr_i64(&n, "marR")).unwrap_or(0);
    let indent = p_pr.and_then(|n| attr_i64(&n, "indent")).unwrap_or_else(|| {
        if has_explicit_bullet { -342900 } else { 0 }
    });

    let space_before = p_pr.and_then(|n| {
        child(n, "spcBef").and_then(|s| child(s, "spcPts")).and_then(|s| attr_i64(&s, "val"))
    }).or(body_default_space_before);
    let space_after = p_pr.and_then(|n| {
        child(n, "spcAft").and_then(|s| child(s, "spcPts")).and_then(|s| attr_i64(&s, "val"))
    }).or(body_default_space_after);

    let space_line = p_pr.and_then(|n| {
        let spc = child(n, "lnSpc")?;
        if let Some(pct) = child(spc, "spcPct") {
            attr_f64(&pct, "val").map(|v| SpaceLine::Pct { val: v })
        } else {
            child(spc, "spcPts")
                .and_then(|pts| attr_f64(&pts, "val"))
                .map(|v| SpaceLine::Pts { val: v / 100.0 }) // hundredths of pt → pt
        }
    }).or_else(|| body_default_line_spacing.map(|v| SpaceLine::Pct { val: v }));

    let bullet = parse_bullet(p_pr, theme);

    // Tab stops from pPr > tabLst
    let tab_stops: Vec<TabStop> = p_pr
        .and_then(|n| child(n, "tabLst"))
        .map(|tab_lst| {
            tab_lst
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "tab")
                .filter_map(|tab| {
                    let pos  = attr_i64(&tab, "pos")?;
                    let algn = attr(&tab, "algn").unwrap_or_else(|| "l".into());
                    Some(TabStop { pos, algn })
                })
                .collect()
        })
        .unwrap_or_default();

    // Paragraph-level default run properties (pPr > defRPr)
    let def_rpr        = p_pr.and_then(|n| child(n, "defRPr"));
    let def_font_size  = def_rpr.and_then(|n| attr_f64(&n, "sz")).map(|v| v / 100.0);
    let def_color      = def_rpr.and_then(|n| child(n, "solidFill")).and_then(|n| parse_color_node(n, theme));
    let def_bold       = def_rpr.and_then(|n| attr(&n, "b")).map(|v| v == "1" || v == "true");
    let def_italic     = def_rpr.and_then(|n| attr(&n, "i")).map(|v| v == "1" || v == "true");
    let def_font_family = def_rpr.and_then(|n| child(n, "latin")).and_then(|n| attr(&n, "typeface"))
        .map(|tf| resolve_theme_typeface(&tf, theme));

    let mut runs = Vec::new();
    for node in p_node.children().filter(|n| n.is_element()) {
        match node.tag_name().name() {
            "r" => {
                if let Some(run) = parse_run(node, def_rpr, theme) {
                    runs.push(TextRun::Text(run));
                }
            }
            "br" => runs.push(TextRun::Break),
            // Field elements (e.g. slide number, date): parse like a run but tag the field type
            "fld" => {
                let fld_type = attr(&node, "type").unwrap_or_default().to_string();
                let text = child(node, "t").and_then(|t| t.text()).unwrap_or("").to_string();
                let r_pr = child(node, "rPr");
                let font_size = r_pr.and_then(|n| attr_f64(&n, "sz")).map(|v| v / 100.0);
                let color = r_pr.and_then(|n| child(n, "solidFill")).and_then(|n| parse_color_node(n, theme));
                let bold = r_pr.and_then(|n| attr(&n, "b")).map(|v| v == "1" || v == "true");
                let italic = r_pr.and_then(|n| attr(&n, "i")).map(|v| v == "1" || v == "true");
                let font_family = r_pr.and_then(|n| child(n, "latin")).and_then(|n| attr(&n, "typeface"))
                    .map(|tf| resolve_theme_typeface(&tf, theme));
                runs.push(TextRun::Text(TextRunData {
                    text,
                    bold,
                    italic,
                    underline: false,
                    strikethrough: false,
                    font_size,
                    color,
                    font_family,
                    baseline: None,
                    field_type: if fld_type == "slidenum" { Some("slidenum".to_string()) } else { None },
                }));
            }
            _ => {}
        }
    }

    // For paragraphs with no visible text content, use endParaRPr sz to set line height.
    // This ensures empty spacer paragraphs have the correct height (e.g. between sections).
    let end_rpr = child(p_node, "endParaRPr");
    let has_text = runs.iter().any(|r| matches!(r, TextRun::Text(_)));
    let def_font_size = def_font_size.or_else(|| {
        if !has_text {
            end_rpr.and_then(|n| attr_f64(&n, "sz")).map(|v| v / 100.0)
        } else {
            None
        }
    });

    Paragraph {
        alignment, mar_l, mar_r, indent,
        space_before, space_after, space_line,
        lvl, bullet,
        def_font_size, def_color, def_bold, def_italic, def_font_family,
        tab_stops, runs,
    }
}

/// Parse bullet specification from pPr node.
fn parse_bullet(
    p_pr: Option<roxmltree::Node<'_, '_>>,
    theme: &HashMap<String, String>,
) -> Bullet {
    let p_pr = match p_pr {
        Some(n) => n,
        None    => return Bullet::Inherit,
    };

    // Explicit "no bullet"
    if child(p_pr, "buNone").is_some() {
        return Bullet::None;
    }

    // Character bullet
    if let Some(bu_char) = child(p_pr, "buChar") {
        let ch = attr(&bu_char, "char").unwrap_or_else(|| "\u{2022}".into()); // •
        let color = child(p_pr, "buClr").and_then(|n| parse_color_node(n, theme));
        // buSzPct val is in thousandths of a percent: 100000 = 100%
        let size_pct = child(p_pr, "buSzPct")
            .and_then(|n| attr_f64(&n, "val"))
            .map(|v| v / 1000.0);
        let font_family = child(p_pr, "buFont").and_then(|n| attr(&n, "typeface"))
            .map(|tf| resolve_theme_typeface(&tf, theme));
        return Bullet::Char { ch, color, size_pct, font_family };
    }

    // Auto-numbered bullet
    if let Some(bu_auto) = child(p_pr, "buAutoNum") {
        let num_type = attr(&bu_auto, "type").unwrap_or_else(|| "arabicPeriod".into());
        let start_at = attr(&bu_auto, "startAt").and_then(|v| v.parse().ok());
        return Bullet::AutoNum { num_type, start_at };
    }

    Bullet::Inherit
}

fn parse_run(
    r_node: roxmltree::Node<'_, '_>,
    def_rpr: Option<roxmltree::Node<'_, '_>>,
    theme: &HashMap<String, String>,
) -> Option<TextRunData> {
    let t_node = child(r_node, "t")?;
    let text  = t_node.text().unwrap_or("").to_owned();
    let r_pr  = child(r_node, "rPr");

    // Attribute with rPr → defRPr fallback; None means "not set" (inherit from body/layout defaults)
    let bold = r_pr.and_then(|n| attr(&n, "b"))
        .or_else(|| def_rpr.and_then(|n| attr(&n, "b")))
        .map(|v| v == "1" || v == "true");
    let italic = r_pr.and_then(|n| attr(&n, "i"))
        .or_else(|| def_rpr.and_then(|n| attr(&n, "i")))
        .map(|v| v == "1" || v == "true");
    let underline = r_pr.and_then(|n| attr(&n, "u"))
        .or_else(|| def_rpr.and_then(|n| attr(&n, "u")))
        .map(|v| v != "none").unwrap_or(false);

    // strikethrough: "sngStrike" or "dblStrike" → true
    let strikethrough = r_pr.and_then(|n| attr(&n, "strike"))
        .or_else(|| def_rpr.and_then(|n| attr(&n, "strike")))
        .map(|v| v == "sngStrike" || v == "dblStrike")
        .unwrap_or(false);

    // sz in hundredths of a point
    let font_size = r_pr.and_then(|n| attr_f64(&n, "sz"))
        .or_else(|| def_rpr.and_then(|n| attr_f64(&n, "sz")))
        .map(|v| v / 100.0);

    let color = r_pr.and_then(|n| child(n, "solidFill"))
        .and_then(|n| parse_color_node(n, theme))
        .or_else(|| {
            def_rpr.and_then(|n| child(n, "solidFill"))
                   .and_then(|n| parse_color_node(n, theme))
        });

    let font_family = r_pr.and_then(|n| child(n, "latin")).and_then(|n| attr(&n, "typeface"))
        .or_else(|| def_rpr.and_then(|n| child(n, "latin")).and_then(|n| attr(&n, "typeface")))
        .map(|tf| resolve_theme_typeface(&tf, theme));

    // baseline in thousandths of a point; 30000=superscript, -25000=subscript (OOXML typical)
    let baseline = r_pr.and_then(|n| attr(&n, "baseline"))
        .and_then(|v| v.parse::<i32>().ok())
        .filter(|&v| v != 0);

    Some(TextRunData { text, bold, italic, underline, strikethrough, font_size, color, font_family, baseline, field_type: None })
}

// ===========================
//  Chart parsing
// ===========================

/// Parse a legacy OOXML chart (c: namespace) — barChart / lineChart etc.
fn parse_legacy_chart(xml: &str, theme: &HashMap<String, String>) -> Option<ChartElement> {
    let doc = roxmltree::Document::parse(xml).ok()?;
    let root = doc.root_element();

    // Determine chart type by finding the first recognized chart element
    let find_chart = |name: &str| root.descendants()
        .find(|n| n.is_element() && n.tag_name().name() == name);

    let chart_type = if let Some(bc) = find_chart("barChart") {
        let grouping = bc.children()
            .find(|c| c.is_element() && c.tag_name().name() == "grouping")
            .and_then(|n| attr(&n, "val"))
            .unwrap_or_else(|| "clustered".into());
        let bar_dir = bc.children()
            .find(|c| c.is_element() && c.tag_name().name() == "barDir")
            .and_then(|n| attr(&n, "val"))
            .unwrap_or_else(|| "col".into());
        let horizontal = bar_dir == "bar";
        match (grouping.as_str(), horizontal) {
            ("stacked" | "percentStacked", false) => "stackedBar".to_string(),
            ("stacked" | "percentStacked", true)  => "stackedBarH".to_string(),
            (_, false) => "clusteredBar".to_string(),
            (_, true)  => "clusteredBarH".to_string(),
        }
    } else if let Some(lc) = find_chart("lineChart") {
        let grouping = lc.children()
            .find(|c| c.is_element() && c.tag_name().name() == "grouping")
            .and_then(|n| attr(&n, "val"))
            .unwrap_or_else(|| "standard".into());
        match grouping.as_str() {
            "stacked" => "stackedLine".to_string(),
            "percentStacked" => "stackedLinePct".to_string(),
            _ => "line".to_string(),
        }
    } else if find_chart("pieChart").is_some() {
        "pie".to_string()
    } else if find_chart("doughnutChart").is_some() {
        "doughnut".to_string()
    } else if let Some(ac) = find_chart("areaChart") {
        let grouping = ac.children()
            .find(|c| c.is_element() && c.tag_name().name() == "grouping")
            .and_then(|n| attr(&n, "val"))
            .unwrap_or_else(|| "standard".into());
        match grouping.as_str() {
            "stacked" => "stackedArea".to_string(),
            _ => "area".to_string(),
        }
    } else if find_chart("scatterChart").is_some() {
        "scatter".to_string()
    } else if find_chart("bubbleChart").is_some() {
        "bubble".to_string()
    } else if find_chart("radarChart").is_some() {
        "radar".to_string()
    } else {
        "unknown".to_string()
    };

    // Title text
    let title = root.descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "title")
        .and_then(|title_node| {
            let texts: Vec<String> = title_node.descendants()
                .filter(|n| n.is_element() && n.tag_name().name() == "t")
                .filter_map(|n| n.text().map(|t| t.to_string()))
                .collect();
            if texts.is_empty() { None } else { Some(texts.join("")) }
        });

    // val axis max
    let val_max = root.descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "valAx")
        .and_then(|ax| ax.descendants().find(|n| n.is_element() && n.tag_name().name() == "max"))
        .and_then(|n| attr(&n, "val"))
        .and_then(|v| v.parse::<f64>().ok());

    // Series
    let plot_area = root.descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "plotArea")?;

    let ser_nodes: Vec<_> = plot_area.descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "ser")
        .collect();

    if ser_nodes.is_empty() {
        return None;
    }

    // Helper: collect <c:pt> values from a cache node (strCache or numCache)
    let collect_pt_strings = |cache: roxmltree::Node<'_, '_>| -> Vec<String> {
        cache.children()
            .filter(|n| n.is_element() && n.tag_name().name() == "pt")
            .filter_map(|pt| pt.children().find(|n| n.is_element() && n.tag_name().name() == "v"))
            .filter_map(|v| v.text().map(|t| t.to_string()))
            .collect()
    };

    // Categories from first series's <c:cat> — supports strCache and numCache
    let categories: Vec<String> = ser_nodes[0]
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "cat")
        .and_then(|cat| {
            cat.descendants()
                .find(|n| n.is_element() && (n.tag_name().name() == "strCache" || n.tag_name().name() == "numCache"))
        })
        .map(|cache| collect_pt_strings(cache))
        .unwrap_or_default();

    let pt_count = categories.len().max(1);

    let series: Vec<ChartSeriesData> = ser_nodes.iter().map(|ser| {
        // Series name from <c:tx>
        let name = ser.children()
            .find(|n| n.is_element() && n.tag_name().name() == "tx")
            .and_then(|tx| tx.descendants().find(|n| n.is_element() && (n.tag_name().name() == "strCache" || n.tag_name().name() == "numCache")))
            .and_then(|cache| {
                cache.children()
                    .find(|n| n.is_element() && n.tag_name().name() == "pt")
                    .and_then(|pt| pt.children().find(|n| n.is_element() && n.tag_name().name() == "v"))
                    .and_then(|v| v.text().map(|t| t.to_string()))
            })
            .unwrap_or_default();

        // Values: use <c:val> or <c:yVal> (scatter), falling back to <c:numCache> anywhere
        let val_cache = ser.descendants()
            .find(|n| n.is_element() && (n.tag_name().name() == "val" || n.tag_name().name() == "yVal"))
            .and_then(|v| v.descendants().find(|n| n.is_element() && n.tag_name().name() == "numCache"))
            .or_else(|| ser.descendants().find(|n| n.is_element() && n.tag_name().name() == "numCache"));

        let mut values: Vec<Option<f64>> = vec![None; pt_count];
        if let Some(cache) = val_cache {
            for pt in cache.children().filter(|n| n.is_element() && n.tag_name().name() == "pt") {
                let idx: usize = attr(&pt, "idx").and_then(|v| v.parse().ok()).unwrap_or(0);
                let val: Option<f64> = pt.children()
                    .find(|n| n.is_element() && n.tag_name().name() == "v")
                    .and_then(|v| v.text())
                    .and_then(|t| t.parse().ok());
                if idx < values.len() {
                    values[idx] = val;
                }
            }
        }

        // Series color from spPr > solidFill
        let color = ser.children()
            .find(|n| n.is_element() && n.tag_name().name() == "spPr")
            .and_then(|sp| sp.children().find(|n| n.is_element() && n.tag_name().name() == "solidFill"))
            .and_then(|fill| parse_color_node(fill, theme));

        // Per-data-point colors from <c:dPt> (important for pie charts)
        let data_point_colors: Vec<Option<String>> = (0..pt_count).map(|i| {
            ser.children()
                .filter(|n| n.is_element() && n.tag_name().name() == "dPt")
                .find(|dpt| attr(dpt, "idx").and_then(|v| v.parse::<usize>().ok()) == Some(i))
                .and_then(|dpt| dpt.descendants().find(|n| n.is_element() && n.tag_name().name() == "solidFill"))
                .and_then(|fill| parse_color_node(fill, theme))
        }).collect();

        let has_dpt_colors = data_point_colors.iter().any(|c| c.is_some());
        ChartSeriesData {
            name, values, color,
            data_point_colors: if has_dpt_colors { Some(data_point_colors) } else { None },
        }
    }).collect();

    // Check if data labels (showVal) are enabled — at chart level or in any series
    let show_data_labels = root.descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "dLbls")
        .any(|dLbls| {
            dLbls.children()
                .any(|c| c.is_element() && c.tag_name().name() == "showVal"
                    && attr(&c, "val").as_deref() == Some("1"))
        });

    Some(ChartElement {
        x: 0, y: 0, width: 0, height: 0,
        chart_type,
        title,
        categories,
        series,
        val_max,
        subtotal_indices: vec![],
        show_data_labels,
    })
}

/// Parse a modern chartEx (cx: namespace) — waterfall, treemap, etc.
fn parse_chartex(xml: &str, theme: &HashMap<String, String>) -> Option<ChartElement> {
    let doc = roxmltree::Document::parse(xml).ok()?;
    let root = doc.root_element();

    // Chart type from series layoutId attribute
    let series_node = root.descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "series")?;
    let layout_id = attr(&series_node, "layoutId").unwrap_or_default();
    let chart_type = layout_id; // "waterfall", "treemap", etc.

    // Categories from chartData > data > strDim[@type="cat"] > lvl > pt
    let categories: Vec<String> = root.descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "strDim"
            && attr(n, "type").as_deref() == Some("cat"))
        .and_then(|dim| dim.descendants().find(|n| n.is_element() && n.tag_name().name() == "lvl"))
        .map(|lvl| {
            lvl.children()
                .filter(|n| n.is_element() && n.tag_name().name() == "pt")
                .filter_map(|pt| pt.text().map(|t| t.replace('\n', " ")))
                .collect()
        })
        .unwrap_or_default();

    let pt_count = categories.len().max(1);

    // Values from chartData > data > numDim[@type="val"] > lvl > pt
    let raw_values: Vec<Option<f64>> = root.descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "numDim"
            && attr(n, "type").as_deref() == Some("val"))
        .and_then(|dim| dim.descendants().find(|n| n.is_element() && n.tag_name().name() == "lvl"))
        .map(|lvl| {
            let mut vals: Vec<Option<f64>> = vec![None; pt_count];
            for (i, pt) in lvl.children()
                .filter(|n| n.is_element() && n.tag_name().name() == "pt")
                .enumerate()
            {
                if i < vals.len() {
                    vals[i] = pt.text().and_then(|t| t.parse().ok());
                }
            }
            vals
        })
        .unwrap_or_else(|| vec![None; pt_count]);

    // Subtotal indices (idx=0 is always implicit; add from cx:subtotals)
    let mut subtotal_indices: Vec<u32> = vec![0];
    if let Some(subtotals_node) = series_node.descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "subtotals") {
        for idx_node in subtotals_node.children()
            .filter(|n| n.is_element() && n.tag_name().name() == "idx") {
            if let Some(v) = attr(&idx_node, "val").and_then(|v| v.parse::<u32>().ok()) {
                if v != 0 {
                    subtotal_indices.push(v);
                }
            }
        }
    }

    // Series color (first dataPt or series spPr)
    let color = series_node.children()
        .find(|n| n.is_element() && n.tag_name().name() == "spPr")
        .and_then(|sp| sp.children().find(|n| n.is_element() && n.tag_name().name() == "solidFill"))
        .and_then(|fill| parse_color_node(fill, theme));

    let series = vec![ChartSeriesData {
        name: String::new(),
        values: raw_values,
        color,
        data_point_colors: None,
    }];

    Some(ChartElement {
        x: 0, y: 0, width: 0, height: 0,
        chart_type,
        title: None,
        categories,
        series,
        val_max: None,
        subtotal_indices,
        show_data_labels: false,
    })
}

// ===========================
//  Placeholder defaults
// ===========================

/// OOXML spec default positions for common placeholder types.
/// Values are in EMU, assuming a 9144000×6858000 slide (10"×7.5").
fn default_placeholder_transform(ph_type: &str) -> Transform {
    match ph_type {
        "title" | "ctrTitle" => Transform {
            x: 457200,   y: 274638,   cx: 8229600, cy: 1143000, rot: 0.0, flip_h: false, flip_v: false,
        },
        "subTitle" => Transform {
            x: 457200,   y: 1600200,  cx: 8229600, cy: 899160,  rot: 0.0, flip_h: false, flip_v: false,
        },
        "dt" => Transform {
            x: 0,        y: 6261600,  cx: 2286000, cy: 596900,  rot: 0.0, flip_h: false, flip_v: false,
        },
        "ftr" => Transform {
            x: 2972400,  y: 6261600,  cx: 3086100, cy: 596900,  rot: 0.0, flip_h: false, flip_v: false,
        },
        "sldNum" => Transform {
            x: 6629400,  y: 6261600,  cx: 2057400, cy: 596900,  rot: 0.0, flip_h: false, flip_v: false,
        },
        // "body" and everything else: full-width content area below title
        _ => Transform {
            x: 457200,   y: 1600200,  cx: 8229600, cy: 4525963, rot: 0.0, flip_h: false, flip_v: false,
        },
    }
}

// ===========================
//  Placeholder detection
// ===========================

/// Returns true if the node contains a `p:ph` descendant.
fn is_placeholder(node: roxmltree::Node<'_, '_>) -> bool {
    node.descendants()
        .any(|n| n.is_element() && n.tag_name().name() == "ph")
}

// ===========================
//  Shape parsing  (p:sp)
// ===========================

fn parse_shape(
    sp_node: roxmltree::Node<'_, '_>,
    lph: &LayoutPlaceholders,
    theme: &HashMap<String, String>,
    group_fill: Option<&Fill>,
) -> Option<ShapeElement> {
    // --- Placeholder info (for layout fallback) ---
    let ph_node = sp_node
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "ph");
    let ph_type = ph_node
        .as_ref()
        .and_then(|n| attr(n, "type"))
        .unwrap_or_else(|| "body".into());
    let ph_idx: Option<u32> = ph_node
        .as_ref()
        .and_then(|n| attr(n, "idx"))
        .and_then(|v| v.parse().ok());

    // --- Transform: slide xfrm OR layout fallback ---
    let sp_pr = child(sp_node, "spPr");
    let slide_xfrm = sp_pr.and_then(|p| child(p, "xfrm"));

    let t: Transform = if let Some(xfrm) = slide_xfrm {
        parse_xfrm(xfrm)
    } else if ph_node.is_some() {
        match lph.lookup(&ph_type, ph_idx) {
            Some(lt) => lt.clone(),
            None => default_placeholder_transform(&ph_type),
        }
    } else {
        return None; // non-placeholder with no xfrm — skip
    };

    // cx=0 → skip.
    // cy=0 means "auto-height": keep 0 when anchor="b" (renderer grows shape upward from off_y),
    // otherwise use a generous fallback so text has room to render.
    if t.cx == 0 {
        return None;
    }
    let inherited_anchor: Option<String> = if ph_node.is_some() {
        lph.lookup_anchor(&ph_type)
    } else {
        None
    };
    let is_bottom_anchor = inherited_anchor.as_deref() == Some("b")
        || child(sp_node, "txBody")
            .and_then(|tb| child(tb, "bodyPr"))
            .and_then(|bp| attr(&bp, "anchor"))
            .map(|a| a == "b")
            .unwrap_or(false);
    let cy = if t.cy == 0 {
        if is_bottom_anchor { 0_i64 } else { 2_000_000_i64 }
    } else {
        t.cy
    };

    // custGeom takes priority over prstGeom
    let cust_geom_node = sp_pr.and_then(|p| child(p, "custGeom"));
    let prst_geom_node = sp_pr.and_then(|p| child(p, "prstGeom"));
    let geometry = if cust_geom_node.is_some() {
        "custGeom".into()
    } else {
        prst_geom_node
            .and_then(|n| attr(&n, "prst"))
            .unwrap_or_else(|| "rect".into())
    };
    let cust_geom = cust_geom_node.map(|n| parse_cust_geom(n));

    // Parse adjustment values from prstGeom avLst (e.g. trapezoid inset)
    // Collect all gd elements; first is adj (name="adj" or "adj1"), second is adj2
    let parse_gd_val = |gd: roxmltree::Node<'_, '_>| -> Option<f64> {
        attr(&gd, "fmla")
            .and_then(|f| f.strip_prefix("val ").map(|s| s.to_owned()))
            .and_then(|s| s.parse::<f64>().ok())
    };
    let av_node = prst_geom_node.and_then(|n| child(n, "avLst"));
    let gd_nodes: Vec<_> = av_node
        .map(|av| av.children().filter(|n| n.is_element() && n.tag_name().name() == "gd").collect())
        .unwrap_or_default();
    // First gd = adj (match by name "adj" or "adj1", fallback to position 0)
    let adj: Option<f64> = gd_nodes
        .iter()
        .find(|n| matches!(attr(n, "name").as_deref(), Some("adj") | Some("adj1")))
        .or_else(|| gd_nodes.first())
        .and_then(|n| parse_gd_val(*n));
    // Second gd = adj2 (match by name "adj2", fallback to position 1)
    let adj2: Option<f64> = gd_nodes
        .iter()
        .find(|n| attr(n, "name").as_deref() == Some("adj2"))
        .or_else(|| gd_nodes.get(1))
        .and_then(|n| parse_gd_val(*n));
    // Third gd = adj3 (match by name "adj3", fallback to position 2)
    let adj3: Option<f64> = gd_nodes
        .iter()
        .find(|n| attr(n, "name").as_deref() == Some("adj3"))
        .or_else(|| gd_nodes.get(2))
        .and_then(|n| parse_gd_val(*n));
    // Fourth gd = adj4 (match by name "adj4", fallback to position 3)
    let adj4: Option<f64> = gd_nodes
        .iter()
        .find(|n| attr(n, "name").as_deref() == Some("adj4"))
        .or_else(|| gd_nodes.get(3))
        .and_then(|n| parse_gd_val(*n));

    // --- Shape style (p:style) provides fill/stroke/text-color fallbacks ---
    let style_node = child(sp_node, "style");

    // fillRef idx=0 → explicit no-fill; idx>0 → use referenced color as solid fill
    let style_fill: Option<Fill> = style_node
        .and_then(|s| child(s, "fillRef"))
        .and_then(|fr| {
            let idx: u32 = attr(&fr, "idx").and_then(|v| v.parse().ok()).unwrap_or(1);
            if idx == 0 {
                Some(Fill::None)
            } else {
                parse_color_node(fr, theme).map(|c| Fill::Solid { color: c })
            }
        });

    // lnRef idx=0 → no line; idx>0 → use referenced color
    let style_stroke: Option<Stroke> = style_node
        .and_then(|s| child(s, "lnRef"))
        .and_then(|lr| {
            let idx: u32 = attr(&lr, "idx").and_then(|v| v.parse().ok()).unwrap_or(1);
            if idx == 0 { None } else {
                parse_color_node(lr, theme).map(|c| Stroke { color: c, width: 9525, dash_style: None })
            }
        });

    // fontRef → default text color for this shape
    let default_text_color: Option<String> = style_node
        .and_then(|s| child(s, "fontRef"))
        .and_then(|fr| parse_color_node(fr, theme));

    // spPr fill: grpFill means inherit from parent group; explicit fill overrides style.
    // Note: Some(Fill::None) (noFill in spPr) must NOT be overridden by style.
    let sp_pr_has_grp_fill = sp_pr.and_then(|p| child(p, "grpFill")).is_some();
    let fill = if sp_pr_has_grp_fill {
        group_fill.cloned()
    } else {
        sp_pr.and_then(|p| parse_fill(p, theme)).or(style_fill)
    };

    // spPr stroke: if ln element is present, respect it (even if noFill → None);
    // otherwise fall back to layout placeholder stroke, then style stroke.
    let stroke = if sp_pr.and_then(|p| child(p, "ln")).is_some() {
        sp_pr.and_then(|p| child(p, "ln")).and_then(|n| parse_stroke(n, theme))
    } else if ph_node.is_some() {
        lph.lookup_stroke(&ph_type, ph_idx).or(style_stroke)
    } else {
        style_stroke
    };

    // Inherited defaults from layout/master for this placeholder type/idx
    let (inherited_font_size, inherited_bold, inherited_italic, inherited_anchor, inherited_alignment,
         inherited_space_before, inherited_space_after, inherited_line_spacing) = if ph_node.is_some() {
        (
            lph.lookup_font_size(&ph_type, ph_idx),
            lph.lookup_bold(&ph_type),
            lph.lookup_italic(&ph_type),
            lph.lookup_anchor(&ph_type),
            lph.lookup_alignment(&ph_type),
            lph.lookup_space_before(&ph_type),
            lph.lookup_space_after(&ph_type),
            lph.lookup_line_spacing(&ph_type, ph_idx),
        )
    } else {
        (None, None, None, None, None, None, None, None)
    };

    let text_body = child(sp_node, "txBody")
        .map(|n| parse_text_body(n, theme, inherited_font_size, inherited_bold, inherited_italic, inherited_anchor, inherited_alignment, inherited_space_before, inherited_space_after, inherited_line_spacing));

    // Shadow from spPr > effectLst > outerShdw
    let shadow = sp_pr
        .and_then(|p| child(p, "effectLst"))
        .and_then(|n| parse_shadow(n, theme));

    Some(ShapeElement {
        x: t.x, y: t.y, width: t.cx, height: cy,
        rotation: t.rot, flip_h: t.flip_h, flip_v: t.flip_v,
        geometry, fill, stroke, text_body, default_text_color, cust_geom, adj, adj2, adj3, adj4, shadow,
    })
}

// ===========================
//  Picture parsing  (p:pic)
// ===========================

fn parse_picture(
    pic_node: roxmltree::Node<'_, '_>,
    slide_dir: &str,
    rels: &HashMap<String, String>,
    zip: &mut PptxZip<'_>,
) -> Option<PictureElement> {
    let sp_pr = child(pic_node, "spPr")?;
    let xfrm_node = child(sp_pr, "xfrm")?;
    let t = parse_xfrm(xfrm_node);

    if t.cx == 0 || t.cy == 0 {
        return None; // pictures always need explicit dimensions
    }

    let r_id = child(pic_node, "blipFill")
        .and_then(|bf| child(bf, "blip"))
        .and_then(|b| attr_r(&b, "embed"))?;

    let rel_target = rels.get(&r_id)?;
    let image_path = resolve_path(slide_dir, rel_target);

    let image_bytes = read_zip_bytes(zip, &image_path)?;
    let mime = mime_from_ext(&image_path);
    let data_url = format!("data:{mime};base64,{}", B64.encode(&image_bytes));

    Some(PictureElement {
        x: t.x, y: t.y, width: t.cx, height: t.cy,
        rotation: t.rot, flip_h: t.flip_h, flip_v: t.flip_v, data_url,
    })
}

fn mime_from_ext(path: &str) -> &'static str {
    match path.rsplit('.').next().unwrap_or("").to_ascii_lowercase().as_str() {
        "png"  => "image/png",
        "jpg" | "jpeg" => "image/jpeg",
        "gif"  => "image/gif",
        "bmp"  => "image/bmp",
        "svg"  => "image/svg+xml",
        "webp" => "image/webp",
        _      => "application/octet-stream",
    }
}

// ===========================
//  Slide background
// ===========================

fn parse_background(
    c_sld: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<Fill> {
    let bg = child(c_sld, "bg")?;
    // bgPr contains an explicit fill specification
    if let Some(bg_pr) = child(bg, "bgPr") {
        return parse_fill(bg_pr, theme);
    }
    // bgRef references a theme background style; its child is a color element
    if let Some(bg_ref) = child(bg, "bgRef") {
        return parse_color_node(bg_ref, theme).map(|c| Fill::Solid { color: c });
    }
    None
}

// ===========================
//  Table parsing
// ===========================

/// Parse ppt/tableStyles.xml into a map of styleId → TableStyleDef.
fn parse_table_styles_xml(xml: &str, theme: &HashMap<String, String>) -> HashMap<String, TableStyleDef> {
    let mut map = HashMap::new();
    let Ok(doc) = roxmltree::Document::parse(xml) else { return map; };
    let root = doc.root_element();
    for style_node in root.children().filter(|n| n.is_element() && n.tag_name().name() == "tblStyle") {
        let Some(style_id) = attr(&style_node, "styleId") else { continue };
        let style_id = style_id.to_string();
        let mut def = TableStyleDef::default();

        let parse_tc_style = |role: roxmltree::Node<'_, '_>|
            -> (Option<Fill>, Option<Stroke>, Option<Stroke>, Option<Stroke>, Option<Stroke>, Option<Stroke>)
        {
            let tc_style = match child(role, "tcStyle") {
                Some(n) => n,
                None    => return (None, None, None, None, None, None),
            };
            let fill = parse_fill(tc_style, theme);
            let tc_bdr = child(tc_style, "tcBdr");
            let parse_side = |side: &str| -> Option<Stroke> {
                tc_bdr.and_then(|b| child(b, side))
                    .and_then(|n| child(n, "ln"))
                    .and_then(|n| parse_stroke(n, theme))
            };
            let inside_h = parse_side("insideH");
            let inside_v = parse_side("insideV");
            let border_b = parse_side("bottom");
            let outer_h  = parse_side("top");
            let outer_v  = parse_side("left");
            (fill, inside_h, inside_v, border_b, outer_h, outer_v)
        };

        if let Some(whole) = child(style_node, "wholeTbl") {
            let (fill, ih, iv, _, oh, ov) = parse_tc_style(whole);
            def.whole_fill = fill;
            def.whole_inside_h = ih;
            def.whole_inside_v = iv;
            def.whole_outer_h  = oh;
            def.whole_outer_v  = ov;
        }
        if let Some(band) = child(style_node, "band1H") {
            let (fill, _, _, _, _, _) = parse_tc_style(band);
            def.band1h_fill = fill;
        }
        if let Some(band) = child(style_node, "band2H") {
            let (fill, _, _, _, _, _) = parse_tc_style(band);
            def.band2h_fill = fill;
        }
        if let Some(first) = child(style_node, "firstRow") {
            let (fill, _, _, border_b, _, _) = parse_tc_style(first);
            def.first_row_fill = fill;
            def.first_row_border_b = border_b;
        }
        if let Some(last) = child(style_node, "lastRow") {
            let (fill, _, _, _, _, _) = parse_tc_style(last);
            def.last_row_fill = fill;
        }
        if let Some(first) = child(style_node, "firstCol") {
            let (fill, _, _, _, _, _) = parse_tc_style(first);
            def.first_col_fill = fill;
        }
        if let Some(last) = child(style_node, "lastCol") {
            let (fill, _, _, _, _, _) = parse_tc_style(last);
            def.last_col_fill = fill;
        }

        map.insert(style_id, def);
    }
    map
}

fn parse_table(
    tbl: roxmltree::Node<'_, '_>,
    t: &Transform,
    theme: &HashMap<String, String>,
    zip: &mut PptxZip<'_>,
) -> Option<TableElement> {
    // Parse tblPr attributes and look up table style
    let tbl_pr = child(tbl, "tblPr");
    let style_id = tbl_pr
        .and_then(|n| child(n, "tableStyleId"))
        .and_then(|n| n.text())
        .map(|s| s.to_string());
    let flag = |attr_name: &str| -> bool {
        tbl_pr.and_then(|n| attr(&n, attr_name))
            .map(|v| v == "1" || v == "true").unwrap_or(false)
    };
    let first_row = flag("firstRow");
    let last_row  = flag("lastRow");
    let band_row  = flag("bandRow");
    let first_col = flag("firstCol");
    let last_col  = flag("lastCol");

    // Load style definitions once
    let table_styles_xml = read_zip_str(zip, "ppt/tableStyles.xml").ok();
    let table_styles = table_styles_xml.as_deref()
        .map(|xml| parse_table_styles_xml(xml, theme))
        .unwrap_or_default();
    let style = style_id.as_deref().and_then(|id| table_styles.get(id));

    let cols: Vec<i64> = tbl
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "tblGrid")
        .map(|grid| {
            grid.children()
                .filter(|n| n.is_element() && n.tag_name().name() == "gridCol")
                .filter_map(|n| attr_i64(&n, "w"))
                .collect()
        })
        .unwrap_or_default();

    if cols.is_empty() {
        return None;
    }

    let col_count = cols.len();
    let last_col_idx = col_count.saturating_sub(1);

    let mut rows: Vec<TableRow> = tbl
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "tr")
        .map(|tr| parse_table_row(tr, theme))
        .collect();

    let row_count = rows.len();
    let last_row_idx = row_count.saturating_sub(1);

    // Apply table style fills and borders to each cell
    for (ri, row) in rows.iter_mut().enumerate() {
        for (ci, cell) in row.cells.iter_mut().enumerate() {
            if let Some(s) = style {
                // ── Fill cascade ────────────────────────────────────────────
                let mut effective_fill = s.whole_fill.clone();

                if band_row {
                    // Determine band index excluding firstRow header if present
                    let band_ri = ri.saturating_sub(if first_row { 1 } else { 0 });
                    if !(first_row && ri == 0) {
                        if band_ri % 2 == 0 {
                            if let Some(f) = s.band1h_fill.clone() { effective_fill = Some(f); }
                        } else if let Some(f) = s.band2h_fill.clone() {
                            effective_fill = Some(f);
                        }
                    }
                }
                if first_row && ri == 0 {
                    if let Some(f) = s.first_row_fill.clone() { effective_fill = Some(f); }
                }
                if last_row && ri == last_row_idx {
                    if let Some(f) = s.last_row_fill.clone() { effective_fill = Some(f); }
                }
                if first_col && ci == 0 {
                    if let Some(f) = s.first_col_fill.clone() { effective_fill = Some(f); }
                }
                if last_col && ci == last_col_idx {
                    if let Some(f) = s.last_col_fill.clone() { effective_fill = Some(f); }
                }
                // Cell's own tcPr fill wins
                if cell.fill.is_none() {
                    cell.fill = effective_fill;
                }

                // ── Border cascade (style provides inside and outer borders) ──
                // Outer top edge
                if cell.border_t.is_none() && ri == 0 {
                    cell.border_t = s.whole_outer_h.clone();
                }
                // Inner horizontal separator between rows
                if cell.border_t.is_none() && ri > 0 {
                    cell.border_t = s.whole_inside_h.clone();
                }
                // Outer bottom edge
                if cell.border_b.is_none() && ri == last_row_idx {
                    cell.border_b = s.whole_outer_h.clone();
                }
                // Inner bottom separator; firstRow gets its own bottom definition
                if cell.border_b.is_none() {
                    if first_row && ri == 0 {
                        cell.border_b = s.first_row_border_b.clone()
                            .or_else(|| s.whole_inside_h.clone());
                    } else if ri < last_row_idx {
                        cell.border_b = s.whole_inside_h.clone();
                    }
                }
                // Outer left edge
                if cell.border_l.is_none() && ci == 0 {
                    cell.border_l = s.whole_outer_v.clone();
                }
                // Inner vertical separator between cols
                if cell.border_l.is_none() && ci > 0 {
                    cell.border_l = s.whole_inside_v.clone();
                }
                // Outer right edge
                if cell.border_r.is_none() && ci == last_col_idx {
                    cell.border_r = s.whole_outer_v.clone();
                }
                // Inner right separator
                if cell.border_r.is_none() && ci < last_col_idx {
                    cell.border_r = s.whole_inside_v.clone();
                }
            } else {
                // ── Fallback for built-in styles not defined in tableStyles.xml ──
                // Approximate "Medium Style 2": accent1 header fill + thin outer box + row separators.
                let thin = Stroke { color: "A0A096".to_string(), width: 9525, dash_style: None };
                if cell.fill.is_none() && first_row && ri == 0 {
                    if let Some(color) = theme.get("accent1") {
                        cell.fill = Some(Fill::Solid { color: color.clone() });
                    }
                }
                // Outer top
                if cell.border_t.is_none() && ri == 0 {
                    cell.border_t = Some(thin.clone());
                }
                // Inner horizontal separators
                if cell.border_t.is_none() && ri > 0 {
                    cell.border_t = Some(thin.clone());
                }
                // Outer bottom
                if cell.border_b.is_none() && ri == last_row_idx {
                    cell.border_b = Some(thin.clone());
                }
                // Outer left edge
                if cell.border_l.is_none() && ci == 0 {
                    cell.border_l = Some(thin.clone());
                }
                // Outer right edge
                if cell.border_r.is_none() && ci == last_col_idx {
                    cell.border_r = Some(thin.clone());
                }
            }
        }
    }

    Some(TableElement {
        x: t.x, y: t.y, width: t.cx, height: t.cy,
        cols, rows,
    })
}

fn parse_table_row(
    tr: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> TableRow {
    let height = attr_i64(&tr, "h").unwrap_or(0);
    let cells: Vec<TableCell> = tr
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "tc")
        .map(|tc| parse_table_cell(tc, theme))
        .collect();
    TableRow { height, cells }
}

fn parse_table_cell(
    tc: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> TableCell {
    let tc_pr = child(tc, "tcPr");
    // tcPr > anchor controls vertical text alignment within the cell
    let anchor = tc_pr.and_then(|n| attr(&n, "anchor")).map(|a| a.to_string());
    let text_body = child(tc, "txBody").map(|n| parse_text_body(n, theme, None, None, None, anchor, None, None, None, None));

    let fill = tc_pr.and_then(|n| parse_fill(n, theme));

    let border_l = tc_pr.and_then(|n| child(n, "lnL")).and_then(|n| parse_stroke(n, theme));
    let border_r = tc_pr.and_then(|n| child(n, "lnR")).and_then(|n| parse_stroke(n, theme));
    let border_t = tc_pr.and_then(|n| child(n, "lnT")).and_then(|n| parse_stroke(n, theme));
    let border_b = tc_pr.and_then(|n| child(n, "lnB")).and_then(|n| parse_stroke(n, theme));

    let grid_span: u32 = attr(&tc, "gridSpan").and_then(|v| v.parse().ok()).unwrap_or(1);
    let row_span: u32  = attr(&tc, "rowSpan").and_then(|v| v.parse().ok()).unwrap_or(1);
    let h_merge = attr(&tc, "hMerge").map(|v| v == "1" || v == "true").unwrap_or(false);
    let v_merge = attr(&tc, "vMerge").map(|v| v == "1" || v == "true").unwrap_or(false);

    TableCell {
        text_body, fill,
        border_l, border_r, border_t, border_b,
        grid_span, row_span, h_merge, v_merge,
    }
}

// ===========================
//  Slide parser
// ===========================

fn parse_slide(
    xml: &str,
    layout_xml: Option<&str>,
    layout_rels: &HashMap<String, String>,
    layout_dir: &str,
    master_bg: Option<Fill>,
    master_font_sizes: &HashMap<String, f64>,
    master_anchors: &HashMap<String, String>,
    master_transforms: &HashMap<String, Transform>,
    master_alignments: &HashMap<String, String>,
    master_space_before: &HashMap<String, i64>,
    master_space_after: &HashMap<String, i64>,
    master_line_spacing: &HashMap<String, f64>,
    index: usize,
    rels: &HashMap<String, String>,
    zip: &mut PptxZip<'_>,
    theme: &HashMap<String, String>,
) -> Result<Slide, Box<dyn std::error::Error>> {
    let lph = layout_xml
        .map(|x| parse_layout_placeholders(x, master_font_sizes, master_anchors, master_transforms, master_alignments, master_space_before, master_space_after, master_line_spacing, theme))
        .unwrap_or_default();

    let doc = roxmltree::Document::parse(xml)?;
    let root = doc.root_element(); // <p:sld>
    let c_sld = child(root, "cSld");

    // Background chain: slide → layout → master
    let background = c_sld
        .and_then(|n| parse_background(n, theme))
        .or_else(|| {
            layout_xml.and_then(|lx| {
                let doc2 = roxmltree::Document::parse(lx).ok()?;
                child(doc2.root_element(), "cSld")
                    .and_then(|n| parse_background(n, theme))
            })
        })
        .or(master_bg);

    let sp_tree = c_sld
        .and_then(|n| child(n, "spTree"))
        .ok_or("missing spTree")?;

    let slide_dir = "ppt/slides";
    let mut elements = Vec::new();

    // ── Layout non-placeholder shapes (rendered BEFORE slide shapes) ──────
    // These are decorative background elements defined in the slide layout
    // (e.g. coloured bands, logos) that are not placeholder anchors.
    if let Some(lxml) = layout_xml {
        if let Ok(ldoc) = roxmltree::Document::parse(lxml) {
            let lroot = ldoc.root_element();
            if let Some(lsp_tree) = child(lroot, "cSld").and_then(|n| child(n, "spTree")) {
                let empty_lph = LayoutPlaceholders::default();
                for node in lsp_tree.children().filter(|n| n.is_element()) {
                    parse_sp_tree_node(
                        node, &empty_lph, layout_dir, layout_rels,
                        zip, theme, &mut elements,
                        true, // skip placeholder shapes
                        None, // no inherited group fill at top level
                    );
                }
            }
        }
    }

    // ── Slide shapes ─────────────────────────────────────────────────────
    for node in sp_tree.children().filter(|n| n.is_element()) {
        parse_sp_tree_node(node, &lph, slide_dir, rels, zip, theme, &mut elements, false, None);
    }

    Ok(Slide { index, slide_number: index + 1, background, elements })
}

fn parse_sp_tree_node(
    node: roxmltree::Node<'_, '_>,
    lph: &LayoutPlaceholders,
    slide_dir: &str,
    rels: &HashMap<String, String>,
    zip: &mut PptxZip<'_>,
    theme: &HashMap<String, String>,
    out: &mut Vec<SlideElement>,
    skip_placeholders: bool,
    group_fill: Option<&Fill>,
) {
    match node.tag_name().name() {
        "sp" => {
            if skip_placeholders && is_placeholder(node) {
                return;
            }
            // Image-filled shape: spPr > blipFill > blip r:embed → render as PictureElement
            let sp_pr_node = child(node, "spPr");
            let blip_rid = sp_pr_node
                .and_then(|p| child(p, "blipFill"))
                .and_then(|bf| child(bf, "blip"))
                .and_then(|b| attr_r(&b, "embed"));
            if let Some(rid) = blip_rid {
                if let Some(xfrm_node) = sp_pr_node.and_then(|p| child(p, "xfrm")) {
                    let t = parse_xfrm(xfrm_node);
                    if t.cx > 0 && t.cy > 0 {
                        if let Some(target) = rels.get(&rid) {
                            let image_path = resolve_path(slide_dir, target);
                            if let Some(bytes) = read_zip_bytes(zip, &image_path) {
                                let mime = mime_from_ext(&image_path);
                                let data_url = format!("data:{mime};base64,{}", B64.encode(&bytes));
                                out.push(SlideElement::Picture(PictureElement {
                                    x: t.x, y: t.y, width: t.cx, height: t.cy,
                                    rotation: t.rot, flip_h: t.flip_h, flip_v: t.flip_v, data_url,
                                }));
                                return;
                            }
                        }
                    }
                }
            }
            if let Some(shape) = parse_shape(node, lph, theme, group_fill) {
                out.push(SlideElement::Shape(shape));
            }
        }
        "pic" => {
            if let Some(pic) = parse_picture(node, slide_dir, rels, zip) {
                out.push(SlideElement::Picture(pic));
            } else {
                // Placeholder pic: no xfrm in spPr — position comes from layout by_idx
                let ph_idx = node.descendants()
                    .find(|n| n.is_element() && n.tag_name().name() == "ph")
                    .and_then(|ph| attr(&ph, "idx"))
                    .and_then(|s| s.parse::<u32>().ok());
                if let Some(idx) = ph_idx {
                    if let Some(t) = lph.by_idx.get(&idx) {
                        let r_id = child(node, "blipFill")
                            .and_then(|bf| child(bf, "blip"))
                            .and_then(|b| attr_r(&b, "embed"));
                        if let Some(rid) = r_id {
                            if let Some(rel_target) = rels.get(&rid) {
                                let image_path = resolve_path(slide_dir, rel_target);
                                if let Some(image_bytes) = read_zip_bytes(zip, &image_path) {
                                    let mime = mime_from_ext(&image_path);
                                    let data_url = format!("data:{mime};base64,{}", B64.encode(&image_bytes));
                                    out.push(SlideElement::Picture(PictureElement {
                                        x: t.x, y: t.y, width: t.cx, height: t.cy,
                                        rotation: t.rot, flip_h: t.flip_h, flip_v: t.flip_v,
                                        data_url,
                                    }));
                                }
                            }
                        }
                    }
                }
            }
        }
        "AlternateContent" => {
            // mc:AlternateContent wraps modern elements (e.g. chartEx inside grpSp).
            // Process mc:Choice first; mc:Fallback is a lower-fidelity version we skip.
            let choice = node.children()
                .find(|n| n.is_element() && n.tag_name().name() == "Choice")
                .or_else(|| node.children().find(|n| n.is_element() && n.tag_name().name() == "Fallback"));
            if let Some(choice_node) = choice {
                for child_node in choice_node.children().filter(|n| n.is_element()) {
                    parse_sp_tree_node(child_node, lph, slide_dir, rels, zip, theme, out, skip_placeholders, group_fill);
                }
            }
        }
        "graphicFrame" => {
            let xfrm_node = child(node, "xfrm");
            let t = xfrm_node.map(parse_xfrm).unwrap_or_default();

            // Table
            let tbl_node = node
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "tbl");
            if let Some(tbl_node) = tbl_node {
                if let Some(table) = parse_table(tbl_node, &t, theme, zip) {
                    out.push(SlideElement::Table(table));
                }
                return;
            }

            // Chart
            if let Some(gd) = node.descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "graphicData") {
                let uri = attr(&gd, "uri").unwrap_or_default();
                // Both <c:chart> and <cx:chart> share the local name "chart"
                let chart_rid = gd.descendants()
                    .find(|n| n.is_element() && n.tag_name().name() == "chart")
                    .and_then(|n| attr_r(&n, "id"));
                if let Some(rid) = chart_rid {
                    if let Some(rel_target) = rels.get(&rid) {
                        let chart_path = resolve_path(slide_dir, rel_target);
                        if let Ok(chart_xml) = read_zip_str(zip, &chart_path) {
                            let chart_opt = if uri.contains("chartex") || uri.contains("chartEx") {
                                parse_chartex(&chart_xml, theme)
                            } else {
                                parse_legacy_chart(&chart_xml, theme)
                            };
                            if let Some(mut chart) = chart_opt {
                                chart.x = t.x; chart.y = t.y;
                                chart.width = t.cx; chart.height = t.cy;
                                out.push(SlideElement::Chart(chart));
                            }
                        }
                    }
                }
            }
        }
        "grpSp" => {
            let grp_sp_pr = child(node, "grpSpPr");
            let gt: Option<GroupTransform> = grp_sp_pr
                .and_then(|pr| child(pr, "xfrm"))
                .map(|xfrm| {
                    let off    = child(xfrm, "off");
                    let ext    = child(xfrm, "ext");
                    let ch_off = child(xfrm, "chOff");
                    let ch_ext = child(xfrm, "chExt");
                    GroupTransform {
                        x:     off.and_then(|n| attr_i64(&n, "x")).unwrap_or(0),
                        y:     off.and_then(|n| attr_i64(&n, "y")).unwrap_or(0),
                        cx:    ext.and_then(|n| attr_i64(&n, "cx")).unwrap_or(0),
                        cy:    ext.and_then(|n| attr_i64(&n, "cy")).unwrap_or(0),
                        ch_x:  ch_off.and_then(|n| attr_i64(&n, "x")).unwrap_or(0),
                        ch_y:  ch_off.and_then(|n| attr_i64(&n, "y")).unwrap_or(0),
                        ch_cx: ch_ext.and_then(|n| attr_i64(&n, "cx")).unwrap_or(0),
                        ch_cy: ch_ext.and_then(|n| attr_i64(&n, "cy")).unwrap_or(0),
                        flip_h: attr(&xfrm, "flipH").map(|v| v == "1" || v == "true").unwrap_or(false),
                        flip_v: attr(&xfrm, "flipV").map(|v| v == "1" || v == "true").unwrap_or(false),
                        rot: attr_f64(&xfrm, "rot").unwrap_or(0.0) / 60000.0,
                    }
                });

            // Determine the fill to propagate to child shapes that use grpFill.
            // - Group has solidFill/noFill → use that as child group fill
            // - Group has grpFill → inherit from parent
            // - Group has no fill → inherit from parent
            let grp_has_grp_fill = grp_sp_pr.and_then(|pr| child(pr, "grpFill")).is_some();
            let grp_explicit_fill = grp_sp_pr.and_then(|pr| parse_fill(pr, theme));
            let child_group_fill: Option<Fill> = if grp_has_grp_fill {
                group_fill.cloned()
            } else if let Some(f) = grp_explicit_fill {
                Some(f)
            } else {
                group_fill.cloned()
            };

            let start = out.len();
            for child_node in node.children().filter(|n| n.is_element()) {
                parse_sp_tree_node(
                    child_node, lph, slide_dir, rels, zip, theme, out,
                    skip_placeholders, child_group_fill.as_ref(),
                );
            }
            if let Some(gt) = gt {
                for el in &mut out[start..] {
                    apply_group_transform_to_element(el, &gt);
                }
            }
        }
        "cxnSp" => {
            // Connector shape: parse as a line/shape element
            if skip_placeholders && is_placeholder(node) {
                return;
            }
            if let Some(shape) = parse_connector(node, theme) {
                out.push(SlideElement::Shape(shape));
            }
        }
        _ => {}
    }
}

/// Parse a connector shape (p:cxnSp) as a ShapeElement with line geometry.
fn parse_connector(
    node: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<ShapeElement> {
    let sp_pr = child(node, "spPr")?;
    let xfrm = child(sp_pr, "xfrm")?;
    let t = parse_xfrm(xfrm);
    if t.cx == 0 && t.cy == 0 {
        return None;
    }

    // Style-based stroke fallback
    let style_node = child(node, "style");
    let style_stroke: Option<Stroke> = style_node
        .and_then(|s| child(s, "lnRef"))
        .and_then(|lr| {
            let idx: u32 = attr(&lr, "idx").and_then(|v| v.parse().ok()).unwrap_or(1);
            if idx == 0 { None } else {
                parse_color_node(lr, theme).map(|c| Stroke { color: c, width: 9525, dash_style: None })
            }
        });

    let stroke = if child(sp_pr, "ln").is_some() {
        child(sp_pr, "ln").and_then(|n| parse_stroke(n, theme))
    } else {
        style_stroke
    };

    let shadow = child(sp_pr, "effectLst")
        .and_then(|n| parse_shadow(n, theme));

    let cy = if t.cy == 0 { 1 } else { t.cy };

    Some(ShapeElement {
        x: t.x, y: t.y, width: t.cx, height: cy,
        rotation: t.rot, flip_h: t.flip_h, flip_v: t.flip_v,
        geometry: "line".to_owned(),
        fill: None,
        stroke,
        text_body: None,
        default_text_color: None,
        cust_geom: None,
        adj: None,
        adj2: None,
        adj3: None,
        adj4: None,
        shadow,
    })
}

// ===========================
//  Presentation parser
// ===========================

fn parse_presentation(data: &[u8]) -> Result<Presentation, Box<dyn std::error::Error>> {
    let cursor = Cursor::new(data);
    let mut zip = zip::ZipArchive::new(cursor)?;

    // --- presentation.xml ---
    let pres_xml = read_zip_str(&mut zip, "ppt/presentation.xml")?;
    let pres_doc = roxmltree::Document::parse(&pres_xml)?;
    let pres_root = pres_doc.root_element();

    let sld_sz = child(pres_root, "sldSz");
    let slide_width  = sld_sz.and_then(|n| attr_i64(&n, "cx")).unwrap_or(9_144_000);
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
    let pres_rels_xml = read_zip_str(&mut zip, "ppt/_rels/presentation.xml.rels")?;
    let pres_rels = parse_rels(&pres_rels_xml);

    // --- Theme colors ---
    let theme_xml = find_rel_target_by_type(&pres_rels_xml, "/theme")
        .map(|t| resolve_path("ppt", &t))
        .and_then(|path| read_zip_str(&mut zip, &path).ok())
        .unwrap_or_default();
    let theme = parse_theme_colors(&theme_xml);

    // --- First slide master: background + font size defaults ---
    let master_xml_opt: Option<String> = find_rel_target_by_type(&pres_rels_xml, "/slideMaster")
        .map(|t| resolve_path("ppt", &t))
        .and_then(|path| read_zip_str(&mut zip, &path).ok());

    let master_bg: Option<Fill> = master_xml_opt.as_deref().and_then(|master_xml| {
        let doc = roxmltree::Document::parse(master_xml).ok()?;
        child(doc.root_element(), "cSld")
            .and_then(|n| parse_background(n, &theme))
    });

    let master_font_sizes: HashMap<String, f64> = master_xml_opt
        .as_deref()
        .map(|xml| parse_master_font_sizes(xml))
        .unwrap_or_default();

    let master_anchors: HashMap<String, String> = master_xml_opt
        .as_deref()
        .map(|xml| parse_master_anchors(xml))
        .unwrap_or_default();

    let master_transforms: HashMap<String, Transform> = master_xml_opt
        .as_deref()
        .map(|xml| parse_master_transforms(xml))
        .unwrap_or_default();

    let master_alignments: HashMap<String, String> = master_xml_opt
        .as_deref()
        .map(|xml| parse_master_alignments(xml))
        .unwrap_or_default();

    let (master_space_before, master_space_after, master_line_spacing): (HashMap<String, i64>, HashMap<String, i64>, HashMap<String, f64>) = master_xml_opt
        .as_deref()
        .map(|xml| parse_master_txstyle_spacing(xml))
        .unwrap_or_default();

    // Pre-collect slide XMLs, their rels, the layout XML, and layout rels
    struct SlideRaw {
        index: usize,
        slide_xml: String,
        slide_rels: HashMap<String, String>,
        layout_xml: Option<String>,
        layout_rels: HashMap<String, String>,
        layout_dir: String,
    }

    let mut raw_slides: Vec<SlideRaw> = Vec::new();

    for (idx, r_id) in slide_rids.iter().enumerate() {
        let rel_target = match pres_rels.get(r_id) {
            Some(t) => t.clone(),
            None => continue,
        };
        let slide_path = format!("ppt/{rel_target}");
        let slide_file = rel_target.split('/').last().unwrap_or("slide.xml").to_owned();
        let rels_path = format!("ppt/slides/_rels/{slide_file}.rels");

        let slide_xml = read_zip_str(&mut zip, &slide_path)?;
        let slide_rels_xml = read_zip_str(&mut zip, &rels_path).unwrap_or_default();
        let slide_rels = parse_rels(&slide_rels_xml);

        // Layout XML
        let layout_path = find_rel_target_by_type(&slide_rels_xml, "/slideLayout")
            .map(|target| resolve_path("ppt/slides", &target));

        let layout_xml = layout_path.as_deref()
            .and_then(|path| read_zip_str(&mut zip, path).ok());

        // Layout rels (for resolving images inside the layout)
        let layout_rels = layout_path.as_deref()
            .and_then(|path| {
                let file = path.split('/').last().unwrap_or("layout.xml");
                let rels_p = format!("ppt/slideLayouts/_rels/{file}.rels");
                read_zip_str(&mut zip, &rels_p).ok()
            })
            .map(|xml| parse_rels(&xml))
            .unwrap_or_default();

        let layout_dir = layout_path
            .as_deref()
            .and_then(|p| p.rsplit_once('/').map(|(dir, _)| dir.to_owned()))
            .unwrap_or_else(|| "ppt/slideLayouts".to_owned());

        raw_slides.push(SlideRaw {
            index: idx, slide_xml, slide_rels,
            layout_xml, layout_rels, layout_dir,
        });
    }

    let mut slides = Vec::new();
    for raw in &raw_slides {
        let slide = parse_slide(
            &raw.slide_xml,
            raw.layout_xml.as_deref(),
            &raw.layout_rels,
            &raw.layout_dir,
            master_bg.clone(),
            &master_font_sizes,
            &master_anchors,
            &master_transforms,
            &master_alignments,
            &master_space_before,
            &master_space_after,
            &master_line_spacing,
            raw.index,
            &raw.slide_rels,
            &mut zip,
            &theme,
        )?;
        slides.push(slide);
    }

    let default_text_color = theme.get("dk1").cloned();
    Ok(Presentation { slide_width, slide_height, slides, default_text_color })
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn test_parse_chartex() {
        let xml = std::fs::read_to_string("../public/sample-2.pptx").ok()
            .and_then(|_| None::<String>)
            .unwrap_or_else(|| {
                // read from zip directly
                let data = std::fs::read("../public/sample-2.pptx").unwrap();
                let cursor = std::io::Cursor::new(data.as_slice());
                let mut zip = zip::ZipArchive::new(cursor).unwrap();
                let mut s = String::new();
                zip.by_name("ppt/charts/chartEx1.xml").unwrap().read_to_string(&mut s).unwrap();
                s
            });
        let theme = HashMap::new();
        let result = parse_chartex(&xml, &theme);
        println!("parse_chartex result: {:?}", result.is_some());
        if let Some(ref c) = result {
            println!("  chart_type: {}", c.chart_type);
            println!("  categories: {:?}", c.categories);
            println!("  series len: {}", c.series.len());
            if !c.series.is_empty() {
                println!("  values: {:?}", c.series[0].values);
            }
            println!("  subtotal_indices: {:?}", c.subtotal_indices);
        }
        assert!(result.is_some(), "parse_chartex should succeed");
    }

    #[test]
    fn test_slide8_chart_rid() {
        let data = std::fs::read("../public/sample-2.pptx").unwrap();
        let cursor = std::io::Cursor::new(data.as_slice());
        let mut zip = zip::ZipArchive::new(cursor).unwrap();
        let mut slide_xml = String::new();
        zip.by_name("ppt/slides/slide8.xml").unwrap().read_to_string(&mut slide_xml).unwrap();
        
        let doc = roxmltree::Document::parse(&slide_xml).unwrap();
        let root = doc.root_element();
        
        for gf in root.descendants().filter(|n| n.is_element() && n.tag_name().name() == "graphicFrame") {
            println!("Found graphicFrame");
            if let Some(gd) = gf.descendants().find(|n| n.is_element() && n.tag_name().name() == "graphicData") {
                let uri = attr(&gd, "uri").unwrap_or_default();
                println!("  graphicData uri: {}", uri);
                if let Some(chart_node) = gd.descendants().find(|n| n.is_element() && n.tag_name().name() == "chart") {
                    println!("  chart node found, tag: {:?}", chart_node.tag_name());
                    for a in chart_node.attributes() {
                        println!("  attr: name={} ns={:?} val={}", a.name(), a.namespace(), a.value());
                    }
                    let rid = attr_r(&chart_node, "id");
                    println!("  attr_r id: {:?}", rid);
                }
            }
        }
    }

    #[test]
    fn test_slide8_full_parse() {
        let data = std::fs::read("../public/sample-2.pptx").unwrap();
        let pres = parse_presentation(&data).unwrap();
        let slide = &pres.slides[7]; // 0-indexed, slide 8
        println!("Slide 8 elements: {}", slide.elements.len());
        for (i, el) in slide.elements.iter().enumerate() {
            match el {
                SlideElement::Chart(c) => println!("  [{}] CHART type={} cats={}", i, c.chart_type, c.categories.len()),
                SlideElement::Shape(s) => println!("  [{}] shape x={}", i, s.x),
                SlideElement::Table(_) => println!("  [{}] table", i),
                SlideElement::Picture(_) => println!("  [{}] picture", i),
            }
        }
    }

    #[test]
    fn test_slide8_chartex_pipeline() {
        let data = std::fs::read("../public/sample-2.pptx").unwrap();
        let cursor = std::io::Cursor::new(data.as_slice());
        let mut zip = zip::ZipArchive::new(cursor).unwrap();
        
        let mut rels_xml = String::new();
        zip.by_name("ppt/slides/_rels/slide8.xml.rels").unwrap().read_to_string(&mut rels_xml).unwrap();
        let rels = parse_rels(&rels_xml);
        println!("rels: {:?}", rels);
        
        let chart_path = resolve_path("ppt/slides", "../charts/chartEx1.xml");
        println!("chart_path: {}", chart_path);
        
        let result = read_zip_str(&mut zip, &chart_path);
        println!("read_zip_str ok: {}", result.is_ok());
        
        if let Ok(chart_xml) = result {
            let theme = HashMap::new();
            let r = parse_chartex(&chart_xml, &theme);
            println!("parse_chartex: {:?}", r.is_some());
        }
    }


}
