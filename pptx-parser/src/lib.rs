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
}

#[derive(Serialize, Deserialize, Debug)]
#[serde(rename_all = "camelCase")]
struct Slide {
    index: usize,
    background: Option<Fill>,
    elements: Vec<SlideElement>,
}

#[derive(Serialize, Deserialize, Debug)]
#[serde(tag = "type", rename_all = "camelCase")]
enum SlideElement {
    Shape(ShapeElement),
    Picture(PictureElement),
    Table(TableElement),
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
#[serde(tag = "fillType", rename_all = "camelCase")]
enum Fill {
    Solid { color: String },
    None,
}

#[derive(Serialize, Deserialize, Debug)]
#[serde(rename_all = "camelCase")]
struct Stroke {
    color: String,
    width: i64,
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
    /// Text insets in EMU. Defaults: lIns=rIns=91440, tIns=bIns=45720
    l_ins: i64,
    r_ins: i64,
    t_ins: i64,
    b_ins: i64,
    /// Whether text wraps inside the bounding box ("square") or not ("none")
    wrap: String,
    /// Text direction from bodyPr vert attribute: "horz" | "vert" | "vert270" | "eaVert" etc.
    vert: String,
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
    Char {
        #[serde(rename = "char")]
        ch: String,
        color: Option<String>,
        /// Size as % of text size (100.0 = same size)
        size_pct: Option<f64>,
        font_family: Option<String>,
    },
    /// Auto-numbered bullet (buAutoNum)
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
    bold: bool,
    italic: bool,
    underline: bool,
    font_size: Option<f64>,
    color: Option<String>,
    font_family: Option<String>,
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
    let clr_scheme = match doc
        .root_element()
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
    map
}

/// Resolve a color node (solidFill child / run rPr child) to a hex string.
/// Handles srgbClr, sysClr, prstClr, and schemeClr (with transform support).
fn parse_color_node(
    node: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<String> {
    for c in node.children().filter(|n| n.is_element()) {
        match c.tag_name().name() {
            "srgbClr" => return attr(&c, "val"),
            "sysClr"  => return attr(&c, "lastClr"),
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

/// Apply OOXML color transforms (lumMod, lumOff, shade, tint) to a base hex color.
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
            _ => {}
        }
    }

    let r = (rf.clamp(0.0, 1.0) * 255.0).round() as u8;
    let g = (gf.clamp(0.0, 1.0) * 255.0).round() as u8;
    let b = (bf.clamp(0.0, 1.0) * 255.0).round() as u8;
    format!("{:02X}{:02X}{:02X}", r, g, b)
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
                // Use the first gradient stop's color as a solid-fill approximation.
                let first_gs = c
                    .descendants()
                    .find(|n| n.is_element() && n.tag_name().name() == "gs");
                if let Some(gs) = first_gs {
                    if let Some(color) = parse_color_node(gs, theme) {
                        return Some(Fill::Solid { color });
                    }
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
    // If no explicit solidFill or color is unresolvable, return None so the
    // caller can fall back to the shape style stroke color.
    let color = child(ln_node, "solidFill")
        .and_then(|n| parse_color_node(n, theme))?;
    Some(Stroke { color, width })
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
}

impl GroupTransform {
    fn apply_to_transform(&self, t: Transform) -> Transform {
        let sx = if self.ch_cx != 0 { self.cx as f64 / self.ch_cx as f64 } else { 1.0 };
        let sy = if self.ch_cy != 0 { self.cy as f64 / self.ch_cy as f64 } else { 1.0 };
        // If the group is flipped, mirror child positions in child coordinate space
        // before applying the normal scale+translate.
        let child_x = if self.flip_h {
            self.ch_x + self.ch_cx - t.x - t.cx
        } else {
            t.x
        };
        let child_y = if self.flip_v {
            self.ch_y + self.ch_cy - t.y - t.cy
        } else {
            t.y
        };
        Transform {
            x:  ((child_x - self.ch_x) as f64 * sx + self.x as f64).round() as i64,
            y:  ((child_y - self.ch_y) as f64 * sy + self.y as f64).round() as i64,
            cx: (t.cx as f64 * sx).round() as i64,
            cy: (t.cy as f64 * sy).round() as i64,
            rot: t.rot,
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
            s.flip_h = nt.flip_h; s.flip_v = nt.flip_v;
        }
        SlideElement::Picture(p) => {
            let t = Transform { x: p.x, y: p.y, cx: p.width, cy: p.height, rot: p.rotation, flip_h: p.flip_h, flip_v: p.flip_v };
            let nt = gt.apply_to_transform(t);
            p.x = nt.x; p.y = nt.y; p.width = nt.cx; p.height = nt.cy;
            p.flip_h = nt.flip_h; p.flip_v = nt.flip_v;
        }
        SlideElement::Table(tbl) => {
            let t = Transform { x: tbl.x, y: tbl.y, cx: tbl.width, cy: tbl.height, rot: 0.0, flip_h: false, flip_v: false };
            let nt = gt.apply_to_transform(t);
            tbl.x = nt.x; tbl.y = nt.y; tbl.width = nt.cx; tbl.height = nt.cy;
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
}

impl LayoutPlaceholders {
    fn lookup(&self, ph_type: &str, ph_idx: Option<u32>) -> Option<&Transform> {
        ph_idx
            .and_then(|i| self.by_idx.get(&i))
            .or_else(|| self.by_type.get(ph_type))
            .or_else(|| {
                if ph_type == "body" { self.by_type.get("") } else { None }
            })
    }
}

fn parse_layout_placeholders(layout_xml: &str) -> LayoutPlaceholders {
    let mut lph = LayoutPlaceholders::default();
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
        let xfrm = match child(sp_pr, "xfrm") {
            Some(n) => n,
            None => continue,
        };
        let t = parse_xfrm(xfrm);

        if let Some(ph) = ph_node {
            let ph_type = attr(&ph, "type").unwrap_or_default();
            let ph_idx: Option<u32> = attr(&ph, "idx").and_then(|v| v.parse().ok());

            if let Some(idx) = ph_idx {
                lph.by_idx.entry(idx).or_insert_with(|| t.clone());
            }
            lph.by_type.entry(ph_type).or_insert(t);
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
) -> TextBody {
    let body_pr = child(tx_body, "bodyPr");
    let vertical_anchor = body_pr
        .and_then(|n| attr(&n, "anchor"))
        .unwrap_or_else(|| "t".into());
    // Text insets (EMU). OOXML defaults: lIns=rIns=91440, tIns=bIns=45720
    let l_ins = body_pr.and_then(|n| attr_i64(&n, "lIns")).unwrap_or(91_440);
    let r_ins = body_pr.and_then(|n| attr_i64(&n, "rIns")).unwrap_or(91_440);
    let t_ins = body_pr.and_then(|n| attr_i64(&n, "tIns")).unwrap_or(45_720);
    let b_ins = body_pr.and_then(|n| attr_i64(&n, "bIns")).unwrap_or(45_720);
    let wrap = body_pr.and_then(|n| attr(&n, "wrap")).unwrap_or_else(|| "square".into());
    let vert = body_pr.and_then(|n| attr(&n, "vert")).unwrap_or_else(|| "horz".into());

    // Default font size from lstStyle > lvl1pPr > defRPr sz
    let default_font_size = child(tx_body, "lstStyle")
        .and_then(|ls| child(ls, "lvl1pPr"))
        .and_then(|lp| child(lp, "defRPr"))
        .and_then(|rp| attr_f64(&rp, "sz"))
        .map(|v| v / 100.0);

    let paragraphs = children_vec(tx_body, "p")
        .into_iter()
        .map(|p| parse_paragraph(p, theme))
        .collect();

    TextBody { vertical_anchor, paragraphs, default_font_size, l_ins, r_ins, t_ins, b_ins, wrap, vert }
}

fn parse_paragraph(
    p_node: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Paragraph {
    let p_pr = child(p_node, "pPr");

    let alignment  = p_pr.and_then(|n| attr(&n, "algn")).unwrap_or_else(|| "l".into());
    let lvl: u32   = p_pr.and_then(|n| attr(&n, "lvl")).and_then(|v| v.parse().ok()).unwrap_or(0);
    let mar_l      = p_pr.and_then(|n| attr_i64(&n, "marL")).unwrap_or(0);
    let mar_r      = p_pr.and_then(|n| attr_i64(&n, "marR")).unwrap_or(0);
    let indent     = p_pr.and_then(|n| attr_i64(&n, "indent")).unwrap_or(0);

    let space_before = p_pr.and_then(|n| {
        child(n, "spcBef").and_then(|s| child(s, "spcPts")).and_then(|s| attr_i64(&s, "val"))
    });
    let space_after = p_pr.and_then(|n| {
        child(n, "spcAft").and_then(|s| child(s, "spcPts")).and_then(|s| attr_i64(&s, "val"))
    });

    let space_line = p_pr.and_then(|n| {
        let spc = child(n, "lnSpc")?;
        if let Some(pct) = child(spc, "spcPct") {
            attr_f64(&pct, "val").map(|v| SpaceLine::Pct { val: v })
        } else {
            child(spc, "spcPts")
                .and_then(|pts| attr_f64(&pts, "val"))
                .map(|v| SpaceLine::Pts { val: v / 100.0 }) // hundredths of pt → pt
        }
    });

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
    let def_font_family = def_rpr.and_then(|n| child(n, "latin")).and_then(|n| attr(&n, "typeface"));

    let mut runs = Vec::new();
    for node in p_node.children().filter(|n| n.is_element()) {
        match node.tag_name().name() {
            "r" => {
                if let Some(run) = parse_run(node, def_rpr, theme) {
                    runs.push(TextRun::Text(run));
                }
            }
            "br" => runs.push(TextRun::Break),
            _ => {}
        }
    }

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
        let font_family = child(p_pr, "buFont").and_then(|n| attr(&n, "typeface"));
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

    // Attribute with rPr → defRPr fallback
    let bold = r_pr.and_then(|n| attr(&n, "b"))
        .or_else(|| def_rpr.and_then(|n| attr(&n, "b")))
        .map(|v| v == "1" || v == "true").unwrap_or(false);
    let italic = r_pr.and_then(|n| attr(&n, "i"))
        .or_else(|| def_rpr.and_then(|n| attr(&n, "i")))
        .map(|v| v == "1" || v == "true").unwrap_or(false);
    let underline = r_pr.and_then(|n| attr(&n, "u"))
        .or_else(|| def_rpr.and_then(|n| attr(&n, "u")))
        .map(|v| v != "none").unwrap_or(false);

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
        .or_else(|| def_rpr.and_then(|n| child(n, "latin")).and_then(|n| attr(&n, "typeface")));

    Some(TextRunData { text, bold, italic, underline, font_size, color, font_family })
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

    // cx=0 → skip. cy=0 means "auto-height" in OOXML; use a generous fallback.
    if t.cx == 0 {
        return None;
    }
    let cy = if t.cy == 0 { 2_000_000_i64 } else { t.cy };

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

    // First adjustment value from prstGeom avLst (e.g. trapezoid inset)
    let adj: Option<f64> = prst_geom_node
        .and_then(|n| child(n, "avLst"))
        .and_then(|av| {
            av.children()
                .filter(|n| n.is_element() && n.tag_name().name() == "gd")
                .find(|n| attr(n, "name").as_deref() == Some("adj"))
        })
        .and_then(|gd| {
            // fmla is like "val 25000"; extract the numeric part
            attr(&gd, "fmla")
                .and_then(|f| f.strip_prefix("val ").map(|s| s.to_owned()))
                .and_then(|s| s.parse::<f64>().ok())
        });

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
                parse_color_node(lr, theme).map(|c| Stroke { color: c, width: 9525 })
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
    // only fall back to style when spPr has no ln at all.
    let stroke = if sp_pr.and_then(|p| child(p, "ln")).is_some() {
        sp_pr.and_then(|p| child(p, "ln")).and_then(|n| parse_stroke(n, theme))
    } else {
        style_stroke
    };

    let text_body = child(sp_node, "txBody").map(|n| parse_text_body(n, theme));

    Some(ShapeElement {
        x: t.x, y: t.y, width: t.cx, height: cy,
        rotation: t.rot, flip_h: t.flip_h, flip_v: t.flip_v,
        geometry, fill, stroke, text_body, default_text_color, cust_geom, adj,
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

fn parse_table(
    tbl: roxmltree::Node<'_, '_>,
    t: &Transform,
    theme: &HashMap<String, String>,
) -> Option<TableElement> {
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

    let rows: Vec<TableRow> = tbl
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "tr")
        .map(|tr| parse_table_row(tr, theme))
        .collect();

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
    let text_body = child(tc, "txBody").map(|n| parse_text_body(n, theme));

    let tc_pr = child(tc, "tcPr");
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
    index: usize,
    rels: &HashMap<String, String>,
    zip: &mut PptxZip<'_>,
    theme: &HashMap<String, String>,
) -> Result<Slide, Box<dyn std::error::Error>> {
    let lph = layout_xml
        .map(|x| parse_layout_placeholders(x))
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

    Ok(Slide { index, background, elements })
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
            if let Some(shape) = parse_shape(node, lph, theme, group_fill) {
                out.push(SlideElement::Shape(shape));
            }
        }
        "pic" => {
            if let Some(pic) = parse_picture(node, slide_dir, rels, zip) {
                out.push(SlideElement::Picture(pic));
            }
        }
        "graphicFrame" => {
            let xfrm_node = child(node, "xfrm");
            let t = xfrm_node.map(parse_xfrm).unwrap_or_default();
            let tbl_node = node
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "tbl");
            if let Some(tbl_node) = tbl_node {
                if let Some(table) = parse_table(tbl_node, &t, theme) {
                    out.push(SlideElement::Table(table));
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
        _ => {}
    }
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

    // --- First slide master background (fallback for slides without their own bg) ---
    let master_bg: Option<Fill> = find_rel_target_by_type(&pres_rels_xml, "/slideMaster")
        .map(|t| resolve_path("ppt", &t))
        .and_then(|path| read_zip_str(&mut zip, &path).ok())
        .and_then(|master_xml| {
            let doc = roxmltree::Document::parse(&master_xml).ok()?;
            child(doc.root_element(), "cSld")
                .and_then(|n| parse_background(n, &theme))
        });

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
            raw.index,
            &raw.slide_rels,
            &mut zip,
            &theme,
        )?;
        slides.push(slide);
    }

    Ok(Presentation { slide_width, slide_height, slides })
}
