use wasm_bindgen::prelude::*;
use serde::Serialize;
use std::collections::HashMap;
use std::io::{Cursor, Read};

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Workbook {
    pub sheets: Vec<SheetMeta>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SheetMeta {
    pub name: String,
    pub sheet_id: u32,
    pub r_id: String,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct MergeCell {
    pub top: u32,
    pub left: u32,
    pub bottom: u32,
    pub right: u32,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Worksheet {
    pub name: String,
    pub rows: Vec<Row>,
    pub col_widths: HashMap<u32, f64>,
    pub row_heights: HashMap<u32, f64>,
    pub default_col_width: f64,
    pub default_row_height: f64,
    pub merge_cells: Vec<MergeCell>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Row {
    pub index: u32,
    pub height: Option<f64>,
    pub cells: Vec<Cell>,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Cell {
    pub col: u32,
    pub row: u32,
    pub col_ref: String,
    pub value: CellValue,
    pub style_index: u32,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase", tag = "type")]
pub enum CellValue {
    #[default]
    Empty,
    Text { text: String },
    Number { number: f64 },
    Bool { bool: bool },
    Error { error: String },
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Styles {
    pub fonts: Vec<Font>,
    pub fills: Vec<Fill>,
    pub borders: Vec<Border>,
    pub cell_xfs: Vec<CellXf>,
    pub num_fmts: Vec<NumFmt>,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Font {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub size: f64,
    pub color: Option<String>,
    pub name: Option<String>,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Fill {
    pub pattern_type: String,
    pub fg_color: Option<String>,
    pub bg_color: Option<String>,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Border {
    pub left: Option<BorderEdge>,
    pub right: Option<BorderEdge>,
    pub top: Option<BorderEdge>,
    pub bottom: Option<BorderEdge>,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct BorderEdge {
    pub style: String,
    pub color: Option<String>,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct CellXf {
    pub font_id: u32,
    pub fill_id: u32,
    pub border_id: u32,
    pub num_fmt_id: u32,
    pub align_h: Option<String>,
    pub align_v: Option<String>,
    pub wrap_text: bool,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct NumFmt {
    pub num_fmt_id: u32,
    pub format_code: String,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ParsedWorkbook {
    pub workbook: Workbook,
    pub styles: Styles,
    pub shared_strings: Vec<String>,
}

// Excel built-in indexed color palette (indices 0-63)
// Standard Excel 2003 color palette
const INDEXED_COLORS: &[&str] = &[
    "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF", // 0-7
    "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF", // 8-15
    "#800000", "#008000", "#000080", "#808000", "#800080", "#008080", "#C0C0C0", "#808080", // 16-23
    "#9999FF", "#993366", "#FFFFCC", "#CCFFFF", "#660066", "#FF8080", "#0066CC", "#CCCCFF", // 24-31
    "#000080", "#FF00FF", "#FFFF00", "#00FFFF", "#800080", "#800000", "#008080", "#0000FF", // 32-39
    "#00CCFF", "#CCFFFF", "#CCFFCC", "#FFFF99", "#99CCFF", "#FF99CC", "#CC99FF", "#FFCC99", // 40-47
    "#3366FF", "#33CCCC", "#99CC00", "#FFCC00", "#FF9900", "#FF6600", "#666699", "#969696", // 48-55
    "#003366", "#339966", "#003300", "#333300", "#993300", "#993366", "#333399", "#333333", // 56-63
];

#[wasm_bindgen]
pub fn parse_xlsx(data: &[u8]) -> Result<String, JsValue> {
    console_error_panic_hook::set_once();
    parse_xlsx_inner(data)
        .map(|wb| serde_json::to_string(&wb).unwrap())
        .map_err(|e| JsValue::from_str(&e))
}

#[wasm_bindgen]
pub fn parse_sheet(data: &[u8], sheet_index: u32, name: &str) -> Result<String, JsValue> {
    console_error_panic_hook::set_once();
    let cursor = Cursor::new(data);
    let mut archive = zip::ZipArchive::new(cursor).map_err(|e| e.to_string())?;

    let workbook_xml = read_zip_entry(&mut archive, "xl/workbook.xml")?;
    let wb_doc = roxmltree::Document::parse(&workbook_xml).map_err(|e| e.to_string())?;
    let sheets = parse_workbook_sheets(&wb_doc);

    let sheet_meta = sheets
        .get(sheet_index as usize)
        .ok_or_else(|| format!("sheet index {} out of range", sheet_index))?;

    // Resolve rId → target path from workbook.xml.rels
    let rels_xml = read_zip_entry(&mut archive, "xl/_rels/workbook.xml.rels")?;
    let rels_doc = roxmltree::Document::parse(&rels_xml).map_err(|e| e.to_string())?;
    let sheet_path = resolve_sheet_path(&rels_doc, &sheet_meta.r_id)
        .ok_or_else(|| format!("rId {} not found in rels", sheet_meta.r_id))?;

    let shared_strings = read_shared_strings(&mut archive);
    let sheet_xml = read_zip_entry(&mut archive, &format!("xl/{}", sheet_path))?;
    let ws = parse_worksheet(&sheet_xml, &shared_strings, name)
        .map_err(|e| e.to_string())?;

    serde_json::to_string(&ws).map_err(|e| JsValue::from_str(&e.to_string()))
}

fn parse_xlsx_inner(data: &[u8]) -> Result<ParsedWorkbook, String> {
    let cursor = Cursor::new(data);
    let mut archive = zip::ZipArchive::new(cursor).map_err(|e| e.to_string())?;

    let workbook_xml = read_zip_entry(&mut archive, "xl/workbook.xml")?;
    let wb_doc = roxmltree::Document::parse(&workbook_xml).map_err(|e| e.to_string())?;
    let sheets = parse_workbook_sheets(&wb_doc);

    let shared_strings = read_shared_strings(&mut archive);
    let theme_colors = parse_theme_colors(&mut archive);
    let styles = parse_styles(&mut archive, &theme_colors)?;

    Ok(ParsedWorkbook {
        workbook: Workbook { sheets },
        styles,
        shared_strings,
    })
}

fn read_zip_entry(archive: &mut zip::ZipArchive<Cursor<&[u8]>>, name: &str) -> Result<String, String> {
    let mut file = archive
        .by_name(name)
        .map_err(|e| format!("entry '{}' not found: {}", name, e))?;
    let mut buf = String::new();
    file.read_to_string(&mut buf).map_err(|e| e.to_string())?;
    Ok(buf)
}

fn parse_theme_colors(archive: &mut zip::ZipArchive<Cursor<&[u8]>>) -> Vec<String> {
    let Ok(xml) = read_zip_entry(archive, "xl/theme/theme1.xml") else {
        return Vec::new();
    };
    let Ok(doc) = roxmltree::Document::parse(&xml) else {
        return Vec::new();
    };
    let a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main";

    // Find clrScheme node and collect child color elements in order
    // OOXML order: dk1, lt1, dk2, lt2, accent1, accent2, accent3, accent4, accent5, accent6, hlink, folHlink
    let mut colors: Vec<String> = Vec::new();
    for node in doc.descendants() {
        if node.tag_name().name() == "clrScheme" && node.tag_name().namespace() == Some(a_ns) {
            for child in node.children() {
                if !child.is_element() { continue; }
                // Each child is a color slot; its first child element holds the actual color
                for color_node in child.children() {
                    if !color_node.is_element() { continue; }
                    let hex = match color_node.tag_name().name() {
                        "srgbClr" => {
                            color_node.attribute("val").map(|v| format!("#{}", v.to_uppercase()))
                        }
                        "sysClr" => {
                            color_node.attribute("lastClr").map(|v| format!("#{}", v.to_uppercase()))
                        }
                        _ => None,
                    };
                    if let Some(h) = hex {
                        colors.push(h);
                        break;
                    }
                }
            }
            break;
        }
    }
    colors
}

/// Convert hex color + tint to resulting hex color using HLS model.
/// tint > 0: lighten; tint < 0: darken.
fn apply_tint(hex: &str, tint: f64) -> String {
    let hex = hex.trim_start_matches('#');
    if hex.len() < 6 { return format!("#{}", hex); }
    let r = u8::from_str_radix(&hex[0..2], 16).unwrap_or(0) as f64 / 255.0;
    let g = u8::from_str_radix(&hex[2..4], 16).unwrap_or(0) as f64 / 255.0;
    let b = u8::from_str_radix(&hex[4..6], 16).unwrap_or(0) as f64 / 255.0;

    // RGB → HLS
    let max = r.max(g).max(b);
    let min = r.min(g).min(b);
    let l = (max + min) / 2.0;
    let s = if max == min {
        0.0
    } else if l < 0.5 {
        (max - min) / (max + min)
    } else {
        (max - min) / (2.0 - max - min)
    };
    let h = if max == min {
        0.0
    } else if max == r {
        (g - b) / (max - min) / 6.0
    } else if max == g {
        ((b - r) / (max - min) + 2.0) / 6.0
    } else {
        ((r - g) / (max - min) + 4.0) / 6.0
    };
    let h = if h < 0.0 { h + 1.0 } else { h };

    // Apply tint to luminance
    let new_l = if tint > 0.0 {
        l * (1.0 - tint) + tint
    } else {
        l * (1.0 + tint)
    };

    // HLS → RGB
    let (nr, ng, nb) = hls_to_rgb(h, new_l, s);
    format!("#{:02X}{:02X}{:02X}", (nr * 255.0).round() as u8, (ng * 255.0).round() as u8, (nb * 255.0).round() as u8)
}

fn hls_to_rgb(h: f64, l: f64, s: f64) -> (f64, f64, f64) {
    if s == 0.0 {
        return (l, l, l);
    }
    let q = if l < 0.5 { l * (1.0 + s) } else { l + s - l * s };
    let p = 2.0 * l - q;
    let r = hue_to_rgb(p, q, h + 1.0 / 3.0);
    let g = hue_to_rgb(p, q, h);
    let b = hue_to_rgb(p, q, h - 1.0 / 3.0);
    (r, g, b)
}

fn hue_to_rgb(p: f64, q: f64, mut t: f64) -> f64 {
    if t < 0.0 { t += 1.0; }
    if t > 1.0 { t -= 1.0; }
    if t < 1.0 / 6.0 { return p + (q - p) * 6.0 * t; }
    if t < 1.0 / 2.0 { return q; }
    if t < 2.0 / 3.0 { return p + (q - p) * (2.0 / 3.0 - t) * 6.0; }
    p
}

fn parse_color(node: &roxmltree::Node, theme_colors: &[String]) -> Option<String> {
    // rgb attribute (ARGB: 8 chars, drop alpha; or 6-char RGB)
    if let Some(rgb) = node.attribute("rgb") {
        if rgb.len() == 8 {
            return Some(format!("#{}", &rgb[2..].to_uppercase()));
        }
        return Some(format!("#{}", rgb.to_uppercase()));
    }

    // theme attribute → resolve from theme color array + optional tint
    if let Some(theme_str) = node.attribute("theme") {
        if let Ok(idx) = theme_str.parse::<usize>() {
            if let Some(base) = theme_colors.get(idx) {
                let tint = node.attribute("tint").and_then(|s| s.parse::<f64>().ok()).unwrap_or(0.0);
                if tint == 0.0 {
                    return Some(base.clone());
                }
                return Some(apply_tint(base, tint));
            }
        }
    }

    // indexed attribute → Excel built-in palette
    if let Some(indexed_str) = node.attribute("indexed") {
        if let Ok(idx) = indexed_str.parse::<usize>() {
            // indices 64 (foreground) and 65 (background) are special: use black/white
            let color = match idx {
                64 => "#000000",
                65 => "#FFFFFF",
                _ => INDEXED_COLORS.get(idx).copied().unwrap_or("#000000"),
            };
            return Some(color.to_string());
        }
    }

    None
}

fn parse_workbook_sheets(doc: &roxmltree::Document) -> Vec<SheetMeta> {
    let mut sheets = Vec::new();
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    for node in doc.descendants() {
        if node.tag_name().name() == "sheet" && node.tag_name().namespace() == Some(ns) {
            let name = node.attribute("name").unwrap_or("Sheet").to_string();
            let sheet_id = node
                .attribute("sheetId")
                .and_then(|v| v.parse().ok())
                .unwrap_or(1);
            let r_id = node
                .attribute((r_ns, "id"))
                .unwrap_or("")
                .to_string();
            sheets.push(SheetMeta { name, sheet_id, r_id });
        }
    }
    sheets
}

fn resolve_sheet_path(doc: &roxmltree::Document, r_id: &str) -> Option<String> {
    let ns = "http://schemas.openxmlformats.org/package/2006/relationships";
    for node in doc.descendants() {
        if node.tag_name().name() == "Relationship" && node.tag_name().namespace() == Some(ns) {
            if node.attribute("Id") == Some(r_id) {
                return node.attribute("Target").map(|s| s.to_string());
            }
        }
    }
    None
}

fn read_shared_strings(archive: &mut zip::ZipArchive<Cursor<&[u8]>>) -> Vec<String> {
    let Ok(xml) = read_zip_entry(archive, "xl/sharedStrings.xml") else {
        return Vec::new();
    };
    let Ok(doc) = roxmltree::Document::parse(&xml) else {
        return Vec::new();
    };
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let mut strings = Vec::new();
    for si in doc.descendants() {
        if si.tag_name().name() == "si" && si.tag_name().namespace() == Some(ns) {
            let mut text = String::new();
            for t in si.descendants() {
                if t.tag_name().name() == "t" && t.tag_name().namespace() == Some(ns) {
                    if let Some(s) = t.text() {
                        text.push_str(s);
                    }
                }
            }
            strings.push(text);
        }
    }
    strings
}

fn parse_styles(archive: &mut zip::ZipArchive<Cursor<&[u8]>>, theme_colors: &[String]) -> Result<Styles, String> {
    let xml = read_zip_entry(archive, "xl/styles.xml")?;
    let doc = roxmltree::Document::parse(&xml).map_err(|e| e.to_string())?;
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    let num_fmts = parse_num_fmts(&doc, ns);
    let fonts = parse_fonts(&doc, ns, theme_colors);
    let fills = parse_fills(&doc, ns, theme_colors);
    let borders = parse_borders(&doc, ns, theme_colors);
    let cell_xfs = parse_cell_xfs(&doc, ns);

    Ok(Styles { fonts, fills, borders, cell_xfs, num_fmts })
}

fn parse_num_fmts(doc: &roxmltree::Document, ns: &str) -> Vec<NumFmt> {
    let mut fmts = Vec::new();
    for node in doc.descendants() {
        if node.tag_name().name() == "numFmts" && node.tag_name().namespace() == Some(ns) {
            for child in node.children() {
                if child.tag_name().name() != "numFmt" { continue; }
                let num_fmt_id = child.attribute("numFmtId").and_then(|v| v.parse().ok()).unwrap_or(0);
                let format_code = child.attribute("formatCode").unwrap_or("").to_string();
                fmts.push(NumFmt { num_fmt_id, format_code });
            }
            break;
        }
    }
    fmts
}

fn parse_fonts(doc: &roxmltree::Document, ns: &str, theme_colors: &[String]) -> Vec<Font> {
    let mut fonts = Vec::new();
    for fonts_node in doc.descendants() {
        if fonts_node.tag_name().name() == "fonts" && fonts_node.tag_name().namespace() == Some(ns) {
            for font_node in fonts_node.children() {
                if font_node.tag_name().name() != "font" { continue; }
                let mut f = Font { size: 11.0, ..Default::default() };
                for child in font_node.children() {
                    match child.tag_name().name() {
                        "b" => f.bold = true,
                        "i" => f.italic = true,
                        "u" => f.underline = true,
                        "sz" => {
                            if let Some(v) = child.attribute("val").and_then(|s| s.parse().ok()) {
                                f.size = v;
                            }
                        }
                        "name" => {
                            f.name = child.attribute("val").map(|s| s.to_string());
                        }
                        "color" => {
                            f.color = parse_color(&child, theme_colors);
                        }
                        _ => {}
                    }
                }
                fonts.push(f);
            }
            break;
        }
    }
    fonts
}

fn parse_fills(doc: &roxmltree::Document, ns: &str, theme_colors: &[String]) -> Vec<Fill> {
    let mut fills = Vec::new();
    for fills_node in doc.descendants() {
        if fills_node.tag_name().name() == "fills" && fills_node.tag_name().namespace() == Some(ns) {
            for fill_node in fills_node.children() {
                if fill_node.tag_name().name() != "fill" { continue; }
                let mut f = Fill::default();
                for pf in fill_node.children() {
                    if pf.tag_name().name() == "patternFill" {
                        f.pattern_type = pf.attribute("patternType").unwrap_or("none").to_string();
                        for color_node in pf.children() {
                            match color_node.tag_name().name() {
                                "fgColor" => f.fg_color = parse_color(&color_node, theme_colors),
                                "bgColor" => f.bg_color = parse_color(&color_node, theme_colors),
                                _ => {}
                            }
                        }
                    }
                }
                fills.push(f);
            }
            break;
        }
    }
    fills
}

fn parse_borders(doc: &roxmltree::Document, ns: &str, theme_colors: &[String]) -> Vec<Border> {
    let mut borders = Vec::new();
    for borders_node in doc.descendants() {
        if borders_node.tag_name().name() == "borders" && borders_node.tag_name().namespace() == Some(ns) {
            for border_node in borders_node.children() {
                if border_node.tag_name().name() != "border" { continue; }
                let mut b = Border::default();
                for edge_node in border_node.children() {
                    let style = edge_node.attribute("style").unwrap_or("").to_string();
                    if style.is_empty() { continue; }
                    let color = edge_node.children().find(|c| c.is_element()).and_then(|c| parse_color(&c, theme_colors));
                    let edge = Some(BorderEdge { style, color });
                    match edge_node.tag_name().name() {
                        "left" => b.left = edge,
                        "right" => b.right = edge,
                        "top" => b.top = edge,
                        "bottom" => b.bottom = edge,
                        _ => {}
                    }
                }
                borders.push(b);
            }
            break;
        }
    }
    borders
}

fn parse_cell_xfs(doc: &roxmltree::Document, ns: &str) -> Vec<CellXf> {
    let mut xfs = Vec::new();
    for xfs_node in doc.descendants() {
        if xfs_node.tag_name().name() == "cellXfs" && xfs_node.tag_name().namespace() == Some(ns) {
            for xf_node in xfs_node.children() {
                if xf_node.tag_name().name() != "xf" { continue; }
                let font_id = xf_node.attribute("fontId").and_then(|v| v.parse().ok()).unwrap_or(0);
                let fill_id = xf_node.attribute("fillId").and_then(|v| v.parse().ok()).unwrap_or(0);
                let border_id = xf_node.attribute("borderId").and_then(|v| v.parse().ok()).unwrap_or(0);
                let num_fmt_id = xf_node.attribute("numFmtId").and_then(|v| v.parse().ok()).unwrap_or(0);
                let mut align_h = None;
                let mut align_v = None;
                let mut wrap_text = false;
                for child in xf_node.children() {
                    if child.tag_name().name() == "alignment" {
                        align_h = child.attribute("horizontal").map(|s| s.to_string());
                        align_v = child.attribute("vertical").map(|s| s.to_string());
                        wrap_text = child.attribute("wrapText").map(|v| v == "1" || v == "true").unwrap_or(false);
                    }
                }
                xfs.push(CellXf { font_id, fill_id, border_id, num_fmt_id, align_h, align_v, wrap_text });
            }
            break;
        }
    }
    xfs
}

fn parse_worksheet(xml: &str, shared_strings: &[String], name: &str) -> Result<Worksheet, String> {
    let doc = roxmltree::Document::parse(xml).map_err(|e| e.to_string())?;
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    let mut rows = Vec::new();
    let mut col_widths: HashMap<u32, f64> = HashMap::new();
    let mut row_heights: HashMap<u32, f64> = HashMap::new();
    let mut merge_cells: Vec<MergeCell> = Vec::new();
    let mut default_col_width = 8.43;
    let mut default_row_height = 15.0;

    for node in doc.descendants() {
        match node.tag_name().name() {
            "sheetFormatPr" if node.tag_name().namespace() == Some(ns) => {
                if let Some(v) = node.attribute("defaultColWidth").and_then(|s| s.parse().ok()) {
                    default_col_width = v;
                }
                if let Some(v) = node.attribute("defaultRowHeight").and_then(|s| s.parse().ok()) {
                    default_row_height = v;
                }
            }
            "col" if node.tag_name().namespace() == Some(ns) => {
                let custom = node.attribute("customWidth").map(|v| v == "1").unwrap_or(false);
                if !custom { continue; }
                let min: u32 = node.attribute("min").and_then(|s| s.parse().ok()).unwrap_or(1);
                let max: u32 = node.attribute("max").and_then(|s| s.parse().ok()).unwrap_or(1);
                // Cap range to avoid storing 16K entries for max=16384 ranges
                let max = max.min(min + 255);
                let width: f64 = node.attribute("width").and_then(|s| s.parse().ok()).unwrap_or(default_col_width);
                for c in min..=max {
                    col_widths.insert(c, width);
                }
            }
            "mergeCell" if node.tag_name().namespace() == Some(ns) => {
                if let Some(r) = node.attribute("ref") {
                    let parts: Vec<&str> = r.split(':').collect();
                    if parts.len() == 2 {
                        let (left, top) = parse_cell_ref(parts[0]);
                        let (right, bottom) = parse_cell_ref(parts[1]);
                        merge_cells.push(MergeCell { top, left, bottom, right });
                    }
                }
            }
            "row" if node.tag_name().namespace() == Some(ns) => {
                let row_idx: u32 = node.attribute("r").and_then(|s| s.parse().ok()).unwrap_or(0);
                let height: Option<f64> = node.attribute("ht").and_then(|s| s.parse().ok());
                if let Some(h) = height {
                    row_heights.insert(row_idx, h);
                }
                let cells = parse_row_cells(&node, shared_strings, ns);
                rows.push(Row { index: row_idx, height, cells });
            }
            _ => {}
        }
    }

    Ok(Worksheet {
        name: name.to_string(),
        rows,
        col_widths,
        row_heights,
        default_col_width,
        default_row_height,
        merge_cells,
    })
}

fn parse_row_cells(row_node: &roxmltree::Node, shared_strings: &[String], ns: &str) -> Vec<Cell> {
    let mut cells = Vec::new();
    for c_node in row_node.children() {
        if c_node.tag_name().name() != "c" || c_node.tag_name().namespace() != Some(ns) {
            continue;
        }
        let cell_ref = c_node.attribute("r").unwrap_or("A1").to_string();
        let (col, row) = parse_cell_ref(&cell_ref);
        let cell_type = c_node.attribute("t").unwrap_or("");
        let style_index: u32 = c_node.attribute("s").and_then(|s| s.parse().ok()).unwrap_or(0);

        let v_text = c_node
            .children()
            .find(|n| n.tag_name().name() == "v")
            .and_then(|n| n.text())
            .unwrap_or("")
            .to_string();

        let value = if v_text.is_empty() {
            CellValue::Empty
        } else {
            match cell_type {
                "s" => {
                    let idx: usize = v_text.parse().unwrap_or(0);
                    let text = shared_strings.get(idx).cloned().unwrap_or_default();
                    CellValue::Text { text }
                }
                "str" | "inlineStr" => CellValue::Text { text: v_text },
                "b" => CellValue::Bool { bool: v_text == "1" || v_text == "true" },
                "e" => CellValue::Error { error: v_text },
                _ => {
                    if let Ok(n) = v_text.parse::<f64>() {
                        CellValue::Number { number: n }
                    } else {
                        CellValue::Text { text: v_text }
                    }
                }
            }
        };

        cells.push(Cell { col, row, col_ref: cell_ref, value, style_index });
    }
    cells
}

fn parse_cell_ref(r: &str) -> (u32, u32) {
    let col_str: String = r.chars().take_while(|c| c.is_ascii_alphabetic()).collect();
    let row_str: String = r.chars().skip_while(|c| c.is_ascii_alphabetic()).collect();
    let col = col_str.chars().fold(0u32, |acc, c| acc * 26 + (c as u32 - 'A' as u32 + 1));
    let row = row_str.parse().unwrap_or(1);
    (col, row)
}
