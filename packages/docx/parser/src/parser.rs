use std::collections::HashMap;
use std::io::Read;
use zip::ZipArchive;
use roxmltree::Document as XmlDoc;
use base64::Engine;
use base64::engine::general_purpose::STANDARD as B64;

use crate::xml_util::*;
use crate::types::*;
use crate::styles::{StyleMap, parse_para_fmt, parse_run_fmt, ParaFmt, RunFmt};
use crate::numbering::NumberingMap;

const DEFAULT_FONT_SIZE: f64 = 10.0; // pt fallback

pub fn parse(data: &[u8]) -> Result<Document, String> {
    let cursor = std::io::Cursor::new(data);
    let mut zip = ZipArchive::new(cursor).map_err(|e| e.to_string())?;

    let style_map = read_zip_entry(&mut zip, "word/styles.xml")
        .map(|s| StyleMap::parse(&s))
        .unwrap_or_else(|_| StyleMap::parse(""));

    let mut num_map = read_zip_entry(&mut zip, "word/numbering.xml")
        .map(|s| NumberingMap::parse(&s))
        .unwrap_or_default();

    let rels = read_zip_entry(&mut zip, "word/_rels/document.xml.rels")
        .unwrap_or_default();

    // Build relationship map: rId → target path
    let rel_map = parse_rels(&rels);

    // Load media images as base64
    let mut media_map: HashMap<String, String> = HashMap::new();
    for (rid, target) in &rel_map {
        if target.contains("media/") || target.contains("image") {
            let path = if target.starts_with('/') {
                target.trim_start_matches('/').to_string()
            } else {
                format!("word/{}", target)
            };
            if let Ok(bytes) = read_zip_bytes(&mut zip, &path) {
                let mime = if path.ends_with(".png") { "image/png" }
                    else if path.ends_with(".jpg") || path.ends_with(".jpeg") { "image/jpeg" }
                    else if path.ends_with(".gif") { "image/gif" }
                    else { "image/png" };
                let b64 = B64.encode(&bytes);
                media_map.insert(rid.clone(), format!("data:{};base64,{}", mime, b64));
            }
        }
    }

    let doc_xml = read_zip_entry(&mut zip, "word/document.xml")?;
    let xml_doc = XmlDoc::parse(&doc_xml).map_err(|e| e.to_string())?;

    let body_node = xml_doc.root_element()
        .descendants()
        .find(|n| n.tag_name().name() == "body")
        .ok_or("No body element")?;

    let sect_pr = body_node.children()
        .filter(|n| n.is_element())
        .last()
        .filter(|n| n.tag_name().name() == "sectPr");

    let section = parse_section(sect_pr);

    let mut body: Vec<BodyElement> = Vec::new();

    for child in body_node.children().filter(|n| n.is_element()) {
        match child.tag_name().name() {
            "p" => {
                let result = parse_paragraph(child, &style_map, &mut num_map, &media_map);
                // Check if this paragraph contains only a page break
                if result.runs.len() == 1 {
                    if let DocRun::Break { break_type: BreakType::Page } = &result.runs[0] {
                        body.push(BodyElement::PageBreak);
                        continue;
                    }
                }
                body.push(BodyElement::Paragraph(result));
            }
            "tbl" => {
                let tbl = parse_table(child, &style_map, &mut num_map, &media_map);
                body.push(BodyElement::Table(tbl));
            }
            "sectPr" => {} // already handled
            _ => {}
        }
    }

    Ok(Document { section, body })
}

fn parse_section(sect_pr: Option<roxmltree::Node>) -> SectionProps {
    let default = SectionProps {
        page_width: 612.0,
        page_height: 792.0,
        margin_top: 72.0,
        margin_right: 72.0,
        margin_bottom: 72.0,
        margin_left: 72.0,
    };

    let Some(sp) = sect_pr else { return default };

    let mut props = default;
    if let Some(pg_sz) = child_w(sp, "pgSz") {
        if let Some(w) = attr_w(pg_sz, "w") { props.page_width = twips_to_pt(&w); }
        if let Some(h) = attr_w(pg_sz, "h") { props.page_height = twips_to_pt(&h); }
    }
    if let Some(pg_mar) = child_w(sp, "pgMar") {
        if let Some(v) = attr_w(pg_mar, "top") { props.margin_top = twips_to_pt(&v); }
        if let Some(v) = attr_w(pg_mar, "right") { props.margin_right = twips_to_pt(&v); }
        if let Some(v) = attr_w(pg_mar, "bottom") { props.margin_bottom = twips_to_pt(&v); }
        if let Some(v) = attr_w(pg_mar, "left") { props.margin_left = twips_to_pt(&v); }
    }
    props
}

fn parse_paragraph(
    node: roxmltree::Node,
    style_map: &StyleMap,
    num_map: &mut NumberingMap,
    media_map: &HashMap<String, String>,
) -> DocParagraph {
    // Get style ID from pPr/pStyle
    let ppr_node = child_w(node, "pPr");
    let style_id = ppr_node
        .and_then(|p| child_w(p, "pStyle"))
        .and_then(|s| attr_w(s, "val"));

    // Resolve base formatting from style
    let (mut base_para, mut base_run) = style_map.resolve_para(style_id.as_deref());

    // Apply direct paragraph formatting overrides
    if let Some(ppr) = ppr_node {
        let direct = parse_para_fmt(ppr);
        apply_direct_para(&mut base_para, &direct);
        // Also merge direct rPr
        if let Some(rpr) = child_w(ppr, "rPr") {
            let direct_run = parse_run_fmt(rpr);
            apply_direct_run(&mut base_run, &direct_run);
        }
    }

    let alignment = base_para.alignment.as_deref().map(normalize_align).unwrap_or("left").to_string();
    let indent_left = base_para.indent_left.unwrap_or(0.0);
    let indent_right = base_para.indent_right.unwrap_or(0.0);
    let indent_first = base_para.indent_first.unwrap_or(0.0);
    let space_before = base_para.space_before.unwrap_or(0.0);
    let space_after = base_para.space_after.unwrap_or(0.0);
    let line_spacing = base_para.line_spacing_val.map(|v| LineSpacing {
        value: v,
        rule: base_para.line_spacing_rule.clone().unwrap_or_else(|| "auto".to_string()),
    });

    // Numbering
    let numbering = if let (Some(num_id), Some(num_level)) = (base_para.num_id, base_para.num_level) {
        if num_id != 0 {
            let counter = num_map.advance(num_id, num_level);
            let text = num_map.resolve_text(num_id, num_level, counter);
            let lvl = num_map.get_level(num_id, num_level);
            let (ind_left, tab) = lvl.map(|l| (l.indent_left, l.tab)).unwrap_or((36.0, 36.0));
            Some(NumberingInfo {
                num_id,
                level: num_level,
                format: lvl.map(|l| l.format.clone()).unwrap_or_else(|| "decimal".to_string()),
                text,
                indent_left: ind_left,
                tab,
            })
        } else { None }
    } else { None };

    // Parse runs
    let mut runs = vec![];
    parse_para_content(node, &base_run, style_map, media_map, &mut runs);

    DocParagraph {
        alignment,
        indent_left,
        indent_right,
        indent_first,
        space_before,
        space_after,
        line_spacing,
        numbering,
        runs,
    }
}

fn parse_para_content(
    node: roxmltree::Node,
    base_run: &RunFmt,
    style_map: &StyleMap,
    media_map: &HashMap<String, String>,
    runs: &mut Vec<DocRun>,
) {
    for child in node.children().filter(|n| n.is_element()) {
        match child.tag_name().name() {
            "r" => {
                parse_run(child, base_run, style_map, media_map, runs);
            }
            "hyperlink" => {
                // Recurse into hyperlink, marking runs as links
                for r in child.children().filter(|n| n.is_element() && n.tag_name().name() == "r") {
                    parse_run_as_link(r, base_run, style_map, media_map, runs);
                }
            }
            "ins" | "del" | "smartTag" => {
                // Recurse into tracked changes / smart tags
                parse_para_content(child, base_run, style_map, media_map, runs);
            }
            _ => {}
        }
    }
}

fn parse_run(
    node: roxmltree::Node,
    base_run: &RunFmt,
    style_map: &StyleMap,
    media_map: &HashMap<String, String>,
    runs: &mut Vec<DocRun>,
) {
    parse_run_inner(node, base_run, style_map, media_map, runs, false);
}

fn parse_run_as_link(
    node: roxmltree::Node,
    base_run: &RunFmt,
    style_map: &StyleMap,
    media_map: &HashMap<String, String>,
    runs: &mut Vec<DocRun>,
) {
    parse_run_inner(node, base_run, style_map, media_map, runs, true);
}

fn parse_run_inner(
    node: roxmltree::Node,
    base_run: &RunFmt,
    style_map: &StyleMap,
    media_map: &HashMap<String, String>,
    runs: &mut Vec<DocRun>,
    is_link: bool,
) {
    // Merge run-level formatting
    let rpr_node = child_w(node, "rPr");
    let mut fmt = base_run.clone();

    // Apply rStyle
    if let Some(rpr) = rpr_node {
        if let Some(rs) = child_w(rpr, "rStyle").and_then(|n| attr_w(n, "val")) {
            let (_, style_run) = style_map.resolve_para(Some(&rs));
            apply_direct_run(&mut fmt, &style_run);
        }
        let direct = parse_run_fmt(rpr);
        apply_direct_run(&mut fmt, &direct);
    }

    let bold = fmt.bold.unwrap_or(false);
    let italic = fmt.italic.unwrap_or(false);
    let underline = fmt.underline.unwrap_or(false) || is_link;
    let strikethrough = fmt.strikethrough.unwrap_or(false);
    let font_size = fmt.font_size.unwrap_or(DEFAULT_FONT_SIZE);
    let color = if is_link && fmt.color.is_none() {
        Some("0563c1".to_string())
    } else {
        fmt.color.clone()
    };
    let font_family = fmt.font_family_ascii.clone().or(fmt.font_family_east_asia.clone());

    for child in node.children().filter(|n| n.is_element()) {
        match child.tag_name().name() {
            "t" => {
                let text = child.text().unwrap_or("").to_string();
                if !text.is_empty() {
                    runs.push(DocRun::Text(TextRun {
                        text,
                        bold,
                        italic,
                        underline,
                        strikethrough,
                        font_size,
                        color: color.clone(),
                        font_family: font_family.clone(),
                        is_link,
                        background: fmt.background.clone(),
                    }));
                }
            }
            "br" => {
                let break_type = attr_w(child, "type").as_deref().map(|v| match v {
                    "page" => BreakType::Page,
                    "column" => BreakType::Column,
                    _ => BreakType::Line,
                }).unwrap_or(BreakType::Line);
                runs.push(DocRun::Break { break_type });
            }
            "drawing" => {
                if let Some(img) = parse_inline_drawing(child, media_map) {
                    runs.push(DocRun::Image(img));
                }
            }
            _ => {}
        }
    }
}

fn parse_inline_drawing(node: roxmltree::Node, media_map: &HashMap<String, String>) -> Option<ImageRun> {
    // Find wp:inline or wp:anchor
    let inline = node.descendants().find(|n| n.tag_name().name() == "inline")
        .or_else(|| node.descendants().find(|n| n.tag_name().name() == "anchor"))?;

    // Get extent
    let extent = inline.children().find(|n| n.tag_name().name() == "extent")?;
    let cx: f64 = extent.attribute("cx")?.parse().ok()?;
    let cy: f64 = extent.attribute("cy")?.parse().ok()?;
    let width_pt = cx / 12700.0;
    let height_pt = cy / 12700.0;

    // Find blip rId
    let blip = node.descendants().find(|n| n.tag_name().name() == "blip")?;
    let r_id = blip.attribute((R_NS, "embed")).or_else(|| blip.attribute("r:embed"))?;

    let data_url = media_map.get(r_id)?.clone();

    Some(ImageRun { data_url, width_pt, height_pt })
}

// ===== Table parsing =====

fn parse_table(
    node: roxmltree::Node,
    style_map: &StyleMap,
    num_map: &mut NumberingMap,
    media_map: &HashMap<String, String>,
) -> DocTable {
    let tbl_pr = child_w(node, "tblPr");
    let tbl_grid = child_w(node, "tblGrid");

    // Column widths from tblGrid
    let col_widths: Vec<f64> = tbl_grid.map(|g| {
        children_w(g, "gridCol")
            .iter()
            .map(|c| attr_w(*c, "w").map(|v| twips_to_pt(&v)).unwrap_or(72.0))
            .collect()
    }).unwrap_or_default();

    // Table borders
    let borders = tbl_pr.and_then(|p| child_w(p, "tblBorders"))
        .map(|b| parse_table_borders(b))
        .unwrap_or_default();

    // Cell margins
    let (cm_top, cm_bot, cm_left, cm_right) = tbl_pr
        .and_then(|p| child_w(p, "tblCellMar"))
        .map(|m| (
            child_w(m, "top").and_then(|n| attr_w(n, "w")).map(|v| twips_to_pt(&v)).unwrap_or(0.0),
            child_w(m, "bottom").and_then(|n| attr_w(n, "w")).map(|v| twips_to_pt(&v)).unwrap_or(0.0),
            child_w(m, "left").and_then(|n| attr_w(n, "w")).map(|v| twips_to_pt(&v)).unwrap_or(3.6),
            child_w(m, "right").and_then(|n| attr_w(n, "w")).map(|v| twips_to_pt(&v)).unwrap_or(3.6),
        ))
        .unwrap_or((0.0, 0.0, 3.6, 3.6));

    let mut rows = vec![];
    for tr_node in children_w(node, "tr") {
        let row = parse_table_row(tr_node, style_map, num_map, media_map);
        rows.push(row);
    }

    DocTable {
        col_widths,
        rows,
        borders,
        cell_margin_top: cm_top,
        cell_margin_bottom: cm_bot,
        cell_margin_left: cm_left,
        cell_margin_right: cm_right,
    }
}

fn parse_table_row(
    node: roxmltree::Node,
    style_map: &StyleMap,
    num_map: &mut NumberingMap,
    media_map: &HashMap<String, String>,
) -> DocTableRow {
    let tr_pr = child_w(node, "trPr");
    let row_height = tr_pr
        .and_then(|p| child_w(p, "trHeight"))
        .and_then(|h| attr_w(h, "val"))
        .map(|v| twips_to_pt(&v));
    let is_header = tr_pr.and_then(|p| child_w(p, "tblHeader")).is_some();

    let mut cells = vec![];
    for tc_node in children_w(node, "tc") {
        let cell = parse_table_cell(tc_node, style_map, num_map, media_map);
        cells.push(cell);
    }

    DocTableRow { cells, row_height, is_header }
}

fn parse_table_cell(
    node: roxmltree::Node,
    style_map: &StyleMap,
    num_map: &mut NumberingMap,
    media_map: &HashMap<String, String>,
) -> DocTableCell {
    let tc_pr = child_w(node, "tcPr");

    let col_span = tc_pr
        .and_then(|p| child_w(p, "gridSpan"))
        .and_then(|g| attr_w(g, "val"))
        .and_then(|v| v.parse().ok())
        .unwrap_or(1);

    let v_merge = tc_pr.and_then(|p| child_w(p, "vMerge")).map(|m| {
        attr_w(m, "val").map(|v| v == "restart").unwrap_or(true)
    });

    let borders = tc_pr.and_then(|p| child_w(p, "tcBorders"))
        .map(|b| parse_cell_borders(b))
        .unwrap_or_default();

    let background = tc_pr.and_then(|p| child_w(p, "shd"))
        .and_then(|s| attr_w(s, "fill"))
        .filter(|f| f != "auto" && f.len() == 6)
        .map(|f| f.to_lowercase());

    let v_align = tc_pr.and_then(|p| child_w(p, "vAlign"))
        .and_then(|v| attr_w(v, "val"))
        .unwrap_or_else(|| "top".to_string());

    let width_pt = tc_pr.and_then(|p| child_w(p, "tcW"))
        .and_then(|w| {
            let wtype = attr_w(w, "type").unwrap_or_default();
            if wtype == "dxa" {
                attr_w(w, "w").map(|v| twips_to_pt(&v))
            } else { None }
        });

    let mut content = vec![];
    for p_node in children_w(node, "p") {
        content.push(parse_paragraph(p_node, style_map, num_map, media_map));
    }

    DocTableCell { content, col_span, v_merge, borders, background, v_align, width_pt }
}

fn parse_table_borders(node: roxmltree::Node) -> TableBorders {
    TableBorders {
        top: child_w(node, "top").map(parse_border_spec),
        bottom: child_w(node, "bottom").map(parse_border_spec),
        left: child_w(node, "left").map(parse_border_spec),
        right: child_w(node, "right").map(parse_border_spec),
        inside_h: child_w(node, "insideH").map(parse_border_spec),
        inside_v: child_w(node, "insideV").map(parse_border_spec),
    }
}

fn parse_cell_borders(node: roxmltree::Node) -> CellBorders {
    CellBorders {
        top: child_w(node, "top").map(parse_border_spec),
        bottom: child_w(node, "bottom").map(parse_border_spec),
        left: child_w(node, "left").map(parse_border_spec),
        right: child_w(node, "right").map(parse_border_spec),
    }
}

fn parse_border_spec(node: roxmltree::Node) -> BorderSpec {
    let style = attr_w(node, "val").unwrap_or_else(|| "none".to_string());
    let width = attr_w(node, "sz").map(|v| {
        v.parse::<f64>().unwrap_or(4.0) / 8.0  // eighth-points → pt
    }).unwrap_or(0.5);
    let color = attr_w(node, "color").filter(|c| c != "auto").map(|c| c.to_lowercase());
    BorderSpec { width, color, style }
}

// ===== Helpers =====

fn normalize_align(s: &str) -> &str {
    match s {
        "both" | "distribute" => "justify",
        "right" | "end" => "right",
        "center" => "center",
        _ => "left",
    }
}

fn apply_direct_para(base: &mut ParaFmt, direct: &ParaFmt) {
    if direct.alignment.is_some() { base.alignment = direct.alignment.clone(); }
    if direct.indent_left.is_some() { base.indent_left = direct.indent_left; }
    if direct.indent_right.is_some() { base.indent_right = direct.indent_right; }
    if direct.indent_first.is_some() { base.indent_first = direct.indent_first; }
    if direct.space_before.is_some() { base.space_before = direct.space_before; }
    if direct.space_after.is_some() { base.space_after = direct.space_after; }
    if direct.line_spacing_val.is_some() { base.line_spacing_val = direct.line_spacing_val; }
    if direct.line_spacing_rule.is_some() { base.line_spacing_rule = direct.line_spacing_rule.clone(); }
    if direct.num_id.is_some() { base.num_id = direct.num_id; }
    if direct.num_level.is_some() { base.num_level = direct.num_level; }
}

fn apply_direct_run(base: &mut RunFmt, direct: &RunFmt) {
    if direct.bold.is_some() { base.bold = direct.bold; }
    if direct.italic.is_some() { base.italic = direct.italic; }
    if direct.underline.is_some() { base.underline = direct.underline; }
    if direct.strikethrough.is_some() { base.strikethrough = direct.strikethrough; }
    if direct.font_size.is_some() { base.font_size = direct.font_size; }
    if direct.color.is_some() { base.color = direct.color.clone(); }
    if direct.font_family_ascii.is_some() { base.font_family_ascii = direct.font_family_ascii.clone(); }
    if direct.font_family_east_asia.is_some() { base.font_family_east_asia = direct.font_family_east_asia.clone(); }
    if direct.background.is_some() { base.background = direct.background.clone(); }
}

fn parse_rels(xml: &str) -> HashMap<String, String> {
    let mut map = HashMap::new();
    if xml.is_empty() { return map; }
    let doc = match XmlDoc::parse(xml) { Ok(d) => d, Err(_) => return map };
    for rel in doc.root_element().children().filter(|n| n.tag_name().name() == "Relationship") {
        if let (Some(id), Some(target)) = (rel.attribute("Id"), rel.attribute("Target")) {
            map.insert(id.to_string(), target.to_string());
        }
    }
    map
}

fn read_zip_entry(zip: &mut ZipArchive<std::io::Cursor<&[u8]>>, path: &str) -> Result<String, String> {
    let mut entry = zip.by_name(path).map_err(|e| format!("{}: {}", path, e))?;
    let mut s = String::new();
    entry.read_to_string(&mut s).map_err(|e| e.to_string())?;
    Ok(s)
}

fn read_zip_bytes(zip: &mut ZipArchive<std::io::Cursor<&[u8]>>, path: &str) -> Result<Vec<u8>, String> {
    let mut entry = zip.by_name(path).map_err(|e| format!("{}: {}", path, e))?;
    let mut buf = vec![];
    entry.read_to_end(&mut buf).map_err(|e| e.to_string())?;
    Ok(buf)
}
