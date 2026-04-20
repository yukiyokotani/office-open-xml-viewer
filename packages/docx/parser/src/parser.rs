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

type Zip<'a> = ZipArchive<std::io::Cursor<&'a [u8]>>;

/// Section-level header/footer references collected from sectPr.
/// Maps reference type ("default" | "first" | "even") to the target xml path (e.g. "header1.xml").
#[derive(Default)]
struct SectionRefs {
    headers: HashMap<String, String>,
    footers: HashMap<String, String>,
}

pub fn parse(data: &[u8]) -> Result<Document, String> {
    let cursor = std::io::Cursor::new(data);
    let mut zip = ZipArchive::new(cursor).map_err(|e| e.to_string())?;

    let rels_xml = read_zip_entry(&mut zip, "word/_rels/document.xml.rels")
        .unwrap_or_default();
    let rel_map = parse_rels(&rels_xml);

    // Styles are referenced from the document relationships (Target may be
    // "styles.xml" or "styles2.xml"). Fall back to "word/styles.xml" for old files.
    let styles_path = find_rel_target(&rels_xml, "styles")
        .map(|t| if t.starts_with('/') { t.trim_start_matches('/').to_string() } else { format!("word/{}", t) })
        .unwrap_or_else(|| "word/styles.xml".to_string());
    let style_map = read_zip_entry(&mut zip, &styles_path)
        .map(|s| StyleMap::parse(&s))
        .unwrap_or_else(|_| StyleMap::parse(""));

    let numbering_path = find_rel_target(&rels_xml, "numbering")
        .map(|t| if t.starts_with('/') { t.trim_start_matches('/').to_string() } else { format!("word/{}", t) })
        .unwrap_or_else(|| "word/numbering.xml".to_string());
    let mut num_map = read_zip_entry(&mut zip, &numbering_path)
        .map(|s| NumberingMap::parse(&s))
        .unwrap_or_default();

    // Theme is referenced by a relationship with Type ending in "/theme" — resolve
    // to word/<target> and parse the clrScheme.
    let theme = find_rel_target(&rels_xml, "theme")
        .map(|t| {
            let p = if t.starts_with('/') { t.trim_start_matches('/').to_string() } else { format!("word/{}", t) };
            read_zip_entry(&mut zip, &p).map(|s| ThemeColors::parse(&s)).unwrap_or_default()
        })
        .unwrap_or_default();

    let media_map = load_media_map(&mut zip, &rel_map, "word/");

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

    let (section, refs) = parse_section(sect_pr, &rel_map);

    let body = parse_body_elements(body_node, &style_map, &mut num_map, &media_map, &rel_map, &theme);

    let headers = load_header_footer_set(&mut zip, &refs.headers, "hdr", &style_map, &mut num_map, &theme);
    let footers = load_header_footer_set(&mut zip, &refs.footers, "ftr", &style_map, &mut num_map, &theme);

    Ok(Document { section, body, headers, footers })
}

/// Resolve scheme color names (accent1..6, dk1, dk2, lt1, lt2, hlink, folHlink)
/// to hex strings parsed from word/theme/themeN.xml clrScheme.
#[derive(Debug, Default, Clone)]
pub struct ThemeColors {
    map: HashMap<String, String>,
}

impl ThemeColors {
    fn parse(xml: &str) -> Self {
        let mut map: HashMap<String, String> = HashMap::new();
        let doc = match XmlDoc::parse(xml) { Ok(d) => d, Err(_) => return Self { map } };
        let Some(scheme) = doc.descendants().find(|n| n.is_element() && n.tag_name().name() == "clrScheme") else {
            return Self { map };
        };
        for child in scheme.children().filter(|n| n.is_element()) {
            let name = child.tag_name().name().to_string();
            let hex = child.children().filter(|n| n.is_element()).find_map(|n| {
                match n.tag_name().name() {
                    "srgbClr" => n.attribute("val").map(|v| v.to_uppercase()),
                    "sysClr" => n.attribute("lastClr").map(|v| v.to_uppercase()),
                    _ => None,
                }
            });
            if let Some(h) = hex { map.insert(name, h); }
        }
        Self { map }
    }

    fn resolve(&self, scheme_name: &str) -> Option<String> {
        // "bg1"/"bg2"/"tx1"/"tx2" map onto lt1/lt2/dk1/dk2 per spec
        let key = match scheme_name {
            "bg1" => "lt1",
            "bg2" => "lt2",
            "tx1" => "dk1",
            "tx2" => "dk2",
            other => other,
        };
        self.map.get(key).cloned()
    }
}

fn find_rel_target(rels_xml: &str, type_suffix: &str) -> Option<String> {
    if rels_xml.is_empty() { return None; }
    let doc = XmlDoc::parse(rels_xml).ok()?;
    for rel in doc.root_element().children().filter(|n| n.tag_name().name() == "Relationship") {
        if let (Some(ty), Some(target)) = (rel.attribute("Type"), rel.attribute("Target")) {
            if ty.ends_with(&format!("/{}", type_suffix)) {
                return Some(target.to_string());
            }
        }
    }
    None
}

fn parse_body_elements(
    body_node: roxmltree::Node,
    style_map: &StyleMap,
    num_map: &mut NumberingMap,
    media_map: &HashMap<String, String>,
    rel_map: &HashMap<String, String>,
    theme: &ThemeColors,
) -> Vec<BodyElement> {
    let mut body: Vec<BodyElement> = Vec::new();
    // The body-level sectPr (the last element) defines the final section and
    // is not a page break. Mid-body sectPrs (nested in pPr) DO imply a page break.
    let body_children = element_children_flat(body_node);
    let body_level_sect_pr = body_children
        .iter()
        .last()
        .copied()
        .filter(|n| n.tag_name().name() == "sectPr");
    let body_level_sect_id = body_level_sect_pr.map(|n| n.id());

    for child in body_children {
        match child.tag_name().name() {
            "p" => {
                let result = parse_paragraph(child, style_map, num_map, media_map, rel_map, theme);
                let is_page_break_only = result.runs.len() == 1 && matches!(
                    result.runs[0],
                    DocRun::Break { break_type: BreakType::Page }
                );
                if is_page_break_only {
                    body.push(BodyElement::PageBreak);
                    continue;
                }
                body.push(BodyElement::Paragraph(result));
                // A section break inside pPr (that isn't the final body-level sectPr)
                // terminates the current section and starts a new one on a new page.
                let has_mid_section_break = child_w(child, "pPr")
                    .and_then(|ppr| child_w(ppr, "sectPr"))
                    .is_some();
                if has_mid_section_break {
                    body.push(BodyElement::PageBreak);
                }
            }
            "tbl" => {
                let tbl = parse_table(child, style_map, num_map, media_map, rel_map, theme);
                body.push(BodyElement::Table(tbl));
            }
            "sectPr" => {
                // Mid-body loose sectPr (rare) would behave like a page break.
                // The final body-level sectPr only defines section settings — skip it.
                if Some(child.id()) != body_level_sect_id {
                    body.push(BodyElement::PageBreak);
                }
            }
            _ => {}
        }
    }
    body
}

fn load_media_map(
    zip: &mut Zip,
    rel_map: &HashMap<String, String>,
    base_dir: &str,
) -> HashMap<String, String> {
    let mut media_map: HashMap<String, String> = HashMap::new();
    for (rid, target) in rel_map {
        if target.contains("media/") || target.contains("image") {
            let path = if target.starts_with('/') {
                target.trim_start_matches('/').to_string()
            } else {
                format!("{}{}", base_dir, target)
            };
            if let Ok(bytes) = read_zip_bytes(zip, &path) {
                let mime = if path.ends_with(".png") { "image/png" }
                    else if path.ends_with(".jpg") || path.ends_with(".jpeg") { "image/jpeg" }
                    else if path.ends_with(".gif") { "image/gif" }
                    else { "image/png" };
                let b64 = B64.encode(&bytes);
                media_map.insert(rid.clone(), format!("data:{};base64,{}", mime, b64));
            }
        }
    }
    media_map
}

fn load_header_footer_set(
    zip: &mut Zip,
    type_to_target: &HashMap<String, String>,
    root_tag: &str,
    style_map: &StyleMap,
    num_map: &mut NumberingMap,
    theme: &ThemeColors,
) -> HeadersFooters {
    let mut out = HeadersFooters::default();
    for (kind, target) in type_to_target {
        let path = format!("word/{}", target);
        let xml = match read_zip_entry(zip, &path) {
            Ok(s) => s,
            Err(_) => continue,
        };

        // Per-file rels for image resolution
        let stem = target.trim_end_matches(".xml");
        let rels_path = format!("word/_rels/{}.xml.rels", stem);
        let rels_xml = read_zip_entry(zip, &rels_path).unwrap_or_default();
        let local_rel_map = parse_rels(&rels_xml);
        let local_media_map = load_media_map(zip, &local_rel_map, "word/");

        let xml_doc = match XmlDoc::parse(&xml) {
            Ok(d) => d,
            Err(_) => continue,
        };
        let Some(root) = xml_doc.root_element().descendants().find(|n| n.tag_name().name() == root_tag) else {
            continue;
        };

        let body = parse_body_elements(root, style_map, num_map, &local_media_map, &local_rel_map, theme);
        let hf = HeaderFooter { body };
        match kind.as_str() {
            "first" => out.first = Some(hf),
            "even" => out.even = Some(hf),
            _ => out.default = Some(hf),
        }
    }
    out
}

fn parse_section(sect_pr: Option<roxmltree::Node>, rel_map: &HashMap<String, String>) -> (SectionProps, SectionRefs) {
    let default = SectionProps {
        page_width: 612.0,
        page_height: 792.0,
        margin_top: 72.0,
        margin_right: 72.0,
        margin_bottom: 72.0,
        margin_left: 72.0,
        header_distance: 36.0,
        footer_distance: 36.0,
        title_page: false,
        even_and_odd_headers: false,
    };

    let Some(sp) = sect_pr else { return (default, SectionRefs::default()) };

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
        if let Some(v) = attr_w(pg_mar, "header") { props.header_distance = twips_to_pt(&v); }
        if let Some(v) = attr_w(pg_mar, "footer") { props.footer_distance = twips_to_pt(&v); }
    }
    props.title_page = child_w(sp, "titlePg").is_some();

    // Collect header/footer references
    let mut refs = SectionRefs::default();
    for child in sp.children().filter(|n| n.is_element()) {
        let local = child.tag_name().name();
        if local != "headerReference" && local != "footerReference" { continue; }
        let kind = attr_w(child, "type").unwrap_or_else(|| "default".to_string());
        let rid = child.attribute((R_NS, "id"))
            .or_else(|| child.attribute("id"))
            .map(|s| s.to_string());
        let Some(rid) = rid else { continue };
        let Some(target) = rel_map.get(&rid) else { continue };
        let target = target.trim_start_matches('/').to_string();
        if local == "headerReference" {
            refs.headers.insert(kind, target);
        } else {
            refs.footers.insert(kind, target);
        }
    }

    (props, refs)
}

fn parse_paragraph(
    node: roxmltree::Node,
    style_map: &StyleMap,
    num_map: &mut NumberingMap,
    media_map: &HashMap<String, String>,
    rel_map: &HashMap<String, String>,
    theme: &ThemeColors,
) -> DocParagraph {
    // Get style ID from pPr/pStyle. When absent, resolve_para falls back to the
    // paragraph style marked w:default="1" via StyleMap::default_para_style_id.
    let ppr_node = child_w(node, "pPr");
    let explicit_style_id = ppr_node
        .and_then(|p| child_w(p, "pStyle"))
        .and_then(|s| attr_w(s, "val"));

    // Resolve base formatting from style
    let (mut base_para, mut base_run) = style_map.resolve_para(explicit_style_id.as_deref());

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
    let indent_right = base_para.indent_right.unwrap_or(0.0);
    let space_before = base_para.space_before.unwrap_or(0.0);
    let space_after = base_para.space_after.unwrap_or(0.0);
    let line_spacing = base_para.line_spacing_val.map(|v| LineSpacing {
        value: v,
        rule: base_para.line_spacing_rule.clone().unwrap_or_else(|| "auto".to_string()),
    });

    // Numbering — extract level data before advancing counter (avoids borrow conflict)
    let numbering = if let (Some(num_id), Some(num_level)) = (base_para.num_id, base_para.num_level) {
        if num_id != 0 {
            let (format, ind_left, tab) = num_map.get_level(num_id, num_level)
                .map(|l| (l.format.clone(), l.indent_left, l.tab))
                .unwrap_or_else(|| ("decimal".to_string(), 36.0, 18.0));
            let counter = num_map.advance(num_id, num_level);
            let text = num_map.resolve_text(num_id, num_level, counter);
            Some(NumberingInfo { num_id, level: num_level, format, text, indent_left: ind_left, tab })
        } else { None }
    } else { None };

    // Numbering level's pPr/ind overrides the paragraph style's indent
    let (indent_left, indent_first) = if let Some(ref num) = numbering {
        num_map.get_level(num.num_id, num.level)
            .map(|l| (l.indent_left, -l.tab))
            .unwrap_or((base_para.indent_left.unwrap_or(0.0), base_para.indent_first.unwrap_or(0.0)))
    } else {
        (base_para.indent_left.unwrap_or(0.0), base_para.indent_first.unwrap_or(0.0))
    };

    // Parse runs
    let mut runs = vec![];
    parse_para_content(node, &base_run, style_map, media_map, rel_map, theme, &mut runs);

    let tab_stops = base_para.tab_stops.clone().unwrap_or_default().into_iter()
        .map(|(pos, alignment, leader)| TabStop { pos, alignment, leader })
        .collect();

    DocParagraph {
        alignment,
        indent_left,
        indent_right,
        indent_first,
        space_before,
        space_after,
        line_spacing,
        numbering,
        tab_stops,
        runs,
        shading: base_para.shading.clone(),
        page_break_before: base_para.page_break_before.unwrap_or(false),
        contextual_spacing: base_para.contextual_spacing.unwrap_or(false),
        keep_next: base_para.keep_next.unwrap_or(false),
        keep_lines: base_para.keep_lines.unwrap_or(false),
        // ECMA-376 §17.3.1.44: widowControl defaults to true when absent.
        widow_control: base_para.widow_control.unwrap_or(true),
        borders: base_para.para_borders.clone(),
        // Fall back to the document's default paragraph style (w:default="1")
        // rather than the literal "Normal" — international templates often use
        // locale-specific IDs ("a", "標準", etc.) for the default style, and
        // contextualSpacing needs a stable ID to group adjacent paragraphs.
        style_id: explicit_style_id
            .clone()
            .or_else(|| style_map.default_para_style_id().map(str::to_string))
            .or_else(|| Some("Normal".to_string())),
        default_font_size: base_run.font_size,
    }
}

#[derive(Default)]
struct FieldState {
    /// Currently inside a field (between fldChar begin and end).
    active: bool,
    /// Have we passed the `separate` fldChar yet?
    past_separate: bool,
    /// Accumulated instruction text (PAGE, NUMPAGES, etc.)
    instruction: String,
    /// Formatting from the first instrText run — used as the field's display format.
    fmt: Option<RunFmt>,
    /// Fallback text captured between `separate` and `end`.
    fallback: String,
}

fn parse_para_content(
    node: roxmltree::Node,
    base_run: &RunFmt,
    style_map: &StyleMap,
    media_map: &HashMap<String, String>,
    rel_map: &HashMap<String, String>,
    theme: &ThemeColors,
    runs: &mut Vec<DocRun>,
) {
    let mut field = FieldState::default();

    for child in element_children_flat(node) {
        match child.tag_name().name() {
            "r" => {
                handle_run_in_para(child, base_run, style_map, media_map, theme, runs, &mut field, None);
            }
            "hyperlink" => {
                // Resolve URL from r:id via relationships
                let href = child.attribute((R_NS, "id"))
                    .or_else(|| child.attribute("id"))
                    .and_then(|rid| rel_map.get(rid).cloned());
                for r in child.children().filter(|n| n.is_element() && n.tag_name().name() == "r") {
                    handle_run_in_para(r, base_run, style_map, media_map, theme, runs, &mut field, Some(href.clone()));
                }
            }
            "ins" | "del" | "smartTag" => {
                parse_para_content(child, base_run, style_map, media_map, rel_map, theme, runs);
            }
            "fldSimple" => {
                let instr = attr_w(child, "instr").unwrap_or_default();
                // Collect formatting from the first contained run (if any)
                let mut fmt = base_run.clone();
                if let Some(r) = child.children().find(|n| n.is_element() && n.tag_name().name() == "r") {
                    if let Some(rpr) = child_w(r, "rPr") {
                        apply_direct_run(&mut fmt, &parse_run_fmt(rpr));
                    }
                }
                let fallback = extract_text_from_runs(child);
                runs.push(make_field_run(&instr, &fmt, &fallback));
            }
            _ => {}
        }
    }
}

fn handle_run_in_para(
    r_node: roxmltree::Node,
    base_run: &RunFmt,
    style_map: &StyleMap,
    media_map: &HashMap<String, String>,
    theme: &ThemeColors,
    runs: &mut Vec<DocRun>,
    field: &mut FieldState,
    // Outer None = not inside a hyperlink. Some(None) = hyperlink without URL. Some(Some(url)) = hyperlink with URL.
    link_href: Option<Option<String>>,
) {
    // Inspect this run for field control characters or instruction text first.
    let mut fld_char_type: Option<String> = None;
    let mut instr_text = String::new();
    for c in r_node.children().filter(|n| n.is_element()) {
        match c.tag_name().name() {
            "fldChar" => {
                if let Some(t) = attr_w(c, "fldCharType") {
                    fld_char_type = Some(t);
                }
            }
            "instrText" => {
                if let Some(t) = c.text() {
                    instr_text.push_str(t);
                }
            }
            _ => {}
        }
    }

    if let Some(ct) = fld_char_type {
        match ct.as_str() {
            "begin" => {
                field.active = true;
                field.past_separate = false;
                field.instruction.clear();
                field.fallback.clear();
                field.fmt = None;
            }
            "separate" => {
                field.past_separate = true;
            }
            "end" => {
                if field.active {
                    let fmt = field.fmt.clone().unwrap_or_else(|| base_run.clone());
                    runs.push(make_field_run(&field.instruction, &fmt, &field.fallback));
                }
                *field = FieldState::default();
            }
            _ => {}
        }
        return;
    }

    if field.active {
        if !field.past_separate {
            // Capture instruction text and remember the formatting of the first instruction run
            if !instr_text.is_empty() {
                field.instruction.push_str(&instr_text);
                if field.fmt.is_none() {
                    let mut fmt = base_run.clone();
                    if let Some(rpr) = child_w(r_node, "rPr") {
                        apply_direct_run(&mut fmt, &parse_run_fmt(rpr));
                    }
                    field.fmt = Some(fmt);
                }
            }
        } else {
            // Fallback/result text between separate and end — accumulate for "other" fields
            for c in r_node.children().filter(|n| n.is_element() && n.tag_name().name() == "t") {
                if let Some(t) = c.text() {
                    field.fallback.push_str(t);
                }
            }
        }
        return;
    }

    // Normal run
    parse_run_inner(r_node, base_run, style_map, media_map, theme, runs, link_href);
}

fn extract_text_from_runs(node: roxmltree::Node) -> String {
    let mut out = String::new();
    for n in node.descendants() {
        if n.is_element() && n.tag_name().name() == "t" {
            if let Some(t) = n.text() {
                out.push_str(t);
            }
        }
    }
    out
}

fn make_field_run(instr: &str, fmt: &RunFmt, fallback: &str) -> DocRun {
    let field_type = classify_field(instr);
    DocRun::Field(FieldRun {
        field_type,
        instruction: instr.trim().to_string(),
        fallback_text: fallback.to_string(),
        bold: fmt.bold.unwrap_or(false),
        italic: fmt.italic.unwrap_or(false),
        underline: fmt.underline.unwrap_or(false),
        strikethrough: fmt.strikethrough.unwrap_or(false),
        font_size: fmt.font_size.unwrap_or(DEFAULT_FONT_SIZE),
        color: fmt.color.clone(),
        font_family: fmt.font_family_ascii.clone().or(fmt.font_family_east_asia.clone()),
        background: fmt.background.clone(),
        vert_align: fmt.vert_align.clone(),
        all_caps: fmt.all_caps.unwrap_or(false),
        small_caps: fmt.small_caps.unwrap_or(false),
        double_strikethrough: fmt.dstrike.unwrap_or(false),
        highlight: fmt.highlight.clone(),
    })
}

fn classify_field(instr: &str) -> String {
    let token = instr.trim().split_whitespace().next().unwrap_or("").to_ascii_uppercase();
    match token.as_str() {
        "PAGE" => "page".to_string(),
        "NUMPAGES" => "numPages".to_string(),
        _ => "other".to_string(),
    }
}

fn parse_run_inner(
    node: roxmltree::Node,
    base_run: &RunFmt,
    style_map: &StyleMap,
    media_map: &HashMap<String, String>,
    theme: &ThemeColors,
    runs: &mut Vec<DocRun>,
    link_href: Option<Option<String>>,
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

    // Skip hidden runs entirely
    if fmt.vanish.unwrap_or(false) { return; }

    let is_link = link_href.is_some();
    let hyperlink = link_href.clone().flatten();

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
    let vert_align = fmt.vert_align.clone();
    let all_caps = fmt.all_caps.unwrap_or(false);
    let small_caps = fmt.small_caps.unwrap_or(false);
    let double_strikethrough = fmt.dstrike.unwrap_or(false);
    let highlight = fmt.highlight.clone();

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
                        vert_align: vert_align.clone(),
                        hyperlink: hyperlink.clone(),
                        all_caps,
                        small_caps,
                        double_strikethrough,
                        highlight: highlight.clone(),
                    }));
                }
            }
            "tab" => {
                // w:tab emits a horizontal tab character; layout handles tab stop alignment.
                runs.push(DocRun::Text(TextRun {
                    text: "\t".to_string(),
                    bold,
                    italic,
                    underline,
                    strikethrough,
                    font_size,
                    color: color.clone(),
                    font_family: font_family.clone(),
                    is_link,
                    background: fmt.background.clone(),
                    vert_align: vert_align.clone(),
                    hyperlink: hyperlink.clone(),
                    all_caps,
                    small_caps,
                    double_strikethrough,
                    highlight: highlight.clone(),
                }));
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
                for r in parse_inline_drawing(child, media_map, theme) {
                    runs.push(r);
                }
            }
            "AlternateContent" => {
                // mc:AlternateContent/mc:Choice may contain w:drawing
                if let Some(choice) = child.children().find(|n| n.tag_name().name() == "Choice") {
                    for inner in choice.children().filter(|n| n.is_element()) {
                        if inner.tag_name().name() == "drawing" {
                            for r in parse_inline_drawing(inner, media_map, theme) {
                                runs.push(r);
                            }
                        }
                    }
                }
            }
            _ => {}
        }
    }
}

fn parse_inline_drawing(
    node: roxmltree::Node,
    media_map: &HashMap<String, String>,
    theme: &ThemeColors,
) -> Vec<DocRun> {
    // Distinguish inline vs anchor
    let is_anchor = node.descendants().any(|n| n.tag_name().name() == "anchor");

    if !is_anchor {
        let container = match node.descendants().find(|n| n.tag_name().name() == "inline") {
            Some(c) => c,
            None => return vec![],
        };
        let extent = match container.children().find(|n| n.tag_name().name() == "extent") {
            Some(e) => e,
            None => return vec![],
        };
        let cx: f64 = match extent.attribute("cx").and_then(|v| v.parse().ok()) {
            Some(v) => v,
            None => return vec![],
        };
        let cy: f64 = match extent.attribute("cy").and_then(|v| v.parse().ok()) {
            Some(v) => v,
            None => return vec![],
        };
        let blip = match node.descendants().find(|n| n.tag_name().name() == "blip") {
            Some(b) => b,
            None => return vec![],
        };
        let r_id = match blip.attribute((R_NS, "embed")).or_else(|| blip.attribute("r:embed")) {
            Some(r) => r,
            None => return vec![],
        };
        let data_url = match media_map.get(r_id) {
            Some(u) => u.clone(),
            None => return vec![],
        };
        return vec![DocRun::Image(ImageRun {
            data_url,
            width_pt: cx / 12700.0,
            height_pt: cy / 12700.0,
            anchor: false,
            anchor_x_pt: 0.0,
            anchor_y_pt: 0.0,
            anchor_x_from_margin: false,
            anchor_y_from_para: false,
            color_replace_from: None,
            wrap_mode: None,
            dist_top: 0.0,
            dist_bottom: 0.0,
            dist_left: 0.0,
            dist_right: 0.0,
            wrap_side: None,
        })];
    }

    // ── Anchor image/shape ─────────────────────────────────
    let container = match node.descendants().find(|n| n.tag_name().name() == "anchor") {
        Some(c) => c,
        None => return vec![],
    };

    // Parse positionH / positionV with relativeFrom
    let (pos_x, x_from_margin) = parse_anchor_pos_h(&container);
    let (pos_y, y_from_para)   = parse_anchor_pos_v(&container);
    let anchor_meta = parse_anchor_wrap(&container);

    // behindDoc="1" flag — renderer uses this to draw shapes before text
    let behind_doc = container.attribute("behindDoc").map(|v| v == "1" || v == "true").unwrap_or(false);

    // Check for wgp (Word Graphics Group) — expands to multiple per-element entries
    if let Some(wgp) = container.descendants().find(|n| n.tag_name().name() == "wgp") {
        let mut out: Vec<DocRun> = Vec::new();
        for img in parse_wgp_images(wgp, media_map, pos_x, x_from_margin, pos_y, y_from_para, &anchor_meta) {
            out.push(DocRun::Image(img));
        }
        for mut shp in parse_wgp_shapes(wgp, theme, pos_x, x_from_margin, pos_y, y_from_para, &anchor_meta) {
            shp.behind_doc = behind_doc;
            out.push(DocRun::Shape(shp));
        }
        return out;
    }

    // wps:wsp directly under the anchor (no wgp wrapper)
    if let Some(wsp) = container.descendants().find(|n| n.tag_name().name() == "wsp") {
        if let Some(mut shp) = parse_wsp_shape(wsp, theme, pos_x, x_from_margin, pos_y, y_from_para, &anchor_meta, 1.0, 1.0, 0.0, 0.0, 0) {
            shp.behind_doc = behind_doc;
            return vec![DocRun::Shape(shp)];
        }
    }

    // Regular single-blip anchor
    let extent = match container.children().find(|n| n.tag_name().name() == "extent") {
        Some(e) => e,
        None => return vec![],
    };
    let cx: f64 = match extent.attribute("cx").and_then(|v| v.parse().ok()) {
        Some(v) => v,
        None => return vec![],
    };
    let cy: f64 = match extent.attribute("cy").and_then(|v| v.parse().ok()) {
        Some(v) => v,
        None => return vec![],
    };
    let blip = match node.descendants().find(|n| n.tag_name().name() == "blip") {
        Some(b) => b,
        None => return vec![],
    };
    let r_id = match blip.attribute((R_NS, "embed")).or_else(|| blip.attribute("r:embed")) {
        Some(r) => r,
        None => return vec![],
    };
    let data_url = match media_map.get(r_id) {
        Some(u) => u.clone(),
        None => return vec![],
    };
    vec![DocRun::Image(ImageRun {
        data_url,
        width_pt: cx / 12700.0,
        height_pt: cy / 12700.0,
        anchor: true,
        anchor_x_pt: pos_x,
        anchor_y_pt: pos_y,
        anchor_x_from_margin: x_from_margin,
        anchor_y_from_para: y_from_para,
        color_replace_from: None,
        wrap_mode: anchor_meta.wrap_mode.clone(),
        dist_top: anchor_meta.dist_top,
        dist_bottom: anchor_meta.dist_bottom,
        dist_left: anchor_meta.dist_left,
        dist_right: anchor_meta.dist_right,
        wrap_side: anchor_meta.wrap_side.clone(),
    })]
}

#[derive(Default, Clone)]
struct AnchorMeta {
    wrap_mode: Option<String>,
    wrap_side: Option<String>,
    dist_top: f64,
    dist_bottom: f64,
    dist_left: f64,
    dist_right: f64,
}

/// Parse wrap element and dist* padding from a wp:anchor container.
fn parse_anchor_wrap(container: &roxmltree::Node) -> AnchorMeta {
    let to_pt = |s: &str| s.parse::<f64>().ok().map(|v| v / 12700.0).unwrap_or(0.0);
    let dist_top = container.attribute("distT").map(to_pt).unwrap_or(0.0);
    let dist_bottom = container.attribute("distB").map(to_pt).unwrap_or(0.0);
    let dist_left = container.attribute("distL").map(to_pt).unwrap_or(0.0);
    let dist_right = container.attribute("distR").map(to_pt).unwrap_or(0.0);

    let mut wrap_mode: Option<String> = None;
    let mut wrap_side: Option<String> = None;

    for child in container.children().filter(|n| n.is_element()) {
        let name = child.tag_name().name();
        match name {
            "wrapSquare"       => { wrap_mode = Some("square".into());       wrap_side = child.attribute("wrapText").map(|s| s.to_string()); break; }
            "wrapTopAndBottom" => { wrap_mode = Some("topAndBottom".into()); break; }
            "wrapNone"         => { wrap_mode = Some("none".into());         break; }
            "wrapTight"        => { wrap_mode = Some("tight".into());        wrap_side = child.attribute("wrapText").map(|s| s.to_string()); break; }
            "wrapThrough"      => { wrap_mode = Some("through".into());      wrap_side = child.attribute("wrapText").map(|s| s.to_string()); break; }
            _ => {}
        }
    }

    AnchorMeta { wrap_mode, wrap_side, dist_top, dist_bottom, dist_left, dist_right }
}

/// Parse positionH — returns (posOffset_pt, needs_margin_offset).
/// "column" and "margin" relative offsets both mean: add marginLeft in the renderer.
fn parse_anchor_pos_h(container: &roxmltree::Node) -> (f64, bool) {
    let pos = match container.children().find(|n| n.tag_name().name() == "positionH") {
        Some(p) => p,
        None => return (0.0, false),
    };
    let rel = pos.attribute("relativeFrom").unwrap_or("page");
    let offset = pos.children()
        .find(|n| n.tag_name().name() == "posOffset")
        .and_then(|n| n.text())
        .and_then(|t| t.parse::<f64>().ok())
        .map(|emu| emu / 12700.0)
        .unwrap_or(0.0);
    let from_margin = matches!(rel, "column" | "margin" | "leftMargin" | "insideMargin");
    (offset, from_margin)
}

/// Parse positionV — returns (posOffset_pt, is_paragraph_relative).
fn parse_anchor_pos_v(container: &roxmltree::Node) -> (f64, bool) {
    let pos = match container.children().find(|n| n.tag_name().name() == "positionV") {
        Some(p) => p,
        None => return (0.0, false),
    };
    let rel = pos.attribute("relativeFrom").unwrap_or("page");
    let offset = pos.children()
        .find(|n| n.tag_name().name() == "posOffset")
        .and_then(|n| n.text())
        .and_then(|t| t.parse::<f64>().ok())
        .map(|emu| emu / 12700.0)
        .unwrap_or(0.0);
    let from_para = matches!(rel, "paragraph" | "line");
    (offset, from_para)
}

/// Expand a wp:wgp group into individual ImageRun entries.
/// Each pic child gets page-relative coordinates: group anchor origin + child offset within group.
fn parse_wgp_images(
    wgp: roxmltree::Node,
    media_map: &HashMap<String, String>,
    anchor_pos_x: f64,
    x_from_margin: bool,
    anchor_pos_y: f64,
    y_from_para: bool,
    anchor_meta: &AnchorMeta,
) -> Vec<ImageRun> {
    let mut results = Vec::new();
    // Iterate all pic descendants in the wgp (covers both direct children and nested grpSp)
    for pic in wgp.descendants().filter(|n| n.tag_name().name() == "pic") {
        // Position and size come from the pic's spPr > a:xfrm
        let sp_pr = match pic.children().find(|n| n.tag_name().name() == "spPr") {
            Some(s) => s,
            None => continue,
        };
        let xfrm = match sp_pr.children().find(|n| n.tag_name().name() == "xfrm") {
            Some(x) => x,
            None => continue,
        };
        let off = match xfrm.children().find(|n| n.tag_name().name() == "off") {
            Some(o) => o,
            None => continue,
        };
        let ext = match xfrm.children().find(|n| n.tag_name().name() == "ext") {
            Some(e) => e,
            None => continue,
        };
        let ox = off.attribute("x").and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0) / 12700.0;
        let oy = off.attribute("y").and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0) / 12700.0;
        let cx = ext.attribute("cx").and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0) / 12700.0;
        let cy = ext.attribute("cy").and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0) / 12700.0;

        if cx <= 0.0 || cy <= 0.0 { continue; }

        // Find the blip inside this pic
        let blip = match pic.descendants().find(|n| n.tag_name().name() == "blip") {
            Some(b) => b,
            None => continue,
        };
        let r_id = match blip.attribute((R_NS, "embed")).or_else(|| blip.attribute("r:embed")) {
            Some(r) => r,
            None => continue,
        };
        let data_url = match media_map.get(r_id) {
            Some(u) => u.clone(),
            None => continue,
        };

        // Parse a:clrChange if present — used to make a specific color transparent.
        // clrFrom specifies the source color; clrTo with alpha=0 means replace with transparent.
        let color_replace_from = blip.children()
            .find(|n| n.tag_name().name() == "clrChange")
            .and_then(|cc| cc.children().find(|n| n.tag_name().name() == "clrFrom"))
            .and_then(|cf| cf.children().find(|n| n.tag_name().name() == "srgbClr"))
            .and_then(|clr| clr.attribute("val").map(|v| v.to_uppercase()));

        results.push(ImageRun {
            data_url,
            width_pt: cx,
            height_pt: cy,
            anchor: true,
            // Combine the group's anchor offset with this pic's offset within the group
            anchor_x_pt: anchor_pos_x + ox,
            anchor_y_pt: anchor_pos_y + oy,
            anchor_x_from_margin: x_from_margin,
            anchor_y_from_para: y_from_para,
            color_replace_from,
            wrap_mode: anchor_meta.wrap_mode.clone(),
            dist_top: anchor_meta.dist_top,
            dist_bottom: anchor_meta.dist_bottom,
            dist_left: anchor_meta.dist_left,
            dist_right: anchor_meta.dist_right,
            wrap_side: anchor_meta.wrap_side.clone(),
        });
    }
    results
}

/// Expand wps:wsp descendants of a wgp into ShapeRun entries. Applies
/// wgp grpSpPr transform (chOff/chExt → off/ext scale) to each child shape.
fn parse_wgp_shapes(
    wgp: roxmltree::Node,
    theme: &ThemeColors,
    anchor_pos_x: f64,
    x_from_margin: bool,
    anchor_pos_y: f64,
    y_from_para: bool,
    anchor_meta: &AnchorMeta,
) -> Vec<ShapeRun> {
    // Read group transform: off/ext (page-relative) vs chOff/chExt (child coord space).
    let grp_xfrm = wgp.descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "grpSpPr")
        .and_then(|gsp| gsp.children().find(|n| n.is_element() && n.tag_name().name() == "xfrm"));
    let (off_x, off_y, ext_cx, ext_cy, ch_off_x, ch_off_y, ch_ext_cx, ch_ext_cy) = grp_xfrm
        .map(|x| {
            let off = x.children().find(|n| n.is_element() && n.tag_name().name() == "off");
            let ext = x.children().find(|n| n.is_element() && n.tag_name().name() == "ext");
            let ch_off = x.children().find(|n| n.is_element() && n.tag_name().name() == "chOff");
            let ch_ext = x.children().find(|n| n.is_element() && n.tag_name().name() == "chExt");
            (
                off.and_then(|o| o.attribute("x").and_then(|v| v.parse::<f64>().ok())).unwrap_or(0.0),
                off.and_then(|o| o.attribute("y").and_then(|v| v.parse::<f64>().ok())).unwrap_or(0.0),
                ext.and_then(|e| e.attribute("cx").and_then(|v| v.parse::<f64>().ok())).unwrap_or(0.0),
                ext.and_then(|e| e.attribute("cy").and_then(|v| v.parse::<f64>().ok())).unwrap_or(0.0),
                ch_off.and_then(|o| o.attribute("x").and_then(|v| v.parse::<f64>().ok())).unwrap_or(0.0),
                ch_off.and_then(|o| o.attribute("y").and_then(|v| v.parse::<f64>().ok())).unwrap_or(0.0),
                ch_ext.and_then(|e| e.attribute("cx").and_then(|v| v.parse::<f64>().ok())).unwrap_or(0.0),
                ch_ext.and_then(|e| e.attribute("cy").and_then(|v| v.parse::<f64>().ok())).unwrap_or(0.0),
            )
        })
        .unwrap_or((0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0));

    let sx = if ch_ext_cx > 0.0 && ext_cx > 0.0 { ext_cx / ch_ext_cx } else { 1.0 };
    let sy = if ch_ext_cy > 0.0 && ext_cy > 0.0 { ext_cy / ch_ext_cy } else { 1.0 };
    // Page-relative offset of the group origin, in EMU. ch_off is subtracted
    // because child coordinates are measured relative to chOff.
    let group_page_off_x_emu = off_x - ch_off_x * sx;
    let group_page_off_y_emu = off_y - ch_off_y * sy;

    let mut results = Vec::new();
    for (idx, wsp) in wgp.descendants().filter(|n| n.is_element() && n.tag_name().name() == "wsp").enumerate() {
        if let Some(shape) = parse_wsp_shape(
            wsp, theme,
            anchor_pos_x, x_from_margin,
            anchor_pos_y, y_from_para,
            anchor_meta,
            sx, sy,
            group_page_off_x_emu / 12700.0, group_page_off_y_emu / 12700.0,
            idx as u32,
        ) {
            results.push(shape);
        }
    }
    results
}

/// Parse a single wps:wsp into ShapeRun. `sx,sy` scale the shape's spPr/xfrm
/// from group child coord space to page EMU; `group_off_pt_*` are the group origin
/// on the page (in pt) so the shape's off.x/off.y (in child coord space) can be
/// translated to page-relative pt. For a standalone wsp (no wgp), pass sx=sy=1, group_off=0.
fn parse_wsp_shape(
    wsp: roxmltree::Node,
    theme: &ThemeColors,
    anchor_pos_x: f64,
    x_from_margin: bool,
    anchor_pos_y: f64,
    y_from_para: bool,
    anchor_meta: &AnchorMeta,
    sx: f64,
    sy: f64,
    group_off_pt_x: f64,
    group_off_pt_y: f64,
    z_order: u32,
) -> Option<ShapeRun> {
    let sp_pr = wsp.children().find(|n| n.is_element() && n.tag_name().name() == "spPr")?;
    let xfrm = sp_pr.children().find(|n| n.is_element() && n.tag_name().name() == "xfrm")?;
    let off = xfrm.children().find(|n| n.is_element() && n.tag_name().name() == "off")?;
    let ext = xfrm.children().find(|n| n.is_element() && n.tag_name().name() == "ext")?;
    let ox = off.attribute("x").and_then(|v| v.parse::<f64>().ok())?;
    let oy = off.attribute("y").and_then(|v| v.parse::<f64>().ok())?;
    let cx = ext.attribute("cx").and_then(|v| v.parse::<f64>().ok())?;
    let cy = ext.attribute("cy").and_then(|v| v.parse::<f64>().ok())?;
    if cx <= 0.0 || cy <= 0.0 { return None; }
    let rotation = xfrm.attribute("rot")
        .and_then(|v| v.parse::<f64>().ok())
        .map(|r| r / 60000.0) // OOXML rotation: 60000ths of a degree
        .unwrap_or(0.0);

    let width_pt = cx * sx / 12700.0;
    let height_pt = cy * sy / 12700.0;
    let anchor_x_pt = anchor_pos_x + group_off_pt_x + ox * sx / 12700.0;
    let anchor_y_pt = anchor_pos_y + group_off_pt_y + oy * sy / 12700.0;

    let cust_geom = sp_pr.children().find(|n| n.is_element() && n.tag_name().name() == "custGeom");
    let subpaths = cust_geom.map(parse_custom_geometry).unwrap_or_default();
    if subpaths.is_empty() { return None; }

    let fill = parse_shape_fill(sp_pr, theme);
    let (stroke, stroke_width) = sp_pr.children()
        .find(|n| n.is_element() && n.tag_name().name() == "ln")
        .map(|ln| {
            let has_no_fill = ln.children().any(|n| n.is_element() && n.tag_name().name() == "noFill");
            if has_no_fill { return (None, 0.0); }
            let color = ln.children()
                .find(|n| n.is_element() && n.tag_name().name() == "solidFill")
                .and_then(|sf| resolve_color_element(sf, theme));
            let w_emu = ln.attribute("w").and_then(|v| v.parse::<f64>().ok()).unwrap_or(9525.0);
            (color, w_emu / 12700.0)
        })
        .unwrap_or((None, 0.0));

    Some(ShapeRun {
        width_pt,
        height_pt,
        anchor_x_pt,
        anchor_y_pt,
        anchor_x_from_margin: x_from_margin,
        anchor_y_from_para: y_from_para,
        behind_doc: false,
        z_order,
        subpaths,
        fill,
        stroke,
        stroke_width,
        rotation,
        wrap_mode: anchor_meta.wrap_mode.clone(),
    })
}

/// Parse a shape's fill (solidFill or gradFill). Returns None for noFill/missing.
fn parse_shape_fill(sp_pr: roxmltree::Node, theme: &ThemeColors) -> Option<ShapeFill> {
    for child in sp_pr.children().filter(|n| n.is_element()) {
        match child.tag_name().name() {
            "solidFill" => {
                return resolve_color_element(child, theme).map(|c| ShapeFill::Solid { color: c });
            }
            "gradFill" => {
                return parse_grad_fill(child, theme);
            }
            "noFill" => return None,
            _ => {}
        }
    }
    None
}

fn parse_grad_fill(node: roxmltree::Node, theme: &ThemeColors) -> Option<ShapeFill> {
    let gs_lst = node.children().find(|n| n.is_element() && n.tag_name().name() == "gsLst")?;
    let mut stops: Vec<GradientStop> = Vec::new();
    for gs in gs_lst.children().filter(|n| n.is_element() && n.tag_name().name() == "gs") {
        let pos = gs.attribute("pos").and_then(|v| v.parse::<f64>().ok()).map(|p| p / 100000.0).unwrap_or(0.0);
        if let Some(color) = resolve_color_element(gs, theme) {
            stops.push(GradientStop { position: pos, color });
        }
    }
    if stops.is_empty() { return None; }

    // Linear direction (a:lin ang = "60000"ths of a degree)
    let (angle, grad_type) = if let Some(lin) = node.children().find(|n| n.is_element() && n.tag_name().name() == "lin") {
        let ang = lin.attribute("ang").and_then(|v| v.parse::<f64>().ok()).map(|a| a / 60000.0).unwrap_or(0.0);
        (ang, "linear".to_string())
    } else if node.children().any(|n| n.is_element() && n.tag_name().name() == "path") {
        (0.0, "radial".to_string())
    } else {
        (0.0, "linear".to_string())
    };

    Some(ShapeFill::Gradient { stops, angle, grad_type })
}

/// Parse <a:custGeom><a:pathLst><a:path w="W" h="H">...</a:path></a:pathLst>.
/// Path coords inside each <a:path> are absolute within W×H; normalize to [0,1].
fn parse_custom_geometry(cust_geom: roxmltree::Node) -> Vec<Vec<PathCmd>> {
    let Some(path_lst) = cust_geom.children().find(|n| n.is_element() && n.tag_name().name() == "pathLst") else {
        return vec![];
    };
    let mut subpaths: Vec<Vec<PathCmd>> = Vec::new();
    for path in path_lst.children().filter(|n| n.is_element() && n.tag_name().name() == "path") {
        let pw = path.attribute("w").and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0);
        let ph = path.attribute("h").and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0);
        if pw <= 0.0 || ph <= 0.0 { continue; }
        let mut cmds: Vec<PathCmd> = Vec::new();
        for cmd in path.children().filter(|n| n.is_element()) {
            let name = cmd.tag_name().name();
            let pts: Vec<(f64, f64)> = cmd.children()
                .filter(|n| n.is_element() && n.tag_name().name() == "pt")
                .filter_map(|p| {
                    let x = p.attribute("x").and_then(|v| v.parse::<f64>().ok())?;
                    let y = p.attribute("y").and_then(|v| v.parse::<f64>().ok())?;
                    Some((x / pw, y / ph))
                })
                .collect();
            match name {
                "moveTo" => { if let Some(p) = pts.first() { cmds.push(PathCmd::MoveTo { x: p.0, y: p.1 }); } }
                "lnTo" => { if let Some(p) = pts.first() { cmds.push(PathCmd::LineTo { x: p.0, y: p.1 }); } }
                "cubicBezTo" => {
                    if pts.len() >= 3 {
                        cmds.push(PathCmd::CubicBezTo {
                            x1: pts[0].0, y1: pts[0].1,
                            x2: pts[1].0, y2: pts[1].1,
                            x:  pts[2].0, y:  pts[2].1,
                        });
                    }
                }
                "arcTo" => {
                    let wr = cmd.attribute("wR").and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0) / pw;
                    let hr = cmd.attribute("hR").and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0) / ph;
                    let st_ang = cmd.attribute("stAng").and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0) / 60000.0;
                    let sw_ang = cmd.attribute("swAng").and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0) / 60000.0;
                    cmds.push(PathCmd::ArcTo { wr, hr, st_ang, sw_ang });
                }
                "close" => cmds.push(PathCmd::Close),
                _ => {}
            }
        }
        if !cmds.is_empty() { subpaths.push(cmds); }
    }
    subpaths
}

/// Resolve a color container (e.g. <a:solidFill>, <a:gs>) into a hex string by
/// inspecting its child: <a:srgbClr>, <a:schemeClr>, or <a:sysClr>, applying
/// any lumMod/lumOff/alpha modifiers declared on the inner color element.
fn resolve_color_element(container: roxmltree::Node, theme: &ThemeColors) -> Option<String> {
    for c in container.children().filter(|n| n.is_element()) {
        let base = match c.tag_name().name() {
            "srgbClr" => c.attribute("val").map(|v| v.to_uppercase()),
            "schemeClr" => c.attribute("val").and_then(|name| theme.resolve(name)),
            "sysClr" => c.attribute("lastClr").map(|v| v.to_uppercase())
                .or_else(|| c.attribute("val").map(|v| v.to_uppercase())),
            _ => None,
        };
        let Some(hex) = base else { continue };
        return Some(apply_color_mods(&hex, c));
    }
    None
}

/// Apply OOXML color modifiers (lumMod / lumOff) declared as child elements.
/// lumMod multiplies L by mod/100000; lumOff adds off/100000 to L. Both are
/// evaluated in HSL space per ECMA-376.
fn apply_color_mods(hex: &str, color_node: roxmltree::Node) -> String {
    let (r, g, b) = hex_to_rgb(hex);
    let (mut h, mut s, mut l) = rgb_to_hsl(r, g, b);

    for m in color_node.children().filter(|n| n.is_element()) {
        let val = m.attribute("val").and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0) / 100000.0;
        match m.tag_name().name() {
            "lumMod" => l *= val,
            "lumOff" => l += val,
            "satMod" => s *= val,
            "satOff" => s += val,
            "hueMod" => h *= val,
            "hueOff" => h += val,
            "shade" => { l *= val; s *= val; }
            "tint" => { l = l + (1.0 - l) * val; }
            _ => {}
        }
    }
    l = l.clamp(0.0, 1.0);
    s = s.clamp(0.0, 1.0);
    let (nr, ng, nb) = hsl_to_rgb(h, s, l);
    format!("{:02X}{:02X}{:02X}", nr, ng, nb)
}

fn hex_to_rgb(hex: &str) -> (u8, u8, u8) {
    let r = u8::from_str_radix(&hex[0..2], 16).unwrap_or(0);
    let g = u8::from_str_radix(&hex[2..4], 16).unwrap_or(0);
    let b = u8::from_str_radix(&hex[4..6], 16).unwrap_or(0);
    (r, g, b)
}

fn rgb_to_hsl(r: u8, g: u8, b: u8) -> (f64, f64, f64) {
    let rf = r as f64 / 255.0;
    let gf = g as f64 / 255.0;
    let bf = b as f64 / 255.0;
    let max = rf.max(gf.max(bf));
    let min = rf.min(gf.min(bf));
    let l = (max + min) / 2.0;
    let d = max - min;
    if d == 0.0 { return (0.0, 0.0, l); }
    let s = if l > 0.5 { d / (2.0 - max - min) } else { d / (max + min) };
    let h = if max == rf {
        ((gf - bf) / d + if gf < bf { 6.0 } else { 0.0 }) / 6.0
    } else if max == gf {
        ((bf - rf) / d + 2.0) / 6.0
    } else {
        ((rf - gf) / d + 4.0) / 6.0
    };
    (h, s, l)
}

fn hsl_to_rgb(h: f64, s: f64, l: f64) -> (u8, u8, u8) {
    if s == 0.0 {
        let v = (l * 255.0).round().clamp(0.0, 255.0) as u8;
        return (v, v, v);
    }
    let q = if l < 0.5 { l * (1.0 + s) } else { l + s - l * s };
    let p = 2.0 * l - q;
    let hue2rgb = |p: f64, q: f64, mut t: f64| -> f64 {
        if t < 0.0 { t += 1.0; }
        if t > 1.0 { t -= 1.0; }
        if t < 1.0 / 6.0 { return p + (q - p) * 6.0 * t; }
        if t < 1.0 / 2.0 { return q; }
        if t < 2.0 / 3.0 { return p + (q - p) * (2.0 / 3.0 - t) * 6.0; }
        p
    };
    let r = hue2rgb(p, q, h + 1.0 / 3.0);
    let g = hue2rgb(p, q, h);
    let b = hue2rgb(p, q, h - 1.0 / 3.0);
    (
        (r * 255.0).round().clamp(0.0, 255.0) as u8,
        (g * 255.0).round().clamp(0.0, 255.0) as u8,
        (b * 255.0).round().clamp(0.0, 255.0) as u8,
    )
}

// ===== Table parsing =====

fn parse_table(
    node: roxmltree::Node,
    style_map: &StyleMap,
    num_map: &mut NumberingMap,
    media_map: &HashMap<String, String>,
    rel_map: &HashMap<String, String>,
    theme: &ThemeColors,
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
    for tr_node in children_w_flat(node, "tr") {
        let row = parse_table_row(tr_node, style_map, num_map, media_map, rel_map, theme);
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
    rel_map: &HashMap<String, String>,
    theme: &ThemeColors,
) -> DocTableRow {
    let tr_pr = child_w(node, "trPr");
    let row_height = tr_pr
        .and_then(|p| child_w(p, "trHeight"))
        .and_then(|h| attr_w(h, "val"))
        .map(|v| twips_to_pt(&v));
    let is_header = tr_pr.and_then(|p| child_w(p, "tblHeader")).is_some();

    let mut cells = vec![];
    for tc_node in children_w_flat(node, "tc") {
        let cell = parse_table_cell(tc_node, style_map, num_map, media_map, rel_map, theme);
        cells.push(cell);
    }

    DocTableRow { cells, row_height, is_header }
}

fn parse_table_cell(
    node: roxmltree::Node,
    style_map: &StyleMap,
    num_map: &mut NumberingMap,
    media_map: &HashMap<String, String>,
    rel_map: &HashMap<String, String>,
    theme: &ThemeColors,
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
    for p_node in children_w_flat(node, "p") {
        content.push(parse_paragraph(p_node, style_map, num_map, media_map, rel_map, theme));
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
    if direct.tab_stops.is_some() { base.tab_stops = direct.tab_stops.clone(); }
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
    if direct.vert_align.is_some() { base.vert_align = direct.vert_align.clone(); }
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

/// Refuse to decompress individual ZIP entries larger than 512 MiB to prevent
/// zip-bomb DoS.
const MAX_ZIP_ENTRY_BYTES: u64 = 512 * 1024 * 1024;

fn read_zip_entry(zip: &mut Zip, path: &str) -> Result<String, String> {
    let mut entry = zip.by_name(path).map_err(|e| format!("{}: {}", path, e))?;
    if entry.size() > MAX_ZIP_ENTRY_BYTES {
        return Err(format!("{}: exceeds size limit", path));
    }
    let mut s = String::new();
    entry.by_ref().take(MAX_ZIP_ENTRY_BYTES).read_to_string(&mut s).map_err(|e| e.to_string())?;
    Ok(s)
}

fn read_zip_bytes(zip: &mut Zip, path: &str) -> Result<Vec<u8>, String> {
    let mut entry = zip.by_name(path).map_err(|e| format!("{}: {}", path, e))?;
    if entry.size() > MAX_ZIP_ENTRY_BYTES {
        return Err(format!("{}: exceeds size limit", path));
    }
    let mut buf = vec![];
    entry.by_ref().take(MAX_ZIP_ENTRY_BYTES).read_to_end(&mut buf).map_err(|e| e.to_string())?;
    Ok(buf)
}
