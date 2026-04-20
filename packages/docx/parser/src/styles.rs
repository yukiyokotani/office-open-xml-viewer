use std::collections::HashMap;
use roxmltree::Document as XmlDoc;
use crate::xml_util::*;

/// Resolved run (character) formatting.
#[derive(Debug, Clone, Default)]
pub struct RunFmt {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
    pub strikethrough: Option<bool>,
    pub font_size: Option<f64>,       // pt
    pub color: Option<String>,        // hex 6
    pub font_family_ascii: Option<String>,
    pub font_family_east_asia: Option<String>,
    pub background: Option<String>,   // hex 6
    /// "super" | "sub" — mapped from w:vertAlign val="superscript|subscript"
    pub vert_align: Option<String>,
    /// All caps (w:caps)
    pub all_caps: Option<bool>,
    /// Small capitals (w:smallCaps)
    pub small_caps: Option<bool>,
    /// Double strikethrough (w:dstrike)
    pub dstrike: Option<bool>,
    /// Hidden text — run should not be rendered (w:vanish / w:webHidden)
    pub vanish: Option<bool>,
    /// Highlight color name: "yellow" | "cyan" | "green" | ... (w:highlight)
    pub highlight: Option<String>,
}

/// Resolved paragraph formatting.
#[derive(Debug, Clone, Default)]
pub struct ParaFmt {
    pub alignment: Option<String>,
    pub indent_left: Option<f64>,     // pt
    pub indent_right: Option<f64>,    // pt
    pub indent_first: Option<f64>,    // pt
    pub space_before: Option<f64>,    // pt
    pub space_after: Option<f64>,     // pt
    pub line_spacing_val: Option<f64>,
    pub line_spacing_rule: Option<String>,
    pub num_id: Option<u32>,
    pub num_level: Option<u32>,
    /// Explicit tab stops (pos_pt, alignment, leader). None = inherit from parent style chain.
    pub tab_stops: Option<Vec<(f64, String, String)>>,
    /// merged run defaults from pPr/rPr
    pub run: RunFmt,
    pub based_on: Option<String>,
    /// Paragraph background hex color (w:shd fill on paragraph)
    pub shading: Option<String>,
    /// Force page break before paragraph (w:pageBreakBefore)
    pub page_break_before: Option<bool>,
    /// Suppress spacing between adjacent same-style paragraphs (w:contextualSpacing)
    pub contextual_spacing: Option<bool>,
    /// Keep paragraph on same page as the next paragraph (w:keepNext)
    pub keep_next: Option<bool>,
    /// Keep all lines of this paragraph on the same page (w:keepLines)
    pub keep_lines: Option<bool>,
    /// Widow/orphan control (w:widowControl). Default per spec: true.
    pub widow_control: Option<bool>,
    /// Paragraph border edges (w:pBdr)
    pub para_borders: Option<crate::types::ParagraphBorders>,
}

#[derive(Debug, Default)]
pub struct StyleDef {
    pub para: ParaFmt,
    pub run: RunFmt,
    pub based_on: Option<String>,
}

pub struct StyleMap {
    styles: HashMap<String, StyleDef>,
    defaults_para: ParaFmt,
    defaults_run: RunFmt,
    /// styleId of the style with w:default="1" and w:type="paragraph".
    /// Applied to paragraphs that have no explicit pStyle.
    default_para_style_id: Option<String>,
}

impl StyleMap {
    /// Style ID of the paragraph style marked `w:default="1"` in styles.xml.
    /// International templates may use non-English IDs (e.g. "a", "標準").
    pub fn default_para_style_id(&self) -> Option<&str> {
        self.default_para_style_id.as_deref()
    }

    pub fn parse(xml: &str) -> Self {
        let doc = match XmlDoc::parse(xml) {
            Ok(d) => d,
            Err(_) => return Self::empty(),
        };
        let root = doc.root_element();
        let mut styles: HashMap<String, StyleDef> = HashMap::new();
        let mut defaults_para = ParaFmt::default();
        let mut defaults_run = RunFmt::default();

        // Parse docDefaults
        if let Some(dd) = child_w(root, "docDefaults") {
            if let Some(rpr_def) = child_w(dd, "rPrDefault").and_then(|n| child_w(n, "rPr")) {
                defaults_run = parse_run_fmt(rpr_def);
            }
            if let Some(ppr_def) = child_w(dd, "pPrDefault").and_then(|n| child_w(n, "pPr")) {
                defaults_para = parse_para_fmt(ppr_def);
            }
        }

        // Parse each style
        let mut default_para_style_id: Option<String> = None;
        for style_node in children_w(root, "style") {
            let Some(style_id) = attr_w(style_node, "styleId") else { continue };
            let style_type = attr_w(style_node, "type").unwrap_or_default();
            if style_type != "paragraph" && style_type != "character" { continue; }

            if style_type == "paragraph"
                && attr_w(style_node, "default").as_deref() == Some("1")
            {
                default_para_style_id = Some(style_id.clone());
            }

            let based_on = child_w(style_node, "basedOn").and_then(|n| attr_w(n, "val"));

            let para = if let Some(ppr) = child_w(style_node, "pPr") {
                parse_para_fmt(ppr)
            } else {
                ParaFmt::default()
            };

            let run = if let Some(rpr) = child_w(style_node, "rPr") {
                parse_run_fmt(rpr)
            } else {
                RunFmt::default()
            };

            styles.insert(style_id, StyleDef { para, run, based_on });
        }

        StyleMap { styles, defaults_para, defaults_run, default_para_style_id }
    }

    fn empty() -> Self {
        StyleMap {
            styles: HashMap::new(),
            defaults_para: ParaFmt::default(),
            defaults_run: RunFmt::default(),
            default_para_style_id: None,
        }
    }

    /// Resolve all formatting for a paragraph style ID, merging inherited chain.
    /// Priority (lowest to highest): docDefaults → basedOn chain → style itself.
    /// Within each level: style rPr then pPr/rPr (both are paragraph-level run defaults).
    pub fn resolve_para(&self, style_id: Option<&str>) -> (ParaFmt, RunFmt) {
        let mut merged_para = ParaFmt::default();
        let mut merged_run = RunFmt::default();

        apply_para(&mut merged_para, &self.defaults_para);
        apply_run(&mut merged_run, &self.defaults_run);

        // Use explicit pStyle if present, otherwise fall back to the
        // paragraph style marked w:default="1" (typically "Normal").
        let effective_id = style_id
            .map(str::to_string)
            .or_else(|| self.default_para_style_id.clone());
        if let Some(id) = effective_id.as_deref() {
            self.apply_style_chain(id, &mut merged_para, &mut merged_run);
        }

        (merged_para, merged_run)
    }

    fn apply_style_chain(&self, id: &str, merged_para: &mut ParaFmt, merged_run: &mut RunFmt) {
        if let Some(def) = self.styles.get(id) {
            if let Some(base) = def.based_on.clone() {
                self.apply_style_chain(&base, merged_para, merged_run);
            }
            apply_para(merged_para, &def.para);
            apply_run(merged_run, &def.run);
            // pPr/rPr (paragraph mark run properties) also apply to runs
            apply_run(merged_run, &def.para.run);
        }
    }
}

fn apply_para(dst: &mut ParaFmt, src: &ParaFmt) {
    if src.alignment.is_some() { dst.alignment = src.alignment.clone(); }
    if src.indent_left.is_some() { dst.indent_left = src.indent_left; }
    if src.indent_right.is_some() { dst.indent_right = src.indent_right; }
    if src.indent_first.is_some() { dst.indent_first = src.indent_first; }
    if src.space_before.is_some() { dst.space_before = src.space_before; }
    if src.space_after.is_some() { dst.space_after = src.space_after; }
    if src.line_spacing_val.is_some() { dst.line_spacing_val = src.line_spacing_val; }
    if src.line_spacing_rule.is_some() { dst.line_spacing_rule = src.line_spacing_rule.clone(); }
    if src.num_id.is_some() { dst.num_id = src.num_id; }
    if src.num_level.is_some() { dst.num_level = src.num_level; }
    if src.tab_stops.is_some() { dst.tab_stops = src.tab_stops.clone(); }
    if src.shading.is_some() { dst.shading = src.shading.clone(); }
    if src.page_break_before.is_some() { dst.page_break_before = src.page_break_before; }
    if src.contextual_spacing.is_some() { dst.contextual_spacing = src.contextual_spacing; }
    if src.keep_next.is_some() { dst.keep_next = src.keep_next; }
    if src.keep_lines.is_some() { dst.keep_lines = src.keep_lines; }
    if src.widow_control.is_some() { dst.widow_control = src.widow_control; }
    if src.para_borders.is_some() { dst.para_borders = src.para_borders.clone(); }
}

fn apply_run(dst: &mut RunFmt, src: &RunFmt) {
    if src.bold.is_some() { dst.bold = src.bold; }
    if src.italic.is_some() { dst.italic = src.italic; }
    if src.underline.is_some() { dst.underline = src.underline; }
    if src.strikethrough.is_some() { dst.strikethrough = src.strikethrough; }
    if src.font_size.is_some() { dst.font_size = src.font_size; }
    if src.color.is_some() { dst.color = src.color.clone(); }
    if src.font_family_ascii.is_some() { dst.font_family_ascii = src.font_family_ascii.clone(); }
    if src.font_family_east_asia.is_some() { dst.font_family_east_asia = src.font_family_east_asia.clone(); }
    if src.background.is_some() { dst.background = src.background.clone(); }
    if src.vert_align.is_some() { dst.vert_align = src.vert_align.clone(); }
    if src.all_caps.is_some() { dst.all_caps = src.all_caps; }
    if src.small_caps.is_some() { dst.small_caps = src.small_caps; }
    if src.dstrike.is_some() { dst.dstrike = src.dstrike; }
    if src.vanish.is_some() { dst.vanish = src.vanish; }
    if src.highlight.is_some() { dst.highlight = src.highlight.clone(); }
}

pub fn parse_para_fmt(ppr: roxmltree::Node) -> ParaFmt {
    let mut fmt = ParaFmt::default();

    // Alignment
    if let Some(jc) = child_w(ppr, "jc") {
        fmt.alignment = attr_w(jc, "val");
    }

    // Spacing
    if let Some(sp) = child_w(ppr, "spacing") {
        if let Some(v) = attr_w(sp, "before") { fmt.space_before = Some(twips_to_pt(&v)); }
        if let Some(v) = attr_w(sp, "after") { fmt.space_after = Some(twips_to_pt(&v)); }
        if let Some(v) = attr_w(sp, "line") {
            let rule = attr_w(sp, "lineRule").unwrap_or_else(|| "auto".to_string());
            let raw: f64 = v.parse().unwrap_or(240.0);
            // OOXML encodes line spacing as:
            //   auto      → raw / 240   = multiplier (1.0 = single, 1.5 = 1½, 2.0 = double)
            //   atLeast   → raw / 20    = pt (minimum line height)
            //   exact     → raw / 20    = pt (exact line height)
            // Word tolerates "auto" with very large raw values and treats them as
            // atLeast-twips in practice (seen in the wild for decorative titles).
            // Reinterpret raw > 720 (3x single) in the auto rule as atLeast pt so
            // large fonts render near their natural height instead of 3×+.
            let (val, effective_rule) = match rule.as_str() {
                "exact" => (raw / 20.0, "exact".to_string()),
                "atLeast" => (raw / 20.0, "atLeast".to_string()),
                _ => {
                    if raw > 720.0 {
                        (raw / 20.0, "atLeast".to_string())
                    } else {
                        (raw / 240.0, "auto".to_string())
                    }
                }
            };
            fmt.line_spacing_val = Some(val);
            fmt.line_spacing_rule = Some(effective_rule);
        }
    }

    // Indentation. ECMA-376 §17.3.1.12 allows both the older "left"/"right"
    // attributes and the logical "start"/"end" aliases. In LTR docs these are
    // identical; use either if present, with start/end taking precedence when
    // both appear (logical wins for bidi correctness).
    if let Some(ind) = child_w(ppr, "ind") {
        if let Some(v) = attr_w(ind, "left")  { fmt.indent_left  = Some(twips_to_pt(&v)); }
        if let Some(v) = attr_w(ind, "start") { fmt.indent_left  = Some(twips_to_pt(&v)); }
        if let Some(v) = attr_w(ind, "right") { fmt.indent_right = Some(twips_to_pt(&v)); }
        if let Some(v) = attr_w(ind, "end")   { fmt.indent_right = Some(twips_to_pt(&v)); }
        if let Some(v) = attr_w(ind, "firstLine") { fmt.indent_first = Some(twips_to_pt(&v)); }
        // hanging overrides firstLine per §17.3.1.12 when both are present.
        if let Some(v) = attr_w(ind, "hanging")   { fmt.indent_first = Some(-twips_to_pt(&v)); }
    }

    // Numbering
    if let Some(pnpr) = child_w(ppr, "numPr") {
        // ilvl defaults to 0 when absent
        fmt.num_level = child_w(pnpr, "ilvl")
            .and_then(|n| attr_w(n, "val"))
            .and_then(|v| v.parse().ok())
            .or(Some(0));
        if let Some(nid) = child_w(pnpr, "numId") {
            fmt.num_id = attr_w(nid, "val").and_then(|v| v.parse().ok());
        }
    }

    // Explicit tab stops (pPr/tabs/tab)
    if let Some(tabs_node) = child_w(ppr, "tabs") {
        let mut tabs: Vec<(f64, String, String)> = Vec::new();
        for t in children_w(tabs_node, "tab") {
            let val = attr_w(t, "val").unwrap_or_else(|| "left".to_string());
            // val="clear" removes an inherited tab — MVP: skip (no tab to emit)
            if val == "clear" { continue; }
            let pos = match attr_w(t, "pos").map(|s| twips_to_pt(&s)) {
                Some(p) => p,
                None => continue,
            };
            let leader = attr_w(t, "leader").unwrap_or_else(|| "none".to_string());
            tabs.push((pos, val, leader));
        }
        if !tabs.is_empty() {
            tabs.sort_by(|a, b| a.0.partial_cmp(&b.0).unwrap_or(std::cmp::Ordering::Equal));
            fmt.tab_stops = Some(tabs);
        }
    }

    // pPr/rPr (run defaults within paragraph)
    if let Some(rpr) = child_w(ppr, "rPr") {
        fmt.run = parse_run_fmt(rpr);
    }

    // Paragraph shading
    if let Some(shd) = child_w(ppr, "shd") {
        if let Some(fill) = attr_w(shd, "fill") {
            if fill != "auto" && fill.len() == 6 {
                fmt.shading = Some(fill.to_lowercase());
            }
        }
    }

    // Page break before paragraph
    fmt.page_break_before = bool_prop(ppr, "pageBreakBefore");

    // Contextual spacing
    fmt.contextual_spacing = bool_prop(ppr, "contextualSpacing");

    // keepNext — keep this paragraph on the same page as the next one
    fmt.keep_next = bool_prop(ppr, "keepNext");

    // keepLines — do not split this paragraph's lines across pages
    fmt.keep_lines = bool_prop(ppr, "keepLines");

    // widowControl — avoid leaving a single line at page top/bottom
    // (ECMA-376 default: true; explicit value=0 disables).
    fmt.widow_control = bool_prop(ppr, "widowControl");

    // Paragraph borders (pBdr)
    if let Some(pbdr) = child_w(ppr, "pBdr") {
        use crate::types::{ParagraphBorders, ParaBorderEdge};
        let parse_edge = |name: &str| -> Option<ParaBorderEdge> {
            let node = child_w(pbdr, name)?;
            let style = attr_w(node, "val").unwrap_or_else(|| "none".to_string());
            if style == "none" || style == "nil" { return None; }
            let width = attr_w(node, "sz")
                .and_then(|s| s.parse::<f64>().ok())
                .map(|v| v / 8.0)
                .unwrap_or(0.5);
            let space = attr_w(node, "space")
                .and_then(|s| s.parse::<f64>().ok())
                .unwrap_or(1.0);
            let color = attr_w(node, "color")
                .filter(|c| c != "auto")
                .map(|c| c.to_lowercase());
            Some(ParaBorderEdge { style, color, width, space })
        };
        let borders = ParagraphBorders {
            top: parse_edge("top"),
            bottom: parse_edge("bottom"),
            left: parse_edge("left"),
            right: parse_edge("right"),
            between: parse_edge("between"),
        };
        if borders.top.is_some() || borders.bottom.is_some() || borders.left.is_some() || borders.right.is_some() {
            fmt.para_borders = Some(borders);
        }
    }

    fmt
}

pub fn parse_run_fmt(rpr: roxmltree::Node) -> RunFmt {
    let mut fmt = RunFmt::default();

    fmt.bold = bool_prop(rpr, "b");
    fmt.italic = bool_prop(rpr, "i");
    fmt.strikethrough = bool_prop(rpr, "strike");

    // Underline
    if let Some(u) = child_w(rpr, "u") {
        let val = attr_w(u, "val").unwrap_or_else(|| "single".to_string());
        fmt.underline = Some(val != "none");
    }

    // Font size — w:sz is used for Latin and East Asian (CJK) text.
    // w:szCs is for complex scripts (Arabic/Hebrew RTL text) only; fall back to it when sz is absent.
    if let Some(sz) = child_w(rpr, "sz").or_else(|| child_w(rpr, "szCs")) {
        if let Some(v) = attr_w(sz, "val") {
            fmt.font_size = Some(half_pt_to_pt(&v));
        }
    }

    // Color
    if let Some(col) = child_w(rpr, "color") {
        let val = attr_w(col, "val").unwrap_or_default();
        if val != "auto" && !val.is_empty() {
            fmt.color = Some(val.to_lowercase());
        }
    }

    // Font family. ECMA-376 §17.3.2.26 rFonts supports both direct typeface
    // attributes (ascii/hAnsi/eastAsia/cs) and theme references (asciiTheme,
    // hAnsiTheme, eastAsiaTheme, cstheme). Theme refs are resolved post-parse
    // in parse_document once a Theme is available; here we just record the
    // reference string under the corresponding axis. Direct attributes take
    // precedence over theme refs per spec.
    if let Some(rf) = child_w(rpr, "rFonts") {
        let direct_ascii = attr_w(rf, "ascii").or_else(|| attr_w(rf, "hAnsi"));
        let theme_ascii  = attr_w(rf, "asciiTheme").or_else(|| attr_w(rf, "hAnsiTheme"));
        fmt.font_family_ascii = direct_ascii.or_else(|| theme_ascii.map(|t| format!("@theme:{t}")));

        let direct_ea = attr_w(rf, "eastAsia");
        let theme_ea  = attr_w(rf, "eastAsiaTheme");
        fmt.font_family_east_asia = direct_ea.or_else(|| theme_ea.map(|t| format!("@theme:{t}")));
    }

    // Background highlight
    if let Some(shd) = child_w(rpr, "shd") {
        if let Some(fill) = attr_w(shd, "fill") {
            if fill != "auto" && fill.len() == 6 {
                fmt.background = Some(fill.to_lowercase());
            }
        }
    }

    // Vertical alignment (superscript / subscript)
    if let Some(va) = child_w(rpr, "vertAlign") {
        if let Some(val) = attr_w(va, "val") {
            fmt.vert_align = match val.as_str() {
                "superscript" => Some("super".to_string()),
                "subscript" => Some("sub".to_string()),
                _ => None,
            };
        }
    }

    // All caps / small caps
    fmt.all_caps = bool_prop(rpr, "caps");
    fmt.small_caps = bool_prop(rpr, "smallCaps");

    // Double strikethrough
    fmt.dstrike = bool_prop(rpr, "dstrike");

    // Hidden text (vanish or webHidden)
    fmt.vanish = bool_prop(rpr, "vanish").or_else(|| bool_prop(rpr, "webHidden"));

    // Highlight
    if let Some(hl) = child_w(rpr, "highlight") {
        fmt.highlight = attr_w(hl, "val").filter(|v| v != "none");
    }

    fmt
}
