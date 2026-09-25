use crate::parse_color;
use crate::read_zip_string;
use crate::types::*;
use ooxml_common::depth::parse_guarded;
use ooxml_common::ns::is_x_ns;

/// Resolve the workbook's Normal-style font (family, point size, and style) by
/// following `<cellStyleXfs>[0].fontId` → `<fonts>[fontId]`. Returns `(None,
/// None)` if `xl/styles.xml` is missing or malformed. The renderer uses this
/// to compute the Max Digit Width for column-width pixel conversion
/// (ECMA-376 §18.3.1.13).
pub(crate) type DefaultFont = (Option<String>, Option<f64>, bool, bool);

fn parse_default_font(doc: &roxmltree::Document) -> DefaultFont {
    let mut font_id: usize = 0;
    for n in doc.descendants() {
        if n.tag_name().name() == "cellStyleXfs" && is_x_ns(n.tag_name().namespace()) {
            if let Some(xf) = n
                .children()
                .find(|c| c.is_element() && c.tag_name().name() == "xf")
            {
                font_id = xf
                    .attribute("fontId")
                    .and_then(|s| s.parse().ok())
                    .unwrap_or(0);
            }
            break;
        }
    }
    for fonts_node in doc.descendants() {
        if fonts_node.tag_name().name() != "fonts" || !is_x_ns(fonts_node.tag_name().namespace()) {
            continue;
        }
        if let Some(font_node) = fonts_node
            .children()
            .filter(|c| c.is_element() && c.tag_name().name() == "font")
            .nth(font_id)
        {
            let mut name = None;
            let mut sz = None;
            let mut bold = false;
            let mut italic = false;
            for child in font_node.children() {
                match child.tag_name().name() {
                    "name" => name = child.attribute("val").map(|s| s.to_string()),
                    "sz" => sz = child.attribute("val").and_then(|s| s.parse().ok()),
                    "b" => bold = parse_st_on_off(&child),
                    "i" => italic = parse_st_on_off(&child),
                    _ => {}
                }
            }
            return (name, sz, bold, italic);
        }
        break;
    }
    (None, None, false, false)
}

pub(crate) struct ParsedStylesPart {
    pub(crate) styles: Styles,
    pub(crate) default_font: DefaultFont,
    pub(crate) chart_number_formats: crate::chart::ChartNumberFormatCache,
}

pub(crate) struct ParsedStyleProjection {
    pub(crate) default_font: DefaultFont,
    pub(crate) chart_number_formats: crate::chart::ChartNumberFormatCache,
}

pub(crate) fn parse_styles(
    archive: &mut crate::XlsxZip,
    theme_colors: &[String],
) -> Result<ParsedStylesPart, String> {
    let xml = read_zip_string(archive, "xl/styles.xml")?;
    let doc = parse_guarded(&xml).map_err(|e| e.to_string())?;

    let default_font = parse_default_font(&doc);
    let chart_number_formats = crate::chart::ChartNumberFormatCache::from_document(&doc);
    let num_fmts = parse_num_fmts(&doc);
    let fonts = parse_fonts(&doc, theme_colors);
    let fills = parse_fills(&doc, theme_colors);
    let borders = parse_borders(&doc, theme_colors);
    let cell_xfs = parse_cell_xfs(&doc);
    let dxfs = parse_dxfs(&doc, theme_colors);

    Ok(ParsedStylesPart {
        styles: Styles {
            fonts,
            fills,
            borders,
            cell_xfs,
            num_fmts,
            dxfs,
        },
        default_font,
        chart_number_formats,
    })
}

/// Parse only the workbook-wide style information required while materializing
/// individual sheets. This deliberately avoids constructing fills, borders,
/// DXFs, and other full-workbook output that a sheet projection never returns.
pub(crate) fn parse_style_projection(
    archive: &mut crate::XlsxZip,
) -> Result<ParsedStyleProjection, String> {
    let xml = read_zip_string(archive, "xl/styles.xml")?;
    let doc = parse_guarded(&xml).map_err(|e| e.to_string())?;
    Ok(ParsedStyleProjection {
        default_font: parse_default_font(&doc),
        chart_number_formats: crate::chart::ChartNumberFormatCache::from_document(&doc),
    })
}

pub(crate) fn parse_dxfs(doc: &roxmltree::Document, theme_colors: &[String]) -> Vec<Dxf> {
    doc.descendants()
        .find(|node| node.tag_name().name() == "dxfs" && is_x_ns(node.tag_name().namespace()))
        .map(|dxfs_node| {
            dxfs_node
                .children()
                .filter(|node| node.tag_name().name() == "dxf")
                .map(|dxf_node| parse_dxf(dxf_node, theme_colors))
                .collect()
        })
        .unwrap_or_default()
}

/// One `<dxf>` (ECMA-376 §18.8.14), read in its own document so every
/// namespace declaration in scope applies.
pub(crate) fn parse_dxf(dxf_node: roxmltree::Node, theme_colors: &[String]) -> Dxf {
    let mut d = Dxf::default();
    for child in dxf_node.children() {
        match child.tag_name().name() {
            "font" => {
                let mut f = Font {
                    size: 11.0,
                    ..Default::default()
                };
                let mut toggles = DxfFontToggles::default();
                for fc in child.children() {
                    match fc.tag_name().name() {
                        "b" => {
                            f.bold = parse_st_on_off(&fc);
                            toggles.bold = Some(f.bold);
                        }
                        "i" => {
                            f.italic = parse_st_on_off(&fc);
                            toggles.italic = Some(f.italic);
                        }
                        "u" => {
                            let v = fc.attribute("val").unwrap_or("single");
                            if v != "none" {
                                f.underline = true;
                                if v != "single" {
                                    f.underline_style = Some(v.to_string());
                                }
                            }
                            toggles.underline = Some(v != "none");
                        }
                        "strike" => {
                            f.strike = parse_st_on_off(&fc);
                            toggles.strike = Some(f.strike);
                        }
                        "vertAlign" => {
                            if let Some(v) = fc.attribute("val") {
                                if v != "baseline" {
                                    f.vert_align = Some(v.to_string());
                                }
                            }
                        }
                        "sz" => {
                            if let Some(v) = fc.attribute("val").and_then(|s| s.parse().ok()) {
                                f.size = v;
                            }
                        }
                        "name" => {
                            f.name = fc.attribute("val").map(|s| s.to_string());
                        }
                        "scheme" => {
                            f.scheme = fc
                                .attribute("val")
                                .filter(|value| matches!(*value, "major" | "minor"))
                                .map(str::to_owned);
                        }
                        "charset" => {
                            f.charset = fc
                                .attribute("val")
                                .and_then(|value| value.parse::<u8>().ok());
                        }
                        "color" => {
                            f.color = parse_color(&fc, theme_colors);
                        }
                        _ => {}
                    }
                }
                d.font = Some(f);
                d.font_toggles = Some(toggles);
            }
            "fill" => {
                let mut f = Fill::default();
                for pf in child.children() {
                    if pf.tag_name().name() == "gradientFill" {
                        // ECMA-376 §18.8.24 in a dxf, e.g. a table or
                        // PivotTable style element's gradient.
                        f.gradient = parse_gradient_fill(pf, theme_colors);
                    }
                    if pf.tag_name().name() == "patternFill" {
                        f.pattern_type = pf.attribute("patternType").unwrap_or("solid").to_string();
                        for color_node in pf.children() {
                            match color_node.tag_name().name() {
                                "fgColor" => f.fg_color = parse_color(&color_node, theme_colors),
                                "bgColor" => f.bg_color = parse_color(&color_node, theme_colors),
                                _ => {}
                            }
                        }
                    }
                }
                // In dxf, conditional format fills often only have bgColor; mirror into fgColor
                if f.fg_color.is_none() && f.bg_color.is_some() {
                    f.fg_color = f.bg_color.clone();
                }
                d.fill = Some(f);
            }
            "border" => {
                let mut b = Border::default();
                for edge_node in child.children() {
                    let style = edge_node.attribute("style").unwrap_or("").to_string();
                    if style.is_empty() {
                        continue;
                    }
                    let color = edge_node
                        .children()
                        .find(|c| c.is_element())
                        .and_then(|c| parse_color(&c, theme_colors));
                    let edge = Some(BorderEdge { style, color });
                    match edge_node.tag_name().name() {
                        "left" => b.left = edge,
                        "right" => b.right = edge,
                        "top" => b.top = edge,
                        "bottom" => b.bottom = edge,
                        "horizontal" => b.horizontal = edge,
                        "vertical" => b.vertical = edge,
                        _ => {}
                    }
                }
                d.border = Some(b);
            }
            "numFmt" => {
                let num_fmt_id = child
                    .attribute("numFmtId")
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0);
                let format_code = child.attribute("formatCode").unwrap_or("").to_string();
                d.num_fmt = Some(NumFmt {
                    num_fmt_id,
                    format_code,
                });
            }
            _ => {}
        }
    }
    d
}

pub(crate) fn parse_num_fmts(doc: &roxmltree::Document) -> Vec<NumFmt> {
    let mut fmts = Vec::new();
    for node in doc.descendants() {
        if node.tag_name().name() == "numFmts" && is_x_ns(node.tag_name().namespace()) {
            for child in node.children() {
                if child.tag_name().name() != "numFmt" {
                    continue;
                }
                let num_fmt_id = child
                    .attribute("numFmtId")
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0);
                let format_code = child.attribute("formatCode").unwrap_or("").to_string();
                fmts.push(NumFmt {
                    num_fmt_id,
                    format_code,
                });
            }
            break;
        }
    }
    fmts
}

/// ECMA-376 §22.9.2 ST_OnOff. Toggle elements like `<b/>`, `<i/>`, `<strike/>`
/// accept an optional `val` attribute whose value is "1" / "true" (on) or
/// "0" / "false" (off). When omitted, the presence of the element itself
/// implies "on". A dxf with `<i val="0"/>` therefore means "differential
/// format that *clears* italic", not "set italic to true".
pub(crate) fn parse_st_on_off(node: &roxmltree::Node) -> bool {
    match node.attribute("val") {
        None => true,
        Some(v) => !matches!(v, "0" | "false" | "False" | "FALSE" | "off" | "Off"),
    }
}

pub(crate) fn parse_fonts(doc: &roxmltree::Document, theme_colors: &[String]) -> Vec<Font> {
    let mut fonts = Vec::new();
    for fonts_node in doc.descendants() {
        if fonts_node.tag_name().name() == "fonts" && is_x_ns(fonts_node.tag_name().namespace()) {
            for font_node in fonts_node.children() {
                if font_node.tag_name().name() != "font" {
                    continue;
                }
                let mut f = Font {
                    size: 11.0,
                    ..Default::default()
                };
                for child in font_node.children() {
                    match child.tag_name().name() {
                        "b" => f.bold = parse_st_on_off(&child),
                        "i" => f.italic = parse_st_on_off(&child),
                        "u" => {
                            let v = child.attribute("val").unwrap_or("single");
                            if v != "none" {
                                f.underline = true;
                                if v != "single" {
                                    f.underline_style = Some(v.to_string());
                                }
                            }
                        }
                        "strike" => f.strike = parse_st_on_off(&child),
                        "vertAlign" => {
                            if let Some(v) = child.attribute("val") {
                                if v != "baseline" {
                                    f.vert_align = Some(v.to_string());
                                }
                            }
                        }
                        "sz" => {
                            if let Some(v) = child.attribute("val").and_then(|s| s.parse().ok()) {
                                f.size = v;
                            }
                        }
                        "name" => {
                            f.name = child.attribute("val").map(|s| s.to_string());
                        }
                        "scheme" => {
                            f.scheme = child
                                .attribute("val")
                                .filter(|value| matches!(*value, "major" | "minor"))
                                .map(str::to_owned);
                        }
                        "charset" => {
                            f.charset = child
                                .attribute("val")
                                .and_then(|value| value.parse::<u8>().ok());
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

pub(crate) fn parse_fills(doc: &roxmltree::Document, theme_colors: &[String]) -> Vec<Fill> {
    let mut fills = Vec::new();
    for fills_node in doc.descendants() {
        if fills_node.tag_name().name() == "fills" && is_x_ns(fills_node.tag_name().namespace()) {
            for fill_node in fills_node.children() {
                if fill_node.tag_name().name() != "fill" {
                    continue;
                }
                let mut f = Fill::default();
                for pf in fill_node.children() {
                    match pf.tag_name().name() {
                        "patternFill" => {
                            f.pattern_type =
                                pf.attribute("patternType").unwrap_or("none").to_string();
                            for color_node in pf.children() {
                                match color_node.tag_name().name() {
                                    "fgColor" => {
                                        f.fg_color = parse_color(&color_node, theme_colors)
                                    }
                                    "bgColor" => {
                                        f.bg_color = parse_color(&color_node, theme_colors)
                                    }
                                    _ => {}
                                }
                            }
                        }
                        "gradientFill" => f.gradient = parse_gradient_fill(pf, theme_colors),
                        _ => {}
                    }
                }
                fills.push(f);
            }
            break;
        }
    }
    fills
}

pub(crate) fn parse_borders(doc: &roxmltree::Document, theme_colors: &[String]) -> Vec<Border> {
    let mut borders = Vec::new();
    for borders_node in doc.descendants() {
        if borders_node.tag_name().name() == "borders"
            && is_x_ns(borders_node.tag_name().namespace())
        {
            for border_node in borders_node.children() {
                if border_node.tag_name().name() != "border" {
                    continue;
                }
                let has_diag_up = border_node
                    .attribute("diagonalUp")
                    .map(|v| v == "1" || v == "true")
                    .unwrap_or(false);
                let has_diag_down = border_node
                    .attribute("diagonalDown")
                    .map(|v| v == "1" || v == "true")
                    .unwrap_or(false);
                let mut b = Border::default();
                let mut diag_edge: Option<BorderEdge> = None;
                for edge_node in border_node.children() {
                    let style = edge_node.attribute("style").unwrap_or("").to_string();
                    let color = edge_node
                        .children()
                        .find(|c| c.is_element())
                        .and_then(|c| parse_color(&c, theme_colors));
                    match edge_node.tag_name().name() {
                        "left" if !style.is_empty() => b.left = Some(BorderEdge { style, color }),
                        "right" if !style.is_empty() => b.right = Some(BorderEdge { style, color }),
                        "top" if !style.is_empty() => b.top = Some(BorderEdge { style, color }),
                        "bottom" if !style.is_empty() => {
                            b.bottom = Some(BorderEdge { style, color })
                        }
                        "diagonal" if !style.is_empty() => {
                            diag_edge = Some(BorderEdge { style, color })
                        }
                        _ => {}
                    }
                }
                if has_diag_up {
                    b.diagonal_up = diag_edge.clone();
                }
                if has_diag_down {
                    b.diagonal_down = diag_edge;
                }
                borders.push(b);
            }
            break;
        }
    }
    borders
}

pub(crate) fn parse_cell_xfs(doc: &roxmltree::Document) -> Vec<CellXf> {
    let mut xfs = Vec::new();
    for xfs_node in doc.descendants() {
        if xfs_node.tag_name().name() == "cellXfs" && is_x_ns(xfs_node.tag_name().namespace()) {
            for xf_node in xfs_node.children() {
                if xf_node.tag_name().name() != "xf" {
                    continue;
                }
                let font_id = xf_node
                    .attribute("fontId")
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0);
                let fill_id = xf_node
                    .attribute("fillId")
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0);
                let border_id = xf_node
                    .attribute("borderId")
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0);
                let num_fmt_id = xf_node
                    .attribute("numFmtId")
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0);
                let mut align_h = None;
                let mut align_v = None;
                let mut wrap_text = false;
                let mut indent = None;
                let mut text_rotation = None;
                let mut shrink_to_fit = false;
                let mut reading_order: Option<u32> = None;
                for child in xf_node.children() {
                    if child.tag_name().name() == "alignment" {
                        // ECMA-376 §18.8.1: horizontal defaults to `general`,
                        // so an explicit `general` is the same as omission and
                        // leaves the value-type rule (§18.18.40) to the renderer.
                        align_h = child
                            .attribute("horizontal")
                            .filter(|value| *value != "general")
                            .map(|s| s.to_string());
                        align_v = child.attribute("vertical").map(|s| s.to_string());
                        wrap_text = child
                            .attribute("wrapText")
                            .map(|v| v == "1" || v == "true")
                            .unwrap_or(false);
                        indent = child
                            .attribute("indent")
                            .and_then(|s| s.parse::<u32>().ok())
                            .filter(|&v| v > 0);
                        text_rotation = child
                            .attribute("textRotation")
                            .and_then(|s| s.parse::<u32>().ok())
                            .filter(|&v| v > 0);
                        shrink_to_fit = child
                            .attribute("shrinkToFit")
                            .map(|v| v == "1" || v == "true")
                            .unwrap_or(false);
                        reading_order = child
                            .attribute("readingOrder")
                            .and_then(|s| s.parse::<u32>().ok())
                            .filter(|&v| v > 0);
                    }
                }
                xfs.push(CellXf {
                    font_id,
                    fill_id,
                    border_id,
                    num_fmt_id,
                    align_h,
                    align_v,
                    wrap_text,
                    indent,
                    text_rotation,
                    shrink_to_fit,
                    reading_order,
                });
            }
            break;
        }
    }
    xfs
}

/// ISO/IEC 29500 Strict-conformance fixture (`fix(xlsx): accept Strict
/// namespace URIs across the parser` routed every `x:` element match here
/// through `is_x_ns`). Before that conversion `<fonts>`/`<cellXfs>` were
/// found via a hardcoded Transitional URI, so a Strict `xl/styles.xml` —
/// `xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"` — resolved to
/// empty `fonts/cell_xfs` vectors; this pins that the style *references* a
/// worksheet cell's `s="N"` index into now resolve identically to the
/// Transitional case.
#[cfg(test)]
mod strict_namespace_tests {
    use super::*;

    const X_NS_STRICT: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";

    fn theme() -> Vec<String> {
        vec!["#111111".into(); 12]
    }

    #[test]
    fn cell_font_retains_theme_scheme_and_charset_without_overriding_name() {
        let xml = format!(
            r#"<styleSheet xmlns="{X_NS_STRICT}"><fonts count="2">
          <font><name val="Calibri"/><scheme val="minor"/><charset val="128"/></font>
          <font><name val="Arial"/></font>
        </fonts></styleSheet>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let fonts = parse_fonts(&doc, &theme());
        assert_eq!(fonts[0].name.as_deref(), Some("Calibri"));
        assert_eq!(fonts[0].scheme.as_deref(), Some("minor"));
        assert_eq!(fonts[0].charset, Some(128));
        assert_eq!(fonts[1].name.as_deref(), Some("Arial"));
        assert_eq!(fonts[1].scheme, None);
        assert_eq!(fonts[1].charset, None);
    }

    #[test]
    fn strict_styles_xml_resolves_fonts_fills_and_cell_xfs() {
        let xml = format!(
            r#"<styleSheet xmlns="{ns}">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="14"/><name val="Calibri"/><color rgb="FFFF0000"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf fontId="0" fillId="0" borderId="0"/>
    <xf fontId="1" fillId="1" borderId="0" applyFont="1" applyFill="1">
      <alignment horizontal="center" wrapText="1"/>
    </xf>
  </cellXfs>
</styleSheet>"#,
            ns = X_NS_STRICT,
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();

        let fonts = parse_fonts(&doc, &theme());
        assert_eq!(fonts.len(), 2, "Strict <fonts> must be found via is_x_ns");
        assert!(fonts[1].bold);
        assert_eq!(fonts[1].size, 14.0);
        assert_eq!(fonts[1].name.as_deref(), Some("Calibri"));
        assert_eq!(fonts[1].color.as_deref(), Some("#FF0000"));

        let fills = parse_fills(&doc, &theme());
        assert_eq!(fills.len(), 2, "Strict <fills> must be found via is_x_ns");
        assert_eq!(fills[1].pattern_type, "solid");
        assert_eq!(fills[1].fg_color.as_deref(), Some("#FFFF00"));

        let cell_xfs = parse_cell_xfs(&doc);
        assert_eq!(
            cell_xfs.len(),
            2,
            "Strict <cellXfs> must be found via is_x_ns"
        );
        // A worksheet cell's `s="1"` references cell_xfs[1] — the style
        // reference a real Strict document round-trips through.
        let styled = &cell_xfs[1];
        assert_eq!(styled.font_id, 1);
        assert_eq!(styled.fill_id, 1);
        assert_eq!(styled.align_h.as_deref(), Some("center"));
        assert!(styled.wrap_text);
    }

    #[test]
    fn normal_font_style_follows_cell_style_font_id() {
        let xml = r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fonts count="2">
            <font><sz val="11"/><name val="Calibri"/></font>
            <font><b/><i/><sz val="12"/><name val="Arial"/></font>
          </fonts>
          <cellStyleXfs count="1"><xf fontId="1"/></cellStyleXfs>
        </styleSheet>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        assert_eq!(
            parse_default_font(&doc),
            (Some("Arial".into()), Some(12.0), true, true)
        );
    }

    #[test]
    fn explicit_general_horizontal_alignment_is_the_default() {
        let xml = r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment horizontal="general"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment horizontal="left"/></xf>
  </cellXfs>
</styleSheet>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let cell_xfs = parse_cell_xfs(&doc);
        assert_eq!(cell_xfs[0].align_h, None);
        assert_eq!(cell_xfs[1].align_h.as_deref(), Some("left"));
    }
}

/// ECMA-376 §18.8.24 gradientFill as the model's gradient: linear (default)
/// uses `degree`, path uses top/bottom/left/right as a relative bounding box;
/// children <stop position="n"><color/></stop>. `None` without stops.
fn parse_gradient_fill(pf: roxmltree::Node, theme_colors: &[String]) -> Option<GradientFillSpec> {
    // ECMA-376 §18.8.24 gradientFill — linear (default) uses
    // `degree`, path uses top/bottom/left/right as a relative
    // bounding box; children <stop position="n"><color/></stop>.
    let gtype = pf.attribute("type").unwrap_or("linear").to_string();
    let degree = pf
        .attribute("degree")
        .and_then(|s| s.parse::<f64>().ok())
        .unwrap_or(0.0);
    let left = pf
        .attribute("left")
        .and_then(|s| s.parse::<f64>().ok())
        .unwrap_or(0.0);
    let right = pf
        .attribute("right")
        .and_then(|s| s.parse::<f64>().ok())
        .unwrap_or(0.0);
    let top = pf
        .attribute("top")
        .and_then(|s| s.parse::<f64>().ok())
        .unwrap_or(0.0);
    let bottom = pf
        .attribute("bottom")
        .and_then(|s| s.parse::<f64>().ok())
        .unwrap_or(0.0);
    let mut stops: Vec<GradientStopSpec> = pf
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "stop")
        .filter_map(|stop| {
            let position = stop
                .attribute("position")
                .and_then(|s| s.parse::<f64>().ok())?;
            let color_node = stop
                .children()
                .find(|c| c.is_element() && c.tag_name().name() == "color")?;
            let color = parse_color(&color_node, theme_colors)?;
            Some(GradientStopSpec { position, color })
        })
        .collect();
    stops.sort_by(|a, b| {
        a.position
            .partial_cmp(&b.position)
            .unwrap_or(std::cmp::Ordering::Equal)
    });
    (!stops.is_empty()).then_some(GradientFillSpec {
        gradient_type: gtype,
        degree,
        left,
        right,
        top,
        bottom,
        stops,
    })
}

#[cfg(test)]
mod dxf_gradient_tests {
    #[test]
    fn dxf_gradient_fills_are_parsed_like_cell_gradients() {
        let xml = r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dxfs count="1"><dxf><font><color rgb="FFFFFFFF"/></font><fill><gradientFill degree="90"><stop position="0"><color rgb="FF9F2121"/></stop><stop position="1"><color rgb="FF761818"/></stop></gradientFill></fill></dxf></dxfs></styleSheet>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let dxf = super::parse_dxfs(&doc, &[]).remove(0);
        let gradient = dxf.fill.unwrap().gradient.unwrap();
        assert_eq!(gradient.degree, 90.0);
        assert_eq!(gradient.stops.len(), 2);
        assert_eq!(gradient.stops[0].color, "#9F2121");
    }
}
