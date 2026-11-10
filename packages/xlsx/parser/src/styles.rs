use crate::parse_color;
use crate::read_zip_string;
use crate::types::*;
use ooxml_common::depth::parse_guarded;
use ooxml_common::ns::is_x_ns;

/// Resolve the workbook's default font (family, point size, and style): the
/// first `<fonts>` entry. Returns `(None, None)` if `xl/styles.xml` is missing
/// or malformed. The renderer uses this to compute the Max Digit Width for
/// column-width pixel conversion (ECMA-376 §18.3.1.13).
///
/// Measured with Excel-exported control workbooks: column widths follow
/// `<fonts>[0]` alone. Moving the Normal cell style (`builtinId="0"`) to a
/// larger font, or giving `<cellStyleXfs>[0]` one, leaves the columns
/// unchanged, while a larger `<fonts>[0]` widens them.
pub(crate) type DefaultFont = (Option<String>, Option<f64>, bool, bool);

fn parse_default_font(doc: &roxmltree::Document) -> DefaultFont {
    for fonts_node in doc.descendants() {
        if fonts_node.tag_name().name() != "fonts" || !is_x_ns(fonts_node.tag_name().namespace()) {
            continue;
        }
        if let Some(font_node) = fonts_node
            .children()
            .find(|c| c.is_element() && c.tag_name().name() == "font")
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

/// Point size of the Normal cell style's font: `<cellStyle builtinId="0">`'s
/// `xfId` → `<cellStyleXfs>` → `<fonts>` (ECMA-376 §18.8.7), falling back to
/// `<cellStyleXfs>[0]` when no Normal style is declared. The renderer uses it
/// as the automatic row-height baseline (a font whose line box exceeds the
/// Normal font's may grow a row). Column widths use the default font instead
/// (`parse_default_font`).
fn parse_normal_font_size(doc: &roxmltree::Document) -> Option<f64> {
    let xf_index = normal_style_xf(doc).unwrap_or(0);
    let font_id: usize = doc
        .descendants()
        .find(|n| n.tag_name().name() == "cellStyleXfs" && is_x_ns(n.tag_name().namespace()))?
        .children()
        .filter(|c| c.is_element() && c.tag_name().name() == "xf")
        .nth(xf_index)?
        .attribute("fontId")
        .and_then(|v| v.parse().ok())
        .unwrap_or(0);
    doc.descendants()
        .find(|n| n.tag_name().name() == "fonts" && is_x_ns(n.tag_name().namespace()))?
        .children()
        .filter(|c| c.is_element() && c.tag_name().name() == "font")
        .nth(font_id)?
        .children()
        .find(|c| c.tag_name().name() == "sz")?
        .attribute("val")?
        .parse()
        .ok()
}

pub(crate) struct ParsedStylesPart {
    pub(crate) styles: Styles,
    pub(crate) default_font: DefaultFont,
    pub(crate) normal_font_size: Option<f64>,
    pub(crate) normal_font_color: Option<Option<String>>,
    pub(crate) chart_number_formats: crate::chart::ChartNumberFormatCache,
}

pub(crate) struct ParsedStyleProjection {
    pub(crate) default_font: DefaultFont,
    pub(crate) normal_font_size: Option<f64>,
    pub(crate) normal_font_color: Option<Option<String>>,
    pub(crate) chart_number_formats: crate::chart::ChartNumberFormatCache,
}

pub(crate) fn parse_styles(
    archive: &mut crate::XlsxZip,
    theme_colors: &[String],
) -> Result<ParsedStylesPart, String> {
    let xml = read_zip_string(archive, "xl/styles.xml")?;
    let doc = parse_guarded(&xml).map_err(|e| e.to_string())?;

    let normal_font_size = parse_normal_font_size(&doc);
    let default_font = parse_default_font(&doc);
    let chart_number_formats = crate::chart::ChartNumberFormatCache::from_document(
        &doc,
        &archive.operation()?.limit_reporter()?,
    )?;
    let num_fmts = parse_num_fmts(&doc);
    let fonts = parse_fonts(&doc, theme_colors);
    let fills = parse_fills(&doc, theme_colors);
    let borders = parse_borders(&doc, theme_colors);
    let mut cell_xfs = parse_cell_xfs(&doc);
    for (xf, own) in cell_xfs.iter_mut().zip(own_font_color_flags(&doc)) {
        xf.own_font_color = own;
    }
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
        normal_font_size,
        normal_font_color: normal_font_color_key(&doc),
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
        normal_font_size: parse_normal_font_size(&doc),
        normal_font_color: normal_font_color_key(&doc),
        chart_number_formats: crate::chart::ChartNumberFormatCache::from_document(
            &doc,
            &archive.operation()?.limit_reporter()?,
        )?,
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
                            if matches!(fc.attribute("auto"), Some("1" | "true")) {
                                toggles.automatic_color = Some(true);
                            }
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
                    section_divisors: Vec::new(),
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
                    section_divisors: Vec::new(),
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

/// An authored attribute value in its schema type (whitespace-collapsed and
/// numeric, so `theme="01"` equals `theme="1"`), or its trimmed text when it
/// is not a valid value of that type.
#[derive(Debug, Clone, PartialEq)]
enum Typed<T> {
    Value(T),
    Invalid(String),
}

fn typed<T: std::str::FromStr>(value: Option<&str>) -> Option<Typed<T>> {
    let value = value?.trim();
    Some(
        value
            .parse()
            .map(Typed::Value)
            .unwrap_or_else(|_| Typed::Invalid(value.to_string())),
    )
}

/// A font's `<color>` as authored (ECMA-376 §18.8.3 CT_Color), compared
/// without resolving it: Excel treats `rgb="FF000000"`, `indexed="8"`,
/// `auto="1"` and an absent `<color>` as colors distinct from the Normal
/// font's `theme="1"`, although all of them draw black. A zero `tint` is the
/// same as none.
#[derive(Debug, Clone, PartialEq)]
struct AuthoredColor {
    auto: Typed<bool>,
    rgb: Option<String>,
    theme: Option<Typed<u32>>,
    indexed: Option<Typed<u32>>,
    tint: Option<Typed<f64>>,
}

/// One `<color>` element as authored (see AuthoredColor).
fn authored_color(color: &roxmltree::Node) -> AuthoredColor {
    let rgb = color.attribute("rgb").map(|v| {
        let v = v.trim().to_ascii_uppercase();
        if v.len() == 6 {
            format!("FF{v}")
        } else {
            v
        }
    });
    AuthoredColor {
        // xsd:boolean: 1/true and 0/false; omitted means false.
        auto: match color.attribute("auto").map(str::trim) {
            None | Some("0" | "false") => Typed::Value(false),
            Some("1" | "true") => Typed::Value(true),
            Some(other) => Typed::Invalid(other.to_string()),
        },
        rgb,
        theme: typed(color.attribute("theme")),
        indexed: typed(color.attribute("indexed")),
        // A non-finite tint stays text so equality remains reflexive.
        tint: match typed::<f64>(color.attribute("tint")) {
            Some(Typed::Value(0.0)) => None,
            Some(Typed::Value(t)) if !t.is_finite() => color
                .attribute("tint")
                .map(|v| Typed::Invalid(v.trim().to_string())),
            other => other,
        },
    }
}

/// A canonical key for an authored `<color>`, equal exactly when two colors
/// are the same authored value. Rich-text runs carry it so their color can be
/// compared with the Normal style's once styles are known.
pub(crate) fn authored_color_key(color: &roxmltree::Node) -> String {
    format!("{:?}", authored_color(color))
}

/// The Normal cell style font's authored color key: the outer `None` when
/// Normal cannot be resolved, the inner `None` for a font without `<color>`.
pub(crate) fn normal_font_color_key(doc: &roxmltree::Document) -> Option<Option<String>> {
    let normal_xf = normal_style_xf(doc)?;
    let font_id = doc
        .descendants()
        .find(|n| n.tag_name().name() == "cellStyleXfs" && is_x_ns(n.tag_name().namespace()))?
        .children()
        .filter(|n| n.tag_name().name() == "xf")
        .nth(normal_xf)?
        .attribute("fontId")
        .and_then(|v| v.parse::<u32>().ok())
        .unwrap_or(0) as usize;
    authored_font_colors(doc)
        .get(font_id)
        .map(|color| color.as_ref().map(|c| format!("{c:?}")))
}

/// `<cellStyle builtinId="0">`'s `xfId` (ECMA-376 §18.8.7): the Normal style.
fn normal_style_xf(doc: &roxmltree::Document) -> Option<usize> {
    doc.descendants()
        .find(|n| {
            n.tag_name().name() == "cellStyle"
                && is_x_ns(n.tag_name().namespace())
                && n.attribute("builtinId")
                    .is_some_and(|v| v.trim().parse::<u32>() == Ok(0))
        })
        .and_then(|n| n.attribute("xfId"))
        .and_then(|v| v.trim().parse::<usize>().ok())
}

/// Each `<font>`'s authored color; the inner `None` is a font without a
/// `<color>` element.
fn authored_font_colors(doc: &roxmltree::Document) -> Vec<Option<AuthoredColor>> {
    let Some(fonts) = doc
        .descendants()
        .find(|n| n.tag_name().name() == "fonts" && is_x_ns(n.tag_name().namespace()))
    else {
        return Vec::new();
    };
    fonts
        .children()
        .filter(|n| n.tag_name().name() == "font")
        .map(|font| {
            font.children()
                .find(|c| c.tag_name().name() == "color")
                .map(|color| authored_color(&color))
        })
        .collect()
}

/// `own_font_color` for every `<cellXfs>` entry, in order. The Normal cell
/// style is `<cellStyle builtinId="0">` (ECMA-376 §18.8.7), whose `xfId`
/// selects its `<cellStyleXfs>` entry. A cell's own color is one whose font
/// color, or whose cell style's (`xf/@xfId`) font color, is authored
/// differently from the Normal font color. Excel draws such a color over a
/// table style's element font color — measured with Office-exported controls
/// covering theme / rgb / indexed / auto / absent colors and cell-style
/// inheritance.
///
/// The measurements cover resolvable references only. When the Normal style,
/// a font or a parent style cannot be resolved, the cell keeps the previous
/// precedence (the table color applies) rather than an unmeasured guess.
fn own_font_color_flags(doc: &roxmltree::Document) -> Vec<bool> {
    let colors = authored_font_colors(doc);
    // `fontId` is read as parse_cell_xfs reads it, so the flag describes the
    // font the cell actually renders with.
    let xfs_in = |tag: &str| -> Vec<(usize, Option<usize>)> {
        doc.descendants()
            .find(|n| n.tag_name().name() == tag && is_x_ns(n.tag_name().namespace()))
            .map(|list| {
                list.children()
                    .filter(|n| n.tag_name().name() == "xf")
                    .map(|xf| {
                        let font_id = xf
                            .attribute("fontId")
                            .and_then(|v| v.parse::<u32>().ok())
                            .unwrap_or(0) as usize;
                        let xf_id = match xf.attribute("xfId") {
                            None => Some(0),
                            Some(v) => v.trim().parse().ok(),
                        };
                        (font_id, xf_id)
                    })
                    .collect()
            })
            .unwrap_or_default()
    };
    let style_xfs = xfs_in("cellStyleXfs");
    let normal_xf = normal_style_xf(doc);
    // Normal's authored color, then one "differs from Normal" answer per font
    // (None: the font index does not resolve), so each xf is a lookup.
    let Some(normal) = normal_xf
        .and_then(|id| style_xfs.get(id))
        .and_then(|&(font_id, _)| colors.get(font_id))
    else {
        return Vec::new();
    };
    let differs: Vec<bool> = colors.iter().map(|color| color != normal).collect();
    let font_differs = |font_id: usize| differs.get(font_id).copied();
    let style_differs = |xf_id: Option<usize>| {
        xf_id
            .and_then(|id| style_xfs.get(id))
            .and_then(|&(font_id, _)| font_differs(font_id))
    };
    xfs_in("cellXfs")
        .into_iter()
        .map(
            |(font_id, xf_id)| match (font_differs(font_id), style_differs(xf_id)) {
                (Some(own), Some(style)) => own || style,
                _ => false,
            },
        )
        .collect()
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
                    own_font_color: false,
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
    fn normal_font_size_follows_the_normal_style() {
        let doc = |styles: &str| {
            format!(
                r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts><font><sz val="20"/></font><font><sz val="11"/></font></fonts><cellStyleXfs><xf fontId="0"/><xf fontId="1"/></cellStyleXfs>{styles}</styleSheet>"#
            )
        };
        let size = |xml: String| parse_normal_font_size(&roxmltree::Document::parse(&xml).unwrap());
        // Normal (builtinId 0) at cellStyleXfs[1] → fonts[1] = 11pt.
        assert_eq!(
            size(doc(
                r#"<cellStyles><cellStyle name="Normal" xfId="1" builtinId="0"/></cellStyles>"#
            )),
            Some(11.0)
        );
        // No Normal style declared: cellStyleXfs[0] → fonts[0] = 20pt.
        assert_eq!(size(doc("")), Some(20.0));
    }

    #[test]
    fn default_font_is_the_first_font_whatever_the_normal_style() {
        // Excel sizes columns from <fonts>[0] even when the Normal style
        // (builtinId 0) or <cellStyleXfs>[0] points at another font.
        let xml = r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fonts count="2">
            <font><b/><i/><sz val="20"/><name val="Arial"/></font>
            <font><sz val="11"/><name val="Calibri"/></font>
          </fonts>
          <cellStyleXfs count="1"><xf fontId="1"/></cellStyleXfs>
          <cellStyles><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
        </styleSheet>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        assert_eq!(
            parse_default_font(&doc),
            (Some("Arial".into()), Some(20.0), true, true)
        );
    }

    #[test]
    fn own_font_color_follows_excel_table_style_precedence() {
        // Expectations measured in Excel (PDF export of control workbooks):
        // only a color authored exactly like the Normal font's, with a cell
        // style that keeps it, lets a table style's font color through.
        // Normal is cellStyleXfs[1] here, found through builtinId="0".
        let xml = r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts>
    <font><sz val="11"/><color theme="1"/></font>
    <font><sz val="11"/><color rgb="FF000000"/></font>
    <font><b/><sz val="11"/><color theme="1"/></font>
    <font><sz val="11"/><color rgb="FFFF0000"/></font>
    <font><sz val="11"/></font>
    <font><sz val="11"/><color auto="1"/></font>
    <font><sz val="11"/><color indexed="8"/></font>
    <font><sz val="11"/><color theme="1" tint="0"/></font>
  </fonts>
  <cellStyleXfs><xf fontId="3"/><xf fontId="0"/></cellStyleXfs>
  <cellXfs>
    <xf fontId="0" xfId="1"/><xf fontId="1" xfId="1"/><xf fontId="2" xfId="1"/>
    <xf fontId="3" xfId="1"/><xf fontId="4" xfId="1"/><xf fontId="5" xfId="1"/>
    <xf fontId="6" xfId="1"/><xf fontId="7" xfId="1"/>
    <xf fontId="3" xfId="0"/><xf fontId="0" xfId="0"/>
  </cellXfs>
  <cellStyles><cellStyle name="Red" xfId="0"/><cellStyle name="Normal" xfId="1" builtinId="0"/></cellStyles>
</styleSheet>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        assert_eq!(
            own_font_color_flags(&doc),
            [
                false, // Normal font
                true,  // rgb black
                false, // theme black, bold
                true,  // red
                true,  // no <color>
                true,  // auto
                true,  // indexed black
                false, // theme black, zero tint
                true,  // red from the cell style
                true,  // Normal font under a red cell style
            ]
        );
    }

    #[test]
    fn own_font_color_compares_typed_values_and_fails_safe() {
        let sheet = |fonts: &str, style_xfs: &str, cell_xfs: &str, styles: &str| {
            format!(
                r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts>{fonts}</fonts><cellStyleXfs>{style_xfs}</cellStyleXfs>
  <cellXfs>{cell_xfs}</cellXfs>{styles}</styleSheet>"#
            )
        };
        let flags = |xml: String| own_font_color_flags(&roxmltree::Document::parse(&xml).unwrap());
        let normal =
            r#"<cellStyles><cellStyle name="Normal" xfId="1" builtinId="00"/></cellStyles>"#;
        // Same values spelled differently are the same authored color; the
        // Normal style is still found through a zero-padded builtinId.
        assert_eq!(
            flags(sheet(
                r#"<font><color theme="1" tint=" 0.5 "/></font><font><color theme="01" tint="0.50"/></font><font><color rgb="FFFF0000"/></font>"#,
                r#"<xf fontId="2"/><xf fontId="0"/>"#,
                r#"<xf fontId="1" xfId="1"/><xf fontId="2" xfId="1"/>"#,
                normal,
            )),
            [false, true]
        );
        // Without a resolvable Normal style nothing is classified as owned.
        assert!(flags(sheet(
            r#"<font><color theme="1"/></font><font><color rgb="FFFF0000"/></font>"#,
            r#"<xf fontId="1"/><xf fontId="0"/>"#,
            r#"<xf fontId="0" xfId="1"/>"#,
            "",
        ))
        .is_empty());
        // A non-finite tint compares equal to itself; auto spellings
        // normalize, while an invalid auto value stays distinct.
        assert_eq!(
            flags(sheet(
                r#"<font><color theme="1" tint="NaN" auto="0"/></font><font><color theme="1" tint=" NaN " auto="false"/></font><font><color theme="1" tint="NaN" auto="no"/></font>"#,
                r#"<xf fontId="0"/>"#,
                r#"<xf fontId="0" xfId="0"/><xf fontId="1" xfId="0"/><xf fontId="2" xfId="0"/>"#,
                r#"<cellStyles><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>"#,
            )),
            [false, false, true]
        );
        // A fontId beyond u32 falls back to font 0, as parse_cell_xfs does,
        // so the cell is classified by the red font it renders with.
        assert_eq!(
            flags(sheet(
                r#"<font><color rgb="FFFF0000"/></font><font><color theme="1"/></font>"#,
                r#"<xf fontId="1"/>"#,
                r#"<xf fontId="4294967296" xfId="0"/>"#,
                r#"<cellStyles><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>"#,
            )),
            [true]
        );
        // An unresolvable font or parent style keeps the table color.
        assert_eq!(
            flags(sheet(
                r#"<font><color theme="1"/></font>"#,
                r#"<xf fontId="0"/>"#,
                r#"<xf fontId="9" xfId="0"/><xf fontId="0" xfId="7"/>"#,
                r#"<cellStyles><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>"#,
            )),
            [false, false]
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
