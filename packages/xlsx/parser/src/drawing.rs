use crate::read_zip_string;
use crate::types::*;
use crate::{parse_rels_map, resolve_zip_path};
// Shared DrawingML blip helpers (ECMA-376 §20.1.8.13 + Microsoft 2016 SVG
// extension, MS-ODRAWXML). `mime_from_ext` is the single source of truth for
// `.svg ⇒ image/svg+xml`; `svg_blip_rid` resolves the vector original nested in
// a blip's `<a:extLst>`. Replaces xlsx's former local `mime_from_ext` (a strict
// subset that lacked the `svg` arm and so dropped SVG parts).
use ooxml_common::blip::{
    blip_embed_rid, mime_from_ext, parse_blip_alpha, parse_blip_duotone, parse_src_rect,
    svg_blip_rid,
};
use ooxml_common::depth::{parse_guarded, DepthGuard};
use ooxml_common::drawing::{DrawingGroupSpec, DrawingGroupTransform, DrawingRect};
use ooxml_common::ns::{attr_ns, is_a_ns, is_xdr_ns, relationships};
use ooxml_common::theme::ThemeFormatScheme;
use ooxml_common::units::EMU_PER_PX_96DPI;
use std::collections::HashMap;
// `Cursor` is only used to build in-memory archives in this module's tests; the
// production archive type is `crate::XlsxZip`.
#[cfg(test)]
use std::io::Cursor;

fn parse_xdr_col(value: &str) -> u32 {
    parse_xdr_index(value)
}

fn parse_xdr_row(value: &str) -> u32 {
    parse_xdr_index(value)
}

/// ST_ColID / ST_RowID are non-negative `xsd:int` values. They are not bounded
/// by the worksheet grid in the DrawingML schema, so preserve every conforming
/// off-grid marker as an acquisition fact. Geometry code performs O(log k)
/// culling without materializing intervening cells. Invalid lexical/range
/// values retain the parser's established zero fallback.
fn parse_xdr_index(value: &str) -> u32 {
    value
        .trim()
        .parse::<i32>()
        .ok()
        .filter(|value| *value >= 0)
        .map(|value| value as u32)
        .unwrap_or(0)
}

fn vml_xdr_index(value: i64) -> u32 {
    i32::try_from(value)
        .ok()
        .filter(|value| *value >= 0)
        .map(|value| value as u32)
        .unwrap_or(0)
}

/// ECMA-376 Part 3 §9.3 "given application configuration" for the xlsx parser:
/// the Microsoft chartEx extension (the 2014 payload namespace and Excel's
/// 2015/9/8 MCE capability namespace), the 2010 drawing extensions (a14 — wraps
/// ordinary shapes/charts and OMML `a14:m` equations), and the 2009/9
/// spreadsheet extensions (x14 — slicers and the OLE `objectPr` image preview).
/// All have real render paths here.
pub(crate) fn xlsx_understands_ns(ns: &str) -> bool {
    matches!(
        ns,
        "http://schemas.microsoft.com/office/drawing/2014/chartex"
            | "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"
            | "http://schemas.microsoft.com/office/drawing/2010/main"
            | "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
    )
}

/// True when an `<xdr:sp>` / `<xdr:pic>` / `<xdr:grpSp>` (or `<xdr:cxnSp>`)
/// node's own `<xdr:cNvPr hidden="1">` marks it hidden (ECMA-376 §20.1.2.2.8
/// `CT_NonVisualDrawingProps@hidden`, `xsd:boolean`, default `false`). Only the
/// node's direct non-visual-properties wrapper is inspected, never a
/// descendant, so a group's own props are not confused with a nested shape's.
/// A hidden drawing object is not rendered — skipped at parse time.
pub(crate) fn xdr_node_hidden(node: &roxmltree::Node) -> bool {
    for wrapper in &[
        "nvSpPr",
        "nvPicPr",
        "nvGrpSpPr",
        "nvCxnSpPr",
        "nvGraphicFramePr",
    ] {
        if let Some(nv) = node
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == *wrapper)
        {
            if let Some(cnv) = nv
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "cNvPr")
            {
                return ooxml_common::drawing::nv_props_hidden(cnv);
            }
        }
    }
    false
}

/// Parse `<xdr:twoCellAnchor>` elements from a drawing XML and resolve
/// embedded pictures into data URLs. `drawing_dir` is the folder that
/// contains `drawing_path` so relative `Target`s resolve correctly.
pub(crate) fn parse_drawing_anchors(
    drawing_xml: &str,
    drawing_rels: &HashMap<String, String>,
    drawing_dir: &str,
    archive: &mut crate::XlsxZip,
    // Positional clrScheme palette, for resolving a picture's `<a:duotone>`
    // effect colours (§20.1.8.23).
    theme_colors: &[String],
) -> Vec<ImageAnchor> {
    let Ok(doc) = parse_guarded(drawing_xml) else {
        return Vec::new();
    };
    let mut anchors: Vec<ImageAnchor> = Vec::new();

    for anchor in doc.descendants() {
        if anchor.tag_name().name() != "twoCellAnchor" || !is_xdr_ns(anchor.tag_name().namespace())
        {
            continue;
        }
        let (mut from_col, mut from_col_off, mut from_row, mut from_row_off) =
            (0u32, 0i64, 0u32, 0i64);
        let (mut to_col, mut to_col_off, mut to_row, mut to_row_off) = (0u32, 0i64, 0u32, 0i64);
        let mut pic_rid: Option<String> = None;
        // Microsoft 2016 SVG extension (MS-ODRAWXML): the blip's vector original,
        // nested in `<a:blip><a:extLst>`. The raster `r:embed` (pic_rid) is only a
        // fallback for SVG-incapable clients.
        let mut svg_rid: Option<String> = None;
        let mut native_ext_cx: i64 = 0;
        let mut native_ext_cy: i64 = 0;
        let mut rotation = None;
        let mut flip_h = None;
        let mut flip_v = None;
        // ECMA-376 §20.1.8.55 `<a:srcRect>` source-image crop (None ⇒ uncropped).
        let mut src_rect: Option<SrcRect> = None;
        // ECMA-376 §20.1.8.6 `<a:alphaModFix>` opacity (None ⇒ opaque) and
        // §20.1.8.23 `<a:duotone>` recolour (None ⇒ no effect).
        let mut alpha: Option<f64> = None;
        let mut duotone: Option<Duotone> = None;
        // ECMA-376 §20.5.2.33 `twoCellAnchor@editAs`. Possible values:
        // "twoCell" (default), "oneCell", "absolute". With "oneCell" Excel
        // preserves the picture's saved size from <xdr:spPr><a:xfrm><a:ext>
        // regardless of cell resizing.
        let edit_as = anchor.attribute("editAs").map(|s| s.to_string());

        for child in anchor.children() {
            if !child.is_element() {
                continue;
            }
            match child.tag_name().name() {
                "from" | "to" => {
                    let is_from = child.tag_name().name() == "from";
                    let mut col: u32 = 0;
                    let mut col_off: i64 = 0;
                    let mut row: u32 = 0;
                    let mut row_off: i64 = 0;
                    for c in child.children() {
                        match (c.tag_name().name(), c.text()) {
                            ("col", Some(t)) => col = parse_xdr_col(t),
                            ("colOff", Some(t)) => col_off = t.trim().parse().unwrap_or(0),
                            ("row", Some(t)) => row = parse_xdr_row(t),
                            ("rowOff", Some(t)) => row_off = t.trim().parse().unwrap_or(0),
                            _ => {}
                        }
                    }
                    if is_from {
                        from_col = col;
                        from_col_off = col_off;
                        from_row = row;
                        from_row_off = row_off;
                    } else {
                        to_col = col;
                        to_col_off = col_off;
                        to_row = row;
                        to_row_off = row_off;
                    }
                }
                "pic" => {
                    // §20.1.2.2.8 — a `<xdr:cNvPr hidden="1">` picture is not
                    // rendered. Skip the whole anchor (a twoCellAnchor holds a
                    // single drawing object).
                    if xdr_node_hidden(&child) {
                        continue;
                    }
                    // <xdr:pic><xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill></xdr:pic>
                    let blip_fill = child.children().find(|n| {
                        n.tag_name().name() == "blipFill" && is_xdr_ns(n.tag_name().namespace())
                    });
                    if let Some(bf) = blip_fill {
                        let blip = bf.children().find(|n| {
                            n.tag_name().name() == "blip" && is_a_ns(n.tag_name().namespace())
                        });
                        if let Some(b) = blip {
                            // Raster fallback (`<a:blip r:embed>`); tolerate the
                            // literal `r:embed` form via the shared helper.
                            pic_rid = blip_embed_rid(&b);
                            // Vector original (`<asvg:svgBlip r:embed>` inside the
                            // blip's `<a:extLst>`), matched by namespace-local name.
                            svg_rid = svg_blip_rid(b);
                        }
                        // `<a:srcRect>` is a sibling of `<a:blip>` inside blipFill.
                        src_rect = parse_src_rect(bf);
                        // Blip effects: `<a:alphaModFix>` opacity (§20.1.8.6) and
                        // `<a:duotone>` recolour (§20.1.8.23). The duotone colours
                        // resolve through the shared DrawingML grammar using the
                        // sheet's theme palette.
                        alpha = parse_blip_alpha(bf);
                        duotone = parse_blip_duotone(
                            bf,
                            &XlsxSchemeResolver { theme_colors },
                            ooxml_common::color::TintMode::PowerPointLinear,
                        );
                    }
                    // <xdr:pic><xdr:spPr><a:xfrm><a:ext cx cy>: the picture's
                    // own saved EMU extent. Authoritative when editAs="oneCell".
                    if let Some(sp_pr) = child.children().find(|n| {
                        n.tag_name().name() == "spPr" && is_xdr_ns(n.tag_name().namespace())
                    }) {
                        if let Some(xfrm_n) = sp_pr.children().find(|n| {
                            n.tag_name().name() == "xfrm" && is_a_ns(n.tag_name().namespace())
                        }) {
                            // CT_Transform2D permits xfrm without an ext. The
                            // anchor still supplies the destination bounds, so
                            // preserve its independent transform attributes.
                            rotation = xfrm_n
                                .attribute("rot")
                                .and_then(|value| value.parse::<i32>().ok())
                                .filter(|value| *value != 0)
                                .map(|value| value as f64 / 60000.0);
                            flip_h = matches!(xfrm_n.attribute("flipH"), Some("1") | Some("true"))
                                .then_some(true);
                            flip_v = matches!(xfrm_n.attribute("flipV"), Some("1") | Some("true"))
                                .then_some(true);
                            if let Some(xfrm) = parse_xfrm(&xfrm_n) {
                                native_ext_cx = xfrm.ext_x as i64;
                                native_ext_cy = xfrm.ext_y as i64;
                            }
                        }
                    }
                }
                _ => {}
            }
        }

        // Resolve a blip rId → the media part's zip path (verifying the entry
        // exists, so a dangling rId is dropped exactly as before). The renderer
        // fetches the bytes lazily via `extract_image`; no base64 is inlined.
        let mut resolve = |rid: &str| -> Option<String> {
            let target = drawing_rels.get(rid)?;
            let media_path = resolve_zip_path(drawing_dir, target);
            // Confirm the entry resolves before emitting its path (preserves the
            // previous "drop when bytes are missing" semantics). `index_for_name`
            // reads only the central directory — no inflate, unlike the former
            // `read_zip_bytes` which decompressed the entry only to discard it.
            archive.index_for_name(&media_path)?;
            Some(media_path)
        };

        // Vector original first (so an svg-only picture is never dropped); raster
        // fallback second.
        let svg_image_path = svg_rid.as_deref().and_then(&mut resolve);
        let raster_path = pic_rid.as_deref().and_then(&mut resolve);

        // A picture needs at least one drawable source. Prefer the raster as
        // `image_path` (Excel's compatibility fallback); when no raster is
        // embedded, fall back to the SVG itself so the element is always
        // drawable. Drop only when neither resolves. ECMA-376 §20.1.8.13 +
        // MS-ODRAWXML svgBlip.
        let image_path = match (raster_path, svg_image_path.as_ref()) {
            (Some(raster), _) => raster,
            (None, Some(svg)) => svg.clone(),
            (None, None) => continue,
        };
        let mime_type = mime_from_ext(&image_path).to_string();

        anchors.push(ImageAnchor {
            z_order: anchor
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "pic")
                .map(|n| n.range().start as u64)
                .unwrap_or(anchor.range().start as u64),
            from_col,
            from_col_off,
            from_row,
            from_row_off,
            to_col,
            to_col_off,
            to_row,
            to_row_off,
            edit_as,
            native_ext_cx,
            native_ext_cy,
            rotation,
            flip_h,
            flip_v,
            image_path,
            mime_type,
            svg_image_path,
            src_rect,
            alpha,
            duotone,
        });
    }
    anchors
}

// ─── Shape group parsing ────────────────────────────────────────────────────
//
// ECMA-376 §20.5.2.17 `<xdr:grpSp>` / §20.1.9 DrawingML shapes. Each
// top-level grpSp inside a twoCellAnchor has its own coordinate system:
//   - grpSpPr/xfrm/off,ext     : group's position/size in parent coords
//   - grpSpPr/xfrm/chOff,chExt : origin/extent of the group's child coords
//
// A child sp at child coord (cx, cy) maps to parent coord:
//   parent.x = off.x + (cx - chOff.x) / chExt.cx * ext.cx
//
// For rendering, we chain these transforms down to the top-level grpSp and
// then normalize each leaf shape's rect into [0,1] of the top-level ext.

#[derive(Clone, Copy)]
pub(crate) struct Xfrm {
    off_x: f64,
    off_y: f64,
    ext_x: f64,
    ext_y: f64,
    ch_off_x: f64,
    ch_off_y: f64,
    ch_ext_x: f64,
    ch_ext_y: f64,
    has_ch: bool,
    rot: f64,
    flip_h: bool,
    flip_v: bool,
}

pub(crate) fn parse_xfrm(xfrm_node: &roxmltree::Node) -> Option<Xfrm> {
    let mut off = (0.0_f64, 0.0_f64);
    let mut ext = (0.0_f64, 0.0_f64);
    let mut ch_off = (0.0_f64, 0.0_f64);
    let mut ch_ext = (0.0_f64, 0.0_f64);
    let mut has_ext = false;
    let mut has_ch = false;
    for c in xfrm_node.children() {
        match c.tag_name().name() {
            "off" => {
                off.0 = c.attribute("x").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                off.1 = c.attribute("y").and_then(|s| s.parse().ok()).unwrap_or(0.0);
            }
            "ext" => {
                ext.0 = c
                    .attribute("cx")
                    .and_then(|s| s.parse().ok())
                    .unwrap_or(0.0);
                ext.1 = c
                    .attribute("cy")
                    .and_then(|s| s.parse().ok())
                    .unwrap_or(0.0);
                has_ext = true;
            }
            "chOff" => {
                ch_off.0 = c.attribute("x").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                ch_off.1 = c.attribute("y").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                has_ch = true;
            }
            "chExt" => {
                ch_ext.0 = c
                    .attribute("cx")
                    .and_then(|s| s.parse().ok())
                    .unwrap_or(0.0);
                ch_ext.1 = c
                    .attribute("cy")
                    .and_then(|s| s.parse().ok())
                    .unwrap_or(0.0);
                has_ch = true;
            }
            _ => {}
        }
    }
    if !has_ext {
        return None;
    }
    Some(Xfrm {
        off_x: off.0,
        off_y: off.1,
        ext_x: ext.0,
        ext_y: ext.1,
        ch_off_x: ch_off.0,
        ch_off_y: ch_off.1,
        ch_ext_x: if ch_ext.0 == 0.0 { ext.0 } else { ch_ext.0 },
        ch_ext_y: if ch_ext.1 == 0.0 { ext.1 } else { ch_ext.1 },
        has_ch,
        rot: xfrm_node
            .attribute("rot")
            .and_then(|value| value.parse::<f64>().ok())
            .unwrap_or(0.0)
            / 60000.0,
        flip_h: matches!(xfrm_node.attribute("flipH"), Some("1") | Some("true")),
        flip_v: matches!(xfrm_node.attribute("flipV"), Some("1") | Some("true")),
    })
}

/// Resolves a `<a:schemeClr val>` name to its base theme hex (no `#`) the way
/// xlsx stores its palette, for the shared [`ooxml_common::color::parse_color_node`].
/// `theme_colors` is collected in OOXML clrScheme document order: dk1, lt1, dk2,
/// lt2, accent1..accent6, hlink, folHlink (see `parse_theme_colors`), each
/// `#RRGGBB`. The `#` is stripped here because the shared transform wants clean
/// hex; `parse_solid_fill` re-adds it for the returned color.
///
/// Shared with the chart parser (`chart::extract_solid_fill_in_drawingml`) so a
/// single positional-slot table serves every xlsx DrawingML color path — the
/// chart module previously carried a private copy with `bg2`/`tx2` mapped to the
/// wrong slots (§20.1.6.2).
pub(crate) struct XlsxSchemeResolver<'a> {
    pub(crate) theme_colors: &'a [String],
}

impl ooxml_common::color::ThemeResolver for XlsxSchemeResolver<'_> {
    fn resolve_scheme_color(&self, name: &str) -> Option<String> {
        // Positional index into the clrScheme-order palette. dk1/lt1 and dk2/lt2
        // are the light/dark logical pairs (the earlier local mapping had them
        // swapped, which darkened shapes that painted "lt1", the paper colour).
        let idx = match name {
            "dk1" | "tx1" => 0,
            "lt1" | "bg1" => 1,
            "dk2" | "tx2" => 2,
            "lt2" | "bg2" => 3,
            "accent1" => 4,
            "accent2" => 5,
            "accent3" => 6,
            "accent4" => 7,
            "accent5" => 8,
            "accent6" => 9,
            "hlink" => 10,
            "folHlink" => 11,
            _ => return None,
        };
        self.theme_colors
            .get(idx)
            .map(|hex| hex.trim_start_matches('#').to_owned())
    }
}

/// Resolve a DrawingML `<a:solidFill>` (or any color container) to `#RRGGBB`
/// (uppercase). Thin wrapper over the shared
/// [`ooxml_common::color::parse_color_node`]: the color grammar
/// (srgbClr/sysClr/prstClr/schemeClr + transforms lumMod/tint/…) lives there;
/// [`XlsxSchemeResolver`] supplies the positional theme-palette lookup. xlsx
/// re-applies its own `#`-prefix + uppercase convention on the shared output
/// (which is already uppercase, no `#`). Without the transforms an accent with
/// e.g. `lumMod 20% + lumOff 80%` (a light tint) would render at full strength.
pub(crate) fn parse_solid_fill(
    fill_node: &roxmltree::Node,
    theme_colors: &[String],
) -> Option<String> {
    ooxml_common::color::parse_color_node(
        *fill_node,
        &XlsxSchemeResolver { theme_colors },
        ooxml_common::color::TintMode::PowerPointLinear,
    )
    .map(|hex| format!("#{}", hex.to_uppercase()))
}

fn parse_xlsx_drawingml_fill(
    fill: roxmltree::Node,
    resolver: &(impl ooxml_common::color::ThemeResolver + ?Sized),
) -> Option<ShapeFill> {
    let tint_mode = ooxml_common::color::TintMode::PowerPointLinear;
    match fill.tag_name().name() {
        "solidFill" => {
            ooxml_common::color::parse_color_node(fill, resolver, tint_mode).map(|color| {
                ShapeFill::Solid {
                    color: format!("#{}", color.to_uppercase()),
                }
            })
        }
        "gradFill" => {
            let gradient = ooxml_common::fill::parse_grad_fill(fill, resolver, tint_mode)?;
            Some(ShapeFill::Gradient {
                stops: gradient.stops,
                angle: gradient.angle,
                grad_type: gradient.grad_type,
                scaled: gradient.scaled,
                path: gradient.path,
                fill_to_rect: gradient.fill_to_rect,
                tile_rect: gradient.tile_rect,
                flip: gradient.flip,
                rot_with_shape: gradient.rot_with_shape,
            })
        }
        "pattFill" => {
            let pattern = ooxml_common::fill::parse_patt_fill(fill, resolver, tint_mode);
            Some(ShapeFill::Pattern {
                fg: pattern.fg,
                bg: pattern.bg,
                preset: pattern.preset,
            })
        }
        "noFill" => None,
        _ => None,
    }
}

fn xlsx_fill_fallback_color(fill: &ShapeFill) -> Option<String> {
    match fill {
        ShapeFill::Solid { color } => Some(color.clone()),
        ShapeFill::Gradient { stops, .. } => stops
            .iter()
            .rev()
            .find(|stop| !stop.color.ends_with("00"))
            .or_else(|| stops.last())
            .map(|stop| format!("#{}", stop.color.to_uppercase())),
        ShapeFill::Pattern { fg, .. } => Some(format!("#{}", fg.to_uppercase())),
    }
}

fn resolve_xlsx_fill_ref(
    fill_ref: roxmltree::Node,
    theme_colors: &[String],
    format_scheme: &ThemeFormatScheme,
) -> Option<ShapeFill> {
    use ooxml_common::theme::StyleMatrixLookup;

    let index = fill_ref.attribute("idx")?.parse::<usize>().ok()?;
    let StyleMatrixLookup::Entry(entry) = format_scheme.lookup_fill_ref(index) else {
        return None;
    };
    let entry_xml = entry.to_xml();
    let document = roxmltree::Document::parse(&entry_xml).ok()?;
    let fill = document
        .root_element()
        .children()
        .find(|node| node.is_element())?;
    let base_resolver = XlsxSchemeResolver { theme_colors };
    let reference_color = ooxml_common::color::parse_color_node(
        fill_ref,
        &base_resolver,
        ooxml_common::color::TintMode::PowerPointLinear,
    );
    let resolver = ooxml_common::color::StyleMatrixColorResolver::new(
        &base_resolver,
        reference_color.as_deref(),
    );
    parse_xlsx_drawingml_fill(fill, &resolver)
}

/// Resolve an XLSX shape's `a:lnRef` against the complete DrawingML format
/// scheme. ST_StyleMatrixColumnIndex is one-based and `0` means no style
/// (§20.1.10.57). The optional reference color substitutes `phClr`; fixed
/// colors in the recipe remain valid without it (§20.1.4.2.19).
fn resolve_xlsx_line_ref(
    line_ref: roxmltree::Node,
    theme_colors: &[String],
    format_scheme: &ThemeFormatScheme,
) -> Option<ooxml_common::line::LineProperties> {
    use ooxml_common::theme::StyleMatrixLookup;

    let index = line_ref.attribute("idx")?.parse::<usize>().ok()?;
    let StyleMatrixLookup::Entry(entry) = format_scheme.lookup_line_ref(index) else {
        return None;
    };
    let entry_xml = entry.to_xml();
    let document = roxmltree::Document::parse(&entry_xml).ok()?;
    let line = document.root_element().children().find(|node| {
        node.is_element()
            && node.tag_name().name() == "ln"
            && ooxml_common::ns::is_a_ns(node.tag_name().namespace())
    })?;
    let base_resolver = XlsxSchemeResolver { theme_colors };
    let reference_color = ooxml_common::color::parse_color_node(
        line_ref,
        &base_resolver,
        ooxml_common::color::TintMode::PowerPointLinear,
    );
    let resolver = ooxml_common::color::StyleMatrixColorResolver::new(
        &base_resolver,
        reference_color.as_deref(),
    );
    let mut properties = ooxml_common::line::parse_line_properties(
        line,
        &resolver,
        ooxml_common::color::TintMode::PowerPointLinear,
    );
    if properties.paint.is_none() {
        properties.paint = reference_color
            .map(|color| ooxml_common::line::LinePaint::Solid { color: Some(color) });
    }
    Some(properties)
}

#[derive(Default)]
struct XlsxLineWireProperties {
    color: Option<String>,
    width: i64,
    fill: Option<ShapeStrokeFill>,
    dash_style: Option<String>,
    custom_dash: Vec<ShapeLineDashSegment>,
    line_cap: Option<String>,
    line_join: Option<String>,
    miter_limit: Option<f64>,
    alignment: Option<String>,
    cmpd: Option<String>,
    head_end: Option<ShapeLineEnd>,
    tail_end: Option<ShapeLineEnd>,
}

fn xlsx_line_end(end: &ooxml_common::line::LineEnd) -> Option<ShapeLineEnd> {
    let kind = end.kind.as_deref().unwrap_or("none");
    (kind != "none").then(|| ShapeLineEnd {
        r#type: kind.to_owned(),
        w: end.width.clone().unwrap_or_else(|| "med".to_owned()),
        len: end.length.clone().unwrap_or_else(|| "med".to_owned()),
    })
}

fn xlsx_line_wire_properties(line: &ooxml_common::line::LineProperties) -> XlsxLineWireProperties {
    let (color, fill) = match line.paint.as_ref() {
        Some(ooxml_common::line::LinePaint::Solid { color }) => (color.clone(), None),
        Some(ooxml_common::line::LinePaint::Gradient(Some(gradient))) => {
            let color = gradient
                .stops
                .iter()
                .rev()
                .find(|stop| !stop.color.ends_with("00"))
                .or_else(|| gradient.stops.last())
                .map(|stop| stop.color.clone());
            (
                color,
                Some(ShapeStrokeFill::Gradient {
                    stops: gradient.stops.clone(),
                    angle: gradient.angle,
                    grad_type: gradient.grad_type.clone(),
                    scaled: gradient.scaled,
                    path: gradient.path.clone(),
                    fill_to_rect: gradient.fill_to_rect.clone(),
                    tile_rect: gradient.tile_rect.clone(),
                    flip: gradient.flip.clone(),
                    rot_with_shape: gradient.rot_with_shape,
                }),
            )
        }
        Some(ooxml_common::line::LinePaint::Pattern(pattern)) => (
            Some(pattern.fg.clone()),
            Some(ShapeStrokeFill::Pattern {
                fg: pattern.fg.clone(),
                bg: pattern.bg.clone(),
                preset: pattern.preset.clone(),
            }),
        ),
        Some(ooxml_common::line::LinePaint::NoFill)
        | Some(ooxml_common::line::LinePaint::Gradient(None))
        | None => (None, None),
    };
    let color = color.map(|color| format!("#{}", color.trim_start_matches('#').to_uppercase()));
    let width = color
        .as_ref()
        .map(|_| line.width.unwrap_or(9525))
        .unwrap_or(0);
    let (dash_style, custom_dash) = match line.dash.as_ref() {
        Some(ooxml_common::line::LineDash::Preset(Some(value))) if value != "solid" => {
            (Some(value.clone()), Vec::new())
        }
        Some(ooxml_common::line::LineDash::Custom(stops)) => (
            None,
            stops
                .iter()
                .map(|stop| ShapeLineDashSegment {
                    dash: stop.dash / 100_000.0,
                    space: stop.space / 100_000.0,
                })
                .collect(),
        ),
        _ => (None, Vec::new()),
    };
    let line_cap = line.cap.as_deref().and_then(|cap| match cap {
        "rnd" => Some("round".to_owned()),
        "sq" => Some("square".to_owned()),
        "flat" => Some("butt".to_owned()),
        _ => None,
    });
    let (line_join, miter_limit) = match line.join.as_ref() {
        Some(ooxml_common::line::LineJoin::Round) => (Some("round".to_owned()), None),
        Some(ooxml_common::line::LineJoin::Bevel) => (Some("bevel".to_owned()), None),
        Some(ooxml_common::line::LineJoin::Miter { limit }) => (
            Some("miter".to_owned()),
            limit.map(|value| value as f64 / 100_000.0),
        ),
        None => (None, None),
    };
    XlsxLineWireProperties {
        color,
        width,
        fill,
        dash_style,
        custom_dash,
        line_cap,
        line_join,
        miter_limit,
        alignment: line
            .alignment
            .clone()
            .filter(|value| matches!(value.as_str(), "ctr" | "in")),
        cmpd: line
            .compound
            .clone()
            .filter(|value| value.as_str() != "sng"),
        head_end: line.head_end.as_ref().and_then(xlsx_line_end),
        tail_end: line.tail_end.as_ref().and_then(xlsx_line_end),
    }
}

/// Adapt the shared DrawingML `a:custGeom` grammar to XLSX's existing raw
/// path-coordinate wire model.
fn parse_custom_paths(
    cust_geom: roxmltree::Node<'_, '_>,
    shape_width: f64,
    shape_height: f64,
) -> Vec<PathInfo> {
    use ooxml_common::custom_geometry::{parse_custom_geometry, PathCommand};

    parse_custom_geometry(cust_geom, shape_width, shape_height)
        .paths
        .into_iter()
        .map(|path| {
            let commands = path
                .commands
                .into_iter()
                .map(|command| match command {
                    PathCommand::MoveTo { x, y } => PathCmd::MoveTo { x, y },
                    PathCommand::LineTo { x, y } => PathCmd::LineTo { x, y },
                    PathCommand::CubicBezierTo {
                        x1,
                        y1,
                        x2,
                        y2,
                        x,
                        y,
                    } => PathCmd::CubicBezTo {
                        x1,
                        y1,
                        x2,
                        y2,
                        x3: x,
                        y3: y,
                    },
                    PathCommand::QuadraticBezierTo { x1, y1, x, y } => PathCmd::QuadBezTo {
                        x1,
                        y1,
                        x2: x,
                        y2: y,
                    },
                    PathCommand::ArcTo {
                        wr,
                        hr,
                        st_ang,
                        sw_ang,
                    } => PathCmd::ArcTo {
                        wr,
                        hr,
                        st_ang,
                        sw_ang,
                    },
                    PathCommand::Close => PathCmd::Close,
                })
                .collect();
            PathInfo {
                w: path.width,
                h: path.height,
                commands,
            }
        })
        .collect()
}

/// Parse `<xdr:txBody>` into a `ShapeText`. Returns `None` if the body
/// contains no visible runs. Run formatting follows ECMA-376 §21.1.2.3.1
/// (`<a:rPr>`): `sz` is hundredths of a point, `b="1"` = bold, `i="1"`
/// = italic, `<a:solidFill>` overrides shape-level font color,
/// `<a:latin@typeface>` selects the Latin font face, and `<a:ea@typeface>` /
/// `<a:cs@typeface>` the East-Asian / complex-script faces. The renderer
/// floors the line box by whichever of these declares a tabled (Meiryo /
/// Sakkal Majalla) design line — the common Japanese encoding sets Meiryo on
/// `<a:ea>` while leaving `<a:latin>` default. Per-glyph font switching is
/// still not modeled (a larger change); only the line-box floor uses ea/cs.
/// Point size for an equation: the first `a:rPr@sz` on an actual math RUN
/// (`m:r`), in pt. Scoped to `m:r` so the size on `<m:ctrlPr>` (control
/// properties — structural delimiters, not visible glyphs) is ignored.
fn math_run_size_shape(om: roxmltree::Node) -> Option<f64> {
    om.descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "r")
        .find_map(|r| {
            r.descendants()
                .filter(|n| n.is_element() && n.tag_name().name() == "rPr")
                .find_map(|rpr| rpr.attribute("sz").and_then(|s| s.parse::<f64>().ok()))
        })
        .map(|v| v / 100.0)
}

/// Equation colour: the first `a:rPr/solidFill` on an actual math RUN (`m:r`).
/// Scoped to `m:r` so a colour on `<m:ctrlPr>` (structural control properties,
/// which Excel does NOT use as the visible glyph colour) is ignored — equations
/// with no explicit run colour then fall back to the shape font colour (black).
fn math_run_color_shape(om: roxmltree::Node, theme_colors: &[String]) -> Option<String> {
    om.descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "r")
        .find_map(|r| {
            r.descendants()
                .filter(|n| n.is_element() && n.tag_name().name() == "rPr")
                .find_map(|rpr| {
                    rpr.children()
                        .find(|n| n.is_element() && n.tag_name().name() == "solidFill")
                        .and_then(|sf| parse_solid_fill(&sf, theme_colors))
                })
        })
}

/// Emit `ShapeTextRun::Math` runs from an OMML wrapper inside a `<a:p>`. Handles
/// the four shapes Excel/PowerPoint produce — bare `m:oMath`, `m:oMathPara`
/// (block), the `a14:m` wrapper (local name `m`), and `mc:AlternateContent`
/// (take the `a14` `Choice`, ignore the rasterized `Fallback`). Direct port of
/// the pptx parser's `push_math_runs` (ECMA-376 §22.1); reuses the shared
/// `ooxml_common::math` OMML→AST parser unchanged.
fn push_math_runs_shape(
    node: roxmltree::Node,
    theme_colors: &[String],
    runs: &mut Vec<ShapeTextRun>,
) {
    match node.tag_name().name() {
        "oMath" => {
            let nodes = ooxml_common::math::parse_omath_nodes(node);
            if !nodes.is_empty() {
                runs.push(ShapeTextRun::Math {
                    nodes,
                    display: false,
                    font_size: math_run_size_shape(node),
                    color: math_run_color_shape(node, theme_colors),
                });
            }
        }
        "oMathPara" => {
            for om in node
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "oMath")
            {
                let nodes = ooxml_common::math::parse_omath_nodes(om);
                if !nodes.is_empty() {
                    runs.push(ShapeTextRun::Math {
                        nodes,
                        display: true,
                        font_size: math_run_size_shape(om),
                        color: math_run_color_shape(om, theme_colors),
                    });
                }
            }
        }
        "AlternateContent" => {
            if let Some(sel) =
                ooxml_common::mce::select_alternate_content(node, &xlsx_understands_ns)
            {
                for c in sel.children().filter(|n| n.is_element()) {
                    push_math_runs_shape(c, theme_colors, runs);
                }
            }
        }
        // a14:m wrapper (local name "m") holds an m:oMathPara / m:oMath.
        "m" => {
            for c in node.children().filter(|n| n.is_element()) {
                push_math_runs_shape(c, theme_colors, runs);
            }
        }
        _ => {}
    }
}

pub(crate) fn parse_tx_body(
    tx_body: &roxmltree::Node,
    theme_colors: &[String],
) -> Option<ShapeText> {
    // `<a:bodyPr>` shared grammar (anchor / wrap / insets / autofit) via
    // ooxml_common::text::parse_body_pr over the bare ECMA-376 §21.1.2.1.1
    // defaults — xlsx has no theme objectDefaults / inheritance layer, so
    // BodyPrDefaults::spec() is the whole fallback (anchor `t`, wrap `square`,
    // autofit `none`, insets 91440/45720 EMU). `vert` is parsed but xlsx does
    // not model it. When there is no `<a:bodyPr>` the spec defaults apply.
    let body = tx_body
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "bodyPr")
        .map(|bp| {
            ooxml_common::text::parse_body_pr(bp, &ooxml_common::text::BodyPrDefaults::spec())
        })
        .unwrap_or_else(|| {
            let d = ooxml_common::text::BodyPrDefaults::spec();
            ooxml_common::text::BodyPr {
                anchor: d.anchor,
                wrap: d.wrap,
                vert: d.vert,
                l_ins: d.l_ins,
                t_ins: d.t_ins,
                r_ins: d.r_ins,
                b_ins: d.b_ins,
                auto_fit: d.auto_fit,
                font_scale: None,
                ln_spc_reduction: None,
            }
        });
    let anchor = body.anchor;
    let wrap = body.wrap;
    // Text insets (EMU), §21.1.2.1.1. Emitted always (spec default when the
    // attribute is absent) so the renderer stops using empirical padding.
    let l_ins = body.l_ins;
    let t_ins = body.t_ins;
    let r_ins = body.r_ins;
    let b_ins = body.b_ins;
    // Autofit (spAutoFit / normAutofit / noAutofit); normAutofit also carries the
    // stored fontScale / lnSpcReduction (1000ths of a percent → fraction).
    let auto_fit = body.auto_fit;
    let font_scale = body.font_scale;
    let ln_spc_reduction = body.ln_spc_reduction;
    let mut paragraphs: Vec<ShapeParagraph> = Vec::new();
    for c in tx_body.children().filter(|n| n.is_element()) {
        match c.tag_name().name() {
            "bodyPr" => {}
            "p" => {
                let mut align = String::from("l");
                let mut rtl = false;
                let mut mar_l: Option<i64> = None;
                let mut mar_r: Option<i64> = None;
                let mut indent: Option<i64> = None;
                let mut space_line: Option<SpaceLine> = None;
                let mut runs: Vec<ShapeTextRun> = Vec::new();
                for pc in c.children().filter(|n| n.is_element()) {
                    match pc.tag_name().name() {
                        "pPr" => {
                            if let Some(a) = pc.attribute("algn") {
                                align = a.to_string();
                            }
                            // ECMA-376 §21.1.2.2.7 `<a:pPr@rtl>` — right-to-left
                            // paragraph. Default false (left-to-right).
                            rtl = pc
                                .attribute("rtl")
                                .map(|v| v == "1" || v == "true")
                                .unwrap_or(false);
                            // ECMA-376 §21.1.2.2.7 (`CT_TextParagraphProperties`)
                            // direct indent attributes (EMU). `indent` may be
                            // negative (hanging). Direct-attribute-only — xlsx
                            // text boxes have no lstStyle/level cascade.
                            mar_l = pc.attribute("marL").and_then(|v| v.parse().ok());
                            mar_r = pc.attribute("marR").and_then(|v| v.parse().ok());
                            indent = pc.attribute("indent").and_then(|v| v.parse().ok());
                            // ECMA-376 §21.1.2.2.5 `<a:lnSpc>`: spcPct is a
                            // percentage of the natural single line; spcPts is an
                            // absolute per-line height (raw @val is hundredths of
                            // a point → divide by 100). Shared with pptx.
                            if let Some(ln_spc) = pc
                                .children()
                                .find(|n| n.is_element() && n.tag_name().name() == "lnSpc")
                            {
                                space_line = ooxml_common::text::parse_lnspc(ln_spc);
                            }
                        }
                        "r" => {
                            // Run text + run-level formatting.
                            let mut text = String::new();
                            let mut bold = false;
                            let mut italic = false;
                            let mut size: f64 = 0.0;
                            let mut color: Option<String> = None;
                            let mut font_face: Option<String> = None;
                            // ECMA-376 §21.1.2.3.1 `<a:ea>` / `<a:cs>` typefaces.
                            // Parsed alongside `<a:latin>` so the renderer can
                            // floor the line box by a tabled East-Asian face
                            // (e.g. Meiryo set only on `<a:ea>`).
                            let mut font_face_ea: Option<String> = None;
                            let mut font_face_cs: Option<String> = None;
                            for rc in pc.children().filter(|n| n.is_element()) {
                                match rc.tag_name().name() {
                                    "rPr" => {
                                        bold = rc.attribute("b").map(|v| v == "1").unwrap_or(false);
                                        italic =
                                            rc.attribute("i").map(|v| v == "1").unwrap_or(false);
                                        size = rc
                                            .attribute("sz")
                                            .and_then(|s| s.parse::<f64>().ok())
                                            .map(|v| v / 100.0)
                                            .unwrap_or(0.0);
                                        for rpc in rc.children().filter(|n| n.is_element()) {
                                            match rpc.tag_name().name() {
                                                "solidFill" => {
                                                    color = parse_solid_fill(&rpc, theme_colors);
                                                }
                                                "latin" => {
                                                    font_face =
                                                        rpc.attribute("typeface").map(String::from);
                                                }
                                                "ea" => {
                                                    font_face_ea =
                                                        rpc.attribute("typeface").map(String::from);
                                                }
                                                "cs" => {
                                                    font_face_cs =
                                                        rpc.attribute("typeface").map(String::from);
                                                }
                                                _ => {}
                                            }
                                        }
                                    }
                                    "t" => {
                                        if let Some(t) = rc.text() {
                                            text.push_str(t);
                                        }
                                    }
                                    _ => {}
                                }
                            }
                            if !text.is_empty() {
                                runs.push(ShapeTextRun::Text {
                                    text,
                                    bold,
                                    italic,
                                    size,
                                    color,
                                    font_face,
                                    font_face_ea,
                                    font_face_cs,
                                });
                            }
                        }
                        "br" => {
                            // Soft line break (`<a:br>`).
                            runs.push(ShapeTextRun::Break);
                        }
                        // OMML equations (ECMA-376 §22.1) embedded in shape text:
                        // bare `m:oMath` / `m:oMathPara`, the `a14:m` wrapper
                        // (local name "m"), or `mc:AlternateContent`.
                        "oMath" | "oMathPara" | "AlternateContent" | "m" => {
                            push_math_runs_shape(pc, theme_colors, &mut runs);
                        }
                        _ => {}
                    }
                }
                if !runs.is_empty() {
                    paragraphs.push(ShapeParagraph {
                        align,
                        rtl,
                        mar_l,
                        mar_r,
                        indent,
                        space_line,
                        runs,
                    });
                }
            }
            _ => {}
        }
    }
    if paragraphs.is_empty() {
        None
    } else {
        Some(ShapeText {
            anchor,
            wrap,
            auto_fit,
            font_scale,
            ln_spc_reduction,
            l_ins,
            t_ins,
            r_ins,
            b_ins,
            paragraphs,
        })
    }
}

/// Parse the adjust handles from a `<a:prstGeom>`'s `<a:avLst>` into an
/// `adj1..adj8`-ordered vector. ECMA-376 §19.5.31.3 / §20.1.9.5: each
/// `<a:gd name="adjN" fmla="val X"/>` supplies one handle. Matching is by name
/// (`adj`/`adj1`, `adj2`, …) with a positional fallback, mirroring the pptx
/// parser so the shared `renderPresetShape` engine receives an identical shape.
/// Trailing `None`s are trimmed so a plain rect/ellipse yields an empty vec.
fn parse_preset_adj(prst_geom: &roxmltree::Node) -> Vec<Option<f64>> {
    let av = prst_geom
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "avLst");
    let Some(av) = av else {
        return Vec::new();
    };
    let gd_nodes: Vec<roxmltree::Node> = av
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "gd")
        .collect();
    if gd_nodes.is_empty() {
        return Vec::new();
    }

    let parse_val = |gd: &roxmltree::Node| -> Option<f64> {
        gd.attribute("fmla")
            .and_then(|f| f.strip_prefix("val "))
            .and_then(|s| s.trim().parse::<f64>().ok())
    };

    // When the handles carry `name` attributes, match strictly by name so a
    // missing adjN leaves a `None` gap (its index still aligns with the engine's
    // adj1..adj8 expectations). Positional fallback is used only for legacy
    // unnamed `<a:gd>` lists.
    let named = gd_nodes.iter().any(|n| n.attribute("name").is_some());
    let mut out: Vec<Option<f64>> = (1..=8)
        .map(|i| {
            let by_name = gd_nodes.iter().find(|n| {
                let name = n.attribute("name");
                if i == 1 {
                    matches!(name, Some("adj") | Some("adj1"))
                } else {
                    name == Some(&format!("adj{i}")[..])
                }
            });
            if named {
                by_name.and_then(parse_val)
            } else {
                gd_nodes.get(i - 1).and_then(parse_val)
            }
        })
        .collect();

    while matches!(out.last(), Some(None)) {
        out.pop();
    }
    out
}

pub(crate) fn parse_sp_geom(
    sp_pr: &roxmltree::Node,
    shape_width: f64,
    shape_height: f64,
) -> Option<ShapeGeom> {
    for c in sp_pr.children().filter(|n| n.is_element()) {
        match c.tag_name().name() {
            "prstGeom" => {
                return Some(ShapeGeom::Preset {
                    name: c.attribute("prst").unwrap_or("rect").to_string(),
                    adj: parse_preset_adj(&c),
                });
            }
            "custGeom" => {
                let paths = parse_custom_paths(c, shape_width, shape_height);
                return Some(ShapeGeom::Custom { paths });
            }
            _ => {}
        }
    }
    None
}

/// Recursively walk an `xdr:grpSp` / `xdr:sp` tree, chaining coordinate
/// transforms, and push leaf shapes (normalized to [0,1] of `root_ext`) into
/// `out`.
// The arguments are the running coordinate-transform state threaded through the
// recursion (root offset/extent, accumulated scale/translation, theme + rels);
// bundling them into a struct would only move the same fields elsewhere without
// improving clarity.
//
// `depth` bounds the nested-`<xdr:grpSp>` recursion so a hand-crafted, thousands-
// deep group nest cannot overflow the fixed WASM stack and trap the parse. Past
// the shared limit the group's children are dropped and the rest of the drawing
// still parses (graceful degradation). See `ooxml_common::depth`.
#[allow(clippy::too_many_arguments)]
pub(crate) fn collect_shapes(
    node: &roxmltree::Node,
    root_off_x: f64,
    root_off_y: f64,
    root_ext_x: f64,
    root_ext_y: f64,
    // cumulative Annex L transform from current local coords into root coords
    transform: DrawingGroupTransform,
    theme_colors: &[String],
    theme_format_scheme: &ThemeFormatScheme,
    rid_urls: &HashMap<String, String>,
    out: &mut Vec<ShapeInfo>,
    depth: DepthGuard,
) {
    for child in node.children().filter(|n| n.is_element()) {
        let tag = child.tag_name().name();
        // §20.1.2.2.8 — a shape/pic/group whose own `<xdr:cNvPr hidden="1">`
        // marks it hidden is not rendered. Skipping a `grpSp` here elides its
        // whole subtree (the recursion below never runs for it).
        if matches!(tag, "sp" | "pic" | "grpSp" | "cxnSp") && xdr_node_hidden(&child) {
            continue;
        }
        if tag == "grpSp" {
            // Bound the group nesting: once the shared depth limit is reached,
            // drop this group's subtree rather than recurse into it, so a
            // pathologically deep `<xdr:grpSp>` nest cannot overflow the stack.
            let Some(child_depth) = depth.descend() else {
                continue;
            };
            // Nested grpSp: compose the transform by the group's own xfrm.
            let grp_sp_pr = child
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "grpSpPr");
            let xfrm = grp_sp_pr
                .and_then(|n| {
                    n.children()
                        .find(|c| c.is_element() && c.tag_name().name() == "xfrm")
                })
                .as_ref()
                .and_then(parse_xfrm);
            let child_transform = if let Some(x) = xfrm {
                // Without chOff/chExt, DrawingML defines no distinct child
                // coordinate system; using ext as chExt keeps scale at 1 while
                // still applying the group's offset.
                let child_ext_x = if x.has_ch { x.ch_ext_x } else { x.ext_x };
                let child_ext_y = if x.has_ch { x.ch_ext_y } else { x.ext_y };
                let child_off_x = if x.has_ch { x.ch_off_x } else { 0.0 };
                let child_off_y = if x.has_ch { x.ch_off_y } else { 0.0 };
                transform.compose_group(DrawingGroupSpec {
                    off_x: x.off_x,
                    off_y: x.off_y,
                    ext_x: x.ext_x,
                    ext_y: x.ext_y,
                    child_off_x,
                    child_off_y,
                    child_ext_x,
                    child_ext_y,
                    rotation_degrees: x.rot,
                    flip_h: x.flip_h,
                    flip_v: x.flip_v,
                })
            } else {
                transform
            };
            collect_shapes(
                &child,
                root_off_x,
                root_off_y,
                root_ext_x,
                root_ext_y,
                child_transform,
                theme_colors,
                theme_format_scheme,
                rid_urls,
                out,
                child_depth,
            );
        } else if tag == "sp" {
            let sp_pr = child
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "spPr");
            let Some(sp_pr) = sp_pr else {
                continue;
            };
            let xfrm_node = sp_pr
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "xfrm");
            let Some(xfrm_n) = xfrm_node else {
                continue;
            };
            let Some(xfrm) = parse_xfrm(&xfrm_n) else {
                continue;
            };
            // Shape rect in root coords. Use the cross-format DrawingML mapper
            // so a quarter-turned leaf under non-uniform group scale keeps its
            // transformed centre and swaps the applicable scale axes.
            let root_rect = transform.apply_rect(DrawingRect {
                x: xfrm.off_x,
                y: xfrm.off_y,
                width: xfrm.ext_x,
                height: xfrm.ext_y,
                rotation_degrees: xfrm.rot,
                flip_h: xfrm.flip_h,
                flip_v: xfrm.flip_v,
            });
            let root_x = root_rect.x;
            let root_y = root_rect.y;
            let root_w = root_rect.width;
            let root_h = root_rect.height;

            // Normalize to [0,1] of root ext
            if root_ext_x == 0.0 || root_ext_y == 0.0 {
                continue;
            }
            let nx = (root_x - root_off_x) / root_ext_x;
            let ny = (root_y - root_off_y) / root_ext_y;
            let nw = root_w / root_ext_x;
            let nh = root_h / root_ext_y;

            let geom = parse_sp_geom(&sp_pr, xfrm.ext_x, xfrm.ext_y);
            let Some(geom) = geom else {
                continue;
            };

            // Fill
            let mut fill: Option<ShapeFill> = None;
            let mut has_direct_fill = false;
            let mut has_no_fill = false;
            let fill_resolver = XlsxSchemeResolver { theme_colors };
            for c in sp_pr.children().filter(|n| n.is_element()) {
                match c.tag_name().name() {
                    "solidFill" | "gradFill" | "pattFill" => {
                        has_direct_fill = true;
                        fill = parse_xlsx_drawingml_fill(c, &fill_resolver);
                    }
                    "noFill" => {
                        has_direct_fill = true;
                        has_no_fill = true;
                        fill = None;
                    }
                    _ => {}
                }
            }

            // <xdr:style> drives fallbacks: <a:fillRef> supplies a fill when
            // <xdr:spPr> didn't, and <a:fontRef> supplies the run-default text
            // color (ECMA-376 §20.5.2.30 `<xdr:style>`). Real-world text boxes
            // saved by Excel often leave `<xdr:spPr>` without `<a:solidFill>`
            // and rely on the style's fillRef + fontRef pair (e.g. accent1
            // background + lt1 white text). We resolve scheme colors here
            // against the workbook theme and apply them as fallbacks.
            let style_node = child
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "style");
            let style_fill = style_node
                .as_ref()
                .and_then(|s| {
                    s.children()
                        .find(|n| n.is_element() && n.tag_name().name() == "fillRef")
                })
                .and_then(|n| resolve_xlsx_fill_ref(n, theme_colors, theme_format_scheme));
            let style_text_color = style_node
                .as_ref()
                .and_then(|s| {
                    s.children()
                        .find(|n| n.is_element() && n.tag_name().name() == "fontRef")
                })
                .and_then(|n| parse_solid_fill(&n, theme_colors));
            if !has_direct_fill && !has_no_fill {
                fill = style_fill;
            }
            let fill_color = fill.as_ref().and_then(xlsx_fill_fallback_color);

            // CT_ShapeStyle supplies a complete line recipe. CT_LineProperties
            // is a component-wise cascade: local `a:ln@w`, paint, dash, cap,
            // join and ends each override only the corresponding theme value
            // (ECMA-376 §20.1.4.2.19, §20.1.2.2.24). This matters even for
            // XLSX's current solid-stroke wire: a local width must not discard
            // the inherited paint, and a fixed theme color needs no lnRef child.
            let style_line = style_node
                .as_ref()
                .and_then(|style| {
                    style
                        .children()
                        .find(|n| n.is_element() && n.tag_name().name() == "lnRef")
                })
                .and_then(|line_ref| {
                    resolve_xlsx_line_ref(line_ref, theme_colors, theme_format_scheme)
                });
            let local_line = sp_pr
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "ln")
                .map(|line| {
                    ooxml_common::line::parse_line_properties(
                        line,
                        &XlsxSchemeResolver { theme_colors },
                        ooxml_common::color::TintMode::PowerPointLinear,
                    )
                });
            let effective_line = match (local_line, style_line) {
                (Some(local), Some(style)) => Some(local.with_fallback(&style)),
                (Some(local), None) => Some(local),
                (None, style) => style,
            };
            let stroke = effective_line
                .as_ref()
                .map(xlsx_line_wire_properties)
                .unwrap_or_default();

            // Text body (txBox shapes carry visible text inside
            // `<xdr:txBody>`; non-textbox shapes may also have one).
            let mut text = child
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "txBody")
                .and_then(|tb| parse_tx_body(&tb, theme_colors));
            if let (Some(t), Some(default_color)) = (text.as_mut(), style_text_color) {
                for p in t.paragraphs.iter_mut() {
                    for r in p.runs.iter_mut() {
                        match r {
                            ShapeTextRun::Text { color, .. } | ShapeTextRun::Math { color, .. } => {
                                if color.is_none() {
                                    *color = Some(default_color.clone());
                                }
                            }
                            ShapeTextRun::Break => {}
                        }
                    }
                }
            }

            out.push(ShapeInfo {
                z_order: child.range().start as u64,
                x: nx,
                y: ny,
                w: nw,
                h: nh,
                rot: root_rect.rotation_degrees,
                flip_h: root_rect.flip_h,
                flip_v: root_rect.flip_v,
                fill_color,
                fill,
                stroke_color: stroke.color,
                stroke_width: stroke.width,
                stroke_fill: stroke.fill,
                stroke_dash_style: stroke.dash_style,
                stroke_custom_dash: stroke.custom_dash,
                stroke_line_cap: stroke.line_cap,
                stroke_line_join: stroke.line_join,
                stroke_miter_limit: stroke.miter_limit,
                stroke_alignment: stroke.alignment,
                stroke_cmpd: stroke.cmpd,
                stroke_head_end: stroke.head_end,
                stroke_tail_end: stroke.tail_end,
                geom,
                text,
            });
        } else if tag == "pic" {
            // `<xdr:pic>` leaf inside a group (ECMA-376 §20.5.2.17). The image
            // binary is resolved via the drawing's .rels file; `rid_urls` maps
            // each r:id to its zip path inside the package.
            let sp_pr = child
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "spPr");
            let Some(sp_pr) = sp_pr else {
                continue;
            };
            let xfrm_node = sp_pr
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "xfrm");
            let Some(xfrm_n) = xfrm_node else {
                continue;
            };
            let Some(xfrm) = parse_xfrm(&xfrm_n) else {
                continue;
            };
            // Resolve `<a:blip>`. The raster fallback rides in `r:embed`; the
            // vector original (Microsoft 2016 svgBlip extension, MS-ODRAWXML) is
            // nested in `<a:blip><a:extLst>`. `rid_urls` maps every media rId
            // (raster *and* svg) to its zip path inside the package.
            let blip = child
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "blip");
            let pic_rid = blip.as_ref().and_then(blip_embed_rid);
            let svg_rid = blip.and_then(svg_blip_rid);

            let svg_image_path = svg_rid.as_deref().and_then(|r| rid_urls.get(r)).cloned();
            let raster_path = pic_rid.as_deref().and_then(|r| rid_urls.get(r)).cloned();

            // `<a:srcRect>` crop, `<a:alphaModFix>` opacity and `<a:duotone>`
            // recolour all live in the leaf's `<xdr:blipFill>` (§20.1.8.55 /
            // §20.1.8.6 / §20.1.8.23).
            let leaf_blip_fill = child
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "blipFill");
            let src_rect = leaf_blip_fill.and_then(parse_src_rect);
            let alpha = leaf_blip_fill.and_then(parse_blip_alpha);
            let duotone = leaf_blip_fill.and_then(|bf| {
                parse_blip_duotone(
                    bf,
                    &XlsxSchemeResolver { theme_colors },
                    ooxml_common::color::TintMode::PowerPointLinear,
                )
            });

            // Prefer the raster as `image_path`; fall back to the SVG when no
            // raster is embedded so an svg-only leaf is never dropped. Drop only
            // when neither resolves. ECMA-376 §20.1.8.13 + MS-ODRAWXML svgBlip.
            let image_path = match (raster_path, svg_image_path.as_ref()) {
                (Some(raster), _) => raster,
                (None, Some(svg)) => svg.clone(),
                (None, None) => continue,
            };
            let mime_type = mime_from_ext(&image_path).to_string();

            let root_rect = transform.apply_rect(DrawingRect {
                x: xfrm.off_x,
                y: xfrm.off_y,
                width: xfrm.ext_x,
                height: xfrm.ext_y,
                rotation_degrees: xfrm.rot,
                flip_h: xfrm.flip_h,
                flip_v: xfrm.flip_v,
            });
            let root_x = root_rect.x;
            let root_y = root_rect.y;
            let root_w = root_rect.width;
            let root_h = root_rect.height;
            if root_ext_x == 0.0 || root_ext_y == 0.0 {
                continue;
            }
            let nx = (root_x - root_off_x) / root_ext_x;
            let ny = (root_y - root_off_y) / root_ext_y;
            let nw = root_w / root_ext_x;
            let nh = root_h / root_ext_y;
            if nw <= 0.0 || nh <= 0.0 {
                continue;
            }

            out.push(ShapeInfo {
                z_order: child.range().start as u64,
                x: nx,
                y: ny,
                w: nw,
                h: nh,
                rot: root_rect.rotation_degrees,
                flip_h: root_rect.flip_h,
                flip_v: root_rect.flip_v,
                fill_color: None,
                fill: None,
                stroke_color: None,
                stroke_width: 0,
                stroke_fill: None,
                stroke_dash_style: None,
                stroke_custom_dash: Vec::new(),
                stroke_line_cap: None,
                stroke_line_join: None,
                stroke_miter_limit: None,
                stroke_alignment: None,
                stroke_cmpd: None,
                stroke_head_end: None,
                stroke_tail_end: None,
                geom: ShapeGeom::Image {
                    image_path,
                    mime_type,
                    svg_image_path,
                    src_rect,
                    alpha,
                    duotone,
                },
                text: None,
            });
        }
        // Ignore `xdr:cxnSp` / text-only elements for this minimal pass.
    }
}

pub(crate) fn parse_shape_anchors(
    drawing_xml: &str,
    theme_colors: &[String],
    theme_format_scheme: &ThemeFormatScheme,
    rid_urls: &HashMap<String, String>,
) -> Vec<ShapeAnchor> {
    let Ok(doc) = parse_guarded(drawing_xml) else {
        return Vec::new();
    };
    let mut anchors: Vec<ShapeAnchor> = Vec::new();

    for anchor in doc.descendants() {
        let anchor_tag = anchor.tag_name().name();
        // ECMA-376 §20.5.2: a drawing shape is anchored with either
        // `twoCellAnchor` (from+to) or `oneCellAnchor` (from + a saved `<ext>`
        // size). Excel authors equation text boxes as oneCellAnchor, so we must
        // accept both or those shapes (and their math) are silently dropped.
        if (anchor_tag != "twoCellAnchor" && anchor_tag != "oneCellAnchor")
            || !is_xdr_ns(anchor.tag_name().namespace())
        {
            continue;
        }

        // Parse from/to anchor rect (shared between grpSp and stand-alone sp paths)
        let (mut from_col, mut from_col_off, mut from_row, mut from_row_off) =
            (0u32, 0i64, 0u32, 0i64);
        let (mut to_col, mut to_col_off, mut to_row, mut to_row_off) = (0u32, 0i64, 0u32, 0i64);
        // ECMA-376 §20.5.2.33 `twoCellAnchor@editAs` — see ImageAnchor parsing
        // path for semantics. `"oneCell"` instructs the renderer to preserve
        // the group's saved EMU size instead of resizing with the cell rect.
        // oneCellAnchor has no `<to>`; its size is the saved EMU `<ext>` (which
        // equals the shape's own xfrm ext), positioned at `<from>` — i.e. the
        // "oneCell" (move-but-don't-size) semantics the renderer already
        // implements via nativeExtCx/Cy. Force editAs accordingly so the
        // renderer sizes from the saved ext rather than a (missing) `to` rect.
        let edit_as = if anchor_tag == "oneCellAnchor" {
            Some("oneCell".to_string())
        } else {
            anchor.attribute("editAs").map(|s| s.to_string())
        };
        let native_ext_cx: i64;
        let native_ext_cy: i64;
        for c in anchor.children() {
            if !c.is_element() {
                continue;
            }
            if c.tag_name().name() == "from" || c.tag_name().name() == "to" {
                let is_from = c.tag_name().name() == "from";
                let mut col: u32 = 0;
                let mut col_off: i64 = 0;
                let mut row: u32 = 0;
                let mut row_off: i64 = 0;
                for cc in c.children() {
                    match (cc.tag_name().name(), cc.text()) {
                        ("col", Some(t)) => col = parse_xdr_col(t),
                        ("colOff", Some(t)) => col_off = t.trim().parse().unwrap_or(0),
                        ("row", Some(t)) => row = parse_xdr_row(t),
                        ("rowOff", Some(t)) => row_off = t.trim().parse().unwrap_or(0),
                        _ => {}
                    }
                }
                if is_from {
                    from_col = col;
                    from_col_off = col_off;
                    from_row = row;
                    from_row_off = row_off;
                } else {
                    to_col = col;
                    to_col_off = col_off;
                    to_row = row;
                    to_row_off = row_off;
                }
            }
        }

        // Excel wraps a modern shape in an anchor-level `<mc:AlternateContent>`
        // (`<mc:Choice Requires="…">` = the live shape, `<mc:Fallback>` = a
        // legacy fallback). Per ECMA-376 Part 3 §9.3, unwrap the first understood
        // Choice or fall through to the Fallback so the grpSp/sp/pic lookups below
        // see the selected shape. (Equations inside the shape text body are handled
        // separately by `push_math_runs_shape`; this is the shape-level wrapper.)
        let content = anchor
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "AlternateContent")
            .and_then(|ac| ooxml_common::mce::select_alternate_content(ac, &xlsx_understands_ns))
            .unwrap_or(anchor);

        // Two top-level layouts ECMA-376 allows under <xdr:twoCellAnchor>:
        //   (a) <xdr:grpSp> wrapping a tree of nested groups + leaves; and
        //   (b) a single <xdr:sp> / <xdr:pic> directly under the anchor
        //       (no grouping wrapper). The grpSp path uses the group's xfrm
        //       to define the anchor's drawing-coord system; the stand-alone
        //       path treats the shape as filling 100 % of the anchor rect.
        let mut shapes: Vec<ShapeInfo> = Vec::new();
        if let Some(grp) = content
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "grpSp")
        {
            // §20.1.2.2.8 — a hidden top-level group elides its whole subtree.
            if xdr_node_hidden(&grp) {
                continue;
            }
            let grp_sp_pr = grp
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "grpSpPr");
            let xfrm = grp_sp_pr
                .and_then(|n| {
                    n.children()
                        .find(|c| c.is_element() && c.tag_name().name() == "xfrm")
                })
                .as_ref()
                .and_then(parse_xfrm);
            let Some(root) = xfrm else {
                continue;
            };
            if !root.has_ch || root.ch_ext_x == 0.0 || root.ch_ext_y == 0.0 {
                continue;
            }

            // Top-level grpSp ext is the group's saved on-sheet EMU size —
            // authoritative when editAs="oneCell".
            native_ext_cx = root.ext_x as i64;
            native_ext_cy = root.ext_y as i64;

            // Map child coords → root coords through the same shared DrawingML
            // transform contract used by DOCX and PPTX.
            let root_transform = DrawingGroupTransform::from_group(DrawingGroupSpec {
                off_x: root.off_x,
                off_y: root.off_y,
                ext_x: root.ext_x,
                ext_y: root.ext_y,
                child_off_x: root.ch_off_x,
                child_off_y: root.ch_off_y,
                child_ext_x: root.ch_ext_x,
                child_ext_y: root.ch_ext_y,
                rotation_degrees: root.rot,
                flip_h: root.flip_h,
                flip_v: root.flip_v,
            });

            collect_shapes(
                &grp,
                root.off_x,
                root.off_y,
                root.ext_x,
                root.ext_y,
                root_transform,
                theme_colors,
                theme_format_scheme,
                rid_urls,
                &mut shapes,
                DepthGuard::root(),
            );
        } else if let Some(sp) = content.children().find(|n| {
            if !n.is_element() {
                return false;
            }
            let t = n.tag_name().name();
            // A standalone `<xdr:sp>` always renders through the shape path. A
            // standalone `<xdr:pic>` under a `twoCellAnchor`, however, is ALSO a
            // plain image anchor that `parse_drawing_anchors` already emits into
            // `ws.images` — WITH its `<a:srcRect>` crop, svgBlip vector original,
            // and `editAs` size. Capturing it here too would draw the picture a
            // second time as an uncropped full-image shape, overwriting the
            // cropped anchor draw (the sample-27 regression). So a `twoCellAnchor`
            // pic is owned solely by `ws.images`; we only keep a standalone pic
            // for `oneCellAnchor`, which `parse_drawing_anchors` does not handle.
            t == "sp" || (t == "pic" && anchor_tag == "oneCellAnchor")
        }) {
            // §20.1.2.2.8 — a hidden stand-alone shape/pic is not rendered.
            if xdr_node_hidden(&sp) {
                continue;
            }
            // Stand-alone sp/pic: the shape's own xfrm gives its absolute EMU
            // rect, but for our rendering pipeline the anchor's from/to
            // already defines the on-sheet rect, and the leaf occupies it
            // 100 %. Build a synthetic root coord-system whose origin matches
            // the shape's xfrm so collect_shapes normalizes the leaf to (0,0)
            // (1,1).
            let sp_pr = sp
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "spPr");
            let Some(sp_pr_node) = sp_pr else {
                continue;
            };
            let xfrm_node = sp_pr_node
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "xfrm");
            let Some(xfrm_n) = xfrm_node else {
                continue;
            };
            let Some(xfrm) = parse_xfrm(&xfrm_n) else {
                continue;
            };
            if xfrm.ext_x == 0.0 || xfrm.ext_y == 0.0 {
                continue;
            }
            native_ext_cx = xfrm.ext_x as i64;
            native_ext_cy = xfrm.ext_y as i64;
            collect_shapes(
                &content,
                xfrm.off_x,
                xfrm.off_y,
                xfrm.ext_x,
                xfrm.ext_y,
                DrawingGroupTransform::IDENTITY,
                theme_colors,
                theme_format_scheme,
                rid_urls,
                &mut shapes,
                DepthGuard::root(),
            );
        } else {
            continue;
        }

        if shapes.is_empty() {
            continue;
        }

        anchors.push(ShapeAnchor {
            from_col,
            from_col_off,
            from_row,
            from_row_off,
            to_col,
            to_col_off,
            to_row,
            to_row_off,
            edit_as,
            native_ext_cx,
            native_ext_cy,
            shapes,
        });
    }
    anchors
}

pub(crate) fn load_sheet_shape_groups(
    archive: &mut crate::XlsxZip,
    sheet_path: &str,
    theme_colors: &[String],
    theme_format_scheme: &ThemeFormatScheme,
) -> Vec<ShapeAnchor> {
    let Some((sheet_dir, sheet_file)) = sheet_path.rsplit_once('/') else {
        return Vec::new();
    };
    let sheet_rels_path = format!("xl/{}/_rels/{}.rels", sheet_dir, sheet_file);
    let Ok(sheet_rels_xml) = read_zip_string(archive, &sheet_rels_path) else {
        return Vec::new();
    };
    let Ok(rels_doc) = parse_guarded(&sheet_rels_xml) else {
        return Vec::new();
    };
    let mut drawing_targets: Vec<String> = Vec::new();
    for rel in rels_doc
        .root_element()
        .children()
        .filter(|n| n.is_element())
    {
        if rel.attribute("Type").unwrap_or("").ends_with("/drawing") {
            if let Some(t) = rel.attribute("Target") {
                drawing_targets.push(t.to_string());
            }
        }
    }
    let mut all: Vec<ShapeAnchor> = Vec::new();
    for target in drawing_targets {
        let drawing_path = resolve_zip_path(&format!("xl/{}", sheet_dir), &target);
        let Ok(drawing_xml) = read_zip_string(archive, &drawing_path) else {
            continue;
        };
        let rid_urls = build_drawing_rid_urls(archive, &drawing_path);
        all.extend(parse_shape_anchors(
            &drawing_xml,
            theme_colors,
            theme_format_scheme,
            &rid_urls,
        ));
    }
    all
}

/// Build a `HashMap<rId, zip-path>` for every image (png/jpg/…) target in
/// a drawing's `.rels` file. Used by `collect_shapes` to resolve `<xdr:pic>`
/// leaves inside a group. Mirrors the logic in `parse_drawing_anchors`; the
/// renderer fetches each referenced image's bytes lazily via `extract_image`,
/// so this maps to zip paths rather than inlining base64.
pub(crate) fn build_drawing_rid_urls(
    archive: &mut crate::XlsxZip,
    drawing_path: &str,
) -> HashMap<String, String> {
    let Some((drawing_dir, drawing_file)) = drawing_path.rsplit_once('/') else {
        return HashMap::new();
    };
    let rels_path = format!("{}/_rels/{}.rels", drawing_dir, drawing_file);
    let rels = read_zip_string(archive, &rels_path)
        .ok()
        .map(|xml| parse_rels_map(&xml))
        .unwrap_or_default();

    let mut result: HashMap<String, String> = HashMap::new();
    for (rid, target) in rels {
        let lower = target.to_lowercase();
        // `.svg` is included so the Microsoft svgBlip extension's vector original
        // is collected (it was previously excluded, dropping svg-only pictures).
        if !(lower.ends_with(".png")
            || lower.ends_with(".jpg")
            || lower.ends_with(".jpeg")
            || lower.ends_with(".gif")
            || lower.ends_with(".bmp")
            || lower.ends_with(".webp")
            || lower.ends_with(".svg"))
        {
            continue;
        }
        let media_path = resolve_zip_path(drawing_dir, &target);
        // Only emit the path when the entry actually resolves (preserves the
        // previous behavior of dropping rIds whose bytes are missing).
        // `index_for_name` checks the central directory only — no inflate, unlike
        // the former `read_zip_bytes` which decompressed the entry to discard it.
        if archive.index_for_name(&media_path).is_some() {
            result.insert(rid, media_path);
        }
    }
    result
}

/// Given a sheet path (e.g. "worksheets/sheet1.xml"), locate and parse
/// its drawing(s), and return all image anchors found.
pub(crate) fn load_sheet_images(
    archive: &mut crate::XlsxZip,
    sheet_path: &str, // e.g. "worksheets/sheet1.xml"
    // Positional clrScheme palette, for resolving a picture's `<a:duotone>`
    // effect colours (§20.1.8.23) via the shared DrawingML colour grammar.
    theme_colors: &[String],
) -> Vec<ImageAnchor> {
    // sheet rels path:  xl/worksheets/_rels/sheet1.xml.rels
    let Some((sheet_dir, sheet_file)) = sheet_path.rsplit_once('/') else {
        return Vec::new();
    };
    let sheet_rels_path = format!("xl/{}/_rels/{}.rels", sheet_dir, sheet_file);
    let Ok(sheet_rels_xml) = read_zip_string(archive, &sheet_rels_path) else {
        return Vec::new();
    };

    // Find all drawing relationships
    let Ok(rels_doc) = parse_guarded(&sheet_rels_xml) else {
        return Vec::new();
    };
    let mut drawing_targets: Vec<String> = Vec::new();
    for rel in rels_doc
        .root_element()
        .children()
        .filter(|n| n.is_element())
    {
        let rel_type = rel.attribute("Type").unwrap_or("");
        if rel_type.ends_with("/drawing") {
            if let Some(t) = rel.attribute("Target") {
                drawing_targets.push(t.to_string());
            }
        }
    }
    if drawing_targets.is_empty() {
        return Vec::new();
    }

    let mut all_anchors: Vec<ImageAnchor> = Vec::new();
    for target in drawing_targets {
        // sheet_dir is "worksheets", target typically "../drawings/drawing1.xml"
        // base dir for the drawing = "xl/worksheets" + "../drawings" → "xl/drawings"
        let drawing_path = resolve_zip_path(&format!("xl/{}", sheet_dir), &target);
        let Ok(drawing_xml) = read_zip_string(archive, &drawing_path) else {
            continue;
        };
        // Drawing rels:  xl/drawings/_rels/drawing1.xml.rels
        let Some((drawing_dir, drawing_file)) = drawing_path.rsplit_once('/') else {
            continue;
        };
        let drawing_rels_path = format!("{}/_rels/{}.rels", drawing_dir, drawing_file);
        let drawing_rels = read_zip_string(archive, &drawing_rels_path)
            .ok()
            .map(|xml| parse_rels_map(&xml))
            .unwrap_or_default();

        let mut anchors = parse_drawing_anchors(
            &drawing_xml,
            &drawing_rels,
            drawing_dir,
            archive,
            theme_colors,
        );
        all_anchors.append(&mut anchors);
    }
    all_anchors
}

/// The two-cell `<from>`/`<to>` marker rect of an OLE-object anchor, in the same
/// units the renderer's `ImageAnchor` uses (col/row indices + EMU offsets).
struct AnchorRect {
    from_col: u32,
    from_col_off: i64,
    from_row: u32,
    from_row_off: i64,
    to_col: u32,
    to_col_off: i64,
    to_row: u32,
    to_row_off: i64,
}

/// Parse a `<objectPr><anchor>` (or drawing-style) two-cell marker: `<from>`/
/// `<to>` each carrying `col`/`colOff`/`row`/`rowOff` (CT_Marker, EMU offsets),
/// matched by local name so namespace prefixes don't matter.
fn parse_two_cell_marker(anchor: &roxmltree::Node) -> AnchorRect {
    let (mut from_col, mut from_col_off, mut from_row, mut from_row_off) = (0u32, 0i64, 0u32, 0i64);
    let (mut to_col, mut to_col_off, mut to_row, mut to_row_off) = (0u32, 0i64, 0u32, 0i64);
    for marker in anchor.children().filter(|n| n.is_element()) {
        let is_from = match marker.tag_name().name() {
            "from" => true,
            "to" => false,
            _ => continue,
        };
        let (mut col, mut col_off, mut row, mut row_off) = (0u32, 0i64, 0u32, 0i64);
        for c in marker.children() {
            match (c.tag_name().name(), c.text()) {
                ("col", Some(t)) => col = parse_xdr_col(t),
                ("colOff", Some(t)) => col_off = t.trim().parse().unwrap_or(0),
                ("row", Some(t)) => row = parse_xdr_row(t),
                ("rowOff", Some(t)) => row_off = t.trim().parse().unwrap_or(0),
                _ => {}
            }
        }
        if is_from {
            (from_col, from_col_off, from_row, from_row_off) = (col, col_off, row, row_off);
        } else {
            (to_col, to_col_off, to_row, to_row_off) = (col, col_off, row, row_off);
        }
    }
    AnchorRect {
        from_col,
        from_col_off,
        from_row,
        from_row_off,
        to_col,
        to_col_off,
        to_row,
        to_row_off,
    }
}

/// Parse a legacy VML `<x:ClientData><x:Anchor>` into the two-cell rect. The
/// element's text is an 8-value CSV (`[MS-OI29500] 2.1.639 / VML §14.2`):
/// `LeftCol, LeftOffsetPx, TopRow, TopOffsetPx, RightCol, RightOffsetPx,
/// BottomRow, BottomOffsetPx`. Columns/rows are 0-based cell indices; the four
/// offsets are **pixels** and are converted to EMU via [`EMU_PER_PX_96DPI`].
/// The multiplication saturates instead of wrapping/panicking so a
/// pathologically large offset in a malformed or hostile CSV clamps to
/// `i64::MAX` rather than overflowing.
/// Returns `None` when the CSV is malformed (< 8 numeric fields).
fn parse_vml_client_anchor(client_data: &roxmltree::Node) -> Option<AnchorRect> {
    let anchor = client_data
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "Anchor")?;
    let text = anchor.text()?;
    let vals: Vec<i64> = text
        .split(',')
        .filter_map(|s| s.trim().parse::<i64>().ok())
        .collect();
    if vals.len() < 8 {
        return None;
    }
    Some(AnchorRect {
        from_col: vml_xdr_index(vals[0]),
        from_col_off: vals[1].saturating_mul(EMU_PER_PX_96DPI),
        from_row: vml_xdr_index(vals[2]),
        from_row_off: vals[3].saturating_mul(EMU_PER_PX_96DPI),
        to_col: vml_xdr_index(vals[4]),
        to_col_off: vals[5].saturating_mul(EMU_PER_PX_96DPI),
        to_row: vml_xdr_index(vals[6]),
        to_row_off: vals[7].saturating_mul(EMU_PER_PX_96DPI),
    })
}

/// Locate the OLE preview part inside a parsed legacy VML drawing, keyed by
/// `oleObject@shapeId`. Excel connects the worksheet's `<oleObject shapeId="N">`
/// to a VML `<v:shape id="_x0000_s{N}">`; that shape's `<v:imagedata>` carries
/// the preview image rId (`o:relid`, or `r:id`) which resolves through the *VML
/// part's own* rels (`xl/drawings/_rels/vmlDrawingK.vml.rels`), NOT the sheet
/// rels. Returns the resolved image zip path + the shape's `<x:ClientData>`
/// (for the `<x:Anchor>` fallback) when the preview resolves to an image part.
///
/// Comment (note) shapes in the same VML are `type="#_x0000_t202"` /
/// `ObjectType="Note"` and carry no `<v:imagedata>`; they never match a
/// `shapeId` used by `<oleObject>` and are additionally filtered by the
/// imagedata requirement, so a comments-only VML yields nothing here.
fn vml_ole_preview<'a>(
    vml_doc: &'a roxmltree::Document<'a>,
    vml_rels: &HashMap<String, String>,
    vml_dir: &str,
    shape_id: &str,
    archive: &mut crate::XlsxZip,
) -> Option<(String, roxmltree::Node<'a, 'a>)> {
    let target_id = format!("_x0000_s{shape_id}");
    let shape = vml_doc.descendants().find(|n| {
        n.is_element()
            && n.tag_name().name() == "shape"
            && n.attribute("id") == Some(&target_id[..])
    })?;
    let imagedata = shape
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "imagedata")?;
    // VML `<v:imagedata>` names the preview via `o:relid` (Excel's usual form)
    // or `r:id`; both resolve through the VML part's own rels.
    let rid = imagedata
        .attribute(("urn:schemas-microsoft-com:office:office", "relid"))
        .or_else(|| imagedata.attribute("relid"))
        .or_else(|| {
            attr_ns(
                &imagedata,
                relationships::TRANSITIONAL,
                relationships::STRICT,
                "id",
            )
        })?;
    let target = vml_rels.get(rid)?;
    let media_path = resolve_zip_path(vml_dir, target);
    archive.index_for_name(&media_path)?;
    // Preview parts are metafiles/bitmaps (EMF/WMF/PNG…); require an image MIME
    // so a mis-typed rels target can never feed non-image bytes to the renderer.
    if !mime_from_ext(&media_path).starts_with("image/") {
        return None;
    }
    let client_data = shape
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "ClientData")?;
    Some((media_path, client_data))
}

/// Resolve the sheet's legacy VML drawing part (`<legacyDrawing r:id>`,
/// §18.3.1.55 → the sheet rels `vmlDrawing` relationship), read it, and return
/// its XML + its own rels map. Read once per sheet and only when there is at
/// least one `<oleObject>` to place (so a comments-only sheet pays no extra zip
/// read). `None` when there is no legacyDrawing, the part is missing, or it
/// fails to parse. Whether the VML part carries a `<?xml?>` declaration varies
/// by producer; either way it is well-formed XML, so roxmltree parses it as-is
/// (the declaration-less case is verified against real Excel output — see the
/// `roxmltree_parses_declarationless_vml` test); no preprocessing is needed.
fn load_legacy_vml_drawing(
    doc: &roxmltree::Document,
    sheet_rels: &HashMap<String, String>,
    sheet_dir: &str,
    archive: &mut crate::XlsxZip,
) -> Option<(String, HashMap<String, String>, String)> {
    let legacy = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "legacyDrawing")?;
    let rid = attr_ns(
        &legacy,
        relationships::TRANSITIONAL,
        relationships::STRICT,
        "id",
    )?;
    let target = sheet_rels.get(rid)?;
    let vml_path = resolve_zip_path(sheet_dir, target); // e.g. xl/drawings/vmlDrawing1.vml
    let vml_xml = read_zip_string(archive, &vml_path).ok()?;
    // Directory + file split for the VML part's own rels.
    let (vml_dir, vml_file) = vml_path.rsplit_once('/')?;
    let vml_rels_path = format!("{vml_dir}/_rels/{vml_file}.rels");
    let vml_rels = read_zip_string(archive, &vml_rels_path)
        .ok()
        .map(|xml| parse_rels_map(&xml))
        .unwrap_or_default();
    Some((vml_xml, vml_rels, vml_dir.to_string()))
}

/// Parse worksheet `<oleObjects>` (the collection element, ECMA-376
/// §18.3.1.60) into preview `ImageAnchor`s. An embedded OLE object we can't run
/// still needs a visible on-sheet representation.
///
/// **Where the preview image lives.** Real Excel output stores the on-sheet
/// preview (typically an EMF metafile) in a legacy VML drawing, NOT in
/// `objectPr@r:id`. Each `<oleObject shapeId="N">` is connected to a VML
/// `<v:shape id="_x0000_s{N}"><v:imagedata o:relid>` in the sheet's
/// `vmlDrawingK.vml` (reached via `<legacyDrawing r:id>` §18.3.1.55), whose rId
/// resolves through the *VML part's* own rels — exactly like Word's
/// `<w:object><v:shape><v:imagedata>` path (see docx `parse_object_ole_image`).
/// The parts differ (Word's OLE VML lives inline in the document; Excel's lives
/// in a standalone vmlDrawing keyed by shapeId), so the code is not shared; only
/// the CSS-dimension/VML grammar is conceptually common.
///
/// **Preview-source priority.** For each object:
///   1. If `objectPr@r:id` (§18.3.1.56) resolves to a genuine **image** part,
///      use it. The spec text says this relationship targets the *object data*
///      part, and real Excel does point it at a `.bin`, but a spec-faithful
///      producer *could* emit an image here, so we honour that first and gate it
///      by an `image/*` MIME check (a `.bin` ⇒ `application/octet-stream` is
///      rejected, never fed to the renderer).
///   2. Otherwise fall back to the **vmlDrawing** preview keyed by
///      `oleObject@shapeId` (the real-Excel path).
///
/// An object that resolves neither is skipped (icon-only / data-only /
/// link-only), matching the prior silent-skip.
///
/// **Anchor priority.** When present, the `<objectPr><anchor>` two-cell markers
/// (EMU, exact) win. Otherwise the VML `<x:ClientData><x:Anchor>` 8-value
/// pixel CSV is converted to a two-cell rect. An object with neither is skipped.
///
/// `<oleObject>` is commonly wrapped in `mc:AlternateContent`; we skip any
/// oleObject nested in an `mc:Fallback` so each object contributes exactly one
/// anchor (no double-draw).
///
/// The vmlDrawing is loaded once per sheet and only when at least one
/// `<oleObject>` exists (a comments-only sheet — whose VML holds only
/// `type="#_x0000_t202"` / `ObjectType="Note"` shapes — pays no extra zip read).
pub(crate) fn parse_ole_object_anchors(
    sheet_xml: &str,
    sheet_rels: &HashMap<String, String>,
    sheet_dir: &str,
    archive: &mut crate::XlsxZip,
) -> Vec<ImageAnchor> {
    let Ok(doc) = parse_guarded(sheet_xml) else {
        return Vec::new();
    };

    // Collect the ECMA-376 Part 3 §9.3-selected oleObjects up front so we can
    // decide whether to read the vmlDrawing at all.
    let ole_objects: Vec<roxmltree::Node> = doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "oleObject")
        .filter(|n| ole_object_in_selected_mce_branch(*n))
        .collect();
    if ole_objects.is_empty() {
        return Vec::new();
    }

    // Lazily load the sheet's legacy VML drawing (once) — only reached because
    // there is at least one oleObject to place. `vml` owns the XML string; the
    // parsed document borrows from it, so both live to the end of the function.
    let vml = load_legacy_vml_drawing(&doc, sheet_rels, sheet_dir, archive);
    let vml_parsed = vml
        .as_ref()
        .and_then(|(xml, rels, dir)| parse_guarded(xml).ok().map(|d| (d, rels, dir)));

    let mut anchors: Vec<ImageAnchor> = Vec::new();
    for ole in ole_objects {
        // (1) Spec-faithful producer path: an image-typed `objectPr@r:id`.
        let object_pr = ole
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "objectPr");
        let object_pr_preview = object_pr.and_then(|pr| {
            let rid = attr_ns(
                &pr,
                relationships::TRANSITIONAL,
                relationships::STRICT,
                "id",
            )?;
            let target = sheet_rels.get(rid)?;
            let media_path = resolve_zip_path(sheet_dir, target);
            archive.index_for_name(&media_path)?;
            // §18.3.1.56: this relationship nominally targets the object *data*
            // part. Only genuine image parts enter the picture pipeline; a
            // non-image target (`.bin` ⇒ application/octet-stream) is rejected.
            if !mime_from_ext(&media_path).starts_with("image/") {
                return None;
            }
            Some(media_path)
        });

        // (2) Real-Excel path: the vmlDrawing preview keyed by `shapeId`.
        let shape_id = ole.attribute("shapeId");
        let vml_preview = match (&object_pr_preview, shape_id, &vml_parsed) {
            // objectPr already resolved an image — no need to touch the VML.
            (Some(_), _, _) => None,
            (None, Some(sid), Some((vml_doc, vml_rels, vml_dir))) => {
                vml_ole_preview(vml_doc, vml_rels, vml_dir, sid, archive)
            }
            _ => None,
        };

        let image_path = match (&object_pr_preview, &vml_preview) {
            (Some(p), _) => p.clone(),
            (None, Some((p, _))) => p.clone(),
            (None, None) => continue,
        };
        let mime_type = mime_from_ext(&image_path).to_string();

        // Anchor: `<objectPr><anchor>` (EMU) wins; else the VML `<x:Anchor>`
        // pixel CSV. Skip when neither places the object.
        let rect = object_pr
            .and_then(|pr| {
                pr.children()
                    .find(|n| n.is_element() && n.tag_name().name() == "anchor")
            })
            .map(|a| parse_two_cell_marker(&a))
            .or_else(|| {
                vml_preview
                    .as_ref()
                    .and_then(|(_, client_data)| parse_vml_client_anchor(client_data))
            });
        let Some(rect) = rect else {
            continue;
        };

        anchors.push(ImageAnchor {
            z_order: ole.range().start as u64,
            from_col: rect.from_col,
            from_col_off: rect.from_col_off,
            from_row: rect.from_row,
            from_row_off: rect.from_row_off,
            to_col: rect.to_col,
            to_col_off: rect.to_col_off,
            to_row: rect.to_row,
            to_row_off: rect.to_row_off,
            // OLE previews always place via the two-cell rect. This "twoCell" is
            // a convenience default in the ST_EditAs value space (§20.5.3.3), not
            // a faithful conversion of the anchor's own moveWithCells/
            // sizeWithCells booleans (CT_ObjectAnchor); those live in a different
            // value space and are not mapped here.
            edit_as: Some("twoCell".to_string()),
            native_ext_cx: 0,
            native_ext_cy: 0,
            rotation: None,
            flip_h: None,
            flip_v: None,
            image_path,
            mime_type,
            svg_image_path: None,
            src_rect: None,
            // OLE object previews carry no blip effects.
            alpha: None,
            duotone: None,
        });
    }
    anchors
}

/// True when `ole` (an `<oleObject>`) is not wrapped in `<mc:AlternateContent>`,
/// or lies inside the Choice/Fallback branch that ECMA-376 Part 3 §9.3 selects.
/// Replaces the old blanket "skip anything under a Fallback" rule so an
/// un-understood Choice correctly yields the Fallback's oleObject (issue #787).
fn ole_object_in_selected_mce_branch(ole: roxmltree::Node) -> bool {
    let Some(ac) = ole
        .ancestors()
        .find(|a| a.is_element() && a.tag_name().name() == "AlternateContent")
    else {
        return true; // bare oleObject, not under MCE
    };
    let branch = ole
        .ancestors()
        .find(|a| a.is_element() && matches!(a.tag_name().name(), "Choice" | "Fallback"));
    match (
        ooxml_common::mce::select_alternate_content(ac, &xlsx_understands_ns),
        branch,
    ) {
        (Some(sel), Some(b)) => sel == b,
        _ => false,
    }
}

/// Load a sheet's OLE-object preview images (the `<oleObjects>` collection,
/// §18.3.1.60) as `ImageAnchor`s,
/// wiring the worksheet XML + its rels. Sibling of `load_sheet_images`; the
/// caller appends the result to `ws.images` so previews draw through the same
/// pipeline as ordinary pictures.
pub(crate) fn load_sheet_ole_images(
    archive: &mut crate::XlsxZip,
    sheet_path: &str, // e.g. "worksheets/sheet1.xml"
    sheet_xml: &str,
) -> Vec<ImageAnchor> {
    let Some((sheet_dir, sheet_file)) = sheet_path.rsplit_once('/') else {
        return Vec::new();
    };
    let sheet_rels_path = format!("xl/{}/_rels/{}.rels", sheet_dir, sheet_file);
    let sheet_rels = read_zip_string(archive, &sheet_rels_path)
        .ok()
        .map(|xml| parse_rels_map(&xml))
        .unwrap_or_default();
    if sheet_rels.is_empty() {
        return Vec::new();
    }
    let base_dir = format!("xl/{}", sheet_dir);
    parse_ole_object_anchors(sheet_xml, &sheet_rels, &base_dir, archive)
}

#[cfg(test)]
mod math_tests {
    use super::*;
    use ooxml_common::math::nodes_to_text;

    const NS: &str = r#"xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
        xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math""#;

    /// Excel stores "Insert > Equation" as OMML inside the shared DrawingML
    /// `<xdr:txBody>` grammar — a block equation as `mc:AlternateContent` →
    /// `a14:m` → `m:oMathPara` (ECMA-376 §22.1). The Fallback (a rasterized
    /// picture) must be ignored so the equation isn't double-rendered.
    #[test]
    fn parses_block_math_from_alternatecontent() {
        let xml = format!(
            r#"<xdr:txBody {NS}>
              <a:bodyPr/>
              <a:p>
                <mc:AlternateContent>
                  <mc:Choice Requires="a14"><a14:m>
                    <m:oMathPara><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></m:oMathPara>
                  </a14:m></mc:Choice>
                  <mc:Fallback><a:r><a:t>fallback</a:t></a:r></mc:Fallback>
                </mc:AlternateContent>
              </a:p>
            </xdr:txBody>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let text = parse_tx_body(&doc.root_element(), &[]).expect("txBody parses");
        assert_eq!(text.paragraphs.len(), 1);
        let runs = &text.paragraphs[0].runs;
        assert_eq!(runs.len(), 1, "one math run, fallback ignored");
        match &runs[0] {
            ShapeTextRun::Math { display, nodes, .. } => {
                assert!(*display, "oMathPara → display (block) math");
                assert_eq!(nodes_to_text(nodes), "x");
            }
            other => panic!("expected Math run, got {other:?}"),
        }
    }

    /// Inline math is a bare `a14:m` (local name "m") directly under `a:p`,
    /// holding `m:oMath` (not oMathPara). It extracts as inline (display:false)
    /// and inherits size + colour from the math run's `a:rPr`.
    #[test]
    fn parses_inline_bare_a14m_with_size_and_color() {
        let xml = format!(
            r#"<xdr:txBody {NS}>
              <a:bodyPr/>
              <a:p>
                <a14:m><m:oMath><m:r>
                  <a:rPr sz="2800" i="1"><a:solidFill><a:srgbClr val="7030A0"/></a:solidFill></a:rPr>
                  <m:t>n</m:t>
                </m:r></m:oMath></a14:m>
              </a:p>
            </xdr:txBody>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let text = parse_tx_body(&doc.root_element(), &[]).expect("txBody parses");
        let runs = &text.paragraphs[0].runs;
        assert_eq!(runs.len(), 1);
        match &runs[0] {
            ShapeTextRun::Math {
                display,
                font_size,
                color,
                nodes,
            } => {
                assert!(!*display, "bare a14:m + m:oMath → inline");
                assert_eq!(*font_size, Some(28.0), "size from math run rPr sz/100");
                assert_eq!(
                    color.as_deref(),
                    Some("#7030A0"),
                    "colour from math run rPr solidFill (xlsx #-prefixed convention)"
                );
                assert_eq!(nodes_to_text(nodes), "n");
            }
            other => panic!("expected Math run, got {other:?}"),
        }
    }

    /// Text + break + math coexist in one paragraph; the run order is preserved
    /// and `<a:br>` becomes a `Break` variant.
    #[test]
    fn mixes_text_break_and_math_runs() {
        let xml = format!(
            r#"<xdr:txBody {NS}>
              <a:p>
                <a:r><a:rPr sz="1100"/><a:t>E=</a:t></a:r>
                <a14:m><m:oMath><m:r><m:t>mc</m:t></m:r></m:oMath></a14:m>
                <a:br/>
                <a:r><a:t>done</a:t></a:r>
              </a:p>
            </xdr:txBody>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let text = parse_tx_body(&doc.root_element(), &[]).expect("txBody parses");
        let runs = &text.paragraphs[0].runs;
        assert_eq!(runs.len(), 4, "text, math, break, text");
        assert!(matches!(runs[0], ShapeTextRun::Text { .. }));
        assert!(matches!(runs[1], ShapeTextRun::Math { display: false, .. }));
        assert!(matches!(runs[2], ShapeTextRun::Break));
        assert!(matches!(runs[3], ShapeTextRun::Text { .. }));
    }

    /// `<a:pPr rtl="1">` (ECMA-376 §21.1.2.2.7) marks the paragraph as
    /// right-to-left.
    #[test]
    fn parses_paragraph_rtl_attribute() {
        let xml = format!(
            r#"<xdr:txBody {NS}>
              <a:p>
                <a:pPr rtl="1"/>
                <a:r><a:t>שלום</a:t></a:r>
              </a:p>
            </xdr:txBody>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let text = parse_tx_body(&doc.root_element(), &[]).expect("txBody parses");
        assert_eq!(text.paragraphs.len(), 1);
        assert!(text.paragraphs[0].rtl, "pPr rtl=\"1\" → rtl true");
    }

    /// Absent `@rtl` defaults to false (left-to-right).
    #[test]
    fn paragraph_rtl_defaults_false_when_absent() {
        let xml = format!(
            r#"<xdr:txBody {NS}>
              <a:p>
                <a:pPr algn="ctr"/>
                <a:r><a:t>hello</a:t></a:r>
              </a:p>
            </xdr:txBody>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let text = parse_tx_body(&doc.root_element(), &[]).expect("txBody parses");
        assert_eq!(text.paragraphs.len(), 1);
        assert!(!text.paragraphs[0].rtl, "absent @rtl → rtl false");
    }

    /// `<a:pPr marL marR indent>` (ECMA-376 §21.1.2.2.7,
    /// `CT_TextParagraphProperties`) are the direct paragraph indent attributes
    /// in EMU. `indent` may be negative (hanging). They parse onto the
    /// `ShapeParagraph` as `Option<i64>` so an absent attribute stays `None`.
    #[test]
    fn parses_paragraph_indent_attributes() {
        let xml = format!(
            r#"<xdr:txBody {NS}>
              <a:p>
                <a:pPr marL="457200" marR="91440" indent="-228600"/>
                <a:r><a:t>indented</a:t></a:r>
              </a:p>
            </xdr:txBody>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let text = parse_tx_body(&doc.root_element(), &[]).expect("txBody parses");
        assert_eq!(text.paragraphs.len(), 1);
        let p = &text.paragraphs[0];
        assert_eq!(p.mar_l, Some(457200), "marL parses as EMU i64");
        assert_eq!(p.mar_r, Some(91440), "marR parses as EMU i64");
        assert_eq!(
            p.indent,
            Some(-228600),
            "indent parses (negative = hanging)"
        );
    }

    /// Absent indent attributes (or absent `<a:pPr>` entirely) leave all three
    /// fields `None` so the JSON output stays byte-identical (additive).
    #[test]
    fn paragraph_indent_defaults_none_when_absent() {
        let xml = format!(
            r#"<xdr:txBody {NS}>
              <a:p>
                <a:r><a:t>plain</a:t></a:r>
              </a:p>
            </xdr:txBody>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let text = parse_tx_body(&doc.root_element(), &[]).expect("txBody parses");
        assert_eq!(text.paragraphs.len(), 1);
        let p = &text.paragraphs[0];
        assert_eq!(p.mar_l, None, "absent marL → None");
        assert_eq!(p.mar_r, None, "absent marR → None");
        assert_eq!(p.indent, None, "absent indent → None");

        // None of the three keys should appear in the serialized JSON.
        let v: serde_json::Value = serde_json::to_value(p).unwrap();
        assert!(v.get("marL").is_none(), "marL omitted when None");
        assert!(v.get("marR").is_none(), "marR omitted when None");
        assert!(v.get("indent").is_none(), "indent omitted when None");
        // spaceLine is additive/omitted when absent; autoFit always present.
        assert!(p.space_line.is_none(), "absent lnSpc → None");
        assert!(v.get("spaceLine").is_none(), "spaceLine omitted when None");
        let vt: serde_json::Value = serde_json::to_value(&text).unwrap();
        assert_eq!(vt["autoFit"], "none", "no autofit child → autoFit=none");
        assert!(
            vt.get("fontScale").is_none(),
            "fontScale omitted when unset"
        );
        assert!(
            vt.get("lnSpcReduction").is_none(),
            "lnSpcReduction omitted when unset"
        );
    }

    /// `<a:pPr>/<a:lnSpc>` line spacing (ECMA-376 §21.1.2.2.5): spcPct → a
    /// percent SpaceLine (raw @val), spcPts → a points SpaceLine (raw hundredths
    /// of a point divided by 100). Mirrors the pptx SpaceLine JSON contract.
    #[test]
    fn parses_lnspc_pct_and_pts() {
        let xml = format!(
            r#"<xdr:txBody {NS}>
              <a:p>
                <a:pPr><a:lnSpc><a:spcPct val="150000"/></a:lnSpc></a:pPr>
                <a:r><a:t>pct</a:t></a:r>
              </a:p>
              <a:p>
                <a:pPr><a:lnSpc><a:spcPts val="1800"/></a:lnSpc></a:pPr>
                <a:r><a:t>pts</a:t></a:r>
              </a:p>
            </xdr:txBody>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let text = parse_tx_body(&doc.root_element(), &[]).expect("txBody parses");
        assert_eq!(text.paragraphs.len(), 2);

        let vp0: serde_json::Value = serde_json::to_value(&text.paragraphs[0]).unwrap();
        assert_eq!(vp0["spaceLine"]["type"], "pct");
        assert_eq!(vp0["spaceLine"]["val"], 150000.0);

        let vp1: serde_json::Value = serde_json::to_value(&text.paragraphs[1]).unwrap();
        assert_eq!(vp1["spaceLine"]["type"], "pts");
        // 1800 hundredths of a point → 18 pt.
        assert_eq!(vp1["spaceLine"]["val"], 18.0);
    }

    /// `<a:bodyPr>/<a:normAutofit>` (ECMA-376 §21.1.2.1.3): autoFit="norm" plus
    /// the stored fontScale / lnSpcReduction (1000ths of a percent → fraction).
    #[test]
    fn parses_normautofit_scales() {
        let xml = format!(
            r#"<xdr:txBody {NS}>
              <a:bodyPr><a:normAutofit fontScale="62500" lnSpcReduction="20000"/></a:bodyPr>
              <a:p><a:r><a:t>x</a:t></a:r></a:p>
            </xdr:txBody>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let text = parse_tx_body(&doc.root_element(), &[]).expect("txBody parses");
        let v: serde_json::Value = serde_json::to_value(&text).unwrap();
        assert_eq!(v["autoFit"], "norm");
        assert_eq!(v["fontScale"], 0.625);
        assert_eq!(v["lnSpcReduction"], 0.2);
    }

    /// `<a:bodyPr>/<a:spAutoFit>` → autoFit="sp" (no scales); `<a:noAutofit>` →
    /// autoFit="none". Both leave rendering unchanged (renderer applies neither).
    #[test]
    fn parses_spautofit_and_noautofit() {
        let sp_xml = format!(
            r#"<xdr:txBody {NS}>
              <a:bodyPr><a:spAutoFit/></a:bodyPr>
              <a:p><a:r><a:t>x</a:t></a:r></a:p>
            </xdr:txBody>"#
        );
        let doc = roxmltree::Document::parse(&sp_xml).unwrap();
        let text = parse_tx_body(&doc.root_element(), &[]).expect("txBody parses");
        let v: serde_json::Value = serde_json::to_value(&text).unwrap();
        assert_eq!(v["autoFit"], "sp");
        assert!(
            v.get("fontScale").is_none(),
            "spAutoFit stores no fontScale"
        );

        let no_xml = format!(
            r#"<xdr:txBody {NS}>
              <a:bodyPr><a:noAutofit/></a:bodyPr>
              <a:p><a:r><a:t>x</a:t></a:r></a:p>
            </xdr:txBody>"#
        );
        let doc = roxmltree::Document::parse(&no_xml).unwrap();
        let text = parse_tx_body(&doc.root_element(), &[]).expect("txBody parses");
        let v: serde_json::Value = serde_json::to_value(&text).unwrap();
        assert_eq!(v["autoFit"], "none");
    }

    /// `ShapeTextRun` uses an enum-level `#[serde(tag = "type", rename_all =
    /// "camelCase")]`, which renames only the variant tags — not the fields. The
    /// `Text` variant's `font_face` and the `Math` variant's `font_size`
    /// therefore need per-variant `rename_all` to serialize as the camelCase
    /// keys the TS renderer reads (`run.fontFace` / `run.fontSize`). This locks
    /// the JSON contract so the keys never regress to snake_case (which the
    /// renderer reads as `undefined`). Same bug class as the pptx serde fix
    /// (PR #489) and the xlsx ArcTo fix (PR #491).
    #[test]
    fn serializes_text_run_font_face_as_camel_case() {
        let xml = format!(
            r#"<xdr:txBody {NS}>
              <a:p>
                <a:r>
                  <a:rPr sz="1400"><a:latin typeface="Calibri"/></a:rPr>
                  <a:t>hi</a:t>
                </a:r>
              </a:p>
            </xdr:txBody>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let text = parse_tx_body(&doc.root_element(), &[]).expect("txBody parses");
        let run = &text.paragraphs[0].runs[0];
        assert!(
            matches!(run, ShapeTextRun::Text { .. }),
            "expected Text run"
        );

        let v: serde_json::Value = serde_json::to_value(run).unwrap();
        assert_eq!(v["type"], "text", "tag key is `type`");
        assert_eq!(
            v["fontFace"], "Calibri",
            "font_face must serialize as camelCase `fontFace` (renderer reads run.fontFace)"
        );
        assert!(
            v.get("font_face").is_none(),
            "snake_case `font_face` must not appear"
        );
    }

    /// ECMA-376 §21.1.2.3.1 `<a:ea>` (East-Asian) / `<a:cs>` (complex-script)
    /// typefaces parse onto the run alongside `<a:latin>`. The common Japanese
    /// encoding sets Meiryo on `<a:ea>` while leaving `<a:latin>` default; the
    /// renderer floors the line box by the tabled face. They serialize as
    /// camelCase `fontFaceEa` / `fontFaceCs` and are additive (omitted when
    /// absent), so shapes without ea/cs stay byte-identical.
    #[test]
    fn parses_ea_and_cs_typefaces() {
        let xml = format!(
            r#"<xdr:txBody {NS}>
              <a:p>
                <a:r>
                  <a:rPr sz="1400">
                    <a:ea typeface="Meiryo"/>
                    <a:cs typeface="Sakkal Majalla"/>
                  </a:rPr>
                  <a:t>あ</a:t>
                </a:r>
              </a:p>
            </xdr:txBody>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let text = parse_tx_body(&doc.root_element(), &[]).expect("txBody parses");
        let run = &text.paragraphs[0].runs[0];
        match run {
            ShapeTextRun::Text {
                font_face,
                font_face_ea,
                font_face_cs,
                ..
            } => {
                assert!(font_face.is_none(), "latin left default → font_face None");
                assert_eq!(
                    font_face_ea.as_deref(),
                    Some("Meiryo"),
                    "<a:ea> → fontFaceEa"
                );
                assert_eq!(
                    font_face_cs.as_deref(),
                    Some("Sakkal Majalla"),
                    "<a:cs> → fontFaceCs"
                );
            }
            other => panic!("expected Text run, got {other:?}"),
        }

        let v: serde_json::Value = serde_json::to_value(run).unwrap();
        assert_eq!(
            v["fontFaceEa"], "Meiryo",
            "serializes as camelCase fontFaceEa"
        );
        assert_eq!(
            v["fontFaceCs"], "Sakkal Majalla",
            "serializes as camelCase fontFaceCs"
        );
    }

    /// Additive guarantee: a run with no `<a:ea>` / `<a:cs>` omits both keys
    /// from JSON entirely (serde `skip_serializing_if = "Option::is_none"`), so
    /// existing shape JSON stays byte-identical.
    #[test]
    fn omits_ea_and_cs_when_absent() {
        let xml = format!(
            r#"<xdr:txBody {NS}>
              <a:p>
                <a:r>
                  <a:rPr sz="1400"><a:latin typeface="Calibri"/></a:rPr>
                  <a:t>hi</a:t>
                </a:r>
              </a:p>
            </xdr:txBody>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let text = parse_tx_body(&doc.root_element(), &[]).expect("txBody parses");
        let run = &text.paragraphs[0].runs[0];
        let v: serde_json::Value = serde_json::to_value(run).unwrap();
        assert!(
            v.get("fontFaceEa").is_none(),
            "fontFaceEa omitted when absent"
        );
        assert!(
            v.get("fontFaceCs").is_none(),
            "fontFaceCs omitted when absent"
        );
    }

    /// Companion to the Text-run case: the `Math` variant's `font_size` must
    /// serialize as `fontSize` (renderer reads `run.fontSize`). Without the
    /// per-variant `rename_all` the equation falls back to the inherited size.
    #[test]
    fn serializes_math_run_font_size_as_camel_case() {
        let xml = format!(
            r#"<xdr:txBody {NS}>
              <a:p>
                <a14:m><m:oMath><m:r>
                  <a:rPr sz="2800"/>
                  <m:t>n</m:t>
                </m:r></m:oMath></a14:m>
              </a:p>
            </xdr:txBody>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let text = parse_tx_body(&doc.root_element(), &[]).expect("txBody parses");
        let run = &text.paragraphs[0].runs[0];
        assert!(
            matches!(run, ShapeTextRun::Math { .. }),
            "expected Math run"
        );

        let v: serde_json::Value = serde_json::to_value(run).unwrap();
        assert_eq!(v["type"], "math", "tag key is `type`");
        assert_eq!(
            v["fontSize"], 28.0,
            "font_size must serialize as camelCase `fontSize` (renderer reads run.fontSize)"
        );
        assert!(
            v.get("font_size").is_none(),
            "snake_case `font_size` must not appear"
        );
    }
}

#[cfg(test)]
mod style_lnref_tests {
    use super::*;

    const NS: &str = r#"xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main""#;

    /// dk1, lt1, dk2, lt2, accent1 — enough slots for schemeClr accent1 (idx 4).
    fn theme() -> Vec<String> {
        ["#000000", "#FFFFFF", "#222222", "#EEEEEE", "#4472C4"]
            .iter()
            .map(|s| s.to_string())
            .collect()
    }

    fn format_scheme() -> ThemeFormatScheme {
        ThemeFormatScheme::parse(
            r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:themeElements><a:fmtScheme name="s"><a:lnStyleLst>
                <a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
                <a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
                <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
              </a:lnStyleLst></a:fmtScheme></a:themeElements>
            </a:theme>"#,
        )
    }

    fn shape_of(sp_pr_inner: &str, style: &str) -> (Option<String>, i64) {
        let shape = shape_info_with_scheme(sp_pr_inner, style, &format_scheme());
        (shape.stroke_color, shape.stroke_width)
    }

    fn shape_of_with_scheme(
        sp_pr_inner: &str,
        style: &str,
        scheme: &ThemeFormatScheme,
    ) -> (Option<String>, i64) {
        let shape = shape_info_with_scheme(sp_pr_inner, style, scheme);
        (shape.stroke_color, shape.stroke_width)
    }

    fn shape_info_with_scheme(
        sp_pr_inner: &str,
        style: &str,
        scheme: &ThemeFormatScheme,
    ) -> ShapeInfo {
        let xml = format!(
            r#"<xdr:wsDr {NS}><xdr:twoCellAnchor>
              <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
              <xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
              <xdr:sp>
                <xdr:nvSpPr><xdr:cNvPr id="2" name="P"/><xdr:cNvSpPr/></xdr:nvSpPr>
                <xdr:spPr>
                  <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                  <a:prstGeom prst="parallelogram"><a:avLst/></a:prstGeom>
                  {sp_pr_inner}
                </xdr:spPr>
                {style}
              </xdr:sp>
              <xdr:clientData/>
            </xdr:twoCellAnchor></xdr:wsDr>"#
        );
        let mut anchors = parse_shape_anchors(&xml, &theme(), scheme, &HashMap::new());
        assert_eq!(anchors.len(), 1, "one anchor expected");
        anchors.remove(0).shapes.remove(0)
    }

    /// §20.1.4.2.19 — with no explicit `<a:ln>`, the outline comes from the
    /// theme style matrix: lnRef idx picks the lnStyleLst width, the lnRef's
    /// child scheme colour substitutes phClr.
    #[test]
    fn lnref_supplies_stroke_when_sppr_has_no_ln() {
        let (color, width) = shape_of(
            "",
            r#"<xdr:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef></xdr:style>"#,
        );
        assert_eq!(color.as_deref(), Some("#4472C4"));
        assert_eq!(width, 12_700);
    }

    #[test]
    fn lnref_without_color_child_uses_fixed_theme_line_color() {
        let scheme = ThemeFormatScheme::parse(
            r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:themeElements><a:fmtScheme name="s"><a:lnStyleLst>
                <a:ln w="25400" cap="rnd" cmpd="dbl">
                  <a:solidFill><a:srgbClr val="112233"/></a:solidFill>
                  <a:prstDash val="dash"/><a:round/>
                  <a:tailEnd type="triangle"/>
                </a:ln>
              </a:lnStyleLst></a:fmtScheme></a:themeElements>
            </a:theme>"#,
        );
        let (color, width) =
            shape_of_with_scheme("", r#"<xdr:style><a:lnRef idx="1"/></xdr:style>"#, &scheme);
        assert_eq!(color.as_deref(), Some("#112233"));
        assert_eq!(width, 25_400);
    }

    #[test]
    fn local_width_overrides_theme_without_discarding_its_paint() {
        let (color, width) = shape_of(
            r#"<a:ln w="28575"/>"#,
            r#"<xdr:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef></xdr:style>"#,
        );
        assert_eq!(color.as_deref(), Some("#4472C4"));
        assert_eq!(width, 28_575);
    }

    #[test]
    fn complete_line_properties_map_to_the_additive_shape_wire() {
        let xml = r#"<a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          w="25400" cap="rnd" cmpd="dbl" algn="in">
          <a:gradFill><a:gsLst>
            <a:gs pos="0"><a:srgbClr val="112233"/></a:gs>
            <a:gs pos="100000"><a:srgbClr val="AABBCC"/></a:gs>
          </a:gsLst><a:lin ang="5400000"/></a:gradFill>
          <a:custDash><a:ds d="125000" sp="75000"/></a:custDash>
          <a:miter lim="800000"/>
          <a:headEnd type="triangle" w="lg" len="sm"/>
        </a:ln>"#;
        let document = roxmltree::Document::parse(xml).unwrap();
        let line = ooxml_common::line::parse_line_properties(
            document.root_element(),
            &XlsxSchemeResolver {
                theme_colors: &theme(),
            },
            ooxml_common::color::TintMode::PowerPointLinear,
        );
        let wire = xlsx_line_wire_properties(&line);

        assert_eq!(wire.color.as_deref(), Some("#AABBCC"));
        assert_eq!(wire.width, 25_400);
        assert!(matches!(wire.fill, Some(ShapeStrokeFill::Gradient { .. })));
        assert_eq!(wire.custom_dash.len(), 1);
        assert_eq!(wire.line_cap.as_deref(), Some("round"));
        assert_eq!(wire.line_join.as_deref(), Some("miter"));
        assert_eq!(wire.miter_limit, Some(8.0));
        assert_eq!(wire.alignment.as_deref(), Some("in"));
        assert_eq!(wire.cmpd.as_deref(), Some("dbl"));
        assert_eq!(
            wire.head_end.as_ref().map(|end| end.r#type.as_str()),
            Some("triangle")
        );
    }

    /// An explicit `<a:ln><a:noFill/>` is a hard "no outline" and must beat
    /// the theme lnRef fallback.
    #[test]
    fn explicit_nofill_ln_beats_lnref() {
        let (color, width) = shape_of(
            r#"<a:ln><a:noFill/></a:ln>"#,
            r#"<xdr:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef></xdr:style>"#,
        );
        assert!(color.is_none());
        assert_eq!(width, 0);
    }

    /// An explicit `<a:ln>` with its own fill/width wins over lnRef.
    #[test]
    fn explicit_ln_beats_lnref() {
        let (color, width) = shape_of(
            r#"<a:ln w="28575"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln>"#,
            r#"<xdr:style><a:lnRef idx="3"><a:schemeClr val="accent1"/></a:lnRef></xdr:style>"#,
        );
        assert_eq!(color.as_deref(), Some("#00FF00"));
        assert_eq!(width, 28_575);
    }

    /// A missing style-matrix entry is not an authored line recipe. It remains
    /// distinct from the CT_LineProperties defaults of a selected entry.
    #[test]
    fn lnref_out_of_range_does_not_fabricate_a_line() {
        let (color, width) = shape_of(
            "",
            r#"<xdr:style><a:lnRef idx="9"><a:schemeClr val="accent1"/></a:lnRef></xdr:style>"#,
        );
        assert_eq!(color, None);
        assert_eq!(width, 0);
    }

    /// §20.1.4.2.19 + ST_StyleMatrixColumnIndex (§20.1.10.57): the style-matrix
    /// lists (`lnStyleLst` etc.) are 1-based, so `<a:lnRef idx="0">` references
    /// NO entry — it is the well-known DrawingML encoding for "no line". Excel /
    /// PowerPoint write exactly this when the user sets the shape outline to
    /// "No Line" (e.g. an inserted equation text box in sample-28, which carries
    /// `<a:lnRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef>`). The child
    /// colour is only the phClr substitution that WOULD apply if a line style
    /// were selected; with idx=0 no style is selected, so no outline is drawn.
    /// Must yield no stroke — matching pptx's `shape.rs` (idx==0 → None). Before
    /// the fix xlsx read the colour and drew a spurious 0.75 pt black border.
    #[test]
    fn lnref_idx_zero_is_no_line() {
        let (color, width) = shape_of(
            "",
            r#"<xdr:style><a:lnRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef></xdr:style>"#,
        );
        assert!(color.is_none(), "lnRef idx=0 → no outline");
        assert_eq!(width, 0);
    }

    #[test]
    fn direct_gradient_fill_preserves_shared_geometry() {
        let shape = shape_info_with_scheme(
            r#"<a:gradFill flip="y" rotWithShape="0"><a:gsLst>
                 <a:gs pos="0"><a:srgbClr val="112233"/></a:gs>
                 <a:gs pos="100000"><a:srgbClr val="AABBCC"/></a:gs>
               </a:gsLst><a:path path="circle"><a:fillToRect l="25000"/></a:path>
               <a:tileRect t="50000"/></a:gradFill>"#,
            "",
            &format_scheme(),
        );

        let Some(ShapeFill::Gradient {
            path,
            fill_to_rect,
            tile_rect,
            flip,
            rot_with_shape,
            ..
        }) = shape.fill
        else {
            panic!("expected gradient fill");
        };
        assert_eq!(path.as_deref(), Some("circle"));
        assert_eq!(fill_to_rect.as_ref().map(|rect| rect.l), Some(0.25));
        assert_eq!(tile_rect.as_ref().map(|rect| rect.t), Some(0.5));
        assert_eq!(flip.as_deref(), Some("y"));
        assert_eq!(rot_with_shape, Some(false));
    }

    #[test]
    fn fillref_resolves_pattern_recipe_instead_of_flattening_reference_color() {
        let scheme = ThemeFormatScheme::parse(
            r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:themeElements><a:fmtScheme name="s"><a:fillStyleLst>
                <a:pattFill prst="diagCross">
                  <a:fgClr><a:schemeClr val="phClr"/></a:fgClr>
                  <a:bgClr><a:srgbClr val="F0E0D0"/></a:bgClr>
                </a:pattFill>
              </a:fillStyleLst></a:fmtScheme></a:themeElements>
            </a:theme>"#,
        );
        let shape = shape_info_with_scheme(
            "",
            r#"<xdr:style><a:fillRef idx="1"><a:srgbClr val="102030"/></a:fillRef></xdr:style>"#,
            &scheme,
        );

        assert!(matches!(
            shape.fill,
            Some(ShapeFill::Pattern { fg, bg, preset })
                if fg == "102030" && bg == "F0E0D0" && preset == "diagCross"
        ));
    }

    /// A standalone `<xdr:pic>` directly under a `twoCellAnchor` is a plain image
    /// anchor owned by `ws.images` (parse_drawing_anchors, WITH its `<a:srcRect>`
    /// crop). `parse_shape_anchors` must NOT also capture it — otherwise the
    /// renderer draws the picture twice and the uncropped shape draw overwrites
    /// the cropped anchor draw (the sample-27 double-draw regression).
    #[test]
    fn standalone_twocellanchor_pic_is_not_a_shape_anchor() {
        let xml = format!(
            r#"<xdr:wsDr {NS} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor editAs="oneCell">
              <xdr:from><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
              <xdr:to><xdr:col>11</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
              <xdr:pic>
                <xdr:nvPicPr><xdr:cNvPr id="2" name="P"/><xdr:cNvPicPr/></xdr:nvPicPr>
                <xdr:blipFill><a:blip r:embed="rId1"/><a:srcRect l="32560" r="3829"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
                <xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2905125" cy="2181225"/></a:xfrm>
                  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
              </xdr:pic>
              <xdr:clientData/>
            </xdr:twoCellAnchor></xdr:wsDr>"#
        );
        let anchors = parse_shape_anchors(
            &xml,
            &theme(),
            &ThemeFormatScheme::default(),
            &HashMap::new(),
        );
        assert!(
            anchors.is_empty(),
            "twoCellAnchor pic belongs to ws.images, not ws.shape_groups: {anchors:?}"
        );
    }

    /// A standalone `<xdr:sp>` under the same anchor IS a shape anchor — only the
    /// `pic` is deduped, never a real shape.
    #[test]
    fn standalone_sp_is_still_a_shape_anchor() {
        let (color, _width) = shape_of(
            r#"<a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>"#,
            "",
        );
        assert_eq!(color.as_deref(), Some("#FF0000"), "sp still captured");
    }

    /// A `oneCellAnchor` pic is NOT handled by `parse_drawing_anchors` (which only
    /// scans `twoCellAnchor`), so the shape path must keep capturing it — else the
    /// image is dropped entirely. Its `<a:srcRect>` crop must also be surfaced on
    /// the `ShapeGeom::Image` so the renderer can crop it like a top-level anchor.
    #[test]
    fn standalone_onecellanchor_pic_is_still_captured_with_crop() {
        let xml = format!(
            r#"<xdr:wsDr {NS} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:oneCellAnchor>
              <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
              <xdr:ext cx="914400" cy="914400"/>
              <xdr:pic>
                <xdr:nvPicPr><xdr:cNvPr id="2" name="P"/><xdr:cNvPicPr/></xdr:nvPicPr>
                <xdr:blipFill><a:blip r:embed="rId1"/><a:srcRect l="10000" t="20000"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
                <xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
              </xdr:pic>
              <xdr:clientData/>
            </xdr:oneCellAnchor></xdr:wsDr>"#
        );
        let mut rids = HashMap::new();
        rids.insert("rId1".to_string(), "xl/media/image1.png".to_string());
        let anchors = parse_shape_anchors(&xml, &theme(), &ThemeFormatScheme::default(), &rids);
        assert_eq!(anchors.len(), 1, "oneCellAnchor pic must still be captured");
        match &anchors[0].shapes[0].geom {
            ShapeGeom::Image { src_rect, .. } => {
                let sr = src_rect
                    .as_ref()
                    .expect("leaf pic surfaces its srcRect crop");
                assert!((sr.l - 0.1).abs() < 1e-9);
                assert!((sr.t - 0.2).abs() < 1e-9);
            }
            other => panic!("expected an image-geom shape, got {other:?}"),
        }
    }
}

/// §20.1.2.2.8 — an `<xdr:cNvPr hidden="1">` drawing object is not rendered.
/// Covers shapes (parse_shape_anchors), grouped shapes (collect_shapes) and
/// pictures (parse_drawing_anchors).
#[cfg(test)]
mod hidden_tests {
    use super::*;

    const NS: &str = r#"xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;

    fn theme() -> Vec<String> {
        ["#000000", "#FFFFFF", "#222222", "#EEEEEE", "#4472C4"]
            .iter()
            .map(|s| s.to_string())
            .collect()
    }

    fn sp_anchor(hidden_attr: &str) -> String {
        format!(
            r#"<xdr:wsDr {NS}><xdr:twoCellAnchor>
              <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
              <xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
              <xdr:sp>
                <xdr:nvSpPr><xdr:cNvPr id="2" name="P"{hidden}/><xdr:cNvSpPr/></xdr:nvSpPr>
                <xdr:spPr>
                  <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                  <a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>
                </xdr:spPr>
              </xdr:sp>
              <xdr:clientData/>
            </xdr:twoCellAnchor></xdr:wsDr>"#,
            NS = NS,
            hidden = hidden_attr,
        )
    }

    #[test]
    fn hidden_standalone_shape_is_not_emitted() {
        for attr in [r#" hidden="1""#, r#" hidden="true""#] {
            let anchors = parse_shape_anchors(
                &sp_anchor(attr),
                &theme(),
                &ThemeFormatScheme::default(),
                &HashMap::new(),
            );
            assert!(anchors.is_empty(), "hidden sp emitted (attr={attr})");
        }
    }

    #[test]
    fn visible_standalone_shape_is_emitted() {
        for attr in ["", r#" hidden="0""#, r#" hidden="false""#] {
            let anchors = parse_shape_anchors(
                &sp_anchor(attr),
                &theme(),
                &ThemeFormatScheme::default(),
                &HashMap::new(),
            );
            assert_eq!(anchors.len(), 1, "visible sp dropped (attr={attr})");
        }
    }

    // ── MCE AlternateContent selection (ECMA-376 Part 3 §9.3) ───────────────

    /// A `<xdr:sp>` with the given solid `fill` colour — used to tell which MCE
    /// branch the shape-unwrap site selected.
    fn filled_sp(fill: &str) -> String {
        format!(
            r#"<xdr:sp>
                 <xdr:nvSpPr><xdr:cNvPr id="2" name="S"/><xdr:cNvSpPr/></xdr:nvSpPr>
                 <xdr:spPr>
                   <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                   <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                   <a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>
                 </xdr:spPr>
               </xdr:sp>"#
        )
    }

    /// A `<xdr:twoCellAnchor>` whose `<mc:AlternateContent>` Choice (behind
    /// `requires`) carries a red-filled shape and whose Fallback carries a
    /// blue-filled one. The fill colour of the emitted shape reveals the branch.
    fn alt_content_shape_anchor(requires: &str) -> String {
        format!(
            r#"<xdr:wsDr {NS}
                 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                 xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
                 xmlns:unk="http://example.com/an/extension/we/do/not/understand">
              <xdr:twoCellAnchor>
                <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
                <xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
                <mc:AlternateContent>
                  <mc:Choice Requires="{requires}">{choice}</mc:Choice>
                  <mc:Fallback>{fallback}</mc:Fallback>
                </mc:AlternateContent>
                <xdr:clientData/>
              </xdr:twoCellAnchor></xdr:wsDr>"#,
            NS = NS,
            requires = requires,
            choice = filled_sp("FF0000"),
            fallback = filled_sp("0000FF"),
        )
    }

    /// ECMA-376 Part 3 §9.3(1) — the anchor-level shape wrapper: a Choice whose
    /// `Requires` names an un-understood namespace must not be selected even
    /// though its shape would render; the `<mc:Fallback>` shape is drawn instead.
    /// The old code always unwrapped to the Choice.
    #[test]
    fn mce_shape_unhandled_choice_selects_fallback() {
        let anchors = parse_shape_anchors(
            &alt_content_shape_anchor("unk"),
            &theme(),
            &ThemeFormatScheme::default(),
            &HashMap::new(),
        );
        assert_eq!(anchors.len(), 1, "one anchor");
        assert_eq!(anchors[0].shapes.len(), 1, "one shape");
        assert_eq!(
            anchors[0].shapes[0].fill_color.as_deref(),
            Some("#0000FF"),
            "un-understood Choice must yield the Fallback (blue) shape, not the Choice (red)"
        );
    }

    /// §9.3(1) — `a14` (drawing 2010) IS understood; the real
    /// `<mc:Choice Requires="a14">` that Excel wraps ordinary shapes in
    /// (sample-1 / sample-28) is selected, the Fallback dropped. Guards those
    /// slides against a regression.
    #[test]
    fn mce_shape_understood_a14_choice_selected() {
        let anchors = parse_shape_anchors(
            &alt_content_shape_anchor("a14"),
            &theme(),
            &ThemeFormatScheme::default(),
            &HashMap::new(),
        );
        assert_eq!(anchors.len(), 1, "one anchor");
        assert_eq!(anchors[0].shapes.len(), 1, "one shape");
        assert_eq!(
            anchors[0].shapes[0].fill_color.as_deref(),
            Some("#FF0000"),
            "understood a14 Choice must be selected (red), Fallback dropped"
        );
    }

    #[test]
    fn hidden_group_elides_children() {
        // A visible group with one hidden and one visible child leaf emits only
        // the visible child. A fully hidden group emits nothing.
        let group = |grp_hidden: &str, child_hidden: &str| {
            format!(
                r#"<xdr:wsDr {NS}><xdr:twoCellAnchor>
                  <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
                  <xdr:to><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
                  <xdr:grpSp>
                    <xdr:nvGrpSpPr><xdr:cNvPr id="2" name="G"{grp_hidden}/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>
                    <xdr:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="1828800"/>
                      <a:chOff x="0" y="0"/><a:chExt cx="1828800" cy="1828800"/></a:xfrm></xdr:grpSpPr>
                    <xdr:sp>
                      <xdr:nvSpPr><xdr:cNvPr id="3" name="C1"{child_hidden}/><xdr:cNvSpPr/></xdr:nvSpPr>
                      <xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm>
                        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                        <a:ln w="6350"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln></xdr:spPr></xdr:sp>
                    <xdr:sp>
                      <xdr:nvSpPr><xdr:cNvPr id="4" name="C2"/><xdr:cNvSpPr/></xdr:nvSpPr>
                      <xdr:spPr><a:xfrm><a:off x="457200" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm>
                        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                        <a:ln w="6350"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln></xdr:spPr></xdr:sp>
                  </xdr:grpSp>
                  <xdr:clientData/>
                </xdr:twoCellAnchor></xdr:wsDr>"#,
                NS = NS,
                grp_hidden = grp_hidden,
                child_hidden = child_hidden,
            )
        };

        // Visible group, one hidden child → only the visible child leaf remains.
        let a = parse_shape_anchors(
            &group("", r#" hidden="1""#),
            &theme(),
            &ThemeFormatScheme::default(),
            &HashMap::new(),
        );
        assert_eq!(a.len(), 1);
        assert_eq!(a[0].shapes.len(), 1, "hidden child leaked");

        // Hidden group → no anchor at all.
        let b = parse_shape_anchors(
            &group(r#" hidden="1""#, ""),
            &theme(),
            &ThemeFormatScheme::default(),
            &HashMap::new(),
        );
        assert!(b.is_empty(), "hidden group leaked");
    }

    // ── Group-shape recursion depth guard (RB2) ────────────────────────────

    /// A twoCellAnchor holding `<xdr:grpSp>` nested `levels` deep, with one leaf
    /// `<xdr:sp>` at the bottom. Each `<xdr:grpSp>` is a real `collect_shapes`
    /// recursion.
    fn nested_group_anchor(levels: usize) -> String {
        let leaf = concat!(
            r#"<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="99" name="Leaf"/><xdr:cNvSpPr/></xdr:nvSpPr>"#,
            r#"<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm>"#,
            r#"<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>"#,
            r#"<a:ln w="6350"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln></xdr:spPr></xdr:sp>"#,
        );
        let open = concat!(
            r#"<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="2" name="G"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>"#,
            r#"<xdr:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="1828800"/>"#,
            r#"<a:chOff x="0" y="0"/><a:chExt cx="1828800" cy="1828800"/></a:xfrm></xdr:grpSpPr>"#,
        );
        let mut inner = leaf.to_string();
        for _ in 0..levels {
            inner = format!("{open}{inner}</xdr:grpSp>");
        }
        format!(
            r#"<xdr:wsDr {NS}><xdr:twoCellAnchor>
              <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
              <xdr:to><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
              {inner}
              <xdr:clientData/>
            </xdr:twoCellAnchor></xdr:wsDr>"#,
            NS = NS,
        )
    }

    fn anchors_deep(levels: usize) -> Vec<ShapeAnchor> {
        // roxmltree::Document::parse itself recurses on nesting depth, so parse on
        // a generous stack; this test targets OUR collect_shapes guard. (The
        // roxmltree layer is bounded separately by the raw-XML depth pre-check in
        // `ooxml_common::depth`.)
        std::thread::Builder::new()
            .stack_size(256 * 1024 * 1024)
            .spawn(move || {
                parse_shape_anchors(
                    &nested_group_anchor(levels),
                    &theme(),
                    &ThemeFormatScheme::default(),
                    &HashMap::new(),
                )
            })
            .unwrap()
            .join()
            .unwrap()
    }

    #[test]
    fn deeply_nested_groups_do_not_trap() {
        // ~2 000 nested groups is ~30× the depth limit; pre-guard this recurses
        // 2 000 frames and traps on the WASM stack. The guard caps the descent so
        // parsing RETURNS instead of aborting. A group deeper than the limit drops
        // its subtree (the single leaf below it does not survive) — the assertion
        // is that we reach here with a bounded result and no trap.
        let anchors = anchors_deep(2_000);
        let leaves: usize = anchors.iter().map(|a| a.shapes.len()).sum();
        assert!(leaves <= 1, "guard should bound the collected leaves");
    }

    #[test]
    fn shallow_nested_groups_keep_their_leaf() {
        // A modest nest (well under the limit) parses in full: the leaf survives.
        let anchors = anchors_deep(10);
        assert_eq!(anchors.len(), 1, "expected the one anchor");
        assert_eq!(
            anchors[0].shapes.len(),
            1,
            "leaf under a shallow group nest must survive"
        );
    }

    // ── RB2 neutralization (roxmltree layer): a pathologically deep drawing XML
    //    is rejected by the depth pre-check in `parse_guarded` BEFORE roxmltree's
    //    recursive tree builder runs, so `parse_shape_anchors` returns gracefully
    //    instead of trapping. Unlike the two tests above (which spawn a 256 MB
    //    stack to exercise OUR collect_shapes guard), this runs on the DEFAULT
    //    (small) test-thread stack ON PURPOSE: handing 5 000-deep XML straight to
    //    roxmltree here would overflow and abort the process. That it returns is
    //    the guarantee.
    #[test]
    fn deeply_nested_drawing_xml_is_rejected_not_trapped() {
        let mut xml = format!(r#"<xdr:wsDr {NS}>"#, NS = NS);
        // 5 000 levels — ~20× MAX_XML_DEPTH and far past roxmltree's ~1 000-deep
        // small-stack overflow threshold.
        for _ in 0..5_000 {
            xml.push_str("<xdr:grpSp>");
        }
        for _ in 0..5_000 {
            xml.push_str("</xdr:grpSp>");
        }
        xml.push_str("</xdr:wsDr>");

        // Must return (not trap); the rejected part yields no anchors.
        let anchors = parse_shape_anchors(
            &xml,
            &theme(),
            &ThemeFormatScheme::default(),
            &HashMap::new(),
        );
        assert!(
            anchors.is_empty(),
            "an over-deep drawing XML must be rejected, yielding no anchors"
        );
    }
}

#[cfg(test)]
mod geom_tests {
    use super::*;

    const NS: &str = r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main""#;

    fn geom_of(sp_pr_xml: &str) -> ShapeGeom {
        let doc = roxmltree::Document::parse(sp_pr_xml).unwrap();
        parse_sp_geom(&doc.root_element(), 100.0, 100.0).expect("sp_pr has geometry")
    }

    /// A preset with no `<a:avLst>` yields an empty adjust vector — the engine
    /// then falls back to each shape's declared defaults.
    #[test]
    fn preset_without_avlst_has_empty_adj() {
        let geom = geom_of(&format!(
            r#"<a:spPr {NS}><a:prstGeom prst="parallelogram"/></a:spPr>"#
        ));
        match geom {
            ShapeGeom::Preset { name, adj } => {
                assert_eq!(name, "parallelogram");
                assert!(adj.is_empty(), "no avLst → empty adj, got {adj:?}");
            }
            other => panic!("expected Preset, got {other:?}"),
        }
    }

    #[test]
    fn custom_geometry_resolves_guides_and_preserves_quadratic_bezier() {
        let geom = geom_of(&format!(
            r#"<a:spPr {NS}><a:custGeom>
          <a:gdLst><a:gd name="half" fmla="*/ w 1 2"/></a:gdLst>
          <a:pathLst><a:path>
            <a:moveTo><a:pt x="0" y="0"/></a:moveTo>
            <a:quadBezTo><a:pt x="half" y="b"/><a:pt x="r" y="0"/></a:quadBezTo>
          </a:path></a:pathLst>
        </a:custGeom></a:spPr>"#
        ));
        let ShapeGeom::Custom { paths } = geom else {
            panic!("expected custom geometry")
        };
        assert_eq!(paths[0].w, 100.0);
        assert_eq!(paths[0].h, 100.0);
        assert!(matches!(
            paths[0].commands[1],
            PathCmd::QuadBezTo { x1, y1, x2, y2 }
                if (x1 - 50.0).abs() < 1e-9 && (y1 - 100.0).abs() < 1e-9
                    && (x2 - 100.0).abs() < 1e-9 && y2.abs() < 1e-9
        ));
    }

    /// `<a:gd name="adj" fmla="val X"/>` is read into the first adjust slot.
    #[test]
    fn preset_reads_single_named_adj() {
        let geom = geom_of(&format!(
            r#"<a:spPr {NS}><a:prstGeom prst="parallelogram">
                 <a:avLst><a:gd name="adj" fmla="val 41667"/></a:avLst>
               </a:prstGeom></a:spPr>"#
        ));
        match geom {
            ShapeGeom::Preset { name, adj } => {
                assert_eq!(name, "parallelogram");
                assert_eq!(adj, vec![Some(41667.0)]);
            }
            other => panic!("expected Preset, got {other:?}"),
        }
    }

    /// Multiple named handles (`adj1`/`adj2`) land in declaration order; a gap
    /// produces a `None` placeholder so later indices stay aligned.
    #[test]
    fn preset_reads_named_adj1_adj2_and_fills_gaps() {
        let geom = geom_of(&format!(
            r#"<a:spPr {NS}><a:prstGeom prst="wedgeRectCallout">
                 <a:avLst>
                   <a:gd name="adj1" fmla="val -20000"/>
                   <a:gd name="adj3" fmla="val 55000"/>
                 </a:avLst>
               </a:prstGeom></a:spPr>"#
        ));
        match geom {
            ShapeGeom::Preset { adj, .. } => {
                // adj1 set, adj2 missing → None, adj3 set, rest trimmed.
                assert_eq!(adj, vec![Some(-20000.0), None, Some(55000.0)]);
            }
            other => panic!("expected Preset, got {other:?}"),
        }
    }

    /// Unnamed `<a:gd>` handles fall back to declaration position.
    #[test]
    fn preset_positional_fallback_for_unnamed_gd() {
        let geom = geom_of(&format!(
            r#"<a:spPr {NS}><a:prstGeom prst="roundRect">
                 <a:avLst><a:gd fmla="val 16667"/></a:avLst>
               </a:prstGeom></a:spPr>"#
        ));
        match geom {
            ShapeGeom::Preset { adj, .. } => assert_eq!(adj, vec![Some(16667.0)]),
            other => panic!("expected Preset, got {other:?}"),
        }
    }
}

#[cfg(test)]
mod blip_svg_tests {
    // Microsoft 2016 SVG extension (MS-ODRAWXML) on a `<xdr:pic>` blip: the real
    // vector image rides in `<a:blip><a:extLst><a:ext
    // uri="{96DAC541-…}"><asvg:svgBlip r:embed>`, while `<a:blip r:embed>` is a
    // PNG/JPEG *fallback* (which may be absent for a pure-SVG insert). The parser
    // must keep emitting the raster fallback as `image_path` (regression-safe)
    // while additionally surfacing the SVG body in `svg_image_path`, and must
    // never drop a picture that carries only the svgBlip. The model now carries
    // zip paths + mime (no inlined base64); the renderer fetches bytes lazily via
    // `extract_image`. Mirrors the pptx/docx parser tests.
    use super::*;

    // 1×1 transparent PNG (smallest valid PNG).
    const PNG_1X1: &[u8] = &[
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44,
        0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F,
        0x15, 0xC4, 0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x62, 0x00,
        0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49,
        0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
    ];
    const SVG: &[u8] =
        br##"<svg xmlns="http://www.w3.org/2000/svg" width="2" height="2"><rect width="2" height="2" fill="#0a0"/></svg>"##;

    /// Build a tiny zip with the two media parts a `<xdr:pic>` blip references
    /// (a PNG fallback and an SVG body) under `xl/media/`, so
    /// `parse_drawing_anchors` can be driven with a hand-rolled rels map.
    fn build_media_zip(png: &[u8], svg: &[u8]) -> Vec<u8> {
        use std::io::Write;
        use zip::write::SimpleFileOptions;
        let mut buf = Vec::new();
        {
            let cursor = Cursor::new(&mut buf);
            let mut zw = zip::ZipWriter::new(cursor);
            let opts = SimpleFileOptions::default();
            zw.start_file("xl/media/image1.png", opts).unwrap();
            zw.write_all(png).unwrap();
            zw.start_file("xl/media/image2.svg", opts).unwrap();
            zw.write_all(svg).unwrap();
            zw.finish().unwrap();
        }
        buf
    }

    /// `<xdr:wsDr>` wrapping a single `twoCellAnchor` picture whose `<a:blip>`
    /// inner XML is supplied by the caller (so each test varies only the blip).
    fn drawing_xml(blip_inner: &str) -> String {
        format!(
            r#"<xdr:wsDr
  xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main">
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="2" name="P"/><xdr:cNvPicPr/></xdr:nvPicPr>
      <xdr:blipFill>{blip_inner}<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
      <xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="300000" cy="300000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#
        )
    }

    fn parse_one(blip_inner: &str, rels: &HashMap<String, String>) -> ImageAnchor {
        let xml = drawing_xml(blip_inner);
        let data = build_media_zip(PNG_1X1, SVG);
        let cursor = Cursor::new(data.clone());
        let mut archive = crate::XlsxZip::new(cursor).unwrap();
        let anchors = parse_drawing_anchors(&xml, rels, "xl/drawings", &mut archive, &[]);
        assert_eq!(anchors.len(), 1, "exactly one picture anchor expected");
        anchors.into_iter().next().unwrap()
    }

    /// §20.1.2.2.8 — a `<xdr:pic>` whose `<xdr:cNvPr hidden="1">` marks it hidden
    /// is not emitted as an image anchor. A visible pic still resolves.
    #[test]
    fn hidden_picture_is_not_emitted() {
        let mut rels = HashMap::new();
        rels.insert("rId1".to_string(), "../media/image1.png".to_string());
        let data = build_media_zip(PNG_1X1, SVG);
        for (hidden_attr, expect_len) in [(r#" hidden="1""#, 0usize), ("", 1usize)] {
            let xml = format!(
                r#"<xdr:wsDr
  xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="2" name="P"{hidden}/><xdr:cNvPicPr/></xdr:nvPicPr>
      <xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
      <xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="300000" cy="300000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#,
                hidden = hidden_attr,
            );
            let mut archive = crate::XlsxZip::new(Cursor::new(data.clone())).unwrap();
            let anchors = parse_drawing_anchors(&xml, &rels, "xl/drawings", &mut archive, &[]);
            assert_eq!(anchors.len(), expect_len, "hidden={hidden_attr}");
        }
    }

    /// A `<xdr:pic>` carrying a PNG fallback blip plus an `asvg:svgBlip` extension
    /// must keep the PNG on `image_path` and surface the SVG on `svg_image_path`,
    /// both as zip paths with mime derived from the extension (no base64). The
    /// svgBlip uses a different prefix (asvg:) on purpose — matching is by
    /// namespace-local name, so the prefix must not matter.
    #[test]
    fn picture_with_svg_blip_extension_emits_both_paths() {
        let blip = r#"<a:blip r:embed="rIdPng">
          <a:extLst>
            <a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">
              <asvg:svgBlip r:embed="rIdSvg"/>
            </a:ext>
          </a:extLst>
        </a:blip>"#;
        let mut rels = HashMap::new();
        rels.insert("rIdPng".to_string(), "../media/image1.png".to_string());
        rels.insert("rIdSvg".to_string(), "../media/image2.svg".to_string());

        let anchor = parse_one(blip, &rels);

        // PNG fallback preserved on image_path (regression-safe). No base64.
        assert_eq!(anchor.image_path, "xl/media/image1.png");
        assert_eq!(anchor.mime_type, "image/png");
        assert!(
            !anchor.image_path.contains(";base64,"),
            "image_path must be a zip path, not a data URL"
        );
        // The SVG part is surfaced separately as a zip path.
        assert_eq!(
            anchor.svg_image_path.as_deref(),
            Some("xl/media/image2.svg")
        );
        // native ext is retained alongside the new path refs.
        assert_eq!(anchor.native_ext_cx, 300000);
        assert_eq!(anchor.native_ext_cy, 300000);
    }

    /// End-to-end wiring guard: a `<xdr:pic>` whose `<xdr:blipFill>` carries an
    /// `<a:srcRect>` sibling of the blip must surface the crop on the parsed
    /// `ImageAnchor.src_rect` (sample-27's horizontal crop). Catches a regression
    /// that drops the `parse_src_rect(bf)` wiring from the pic branch.
    #[test]
    fn picture_with_src_rect_surfaces_crop() {
        let blip = r#"<a:blip r:embed="rIdPng"/><a:srcRect l="32560" r="3829"/>"#;
        let mut rels = HashMap::new();
        rels.insert("rIdPng".to_string(), "../media/image1.png".to_string());

        let anchor = parse_one(blip, &rels);

        let sr = anchor
            .src_rect
            .as_ref()
            .expect("srcRect surfaced on the anchor");
        assert!((sr.l - 0.3256).abs() < 1e-9);
        assert!((sr.r - 0.03829).abs() < 1e-9);
        assert_eq!(sr.t, 0.0);
        assert_eq!(sr.b, 0.0);

        // It serializes as camelCase `srcRect` so the TS renderer can read it.
        let json = serde_json::to_string(&anchor).unwrap();
        assert!(json.contains("\"srcRect\""), "emits srcRect: {json}");
    }

    /// Part 1 §20.1.7.6 (`xfrm`, CT_Transform2D) and §20.1.10.3 ST_Angle: `rot` uses
    /// signed 60000ths of a degree while flips are independent booleans. The
    /// transform remains meaningful when optional `ext` is absent because the
    /// spreadsheet anchor supplies the painted rectangle.
    #[test]
    fn picture_xfrm_preserves_rotation_and_flips_without_extent() {
        let mut rels = HashMap::new();
        rels.insert("rIdPng".to_string(), "../media/image1.png".to_string());
        let xml = drawing_xml(r#"<a:blip r:embed="rIdPng"/>"#)
            .replace(
                "<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"300000\" cy=\"300000\"/></a:xfrm>",
                "<a:xfrm rot=\"-5400000\" flipH=\"true\" flipV=\"1\"><a:off x=\"0\" y=\"0\"/></a:xfrm>",
            );
        let data = build_media_zip(PNG_1X1, SVG);
        let mut archive = crate::XlsxZip::new(Cursor::new(data)).unwrap();
        let anchor = parse_drawing_anchors(&xml, &rels, "xl/drawings", &mut archive, &[])
            .into_iter().next().unwrap();

        assert_eq!(anchor.native_ext_cx, 0);
        assert_eq!(anchor.native_ext_cy, 0);
        assert_eq!(anchor.rotation, Some(-90.0));
        assert_eq!(anchor.flip_h, Some(true));
        assert_eq!(anchor.flip_v, Some(true));
    }

    #[test]
    fn picture_xfrm_omits_identity_and_rejects_non_integer_rotation() {
        let mut rels = HashMap::new();
        rels.insert("rIdPng".to_string(), "../media/image1.png".to_string());
        for rot in ["0", "NaN", "inf", "5400000.5", "2147483648"] {
            let xml = drawing_xml(r#"<a:blip r:embed="rIdPng"/>"#)
                .replace("<a:xfrm>", &format!("<a:xfrm rot=\"{rot}\" flipH=\"false\" flipV=\"0\">"));
            let data = build_media_zip(PNG_1X1, SVG);
            let mut archive = crate::XlsxZip::new(Cursor::new(data)).unwrap();
            let anchor = parse_drawing_anchors(&xml, &rels, "xl/drawings", &mut archive, &[])
                .into_iter().next().unwrap();
            assert_eq!(anchor.rotation, None, "rot={rot}");
            assert_eq!(anchor.flip_h, None);
            assert_eq!(anchor.flip_v, None);
            let json = serde_json::to_string(&anchor).unwrap();
            assert!(!json.contains("rotation"), "rot={rot}: {json}");
            assert!(!json.contains("flipH"), "rot={rot}: {json}");
            assert!(!json.contains("flipV"), "rot={rot}: {json}");
        }
    }

    /// An uncropped `<xdr:pic>` (no `<a:srcRect>`) leaves `src_rect == None`, and
    /// the serialized JSON omits the key entirely (skip_serializing_if), so the
    /// common case stays on the cheap full-blip draw path.
    #[test]
    fn picture_without_src_rect_omits_crop() {
        let blip = r#"<a:blip r:embed="rIdPng"/>"#;
        let mut rels = HashMap::new();
        rels.insert("rIdPng".to_string(), "../media/image1.png".to_string());

        let anchor = parse_one(blip, &rels);

        assert!(anchor.src_rect.is_none());
        let json = serde_json::to_string(&anchor).unwrap();
        assert!(
            !json.contains("srcRect"),
            "omits srcRect when absent: {json}"
        );
    }

    /// A `<xdr:pic>` whose `<a:blip>` carries ONLY the `asvg:svgBlip` extension —
    /// no raster `r:embed` fallback (an icon inserted as a pure SVG) — must still
    /// parse. Previously the media filter excluded `.svg` and the resolution
    /// required a raster embed, so the whole picture was silently dropped.
    #[test]
    fn picture_with_only_svg_blip_and_no_raster_embed_still_parses() {
        let blip = r#"<a:blip>
          <a:extLst>
            <a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">
              <asvg:svgBlip r:embed="rIdSvg"/>
            </a:ext>
          </a:extLst>
        </a:blip>"#;
        let mut rels = HashMap::new();
        rels.insert("rIdSvg".to_string(), "../media/image2.svg".to_string());

        let anchor = parse_one(blip, &rels);

        // With no raster embed, image_path falls back to the SVG path itself so
        // the element is always drawable …
        assert_eq!(
            anchor.image_path, "xl/media/image2.svg",
            "image_path must fall back to the SVG path when no raster is embedded"
        );
        assert_eq!(
            anchor.mime_type, "image/svg+xml",
            "mime must follow the fallback SVG path's extension"
        );
        // … and svg_image_path carries the same vector source.
        assert_eq!(
            anchor.svg_image_path.as_deref(),
            Some("xl/media/image2.svg"),
            "svg_image_path must be Some for a pure-SVG picture"
        );
    }

    /// A plain `<xdr:pic>` with no svgBlip extension must leave `svg_image_path`
    /// as None (and still emit the PNG `image_path`) — guards against the new
    /// branch firing spuriously.
    #[test]
    fn picture_without_svg_blip_has_no_svg_path() {
        let blip = r#"<a:blip r:embed="rIdPng"/>"#;
        let mut rels = HashMap::new();
        rels.insert("rIdPng".to_string(), "../media/image1.png".to_string());

        let anchor = parse_one(blip, &rels);
        assert_eq!(anchor.image_path, "xl/media/image1.png");
        assert_eq!(anchor.mime_type, "image/png");
        assert!(
            anchor.svg_image_path.is_none(),
            "svg_image_path must be None without an svgBlip extension"
        );
    }

    /// Struct-serialize guard: `ImageAnchor` serializes the new camelCase path
    /// refs (`imagePath`/`mimeType`/`svgImagePath`) and retains
    /// `nativeExtCx`/`nativeExtCy`, and never emits the old `dataUrl` /
    /// `;base64,` form. Fixture-free so it always runs in CI.
    #[test]
    fn image_anchor_serializes_path_refs_not_data_url() {
        let anchor = ImageAnchor {
            z_order: 42,
            from_col: 1,
            from_col_off: 0,
            from_row: 1,
            from_row_off: 0,
            to_col: 3,
            to_col_off: 0,
            to_row: 5,
            to_row_off: 0,
            edit_as: Some("oneCell".to_string()),
            native_ext_cx: 300000,
            native_ext_cy: 200000,
            rotation: None,
            flip_h: None,
            flip_v: None,
            image_path: "xl/media/image1.png".to_string(),
            mime_type: "image/png".to_string(),
            svg_image_path: Some("xl/media/image2.svg".to_string()),
            src_rect: None,
            alpha: None,
            duotone: None,
        };
        let json = serde_json::to_string(&anchor).unwrap();
        assert!(json.contains("\"imagePath\":\"xl/media/image1.png\""));
        assert!(json.contains("\"mimeType\":\"image/png\""));
        assert!(json.contains("\"svgImagePath\":\"xl/media/image2.svg\""));
        assert!(json.contains("\"nativeExtCx\":300000"));
        assert!(json.contains("\"nativeExtCy\":200000"));
        assert!(!json.contains("dataUrl"), "must not emit dataUrl: {json}");
        assert!(!json.contains(";base64,"), "must not inline base64: {json}");
    }

    /// Struct-serialize guard for the group-leaf `ShapeGeom::Image` variant:
    /// same camelCase path refs, no `dataUrl` / base64.
    #[test]
    fn shape_geom_image_serializes_path_refs_not_data_url() {
        let geom = ShapeGeom::Image {
            image_path: "xl/media/image1.png".to_string(),
            mime_type: "image/png".to_string(),
            svg_image_path: None,
            src_rect: None,
            alpha: None,
            duotone: None,
        };
        let json = serde_json::to_string(&geom).unwrap();
        assert!(json.contains("\"type\":\"image\""));
        assert!(json.contains("\"imagePath\":\"xl/media/image1.png\""));
        assert!(json.contains("\"mimeType\":\"image/png\""));
        assert!(!json.contains("dataUrl"), "must not emit dataUrl: {json}");
        assert!(!json.contains(";base64,"), "must not inline base64: {json}");
    }
}

#[cfg(test)]
mod custom_path_arc_tests {
    use super::*;

    const NS: &str = r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main""#;

    /// `PathCmd`'s enum-level `rename_all = "camelCase"` (tag = "op") renames
    /// only the variant tags, not the fields. Without a per-variant
    /// `rename_all`, `ArcTo { st_ang, sw_ang }` serialized as snake_case keys
    /// `st_ang`/`sw_ang`, which the TS renderer never reads — it reads
    /// `cmd.stAng`/`cmd.swAng` (`renderer.ts`), so the values came back
    /// `undefined` → `NaN` and the arc failed to draw. Same root cause as the
    /// pptx fix in PR #489. This guards the JSON keys the renderer relies on.
    #[test]
    fn arc_to_serializes_angle_fields_as_camelcase() {
        // Non-degenerate arc (wR/hR > 0) — the renderer's degenerate-arc guard
        // (`if (rx <= 0 || ry <= 0) break;`) short-circuits before the angles
        // are read, so only non-degenerate arcs surface the missing keys.
        let xml = format!(
            r#"<a:custGeom {NS}><a:pathLst><a:path w="100" h="100">
                 <a:moveTo><a:pt x="0" y="50"/></a:moveTo>
                 <a:arcTo wR="50" hR="50" stAng="0" swAng="5400000"/>
               </a:path></a:pathLst></a:custGeom>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let path = parse_custom_paths(doc.root_element(), 100.0, 100.0)
            .into_iter()
            .next()
            .expect("custom geometry contains one path");

        // The parser keeps angles as raw 60000ths of a degree (the xlsx
        // convention — the TS renderer divides by 60000). swAng = 5_400_000 =
        // 90°. wR/hR stay in path-coord units.
        let arc = path
            .commands
            .iter()
            .find(|c| matches!(c, PathCmd::ArcTo { .. }))
            .expect("path contains an arcTo command");
        match arc {
            PathCmd::ArcTo {
                wr,
                hr,
                st_ang,
                sw_ang,
            } => {
                assert_eq!(*wr, 50.0);
                assert_eq!(*hr, 50.0);
                assert_eq!(*st_ang, 0.0);
                assert_eq!(
                    *sw_ang, 5_400_000.0,
                    "raw 60000ths preserved (renderer divides by 60000)"
                );
            }
            other => panic!("expected ArcTo, got {other:?}"),
        }

        let value = serde_json::to_value(arc).expect("arcTo serializes");
        let obj = value
            .as_object()
            .expect("arcTo serializes to a JSON object");

        // Tag key is "op" (not "cmd"), variant tag is camelCase "arcTo".
        assert_eq!(
            obj.get("op").and_then(|v| v.as_str()),
            Some("arcTo"),
            "variant tag serializes under key \"op\" as camelCase \"arcTo\""
        );

        // The angle fields must be camelCase — the TS renderer reads
        // cmd.stAng / cmd.swAng. snake_case keys here regress the bug.
        assert!(
            obj.contains_key("stAng"),
            "stAng must be camelCase, got keys: {:?}",
            obj.keys().collect::<Vec<_>>()
        );
        assert!(
            obj.contains_key("swAng"),
            "swAng must be camelCase, got keys: {:?}",
            obj.keys().collect::<Vec<_>>()
        );
        assert!(
            !obj.contains_key("st_ang") && !obj.contains_key("sw_ang"),
            "snake_case angle keys must not appear (regresses the NaN bug)"
        );
        assert_eq!(obj.get("swAng").and_then(|v| v.as_f64()), Some(5_400_000.0));
    }
}

#[cfg(test)]
mod src_rect_tests {
    use super::*;

    const NS: &str = r#"xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;

    /// sample-27: `<a:srcRect l="32560" r="3829"/>` ⇒ left 0.3256, right 0.03829,
    /// top/bottom 0 (absent edges default 0). Edge attrs are ST_Percentage in
    /// 1000ths of a percent, so the fraction is the raw value / 100000.
    #[test]
    fn parses_horizontal_crop_fractions() {
        let xml = format!(
            r#"<xdr:blipFill {NS}>
              <a:blip r:embed="rId1"/>
              <a:srcRect l="32560" r="3829"/>
              <a:stretch><a:fillRect/></a:stretch>
            </xdr:blipFill>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let sr = parse_src_rect(doc.root_element()).expect("srcRect present");
        assert!((sr.l - 0.3256).abs() < 1e-9, "l = l_attr / 100000");
        assert!((sr.r - 0.03829).abs() < 1e-9, "r = r_attr / 100000");
        assert_eq!(sr.t, 0.0, "absent top defaults to 0");
        assert_eq!(sr.b, 0.0, "absent bottom defaults to 0");
    }

    /// No `<a:srcRect>` ⇒ None (uncropped picture, the common case).
    #[test]
    fn absent_src_rect_is_none() {
        let xml = format!(
            r#"<xdr:blipFill {NS}><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        assert!(parse_src_rect(doc.root_element()).is_none());
    }

    /// An all-zero `<a:srcRect>` ⇒ None: an explicit no-op crop must not push the
    /// renderer onto the 9-arg sub-rect path (which would be an identity draw).
    #[test]
    fn all_zero_src_rect_is_none() {
        let xml = format!(
            r#"<xdr:blipFill {NS}><a:blip r:embed="rId1"/><a:srcRect l="0" t="0" r="0" b="0"/></xdr:blipFill>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        assert!(parse_src_rect(doc.root_element()).is_none());
    }
}

#[cfg(test)]
mod ole_object_tests {
    use super::*;
    use std::io::{Cursor, Write};

    // A minimal EMF-signature blob is unnecessary here — `parse_ole_object_anchors`
    // only needs the media part to EXIST in the archive (index_for_name), the
    // renderer decodes lazily. Any bytes suffice.
    fn archive_with(parts: &[(&str, &[u8])]) -> crate::XlsxZip {
        use zip::write::SimpleFileOptions;
        let mut buf = Vec::new();
        {
            let mut zw = zip::ZipWriter::new(Cursor::new(&mut buf));
            let o = SimpleFileOptions::default();
            for (path, bytes) in parts {
                zw.start_file(*path, o).unwrap();
                zw.write_all(bytes).unwrap();
            }
            zw.finish().unwrap();
        }
        crate::XlsxZip::new(Cursor::new(buf)).unwrap()
    }

    const OLE_NS: &str = concat!(
        r#"xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" "#,
        r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" "#,
        r#"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" "#,
        // x14 (spreadsheet 2009/9) is the namespace Excel binds on the real
        // `<mc:Choice Requires="x14">` around an OLE object; unk is a namespace
        // no parser understands. Declaring both lets `select_alternate_content`
        // resolve the Requires prefixes to URIs.
        r#"xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" "#,
        r#"xmlns:unk="http://example.com/an/extension/we/do/not/understand""#,
    );

    /// A `<mc:AlternateContent>` OLE wrapper whose Choice (behind `requires`) and
    /// Fallback each carry an `<oleObject>` with a resolvable image `objectPr`
    /// preview — `rIdChoice` → choice.emf, `rIdFallback` → fallback.emf. Returns
    /// the emitted anchors so a test can assert WHICH branch was selected.
    fn ole_alt_content_anchors(requires: &str) -> Vec<ImageAnchor> {
        let object = |rid_prev: &str| {
            format!(
                r#"<oleObject progId="Excel.Sheet.12" shapeId="1025" r:id="rIdData">
                     <objectPr defaultSize="0" r:id="{rid_prev}">
                       <anchor moveWithCells="1">
                         <from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></from>
                         <to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>7</xdr:row><xdr:rowOff>0</xdr:rowOff></to>
                       </anchor>
                     </objectPr>
                   </oleObject>"#
            )
        };
        let sheet_xml = format!(
            r#"<worksheet {ns}><sheetData/>
              <oleObjects>
                <mc:AlternateContent>
                  <mc:Choice Requires="{requires}">{choice}</mc:Choice>
                  <mc:Fallback>{fallback}</mc:Fallback>
                </mc:AlternateContent>
              </oleObjects>
            </worksheet>"#,
            ns = OLE_NS,
            requires = requires,
            choice = object("rIdChoice"),
            fallback = object("rIdFallback"),
        );
        let mut rels = HashMap::new();
        rels.insert("rIdChoice".to_string(), "../media/choice.emf".to_string());
        rels.insert(
            "rIdFallback".to_string(),
            "../media/fallback.emf".to_string(),
        );
        let mut archive = archive_with(&[
            ("xl/media/choice.emf", b"emf-c"),
            ("xl/media/fallback.emf", b"emf-f"),
        ]);
        parse_ole_object_anchors(&sheet_xml, &rels, "xl/worksheets", &mut archive)
    }

    /// ECMA-376 Part 3 §9.3(1) — an `<oleObject>` Choice whose `Requires` names a
    /// namespace the parser does NOT understand must not be selected, even though
    /// its `objectPr` preview resolves; the `<mc:Fallback>`'s (equally drawable)
    /// oleObject is selected instead. The old collector always took the
    /// non-Fallback twin, silently dropping a renderable Fallback.
    #[test]
    fn ole_unhandled_choice_selects_fallback_preview() {
        let anchors = ole_alt_content_anchors("unk");
        assert_eq!(anchors.len(), 1, "exactly one preview anchor");
        assert_eq!(
            anchors[0].image_path, "xl/media/fallback.emf",
            "un-understood Choice must yield the Fallback preview, not the Choice's"
        );
    }

    /// §9.3(1) — the `x14` namespace IS understood by the xlsx parser (it drives
    /// slicers and the OLE `objectPr` preview), so a real
    /// `<mc:Choice Requires="x14">` is selected and the Fallback dropped. Guards
    /// against a regression on Excel's canonical OLE output.
    #[test]
    fn ole_understood_x14_choice_selected_over_fallback() {
        let anchors = ole_alt_content_anchors("x14");
        assert_eq!(anchors.len(), 1, "exactly one preview anchor");
        assert_eq!(
            anchors[0].image_path, "xl/media/choice.emf",
            "understood x14 Choice must be selected, Fallback dropped"
        );
    }

    /// A worksheet `<oleObjects>` wrapped in `mc:AlternateContent` where the
    /// Choice carries `<oleObject><objectPr r:id><anchor><from>/<to>>` and the
    /// Fallback carries a bare `<oleObject>` (no objectPr). Exactly ONE preview
    /// ImageAnchor must be emitted — the Choice's — never a second from Fallback.
    #[test]
    fn ole_object_emits_single_preview_anchor_from_choice() {
        let sheet_xml = format!(
            r#"<worksheet {ns}><sheetData/>
              <oleObjects>
                <mc:AlternateContent>
                  <mc:Choice Requires="x14">
                    <oleObject progId="Excel.Sheet.12" shapeId="1025" r:id="rIdData">
                      <objectPr defaultSize="0" r:id="rIdPrev">
                        <anchor moveWithCells="1">
                          <from><xdr:col>1</xdr:col><xdr:colOff>10</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>20</xdr:rowOff></from>
                          <to><xdr:col>4</xdr:col><xdr:colOff>30</xdr:colOff><xdr:row>7</xdr:row><xdr:rowOff>40</xdr:rowOff></to>
                        </anchor>
                      </objectPr>
                    </oleObject>
                  </mc:Choice>
                  <mc:Fallback>
                    <oleObject progId="Excel.Sheet.12" shapeId="1025" r:id="rIdData"/>
                  </mc:Fallback>
                </mc:AlternateContent>
              </oleObjects>
            </worksheet>"#,
            ns = OLE_NS,
        );
        let mut rels = HashMap::new();
        rels.insert("rIdPrev".to_string(), "../media/image1.emf".to_string());
        let mut archive = archive_with(&[("xl/media/image1.emf", b"emfbytes")]);

        let anchors = parse_ole_object_anchors(&sheet_xml, &rels, "xl/worksheets", &mut archive);
        assert_eq!(
            anchors.len(),
            1,
            "Choice+Fallback must yield exactly one preview anchor, got {}",
            anchors.len()
        );
        let a = &anchors[0];
        assert_eq!(a.image_path, "xl/media/image1.emf");
        assert_eq!(a.mime_type, "image/emf");
        assert_eq!((a.from_col, a.from_row), (1, 2));
        assert_eq!((a.from_col_off, a.from_row_off), (10, 20));
        assert_eq!((a.to_col, a.to_row), (4, 7));
        assert_eq!((a.to_col_off, a.to_row_off), (30, 40));
    }

    /// A bare `<oleObjects>` (not wrapped in AlternateContent) with a single
    /// `<oleObject>` whose `objectPr@r:id` resolves is emitted directly.
    #[test]
    fn unwrapped_ole_object_emits_anchor() {
        let sheet_xml = format!(
            r#"<worksheet {ns}><sheetData/>
              <oleObjects>
                <oleObject progId="Equation.3" shapeId="2" r:id="rIdData">
                  <objectPr r:id="rIdPrev">
                    <anchor>
                      <from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></from>
                      <to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></to>
                    </anchor>
                  </objectPr>
                </oleObject>
              </oleObjects>
            </worksheet>"#,
            ns = OLE_NS,
        );
        let mut rels = HashMap::new();
        rels.insert("rIdPrev".to_string(), "../media/image2.wmf".to_string());
        let mut archive = archive_with(&[("xl/media/image2.wmf", b"wmf")]);

        let anchors = parse_ole_object_anchors(&sheet_xml, &rels, "xl/worksheets", &mut archive);
        assert_eq!(anchors.len(), 1);
        assert_eq!(anchors[0].image_path, "xl/media/image2.wmf");
        assert_eq!(anchors[0].mime_type, "image/wmf");
    }

    /// An `<oleObject>` whose `<objectPr>` is absent, or whose `r:id` doesn't
    /// resolve, is skipped (no zero-sized / path-less anchor).
    #[test]
    fn ole_object_without_objectpr_or_resolvable_rid_is_skipped() {
        let sheet_xml = format!(
            r#"<worksheet {ns}><sheetData/>
              <oleObjects>
                <oleObject shapeId="3" r:id="rIdData"/>
                <oleObject shapeId="4" r:id="rIdData2">
                  <objectPr r:id="rIdMissing">
                    <anchor><from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></from>
                    <to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></to></anchor>
                  </objectPr>
                </oleObject>
              </oleObjects>
            </worksheet>"#,
            ns = OLE_NS,
        );
        let rels = HashMap::new(); // rIdMissing not present
        let mut archive = archive_with(&[]);
        let anchors = parse_ole_object_anchors(&sheet_xml, &rels, "xl/worksheets", &mut archive);
        assert!(anchors.is_empty(), "no resolvable preview ⇒ no anchor");
    }

    /// Defensive guard (§18.3.1.56): when `objectPr@r:id` resolves to the object
    /// *data* part (a `.bin` ⇒ `application/octet-stream`) rather than an image,
    /// it must be skipped so object data never enters the picture pipeline.
    #[test]
    fn ole_object_pointing_at_bin_data_part_is_skipped() {
        let sheet_xml = format!(
            r#"<worksheet {ns}><sheetData/>
              <oleObjects>
                <oleObject progId="Excel.Sheet.12" shapeId="5" r:id="rIdData">
                  <objectPr r:id="rIdBin">
                    <anchor>
                      <from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></from>
                      <to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></to>
                    </anchor>
                  </objectPr>
                </oleObject>
              </oleObjects>
            </worksheet>"#,
            ns = OLE_NS,
        );
        let mut rels = HashMap::new();
        // objectPr@r:id points at the embedded object data (.bin), which is what
        // the ECMA-376 §18.3.1.56 text actually says this relationship targets.
        rels.insert(
            "rIdBin".to_string(),
            "../embeddings/oleObject1.bin".to_string(),
        );
        let mut archive = archive_with(&[("xl/embeddings/oleObject1.bin", b"\x00\x01datablob")]);

        let anchors = parse_ole_object_anchors(&sheet_xml, &rels, "xl/worksheets", &mut archive);
        assert!(
            anchors.is_empty(),
            "a .bin (non-image) objectPr target must be skipped, got {} anchor(s)",
            anchors.len()
        );
    }

    // ─── vmlDrawing preview path (real Excel output) ────────────────────────

    /// A standalone legacy VML drawing part, root `<xml>` (no `<?xml?>`
    /// declaration — exactly how Excel writes it). `shapes` is the raw inner
    /// markup for the `<v:shape>`s under the root.
    fn vml_part(shapes: &str) -> String {
        format!(
            concat!(
                r#"<xml xmlns:v="urn:schemas-microsoft-com:vml" "#,
                r#"xmlns:o="urn:schemas-microsoft-com:office:office" "#,
                r#"xmlns:x="urn:schemas-microsoft-com:office:excel" "#,
                r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{}</xml>"#
            ),
            shapes
        )
    }

    fn rels_part(pairs: &[(&str, &str)]) -> String {
        let mut s = String::from(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        );
        for (id, target) in pairs {
            s.push_str(&format!(
                r#"<Relationship Id="{id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{target}"/>"#
            ));
        }
        s.push_str("</Relationships>");
        s
    }

    /// Canonical Excel structure: `<oleObject shapeId>` (no image-typed
    /// objectPr) + `<legacyDrawing r:id>` → vmlDrawing part → VML rels → EMF.
    /// The `<v:imagedata o:relid>` preview must surface as one ImageAnchor.
    #[test]
    fn vml_drawing_preview_keyed_by_shape_id_emits_anchor() {
        let sheet_xml = format!(
            r#"<worksheet {ns}><sheetData/>
              <oleObjects>
                <oleObject progId="Package" shapeId="1025" r:id="rIdData">
                  <objectPr defaultSize="0" r:id="rIdData"/>
                </oleObject>
              </oleObjects>
              <legacyDrawing r:id="rIdVml"/>
            </worksheet>"#,
            ns = OLE_NS,
        );
        // Sheet rels: the objectPr@r:id points at object DATA (.bin, rejected),
        // the legacyDrawing points at the vmlDrawing part.
        let mut sheet_rels = HashMap::new();
        sheet_rels.insert(
            "rIdData".to_string(),
            "../embeddings/oleObject1.bin".to_string(),
        );
        sheet_rels.insert(
            "rIdVml".to_string(),
            "../drawings/vmlDrawing1.vml".to_string(),
        );

        let vml = vml_part(
            r##"<v:shape id="_x0000_s1025" type="#_x0000_t75" style='width:96pt;height:48pt'>
                 <v:imagedata o:relid="rIdImg" o:title=""/>
                 <x:ClientData ObjectType="Pict">
                   <x:Anchor>1, 5, 2, 6, 4, 7, 7, 8</x:Anchor>
                 </x:ClientData>
               </v:shape>"##,
        );
        let vml_rels = rels_part(&[("rIdImg", "../media/image1.emf")]);
        let mut archive = archive_with(&[
            ("xl/embeddings/oleObject1.bin", b"\x00datablob"),
            ("xl/drawings/vmlDrawing1.vml", vml.as_bytes()),
            (
                "xl/drawings/_rels/vmlDrawing1.vml.rels",
                vml_rels.as_bytes(),
            ),
            ("xl/media/image1.emf", b"emfbytes"),
        ]);

        let anchors =
            parse_ole_object_anchors(&sheet_xml, &sheet_rels, "xl/worksheets", &mut archive);
        assert_eq!(anchors.len(), 1, "one vmlDrawing preview anchor expected");
        let a = &anchors[0];
        assert_eq!(a.image_path, "xl/media/image1.emf");
        assert_eq!(a.mime_type, "image/emf");
        // x:Anchor CSV: L=1,LOff=5,T=2,TOff=6,R=4,ROff=7,B=7,BOff=8 (px→EMU ×9525)
        assert_eq!((a.from_col, a.from_row), (1, 2));
        assert_eq!(
            (a.from_col_off, a.from_row_off),
            (5 * 9525, 6 * 9525),
            "x:Anchor px offsets convert to EMU via ×9525"
        );
        assert_eq!((a.to_col, a.to_row), (4, 7));
        assert_eq!((a.to_col_off, a.to_row_off), (7 * 9525, 8 * 9525));
    }

    /// A pathologically large `<x:Anchor>` pixel offset (e.g. a malformed or
    /// hostile CSV) must saturate the px→EMU conversion rather than panic
    /// (debug overflow check) or silently wrap (release build).
    #[test]
    fn vml_client_anchor_px_to_emu_saturates_on_overflow() {
        let xml = format!(
            r#"<x:ClientData xmlns:x="urn:schemas-microsoft-com:office:excel" ObjectType="Pict">
                 <x:Anchor>0, {big}, 0, 0, 1, 0, 1, 0</x:Anchor>
               </x:ClientData>"#,
            big = i64::MAX
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let anchor = parse_vml_client_anchor(&doc.root_element())
            .expect("8 numeric fields must parse even when one overflows on conversion");
        assert_eq!(
            anchor.from_col_off,
            i64::MAX,
            "saturating_mul must clamp to i64::MAX instead of panicking or wrapping"
        );
        // The other (well-behaved) offsets are unaffected.
        assert_eq!(anchor.to_col_off, 0);
    }

    #[test]
    fn drawing_marker_indices_preserve_conforming_off_grid_xsd_int_values() {
        assert_eq!(parse_xdr_col("16384"), 16_384);
        assert_eq!(parse_xdr_row("1048576"), 1_048_576);
        assert_eq!(parse_xdr_col(&i32::MAX.to_string()), i32::MAX as u32);
        assert_eq!(parse_xdr_row("-1"), 0);
        assert_eq!(parse_xdr_col("2147483648"), 0);
    }

    #[test]
    fn vml_anchor_preserves_conforming_off_grid_indices_without_wrapping() {
        let xml = r#"<x:ClientData xmlns:x="urn:schemas-microsoft-com:office:excel" ObjectType="Pict">
          <x:Anchor>16384, 0, 1048576, 0, 20000, 0, 2000000, 0</x:Anchor>
        </x:ClientData>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let anchor = parse_vml_client_anchor(&doc.root_element()).unwrap();
        assert_eq!((anchor.from_col, anchor.from_row), (16_384, 1_048_576));
        assert_eq!((anchor.to_col, anchor.to_row), (20_000, 2_000_000));
    }

    /// The `<objectPr><anchor>` two-cell markers (EMU, exact) win over the VML
    /// `<x:Anchor>` pixel CSV when both are present, but the preview IMAGE still
    /// comes from the vmlDrawing (objectPr@r:id is object data).
    #[test]
    fn objectpr_anchor_wins_over_vml_anchor() {
        let sheet_xml = format!(
            r#"<worksheet {ns}><sheetData/>
              <oleObjects>
                <oleObject shapeId="1026" r:id="rIdData">
                  <objectPr r:id="rIdData">
                    <anchor>
                      <from><xdr:col>3</xdr:col><xdr:colOff>111</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>222</xdr:rowOff></from>
                      <to><xdr:col>8</xdr:col><xdr:colOff>333</xdr:colOff><xdr:row>9</xdr:row><xdr:rowOff>444</xdr:rowOff></to>
                    </anchor>
                  </objectPr>
                </oleObject>
              </oleObjects>
              <legacyDrawing r:id="rIdVml"/>
            </worksheet>"#,
            ns = OLE_NS,
        );
        let mut sheet_rels = HashMap::new();
        sheet_rels.insert(
            "rIdData".to_string(),
            "../embeddings/oleObject1.bin".to_string(),
        );
        sheet_rels.insert(
            "rIdVml".to_string(),
            "../drawings/vmlDrawing1.vml".to_string(),
        );
        let vml = vml_part(
            r##"<v:shape id="_x0000_s1026" type="#_x0000_t75">
                 <v:imagedata o:relid="rIdImg"/>
                 <x:ClientData ObjectType="Pict"><x:Anchor>0, 0, 0, 0, 1, 0, 1, 0</x:Anchor></x:ClientData>
               </v:shape>"##,
        );
        let vml_rels = rels_part(&[("rIdImg", "../media/image1.wmf")]);
        let mut archive = archive_with(&[
            ("xl/embeddings/oleObject1.bin", b"data"),
            ("xl/drawings/vmlDrawing1.vml", vml.as_bytes()),
            (
                "xl/drawings/_rels/vmlDrawing1.vml.rels",
                vml_rels.as_bytes(),
            ),
            ("xl/media/image1.wmf", b"wmf"),
        ]);
        let anchors =
            parse_ole_object_anchors(&sheet_xml, &sheet_rels, "xl/worksheets", &mut archive);
        assert_eq!(anchors.len(), 1);
        let a = &anchors[0];
        assert_eq!(a.image_path, "xl/media/image1.wmf");
        // objectPr EMU anchor wins — NOT the VML pixel CSV.
        assert_eq!(
            (a.from_col, a.from_col_off, a.from_row, a.from_row_off),
            (3, 111, 3, 222)
        );
        assert_eq!(
            (a.to_col, a.to_col_off, a.to_row, a.to_row_off),
            (8, 333, 9, 444)
        );
    }

    /// A comments-only vmlDrawing (`type="#_x0000_t202"` / `ObjectType="Note"`,
    /// no `<v:imagedata>`) whose shapeId never matches an oleObject must emit
    /// nothing. Guards against the note VML being mistaken for a preview.
    #[test]
    fn comment_only_vml_emits_nothing() {
        let sheet_xml = format!(
            r#"<worksheet {ns}><sheetData/>
              <oleObjects>
                <oleObject shapeId="1025" r:id="rIdData">
                  <objectPr r:id="rIdData"/>
                </oleObject>
              </oleObjects>
              <legacyDrawing r:id="rIdVml"/>
            </worksheet>"#,
            ns = OLE_NS,
        );
        let mut sheet_rels = HashMap::new();
        sheet_rels.insert(
            "rIdData".to_string(),
            "../embeddings/oleObject1.bin".to_string(),
        );
        sheet_rels.insert(
            "rIdVml".to_string(),
            "../drawings/vmlDrawing1.vml".to_string(),
        );
        // A Note shape with a DIFFERENT shapeId (6146) and no imagedata.
        let vml = vml_part(
            r##"<v:shape id="_x0000_s6146" type="#_x0000_t202" style='visibility:hidden'>
                 <v:textbox><div></div></v:textbox>
                 <x:ClientData ObjectType="Note"><x:Anchor>11, 15, 4, 7, 22, 11, 7, 15</x:Anchor></x:ClientData>
               </v:shape>"##,
        );
        let mut archive = archive_with(&[
            ("xl/embeddings/oleObject1.bin", b"data"),
            ("xl/drawings/vmlDrawing1.vml", vml.as_bytes()),
        ]);
        let anchors =
            parse_ole_object_anchors(&sheet_xml, &sheet_rels, "xl/worksheets", &mut archive);
        assert!(
            anchors.is_empty(),
            "note VML must not be picked up as an OLE preview, got {}",
            anchors.len()
        );
    }

    /// shapeId mismatch, missing imagedata, and a missing VML rels part each
    /// skip cleanly (no anchor, no panic).
    #[test]
    fn vml_preview_skips_on_mismatch_or_missing_rels() {
        // Case A: shapeId in the sheet (999) does not match the VML shape (1025).
        let sheet = |sid: &str| {
            format!(
                r#"<worksheet {ns}><sheetData/>
                  <oleObjects><oleObject shapeId="{sid}" r:id="rIdData"><objectPr r:id="rIdData"/></oleObject></oleObjects>
                  <legacyDrawing r:id="rIdVml"/>
                </worksheet>"#,
                ns = OLE_NS,
                sid = sid,
            )
        };
        let mut sheet_rels = HashMap::new();
        sheet_rels.insert(
            "rIdData".to_string(),
            "../embeddings/oleObject1.bin".to_string(),
        );
        sheet_rels.insert(
            "rIdVml".to_string(),
            "../drawings/vmlDrawing1.vml".to_string(),
        );
        let vml_ok = vml_part(
            r##"<v:shape id="_x0000_s1025" type="#_x0000_t75"><v:imagedata o:relid="rIdImg"/>
                 <x:ClientData ObjectType="Pict"><x:Anchor>0,0,0,0,1,0,1,0</x:Anchor></x:ClientData></v:shape>"##,
        );
        let vml_rels = rels_part(&[("rIdImg", "../media/image1.emf")]);

        // A: mismatch → nothing.
        let mut arch_a = archive_with(&[
            ("xl/embeddings/oleObject1.bin", b"d"),
            ("xl/drawings/vmlDrawing1.vml", vml_ok.as_bytes()),
            (
                "xl/drawings/_rels/vmlDrawing1.vml.rels",
                vml_rels.as_bytes(),
            ),
            ("xl/media/image1.emf", b"emf"),
        ]);
        assert!(
            parse_ole_object_anchors(&sheet("999"), &sheet_rels, "xl/worksheets", &mut arch_a)
                .is_empty(),
            "shapeId mismatch ⇒ no anchor"
        );

        // B: matching shape but NO imagedata → nothing.
        let vml_no_img = vml_part(
            r##"<v:shape id="_x0000_s1025" type="#_x0000_t75"><x:ClientData ObjectType="Pict"><x:Anchor>0,0,0,0,1,0,1,0</x:Anchor></x:ClientData></v:shape>"##,
        );
        let mut arch_b = archive_with(&[
            ("xl/embeddings/oleObject1.bin", b"d"),
            ("xl/drawings/vmlDrawing1.vml", vml_no_img.as_bytes()),
            (
                "xl/drawings/_rels/vmlDrawing1.vml.rels",
                vml_rels.as_bytes(),
            ),
            ("xl/media/image1.emf", b"emf"),
        ]);
        assert!(
            parse_ole_object_anchors(&sheet("1025"), &sheet_rels, "xl/worksheets", &mut arch_b)
                .is_empty(),
            "no imagedata ⇒ no anchor"
        );

        // C: imagedata present but the VML rels part is absent → relid unresolved.
        let mut arch_c = archive_with(&[
            ("xl/embeddings/oleObject1.bin", b"d"),
            ("xl/drawings/vmlDrawing1.vml", vml_ok.as_bytes()),
            ("xl/media/image1.emf", b"emf"),
        ]);
        assert!(
            parse_ole_object_anchors(&sheet("1025"), &sheet_rels, "xl/worksheets", &mut arch_c)
                .is_empty(),
            "missing vmlDrawing rels ⇒ relid unresolved ⇒ no anchor"
        );
    }

    /// Declaration-less VML (root `<xml>`, no `<?xml?>`), as some producers
    /// write it, must parse with roxmltree without any preprocessing. This is
    /// the invariant `load_legacy_vml_drawing` relies on.
    #[test]
    fn roxmltree_parses_declarationless_vml() {
        let vml = vml_part(
            r##"<v:shape id="_x0000_s1025" type="#_x0000_t75"><v:imagedata o:relid="rIdImg"/></v:shape>"##,
        );
        assert!(
            !vml.starts_with("<?xml"),
            "fixture must have no XML declaration"
        );
        let doc = roxmltree::Document::parse(&vml).expect("declaration-less VML must parse");
        assert_eq!(doc.root_element().tag_name().name(), "xml");
    }
}

/// ISO/IEC 29500 Strict-conformance fixture (`fix(xlsx): accept Strict
/// namespace URIs across the parser` routed `parse_drawing_anchors`'s
/// `xdr:`/`a:` element matching through `is_xdr_ns`/`is_a_ns`, and the blip's
/// `r:embed` through the shared `blip_embed_rid`/`attr_ns` helpers). Before
/// that conversion `<xdr:twoCellAnchor>` was found via a hardcoded
/// Transitional URI, so a Strict `xl/drawings/drawing1.xml` — declaring
/// `xdr:`/`a:`/`r:` under the `http://purl.oclc.org/ooxml/…` URIs — resolved
/// to zero anchors; this pins that a picture's `r:embed` blip still resolves
/// to its media part.
#[cfg(test)]
mod strict_namespace_tests {
    use super::*;
    use std::io::Write;
    use zip::write::SimpleFileOptions;

    const XDR_NS_STRICT: &str = "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
    const A_NS_STRICT: &str = "http://purl.oclc.org/ooxml/drawingml/main";
    const R_NS_STRICT: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";

    // 1×1 transparent PNG (smallest valid PNG), same fixture as `blip_svg_tests`.
    const PNG_1X1: &[u8] = &[
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44,
        0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F,
        0x15, 0xC4, 0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x62, 0x00,
        0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49,
        0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
    ];

    fn build_media_zip(png: &[u8]) -> Vec<u8> {
        let mut buf = Vec::new();
        {
            let cursor = Cursor::new(&mut buf);
            let mut zw = zip::ZipWriter::new(cursor);
            let opts = SimpleFileOptions::default();
            zw.start_file("xl/media/image1.png", opts).unwrap();
            zw.write_all(png).unwrap();
            zw.finish().unwrap();
        }
        buf
    }

    #[test]
    fn strict_drawing_anchor_resolves_blip_embed() {
        let xml = format!(
            r#"<xdr:wsDr xmlns:xdr="{xdr}" xmlns:a="{a}" xmlns:r="{r}">
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="2" name="P"/><xdr:cNvPicPr/></xdr:nvPicPr>
      <xdr:blipFill><a:blip r:embed="rIdPng"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
      <xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="300000" cy="300000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#,
            xdr = XDR_NS_STRICT,
            a = A_NS_STRICT,
            r = R_NS_STRICT,
        );

        let mut rels = HashMap::new();
        rels.insert("rIdPng".to_string(), "../media/image1.png".to_string());

        let data = build_media_zip(PNG_1X1);
        let mut archive = crate::XlsxZip::new(Cursor::new(data)).unwrap();
        let anchors = parse_drawing_anchors(&xml, &rels, "xl/drawings", &mut archive, &[]);

        assert_eq!(
            anchors.len(),
            1,
            "Strict xdr:twoCellAnchor must be found via is_xdr_ns"
        );
        assert_eq!(anchors[0].image_path, "xl/media/image1.png");
        assert_eq!(anchors[0].mime_type, "image/png");
    }
}

#[cfg(test)]
mod group_transform_contract_tests {
    use super::*;

    #[test]
    fn quarter_turned_leaf_uses_shared_non_uniform_group_mapping() {
        let xml = r#"<xdr:grpSp xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                                  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <xdr:sp>
            <xdr:spPr>
              <a:xfrm rot="5400000">
                <a:off x="0" y="50800"/>
                <a:ext cx="127000" cy="25400"/>
              </a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            </xdr:spPr>
          </xdr:sp>
        </xdr:grpSp>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let mut shapes = Vec::new();
        collect_shapes(
            &doc.root_element(),
            0.0,
            0.0,
            127_000.0,
            254_000.0,
            DrawingGroupTransform::from_group(DrawingGroupSpec {
                off_x: 0.0,
                off_y: 0.0,
                ext_x: 127_000.0,
                ext_y: 254_000.0,
                child_off_x: 0.0,
                child_off_y: 0.0,
                child_ext_x: 127_000.0,
                child_ext_y: 127_000.0,
                rotation_degrees: 0.0,
                flip_h: false,
                flip_v: false,
            }),
            &[],
            &ThemeFormatScheme::default(),
            &HashMap::new(),
            &mut shapes,
            DepthGuard::root(),
        );

        assert_eq!(shapes.len(), 1);
        let shape = &shapes[0];
        assert!((shape.x - 0.0).abs() < 1e-9);
        assert!((shape.y - 0.4).abs() < 1e-9);
        assert!((shape.w - 1.0).abs() < 1e-9);
        assert!((shape.h - 0.2).abs() < 1e-9);
        assert_eq!(shape.rot, 90.0);
    }

    #[test]
    fn xfrm_preserves_group_rotation_and_flip_for_shared_mapping() {
        let doc = roxmltree::Document::parse(
            r#"<a:xfrm xmlns:a="urn:a" rot="5400000" flipH="1">
                 <a:off x="0" y="0"/><a:ext cx="200" cy="100"/>
                 <a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/>
               </a:xfrm>"#,
        )
        .unwrap();
        let group = parse_xfrm(&doc.root_element()).unwrap();
        let transform = DrawingGroupTransform::from_group(DrawingGroupSpec {
            off_x: group.off_x,
            off_y: group.off_y,
            ext_x: group.ext_x,
            ext_y: group.ext_y,
            child_off_x: group.ch_off_x,
            child_off_y: group.ch_off_y,
            child_ext_x: group.ch_ext_x,
            child_ext_y: group.ch_ext_y,
            rotation_degrees: group.rot,
            flip_h: group.flip_h,
            flip_v: group.flip_v,
        });
        let mapped = transform.apply_rect(DrawingRect {
            x: 10.0,
            y: 20.0,
            width: 20.0,
            height: 10.0,
            rotation_degrees: 15.0,
            flip_h: false,
            flip_v: true,
        });

        assert!((mapped.x - 105.0).abs() < 1e-9);
        assert!((mapped.y - 105.0).abs() < 1e-9);
        assert!((mapped.width - 40.0).abs() < 1e-9);
        assert!((mapped.height - 10.0).abs() < 1e-9);
        assert!((mapped.rotation_degrees - 75.0).abs() < 1e-9);
        assert!(mapped.flip_h);
        assert!(mapped.flip_v);
    }
}
