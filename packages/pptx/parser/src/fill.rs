//! Fill / colour / stroke / effect / 3-D scene / custom-geometry parsing
//! (pptx-specific DrawingML). Extracted verbatim from `lib.rs`. The general
//! colour-node grammar (`parse_color_node` / `parse_color_node_tint`) lives here;
//! it uses the `PptxSchemeResolver` (from `theme`) for `<a:schemeClr>` lookups.
//! Shared XML helpers (`child`, `children_vec`, `attr`, `attr_r`, `attr_i64`,
//! `attr_f64`) stay in `lib.rs` and are imported here.

use crate::theme::{
    theme_relationship_path, PptxRawSchemeResolver, PptxSchemeResolver, PptxThemeSource,
};
use crate::types::*;
use crate::{attr, attr_f64, attr_i64, attr_r, child, parse_preflighted_pptx_xml};
use ooxml_common::blip::{mime_from_ext, parse_blip_duotone, parse_src_rect};
use ooxml_common::color::ThemeResolver;
use std::collections::HashMap;

/// Parse `<a:blip><a:alphaModFix amt="..."/></a:blip>` from a blipFill node
/// (ECMA-376 §20.1.8.6). Thin re-export of the shared
/// [`ooxml_common::blip::parse_blip_alpha`] so the three formats read the blip
/// alpha identically (previously a pptx-local copy). Returns the fraction
/// `amt/100000` when present and < 1.0; `None` otherwise.
pub(crate) use ooxml_common::blip::parse_blip_alpha;
pub(crate) use ooxml_common::fill::{parse_fill_rect, parse_tile};

// ===========================
//  Color parsing
// ===========================

/// Resolve a color node (solidFill child / run rPr child) to a hex string.
/// Handles srgbClr, sysClr, prstClr, and schemeClr (with transform support).
pub(crate) fn parse_color_node(
    node: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<String> {
    parse_color_node_tint(node, theme, ooxml_common::color::TintMode::PowerPointLinear)
}

/// Like `parse_color_node`, but lets the caller pick how `<a:tint>` is interpreted.
/// Table styles (`<a:tcStyle>` band fills) use `TintMode::WordLiteral` — the literal
/// ECMA-376 §20.1.2.3.34 definition (`val·input + (1-val)·white`, so a 20% tint is a
/// near-white wash) — which is how PowerPoint renders table band tints. The SmartArt
/// accent-recolor path keeps `PowerPointLinear` (see `apply_color_transforms`).
///
/// Thin wrapper over the shared [`ooxml_common::color::parse_color_node`]: the
/// grammar + transforms live there; [`PptxSchemeResolver`] supplies the
/// pptx-specific theme-slot lookup. Output is unchanged (uppercase hex, no `#`).
pub(crate) fn parse_color_node_tint(
    node: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    tint_mode: ooxml_common::color::TintMode,
) -> Option<String> {
    ooxml_common::color::parse_color_node(node, &PptxSchemeResolver { theme }, tint_mode)
}

// ===========================
//  Fill / Stroke parsing
// ===========================

pub(crate) fn parse_fill(
    node: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<Fill> {
    parse_fill_tint(node, theme, ooxml_common::color::TintMode::PowerPointLinear)
}

/// Parse DrawingML fill properties with the caller-selected tint semantics.
/// Presentation fills use PowerPoint's linear-light tint interpolation. A few
/// specialized callers, such as table styles, select their own tint semantics.
fn parse_fill_tint(
    node: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    tint_mode: ooxml_common::color::TintMode,
) -> Option<Fill> {
    parse_fill_with_resolver(node, &PptxSchemeResolver { theme }, tint_mode)
}

fn parse_fill_with_resolver<R: ThemeResolver + ?Sized>(
    node: roxmltree::Node<'_, '_>,
    resolver: &R,
    tint_mode: ooxml_common::color::TintMode,
) -> Option<Fill> {
    for c in node.children().filter(|n| n.is_element()) {
        match c.tag_name().name() {
            "solidFill" => {
                // If the color resolves, use it. If not (e.g. phClr with no theme slot),
                // return None so the caller can fall back to the shape style color.
                if let Some(color) = ooxml_common::color::parse_color_node(c, resolver, tint_mode) {
                    return Some(Fill::Solid { color });
                }
                // Unresolvable → don't default to black; let fallback logic handle it
            }
            "noFill" => return Some(Fill::None),
            "pattFill" => {
                // ECMA-376 §20.1.8.40 — preset pattern with fg/bg colours.
                // Shared parse (ooxml_common::fill); colors resolve with pptx's
                // PowerPointLinear tint via PptxSchemeResolver.
                let ooxml_common::fill::PatternFill { fg, bg, preset } =
                    ooxml_common::fill::parse_patt_fill(c, resolver, tint_mode);
                return Some(Fill::Pattern { fg, bg, preset });
            }
            "gradFill" => {
                // Shared parse (ooxml_common::fill). Returns None when there are
                // no resolvable stops, so we keep scanning sibling fill elements.
                if let Some(g) = ooxml_common::fill::parse_grad_fill(c, resolver, tint_mode) {
                    return Some(Fill::Gradient {
                        stops: g.stops,
                        angle: g.angle,
                        grad_type: g.grad_type,
                        scaled: g.scaled,
                        path: g.path,
                        fill_to_rect: g.fill_to_rect,
                        tile_rect: g.tile_rect,
                        flip: g.flip,
                        rot_with_shape: g.rot_with_shape,
                    });
                }
            }
            _ => {}
        }
    }
    None
}

/// Resolve a fill/background style reference through the structured shared
/// theme model. Fixed scheme colors are read from the authored scheme while
/// only `phClr` is substituted from the reference's effective mapped color.
pub(crate) fn parse_style_matrix_fill_from_source(
    style_ref: roxmltree::Node<'_, '_>,
    theme_source: &(impl PptxThemeSource + ?Sized),
) -> Option<Fill> {
    use ooxml_common::color::{StyleMatrixColorResolver, TintMode};
    use ooxml_common::theme::StyleMatrixLookup;

    let idx = attr(&style_ref, "idx")?.parse::<usize>().ok()?;
    let Some(format_scheme) = theme_source.format_scheme() else {
        return parse_style_matrix_fill(style_ref, theme_source.colors(), false);
    };
    let entry = match format_scheme.lookup_fill_ref(idx) {
        StyleMatrixLookup::NoStyle => return Some(Fill::None),
        StyleMatrixLookup::Missing => return None,
        StyleMatrixLookup::Entry(entry) => entry,
    };
    let entry_xml = entry.to_xml();
    let document = roxmltree::Document::parse(&entry_xml).ok()?;
    let theme = theme_source.colors();
    let placeholder_color = parse_color_node_tint(style_ref, theme, TintMode::PowerPointLinear);
    let raw_resolver = PptxRawSchemeResolver { theme };
    let resolver = StyleMatrixColorResolver::new(&raw_resolver, placeholder_color.as_deref());
    if let Some(blip_fill) = child(document.root_element(), "blipFill") {
        let mut resolve_blip = |relationship_id: &str| {
            theme_relationship_path(theme, relationship_id).map(str::to_owned)
        };
        if let Some(fill) =
            parse_blip_fill_with_color_resolver(blip_fill, &resolver, &mut resolve_blip)
        {
            return Some(fill);
        }
    }
    parse_fill_with_resolver(
        document.root_element(),
        &resolver,
        TintMode::PowerPointLinear,
    )
}

struct StyleMatrixSchemeResolver<'a> {
    theme: &'a HashMap<String, String>,
    placeholder_color: Option<&'a str>,
}

impl ThemeResolver for StyleMatrixSchemeResolver<'_> {
    fn resolve_scheme_color(&self, name: &str) -> Option<String> {
        if name == "phClr" {
            return self.placeholder_color.map(str::to_owned);
        }
        PptxRawSchemeResolver { theme: self.theme }.resolve_scheme_color(name)
    }
}

/// Resolve a shape `fillRef` or slide/master `bgRef` through the theme's format
/// style matrix. `phClr` inside the selected style is substituted with the
/// reference element's colour before its own transforms are applied.
///
/// ECMA-376 Part 1 §19.3.1.3: bgRef 1..999 indexes fillStyleLst, 1001+
/// indexes bgFillStyleLst (1001 = first); 0 and 1000 mean no background.
pub(crate) fn parse_style_matrix_fill(
    style_ref: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    _background: bool,
) -> Option<Fill> {
    use ooxml_common::color::TintMode::PowerPointLinear;

    let idx = attr(&style_ref, "idx")?.parse::<u32>().ok()?;
    // ECMA-376 §20.1.4.2.10 and §19.3.1.3 share one index space.
    let key = match idx {
        0 | 1000 => return Some(Fill::None),
        1..=999 => format!("+fillStyle-{idx}"),
        _ => format!("+bgFillStyle-{}", idx - 1000),
    };
    let fragment = theme.get(&key)?;

    let placeholder_color = parse_color_node_tint(style_ref, theme, PowerPointLinear);

    let wrapped = format!(
        r#"<root xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{fragment}</root>"#
    );
    let doc = parse_preflighted_pptx_xml(&wrapped).ok()?;
    let resolver = StyleMatrixSchemeResolver {
        theme,
        placeholder_color: placeholder_color.as_deref(),
    };
    if let Some(blip_fill) = child(doc.root_element(), "blipFill") {
        let mut resolve_blip = |relationship_id: &str| {
            theme_relationship_path(theme, relationship_id).map(str::to_owned)
        };
        if let Some(fill) =
            parse_blip_fill_with_color_resolver(blip_fill, &resolver, &mut resolve_blip)
        {
            return Some(fill);
        }
    }
    parse_fill_with_resolver(doc.root_element(), &resolver, PowerPointLinear)
}

/// ECMA-376 §20.1.8.14 `a:blipFill` → `Fill::Image`. The `resolve_blip`
/// closure maps the `<a:blip r:embed>` rId to the blip's embedded **zip path**
/// using the caller's rels (each inheritance level resolves against its own
/// part); the mime is derived from that path. The renderer fetches the bytes
/// lazily by path rather than from an inlined data URL.
///
/// Both fill-modes are honoured and mutually exclusive:
/// - `stretch` (§20.1.8.56): the `fillRect` (§20.1.8.30) is captured so the
///   renderer can place the (possibly overscanned) image into the box.
/// - `tile` (§20.1.8.58): the tile offset/scale/flip/align descriptor is
///   captured so the renderer can repeat the blip at its native (scaled) size.
///
/// When neither child is present the blip defaults to full-box placement
/// (stretch with no fillRect).
///
/// `theme` resolves the `<a:duotone>` (§20.1.8.23) endpoint colours through the
/// slide palette (PowerPoint linear tint), so a picture FILL recolours exactly
/// like a `<p:pic>` picture element does.
pub(crate) fn parse_blip_fill<F: FnMut(&str) -> Option<String>>(
    blip_fill: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    resolve_blip: &mut F,
) -> Option<Fill> {
    parse_blip_fill_with_color_resolver(blip_fill, &PptxSchemeResolver { theme }, resolve_blip)
}

fn parse_blip_fill_with_color_resolver<
    R: ThemeResolver + ?Sized,
    F: FnMut(&str) -> Option<String>,
>(
    blip_fill: roxmltree::Node<'_, '_>,
    color_resolver: &R,
    resolve_blip: &mut F,
) -> Option<Fill> {
    let r_id = child(blip_fill, "blip").and_then(|b| attr_r(&b, "embed"))?;
    let image_path = resolve_blip(&r_id)?;
    let mime_type = mime_from_ext(&image_path).to_owned();
    let alpha = parse_blip_alpha(blip_fill);
    // §20.1.8.23 duotone recolour, resolved through the theme with PowerPoint's
    // linear tint (same call the `<p:pic>` paths use). `None` ⇒ no effect.
    let duotone = parse_blip_duotone(
        blip_fill,
        color_resolver,
        ooxml_common::color::TintMode::PowerPointLinear,
    );
    // §20.1.8.58 tile takes precedence when present (stretch/tile are an
    // either-or choice in CT_BlipFillProperties).
    if let Some(tile_node) = child(blip_fill, "tile") {
        return Some(Fill::Image {
            image_path,
            mime_type,
            src_rect: parse_src_rect(blip_fill),
            fill_rect: None,
            tile: Some(parse_tile(tile_node)),
            alpha,
            duotone,
        });
    }
    let fill_rect = child(blip_fill, "stretch").and_then(parse_fill_rect);
    Some(Fill::Image {
        image_path,
        mime_type,
        src_rect: parse_src_rect(blip_fill),
        fill_rect,
        tile: None,
        alpha,
        duotone,
    })
}

fn canvas_line_cap(cap: &str) -> Option<String> {
    match cap {
        "rnd" => Some("round".to_owned()),
        "sq" => Some("square".to_owned()),
        "flat" => Some("butt".to_owned()),
        _ => None,
    }
}

pub(crate) fn line_properties_to_stroke(
    line: &ooxml_common::line::LineProperties,
    fallback_color: Option<String>,
) -> Option<Stroke> {
    use ooxml_common::line::{LineDash, LineEnd, LineJoin, LinePaint};

    let (color, fill) = match line.paint.as_ref()? {
        LinePaint::NoFill => return None,
        LinePaint::Solid { color } => (color.clone().or(fallback_color)?, None),
        LinePaint::Gradient(Some(gradient)) => {
            let color = gradient
                .stops
                .iter()
                .rev()
                .find(|stop| !stop.color.ends_with("00"))
                .or_else(|| gradient.stops.last())?
                .color
                .clone();
            (
                color,
                Some(Fill::Gradient {
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
        LinePaint::Gradient(None) => return None,
        LinePaint::Pattern(pattern) => (
            pattern.fg.clone(),
            Some(Fill::Pattern {
                fg: pattern.fg.clone(),
                bg: pattern.bg.clone(),
                preset: pattern.preset.clone(),
            }),
        ),
    };
    let dash_style = match line.dash.as_ref() {
        Some(LineDash::Preset(Some(value))) if value != "solid" => Some(value.clone()),
        _ => None,
    };
    let custom_dash = match line.dash.as_ref() {
        Some(LineDash::Custom(stops)) => stops
            .iter()
            .map(|stop| StrokeDashSegment {
                dash: stop.dash / 100_000.0,
                space: stop.space / 100_000.0,
            })
            .collect(),
        _ => Vec::new(),
    };
    let (line_join, miter_limit) = match line.join.as_ref() {
        Some(LineJoin::Round) => (Some("round".to_owned()), None),
        Some(LineJoin::Bevel) => (Some("bevel".to_owned()), None),
        Some(LineJoin::Miter { limit }) => (
            Some("miter".to_owned()),
            limit.map(|value| value as f64 / 100_000.0),
        ),
        None => (None, None),
    };
    let arrow = |end: &LineEnd| {
        let kind = end.kind.clone().unwrap_or_else(|| "none".to_owned());
        (kind != "none").then(|| ArrowEnd {
            kind,
            w: end.width.clone().unwrap_or_else(|| "med".to_owned()),
            len: end.length.clone().unwrap_or_else(|| "med".to_owned()),
        })
    };
    Some(Stroke {
        color,
        width: line.width.unwrap_or(9525),
        fill,
        dash_style,
        custom_dash,
        line_cap: line.cap.as_deref().and_then(canvas_line_cap),
        line_join,
        miter_limit,
        alignment: line
            .alignment
            .clone()
            .filter(|value| matches!(value.as_str(), "ctr" | "in")),
        head_end: line.head_end.as_ref().and_then(arrow),
        tail_end: line.tail_end.as_ref().and_then(arrow),
        cmpd: line.compound.clone().filter(|value| value != "sng"),
    })
}

pub(crate) fn parse_stroke(
    ln_node: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<Stroke> {
    let properties = ooxml_common::line::parse_line_properties(
        ln_node,
        &PptxSchemeResolver { theme },
        ooxml_common::color::TintMode::PowerPointLinear,
    );
    line_properties_to_stroke(&properties, None)
}

// ===========================
//  Shadow parsing
// ===========================

/// Parse spPr > effectLst > outerShdw into a Shadow.
pub(crate) fn parse_shadow(
    effect_lst: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<Shadow> {
    parse_shadow_node(child(effect_lst, "outerShdw")?, theme)
}

/// Parse spPr > effectLst > innerShdw into a Shadow. ECMA-376 §20.1.8.21
/// (CT_InnerShadowEffect) — same field shape as outerShdw, semantics differ
/// at render time (cast inward).
#[cfg(test)]
pub(crate) fn parse_inner_shadow(
    effect_lst: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<Shadow> {
    parse_shadow_node(child(effect_lst, "innerShdw")?, theme)
}

/// Shared field reader for innerShdw / outerShdw. Both elements expose
/// blurRad, dist, dir, and a color child with optional alphaModFix.
pub(crate) fn parse_shadow_node(
    n: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<Shadow> {
    parse_shadow_node_with_resolver(
        n,
        &PptxSchemeResolver { theme },
        ooxml_common::color::TintMode::PowerPointLinear,
    )
}

fn parse_shadow_node_with_resolver<R: ThemeResolver + ?Sized>(
    n: roxmltree::Node<'_, '_>,
    resolver: &R,
    tint_mode: ooxml_common::color::TintMode,
) -> Option<Shadow> {
    let blur = attr_i64(&n, "blurRad").unwrap_or(0);
    let dist = attr_i64(&n, "dist").unwrap_or(0);
    let dir = attr_f64(&n, "dir").unwrap_or(0.0) / 60_000.0;
    // CT_OuterShadowEffect (§20.1.8.45). These attributes do not exist on
    // CT_InnerShadowEffect, so the shared reader keeps them optional.
    let sx = attr_f64(&n, "sx").map(|value| value / 100_000.0);
    let sy = attr_f64(&n, "sy").map(|value| value / 100_000.0);
    let kx = attr_f64(&n, "kx").map(|value| value / 60_000.0);
    let ky = attr_f64(&n, "ky").map(|value| value / 60_000.0);
    let algn = attr(&n, "algn");
    let rot_with_shape =
        attr(&n, "rotWithShape").map(|value| value == "1" || value.eq_ignore_ascii_case("true"));

    let color_str = ooxml_common::color::parse_color_node(n, resolver, tint_mode)
        .unwrap_or_else(|| "000000".to_owned());
    let (color, alpha) = if color_str.len() >= 8 {
        let a = u8::from_str_radix(&color_str[6..8], 16).unwrap_or(255) as f64 / 255.0;
        (color_str[..6].to_owned(), a)
    } else {
        (color_str, 1.0)
    };

    Some(Shadow {
        color,
        alpha,
        blur,
        dist,
        dir,
        sx,
        sy,
        kx,
        ky,
        algn,
        rot_with_shape,
    })
}

/// Parse spPr > effectLst > glow into a Glow effect — ECMA-376 §20.1.8.17
/// (CT_GlowEffect): a coloured halo with a blur radius, no offset.
#[cfg(test)]
pub(crate) fn parse_glow(
    effect_lst: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<Glow> {
    let g = child(effect_lst, "glow")?;
    parse_glow_node_with_resolver(
        g,
        &PptxSchemeResolver { theme },
        ooxml_common::color::TintMode::PowerPointLinear,
    )
}

fn parse_glow_node_with_resolver<R: ThemeResolver + ?Sized>(
    g: roxmltree::Node<'_, '_>,
    resolver: &R,
    tint_mode: ooxml_common::color::TintMode,
) -> Option<Glow> {
    let radius = attr_i64(&g, "rad").unwrap_or(0);
    let color_str = ooxml_common::color::parse_color_node(g, resolver, tint_mode)
        .unwrap_or_else(|| "000000".to_owned());
    let (color, alpha) = if color_str.len() >= 8 {
        let a = u8::from_str_radix(&color_str[6..8], 16).unwrap_or(255) as f64 / 255.0;
        (color_str[..6].to_owned(), a)
    } else {
        (color_str, 1.0)
    };
    Some(Glow {
        color,
        alpha,
        radius,
    })
}

/// Parse spPr > effectLst > softEdge into a SoftEdge — ECMA-376 §20.1.8.31.
pub(crate) fn parse_soft_edge(effect_lst: roxmltree::Node<'_, '_>) -> Option<SoftEdge> {
    let n = child(effect_lst, "softEdge")?;
    let radius = attr_i64(&n, "rad").unwrap_or(0);
    Some(SoftEdge { radius })
}

/// Parse spPr > effectLst > reflection — ECMA-376 §20.1.8.27. Defaults
/// follow the spec table: blur=0, dist=0, dir=0, stA=100000 (=1.0),
/// stPos=0, endA=0, endPos=100000 (=1.0), sx=100000, sy=-100000.
pub(crate) fn parse_reflection(effect_lst: roxmltree::Node<'_, '_>) -> Option<Reflection> {
    let r = child(effect_lst, "reflection")?;
    let pct = |name: &str, default: f64| -> f64 {
        attr_f64(&r, name).map(|v| v / 100_000.0).unwrap_or(default)
    };
    Some(Reflection {
        blur: attr_i64(&r, "blurRad").unwrap_or(0),
        dist: attr_i64(&r, "dist").unwrap_or(0),
        dir: attr_f64(&r, "dir").unwrap_or(0.0) / 60_000.0,
        st_a: pct("stA", 1.0),
        st_pos: pct("stPos", 0.0),
        end_a: pct("endA", 0.0),
        end_pos: pct("endPos", 1.0),
        sx: pct("sx", 1.0),
        sy: pct("sy", -1.0),
    })
}

/// Effects pulled from `spPr > effectLst`. The five members are independent
/// siblings inside `CT_EffectList` — ECMA-376 §20.1.8.16. Used by both shapes
/// (`p:sp`) and pictures (`p:pic`): `p:spPr` is `CT_ShapeProperties` in both
/// cases (§19.3.1.37), so `effectLst` applies equally to images.
#[derive(Default)]
pub(crate) struct EffectLst {
    pub(crate) shadow: Option<Shadow>,
    pub(crate) inner_shadow: Option<Shadow>,
    pub(crate) glow: Option<Glow>,
    pub(crate) soft_edge: Option<SoftEdge>,
    pub(crate) reflection: Option<Reflection>,
}

/// One entry in theme `effectStyleLst` (ECMA-376 §20.1.4.1.11).
/// `scene3d` and `sp3d` are peers of the effect property choice and must not be
/// discarded when a shape resolves `effectRef`.
#[derive(Default)]
pub(crate) struct StyleMatrixEffects {
    pub(crate) effects: EffectLst,
    pub(crate) scene3d: Option<Scene3d>,
    pub(crate) sp3d: Option<Sp3d>,
}

/// Read every `effectLst` child shapes and pictures share. `effect_lst` is the
/// optional `<a:effectLst>` node; missing nodes yield an all-`None` result.
pub(crate) fn parse_effect_lst(
    effect_lst: Option<roxmltree::Node<'_, '_>>,
    theme: &HashMap<String, String>,
) -> EffectLst {
    parse_effect_lst_with_resolver(
        effect_lst,
        &PptxSchemeResolver { theme },
        ooxml_common::color::TintMode::PowerPointLinear,
    )
}

fn parse_effect_lst_with_resolver<R: ThemeResolver + ?Sized>(
    effect_lst: Option<roxmltree::Node<'_, '_>>,
    resolver: &R,
    tint_mode: ooxml_common::color::TintMode,
) -> EffectLst {
    EffectLst {
        shadow: effect_lst
            .and_then(|node| child(node, "outerShdw"))
            .and_then(|node| parse_shadow_node_with_resolver(node, resolver, tint_mode)),
        inner_shadow: effect_lst
            .and_then(|node| child(node, "innerShdw"))
            .and_then(|node| parse_shadow_node_with_resolver(node, resolver, tint_mode)),
        glow: effect_lst
            .and_then(|node| child(node, "glow"))
            .and_then(|node| parse_glow_node_with_resolver(node, resolver, tint_mode)),
        soft_edge: effect_lst.and_then(parse_soft_edge),
        reflection: effect_lst.and_then(parse_reflection),
    }
}

/// Resolve `p:style/a:effectRef` through the theme format matrix.
///
/// `effectRef@idx` is one-based into `a:effectStyleLst`. Any `phClr` inside
/// that effect style is supplied by the color child of the reference before
/// the ordinary DrawingML transforms are applied.
pub(crate) fn parse_style_matrix_effects(
    effect_ref: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> StyleMatrixEffects {
    let Some(idx) = attr(&effect_ref, "idx").and_then(|value| value.parse::<u32>().ok()) else {
        return StyleMatrixEffects::default();
    };
    if idx == 0 {
        return StyleMatrixEffects::default();
    }
    let Some(fragment) = theme.get(&format!("+effectStyle-{idx}")) else {
        return StyleMatrixEffects::default();
    };

    let placeholder_color = parse_color_node_tint(
        effect_ref,
        theme,
        ooxml_common::color::TintMode::PowerPointLinear,
    );
    let wrapped = format!(
        r#"<root xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">{fragment}</root>"#
    );
    let Ok(doc) = parse_preflighted_pptx_xml(&wrapped) else {
        return StyleMatrixEffects::default();
    };
    let effect_style = doc
        .root_element()
        .descendants()
        .find(|node| node.is_element() && node.tag_name().name() == "effectStyle");
    let effect_properties =
        effect_style.and_then(|node| child(node, "effectLst").or_else(|| child(node, "effectDag")));
    let resolver = StyleMatrixSchemeResolver {
        theme,
        placeholder_color: placeholder_color.as_deref(),
    };
    StyleMatrixEffects {
        effects: parse_effect_lst_with_resolver(
            effect_properties,
            &resolver,
            ooxml_common::color::TintMode::PowerPointLinear,
        ),
        scene3d: effect_style.and_then(parse_scene3d),
        sp3d: effect_style.and_then(|node| {
            parse_sp3d_with_resolver(
                node,
                &resolver,
                ooxml_common::color::TintMode::PowerPointLinear,
            )
        }),
    }
}

// ===========================
//  3D scene parsing (scene3d / sp3d)
// ===========================

/// Parse `<a:rot>` (`CT_SphereCoords`, ECMA-376 §20.1.5.11). Angles are stored
/// in the XML as 60000ths of a degree; we convert to degrees. All three
/// attributes are required by the schema, but we default missing ones to 0 to
/// stay tolerant of malformed input.
pub(crate) fn parse_rot3d(rot: roxmltree::Node<'_, '_>) -> Rot3d {
    let deg = |name: &str| attr_f64(&rot, name).unwrap_or(0.0) / 60_000.0;
    Rot3d {
        lat: deg("lat"),
        lon: deg("lon"),
        rev: deg("rev"),
    }
}

/// Parse `<a:scene3d>` (`CT_Scene3D`, ECMA-376 §20.1.4.1.41). Requires a
/// `<a:camera>` child (§20.1.5.5); `<a:lightRig>` is optional for our purposes
/// (Phase A renders the camera only). Returns None when no camera is present.
pub(crate) fn parse_scene3d(sppr: roxmltree::Node<'_, '_>) -> Option<Scene3d> {
    let scene = child(sppr, "scene3d")?;
    let cam = child(scene, "camera")?;
    let camera = Camera3d {
        prst: attr(&cam, "prst")?,
        // §20.1.5.5: fov is an ST_FOVAngle in 60000ths of a degree.
        fov: attr_f64(&cam, "fov").map(|v| v / 60_000.0),
        // zoom is an ST_PositivePercentage (100000 = 100%).
        zoom: attr_f64(&cam, "zoom").map(|v| v / 100_000.0),
        rot: child(cam, "rot").map(parse_rot3d),
    };
    let light_rig = child(scene, "lightRig").and_then(|lr| {
        Some(LightRig {
            rig: attr(&lr, "rig")?,
            dir: attr(&lr, "dir")?,
            rot: child(lr, "rot").map(parse_rot3d),
        })
    });
    Some(Scene3d { camera, light_rig })
}

/// Parse `<a:bevel>` (`CT_Bevel`, ECMA-376 §20.1.5.3). `w`/`h` default to
/// 76200 EMU and `prst` to "circle" per the schema.
pub(crate) fn parse_bevel3d(bevel: roxmltree::Node<'_, '_>) -> Bevel3d {
    Bevel3d {
        w: attr_i64(&bevel, "w").unwrap_or(76_200),
        h: attr_i64(&bevel, "h").unwrap_or(76_200),
        prst: attr(&bevel, "prst").unwrap_or_else(|| "circle".into()),
    }
}

/// Parse `<a:sp3d>` (`CT_Shape3D`, ECMA-376 §20.1.5.12). Defaults follow the
/// schema: z=0, extrusionH=0, contourW=0, prstMaterial="warmMatte". Parsed in
/// full but not rendered in Phase A.
pub(crate) fn parse_sp3d(
    sppr: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<Sp3d> {
    parse_sp3d_with_resolver(
        sppr,
        &PptxSchemeResolver { theme },
        ooxml_common::color::TintMode::PowerPointLinear,
    )
}

fn parse_sp3d_with_resolver<R: ThemeResolver + ?Sized>(
    sppr: roxmltree::Node<'_, '_>,
    resolver: &R,
    tint_mode: ooxml_common::color::TintMode,
) -> Option<Sp3d> {
    let n = child(sppr, "sp3d")?;
    let contour_clr = child(n, "contourClr")
        .and_then(|c| ooxml_common::color::parse_color_node(c, resolver, tint_mode));
    let extrusion_clr = child(n, "extrusionClr")
        .and_then(|c| ooxml_common::color::parse_color_node(c, resolver, tint_mode));
    Some(Sp3d {
        z: attr_i64(&n, "z").unwrap_or(0),
        extrusion_h: attr_i64(&n, "extrusionH").unwrap_or(0),
        contour_w: attr_i64(&n, "contourW").unwrap_or(0),
        contour_clr,
        extrusion_clr,
        prst_material: attr(&n, "prstMaterial").unwrap_or_else(|| "warmMatte".into()),
        bevel_t: child(n, "bevelT").map(parse_bevel3d),
        bevel_b: child(n, "bevelB").map(parse_bevel3d),
    })
}

// ===========================
//  Custom geometry parsing
// ===========================

/// Parse custGeom > pathLst into a list of sub-paths (one per <a:path> element).
pub(crate) fn parse_cust_geom(
    cust_geom: roxmltree::Node<'_, '_>,
    shape_w: f64,
    shape_h: f64,
) -> Vec<Vec<PathCmd>> {
    use ooxml_common::custom_geometry::{parse_custom_geometry, PathCommand};

    parse_custom_geometry(cust_geom, shape_w, shape_h)
        .paths
        .into_iter()
        .map(|path| {
            path.commands
                .into_iter()
                .map(|command| match command {
                    PathCommand::MoveTo { x, y } => PathCmd::MoveTo {
                        x: x / path.width,
                        y: y / path.height,
                    },
                    PathCommand::LineTo { x, y } => PathCmd::LineTo {
                        x: x / path.width,
                        y: y / path.height,
                    },
                    PathCommand::CubicBezierTo {
                        x1,
                        y1,
                        x2,
                        y2,
                        x,
                        y,
                    } => PathCmd::CubicBezTo {
                        x1: x1 / path.width,
                        y1: y1 / path.height,
                        x2: x2 / path.width,
                        y2: y2 / path.height,
                        x: x / path.width,
                        y: y / path.height,
                    },
                    PathCommand::QuadraticBezierTo { x1, y1, x, y } => PathCmd::QuadBezTo {
                        x1: x1 / path.width,
                        y1: y1 / path.height,
                        x: x / path.width,
                        y: y / path.height,
                    },
                    PathCommand::ArcTo {
                        wr,
                        hr,
                        st_ang,
                        sw_ang,
                    } => PathCmd::ArcTo {
                        wr: wr / path.width,
                        hr: hr / path.height,
                        st_ang: st_ang / 60000.0,
                        sw_ang: sw_ang / 60000.0,
                    },
                    PathCommand::Close => PathCmd::Close,
                })
                .collect()
        })
        .collect()
}

// ===========================
//  Transform (a:xfrm)
// ===========================

pub(crate) fn parse_xfrm(xfrm: roxmltree::Node<'_, '_>) -> Transform {
    let rot = attr_f64(&xfrm, "rot").unwrap_or(0.0) / 60000.0;
    let flip_h = attr(&xfrm, "flipH")
        .map(|v| v == "1" || v == "true")
        .unwrap_or(false);
    let flip_v = attr(&xfrm, "flipV")
        .map(|v| v == "1" || v == "true")
        .unwrap_or(false);
    let off = child(xfrm, "off");
    let ext = child(xfrm, "ext");
    Transform {
        x: off.and_then(|n| attr_i64(&n, "x")).unwrap_or(0),
        y: off.and_then(|n| attr_i64(&n, "y")).unwrap_or(0),
        cx: ext.and_then(|n| attr_i64(&n, "cx")).unwrap_or(0),
        cy: ext.and_then(|n| attr_i64(&n, "cy")).unwrap_or(0),
        rot,
        flip_h,
        flip_v,
    }
}

// ===========================
//  Slide background
// ===========================

/// ECMA-376 §19.3.1.1 `p:bg`. `resolve_blip` maps a `<a:blip r:embed>` rId to a
/// base64 data URL using the rels + zip of the part this `c_sld` belongs to
/// (slide / layout / master), so an image background (§20.1.8.14) is resolved
/// against the correct relationship base.
pub(crate) fn parse_background<F: FnMut(&str) -> Option<String>>(
    c_sld: roxmltree::Node<'_, '_>,
    theme_source: &(impl PptxThemeSource + ?Sized),
    resolve_blip: &mut F,
) -> Option<Fill> {
    let theme = theme_source.colors();
    let bg = child(c_sld, "bg")?;
    // bgPr contains an explicit fill specification
    if let Some(bg_pr) = child(bg, "bgPr") {
        // §20.1.8.14 — an image background lives in `bgPr > blipFill`. Try it
        // first so the embedded blip is resolved; fall back to the generic
        // solid/gradient/pattern parser for non-image bgPr fills.
        if let Some(blip_fill) = child(bg_pr, "blipFill") {
            if let Some(fill) = parse_blip_fill(blip_fill, theme, resolve_blip) {
                return Some(fill);
            }
        }
        return parse_fill_tint(
            bg_pr,
            theme,
            ooxml_common::color::TintMode::PowerPointLinear,
        );
    }
    // bgRef references a theme background style; its child is a color element
    if let Some(bg_ref) = child(bg, "bgRef") {
        return parse_style_matrix_fill_from_source(bg_ref, theme_source)
            .or_else(|| parse_color_node(bg_ref, theme).map(|c| Fill::Solid { color: c }));
    }
    None
}

/// Resolve a table-style `<a:fill>` wrapper's colour. PowerPoint applies the
/// ECMA-376 §20.1.2.3.34 retained-input tint in linear sRGB for these DrawingML
/// fills, just as it does for other presentation fills. Gradient/pattern/blip
/// fills (rare in table styles) defer to the generic `parse_fill`.
pub(crate) fn parse_table_style_fill(
    fill_wrapper: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<Fill> {
    use ooxml_common::color::TintMode::PowerPointLinear;
    for c in fill_wrapper.children().filter(|n| n.is_element()) {
        match c.tag_name().name() {
            "noFill" => return Some(Fill::None),
            "solidFill" => {
                return parse_color_node_tint(c, theme, PowerPointLinear)
                    .map(|color| Fill::Solid { color });
            }
            _ => {}
        }
    }
    parse_fill(fill_wrapper, theme)
}
