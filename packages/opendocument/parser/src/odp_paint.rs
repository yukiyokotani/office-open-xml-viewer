//! Area fill, line and effect properties of drawing objects (ODF 1.3
//! Part 3 §20: `style:graphic-properties` and `style:drawing-page-properties`)
//! mapped onto the PPTX paint model.

use crate::odp_media::Media;
use crate::odp_style::{Nd, Props, Sheet};
use crate::odp_units::{
    angle_degrees, color, hex, intensity, length_emu, percent, rgb, with_alpha,
};
use crate::xml::{NS_DRAW, NS_FO, NS_LOEXT, NS_STYLE, NS_SVG, NS_XLINK};
use ooxml_common::effect::{Glow, Shadow, SoftEdge};
use ooxml_common::fill::{FillRect, GradStop, TileInfo};
use pptx_model::{Fill, Stroke, StrokeDashSegment};

pub struct PaintContext<'a, 'm, 'p> {
    pub sheet: &'a Sheet<'a>,
    pub media: &'m mut Media<'p>,
    /// The package file holding the styles (relative image references).
    pub base: &'a str,
    /// The slide background, for `loext:fill-use-slide-background`.
    pub slide_background: Option<&'a Fill>,
}

/// Opacity of the area (`draw:opacity`, §20.92), 1 when absent.
fn uniform_opacity(props: &Props<'_>) -> f64 {
    props
        .get(NS_DRAW, "opacity")
        .and_then(percent)
        .unwrap_or(1.0)
        .clamp(0.0, 1.0)
}

/// Resolve the area fill (§20.95 `draw:fill`) of a property cascade.
/// `None` means no fill is painted.
pub fn fill(
    props: &Props<'_>,
    ctx: &mut PaintContext<'_, '_, '_>,
    area: (f64, f64),
) -> Option<Fill> {
    if props.get(NS_LOEXT, "fill-use-slide-background") == Some("true") {
        return ctx.slide_background.cloned();
    }
    let kind = props.get(NS_DRAW, "fill").unwrap_or("none");
    let opacity = uniform_opacity(props);
    let opacity_gradient = props
        .get(NS_DRAW, "opacity-name")
        .and_then(|name| ctx.sheet.named("opacity", name));
    let base = match kind {
        "solid" => {
            let c = props
                .get(NS_DRAW, "fill-color")
                .and_then(color)
                .unwrap_or_else(|| "FFFFFF".to_string());
            match opacity_gradient {
                Some(og) => {
                    return Some(opacity_gradient_fill(&c, og));
                }
                None => Fill::Solid {
                    color: with_alpha(&c, opacity),
                },
            }
        }
        "gradient" => {
            let g = props
                .get(NS_DRAW, "fill-gradient-name")
                .and_then(|name| ctx.sheet.named("gradient", name))?;
            match opacity_gradient.and_then(|og| combined_gradient(g, og)) {
                Some(fill) => fill,
                None => {
                    let mut fill = gradient_fill(g, props)?;
                    apply_opacity(&mut fill, opacity, opacity_gradient);
                    fill
                }
            }
        }
        "hatch" => {
            let h = props
                .get(NS_DRAW, "fill-hatch-name")
                .and_then(|name| ctx.sheet.named("hatch", name))?;
            let bg = if props.get(NS_DRAW, "fill-hatch-solid") == Some("true") {
                props
                    .get(NS_DRAW, "fill-color")
                    .and_then(color)
                    .map(|c| with_alpha(&c, opacity))
                    .unwrap_or_else(|| "FFFFFF".to_string())
            } else {
                "FFFFFF00".to_string()
            };
            // The SVG is a function of the hatch's attributes, the area colour,
            // the opacity and the area size: the same ones reuse one picture.
            let attr = |name: &str| h.attribute((NS_DRAW, name)).unwrap_or("");
            let description = format!(
                "hatch|{}|{}|{}|{}|{bg}|{opacity:.3}|{:.0}|{:.0}",
                attr("color"),
                attr("distance"),
                attr("rotation"),
                attr("style"),
                area.0,
                area.1
            );
            let generated = ctx.media.generated("hatch", description, area, || {
                hatch_svg(h, &bg, opacity, area)
            });
            match generated {
                Some((path, resource)) => Fill::Image {
                    image_path: path.clone(),
                    mime_type: resource.mime.to_string(),
                    svg_image_path: Some(path),
                    dpi: None,
                    rot_with_shape: Some(true),
                    src_rect: None,
                    fill_rect: None,
                    stretch: true,
                    tile: None,
                    alpha: None,
                    duotone: None,
                    blip_effects: Vec::new(),
                },
                None => hatch_fill(h, bg, opacity),
            }
        }
        "bitmap" => {
            let image = props
                .get(NS_DRAW, "fill-image-name")
                .and_then(|name| ctx.sheet.named("fill-image", name))?;
            let href = image.attribute((NS_XLINK, "href"))?;
            let (path, resource) = ctx.media.admit(ctx.base, href)?;
            image_fill(path, &resource, props, opacity, area)
        }
        _ => return None,
    };
    Some(base)
}

/// Linear interpolation of colours a→b at t.
fn mix(a: &str, b: &str, t: f64) -> String {
    let (ar, ag, ab) = rgb(a);
    let (br, bg, bb) = rgb(b);
    hex(ar + (br - ar) * t, ag + (bg - ag) * t, ab + (bb - ab) * t)
}

/// One gradient description shared by colour and opacity gradients
/// (§19.224–§19.228 / §16.40 `draw:opacity`).
struct Ramp {
    style: String,
    angle: f64,
    border: f64,
    cx: f64,
    cy: f64,
    /// (offset from start colour 0 → end colour 1, colour or opacity).
    stops: Vec<(f64, f64, String)>,
}

/// LibreOffice writes `draw:angle` with a unit; a unitless value is read in
/// tenths of a degree, as LibreOffice does for its earlier files.
fn gradient_angle(node: Nd<'_>) -> f64 {
    match node.attribute((NS_DRAW, "angle")) {
        Some(v)
            if v.trim()
                .chars()
                .all(|c| c.is_ascii_digit() || c == '-' || c == '.') =>
        {
            v.trim().parse::<f64>().map(|t| t / 10.0).unwrap_or(0.0)
        }
        Some(v) => angle_degrees(v).unwrap_or(0.0),
        None => 0.0,
    }
}

fn ramp(node: Nd<'_>) -> Ramp {
    let pct = |name: &str, default: f64| {
        node.attribute((NS_DRAW, name))
            .and_then(percent)
            .unwrap_or(default)
    };
    Ramp {
        style: node
            .attribute((NS_DRAW, "style"))
            .unwrap_or("linear")
            .to_string(),
        angle: gradient_angle(node),
        border: pct("border", 0.0).clamp(0.0, 1.0),
        cx: pct("cx", 0.5),
        cy: pct("cy", 0.5),
        stops: Vec::new(),
    }
}

/// Implementation bound on the stops of one colour or opacity gradient
/// (later stops are not read), as in the shared fill reader.
const MAX_GRADIENT_STOPS: usize = 256;

fn color_ramp(g: Nd<'_>) -> Option<Ramp> {
    let mut r = ramp(g);
    let stops: Vec<(f64, f64, String)> = g
        .children()
        .filter(|c| crate::xml::is(*c, NS_LOEXT, "gradient-stop"))
        .take(MAX_GRADIENT_STOPS)
        .filter_map(|c| {
            let offset: f64 = c.attribute((NS_SVG, "offset"))?.trim().parse().ok()?;
            let value = c.attribute((NS_LOEXT, "color-value")).and_then(color)?;
            Some((offset.clamp(0.0, 1.0), 1.0, value))
        })
        .collect();
    let start_intensity = g
        .attribute((NS_DRAW, "start-intensity"))
        .and_then(percent)
        .unwrap_or(1.0);
    let end_intensity = g
        .attribute((NS_DRAW, "end-intensity"))
        .and_then(percent)
        .unwrap_or(1.0);
    if stops.len() >= 2 {
        // Multi-colour gradients (LibreOffice 7.6 `loext:gradient-stop`);
        // the intensities scale along the axis as for two colours.
        r.stops = stops
            .into_iter()
            .map(|(o, a, c)| {
                let f = start_intensity + (end_intensity - start_intensity) * o;
                (o, a, intensity(&c, f))
            })
            .collect();
    } else {
        let start = g
            .attribute((NS_DRAW, "start-color"))
            .and_then(color)
            .unwrap_or_else(|| "000000".into());
        let end = g
            .attribute((NS_DRAW, "end-color"))
            .and_then(color)
            .unwrap_or_else(|| "FFFFFF".into());
        r.stops = vec![
            (0.0, 1.0, intensity(&start, start_intensity)),
            (1.0, 1.0, intensity(&end, end_intensity)),
        ];
    }
    Some(r)
}

fn opacity_ramp(o: Nd<'_>) -> Ramp {
    let mut r = ramp(o);
    let stops: Vec<(f64, f64, String)> = o
        .children()
        .filter(|c| crate::xml::is(*c, NS_LOEXT, "opacity-stop"))
        .take(MAX_GRADIENT_STOPS)
        .filter_map(|c| {
            let offset: f64 = c.attribute((NS_SVG, "offset"))?.trim().parse().ok()?;
            let value: f64 = c.attribute((NS_SVG, "stop-opacity"))?.trim().parse().ok()?;
            Some((offset.clamp(0.0, 1.0), value.clamp(0.0, 1.0), String::new()))
        })
        .collect();
    if stops.len() >= 2 {
        r.stops = stops;
    } else {
        let start = o
            .attribute((NS_DRAW, "start"))
            .and_then(percent)
            .unwrap_or(0.0);
        let end = o
            .attribute((NS_DRAW, "end"))
            .and_then(percent)
            .unwrap_or(1.0);
        r.stops = vec![(0.0, start, String::new()), (1.0, end, String::new())];
    }
    r
}

/// Sample a ramp at `t`, searching from window `*from` (advanced to the
/// window found, for ascending `t`). The stops are in file order, as the
/// unsorted search before them found them.
fn sample_from(stops: &[(f64, f64, String)], t: f64, from: &mut usize) -> (f64, String) {
    let first = &stops[0];
    if t <= first.0 {
        return (first.1, first.2.clone());
    }
    for (k, pair) in stops.windows(2).enumerate().skip(*from) {
        let (a, b) = (&pair[0], &pair[1]);
        *from = k;
        if t <= b.0 {
            let span = b.0 - a.0;
            let u = if span > 0.0 { (t - a.0) / span } else { 1.0 };
            let colour = if a.2.is_empty() {
                String::new()
            } else {
                mix(&a.2, &b.2, u)
            };
            return (a.1 + (b.1 - a.1) * u, colour);
        }
    }
    let last = stops.last().unwrap();
    (last.1, last.2.clone())
}

/// Express a ramp in the DrawingML gradient model. Returns the fill with
/// stops positioned along the DrawingML axis (0 = start of the linear axis,
/// or the centre of a path gradient), using `paint(offset)` for each stop.
/// Stops are (offset from the start colour, colour); equal offsets keep
/// hard colour steps.
fn ramp_fill(r: &Ramp, ramp_stops: &[(f64, String)]) -> Fill {
    let first = ramp_stops.first().map(|s| s.1.clone()).unwrap_or_default();
    let b = r.border;
    let mut stops: Vec<GradStop> = Vec::new();
    let push = |stops: &mut Vec<GradStop>, position: f64, color: String| {
        stops.push(GradStop {
            position: position.clamp(0.0, 1.0),
            color,
        });
    };
    // ODF: 0° linear runs from the top (start colour) to the bottom, and the
    // angle turns the axis counter-clockwise. DrawingML 90° runs top→bottom
    // and turns clockwise.
    let ooxml_angle = (90.0 - r.angle).rem_euclid(360.0);
    match r.style.as_str() {
        "axial" => {
            // Start colour at both edges, end colour on the centre line; the
            // border holds the start colour at each edge.
            for (o, c) in ramp_stops {
                push(&mut stops, (b + (1.0 - b) * o) * 0.5, c.clone());
            }
            for (o, c) in ramp_stops.iter().rev() {
                push(&mut stops, 1.0 - (b + (1.0 - b) * o) * 0.5, c.clone());
            }
            if b > 0.0 {
                stops.insert(
                    0,
                    GradStop {
                        position: 0.0,
                        color: first.clone(),
                    },
                );
                stops.push(GradStop {
                    position: 1.0,
                    color: first.clone(),
                });
            }
            Fill::Gradient {
                stops,
                angle: ooxml_angle,
                grad_type: "linear".into(),
                scaled: Some(false),
                path: None,
                fill_to_rect: None,
                tile_rect: None,
                flip: None,
                rot_with_shape: Some(true),
            }
        }
        "radial" | "ellipsoid" | "square" | "rectangular" => {
            // Start colour outside, end colour at the centre (cx, cy); the
            // border is the outer band of start colour.
            for (o, c) in ramp_stops.iter().rev() {
                push(&mut stops, (1.0 - b) * (1.0 - o), c.clone());
            }
            if b > 0.0 {
                stops.push(GradStop {
                    position: 1.0,
                    color: first.clone(),
                });
            }
            let path = if matches!(r.style.as_str(), "square" | "rectangular") {
                "rect"
            } else {
                "circle"
            };
            Fill::Gradient {
                stops,
                angle: 0.0,
                grad_type: "radial".into(),
                scaled: None,
                path: Some(path.into()),
                fill_to_rect: Some(FillRect {
                    l: r.cx,
                    t: r.cy,
                    r: 1.0 - r.cx,
                    b: 1.0 - r.cy,
                }),
                tile_rect: None,
                flip: None,
                rot_with_shape: Some(true),
            }
        }
        _ => {
            if b > 0.0 {
                push(&mut stops, 0.0, first.clone());
            }
            for (o, c) in ramp_stops {
                push(&mut stops, b + (1.0 - b) * o, c.clone());
            }
            Fill::Gradient {
                stops,
                angle: ooxml_angle,
                grad_type: "linear".into(),
                scaled: Some(false),
                path: None,
                fill_to_rect: None,
                tile_rect: None,
                flip: None,
                rot_with_shape: Some(true),
            }
        }
    }
}

fn gradient_fill(g: Nd<'_>, _props: &Props<'_>) -> Option<Fill> {
    let r = color_ramp(g)?;
    Some(ramp_fill(
        &r,
        &r.stops
            .iter()
            .map(|s| (s.0, s.2.clone()))
            .collect::<Vec<_>>(),
    ))
}

fn same_geometry(a: &Ramp, b: &Ramp) -> bool {
    a.style == b.style
        && (a.angle - b.angle).abs() < 1e-6
        && (a.border - b.border).abs() < 1e-6
        && (a.cx - b.cx).abs() < 1e-6
        && (a.cy - b.cy).abs() < 1e-6
}

/// A solid colour under an opacity gradient (§20.93 `draw:opacity-name`):
/// a gradient of the same colour whose stops carry the opacity.
fn opacity_gradient_fill(c: &str, og: Nd<'_>) -> Fill {
    let r = opacity_ramp(og);
    ramp_fill(
        &r,
        &r.stops
            .iter()
            .map(|s| (s.0, with_alpha(c, s.1)))
            .collect::<Vec<_>>(),
    )
}

fn apply_opacity(fill: &mut Fill, opacity: f64, gradient: Option<Nd<'_>>) {
    let Fill::Gradient { stops, .. } = fill else {
        return;
    };
    match gradient {
        None => {
            for s in stops.iter_mut() {
                s.color = with_alpha(&s.color, opacity);
            }
        }
        Some(og) => {
            // Opacity is sampled at each colour stop's axis position. The two
            // gradients share their geometry in practice (LibreOffice writes
            // the transparency gradient beside the colour gradient it was
            // authored with); a different geometry uses its mean opacity.
            let r = opacity_ramp(og);
            let mean = r.stops.iter().map(|s| s.1).sum::<f64>() / r.stops.len() as f64;
            let _ = same_geometry;
            for s in stops.iter_mut() {
                s.color = with_alpha(&s.color, mean);
            }
        }
    }
}

/// Combine a colour gradient and its transparency gradient when both share a
/// geometry: the union of their stop offsets, each painted with both.
pub fn combined_gradient(g: Nd<'_>, og: Nd<'_>) -> Option<Fill> {
    let colour = color_ramp(g)?;
    let opacity = opacity_ramp(og);
    if !same_geometry(&colour, &opacity) {
        return None;
    }
    let mut offsets: Vec<f64> = colour
        .stops
        .iter()
        .chain(opacity.stops.iter())
        .map(|s| s.0)
        .collect();
    offsets.sort_by(|a, b| a.total_cmp(b));
    offsets.dedup_by(|a, b| (*a - *b).abs() < 1e-9);
    // Both ramps are sampled at ascending offsets, each from where the
    // previous sample was found (linear in the stop counts).
    let (mut ci, mut oi) = (0usize, 0usize);
    let painted: Vec<(f64, String)> = offsets
        .iter()
        .map(|&o| {
            let c = sample_from(&colour.stops, o, &mut ci).1;
            let a = sample_from(&opacity.stops, o, &mut oi).0;
            (o, with_alpha(&c, a))
        })
        .collect();
    Some(ramp_fill(&colour, &painted))
}

/// `draw:hatch` (§16.42) drawn exactly for an area of `area` EMU: parallel
/// hairlines `draw:distance` apart at `draw:rotation` (counter-clockwise),
/// a second family at +90° for `double` and a third at +45° for `triple`,
/// over the area colour when `draw:fill-hatch-solid` is set. Lines run
/// through the area centre, as LibreOffice's hatch texture is laid out on the
/// object range. The SVG stretches over the area like a picture fill, so it
/// turns and mirrors with the shape. `None` for an empty area.
fn hatch_svg(h: Nd<'_>, bg: &str, opacity: f64, area: (f64, f64)) -> Option<Vec<u8>> {
    let (w, ht) = area;
    if !(w > 0.0 && ht > 0.0) {
        return None;
    }
    let fg = h
        .attribute((NS_DRAW, "color"))
        .and_then(color)
        .unwrap_or_else(|| "000000".into());
    let distance = h
        .attribute((NS_DRAW, "distance"))
        .and_then(length_emu)
        .unwrap_or(0.1 * 360_000.0)
        .max(9_525.0);
    let rotation = match h.attribute((NS_DRAW, "rotation")) {
        Some(v) if v.trim().chars().all(|c| c.is_ascii_digit() || c == '-') => {
            v.trim().parse::<f64>().unwrap_or(0.0) / 10.0
        }
        Some(v) => angle_degrees(v).unwrap_or(0.0),
        None => 0.0,
    };
    let families: &[f64] = match h.attribute((NS_DRAW, "style")).unwrap_or("single") {
        "double" => &[0.0, 90.0],
        "triple" => &[0.0, 90.0, 45.0],
        _ => &[0.0],
    };
    // Implementation bound on the lines drawn (very fine hatches on large
    // areas): beyond it the distance is widened.
    let half_diag = 0.5 * w.hypot(ht);
    let mut step = distance;
    let per_family = (2.0 * half_diag / step).ceil();
    if per_family > 4000.0 {
        step = 2.0 * half_diag / 4000.0;
    }
    let mut d = String::new();
    let (cx, cy) = (w / 2.0, ht / 2.0);
    for extra in families {
        // Counter-clockwise on the page: direction (cos, -sin).
        let a = (rotation + extra).to_radians();
        let (dx, dy) = (a.cos(), -a.sin());
        let (nx, ny) = (-dy, dx);
        let count = (half_diag / step).ceil() as i64;
        for k in -count..=count {
            let off = k as f64 * step;
            let (px, py) = (cx + nx * off, cy + ny * off);
            let (x0, y0) = (px - dx * half_diag, py - dy * half_diag);
            let (x1, y1) = (px + dx * half_diag, py + dy * half_diag);
            d.push_str(&format!("M{x0:.0} {y0:.0}L{x1:.0} {y1:.0}"));
        }
    }
    let background = if bg.len() == 6 {
        format!(r##"<rect width="{w:.0}" height="{ht:.0}" fill="#{bg}"/>"##)
    } else if bg.len() == 8 && !bg.ends_with("00") {
        let alpha = u8::from_str_radix(&bg[6..8], 16).unwrap_or(255) as f64 / 255.0;
        format!(
            r##"<rect width="{w:.0}" height="{ht:.0}" fill="#{}" fill-opacity="{alpha:.3}"/>"##,
            &bg[..6]
        )
    } else {
        String::new()
    };
    let svg = format!(
        r##"<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {w:.0} {ht:.0}" width="{w:.0}" height="{ht:.0}" preserveAspectRatio="none">{background}<path d="{d}" stroke="#{fg}" stroke-opacity="{opacity:.3}" stroke-width="9525" fill="none"/></svg>"##
    );
    Some(svg.into_bytes())
}

/// `draw:hatch` (§16.41) as the nearest DrawingML preset pattern: line
/// direction (`draw:rotation`) and single/double/triple style select it.
fn hatch_fill(h: Nd<'_>, bg: String, opacity: f64) -> Fill {
    let fg = h
        .attribute((NS_DRAW, "color"))
        .and_then(color)
        .map(|c| with_alpha(&c, opacity))
        .unwrap_or_else(|| "000000".into());
    let rotation = match h.attribute((NS_DRAW, "rotation")) {
        Some(v) if v.trim().chars().all(|c| c.is_ascii_digit() || c == '-') => {
            v.trim().parse::<f64>().unwrap_or(0.0) / 10.0
        }
        Some(v) => angle_degrees(v).unwrap_or(0.0),
        None => 0.0,
    }
    .rem_euclid(180.0);
    let style = h.attribute((NS_DRAW, "style")).unwrap_or("single");
    let near = |target: f64| {
        (rotation - target).abs() <= 22.5 || (rotation - target - 180.0).abs() <= 22.5
    };
    let preset = match style {
        "double" | "triple" => {
            if near(45.0) || near(135.0) {
                "diagCross"
            } else {
                "cross"
            }
        }
        _ => {
            if near(90.0) {
                "vert"
            } else if near(45.0) {
                "upDiag"
            } else if near(135.0) {
                "dnDiag"
            } else {
                "horz"
            }
        }
    };
    Fill::Pattern {
        fg,
        bg,
        preset: preset.into(),
    }
}

/// Bitmap area fill (§20.88 `draw:fill-image-name`, §20.124 `style:repeat`).
fn image_fill(
    path: String,
    resource: &crate::odp_media::Resource,
    props: &Props<'_>,
    opacity: f64,
    area: (f64, f64),
) -> Fill {
    let repeat = props.get(NS_STYLE, "repeat").unwrap_or("repeat");
    let natural = resource.natural_emu();
    let size = |name: &str, natural: Option<f64>| -> Option<f64> {
        let v = props.get(NS_DRAW, name)?;
        if let Some(p) = percent(v) {
            return natural.map(|n| n * p);
        }
        length_emu(v).filter(|v| *v > 0.0)
    };
    let width = size("fill-image-width", natural.map(|n| n.0));
    let height = size("fill-image-height", natural.map(|n| n.1));
    let ref_point = props
        .get(NS_DRAW, "fill-image-ref-point")
        .unwrap_or("center");
    let algn = match ref_point {
        "top-left" => "tl",
        "top" => "t",
        "top-right" => "tr",
        "left" => "l",
        "right" => "r",
        "bottom-left" => "bl",
        "bottom" => "b",
        "bottom-right" => "br",
        _ => "ctr",
    };
    let alpha = (opacity < 1.0).then_some(opacity);
    let mut fill = Fill::Image {
        image_path: path,
        mime_type: resource.mime.to_string(),
        svg_image_path: None,
        dpi: None,
        rot_with_shape: Some(true),
        src_rect: None,
        fill_rect: None,
        stretch: false,
        tile: None,
        alpha,
        duotone: None,
        blip_effects: Vec::new(),
    };
    let Fill::Image {
        stretch,
        tile,
        dpi,
        fill_rect,
        ..
    } = &mut fill
    else {
        unreachable!()
    };
    match repeat {
        "stretch" => *stretch = true,
        "no-repeat" => {
            // One copy at its size, placed at the reference point, as
            // stretch insets relative to the area.
            *stretch = true;
            *fill_rect = no_repeat_rect(area, width.zip(height), algn);
        }
        _ => {
            // Tile period: the stated size, or the natural size. The renderer
            // sizes a tile as natural pixels × 72 / dpi × sx (points).
            let (sx, sy) = match (resource.pixels, width, height) {
                (Some((pw, ph)), Some(w), Some(h)) if pw > 0 && ph > 0 => {
                    *dpi = Some(96);
                    (w / (pw as f64 * 9525.0), h / (ph as f64 * 9525.0))
                }
                _ => {
                    *dpi = resource.dpi.map(|d| d.0.round().max(1.0) as u32);
                    (1.0, 1.0)
                }
            };
            let offset = |name: &str| props.get(NS_DRAW, name).and_then(percent).unwrap_or(0.0);
            let (ox, oy) = (
                offset("fill-image-ref-point-x"),
                offset("fill-image-ref-point-y"),
            );
            *tile = Some(TileInfo {
                tx: Some((ox * width.unwrap_or(0.0)).round() as i64),
                ty: Some((oy * height.unwrap_or(0.0)).round() as i64),
                sx: Some(sx),
                sy: Some(sy),
                flip: None,
                algn: Some(algn.into()),
            });
        }
    }
    fill
}

fn no_repeat_rect(area: (f64, f64), size: Option<(f64, f64)>, algn: &str) -> Option<FillRect> {
    let (w, h) = area;
    let (iw, ih) = size?;
    if w <= 0.0 || h <= 0.0 || iw <= 0.0 || ih <= 0.0 {
        return None;
    }
    let x = if algn.ends_with('l') {
        0.0
    } else if algn.ends_with('r') {
        w - iw
    } else {
        (w - iw) / 2.0
    };
    let y = if algn.starts_with('t') {
        0.0
    } else if algn.starts_with('b') {
        h - ih
    } else {
        (h - ih) / 2.0
    };
    Some(FillRect {
        l: x / w,
        t: y / h,
        r: (w - x - iw) / w,
        b: (h - y - ih) / h,
    })
}

/// The resolved line of a drawing object (§20.164 `draw:stroke`).
pub struct Line<'a> {
    pub stroke: Option<Stroke>,
    /// Marker paths for the start and end, with their width in EMU and
    /// whether they are centred on the end point.
    pub start_marker: Option<Marker<'a>>,
    pub end_marker: Option<Marker<'a>>,
}

#[derive(Clone)]
pub struct Marker<'a> {
    pub view_box: (f64, f64, f64, f64),
    /// The marker's `svg:d`, borrowed from its definition (parsed once per
    /// definition by the reader).
    pub d: &'a str,
    pub width: f64,
    pub center: bool,
    pub color: String,
}

/// Hairline width: LibreOffice draws a zero-width line one device pixel
/// wide; the model's thinnest line is used (1 px at 96 dpi).
const HAIRLINE_EMU: i64 = 9525;

pub fn line<'a>(props: &Props<'a>, sheet: &Sheet<'a>) -> Line<'a> {
    let kind = props.get(NS_DRAW, "stroke").unwrap_or("solid");
    if kind == "none" {
        return Line {
            stroke: None,
            start_marker: None,
            end_marker: None,
        };
    }
    let opacity = props
        .get(NS_SVG, "stroke-opacity")
        .and_then(|v| percent(v).or_else(|| v.parse::<f64>().ok()))
        .unwrap_or(1.0);
    let base = props
        .get(NS_SVG, "stroke-color")
        .and_then(color)
        .unwrap_or_else(|| "000000".into());
    let colour = with_alpha(&base, opacity);
    let width_emu = props
        .get(NS_SVG, "stroke-width")
        .and_then(length_emu)
        .unwrap_or(0.0);
    let width = if width_emu <= 0.0 {
        HAIRLINE_EMU
    } else {
        width_emu.round() as i64
    };
    let line_cap = props.get(NS_SVG, "stroke-linecap").map(|cap| match cap {
        "round" => "round".to_string(),
        "square" => "square".to_string(),
        _ => "butt".to_string(),
    });
    let line_join = props
        .get(NS_DRAW, "stroke-linejoin")
        .and_then(|join| match join {
            "round" => Some("round".to_string()),
            "bevel" | "none" => Some("bevel".to_string()),
            "miter" => Some("miter".to_string()),
            _ => None,
        });
    let mut custom_dash = Vec::new();
    let mut cap = line_cap;
    if kind == "dash" {
        if let Some(dash) = props
            .get(NS_DRAW, "stroke-dash")
            .and_then(|name| sheet.named("stroke-dash", name))
        {
            custom_dash = dash_segments(dash, width as f64);
            if dash.attribute((NS_DRAW, "style")) == Some("round") && cap.is_none() {
                cap = Some("round".to_string());
            }
        }
    }
    let marker = |which: &str| -> Option<Marker<'a>> {
        let name = props.get(NS_DRAW, &format!("marker-{which}"))?;
        let node = sheet.named("marker", name)?;
        let view_box = node
            .attribute((NS_SVG, "viewBox"))
            .and_then(crate::odp_geometry::parse_view_box)?;
        let d = node.attribute((NS_SVG, "d"))?;
        let width = props
            .get(NS_DRAW, &format!("marker-{which}-width"))
            .and_then(length_emu)
            .unwrap_or(0.3 * 360_000.0);
        let center = props.get(NS_DRAW, &format!("marker-{which}-center")) == Some("true");
        Some(Marker {
            view_box,
            d,
            width,
            center,
            color: colour.clone(),
        })
    };
    Line {
        stroke: Some(Stroke {
            color: colour.clone(),
            width,
            fill: None,
            dash_style: None,
            custom_dash,
            line_cap: cap,
            line_join,
            miter_limit: None,
            alignment: None,
            head_end: None,
            tail_end: None,
            cmpd: None,
        }),
        start_marker: marker("start"),
        end_marker: marker("end"),
    }
}

/// `draw:stroke-dash` (§16.36): `dots1` dashes of `dots1-length`, then
/// `dots2` of `dots2-length`, each followed by `distance`. Lengths are
/// absolute or a percentage of the line width; a missing length is a dot
/// as long as the line is wide.
fn dash_segments(dash: Nd<'_>, width: f64) -> Vec<StrokeDashSegment> {
    let rel = |name: &str, default: f64| -> f64 {
        match dash.attribute((NS_DRAW, name)) {
            Some(v) => {
                if let Some(p) = percent(v) {
                    p
                } else {
                    length_emu(v).map(|l| l / width.max(1.0)).unwrap_or(default)
                }
            }
            None => default,
        }
    };
    let count = |name: &str| -> usize {
        dash.attribute((NS_DRAW, name))
            .and_then(|v| v.parse::<usize>().ok())
            .unwrap_or(0)
            .min(32)
    };
    let distance = rel("distance", 1.0).max(0.0);
    let mut out = Vec::new();
    for _ in 0..count("dots1") {
        out.push(StrokeDashSegment {
            dash: rel("dots1-length", 1.0).max(0.0),
            space: distance,
        });
    }
    for _ in 0..count("dots2") {
        out.push(StrokeDashSegment {
            dash: rel("dots2-length", 1.0).max(0.0),
            space: distance,
        });
    }
    out
}

/// Shape shadow (§20.144 `draw:shadow`, offsets, colour, opacity and
/// LibreOffice's `loext:shadow-blur`).
pub fn shadow(props: &Props<'_>) -> Option<Shadow> {
    if props.get(NS_DRAW, "shadow") != Some("visible") {
        return None;
    }
    let ox = props
        .get(NS_DRAW, "shadow-offset-x")
        .and_then(length_emu)
        .unwrap_or(0.0);
    let oy = props
        .get(NS_DRAW, "shadow-offset-y")
        .and_then(length_emu)
        .unwrap_or(0.0);
    let colour = props
        .get(NS_DRAW, "shadow-color")
        .and_then(color)
        .unwrap_or_else(|| "808080".into());
    let alpha = props
        .get(NS_DRAW, "shadow-opacity")
        .and_then(percent)
        .unwrap_or(1.0);
    let blur = props
        .get(NS_LOEXT, "shadow-blur")
        .and_then(length_emu)
        .unwrap_or(0.0);
    Some(Shadow {
        color: colour,
        alpha,
        blur: blur.round() as i64,
        dist: ox.hypot(oy).round() as i64,
        dir: oy.atan2(ox).to_degrees().rem_euclid(360.0),
        sx: None,
        sy: None,
        kx: None,
        ky: None,
        algn: None,
        rot_with_shape: Some(false),
    })
}

/// LibreOffice glow and soft edge (`loext:glow-radius`, `loext:softedge-radius`).
pub fn glow(props: &Props<'_>) -> Option<Glow> {
    let radius = props.get(NS_LOEXT, "glow-radius").and_then(length_emu)?;
    if radius <= 0.0 {
        return None;
    }
    let colour = props.get(NS_LOEXT, "glow-color").and_then(color)?;
    let transparency = props
        .get(NS_LOEXT, "glow-transparency")
        .and_then(percent)
        .unwrap_or(0.0);
    Some(Glow {
        color: colour,
        alpha: 1.0 - transparency,
        radius: radius.round() as i64,
    })
}

pub fn soft_edge(props: &Props<'_>) -> Option<SoftEdge> {
    let radius = props
        .get(NS_LOEXT, "softedge-radius")
        .and_then(length_emu)?;
    (radius > 0.0).then(|| SoftEdge {
        radius: radius.round() as i64,
    })
}

/// Page background (§20.4 `style:drawing-page-properties` fill).
pub fn page_fill(
    props: &Props<'_>,
    ctx: &mut PaintContext<'_, '_, '_>,
    area: (f64, f64),
) -> Option<Fill> {
    fill(props, ctx, area)
}

/// Whether the cascade states its own area fill (`draw:fill`).
pub fn has_fill_attribute(props: &Props<'_>) -> bool {
    props.get(NS_DRAW, "fill").is_some()
}

/// `fo:background-color` of paragraph / cell properties.
pub fn background_color(props: &Props<'_>) -> Option<String> {
    let v = props.get(NS_FO, "background-color")?;
    if v == "transparent" {
        return None;
    }
    color(v)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn linear_gradient_maps_angle_and_border() {
        let xml = r##"<draw:gradient xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" draw:name="g" draw:style="linear" draw:start-color="#000000" draw:end-color="#ffffff" draw:angle="0deg" draw:border="50%"/>"##;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let fill = gradient_fill(doc.root_element(), &Props::new()).unwrap();
        let Fill::Gradient {
            stops,
            angle,
            grad_type,
            ..
        } = fill
        else {
            panic!()
        };
        assert_eq!(grad_type, "linear");
        assert_eq!(angle, 90.0);
        assert_eq!(
            stops.iter().map(|s| s.position).collect::<Vec<_>>(),
            vec![0.0, 0.5, 1.0]
        );
        assert_eq!(stops[1].color, "000000");
        assert_eq!(stops[2].color, "FFFFFF");
    }

    #[test]
    fn radial_gradient_puts_end_colour_at_the_centre() {
        let xml = r##"<draw:gradient xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" draw:name="g" draw:style="radial" draw:cx="25%" draw:cy="50%" draw:start-color="#ff0000" draw:end-color="#0000ff"/>"##;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let Fill::Gradient {
            stops,
            path,
            fill_to_rect,
            ..
        } = gradient_fill(doc.root_element(), &Props::new()).unwrap()
        else {
            panic!()
        };
        assert_eq!(path.as_deref(), Some("circle"));
        assert_eq!(stops[0].color, "0000FF");
        assert_eq!(stops.last().unwrap().color, "FF0000");
        assert_eq!(fill_to_rect.unwrap().l, 0.25);
    }
}
