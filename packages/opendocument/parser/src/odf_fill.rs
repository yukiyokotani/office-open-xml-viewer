//! Area fills of drawing styles (ODF Part 3 §20.130 `draw:fill`): solid
//! colours, gradients (§16.40 `draw:gradient`), hatches (§16.42
//! `draw:hatch`) and pictures (§16.41 `draw:fill-image`), resolved from a
//! graphic style chain into a format-neutral description. The spreadsheet
//! reader maps it to the workbook drawing model; the presentation reader can
//! map it to its own.
//!
//! The opacity of the area (§20.152 `draw:opacity`, or the opacity gradient
//! §20.153 `draw:opacity-name` names, which LibreOffice applies instead of
//! it) is part of the description: colours of a translucent fill are
//! `RRGGBBAA`, and hatches and pictures carry their opacity. A solid colour
//! under an opacity gradient is a gradient of that colour; a colour gradient
//! under an opacity gradient of the same geometry has both in its stops.
//! Two gradients of different geometry are not one gradient: the area then
//! takes the mean opacity, as do hatches and pictures.
//!
//! Gradients are expressed as DrawingML gradients (ECMA-376 §20.1.8.33
//! `a:gradFill`): stop lists (`ooxml_common::fill::GradStop`), a linear angle
//! (§20.1.8.41, degrees clockwise from left-to-right) or a path shade
//! (§20.1.8.46) with its focus rectangle (§20.1.8.30 `a:fillToRect`).

use crate::iri::Target;
use crate::odf_props::{get, hex_color, length_pt, percent, Style};
use crate::xml::{self, NS_DRAW, NS_LOEXT, NS_SVG, NS_XLINK};
use ooxml_common::fill::{FillRect, GradStop};
use std::collections::HashMap;

/// Implementation bound on the stops of one gradient; more reject the
/// document.
const MAX_GRADIENT_STOPS: usize = 256;

/// A gradient's geometry in DrawingML terms.
#[derive(Clone, Debug, PartialEq)]
pub enum GradientShape {
    /// Linear, `angle` in degrees clockwise from left-to-right.
    Linear { angle: f64 },
    /// A path shade (`circle` or `rect`) radiating from `focus`.
    Path { path: &'static str, focus: FillRect },
}

#[derive(Clone, Debug, PartialEq)]
pub struct Gradient {
    /// Stops along the axis (linear) or from the focus outwards (path),
    /// colours `RRGGBB`, or `RRGGBBAA` where translucent.
    pub stops: Vec<GradStop>,
    pub shape: GradientShape,
    /// `draw:angle` of a path gradient, degrees counter-clockwise; DrawingML
    /// path shades have no rotation of their own.
    pub rotation: f64,
}

/// §16.42 `draw:hatch`: parallel lines, crossed for `double`, with a
/// diagonal for `triple`.
#[derive(Clone, Debug, PartialEq)]
pub struct Hatch {
    pub color: String,
    /// The area colour under the lines (`draw:fill-hatch-solid`).
    pub background: Option<String>,
    /// `draw:rotation` of the first line family, as LibreOffice draws it:
    /// degrees counter-clockwise from the horizontal. §19.209 words the
    /// attribute as clockwise from the vertical axis, but LibreOffice, whose
    /// files these are, draws rotation 0 as horizontal lines, 30° as lines
    /// rising to the right and 90° as vertical lines (measured with
    /// LibreOffice 26.8 on `draw:rotation` 0, 300, `30deg` and 900); the
    /// reader follows it. A `double` hatch adds the perpendicular family
    /// and a `triple` hatch the family 45° further counter-clockwise
    /// (measured at 30°: families at 30°, 120° and 75°).
    pub angle: f64,
    /// Distance between lines, in points.
    pub distance_pt: f64,
    /// 1 (single), 2 (double), 3 (triple).
    pub lines: u8,
    /// Opacity of the lines and the background, 0 to 1.
    pub opacity: f64,
}

/// How a picture covers its area (§20.341 `style:repeat`).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PictureLayout {
    /// `stretch`: scaled to the area.
    Stretch,
    /// `repeat` (the default LibreOffice applies): repeated at its size.
    Tile,
    /// `no-repeat`: drawn once at its size.
    Once,
}

/// One edge of a picture's size (§20.128 `draw:fill-image-width`).
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum PictureSize {
    Pt(f64),
    /// A percentage, of the filled area's edge.
    Fraction(f64),
}

/// §20.175 `draw:tile-repeat-offset`: every second row (`horizontal`) or
/// column (`vertical`) of tiles moved by a fraction of the tile.
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum TileShift {
    Rows(f64),
    Columns(f64),
}

/// §16.41 `draw:fill-image` as an area fill. Its geometry is what
/// LibreOffice 26.8 draws (measured on tiled, single and stretched fills):
/// a picture that is not stretched has its own size, sits where its
/// reference point puts it (the centre by default), and, when repeated, is
/// moved by the reference-point offsets and staggered by the tile offset.
/// LibreOffice honours the reference point of a single picture too, where
/// §20.125 limits it to repeated ones; its offsets and stagger apply to
/// repeated pictures only, as §20.126 and §20.175 say.
#[derive(Clone, Debug, PartialEq)]
pub struct Picture {
    /// Package path of the image.
    pub path: String,
    pub layout: PictureLayout,
    /// The size of the picture when it is not stretched
    /// (`draw:fill-image-width` / `-height`); `None`: its natural size.
    pub size: (Option<PictureSize>, Option<PictureSize>),
    /// §20.125 `draw:fill-image-ref-point`: where a single picture, or one
    /// tile of a repeated one, sits in the area, as fractions of the free
    /// width and height (0, 0.5 or 1).
    pub align: (f64, f64),
    /// §20.126 / §20.127 `draw:fill-image-ref-point-x` / `-y`: repeated
    /// tiles moved right and down by these fractions of the tile.
    pub offset: (f64, f64),
    pub shift: Option<TileShift>,
    /// Opacity of the picture, 0 to 1.
    pub opacity: f64,
}

#[derive(Clone, Debug, PartialEq)]
pub enum Fill {
    None,
    /// `RRGGBB`, or `RRGGBBAA` where translucent.
    Solid(String),
    Gradient(Gradient),
    Hatch(Hatch),
    Picture(Picture),
}

#[derive(Clone, Debug)]
struct GradientDef {
    style: String,
    start: String,
    end: String,
    start_intensity: f64,
    end_intensity: f64,
    angle: f64,
    border: f64,
    cx: f64,
    cy: f64,
    stops: Vec<(f64, String)>,
}

/// §16.40.? `draw:opacity`: a gradient of opacities, with the geometry
/// attributes of a colour gradient.
#[derive(Clone, Debug)]
struct OpacityDef {
    style: String,
    angle: f64,
    border: f64,
    cx: f64,
    cy: f64,
    /// Offsets from the start side and their opacities, 0 to 1.
    stops: Vec<(f64, f64)>,
}

/// Values laid out along a gradient's axis (linear) or from its focus
/// outwards (path), positions ascending.
struct Axis<T> {
    stops: Vec<(f64, T)>,
    shape: GradientShape,
    rotation: f64,
}

/// The opacity of an area.
enum Opacity {
    Uniform(f64),
    Graded(Axis<f64>),
}

#[derive(Clone, Debug)]
struct HatchDef {
    style: String,
    color: String,
    distance_pt: f64,
    angle: f64,
}

/// The named fill definitions of a document (`office:styles` children).
#[derive(Default)]
pub struct FillDefs {
    gradients: HashMap<String, GradientDef>,
    hatches: HashMap<String, HatchDef>,
    images: HashMap<String, String>,
    opacities: HashMap<String, OpacityDef>,
    dashes: HashMap<String, &'static str>,
    markers: HashMap<String, String>,
}

/// `draw:angle` (§19.218): an angle with a unit, or a bare number in tenths
/// of a degree (ODF 1.2 and LibreOffice's older output).
fn angle(value: Option<&str>) -> f64 {
    let Some(value) = value.map(str::trim) else {
        return 0.0;
    };
    for (unit, factor) in [
        ("deg", 1.0),
        ("grad", 0.9),
        ("rad", 180.0 / std::f64::consts::PI),
    ] {
        if let Some(number) = value.strip_suffix(unit) {
            return number.trim().parse::<f64>().map_or(0.0, |v| v * factor);
        }
    }
    value.parse::<f64>().map_or(0.0, |v| v / 10.0)
}

fn scaled(hex: &str, intensity: f64) -> String {
    if (intensity - 1.0).abs() < f64::EPSILON {
        return hex.to_string();
    }
    let channel = |at: usize| {
        let v = u8::from_str_radix(&hex[at..at + 2], 16).unwrap_or(0) as f64;
        (v * intensity).round().clamp(0.0, 255.0) as u8
    };
    format!("{:02X}{:02X}{:02X}", channel(0), channel(2), channel(4))
}

impl FillDefs {
    /// Read the definitions below `office:styles` of the package file
    /// `file_entry` (`styles.xml`), against which picture references
    /// resolve (ODF Part 2 §3.7).
    pub fn add(&mut self, styles: roxmltree::Node<'_, '_>, file_entry: &str) -> Result<(), String> {
        for node in styles.children().filter(|n| n.is_element()) {
            crate::budget::tick(1)?;
            let Some(name) = node.attribute((NS_DRAW, "name")) else {
                continue;
            };
            let attr = |key: &str| node.attribute((NS_DRAW, key));
            if xml::is(node, NS_DRAW, "gradient") {
                // LibreOffice 7.6 multi-colour stops: `loext:gradient-stop`
                // with `svg:offset` and `loext:color-value`.
                let stop_nodes: Vec<_> = node
                    .children()
                    .filter(|n| xml::is(*n, NS_LOEXT, "gradient-stop"))
                    .collect();
                if stop_nodes.len() > MAX_GRADIENT_STOPS {
                    return Err(crate::error::resource("gradient stop"));
                }
                let stops = stop_nodes
                    .into_iter()
                    .filter_map(|stop| {
                        let offset = stop
                            .attribute((NS_SVG, "offset"))?
                            .trim()
                            .parse::<f64>()
                            .ok()
                            .filter(|v| v.is_finite())?;
                        let color = stop
                            .attribute((NS_LOEXT, "color-value"))
                            .and_then(hex_color)?;
                        Some((offset.clamp(0.0, 1.0), color))
                    })
                    .collect();
                self.gradients.insert(
                    name.to_string(),
                    GradientDef {
                        style: attr("style").unwrap_or("linear").to_string(),
                        start: attr("start-color")
                            .and_then(hex_color)
                            .unwrap_or_else(|| "000000".into()),
                        end: attr("end-color")
                            .and_then(hex_color)
                            .unwrap_or_else(|| "FFFFFF".into()),
                        start_intensity: attr("start-intensity").and_then(percent).unwrap_or(1.0),
                        end_intensity: attr("end-intensity").and_then(percent).unwrap_or(1.0),
                        angle: angle(attr("angle")),
                        border: attr("border")
                            .and_then(percent)
                            .unwrap_or(0.0)
                            .clamp(0.0, 1.0),
                        cx: attr("cx").and_then(percent).unwrap_or(0.5).clamp(0.0, 1.0),
                        cy: attr("cy").and_then(percent).unwrap_or(0.5).clamp(0.0, 1.0),
                        stops,
                    },
                );
            } else if xml::is(node, NS_DRAW, "stroke-dash") {
                self.dashes.insert(name.to_string(), dash_preset(node));
            } else if xml::is(node, NS_DRAW, "marker") {
                if let Some(d) = node.attribute((NS_SVG, "d")) {
                    self.markers.insert(name.to_string(), d.to_string());
                }
            } else if xml::is(node, NS_DRAW, "opacity") {
                // LibreOffice 7.6 multi-step opacity: `loext:opacity-stop`
                // with `svg:offset` and `loext:stop-opacity` (0 to 1).
                let stop_nodes: Vec<_> = node
                    .children()
                    .filter(|n| xml::is(*n, NS_LOEXT, "opacity-stop"))
                    .collect();
                if stop_nodes.len() > MAX_GRADIENT_STOPS {
                    return Err(crate::error::resource("gradient stop"));
                }
                let unit = |v: f64| v.is_finite().then(|| v.clamp(0.0, 1.0));
                let mut stops: Vec<(f64, f64)> = stop_nodes
                    .into_iter()
                    .filter_map(|stop| {
                        let number = |name: (&str, &str)| {
                            stop.attribute(name)?
                                .trim()
                                .parse::<f64>()
                                .ok()
                                .and_then(unit)
                        };
                        Some((
                            number((NS_SVG, "offset"))?,
                            number((NS_LOEXT, "stop-opacity"))?,
                        ))
                    })
                    .collect();
                if stops.len() < 2 {
                    let end = |key: &str| attr(key).and_then(percent).and_then(unit).unwrap_or(1.0);
                    stops = vec![(0.0, end("start")), (1.0, end("end"))];
                }
                self.opacities.insert(
                    name.to_string(),
                    OpacityDef {
                        style: attr("style").unwrap_or("linear").to_string(),
                        angle: angle(attr("angle")),
                        border: attr("border")
                            .and_then(percent)
                            .unwrap_or(0.0)
                            .clamp(0.0, 1.0),
                        cx: attr("cx").and_then(percent).unwrap_or(0.5).clamp(0.0, 1.0),
                        cy: attr("cy").and_then(percent).unwrap_or(0.5).clamp(0.0, 1.0),
                        stops,
                    },
                );
            } else if xml::is(node, NS_DRAW, "hatch") {
                self.hatches.insert(
                    name.to_string(),
                    HatchDef {
                        style: attr("style").unwrap_or("single").to_string(),
                        color: attr("color")
                            .and_then(hex_color)
                            .unwrap_or_else(|| "000000".into()),
                        distance_pt: attr("distance")
                            .and_then(length_pt)
                            .unwrap_or(2.0)
                            .max(0.25),
                        angle: angle(attr("rotation")),
                    },
                );
            } else if xml::is(node, NS_DRAW, "fill-image") {
                // §16.41: the image is a package file named by an IRI
                // relative to the containing file (Part 2 §3.7).
                if let Some(Target::Package { path, .. }) = node
                    .attribute((NS_XLINK, "href"))
                    .and_then(|href| crate::iri::resolve(file_entry, href))
                {
                    if !path.is_empty() {
                        self.images.insert(name.to_string(), path);
                    }
                }
            }
        }
        Ok(())
    }

    /// The DrawingML preset of the dash style `name` (`draw:stroke-dash`).
    pub fn dash(&self, name: &str) -> Option<&'static str> {
        self.dashes.get(name).copied()
    }

    /// The line ends (§16.42.8 `draw:marker`): each name with its path
    /// data. A reader reads each path once, not once per line that uses it.
    pub fn markers(&self) -> impl Iterator<Item = (&str, &str)> {
        self.markers
            .iter()
            .map(|(name, d)| (name.as_str(), d.as_str()))
    }

    /// The fill of a graphic style chain (`graphic:` properties), with its
    /// opacity.
    pub fn resolve(&self, chain: &[&Style]) -> Fill {
        let property = |key: &str| get(chain, &format!("graphic:draw:{key}"));
        let opacity = match property("opacity-name").and_then(|name| self.opacities.get(name)) {
            Some(def) => Opacity::Graded(lay_out(
                &def.style,
                (def.angle, def.border, def.cx, def.cy),
                def.stops.clone(),
            )),
            None => Opacity::Uniform(
                property("opacity")
                    .and_then(percent)
                    .filter(|v| v.is_finite())
                    .map_or(1.0, |v| v.clamp(0.0, 1.0)),
            ),
        };
        translucent(self.opaque(chain), opacity)
    }

    fn opaque(&self, chain: &[&Style]) -> Fill {
        let property = |key: &str| get(chain, &format!("graphic:draw:{key}"));
        let solid = || property("fill-color").and_then(hex_color);
        match property("fill") {
            Some("none") => Fill::None,
            Some("gradient") => property("fill-gradient-name")
                .and_then(|name| self.gradients.get(name))
                .map_or_else(
                    || solid().map_or(Fill::None, Fill::Solid),
                    |def| Fill::Gradient(gradient(def)),
                ),
            Some("hatch") => property("fill-hatch-name")
                .and_then(|name| self.hatches.get(name))
                .map_or_else(
                    || solid().map_or(Fill::None, Fill::Solid),
                    |def| {
                        Fill::Hatch(Hatch {
                            color: def.color.clone(),
                            background: (property("fill-hatch-solid") == Some("true"))
                                .then(solid)
                                .flatten(),
                            angle: def.angle,
                            distance_pt: def.distance_pt,
                            lines: match def.style.as_str() {
                                "double" => 2,
                                "triple" => 3,
                                _ => 1,
                            },
                            opacity: 1.0,
                        })
                    },
                ),
            Some("bitmap") => property("fill-image-name")
                .and_then(|name| self.images.get(name))
                .map(|path| {
                    let size = |key: &str| {
                        let value = property(key)?;
                        length_pt(value)
                            .map(PictureSize::Pt)
                            .or_else(|| percent(value).map(PictureSize::Fraction))
                            .filter(|size| match size {
                                PictureSize::Pt(v) | PictureSize::Fraction(v) => {
                                    v.is_finite() && *v > 0.0
                                }
                            })
                    };
                    // A move by whole tiles changes nothing: offsets and
                    // stagger are kept as positions within one tile.
                    let within_tile = |v: f64| v.rem_euclid(1.0);
                    let fraction = |key: &str| {
                        property(key)
                            .and_then(percent)
                            .filter(|v| v.is_finite())
                            .map_or(0.0, within_tile)
                    };
                    let layout = match get(chain, "graphic:style:repeat") {
                        Some("stretch") => PictureLayout::Stretch,
                        Some("no-repeat") => PictureLayout::Once,
                        _ => PictureLayout::Tile,
                    };
                    let tiled = layout == PictureLayout::Tile;
                    let reference = property("fill-image-ref-point").unwrap_or("center");
                    Fill::Picture(Picture {
                        path: path.clone(),
                        layout,
                        size: (size("fill-image-width"), size("fill-image-height")),
                        align: (
                            if reference.ends_with("left") {
                                0.0
                            } else if reference.ends_with("right") {
                                1.0
                            } else {
                                0.5
                            },
                            if reference.starts_with("top") {
                                0.0
                            } else if reference.starts_with("bottom") {
                                1.0
                            } else {
                                0.5
                            },
                        ),
                        offset: if tiled {
                            (
                                fraction("fill-image-ref-point-x"),
                                fraction("fill-image-ref-point-y"),
                            )
                        } else {
                            (0.0, 0.0)
                        },
                        shift: property("tile-repeat-offset")
                            .filter(|_| tiled)
                            .and_then(|value| {
                                let mut parts = value.split_whitespace();
                                let amount = percent(parts.next()?)
                                    .filter(|v| v.is_finite())
                                    .map(within_tile)
                                    .filter(|v| *v != 0.0)?;
                                match parts.next()? {
                                    "horizontal" => Some(TileShift::Rows(amount)),
                                    "vertical" => Some(TileShift::Columns(amount)),
                                    _ => None,
                                }
                            }),
                        opacity: 1.0,
                    })
                })
                .unwrap_or_else(|| solid().map_or(Fill::None, Fill::Solid)),
            _ => solid().map_or(Fill::None, Fill::Solid),
        }
    }
}

/// `hex` with the alpha of `opacity`, when it is not opaque.
pub(crate) fn with_alpha(hex: &str, opacity: f64) -> String {
    let alpha = (opacity.clamp(0.0, 1.0) * 255.0).round() as u8;
    match hex.get(..6) {
        Some(rgb) if alpha < 255 => format!("{rgb}{alpha:02X}"),
        _ => hex.to_string(),
    }
}

/// The value of ascending `stops` at `position`, by `mix(a, b, t)` between
/// the stops around it.
fn sample<T: Clone>(stops: &[(f64, T)], position: f64, mix: impl Fn(&T, &T, f64) -> T) -> T {
    let after = stops.partition_point(|(at, _)| *at <= position);
    match (after.checked_sub(1).map(|i| &stops[i]), stops.get(after)) {
        (Some((from, a)), Some((to, b))) if to > from => mix(a, b, (position - from) / (to - from)),
        (Some((_, a)), _) => a.clone(),
        (None, Some((_, b))) => b.clone(),
        (None, None) => unreachable!("a gradient has stops"),
    }
}

fn mix_number(a: &f64, b: &f64, t: f64) -> f64 {
    a + (b - a) * t
}

fn mix_colour(a: &str, b: &str, t: f64) -> String {
    let channel = |at: usize| {
        let of = |hex: &str| {
            hex.get(at..at + 2)
                .and_then(|v| u8::from_str_radix(v, 16).ok())
                .unwrap_or(0) as f64
        };
        mix_number(&of(a), &of(b), t).round().clamp(0.0, 255.0) as u8
    };
    format!("{:02X}{:02X}{:02X}", channel(0), channel(2), channel(4))
}

/// The mean of an opacity gradient along its axis.
fn mean(stops: &[(f64, f64)]) -> f64 {
    let mut total = 0.0;
    let mut previous = (0.0, stops.first().map_or(1.0, |s| s.1));
    for &(at, value) in stops
        .iter()
        .chain([&(1.0, stops.last().map_or(1.0, |s| s.1))])
    {
        total += (at - previous.0).max(0.0) * (value + previous.1) / 2.0;
        previous = (at, value);
    }
    total.clamp(0.0, 1.0)
}

/// `fill` seen through `opacity`.
fn translucent(fill: Fill, opacity: Opacity) -> Fill {
    let uniform = |opacity: &Opacity| match opacity {
        Opacity::Uniform(value) => *value,
        Opacity::Graded(axis) => mean(&axis.stops),
    };
    match (fill, opacity) {
        (Fill::None, _) => Fill::None,
        (Fill::Solid(colour), Opacity::Graded(axis)) => Fill::Gradient(Gradient {
            stops: axis
                .stops
                .iter()
                .map(|(at, value)| stop(*at, &with_alpha(&colour, *value)))
                .collect(),
            shape: axis.shape,
            rotation: axis.rotation,
        }),
        (Fill::Solid(colour), opacity) => Fill::Solid(with_alpha(&colour, uniform(&opacity))),
        (Fill::Gradient(mut gradient), opacity) => {
            match opacity {
                // Both gradients along one axis: every stop of either, with
                // the other's value there.
                Opacity::Graded(axis)
                    if axis.shape == gradient.shape && axis.rotation == gradient.rotation =>
                {
                    let colours: Vec<(f64, String)> = gradient
                        .stops
                        .iter()
                        .map(|s| (s.position, s.color.clone()))
                        .collect();
                    let mut stops: Vec<GradStop> = colours
                        .iter()
                        .map(|(at, colour)| {
                            stop(
                                *at,
                                &with_alpha(colour, sample(&axis.stops, *at, mix_number)),
                            )
                        })
                        .chain(axis.stops.iter().map(|(at, value)| {
                            stop(
                                *at,
                                &with_alpha(
                                    &sample(&colours, *at, |a, b, t| mix_colour(a, b, t)),
                                    *value,
                                ),
                            )
                        }))
                        .collect();
                    stops.sort_by(|a, b| a.position.total_cmp(&b.position));
                    gradient.stops = stops;
                }
                opacity => {
                    let value = uniform(&opacity);
                    for stop in &mut gradient.stops {
                        stop.color = with_alpha(&stop.color, value);
                    }
                }
            }
            Fill::Gradient(gradient)
        }
        (Fill::Hatch(hatch), opacity) => Fill::Hatch(Hatch {
            opacity: uniform(&opacity),
            ..hatch
        }),
        (Fill::Picture(picture), opacity) => Fill::Picture(Picture {
            opacity: uniform(&opacity),
            ..picture
        }),
    }
}

/// A `draw:stroke-dash` (§16.41): dot runs and distance relative to the
/// line width, matched to the DrawingML preset with the same pattern
/// (ECMA-376 §20.1.10.48).
pub(crate) fn dash_preset(node: roxmltree::Node<'_, '_>) -> &'static str {
    let length = |key: &str| -> f64 {
        node.attribute((NS_DRAW, key))
            .and_then(|v| percent(v).or_else(|| length_pt(v).map(|pt| pt / 0.75)))
            .unwrap_or(1.0)
    };
    let count = |key: &str| -> u32 {
        node.attribute((NS_DRAW, key))
            .and_then(|v| v.parse().ok())
            .unwrap_or(0)
    };
    let (d1, l1, d2, l2) = (
        count("dots1"),
        length("dots1-length"),
        count("dots2"),
        length("dots2-length"),
    );
    let gap = length("distance");
    match (d1, d2) {
        (_, 0) | (0, _) => {
            let dash = if d1 > 0 { l1 } else { l2 };
            if dash <= 1.5 {
                "sysDot"
            } else if dash >= 6.0 {
                "lgDash"
            } else if gap <= 1.5 {
                "sysDash"
            } else {
                "dash"
            }
        }
        (_, n) => {
            let long = l1.max(l2);
            match (long >= 6.0, n >= 2 || d1 >= 2) {
                (true, true) => "lgDashDotDot",
                (true, false) => "lgDashDot",
                (false, true) => "sysDashDotDot",
                (false, false) => {
                    if gap <= 1.5 {
                        "sysDashDot"
                    } else {
                        "dashDot"
                    }
                }
            }
        }
    }
}

fn stop(position: f64, color: &str) -> GradStop {
    GradStop {
        position: position.clamp(0.0, 1.0),
        color: color.to_string(),
    }
}

/// §16.40.? `draw:gradient` styles as DrawingML gradients. LibreOffice's
/// geometry: `linear` runs from the start colour at the top to the end
/// colour at the bottom before rotating by `draw:angle` counter-clockwise;
/// `axial` has the start colour at both edges and the end colour on the
/// axis; the other styles have the start colour at the outside and the end
/// colour at the centre (`draw:cx`/`draw:cy`). `draw:border` keeps that
/// share of the start side in the start colour.
fn gradient(def: &GradientDef) -> Gradient {
    let start = scaled(&def.start, def.start_intensity);
    let end = scaled(&def.end, def.end_intensity);
    // Multi-colour stops (LibreOffice 7.6 `loext:gradient-stop`) run from
    // the start side; otherwise the two colours.
    let colours: Vec<(f64, String)> = if def.stops.len() >= 2 {
        def.stops.clone()
    } else {
        vec![(0.0, start), (1.0, end)]
    };
    let axis = lay_out(&def.style, (def.angle, def.border, def.cx, def.cy), colours);
    Gradient {
        stops: axis.stops.iter().map(|(at, c)| stop(*at, c)).collect(),
        shape: axis.shape,
        rotation: axis.rotation,
    }
}

/// Lay values given from the start side of a `style` gradient (colours, or
/// the opacities of a `draw:opacity`) out in DrawingML terms.
fn lay_out<T: Clone>(
    style: &str,
    (angle, border, cx, cy): (f64, f64, f64, f64),
    mut values: Vec<(f64, T)>,
) -> Axis<T> {
    values.sort_by(|a, b| a.0.total_cmp(&b.0));
    let at = |position: f64, value: &T| (position.clamp(0.0, 1.0), value.clone());
    // Map an offset in [0, 1] from the start side, after the border.
    let along = |t: f64| border + (1.0 - border) * t;
    match style {
        "linear" => {
            let mut stops = Vec::new();
            if border > 0.0 {
                stops.push(at(0.0, &values[0].1));
            }
            stops.extend(values.iter().map(|(t, c)| at(along(*t), c)));
            Axis {
                stops,
                // Start at the top (DrawingML 90°), rotated counter-clockwise.
                shape: GradientShape::Linear {
                    angle: (90.0 - angle).rem_euclid(360.0),
                },
                rotation: 0.0,
            }
        }
        "axial" => {
            let half = |t: f64| along(t) / 2.0;
            let mut stops = Vec::new();
            if border > 0.0 {
                stops.push(at(0.0, &values[0].1));
            }
            stops.extend(values.iter().map(|(t, c)| at(half(*t), c)));
            stops.extend(values.iter().rev().map(|(t, c)| at(1.0 - half(*t), c)));
            if border > 0.0 {
                stops.push(at(1.0, &values[0].1));
            }
            Axis {
                stops,
                shape: GradientShape::Linear {
                    angle: (90.0 - angle).rem_euclid(360.0),
                },
                rotation: 0.0,
            }
        }
        style => {
            // Path shades run from the focus (position 0) outwards; ODF's
            // start value is outside, so the stops are reversed.
            let mut stops: Vec<(f64, T)> = values
                .iter()
                .rev()
                .map(|(t, c)| at((1.0 - along(*t)).max(0.0), c))
                .collect();
            if border > 0.0 {
                stops.push(at(1.0, &values[0].1));
            }
            stops.sort_by(|a, b| a.0.total_cmp(&b.0));
            Axis {
                stops,
                shape: GradientShape::Path {
                    // §19.218.2 spells the rectangular blend `rectangle` in
                    // its definitions and `rectangular` in its value list
                    // and schema (which LibreOffice writes): both are read.
                    path: if matches!(style, "square" | "rectangular" | "rectangle") {
                        "rect"
                    } else {
                        "circle"
                    },
                    focus: FillRect {
                        l: cx,
                        t: cy,
                        r: 1.0 - cx,
                        b: 1.0 - cy,
                    },
                },
                rotation: angle,
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::odf_props::StyleStore;

    fn resolve(defs: &str, style: &str) -> Fill {
        let xml = format!(
            r##"<office:styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:draw="{NS_DRAW}" xmlns:xlink="{NS_XLINK}" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0">{defs}<style:style style:name="g" style:family="graphic"><style:graphic-properties {style}/></style:style></office:styles>"##
        );
        let document = roxmltree::Document::parse(&xml).unwrap();
        let mut fills = FillDefs::default();
        fills.add(document.root_element(), "styles.xml").unwrap();
        let mut store = StyleStore::default();
        store.add(document.root_element(), false).unwrap();
        fills.resolve(&store.chain("graphic", Some("g")))
    }

    #[test]
    fn linear_axial_and_path_gradients() {
        let linear = resolve(
            r##"<draw:gradient draw:name="G" draw:style="linear" draw:start-color="#ff0000" draw:end-color="#0000ff" draw:angle="90deg" draw:border="20%" draw:start-intensity="50%"/>"##,
            r#"draw:fill="gradient" draw:fill-gradient-name="G""#,
        );
        let Fill::Gradient(g) = linear else {
            panic!("{linear:?}")
        };
        assert_eq!(g.shape, GradientShape::Linear { angle: 0.0 });
        assert_eq!(
            g.stops
                .iter()
                .map(|s| (s.position, s.color.as_str()))
                .collect::<Vec<_>>(),
            [(0.0, "800000"), (0.2, "800000"), (1.0, "0000FF")]
        );
        let axial = resolve(
            r##"<draw:gradient draw:name="A" draw:style="axial" draw:start-color="#000000" draw:end-color="#ffffff" draw:angle="0"/>"##,
            r#"draw:fill="gradient" draw:fill-gradient-name="A""#,
        );
        let Fill::Gradient(a) = axial else { panic!() };
        assert_eq!(
            a.stops
                .iter()
                .map(|s| (s.position, s.color.as_str()))
                .collect::<Vec<_>>(),
            [
                (0.0, "000000"),
                (0.5, "FFFFFF"),
                (0.5, "FFFFFF"),
                (1.0, "000000")
            ]
        );
        assert_eq!(a.shape, GradientShape::Linear { angle: 90.0 });
        let radial = resolve(
            r##"<draw:gradient draw:name="R" draw:style="square" draw:cx="25%" draw:cy="75%" draw:start-color="#000000" draw:end-color="#ffffff" draw:border="10%"/>"##,
            r#"draw:fill="gradient" draw:fill-gradient-name="R""#,
        );
        let Fill::Gradient(r) = radial else { panic!() };
        assert_eq!(
            r.shape,
            GradientShape::Path {
                path: "rect",
                focus: FillRect {
                    l: 0.25,
                    t: 0.75,
                    r: 0.75,
                    b: 0.25
                }
            }
        );
        assert_eq!(r.stops.first().map(|s| s.color.as_str()), Some("FFFFFF"));
        assert_eq!(
            r.stops.last().map(|s| (s.position, s.color.as_str())),
            Some((1.0, "000000"))
        );
    }

    #[test]
    fn hatches_and_pictures() {
        let hatch = resolve(
            r##"<draw:hatch draw:name="H" draw:style="double" draw:color="#00ff00" draw:distance="0.1cm" draw:rotation="450"/>"##,
            r##"draw:fill="hatch" draw:fill-hatch-name="H" draw:fill-hatch-solid="true" draw:fill-color="#ffffff""##,
        );
        assert_eq!(
            hatch,
            Fill::Hatch(Hatch {
                color: "00FF00".into(),
                background: Some("FFFFFF".into()),
                angle: 45.0,
                distance_pt: length_pt("0.1cm").unwrap(),
                lines: 2,
                opacity: 1.0,
            })
        );
        let picture = resolve(
            r#"<draw:fill-image draw:name="P" xlink:href="Pictures/p.png"/>"#,
            r#"draw:fill="bitmap" draw:fill-image-name="P" style:repeat="stretch""#,
        );
        let stretched = Picture {
            path: "Pictures/p.png".into(),
            layout: PictureLayout::Stretch,
            size: (None, None),
            align: (0.5, 0.5),
            offset: (0.0, 0.0),
            shift: None,
            opacity: 1.0,
        };
        assert_eq!(picture, Fill::Picture(stretched.clone()));
        // §20.341: without `style:repeat` the picture repeats (LibreOffice's
        // default), and `no-repeat` draws it once, not stretched.
        let tiled = resolve(
            r#"<draw:fill-image draw:name="P" xlink:href="Pictures/p.png"/>"#,
            r#"draw:fill="bitmap" draw:fill-image-name="P" draw:fill-image-width="1cm" draw:fill-image-height="25%" draw:fill-image-ref-point-x="25%" draw:tile-repeat-offset="50% horizontal""#,
        );
        assert_eq!(
            tiled,
            Fill::Picture(Picture {
                layout: PictureLayout::Tile,
                size: (
                    length_pt("1cm").map(PictureSize::Pt),
                    Some(PictureSize::Fraction(0.25))
                ),
                offset: (0.25, 0.0),
                shift: Some(TileShift::Rows(0.5)),
                ..stretched.clone()
            })
        );
        // A single picture keeps its reference point (as LibreOffice
        // places it) but is neither moved nor staggered.
        let once = resolve(
            r#"<draw:fill-image draw:name="P" xlink:href="Pictures/p.png"/>"#,
            r#"draw:fill="bitmap" draw:fill-image-name="P" style:repeat="no-repeat" draw:fill-image-ref-point="top-right" draw:fill-image-ref-point-x="25%" draw:tile-repeat-offset="50% vertical""#,
        );
        assert_eq!(
            once,
            Fill::Picture(Picture {
                layout: PictureLayout::Once,
                align: (1.0, 0.0),
                ..stretched
            })
        );
    }

    #[test]
    fn opacity_is_part_of_the_fill() {
        let stops = |fill: Fill| match fill {
            Fill::Gradient(g) => g
                .stops
                .iter()
                .map(|s| (s.position, s.color.clone()))
                .collect::<Vec<_>>(),
            other => panic!("{other:?}"),
        };
        // §20.152: a uniform opacity is the alpha of every colour.
        assert_eq!(
            resolve(
                "",
                r##"draw:fill="solid" draw:fill-color="#ff0000" draw:opacity="50%""##
            ),
            Fill::Solid("FF000080".into())
        );
        assert_eq!(
            resolve(
                "",
                r##"draw:fill="solid" draw:fill-color="#ff0000" draw:opacity="100%""##
            ),
            Fill::Solid("FF0000".into())
        );
        let gradient = r##"<draw:gradient draw:name="G" draw:style="linear" draw:start-color="#ff0000" draw:end-color="#0000ff"/>"##;
        assert_eq!(
            stops(resolve(
                gradient,
                r#"draw:fill="gradient" draw:fill-gradient-name="G" draw:opacity="20%""#
            )),
            [(0.0, "FF000033".into()), (1.0, "0000FF33".into())]
        );
        // §20.153: an opacity gradient replaces the uniform opacity. Over a
        // solid colour it is a gradient of that colour.
        let fade = r##"<draw:opacity draw:name="T" draw:style="linear" draw:start="100%" draw:end="0%" draw:border="50%"/>"##;
        assert_eq!(
            stops(resolve(
                fade,
                r##"draw:fill="solid" draw:fill-color="#00ff00" draw:opacity="30%" draw:opacity-name="T""##
            )),
            [
                (0.0, "00FF00".into()),
                (0.5, "00FF00".into()),
                (1.0, "00FF0000".into())
            ]
        );
        // Along the same axis as a colour gradient, both are in the stops.
        assert_eq!(
            stops(resolve(
                &format!("{gradient}{fade}"),
                r#"draw:fill="gradient" draw:fill-gradient-name="G" draw:opacity-name="T""#
            )),
            [
                (0.0, "FF0000".into()),
                (0.0, "FF0000".into()),
                (0.5, "800080".into()),
                (1.0, "0000FF00".into()),
                (1.0, "0000FF00".into())
            ]
        );
        // Across it, the area takes the mean opacity (here 75%).
        let across = r##"<draw:opacity draw:name="T" draw:style="linear" draw:start="100%" draw:end="0%" draw:border="50%" draw:angle="90deg"/>"##;
        assert_eq!(
            stops(resolve(
                &format!("{gradient}{across}"),
                r#"draw:fill="gradient" draw:fill-gradient-name="G" draw:opacity-name="T""#
            )),
            [(0.0, "FF0000BF".into()), (1.0, "0000FFBF".into())]
        );
        let hatch = resolve(
            r##"<draw:hatch draw:name="H" draw:color="#000000" draw:distance="0.1cm"/>"##,
            r#"draw:fill="hatch" draw:fill-hatch-name="H" draw:opacity="40%""#,
        );
        assert!(matches!(hatch, Fill::Hatch(Hatch { opacity, .. }) if opacity == 0.4));
    }

    #[test]
    fn picture_references_resolve_as_package_iris() {
        let encoded = resolve(
            r#"<draw:fill-image draw:name="P" xlink:href="./Pictures/a%20b.png"/>"#,
            r#"draw:fill="bitmap" draw:fill-image-name="P" style:repeat="stretch""#,
        );
        assert!(matches!(
            encoded,
            Fill::Picture(Picture { path, .. }) if path == "Pictures/a b.png"
        ));
        // Outside the package: no picture, the solid colour applies.
        let external = resolve(
            r#"<draw:fill-image draw:name="P" xlink:href="../p.png"/>"#,
            r##"draw:fill="bitmap" draw:fill-image-name="P" draw:fill-color="#00ff00""##,
        );
        assert_eq!(external, Fill::Solid("00FF00".into()));
    }

    #[test]
    fn rectangle_spellings_and_namespaced_stops() {
        for style in ["rectangle", "rectangular"] {
            let fill = resolve(
                &format!(
                    r##"<draw:gradient draw:name="R" draw:style="{style}" draw:start-color="#000000" draw:end-color="#ffffff"/>"##
                ),
                r#"draw:fill="gradient" draw:fill-gradient-name="R""#,
            );
            let Fill::Gradient(g) = fill else { panic!() };
            assert!(
                matches!(g.shape, GradientShape::Path { path: "rect", .. }),
                "{style}"
            );
        }
        // Stops count only in their own namespaces (`loext:gradient-stop`,
        // `svg:offset`, `loext:color-value`).
        let stops = resolve(
            &format!(
                r##"<draw:gradient draw:name="S" draw:style="linear" xmlns:loext="{NS_LOEXT}" xmlns:x="urn:x"><loext:gradient-stop svg:offset="0" loext:color-value="#ff0000"/><x:gradient-stop svg:offset="0.5" loext:color-value="#00ff00"/><loext:gradient-stop svg:offset="1" loext:color-value="#0000ff"/></draw:gradient>"##
            ),
            r#"draw:fill="gradient" draw:fill-gradient-name="S""#,
        );
        let Fill::Gradient(g) = stops else { panic!() };
        assert_eq!(
            g.stops.iter().map(|s| s.color.as_str()).collect::<Vec<_>>(),
            ["FF0000", "0000FF"]
        );
    }

    #[test]
    fn too_many_gradient_stops_reject() {
        let stops: String = (0..=MAX_GRADIENT_STOPS)
            .map(|_| r##"<loext:gradient-stop svg:offset="0" loext:color-value="#ff0000"/>"##)
            .collect();
        let xml = format!(
            r##"<office:styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="{NS_DRAW}" xmlns:svg="{NS_SVG}" xmlns:loext="{NS_LOEXT}"><draw:gradient draw:name="G">{stops}</draw:gradient></office:styles>"##
        );
        let document = roxmltree::Document::parse(&xml).unwrap();
        assert_eq!(
            FillDefs::default().add(document.root_element(), "styles.xml"),
            Err(crate::error::resource("gradient stop"))
        );
    }
}
