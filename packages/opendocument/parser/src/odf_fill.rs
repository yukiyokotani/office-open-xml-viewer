//! Area fills of drawing styles (ODF Part 3 §20.130 `draw:fill`): solid
//! colours, gradients (§16.40 `draw:gradient`), hatches (§16.42
//! `draw:hatch`) and pictures (§16.41 `draw:fill-image`), resolved from a
//! graphic style chain into a format-neutral description. The spreadsheet
//! reader maps it to the workbook drawing model; the presentation reader can
//! map it to its own.
//!
//! Gradients are expressed as DrawingML gradients (ECMA-376 §20.1.8.33
//! `a:gradFill`): stop lists (`ooxml_common::fill::GradStop`), a linear angle
//! (§20.1.8.41, degrees clockwise from left-to-right) or a path shade
//! (§20.1.8.46) with its focus rectangle (§20.1.8.30 `a:fillToRect`).

use crate::odf_props::{get, hex_color, length_pt, percent, Style};
use crate::xml::{self, NS_DRAW, NS_XLINK};
use ooxml_common::fill::{FillRect, GradStop};
use std::collections::HashMap;

/// Implementation bound on the stops of one gradient.
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
    /// colours `RRGGBB`.
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
    /// Degrees counter-clockwise of the first line family.
    pub angle: f64,
    /// Distance between lines, in points.
    pub distance_pt: f64,
    /// 1 (single), 2 (double), 3 (triple).
    pub lines: u8,
}

/// §16.41 `draw:fill-image` with the §20.240 `style:repeat` mode.
#[derive(Clone, Debug, PartialEq)]
pub struct Picture {
    /// Package path of the image.
    pub path: String,
    /// `style:repeat="repeat"`: the tile size in points, if authored.
    pub tile: Option<(Option<f64>, Option<f64>)>,
}

#[derive(Clone, Debug, PartialEq)]
pub enum Fill {
    None,
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
    /// Read the definitions below `office:styles`.
    pub fn add(&mut self, styles: roxmltree::Node<'_, '_>) {
        for node in styles.children().filter(|n| n.is_element()) {
            let Some(name) = node.attribute((NS_DRAW, "name")) else {
                continue;
            };
            let attr = |key: &str| node.attribute((NS_DRAW, key));
            if xml::is(node, NS_DRAW, "gradient") {
                let stops = node
                    .children()
                    .filter(|n| n.is_element() && n.tag_name().name() == "gradient-stop")
                    .take(MAX_GRADIENT_STOPS)
                    .filter_map(|stop| {
                        let offset = stop
                            .attributes()
                            .find(|a| a.name() == "offset")?
                            .value()
                            .parse::<f64>()
                            .ok()?;
                        let color = stop
                            .attributes()
                            .find(|a| a.name() == "color-value")
                            .and_then(|a| hex_color(a.value()))?;
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
                if let Some(href) = node.attribute((NS_XLINK, "href")) {
                    self.images.insert(name.to_string(), href.to_string());
                }
            }
        }
    }

    /// The fill of a graphic style chain (`graphic:` properties).
    pub fn resolve(&self, chain: &[&Style]) -> Fill {
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
                        })
                    },
                ),
            Some("bitmap") => property("fill-image-name")
                .and_then(|name| self.images.get(name))
                .and_then(|href| {
                    let path = href.trim_start_matches("./");
                    (!path.contains("://") && !path.contains("..") && !path.is_empty()).then(|| {
                        let repeat = get(chain, "graphic:style:repeat").unwrap_or("repeat");
                        Fill::Picture(Picture {
                            path: path.to_string(),
                            tile: (repeat == "repeat").then(|| {
                                (
                                    property("fill-image-width")
                                        .and_then(length_pt)
                                        .filter(|v| *v > 0.0),
                                    property("fill-image-height")
                                        .and_then(length_pt)
                                        .filter(|v| *v > 0.0),
                                )
                            }),
                        })
                    })
                })
                .unwrap_or_else(|| solid().map_or(Fill::None, Fill::Solid)),
            _ => solid().map_or(Fill::None, Fill::Solid),
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
    let mut colours: Vec<(f64, String)> = if def.stops.len() >= 2 {
        def.stops.clone()
    } else {
        vec![(0.0, start), (1.0, end)]
    };
    colours.sort_by(|a, b| a.0.total_cmp(&b.0));
    let border = def.border;
    // Map an offset in [0, 1] from the start side, after the border.
    let along = |t: f64| border + (1.0 - border) * t;
    match def.style.as_str() {
        "linear" => {
            let mut stops = Vec::new();
            if border > 0.0 {
                stops.push(stop(0.0, &colours[0].1));
            }
            stops.extend(colours.iter().map(|(t, c)| stop(along(*t), c)));
            Gradient {
                stops,
                // Start at the top (DrawingML 90°), rotated counter-clockwise.
                shape: GradientShape::Linear {
                    angle: (90.0 - def.angle).rem_euclid(360.0),
                },
                rotation: 0.0,
            }
        }
        "axial" => {
            let half = |t: f64| along(t) / 2.0;
            let mut stops = Vec::new();
            if border > 0.0 {
                stops.push(stop(0.0, &colours[0].1));
            }
            stops.extend(colours.iter().map(|(t, c)| stop(half(*t), c)));
            stops.extend(colours.iter().rev().map(|(t, c)| stop(1.0 - half(*t), c)));
            if border > 0.0 {
                stops.push(stop(1.0, &colours[0].1));
            }
            Gradient {
                stops,
                shape: GradientShape::Linear {
                    angle: (90.0 - def.angle).rem_euclid(360.0),
                },
                rotation: 0.0,
            }
        }
        style => {
            // Path shades run from the focus (position 0) outwards; ODF's
            // start colour is outside, so the stops are reversed.
            let mut stops: Vec<GradStop> = colours
                .iter()
                .rev()
                .map(|(t, c)| stop((1.0 - along(*t)).max(0.0), c))
                .collect();
            if border > 0.0 {
                stops.push(stop(1.0, &colours[0].1));
            }
            stops.sort_by(|a, b| a.position.total_cmp(&b.position));
            Gradient {
                stops,
                shape: GradientShape::Path {
                    path: if matches!(style, "square" | "rectangular") {
                        "rect"
                    } else {
                        "circle"
                    },
                    focus: FillRect {
                        l: def.cx,
                        t: def.cy,
                        r: 1.0 - def.cx,
                        b: 1.0 - def.cy,
                    },
                },
                rotation: def.angle,
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
        fills.add(document.root_element());
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
                lines: 2
            })
        );
        let picture = resolve(
            r#"<draw:fill-image draw:name="P" xlink:href="Pictures/p.png"/>"#,
            r#"draw:fill="bitmap" draw:fill-image-name="P" style:repeat="stretch""#,
        );
        assert_eq!(
            picture,
            Fill::Picture(Picture {
                path: "Pictures/p.png".into(),
                tile: None
            })
        );
        let tiled = resolve(
            r#"<draw:fill-image draw:name="P" xlink:href="Pictures/p.png"/>"#,
            r#"draw:fill="bitmap" draw:fill-image-name="P" draw:fill-image-width="1cm""#,
        );
        assert!(matches!(
            tiled,
            Fill::Picture(Picture {
                tile: Some((Some(_), None)),
                ..
            })
        ));
    }
}
