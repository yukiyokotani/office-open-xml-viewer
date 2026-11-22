//! ECMA-376 Part 4 §19.1.2.5: VML brushes acquired into retained fills.
//! The shape instance overrides the template per property, not per element.
use super::*;

pub(super) fn property<'a>(
    shape: roxmltree::Node<'a, '_>,
    definition: Option<roxmltree::Node<'a, '_>>,
    name: &str,
) -> Option<&'a str> {
    let read = |host| {
        vml_direct_child(host, "fill")
            .and_then(|n| n.attributes().find(|a| a.name() == name).map(|a| a.value()))
            // AG_ShapeAttributes exposes the same primary-fill opacity;
            // a direct decoration overrides its host before inheritance.
            .or_else(|| (name == "opacity").then(|| host.attribute(name)).flatten())
    };
    read(shape).or_else(|| definition.and_then(read))
}
fn fraction(raw: &str) -> Option<f64> {
    raw.trim()
        .strip_suffix('%')
        .and_then(|s| s.parse::<f64>().ok())
        .map(|v| v / 100.0)
        .or_else(|| parse_vml_opacity(raw.trim()))
        .filter(|v| v.is_finite())
}
fn vector(raw: &str) -> Option<(f64, f64)> {
    let (x, y) = raw.split_once(',')?;
    Some((fraction(x)?, fraction(y)?))
}

pub(super) fn enabled(shape: roxmltree::Node, definition: Option<roxmltree::Node>) -> bool {
    let layer = |host| {
        vml_direct_child(host, "fill")
            .and_then(|n| n.attribute("on"))
            .or_else(|| host.attribute("filled"))
            .and_then(parse_vml_true_false)
    };
    layer(shape)
        .or_else(|| definition.and_then(layer))
        .unwrap_or(true)
}

pub(super) fn resolve(
    shape: roxmltree::Node,
    definition: Option<roxmltree::Node>,
    media: &HashMap<String, String>,
) -> ResolvedVmlFill {
    let attr = |name| property(shape, definition, name);
    let enabled = enabled(shape, definition);
    let opacity = attr("opacity").and_then(fraction);
    if !enabled {
        return ResolvedVmlFill {
            fill: None,
            opacity: None,
            incomplete: false,
        };
    }
    let color = vml_direct_child(shape, "fill")
        .and_then(|n| n.attribute("color"))
        .or_else(|| shape.attribute("fillcolor"))
        .or_else(|| {
            definition
                .and_then(|n| vml_direct_child(n, "fill"))
                .and_then(|n| n.attribute("color"))
        })
        .or_else(|| definition.and_then(|n| n.attribute("fillcolor")))
        .and_then(vml_color_hex6)
        .unwrap_or_else(|| "FFFFFF".into());
    let color2 = attr("color2")
        .and_then(vml_color_hex6)
        .unwrap_or_else(|| "FFFFFF".into());
    let kind = attr("type").unwrap_or("solid");
    let fill = match kind {
        "solid" => Some(ShapeFill::Solid { color }),
        "gradient" | "gradientradial" => {
            // Endpoint opacity belongs to each stop; it must not additionally
            // attenuate the whole brush and thereby multiply alpha twice.
            let primary = vml_color_alpha(color, opacity);
            let secondary = vml_color_alpha(color2, attr("opacity2").and_then(fraction));
            let mut stops = vec![GradientStop {
                position: 0.0,
                color: primary.clone(),
            }];
            if let Some(colors) = attr("colors") {
                for pair in colors.split([',', ';']) {
                    if let Some((position, color)) = pair.trim().split_once(char::is_whitespace) {
                        if let (Some(position), Some(color)) =
                            (fraction(position), vml_color_hex6(color.trim()))
                        {
                            stops.push(GradientStop {
                                position: position.clamp(0.0, 1.0),
                                color: vml_color_alpha(
                                    color,
                                    Some({
                                        let primary = opacity.unwrap_or(1.0);
                                        primary
                                            + (attr("opacity2").and_then(fraction).unwrap_or(1.0)
                                                - primary)
                                                * position
                                    }),
                                ),
                            });
                        }
                    }
                }
            }
            stops.push(GradientStop {
                position: 1.0,
                color: secondary,
            });
            stops.sort_by(|a, b| a.position.total_cmp(&b.position));
            // Part 4 ST_FillMethod names sigma/linear-sigma without a numeric
            // transfer function. Library policy uses the shared stop-linear
            // interpolation; this is not an Office sigma compatibility claim.
            let focus = attr("focus")
                .and_then(fraction)
                .unwrap_or(0.0)
                .clamp(-1.0, 1.0);
            if kind == "gradient" && focus != 0.0 {
                let mut focused = Vec::new();
                let left = if focus > 0.0 { 1.0 - focus } else { -focus };
                let right = 1.0 - left;
                if left > 0.0 {
                    for stop in &stops {
                        focused.push(GradientStop {
                            position: if focus > 0.0 {
                                stop.position * left
                            } else {
                                (1.0 - stop.position) * left
                            },
                            color: stop.color.clone(),
                        });
                    }
                }
                if right > 0.0 {
                    for stop in &stops {
                        focused.push(GradientStop {
                            position: if focus > 0.0 {
                                1.0 - stop.position * right
                            } else {
                                left + stop.position * right
                            },
                            color: stop.color.clone(),
                        });
                    }
                }
                focused.sort_by(|a, b| a.position.total_cmp(&b.position));
                focused.dedup_by(|a, b| a.position == b.position && a.color == b.color);
                stops = focused;
            }
            let radial = kind == "gradientradial";
            let (x, y) = attr("focusposition").and_then(vector).unwrap_or((0.0, 0.0));
            let (w, h) = attr("focussize").and_then(vector).unwrap_or((0.0, 0.0));
            // VML primary is the outside of the rectangular radial fill;
            // DrawingML starts at the focus rectangle and travels outward.
            if radial {
                for stop in &mut stops {
                    stop.position = 1.0 - stop.position;
                }
                stops.reverse();
            }
            Some(ShapeFill::Gradient {
                stops,
                // VML's zero gradient runs bottom -> top; positive angles
                // rotate counterclockwise (Part 4 color2/angle examples).
                angle: (270.0
                    - attr("angle")
                        .and_then(|s| s.parse::<f64>().ok())
                        .unwrap_or(0.0))
                .rem_euclid(360.0),
                grad_type: if radial { "radial" } else { "linear" }.into(),
                scaled: Some(false),
                // Part 4 §19.1.2.14: gradientshapeok selects concentric
                // copies of the actual path, rather than its bounding rectangle.
                path: radial.then(|| {
                    if vml_geometry::path_attr(shape, definition, "gradientshapeok")
                        .and_then(parse_vml_true_false)
                        == Some(true)
                    {
                        "shape"
                    } else {
                        "rect"
                    }
                    .into()
                }),
                fill_to_rect: radial.then_some(ooxml_common::fill::FillRect {
                    l: x,
                    t: y,
                    r: 1.0 - x - w,
                    b: 1.0 - y - h,
                }),
                tile_rect: None,
                flip: None,
                rot_with_shape: Some(
                    attr("rotate")
                        .and_then(parse_vml_true_false)
                        .unwrap_or(false),
                ),
            })
        }
        "tile" | "pattern" | "frame" => {
            let source = vml_direct_child(shape, "fill")
                .filter(|n| n.attribute("src").is_some() || vml_image_relationship(*n).is_some())
                .or_else(|| definition.and_then(|n| vml_direct_child(n, "fill")));
            source.and_then(|n| vml_image_source(n, media)).map(|path| {
                let size = attr("size")
                    .and_then(|s| s.split_once(','))
                    .and_then(|(x, y)| Some((length(x)?, length(y)?)));
                let origin = attr("origin").and_then(vector).unwrap_or((0.5, 0.5));
                let position = attr("position").and_then(vector).unwrap_or((0.5, 0.5));
                ShapeFill::Image {
                    mime_type: mime_from_ext(&path).into(),
                    image_path: path,
                    svg_image_path: None,
                    src_rect: None,
                    fill_rect: None,
                    tile: None,
                    alpha: opacity,
                    // A pattern recolours the image by its gray level: Windows
                    // Word 16.0.20430 paints black pixels color2 and white
                    // pixels color, blending grays linearly (#1896 controls:
                    // checker, dots, gray ramp and two-gray 8x8 images). As a
                    // duotone (dark -> clr1, light -> clr2) that is
                    // clr1 = color2, clr2 = color.
                    duotone: (kind == "pattern").then_some(Duotone {
                        clr1: color2,
                        clr2: color,
                    }),
                    blip_effects: Vec::new(),
                    dpi: None,
                    vml: Some(VmlImageFill {
                        repeat: kind != "frame",
                        size,
                        origin,
                        position,
                        aspect: attr("aspect").unwrap_or("ignore").into(),
                        align_shape: attr("alignshape")
                            .and_then(parse_vml_true_false)
                            .unwrap_or(true),
                    }),
                }
            })
        }
        _ => None,
    };
    ResolvedVmlFill {
        incomplete: fill.is_none(),
        fill,
        opacity: if matches!(
            kind,
            "gradient" | "gradientradial" | "tile" | "pattern" | "frame"
        ) {
            None
        } else {
            opacity
        },
    }
}
pub(super) fn length(raw: &str) -> Option<VmlFillLength> {
    let raw = raw.trim();
    let is_fraction =
        raw.ends_with('%') || raw.ends_with('f') || !raw.chars().any(|c| c.is_ascii_alphabetic());
    Some(VmlFillLength {
        value: if is_fraction {
            fraction(raw)?
        } else {
            parse_vml_stroke_weight_pt(raw)?
        },
        fraction: is_fraction,
    })
}

/// Part 4 §19.1.2.21: stroke textures use the same decoded brush authority
/// as shape fills, with their own per-property inheritance and image frame.
pub(super) fn stroke_image(
    shape: roxmltree::Node,
    definition: Option<roxmltree::Node>,
    media: &HashMap<String, String>,
) -> Option<ShapeFill> {
    let attr = |name| vml_decoration_attr(shape, definition, "stroke", name);
    let source = vml_direct_child(shape, "stroke")
        .filter(|n| n.attribute("src").is_some() || vml_image_relationship(*n).is_some())
        .or_else(|| definition.and_then(|n| vml_direct_child(n, "stroke")));
    pen_image(attr, source, media)
}

pub(super) fn pen_image<'a>(
    attr: impl Fn(&'static str) -> Option<&'a str>,
    source: Option<roxmltree::Node>,
    media: &HashMap<String, String>,
) -> Option<ShapeFill> {
    let kind = attr("filltype").unwrap_or("solid");
    if !matches!(kind, "tile" | "pattern" | "frame") {
        return None;
    }
    let path = vml_image_source(source?, media)?;
    Some(ShapeFill::Image {
        mime_type: mime_from_ext(&path).into(),
        image_path: path,
        svg_image_path: None,
        src_rect: None,
        fill_rect: None,
        tile: None,
        alpha: attr("opacity").and_then(fraction),
        duotone: None,
        blip_effects: Vec::new(),
        dpi: None,
        vml: Some(VmlImageFill {
            repeat: kind == "tile",
            size: attr("imagesize")
                .and_then(|s| s.split_once(','))
                .and_then(|(x, y)| Some((length(x)?, length(y)?))),
            origin: (0.0, 0.0),
            position: (0.0, 0.0),
            aspect: attr("imageaspect").unwrap_or("ignore").into(),
            align_shape: attr("imagealignshape")
                .and_then(parse_vml_true_false)
                .unwrap_or(true),
        }),
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    fn fill(xml: &str) -> serde_json::Value {
        let xml = format!(
            r##"<root xmlns:v="urn:schemas-microsoft-com:vml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:o="urn:schemas-microsoft-com:office:office">{xml}</root>"##
        );
        let doc = XmlDoc::parse(&xml).unwrap();
        let definition = doc.descendants().find(|n| n.has_tag_name("type"));
        let shape = doc.descendants().find(|n| n.has_tag_name("shape")).unwrap();
        serde_json::to_value(
            resolve(
                shape,
                definition,
                &HashMap::from([("pic".into(), "word/media/a.png".into())]),
            )
            .fill,
        )
        .unwrap()
    }
    #[test]
    fn fill_properties_cascade_individually_with_stop_opacity_and_focus() {
        let fill = fill(
            r##"<type><v:fill type="gradient" color="red" color2="blue" angle="90" opacity="0" o:opacity2=".5"/></type><shape fillcolor="lime"><v:fill colors="32768f yellow" focus="100%"/></shape>"##,
        );
        assert_eq!(fill["fillType"], "gradient");
        assert_eq!(fill["angle"], 180.0);
        assert_eq!(fill["stops"][0]["color"], "0000ff80");
        assert_eq!(fill["stops"][2]["color"], "00ff0000");
        let secondary = self::fill(
            r#"<shape><v:fill type="gradient" color="red" color2="blue" o:opacity2=".5" colors="50% yellow"/></shape>"#,
        );
        assert_eq!(secondary["stops"][1]["color"], "ffff00bf");
    }
    #[test]
    fn radial_focus_and_inherited_image_resources_retain_geometry() {
        let radial = fill(
            r#"<shape><v:fill type="gradientradial" color="red" color2="blue" focusposition="25%,50%" focussize="10%,20%"/></shape>"#,
        );
        assert_eq!(radial["fillToRect"]["l"], 0.25);
        assert_eq!(radial["fillToRect"]["t"], 0.5);
        assert_eq!(radial["stops"][0]["color"], "0000ff");
        let image = fill(
            r#"<type><v:fill type="pattern" r:id="pic" color="red" color2="blue" size="20pt,25%"/></type><shape><v:fill id="metadata" position="25%,50%"/></shape>"#,
        );
        assert_eq!(image["imagePath"], "word/media/a.png");
        // Black pixels take color2, white pixels color.
        assert_eq!(image["duotone"]["clr1"], "0000ff");
        assert_eq!(image["duotone"]["clr2"], "ff0000");
        assert_eq!(image["vml"]["size"][0]["value"], 20.0);
        assert_eq!(image["vml"]["size"][1]["value"], 0.25);
        assert_eq!(image["vml"]["position"][0], 0.25);
    }
}
