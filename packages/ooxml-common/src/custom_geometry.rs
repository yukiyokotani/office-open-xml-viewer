//! Format-agnostic DrawingML custom-geometry grammar.
//!
//! ECMA-376 Part 1 §20.1.9 embeds the same `a:custGeom` vocabulary in Word,
//! Excel, and PowerPoint. This module resolves its ordered guide program and
//! path commands once; format parsers only adapt the result to their existing
//! wire models.

use std::collections::HashMap;

const FULL_CIRCLE: f64 = 21_600_000.0;
const ANGLE_TO_RAD: f64 = std::f64::consts::TAU / FULL_CIRCLE;

#[derive(Debug, Clone, PartialEq)]
pub struct CustomGeometry {
    pub paths: Vec<CustomPath>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct CustomPath {
    /// Effective path coordinate-system width. Omitted/zero `a:path@w` uses
    /// the containing shape extent as required by the host shape space.
    pub width: f64,
    /// Effective path coordinate-system height.
    pub height: f64,
    /// `a:path@fill` (ECMA-376 20.1.9.15, ST_PathFillMode 20.1.10.37):
    /// `none`, `lighten`, `lightenLess`, `darken` or `darkenLess`. `None` is
    /// the default `norm` (and any unrecognized value).
    pub fill: Option<String>,
    /// `a:path@stroke`, default true.
    pub stroke: bool,
    pub commands: Vec<PathCommand>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum PathCommand {
    MoveTo {
        x: f64,
        y: f64,
    },
    LineTo {
        x: f64,
        y: f64,
    },
    CubicBezierTo {
        x1: f64,
        y1: f64,
        x2: f64,
        y2: f64,
        x: f64,
        y: f64,
    },
    QuadraticBezierTo {
        x1: f64,
        y1: f64,
        x: f64,
        y: f64,
    },
    /// Radii are in path coordinates; angles are DrawingML 60000ths of a degree.
    ArcTo {
        wr: f64,
        hr: f64,
        st_ang: f64,
        sw_ang: f64,
    },
    Close,
}

fn child<'a, 'input>(
    node: roxmltree::Node<'a, 'input>,
    name: &str,
) -> Option<roxmltree::Node<'a, 'input>> {
    node.children()
        .find(|child| child.is_element() && child.tag_name().name() == name)
}

fn attr_f64(node: roxmltree::Node<'_, '_>, name: &str) -> Option<f64> {
    let value = node.attribute(name)?.parse::<f64>().ok()?;
    value.is_finite().then_some(value)
}

fn resolve(token: &str, env: &HashMap<String, f64>) -> Option<f64> {
    env.get(token)
        .copied()
        .or_else(|| token.parse::<f64>().ok())
        .filter(|value| value.is_finite())
}

fn evaluate_formula(formula: &str, env: &HashMap<String, f64>) -> Option<f64> {
    let mut tokens = formula.split_whitespace();
    let op = tokens.next()?;
    let args: Vec<f64> = tokens
        .map(|token| resolve(token, env))
        .collect::<Option<_>>()?;
    let arg = |index: usize| args.get(index).copied();
    let value = match op {
        "val" => arg(0)?,
        "*/" => arg(0)? * arg(1)? / arg(2)?,
        "+-" => arg(0)? + arg(1)? - arg(2)?,
        "+/" => (arg(0)? + arg(1)?) / arg(2)?,
        "?:" => {
            if arg(0)? > 0.0 {
                arg(1)?
            } else {
                arg(2)?
            }
        }
        "abs" => arg(0)?.abs(),
        "at2" => arg(1)?.atan2(arg(0)?) / ANGLE_TO_RAD,
        "cat2" => arg(0)? * arg(2)?.atan2(arg(1)?).cos(),
        "cos" => arg(0)? * (arg(1)? * ANGLE_TO_RAD).cos(),
        "max" => arg(0)?.max(arg(1)?),
        "min" => arg(0)?.min(arg(1)?),
        "mod" => (arg(0)?.powi(2) + arg(1)?.powi(2) + arg(2)?.powi(2)).sqrt(),
        "pin" => {
            let (low, value, high) = (arg(0)?, arg(1)?, arg(2)?);
            if value < low {
                low
            } else if value > high {
                high
            } else {
                value
            }
        }
        "sat2" => arg(0)? * arg(2)?.atan2(arg(1)?).sin(),
        "sin" => arg(0)? * (arg(1)? * ANGLE_TO_RAD).sin(),
        "sqrt" => arg(0)?.max(0.0).sqrt(),
        "tan" => arg(0)? * (arg(1)? * ANGLE_TO_RAD).tan(),
        _ => return None,
    };
    value.is_finite().then_some(value)
}

fn guide_environment(
    cust_geom: roxmltree::Node<'_, '_>,
    shape_width: f64,
    shape_height: f64,
) -> HashMap<String, f64> {
    let first_path = child(cust_geom, "pathLst").and_then(|paths| {
        paths
            .children()
            .find(|node| node.is_element() && node.tag_name().name() == "path")
    });
    let width = if shape_width > 0.0 {
        shape_width
    } else {
        first_path
            .and_then(|path| attr_f64(path, "w"))
            .unwrap_or(1.0)
            .max(1.0)
    };
    let height = if shape_height > 0.0 {
        shape_height
    } else {
        first_path
            .and_then(|path| attr_f64(path, "h"))
            .unwrap_or(1.0)
            .max(1.0)
    };
    let short_side = width.min(height);
    let long_side = width.max(height);
    let mut env = HashMap::new();

    macro_rules! builtins {
        ($($name:expr => $value:expr),* $(,)?) => { $(env.insert($name.to_owned(), $value);)* };
    }
    builtins! {
        "w" => width, "h" => height,
        "l" => 0.0, "t" => 0.0, "r" => width, "b" => height,
        "hc" => width / 2.0, "vc" => height / 2.0,
        "wd2" => width / 2.0, "wd3" => width / 3.0, "wd4" => width / 4.0,
        "wd5" => width / 5.0, "wd6" => width / 6.0, "wd8" => width / 8.0,
        "wd10" => width / 10.0, "wd12" => width / 12.0,
        "wd16" => width / 16.0, "wd32" => width / 32.0,
        "hd2" => height / 2.0, "hd3" => height / 3.0, "hd4" => height / 4.0,
        "hd5" => height / 5.0, "hd6" => height / 6.0, "hd8" => height / 8.0,
        "hd10" => height / 10.0, "hd12" => height / 12.0,
        "hd16" => height / 16.0, "hd32" => height / 32.0,
        "ss" => short_side, "ssd2" => short_side / 2.0, "ssd4" => short_side / 4.0,
        "ssd6" => short_side / 6.0, "ssd8" => short_side / 8.0,
        "ssd16" => short_side / 16.0, "ssd32" => short_side / 32.0,
        "ls" => long_side, "lsd2" => long_side / 2.0, "lsd4" => long_side / 4.0,
        "lsd6" => long_side / 6.0, "lsd8" => long_side / 8.0,
        "lsd16" => long_side / 16.0, "lsd32" => long_side / 32.0,
        "cd" => FULL_CIRCLE, "cd2" => FULL_CIRCLE / 2.0,
        "cd4" => FULL_CIRCLE / 4.0, "cd8" => FULL_CIRCLE / 8.0,
        "3cd4" => 3.0 * FULL_CIRCLE / 4.0, "3cd8" => 3.0 * FULL_CIRCLE / 8.0,
        "5cd8" => 5.0 * FULL_CIRCLE / 8.0, "7cd8" => 7.0 * FULL_CIRCLE / 8.0,
    }

    for list_name in ["avLst", "gdLst"] {
        let Some(list) = child(cust_geom, list_name) else {
            continue;
        };
        for guide in list
            .children()
            .filter(|node| node.is_element() && node.tag_name().name() == "gd")
        {
            let (Some(name), Some(formula)) = (guide.attribute("name"), guide.attribute("fmla"))
            else {
                continue;
            };
            if let Some(value) = evaluate_formula(formula, &env) {
                env.insert(name.to_owned(), value);
            }
        }
    }
    env
}

fn geometry_attr(
    node: roxmltree::Node<'_, '_>,
    name: &str,
    env: &HashMap<String, f64>,
) -> Option<f64> {
    resolve(node.attribute(name)?, env)
}

fn point(node: roxmltree::Node<'_, '_>, env: &HashMap<String, f64>) -> Option<(f64, f64)> {
    Some((
        geometry_attr(node, "x", env)?,
        geometry_attr(node, "y", env)?,
    ))
}

fn parse_command(node: roxmltree::Node<'_, '_>, env: &HashMap<String, f64>) -> Option<PathCommand> {
    let points: Vec<_> = node
        .children()
        .filter(|child| child.is_element() && child.tag_name().name() == "pt")
        .map(|child| point(child, env))
        .collect::<Option<_>>()?;
    match node.tag_name().name() {
        "moveTo" => points.first().map(|&(x, y)| PathCommand::MoveTo { x, y }),
        "lnTo" => points.first().map(|&(x, y)| PathCommand::LineTo { x, y }),
        "cubicBezTo" if points.len() >= 3 => Some(PathCommand::CubicBezierTo {
            x1: points[0].0,
            y1: points[0].1,
            x2: points[1].0,
            y2: points[1].1,
            x: points[2].0,
            y: points[2].1,
        }),
        "quadBezTo" if points.len() >= 2 => Some(PathCommand::QuadraticBezierTo {
            x1: points[0].0,
            y1: points[0].1,
            x: points[1].0,
            y: points[1].1,
        }),
        "arcTo" => Some(PathCommand::ArcTo {
            wr: geometry_attr(node, "wR", env).unwrap_or(0.0),
            hr: geometry_attr(node, "hR", env).unwrap_or(0.0),
            st_ang: geometry_attr(node, "stAng", env).unwrap_or(0.0),
            sw_ang: geometry_attr(node, "swAng", env).unwrap_or(0.0),
        }),
        "close" => Some(PathCommand::Close),
        _ => None,
    }
}

/// Parse `a:custGeom` according to ECMA-376 Part 1 §20.1.9.
pub fn parse_custom_geometry(
    cust_geom: roxmltree::Node<'_, '_>,
    shape_width: f64,
    shape_height: f64,
) -> CustomGeometry {
    let Some(path_list) = child(cust_geom, "pathLst") else {
        return CustomGeometry { paths: Vec::new() };
    };
    let env = guide_environment(cust_geom, shape_width, shape_height);
    let paths = path_list
        .children()
        .filter(|node| node.is_element() && node.tag_name().name() == "path")
        .map(|path| {
            let width = attr_f64(path, "w")
                .filter(|value| *value > 0.0)
                .unwrap_or(shape_width.max(1.0));
            let height = attr_f64(path, "h")
                .filter(|value| *value > 0.0)
                .unwrap_or(shape_height.max(1.0));
            let commands = path
                .children()
                .filter(|node| node.is_element())
                .filter_map(|node| parse_command(node, &env))
                .collect();
            let fill = path
                .attribute("fill")
                .filter(|mode| {
                    matches!(
                        *mode,
                        "none" | "lighten" | "lightenLess" | "darken" | "darkenLess"
                    )
                })
                .map(str::to_owned);
            // xsd:boolean
            let stroke = !matches!(path.attribute("stroke"), Some("0" | "false"));
            CustomPath {
                width,
                height,
                fill,
                stroke,
                commands,
            }
        })
        .collect();
    CustomGeometry { paths }
}

#[cfg(test)]
mod tests {
    use super::{parse_custom_geometry, PathCommand};

    #[test]
    fn resolves_ordered_guides_and_quadratic_beziers() {
        let xml = r#"<a:custGeom xmlns:a="urn:a">
          <a:gdLst>
            <a:gd name="quarter" fmla="*/ w 1 4"/>
            <a:gd name="threeQuarter" fmla="+- w 0 quarter"/>
          </a:gdLst>
          <a:pathLst><a:path w="100" h="200">
            <a:moveTo><a:pt x="quarter" y="0"/></a:moveTo>
            <a:quadBezTo><a:pt x="hc" y="h"/><a:pt x="threeQuarter" y="0"/></a:quadBezTo>
          </a:path></a:pathLst>
        </a:custGeom>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let geometry = parse_custom_geometry(doc.root_element(), 100.0, 200.0);
        assert_eq!(
            geometry.paths[0].commands,
            vec![
                PathCommand::MoveTo { x: 25.0, y: 0.0 },
                PathCommand::QuadraticBezierTo {
                    x1: 50.0,
                    y1: 200.0,
                    x: 75.0,
                    y: 0.0,
                },
            ]
        );
    }

    #[test]
    fn path_fill_mode_and_stroke_flags_are_retained() {
        let xml = r#"<custGeom><pathLst>
          <path w="1" h="1"/>
          <path w="1" h="1" fill="none" stroke="0"/>
          <path w="1" h="1" fill="darkenLess" stroke="true"/>
          <path w="1" h="1" fill="norm" stroke="false"/>
          <path w="1" h="1" fill="bogus"/>
        </pathLst></custGeom>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let geometry = parse_custom_geometry(doc.root_element(), 1.0, 1.0);
        let flags: Vec<_> = geometry
            .paths
            .iter()
            .map(|p| (p.fill.as_deref(), p.stroke))
            .collect();
        assert_eq!(
            flags,
            [
                (None, true),
                (Some("none"), false),
                (Some("darkenLess"), true),
                (None, false),
                (None, true),
            ]
        );
    }

    #[test]
    fn omitted_path_size_uses_shape_coordinate_space() {
        let xml = r#"<custGeom><pathLst><path>
          <moveTo><pt x="0" y="0"/></moveTo><lnTo><pt x="r" y="b"/></lnTo>
        </path></pathLst></custGeom>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let geometry = parse_custom_geometry(doc.root_element(), 300.0, 150.0);
        assert_eq!(geometry.paths[0].width, 300.0);
        assert_eq!(geometry.paths[0].height, 150.0);
        assert_eq!(
            geometry.paths[0].commands[1],
            PathCommand::LineTo { x: 300.0, y: 150.0 }
        );
    }
}
