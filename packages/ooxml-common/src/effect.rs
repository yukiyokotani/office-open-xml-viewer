//! Shared DrawingML effect-list parsing.
//!
//! Shape properties, chart styles, pictures, and text all embed the same
//! `CT_EffectList` grammar (ECMA-376 Part 1 §20.1.8.16). Keeping the five
//! effects currently representable by the Canvas model here prevents host
//! parsers from drifting on defaults, units, or colour resolution.

use crate::color::{parse_color_node, ThemeResolver, TintMode};
use roxmltree::Node;
use serde::{Deserialize, Serialize};

fn child<'a>(node: Node<'a, 'a>, name: &str) -> Option<Node<'a, 'a>> {
    node.children()
        .find(|candidate| candidate.is_element() && candidate.tag_name().name() == name)
}

fn attr_i64(node: Node<'_, '_>, name: &str) -> Option<i64> {
    node.attribute(name)?.parse().ok()
}

fn attr_f64(node: Node<'_, '_>, name: &str) -> Option<f64> {
    node.attribute(name)?.parse().ok()
}

/// ECMA-376 §20.1.8.45/40 outer or inner shadow.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct Shadow {
    pub color: String,
    pub alpha: f64,
    pub blur: i64,
    pub dist: i64,
    /// Direction in degrees clockwise from East.
    pub dir: f64,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sx: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sy: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub kx: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ky: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub algn: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rot_with_shape: Option<bool>,
}

/// ECMA-376 §20.1.8.17 coloured glow halo.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct Glow {
    pub color: String,
    pub alpha: f64,
    pub radius: i64,
}

/// ECMA-376 §20.1.8.31 feathered edge.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct SoftEdge {
    pub radius: i64,
}

/// ECMA-376 §20.1.8.27 mirrored reflection.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct Reflection {
    pub blur: i64,
    pub dist: i64,
    /// Direction in degrees clockwise from East.
    pub dir: f64,
    pub st_a: f64,
    pub st_pos: f64,
    pub end_a: f64,
    pub end_pos: f64,
    pub sx: f64,
    pub sy: f64,
}

/// The five independent `CT_EffectList` members represented by the shared
/// Canvas wire model. `unsupported` records other concrete effect children so
/// callers can suppress lower-precedence defaults instead of guessing.
#[derive(Serialize, Deserialize, Debug, Clone, Default, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct DrawingMlEffects {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shadow: Option<Shadow>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inner_shadow: Option<Shadow>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub glow: Option<Glow>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub soft_edge: Option<SoftEdge>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub reflection: Option<Reflection>,
    /// At least one authored effect child is outside the supported subset.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub unsupported: bool,
}

fn try_resolved_color<R: ThemeResolver + ?Sized>(
    node: Node<'_, '_>,
    resolver: &R,
    tint_mode: TintMode,
) -> Option<(String, f64)> {
    let color = parse_color_node(node, resolver, tint_mode)?;
    if color.len() >= 8 {
        let alpha = u8::from_str_radix(&color[6..8], 16).unwrap_or(255) as f64 / 255.0;
        Some((color[..6].to_owned(), alpha))
    } else {
        Some((color, 1.0))
    }
}

/// Parse one `<a:outerShdw>` or `<a:innerShdw>` element.
pub fn parse_shadow_effect<R: ThemeResolver + ?Sized>(
    node: Node<'_, '_>,
    resolver: &R,
    tint_mode: TintMode,
) -> Shadow {
    let (color, alpha) =
        try_resolved_color(node, resolver, tint_mode).unwrap_or_else(|| ("000000".to_owned(), 1.0));
    Shadow {
        color,
        alpha,
        blur: attr_i64(node, "blurRad").unwrap_or(0),
        dist: attr_i64(node, "dist").unwrap_or(0),
        dir: attr_f64(node, "dir").unwrap_or(0.0) / 60_000.0,
        sx: attr_f64(node, "sx").map(|value| value / 100_000.0),
        sy: attr_f64(node, "sy").map(|value| value / 100_000.0),
        kx: attr_f64(node, "kx").map(|value| value / 60_000.0),
        ky: attr_f64(node, "ky").map(|value| value / 60_000.0),
        algn: node.attribute("algn").map(ToOwned::to_owned),
        rot_with_shape: node
            .attribute("rotWithShape")
            .map(|value| value == "1" || value.eq_ignore_ascii_case("true")),
    }
}

/// Parse one `<a:effectLst>`. The caller owns precedence: in particular, an
/// authored empty list is a meaningful clear and must remain distinguishable
/// from an absent effect component.
pub fn parse_effect_list<R: ThemeResolver + ?Sized>(
    effect_list: Node<'_, '_>,
    resolver: &R,
    tint_mode: TintMode,
) -> DrawingMlEffects {
    let supported = ["outerShdw", "innerShdw", "glow", "softEdge", "reflection"];
    let mut result = DrawingMlEffects {
        unsupported: effect_list
            .children()
            .any(|node| node.is_element() && !supported.contains(&node.tag_name().name())),
        ..DrawingMlEffects::default()
    };
    result.shadow = child(effect_list, "outerShdw").and_then(|node| {
        if try_resolved_color(node, resolver, tint_mode).is_none() {
            result.unsupported = true;
            None
        } else {
            Some(parse_shadow_effect(node, resolver, tint_mode))
        }
    });
    result.inner_shadow = child(effect_list, "innerShdw").and_then(|node| {
        if try_resolved_color(node, resolver, tint_mode).is_none() {
            result.unsupported = true;
            None
        } else {
            Some(parse_shadow_effect(node, resolver, tint_mode))
        }
    });
    result.glow = child(effect_list, "glow").and_then(|node| {
        if try_resolved_color(node, resolver, tint_mode).is_none() {
            result.unsupported = true;
            None
        } else {
            Some(parse_glow_effect(node, resolver, tint_mode))
        }
    });
    result.soft_edge = child(effect_list, "softEdge").map(parse_soft_edge_effect);
    result.reflection = child(effect_list, "reflection").map(parse_reflection_effect);
    result
}

/// Parse one `<a:glow>` element.
pub fn parse_glow_effect<R: ThemeResolver + ?Sized>(
    node: Node<'_, '_>,
    resolver: &R,
    tint_mode: TintMode,
) -> Glow {
    let (color, alpha) =
        try_resolved_color(node, resolver, tint_mode).unwrap_or_else(|| ("000000".to_owned(), 1.0));
    Glow {
        color,
        alpha,
        radius: attr_i64(node, "rad").unwrap_or(0),
    }
}

/// Parse one `<a:softEdge>` element.
pub fn parse_soft_edge_effect(node: Node<'_, '_>) -> SoftEdge {
    SoftEdge {
        radius: attr_i64(node, "rad").unwrap_or(0),
    }
}

/// Parse one `<a:reflection>` element with the CT_ReflectionEffect defaults.
pub fn parse_reflection_effect(node: Node<'_, '_>) -> Reflection {
    let pct = |name: &str, default: f64| {
        attr_f64(node, name)
            .map(|value| value / 100_000.0)
            .unwrap_or(default)
    };
    Reflection {
        blur: attr_i64(node, "blurRad").unwrap_or(0),
        dist: attr_i64(node, "dist").unwrap_or(0),
        dir: attr_f64(node, "dir").unwrap_or(0.0) / 60_000.0,
        st_a: pct("stA", 1.0),
        st_pos: pct("stPos", 0.0),
        end_a: pct("endA", 0.0),
        end_pos: pct("endPos", 1.0),
        sx: pct("sx", 1.0),
        sy: pct("sy", -1.0),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::color::ThemeResolver;

    struct FixtureResolver;

    impl ThemeResolver for FixtureResolver {
        fn resolve_scheme_color(&self, name: &str) -> Option<String> {
            (name == "accent1").then(|| "123456".to_owned())
        }
    }

    #[test]
    fn parses_supported_effects_and_drawingml_defaults() {
        let doc = roxmltree::Document::parse(
            r#"<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:outerShdw blurRad="12700" dist="25400" dir="5400000" sx="50000" rotWithShape="0">
                <a:schemeClr val="accent1"><a:alpha val="50000"/></a:schemeClr>
              </a:outerShdw>
              <a:innerShdw><a:srgbClr val="ABCDEF"/></a:innerShdw>
              <a:glow rad="38100"><a:srgbClr val="FEDCBA"/></a:glow>
              <a:softEdge rad="6350"/>
              <a:reflection/>
            </a:effectLst>"#,
        )
        .unwrap();
        let effects = parse_effect_list(
            doc.root_element(),
            &FixtureResolver,
            TintMode::PowerPointLinear,
        );
        let shadow = effects.shadow.unwrap();
        assert_eq!(shadow.color, "123456");
        assert!((shadow.alpha - 0.5).abs() < 0.01);
        assert_eq!(shadow.dir, 90.0);
        assert_eq!(shadow.sx, Some(0.5));
        assert_eq!(shadow.rot_with_shape, Some(false));
        assert_eq!(effects.inner_shadow.unwrap().color, "ABCDEF");
        assert_eq!(effects.glow.unwrap().radius, 38100);
        assert_eq!(effects.soft_edge.unwrap().radius, 6350);
        assert_eq!(
            effects.reflection.unwrap(),
            Reflection {
                blur: 0,
                dist: 0,
                dir: 0.0,
                st_a: 1.0,
                st_pos: 0.0,
                end_a: 0.0,
                end_pos: 1.0,
                sx: 1.0,
                sy: -1.0,
            }
        );
        assert!(!effects.unsupported);
    }

    #[test]
    fn retains_unsupported_provenance_without_inventing_a_supported_effect() {
        let doc = roxmltree::Document::parse(
            r#"<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:blur rad="12700"/>
            </a:effectLst>"#,
        )
        .unwrap();
        let effects = parse_effect_list(
            doc.root_element(),
            &FixtureResolver,
            TintMode::PowerPointLinear,
        );
        assert!(effects.unsupported);
        assert!(effects.shadow.is_none());
        assert!(effects.glow.is_none());

        let unresolved = roxmltree::Document::parse(
            r#"<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:outerShdw><a:schemeClr val="missing"/></a:outerShdw>
            </a:effectLst>"#,
        )
        .unwrap();
        let effects = parse_effect_list(
            unresolved.root_element(),
            &FixtureResolver,
            TintMode::PowerPointLinear,
        );
        assert!(effects.unsupported);
        assert!(effects.shadow.is_none(), "must not invent a black shadow");
    }
}
