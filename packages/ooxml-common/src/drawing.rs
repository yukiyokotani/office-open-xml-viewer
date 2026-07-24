//! Shared DrawingML non-visual-properties helpers (`cNvPr` / `docPr`).
//!
//! Every OOXML host embeds the same DrawingML non-visual drawing properties
//! (ECMA-376 §20.1.2.2.8 `CT_NonVisualDrawingProps`, exposed as `<*:cNvPr>` in
//! pptx/xlsx drawings and as `<wp:docPr>` on a WordprocessingML `<wp:inline>` /
//! `<wp:anchor>`). Its `hidden` attribute (`xsd:boolean`, default `false`)
//! marks a drawing object that must **not** be rendered. This predicate is the
//! single source of truth for reading that flag across the three parsers.

use roxmltree::Node;

/// Parse an `xsd:boolean` attribute value. Per the W3C XML Schema lexical
/// space, the four valid literals are `true` / `false` / `1` / `0`; any other
/// text is not a valid boolean and yields `None` (callers apply the schema
/// default themselves). Surrounding whitespace is tolerated.
pub fn parse_xsd_bool(value: &str) -> Option<bool> {
    match value.trim() {
        "true" | "1" => Some(true),
        "false" | "0" => Some(false),
        _ => None,
    }
}

/// True when a DrawingML non-visual-properties node
/// (`<*:cNvPr>` or `<wp:docPr>`) carries `hidden` set to a truthy
/// `xsd:boolean` (§20.1.2.2.8 / §20.4.2.5). Absent or `false`/`0` → not hidden
/// (the schema default). A hidden drawing object is not rendered.
pub fn nv_props_hidden(nv_props: Node) -> bool {
    nv_props
        .attribute("hidden")
        .and_then(parse_xsd_bool)
        .unwrap_or(false)
}

/// One DrawingML group-transform step (ECMA-376 §20.1.7.5 and Annex L
/// §L.4.7.4): child bounding box, parent bounding box, then flip and rotation.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct DrawingGroupSpec {
    pub off_x: f64,
    pub off_y: f64,
    pub ext_x: f64,
    pub ext_y: f64,
    pub child_off_x: f64,
    pub child_off_y: f64,
    pub child_ext_x: f64,
    pub child_ext_y: f64,
    pub rotation_degrees: f64,
    pub flip_h: bool,
    pub flip_v: bool,
}

/// Cumulative DrawingML group hierarchy. Annex L §L.4.7.4–§L.4.7.6 requires
/// two related representations: scale/flip/rotation compose independently for
/// the leaf's effective geometry, while the full per-level matrices are applied
/// to the leaf's original centre to determine translation.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct DrawingGroupTransform {
    pub scale_x: f64,
    pub scale_y: f64,
    pub rotation_degrees: f64,
    pub flip_h: bool,
    pub flip_v: bool,
    m11: f64,
    m12: f64,
    m21: f64,
    m22: f64,
    tx: f64,
    ty: f64,
    non_neutral_group_levels: u32,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub struct DrawingRect {
    pub x: f64,
    pub y: f64,
    pub width: f64,
    pub height: f64,
    pub rotation_degrees: f64,
    pub flip_h: bool,
    pub flip_v: bool,
}

impl DrawingGroupTransform {
    pub const IDENTITY: Self = Self {
        scale_x: 1.0,
        scale_y: 1.0,
        rotation_degrees: 0.0,
        flip_h: false,
        flip_v: false,
        m11: 1.0,
        m12: 0.0,
        m21: 0.0,
        m22: 1.0,
        tx: 0.0,
        ty: 0.0,
        non_neutral_group_levels: 0,
    };

    pub fn from_group(spec: DrawingGroupSpec) -> Self {
        let scale_x = if spec.child_ext_x != 0.0 {
            spec.ext_x / spec.child_ext_x
        } else {
            1.0
        };
        let scale_y = if spec.child_ext_y != 0.0 {
            spec.ext_y / spec.child_ext_y
        } else {
            1.0
        };
        let radians = spec.rotation_degrees.to_radians();
        let cos = radians.cos();
        let sin = radians.sin();
        let fx = if spec.flip_h { -1.0 } else { 1.0 };
        let fy = if spec.flip_v { -1.0 } else { 1.0 };
        let m11 = cos * fx * scale_x;
        let m12 = -sin * fy * scale_y;
        let m21 = sin * fx * scale_x;
        let m22 = cos * fy * scale_y;
        let center_x = spec.off_x + spec.ext_x / 2.0;
        let center_y = spec.off_y + spec.ext_y / 2.0;
        let scaled_origin_x = spec.off_x - spec.child_off_x * scale_x;
        let scaled_origin_y = spec.off_y - spec.child_off_y * scale_y;
        let dx = scaled_origin_x - center_x;
        let dy = scaled_origin_y - center_y;
        let tx = center_x + cos * fx * dx - sin * fy * dy;
        let ty = center_y + sin * fx * dx + cos * fy * dy;
        Self {
            scale_x,
            scale_y,
            rotation_degrees: spec.rotation_degrees,
            flip_h: spec.flip_h,
            flip_v: spec.flip_v,
            m11,
            m12,
            m21,
            m22,
            tx,
            ty,
            non_neutral_group_levels: u32::from(!group_transform_is_neutral(spec)),
        }
    }

    pub fn compose_group(self, spec: DrawingGroupSpec) -> Self {
        let child = Self::from_group(spec);
        Self {
            scale_x: self.scale_x * child.scale_x,
            scale_y: self.scale_y * child.scale_y,
            rotation_degrees: self.rotation_degrees + child.rotation_degrees,
            flip_h: self.flip_h ^ child.flip_h,
            flip_v: self.flip_v ^ child.flip_v,
            m11: self.m11 * child.m11 + self.m12 * child.m21,
            m12: self.m11 * child.m12 + self.m12 * child.m22,
            m21: self.m21 * child.m11 + self.m22 * child.m21,
            m22: self.m21 * child.m12 + self.m22 * child.m22,
            tx: self.m11 * child.tx + self.m12 * child.ty + self.tx,
            ty: self.m21 * child.tx + self.m22 * child.ty + self.ty,
            non_neutral_group_levels: self.non_neutral_group_levels
                + child.non_neutral_group_levels,
        }
    }

    /// Count group levels that alter scale, rotation, or reflection. Pure
    /// translation wrappers are neutral because they do not affect a leaf's
    /// scale/rotation ordering. OOXML transform inputs are authored integers
    /// below f64's exact-integer range, so exact neutral comparisons are stable.
    pub fn non_neutral_group_levels(self) -> u32 {
        self.non_neutral_group_levels
    }

    pub fn map_point(self, x: f64, y: f64) -> (f64, f64) {
        (
            self.m11 * x + self.m12 * y + self.tx,
            self.m21 * x + self.m22 * y + self.ty,
        )
    }

    /// Apply the Annex L nested-transform rendering procedure to a leaf's
    /// authored bounding box and own rotation/flip.
    pub fn apply_rect(self, rect: DrawingRect) -> DrawingRect {
        let (center_x, center_y) =
            self.map_point(rect.x + rect.width / 2.0, rect.y + rect.height / 2.0);
        let mapped_width = rect.width * self.scale_x;
        let mapped_height = rect.height * self.scale_y;
        DrawingRect {
            x: center_x - mapped_width / 2.0,
            y: center_y - mapped_height / 2.0,
            width: mapped_width,
            height: mapped_height,
            rotation_degrees: self.rotation_degrees + rect.rotation_degrees,
            flip_h: self.flip_h ^ rect.flip_h,
            flip_v: self.flip_v ^ rect.flip_v,
        }
    }
}

fn group_transform_is_neutral(spec: DrawingGroupSpec) -> bool {
    let scale_x = if spec.child_ext_x != 0.0 {
        spec.ext_x / spec.child_ext_x
    } else {
        1.0
    };
    let scale_y = if spec.child_ext_y != 0.0 {
        spec.ext_y / spec.child_ext_y
    } else {
        1.0
    };
    !spec.flip_h
        && !spec.flip_v
        && spec.rotation_degrees.rem_euclid(360.0) == 0.0
        && scale_x == 1.0
        && scale_y == 1.0
}

#[cfg(test)]
mod tests {
    use super::*;
    use roxmltree::Document;

    #[test]
    fn parse_xsd_bool_accepts_all_four_literals() {
        assert_eq!(parse_xsd_bool("true"), Some(true));
        assert_eq!(parse_xsd_bool("1"), Some(true));
        assert_eq!(parse_xsd_bool("false"), Some(false));
        assert_eq!(parse_xsd_bool("0"), Some(false));
        // Whitespace tolerated.
        assert_eq!(parse_xsd_bool("  1 "), Some(true));
        // Anything else is not a valid boolean.
        assert_eq!(parse_xsd_bool("yes"), None);
        assert_eq!(parse_xsd_bool(""), None);
        assert_eq!(parse_xsd_bool("TRUE"), None); // case-sensitive per XSD
    }

    fn node_from(xml: &str) -> Document<'_> {
        Document::parse(xml).unwrap()
    }

    #[test]
    fn nv_props_hidden_reads_boolean_default_false() {
        // hidden="1" and hidden="true" → hidden.
        for attr in ["1", "true"] {
            let xml = format!(r#"<cNvPr id="2" name="x" hidden="{attr}"/>"#);
            let doc = node_from(&xml);
            assert!(nv_props_hidden(doc.root_element()), "hidden={attr}");
        }
        // hidden="0" / "false" / absent → not hidden.
        for xml in [
            r#"<cNvPr id="2" name="x" hidden="0"/>"#,
            r#"<cNvPr id="2" name="x" hidden="false"/>"#,
            r#"<cNvPr id="2" name="x"/>"#,
        ] {
            let doc = node_from(xml);
            assert!(!nv_props_hidden(doc.root_element()), "xml={xml}");
        }
    }

    #[test]
    fn group_transform_keeps_scale_axes_independent_of_child_rotation() {
        let transform = DrawingGroupTransform::from_group(DrawingGroupSpec {
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
        });
        let mapped = transform.apply_rect(DrawingRect {
            x: 0.0,
            y: 50_800.0,
            width: 127_000.0,
            height: 25_400.0,
            rotation_degrees: 90.0,
            flip_h: false,
            flip_v: false,
        });

        assert!((mapped.x - 0.0).abs() < 1e-6);
        assert!((mapped.y - 101_600.0).abs() < 1e-6);
        assert!((mapped.width - 127_000.0).abs() < 1e-6);
        assert!((mapped.height - 50_800.0).abs() < 1e-6);
        assert!((mapped.rotation_degrees - 90.0).abs() < 1e-6);
    }

    #[test]
    fn nested_group_transform_composes_scale_and_translation() {
        let outer = DrawingGroupTransform::from_group(DrawingGroupSpec {
            off_x: 10.0,
            off_y: 20.0,
            ext_x: 200.0,
            ext_y: 300.0,
            child_off_x: 5.0,
            child_off_y: 10.0,
            child_ext_x: 100.0,
            child_ext_y: 100.0,
            rotation_degrees: 0.0,
            flip_h: false,
            flip_v: false,
        });
        let nested = outer.compose_group(DrawingGroupSpec {
            off_x: 30.0,
            off_y: 40.0,
            ext_x: 50.0,
            ext_y: 80.0,
            child_off_x: 10.0,
            child_off_y: 20.0,
            child_ext_x: 25.0,
            child_ext_y: 40.0,
            rotation_degrees: 0.0,
            flip_h: false,
            flip_v: false,
        });
        let point = nested.map_point(10.0, 20.0);

        assert!((point.0 - 60.0).abs() < 1e-6);
        assert!((point.1 - 110.0).abs() < 1e-6);
    }

    #[test]
    fn group_rotation_and_flip_map_center_and_compose_leaf_properties() {
        let transform = DrawingGroupTransform::from_group(DrawingGroupSpec {
            off_x: 0.0,
            off_y: 0.0,
            ext_x: 200.0,
            ext_y: 100.0,
            child_off_x: 0.0,
            child_off_y: 0.0,
            child_ext_x: 100.0,
            child_ext_y: 100.0,
            rotation_degrees: 90.0,
            flip_h: true,
            flip_v: false,
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

        assert!((mapped.x - 105.0).abs() < 1e-6);
        assert!((mapped.y - 105.0).abs() < 1e-6);
        assert!((mapped.width - 40.0).abs() < 1e-6);
        assert!((mapped.height - 10.0).abs() < 1e-6);
        assert!((mapped.rotation_degrees - 105.0).abs() < 1e-6);
        assert!(mapped.flip_h);
        assert!(mapped.flip_v);
    }
}
