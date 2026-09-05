//! OfficeArt preset geometry and solid paint, without renderer extensions.
use super::{scheme, unsupported};

#[derive(Clone, Copy, Default)]
pub(super) struct Paint {
    details: crate::officeart::stroke::Details,
    pub custom_geometry: bool,
    fill_ok: Option<bool>,
    line_ok: Option<bool>,
    fill_rect: Option<bool>,
    fill_type: Option<u32>,
    fill: Option<u32>,
    fill_blip: Option<u32>,
    fill_alpha: Option<u32>,
    filled: Option<bool>,
    line_type: Option<u32>,
    line: Option<u32>,
    line_alpha: Option<u32>,
    lined: Option<bool>,
    width: Option<u32>,
    dash: Option<u32>,
}
impl Paint {
    /// Explicit hspMaster supplies defaults (MS-ODRAW 1.1, 2.2.40, 2.3.2.1).
    /// Preserve absence until the full chain is resolved, especially Boolean
    /// use bits: explicit false must win over inherited true. Colors stay in
    /// their source representation until the destination slide resolves them.
    pub fn inherit(&self, parent: &Self) -> Self {
        Self {
            details: self.details.inherit(&parent.details),
            // Unsupported inherited adjustments/paths must not turn into an
            // invented unadjusted preset. Explicit paths are decoded separately.
            custom_geometry: self.custom_geometry || parent.custom_geometry,
            fill_ok: self.fill_ok.or(parent.fill_ok),
            line_ok: self.line_ok.or(parent.line_ok),
            fill_rect: self.fill_rect.or(parent.fill_rect),
            fill_type: self.fill_type.or(parent.fill_type),
            fill: self.fill.or(parent.fill),
            fill_blip: self.fill_blip.or(parent.fill_blip),
            fill_alpha: self.fill_alpha.or(parent.fill_alpha),
            filled: self.filled.or(parent.filled),
            line_type: self.line_type.or(parent.line_type),
            line: self.line.or(parent.line),
            line_alpha: self.line_alpha.or(parent.line_alpha),
            lined: self.lined.or(parent.lined),
            width: self.width.or(parent.width),
            dash: self.dash.or(parent.dash),
        }
    }
    pub fn property(&mut self, id: u16, value: u32) -> Result<(), String> {
        // MS-ODRAW 2.3.6: customized vertices/segments/adjustments require their
        // own geometry conversion; never apply a preset over these overrides.
        match id {
            0x145..=0x150 => self.custom_geometry = true,
            // MS-ODRAW 2.3.6.31: geometry can veto paint independently of
            // fill/line style. Its use bits must not override those style bits.
            0x17f => {
                if value & 0x00010000 != 0 {
                    self.fill_ok = Some(value & 1 != 0);
                }
                if value & 0x00080000 != 0 {
                    self.line_ok = Some(value & 8 != 0);
                }
            }
            0x180 => self.fill_type = Some(value),
            0x181 => self.fill = Some(value),
            // Caller validates fBid on this one-based BStore reference.
            0x4186 => self.fill_blip = Some(value),
            0x182 | 0x1c1 => {
                if value > 65536 {
                    return Err(unsupported("invalid PowerPoint paint opacity"));
                }
                if id == 0x182 {
                    self.fill_alpha = Some(value);
                } else {
                    self.line_alpha = Some(value);
                }
            }
            // Boolean property's high word contains use bits, low word values.
            // MS-ODRAW 2.3.7.43 and 2.3.8.38: unused values cannot override paint.
            0x1bf => {
                if value & 0x00100000 != 0 {
                    self.filled = Some(value & 0x10 != 0);
                }
                if value & 0x00020000 != 0 {
                    self.fill_rect = Some(value & 2 != 0);
                }
            }
            0x1ff if value & 0x00080000 != 0 => self.lined = Some(value & 8 != 0),
            0x1c0 => self.line = Some(value),
            0x1c4 => self.line_type = Some(value),
            0x1cb => {
                if value > 0x132f540 {
                    return Err(unsupported("invalid PowerPoint line width"));
                }
                self.width = Some(value);
            }
            0x1ce => self.dash = Some(value),
            _ => self.details.property(id, value)?,
        }
        Ok(())
    }
    pub fn geometry(&self, kind: u16) -> Option<&'static str> {
        if self.custom_geometry {
            return None;
        }
        // MS-ODRAW 2.4.24 -> ECMA-376 ST_ShapeType. Only presets whose
        // unadjusted outlines correspond directly are included here.
        match kind {
            1 | 202 => Some("rect"),
            3 => Some("ellipse"),
            4 => Some("diamond"),
            5 => Some("triangle"),
            6 => Some("rtTriangle"),
            20 => Some("line"),
            _ => None,
        }
    }
    /// Background paint has no geometry or line, and must not acquire a fake
    /// preset merely to extract its fill (PresentationML CT_BackgroundProperties).
    pub fn background_fill(&self, scheme: Option<&scheme::Scheme>) -> Option<String> {
        if !self.filled.unwrap_or(true) || !self.fill_ok.unwrap_or(true) {
            return Some("<a:noFill/>".into());
        }
        if self.fill_rect.unwrap_or(false) || self.fill_type.unwrap_or(0) != 0 {
            return None;
        }
        solid(
            self.fill.unwrap_or(0xffffff),
            self.fill_alpha.unwrap_or(65536),
            scheme,
        )
    }
    pub fn background_image(&self) -> Option<(u32, u32)> {
        (self.fill_type == Some(3)
            && self.fill_blip.unwrap_or(0) != 0
            && !self.fill_rect.unwrap_or(false)
            && self.filled.unwrap_or(true)
            && self.fill_ok.unwrap_or(true))
        .then_some((
            self.fill_blip.unwrap_or(0),
            self.fill_alpha.unwrap_or(65536),
        ))
    }
    #[cfg(test)]
    fn xml(&self, kind: u16) -> String {
        self.xml_with_scheme(kind, None)
    }
    pub fn xml_with_scheme(&self, kind: u16, scheme: Option<&scheme::Scheme>) -> String {
        let no_paint = "<a:noFill/><a:ln><a:noFill/></a:ln>";
        if self.geometry(kind).is_none() {
            return no_paint.into();
        }
        self.xml_with_custom_geometry(scheme, kind != 20, true)
    }
    /// The caller has reconstructed explicit custom paths. Geometry and path
    /// paint vetoes stay separate from style inheritance and property defaults.
    pub fn xml_with_custom_geometry(
        &self,
        scheme: Option<&scheme::Scheme>,
        allow_fill: bool,
        allow_line: bool,
    ) -> String {
        // Direct and explicitly linked master paint are reconstructed. Drawing
        // defaults and unlinked masters remain absent, with a conversion warning.
        // Within an explicit layer, use the normative MS-ODRAW property defaults.
        let fill_set = self.fill.is_some()
            || self.filled.is_some()
            || self.fill_type.is_some()
            || self.fill_alpha.is_some();
        let line_set = self.line.is_some()
            || self.lined.is_some()
            || self.line_type.is_some()
            || self.line_alpha.is_some()
            || self.width.is_some()
            || self.dash.is_some()
            || self.details.specified();
        let fill = if allow_fill
            && fill_set
            && self.filled.unwrap_or(true)
            && self.fill_ok.unwrap_or(true)
            && !self.fill_rect.unwrap_or(false)
            && self.fill_type.unwrap_or(0) == 0
        {
            solid(
                self.fill.unwrap_or(0xffffff),
                self.fill_alpha.unwrap_or(65536),
                scheme,
            )
        } else {
            None
        };
        let line = if allow_line
            && line_set
            && self.lined.unwrap_or(true)
            && self.line_ok.unwrap_or(true)
            && self.line_type.unwrap_or(0) == 0
            && self.dash.unwrap_or(0) == 0
        {
            solid(
                self.line.unwrap_or(0),
                self.line_alpha.unwrap_or(65536),
                scheme,
            )
        } else {
            None
        };
        let mut xml = fill.unwrap_or_else(|| "<a:noFill/>".into());
        if let Some(line) = line {
            xml.push_str(&format!(
                "<a:ln w=\"{}\" cap=\"{}\">{line}{}</a:ln>",
                self.width.unwrap_or(9525),
                self.details.cap(),
                self.details.children_xml()
            ));
        } else {
            xml.push_str("<a:ln><a:noFill/></a:ln>");
        }
        xml
    }
}

fn solid(color: u32, opacity: u32, scheme: Option<&scheme::Scheme>) -> Option<String> {
    let color = scheme::drawing(color, scheme)?;
    let mut xml = format!(
        "<a:solidFill><a:srgbClr val=\"{:02X}{:02X}{:02X}\">",
        color & 255,
        (color >> 8) & 255,
        (color >> 16) & 255
    );
    if opacity != 65536 {
        xml.push_str(&format!(
            "<a:alpha val=\"{}\"/>",
            (u64::from(opacity) * 100000 + 32768) / 65536
        ));
    }
    xml.push_str("</a:srgbClr></a:solidFill>");
    Some(xml)
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn line_decorations_caps_and_joins_reach_drawingml() {
        let mut p = Paint::default();
        for (id, value) in [
            (0x1c0, 0),
            (0x1d0, 1),
            (0x1d1, 5),
            (0x1d2, 0),
            (0x1d3, 2),
            (0x1d4, 2),
            (0x1d5, 0),
            (0x1d6, 0),
            (0x1d7, 0),
        ] {
            p.property(id, value).unwrap();
        }
        let xml = p.xml(20);
        assert!(xml.contains("cap=\"rnd\""));
        assert!(xml.contains("<a:bevel/>"));
        assert!(xml.contains("<a:headEnd type=\"triangle\" w=\"sm\" len=\"lg\"/>"));
        assert!(xml.contains("<a:tailEnd type=\"arrow\" w=\"lg\" len=\"sm\"/>"));
        p.property(0x1ff, 0x00080000).unwrap();
        assert_eq!(p.xml(1), "<a:noFill/><a:ln><a:noFill/></a:ln>");
    }

    #[test]
    fn decorations_inherit_with_explicit_none_and_miter_defaults() {
        let mut parent = Paint::default();
        for (id, value) in [(0x1c0, 0), (0x1d0, 3), (0x1d1, 4), (0x1d6, 1), (0x1d7, 1)] {
            parent.property(id, value).unwrap();
        }
        let mut child = Paint::default();
        child.property(0x1d0, 0).unwrap();
        child.property(0x1d4, 0).unwrap();
        let xml = child.inherit(&parent).xml(20);
        assert!(!xml.contains("type=\"diamond\""));
        assert!(xml.contains("<a:tailEnd type=\"oval\" w=\"sm\" len=\"med\"/>"));
        assert!(xml.contains("cap=\"sq\""));
        assert!(xml.contains("<a:miter lim=\"800000\"/>"));
        child.property(0x1cc, 0x00018000).unwrap();
        assert!(child
            .inherit(&parent)
            .xml(20)
            .contains("<a:miter lim=\"150000\"/>"));
    }

    #[test]
    fn arrow_editability_flags_do_not_hide_end_decorations() {
        let mut p = Paint::default();
        p.property(0x1d1, 1).unwrap();
        // MS-ODRAW 2.3.8.38: fArrowheadsOK controls editing, not rendering.
        p.property(0x1ff, 0x00100000).unwrap();
        assert!(p.xml(20).contains("<a:tailEnd type=\"triangle\""));
        assert!(p.xml(20).contains("<a:round/>"));
    }

    #[test]
    fn inherited_paint_keeps_values_independent_from_boolean_use_bits() {
        let mut master = Paint::default();
        for (id, value) in [
            (0x181, 255),
            (0x182, 32768),
            (0x1c0, 0xff0000),
            (0x1cb, 25400),
        ] {
            master.property(id, value).unwrap();
        }
        let mut local = Paint::default();
        local.property(0x1bf, 0).unwrap(); // Unused false is not an override.
        let inherited = local.inherit(&master);
        let xml = inherited.xml(1);
        assert!(xml.contains("FF0000"));
        assert!(xml.contains("0000FF"));
        assert!(xml.contains("50000"));
        assert!(xml.contains("w=\"25400\""));
        local.property(0x1bf, 0x00100000).unwrap();
        local.property(0x1ff, 0x00080000).unwrap();
        assert_eq!(
            local.inherit(&master).xml(1),
            "<a:noFill/><a:ln><a:noFill/></a:ln>"
        );
        local.property(0x1bf, 0x00100010).unwrap();
        local.property(0x181, 0xff00).unwrap();
        let xml = local.inherit(&master).xml(1);
        assert!(xml.contains("00FF00"));
        assert!(xml.contains("50000"));
        assert!(!xml.contains("0000FF"));
    }

    #[test]
    fn unsupported_inherited_paint_is_not_replaced_with_solid_defaults() {
        let mut master = Paint::default();
        master.property(0x180, 4).unwrap(); // Gradient.
        master.property(0x1c0, 255).unwrap();
        master.property(0x1ce, 1).unwrap(); // Dashed line, not yet supported.
        let mut local = Paint::default();
        local.property(0x181, 0xff00).unwrap();
        assert_eq!(
            local.inherit(&master).xml(1),
            "<a:noFill/><a:ln><a:noFill/></a:ln>"
        );
        local.property(0x180, 0).unwrap();
        assert!(local.inherit(&master).xml(1).contains("00FF00"));
        master.property(0x1bf, 0x00020002).unwrap();
        assert!(!local.inherit(&master).xml(1).contains("00FF00"));
        local.property(0x1bf, 0x00020000).unwrap(); // Explicitly clear inherited fill rectangle.
        assert!(local.inherit(&master).xml(1).contains("00FF00"));
        master.custom_geometry = true;
        assert!(local.inherit(&master).geometry(1).is_none());
    }

    #[test]
    fn inherited_scheme_color_resolves_at_destination_not_master() {
        let mut master = Paint::default();
        master.property(0x181, 0x08000004).unwrap();
        let inherited = Paint::default().inherit(&master);
        let mut scheme = [0; 8];
        scheme[4] = 0x563412;
        assert!(inherited
            .xml_with_scheme(1, Some(&scheme))
            .contains("123456"));
        scheme[4] = 0xabcdef;
        assert!(inherited
            .xml_with_scheme(1, Some(&scheme))
            .contains("EFCDAB"));
    }
    #[test]
    fn background_fill_has_no_geometry_or_line_and_respects_use_bits() {
        let mut p = Paint::default();
        p.property(0x181, 0x112233).unwrap();
        p.property(0x182, 32768).unwrap();
        p.property(0x1c0, 0xff).unwrap();
        p.custom_geometry = true;
        let xml = p.background_fill(None).unwrap();
        assert!(xml.contains("332211"));
        assert!(xml.contains("50000"));
        assert!(!xml.contains("a:ln"));
        p.property(0x180, 3).unwrap();
        p.property(0x4186, 9).unwrap();
        assert_eq!(p.background_image(), Some((9, 32768)));
        assert!(p.background_fill(None).is_none());
        p.property(0x1bf, 0x00100000).unwrap();
        assert!(p.background_image().is_none());
        assert_eq!(p.background_fill(None).unwrap(), "<a:noFill/>");
    }
    #[test]
    fn maps_only_supported_unmodified_presets() {
        let mut p = Paint::default();
        for (kind, name) in [
            (1, "rect"),
            (3, "ellipse"),
            (4, "diamond"),
            (5, "triangle"),
            (6, "rtTriangle"),
            (20, "line"),
            (202, "rect"),
        ] {
            assert_eq!(p.geometry(kind), Some(name));
        }
        assert_eq!(p.geometry(0), None);
        p.property(0x147, 100).unwrap();
        assert_eq!(p.geometry(5), None);
    }
    #[test]
    fn literal_colors_width_and_fixed_point_opacity_survive() {
        let mut p = Paint::default();
        for (id, value) in [
            (0x181, 0x00563412),
            (0x182, 32768),
            (0x1c0, 0x00efcdab),
            (0x1cb, 25400),
        ] {
            p.property(id, value).unwrap();
        }
        let xml = p.xml(3);
        assert!(xml.contains("val=\"123456\"><a:alpha val=\"50000\"/>"));
        assert!(xml.contains("<a:ln w=\"25400\" cap=\"flat\">"));
        assert!(xml.contains("val=\"ABCDEF\""));
    }
    #[test]
    fn boolean_use_bits_control_suppression_not_the_unused_values() {
        let mut p = Paint::default();
        p.property(0x181, 0x000000ff).unwrap();
        p.property(0x1c0, 0).unwrap();
        p.property(0x1bf, 0).unwrap();
        p.property(0x1ff, 0).unwrap();
        assert!(p.xml(1).contains("FF0000"));
        p.property(0x1bf, 0x00100000).unwrap();
        p.property(0x1ff, 0x00080000).unwrap();
        assert_eq!(p.xml(1), "<a:noFill/><a:ln><a:noFill/></a:ln>");
    }
    #[test]
    fn does_not_invent_scheme_colors_gradient_fills_or_unknown_geometry() {
        let mut p = Paint::default();
        p.property(0x181, 0x08000005).unwrap();
        assert!(!p.xml(1).contains("srgbClr"));
        p.property(0x181, 255).unwrap();
        p.property(0x180, 4).unwrap();
        assert!(!p.xml(1).contains("FF0000"));
        p.property(0x180, 0).unwrap();
        assert!(!p.xml(0).contains("FF0000"));
        assert!(!p.xml(20).contains("FF0000"));
    }
    #[test]
    fn validates_opacity_and_line_width_without_clamping() {
        let mut p = Paint::default();
        assert!(p.property(0x182, 65537).is_err());
        assert!(p.property(0x1c1, u32::MAX).is_err());
        assert!(p.property(0x1cb, u32::MAX).is_err());
    }

    #[test]
    fn geometry_vetoes_are_independent_and_custom_fill_rects_are_not_invented() {
        let mut p = Paint::default();
        p.property(0x181, 255).unwrap();
        p.property(0x1c0, 0).unwrap();
        p.property(0x17f, 0x00090000).unwrap();
        p.property(0x1bf, 0x00100010).unwrap();
        p.property(0x1ff, 0x00080008).unwrap();
        assert_eq!(p.xml(1), "<a:noFill/><a:ln><a:noFill/></a:ln>");
        p.property(0x17f, 0x00090009).unwrap();
        assert!(p.xml(1).contains("FF0000"));
        p.property(0x1bf, 0x00020002).unwrap();
        assert!(!p.xml(1).contains("FF0000"));
    }
}
