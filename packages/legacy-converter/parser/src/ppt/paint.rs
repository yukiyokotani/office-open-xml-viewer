//! Direct OfficeArt preset geometry and solid paint, without renderer extensions.
use super::{scheme, unsupported};

#[derive(Default)]
pub(super) struct Paint {
    pub custom_geometry: bool,
    fill_ok: Option<bool>,
    line_ok: Option<bool>,
    fill_rect: bool,
    fill_type: Option<u32>,
    fill: Option<u32>,
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
                    self.fill_rect = value & 2 != 0;
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
            _ => {}
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
    #[cfg(test)]
    fn xml(&self, kind: u16) -> String {
        self.xml_with_scheme(kind, None)
    }
    pub fn xml_with_scheme(&self, kind: u16, scheme: Option<&scheme::Scheme>) -> String {
        let no_paint = "<a:noFill/><a:ln><a:noFill/></a:ln>";
        if self.geometry(kind).is_none() {
            return no_paint.into();
        }
        // Only direct paint is reconstructed until drawing/master defaults are
        // resolved. An absent paint layer stays omitted, with a conversion warning.
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
            || self.dash.is_some();
        let fill = if kind != 20
            && fill_set
            && self.filled.unwrap_or(true)
            && self.fill_ok.unwrap_or(true)
            && !self.fill_rect
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
        let line = if line_set
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
                "<a:ln w=\"{}\">{line}</a:ln>",
                self.width.unwrap_or(9525)
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
        assert!(xml.contains("<a:ln w=\"25400\">"));
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
