//! OfficeArt line-end and join/cap properties (MS-ODRAW 2.3.8.15/20-27,
//! 2.4.16-20) mapped to DrawingML (ECMA-376 20.1.8.9/38/43/52/57,
//! 20.1.10.31-34). No geometry, colors, or renderer-specific policy.
use super::unsupported;

/// MS-ODRAW 2.3.8.17 / 2.4.15 and ECMA-376 20.1.10.49 give the same
/// repeating bit patterns for these names. Custom lineDashStyle is separate.
pub(crate) fn preset_dash(value: u32) -> Option<&'static str> {
    [
        "solid",
        "sysDash",
        "sysDot",
        "sysDashDot",
        "sysDashDotDot",
        "dot",
        "dash",
        "lgDash",
        "dashDot",
        "lgDashDot",
        "lgDashDotDot",
    ]
    .get(value as usize)
    .copied()
}

#[derive(Clone, Copy, Default)]
pub(crate) struct Details {
    // Property order: start/end kind, start width/length, end width/length,
    // join, cap. Preserve absent versus explicit zero through inheritance.
    values: [Option<u8>; 8],
    miter: Option<u32>,
}
impl Details {
    pub fn property(&mut self, id: u16, value: u32) -> Result<(), String> {
        match id {
            0x1cc => {
                // This writer uses Transitional integer percentages, whose
                // ST_PercentageDecimal base is xsd:int. Reject unrepresentable
                // values rather than clamp them or emit invalid DrawingML.
                if value > i32::MAX as u32 || miter_percentage(value) > i32::MAX as u64 {
                    return Err(unsupported(
                        "OfficeArt line miter limit outside output range",
                    ));
                }
                self.miter = Some(value);
            }
            // MSOLINEEND's chevron values MUST be ignored, not mapped to
            // another arrow or mistaken for an explicit no-end override.
            0x1d0 | 0x1d1 if matches!(value, 6 | 7) => {}
            0x1d0..=0x1d7 => {
                let maximum = if id <= 0x1d1 { 5 } else { 2 };
                if value > maximum {
                    return Err(unsupported("invalid OfficeArt line decoration"));
                }
                self.values[usize::from(id - 0x1d0)] = Some(value as u8);
            }
            _ => {}
        }
        Ok(())
    }
    pub fn inherit(&self, parent: &Self) -> Self {
        Self {
            values: std::array::from_fn(|i| self.values[i].or(parent.values[i])),
            miter: self.miter.or(parent.miter),
        }
    }
    pub fn specified(&self) -> bool {
        self.miter.is_some() || self.values.iter().any(Option::is_some)
    }
    pub fn cap(&self) -> &'static str {
        // MS-ODRAW defaults: flat cap and round join.
        ["rnd", "sq", "flat"][usize::from(self.values[7].unwrap_or(2))]
    }
    pub fn children_xml(&self) -> String {
        let mut xml = match self.values[6].unwrap_or(2) {
            0 => "<a:bevel/>".to_string(),
            1 => {
                // Signed 16.16 ratio -> DrawingML thousandths of a percent.
                let limit = miter_percentage(self.miter.unwrap_or(0x80000));
                format!("<a:miter lim=\"{limit}\"/>")
            }
            _ => "<a:round/>".to_string(),
        };
        for (index, tag) in [(0, "headEnd"), (1, "tailEnd")] {
            let kind = ["none", "triangle", "stealth", "diamond", "oval", "arrow"]
                [usize::from(self.values[index].unwrap_or(0))];
            let sizes = ["sm", "med", "lg"];
            let width = sizes[usize::from(self.values[2 + index * 2].unwrap_or(1))];
            let length = sizes[usize::from(self.values[3 + index * 2].unwrap_or(1))];
            xml.push_str(&format!(
                "<a:{tag} type=\"{kind}\" w=\"{width}\" len=\"{length}\"/>"
            ));
        }
        xml
    }
}

fn miter_percentage(value: u32) -> u64 {
    (u64::from(value) * 100000 + 32768) / 65536
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn invalid_dash_enums_cannot_become_a_solid_line() {
        assert_eq!(preset_dash(0), Some("solid"));
        assert_eq!(preset_dash(10), Some("lgDashDotDot"));
        assert_eq!(preset_dash(11), None);
        assert_eq!(preset_dash(u32::MAX), None);
    }
    #[test]
    fn maps_all_supported_ends_and_dimensions() {
        for (kind, name) in ["none", "triangle", "stealth", "diamond", "oval", "arrow"]
            .iter()
            .enumerate()
        {
            for (width, w) in ["sm", "med", "lg"].iter().enumerate() {
                for (length, len) in ["sm", "med", "lg"].iter().enumerate() {
                    let mut d = Details::default();
                    for (id, value) in [
                        (0x1d0, kind),
                        (0x1d1, kind),
                        (0x1d2, width),
                        (0x1d3, length),
                        (0x1d4, width),
                        (0x1d5, length),
                    ] {
                        d.property(id, value as u32).unwrap();
                    }
                    let xml = d.children_xml();
                    for tag in ["headEnd", "tailEnd"] {
                        assert!(xml.contains(&format!(
                            "<a:{tag} type=\"{name}\" w=\"{w}\" len=\"{len}\"/>"
                        )));
                    }
                }
            }
        }
    }
    #[test]
    fn validates_enums_without_turning_ignored_values_into_overrides() {
        let mut parent = Details::default();
        parent.property(0x1d0, 1).unwrap();
        for ignored in [6, 7] {
            let mut child = Details::default();
            child.property(0x1d0, ignored).unwrap();
            assert!(!child.specified());
            assert!(child
                .inherit(&parent)
                .children_xml()
                .contains("type=\"triangle\""));
        }
        for id in 0x1d0..=0x1d7 {
            let invalid = if id <= 0x1d1 { 8 } else { 3 };
            assert!(Details::default().property(id, invalid).is_err());
        }
        assert!(Details::default().property(0x1cc, u32::MAX).is_err());
        let mut child = Details::default();
        child.property(0x1d0, 0).unwrap();
        assert!(!child
            .inherit(&parent)
            .children_xml()
            .contains("type=\"triangle\""));
    }
    #[test]
    fn retains_cap_join_and_fixed_point_miter_values_independently() {
        for (cap, name) in ["rnd", "sq", "flat"].iter().enumerate() {
            for (join, tag) in ["<a:bevel/>", "<a:miter lim=\"150000\"/>", "<a:round/>"]
                .iter()
                .enumerate()
            {
                let mut d = Details::default();
                d.property(0x1d7, cap as u32).unwrap();
                d.property(0x1d6, join as u32).unwrap();
                d.property(0x1cc, 0x18000).unwrap();
                assert_eq!(d.cap(), *name);
                assert!(d.children_xml().starts_with(tag));
            }
        }
        let d = Details::default();
        assert_eq!(d.cap(), "flat");
        assert!(d.children_xml().starts_with("<a:round/>"));
    }

    #[test]
    fn miter_conversion_stays_in_the_transitional_integer_percentage_range() {
        // Include values that round to, rather than exceed, xsd:int::MAX.
        let maximum = ((i32::MAX as u64 * 65536 + 32767) / 100000) as u32;
        let mut d = Details::default();
        d.property(0x1d6, 1).unwrap();
        d.property(0x1cc, maximum).unwrap();
        assert!(d
            .children_xml()
            .starts_with("<a:miter lim=\"2147483647\"/>"));
        assert!(d.property(0x1cc, maximum + 1).is_err());
        d.property(0x1cc, 0).unwrap();
        assert!(d.children_xml().starts_with("<a:miter lim=\"0\"/>"));
    }
}
