//! MS-DOC 2.6.3 / 2.9.208 / XAS_plusOne / YAS_plusOne to ECMA-376 17.4.57.
//! Preserve authored positioning; do not compensate for renderer layout.
use super::{boolean, nonnegative, signed};

#[derive(Clone, Default)]
pub struct Position {
    anchors: Option<u8>,
    x: i32,
    y: i32,
    // Physical left, top, right, bottom; no bidi-dependent permutation.
    distances: [i32; 4],
    no_overlap: bool,
}

impl Position {
    pub fn apply(&mut self, code: u16, b: &[u8]) -> Result<bool, String> {
        match code {
            // Padding MUST be ignored (MS-DOC 2.9.208).
            0x360d => self.anchors = Some(b[0] & 0xf0),
            0x940e | 0x940f => {
                let value = signed(b)?;
                if !(-31679..=31681).contains(&value) {
                    return Err(super::unsupported("invalid Word table position"));
                }
                if code == 0x940e {
                    self.x = value;
                } else {
                    self.y = value;
                }
            }
            0x9410 => self.distances[0] = nonnegative(b)?,
            0x9411 => self.distances[1] = nonnegative(b)?,
            0x941e => self.distances[2] = nonnegative(b)?,
            0x941f => self.distances[3] = nonnegative(b)?,
            0x3465 => self.no_overlap = boolean(b[0])?,
            _ => return Ok(false),
        }
        Ok(true)
    }

    pub fn xml(&self) -> String {
        let pc = self.anchors.unwrap_or(0);
        let vertical = (pc >> 4) & 3;
        let horizontal = pc >> 6;
        // MS-DOC 2.7.13 Copts identifies tblpPr with
        // nondefault position/wrapping properties. No-overlap alone is not one.
        let active = self.anchors.is_some()
            || self.x != 0
            || self.y != 0
            || self.distances.iter().any(|v| *v != 0);
        let mut xml = String::new();
        if active && vertical != 3 && horizontal != 3 {
            xml.push_str(&format!("<w:tblpPr w:horzAnchor=\"{}\" w:vertAnchor=\"{}\" {} {} w:leftFromText=\"{}\" w:topFromText=\"{}\" w:rightFromText=\"{}\" w:bottomFromText=\"{}\"/>",
                ["text", "margin", "page"][horizontal as usize],
                ["margin", "page", "text"][vertical as usize],
                coordinate("X", self.x), coordinate("Y", self.y),
                self.distances[0], self.distances[1], self.distances[2], self.distances[3]));
        }
        if self.no_overlap {
            xml.push_str("<w:tblOverlap w:val=\"never\"/>");
        }
        xml
    }
}

fn coordinate(axis: &str, value: i32) -> String {
    let special = match (axis, value) {
        ("X", 0) => Some("left"),
        ("X", -4) => Some("center"),
        ("X", -8) => Some("right"),
        ("X", -12) => Some("inside"),
        ("X", -16) => Some("outside"),
        ("Y", 0) => Some("inline"),
        ("Y", -4) => Some("top"),
        ("Y", -8) => Some("center"),
        ("Y", -12) => Some("bottom"),
        ("Y", -16) => Some("inside"),
        ("Y", -20) => Some("outside"),
        _ => None,
    };
    if let Some(value) = special {
        format!("w:tblp{axis}Spec=\"{value}\"")
    } else {
        // MS-DOC 2.9.351/357: distances are stored one greater than twips.
        format!("w:tblp{axis}=\"{}\"", value - 1)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn anchors_and_ignored_padding_cover_the_complete_byte_domain() {
        for bits in 0..=255u8 {
            let mut p = Position::default();
            p.apply(0x360d, &[bits]).unwrap();
            let x = bits >> 6;
            let y = (bits >> 4) & 3;
            if x == 3 || y == 3 {
                assert!(p.xml().is_empty());
            } else {
                assert!(p.xml().contains(&format!(
                    "w:horzAnchor=\"{}\"",
                    ["text", "margin", "page"][x as usize]
                )));
                assert!(p.xml().contains(&format!(
                    "w:vertAnchor=\"{}\"",
                    ["margin", "page", "text"][y as usize]
                )));
            }
        }
        assert!(Position::default().xml().is_empty());
    }
    #[test]
    fn symbolic_positions_and_signed_twips_do_not_share_the_minus_one_path() {
        for (axis, labels) in [
            ("X", vec!["left", "center", "right", "inside", "outside"]),
            (
                "Y",
                vec!["inline", "top", "center", "bottom", "inside", "outside"],
            ),
        ] {
            for (i, label) in labels.iter().enumerate() {
                assert_eq!(
                    coordinate(axis, -(i as i32) * 4),
                    format!("w:tblp{axis}Spec=\"{label}\"")
                );
            }
            for n in [-31679, -21, -3, -1, 1, 2, 721, 31681] {
                assert_eq!(coordinate(axis, n), format!("w:tblp{axis}=\"{}\"", n - 1));
            }
        }
    }
    #[test]
    fn distances_and_overlap_keep_physical_sides_and_validate_bounds() {
        let mut p = Position::default();
        for (code, n) in [(0x9410, 12u16), (0x9411, 34), (0x941e, 56), (0x941f, 78)] {
            p.apply(code, &n.to_le_bytes()).unwrap();
            assert!(p.apply(code, &31681u16.to_le_bytes()).is_err());
        }
        assert!(p.xml().contains("w:leftFromText=\"12\" w:topFromText=\"34\" w:rightFromText=\"56\" w:bottomFromText=\"78\""));
        for code in [0x940e, 0x940f] {
            for n in [-32768i16, -31680, 31682, 32767] {
                assert!(p.apply(code, &n.to_le_bytes()).is_err());
            }
        }
        let mut p = Position::default();
        p.apply(0x3465, &[1]).unwrap();
        assert_eq!(p.xml(), "<w:tblOverlap w:val=\"never\"/>");
        p.apply(0x3465, &[0]).unwrap();
        assert!(p.xml().is_empty());
        assert!(p.apply(0x3465, &[2]).is_err());
    }
}
