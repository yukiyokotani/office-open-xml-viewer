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
        let mut xml = String::new();
        if let Some((horizontal, vertical)) = self.active_anchors() {
            let (horizontal_anchor, vertical_anchor) = anchors(horizontal, vertical);
            xml.push_str(&format!("<w:tblpPr w:horzAnchor=\"{}\" w:vertAnchor=\"{}\" {} {} w:leftFromText=\"{}\" w:topFromText=\"{}\" w:rightFromText=\"{}\" w:bottomFromText=\"{}\"/>",
                horizontal_anchor, vertical_anchor,
                coordinate("X", self.x), coordinate("Y", self.y),
                self.distances[0], self.distances[1], self.distances[2], self.distances[3]));
        }
        if self.no_overlap {
            xml.push_str("<w:tblOverlap w:val=\"never\"/>");
        }
        xml
    }

    #[cfg(feature = "direct-doc")]
    pub(in crate::doc) fn direct(&self) -> (Option<docx_model::TblpPr>, Option<String>) {
        let position = self.active_anchors().map(|(horizontal, vertical)| {
            let (tblp_x, tblp_x_spec) = direct_coordinate("X", self.x);
            let (tblp_y, tblp_y_spec) = direct_coordinate("Y", self.y);
            let (horizontal_anchor, vertical_anchor) = anchors(horizontal, vertical);
            docx_model::TblpPr {
                left_from_text: f64::from(self.distances[0]) / 20.0,
                top_from_text: f64::from(self.distances[1]) / 20.0,
                right_from_text: f64::from(self.distances[2]) / 20.0,
                bottom_from_text: f64::from(self.distances[3]) / 20.0,
                horz_anchor: horizontal_anchor.into(),
                // The canonical XML adapter always emits horzAnchor and an X
                // coordinate for every retained tblpPr, including Y-only and
                // wrap-distance-only DOC inputs. Match that authored wire.
                horz_specified: true,
                vert_anchor: vertical_anchor.into(),
                tblp_x,
                tblp_y,
                tblp_x_spec,
                tblp_y_spec,
            }
        });
        (position, self.no_overlap.then(|| "never".into()))
    }

    /// MS-DOC 2.7.13 Copts: nondefault position/wrapping facts create tblpPr;
    /// no-overlap alone does not. Reserved anchor values suppress placement.
    fn active_anchors(&self) -> Option<(u8, u8)> {
        let pc = self.anchors.unwrap_or(0);
        let vertical = (pc >> 4) & 3;
        let horizontal = pc >> 6;
        let active = self.anchors.is_some()
            || self.x != 0
            || self.y != 0
            || self.distances.iter().any(|value| *value != 0);
        (active && vertical != 3 && horizontal != 3).then_some((horizontal, vertical))
    }
}

fn anchors(horizontal: u8, vertical: u8) -> (&'static str, &'static str) {
    (
        ["text", "margin", "page"][horizontal as usize],
        ["margin", "page", "text"][vertical as usize],
    )
}

#[cfg(feature = "direct-doc")]
fn direct_coordinate(axis: &str, value: i32) -> (f64, Option<String>) {
    match resolved_coordinate(axis, value) {
        Coordinate::Spec(value) => (0.0, Some(value.into())),
        Coordinate::Offset(value) => (f64::from(value) / 20.0, None),
    }
}

fn coordinate(axis: &str, value: i32) -> String {
    match resolved_coordinate(axis, value) {
        Coordinate::Spec(value) => format!("w:tblp{axis}Spec=\"{value}\""),
        Coordinate::Offset(value) => format!("w:tblp{axis}=\"{value}\""),
    }
}

enum Coordinate {
    Spec(&'static str),
    Offset(i32),
}

fn resolved_coordinate(axis: &str, value: i32) -> Coordinate {
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
    match special {
        Some(value) => Coordinate::Spec(value),
        // MS-DOC 2.9.351/357: distances are stored one greater than twips.
        None => Coordinate::Offset(value - 1),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[cfg(feature = "direct-doc")]
    use std::io::{Cursor, Write};
    #[cfg(feature = "direct-doc")]
    use zip::write::SimpleFileOptions;

    #[cfg(feature = "direct-doc")]
    fn parsed_table(xml: &str) -> serde_json::Value {
        let document = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblPr>{xml}</w:tblPr><w:tblGrid/><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:body></w:document>"#
        );
        let mut bytes = Vec::new();
        {
            let mut archive = zip::ZipWriter::new(Cursor::new(&mut bytes));
            archive
                .start_file("word/document.xml", SimpleFileOptions::default())
                .unwrap();
            archive.write_all(document.as_bytes()).unwrap();
            archive.finish().unwrap();
        }
        let json: serde_json::Value =
            serde_json::from_str(&docx_parser::parse_docx_native(&bytes).unwrap()).unwrap();
        json["body"][0].clone()
    }
    #[cfg(feature = "direct-doc")]
    #[test]
    fn direct_position_uses_the_serializer_mapping_and_preserves_disabled_authorship() {
        let mut p = Position::default();
        p.apply(0x360d, &[0x21]).unwrap();
        p.apply(0x940e, &721i16.to_le_bytes()).unwrap();
        p.apply(0x940f, &(-8i16).to_le_bytes()).unwrap();
        p.apply(0x9410, &20u16.to_le_bytes()).unwrap();
        p.apply(0x3465, &[1]).unwrap();
        let (position, overlap) = p.direct();
        let position = position.unwrap();
        assert_eq!(position.horz_anchor, "text");
        assert_eq!(position.vert_anchor, "text");
        assert_eq!(position.tblp_x, 36.0);
        assert_eq!(position.tblp_y_spec.as_deref(), Some("center"));
        assert_eq!(position.left_from_text, 1.0);
        assert!(position.horz_specified);
        assert_eq!(overlap.as_deref(), Some("never"));
        let parsed = parsed_table(&p.xml());
        assert_eq!(serde_json::to_value(position).unwrap(), parsed["tblpPr"]);
        assert_eq!(parsed["overlap"], "never");

        for code in [0x940f, 0x9410] {
            let mut canonical = Position::default();
            canonical.apply(code, &20i16.to_le_bytes()).unwrap();
            let position = canonical.direct().0.unwrap();
            assert!(position.horz_specified, "code {code:#x}");
            assert_eq!(
                serde_json::to_value(position).unwrap(),
                parsed_table(&canonical.xml())["tblpPr"],
                "code {code:#x}"
            );
        }

        let mut disabled = Position::default();
        disabled.apply(0x360d, &[0xf0]).unwrap();
        disabled.apply(0x3465, &[1]).unwrap();
        assert!(disabled.direct().0.is_none());
        assert_eq!(disabled.direct().1.as_deref(), Some("never"));
    }
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
