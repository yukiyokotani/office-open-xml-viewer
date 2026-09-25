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
    pub(in crate::doc) fn reset_no_overlap_at_tistd(&mut self) {
        self.no_overlap = false;
    }

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
                // Every retained position states its horizontal anchor and X
                // coordinate, including Y-only and wrap-distance-only DOC
                // inputs, where both take their MS-DOC defaults.
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

    /// Reject the active positions whose DOC display is not established.
    ///
    /// * A vertical position of zero is the ST_YAlign `inline` value
    ///   ([MS-DOC] 2.6.3 sprmTDyaAbs). Its meaning for an absolutely
    ///   positioned table is not specified.
    /// * Word ignores an OOXML `tblpPr` whose offsets are zero with a text
    ///   horizontal anchor and a non-text vertical anchor ([MS-OI29500]
    ///   2.1.162). Whether Word's DOC reader applies the same exception to the
    ///   equivalent DOC values (left or zero X, zero Y, column and margin/page
    ///   anchors) has not been observed.
    pub(in crate::doc) fn check_direct_floating(&self) -> Result<(), String> {
        let Some((horizontal, vertical)) = self.active_anchors() else {
            return Ok(());
        };
        if self.y == 0 {
            return Err(super::unsupported(
                "direct DOC model cannot position a table with inline vertical alignment",
            ));
        }
        if horizontal == 0 && vertical != 2 && matches!(self.x, 0 | 1) && self.y == 1 {
            return Err(super::unsupported(
                "direct DOC model cannot classify a zero-offset positioned table",
            ));
        }
        Ok(())
    }

    /// True when a table paragraph's frame properties (MS-DOC 2.6.2) repeat
    /// this positioned table's own placement: the same PositionCodeOperand
    /// anchors, the same XAS_plusOne/YAS_plusOne X and Y (including special
    /// values), automatic size, around-wrapping, the same no-overlap flag and
    /// frame distances equal to the table's leading (left and top) distances.
    ///
    /// [MS-DOC] 2.4.3 consults cell-paragraph frame properties for table
    /// identity only when neither row specifies nondefault table position
    /// properties, so a positioned table's own position is authoritative.
    /// Word writes the frame form as the pre-TAP (Word 97) mirror of the same
    /// table: in the paired OOXML documents of the corpus files carrying it,
    /// Word keeps only the table's tblpPr, and nested tables inside the
    /// positioned table carry the same frame values. A paragraph frame has
    /// one horizontal and one vertical distance; the mirrors observed carry
    /// the table's left and top distances (equal to the right and bottom in
    /// the symmetric cases, 0 against a right-only/bottom-only table).
    pub(in crate::doc) fn matches_cell_frame(
        &self,
        frame: crate::doc::paragraph::TableParagraphFrame,
    ) -> bool {
        let Some((horizontal, vertical)) = self.active_anchors() else {
            return false;
        };
        frame
            .position_code
            .is_some_and(|code| (code >> 6, (code >> 4) & 3) == (horizontal, vertical))
            && i32::from(frame.dxa_abs) == self.x
            && i32::from(frame.dya_abs) == self.y
            && i32::from(frame.dxa_from_text) == self.distances[0]
            && i32::from(frame.dya_from_text) == self.distances[1]
            && frame.no_allow_overlap == self.no_overlap
            && frame.auto_size
            && frame.wrap == 2
            && !frame.drop_cap_or_text_flow
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

fn direct_coordinate(axis: &str, value: i32) -> (f64, Option<String>) {
    match resolved_coordinate(axis, value) {
        Coordinate::Spec(value) => (0.0, Some(value.into())),
        Coordinate::Offset(value) => (f64::from(value) / 20.0, None),
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
    #[test]
    fn direct_position_maps_anchors_coordinates_and_preserves_disabled_authorship() {
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
        assert_eq!(
            serde_json::to_value(position).unwrap(),
            serde_json::json!({
                "bottomFromText": 0.0, "horzAnchor": "text", "horzSpecified": true,
                "leftFromText": 1.0, "rightFromText": 0.0, "tblpX": 36.0, "tblpY": 0.0,
                "tblpYSpec": "center", "topFromText": 0.0, "vertAnchor": "text"
            })
        );

        // A Y-only or distance-only position still retains the horizontal
        // anchor and an X coordinate (the default margin/left placement).
        for (code, expected) in [
            (
                0x940f,
                serde_json::json!({
                    "bottomFromText": 0.0, "horzAnchor": "text", "horzSpecified": true,
                    "leftFromText": 0.0, "rightFromText": 0.0, "tblpX": 0.0,
                    "tblpXSpec": "left", "tblpY": 0.95, "topFromText": 0.0,
                    "vertAnchor": "margin"
                }),
            ),
            (
                0x9410,
                serde_json::json!({
                    "bottomFromText": 0.0, "horzAnchor": "text", "horzSpecified": true,
                    "leftFromText": 1.0, "rightFromText": 0.0, "tblpX": 0.0,
                    "tblpXSpec": "left", "tblpY": 0.0, "tblpYSpec": "inline",
                    "topFromText": 0.0, "vertAnchor": "margin"
                }),
            ),
        ] {
            let mut canonical = Position::default();
            canonical.apply(code, &20i16.to_le_bytes()).unwrap();
            let position = canonical.direct().0.unwrap();
            assert_eq!(
                serde_json::to_value(position).unwrap(),
                expected,
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
            let position = p.direct().0;
            if x == 3 || y == 3 {
                assert!(position.is_none());
            } else {
                let position = position.unwrap();
                assert_eq!(position.horz_anchor, ["text", "margin", "page"][x as usize]);
                assert_eq!(position.vert_anchor, ["margin", "page", "text"][y as usize]);
            }
        }
        assert!(Position::default().direct().0.is_none());
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
                    direct_coordinate(axis, -(i as i32) * 4),
                    (0.0, Some(label.to_string()))
                );
            }
            for n in [-31679, -21, -3, -1, 1, 2, 721, 31681] {
                assert_eq!(direct_coordinate(axis, n), (f64::from(n - 1) / 20.0, None));
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
        let position = p.direct().0.unwrap();
        assert_eq!(
            (
                position.left_from_text,
                position.top_from_text,
                position.right_from_text,
                position.bottom_from_text
            ),
            (0.6, 1.7, 2.8, 3.9)
        );
        for code in [0x940e, 0x940f] {
            for n in [-32768i16, -31680, 31682, 32767] {
                assert!(p.apply(code, &n.to_le_bytes()).is_err());
            }
        }
        let mut p = Position::default();
        p.apply(0x3465, &[1]).unwrap();
        // No-overlap alone does not create a positioned table.
        assert!(p.direct().0.is_none());
        assert_eq!(p.direct().1.as_deref(), Some("never"));
        p.apply(0x3465, &[0]).unwrap();
        assert!(p.direct().0.is_none() && p.direct().1.is_none());
        assert!(p.apply(0x3465, &[2]).is_err());
    }
}
