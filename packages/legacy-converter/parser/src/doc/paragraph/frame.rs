//! MS-DOC 2.6.2 paragraph frame (positioned paragraph) properties.
//!
//! The raw SPRM facts are retained exactly as decoded so the direct projection
//! can decide, once the complete style/direct/list cascade is known, whether
//! they form an ECMA-376 17.3.1.11 `framePr` the DOCX model can represent.
//! Only sub-cases with a documented meaning are projected; the rest stay
//! fail-closed (see [`Frame::direct`]).

use super::super::{u16_at, unsupported};

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub(super) struct Frame {
    /// True once any frame SPRM was applied anywhere in the cascade.
    applied: bool,
    /// sprmPDxaAbs XAS_plusOne raw value (MS-DOC 2.9.351). Default 0 = left.
    dxa_abs: i16,
    /// sprmPDyaAbs YAS_plusOne raw value (MS-DOC 2.9.357). Default 0 = inline.
    dya_abs: i16,
    /// sprmPDxaWidth XAS_nonNeg; 0 = automatic width.
    width: u16,
    /// sprmPWHeightAbs WHeightAbs (MS-DOC 2.9.345) raw bits.
    height: u16,
    /// sprmPPc PositionCodeOperand (MS-DOC 2.9.208) raw byte. MS-DOC states
    /// no default anchor, so its absence is retained rather than assumed.
    position_code: Option<u8>,
    /// sprmPWr; MS-DOC 2.6.2 default 0 (ST_Wrap auto).
    wrap: u8,
    /// sprmPDxaFromText / sprmPDyaFromText XAS_nonNeg/YAS_nonNeg twips.
    dxa_from_text: u16,
    dya_from_text: u16,
    /// sprmPFLocked (ECMA-376 framePr@anchorLock).
    locked: bool,
    /// sprmPFNoAllowOverlap. The DOCX frame model has no counterpart.
    no_allow_overlap: bool,
    /// sprmPDcs DCS (MS-DOC 2.9.51) raw value.
    drop_cap: u16,
    /// sprmPFrameTextFlow. The DOCX frame model has no counterpart.
    text_flow: bool,
}

fn nonnegative_twips(operand: &[u8]) -> Result<u16, String> {
    let value = u16_at(operand, 0)?;
    // MS-DOC 2.9.350 XAS_nonNeg and 2.9.358 YAS_nonNeg: at most 31680.
    if value > 31680 {
        return Err(unsupported("invalid Word frame distance"));
    }
    Ok(value)
}

fn plus_one(operand: &[u8]) -> Result<i16, String> {
    let value = u16_at(operand, 0)? as i16;
    // MS-DOC 2.9.351/357: -31679..=31681. The special relative-position
    // codes (0 and small negative multiples of 4) lie inside this range.
    if !(-31679..=31681).contains(&value) {
        return Err(unsupported("invalid Word frame position"));
    }
    Ok(value)
}

fn bool8(operand: &[u8]) -> Result<bool, String> {
    match operand.first() {
        Some(0) => Ok(false),
        Some(1) => Ok(true),
        _ => Err(unsupported("invalid Word frame boolean")),
    }
}

impl Frame {
    /// Apply one frame SPRM. Returns `Ok(false)` for codes outside this set.
    pub(super) fn apply(&mut self, code: u16, operand: &[u8]) -> Result<bool, String> {
        match code {
            0x8418 => self.dxa_abs = plus_one(operand)?,
            0x8419 => self.dya_abs = plus_one(operand)?,
            0x841a => self.width = nonnegative_twips(operand)?,
            0x442b => {
                let raw = u16_at(operand, 0)?;
                let height = raw & 0x7fff;
                // MS-DOC 2.9.345: YAS_nonNeg height; fMinHeight requires a
                // nonzero height.
                if height > 31680 || (raw & 0x8000 != 0 && height == 0) {
                    return Err(unsupported("invalid Word frame height"));
                }
                self.height = raw;
            }
            0x261b => {
                // MS-DOC 2.9.208: the four padding bits MUST be ignored.
                self.position_code = Some(
                    *operand
                        .first()
                        .ok_or_else(|| unsupported("truncated Word frame anchor"))?,
                );
            }
            0x2423 => {
                let value = *operand
                    .first()
                    .ok_or_else(|| unsupported("truncated Word frame wrap"))?;
                if value > 5 {
                    return Err(unsupported("invalid Word frame wrap"));
                }
                self.wrap = value;
            }
            0x842f => self.dxa_from_text = nonnegative_twips(operand)?,
            0x842e => self.dya_from_text = nonnegative_twips(operand)?,
            0x2430 => self.locked = bool8(operand)?,
            0x2462 => self.no_allow_overlap = bool8(operand)?,
            0x442c => {
                let raw = u16_at(operand, 0)?;
                let kind = raw & 7;
                let lines = (raw >> 3) & 31;
                // MS-DOC 2.9.51: fdct 1 or 2 with cl 1..=10. A zero DCS is
                // the documented default (no drop cap); its reserved byte
                // MUST be ignored.
                if kind > 2 || (kind != 0 && !(1..=10).contains(&lines)) {
                    return Err(unsupported("invalid Word drop cap"));
                }
                self.drop_cap = raw & 0xff;
            }
            0x443a => {
                let _ = u16_at(operand, 0)?;
                self.text_flow = true;
            }
            _ => return Ok(false),
        }
        self.applied = true;
        Ok(true)
    }
}

/// Raw frame facts of a table paragraph, compared against the table's own
/// position by `table::Position::matches_cell_frame`.
#[cfg(feature = "direct-doc")]
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(in crate::doc) struct TableParagraphFrame {
    pub(in crate::doc) position_code: Option<u8>,
    pub(in crate::doc) dxa_abs: i16,
    pub(in crate::doc) dya_abs: i16,
    pub(in crate::doc) auto_size: bool,
    pub(in crate::doc) wrap: u8,
    pub(in crate::doc) dxa_from_text: u16,
    pub(in crate::doc) dya_from_text: u16,
    pub(in crate::doc) no_allow_overlap: bool,
    pub(in crate::doc) drop_cap_or_text_flow: bool,
}

#[cfg(feature = "direct-doc")]
impl Frame {
    pub(in crate::doc) fn table_paragraph_facts(&self) -> Option<TableParagraphFrame> {
        self.applied.then_some(TableParagraphFrame {
            position_code: self.position_code,
            dxa_abs: self.dxa_abs,
            dya_abs: self.dya_abs,
            auto_size: self.width == 0 && self.height == 0,
            wrap: self.wrap,
            dxa_from_text: self.dxa_from_text,
            dya_from_text: self.dya_from_text,
            no_allow_overlap: self.no_allow_overlap,
            drop_cap_or_text_flow: self.drop_cap & 7 != 0 || self.text_flow,
        })
    }
}

/// Why a frame cannot be projected. Every variant keeps the file fail-closed.
#[cfg(feature = "direct-doc")]
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(in crate::doc) enum FrameGap {
    /// MS-DOC 2.6.2 gives sprmPPc no default anchor.
    MissingAnchor,
    /// PositionCodeOperand value 3: "not absolutely positioned".
    UnpositionedAnchor,
    /// sprmPDyaAbs 0 (inline). The DOCX renderer resolves ST_YAlign inline
    /// as the band start; Word's inline frame placement is not documented.
    InlineVerticalPosition,
    /// sprmPFNoAllowOverlap = 1 or sprmPFrameTextFlow: no DOCX model field.
    UnrepresentableProperty,
}

#[cfg(feature = "direct-doc")]
impl Frame {
    /// Project the resolved frame facts onto ECMA-376 17.3.1.11 `framePr`
    /// semantics carried by `docx_model::FramePr`.
    ///
    /// Presence: ECMA-376 makes `framePr` presence the frame criterion. The
    /// binary format has no presence bit, so any applied frame SPRM marks the
    /// paragraph as framed. Resetting SPRMs that would restore every default
    /// are not interpreted as frame removal; they fail closed through the
    /// anchor/inline gaps below instead of guessing Word's criterion.
    pub(in crate::doc) fn direct(&self) -> Result<Option<docx_model::FramePr>, FrameGap> {
        if !self.applied {
            return Ok(None);
        }
        if self.no_allow_overlap || self.text_flow {
            return Err(FrameGap::UnrepresentableProperty);
        }
        let code = self.position_code.ok_or(FrameGap::MissingAnchor)?;
        // MS-DOC 2.9.208: bits 4-5 pcVert, bits 6-7 pcHorz.
        let vertical = (code >> 4) & 3;
        let horizontal = (code >> 6) & 3;
        let v_anchor = match vertical {
            // 0: relative to the top page margin; 1: the page top edge;
            // 2: the bottom of the preceding paragraph, i.e. the frame
            // paragraph's own flow position (ST_VAnchor text).
            0 => "margin",
            1 => "page",
            2 => "text",
            _ => return Err(FrameGap::UnpositionedAnchor),
        };
        let h_anchor = match horizontal {
            // 0: the current column (ST_HAnchor text); 1: the left page
            // margin; 2: the left page edge.
            0 => "text",
            1 => "margin",
            2 => "page",
            _ => return Err(FrameGap::UnpositionedAnchor),
        };
        let kind = self.drop_cap & 7;
        let drop_cap = match kind {
            1 => "drop",
            2 => "margin",
            _ => "none",
        };
        let lines = if kind == 0 {
            1
        } else {
            u32::from((self.drop_cap >> 3) & 31)
        };
        // MS-DOC 2.6.2 sprmPDxaAbs special values mirror ST_XAlign; the
        // documented default (0) is Left.
        let (x, x_align) = match self.dxa_abs {
            0 => (None, Some("left")),
            -4 => (None, Some("center")),
            -8 => (None, Some("right")),
            -12 => (None, Some("inside")),
            -16 => (None, Some("outside")),
            value => (Some(f64::from(i32::from(value) - 1) / 20.0), None),
        };
        // ECMA-376 17.3.1.11 ignores y/yAlign for a drop cap; only then is
        // the undocumented inline placement irrelevant.
        let (y, y_align) = match self.dya_abs {
            0 if kind != 0 => (None, None),
            0 => return Err(FrameGap::InlineVerticalPosition),
            -4 => (None, Some("top")),
            -8 => (None, Some("center")),
            -12 => (None, Some("bottom")),
            -16 => (None, Some("inside")),
            -20 => (None, Some("outside")),
            value => (Some(f64::from(i32::from(value) - 1) / 20.0), None),
        };
        let height = self.height & 0x7fff;
        let (h, h_rule) = if height == 0 {
            (None, "auto")
        } else if self.height & 0x8000 != 0 {
            (Some(f64::from(height) / 20.0), "atLeast")
        } else {
            (Some(f64::from(height) / 20.0), "exact")
        };
        Ok(Some(docx_model::FramePr {
            anchor_lock: self.locked,
            drop_cap: drop_cap.into(),
            lines,
            // MS-DOC 2.6.2 sprmPWr values correspond to ST_Wrap.
            wrap: ["auto", "notBeside", "around", "none", "tight", "through"]
                [usize::from(self.wrap)]
            .into(),
            h_anchor: h_anchor.into(),
            v_anchor: v_anchor.into(),
            h_rule: h_rule.into(),
            h_space: f64::from(self.dxa_from_text) / 20.0,
            v_space: f64::from(self.dya_from_text) / 20.0,
            w: (self.width != 0).then(|| f64::from(self.width) / 20.0),
            h,
            x,
            y,
            x_align: x_align.map(str::to_string),
            y_align: y_align.map(str::to_string),
        }))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn frame(entries: &[(u16, &[u8])]) -> Frame {
        let mut frame = Frame::default();
        for (code, operand) in entries {
            assert!(frame.apply(*code, operand).unwrap(), "{code:04x}");
        }
        frame
    }

    #[test]
    fn frame_operands_are_validated() {
        let mut value = Frame::default();
        assert!(!value.apply(0x2403, &[0]).unwrap());
        assert!(!value.applied);
        for (code, operand) in [
            (0x8418, (-31680i16).to_le_bytes().to_vec()),
            (0x8419, 31682i16.to_le_bytes().to_vec()),
            (0x841a, 31681u16.to_le_bytes().to_vec()),
            (0x842f, 31681u16.to_le_bytes().to_vec()),
            (0x842e, 31681u16.to_le_bytes().to_vec()),
            (0x442b, 0x8000u16.to_le_bytes().to_vec()),
            (0x442b, 31681u16.to_le_bytes().to_vec()),
            (0x2423, vec![6]),
            (0x2430, vec![2]),
            (0x2462, vec![2]),
            (0x442c, vec![3, 0]),
            (0x442c, vec![1 | (11 << 3), 0]),
            (0x442c, vec![1, 0]),
            (0x261b, vec![]),
            (0x8418, vec![1]),
        ] {
            assert!(
                Frame::default().apply(code, &operand).is_err(),
                "{code:04x}"
            );
        }
        // A zero DCS is the documented no-drop-cap default; reserved bits in
        // the second byte are ignored.
        assert!(Frame::default().apply(0x442c, &[0, 0xff]).unwrap());
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn positioned_frame_projects_documented_anchor_position_and_size_facts() {
        // pcVert=2 (paragraph -> text), pcHorz=1 (margin).
        let value = frame(&[
            (0x261b, &[0x60]),
            (0x2423, &[2]),
            (0x8418, &(-8i16).to_le_bytes()),
            (0x8419, &2971i16.to_le_bytes()),
            (0x841a, &2880u16.to_le_bytes()),
            (0x442b, &(0x8000u16 | 400).to_le_bytes()),
            (0x842f, &180u16.to_le_bytes()),
            (0x842e, &40u16.to_le_bytes()),
            (0x2430, &[1]),
        ])
        .direct()
        .unwrap()
        .unwrap();
        assert_eq!(
            (value.h_anchor.as_str(), value.v_anchor.as_str()),
            ("margin", "text")
        );
        assert_eq!(value.wrap, "around");
        assert_eq!((value.x, value.x_align.as_deref()), (None, Some("right")));
        assert_eq!((value.y, value.y_align.as_deref()), (Some(148.5), None));
        assert_eq!(
            (value.w, value.h, value.h_rule.as_str()),
            (Some(144.0), Some(20.0), "atLeast")
        );
        assert_eq!((value.h_space, value.v_space), (9.0, 2.0));
        assert!(value.anchor_lock);
        assert_eq!((value.drop_cap.as_str(), value.lines), ("none", 1));

        // pcVert=1 (page), pcHorz=2 (page); absolute signed positions are
        // decremented by one twip (MS-DOC 2.9.351/357).
        let value = frame(&[
            (0x261b, &[0x90]),
            (0x8418, &(-284i16).to_le_bytes()),
            (0x8419, &1i16.to_le_bytes()),
            (0x442b, &200u16.to_le_bytes()),
        ])
        .direct()
        .unwrap()
        .unwrap();
        assert_eq!(
            (value.h_anchor.as_str(), value.v_anchor.as_str()),
            ("page", "page")
        );
        assert_eq!(value.wrap, "auto");
        assert_eq!((value.x, value.y), (Some(-14.25), Some(0.0)));
        assert_eq!(
            (value.h, value.h_rule.as_str(), value.w),
            (Some(10.0), "exact", None)
        );
        for (raw, align) in [
            (0, "left"),
            (-4, "center"),
            (-12, "inside"),
            (-16, "outside"),
        ] {
            let value = frame(&[
                (0x261b, &[0x00]),
                (0x8418, &(raw as i16).to_le_bytes()),
                (0x8419, &(-20i16).to_le_bytes()),
            ])
            .direct()
            .unwrap()
            .unwrap();
            assert_eq!(value.x_align.as_deref(), Some(align));
            assert_eq!(value.y_align.as_deref(), Some("outside"));
            assert_eq!(
                (value.h_anchor.as_str(), value.v_anchor.as_str()),
                ("text", "margin")
            );
        }
        for (raw, align) in [
            (-4i16, "top"),
            (-8, "center"),
            (-12, "bottom"),
            (-16, "inside"),
        ] {
            let value = frame(&[(0x261b, &[0x50]), (0x8419, &raw.to_le_bytes())])
                .direct()
                .unwrap()
                .unwrap();
            assert_eq!(value.y_align.as_deref(), Some(align));
        }
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn drop_cap_ignores_vertical_position_and_records_lines() {
        let value = frame(&[(0x261b, &[0x20]), (0x2423, &[2]), (0x442c, &[0x19, 0])])
            .direct()
            .unwrap()
            .unwrap();
        assert_eq!((value.drop_cap.as_str(), value.lines), ("drop", 3));
        assert_eq!((value.y, value.y_align), (None, None));
        let value = frame(&[(0x261b, &[0x20]), (0x442c, &[2 | (10 << 3), 0])])
            .direct()
            .unwrap()
            .unwrap();
        assert_eq!((value.drop_cap.as_str(), value.lines), ("margin", 10));
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn implicit_or_unrepresentable_frames_stay_fail_closed() {
        assert!(Frame::default().direct().unwrap().is_none());
        assert_eq!(
            frame(&[(0x8419, &2i16.to_le_bytes())]).direct().err(),
            Some(FrameGap::MissingAnchor)
        );
        for code in [0x30u8, 0xc0] {
            assert_eq!(
                frame(&[(0x261b, &[code]), (0x8419, &2i16.to_le_bytes())])
                    .direct()
                    .err(),
                Some(FrameGap::UnpositionedAnchor)
            );
        }
        assert_eq!(
            frame(&[(0x261b, &[0x20])]).direct().err(),
            Some(FrameGap::InlineVerticalPosition)
        );
        assert_eq!(
            frame(&[
                (0x261b, &[0x20]),
                (0x8419, &2i16.to_le_bytes()),
                (0x2462, &[1])
            ])
            .direct()
            .err(),
            Some(FrameGap::UnrepresentableProperty)
        );
        assert!(frame(&[
            (0x261b, &[0x20]),
            (0x8419, &2i16.to_le_bytes()),
            (0x2462, &[0])
        ])
        .direct()
        .unwrap()
        .is_some());
        assert_eq!(
            frame(&[
                (0x261b, &[0x20]),
                (0x8419, &2i16.to_le_bytes()),
                (0x443a, &[0, 0])
            ])
            .direct()
            .err(),
            Some(FrameGap::UnrepresentableProperty)
        );
    }
}
