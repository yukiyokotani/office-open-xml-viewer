//! Native (direct-model) admission of table-style selection and of the table
//! preference SPRMs that have no separate projection.
//!
//! [MS-DOC] 2.6.3 sprmTIstd selects a table style, "fetch[es] the complete set
//! of table properties from that style" and applies them "while preserving the
//! previous values" of an explicit list (revision, direction, rsid, position
//! and wrapping, gap, row height, preferred table width, autofit and the TLP
//! grfatl). Every other table property authored before the TIstd is replaced.
//! The style's TAPX/PAPX/CHPX are validated and projected by the table-style
//! profile, which retains its own per-property gates. This module decides
//! only whether the row's own Prl order around TIstd is one whose result is
//! established:
//!
//! * a property on the preserved list keeps its value in either order;
//! * a property whose replacement is implemented by an applier in this module
//!   family (row alignment, header, cant-split, no-overlap, default/cell
//!   margins, modern borders, shading) is reset by that applier;
//! * the cell geometry carried by TDxaLeft, TDefTable, TInsert, TDxaCol and the
//!   D635 preferred cell width is NOT replaced: current Word controls recorded
//!   under DOC-44, DOC-129, DOC-184 and DOC-193 show it surviving TIstd.
//!
//! Any other table property authored before a TIstd keeps the table gated,
//! because its replacement is neither implemented nor observed.

use super::{range, PreferredWidth, Row};
use crate::doc::unsupported;

const T_ISTD: u16 = 0x563a;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(in crate::doc) enum NativeAdmissionApply {
    Unhandled,
    Handled,
    HandledUnsupported,
}

/// Table SPRMs whose value is determined when a later TIstd applies. See the
/// module documentation for the three evidence classes.
const ESTABLISHED_BEFORE_TISTD: [u16; 38] = [
    // [MS-DOC] 2.6.3 preserved properties (sprmTWall/sprmTPropRMark are
    // revision state with their own gates and are deliberately absent).
    0x560b, // sprmTFBiDi
    0x7479, // sprmTRsid
    0x360d, // sprmTPc
    0x940e, // sprmTDxaAbs
    0x940f, // sprmTDyaAbs
    0x9410, // sprmTDxaFromText
    0x9411, // sprmTDyaFromText
    0x941e, // sprmTDxaFromTextRight
    0x941f, // sprmTDyaFromTextBottom
    0x9602, // sprmTDxaGapHalf
    0x9407, // sprmTDyaRowHeight
    0xf614, // sprmTTableWidth
    0x3615, // sprmTFAutofit
    0x740a, // sprmTTlp
    // Replaced properties whose reset is implemented by the native appliers.
    0x5400, // sprmTJc90
    0x548a, // sprmTJc
    0x3404, // sprmTTableHeader
    0x3466, // sprmTFCantSplit
    0x3403, // sprmTFCantSplit90 (ignored by current Word, see apply_native_cant_split)
    0x3465, // sprmTFNoAllowOverlap
    0xd632, // sprmTCellPadding
    0xd634, // sprmTCellPaddingDefault
    0xd613, // sprmTTableBorders
    0xd62f, // sprmTSetBrc
    0xd609, // sprmTDefTableShd80 (ignored by style-capable readers)
    0xd612, // sprmTDefTableShd
    0xd616, // sprmTDefTableShd2nd
    0xd60c, // sprmTDefTableShd3rd
    0xd660, // sprmTSetShdTable
    0xd670, // sprmTDefTableShdRaw
    0xd671, // sprmTDefTableShdRaw2nd
    0xd672, // sprmTDefTableShdRaw3rd
    // Cell geometry observed to survive TIstd in current Word.
    0x9601, // sprmTDxaLeft
    0xd608, // sprmTDefTable
    0x7621, // sprmTInsert
    0x7623, // sprmTDxaCol
    0xd635, // sprmTCellWidth
    T_ISTD,
];

/// Shading records whose TIstd reset is implemented only by the style-aware
/// shading applier (nFib > 0x00D9 readers, [MS-DOC] 2.6.3). For older files
/// the compatibility arrays remain live cell shading that TIstd does not reset.
const STYLE_AWARE_SHADING: [u16; 8] = [
    0xd609, 0xd612, 0xd616, 0xd60c, 0xd660, 0xd670, 0xd671, 0xd672,
];

/// Per-row-chain state for the native admission decisions. The caller creates
/// one value per `table_properties_native` acquisition and presents every Prl
/// to [`Self::observe`] in application order, before any applier consumes it.
#[derive(Default)]
pub(in crate::doc) struct NativeAdmission {
    style_aware_shading: bool,
    unestablished_before_tistd: bool,
    last_tistd_blocked: bool,
}

impl NativeAdmission {
    pub(in crate::doc) fn new(style_aware_shading: bool) -> Self {
        Self {
            style_aware_shading,
            ..Self::default()
        }
    }

    pub(in crate::doc) fn observe(&mut self, code: u16) {
        if (code >> 10) & 7 != 5 {
            return;
        }
        if code == T_ISTD {
            self.last_tistd_blocked = self.unestablished_before_tistd;
        } else if !ESTABLISHED_BEFORE_TISTD.contains(&code)
            || (!self.style_aware_shading && STYLE_AWARE_SHADING.contains(&code))
        {
            self.unestablished_before_tistd = true;
        }
    }

    pub(in crate::doc) fn apply(
        &mut self,
        row: &mut Row,
        code: u16,
        operand: &[u8],
    ) -> Result<NativeAdmissionApply, String> {
        match code {
            T_ISTD => {
                // TC80 cell geometry survives TIstd in current Word (module
                // documentation), but no control shows whether a TC80 textFlow
                // (2.9.317 TCGRF) authored before the selection survives or is
                // replaced: a style cannot carry text flow (2.9.340 excludes
                // sprmTTextFlow from UpxTapx). Keep that order gated.
                let text_flow_before = row.cells.iter().any(|cell| cell_text_flow(cell.flags) != 0);
                // Row::apply records the last selection and discards the
                // prepared shading layers; it keeps returning false for the
                // XML conversion, which does not interpret table styles.
                row.apply(code, operand)?;
                Ok(if self.last_tistd_blocked || text_flow_before {
                    NativeAdmissionApply::HandledUnsupported
                } else {
                    NativeAdmissionApply::Handled
                })
            }
            0x740a => {
                // [MS-DOC] 2.9.326: itl is historical auto-format metadata;
                // grfatl selects the conditional formats (2.9.69) consumed by
                // table_style_condition::Options. Row::apply validates the TLP.
                row.apply(code, operand)?;
                Ok(NativeAdmissionApply::Handled)
            }
            // [MS-DOC] 2.6.3 sprmTRsid: a revision save ID (ECMA-376 Part 1
            // 17.15.1.70) associated with table formatting. It has no
            // presentation semantics; the Sprm grammar fixes its 4-byte size.
            0x7479 => Ok(NativeAdmissionApply::Handled),
            0xf661 => {
                let indent = PreferredIndent::read(operand)?;
                row.preferred_indent = Some(indent);
                Ok(NativeAdmissionApply::Handled)
            }
            0xf617 | 0xf618 => {
                let value = PreferredWidth::part(operand)?;
                if code == 0xf617 {
                    row.preferred_before = Some(value);
                } else {
                    row.preferred_after = Some(value);
                }
                Ok(NativeAdmissionApply::Handled)
            }
            0xd639 => {
                // [MS-DOC] 2.9.28 CellRangeNoWrap: cb MUST be 3.
                if operand.len() != 4 || operand[0] != 3 {
                    return Err(unsupported("invalid Word cell no-wrap operand"));
                }
                let value = match operand[3] {
                    0 => false,
                    1 => true,
                    _ => return Err(unsupported("invalid Word cell no-wrap boolean")),
                };
                let cells = range(&operand[1..], row.cells.len())?;
                for cell in &mut row.cells[cells] {
                    cell.no_wrap = value;
                }
                Ok(NativeAdmissionApply::Handled)
            }
            0xd642 => {
                // [MS-DOC] 2.9.26 CellHideMarkOperand: cb MUST be 3, then an
                // ItcFirstLim and a Bool8. The projection is ECMA-376
                // §17.4.21 hideMark, which Word applies per cell (see the
                // DOCX table layout): its PDFs of sample-26 drop a hideMark
                // cell's final empty paragraph although the row has content.
                if operand.len() != 4 || operand[0] != 3 {
                    return Err(unsupported("invalid Word cell hide-mark operand"));
                }
                let value = match operand[3] {
                    0 => false,
                    1 => true,
                    _ => return Err(unsupported("invalid Word cell hide-mark boolean")),
                };
                let cells = range(&operand[1..], row.cells.len())?;
                for cell in &mut row.cells[cells] {
                    cell.hide_mark = value;
                }
                Ok(NativeAdmissionApply::Handled)
            }
            0x7629 => {
                // [MS-DOC] 2.6.3 sprmTTextFlow: a CellRangeTextFlow (2.9.29),
                // an ItcFirstLim (2.9.123) followed by a 2-byte TextFlow
                // (2.9.323), with no cb prefix. The value lives in the same
                // TCGRF textFlow field (2.9.317) that a TC80 authors, so a
                // later TDefTable replaces it and projection reads one place.
                let [first, lim, low, high] = operand else {
                    return Err(unsupported("invalid Word cell text flow operand"));
                };
                let text_flow = u16::from_le_bytes([*low, *high]);
                if !matches!(text_flow, 0 | 1 | 3 | 4 | 5) {
                    return Err(unsupported("invalid Word cell text flow value"));
                }
                let cells = range(&[*first, *lim], row.cells.len())?;
                for cell in &mut row.cells[cells] {
                    cell.flags = (cell.flags & !TCGRF_TEXT_FLOW) | (text_flow << 2);
                }
                Ok(NativeAdmissionApply::Handled)
            }
            _ => Ok(NativeAdmissionApply::Unhandled),
        }
    }
}

/// [MS-DOC] 2.9.317 TCGRF bits 2..=4: the cell's TextFlow (2.9.323).
const TCGRF_TEXT_FLOW: u16 = 7 << 2;

pub(in crate::doc) fn cell_text_flow(flags: u16) -> u16 {
    (flags & TCGRF_TEXT_FLOW) >> 2
}

/// [MS-DOC] 2.9.102 FtsWWidth_Indent, the preferred leading indent written by
/// sprmTWidthIndent.
///
/// The displayed horizontal origin of a DOC table is the physical one given by
/// sprmTDxaLeft/sprmTDxaGapHalf or the first TDefTable boundary (2.6.3: "the
/// location of the horizontal origin of the table"), which the direct model
/// already projects as the table indent. Word's own PDF exports of two
/// left-to-right documents whose preferred indent differs from that origin
/// (once by exactly the left default cell margin, once with nil margins) place
/// the table borders at the physical origin, while the paired OOXML documents
/// carry the preferred value as `w:tblInd` and display at that value.
/// The preference is therefore validated and retained, but does not replace
/// the physical origin. Right-to-left tables are not covered by that evidence
/// and stay gated at projection.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(in crate::doc) enum PreferredIndent {
    Nil,
    Auto,
    Dxa(i16),
}

impl PreferredIndent {
    pub(in crate::doc) fn read(bytes: &[u8]) -> Result<Self, String> {
        let [fts, low, high] = bytes else {
            return Err(unsupported("invalid Word preferred indent length"));
        };
        let value = i16::from_le_bytes([*low, *high]);
        match fts {
            0 | 1 if value != 0 => Err(unsupported("nonzero Word nil or automatic indent")),
            0 => Ok(Self::Nil),
            1 => Ok(Self::Auto),
            3 if (-31_560..=31_680).contains(&i32::from(value)) => Ok(Self::Dxa(value)),
            3 => Err(unsupported("Word preferred indent outside range")),
            // ftsPercent and ftsDxaSys are explicitly disallowed.
            _ => Err(unsupported("invalid Word preferred indent unit")),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::doc::table::Cell;

    fn row(cells: usize) -> Row {
        Row {
            cells: vec![Cell::default(); cells],
            ..Row::default()
        }
    }

    #[test]
    fn tistd_is_admitted_after_preserved_reset_and_geometry_properties_only() {
        for before in ESTABLISHED_BEFORE_TISTD {
            let mut admission = NativeAdmission::new(true);
            let mut target = row(1);
            admission.observe(before);
            admission.observe(T_ISTD);
            assert_eq!(
                admission
                    .apply(&mut target, T_ISTD, &3u16.to_le_bytes())
                    .unwrap(),
                NativeAdmissionApply::Handled,
                "{before:#06x}"
            );
            assert_eq!(target.table_style, Some(3));
        }
        // Vertical alignment is replaced by TIstd, but that replacement is not
        // implemented; no control shows whether TIstd resets a preceding
        // hide-mark, which a style cannot carry (2.9.340).
        for before in [0xd62c, 0xd642, 0xf661, 0xd639, 0x5622, 0xd605] {
            let mut admission = NativeAdmission::new(true);
            let mut target = row(1);
            admission.observe(before);
            admission.observe(T_ISTD);
            assert_eq!(
                admission
                    .apply(&mut target, T_ISTD, &3u16.to_le_bytes())
                    .unwrap(),
                NativeAdmissionApply::HandledUnsupported,
                "{before:#06x}"
            );
        }
        // The same properties after the selection are not replaced by it.
        let mut admission = NativeAdmission::new(true);
        let mut target = row(1);
        admission.observe(T_ISTD);
        assert_eq!(
            admission
                .apply(&mut target, T_ISTD, &3u16.to_le_bytes())
                .unwrap(),
            NativeAdmissionApply::Handled
        );
        admission.observe(0xd62c);
        // Before Word 2000 the compatibility shading arrays are live cell
        // shading, whose replacement by TIstd is not implemented.
        let mut admission = NativeAdmission::new(false);
        admission.observe(0xd612);
        admission.observe(T_ISTD);
        assert!(admission.last_tistd_blocked);
        // Paragraph SPRMs in the same chain never influence the decision.
        let mut admission = NativeAdmission::new(true);
        admission.observe(0x2416);
        admission.observe(T_ISTD);
        assert!(!admission.last_tistd_blocked);
    }

    #[test]
    fn each_selection_is_decided_by_the_properties_preceding_it() {
        let mut admission = NativeAdmission::new(true);
        let mut target = row(1);
        admission.observe(T_ISTD);
        admission.observe(0xd62c);
        admission.observe(T_ISTD);
        assert_eq!(
            admission
                .apply(&mut target, T_ISTD, &4u16.to_le_bytes())
                .unwrap(),
            NativeAdmissionApply::HandledUnsupported
        );
    }

    #[test]
    fn style_options_and_revision_ids_are_admitted_without_projection() {
        let mut admission = NativeAdmission::new(true);
        let mut target = row(1);
        assert_eq!(
            admission
                .apply(&mut target, 0x740a, &[0xff, 0xff, 0xa0, 0x04])
                .unwrap(),
            NativeAdmissionApply::Handled
        );
        assert_eq!(target.table_style_options, Some(0x04a0));
        assert!(admission.apply(&mut target, 0x740a, &[0, 0, 0]).is_err());
        assert_eq!(
            admission.apply(&mut target, 0x7479, &[1, 2, 3, 4]).unwrap(),
            NativeAdmissionApply::Handled
        );
    }

    #[test]
    fn hide_mark_applies_to_its_cell_range_and_validates_the_operand() {
        let mut admission = NativeAdmission::new(true);
        let mut target = row(3);
        assert_eq!(
            admission.apply(&mut target, 0xd642, &[3, 0, 2, 1]).unwrap(),
            NativeAdmissionApply::Handled
        );
        assert_eq!(
            target
                .cells
                .iter()
                .map(|cell| cell.hide_mark)
                .collect::<Vec<_>>(),
            [true, true, false]
        );
        admission.apply(&mut target, 0xd642, &[3, 1, 2, 0]).unwrap();
        assert!(!target.cells[1].hide_mark);
        for invalid in [
            vec![3, 0, 1],
            vec![2, 0, 1, 1],
            vec![3, 0, 1, 2],
            vec![3, 2, 1, 1],
            vec![3, 0, 4, 1],
        ] {
            assert!(
                admission.apply(&mut target, 0xd642, &invalid).is_err(),
                "{invalid:?}"
            );
        }
    }

    #[test]
    fn preferred_indent_follows_the_ftswwidth_indent_domain() {
        assert_eq!(
            PreferredIndent::read(&[0, 0, 0]).unwrap(),
            PreferredIndent::Nil
        );
        assert_eq!(
            PreferredIndent::read(&[1, 0, 0]).unwrap(),
            PreferredIndent::Auto
        );
        assert_eq!(
            PreferredIndent::read(&[3, 0x1c, 0xfd]).unwrap(),
            PreferredIndent::Dxa(-740)
        );
        let [low, high] = (-31_560i16).to_le_bytes();
        assert_eq!(
            PreferredIndent::read(&[3, low, high]).unwrap(),
            PreferredIndent::Dxa(-31_560)
        );
        for invalid in [
            vec![0, 1, 0],
            vec![1, 0, 1],
            vec![2, 0, 0],
            vec![0x13, 0, 0],
            vec![3, 0xc1, 0x7b],
            vec![3, 0],
        ] {
            assert!(PreferredIndent::read(&invalid).is_err(), "{invalid:?}");
        }
        assert!(PreferredIndent::read(&[3, 0x47, 0x84]).is_err());

        let mut admission = NativeAdmission::new(true);
        let mut target = row(1);
        admission.apply(&mut target, 0xf661, &[3, 0x6d, 0]).unwrap();
        assert_eq!(target.preferred_indent, Some(PreferredIndent::Dxa(109)));
    }

    #[test]
    fn preferred_before_and_after_retain_table_part_widths() {
        let mut admission = NativeAdmission::new(true);
        let mut target = row(2);
        admission.apply(&mut target, 0xf617, &[0, 0, 0]).unwrap();
        admission
            .apply(&mut target, 0xf618, &[3, 0x18, 0x15])
            .unwrap();
        assert_eq!(target.preferred_before, Some(None));
        assert_eq!(
            target.preferred_after,
            Some(Some(PreferredWidth::Dxa(5400)))
        );
        assert!(admission
            .apply(&mut target, 0xf617, &[3, 0xc1, 0x7b])
            .is_err());
    }

    #[test]
    fn text_flow_sets_the_tcgrf_field_of_its_cell_range_and_validates_the_operand() {
        let mut admission = NativeAdmission::new(true);
        let mut target = row(3);
        target.cells[1].flags = 0x0183; // horzMerge 3, vertAlign 3: preserved
        admission.apply(&mut target, 0x7629, &[1, 3, 5, 0]).unwrap();
        assert_eq!(
            target
                .cells
                .iter()
                .map(|cell| cell_text_flow(cell.flags))
                .collect::<Vec<_>>(),
            [0, 5, 5]
        );
        assert_eq!(target.cells[1].flags & !TCGRF_TEXT_FLOW, 0x0183);
        admission.apply(&mut target, 0x7629, &[2, 3, 0, 0]).unwrap();
        assert_eq!(cell_text_flow(target.cells[2].flags), 0);
        for invalid in [
            vec![0, 1, 2, 0],
            vec![0, 1, 6, 0],
            vec![0, 1, 7, 0],
            vec![0, 1, 1, 1],
            vec![0, 4, 1, 0],
            vec![2, 1, 1, 0],
            vec![0, 1, 1],
            vec![4, 0, 1, 1, 0],
        ] {
            assert!(
                admission.apply(&mut target, 0x7629, &invalid).is_err(),
                "{invalid:?}"
            );
        }
    }

    #[test]
    fn tistd_after_authored_text_flow_stays_gated() {
        // sprmTTextFlow before TIstd is not an established property.
        let mut admission = NativeAdmission::new(true);
        admission.observe(0x7629);
        admission.observe(T_ISTD);
        let mut target = row(1);
        admission.apply(&mut target, 0x7629, &[0, 1, 1, 0]).unwrap();
        assert_eq!(
            admission
                .apply(&mut target, T_ISTD, &11u16.to_le_bytes())
                .unwrap(),
            NativeAdmissionApply::HandledUnsupported
        );
        // A TC80 textFlow (via TDefTable) before TIstd has no survival control.
        let mut admission = NativeAdmission::new(true);
        admission.observe(0xd608);
        admission.observe(T_ISTD);
        let mut target = row(1);
        target.cells[0].flags = 3 << 2;
        assert_eq!(
            admission
                .apply(&mut target, T_ISTD, &11u16.to_le_bytes())
                .unwrap(),
            NativeAdmissionApply::HandledUnsupported
        );
    }

    #[test]
    fn no_wrap_applies_to_its_cell_range_and_validates_the_operand() {
        let mut admission = NativeAdmission::new(true);
        let mut target = row(3);
        admission.apply(&mut target, 0xd639, &[3, 1, 3, 1]).unwrap();
        assert_eq!(
            target
                .cells
                .iter()
                .map(|cell| cell.no_wrap)
                .collect::<Vec<_>>(),
            [false, true, true]
        );
        admission.apply(&mut target, 0xd639, &[3, 2, 3, 0]).unwrap();
        assert!(!target.cells[2].no_wrap);
        for invalid in [
            vec![2, 0, 1],
            vec![3, 0, 1, 2],
            vec![3, 0, 4, 1],
            vec![3, 2, 1, 1],
        ] {
            assert!(
                admission.apply(&mut target, 0xd639, &invalid).is_err(),
                "{invalid:?}"
            );
        }
    }
}
