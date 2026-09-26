//! Bounded native acquisition for explicit DOC table cell geometry.
//!
//! [MS-DOC] 2.6.3 and 2.9.322 distinguish widths supplied when a cell is
//! created from later `sprmTDxaCol` overrides. Current Word controls establish
//! that such overrides survive a later same-count `sprmTDefTable` in the
//! fixed, unmerged profile retained here. Other structural interactions remain
//! admission-gated rather than being inferred from raw Prl order.
//! Current Word controls also establish last-range-wins vertical alignment,
//! including explicit top resets and conflicts with later same-count TC80
//! descriptor alignments, across first/last/disjoint/overlapping three-cell
//! ranges. Paired
//! controls without TDxaCol establish that alignment persistence is independent
//! of width overrides; a TC80-only control confirms the descriptor is active.

use super::{nonnegative, range, u16_at, unsupported, Row};

const MAX_CELLS: usize = 63;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(in crate::doc) enum NativeGeometryApply {
    Unhandled,
    Handled,
    HandledUnsupported,
}

pub(in crate::doc) struct NativeGeometry {
    overrides: [Option<i32>; MAX_CELLS],
    alignments: [Option<u8>; MAX_CELLS],
    count: Option<usize>,
    eligible: bool,
    unresolved_overrides: bool,
    prior_source_overrides: bool,
    preferred_competition: bool,
    persistence_used: bool,
}

impl Default for NativeGeometry {
    fn default() -> Self {
        Self {
            overrides: [None; MAX_CELLS],
            alignments: [None; MAX_CELLS],
            count: None,
            eligible: false,
            unresolved_overrides: false,
            prior_source_overrides: false,
            preferred_competition: false,
            persistence_used: false,
        }
    }
}

impl NativeGeometry {
    pub(in crate::doc) fn begin_source(&mut self) {
        self.prior_source_overrides |= self.has_overrides() || self.unresolved_overrides;
        self.overrides.fill(None);
        self.alignments.fill(None);
        self.count = None;
        self.eligible = false;
        self.unresolved_overrides = false;
    }

    pub(in crate::doc) fn apply(
        &mut self,
        row: &mut Row,
        code: u16,
        operand: &[u8],
    ) -> Result<NativeGeometryApply, String> {
        match code {
            0xd608 => self.apply_definition(row, operand),
            0x7623 => self.apply_width(row, operand),
            0x3615 if operand.first() == Some(&0) => {
                row.apply(code, operand)?;
                Ok(NativeGeometryApply::Handled)
            }
            0x560b | 0x5664 if u16_at(operand, 0)? == 0 => {
                row.apply(code, operand)?;
                Ok(NativeGeometryApply::Handled)
            }
            0xd635 => {
                let competed =
                    self.has_overrides() || self.unresolved_overrides || self.persistence_used;
                row.apply(code, operand)?;
                self.preferred_competition |= competed;
                Ok(if self.persistence_used {
                    NativeGeometryApply::HandledUnsupported
                } else {
                    NativeGeometryApply::Handled
                })
            }
            0xd62b => {
                // MS-DOC 2.6.3 defines zero as unmerged. Native fixed-row
                // controls confirm clearing an already-unmerged cell is a no-op.
                let unchanged = operand.first() == Some(&2)
                    && operand.get(2) == Some(&0)
                    && row
                        .cells
                        .get(usize::from(operand[1]))
                        .is_some_and(|cell| cell.flags & (3 << 5) == 0);
                if !row.apply(code, operand)? {
                    return Ok(NativeGeometryApply::HandledUnsupported);
                }
                Ok(self.finish_structural_change(!unchanged))
            }
            0xd62c => {
                if operand.is_empty() {
                    return Err(unsupported("short Word cell alignment"));
                }
                if !row.apply(code, operand)? {
                    return Ok(NativeGeometryApply::HandledUnsupported);
                }
                let cells = range(&operand[1..], row.cells.len())?;
                if cells.is_empty() {
                    return Ok(NativeGeometryApply::Handled);
                }
                let alignment = operand[3]; // Row::apply validated cb and VerticalAlign.
                if row.cells.iter().any(|cell| cell.preferred.is_some()) {
                    self.preferred_competition = true;
                }
                if self.count == Some(row.cells.len()) && self.eligible {
                    for slot in &mut self.alignments[cells] {
                        *slot = Some(alignment);
                    }
                } else {
                    self.unresolved_overrides = true;
                }
                Ok(if self.persistence_used && self.count.is_none() {
                    NativeGeometryApply::HandledUnsupported
                } else {
                    NativeGeometryApply::Handled
                })
            }
            0x5624 | 0x5625 => {
                // MS-DOC 2.6.3 explicitly makes TSplit on unmerged cells a
                // no-op. Native controls also confirm an empty TMerge range;
                // other merge operations remain outside persistence support.
                let cells = range(operand, row.cells.len())?;
                let changed = if code == 0x5625 {
                    row.cells[cells.clone()]
                        .iter()
                        .any(|cell| cell.flags & 3 != 0)
                } else {
                    !cells.is_empty()
                };
                if !row.apply(code, operand)? {
                    return Ok(NativeGeometryApply::HandledUnsupported);
                }
                Ok(self.finish_structural_change(changed))
            }
            0x7621 | 0x5622 | 0x3615 | 0x560b | 0x5664 => {
                let had_overrides = self.has_overrides();
                if !row.apply(code, operand)? {
                    return Ok(NativeGeometryApply::HandledUnsupported);
                }
                self.overrides.fill(None);
                self.alignments.fill(None);
                self.eligible = false;
                self.unresolved_overrides |= had_overrides;
                Ok(if self.persistence_used {
                    NativeGeometryApply::HandledUnsupported
                } else {
                    NativeGeometryApply::Handled
                })
            }
            _ => Ok(NativeGeometryApply::Unhandled),
        }
    }

    fn finish_structural_change(&mut self, changed: bool) -> NativeGeometryApply {
        if !changed {
            return NativeGeometryApply::Handled;
        }
        let had_overrides = self.has_overrides();
        self.overrides.fill(None);
        self.alignments.fill(None);
        self.eligible = false;
        self.unresolved_overrides |= had_overrides;
        if self.persistence_used {
            NativeGeometryApply::HandledUnsupported
        } else {
            NativeGeometryApply::Handled
        }
    }

    fn has_overrides(&self) -> bool {
        self.overrides.iter().any(Option::is_some) || self.alignments.iter().any(Option::is_some)
    }

    fn apply_definition(
        &mut self,
        row: &mut Row,
        operand: &[u8],
    ) -> Result<NativeGeometryApply, String> {
        let count = usize::from(
            *operand
                .get(2)
                .ok_or_else(|| unsupported("short Word table definition"))?,
        );
        if count > MAX_CELLS {
            return Err(unsupported("too many Word table cells"));
        }
        let boundary_end = 3 + (count + 1) * 2;
        let descriptors = operand
            .get(boundary_end..)
            .ok_or_else(|| unsupported("short Word table boundaries"))?;
        // The bounded profile permits only TC80 vertical-alignment bits. Other
        // flags (merge, text flow, preferred sizing, and related categories)
        // remain unsupported. Borders are stored outside this flag word.
        let definition_eligible = !row.autofit
            && !row.bidi
            && descriptors.len() % 20 == 0
            && descriptors.chunks_exact(20).take(count).all(|tc| {
                let flags = u16_at(tc, 0).unwrap_or(u16::MAX);
                flags & !(3 << 7) == 0 && (flags >> 7) & 3 <= 2
            });
        let had_overrides = self.has_overrides();
        // Count changes stay gated: native 3-to-2-to-3 TDef controls retained
        // widths even on the temporarily absent slot, whereas explicit
        // TDelete/TInsert discarded widths on replacement cells. Truncating
        // the override array as if both operations deleted cells is unsound.
        let can_reapply = had_overrides
            && !self.prior_source_overrides
            && self.count == Some(count)
            && self.eligible
            && !self.preferred_competition
            && definition_eligible;

        row.apply(0xd608, operand)?;
        if can_reapply {
            for (cell, width) in row.cells.iter_mut().zip(self.overrides.iter().copied()) {
                if let Some(width) = width {
                    cell.width = width;
                }
            }
            for (cell, alignment) in row.cells.iter_mut().zip(self.alignments.iter().copied()) {
                if let Some(alignment) = alignment {
                    cell.flags = (cell.flags & !(3 << 7)) | (u16::from(alignment) << 7);
                }
            }
            self.persistence_used = true;
        }

        let unsupported_competition = self.prior_source_overrides
            || self.unresolved_overrides
            || self.preferred_competition
            || (had_overrides && !can_reapply);
        if !can_reapply {
            self.overrides.fill(None);
            self.alignments.fill(None);
        }
        self.count = Some(count);
        self.eligible = definition_eligible;
        self.unresolved_overrides = false;
        self.preferred_competition = false;
        self.prior_source_overrides = false;
        Ok(if unsupported_competition {
            NativeGeometryApply::HandledUnsupported
        } else {
            NativeGeometryApply::Handled
        })
    }

    fn apply_width(
        &mut self,
        row: &mut Row,
        operand: &[u8],
    ) -> Result<NativeGeometryApply, String> {
        let cells = range(operand, row.cells.len())?;
        let width = nonnegative(&operand[2..])?;
        row.apply(0x7623, operand)?;
        if row.cells.iter().any(|cell| cell.preferred.is_some()) {
            self.preferred_competition = true;
        }
        if self.count == Some(row.cells.len()) && self.eligible {
            for slot in &mut self.overrides[cells] {
                *slot = Some(width);
            }
        } else {
            self.unresolved_overrides = true;
        }
        Ok(if self.persistence_used && self.count.is_none() {
            NativeGeometryApply::HandledUnsupported
        } else {
            NativeGeometryApply::Handled
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn definition(boundaries: &[i16]) -> Vec<u8> {
        let count = boundaries.len() - 1;
        let mut operand = vec![0, 0, count as u8];
        for boundary in boundaries {
            operand.extend(boundary.to_le_bytes());
        }
        operand.resize(operand.len() + count * 20, 0);
        let cb = (operand.len() - 1) as u16;
        operand[..2].copy_from_slice(&cb.to_le_bytes());
        operand
    }

    fn definition_with_alignment(boundaries: &[i16], cell: usize, alignment: u8) -> Vec<u8> {
        let mut operand = definition(boundaries);
        let descriptor = 3 + boundaries.len() * 2 + cell * 20;
        operand[descriptor..descriptor + 2]
            .copy_from_slice(&(u16::from(alignment) << 7).to_le_bytes());
        operand
    }

    #[test]
    fn same_count_definition_preserves_composed_explicit_ranges() {
        let mut row = Row::default();
        let mut geometry = NativeGeometry::default();
        geometry.begin_source();
        geometry
            .apply(&mut row, 0xd608, &definition(&[0, 1000, 2000, 3000]))
            .unwrap();
        geometry.apply(&mut row, 0x7623, &[0, 2, 0xd0, 7]).unwrap();
        geometry
            .apply(&mut row, 0x7623, &[1, 3, 0xb8, 0xb])
            .unwrap();
        geometry
            .apply(&mut row, 0xd608, &definition(&[0, 2500, 5500, 10500]))
            .unwrap();
        assert_eq!(
            row.cells.iter().map(|cell| cell.width).collect::<Vec<_>>(),
            [2000, 3000, 3000]
        );
    }

    #[test]
    fn changed_count_and_source_boundary_are_explicitly_unsupported() {
        for source_boundary in [false, true] {
            let mut row = Row::default();
            let mut geometry = NativeGeometry::default();
            geometry.begin_source();
            geometry
                .apply(&mut row, 0xd608, &definition(&[0, 1000, 2000]))
                .unwrap();
            geometry.apply(&mut row, 0x7623, &[0, 1, 0xd0, 7]).unwrap();
            if source_boundary {
                geometry.begin_source();
            }
            let result = geometry
                .apply(
                    &mut row,
                    0xd608,
                    &definition(if source_boundary {
                        &[0, 1000, 2000]
                    } else {
                        &[0, 1000, 2000, 3000]
                    }),
                )
                .unwrap();
            assert_eq!(result, NativeGeometryApply::HandledUnsupported);
        }
    }

    #[test]
    fn structural_change_before_width_keeps_later_definition_unsupported() {
        let mut row = Row::default();
        let mut geometry = NativeGeometry::default();
        geometry.begin_source();
        geometry
            .apply(&mut row, 0xd608, &definition(&[0, 1000, 2000, 3000]))
            .unwrap();
        assert_eq!(
            geometry.apply(&mut row, 0x5624, &[0, 2]).unwrap(),
            NativeGeometryApply::Handled
        );
        geometry.apply(&mut row, 0x7623, &[0, 1, 0xd0, 7]).unwrap();
        assert_eq!(
            geometry
                .apply(&mut row, 0xd608, &definition(&[0, 1000, 2000, 3000]))
                .unwrap(),
            NativeGeometryApply::HandledUnsupported
        );
    }

    #[test]
    fn false_autofit_keeps_fixed_profile_and_true_only_gates_a_competing_definition() {
        for (value, expected) in [
            (0, NativeGeometryApply::Handled),
            (1, NativeGeometryApply::HandledUnsupported),
        ] {
            let mut row = Row::default();
            let mut geometry = NativeGeometry::default();
            geometry.begin_source();
            geometry
                .apply(&mut row, 0xd608, &definition(&[0, 1000, 2000]))
                .unwrap();
            geometry.apply(&mut row, 0x7623, &[0, 1, 0xd0, 7]).unwrap();
            assert_eq!(
                geometry.apply(&mut row, 0x3615, &[value]).unwrap(),
                NativeGeometryApply::Handled
            );
            assert_eq!(
                geometry
                    .apply(&mut row, 0xd608, &definition(&[0, 1000, 2000]))
                    .unwrap(),
                expected
            );
        }
    }

    #[test]
    fn vertical_alignment_ranges_persist_with_last_direct_value() {
        for (operations, expected) in [
            (vec![(1, 2, 1), (1, 2, 0)], [0, 0, 0]),
            (vec![(1, 2, 0), (1, 2, 1)], [0, 1, 0]),
            (vec![(0, 2, 1), (1, 3, 2)], [1, 2, 2]),
            (vec![(1, 3, 2), (0, 2, 1)], [1, 1, 2]),
            (vec![(0, 1, 1), (2, 3, 2)], [1, 0, 2]),
        ] {
            let mut row = Row::default();
            let mut geometry = NativeGeometry::default();
            geometry.begin_source();
            geometry
                .apply(&mut row, 0xd608, &definition(&[0, 1000, 2000, 3000]))
                .unwrap();
            for (first, limit, alignment) in operations {
                geometry
                    .apply(&mut row, 0xd62c, &[3, first, limit, alignment])
                    .unwrap();
            }
            assert_eq!(
                geometry
                    .apply(&mut row, 0xd608, &definition(&[0, 1500, 3500, 6000]))
                    .unwrap(),
                NativeGeometryApply::Handled
            );
            assert_eq!(
                row.cells
                    .iter()
                    .map(|cell| (cell.flags >> 7) & 3)
                    .collect::<Vec<_>>(),
                expected
            );
        }
    }

    #[test]
    fn direct_alignment_overrides_later_descriptor_alignment() {
        for (direct, descriptor_alignment, expected) in [(1, 2, 1), (0, 1, 0)] {
            let mut row = Row::default();
            let mut geometry = NativeGeometry::default();
            geometry.begin_source();
            geometry
                .apply(&mut row, 0xd608, &definition(&[0, 1000, 2000, 3000]))
                .unwrap();
            geometry
                .apply(&mut row, 0xd62c, &[3, 1, 2, direct])
                .unwrap();
            geometry
                .apply(
                    &mut row,
                    0xd608,
                    &definition_with_alignment(&[0, 1500, 3500, 6000], 1, descriptor_alignment),
                )
                .unwrap();
            assert_eq!((row.cells[1].flags >> 7) & 3, expected);
        }

        let mut row = Row::default();
        let mut geometry = NativeGeometry::default();
        geometry.begin_source();
        geometry
            .apply(&mut row, 0xd608, &definition(&[0, 1000, 2000]))
            .unwrap();
        geometry.apply(&mut row, 0xd62c, &[3, 1, 1, 1]).unwrap();
        assert_eq!(
            geometry
                .apply(&mut row, 0xd608, &definition(&[0, 1500, 3000]))
                .unwrap(),
            NativeGeometryApply::Handled
        );
        assert_eq!(
            geometry.apply(&mut row, 0xd62c, &[2, 0, 1]).unwrap(),
            NativeGeometryApply::HandledUnsupported
        );
        assert!(geometry.apply(&mut row, 0xd62c, &[3, 0, 1, 3]).is_err());

        let mut row = Row::default();
        let mut geometry = NativeGeometry::default();
        geometry.begin_source();
        assert_eq!(
            geometry.apply(&mut row, 0xd62c, &[3, 0, 0, 1]).unwrap(),
            NativeGeometryApply::Handled
        );

        geometry
            .apply(&mut row, 0xd608, &definition(&[0, 1000, 2000]))
            .unwrap();
        geometry
            .apply(&mut row, 0xd635, &[5, 0, 1, 3, 100, 0])
            .unwrap();
        geometry.apply(&mut row, 0xd62c, &[3, 0, 1, 1]).unwrap();
        assert_eq!(
            geometry
                .apply(&mut row, 0xd608, &definition(&[0, 1500, 3000]))
                .unwrap(),
            NativeGeometryApply::HandledUnsupported
        );
    }

    #[test]
    fn unmerged_split_and_clear_operations_do_not_invalidate_alignment() {
        for (code, operand) in [
            (0x5625, vec![0, 2]),
            (0x5624, vec![1, 1]),
            (0xd62b, vec![2, 0, 0]),
        ] {
            let mut row = Row::default();
            let mut geometry = NativeGeometry::default();
            geometry.begin_source();
            geometry
                .apply(&mut row, 0xd608, &definition(&[0, 1000, 2000]))
                .unwrap();
            geometry.apply(&mut row, 0xd62c, &[3, 0, 1, 1]).unwrap();
            assert_eq!(
                geometry.apply(&mut row, code, &operand).unwrap(),
                NativeGeometryApply::Handled
            );
            assert_eq!(
                geometry
                    .apply(&mut row, 0xd608, &definition(&[0, 1500, 3000]))
                    .unwrap(),
                NativeGeometryApply::Handled
            );
            assert_eq!((row.cells[0].flags >> 7) & 3, 1);
        }

        let mut row = Row::default();
        let mut geometry = NativeGeometry::default();
        geometry.begin_source();
        geometry
            .apply(&mut row, 0xd608, &definition(&[0, 1000, 2000]))
            .unwrap();
        geometry.apply(&mut row, 0xd62c, &[3, 0, 1, 1]).unwrap();
        geometry
            .apply(&mut row, 0xd608, &definition(&[0, 1500, 3000]))
            .unwrap();
        assert_eq!(
            geometry.apply(&mut row, 0x5625, &[0, 2]).unwrap(),
            NativeGeometryApply::Handled
        );
        assert_eq!((row.cells[0].flags >> 7) & 3, 1);
    }

    #[test]
    fn all_sixty_three_cells_are_bounded_and_ranges_remain_strict() {
        let boundaries = (0..=63).map(|value| value * 10).collect::<Vec<i16>>();
        let mut row = Row::default();
        let mut geometry = NativeGeometry::default();
        geometry.begin_source();
        geometry
            .apply(&mut row, 0xd608, &definition(&boundaries))
            .unwrap();
        geometry.apply(&mut row, 0x7623, &[0, 63, 20, 0]).unwrap();
        geometry
            .apply(&mut row, 0xd608, &definition(&boundaries))
            .unwrap();
        assert!(row.cells.iter().all(|cell| cell.width == 20));
        assert!(geometry.apply(&mut row, 0x7623, &[62, 64, 1, 0]).is_err());
    }

    #[test]
    fn only_post_persistence_competitors_gate_existing_simple_sequences() {
        fn prepared() -> (Row, NativeGeometry) {
            let mut row = Row::default();
            let mut geometry = NativeGeometry::default();
            geometry.begin_source();
            geometry
                .apply(&mut row, 0xd608, &definition(&[0, 1000, 2000]))
                .unwrap();
            geometry.apply(&mut row, 0x7623, &[0, 1, 0xd0, 7]).unwrap();
            (row, geometry)
        }

        let (mut simple_row, mut simple) = prepared();
        assert_eq!(
            simple.apply(&mut simple_row, 0x5624, &[0, 2]).unwrap(),
            NativeGeometryApply::Handled,
            "TDxa followed by merge keeps its prior acquisition behavior"
        );

        for (code, operand) in [
            (0x5624, vec![0, 2]),
            (0x3615, vec![1]),
            (0x560b, 1u16.to_le_bytes().to_vec()),
            (0xd635, vec![5, 0, 1, 3, 100, 0]),
        ] {
            let (mut row, mut geometry) = prepared();
            geometry
                .apply(&mut row, 0xd608, &definition(&[0, 1000, 2000]))
                .unwrap();
            assert_eq!(
                geometry.apply(&mut row, code, &operand).unwrap(),
                NativeGeometryApply::HandledUnsupported,
                "late competitor {code:#x}"
            );
        }

        let (mut row, mut geometry) = prepared();
        geometry
            .apply(&mut row, 0xd635, &[5, 0, 1, 3, 100, 0])
            .unwrap();
        assert_eq!(
            geometry
                .apply(&mut row, 0xd608, &definition(&[0, 1000, 2000]))
                .unwrap(),
            NativeGeometryApply::HandledUnsupported,
            "preferred width competing before repeated TDef remains gated"
        );

        let mut row = Row::default();
        let mut geometry = NativeGeometry::default();
        geometry.begin_source();
        geometry
            .apply(&mut row, 0xd608, &definition(&[0, 1000, 2000]))
            .unwrap();
        geometry
            .apply(&mut row, 0xd635, &[5, 0, 1, 3, 100, 0])
            .unwrap();
        assert_eq!(
            geometry
                .apply(&mut row, 0xd608, &definition(&[0, 1000, 2000]))
                .unwrap(),
            NativeGeometryApply::Handled,
            "preferred-width replacement without TDxa keeps prior behavior"
        );

        let (mut row, mut geometry) = prepared();
        geometry
            .apply(&mut row, 0xd608, &definition(&[0, 1000, 2000]))
            .unwrap();
        geometry.begin_source();
        assert_eq!(
            geometry.apply(&mut row, 0x5624, &[0, 2]).unwrap(),
            NativeGeometryApply::HandledUnsupported,
            "persistence remains sticky across source transitions"
        );
    }

    #[test]
    fn bidi_before_definition_cannot_enter_ltr_persistence_profile() {
        let mut row = Row::default();
        let mut geometry = NativeGeometry::default();
        geometry.begin_source();
        geometry
            .apply(&mut row, 0x560b, &1u16.to_le_bytes())
            .unwrap();
        geometry
            .apply(&mut row, 0xd608, &definition(&[0, 1000, 2000]))
            .unwrap();
        geometry.apply(&mut row, 0x7623, &[0, 1, 0xd0, 7]).unwrap();
        assert_eq!(
            geometry
                .apply(&mut row, 0xd608, &definition(&[0, 1000, 2000]))
                .unwrap(),
            NativeGeometryApply::HandledUnsupported
        );
    }
}
