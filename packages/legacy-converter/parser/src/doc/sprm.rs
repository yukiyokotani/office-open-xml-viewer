//! Bounded framing shared by Word character and paragraph property readers.
//! [MS-DOC] 2.2.5 and 2.6 (Sprm); variable-length tab/table exceptions.

use super::{u16_at, u32_at, unsupported};
use std::collections::BTreeSet;

/// [MS-DOC] 2.9.210 PrcData limits cbGrpprl to this many bytes.
const MAX_PRC_DATA_GRPPRL_BYTES: usize = 0x3fa2;

#[derive(Clone, Copy)]
pub enum TopLevelFilter {
    All,
    Paragraph,
}

/// Traverse one property array. `Paragraph` treats `bytes` as piece properties
/// (see `paragraph_properties_appended`); `All` treats them as a direct array.
/// Share the traversal so paragraph layout and table structure see the same data.
pub fn paragraph_properties<'a>(
    bytes: &'a [u8],
    data: &'a [u8],
    budget: &mut Budget,
    top_level_filter: TopLevelFilter,
    mut apply: impl FnMut(u16, &[u8], &mut Budget) -> Result<(), String>,
) -> Result<(), String> {
    let (direct, appended) = match top_level_filter {
        TopLevelFilter::All => (bytes, None),
        TopLevelFilter::Paragraph => (&[][..], Some(bytes)),
    };
    paragraph_properties_appended(
        direct,
        appended,
        data,
        budget,
        top_level_filter,
        |code, operand, budget, _| apply(code, operand, budget),
    )
}

/// Traverse direct PAPX properties, then the paragraph's Pcd.Prm properties.
///
/// Normative basis: [MS-DOC] 2.4.6.1 steps 4-5 place the piece properties after
/// the direct grpprl, and 2.6.2 lets sprmPHugePapx/sprmPTableProps replace the
/// rest of "the array that contained" them with a Data-stream PrcData
/// (2.9.210); sprmPHugePapx is only honored as the first Prl.
///
/// Observed Word 16.113 (macOS) behavior, from native controls that each
/// changed one Word-saved DOC row mark or table-cell paragraph and compared
/// Word's PDF plus its saved DOCX (cell widths, table style, alignment), with
/// a sprmPJc complex-piece witness proving the split pieces were active:
/// - A direct sprmPTableProps is followed only as the first Prl of the direct
///   grpprl. A later one is ignored and the following Prls still apply, the
///   same rule 2.6.2 states for sprmPHugePapx. PrcData chains restart the
///   first-position rule.
/// - Piece properties (Prm1 and Prm0) are not part of the replaced array:
///   they still apply after a direct first-position sprmPTableProps and its
///   PrcData (a paragraph jc from either PRM form and a row TDefTable/TIstd
///   survived).
/// - Piece sprmPTableProps/sprmPHugePapx are never followed, even when they
///   are the first Prl of the paragraph (empty direct PAPX, BxPap.bOffset 0).
///   Word ignores them and still applies the following piece Prls.
/// - With `TopLevelFilter::Paragraph`, piece table SPRMs (sgc 5) are kept with
///   paragraph SPRMs (sgc 1): Word applied a piece TDefTable and TIstd to the
///   row mark. Other groups (character, picture, section) stay excluded, as in
///   2.4.6.1 step 5. Simple Prm0 values only encode paragraph/character SPRMs.
/// The last callback argument reports whether the Prl came from the piece.
pub fn paragraph_properties_appended<'a>(
    direct: &'a [u8],
    appended: Option<&'a [u8]>,
    data: &'a [u8],
    budget: &mut Budget,
    appended_filter: TopLevelFilter,
    mut apply: impl FnMut(u16, &[u8], &mut Budget, bool) -> Result<(), String>,
) -> Result<(), String> {
    let mut visited = BTreeSet::new();
    let mut bytes = direct;
    loop {
        let mut sprms = Sprms::new(bytes);
        let mut next = None;
        let mut first = true;
        while let Some((code, operand)) = sprms.next(budget)? {
            if matches!(code, 0x646b | 0x6646) {
                if first {
                    next = Some(prc_data(operand, data, &mut visited)?);
                    break;
                }
                continue;
            }
            apply(code, operand, budget, false)?;
            first = false;
        }
        match next {
            Some(value) => bytes = value,
            None => break,
        }
    }
    let Some(piece) = appended else {
        return Ok(());
    };
    let mut sprms = Sprms::new(piece);
    while let Some((code, operand)) = sprms.next(budget)? {
        if matches!(code, 0x646b | 0x6646)
            || (matches!(appended_filter, TopLevelFilter::Paragraph)
                && !matches!((code >> 10) & 7, 1 | 5))
        {
            continue;
        }
        apply(code, operand, budget, true)?;
    }
    Ok(())
}

/// Resolve one sprmPHugePapx/sprmPTableProps operand to its PrcData grpprl.
fn prc_data<'a>(
    operand: &[u8],
    data: &'a [u8],
    visited: &mut BTreeSet<usize>,
) -> Result<&'a [u8], String> {
    let offset = u32_at(operand, 0)? as usize;
    if visited.len() >= 64 || !visited.insert(offset) {
        return Err(unsupported("cyclic or excessive Word paragraph data chain"));
    }
    let record = data
        .get(offset..)
        .ok_or_else(|| unsupported("Word paragraph data offset outside Data stream"))?;
    let size = u16_at(record, 0)? as usize;
    if size < 10 {
        return Err(unsupported("short Word paragraph data record"));
    }
    if size > MAX_PRC_DATA_GRPPRL_BYTES {
        return Err(unsupported("oversized Word paragraph data record"));
    }
    record
        .get(2..2 + size)
        .ok_or_else(|| unsupported("Word paragraph properties outside Data stream"))
}

pub struct Budget(usize);
impl Default for Budget {
    fn default() -> Self {
        Self(4_000_000)
    }
}
impl Budget {
    pub fn take(&mut self) -> Result<(), String> {
        self.take_many(1)
    }
    #[cfg(test)]
    pub(super) fn remaining(&self) -> usize {
        self.0
    }
    fn take_many(&mut self, amount: usize) -> Result<(), String> {
        self.0 = self
            .0
            .checked_sub(amount)
            .ok_or_else(|| unsupported("Word formatting operation budget exceeded"))?;
        Ok(())
    }
}

pub struct Sprms<'a> {
    bytes: &'a [u8],
}
impl<'a> Sprms<'a> {
    pub fn new(bytes: &'a [u8]) -> Self {
        Self { bytes }
    }
    pub fn next(&mut self, budget: &mut Budget) -> Result<Option<(u16, &'a [u8])>, String> {
        if self.bytes.is_empty() {
            return Ok(None);
        }
        budget.take()?;
        let code = u16_at(self.bytes, 0)?;
        let bytes = &self.bytes[2..];
        let size = match code >> 13 {
            0 | 1 => 1,
            2 | 4 | 5 => 2,
            3 => 4,
            7 => 3,
            _ if code == 0xd608 => {
                let cb = u16_at(bytes, 0)? as usize;
                if cb == 0 {
                    return Err(unsupported("invalid Word table property size"));
                }
                cb + 1
            }
            _ if code == 0xc615 && bytes.first() == Some(&255) => {
                let deleted = *bytes
                    .get(1)
                    .ok_or_else(|| unsupported("truncated Word tab property"))?
                    as usize;
                let added = *bytes
                    .get(2 + deleted * 4)
                    .ok_or_else(|| unsupported("truncated Word tab additions"))?
                    as usize;
                3 + deleted * 4 + added * 3
            }
            _ => {
                1 + *bytes
                    .first()
                    .ok_or_else(|| unsupported("truncated Word variable property"))?
                    as usize
            }
        };
        let operand = bytes
            .get(..size)
            .ok_or_else(|| unsupported("truncated Word formatting operand"))?;
        if matches!(
            code,
            0xc60d | 0xc615 | 0xd609 | 0xd612 | 0xd616 | 0xd60c | 0xd62d | 0xd62e | 0xd660
        ) {
            // Charge variable tab edits, not just their enclosing SPRM. Range
            // deletion is logarithmic plus removed entries, not a full-set scan.
            // Table shading arrays/ranges also perform per-cell work (at most
            // 63 cells for a range); account for it below before expansion.
            budget.take_many(operand.len())?;
            // DefTableShd replaces its segment, including omitted non-shaded
            // trailing cells, so even an empty array has bounded reset work.
            budget.take_many(match code {
                0xd612 | 0xd616 => 22,
                0xd60c => 19,
                _ => 0,
            })?;
            if matches!(code, 0xd62d | 0xd62e) && operand.len() >= 3 {
                budget.take_many(usize::from(operand[2].saturating_sub(operand[1])))?;
            }
        }
        self.bytes = &bytes[size..];
        Ok(Some((code, operand)))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn valid_grpprl(size: usize) -> Vec<u8> {
        let four_byte_sprms = size % 3;
        let mut bytes = Vec::with_capacity(size);
        for _ in 0..four_byte_sprms {
            bytes.extend([0x00, 0x46, 0, 0]);
        }
        while bytes.len() < size {
            bytes.extend([0x07, 0x24, 0]);
        }
        assert_eq!(bytes.len(), size);
        bytes
    }

    fn prc_data(size: usize) -> Vec<u8> {
        let mut bytes = Vec::with_capacity(size + 2);
        bytes.extend(u16::try_from(size).unwrap().to_le_bytes());
        bytes.extend(valid_grpprl(size));
        bytes
    }

    #[test]
    fn prc_data_enforces_the_cb_grpprl_maximum_at_its_exact_boundary() {
        let reference = [0x46, 0x66, 0, 0, 0, 0];
        for size in [0x3fa1, 0x3fa2] {
            let data = prc_data(size);
            let mut applied = 0;
            paragraph_properties(
                &reference,
                &data,
                &mut Budget::default(),
                TopLevelFilter::All,
                |_, _, _| {
                    applied += 1;
                    Ok(())
                },
            )
            .unwrap();
            assert!(applied > 0);
        }

        let data = prc_data(0x3fa3);
        let mut applied = 0;
        let error = paragraph_properties(
            &reference,
            &data,
            &mut Budget::default(),
            TopLevelFilter::All,
            |_, _, _| {
                applied += 1;
                Ok(())
            },
        )
        .unwrap_err();
        assert!(error.contains("oversized Word paragraph data"));
        assert_eq!(applied, 0);
    }

    #[test]
    fn shading_work_is_charged_before_cell_expansion() {
        let mut budget = Budget(76); // 1 SPRM + 13 operand bytes + 63 selected cells = 77.
        let mut bytes = vec![0x2d, 0xd6, 12, 0, 63];
        bytes.extend([0u8; 10]);
        assert!(Sprms::new(&bytes)
            .next(&mut budget)
            .unwrap_err()
            .contains("budget"));
        let mut budget = Budget(77);
        assert!(Sprms::new(&bytes).next(&mut budget).unwrap().is_some());
        assert_eq!(budget.0, 0);
    }

    fn traced(
        direct: &[u8],
        piece: Option<&[u8]>,
        data: &[u8],
        budget: &mut Budget,
    ) -> Result<Vec<(u16, u8, bool)>, String> {
        let mut applied = Vec::new();
        paragraph_properties_appended(
            direct,
            piece,
            data,
            budget,
            TopLevelFilter::Paragraph,
            |code, operand, _, from_piece| {
                applied.push((code, operand[0], from_piece));
                Ok(())
            },
        )?;
        Ok(applied)
    }

    #[test]
    fn direct_redirects_are_followed_only_from_the_first_prl() {
        let data = [
            12, 0, 0x03, 0x24, 2, 0x07, 0x24, 0, 0x07, 0x24, 0, 0x07, 0x24, 0,
        ];
        let first = [0x6b, 0x64, 0, 0, 0, 0, 0x03, 0x24, 1];
        let expected = [
            (0x2403, 2, false),
            (0x2407, 0, false),
            (0x2407, 0, false),
            (0x2407, 0, false),
        ];
        assert_eq!(
            traced(&first, None, &data, &mut Budget(5)).unwrap(),
            expected
        );
        assert!(traced(&first, None, &data, &mut Budget(4))
            .unwrap_err()
            .contains("budget"));

        // A later PTableProps/PHugePapx is ignored without reading Data.
        for redirect in [0x646b_u16, 0x6646] {
            let [low, high] = redirect.to_le_bytes();
            let later = [
                0x07, 0x24, 1, low, high, 0xff, 0xff, 0xff, 0xff, 0x03, 0x24, 1,
            ];
            assert_eq!(
                traced(&later, None, &[], &mut Budget::default()).unwrap(),
                [(0x2407, 1, false), (0x2403, 1, false)]
            );
        }

        // Each PrcData restarts the first-position rule, so chains are bounded.
        let cycle = [12, 0, 0x6b, 0x64, 0, 0, 0, 0, 0x07, 0x24, 0, 0x07, 0x24, 0];
        assert!(traced(
            &[0x46, 0x66, 0, 0, 0, 0],
            None,
            &cycle,
            &mut Budget::default()
        )
        .unwrap_err()
        .contains("cyclic"));
    }

    #[test]
    fn piece_properties_follow_the_direct_chain_without_redirects() {
        let data = [
            12, 0, 0x03, 0x24, 1, 0x07, 0x24, 0, 0x07, 0x24, 0, 0x07, 0x24, 0,
        ];
        let piece_alignment = [0x61, 0x24, 2];
        assert_eq!(
            traced(
                &[0x6b, 0x64, 0, 0, 0, 0],
                Some(&piece_alignment),
                &data,
                &mut Budget::default()
            )
            .unwrap(),
            [
                (0x2403, 1, false),
                (0x2407, 0, false),
                (0x2407, 0, false),
                (0x2407, 0, false),
                (0x2461, 2, true)
            ]
        );

        // Piece redirects are ignored even as the paragraph's first Prl; table
        // SPRMs are kept and character SPRMs are excluded.
        let piece = [
            0x46, 0x66, 0xff, 0xff, 0xff, 0xff, // PHugePapx
            0x6b, 0x64, 0xff, 0xff, 0xff, 0xff, // PTableProps
            0x08, 0xd6, 1, 0, // TDefTable
            0x35, 0x08, 1, // CFBold
            0x03, 0x24, 1,
        ];
        assert_eq!(
            traced(&[], Some(&piece), &[], &mut Budget::default()).unwrap(),
            [(0xd608, 1, true), (0x2403, 1, true)]
        );
    }

    #[test]
    fn extended_tabs_are_framed_without_consuming_the_following_sprm() {
        let mut b = vec![0x15, 0xc6, 255, 64];
        b.extend([0u8; 256]);
        b.push(0);
        b.extend([0x07, 0x24, 1]);
        let mut p = Sprms::new(&b);
        let mut budget = Budget::default();
        assert_eq!(p.next(&mut budget).unwrap().unwrap().1.len(), 259);
        assert_eq!(p.next(&mut budget).unwrap(), Some((0x2407, &[1][..])));
        assert!(Sprms::new(&b).next(&mut Budget(259)).is_err());
    }
    #[test]
    fn skips_unknown_variable_operands_and_rejects_truncation() {
        let mut budget = Budget::default();
        let mut p = Sprms::new(&[0x71, 0xca, 2, 9, 8, 0x35, 8, 1]);
        assert_eq!(p.next(&mut budget).unwrap(), Some((0xca71, &[2, 9, 8][..])));
        assert_eq!(p.next(&mut budget).unwrap(), Some((0x0835, &[1][..])));
        assert!(p.next(&mut budget).unwrap().is_none());
        assert!(Sprms::new(&[0x71, 0xca, 3, 9]).next(&mut budget).is_err());
        assert!(Sprms::new(&[1]).next(&mut budget).is_err());
    }
}
