//! Worksheet AutoFilter ranges (ECMA-376 18.3.1.2 `autoFilter@ref`).
//!
//! BIFF8 keeps the filtered range in the sheet's built-in `_FilterDatabase`
//! name (MS-XLS 2.4.150 Lbl, fBuiltin with name 0x0D, whose itab is the
//! one-based sheet index) and marks the sheet as filtered with 2.4.8
//! AutoFilterInfo, whose column count must match that range. The XLSX model
//! keeps only the range, from which the renderer draws the filter buttons, as
//! the XLSX parser does. Active criteria (2.4.6 AutoFilter, 2.4.7
//! AutoFilter12, 2.4.117 FilterMode) also change how Excel draws the filtered
//! columns' buttons, which that range cannot express, so they fail closed.
//!
//! Excel inserts one drop-down object per filter column itself (Obj ot 20
//! with FtCmo.fUIObj, "can only be automatically inserted by the
//! application") and anchors it with fMove set but fSize clear, which
//! OfficeArtClientAnchorSheet (2.5.193) otherwise forbids. Those objects are
//! the filter's buttons, not drawn content: Excel's own XLSX of the same
//! workbook has an `autoFilter` over the named range and no drawing objects
//! for them. [`owns_button`] identifies them for the drawing validation.

use super::{u16_at, unsupported, Record, EOF};
use std::collections::BTreeMap;

/// `_FilterDatabase` ranges by zero-based sheet index.
#[derive(Default)]
pub(super) struct Databases(BTreeMap<usize, xlsx_model::CellRange>);

impl Databases {
    pub(super) fn parse(records: &[Record<'_>]) -> Result<Self, String> {
        let globals: Vec<&Record<'_>> = records.iter().take_while(|r| r.kind != EOF).collect();
        // EXTERNSHEET XTIs of the workbook's own SupBook (cch 0x0401).
        let mut own_books = Vec::new();
        let mut xtis = Vec::new();
        for record in &globals {
            match record.kind {
                0x01ae => own_books.push(u16_at(record.data, 2).is_ok_and(|cch| cch == 0x0401)),
                0x0017 => {
                    let count = usize::from(u16_at(record.data, 0)?);
                    for index in 0..count {
                        let at = 2 + index * 6;
                        xtis.push((
                            u16_at(record.data, at)?,
                            u16_at(record.data, at + 2)?,
                            u16_at(record.data, at + 4)?,
                        ));
                    }
                }
                _ => {}
            }
        }
        let mut result = Self::default();
        for record in globals.iter().filter(|r| r.kind == 0x0018) {
            let data = record.data;
            let flags = u16_at(data, 0)?;
            let chars = usize::from(*data.get(3).ok_or_else(truncated)?);
            // fBuiltin with the one-character built-in name 0x0D.
            if flags & 0x0020 == 0 || chars != 1 {
                continue;
            }
            let high = *data.get(14).ok_or_else(truncated)?;
            let code = match high {
                0 => u16::from(*data.get(15).ok_or_else(truncated)?),
                1 => u16_at(data, 15)?,
                _ => return Err(unsupported("invalid BIFF defined name")),
            };
            if code != 0x0d {
                continue;
            }
            let itab = usize::from(u16_at(data, 8)?);
            let cce = usize::from(u16_at(data, 4)?);
            let start = 15 + if high == 0 { 1 } else { 2 };
            let rgce = data.get(start..start + cce).ok_or_else(truncated)?;
            let sheet = itab
                .checked_sub(1)
                .ok_or_else(|| unsupported("BIFF filter database without a sheet"))?;
            let range = area(rgce, sheet, &own_books, &xtis)?;
            if result.0.insert(sheet, range).is_some() {
                return Err(unsupported("duplicate BIFF filter database"));
            }
        }
        Ok(result)
    }

    pub(super) fn range(&self, sheet: usize) -> Option<&xlsx_model::CellRange> {
        self.0.get(&sheet)
    }
}

/// The AutoFilterInfo column count must span the range; criteria reject.
pub(super) fn check(
    range: &xlsx_model::CellRange,
    columns: u16,
    criteria: bool,
) -> Result<(), String> {
    if criteria {
        return Err(unsupported("XLS AutoFilter criteria are not projected"));
    }
    if range.right - range.left + 1 != u32::from(columns) {
        return Err(unsupported(
            "XLS AutoFilter column count disagrees with its range",
        ));
    }
    Ok(())
}

/// Whether a drop-down object anchored at this zero-based cell is one of the
/// filter's column buttons: on the header row, inside the filter columns.
pub(super) fn owns_button(range: &xlsx_model::CellRange, column: u16, row: u16) -> bool {
    let (column, row) = (u32::from(column) + 1, u32::from(row) + 1);
    row == range.top && (range.left..=range.right).contains(&column)
}

fn truncated() -> String {
    unsupported("truncated BIFF filter database")
}

/// A single PtgArea3d (2.5.198.28) or PtgRef3d (2.5.198.85) on `sheet`.
fn area(
    rgce: &[u8],
    sheet: usize,
    own_books: &[bool],
    xtis: &[(u16, u16, u16)],
) -> Result<xlsx_model::CellRange, String> {
    let token = *rgce.first().ok_or_else(truncated)?;
    let (rows, cols, size) = match token & 0x9f {
        0x1b => (
            (u16_at(rgce, 3)?, u16_at(rgce, 5)?),
            (u16_at(rgce, 7)? & 0x3fff, u16_at(rgce, 9)? & 0x3fff),
            11,
        ),
        0x1a => {
            let row = u16_at(rgce, 3)?;
            let col = u16_at(rgce, 5)? & 0x3fff;
            ((row, row), (col, col), 7)
        }
        _ => return Err(unsupported("unsupported BIFF filter database formula")),
    };
    if !matches!(token >> 5 & 3, 1..=3) || rgce.len() != size {
        return Err(unsupported("unsupported BIFF filter database formula"));
    }
    let xti = xtis
        .get(usize::from(u16_at(rgce, 1)?))
        .ok_or_else(|| unsupported("BIFF filter database names no sheet"))?;
    let own = own_books.get(usize::from(xti.0)).copied().unwrap_or(false);
    if !own || xti.1 != xti.2 || usize::from(xti.1) != sheet {
        return Err(unsupported("BIFF filter database on another sheet"));
    }
    if rows.0 > rows.1 || cols.0 > cols.1 || cols.1 > 0x00ff {
        return Err(unsupported("invalid BIFF filter database range"));
    }
    Ok(xlsx_model::CellRange {
        top: u32::from(rows.0) + 1,
        left: u32::from(cols.0) + 1,
        bottom: u32::from(rows.1) + 1,
        right: u32::from(cols.1) + 1,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn filter_databases_resolve_to_their_sheet_ranges() {
        let supbook = [2u8, 0, 1, 4];
        let externsheet = [1u8, 0, 0, 0, 7, 0, 7, 0];
        let mut name = vec![0x21, 0, 0, 1, 11, 0, 0, 0, 8, 0, 0, 0, 0, 0, 0, 0x0d];
        name.extend([0x3b, 0, 0, 3, 0, 48, 0, 1, 0, 8, 0]);
        let records = [
            Record {
                kind: 0x01ae,
                offset: 0,
                data: &supbook,
            },
            Record {
                kind: 0x0017,
                offset: 1,
                data: &externsheet,
            },
            Record {
                kind: 0x0018,
                offset: 2,
                data: &name,
            },
        ];
        let databases = Databases::parse(&records).unwrap();
        let range = databases.range(7).unwrap();
        assert_eq!(
            (range.top, range.left, range.bottom, range.right),
            (4, 2, 49, 9)
        );
        // A database whose reference names another sheet rejects.
        let mut other = name.clone();
        other[8] = 3;
        let records = [
            Record {
                kind: 0x01ae,
                offset: 0,
                data: &supbook,
            },
            Record {
                kind: 0x0017,
                offset: 1,
                data: &externsheet,
            },
            Record {
                kind: 0x0018,
                offset: 2,
                data: &other,
            },
        ];
        assert!(Databases::parse(&records).is_err());
    }

    #[test]
    fn autofilter_buttons_counts_and_criteria() {
        let range = xlsx_model::CellRange {
            top: 4,
            left: 2,
            bottom: 49,
            right: 9,
        };
        assert!(owns_button(&range, 1, 3) && owns_button(&range, 8, 3));
        assert!(!owns_button(&range, 9, 3) && !owns_button(&range, 1, 4));
        assert!(check(&range, 8, false).is_ok());
        assert!(check(&range, 7, false)
            .unwrap_err()
            .contains("column count"));
        assert!(check(&range, 8, true).unwrap_err().contains("criteria"));
    }
}
