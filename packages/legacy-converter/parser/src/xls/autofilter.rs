//! Worksheet AutoFilter ranges and their drop-down objects. MS-XLS 2.4.8
//! (AutoFilterInfo), 2.4.150 (Lbl, built-in name 0x0D _FilterDatabase),
//! 2.5.198.28 (PtgArea3d), 2.4.105/2.4.271 (ExternSheet/SupBook) and 2.4.181
//! (Obj ot 20 with FtCmo.fUIObj).
//!
//! Excel inserts one drop-down object per AutoFilter column itself (fUIObj,
//! "can only be automatically inserted by the application") and anchors it
//! with fMove set but fSize clear, which OfficeArtClientAnchorSheet (2.5.193)
//! otherwise forbids. Those objects are the AutoFilter's buttons, not drawn
//! content: Excel's own XLSX of the same workbook has an `autoFilter` element
//! over the named range and no drawing objects for them, so the direct model
//! projects the range and lets the worksheet renderer draw the buttons.
use super::{parse_bound_sheet, u16_at, unsupported, Record, BOF, BOUNDSHEET8, EOF};
use std::collections::BTreeMap;

/// Zero-based inclusive cell range.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct AutoFilter {
    pub first_row: u16,
    pub last_row: u16,
    pub first_column: u16,
    pub last_column: u16,
}

impl AutoFilter {
    pub(super) fn model(self) -> xlsx_model::CellRange {
        xlsx_model::CellRange {
            top: u32::from(self.first_row) + 1,
            left: u32::from(self.first_column) + 1,
            bottom: u32::from(self.last_row) + 1,
            right: u32::from(self.last_column) + 1,
        }
    }

    /// Whether a drop-down object anchored at this cell is one of the
    /// filter's column buttons (the header row, inside the filter columns).
    pub(super) fn owns_button(self, column: u16, row: u16) -> bool {
        row == self.first_row && (self.first_column..=self.last_column).contains(&column)
    }
}

/// AutoFilter ranges by zero-based BoundSheet tab index.
pub(super) fn workbook(records: &[Record<'_>]) -> Result<BTreeMap<usize, AutoFilter>, String> {
    let globals_end = records
        .iter()
        .position(|record| record.kind == EOF)
        .ok_or_else(|| unsupported("missing BIFF global EOF"))?;
    let mut self_books = Vec::new();
    let mut xti = Vec::new();
    let mut names = BTreeMap::new();
    let mut sheets = Vec::new();
    for record in &records[..globals_end] {
        match record.kind {
            0x01ae => self_books.push(u16_at(record.data, 2).is_ok_and(|cch| cch == 0x0401)),
            0x0017 => {
                let count = usize::from(u16_at(record.data, 0)?);
                for index in 0..count {
                    let at = 2 + index * 6;
                    xti.push((
                        u16_at(record.data, at)?,
                        u16_at(record.data, at + 2)?,
                        u16_at(record.data, at + 4)?,
                    ));
                }
            }
            0x0018 => {
                if let Some((tab, range)) = filter_database(record.data)? {
                    names.insert(tab, range);
                }
            }
            BOUNDSHEET8 => sheets.push(parse_bound_sheet(record.data)?.offset),
            _ => {}
        }
    }
    let mut filters = BTreeMap::new();
    for (tab, &offset) in sheets.iter().enumerate() {
        let Ok(start) = records.binary_search_by_key(&offset, |record| record.offset) else {
            continue;
        };
        let Some(columns) = filter_info(&records[start..])? else {
            continue;
        };
        let (ixti, range) = names
            .get(&tab)
            .copied()
            .ok_or_else(|| unsupported("XLS AutoFilter without its _FilterDatabase range"))?;
        let resolved = xti
            .get(usize::from(ixti))
            .filter(|(book, first, last)| {
                first == last
                    && usize::from(*first) == tab
                    && self_books.get(usize::from(*book)).copied().unwrap_or(false)
            })
            .is_some();
        if !resolved {
            return Err(unsupported(
                "XLS AutoFilter range does not refer to its own sheet",
            ));
        }
        if usize::from(range.last_column - range.first_column) + 1 != usize::from(columns) {
            return Err(unsupported(
                "XLS AutoFilter column count disagrees with its range",
            ));
        }
        filters.insert(tab, range);
    }
    Ok(filters)
}

/// A built-in _FilterDatabase Lbl (fBuiltin, name 0x0D) whose formula is one
/// PtgArea3d: its one-based itab as a tab index, the XTI and the range.
fn filter_database(data: &[u8]) -> Result<Option<(usize, (u16, AutoFilter))>, String> {
    if data.len() < 15 {
        return Err(unsupported("truncated BIFF Lbl record"));
    }
    let grbit = u16_at(data, 0)?;
    let characters = data[3];
    let formula = usize::from(u16_at(data, 4)?);
    let itab = u16_at(data, 8)?;
    // fBuiltin with the one-character built-in name 0x0D, not a string name.
    if grbit & 0x20 == 0 || characters != 1 || data[14] != 0 || data.get(15) != Some(&0x0d) {
        return Ok(None);
    }
    if itab == 0 {
        return Err(unsupported("XLS _FilterDatabase without its sheet"));
    }
    let rgce = data
        .get(16..16 + formula)
        .ok_or_else(|| unsupported("truncated BIFF Lbl formula"))?;
    if rgce.len() != 11 || !matches!(rgce[0], 0x3b | 0x5b | 0x7b) {
        return Err(unsupported("XLS _FilterDatabase is not one 3-D area"));
    }
    let range = AutoFilter {
        first_row: u16_at(rgce, 3)?,
        last_row: u16_at(rgce, 5)?,
        first_column: u16_at(rgce, 7)? & 0x3fff,
        last_column: u16_at(rgce, 9)? & 0x3fff,
    };
    if range.first_row > range.last_row || range.first_column > range.last_column {
        return Err(unsupported("invalid XLS _FilterDatabase range"));
    }
    Ok(Some((usize::from(itab - 1), (u16_at(rgce, 1)?, range))))
}

/// AutoFilterInfo's column count in a worksheet substream. Filter criteria
/// (AutoFilter, AutoFilter12, FilterMode) change the buttons and the hidden
/// rows and are not projected.
fn filter_info(records: &[Record<'_>]) -> Result<Option<u16>, String> {
    let mut depth = 0usize;
    let mut columns = None;
    for record in records {
        match record.kind {
            BOF => depth += 1,
            EOF => {
                depth -= 1;
                if depth == 0 {
                    break;
                }
            }
            _ if depth != 1 => {}
            0x009d => {
                if columns.is_some() {
                    return Err(unsupported("duplicate XLS AutoFilterInfo"));
                }
                columns = Some(u16_at(record.data, 0)?);
            }
            0x009b | 0x009e | 0x087e => {
                return Err(unsupported("XLS AutoFilter criteria are not projected"))
            }
            _ => {}
        }
    }
    Ok(columns)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn records(data: &[(u16, Vec<u8>)]) -> Vec<Record<'_>> {
        let mut offset = 0;
        data.iter()
            .map(|(kind, data)| {
                let record = Record {
                    kind: *kind,
                    offset,
                    data,
                };
                offset += 4 + data.len();
                record
            })
            .collect()
    }

    fn lbl(itab: u16, ixti: u16, rows: (u16, u16), columns: (u16, u16)) -> Vec<u8> {
        let mut data = vec![0x21, 0, 0, 1, 11, 0, 0, 0];
        data.extend_from_slice(&itab.to_le_bytes());
        data.extend_from_slice(&[0, 0, 0, 0, 0, 0x0d, 0x3b]);
        for value in [ixti, rows.0, rows.1, columns.0, columns.1] {
            data.extend_from_slice(&value.to_le_bytes());
        }
        data
    }

    /// Globals with one self SupBook, XTI 0 -> tab `xti_tab`, a
    /// _FilterDatabase for tab 0 and one worksheet with `sheet` records.
    fn workbook_records(xti_tab: u16, sheet: &[(u16, Vec<u8>)]) -> Vec<(u16, Vec<u8>)> {
        let mut data = vec![
            (BOF, vec![0, 6, 5, 0]),
            (BOUNDSHEET8, vec![0, 0, 0, 0, 0, 0, 1, 0, b'S']),
            (0x01ae, vec![1, 0, 1, 4]),
            (
                0x0017,
                [1u16, 0, xti_tab, xti_tab]
                    .into_iter()
                    .flat_map(u16::to_le_bytes)
                    .collect(),
            ),
            (0x0018, lbl(1, 0, (3, 48), (1, 8))),
            (EOF, vec![]),
        ];
        let offset: usize = data.iter().map(|(_, d)| d.len() + 4).sum();
        data[1].1[..4].copy_from_slice(&(offset as u32).to_le_bytes());
        data.push((BOF, vec![0, 6, 0x10, 0]));
        data.extend_from_slice(sheet);
        data.push((EOF, vec![]));
        data
    }

    #[test]
    fn autofilter_range_comes_from_the_sheets_filter_database() {
        let owned = workbook_records(0, &[(0x009d, vec![8, 0])]);
        let filters = workbook(&records(&owned)).unwrap();
        let filter = filters[&0];
        let model = filter.model();
        assert_eq!(
            (model.top, model.left, model.bottom, model.right),
            (4, 2, 49, 9)
        );
        assert!(filter.owns_button(1, 3) && filter.owns_button(8, 3));
        assert!(!filter.owns_button(9, 3) && !filter.owns_button(1, 4));
        // No AutoFilterInfo: the name alone (e.g. an advanced filter) draws nothing.
        assert!(workbook(&records(&workbook_records(0, &[])))
            .unwrap()
            .is_empty());
    }

    #[test]
    fn autofilter_criteria_and_inconsistent_ranges_fail_closed() {
        for (xti_tab, sheet, expected) in [
            (0, vec![(0x009d, vec![7, 0])], "column count"),
            (1, vec![(0x009d, vec![8, 0])], "own sheet"),
            (
                0,
                vec![(0x009d, vec![8, 0]), (0x009e, vec![0; 24])],
                "criteria",
            ),
            (0, vec![(0x009b, vec![]), (0x009d, vec![8, 0])], "criteria"),
        ] {
            let owned = workbook_records(xti_tab, &sheet);
            let error = workbook(&records(&owned)).unwrap_err();
            assert!(error.contains(expected), "{expected}: {error}");
        }
    }
}
