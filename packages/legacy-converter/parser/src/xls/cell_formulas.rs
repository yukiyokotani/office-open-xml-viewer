//! Cell formula text of the direct XLS model, as the XLSX parser keeps the
//! text of ECMA-376 18.3.1.40 `<f>`: a cell's own formula, the anchor cell
//! of a shared or array formula (whose text SpreadsheetML stores once, on
//! that cell), and no text for the other cells of a shared or array range
//! or for data tables (`t="dataTable"` carries no text). As in the XLSX
//! model the text is informational only: formulas are never calculated, and
//! the renderer shows each cell's cached value, TODAY()/NOW() included.
//!
//! MS-XLS 2.4.127 Formula carries a CellParsedFormula; one that is only a
//! PtgExp (2.5.198.58) or PtgTbl refers to the 2.4.260 ShrFmla, 2.4.4 Array or 2.4.313
//! Table record that follows its anchor cell. Shared formulas use relative
//! PtgRefN/PtgAreaN tokens, written for the anchor cell as the Excel-saved
//! .xlsx counterparts do (observed with a shared formula whose members each
//! refer to a neighbouring cell).
//! Formulas the decompiler cannot express reject the workbook.

use super::conditional::{decompile_cell, Externs};
use super::{u16_at, unsupported};
use std::collections::BTreeMap;

fn truncated() -> String {
    unsupported("truncated XLS cell formula")
}

/// Project FORMULA (0x0006), SHRFMLA (0x04BC), ARRAY (0x0221) and TABLE
/// (0x0236) records, in stream order, to formula text by cell.
pub(super) fn project<'a>(
    records: impl Iterator<Item = (u16, &'a [u8])>,
    externs: &Externs,
) -> Result<BTreeMap<(u16, u16), String>, String> {
    let mut output = BTreeMap::new();
    // The anchor FORMULA whose PtgExp names itself, awaiting its record.
    let mut anchor: Option<(u16, u16)> = None;
    for (kind, data) in records {
        match kind {
            0x0006 => {
                if anchor.is_some() {
                    return Err(unsupported("XLS shared formula lacks its definition"));
                }
                let row = u16_at(data, 0)?;
                let col = u16_at(data, 2)?;
                let cce = usize::from(u16_at(data, 20)?);
                let rgce = data.get(22..22 + cce).ok_or_else(truncated)?;
                let extra = data.get(22 + cce..).ok_or_else(truncated)?;
                if matches!(rgce.first(), Some(0x01 | 0x02)) {
                    // PtgExp / PtgTbl (2.5.198.92): a cell of a shared,
                    // array or data-table formula.
                    if rgce.len() != 5 || !extra.is_empty() {
                        return Err(unsupported("invalid XLS formula reference"));
                    }
                    if (u16_at(rgce, 1)?, u16_at(rgce, 3)?) == (row, col) {
                        anchor = Some((row, col));
                    }
                    continue;
                }
                output.insert(
                    (row, col),
                    decompile_cell(rgce, extra, (row, col), externs)?,
                );
            }
            0x04bc | 0x0221 => {
                let cell = anchor
                    .take()
                    .ok_or_else(|| unsupported("XLS shared formula without its anchor"))?;
                let offset = if kind == 0x04bc { 8 } else { 12 };
                let cce = usize::from(u16_at(data, offset)?);
                let rgce = data
                    .get(offset + 2..offset + 2 + cce)
                    .ok_or_else(truncated)?;
                let extra = data.get(offset + 2 + cce..).ok_or_else(truncated)?;
                output.insert(cell, decompile_cell(rgce, extra, cell, externs)?);
            }
            0x0236 => {
                anchor
                    .take()
                    .ok_or_else(|| unsupported("XLS data table without its anchor"))?;
            }
            _ => return Err(unsupported("unexpected XLS formula record")),
        }
    }
    if anchor.is_some() {
        return Err(unsupported("XLS shared formula lacks its definition"));
    }
    Ok(output)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn formula(row: u16, col: u16, rgce: &[u8]) -> Vec<u8> {
        let mut data = Vec::new();
        for value in [row, col, 15] {
            data.extend(value.to_le_bytes());
        }
        data.extend([0u8; 8]);
        data.extend(0u16.to_le_bytes());
        data.extend(0u32.to_le_bytes());
        data.extend((rgce.len() as u16).to_le_bytes());
        data.extend(rgce);
        data
    }

    #[test]
    fn own_shared_and_table_formulas_follow_spreadsheetml() {
        // A1: TODAY(); B1..C1 share RefN(0,-1)+1 anchored at B1; D1 is a
        // data table.
        let today = formula(0, 0, &[0x41, 0xdd, 0x00]);
        let anchor = formula(0, 1, &[0x01, 0, 0, 1, 0]);
        let mut shared = vec![0u8, 0, 0, 0, 1, 2, 0, 2];
        let rgce = [0x4c, 0, 0, 0xff, 0xff, 0x1e, 1, 0, 0x03];
        shared.extend((rgce.len() as u16).to_le_bytes());
        shared.extend(rgce);
        let member = formula(0, 2, &[0x01, 0, 0, 1, 0]);
        let table_anchor = formula(0, 3, &[0x02, 0, 0, 3, 0]);
        let table = [0u8; 16];
        let records = [
            (0x0006u16, today.as_slice()),
            (0x0006, anchor.as_slice()),
            (0x04bc, shared.as_slice()),
            (0x0006, member.as_slice()),
            (0x0006, table_anchor.as_slice()),
            (0x0236, table.as_slice()),
        ];
        let output = project(records.into_iter(), &Externs::default()).unwrap();
        assert_eq!(output.get(&(0, 0)).map(String::as_str), Some("TODAY()"));
        assert_eq!(output.get(&(0, 1)).map(String::as_str), Some("A1+1"));
        assert!(!output.contains_key(&(0, 2)) && !output.contains_key(&(0, 3)));
        // An anchor without its definition rejects.
        assert!(project(
            [(0x0006u16, anchor.as_slice())].into_iter(),
            &Externs::default()
        )
        .is_err());
    }
}
