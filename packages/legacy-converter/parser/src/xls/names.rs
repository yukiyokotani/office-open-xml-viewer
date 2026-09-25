//! MS-XLS 2.4.150 Lbl records as the XLSX model's defined names, as the
//! XLSX parser reads ECMA-376 18.2.5 `definedName` for a sheet: workbook
//! names plus the names local to that sheet (itab), with the formula text.
//! Built-in names take the `_xlnm.` prefix SpreadsheetML gives them
//! (18.2.5, e.g. `_xlnm.Print_Area`, `_xlnm._FilterDatabase`).
//!
//! The renderer resolves these names in conditional-formatting formulas
//! and internal hyperlinks.

use super::conditional::{decompile_name, Externs};
use super::{u16_at, unsupported, Record, EOF};

fn truncated() -> String {
    unsupported("truncated XLS defined name")
}

/// Built-in names (2.4.150 Name) by their one-character code.
const BUILTIN: [&str; 14] = [
    "Consolidate_Area",
    "Auto_Open",
    "Auto_Close",
    "Extract",
    "Database",
    "Criteria",
    "Print_Area",
    "Print_Titles",
    "Recorder",
    "Data_Form",
    "Auto_Activate",
    "Auto_Deactivate",
    "Sheet_Title",
    "_FilterDatabase",
];

/// Workbook defined names with their scope (`None` = workbook).
#[derive(Default)]
pub(super) struct Names(Vec<(Option<usize>, String, String)>);

impl Names {
    pub(super) fn parse(records: &[Record<'_>], externs: &Externs) -> Result<Self, String> {
        let mut names = Vec::new();
        for record in records
            .iter()
            .take_while(|r| r.kind != EOF)
            .filter(|r| r.kind == 0x0018)
        {
            let data = record.data;
            let flags = u16_at(data, 0)?;
            let count = usize::from(*data.get(3).ok_or_else(truncated)?);
            let cce = usize::from(u16_at(data, 4)?);
            let itab = usize::from(u16_at(data, 8)?);
            let high = *data.get(14).ok_or_else(truncated)?;
            let width = match high {
                0 => 1,
                1 => 2,
                _ => return Err(unsupported("invalid XLS defined name")),
            };
            let raw = data.get(15..15 + count * width).ok_or_else(truncated)?;
            let text: String = if width == 1 {
                raw.iter().map(|&byte| char::from(byte)).collect()
            } else {
                let units: Vec<u16> = raw
                    .chunks_exact(2)
                    .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
                    .collect();
                String::from_utf16(&units).map_err(|_| truncated())?
            };
            let name = if flags & 0x0020 != 0 {
                let code = text.chars().next().map(u32::from).unwrap_or(u32::MAX);
                let builtin = usize::try_from(code)
                    .ok()
                    .and_then(|code| BUILTIN.get(code))
                    .filter(|_| count == 1)
                    .ok_or_else(|| unsupported("unknown XLS built-in name"))?;
                format!("_xlnm.{builtin}")
            } else {
                text
            };
            let start = 15 + count * width;
            let rgce = data.get(start..start + cce).ok_or_else(truncated)?;
            // Macro names (fProc) name procedures, not formulas or ranges;
            // Excel's .xlsx counterparts omit them (sample-1, sample-12:
            // the hidden `_xlfn.IFERROR` future-function names). A name
            // without a formula defines nothing a formula can resolve.
            if flags & 0x0008 != 0 || rgce.is_empty() {
                continue;
            }
            let extra = data.get(start + cce..).ok_or_else(truncated)?;
            let (formula, used) = decompile_name(rgce, extra, externs)?;
            if used != extra.len() {
                return Err(unsupported("unexpected XLS defined name tail"));
            }
            names.push((itab.checked_sub(1), name, formula));
        }
        Ok(Self(names))
    }

    /// Names visible on the zero-based sheet `index`.
    pub(super) fn for_sheet(&self, index: usize) -> Vec<xlsx_model::DefinedName> {
        self.0
            .iter()
            .filter(|(scope, _, _)| scope.is_none_or(|scope| scope == index))
            .map(|(_, name, formula)| xlsx_model::DefinedName {
                name: name.clone(),
                formula: formula.clone(),
            })
            .collect()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn built_in_names_and_scopes_follow_spreadsheetml() {
        // _FilterDatabase (0x0D) local to sheet 2 with a PtgInt formula,
        // a workbook name "Rate" = 5, and a macro name that is skipped.
        let builtin = [
            0x21u8, 0, 0, 1, 3, 0, 0, 0, 2, 0, 0, 0, 0, 0, 0, 0x0d, 0x1e, 1, 0,
        ];
        let user = [
            0u8, 0, 0, 4, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, b'R', b'a', b't', b'e', 0x1e, 5, 0,
        ];
        let mac = [
            0x0au8, 0, 0, 1, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, b'M', 0x1e, 5, 0,
        ];
        let records = [
            Record {
                kind: 0x0018,
                offset: 0,
                data: &builtin,
            },
            Record {
                kind: 0x0018,
                offset: 1,
                data: &user,
            },
            Record {
                kind: 0x0018,
                offset: 2,
                data: &mac,
            },
        ];
        let names = Names::parse(&records, &Externs::default()).unwrap();
        let first: Vec<_> = names
            .for_sheet(0)
            .into_iter()
            .map(|n| (n.name, n.formula))
            .collect();
        assert_eq!(first, vec![("Rate".to_string(), "5".to_string())]);
        let second: Vec<_> = names.for_sheet(1).into_iter().map(|n| n.name).collect();
        assert_eq!(second, vec!["_xlnm._FilterDatabase", "Rate"]);
    }
}
