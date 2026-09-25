//! MS-XLS 2.4.96 DVal and 2.4.95 Dv records as the XLSX model's data
//! validations, as the XLSX parser reads ECMA-376 18.3.1.32
//! `dataValidation` (type, operator, formulas, allowBlank, prompt and error
//! texts, sqref).
//!
//! Formulas are decompiled like conditional-formatting formulas, relative
//! to the top-left cell of the first range. An explicit list (fStrLookup)
//! is one string whose items are NUL-separated; SpreadsheetML writes it as
//! a quoted comma-separated list (observed: items "A\0 B" are saved as
//! `"A, B"`). Excel stores an absent title or message as a single
//! NUL character.
//!
//! The error style, IME mode and the show-prompt/show-error switches only
//! govern data entry; neither the XLSX model nor its parser carries them. A
//! list whose in-cell drop-down is suppressed is rejected: the viewer
//! would draw the drop-down button Excel hides.

use super::conditional::{decompile, Externs};
use super::{u16_at, u32_at, unsupported};

fn truncated() -> String {
    unsupported("truncated XLS data validation")
}

/// XLUnicodeString; a lone NUL is Excel's absent string.
fn text(data: &[u8], offset: usize) -> Result<(Option<String>, usize), String> {
    let count = usize::from(u16_at(data, offset)?);
    let (value, size) = match *data.get(offset + 2).ok_or_else(truncated)? {
        0 => (
            data.get(offset + 3..offset + 3 + count)
                .ok_or_else(truncated)?
                .iter()
                .map(|&byte| char::from(byte))
                .collect::<String>(),
            count,
        ),
        1 => {
            let bytes = data
                .get(offset + 3..offset + 3 + count * 2)
                .ok_or_else(truncated)?;
            let units: Vec<u16> = bytes
                .chunks_exact(2)
                .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
                .collect();
            (
                String::from_utf16(&units).map_err(|_| truncated())?,
                count * 2,
            )
        }
        _ => return Err(unsupported("invalid XLS data validation text")),
    };
    let value = (!value.is_empty() && value != "\0").then_some(value);
    Ok((value, 3 + size))
}

fn column(index: u16) -> String {
    let mut index = u32::from(index) + 1;
    let mut letters = Vec::new();
    while index > 0 {
        let rem = (index - 1) % 26;
        letters.push(char::from(b'A' + rem as u8));
        index = (index - 1) / 26;
    }
    letters.iter().rev().collect()
}

/// Project a sheet's DVal and Dv records (kinds 0x01B2 / 0x01BE).
pub(super) fn project<'a>(
    records: impl Iterator<Item = (u16, &'a [u8])>,
    externs: &Externs,
) -> Result<Vec<xlsx_model::DataValidation>, String> {
    let mut output = Vec::new();
    let mut expected: Option<usize> = None;
    for (kind, data) in records {
        match kind {
            0x01b2 => {
                if expected.is_some() || data.len() != 18 {
                    return Err(unsupported("invalid XLS data validation header"));
                }
                let count = usize::try_from(u32_at(data, 14)?).map_err(|_| truncated())?;
                if count > 65534 {
                    return Err(unsupported("invalid XLS data validation count"));
                }
                expected = Some(count);
            }
            0x01be => {
                let left = expected
                    .as_mut()
                    .filter(|left| **left > 0)
                    .ok_or_else(|| unsupported("XLS data validation outside its DVal"))?;
                *left -= 1;
                output.push(dv(data, externs)?);
            }
            _ => return Err(unsupported("unexpected XLS data validation record")),
        }
    }
    if expected.is_some_and(|left| left != 0) {
        return Err(unsupported("XLS DVal lacks its data validations"));
    }
    Ok(output)
}

fn dv(data: &[u8], externs: &Externs) -> Result<xlsx_model::DataValidation, String> {
    let flags = u32_at(data, 0)?;
    let kind = flags & 0x0f;
    let lookup = flags & 0x80 != 0;
    let allow_blank = flags & 0x0100 != 0;
    let suppress = flags & 0x0200 != 0;
    let operator = (flags >> 20) & 0x0f;
    let mut at = 4;
    let mut texts = Vec::with_capacity(4);
    for _ in 0..4 {
        let (value, size) = text(data, at)?;
        texts.push(value);
        at += size;
    }
    let mut formulas = Vec::with_capacity(2);
    for _ in 0..2 {
        let cce = usize::from(u16_at(data, at)?);
        let rgce = data.get(at + 4..at + 4 + cce).ok_or_else(truncated)?;
        formulas.push(rgce);
        at += 4 + cce;
    }
    // SqRefU (2.5.249).
    let count = usize::from(u16_at(data, at)?);
    if count == 0 || count > 432 || data.len() != at + 2 + count * 8 {
        return Err(unsupported("invalid XLS data validation ranges"));
    }
    let mut ranges = Vec::with_capacity(count);
    let mut anchor = None;
    for index in 0..count {
        let base = at + 2 + index * 8;
        let (row_first, row_last) = (u16_at(data, base)?, u16_at(data, base + 2)?);
        let (col_first, col_last) = (u16_at(data, base + 4)?, u16_at(data, base + 6)?);
        if row_first > row_last || col_first > col_last || col_last > 0x00ff {
            return Err(unsupported("invalid XLS data validation range"));
        }
        anchor.get_or_insert((row_first, col_first));
        let first = format!("{}{}", column(col_first), u32::from(row_first) + 1);
        ranges.push(if (row_first, col_first) == (row_last, col_last) {
            first
        } else {
            format!("{first}:{}{}", column(col_last), u32::from(row_last) + 1)
        });
    }
    let anchor = anchor.expect("at least one range");
    let validation_type = match kind {
        0 => None,
        1 => Some("whole"),
        2 => Some("decimal"),
        3 => Some("list"),
        4 => Some("date"),
        5 => Some("time"),
        6 => Some("textLength"),
        7 => Some("custom"),
        _ => return Err(unsupported("invalid XLS data validation type")),
    };
    if kind == 3 && suppress {
        return Err(unsupported(
            "XLS list validation without its drop-down is not representable",
        ));
    }
    let bounded = matches!(kind, 1 | 2 | 4 | 5 | 6);
    // ST_DataValidationOperator; SpreadsheetML omits the default `between`.
    let operator = if bounded {
        match operator {
            0 => None,
            1 => Some("notBetween"),
            2 => Some("equal"),
            3 => Some("notEqual"),
            4 => Some("greaterThan"),
            5 => Some("lessThan"),
            6 => Some("greaterThanOrEqual"),
            7 => Some("lessThanOrEqual"),
            _ => return Err(unsupported("invalid XLS data validation operator")),
        }
    } else {
        None
    };
    let formula = |rgce: &[u8]| -> Result<Option<String>, String> {
        if rgce.is_empty() {
            return Ok(None);
        }
        decompile(rgce, anchor, externs).map(Some)
    };
    let formula1 = if kind == 0 {
        None
    } else if kind == 3 && lookup {
        // One PtgStr of NUL-separated items.
        if formulas[0].first() != Some(&0x17) {
            return Err(unsupported("invalid XLS explicit list validation"));
        }
        formula(formulas[0])?.map(|text| text.replace('\0', ","))
    } else {
        formula(formulas[0])?
    };
    let two = bounded && operator.is_none_or(|op| op == "notBetween");
    let formula2 = if two { formula(formulas[1])? } else { None };
    if (kind != 0 && formula1.is_none()) || (two && formula2.is_none()) {
        return Err(unsupported("XLS data validation lacks its formula"));
    }
    let [prompt_title, error_title, prompt, error_message]: [Option<String>; 4] =
        texts.try_into().expect("four texts");
    Ok(xlsx_model::DataValidation {
        sqref: ranges.join(" "),
        validation_type: validation_type.map(str::to_string),
        operator: operator.map(str::to_string),
        formula1,
        formula2,
        allow_blank,
        prompt_title,
        prompt,
        error_title,
        error_message,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn string(text: &str) -> Vec<u8> {
        let mut data = (text.len() as u16).to_le_bytes().to_vec();
        data.push(0);
        data.extend(text.as_bytes());
        data
    }

    fn dv(flags: u32, formula1: &[u8], ranges: &[[u16; 4]]) -> Vec<u8> {
        let mut data = flags.to_le_bytes().to_vec();
        for text in ["\0", "\0", "Pick one", "\0"] {
            data.extend(string(text));
        }
        data.extend((formula1.len() as u16).to_le_bytes());
        data.extend([0, 0]);
        data.extend(formula1);
        data.extend([0, 0, 0, 0]);
        data.extend((ranges.len() as u16).to_le_bytes());
        for range in ranges {
            for value in range {
                data.extend(value.to_le_bytes());
            }
        }
        data
    }

    fn dval(count: u32) -> Vec<u8> {
        let mut data = vec![0u8; 14];
        data.extend(count.to_le_bytes());
        data
    }

    #[test]
    fn explicit_lists_and_bounds_project_as_spreadsheetml() {
        // List (3) with fStrLookup and fAllowBlank: "a\0b" on C4:C5.
        let list = dv(0x0183, &[0x17, 3, 0, b'a', 0, b'b'], &[[3, 4, 2, 2]]);
        // Whole number (1) >= 0 (operator 6) on C7.
        let whole = dv(0x0060_0001, &[0x1e, 0, 0], &[[6, 6, 2, 2]]);
        let records = [dval(2), list, whole];
        let output = project(
            records
                .iter()
                .enumerate()
                .map(|(i, data)| (if i == 0 { 0x01b2 } else { 0x01be }, data.as_slice())),
            &Externs::default(),
        )
        .unwrap();
        assert_eq!(output[0].sqref, "C4:C5");
        assert_eq!(output[0].validation_type.as_deref(), Some("list"));
        assert_eq!(output[0].formula1.as_deref(), Some("\"a,b\""));
        assert!(output[0].allow_blank);
        assert_eq!(output[0].prompt.as_deref(), Some("Pick one"));
        assert!(output[0].prompt_title.is_none());
        assert_eq!(output[1].operator.as_deref(), Some("greaterThanOrEqual"));
        assert_eq!(output[1].formula1.as_deref(), Some("0"));
        // A list whose drop-down is suppressed, and a missing Dv, reject.
        let suppressed = dv(0x0283, &[0x17, 1, 0, b'a'], &[[0, 0, 0, 0]]);
        let records = [dval(1), suppressed];
        assert!(project(
            records
                .iter()
                .enumerate()
                .map(|(i, data)| { (if i == 0 { 0x01b2 } else { 0x01be }, data.as_slice()) }),
            &Externs::default(),
        )
        .is_err());
        let records = [dval(1)];
        assert!(project(
            records.iter().map(|data| (0x01b2, data.as_slice())),
            &Externs::default()
        )
        .is_err());
    }
}
