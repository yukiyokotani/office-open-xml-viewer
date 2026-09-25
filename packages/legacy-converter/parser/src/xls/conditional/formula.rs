//! Conditional-formatting formulas (MS-XLS 2.5.198.6 CFParsedFormulaNoCCE)
//! decompiled to the SpreadsheetML formula text the XLSX model carries
//! (ECMA-376 18.3.1.10 `cfRule/formula`, 18.17 formula grammar).
//!
//! The Rgce (2.5.198.104) is a postfix token stream. Relative references
//! (PtgRefN 2.5.198.88, PtgAreaN 2.5.198.31) are offsets from the cell being
//! evaluated (RgceLocRel 2.5.198.111, RgceAreaRel 2.5.198.106): rows wrap
//! modulo 0x10000 and columns modulo 0x100. The XLSX text writes them for
//! the top-left cell of the rule's range (`refBound`); the Excel-saved .xlsx
//! counterparts in the private corpus agree token for token.
//!
//! Tokens outside the set below reject the rule instead of guessing a text.

use super::super::{f64_at, u16_at, u32_at, unsupported, Record};
use super::ftab;

/// Analysis ToolPak functions that BIFF8 calls through the add-in SupBook
/// (PtgNameX + PtgFuncVar 0x00FF) and that ECMA-376 18.17.7 defines as
/// predefined worksheet functions, so SpreadsheetML writes them by name.
const TOOLPAK: [&str; 91] = [
    "ACCRINT",
    "ACCRINTM",
    "AMORDEGRC",
    "AMORLINC",
    "BESSELI",
    "BESSELJ",
    "BESSELK",
    "BESSELY",
    "BIN2DEC",
    "BIN2HEX",
    "BIN2OCT",
    "COMPLEX",
    "CONVERT",
    "COUPDAYBS",
    "COUPDAYS",
    "COUPDAYSNC",
    "COUPNCD",
    "COUPNUM",
    "COUPPCD",
    "CUMIPMT",
    "CUMPRINC",
    "DEC2BIN",
    "DEC2HEX",
    "DEC2OCT",
    "DELTA",
    "DISC",
    "DOLLARDE",
    "DOLLARFR",
    "DURATION",
    "EDATE",
    "EFFECT",
    "EOMONTH",
    "ERF",
    "ERFC",
    "FACTDOUBLE",
    "FVSCHEDULE",
    "GCD",
    "GESTEP",
    "HEX2BIN",
    "HEX2DEC",
    "HEX2OCT",
    "IMABS",
    "IMAGINARY",
    "IMARGUMENT",
    "IMCONJUGATE",
    "IMCOS",
    "IMDIV",
    "IMEXP",
    "IMLN",
    "IMLOG10",
    "IMLOG2",
    "IMPOWER",
    "IMPRODUCT",
    "IMREAL",
    "IMSIN",
    "IMSQRT",
    "IMSUB",
    "IMSUM",
    "INTRATE",
    "ISEVEN",
    "ISODD",
    "LCM",
    "MDURATION",
    "MROUND",
    "MULTINOMIAL",
    "NETWORKDAYS",
    "NOMINAL",
    "OCT2BIN",
    "OCT2DEC",
    "OCT2HEX",
    "ODDFPRICE",
    "ODDFYIELD",
    "ODDLPRICE",
    "ODDLYIELD",
    "PRICE",
    "PRICEDISC",
    "PRICEMAT",
    "QUOTIENT",
    "RANDBETWEEN",
    "RECEIVED",
    "SERIESSUM",
    "SQRTPI",
    "TBILLEQ",
    "TBILLPRICE",
    "TBILLYIELD",
    "WEEKNUM",
    "WORKDAY",
    "XIRR",
    "XNPV",
    "YEARFRAC",
    "YIELD",
];

/// Add-in function names reachable through EXTERNSHEET XTIs (2.4.106):
/// for each XTI, the UDF names (AddinUdf 2.5.1) of its SupBook (2.4.271)
/// when that SupBook is the add-in marker (cch 0x3A01).
#[derive(Default)]
pub(in super::super) struct Externs {
    xti_names: Vec<Option<Vec<String>>>,
    /// Lbl (2.4.150) names in record order; built-in names are `None`.
    names: Vec<Option<String>>,
    /// Formula sheet prefix (quoted when needed) per XTI of this workbook's
    /// own SupBook that names exactly one sheet.
    xti_sheets: Vec<Option<String>>,
}

impl Externs {
    pub(in super::super) fn parse(records: &[Record<'_>]) -> Result<Self, String> {
        let mut books: Vec<Option<Vec<String>>> = Vec::new();
        let mut xtis = Vec::new();
        let mut defined = Vec::new();
        let mut own = Vec::new();
        let mut sheet_names = Vec::new();
        let mut xti_tabs = Vec::new();
        for record in records.iter().take_while(|r| r.kind != super::super::EOF) {
            match record.kind {
                0x01ae => {
                    let addin = u16_at(record.data, 2).is_ok_and(|cch| cch == 0x3a01);
                    books.push(addin.then(Vec::new));
                    own.push(u16_at(record.data, 2).is_ok_and(|cch| cch == 0x0401));
                }
                0x0085 => {
                    // BoundSheet8: lbPlyPos, flags, stName (ShortXLUnicodeString).
                    let data = record.data;
                    let count = usize::from(*data.get(6).ok_or_else(truncated)?);
                    let name = match *data.get(7).ok_or_else(truncated)? {
                        0 => data
                            .get(8..8 + count)
                            .ok_or_else(truncated)?
                            .iter()
                            .map(|&byte| char::from(byte))
                            .collect::<String>(),
                        _ => {
                            let bytes = data.get(8..8 + count * 2).ok_or_else(truncated)?;
                            let units: Vec<u16> = bytes
                                .chunks_exact(2)
                                .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
                                .collect();
                            String::from_utf16(&units).map_err(|_| truncated())?
                        }
                    };
                    sheet_names.push(name);
                }
                0x0023 => {
                    // ExternName 2.4.105 of an add-in SupBook: flags (all
                    // zero for a UDF), AddinUdf reserved (4), udfName.
                    let Some(Some(names)) = books.last_mut() else {
                        continue;
                    };
                    let data = record.data;
                    if u16_at(data, 0)? != 0 || u32_at(data, 2)? != 0 {
                        return Err(unsupported("invalid XLS add-in function name"));
                    }
                    let count = usize::from(*data.get(6).ok_or_else(truncated)?);
                    let high = *data.get(7).ok_or_else(truncated)?;
                    let name = match high {
                        0 => data
                            .get(8..8 + count)
                            .ok_or_else(truncated)?
                            .iter()
                            .map(|&byte| char::from(byte))
                            .collect(),
                        1 => {
                            let bytes = data.get(8..8 + count * 2).ok_or_else(truncated)?;
                            let units: Vec<u16> = bytes
                                .chunks_exact(2)
                                .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
                                .collect();
                            String::from_utf16(&units).map_err(|_| truncated())?
                        }
                        _ => return Err(unsupported("invalid XLS add-in function name")),
                    };
                    names.push(name);
                }
                0x0017 => {
                    let count = usize::from(u16_at(record.data, 0)?);
                    for index in 0..count {
                        xtis.push(u16_at(record.data, 2 + index * 6)?);
                        xti_tabs.push((
                            u16_at(record.data, 4 + index * 6)?,
                            u16_at(record.data, 6 + index * 6)?,
                        ));
                    }
                }
                0x0018 => {
                    // Lbl: flags (fBuiltin bit 5), chKey, cch, ..., Name at 14
                    // (XLUnicodeStringNoCch).
                    let data = record.data;
                    let builtin = u16_at(data, 0)? & 0x0020 != 0;
                    let count = usize::from(*data.get(3).ok_or_else(truncated)?);
                    let name = match *data.get(14).ok_or_else(truncated)? {
                        0 => data
                            .get(15..15 + count)
                            .ok_or_else(truncated)?
                            .iter()
                            .map(|&byte| char::from(byte))
                            .collect::<String>(),
                        1 => {
                            let bytes = data.get(15..15 + count * 2).ok_or_else(truncated)?;
                            let units: Vec<u16> = bytes
                                .chunks_exact(2)
                                .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
                                .collect();
                            String::from_utf16(&units).map_err(|_| truncated())?
                        }
                        _ => return Err(unsupported("invalid XLS defined name")),
                    };
                    defined.push((!builtin && !name.is_empty()).then_some(name));
                }
                _ => {}
            }
        }
        Ok(Self {
            xti_names: xtis
                .iter()
                .map(|book| books.get(usize::from(*book)).cloned().flatten())
                .collect(),
            names: defined,
            xti_sheets: xtis
                .iter()
                .zip(&xti_tabs)
                .map(|(book, (first, last))| {
                    let own = own.get(usize::from(*book)).copied().unwrap_or(false);
                    (own && first == last)
                        .then(|| sheet_names.get(usize::from(*first)))
                        .flatten()
                        .map(|name| sheet_prefix(name))
                })
                .collect(),
        })
    }

    fn sheet(&self, xti: u16) -> Option<&str> {
        self.xti_sheets.get(usize::from(xti))?.as_deref()
    }

    /// A one-based Lbl index as the user-defined name it names.
    fn name(&self, index: u32) -> Option<&str> {
        let index = usize::try_from(index).ok()?.checked_sub(1)?;
        self.names.get(index)?.as_deref()
    }

    fn addin(&self, xti: u16, index: u32) -> Option<&str> {
        let names = self.xti_names.get(usize::from(xti))?.as_ref()?;
        let name = names.get(usize::try_from(index).ok()?.checked_sub(1)?)?;
        Some(name.as_str())
    }
}

fn truncated() -> String {
    unsupported("truncated XLS conditional formatting formula")
}

/// A sheet name as a formula prefix: quoted (with doubled apostrophes)
/// unless it is a plain identifier. Excel's own .xlsx files leave names of
/// letters (including CJK), digits, `_` and `.` unquoted (sample-3:
/// `夏休み!$C$4`) and quote names with spaces (`'WATERFALL CHART'!…`); a
/// name that begins with a digit or reads as an A1 reference is quoted.
fn sheet_prefix(name: &str) -> String {
    let identifier = name
        .chars()
        .next()
        .is_some_and(|first| first.is_alphabetic() || first == '_')
        && name
            .chars()
            .all(|c| c.is_alphanumeric() || c == '_' || c == '.');
    let letters = name.chars().take_while(|c| c.is_ascii_alphabetic()).count();
    let cell_like = (1..=3).contains(&letters)
        && name.len() > letters
        && name[letters..].chars().all(|c| c.is_ascii_digit());
    if identifier && !cell_like {
        name.to_string()
    } else {
        format!("'{}'", name.replace('\'', "''"))
    }
}

fn reject() -> String {
    unsupported("unsupported XLS conditional formatting formula token")
}

/// Operator precedence is carried by the token order and explicit PtgParen,
/// so text is built without inserting parentheses.
enum Item {
    Text(String),
    /// A PtgNameX add-in function name, valid only as a UDF call target.
    Function(String),
}

/// Column letters for a zero-based column (at most 0xFF in BIFF8).
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

/// One end of a reference. `relative_offsets` selects RgceLocRel semantics.
fn cell(row: u16, col: u16, anchor: (u16, u16), relative_offsets: bool) -> String {
    let row_relative = col & 0x8000 != 0;
    let col_relative = col & 0x4000 != 0;
    let col = col & 0x3fff;
    let (row, col) = if relative_offsets {
        (
            if row_relative {
                anchor.0.wrapping_add(row)
            } else {
                row
            },
            if col_relative {
                (anchor.1.wrapping_add(col)) & 0x00ff
            } else {
                col
            },
        )
    } else {
        (row, col)
    };
    format!(
        "{}{}{}{}",
        if col_relative { "" } else { "$" },
        column(col),
        if row_relative { "" } else { "$" },
        u32::from(row) + 1
    )
}

/// An area as text; an area spanning every BIFF8 column (0..=0xFF) or row
/// (0..=0xFFFF) with absolute bounds is written as a whole-row or
/// whole-column range, as Excel writes it in SpreadsheetML (sample-1:
/// `$A$3:$IV$3` is saved as `$3:$3`, sample-4: `$C$1:$C$65536` as `$C:$C`).
fn area_text(
    rows: (u16, u16),
    cols: (u16, u16),
    anchor: (u16, u16),
    relative_offsets: bool,
) -> String {
    let col_absolute = cols.0 & 0x4000 == 0 && cols.1 & 0x4000 == 0;
    let row_absolute = cols.0 & 0x8000 == 0 && cols.1 & 0x8000 == 0;
    if col_absolute && cols.0 & 0x3fff == 0 && cols.1 & 0x3fff == 0xff {
        let row = |row: u16, col: u16| {
            let relative = col & 0x8000 != 0;
            let row = if relative && relative_offsets {
                anchor.0.wrapping_add(row)
            } else {
                row
            };
            format!("{}{}", if relative { "" } else { "$" }, u32::from(row) + 1)
        };
        return format!("{}:{}", row(rows.0, cols.0), row(rows.1, cols.1));
    }
    if row_absolute && rows.0 == 0 && rows.1 == 0xffff {
        let col = |col: u16| {
            let relative = col & 0x4000 != 0;
            let index = if relative && relative_offsets {
                anchor.1.wrapping_add(col & 0x3fff) & 0x00ff
            } else {
                col & 0x3fff
            };
            format!("{}{}", if relative { "" } else { "$" }, column(index))
        };
        return format!("{}:{}", col(cols.0), col(cols.1));
    }
    format!(
        "{}:{}",
        cell(rows.0, cols.0, anchor, relative_offsets),
        cell(rows.1, cols.1, anchor, relative_offsets)
    )
}

/// ECMA-376 18.17.2.x number literal: the shortest round-trip decimal.
fn number(value: f64) -> Result<String, String> {
    if !value.is_finite() {
        return Err(unsupported("invalid XLS conditional formatting number"));
    }
    Ok(format!("{value}"))
}

/// Decompile a CFParsedFormulaNoCCE for a rule anchored at `anchor`
/// (zero-based row, column).
pub(in super::super) fn decompile(
    rgce: &[u8],
    anchor: (u16, u16),
    externs: &Externs,
) -> Result<String, String> {
    decompile_with(rgce, &[], anchor, externs, false)
}

/// Decompile a NameParsedFormula (2.5.198.21): its 3-D references are
/// RgceLocRel (2.5.198.85), which SpreadsheetML writes relative to A1;
/// `extra` is the RgbExtra (2.5.198.103) holding array constants.
pub(in super::super) fn decompile_name(
    rgce: &[u8],
    extra: &[u8],
    externs: &Externs,
) -> Result<(String, usize), String> {
    let mut used = 0;
    let text = decompile_inner(rgce, extra, &mut used, (0, 0), externs, true)?;
    Ok((text, used))
}

fn decompile_with(
    rgce: &[u8],
    extra: &[u8],
    anchor: (u16, u16),
    externs: &Externs,
    name: bool,
) -> Result<String, String> {
    let mut used = 0;
    let text = decompile_inner(rgce, extra, &mut used, anchor, externs, name)?;
    if used != extra.len() {
        return Err(reject());
    }
    Ok(text)
}

/// PtgExtraArray (2.5.198.59) at `extra[*used..]` as an ECMA-376
/// 18.17.2 array constant: columns separated by `,`, rows by `;`.
fn array(extra: &[u8], used: &mut usize) -> Result<String, String> {
    let mut at = *used;
    let cols = usize::from(*extra.get(at).ok_or_else(truncated)?) + 1;
    let rows = usize::from(u16_at(extra, at + 1)?) + 1;
    at += 3;
    let mut text = String::from("{");
    for row in 0..rows {
        if row > 0 {
            text.push(';');
        }
        for col in 0..cols {
            if col > 0 {
                text.push(',');
            }
            // SerAr (2.5.198.112).
            match *extra.get(at).ok_or_else(truncated)? {
                0x00 => at += 9,
                0x01 => {
                    text.push_str(&number(f64_at(extra, at + 1)?)?);
                    at += 9;
                }
                0x02 => {
                    let count = usize::from(u16_at(extra, at + 1)?);
                    let (value, size) = match *extra.get(at + 3).ok_or_else(truncated)? {
                        0 => (
                            extra
                                .get(at + 4..at + 4 + count)
                                .ok_or_else(truncated)?
                                .iter()
                                .map(|&byte| char::from(byte))
                                .collect::<String>(),
                            count,
                        ),
                        1 => {
                            let bytes = extra
                                .get(at + 4..at + 4 + count * 2)
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
                        _ => return Err(reject()),
                    };
                    text.push_str(&format!("\"{}\"", value.replace('"', "\"\"")));
                    at += 4 + size;
                }
                0x04 => {
                    text.push_str(if *extra.get(at + 1).ok_or_else(truncated)? != 0 {
                        "TRUE"
                    } else {
                        "FALSE"
                    });
                    at += 9;
                }
                _ => return Err(reject()),
            }
        }
    }
    text.push('}');
    *used = at;
    Ok(text)
}

fn decompile_inner(
    rgce: &[u8],
    extra: &[u8],
    used: &mut usize,
    anchor: (u16, u16),
    externs: &Externs,
    name: bool,
) -> Result<String, String> {
    let mut stack: Vec<Item> = Vec::new();
    let mut space = String::new();
    let mut at = 0;
    let pop = |stack: &mut Vec<Item>| -> Result<String, String> {
        match stack.pop() {
            Some(Item::Text(text)) => Ok(text),
            _ => Err(reject()),
        }
    };
    while at < rgce.len() {
        let token = rgce[at];
        let pending = std::mem::take(&mut space);
        match token {
            0x03..=0x0e => {
                let right = pop(&mut stack)?;
                let left = pop(&mut stack)?;
                let op = [
                    "+", "-", "*", "/", "^", "&", "<", "<=", "=", ">=", ">", "<>",
                ][usize::from(token - 0x03)];
                stack.push(Item::Text(format!("{left}{pending}{op}{right}")));
                at += 1;
            }
            0x12 | 0x13 => {
                let value = pop(&mut stack)?;
                let op = if token == 0x12 { "+" } else { "-" };
                stack.push(Item::Text(format!("{pending}{op}{value}")));
                at += 1;
            }
            0x14 => {
                let value = pop(&mut stack)?;
                stack.push(Item::Text(format!("{value}{pending}%")));
                at += 1;
            }
            0x15 => {
                let value = pop(&mut stack)?;
                stack.push(Item::Text(format!("{pending}({value})")));
                at += 1;
            }
            0x16 => {
                stack.push(Item::Text(pending));
                at += 1;
            }
            0x17 => {
                // PtgStr: ShortXLUnicodeString.
                let count = usize::from(*rgce.get(at + 1).ok_or_else(truncated)?);
                let high = *rgce.get(at + 2).ok_or_else(truncated)?;
                let (text, size) = match high {
                    0 => (
                        rgce.get(at + 3..at + 3 + count)
                            .ok_or_else(truncated)?
                            .iter()
                            .map(|&byte| char::from(byte))
                            .collect::<String>(),
                        count,
                    ),
                    1 => {
                        let bytes = rgce.get(at + 3..at + 3 + count * 2).ok_or_else(truncated)?;
                        let units: Vec<u16> = bytes
                            .chunks_exact(2)
                            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
                            .collect();
                        (
                            String::from_utf16(&units).map_err(|_| truncated())?,
                            count * 2,
                        )
                    }
                    _ => return Err(reject()),
                };
                stack.push(Item::Text(format!(
                    "{pending}\"{}\"",
                    text.replace('"', "\"\"")
                )));
                at += 3 + size;
            }
            0x19 => {
                // PtgAttr family (2.5.198.33ff).
                let kind = *rgce.get(at + 1).ok_or_else(truncated)?;
                match kind {
                    // PtgAttrSemi: volatile marker, no text.
                    0x01 => at += 4,
                    // PtgAttrSpace type 0 (spaces before the next token) and
                    // type 2 (spaces before an opening parenthesis); both
                    // precede the next token's own text.
                    0x40 => {
                        let kind = *rgce.get(at + 2).ok_or_else(truncated)?;
                        let count = usize::from(*rgce.get(at + 3).ok_or_else(truncated)?);
                        if !matches!(kind, 0x00 | 0x02) {
                            return Err(reject());
                        }
                        space = format!("{pending}{}", " ".repeat(count));
                        at += 4;
                    }
                    _ => return Err(reject()),
                }
            }
            0x1c => {
                let error = match *rgce.get(at + 1).ok_or_else(truncated)? {
                    0x00 => "#NULL!",
                    0x07 => "#DIV/0!",
                    0x0f => "#VALUE!",
                    0x17 => "#REF!",
                    0x1d => "#NAME?",
                    0x24 => "#NUM!",
                    0x2a => "#N/A",
                    _ => return Err(reject()),
                };
                stack.push(Item::Text(format!("{pending}{error}")));
                at += 2;
            }
            0x1d => {
                let value = *rgce.get(at + 1).ok_or_else(truncated)?;
                stack.push(Item::Text(format!(
                    "{pending}{}",
                    if value != 0 { "TRUE" } else { "FALSE" }
                )));
                at += 2;
            }
            0x1e => {
                stack.push(Item::Text(format!("{pending}{}", u16_at(rgce, at + 1)?)));
                at += 3;
            }
            0x1f => {
                stack.push(Item::Text(format!(
                    "{pending}{}",
                    number(f64_at(rgce, at + 1)?)?
                )));
                at += 9;
            }
            0x20..=0x7f => match token & 0x1f {
                // PtgFunc: fixed argument count from Ftab.
                0x01 => {
                    let (name, count) = ftab::function(u16_at(rgce, at + 1)?).ok_or_else(reject)?;
                    let count = count.ok_or_else(reject)?;
                    let call = call(&mut stack, name, usize::from(count), &pending)?;
                    stack.push(Item::Text(call));
                    at += 3;
                }
                // PtgFuncVar: explicit count; tab 0x00FF calls the function
                // named by the first argument (a PtgNameX add-in name).
                0x02 => {
                    let count = usize::from(*rgce.get(at + 1).ok_or_else(truncated)? & 0x7f);
                    let tab = u16_at(rgce, at + 2)?;
                    if tab & 0x8000 != 0 {
                        return Err(reject());
                    }
                    let text = if tab == 0x00ff {
                        let arguments =
                            arguments(&mut stack, count.checked_sub(1).ok_or_else(reject)?)?;
                        let Some(Item::Function(name)) = stack.pop() else {
                            return Err(reject());
                        };
                        format!("{pending}{name}({})", arguments.join(","))
                    } else {
                        let (name, _) = ftab::function(tab).ok_or_else(reject)?;
                        call(&mut stack, name, count, &pending)?
                    };
                    stack.push(Item::Text(text));
                    at += 4;
                }
                // PtgRef / PtgRefN.
                0x04 | 0x0c => {
                    let text = cell(
                        u16_at(rgce, at + 1)?,
                        u16_at(rgce, at + 3)?,
                        anchor,
                        token & 0x1f == 0x0c,
                    );
                    stack.push(Item::Text(format!("{pending}{text}")));
                    at += 5;
                }
                // PtgArea / PtgAreaN.
                0x05 | 0x0d => {
                    let relative = token & 0x1f == 0x0d;
                    let text = area_text(
                        (u16_at(rgce, at + 1)?, u16_at(rgce, at + 3)?),
                        (u16_at(rgce, at + 5)?, u16_at(rgce, at + 7)?),
                        anchor,
                        relative,
                    );
                    stack.push(Item::Text(format!("{pending}{text}")));
                    at += 9;
                }
                // PtgArray (2.5.198.32): its values are the next
                // PtgExtraArray of the RgbExtra.
                0x00 => {
                    rgce.get(at + 1..at + 8).ok_or_else(truncated)?;
                    let text = array(extra, used)?;
                    stack.push(Item::Text(format!("{pending}{text}")));
                    at += 8;
                }
                // PtgRef3d / PtgArea3d on one sheet of this workbook.
                0x1a | 0x1b => {
                    let sheet = externs
                        .sheet(u16_at(rgce, at + 1)?)
                        .ok_or_else(|| unsupported("unsupported XLS 3-D reference"))?;
                    let text = if token & 0x1f == 0x1a {
                        cell(u16_at(rgce, at + 3)?, u16_at(rgce, at + 5)?, anchor, name)
                    } else {
                        area_text(
                            (u16_at(rgce, at + 3)?, u16_at(rgce, at + 5)?),
                            (u16_at(rgce, at + 7)?, u16_at(rgce, at + 9)?),
                            anchor,
                            name,
                        )
                    };
                    stack.push(Item::Text(format!("{pending}{sheet}!{text}")));
                    at += if token & 0x1f == 0x1a { 7 } else { 11 };
                }
                // PtgName (2.5.198.76): a user-defined name, written by name
                // as in SpreadsheetML (sample-2: a data-validation list
                // `人リスト`). Built-in names are not projected.
                0x03 => {
                    let name = externs
                        .name(u32_at(rgce, at + 1)?)
                        .ok_or_else(|| unsupported("unsupported XLS formula defined name"))?;
                    stack.push(Item::Text(format!("{pending}{name}")));
                    at += 5;
                }
                // PtgRefErr / PtgAreaErr: #REF!, whose payload MUST be
                // ignored. sample-3 has CF12 rules whose PtgRefErr keeps a
                // relative (0, -1) payload; Excel's .xlsx counterpart prints
                // it as B5, but Excel's own PDF of the .xls shows those rules
                // never apply (no bold day numbers), matching #REF!.
                0x0a => {
                    stack.push(Item::Text(format!("{pending}#REF!")));
                    at += 5;
                }
                0x0b => {
                    stack.push(Item::Text(format!("{pending}#REF!")));
                    at += 9;
                }
                // PtgNameX: CFParsedFormulaNoCCE excludes it, yet Excel
                // stores Analysis ToolPak calls in CF formulas this way
                // (sample corpus: EOMONTH, written as EOMONTH(...) in the
                // Excel-saved .xlsx). Only those predefined names resolve.
                0x19 => {
                    let name = externs
                        .addin(u16_at(rgce, at + 1)?, u32_at(rgce, at + 3)?)
                        .filter(|name| TOOLPAK.contains(name))
                        .ok_or_else(|| {
                            unsupported("unsupported XLS conditional formatting add-in function")
                        })?;
                    if !pending.is_empty() {
                        return Err(reject());
                    }
                    stack.push(Item::Function(name.to_string()));
                    at += 7;
                }
                _ => return Err(reject()),
            },
            _ => return Err(reject()),
        }
    }
    if !space.is_empty() {
        return Err(reject());
    }
    match (stack.pop(), stack.is_empty()) {
        (Some(Item::Text(text)), true) => Ok(text),
        _ => Err(unsupported("malformed XLS conditional formatting formula")),
    }
}

fn arguments(stack: &mut Vec<Item>, count: usize) -> Result<Vec<String>, String> {
    let start = stack.len().checked_sub(count).ok_or_else(reject)?;
    stack
        .drain(start..)
        .map(|item| match item {
            Item::Text(text) => Ok(text),
            Item::Function(_) => Err(reject()),
        })
        .collect()
}

fn call(stack: &mut Vec<Item>, name: &str, count: usize, pending: &str) -> Result<String, String> {
    let arguments = arguments(stack, count)?;
    Ok(format!("{pending}{name}({})", arguments.join(",")))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn decompiles_relative_references_operators_and_functions() {
        // MONTH(B5)<>MONTH(C5) anchored at C5: PtgRefN offsets (0,-1), (0,0).
        let rgce = [
            0x4c, 0x00, 0x00, 0xff, 0xff, 0x41, 0x44, 0x00, 0x4c, 0x00, 0x00, 0x00, 0xc0, 0x41,
            0x44, 0x00, 0x0e,
        ];
        assert_eq!(
            decompile(&rgce, (4, 2), &Externs::default()).unwrap(),
            "MONTH(B5)<>MONTH(C5)"
        );
        // AND($C10="x",I$7>=1) with mixed absolute parts.
        let mut rgce = vec![0x4c, 0x00, 0x00, 0x02, 0x80, 0x17, 0x01, 0x00, b'x', 0x0b];
        rgce.extend([0x4c, 0x06, 0x00, 0x00, 0x40, 0x1e, 0x01, 0x00, 0x0c]);
        rgce.extend([0x42, 0x02, 0x24, 0x00]);
        assert_eq!(
            decompile(&rgce, (9, 8), &Externs::default()).unwrap(),
            "AND($C10=\"x\",I$7>=1)"
        );
    }

    #[test]
    fn resolves_toolpak_addins_and_rejects_other_names() {
        let externs = Externs {
            xti_names: vec![Some(vec!["EOMONTH".into(), "MYUDF".into()])],
            names: vec![Some("List".into()), None],
            xti_sheets: vec![Some("Data".into())],
        };
        // 'Data'!$B$4:$I$49 through a name's PtgArea3d.
        let area = [0x3b, 0, 0, 3, 0, 48, 0, 1, 0xc0, 8, 0xc0];
        assert_eq!(
            decompile_name(&area, &[], &externs).unwrap().0,
            "Data!B4:I49"
        );
        assert_eq!(
            decompile(&[0x23, 1, 0, 0, 0], (0, 0), &externs).unwrap(),
            "List"
        );
        assert!(decompile(&[0x23, 2, 0, 0, 0], (0, 0), &externs).is_err());
        // EOMONTH($I$7,0): NameX, ref, int, FuncVar(3, 0xFF).
        let rgce = [
            0x39, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x24, 0x06, 0x00, 0x08, 0x00, 0x1e, 0x00,
            0x00, 0x42, 0x03, 0xff, 0x00,
        ];
        assert_eq!(
            decompile(&rgce, (0, 0), &externs).unwrap(),
            "EOMONTH($I$7,0)"
        );
        let mut other = rgce;
        other[3] = 2;
        assert!(decompile(&other, (0, 0), &externs).is_err());
        // Unknown tokens (PtgArray) and dangling operands reject.
        assert!(decompile(&[0x20], (0, 0), &externs).is_err());
        assert!(decompile(&[0x1e, 1, 0, 0x1e, 2, 0], (0, 0), &externs).is_err());
    }
}
