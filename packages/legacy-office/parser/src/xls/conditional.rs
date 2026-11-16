//! Direct projection of BIFF8 conditional formatting onto the XLSX model.
//!
//! Records (MS-XLS 2.4.56 CondFmt, 2.4.42 CF, 2.4.57 CondFmt12, 2.4.43 CF12,
//! 2.4.44 CFEx) follow the CONDFMTS rule of the worksheet substream:
//! `*(CONDFMT / CONDFMT12) *(CFEx [CF12])`. A CondFmt's CF rules take their
//! priority, stop-if-true flag, template and Excel 2007 format from the
//! CFEx (CFExNonCF12 2.5.27) that names them by nID and icf; a CFEx with
//! fIsCF12 adds the CF12 that follows it to that CondFmt. Visual CF12 rules
//! use 2.5.32 CFGradient (color scale), 2.5.22 CFDatabar (data bar) and
//! 2.5.36 CFMultistate (icon set), with 2.5.39 CFVO values and 2.5.21
//! CFColor colors; formats are DXFN/DXFN12 (see `dxf`) and formulas are
//! decompiled by `formula`.
//!
//! The XLSX model rules are the ones the XLSX parser produces for the same
//! ECMA-376 18.3.1.x elements; the private corpus pairs each .xls with the
//! .xlsx Excel saved from it, which shows the expected SpreadsheetML.
//!
//! Rules made from a template Excel 2007 added keep the Boolean formula
//! Excel stores for earlier versions (see `condition`); average rules are
//! ECMA-376 aboveAverage rules. A CFVO formula is the `val` text of its
//! `cfvo`. Display options the model's rules cannot state are drawn as the
//! model draws the rule: a data bar or icon set that hides the cell value
//! shows it, and a data bar runs from the left over the whole cell width.
//!
//! What this module does not project yet rejects the workbook: CF rules
//! without a CFEx (no stated priority), inactive rules, filter CF12 rules,
//! stop-if-true on rule types the model cannot mark, activity formulas and
//! strict icon thresholds.

mod dxf;
pub(super) mod formula;
mod ftab;

pub(super) use formula::{decompile, decompile_cell, decompile_name, without_text, Externs};

use super::{f64_at, styles, theme, u16_at, u32_at, unsupported};
use std::collections::BTreeMap;

/// Retained raw records of one worksheet substream, in stream order.
#[derive(Default)]
pub(super) struct Records {
    records: Vec<(u16, Vec<u8>)>,
    bytes: usize,
}

/// Resource policy for retained CF record bytes per workbook sheet.
const MAX_SHEET_BYTES: usize = 16 * 1024 * 1024;

/// Resource policy for differential formats per workbook.
pub(super) const MAX_DXFS: usize = 65_536;

impl Records {
    pub(super) fn push(&mut self, kind: u16, data: &[u8]) -> Result<(), String> {
        self.bytes = self
            .bytes
            .checked_add(data.len())
            .filter(|bytes| *bytes <= MAX_SHEET_BYTES)
            .ok_or_else(|| {
                crate::rejection::resource("XLS conditional formatting byte budget exceeded")
            })?;
        self.records.push((kind, data.to_vec()));
        Ok(())
    }

    pub(super) fn is_empty(&self) -> bool {
        self.records.is_empty()
    }
}

pub(super) struct Context<'a> {
    pub(super) styles: &'a styles::Styles<'a>,
    pub(super) theme: &'a theme::Colors,
    pub(super) externs: &'a Externs,
}

/// A CondFmt whose CF rules wait for their CFEx.
struct Classic<'a> {
    format: usize,
    anchor: (u16, u16),
    rules: Vec<Option<ClassicRule<'a>>>,
}

/// A CF record (2.4.42) before its CFEx supplies priority and template.
struct ClassicRule<'a> {
    ct: u8,
    cp: u8,
    rgce1: &'a [u8],
    rgce2: &'a [u8],
    dxf: xlsx_model::Dxf,
}

/// Project every conditional format of one worksheet; differential formats
/// are appended to the workbook-level `dxfs` the rules index.
pub(super) fn project(
    records: &Records,
    context: &Context<'_>,
    dxfs: &mut Vec<xlsx_model::Dxf>,
) -> Result<Vec<xlsx_model::ConditionalFormat>, String> {
    let mut output: Vec<xlsx_model::ConditionalFormat> = Vec::new();
    let mut classics: BTreeMap<u16, Classic<'_>> = BTreeMap::new();
    // Rules still expected by the latest CondFmt (by nID) or CondFmt12.
    let mut classic_left: Option<(u16, usize)> = None;
    let mut cf12_left: Option<(usize, (u16, u16), usize)> = None;
    // A CFEx with fIsCF12 waiting for its CF12: (format, anchor).
    let mut extension: Option<(usize, (u16, u16))> = None;
    let incomplete = |classic: &Option<(u16, usize)>, cf12: &Option<(usize, (u16, u16), usize)>| {
        classic.is_some_and(|(_, left)| left != 0) || cf12.is_some_and(|(_, _, left)| left != 0)
    };
    for (kind, data) in &records.records {
        if extension.is_some() && *kind != 0x087a {
            return Err(unsupported("XLS CFEx lacks its CF12 record"));
        }
        match *kind {
            0x01b0 | 0x0879 => {
                if incomplete(&classic_left, &cf12_left) {
                    return Err(unsupported("XLS conditional format lacks its rules"));
                }
                let (format, ccf, id, anchor) = if *kind == 0x01b0 {
                    cond_fmt(data)?
                } else {
                    cond_fmt12(data)?
                };
                let index = output.len();
                output.push(format);
                if *kind == 0x01b0 {
                    if classics
                        .insert(
                            id,
                            Classic {
                                format: index,
                                anchor,
                                rules: Vec::with_capacity(ccf),
                            },
                        )
                        .is_some()
                    {
                        return Err(unsupported("duplicate XLS CondFmt identifier"));
                    }
                    classic_left = Some((id, ccf));
                    cf12_left = None;
                } else {
                    cf12_left = Some((index, anchor, ccf));
                    classic_left = None;
                }
            }
            0x01b1 => {
                let (id, left) = classic_left
                    .as_mut()
                    .filter(|(_, left)| *left != 0)
                    .ok_or_else(|| unsupported("XLS CF outside its CondFmt"))?;
                let rule = cf(data, context)?;
                classics
                    .get_mut(id)
                    .expect("registered CondFmt")
                    .rules
                    .push(Some(rule));
                *left -= 1;
            }
            0x087a => {
                let (format, anchor) = if let Some(target) = extension.take() {
                    target
                } else {
                    let (format, anchor, left) = cf12_left
                        .as_mut()
                        .filter(|(_, _, left)| *left != 0)
                        .ok_or_else(|| unsupported("XLS CF12 outside its CondFmt12"))?;
                    *left -= 1;
                    (*format, *anchor)
                };
                let rule = cf12(data, context, anchor, dxfs)?;
                output[format].rules.push(rule);
            }
            0x087b => {
                if incomplete(&classic_left, &cf12_left) {
                    return Err(unsupported("XLS conditional format lacks its rules"));
                }
                if u16_at(data, 0)? != 0x087b {
                    return Err(unsupported("invalid XLS CFEx header"));
                }
                let is_cf12 = u32_at(data, 12)?;
                let id = u16_at(data, 16)?;
                let classic = classics
                    .get_mut(&id)
                    .ok_or_else(|| unsupported("XLS CFEx names no CondFmt"))?;
                match is_cf12 {
                    0 => {
                        let rule = cfex(data, classic, context, dxfs)?;
                        output[classic.format].rules.push(rule);
                    }
                    1 => {
                        if data.len() != 18 {
                            return Err(unsupported("unexpected XLS CFEx tail"));
                        }
                        extension = Some((classic.format, classic.anchor));
                    }
                    _ => return Err(unsupported("invalid XLS CFEx rule kind")),
                }
            }
            _ => return Err(unsupported("unexpected XLS conditional formatting record")),
        }
    }
    if extension.is_some() {
        return Err(unsupported("XLS CFEx lacks its CF12 record"));
    }
    if incomplete(&classic_left, &cf12_left) {
        return Err(unsupported("XLS conditional format lacks its rules"));
    }
    // MS-XLS gives a CF rule's priority only through its CFEx; Excel 2007
    // and later write one for every CF (all CF rules in the private corpus
    // have one). Without it the evaluation order is not stated.
    if classics
        .values()
        .any(|classic| classic.rules.iter().any(Option::is_some))
    {
        return Err(unsupported(
            "XLS CF rule without a CFEx priority is not projected",
        ));
    }
    Ok(output)
}

/// Top-left cell (zero-based row, column) of a Ref8U bound.
fn bound_anchor(data: &[u8], offset: usize) -> Result<(u16, u16), String> {
    let row = u16_at(data, offset)?;
    let col = u16_at(data, offset + 4)?;
    if row > u16_at(data, offset + 2)? || col > u16_at(data, offset + 6)? || col > 0x00ff {
        return Err(unsupported("invalid XLS conditional formatting bound"));
    }
    Ok((row, col))
}

type Group = (xlsx_model::ConditionalFormat, usize, u16, (u16, u16));

/// MS-XLS 2.4.56 CondFmt: ccf, fToughRecalc/nID, refBound, sqref.
fn cond_fmt(data: &[u8]) -> Result<Group, String> {
    let ccf = usize::from(u16_at(data, 0)?);
    if !(1..=3).contains(&ccf) {
        return Err(unsupported("invalid XLS CondFmt rule count"));
    }
    let id = u16_at(data, 2)? >> 1;
    let anchor = bound_anchor(data, 4)?;
    let (sqref, end) = sqref(data, 12)?;
    if end != data.len() {
        return Err(unsupported("unexpected XLS CondFmt tail"));
    }
    Ok((
        xlsx_model::ConditionalFormat {
            sqref,
            rules: Vec::new(),
            first_match: false,
        },
        ccf,
        id,
        anchor,
    ))
}

/// MS-XLS 2.4.57: FrtRefHeaderU, then CondFmtStructure (2.5.56).
fn cond_fmt12(data: &[u8]) -> Result<Group, String> {
    if u16_at(data, 0)? != 0x0879 {
        return Err(unsupported("invalid XLS CondFmt12 header"));
    }
    let ccf = usize::from(u16_at(data, 12)?);
    if ccf == 0 {
        return Err(unsupported("empty XLS CondFmt12"));
    }
    let id = u16_at(data, 14)? >> 1;
    let anchor = bound_anchor(data, 16)?;
    let (sqref, end) = sqref(data, 24)?;
    if end != data.len() {
        return Err(unsupported("unexpected XLS CondFmt12 tail"));
    }
    Ok((
        xlsx_model::ConditionalFormat {
            sqref,
            rules: Vec::new(),
            first_match: false,
        },
        ccf,
        id,
        anchor,
    ))
}

/// MS-XLS 2.5.249 SqRefU of 2.5.209 Ref8U, as one-based XLSX ranges.
fn sqref(data: &[u8], offset: usize) -> Result<(Vec<xlsx_model::CellRange>, usize), String> {
    let count = usize::from(u16_at(data, offset)?);
    if count == 0 || count > 1026 {
        return Err(unsupported(
            "invalid XLS conditional formatting range count",
        ));
    }
    let mut ranges = Vec::with_capacity(count);
    for index in 0..count {
        let at = offset + 2 + index * 8;
        let (row_first, row_last) = (u16_at(data, at)?, u16_at(data, at + 2)?);
        let (col_first, col_last) = (u16_at(data, at + 4)?, u16_at(data, at + 6)?);
        if row_first > row_last || col_first > col_last || col_last > 0x00ff {
            return Err(unsupported("invalid XLS conditional formatting range"));
        }
        ranges.push(xlsx_model::CellRange {
            top: u32::from(row_first) + 1,
            left: u32::from(col_first) + 1,
            bottom: u32::from(row_last) + 1,
            right: u32::from(col_last) + 1,
        });
    }
    Ok((ranges, offset + 2 + count * 8))
}

/// MS-XLS 2.4.42 CF: ct, cp, cce1, cce2, DXFN, rgce1, rgce2.
fn cf<'a>(data: &'a [u8], context: &Context<'_>) -> Result<ClassicRule<'a>, String> {
    let ct = *data.first().ok_or_else(truncated)?;
    let cp = *data.get(1).ok_or_else(truncated)?;
    let cce1 = usize::from(u16_at(data, 2)?);
    let cce2 = usize::from(u16_at(data, 4)?);
    let (dxf, offset) = dxf::classic(data, 6, context)?;
    let rgce1 = slice(data, offset, cce1)?;
    let rgce2 = slice(data, offset + cce1, cce2)?;
    if offset + cce1 + cce2 != data.len() {
        return Err(unsupported("unexpected XLS CF tail"));
    }
    Ok(ClassicRule {
        ct,
        cp,
        rgce1,
        rgce2,
        dxf,
    })
}

/// ECMA-376 18.18.15 ST_ConditionalFormattingOperator for MS-XLS cp.
fn operator(cp: u8) -> Result<(&'static str, usize), String> {
    Ok(match cp {
        1 => ("between", 2),
        2 => ("notBetween", 2),
        3 => ("equal", 1),
        4 => ("notEqual", 1),
        5 => ("greaterThan", 1),
        6 => ("lessThan", 1),
        7 => ("greaterThanOrEqual", 1),
        8 => ("lessThanOrEqual", 1),
        _ => return Err(unsupported("invalid XLS conditional formatting comparison")),
    })
}

/// A cell-value or formula condition (ct 1 or 2 of CF and CF12) with its
/// template, as an XLSX-model rule.
#[allow(clippy::too_many_arguments)]
fn condition(
    ct: u8,
    cp: u8,
    rgce1: &[u8],
    rgce2: &[u8],
    template: u16,
    parameters: &[u8],
    anchor: (u16, u16),
    priority: i32,
    stop_if_true: bool,
    dxf_id: u32,
    context: &Context<'_>,
) -> Result<xlsx_model::CfRule, String> {
    let dxf_id = Some(dxf_id);
    let formula = |rgce: &[u8]| formula::decompile(rgce, anchor, context.externs);
    const TEMPLATE: &str = "unsupported XLS conditional formatting rule template";
    match (ct, template) {
        // Cell value: ECMA-376 cellIs. icfTemplate is 0 for it; any other
        // value does not change what the comparison states.
        (1, _) => {
            if template != 0 {
                crate::rejection::ignored_value(TEMPLATE);
            }
            let (operator, count) = operator(cp)?;
            if (count == 2) == rgce2.is_empty() || rgce1.is_empty() {
                return Err(unsupported("invalid XLS cell-value condition"));
            }
            let mut formulas = vec![formula(rgce1)?];
            if count == 2 {
                formulas.push(formula(rgce2)?);
            }
            Ok(xlsx_model::CfRule::CellIs {
                operator: operator.to_string(),
                formulas,
                dxf_id,
                priority,
                stop_if_true,
            })
        }
        // Formula (1), contains text (8), blanks (9), no blanks (10),
        // errors (11), no errors (12): the rule is its Boolean formula,
        // which Excel also writes as the `formula` of the SpreadsheetML
        // expression / containsText / *Blanks / *Errors rules; the XLSX
        // parser reads all of them as expression rules.
        (2, 1 | 8..=12) => expression(cp, rgce1, rgce2, &formula, dxf_id, priority, stop_if_true),
        // Filter (5) with CFExFilterParams (2.5.26): ECMA-376 top10. The
        // CF formula is Excel's compatibility rendering for older versions
        // (observed: a bottom-N rule stored with a SMALL formula is saved by
        // Excel as `top10 bottom="1" rank="N"`).
        (2, 5) => {
            let flags = *parameters.first().ok_or_else(truncated)?;
            let rank = u16_at(parameters, 1)?;
            let percent = flags & 0x02 != 0;
            if flags & !0x03 != 0 || rank == 0 || rank > if percent { 100 } else { 1000 } {
                return Err(unsupported("invalid XLS top/bottom filter"));
            }
            Ok(xlsx_model::CfRule::Top10 {
                top: flags & 0x01 != 0,
                percent,
                rank: rank.into(),
                dxf_id,
                priority,
                stop_if_true,
            })
        }
        // Above or below average (0x19, 0x1A), or equal to it as well (0x1D,
        // 0x1E), with CFExAveragesTemplateParams (2.5.23): ECMA-376
        // aboveAverage. As for the filter template the CF formula is Excel's
        // rendering for earlier versions.
        (2, 0x19 | 0x1a | 0x1d | 0x1e) => {
            if cp != 0 || rgce1.is_empty() || !rgce2.is_empty() {
                return Err(unsupported("invalid XLS formula condition"));
            }
            // The compatibility formula is not projected, but it is part
            // of the rule: one that cannot be read rejects. One the
            // decompiler merely has no text for does not matter here.
            if let Err(error) = formula(rgce1) {
                if !formula::without_text(&error) {
                    return Err(error);
                }
            }
            // iParam: the number of standard deviations. 2.5.23 lists 0, 1
            // and 2; Excel's rule editor also writes 3 (ECMA-376 stdDev).
            let std_dev = match u16_at(parameters, 0)? {
                0 => None,
                count @ 1..=3 => Some(u32::from(count)),
                _ => crate::rejection::spec_default(TEMPLATE, None),
            };
            Ok(xlsx_model::CfRule::AboveAverage {
                above_average: matches!(template, 0x19 | 0x1d),
                equal_average: matches!(template, 0x1d | 0x1e),
                std_dev,
                dxf_id,
                priority,
                stop_if_true,
            })
        }
        // Unique (7) and duplicate (0x1B) values and the date periods
        // (0x0F..=0x18): SpreadsheetML has uniqueValues, duplicateValues and
        // timePeriod rules, which the model keeps without evaluating them.
        // The CF formula is the Boolean formula Excel stores for versions
        // without the template (COUNTIF over the range, comparisons with
        // TODAY()), so the rule is projected as that expression.
        (2, 7 | 0x0f..=0x18 | 0x1b) => crate::rejection::approximated(
            TEMPLATE,
            expression(cp, rgce1, rgce2, &formula, dxf_id, priority, stop_if_true),
        ),
        // icfTemplate outside the 2.4.43 table, or of another rule type:
        // the rule is the formula condition its ct states.
        (2, _) => {
            crate::rejection::ignored_value(TEMPLATE);
            expression(cp, rgce1, rgce2, &formula, dxf_id, priority, stop_if_true)
        }
        // CF 2.4.42 ct is 1 or 2; `cf12` projects its other types.
        _ => Err(unsupported("invalid XLS conditional formatting rule type")),
    }
}

/// A formula condition (ct 2): ECMA-376 expression.
fn expression(
    cp: u8,
    rgce1: &[u8],
    rgce2: &[u8],
    formula: &dyn Fn(&[u8]) -> Result<String, String>,
    dxf_id: Option<u32>,
    priority: i32,
    stop_if_true: bool,
) -> Result<xlsx_model::CfRule, String> {
    if cp != 0 || rgce1.is_empty() || !rgce2.is_empty() {
        return Err(unsupported("invalid XLS formula condition"));
    }
    Ok(xlsx_model::CfRule::Expression {
        formula: formula(rgce1)?,
        dxf_id,
        priority,
        stop_if_true,
    })
}

fn push_dxf(dxfs: &mut Vec<xlsx_model::Dxf>, dxf: xlsx_model::Dxf) -> Result<u32, String> {
    if dxfs.len() >= MAX_DXFS {
        return Err(crate::rejection::resource(
            "too many XLS conditional formatting formats",
        ));
    }
    dxfs.push(dxf);
    u32::try_from(dxfs.len() - 1).map_err(|_| truncated())
}

/// MS-XLS 2.5.27 CFExNonCF12 at offset 18 of a CFEx record.
fn cfex(
    data: &[u8],
    classic: &mut Classic<'_>,
    context: &Context<'_>,
    dxfs: &mut Vec<xlsx_model::Dxf>,
) -> Result<xlsx_model::CfRule, String> {
    let icf = usize::from(u16_at(data, 18)?);
    let cp = *data.get(20).ok_or_else(truncated)?;
    let template = u16::from(*data.get(21).ok_or_else(truncated)?);
    let priority = i32::from(u16_at(data, 22)?);
    let flags = *data.get(24).ok_or_else(truncated)?;
    let has_dxf = *data.get(25).ok_or_else(truncated)?;
    let (dxf, offset) = match has_dxf {
        0 => (None, 26),
        1 => {
            let (dxf, end) = dxf::dxfn12(data, 26, context)?;
            (
                Some(dxf.ok_or_else(|| unsupported("XLS CFEx format is empty"))?),
                end,
            )
        }
        _ => return Err(unsupported("invalid XLS CFEx format flag")),
    };
    if *data.get(offset).ok_or_else(truncated)? != 16 {
        return Err(unsupported("invalid XLS CFEx template parameter size"));
    }
    let parameters = slice(data, offset + 1, 16)?;
    if offset + 17 != data.len() {
        return Err(unsupported("unexpected XLS CFEx tail"));
    }
    // fActive = 0 disables the rule; SpreadsheetML has no inactive rule.
    if flags & 0x01 == 0 {
        return Err(unsupported("inactive XLS conditional formatting rule"));
    }
    let rule = classic
        .rules
        .get_mut(icf)
        .and_then(Option::take)
        .ok_or_else(|| unsupported("XLS CFEx names no pending CF rule"))?;
    // cp: the CFEx value takes precedence when the two differ (MS-XLS
    // 2.5.27 note 154); it MUST be zero for formula rules.
    let cp = if rule.ct == 1 { cp } else { rule.cp.max(cp) };
    let dxf_id = push_dxf(dxfs, dxf.unwrap_or(rule.dxf))?;
    condition(
        rule.ct,
        cp,
        rule.rgce1,
        rule.rgce2,
        template,
        parameters,
        classic.anchor,
        priority,
        flags & 0x02 != 0,
        dxf_id,
        context,
    )
}

/// MS-XLS 2.4.43 CF12.
fn cf12(
    data: &[u8],
    context: &Context<'_>,
    anchor: (u16, u16),
    dxfs: &mut Vec<xlsx_model::Dxf>,
) -> Result<xlsx_model::CfRule, String> {
    if u16_at(data, 0)? != 0x087a {
        return Err(unsupported("invalid XLS CF12 header"));
    }
    let byte = |offset: usize| -> Result<u8, String> {
        data.get(offset)
            .copied()
            .ok_or_else(|| unsupported("truncated XLS CF12"))
    };
    let ct = byte(12)?;
    let cp = byte(13)?;
    let (cce1, cce2) = (
        usize::from(u16_at(data, 14)?),
        usize::from(u16_at(data, 16)?),
    );
    let (dxf, mut offset) = dxf::dxfn12(data, 18, context)?;
    let rgce1 = slice(data, offset, cce1)?;
    offset += cce1;
    let rgce2 = slice(data, offset, cce2)?;
    offset += cce2;
    let active = usize::from(u16_at(data, offset)?);
    offset += 2 + active;
    let flags = byte(offset)?;
    let priority = i32::from(u16_at(data, offset + 1)?);
    let template = u16_at(data, offset + 3)?;
    if byte(offset + 5)? != 16 {
        return Err(unsupported("invalid XLS CF12 template parameter size"));
    }
    let parameters = slice(data, offset + 6, 16)?;
    let body = data.get(offset + 22..).ok_or_else(truncated)?;
    if active != 0 {
        return Err(unsupported("XLS CF12 activity formula is not supported"));
    }
    if matches!(ct, 1 | 2) {
        // Cell-value and formula rules carry no rgbCT.
        if !body.is_empty() {
            return Err(unsupported("unexpected XLS CF12 tail"));
        }
        let dxf_id = push_dxf(dxfs, dxf.unwrap_or_default())?;
        return condition(
            ct,
            cp,
            rgce1,
            rgce2,
            template,
            parameters,
            anchor,
            priority,
            flags & 0x02 != 0,
            dxf_id,
            context,
        );
    }
    if !matches!(ct, 3 | 4 | 6) {
        return Err(unsupported("XLS CF12 filter rules are not projected yet"));
    }
    // CF12 2.4.43: color scales, data bars and icon sets carry no dxf and no
    // comparison formulas, and fStopIfTrue "MUST be zero when ct is equal to
    // 0x03, 0x04 or 0x06", so their model stop_if_true is always false. (A
    // set flag is invalid binary storage and is rejected rather than honoured
    // as SpreadsheetML would.) A non-empty fmlaActive was rejected above; it is
    // not projected onto the model's active_formula yet, so every projected
    // visual rule is unconditionally active (active_formula None).
    if dxf.is_some() || !rgce1.is_empty() || !rgce2.is_empty() || flags & 0x02 != 0 {
        return Err(unsupported("invalid XLS CF12 visual rule"));
    }
    let expected_template = match ct {
        3 => 2,
        4 => 3,
        _ => 4,
    };
    if template != expected_template {
        return Err(unsupported(
            "XLS CF12 template disagrees with its rule type",
        ));
    }
    let (rule, end) = match ct {
        3 => gradient(body, priority, anchor, context)?,
        4 => databar(body, priority, anchor, context)?,
        _ => multistate(body, priority, anchor, context)?,
    };
    if end != body.len() {
        return Err(unsupported("unexpected XLS CF12 tail"));
    }
    Ok(rule)
}

fn truncated() -> String {
    unsupported("truncated XLS conditional formatting")
}

fn slice(data: &[u8], offset: usize, length: usize) -> Result<&[u8], String> {
    data.get(offset..offset.checked_add(length).ok_or_else(truncated)?)
        .ok_or_else(truncated)
}

/// MS-XLS 2.5.39 CFVO onto the XLSX `cfvo` types the XLSX parser keeps.
/// A CFVOParsedFormula (2.5.198.7: cce, then its rgce) replaces numValue: its
/// text is the `val` of the SpreadsheetML `cfvo`, which holds a formula for
/// every type (ECMA-376 18.3.1.11), and cfvoType 7 is `type="formula"`.
fn cfvo(
    data: &[u8],
    offset: usize,
    anchor: (u16, u16),
    context: &Context<'_>,
) -> Result<(xlsx_model::CfValue, usize), String> {
    let kind = *data.get(offset).ok_or_else(truncated)?;
    let cce = usize::from(u16_at(data, offset + 1)?);
    let mut end = offset + 3;
    let formula = if cce != 0 {
        let rgce = slice(data, end, cce)?;
        end += cce;
        Some(rgce)
    } else {
        None
    };
    let (name, value) = match (kind, formula) {
        // The minimum and maximum take no parameter.
        (2, _) => ("min", None),
        (3, _) => ("max", None),
        (1 | 4 | 5, None) => {
            let value = f64_at(data, end)?;
            end += 8;
            if !value.is_finite() || (kind != 1 && !(0.0..=100.0).contains(&value)) {
                return Err(unsupported("invalid XLS conditional formatting value"));
            }
            (cfvo_type(kind), Some(number(value)))
        }
        (1 | 4 | 5 | 7, Some(rgce)) => (
            cfvo_type(kind),
            Some(formula::decompile(rgce, anchor, context.externs)?),
        ),
        // "If cfvoType is 0x07, then fmla.cce MUST be greater than zero."
        (7, None) => {
            return Err(crate::rejection::corrupt(
                "XLS conditional formatting value lacks its formula",
            ))
        }
        _ => {
            return Err(unsupported(
                "unsupported XLS conditional formatting value type",
            ))
        }
    };
    Ok((
        xlsx_model::CfValue {
            kind: name.to_string(),
            value,
        },
        end,
    ))
}

/// ECMA-376 18.18.12 ST_CfvoType of a cfvoType with a parameter.
fn cfvo_type(kind: u8) -> &'static str {
    match kind {
        1 => "num",
        4 => "percent",
        5 => "percentile",
        _ => "formula",
    }
}

/// The shortest round-trip decimal, as a SpreadsheetML `cfvo@val` string.
fn number(value: f64) -> String {
    format!("{value}")
}

/// MS-XLS 2.5.21 CFColor as the XLSX model's resolved `#RRGGBB`.
fn color(data: &[u8], offset: usize, context: &Context<'_>) -> Result<String, String> {
    cf_color(data, offset, context.styles, context.theme)
}

/// MS-XLS 2.5.21 CFColor as a resolved `#RRGGBB` (also used by SheetExt).
pub(super) fn cf_color(
    data: &[u8],
    offset: usize,
    styles: &styles::Styles<'_>,
    theme: &theme::Colors,
) -> Result<String, String> {
    use ooxml_common::spreadsheet_color::{resolve_color, SpreadsheetColor};
    let kind = u32_at(data, offset)?;
    let value = u32_at(data, offset + 4)?;
    let tint = f64_at(data, offset + 8)?;
    if !(-1.0..=1.0).contains(&tint) {
        return Err(unsupported("invalid XLS conditional formatting color tint"));
    }
    let tint = (tint != 0.0).then_some(tint);
    let resolved = match kind {
        // XCLRINDEXED ColorICV: the workbook Palette or built-in palette.
        1 => {
            let index = u16::try_from(value).map_err(|_| truncated())?;
            let base = styles
                .chart_color(index)
                .ok_or_else(|| unsupported("unsupported XLS conditional formatting color index"))?;
            let hex = base.trim_start_matches('#');
            let channel = |at: usize| u8::from_str_radix(&hex[at..at + 2], 16).unwrap_or(0);
            resolve_color(
                SpreadsheetColor::Argb([0xff, channel(0), channel(2), channel(4)]),
                tint,
                &[],
            )
        }
        // XCLRRGB LongRGBA: red, green, blue, alpha.
        2 => {
            let [r, g, b, a] = value.to_le_bytes();
            resolve_color(SpreadsheetColor::Argb([a, r, g, b]), tint, &[])
        }
        // XCLRTHEMED ColorTheme. As for XFExt colors, Excel's theme index
        // follows SpreadsheetML (0 = lt1, 1 = dk1, 2 = lt2, 3 = dk2) while the
        // parsed scheme is in clrScheme order; swap each light/dark pair.
        3 if value <= 11 => {
            let slot = if value < 4 { value ^ 1 } else { value };
            let argb = theme.argb(slot).ok_or_else(|| {
                unsupported("XLS themed conditional formatting color lacks a theme")
            })?;
            resolve_color(SpreadsheetColor::Argb(argb), tint, &[])
        }
        _ => None,
    };
    resolved.ok_or_else(|| unsupported("unsupported XLS conditional formatting color"))
}

/// MS-XLS 2.5.32 CFGradient onto ECMA-376 18.3.1.16 colorScale.
fn gradient(
    data: &[u8],
    priority: i32,
    anchor: (u16, u16),
    context: &Context<'_>,
) -> Result<(xlsx_model::CfRule, usize), String> {
    let count = usize::from(*data.get(3).ok_or_else(truncated)?);
    let curve = usize::from(*data.get(4).ok_or_else(truncated)?);
    let flags = *data.get(5).ok_or_else(truncated)?;
    // fClamp (SHOULD be 1) and fBackground (MUST be 1): the shared renderer
    // clamps and fills the background.
    if !matches!(count, 2 | 3) || curve != count || flags & 0x03 != 0x03 {
        return Err(unsupported("unsupported XLS color scale"));
    }
    let domains: &[f64] = if count == 2 {
        &[0.0, 1.0]
    } else {
        &[0.0, 0.5, 1.0]
    };
    let mut offset = 6;
    let mut values = Vec::with_capacity(count);
    for domain in domains {
        let (value, end) = cfvo(data, offset, anchor, context)?;
        if f64_at(data, end)? != *domain {
            return Err(unsupported("invalid XLS color scale control point"));
        }
        offset = end + 8;
        values.push(value);
    }
    let mut stops = Vec::with_capacity(count);
    for (value, domain) in values.into_iter().zip(domains) {
        if f64_at(data, offset)? != *domain {
            return Err(unsupported("invalid XLS color scale gradient point"));
        }
        stops.push(xlsx_model::CfStop {
            kind: value.kind,
            value: value.value,
            color: color(data, offset + 8, context)?,
        });
        offset += 24;
    }
    Ok((
        xlsx_model::CfRule::ColorScale {
            stops,
            priority,
            active_formula: None,
            stop_if_true: false,
        },
        offset,
    ))
}

/// MS-XLS 2.5.22 CFDatabar onto ECMA-376 18.3.1.28 dataBar.
fn databar(
    data: &[u8],
    priority: i32,
    anchor: (u16, u16),
    context: &Context<'_>,
) -> Result<(xlsx_model::CfRule, usize), String> {
    let flags = *data.get(3).ok_or_else(truncated)?;
    let (minimum, maximum) = (
        *data.get(4).ok_or_else(truncated)?,
        *data.get(5).ok_or_else(truncated)?,
    );
    // The model's data bar is drawn from the left across the full cell
    // width with the value shown. Excel's PDF of a workbook whose CFDatabar
    // stores flags 0 and lengths 0..100 shows left-to-right gradient bars
    // with the cell values, and Excel stores flags 2 for a bar whose value
    // is hidden ("show bar only"; its PDF prints the bars alone). The model
    // states neither that option, the direction (bit 0) nor other bar
    // lengths, so such a rule is drawn as the model's bar.
    if flags & 0x03 != 0 || minimum != 0 || maximum != 100 {
        crate::rejection::approximated("unsupported XLS data bar display", ());
    }
    let color = color(data, 6, context)?;
    let (min, end) = cfvo(data, 22, anchor, context)?;
    let (max, end) = cfvo(data, end, anchor, context)?;
    Ok((
        xlsx_model::CfRule::DataBar {
            color,
            min,
            max,
            priority,
            // CF12 data bars have no fill-type field; Excel draws them with
            // its 2007 gradient (observed in the same PDF).
            gradient: true,
            active_formula: None,
            // BIFF has no linked x14 extension facts.
            ext_threshold_formula: false,
            stop_if_true: false,
        },
        end,
    ))
}

/// ECMA-376 18.18.42 ST_IconSetType in the order MS-XLS 2.5.36 counts
/// iIconSet (three-icon sets 0..7, four-icon 8..12, five-icon 13..16).
const ICON_SETS: [&str; 17] = [
    "3Arrows",
    "3ArrowsGray",
    "3Flags",
    "3TrafficLights1",
    "3TrafficLights2",
    "3Signs",
    "3Symbols",
    "3Symbols2",
    "4Arrows",
    "4ArrowsGray",
    "4RedToBlack",
    "4Rating",
    "4TrafficLights",
    "5Arrows",
    "5ArrowsGray",
    "5Rating",
    "5Quarters",
];

/// MS-XLS 2.5.36 CFMultistate onto ECMA-376 18.3.1.49 iconSet.
fn multistate(
    data: &[u8],
    priority: i32,
    anchor: (u16, u16),
    context: &Context<'_>,
) -> Result<(xlsx_model::CfRule, usize), String> {
    let states = usize::from(*data.get(3).ok_or_else(truncated)?);
    let set = usize::from(*data.get(4).ok_or_else(truncated)?);
    let flags = *data.get(5).ok_or_else(truncated)?;
    let expected = match set {
        0..=7 => 3,
        8..=12 => 4,
        13..=16 => 5,
        _ => return Err(unsupported("invalid XLS icon set")),
    };
    if states != expected {
        return Err(unsupported("unsupported XLS icon set display"));
    }
    // fIconOnly (bit 0) hides the cell value, which the model's icon set
    // cannot state: the icon is drawn beside the value. The other bits but
    // fReverse (bit 2) are reserved and "MUST be ignored".
    if flags & 0x01 != 0 {
        crate::rejection::approximated("unsupported XLS icon set display", ());
    }
    let mut offset = 6;
    let mut cfvos = Vec::with_capacity(states);
    for _ in 0..states {
        let (value, end) = cfvo(data, offset, anchor, context)?;
        // fEqual = 1 is ECMA-376 cfvo@gte (default true), the only threshold
        // comparison the shared renderer implements.
        if *data.get(end).ok_or_else(truncated)? != 1 {
            return Err(unsupported("XLS icon threshold excluding equal values"));
        }
        slice(data, end + 1, 4)?;
        offset = end + 5;
        cfvos.push(value);
    }
    Ok((
        xlsx_model::CfRule::IconSet {
            icon_set: ICON_SETS[set].to_string(),
            cfvos,
            reverse: flags & 0x04 != 0,
            priority,
            custom_icons: None,
            active_formula: None,
            stop_if_true: false,
        },
        offset,
    ))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn cond_fmt12(ccf: u16) -> Vec<u8> {
        let mut data = vec![0u8; 12];
        data[..2].copy_from_slice(&0x0879u16.to_le_bytes());
        data.extend(ccf.to_le_bytes());
        data.extend(1u16.to_le_bytes());
        data.extend([0u8; 8]);
        data.extend(1u16.to_le_bytes());
        // E11:E21 (zero-based rows 10..20, column 4).
        for value in [10u16, 20, 4, 4] {
            data.extend(value.to_le_bytes());
        }
        data
    }

    fn cf12(ct: u8, template: u16, flags: u8, active: &[u8], body: &[u8]) -> Vec<u8> {
        let mut data = vec![0u8; 12];
        data[..2].copy_from_slice(&0x087au16.to_le_bytes());
        data.extend([ct, 0, 0, 0, 0, 0]);
        data.extend(0u32.to_le_bytes());
        data.extend([0, 0]);
        data.extend((active.len() as u16).to_le_bytes());
        data.extend(active);
        data.push(flags);
        data.extend(7u16.to_le_bytes());
        data.extend(template.to_le_bytes());
        data.push(16);
        data.extend([0u8; 16]);
        data.extend(body);
        data
    }

    fn cfvo(kind: u8, value: Option<f64>) -> Vec<u8> {
        let mut data = vec![kind, 0, 0];
        if let Some(value) = value {
            data.extend(value.to_le_bytes());
        }
        data
    }

    fn rgb(r: u8, g: u8, b: u8) -> Vec<u8> {
        let mut data = 2u32.to_le_bytes().to_vec();
        data.extend([r, g, b, 0xff]);
        data.extend(0f64.to_le_bytes());
        data
    }

    fn project_records(
        records: &[(u16, Vec<u8>)],
    ) -> Result<Vec<xlsx_model::ConditionalFormat>, String> {
        let styles = styles::Styles::parse(&[]).unwrap();
        let theme = theme::Colors::default();
        let mut retained = Records::default();
        for (kind, data) in records {
            retained.push(*kind, data).unwrap();
        }
        project(
            &retained,
            &Context {
                styles: &styles,
                theme: &theme,
                externs: &Externs::default(),
            },
            &mut Vec::new(),
        )
    }

    fn icon_set(flags: u8, equal: u8) -> Vec<u8> {
        let mut body = vec![0, 0, 0, 3, 5, flags];
        for (kind, value) in [(4, Some(0.0)), (1, Some(-20.0)), (1, Some(0.0))] {
            body.extend(cfvo(kind, value));
            body.push(equal);
            body.extend([0u8; 4]);
        }
        body
    }

    #[test]
    fn icon_sets_follow_the_iconset_enumeration_and_cfvo_types() {
        let formats = project_records(&[
            (0x0879, cond_fmt12(1)),
            (0x087a, cf12(6, 4, 0, &[], &icon_set(0x04, 1))),
        ])
        .unwrap();
        assert_eq!(formats.len(), 1);
        let range = &formats[0].sqref[0];
        assert_eq!(
            (range.top, range.bottom, range.left, range.right),
            (11, 21, 5, 5)
        );
        let xlsx_model::CfRule::IconSet {
            icon_set,
            cfvos,
            reverse,
            priority,
            ..
        } = &formats[0].rules[0]
        else {
            panic!("icon set");
        };
        assert_eq!(
            (icon_set.as_str(), *reverse, *priority),
            ("3Signs", true, 7)
        );
        let values: Vec<_> = cfvos
            .iter()
            .map(|value| (value.kind.as_str(), value.value.as_deref()))
            .collect();
        assert_eq!(
            values,
            [
                ("percent", Some("0")),
                ("num", Some("-20")),
                ("num", Some("0"))
            ]
        );
    }

    #[test]
    fn color_scales_and_data_bars_resolve_their_colors() {
        let mut scale = vec![0, 0, 0, 2, 2, 3];
        scale.extend(cfvo(2, None));
        scale.extend(0f64.to_le_bytes());
        scale.extend(cfvo(5, Some(90.0)));
        scale.extend(1f64.to_le_bytes());
        for (domain, color) in [(0.0, rgb(0x38, 0x4d, 0x60)), (1.0, rgb(0xc2, 0xe2, 0xff))] {
            scale.extend(f64::to_le_bytes(domain));
            scale.extend(color);
        }
        let mut bar = vec![0, 0, 0, 0, 0, 100];
        bar.extend(rgb(0x20, 0xa4, 0x72));
        bar.extend(cfvo(1, Some(0.0)));
        bar.extend(cfvo(1, Some(1.0)));
        let formats = project_records(&[
            (0x0879, cond_fmt12(2)),
            (0x087a, cf12(3, 2, 0, &[], &scale)),
            (0x087a, cf12(4, 3, 0, &[], &bar)),
        ])
        .unwrap();
        let rules = &formats[0].rules;
        let xlsx_model::CfRule::ColorScale { stops, .. } = &rules[0] else {
            panic!("color scale");
        };
        assert_eq!(stops[0].color, "#384D60");
        assert_eq!(
            (stops[1].kind.as_str(), stops[1].value.as_deref()),
            ("percentile", Some("90"))
        );
        let xlsx_model::CfRule::DataBar {
            color,
            gradient,
            max,
            ..
        } = &rules[1]
        else {
            panic!("data bar");
        };
        assert_eq!(
            (color.as_str(), *gradient, max.value.as_deref()),
            ("#20A472", true, Some("1"))
        );
    }

    /// MS-XLS 2.4.43 fStopIfTrue is honoured on comparison rules (ECMA-376
    /// 18.3.1.10 stopIfTrue on cellIs), while color scales, data bars and
    /// icon sets project it as false and are unconditionally active.
    #[test]
    fn stop_if_true_reaches_cell_value_rules_and_scales_never_stop() {
        let cell_value = |flags: u8| {
            let mut data = vec![0u8; 12];
            data[..2].copy_from_slice(&0x087au16.to_le_bytes());
            // ct 1 (cell value), cp 5 (greaterThan), cce1 3, cce2 0.
            data.extend([1, 5, 3, 0, 0, 0]);
            data.extend(0u32.to_le_bytes());
            data.extend([0, 0]);
            // rgce1: PtgInt 5.
            data.extend([0x1e, 5, 0]);
            data.extend(0u16.to_le_bytes());
            data.push(flags);
            data.extend(7u16.to_le_bytes());
            data.extend(0u16.to_le_bytes());
            data.push(16);
            data.extend([0u8; 16]);
            data
        };
        for (flags, expected) in [(0x02, true), (0x00, false)] {
            let formats =
                project_records(&[(0x0879, cond_fmt12(1)), (0x087a, cell_value(flags))]).unwrap();
            let xlsx_model::CfRule::CellIs {
                operator,
                formulas,
                stop_if_true,
                ..
            } = &formats[0].rules[0]
            else {
                panic!("cell-value rule");
            };
            assert_eq!(
                (operator.as_str(), formulas.as_slice(), *stop_if_true),
                ("greaterThan", &["5".to_string()][..], expected)
            );
        }
        let formats = project_records(&[
            (0x0879, cond_fmt12(1)),
            (0x087a, cf12(6, 4, 0, &[], &icon_set(0, 1))),
        ])
        .unwrap();
        assert!(matches!(
            &formats[0].rules[0],
            xlsx_model::CfRule::IconSet {
                stop_if_true: false,
                active_formula: None,
                ..
            }
        ));
        // fStopIfTrue MUST be zero on a scale rule: invalid storage fails closed.
        assert!(project_records(&[
            (0x0879, cond_fmt12(1)),
            (0x087a, cf12(6, 4, 0x02, &[], &icon_set(0, 1))),
        ])
        .is_err());
    }

    /// CondFmt (2.4.56) `id` on E11:E21 with `ccf` CF rules.
    fn cond_fmt(id: u16, ccf: u16) -> Vec<u8> {
        let mut data = ccf.to_le_bytes().to_vec();
        data.extend((id << 1).to_le_bytes());
        for value in [10u16, 20, 4, 4, 1, 10, 20, 4, 4] {
            data.extend(value.to_le_bytes());
        }
        data
    }

    /// CF (2.4.42) with a DXFN that changes nothing.
    fn cf(ct: u8, cp: u8, rgce1: &[u8]) -> Vec<u8> {
        let mut data = vec![ct, cp];
        data.extend((rgce1.len() as u16).to_le_bytes());
        data.extend(0u16.to_le_bytes());
        data.extend(0x003f_ffffu32.to_le_bytes());
        data.extend(0u16.to_le_bytes());
        data.extend(rgce1);
        data
    }

    /// CFEx (2.4.44) with a CFExNonCF12 for rule 0 of CondFmt `id`: active,
    /// priority 3, with `dxf` (a DXFN12 body) when not empty.
    fn cf_ex(id: u16, cp: u8, template: u8, parameter: u16, dxf: &[u8]) -> Vec<u8> {
        let mut data = vec![0u8; 12];
        data[..2].copy_from_slice(&0x087bu16.to_le_bytes());
        data.extend(0u32.to_le_bytes());
        data.extend(id.to_le_bytes());
        data.extend(0u16.to_le_bytes());
        data.extend([cp, template]);
        data.extend(3u16.to_le_bytes());
        data.extend([0x01, u8::from(!dxf.is_empty())]);
        if !dxf.is_empty() {
            data.extend((dxf.len() as u32).to_le_bytes());
            data.extend(dxf);
        }
        data.push(16);
        data.extend(parameter.to_le_bytes());
        data.extend([0u8; 14]);
        data
    }

    fn classic_rule(
        ct: u8,
        cp: u8,
        rgce1: &[u8],
        template: u8,
        parameter: u16,
    ) -> Result<xlsx_model::CfRule, String> {
        let mut formats = project_records(&[
            (0x01b0, cond_fmt(1, 1)),
            (0x01b1, cf(ct, cp, rgce1)),
            (0x087b, cf_ex(1, cp, template, parameter, &[])),
        ])?;
        assert_eq!(formats.len(), 1);
        Ok(formats.remove(0).rules.remove(0))
    }

    use crate::rejection::{take_best_effort, BestEffort};
    const TEMPLATE: &str = "unsupported XLS conditional formatting rule template";

    #[test]
    fn average_templates_are_above_average_rules() {
        take_best_effort();
        // PtgBool TRUE stands for Excel's compatibility formula.
        let formula = [0x1d, 1];
        for (template, parameter, expected) in [
            (0x19, 0, (true, false, None)),
            (0x1a, 0, (false, false, None)),
            (0x1d, 0, (true, true, None)),
            (0x1e, 2, (false, true, Some(2))),
            (0x19, 3, (true, false, Some(3))),
        ] {
            let rule = classic_rule(2, 0, &formula, template, parameter).unwrap();
            let xlsx_model::CfRule::AboveAverage {
                above_average,
                equal_average,
                std_dev,
                dxf_id,
                priority,
                stop_if_true,
            } = rule
            else {
                panic!("aboveAverage for template {template:#x}");
            };
            assert_eq!((above_average, equal_average, std_dev), expected);
            assert_eq!((dxf_id, priority, stop_if_true), (Some(0), 3, false));
        }
        assert!(take_best_effort().is_empty());
        // A standard-deviation count the rule cannot have is no offset.
        let rule = classic_rule(2, 0, &formula, 0x19, 9).unwrap();
        assert!(matches!(
            rule,
            xlsx_model::CfRule::AboveAverage {
                above_average: true,
                std_dev: None,
                ..
            }
        ));
        assert_eq!(take_best_effort(), [(BestEffort::IgnoredValue, TEMPLATE)]);
        // The rule still needs its formula slot to be well formed, and the
        // compatibility formula to be readable (a PtgInt without its value).
        assert!(classic_rule(2, 0, &[], 0x19, 0).is_err());
        assert!(classic_rule(2, 0, &[0x1e], 0x19, 0).is_err());
        // A formula the decompiler merely has no text for does not matter.
        assert!(classic_rule(2, 0, &[0x10], 0x19, 0).is_ok());
        take_best_effort();
    }

    #[test]
    fn value_and_date_templates_keep_their_compatibility_formula() {
        take_best_effort();
        // COUNTIF-like stand-in: 1+2 (PtgInt, PtgInt, PtgAdd).
        let formula = [0x1e, 1, 0, 0x1e, 2, 0, 0x03];
        for template in [0x07, 0x0f, 0x12, 0x18, 0x1b] {
            let rule = classic_rule(2, 0, &formula, template, 0).unwrap();
            let xlsx_model::CfRule::Expression {
                formula,
                dxf_id,
                priority,
                stop_if_true,
            } = rule
            else {
                panic!("expression for template {template:#x}");
            };
            assert_eq!(
                (formula.as_str(), dxf_id, priority, stop_if_true),
                ("1+2", Some(0), 3, false)
            );
            assert_eq!(take_best_effort(), [(BestEffort::Approximated, TEMPLATE)]);
        }
        // A template the specification does not list leaves the formula
        // condition its ct states; the formula templates need no decision.
        let rule = classic_rule(2, 0, &formula, 0x06, 0).unwrap();
        assert!(matches!(
            rule,
            xlsx_model::CfRule::Expression { ref formula, .. } if formula == "1+2"
        ));
        assert_eq!(take_best_effort(), [(BestEffort::IgnoredValue, TEMPLATE)]);
        classic_rule(2, 0, &formula, 0x08, 0).unwrap();
        assert!(take_best_effort().is_empty());
        // A cell-value rule is its comparison whatever template it names.
        let rule = classic_rule(1, 5, &[0x1e, 5, 0], 0x1b, 0).unwrap();
        let xlsx_model::CfRule::CellIs {
            operator, formulas, ..
        } = rule
        else {
            panic!("cellIs");
        };
        assert_eq!(
            (operator.as_str(), formulas),
            ("greaterThan", vec!["5".to_string()])
        );
        assert_eq!(take_best_effort(), [(BestEffort::IgnoredValue, TEMPLATE)]);
        classic_rule(1, 5, &[0x1e, 5, 0], 0, 0).unwrap();
        assert!(take_best_effort().is_empty());
    }

    #[test]
    fn value_formulas_are_the_cfvo_val_text() {
        take_best_effort();
        // Thresholds: percent 0, the formula 1+2 (cfvoType 7) and a number
        // given by the formula 5.
        let mut body = vec![0, 0, 0, 3, 5, 0];
        body.extend(cfvo(4, Some(0.0)));
        body.extend([1, 0, 0, 0, 0]);
        body.extend([7, 7, 0, 0x1e, 1, 0, 0x1e, 2, 0, 0x03]);
        body.extend([1, 0, 0, 0, 0]);
        body.extend([1, 3, 0, 0x1e, 5, 0]);
        body.extend([1, 0, 0, 0, 0]);
        let formats =
            project_records(&[(0x0879, cond_fmt12(1)), (0x087a, cf12(6, 4, 0, &[], &body))])
                .unwrap();
        let xlsx_model::CfRule::IconSet { cfvos, .. } = &formats[0].rules[0] else {
            panic!("icon set");
        };
        let values: Vec<_> = cfvos
            .iter()
            .map(|value| (value.kind.as_str(), value.value.as_deref()))
            .collect();
        assert_eq!(
            values,
            [
                ("percent", Some("0")),
                ("formula", Some("1+2")),
                ("num", Some("5"))
            ]
        );
        // A data bar from a formula minimum to the maximum.
        let mut bar = vec![0, 0, 0, 0, 0, 100];
        bar.extend(rgb(0x20, 0xa4, 0x72));
        bar.extend([7, 3, 0, 0x1e, 5, 0]);
        bar.extend(cfvo(3, None));
        let formats =
            project_records(&[(0x0879, cond_fmt12(1)), (0x087a, cf12(4, 3, 0, &[], &bar))])
                .unwrap();
        let xlsx_model::CfRule::DataBar { min, max, .. } = &formats[0].rules[0] else {
            panic!("data bar");
        };
        assert_eq!(
            (min.kind.as_str(), min.value.as_deref()),
            ("formula", Some("5"))
        );
        assert_eq!((max.kind.as_str(), max.value.as_deref()), ("max", None));
        // Reading a formula is not a best-effort decision.
        assert!(take_best_effort().is_empty());
        // cfvoType 7 without a formula is malformed.
        let mut bar = vec![0, 0, 0, 0, 0, 100];
        bar.extend(rgb(0x20, 0xa4, 0x72));
        bar.extend([7, 0, 0]);
        bar.extend(cfvo(3, None));
        assert!(
            project_records(&[(0x0879, cond_fmt12(1)), (0x087a, cf12(4, 3, 0, &[], &bar))])
                .unwrap_err()
                .contains("lacks its formula")
        );
    }

    #[test]
    fn display_options_the_model_cannot_state_draw_the_models_rule() {
        take_best_effort();
        // fIconOnly: the icons are drawn beside the values.
        let plain = project_records(&[
            (0x0879, cond_fmt12(1)),
            (0x087a, cf12(6, 4, 0, &[], &icon_set(0x04, 1))),
        ])
        .unwrap();
        assert!(take_best_effort().is_empty());
        let icon_only = project_records(&[
            (0x0879, cond_fmt12(1)),
            (0x087a, cf12(6, 4, 0, &[], &icon_set(0x05, 1))),
        ])
        .unwrap();
        assert_eq!(
            take_best_effort(),
            [(BestEffort::Approximated, "unsupported XLS icon set display")]
        );
        assert_eq!(
            serde_json::to_value(&icon_only).unwrap(),
            serde_json::to_value(&plain).unwrap()
        );
        assert!(matches!(
            &icon_only[0].rules[0],
            xlsx_model::CfRule::IconSet { reverse: true, icon_set, .. } if icon_set == "3Signs"
        ));
        // A data bar that hides its value (flags 2), runs from the right
        // (flags 1) or has other lengths is the model's bar.
        let bar = |flags: u8, minimum: u8, maximum: u8| {
            let mut bar = vec![0, 0, 0, flags, minimum, maximum];
            bar.extend(rgb(1, 2, 3));
            bar.extend(cfvo(2, None));
            bar.extend(cfvo(3, None));
            project_records(&[(0x0879, cond_fmt12(1)), (0x087a, cf12(4, 3, 0, &[], &bar))]).unwrap()
        };
        let plain = serde_json::to_value(bar(0, 0, 100)).unwrap();
        assert!(take_best_effort().is_empty());
        assert_eq!(plain[0]["rules"][0]["type"], "dataBar");
        assert_eq!(plain[0]["rules"][0]["color"], "#010203");
        for (flags, minimum, maximum) in [(2, 0, 100), (1, 0, 100), (0, 10, 90)] {
            assert_eq!(
                serde_json::to_value(bar(flags, minimum, maximum)).unwrap(),
                plain
            );
            assert_eq!(
                take_best_effort(),
                [(BestEffort::Approximated, "unsupported XLS data bar display")]
            );
        }
    }

    /// DXFN12 body: a DXFN that changes nothing, then an XFExtNoFRT with
    /// these ExtProp entries (type, value).
    fn dxfn12(extensions: &[(u16, Vec<u8>)]) -> Vec<u8> {
        let mut data = 0x003f_ffffu32.to_le_bytes().to_vec();
        data.extend(0u16.to_le_bytes());
        data.extend([0, 0, 0xff, 0xff, 0, 0]);
        data.extend((extensions.len() as u16).to_le_bytes());
        for (kind, value) in extensions {
            data.extend(kind.to_le_bytes());
            data.extend((value.len() as u16 + 4).to_le_bytes());
            data.extend(value);
        }
        data
    }

    /// XFExtGradient (2.5.279) at `degree` with these stops (color type,
    /// value, position).
    fn gradient_extension(degree: f64, stops: &[(u16, u32, f64)]) -> Vec<u8> {
        let mut value = 0u32.to_le_bytes().to_vec();
        for number in [degree, 0.0, 0.0, 0.0, 0.0] {
            value.extend(number.to_le_bytes());
        }
        value.extend((stops.len() as u32).to_le_bytes());
        for (kind, color, position) in stops {
            value.extend(kind.to_le_bytes());
            value.extend(color.to_le_bytes());
            value.extend(position.to_le_bytes());
            value.extend(0f64.to_le_bytes());
        }
        value
    }

    fn rule_dxf(dxf: &[u8]) -> Result<xlsx_model::Dxf, String> {
        let styles = styles::Styles::parse(&[]).unwrap();
        let theme = theme::Colors::default();
        let mut retained = Records::default();
        for (kind, data) in [
            (0x01b0, cond_fmt(1, 1)),
            (0x01b1, cf(1, 5, &[0x1e, 5, 0])),
            (0x087b, cf_ex(1, 5, 0, 0, dxf)),
        ] {
            retained.push(kind, &data).unwrap();
        }
        let mut dxfs = Vec::new();
        project(
            &retained,
            &Context {
                styles: &styles,
                theme: &theme,
                externs: &Externs::default(),
            },
            &mut dxfs,
        )?;
        assert_eq!(dxfs.len(), 1);
        Ok(dxfs.remove(0))
    }

    #[test]
    fn differential_gradient_fills_and_unknown_extensions() {
        take_best_effort();
        // White at 0 to red at 1 (LongRGBA), stored in reverse order.
        let gradient = gradient_extension(90.0, &[(2, 0xff00_00ff, 1.0), (2, 0xffff_ffff, 0.0)]);
        let dxf = rule_dxf(&dxfn12(&[(6, gradient.clone())])).unwrap();
        assert_eq!(
            serde_json::to_value(&dxf.fill).unwrap(),
            serde_json::json!({
                "patternType": "",
                "fgColor": null,
                "bgColor": null,
                "gradient": {
                    "gradientType": "linear",
                    "degree": 90.0,
                    "left": 0.0, "right": 0.0, "top": 0.0, "bottom": 0.0,
                    "stops": [
                        {"position": 0.0, "color": "#FFFFFF"},
                        {"position": 1.0, "color": "#FF0000"}
                    ]
                }
            })
        );
        assert!(take_best_effort().is_empty());
        // An extension without a property here is passed over; the gradient
        // beside it is read.
        let dxf = rule_dxf(&dxfn12(&[(0x000f, vec![2, 0]), (6, gradient.clone())])).unwrap();
        assert!(dxf.fill.unwrap().gradient.is_some());
        assert_eq!(
            take_best_effort(),
            [(
                BestEffort::Skipped,
                "unsupported XLS differential format extension"
            )]
        );
        // A stop of no specified color type is left out; a gradient
        // without any stop is no fill.
        let partial = gradient_extension(0.0, &[(9, 0, 0.0), (2, 0xff00_00ff, 1.0)]);
        let fill = rule_dxf(&dxfn12(&[(6, partial)])).unwrap().fill.unwrap();
        assert_eq!(fill.gradient.unwrap().stops.len(), 1);
        assert_eq!(
            take_best_effort(),
            [(
                BestEffort::IgnoredValue,
                "unsupported XLS differential color"
            )]
        );
        let none = gradient_extension(0.0, &[(9, 0, 0.0)]);
        assert!(rule_dxf(&dxfn12(&[(6, none)])).unwrap().fill.is_none());
        // cGradStops may be zero.
        let empty = gradient_extension(0.0, &[]);
        assert!(rule_dxf(&dxfn12(&[(6, empty)])).unwrap().fill.is_none());
        take_best_effort();
        // A stop list the extension does not hold is malformed.
        let mut truncated = gradient;
        truncated.pop();
        assert!(rule_dxf(&dxfn12(&[(6, truncated)]))
            .unwrap_err()
            .contains("invalid XLS differential gradient fill"));
    }

    #[test]
    fn unprojected_rules_are_rejected() {
        let valid = cf12(6, 4, 0, &[], &icon_set(0, 1));
        for records in [
            // Strict thresholds, activity formula, template mismatch.
            vec![
                (0x0879, cond_fmt12(1)),
                (0x087a, cf12(6, 4, 0, &[], &icon_set(0, 0))),
            ],
            vec![
                (0x0879, cond_fmt12(1)),
                (0x087a, cf12(6, 4, 0, &[0x1d, 1], &icon_set(0, 1))),
            ],
            vec![
                (0x0879, cond_fmt12(1)),
                (0x087a, cf12(6, 3, 0, &[], &icon_set(0, 1))),
            ],
            // Missing and orphan rules.
            vec![(0x0879, cond_fmt12(2)), (0x087a, valid.clone())],
            vec![(0x087a, valid.clone())],
            // A CondFmt without rules; a formula rule without a formula.
            vec![(0x01b0, vec![0; 20])],
            vec![(0x0879, cond_fmt12(1)), (0x087a, cf12(2, 1, 0, &[], &[]))],
        ] {
            assert!(project_records(&records).is_err());
        }
    }
}
