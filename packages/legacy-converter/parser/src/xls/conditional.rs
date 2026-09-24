//! Direct projection of BIFF8 conditional formatting onto the XLSX model.
//!
//! MS-XLS 2.4.57 CondFmt12 with 2.4.43 CF12, whose rgbCT is a 2.5.32
//! CFGradient (color scale), 2.5.22 CFDatabar (data bar) or 2.5.36
//! CFMultistate (icon set); values are 2.5.39 CFVO and colors 2.5.21 CFColor.
//! The XLSX model rules are the ones the XLSX parser produces for the same
//! ECMA-376 18.3.1.x elements; the private corpus pairs each .xls with the
//! .xlsx Excel saved from it, which shows the expected SpreadsheetML.
//!
//! Anything this module does not project rejects the workbook: classic
//! CondFmt/CF/CFEx records and formula rules until their formulas and
//! differential formats are projected, CFVO formulas, activity formulas and
//! display options the shared renderer cannot show (icon-only, data-bar
//! value hiding or direction, non-default data-bar lengths, strict icon
//! thresholds). Conditional formatting is never dropped silently.

use super::{f64_at, styles, theme, u16_at, u32_at, unsupported};

/// Retained raw records of one worksheet substream, in stream order.
#[derive(Default)]
pub(super) struct Records {
    records: Vec<(u16, Vec<u8>)>,
    bytes: usize,
}

/// Resource policy for retained CF record bytes per workbook sheet.
const MAX_SHEET_BYTES: usize = 16 * 1024 * 1024;

impl Records {
    pub(super) fn push(&mut self, kind: u16, data: &[u8]) -> Result<(), String> {
        self.bytes = self
            .bytes
            .checked_add(data.len())
            .filter(|bytes| *bytes <= MAX_SHEET_BYTES)
            .ok_or_else(|| unsupported("XLS conditional formatting byte budget exceeded"))?;
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
}

/// Project every conditional format of one worksheet.
pub(super) fn project(
    records: &Records,
    context: &Context<'_>,
) -> Result<Vec<xlsx_model::ConditionalFormat>, String> {
    let mut output = Vec::new();
    let mut pending: Option<(xlsx_model::ConditionalFormat, usize)> = None;
    for (kind, data) in &records.records {
        match *kind {
            0x0879 => {
                if pending.as_ref().is_some_and(|(_, left)| *left != 0) {
                    return Err(unsupported("XLS CondFmt12 lacks its CF12 records"));
                }
                if let Some((format, _)) = pending.take() {
                    output.push(format);
                }
                pending = Some(cond_fmt12(data)?);
            }
            0x087a => {
                let (format, left) = pending
                    .as_mut()
                    .filter(|(_, left)| *left != 0)
                    .ok_or_else(|| unsupported("XLS CF12 outside its CondFmt12"))?;
                format.rules.push(cf12(data, context)?);
                *left -= 1;
            }
            0x01b0 | 0x01b1 | 0x087b => {
                return Err(unsupported(
                    "XLS CondFmt/CF/CFEx conditional formatting is not projected yet",
                ));
            }
            _ => return Err(unsupported("unexpected XLS conditional formatting record")),
        }
    }
    match pending {
        Some((_, left)) if left != 0 => Err(unsupported("XLS CondFmt12 lacks its CF12 records")),
        Some((format, _)) => {
            output.push(format);
            Ok(output)
        }
        None => Ok(output),
    }
}

/// MS-XLS 2.4.57: FrtRefHeaderU, then CondFmtStructure (2.5.56).
fn cond_fmt12(data: &[u8]) -> Result<(xlsx_model::ConditionalFormat, usize), String> {
    if u16_at(data, 0)? != 0x0879 {
        return Err(unsupported("invalid XLS CondFmt12 header"));
    }
    let ccf = usize::from(u16_at(data, 12)?);
    if ccf == 0 {
        return Err(unsupported("empty XLS CondFmt12"));
    }
    let (sqref, end) = sqref(data, 24)?;
    if end != data.len() {
        return Err(unsupported("unexpected XLS CondFmt12 tail"));
    }
    Ok((
        xlsx_model::ConditionalFormat {
            sqref,
            rules: Vec::new(),
        },
        ccf,
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

/// MS-XLS 2.4.43 CF12.
fn cf12(data: &[u8], context: &Context<'_>) -> Result<xlsx_model::CfRule, String> {
    if u16_at(data, 0)? != 0x087a {
        return Err(unsupported("invalid XLS CF12 header"));
    }
    let byte = |offset: usize| -> Result<u8, String> {
        data.get(offset)
            .copied()
            .ok_or_else(|| unsupported("truncated XLS CF12"))
    };
    let ct = byte(12)?;
    let (cce1, cce2) = (
        usize::from(u16_at(data, 14)?),
        usize::from(u16_at(data, 16)?),
    );
    // DXFN12: cbDxf, then two reserved bytes when it is zero.
    let cb_dxf = usize::try_from(u32_at(data, 18)?).map_err(|_| truncated())?;
    let mut offset = 22usize
        .checked_add(if cb_dxf == 0 { 2 } else { cb_dxf })
        .ok_or_else(truncated)?;
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
    slice(data, offset + 6, 16)?;
    let body = data.get(offset + 22..).ok_or_else(truncated)?;
    let visual = matches!(ct, 3 | 4 | 6);
    if !visual {
        return Err(unsupported(
            "XLS CF12 cell-value, formula and filter rules are not projected yet",
        ));
    }
    // CF12 2.4.43: color scales, data bars and icon sets carry no dxf, no
    // comparison formulas and no stop-if-true; their optional activity
    // formula (fmlaActive) has no XLSX-model equivalent here.
    if cb_dxf != 0 || !rgce1.is_empty() || !rgce2.is_empty() || flags & 0x02 != 0 {
        return Err(unsupported("invalid XLS CF12 visual rule"));
    }
    if active != 0 {
        return Err(unsupported("XLS CF12 activity formula is not supported"));
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
        3 => gradient(body, priority, context)?,
        4 => databar(body, priority, context)?,
        _ => multistate(body, priority)?,
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
fn cfvo(data: &[u8], offset: usize) -> Result<(xlsx_model::CfValue, usize), String> {
    let kind = *data.get(offset).ok_or_else(truncated)?;
    let cce = usize::from(u16_at(data, offset + 1)?);
    if cce != 0 {
        return Err(unsupported(
            "XLS conditional formatting value formula is not supported",
        ));
    }
    let mut end = offset + 3;
    let (name, value) = match kind {
        2 => ("min", None),
        3 => ("max", None),
        1 | 4 | 5 => {
            let value = f64_at(data, end)?;
            end += 8;
            if !value.is_finite() || (kind != 1 && !(0.0..=100.0).contains(&value)) {
                return Err(unsupported("invalid XLS conditional formatting value"));
            }
            let name = match kind {
                1 => "num",
                4 => "percent",
                _ => "percentile",
            };
            (name, Some(number(value)))
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

/// The shortest round-trip decimal, as a SpreadsheetML `cfvo@val` string.
fn number(value: f64) -> String {
    format!("{value}")
}

/// MS-XLS 2.5.21 CFColor as the XLSX model's resolved `#RRGGBB`.
fn color(data: &[u8], offset: usize, context: &Context<'_>) -> Result<String, String> {
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
            let base = context
                .styles
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
            let argb = context.theme.argb(slot).ok_or_else(|| {
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
        let (value, end) = cfvo(data, offset)?;
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
    Ok((xlsx_model::CfRule::ColorScale { stops, priority }, offset))
}

/// MS-XLS 2.5.22 CFDatabar onto ECMA-376 18.3.1.28 dataBar.
fn databar(
    data: &[u8],
    priority: i32,
    context: &Context<'_>,
) -> Result<(xlsx_model::CfRule, usize), String> {
    let flags = *data.get(3).ok_or_else(truncated)?;
    let (minimum, maximum) = (
        *data.get(4).ok_or_else(truncated)?,
        *data.get(5).ok_or_else(truncated)?,
    );
    // The shared renderer draws bars from the left across the full cell
    // width with the value shown. Excel's PDF of a workbook whose CFDatabar
    // stores flags 0 and lengths 0..100 shows left-to-right gradient bars
    // with the cell values, so only that combination is projected; other
    // flags and lengths stay rejected until observed.
    if flags != 0 || minimum != 0 || maximum != 100 {
        return Err(unsupported("unsupported XLS data bar display"));
    }
    let color = color(data, 6, context)?;
    let (min, end) = cfvo(data, 22)?;
    let (max, end) = cfvo(data, end)?;
    Ok((
        xlsx_model::CfRule::DataBar {
            color,
            min,
            max,
            priority,
            // CF12 data bars have no fill-type field; Excel draws them with
            // its 2007 gradient (observed in the same PDF).
            gradient: true,
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
fn multistate(data: &[u8], priority: i32) -> Result<(xlsx_model::CfRule, usize), String> {
    let states = usize::from(*data.get(3).ok_or_else(truncated)?);
    let set = usize::from(*data.get(4).ok_or_else(truncated)?);
    let flags = *data.get(5).ok_or_else(truncated)?;
    let expected = match set {
        0..=7 => 3,
        8..=12 => 4,
        13..=16 => 5,
        _ => return Err(unsupported("invalid XLS icon set")),
    };
    // fIconOnly hides the cell value, which the shared renderer cannot.
    if states != expected || flags & !0x04 != 0 {
        return Err(unsupported("unsupported XLS icon set display"));
    }
    let mut offset = 6;
    let mut cfvos = Vec::with_capacity(states);
    for _ in 0..states {
        let (value, end) = cfvo(data, offset)?;
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
            },
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

    #[test]
    fn unrepresentable_or_unprojected_rules_are_rejected() {
        let valid = cf12(6, 4, 0, &[], &icon_set(0, 1));
        for records in [
            // Icon-only display, strict thresholds, activity formula, template mismatch.
            vec![
                (0x0879, cond_fmt12(1)),
                (0x087a, cf12(6, 4, 0, &[], &icon_set(0x01, 1))),
            ],
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
            // Classic conditional formatting is not projected yet.
            vec![(0x01b0, vec![0; 20])],
            // Formula rules are not projected yet.
            vec![(0x0879, cond_fmt12(1)), (0x087a, cf12(2, 1, 0, &[], &[]))],
        ] {
            assert!(project_records(&records).is_err());
        }
        let mut hidden_value = vec![0, 0, 0, 0x02, 0, 100];
        hidden_value.extend(rgb(1, 2, 3));
        hidden_value.extend(cfvo(2, None));
        hidden_value.extend(cfvo(3, None));
        assert!(project_records(&[
            (0x0879, cond_fmt12(1)),
            (0x087a, cf12(4, 3, 0, &[], &hidden_value)),
        ])
        .is_err());
    }
}
