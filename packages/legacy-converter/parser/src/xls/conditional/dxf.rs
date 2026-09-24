//! Differential formats of conditional-formatting rules: MS-XLS 2.5.95 DXFN
//! (CF), 2.5.96 DXFN12 (CF12, CFExNonCF12) with its 2.5.281 XFExtNoFRT
//! full-color extensions, onto the XLSX model `Dxf` that the XLSX parser
//! builds from ECMA-376 18.8.14 `dxf`.
//!
//! Each "ninch" bit marks a property the rule leaves unchanged, the same as
//! an omitted SpreadsheetML child. Alignment (DXFALC) and protection
//! (DXFProt) are validated for size but carried by neither the XLSX model's
//! `Dxf` nor its parser, so they are not represented for either format.

use super::super::{styles, u16_at, u32_at, unsupported};
use super::Context;

fn truncated() -> String {
    unsupported("truncated XLS differential format")
}

fn byte(data: &[u8], offset: usize) -> Result<u8, String> {
    data.get(offset).copied().ok_or_else(truncated)
}

fn i32_at(data: &[u8], offset: usize) -> Result<i32, String> {
    Ok(u32_at(data, offset)? as i32)
}

/// XLUnicodeString body (fHighByte then characters) of `count` characters.
fn unicode(data: &[u8], offset: usize, count: usize) -> Result<(String, usize), String> {
    match byte(data, offset)? {
        0 => Ok((
            data.get(offset + 1..offset + 1 + count)
                .ok_or_else(truncated)?
                .iter()
                .map(|&value| char::from(value))
                .collect(),
            1 + count,
        )),
        1 => {
            let bytes = data
                .get(offset + 1..offset + 1 + count * 2)
                .ok_or_else(truncated)?;
            let units: Vec<u16> = bytes
                .chunks_exact(2)
                .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
                .collect();
            Ok((
                String::from_utf16(&units).map_err(|_| truncated())?,
                1 + count * 2,
            ))
        }
        _ => Err(unsupported("invalid XLS differential format string")),
    }
}

/// IcvXF/IcvFont (2.5.161/2.5.160) as a dxf color. 0x40 and 0x41 are the
/// default foreground and background, and 0x7FFF the automatic font color;
/// Excel writes all of them as `auto="1"` in the dxfs of the Excel-saved
/// counterparts (sample-3: DXFPat icv 64/65 -> fgColor/bgColor auto), which
/// the XLSX model leaves unset.
fn icv(context: &Context<'_>, value: u16) -> Option<String> {
    match value {
        0x40 | 0x41 | 0x7fff => None,
        value => context.styles.icv_model(value),
    }
}

/// Where each property's color came from before XFExtNoFRT refines it.
#[derive(Default)]
struct Colors {
    font: Option<Option<String>>,
    fill_fg: Option<Option<String>>,
    fill_bg: Option<Option<String>>,
    top: bool,
    bottom: bool,
    left: bool,
    right: bool,
    diagonal: bool,
}

/// Parse a DXFN at `offset`; returns the model and the end offset.
fn dxfn(
    data: &[u8],
    offset: usize,
    context: &Context<'_>,
) -> Result<(xlsx_model::Dxf, usize, DxfnParts), String> {
    let flags = u32_at(data, offset)?;
    let extra = u16_at(data, offset + 4)?;
    let bit = |index: u32| flags & (1 << index) != 0;
    let mut at = offset + 6;
    let mut dxf = xlsx_model::Dxf::default();
    let mut parts = DxfnParts::default();

    // DXFNum 2.5.99.
    if bit(25) {
        let user = extra & 1 != 0;
        let format = if user {
            let size = usize::from(u16_at(data, at)?);
            let count = usize::from(u16_at(data, at + 2)?);
            let (code, used) = unicode(data, at + 4, count)?;
            if size != 4 + used {
                return Err(unsupported("invalid XLS differential number format"));
            }
            at += size;
            context
                .styles
                .format_id(&code)
                .map(|id| xlsx_model::NumFmt {
                    num_fmt_id: id.into(),
                    format_code: code.clone(),
                })
                .or(Some(xlsx_model::NumFmt {
                    num_fmt_id: 0,
                    format_code: code,
                }))
        } else {
            let id = byte(data, at + 1)?;
            at += 2;
            Some(xlsx_model::NumFmt {
                num_fmt_id: id.into(),
                format_code: context.styles.format_code(id.into()).unwrap_or_default(),
            })
        };
        // ifmtNinch leaves the number format unchanged.
        if !bit(19) {
            dxf.num_fmt = format;
        }
    }

    // DXFFntD 2.5.93: 63-byte name area, Stxp, then the ninch words.
    if bit(26) {
        let name_count = usize::from(byte(data, at)?);
        let name = if name_count > 0 {
            Some(unicode(data, at + 1, name_count)?.0)
        } else {
            None
        };
        let stxp = at + 64;
        let height = i32_at(data, stxp)?;
        let ts = u32_at(data, stxp + 4)?;
        let weight = u16_at(data, stxp + 8)?;
        let script = u16_at(data, stxp + 10)?;
        let underline = byte(data, stxp + 12)?;
        let font_icv = i32_at(data, stxp + 16)?;
        let ts_ninch = u32_at(data, stxp + 24)?;
        let script_ninch = u32_at(data, stxp + 28)? != 0;
        let underline_ninch = u32_at(data, stxp + 32)? != 0;
        let weight_ninch = u32_at(data, stxp + 36)? != 0;
        u16_at(data, stxp + 52)?;
        let italic = ts & 0x02 != 0 && ts_ninch & 0x02 == 0;
        let strike = ts & 0x80 != 0 && ts_ninch & 0x80 == 0;
        let bold = !weight_ninch && weight == 0x02bc;
        let (underline, underline_style) = if underline_ninch {
            (false, None)
        } else {
            match underline {
                0x00 => (false, None),
                0x01 => (true, None),
                0x02 => (true, Some("double")),
                0x21 => (true, Some("singleAccounting")),
                0x22 => (true, Some("doubleAccounting")),
                0xff => (false, None),
                _ => return Err(unsupported("invalid XLS differential underline")),
            }
        };
        let vert_align = match (script_ninch, script) {
            (false, 1) => Some("superscript".to_string()),
            (false, 2) => Some("subscript".to_string()),
            _ => None,
        };
        // icvFore -1 leaves the color unchanged; 32767 is the automatic
        // text color, which the XLSX model also leaves unset.
        let color = match font_icv {
            -1 => None,
            0x7fff => Some(None),
            value => Some(icv(
                context,
                u16::try_from(value)
                    .map_err(|_| unsupported("invalid XLS differential font color"))?,
            )),
        };
        parts.colors.font = color.clone();
        dxf.font = Some(xlsx_model::Font {
            bold,
            italic,
            underline,
            strike,
            // The XLSX parser's default when `sz` is absent.
            size: if height > 0 {
                f64::from(height) / 20.0
            } else {
                11.0
            },
            color: color.flatten(),
            name,
            underline_style: underline_style.map(str::to_string),
            vert_align,
        });
        at += 118;
    }

    // DXFALC 2.5.91.
    if bit(27) {
        data.get(at..at + 8).ok_or_else(truncated)?;
        at += 8;
    }

    // DXFBdr 2.5.92.
    if bit(28) {
        let first = u32_at(data, at)?;
        let second = u32_at(data, at + 4)?;
        let edge = |style: u32,
                    icv_index: u32,
                    ninch: bool|
         -> Result<Option<xlsx_model::BorderEdge>, String> {
            if ninch || style == 0 {
                return Ok(None);
            }
            let style = styles::BORDERS
                .get(style as usize)
                .ok_or_else(|| unsupported("invalid XLS differential border style"))?;
            Ok(Some(xlsx_model::BorderEdge {
                style: (*style).to_string(),
                color: icv(context, icv_index as u16),
            }))
        };
        let mut border = xlsx_model::Border {
            left: edge(first & 0xf, (first >> 16) & 0x7f, bit(10))?,
            right: edge((first >> 4) & 0xf, (first >> 23) & 0x7f, bit(11))?,
            top: edge((first >> 8) & 0xf, second & 0x7f, bit(12))?,
            bottom: edge((first >> 12) & 0xf, (second >> 7) & 0x7f, bit(13))?,
            ..Default::default()
        };
        let diagonal = edge(
            (second >> 21) & 0xf,
            (second >> 14) & 0x7f,
            bit(14) && bit(15),
        )?;
        if !bit(14) && first & (1 << 30) != 0 {
            border.diagonal_down = diagonal.clone();
        }
        if !bit(15) && first & (1 << 31) != 0 {
            border.diagonal_up = diagonal;
        }
        parts.colors.left = border.left.is_some();
        parts.colors.right = border.right.is_some();
        parts.colors.top = border.top.is_some();
        parts.colors.bottom = border.bottom.is_some();
        parts.colors.diagonal = border.diagonal_up.is_some() || border.diagonal_down.is_some();
        dxf.border = Some(border);
        at += 8;
    }

    // DXFPat 2.5.102. An omitted pattern (flsNinch) reads as the XLSX
    // parser's `patternFill` default, solid; a dxf fill with only a
    // background color paints it as the foreground, as the parser mirrors.
    if bit(29) {
        let value = u32_at(data, at)?;
        let pattern = if bit(16) {
            "solid"
        } else {
            styles::PATTERNS
                .get(((value >> 10) & 0x3f) as usize)
                .ok_or_else(|| unsupported("invalid XLS differential fill pattern"))?
        };
        let fg = (!bit(17)).then(|| icv(context, ((value >> 16) & 0x7f) as u16));
        let bg = (!bit(18)).then(|| icv(context, ((value >> 23) & 0x7f) as u16));
        parts.colors.fill_fg = fg.clone();
        parts.colors.fill_bg = bg.clone();
        dxf.fill = Some(xlsx_model::Fill {
            pattern_type: pattern.to_string(),
            fg_color: fg.flatten(),
            bg_color: bg.flatten(),
            gradient: None,
        });
        at += 4;
    }

    // DXFProt 2.5.103.
    if bit(30) {
        data.get(at..at + 2).ok_or_else(truncated)?;
        at += 2;
    }
    Ok((dxf, at, parts))
}

#[derive(Default)]
struct DxfnParts {
    colors: Colors,
}

/// DXFN12 at `offset`: cbDxf, then either two reserved bytes (no format)
/// or a DXFN with an optional XFExtNoFRT filling the rest of cbDxf.
pub(super) fn dxfn12(
    data: &[u8],
    offset: usize,
    context: &Context<'_>,
) -> Result<(Option<xlsx_model::Dxf>, usize), String> {
    let size = usize::try_from(u32_at(data, offset)?).map_err(|_| truncated())?;
    if size == 0 {
        data.get(offset + 4..offset + 6).ok_or_else(truncated)?;
        return Ok((None, offset + 6));
    }
    let start = offset + 4;
    let end = start.checked_add(size).ok_or_else(truncated)?;
    let body = data.get(..end).ok_or_else(truncated)?;
    let (mut dxf, at, parts) = dxfn(body, start, context)?;
    if at < end {
        extend(body, at, end, &mut dxf, &parts, context)?;
    } else if at > end {
        return Err(truncated());
    }
    mirror(&mut dxf);
    Ok((Some(dxf), end))
}

/// XFExtNoFRT 2.5.281: full colors refining the palette colors of
/// properties the DXFN specifies.
fn extend(
    data: &[u8],
    offset: usize,
    end: usize,
    dxf: &mut xlsx_model::Dxf,
    parts: &DxfnParts,
    context: &Context<'_>,
) -> Result<(), String> {
    if u16_at(data, offset + 2)? != 0xffff {
        return Err(unsupported("invalid XLS differential format extension"));
    }
    let count = usize::from(u16_at(data, offset + 6)?);
    let mut at = offset + 8;
    for _ in 0..count {
        let kind = u16_at(data, at)?;
        let size = usize::from(u16_at(data, at + 2)?);
        if size < 4 || at + size > end {
            return Err(unsupported("invalid XLS differential format extension"));
        }
        let value = &data[at + 4..at + size];
        at += size;
        match kind {
            4 | 5 | 7..=11 | 13 => {
                if value.len() != 16 {
                    return Err(unsupported("invalid XLS differential extended color"));
                }
                let color = full_color(value, context)?;
                let colors = &parts.colors;
                // An extension refines a color the DXFN specifies; Excel
                // writes the palette fallback next to each one.
                let applied = match kind {
                    4 | 5 => {
                        let fill = dxf.fill.as_mut();
                        let present = if kind == 4 {
                            colors.fill_fg.is_some()
                        } else {
                            colors.fill_bg.is_some()
                        };
                        match (fill, present) {
                            (Some(fill), true) => {
                                if kind == 4 {
                                    fill.fg_color = color;
                                } else {
                                    fill.bg_color = color;
                                }
                                true
                            }
                            _ => false,
                        }
                    }
                    13 => match (dxf.font.as_mut(), colors.font.is_some()) {
                        (Some(font), true) => {
                            font.color = color;
                            true
                        }
                        _ => false,
                    },
                    _ => {
                        let border = dxf.border.as_mut();
                        match border {
                            Some(border) => {
                                let edges: Vec<&mut Option<xlsx_model::BorderEdge>> = match kind {
                                    7 if colors.top => vec![&mut border.top],
                                    8 if colors.bottom => vec![&mut border.bottom],
                                    9 if colors.left => vec![&mut border.left],
                                    10 if colors.right => vec![&mut border.right],
                                    11 if colors.diagonal => {
                                        vec![&mut border.diagonal_up, &mut border.diagonal_down]
                                    }
                                    _ => Vec::new(),
                                };
                                let applied = !edges.is_empty();
                                for edge in edges.into_iter().flatten() {
                                    edge.color = color.clone();
                                }
                                applied
                            }
                            None => false,
                        }
                    }
                };
                if !applied {
                    return Err(unsupported(
                        "XLS differential color extension lacks its base property",
                    ));
                }
            }
            // FontScheme: the XLSX model's Dxf font carries no scheme.
            14 => {
                if value.len() != 2 {
                    return Err(unsupported("invalid XLS differential font scheme"));
                }
            }
            _ => return Err(unsupported("unsupported XLS differential format extension")),
        }
    }
    if at != end {
        return Err(unsupported("unexpected XLS differential format tail"));
    }
    Ok(())
}

/// The XLSX parser mirrors a dxf fill's lone background into the
/// foreground (a solid dxf fill paints its bgColor); keep the same model.
fn mirror(dxf: &mut xlsx_model::Dxf) {
    if let Some(fill) = dxf.fill.as_mut() {
        if fill.fg_color.is_none() {
            fill.fg_color = fill.bg_color.clone();
        }
    }
}

/// A CF record's DXFN (2.5.95) as the model.
pub(super) fn classic(
    data: &[u8],
    offset: usize,
    context: &Context<'_>,
) -> Result<(xlsx_model::Dxf, usize), String> {
    let (mut dxf, end, _) = dxfn(data, offset, context)?;
    mirror(&mut dxf);
    Ok((dxf, end))
}

/// FullColorExt 2.5.155 as a resolved `#RRGGBB` (None: automatic).
fn full_color(value: &[u8], context: &Context<'_>) -> Result<Option<String>, String> {
    use ooxml_common::spreadsheet_color::{resolve_color, SpreadsheetColor};
    let kind = u16_at(value, 0)?;
    // nTintShade in n/32767 steps, as for XFExt (styles/extensions.rs).
    let tint = f64::from(u16_at(value, 2)? as i16) / 32767.0;
    let tint = (tint != 0.0).then_some(tint);
    let data = u32_at(value, 4)?;
    Ok(match kind {
        0 => None,
        1 => {
            let base = icv(context, u16::try_from(data).map_err(|_| truncated())?);
            match (base, tint) {
                (Some(base), Some(_)) => {
                    let hex = base.trim_start_matches('#');
                    let channel = |at: usize| u8::from_str_radix(&hex[at..at + 2], 16).unwrap_or(0);
                    resolve_color(
                        SpreadsheetColor::Argb([0xff, channel(0), channel(2), channel(4)]),
                        tint,
                        &[],
                    )
                }
                (base, _) => base,
            }
        }
        2 => {
            let [r, g, b, a] = data.to_le_bytes();
            resolve_color(SpreadsheetColor::Argb([a, r, g, b]), tint, &[])
        }
        // ColorTheme in SpreadsheetML order over a clrScheme-ordered theme
        // (see styles/extensions.rs for the light/dark swap evidence).
        3 if data <= 11 => {
            let slot = if data < 4 { data ^ 1 } else { data };
            let argb = context
                .theme
                .argb(slot)
                .ok_or_else(|| unsupported("XLS themed differential color lacks a theme"))?;
            resolve_color(SpreadsheetColor::Argb(argb), tint, &[])
        }
        _ => return Err(unsupported("unsupported XLS differential color")),
    })
}
