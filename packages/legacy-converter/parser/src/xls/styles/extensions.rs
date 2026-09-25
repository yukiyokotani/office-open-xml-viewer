//! MS-XLS 2.4.354/355: bind XFExt to the exact XF sequence before using it.
//! ExtProp 2.5.108, FullColorExt 2.5.155, LongRGBA 2.5.178.
use super::super::{u16_at, u32_at, unsupported, Record};
use super::color::ColorIdentity;
use std::collections::{BTreeMap, BTreeSet};

#[derive(Default)]
pub(super) struct Extensions {
    colors: BTreeMap<usize, BTreeMap<u16, ColorIdentity>>,
    indents: BTreeMap<usize, u16>,
    gradients: BTreeMap<usize, Gradient>,
    /// An owned extension carries formatting the model does not resolve: an
    /// indexed color with a tint (the palette fallback lacks the tint) or
    /// an undefined property. Font schemes (0x000E) are not counted: the
    /// XF's Font record already names the resolved face.
    pub(super) unrepresented: bool,
}

/// ExtProp 0x0006 XFExtGradient (2.5.280): an XFPropGradient (2.5.286)
/// and its GradStop (2.5.156) list.
#[derive(Clone, Debug)]
pub(super) struct Gradient {
    pub(super) path: bool,
    pub(super) degree: f64,
    pub(super) left: f64,
    pub(super) right: f64,
    pub(super) top: f64,
    pub(super) bottom: f64,
    pub(super) stops: Vec<(f64, StopColor)>,
}

/// A gradient stop color before palette resolution.
#[derive(Clone, Copy, Debug)]
pub(super) enum StopColor {
    Argb([u8; 4]),
    /// IcvXF with a tint, resolved against the workbook palette.
    Indexed(u16, f64),
}

impl Extensions {
    pub(super) fn parse(records: &[Record<'_>], xfs: &[&[u8]]) -> Result<Self, String> {
        let globals = || records.iter().take_while(|r| r.kind != super::super::EOF);
        let mut checks = globals().filter(|r| r.kind == 0x087c);
        let Some(check) = checks.next() else {
            return Ok(Self::default());
        };
        if checks.next().is_some() || check.data.len() != 20 || u16_at(check.data, 0)? != 0x087c {
            return Err(unsupported("invalid BIFF XF checksum record"));
        }
        // A stale extension must not replace newer palette-based formatting.
        // No extension is admitted without both count and checksum agreement.
        if !(16..=4050).contains(&xfs.len())
            || usize::from(u16_at(check.data, 14)?) != xfs.len()
            || u32_at(check.data, 16)? != checksum(xfs.iter().flat_map(|v| v.iter().copied()))
        {
            return Ok(Self::default());
        }
        let mut theme = None;
        let mut result = Self::default();
        let mut seen = BTreeSet::new();
        for record in globals().filter(|r| r.kind == 0x087d) {
            let data = record.data;
            if data.len() < 20 || u16_at(data, 0)? != 0x087d {
                return Err(unsupported("invalid BIFF XF extension header"));
            }
            let index = usize::from(u16_at(data, 14)?);
            if index >= xfs.len() || !seen.insert(index) {
                return Err(unsupported("invalid or duplicate BIFF XF extension index"));
            }
            // StyleXF reserves this bit; only CellXF uses fHasXFExt (bit25
            // of the border/fill word). Unowned extensions cannot trigger theme
            // inflation, even though their record structure is still validated.
            let owned = u16_at(xfs[index], 4)? & 4 != 0 || xfs[index][17] & 2 != 0;
            let count = usize::from(u16_at(data, 18)?);
            // Resource policy; includes unknown future properties without retaining them.
            if count > 1024 {
                return Err(unsupported("too many BIFF XF extension properties"));
            }
            let mut offset = 20usize;
            let mut colors = BTreeMap::new();
            let mut indent = None;
            let mut gradients_found = None;
            let mut properties = BTreeSet::new();
            for _ in 0..count {
                let kind = u16_at(data, offset)?;
                let size = usize::from(u16_at(data, offset + 2)?);
                if size < 4 || size > data.len().saturating_sub(offset) || !properties.insert(kind)
                {
                    return Err(unsupported("invalid BIFF XF extension property"));
                }
                let value = &data[offset + 4..offset + size];
                offset += size;
                if matches!(kind, 4 | 5 | 7..=11 | 13) {
                    if value.len() != 16 {
                        return Err(unsupported("invalid BIFF extended color size"));
                    }
                    let color_type = u16_at(value, 0)?;
                    if color_type > 4 {
                        return Err(unsupported("invalid BIFF extended color type"));
                    }
                    if owned && color_type == 1 && u16_at(value, 2)? != 0 {
                        result.unrepresented = true;
                    }
                    // Resolve owned theme colors to SML ARGB.
                    // MS-XLS 2.5.49 lists 0..3 as Dark 1, Light 1, Dark 2,
                    // Light 2, but Excel uses the SpreadsheetML order: 0 = lt1,
                    // 1 = dk1, 2 = lt2, 3 = dk2. Evidence: Excel writes a
                    // palette fallback (icvFore) next to every XFExt fill; in
                    // the local corpus theme 0 always pairs with white/silver
                    // (icv 9, 22, 55 as tint darkens), 1 with black (icv 8) or
                    // dark grey (63) under positive tints, 2 with white/silver
                    // and 3 with dark blue (62) -- 211 fills, no exception.
                    // Font colours agree: theme 1 (dark text) dominates, as
                    // theme="1" does in the paired Excel-saved XLSX files.
                    // nTintShade has no stated scale in 2.5.155. Excel writes
                    // its standard tints as n/32767 (26213, 13106, 19660, -8191
                    // and 16383 are 0.8, 0.4, 0.6, -0.25 and 0.5, the dominant
                    // values in a local corpus of Excel-saved workbooks), and
                    // the SpreadsheetML tint algorithm (ECMA-376 §18.8.19)
                    // then applies to the base color.
                    let tint = f64::from(u16_at(value, 2)? as i16) / 32767.0;
                    if owned && color_type == 3 && u32_at(value, 4)? <= 11 {
                        if theme.is_none() {
                            theme = Some(super::super::theme::Colors::parse(records)?);
                        }
                        // Theme slots are stored in clrScheme order (dk1, lt1,
                        // dk2, lt2, accent1..); swap each light/dark pair.
                        let index = u32_at(value, 4)?;
                        let slot = if index < 4 { index ^ 1 } else { index };
                        if let Some(argb) = theme.as_ref().unwrap().argb(slot) {
                            colors.insert(kind, ColorIdentity::Argb(tinted(argb, tint)));
                        }
                    }
                    if color_type == 2 {
                        colors.insert(
                            kind,
                            ColorIdentity::Argb(tinted(
                                [value[7], value[4], value[5], value[6]],
                                tint,
                            )),
                        );
                    }
                } else if !matches!(kind, 0x0006 | 0x000e | 0x000f) {
                    result.unrepresented |= owned;
                } else if kind == 0x0006 {
                    let gradient = gradient(value, owned, &mut theme, records)?;
                    if let Some(gradient) = gradient {
                        gradients_found = Some(gradient);
                    }
                } else if kind == 0x000f {
                    if value.len() != 2 {
                        return Err(unsupported("invalid BIFF extended indentation size"));
                    }
                    let value = u16_at(value, 0)?;
                    if value > 250 {
                        return Err(unsupported("invalid BIFF extended indentation"));
                    }
                    indent = Some(value);
                }
            }
            if offset != data.len() {
                return Err(unsupported("unexpected BIFF XF extension tail"));
            }
            if owned {
                result.colors.insert(index, colors);
                // MS-XLS 2.2.6.1.2.1 permits an XFExt for StyleXF as well as
                // CellXF. ExtProp 0x000F extends either owning XF's cIndent.
                if let Some(indent) = indent {
                    result.indents.insert(index, indent);
                }
                if let Some(gradient) = gradients_found {
                    result.gradients.insert(index, gradient);
                }
            }
        }
        Ok(result)
    }

    pub(super) fn color(&self, index: usize, property: u16) -> Option<ColorIdentity> {
        self.colors.get(&index)?.get(&property).copied()
    }

    pub(super) fn gradient(&self, index: usize) -> Option<&Gradient> {
        self.gradients.get(&index)
    }

    pub(super) fn indent(&self, index: usize) -> Option<u16> {
        self.indents.get(&index).copied()
    }
}

/// XFExtGradient of an XF; `None` for an XF that does not own its
/// extension (its gradient is validated but not applied).
fn gradient(
    value: &[u8],
    owned: bool,
    theme: &mut Option<super::super::theme::Colors>,
    records: &[Record<'_>],
) -> Result<Option<Gradient>, String> {
    use super::super::f64_at;
    let number = |at: usize| -> Result<f64, String> {
        let number = f64_at(value, at)?;
        number
            .is_finite()
            .then_some(number)
            .ok_or_else(|| unsupported("invalid BIFF gradient fill"))
    };
    let kind = u32_at(value, 0)?;
    let count = usize::try_from(u32_at(value, 44)?)
        .map_err(|_| unsupported("invalid BIFF gradient fill"))?;
    if kind > 1 || count > 256 || value.len() != 48 + count * 22 {
        return Err(unsupported("invalid BIFF gradient fill"));
    }
    let mut stops = Vec::with_capacity(count);
    for stop in 0..count {
        let at = 48 + stop * 22;
        let color_type = u16_at(value, at)?;
        let color = u32_at(value, at + 2)?;
        let position = number(at + 6)?;
        let tint = number(at + 14)?;
        if !(0.0..=1.0).contains(&position) || !(-1.0..=1.0).contains(&tint) {
            return Err(unsupported("invalid BIFF gradient stop"));
        }
        let color = match color_type {
            1 => StopColor::Indexed(
                u16::try_from(color).map_err(|_| unsupported("invalid BIFF gradient stop"))?,
                tint,
            ),
            2 => {
                let [r, g, b, a] = color.to_le_bytes();
                StopColor::Argb(tinted([a, r, g, b], tint))
            }
            // ColorTheme in SpreadsheetML order over the clrScheme-ordered
            // theme, as for the XFExt colors above.
            3 if color <= 11 => {
                if !owned {
                    return Ok(None);
                }
                if theme.is_none() {
                    *theme = Some(super::super::theme::Colors::parse(records)?);
                }
                let slot = if color < 4 { color ^ 1 } else { color };
                let argb = theme
                    .as_ref()
                    .unwrap()
                    .argb(slot)
                    .ok_or_else(|| unsupported("BIFF themed gradient stop lacks a theme"))?;
                StopColor::Argb(tinted(argb, tint))
            }
            _ => return Err(unsupported("unsupported BIFF gradient stop color")),
        };
        stops.push((position, color));
    }
    if stops.is_empty() {
        return Err(unsupported("BIFF gradient fill without stops"));
    }
    stops.sort_by(|a, b| a.0.partial_cmp(&b.0).unwrap_or(std::cmp::Ordering::Equal));
    Ok(owned.then_some(Gradient {
        path: kind == 1,
        degree: number(4)?,
        left: number(12)?,
        right: number(20)?,
        top: number(28)?,
        bottom: number(36)?,
        stops,
    }))
}

/// Apply a SpreadsheetML tint (ECMA-376 §18.8.19) through the shared resolver.
fn tinted(argb: [u8; 4], tint: f64) -> [u8; 4] {
    use ooxml_common::spreadsheet_color::{resolve_color, SpreadsheetColor};
    if tint == 0.0 {
        return argb;
    }
    let Some(hex) = resolve_color(SpreadsheetColor::Argb(argb), Some(tint), &[]) else {
        return argb;
    };
    let hex = hex.trim_start_matches('#');
    let channel = |at: usize| u8::from_str_radix(&hex[at..at + 2], 16).unwrap_or(0);
    if hex.len() < 6 {
        return argb;
    }
    [argb[0], channel(0), channel(2), channel(4)]
}

// MS-OSHARED 2.4.3 MsoCrc32Compute: non-reflected MSB-first polynomial
// x^32+x^7+x^5+x^3+x^2+x+1, zero initial remainder, no final complement.
// Streaming avoids allocating a concatenated copy of the XF table.
fn checksum(bytes: impl Iterator<Item = u8>) -> u32 {
    let mut crc = 0u32;
    for byte in bytes {
        crc ^= u32::from(byte) << 24;
        for _ in 0..8 {
            crc = (crc << 1) ^ if crc & 0x8000_0000 != 0 { 0xaf } else { 0 };
        }
    }
    crc
}

#[cfg(test)]
mod tests {
    use super::*;

    fn stop(kind: u16, value: u32, position: f64, tint: f64) -> Vec<u8> {
        let mut data = kind.to_le_bytes().to_vec();
        data.extend(value.to_le_bytes());
        data.extend(position.to_le_bytes());
        data.extend(tint.to_le_bytes());
        data
    }

    #[test]
    fn gradient_fills_read_their_geometry_and_stops() {
        let mut value = 0u32.to_le_bytes().to_vec();
        for number in [90.0f64, 0.0, 0.0, 0.0, 0.0] {
            value.extend(number.to_le_bytes());
        }
        value.extend(2u32.to_le_bytes());
        value.extend(stop(2, 0xff00_00ff, 1.0, 0.0));
        value.extend(stop(2, 0xff00_ff00, 0.0, 0.0));
        let parsed = gradient(&value, true, &mut None, &[]).unwrap().unwrap();
        assert!(!parsed.path);
        assert_eq!(parsed.degree, 90.0);
        // Stops sort by position; LongRGBA is red, green, blue, alpha.
        assert!(matches!(
            parsed.stops[0],
            (0.0, StopColor::Argb([0xff, 0, 0xff, 0]))
        ));
        assert!(matches!(
            parsed.stops[1],
            (1.0, StopColor::Argb([0xff, 0xff, 0, 0]))
        ));
        // An XF that does not own the extension keeps its pattern fill.
        assert!(gradient(&value, false, &mut None, &[]).unwrap().is_none());
        // Truncated stop lists and unknown color types reject.
        assert!(gradient(&value[..value.len() - 1], true, &mut None, &[]).is_err());
        let mut bad = value.clone();
        bad[48] = 9;
        assert!(gradient(&bad, true, &mut None, &[]).is_err());
    }
}
