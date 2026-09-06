//! MS-XLS 2.4.354/355: bind XFExt to the exact XF sequence before using it.
//! ExtProp 2.5.108, FullColorExt 2.5.155, LongRGBA 2.5.178.
use super::super::{u16_at, u32_at, unsupported, Record};
use std::collections::{BTreeMap, BTreeSet};

#[derive(Default)]
pub(super) struct Extensions {
    colors: BTreeMap<usize, BTreeMap<u16, String>>,
    indents: BTreeMap<usize, u16>,
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
                    // Resolve owned, untinted theme colors to SML ARGB. Leave
                    // tint normalization separate; never scale RGB channels.
                    // Interop limitation: the light/dark order in MS-XLS
                    // 2.5.49 conflicts with Office-produced XFExt/palette/PDF
                    // evidence. Keep those four colors' original BIFF fallback
                    // until varied Office probes establish an approved mapping.
                    // Do not guess a swap or let unresolved indices inflate ZIPs.
                    if owned
                        && color_type == 3
                        && u16_at(value, 2)? == 0
                        && (4..=11).contains(&u32_at(value, 4)?)
                    {
                        if theme.is_none() {
                            theme = Some(super::super::theme::Colors::parse(records)?);
                        }
                        if let Some(color) = theme.as_ref().unwrap().rgb(u32_at(value, 4)?) {
                            colors.insert(kind, color);
                        }
                    }
                    if color_type == 2 && u16_at(value, 2)? == 0 {
                        colors.insert(
                            kind,
                            format!(
                                "rgb=\"{:02X}{:02X}{:02X}{:02X}\"",
                                value[7], value[4], value[5], value[6]
                            ),
                        );
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
            }
        }
        Ok(result)
    }

    pub(super) fn color(&self, index: usize, property: u16) -> Option<&str> {
        self.colors.get(&index)?.get(&property).map(String::as_str)
    }

    pub(super) fn indent(&self, index: usize) -> Option<u16> {
        self.indents.get(&index).copied()
    }
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
