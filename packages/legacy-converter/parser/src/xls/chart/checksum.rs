//! ShapePropsStream checksum (MS-XLS 2.4.258, 2.5.236; MS-OSHARED 2.4.3).
//!
//! MS-XLS requires a reader to discard the ShapePropsStream XML whenever its
//! `dwChecksum` differs from the checksum of the related LineFormat,
//! AreaFormat, MarkerFormat and GelFrame records, because another application
//! may have edited those records without updating the XML.
//!
//! The checksum data follows ShapePropsStreamChecksumData and its three parts
//! (2.5.173 line, 2.5.166 interior color, 2.5.128 fill style). Two details are
//! not stated by the specification and were established by recomputing the
//! stored checksum of every ShapePropsStream written by Excel in a local corpus
//! of Excel-saved charts (frames, axes, series, points, markers, drop bars):
//! - the CRC starts from 0;
//! - a LineFormat `lns` of 5 (no line) enters the `pattern` byte as 0xFF.
//! Every non-picture case matched with these rules. Picture-filled frames did
//! not yet match; such a mismatch only selects the BIFF records, which is the
//! specified fallback, so an unmatched case never adopts unverified XML.

/// MS-OSHARED 2.4.3.1 cache: polynomial 0xAF, entries masked to 16 bits.
fn cache() -> [u32; 256] {
    let mut table = [0u32; 256];
    for (index, slot) in table.iter_mut().enumerate() {
        let mut value = (index as u32) << 24;
        for _ in 0..8 {
            value = if value & 0x8000_0000 != 0 {
                (value << 1) ^ 0xaf
            } else {
                value << 1
            };
        }
        *slot = value & 0xffff;
    }
    table
}

/// MS-OSHARED 2.4.3.2 CRC over one byte stream.
pub(super) fn crc(bytes: &[u8]) -> u32 {
    let table = cache();
    let mut value = 0u32;
    for &byte in bytes {
        let index = ((value >> 24) ^ u32::from(byte)) as usize & 0xff;
        value = (value << 8) ^ table[index];
    }
    value
}

/// LinePropertiesForShapePropsStreamChecksum (2.5.173) from LineFormat data.
pub(super) fn line_properties(line_format: &[u8]) -> Option<[u8; 8]> {
    let rgb = line_format.get(0..4)?;
    let lns = u16::from_le_bytes(line_format.get(4..6)?.try_into().ok()?);
    let we = i16::from_le_bytes(line_format.get(6..8)?.try_into().ok()?);
    let flags = u16::from_le_bytes(line_format.get(8..10)?.try_into().ok()?);
    let icv = u16::from_le_bytes(line_format.get(10..12)?.try_into().ok()?);
    let pattern = if lns == 5 { 0xff } else { lns as u8 };
    Some([
        icv as u8,
        pattern,
        we.wrapping_add(1) as u8,
        (flags & 1) as u8,
        rgb[0],
        rgb[1],
        rgb[2],
        rgb[3],
    ])
}

/// InteriorColorPropertiesForShapePropsStreamChecksum (2.5.166). `colors` is
/// the rgbFore/rgbBack pair of the related AreaFormat or MarkerFormat.
pub(super) fn interior_properties(colors: &[u8], area_format: &[u8]) -> Option<[u8; 9]> {
    let colors = colors.get(0..8)?;
    let fls = area_format.get(8)?;
    let mut output = [0u8; 9];
    output[..8].copy_from_slice(colors);
    output[8] = *fls;
    Some(output)
}

const FILL_ORDER: [u16; 39] = [
    0x180, 0x181, 0x182, 0x183, 0x184, 0x185, 0x186, 0x187, 0x188, 0x189, 0x18a, 0x18b, 0x18c,
    0x18d, 0x18e, 0x18f, 0x190, 0x191, 0x192, 0x193, 0x194, 0x195, 0x196, 0x197, 0x198, 0x199,
    0x19a, 0x19b, 0x19c, 0x19e, 0x19f, 0x1a0, 0x1a1, 0x1a2, 0x1a3, 0x1a4, 0x1a5, 0x1a6, 0x1a7,
];

/// MS-ODRAW 2.3.7 defaults, or the fixed values 2.5.128 assigns to the
/// reserved415..423 entries.
fn fill_default(opid: u16) -> u32 {
    match opid {
        0x181 | 0x183 => 0x00ff_ffff,
        0x182 | 0x184 => 0x0001_0000,
        0x185 | 0x1a0 | 0x1a4 => 0x2000_0000,
        0x19c => 0x4000_0003,
        0x19e | 0x19f | 0x1a2 | 0x1a3 | 0x1a6 | 0x1a7 => 0xffff_ffff,
        _ => 0,
    }
}

/// One fill property from a GelFrame table: value and optional complex bytes.
pub(super) struct FillProperty<'a> {
    pub value: u32,
    pub complex: Option<&'a [u8]>,
}

/// FillStylePropertiesForShapePropsStreamChecksum (2.5.128). `lookup` returns
/// a GelFrame fill property by base opid; `boolean` is the fill-style Boolean
/// property (opid 0x1BF) when present.
pub(super) fn fill_style_properties<'a>(
    lookup: impl Fn(u16) -> Option<FillProperty<'a>>,
    boolean: Option<u32>,
    output: &mut Vec<u8>,
) {
    for opid in FILL_ORDER {
        let reserved = matches!(opid, 0x19f | 0x1a1 | 0x1a3 | 0x1a5 | 0x1a6 | 0x1a7);
        let property = (!reserved).then(|| lookup(opid)).flatten();
        let value = property
            .as_ref()
            .map_or_else(|| fill_default(opid), |p| p.value);
        output.extend_from_slice(&u32::from(opid).to_le_bytes());
        output.extend_from_slice(&value.to_le_bytes());
        let complex = property.as_ref().and_then(|p| p.complex);
        match opid {
            // fillBlip_complex_md4uid: rgbUid1 of the embedded OfficeArtBlip.
            0x186 if value > 0 => {
                if let Some(uid) = complex.and_then(|bytes| bytes.get(8..24)) {
                    output.extend_from_slice(uid);
                }
            }
            0x187 | 0x197 if value > 0 => {
                if let Some(bytes) = complex {
                    output.extend_from_slice(bytes);
                }
            }
            _ => {}
        }
    }
    // MS-ODRAW 2.3.7.43: each Boolean applies only when its fUse bit is set;
    // otherwise its default (fFilled and fillShape true, fillUseRect false).
    let bit = |value: u32, flag: u32, used: u32, default: bool| -> u32 {
        if value & used != 0 {
            u32::from(value & flag != 0)
        } else {
            u32::from(default)
        }
    };
    let flags = boolean.unwrap_or(0);
    for (opid, value) in [
        (0x1bbu32, bit(flags, 1 << 4, 1 << 20, true)),
        (0x1bd, bit(flags, 1 << 2, 1 << 18, true)),
        (0x1be, bit(flags, 1 << 1, 1 << 17, false)),
    ] {
        output.extend_from_slice(&opid.to_le_bytes());
        output.extend_from_slice(&value.to_le_bytes());
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn crc_matches_the_shared_office_algorithm() {
        // Frame of an Excel-saved chart: automatic-color hairline with no line
        // pattern (lns 5) and a solid white chart area. The stored checksum
        // is 0x03AC5D6E.
        let line = [0, 0, 0, 0, 5, 0, 0xff, 0xff, 8, 0, 0x4d, 0];
        let area = [0xff, 0xff, 0xff, 0, 0, 0, 0, 0, 1, 0, 0, 0, 9, 0, 0x4d, 0];
        let mut data = line_properties(&line).unwrap().to_vec();
        data.extend(interior_properties(&area, &area).unwrap());
        assert_eq!(data[1], 0xff);
        assert_eq!(crc(&data), 0x03ac_5d6e);
    }
}
