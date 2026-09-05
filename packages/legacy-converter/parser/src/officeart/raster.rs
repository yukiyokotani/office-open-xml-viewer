//! Passive raster BLIP validation shared by binary Office converters.
//! MS-ODRAW 2.2.24-32; W3C PNG IHDR; ITU-T T.81 JPEG frame headers.
use super::Record;

/// Resolve only in-stream BLIPs. `delayed` is the format-defined binary stream,
/// never a file path; DOC inline shapes do not supply a delayed store.
pub(crate) fn read_store_entry<'a>(
    entry: Record<'a>,
    delayed: Option<&'a [u8]>,
    budget: &mut usize,
) -> Result<Option<Image<'a>>, String> {
    if entry.kind != 0xf007 {
        return read(entry, budget);
    }
    let b = entry.payload;
    if entry.version != 2
        || b.len() < 36
        || (entry.instance != u16::from(b[0]) && entry.instance != u16::from(b[1]))
    {
        return Err(unsupported("invalid OfficeArt BLIP store entry"));
    }
    let name = usize::from(b[33]);
    if name % 2 != 0 || 36 + name > b.len() {
        return Err(unsupported("invalid OfficeArt BLIP name length"));
    }
    let number = |offset| u32::from_le_bytes(b[offset..offset + 4].try_into().unwrap()) as usize;
    let size = number(20);
    if number(24) == 0 {
        return Ok(None);
    }
    let source = if 36 + name < b.len() {
        b.get(36 + name..)
            .filter(|s| s.len() == size)
            .ok_or_else(|| unsupported("OfficeArt embedded BLIP size mismatch"))?
    } else {
        let Some(delayed) = delayed else {
            return Ok(None);
        };
        delayed
            .get(number(28)..)
            .and_then(|s| s.get(..size))
            .ok_or_else(|| unsupported("OfficeArt delayed BLIP range out of bounds"))?
    };
    let (blip, end) = super::record_with_end(source, 0, budget, "OfficeArt")?;
    if end != source.len() {
        return Err(unsupported("OfficeArt BLIP record size mismatch"));
    }
    read(blip, budget)
}

const MAX_PIXELS: u64 = 40_000_000;

#[derive(Clone, Copy)]
pub(crate) struct Image<'a> {
    pub bytes: &'a [u8],
    pub extension: &'static str,
}
fn unsupported(message: impl Into<String>) -> String {
    format!("UNSUPPORTED:{}", message.into())
}

pub(crate) fn read<'a>(blip: Record<'a>, budget: &mut usize) -> Result<Option<Image<'a>>, String> {
    let (extension, prefix) = match (blip.kind, blip.instance) {
        (0xf01e, 0x6e0) => ("png", 17),
        (0xf01e, 0x6e1) => ("png", 33),
        (0xf01d, 0x46a | 0x6e2) => ("jpg", 17),
        (0xf01d, 0x46b | 0x6e3) => ("jpg", 33),
        (0xf01d | 0xf01e, _) => return Err(unsupported("invalid OfficeArt raster BLIP instance")),
        _ => return Ok(None), // No WMF/EMF/PICT/DIB or active-object decoding here.
    };
    if blip.version != 0 {
        return Err(unsupported("invalid OfficeArt raster BLIP version"));
    }
    let bytes = blip
        .payload
        .get(prefix..)
        .ok_or_else(|| unsupported("truncated OfficeArt raster BLIP"))?;
    // Admit only the advertised raster encoding. Some producer output puts a
    // different format in a PNG BLIP; omit it rather than relabel, sniff into
    // another decoder, or copy arbitrary bytes into a supported image part.
    if (extension == "png" && !bytes.starts_with(b"\x89PNG\r\n\x1a\n"))
        || (extension == "jpg" && !bytes.starts_with(&[0xff, 0xd8]))
    {
        return Ok(None);
    }
    let (width, height) = if extension == "png" {
        png_size(bytes)?
    } else {
        jpeg_size(bytes, budget)?
    };
    if width == 0
        || height == 0
        || width > 32768
        || height > 32768
        || u64::from(width) * u64::from(height) > MAX_PIXELS
    {
        return Err(unsupported(
            "OfficeArt image dimensions exceed resource limit",
        ));
    }
    Ok(Some(Image { bytes, extension }))
}
pub(crate) fn png_size(b: &[u8]) -> Result<(u32, u32), String> {
    if b.len() < 33 || !b.starts_with(b"\x89PNG\r\n\x1a\n\0\0\0\x0dIHDR") {
        return Err(unsupported("invalid OfficeArt PNG header"));
    }
    // PNG IHDR (W3C PNG section 11.2.2). This is a bounded header check,
    // not a replacement for the ordinary renderer's image decoder.
    let valid_depth = match b[25] {
        0 => matches!(b[24], 1 | 2 | 4 | 8 | 16),
        2 | 4 | 6 => matches!(b[24], 8 | 16),
        3 => matches!(b[24], 1 | 2 | 4 | 8),
        _ => false,
    };
    if !valid_depth || b[26] != 0 || b[27] != 0 || b[28] > 1 {
        return Err(unsupported("invalid OfficeArt PNG IHDR"));
    }
    Ok((
        u32::from_be_bytes(b[16..20].try_into().unwrap()),
        u32::from_be_bytes(b[20..24].try_into().unwrap()),
    ))
}
pub(crate) fn jpeg_size(b: &[u8], budget: &mut usize) -> Result<(u32, u32), String> {
    if !b.starts_with(&[0xff, 0xd8]) {
        return Err(unsupported("invalid OfficeArt JPEG header"));
    }
    let mut position = 2;
    while position < b.len() {
        *budget = budget
            .checked_sub(1)
            .ok_or_else(|| unsupported("OfficeArt JPEG marker work budget exceeded"))?;
        if b[position] != 0xff {
            return Err(unsupported("invalid OfficeArt JPEG marker"));
        }
        position += 1;
        while b.get(position) == Some(&0xff) {
            *budget = budget
                .checked_sub(1)
                .ok_or_else(|| unsupported("OfficeArt JPEG marker work budget exceeded"))?;
            position += 1;
        }
        let marker = *b
            .get(position)
            .ok_or_else(|| unsupported("truncated OfficeArt JPEG marker"))?;
        position += 1;
        if matches!(marker, 0xda | 0xd9) {
            break;
        }
        if marker == 0x01 {
            // Standalone TEM marker (ITU-T T.81, B.1.1.3).
            continue;
        }
        if matches!(marker, 0 | 0xd0..=0xd8) {
            return Err(unsupported("unexpected OfficeArt JPEG marker before frame"));
        }
        // Do not skip an unsupported first frame and accidentally validate the
        // dimensions of a later frame that a decoder would not choose.
        if matches!(marker, 0xc3 | 0xc5..=0xc7 | 0xc9..=0xcb | 0xcd..=0xcf | 0xde) {
            return Err(unsupported("unsupported OfficeArt JPEG frame encoding"));
        }
        let length = b
            .get(position..position + 2)
            .map(|b| u16::from_be_bytes([b[0], b[1]]) as usize)
            .ok_or_else(|| unsupported("truncated OfficeArt JPEG segment"))?;
        let segment = b
            .get(position..)
            .and_then(|b| b.get(..length))
            .filter(|b| b.len() >= 2)
            .ok_or_else(|| unsupported("invalid OfficeArt JPEG segment length"))?;
        if matches!(marker, 0xc0..=0xc2) {
            // ITU-T T.81 B.2.2: Lf = 8 + 3 * Nf. Only 8-bit Huffman
            // baseline/sequential/progressive frames are supported here.
            if segment.len() < 8
                || segment[2] != 8
                || segment[7] == 0
                || segment.len() != 8 + 3 * usize::from(segment[7])
            {
                return Err(unsupported("unsupported OfficeArt JPEG frame"));
            }
            return Ok((
                u32::from(u16::from_be_bytes([segment[5], segment[6]])),
                u32::from(u16::from_be_bytes([segment[3], segment[4]])),
            ));
        }
        position += length;
    }
    Err(unsupported("OfficeArt JPEG lacks a supported frame"))
}
