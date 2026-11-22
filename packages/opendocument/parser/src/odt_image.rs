//! Pictures linked or embedded in a document: media type sniffing and the
//! intrinsic size a crop (§20.172 `fo:clip`) is measured against. Decoding
//! (and its pixel budget) belongs to the shared renderer, which draws
//! raster formats, SVG and the Windows metafiles; formats it cannot draw are
//! left out of the model.

/// The media type of picture bytes by their signature.
pub(super) fn sniff(bytes: &[u8]) -> Option<&'static str> {
    if bytes.starts_with(b"\x89PNG\r\n\x1a\n") {
        return Some("image/png");
    }
    if bytes.starts_with(&[0xFF, 0xD8, 0xFF]) {
        return Some("image/jpeg");
    }
    if bytes.starts_with(b"GIF87a") || bytes.starts_with(b"GIF89a") {
        return Some("image/gif");
    }
    if bytes.starts_with(b"BM") {
        return Some("image/bmp");
    }
    if bytes.starts_with(b"II*\0") || bytes.starts_with(b"MM\0*") {
        return Some("image/tiff");
    }
    if bytes.len() >= 12 && &bytes[0..4] == b"RIFF" && &bytes[8..12] == b"WEBP" {
        return Some("image/webp");
    }
    // The metafile types are the names the shared renderer keys metafile
    // handling on (core `isMetafileMime`: crop frame, text font demand), as
    // the presentation reader and the OOXML readers spell them.
    // EMF: an EMR_HEADER record (type 1) with the " EMF" signature at 40.
    if bytes.len() >= 44 && bytes[0..4] == [1, 0, 0, 0] && &bytes[40..44] == b" EMF" {
        return Some("image/emf");
    }
    // WMF: the placeable header key, or a standard header.
    if bytes.starts_with(&[0xD7, 0xCD, 0xC6, 0x9A])
        || (bytes.len() >= 4 && matches!(bytes[0..4], [1, 0, 9, 0] | [2, 0, 9, 0]))
    {
        return Some("image/wmf");
    }
    // SVG (possibly after an XML declaration, comments or a BOM).
    let head = &bytes[..bytes.len().min(4096)];
    if let Ok(text) = std::str::from_utf8(head) {
        if text.contains("<svg") {
            return Some("image/svg+xml");
        }
    } else if String::from_utf8_lossy(head).contains("<svg") {
        return Some("image/svg+xml");
    }
    None
}

/// Intrinsic picture size in points (pixels at their resolution, 96 dpi by
/// default), from the PNG, JPEG or GIF header, or a metafile's natural size
/// (`odf_props::metafile_size_pt`).
pub(super) fn intrinsic_size(bytes: &[u8]) -> Option<(f64, f64)> {
    if let Some(size) = crate::odf_props::metafile_size_pt(bytes) {
        return Some(size);
    }
    let be16 = |at: usize| -> Option<u32> {
        Some(u32::from(*bytes.get(at)?) << 8 | u32::from(*bytes.get(at + 1)?))
    };
    let be32 = |at: usize| -> Option<u32> { Some(be16(at)? << 16 | be16(at + 2)?) };
    let pt = |pixels: u32, dpi: f64| f64::from(pixels) * 72.0 / dpi;
    if bytes.starts_with(b"\x89PNG\r\n\x1a\n") {
        let width = be32(16)?;
        let height = be32(20)?;
        // pHYs: pixels per metre.
        let mut dpi = (96.0, 96.0);
        let mut at = 8usize;
        while at + 8 <= bytes.len() {
            let length = be32(at)? as usize;
            let kind = bytes.get(at + 4..at + 8)?;
            if kind == b"pHYs" && length >= 9 && bytes.get(at + 16) == Some(&1) {
                let x = f64::from(be32(at + 8)?) * 0.0254;
                let y = f64::from(be32(at + 12)?) * 0.0254;
                if x > 1.0 && y > 1.0 {
                    dpi = (x, y);
                }
                break;
            }
            if kind == b"IDAT" {
                break;
            }
            // Chunk = length, type, data, CRC: the next chunk starts past
            // all of them; a length running past the data ends the scan
            // (checked on every target so nothing wraps on 32-bit hosts).
            match length.checked_add(12).and_then(|step| at.checked_add(step)) {
                Some(next) if next <= bytes.len() => at = next,
                _ => break,
            }
        }
        return Some((pt(width, dpi.0), pt(height, dpi.1)));
    }
    if bytes.starts_with(b"GIF8") {
        let width = u32::from(*bytes.get(6)?) | u32::from(*bytes.get(7)?) << 8;
        let height = u32::from(*bytes.get(8)?) | u32::from(*bytes.get(9)?) << 8;
        return Some((pt(width, 96.0), pt(height, 96.0)));
    }
    if bytes.starts_with(&[0xFF, 0xD8]) {
        let mut at = 2usize;
        let mut dpi = (96.0, 96.0);
        while at + 4 <= bytes.len() {
            if bytes[at] != 0xFF {
                return None;
            }
            let marker = bytes[at + 1];
            if marker == 0xD8 || marker == 0x01 || (0xD0..=0xD7).contains(&marker) {
                at += 2;
                continue;
            }
            let length = be16(at + 2)? as usize;
            if marker == 0xE0 && bytes.get(at + 4..at + 9) == Some(b"JFIF\0") {
                let unit = *bytes.get(at + 11)?;
                let x = f64::from(be16(at + 12)?);
                let y = f64::from(be16(at + 14)?);
                if x > 0.0 && y > 0.0 {
                    match unit {
                        1 => dpi = (x, y),
                        2 => dpi = (x * 2.54, y * 2.54),
                        _ => {}
                    }
                }
            }
            if matches!(marker, 0xC0..=0xC3 | 0xC5..=0xC7 | 0xC9..=0xCB | 0xCD..=0xCF) {
                let height = be16(at + 5)?;
                let width = be16(at + 7)?;
                return Some((pt(width, dpi.0), pt(height, dpi.1)));
            }
            at = at.checked_add(2 + length)?;
        }
    }
    None
}

/// Standard base64 (RFC 4648) with white space ignored, as
/// `office:binary-data` holds it.
pub(super) fn base64(data: &str) -> Option<Vec<u8>> {
    let mut out = Vec::with_capacity(data.len() / 4 * 3);
    let mut buffer = 0u32;
    let mut bits = 0u32;
    for byte in data.bytes() {
        let value = match byte {
            b'A'..=b'Z' => byte - b'A',
            b'a'..=b'z' => byte - b'a' + 26,
            b'0'..=b'9' => byte - b'0' + 52,
            b'+' => 62,
            b'/' => 63,
            b'=' => break,
            b' ' | b'\n' | b'\r' | b'\t' => continue,
            _ => return None,
        };
        buffer = buffer << 6 | u32::from(value);
        bits += 6;
        if bits >= 8 {
            bits -= 8;
            out.push((buffer >> bits) as u8);
            buffer &= (1 << bits) - 1;
        }
    }
    Some(out)
}

/// A minimal PNG header (64 by 32 pixels) for tests.
#[cfg(test)]
pub(super) fn tests_png() -> Vec<u8> {
    let mut png = b"\x89PNG\r\n\x1a\n\0\0\0\rIHDR".to_vec();
    png.extend_from_slice(&[0, 0, 0, 64, 0, 0, 0, 32, 8, 2, 0, 0, 0]);
    png
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn sniffs_formats_and_sizes() {
        let mut png = b"\x89PNG\r\n\x1a\n\0\0\0\rIHDR".to_vec();
        png.extend_from_slice(&[0, 0, 0, 96, 0, 0, 0, 48, 8, 2, 0, 0, 0]);
        assert_eq!(sniff(&png), Some("image/png"));
        assert_eq!(intrinsic_size(&png), Some((72.0, 36.0)));
        assert_eq!(
            sniff(b"<?xml version=\"1.0\"?><svg/>"),
            Some("image/svg+xml")
        );
        assert_eq!(sniff(b"plain"), None);
        assert_eq!(base64("aGk=").unwrap(), b"hi");
        assert_eq!(base64("aG\nk=").unwrap(), b"hi");
    }

    #[test]
    fn png_chunk_lengths_past_the_data_end_the_scan() {
        // IHDR whose length field claims 0xfffffff4 bytes: on a 32-bit host
        // `12 + length` wrapped to 0 and the scan never advanced.
        let mut png = b"\x89PNG\r\n\x1a\n\xff\xff\xff\xf4IHDR".to_vec();
        png.extend_from_slice(&[0, 0, 0, 96, 0, 0, 0, 48, 8, 2, 0, 0, 0]);
        assert_eq!(intrinsic_size(&png), Some((72.0, 36.0)));
    }
}
