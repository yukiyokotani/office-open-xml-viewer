//! Passive image BLIP validation shared by binary Office converters.
//! MS-ODRAW 2.2.24-32; W3C PNG IHDR; ITU-T T.81 JPEG frame headers.
use super::{ByteSpan, Record, RecordSpan};
use std::borrow::Cow;
use std::ops::Range;

/// Resolve only in-stream BLIPs. `delayed` is the format-defined binary stream,
/// never a file path; DOC inline shapes do not supply a delayed store.
pub(crate) fn read_store_entry<'a>(
    entry: Record<'a>,
    delayed: Option<&'a [u8]>,
    budget: &mut usize,
    remaining_bytes: usize,
) -> Result<Option<Image<'a>>, String> {
    if entry.kind != 0xf007 {
        return read(entry, budget, remaining_bytes);
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
    read(blip, budget, remaining_bytes)
}

const MAX_PIXELS: u64 = 40_000_000;

pub(crate) struct Image<'a> {
    pub bytes: Cow<'a, [u8]>,
    pub extension: &'static str,
}

#[derive(Debug, PartialEq, Eq)]
pub(crate) enum ImageSpanBytes {
    Source(ByteSpan),
    Owned(Vec<u8>),
}

#[derive(Debug, PartialEq, Eq)]
pub(crate) struct ImageSpan {
    pub bytes: ImageSpanBytes,
    pub extension: &'static str,
}

pub(super) enum DecodedBytes {
    Source(Range<usize>),
    Owned(Vec<u8>),
}

impl DecodedBytes {
    pub(super) fn view<'a>(&'a self, payload: &'a [u8]) -> &'a [u8] {
        match self {
            Self::Source(range) => &payload[range.clone()],
            Self::Owned(bytes) => bytes,
        }
    }
}

struct DecodedImage {
    bytes: DecodedBytes,
    extension: &'static str,
}
fn unsupported(message: impl Into<String>) -> String {
    format!("UNSUPPORTED:{}", message.into())
}

pub(crate) fn read<'a>(
    blip: Record<'a>,
    budget: &mut usize,
    remaining_bytes: usize,
) -> Result<Option<Image<'a>>, String> {
    Ok(decode(blip, budget, remaining_bytes)?.map(|image| Image {
        bytes: match image.bytes {
            DecodedBytes::Source(range) => Cow::Borrowed(&blip.payload[range]),
            DecodedBytes::Owned(bytes) => Cow::Owned(bytes),
        },
        extension: image.extension,
    }))
}

pub(crate) fn read_span(
    blip: &RecordSpan,
    backing: &[u8],
    budget: &mut usize,
    remaining_bytes: usize,
) -> Result<Option<ImageSpan>, String> {
    let viewed = blip.view(backing)?;
    decode(viewed, budget, remaining_bytes)?
        .map(|image| {
            Ok(ImageSpan {
                bytes: match image.bytes {
                    DecodedBytes::Source(range) => ImageSpanBytes::Source(
                        blip.payload_span()
                            .checked_subrange(range, "OfficeArt BLIP payload")?,
                    ),
                    DecodedBytes::Owned(bytes) => ImageSpanBytes::Owned(bytes),
                },
                extension: image.extension,
            })
        })
        .transpose()
}

fn decode(
    blip: Record<'_>,
    budget: &mut usize,
    remaining_bytes: usize,
) -> Result<Option<DecodedImage>, String> {
    if matches!(blip.kind, 0xf01a | 0xf01b) {
        let extension = if blip.kind == 0xf01a { "emf" } else { "wmf" };
        return Ok(super::metafile::decode(blip, budget, remaining_bytes)?
            .map(|bytes| DecodedImage { bytes, extension }));
    }
    let (extension, prefix) = match (blip.kind, blip.instance) {
        (0xf01e, 0x6e0) => ("png", 17),
        (0xf01e, 0x6e1) => ("png", 33),
        (0xf01d, 0x46a | 0x6e2) => ("jpg", 17),
        (0xf01d, 0x46b | 0x6e3) => ("jpg", 33),
        (0xf01d | 0xf01e, _) => return Err(unsupported("invalid OfficeArt raster BLIP instance")),
        _ => return Ok(None), // No PICT/DIB or active-object decoding here.
    };
    if blip.version != 0 {
        return Err(unsupported("invalid OfficeArt raster BLIP version"));
    }
    let range = prefix..blip.payload.len();
    let bytes = blip
        .payload
        .get(range.clone())
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
    if bytes.len() > remaining_bytes {
        return Err(unsupported("OfficeArt retained media budget exceeded"));
    }
    Ok(Some(DecodedImage {
        bytes: DecodedBytes::Source(range),
        extension,
    }))
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::officeart::{emf_test_blip, record_span_with_end, wmf_test_blip};

    fn record(options: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
        [
            options.to_le_bytes().as_slice(),
            kind.to_le_bytes().as_slice(),
            (payload.len() as u32).to_le_bytes().as_slice(),
            payload,
        ]
        .concat()
    }

    fn png_blip() -> Vec<u8> {
        let mut png = vec![0; 33];
        png[..16].copy_from_slice(b"\x89PNG\r\n\x1a\n\0\0\0\x0dIHDR");
        png[19] = 2;
        png[23] = 3;
        png[24] = 8;
        png[25] = 2;
        record(0x6e00, 0xf01e, &[vec![0; 17], png].concat())
    }

    fn jpeg_blip() -> Vec<u8> {
        let jpeg = [0xff, 0xd8, 0xff, 0xc0, 0, 11, 8, 0, 3, 0, 2, 1, 1, 0x11, 0];
        record(0x46a0, 0xf01d, &[vec![0; 17], jpeg.to_vec()].concat())
    }

    fn uncompress_metafile_record(mut encoded: Vec<u8>, source: &[u8]) -> Vec<u8> {
        let payload_start = 8;
        let header_start = payload_start + 16;
        encoded.truncate(header_start + 34);
        encoded[header_start + 28..header_start + 32]
            .copy_from_slice(&(source.len() as u32).to_le_bytes());
        encoded[header_start + 32] = 0xfe;
        encoded.extend_from_slice(source);
        let payload_len = (encoded.len() - payload_start) as u32;
        encoded[4..8].copy_from_slice(&payload_len.to_le_bytes());
        encoded
    }

    fn assert_parity(encoded: Vec<u8>, expected: &[u8], expected_extension: &str, owned: bool) {
        let (span, _) = record_span_with_end(&encoded, 0, &mut 1, "BLIP").unwrap();
        let mut borrowed_budget = 20;
        let borrowed = read(
            span.view(&encoded).unwrap(),
            &mut borrowed_budget,
            expected.len(),
        )
        .unwrap()
        .unwrap();
        let mut span_budget = 20;
        let spanned = read_span(&span, &encoded, &mut span_budget, expected.len())
            .unwrap()
            .unwrap();
        assert_eq!(borrowed.bytes.as_ref(), expected);
        assert_eq!(borrowed.extension, expected_extension);
        assert_eq!(spanned.extension, expected_extension);
        assert_eq!(span_budget, borrowed_budget);
        match spanned.bytes {
            ImageSpanBytes::Source(source) => {
                assert!(!owned);
                let moved = encoded;
                assert_eq!(source.view(&moved).unwrap(), expected);
            }
            ImageSpanBytes::Owned(bytes) => {
                assert!(owned);
                assert_eq!(bytes, expected);
            }
        }
    }

    #[test]
    fn borrowed_and_spanned_direct_blips_have_matching_bytes_and_budgets() {
        let png = png_blip();
        let png_expected = png[8 + 17..].to_vec();
        assert_parity(png, &png_expected, "png", false);
        let jpeg = jpeg_blip();
        let jpeg_expected = jpeg[8 + 17..].to_vec();
        assert_parity(jpeg, &jpeg_expected, "jpg", false);

        for (extension, fixture) in [("emf", emf_test_blip()), ("wmf", wmf_test_blip())] {
            let (source, compressed) = fixture;
            assert_parity(compressed.clone(), &source, extension, true);
            assert_parity(
                uncompress_metafile_record(compressed, &source),
                &source,
                extension,
                false,
            );
        }
    }

    #[test]
    fn span_reader_preserves_limits_omission_and_backing_bounds() {
        let encoded = png_blip();
        let (span, _) = record_span_with_end(&encoded, 0, &mut 1, "BLIP").unwrap();
        assert!(read_span(&span, &encoded, &mut 10, 32).is_err());
        assert!(read_span(&span, &encoded[..encoded.len() - 1], &mut 10, usize::MAX).is_err());

        let mut mismatched = encoded;
        mismatched[8 + 17] = 0;
        let (span, _) = record_span_with_end(&mismatched, 0, &mut 1, "BLIP").unwrap();
        let mut borrowed_budget = 10;
        let borrowed = read(
            span.view(&mismatched).unwrap(),
            &mut borrowed_budget,
            usize::MAX,
        );
        let mut span_budget = 10;
        let owned = read_span(&span, &mismatched, &mut span_budget, usize::MAX);
        assert!(borrowed.unwrap().is_none());
        assert!(owned.unwrap().is_none());
        assert_eq!(span_budget, borrowed_budget);
    }
}
