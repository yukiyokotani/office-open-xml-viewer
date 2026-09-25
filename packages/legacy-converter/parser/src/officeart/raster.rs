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
    read_store_entry_as(entry, delayed, budget, remaining_bytes, Raster::Advertised)
}

/// Which bytes a PNG/JPEG BLIP may carry. The content-signature variants are
/// host decisions backed by that host's own output: Word writes TIFF data
/// inside PNG BLIPs and reads it back as TIFF (see the DOC picture store), and
/// PowerPoint displays GIF data stored in PNG BLIPs (its PDF export shows the
/// GIF image). No host has evidence for the other combination.
#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) enum Raster {
    Advertised,
    TiffAware,
    GifAware,
}

pub(crate) fn read_store_entry_as<'a>(
    entry: Record<'a>,
    delayed: Option<&'a [u8]>,
    budget: &mut usize,
    remaining_bytes: usize,
    raster: Raster,
) -> Result<Option<Image<'a>>, String> {
    let read = |blip, budget: &mut usize| read_as(blip, budget, remaining_bytes, raster);
    let source = match locate_store_entry(entry)? {
        StoreLocation::Direct => return read(entry, budget),
        StoreLocation::Omit => return Ok(None),
        StoreLocation::Embedded(range) => &entry.payload[range],
        StoreLocation::Delayed { offset, size } => {
            let Some(delayed) = delayed else {
                return Ok(None);
            };
            delayed
                .get(offset..)
                .and_then(|bytes| bytes.get(..size))
                .ok_or_else(|| unsupported("OfficeArt delayed BLIP range out of bounds"))?
        }
    };
    let (blip, end) = super::record_with_end(source, 0, budget, "OfficeArt")?;
    if end != source.len() {
        return Err(unsupported("OfficeArt BLIP record size mismatch"));
    }
    read(blip, budget)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum StoreBacking {
    Primary,
    Delayed,
}

#[derive(Debug, PartialEq, Eq)]
pub(crate) struct StoreImageSpan {
    pub image: ImageSpan,
    pub backing: StoreBacking,
}

impl StoreImageSpan {
    pub(crate) fn view<'a>(
        &'a self,
        primary: &'a [u8],
        delayed: Option<&'a [u8]>,
    ) -> Result<&'a [u8], String> {
        match &self.image.bytes {
            ImageSpanBytes::Owned(bytes) => Ok(bytes),
            ImageSpanBytes::Source(span) => {
                match self.backing {
                    StoreBacking::Primary => span.view(primary),
                    StoreBacking::Delayed => span.view(delayed.ok_or_else(|| {
                        unsupported("OfficeArt delayed BLIP backing is unavailable")
                    })?),
                }
            }
        }
    }
}

pub(crate) fn read_store_entry_span(
    entry: &RecordSpan,
    primary: &[u8],
    delayed: Option<&[u8]>,
    budget: &mut usize,
    remaining_bytes: usize,
) -> Result<Option<StoreImageSpan>, String> {
    read_store_entry_span_as(
        entry,
        primary,
        delayed,
        budget,
        remaining_bytes,
        Raster::Advertised,
    )
}

pub(crate) fn read_store_entry_span_as(
    entry: &RecordSpan,
    primary: &[u8],
    delayed: Option<&[u8]>,
    budget: &mut usize,
    remaining_bytes: usize,
    raster: Raster,
) -> Result<Option<StoreImageSpan>, String> {
    let viewed = entry.view(primary)?;
    let (source, backing, bytes) = match locate_store_entry(viewed)? {
        StoreLocation::Direct => {
            return Ok(
                read_span_as(entry, primary, budget, remaining_bytes, raster)?.map(|image| {
                    StoreImageSpan {
                        image,
                        backing: StoreBacking::Primary,
                    }
                }),
            )
        }
        StoreLocation::Omit => return Ok(None),
        StoreLocation::Embedded(range) => (
            entry
                .payload_span()
                .checked_subrange(range, "OfficeArt embedded BLIP")?,
            StoreBacking::Primary,
            primary,
        ),
        StoreLocation::Delayed { offset, size } => {
            let Some(delayed) = delayed else {
                return Ok(None);
            };
            let end = offset
                .checked_add(size)
                .ok_or_else(|| unsupported("OfficeArt delayed BLIP range out of bounds"))?;
            (
                ByteSpan::new(offset..end, delayed.len(), "OfficeArt delayed BLIP")?,
                StoreBacking::Delayed,
                delayed,
            )
        }
    };
    let offset = source.range().start;
    let (blip, end) = super::record_span_with_end_in(bytes, &source, offset, budget, "OfficeArt")?;
    if end != source.range().end {
        return Err(unsupported("OfficeArt BLIP record size mismatch"));
    }
    Ok(read_span_as(&blip, bytes, budget, remaining_bytes, raster)?
        .map(|image| StoreImageSpan { image, backing }))
}

enum StoreLocation {
    Direct,
    Omit,
    Embedded(Range<usize>),
    Delayed { offset: usize, size: usize },
}

fn locate_store_entry(entry: Record<'_>) -> Result<StoreLocation, String> {
    if entry.kind != 0xf007 {
        return Ok(StoreLocation::Direct);
    }
    let b = entry.payload;
    if entry.version != 2
        || b.len() < 36
        || (entry.instance != u16::from(b[0]) && entry.instance != u16::from(b[1]))
    {
        return Err(unsupported("invalid OfficeArt BLIP store entry"));
    }
    let name = usize::from(b[33]);
    let inline_start = 36 + name;
    if name % 2 != 0 || inline_start > b.len() {
        return Err(unsupported("invalid OfficeArt BLIP name length"));
    }
    let number = |offset| u32::from_le_bytes(b[offset..offset + 4].try_into().unwrap()) as usize;
    let size = number(20);
    if number(24) == 0 {
        return Ok(StoreLocation::Omit);
    }
    if inline_start < b.len() {
        if b.len() - inline_start != size {
            return Err(unsupported("OfficeArt embedded BLIP size mismatch"));
        }
        Ok(StoreLocation::Embedded(inline_start..b.len()))
    } else {
        Ok(StoreLocation::Delayed {
            offset: number(28),
            size,
        })
    }
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
    read_as(blip, budget, remaining_bytes, Raster::Advertised)
}

fn read_as<'a>(
    blip: Record<'a>,
    budget: &mut usize,
    remaining_bytes: usize,
    raster: Raster,
) -> Result<Option<Image<'a>>, String> {
    Ok(
        decode(blip, budget, remaining_bytes, raster)?.map(|image| Image {
            bytes: match image.bytes {
                DecodedBytes::Source(range) => Cow::Borrowed(&blip.payload[range]),
                DecodedBytes::Owned(bytes) => Cow::Owned(bytes),
            },
            extension: image.extension,
        }),
    )
}

pub(crate) fn read_span(
    blip: &RecordSpan,
    backing: &[u8],
    budget: &mut usize,
    remaining_bytes: usize,
) -> Result<Option<ImageSpan>, String> {
    read_span_as(blip, backing, budget, remaining_bytes, Raster::Advertised)
}

fn read_span_as(
    blip: &RecordSpan,
    backing: &[u8],
    budget: &mut usize,
    remaining_bytes: usize,
    raster: Raster,
) -> Result<Option<ImageSpan>, String> {
    let viewed = blip.view(backing)?;
    decode(viewed, budget, remaining_bytes, raster)?
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
    raster: Raster,
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
    let mut extension = extension;
    if (extension == "png" && !bytes.starts_with(b"\x89PNG\r\n\x1a\n"))
        || (extension == "jpg" && !bytes.starts_with(&[0xff, 0xd8]))
    {
        extension = match raster {
            Raster::TiffAware if bytes.starts_with(b"II*\0") || bytes.starts_with(b"MM\0*") => {
                "tiff"
            }
            Raster::GifAware
                if extension == "png"
                    && (bytes.starts_with(b"GIF87a") || bytes.starts_with(b"GIF89a")) =>
            {
                "gif"
            }
            _ => return Ok(None),
        };
    }
    let (width, height) = match extension {
        "png" => png_size(bytes)?,
        "jpg" => jpeg_size(bytes, budget)?,
        "gif" => gif_size(bytes)?,
        _ => tiff_size(bytes, budget)?,
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
/// GIF87a/89a logical screen descriptor (GIF89a specification section 18):
/// the signature and version, then the 16-bit little-endian logical screen
/// width and height. A header check only; the ordinary image decoder remains
/// responsible for the data stream.
pub(crate) fn gif_size(b: &[u8]) -> Result<(u32, u32), String> {
    if b.len() < 13 || !(b.starts_with(b"GIF87a") || b.starts_with(b"GIF89a")) {
        return Err(unsupported("invalid OfficeArt GIF header"));
    }
    Ok((
        u32::from(u16::from_le_bytes([b[6], b[7]])),
        u32::from(u16::from_le_bytes([b[8], b[9]])),
    ))
}

/// Bounded TIFF 6.0 header check: byte order, the magic number, the first
/// IFD inside the data and its ImageWidth (256) and ImageLength (257) tags
/// with SHORT or LONG single values. Decoding stays with the renderer.
pub(crate) fn tiff_size(b: &[u8], budget: &mut usize) -> Result<(u32, u32), String> {
    let invalid = || unsupported("invalid OfficeArt TIFF header");
    let little = match b.get(..4) {
        Some(b"II*\0") => true,
        Some(b"MM\0*") => false,
        _ => return Err(invalid()),
    };
    let u16_at = |at: usize| -> Result<u32, String> {
        let bytes: [u8; 2] = b.get(at..at + 2).ok_or_else(invalid)?.try_into().unwrap();
        Ok(u32::from(if little {
            u16::from_le_bytes(bytes)
        } else {
            u16::from_be_bytes(bytes)
        }))
    };
    let u32_at = |at: usize| -> Result<u32, String> {
        let bytes: [u8; 4] = b.get(at..at + 4).ok_or_else(invalid)?.try_into().unwrap();
        Ok(if little {
            u32::from_le_bytes(bytes)
        } else {
            u32::from_be_bytes(bytes)
        })
    };
    let ifd = u32_at(4)? as usize;
    if ifd < 8 {
        return Err(invalid());
    }
    let count = u16_at(ifd)? as usize;
    *budget = budget
        .checked_sub(count)
        .ok_or_else(|| unsupported("OfficeArt TIFF work budget exceeded"))?;
    let mut size = [None, None];
    for index in 0..count {
        let entry = ifd + 2 + index * 12;
        let tag = u16_at(entry)?;
        if !(256..=257).contains(&tag) {
            continue;
        }
        if u32_at(entry + 4)? != 1 {
            return Err(invalid());
        }
        let value = match u16_at(entry + 2)? {
            3 => u16_at(entry + 8)?,
            4 => u32_at(entry + 8)?,
            _ => return Err(invalid()),
        };
        size[(tag - 256) as usize] = Some(value);
    }
    match size {
        [Some(width), Some(height)] => Ok((width, height)),
        _ => Err(invalid()),
    }
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

    /// A PNG BLIP whose data is a little- or big-endian TIFF with the given
    /// ImageWidth/ImageLength entries.
    fn tiff_in_png(little: bool, entries: &[(u16, u16, u32)]) -> Vec<u8> {
        let u16b = |v: u16| {
            if little {
                v.to_le_bytes()
            } else {
                v.to_be_bytes()
            }
        };
        let u32b = |v: u32| {
            if little {
                v.to_le_bytes()
            } else {
                v.to_be_bytes()
            }
        };
        let mut tiff = if little {
            b"II*\0".to_vec()
        } else {
            b"MM\0*".to_vec()
        };
        tiff.extend(u32b(8));
        tiff.extend(u16b(entries.len() as u16));
        for (tag, kind, value) in entries {
            tiff.extend(u16b(*tag));
            tiff.extend(u16b(*kind));
            tiff.extend(u32b(1));
            if *kind == 3 {
                tiff.extend(u16b(*value as u16));
                tiff.extend([0, 0]);
            } else {
                tiff.extend(u32b(*value));
            }
        }
        tiff.extend([0; 4]);
        record(0x6e00, 0xf01e, &[vec![0; 17], tiff].concat())
    }

    #[test]
    fn gif_data_in_png_blips_is_admitted_only_for_the_gif_aware_host() {
        let gif = |sig: &[u8], w: u16, h: u16| {
            let mut data = sig.to_vec();
            data.extend(w.to_le_bytes());
            data.extend(h.to_le_bytes());
            data.extend([0x80, 0, 0, 0, 0, 0, 0xff, 0xff, 0xff, 0x3b]);
            data
        };
        let blip = |data: Vec<u8>, kind: u16, options: u16| {
            let payload = [vec![0; 17], data].concat();
            [
                options.to_le_bytes().as_slice(),
                kind.to_le_bytes().as_slice(),
                (payload.len() as u32).to_le_bytes().as_slice(),
                &payload,
            ]
            .concat()
        };
        for sig in [b"GIF87a", b"GIF89a"] {
            let bytes = blip(gif(sig, 40, 30), 0xf01e, 0x6e00);
            let (record, _) = crate::officeart::record_with_end(&bytes, 0, &mut 10, "t").unwrap();
            assert!(read(record, &mut 100, usize::MAX).unwrap().is_none());
            assert!(read_as(record, &mut 100, usize::MAX, Raster::TiffAware)
                .unwrap()
                .is_none());
            let image = read_as(record, &mut 100, usize::MAX, Raster::GifAware)
                .unwrap()
                .unwrap();
            assert_eq!(image.extension, "gif");
            assert!(image.bytes.starts_with(sig));
        }
        // A JPEG slot is not read as GIF; zero, oversized or truncated headers
        // are rejected.
        let jpeg = blip(gif(b"GIF89a", 40, 30), 0xf01d, 0x46a0);
        let (record, _) = crate::officeart::record_with_end(&jpeg, 0, &mut 10, "t").unwrap();
        assert!(read_as(record, &mut 100, usize::MAX, Raster::GifAware)
            .unwrap()
            .is_none());
        for data in [
            gif(b"GIF89a", 0, 30),
            gif(b"GIF89a", 40_000, 30),
            b"GIF89a\x01".to_vec(),
        ] {
            let bytes = blip(data, 0xf01e, 0x6e00);
            let (record, _) = crate::officeart::record_with_end(&bytes, 0, &mut 10, "t").unwrap();
            assert!(read_as(record, &mut 100, usize::MAX, Raster::GifAware).is_err());
        }
    }

    #[test]
    fn tiff_data_in_png_blips_is_admitted_only_when_the_host_opts_in() {
        let blip = tiff_in_png(true, &[(256, 3, 40), (257, 4, 30)]);
        let (blip_record, _) = crate::officeart::record_with_end(&blip, 0, &mut 10, "t").unwrap();
        // The shared reader keeps admitting only the advertised encoding.
        assert!(read(blip_record, &mut 100, usize::MAX).unwrap().is_none());
        let image = read_as(blip_record, &mut 100, usize::MAX, Raster::TiffAware)
            .unwrap()
            .unwrap();
        assert_eq!(image.extension, "tiff");
        assert!(image.bytes.starts_with(b"II*\0"));
        let big = tiff_in_png(false, &[(257, 3, 30), (256, 3, 40)]);
        let (blip_record, _) = crate::officeart::record_with_end(&big, 0, &mut 10, "t").unwrap();
        assert_eq!(
            read_as(blip_record, &mut 100, usize::MAX, Raster::TiffAware)
                .unwrap()
                .unwrap()
                .extension,
            "tiff"
        );
        // Missing, zero, oversized or unsupported-type dimensions are rejected.
        for entries in [
            vec![(256u16, 3u16, 40u32)],
            vec![(256, 3, 0), (257, 3, 30)],
            vec![(256, 4, 40_000), (257, 4, 40_000)],
            vec![(256, 2, 40), (257, 3, 30)],
        ] {
            let blip = tiff_in_png(true, &entries);
            let (blip_record, _) =
                crate::officeart::record_with_end(&blip, 0, &mut 10, "t").unwrap();
            assert!(
                read_as(blip_record, &mut 100, usize::MAX, Raster::TiffAware).is_err(),
                "{entries:?}"
            );
        }
        // Other data in a PNG BLIP is still omitted.
        let other = record(0x6e00, 0xf01e, &[vec![0; 17], b"GIF89a".to_vec()].concat());
        let (blip_record, _) = crate::officeart::record_with_end(&other, 0, &mut 10, "t").unwrap();
        assert!(
            read_as(blip_record, &mut 100, usize::MAX, Raster::TiffAware)
                .unwrap()
                .is_none()
        );
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

    fn store_entry(blip: Option<&[u8]>, size: usize, offset: usize, references: u32) -> Vec<u8> {
        let mut payload = vec![0; 36];
        payload[0] = 6;
        payload[20..24].copy_from_slice(&(size as u32).to_le_bytes());
        payload[24..28].copy_from_slice(&references.to_le_bytes());
        payload[28..32].copy_from_slice(&(offset as u32).to_le_bytes());
        if let Some(blip) = blip {
            payload.extend_from_slice(blip);
        }
        record(0x0062, 0xf007, &payload)
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

    #[test]
    fn store_entries_share_framing_and_preserve_primary_or_delayed_identity() {
        let png = png_blip();
        let expected_png = png[8 + 17..].to_vec();
        let (emf, compressed_emf) = emf_test_blip();
        for (blip, expected, owned) in [
            (png.as_slice(), expected_png.as_slice(), false),
            (compressed_emf.as_slice(), emf.as_slice(), true),
        ] {
            for delayed_source in [false, true] {
                let offset = 11;
                let encoded = if delayed_source {
                    store_entry(None, blip.len(), offset, 1)
                } else {
                    store_entry(Some(blip), blip.len(), 0, 1)
                };
                let mut primary = vec![0xaa, 0xbb, 0xcc];
                primary.extend(encoded);
                let mut delayed = vec![0xdd; offset];
                delayed.extend_from_slice(blip);
                let (entry, _) = record_span_with_end(&primary, 3, &mut 1, "BSE").unwrap();

                let mut borrowed_budget = 20;
                let borrowed = read_store_entry(
                    entry.view(&primary).unwrap(),
                    delayed_source.then_some(delayed.as_slice()),
                    &mut borrowed_budget,
                    expected.len(),
                )
                .unwrap()
                .unwrap();
                let mut span_budget = 20;
                let spanned = read_store_entry_span(
                    &entry,
                    &primary,
                    delayed_source.then_some(delayed.as_slice()),
                    &mut span_budget,
                    expected.len(),
                )
                .unwrap()
                .unwrap();
                assert_eq!(borrowed.bytes.as_ref(), expected);
                assert_eq!(spanned.view(&primary, Some(&delayed)).unwrap(), expected);
                assert_eq!(span_budget, borrowed_budget);
                assert_eq!(
                    spanned.backing,
                    if delayed_source {
                        StoreBacking::Delayed
                    } else {
                        StoreBacking::Primary
                    }
                );
                assert_eq!(
                    matches!(&spanned.image.bytes, ImageSpanBytes::Owned(_)),
                    owned
                );

                let moved_primary = primary;
                let moved_delayed = delayed;
                assert_eq!(
                    spanned.view(&moved_primary, Some(&moved_delayed)).unwrap(),
                    expected
                );
                if delayed_source && !owned {
                    let wrong = vec![0x7e; moved_delayed.len()];
                    assert_eq!(
                        spanned.view(&wrong, Some(&moved_delayed)).unwrap(),
                        expected
                    );
                    assert!(spanned.view(&moved_primary, None).is_err());
                }
            }
        }
    }

    #[test]
    fn store_entry_span_preserves_omission_and_exact_source_bounds() {
        let png = png_blip();
        for (references, has_delayed) in [(0, false), (1, false)] {
            let encoded = store_entry(None, png.len(), 7, references);
            let (entry, _) = record_span_with_end(&encoded, 0, &mut 1, "BSE").unwrap();
            assert!(read_store_entry_span(
                &entry,
                &encoded,
                has_delayed.then_some(png.as_slice()),
                &mut 10,
                usize::MAX,
            )
            .unwrap()
            .is_none());
        }

        let encoded = store_entry(None, png.len(), 7, 1);
        let (entry, _) = record_span_with_end(&encoded, 0, &mut 1, "BSE").unwrap();
        assert!(read_store_entry_span(&entry, &encoded, Some(&png), &mut 10, usize::MAX).is_err());
        assert!(read_store_entry_span(
            &entry,
            &encoded[..encoded.len() - 1],
            Some(&png),
            &mut 10,
            usize::MAX,
        )
        .is_err());

        let embedded = store_entry(Some(&png), png.len() - 1, 0, 1);
        let (entry, _) = record_span_with_end(&embedded, 0, &mut 1, "BSE").unwrap();
        assert!(read_store_entry_span(&entry, &embedded, None, &mut 10, usize::MAX).is_err());

        let mut blip_with_trailer = png;
        blip_with_trailer.push(0);
        let embedded = store_entry(Some(&blip_with_trailer), blip_with_trailer.len(), 0, 1);
        let (entry, _) = record_span_with_end(&embedded, 0, &mut 1, "BSE").unwrap();
        assert!(read_store_entry_span(&entry, &embedded, None, &mut 10, usize::MAX).is_err());
    }
}
