//! Image resources (design §7): an image class is admitted only when the
//! reader decodes it completely. PNG, GIF and JPEG, recognised by their
//! signatures, are decoded here with bounded memory; anything else, and any
//! file a decoder rejects, rejects the document.

use crate::error::unsupported;
use std::io::Cursor;

/// Decoded-size bounds: 32768 px per side and 40 M px (the bounds the
/// legacy readers apply to their raster images).
const MAX_SIDE: u32 = 32_768;
const MAX_PIXELS: u64 = 40_000_000;
/// Decoder working memory: the largest admitted image at 8 bytes per pixel
/// (16-bit RGBA PNG), plus room for the decoders' own state.
const MAX_DECODE_BYTES: u64 = MAX_PIXELS * 8 + (1 << 20);
/// Implementation bound on the pixels a document's pictures decode, every
/// frame counted, charged before each frame is decoded.
pub(super) const MAX_DOCUMENT_PIXELS: u64 = 64_000_000;

/// The media type of `bytes` once every frame of it has decoded, charging
/// the decoded pixels against `budget` (the document's remaining pixels).
pub(super) fn decode(bytes: &[u8], budget: &mut u64) -> Result<&'static str, String> {
    if bytes.starts_with(b"\x89PNG\r\n\x1a\n") {
        png(bytes, budget)?;
        Ok("image/png")
    } else if bytes.starts_with(b"GIF87a") || bytes.starts_with(b"GIF89a") {
        gif(bytes, budget)?;
        Ok("image/gif")
    } else if bytes.starts_with(&[0xff, 0xd8, 0xff]) {
        jpeg(bytes, budget)?;
        Ok("image/jpeg")
    } else {
        Err(unsupported("image-format"))
    }
}

fn check_size(width: u64, height: u64) -> Result<(), String> {
    if width == 0
        || height == 0
        || width > u64::from(MAX_SIDE)
        || height > u64::from(MAX_SIDE)
        || width * height > MAX_PIXELS
    {
        return Err(unsupported("image-size"));
    }
    Ok(())
}

fn charge(budget: &mut u64, width: u64, height: u64) -> Result<(), String> {
    *budget = budget
        .checked_sub(width * height)
        .ok_or_else(|| "OpenDocument pictures exceed the decoding limit".to_string())?;
    Ok(())
}

fn png(bytes: &[u8], budget: &mut u64) -> Result<(), String> {
    let decode = || unsupported("image-decode");
    let limits = png::Limits {
        bytes: MAX_DECODE_BYTES as usize,
    };
    let mut decoder = png::Decoder::new_with_limits(Cursor::new(bytes), limits);
    // Chunk CRCs and the zlib Adler-32 are verified: LibreOffice rejects
    // image data with a wrong check value where a browser paints it.
    decoder.ignore_checksums(false);
    let mut reader = decoder.read_info().map_err(|_| decode())?;
    let info = reader.info();
    let (width, height) = (info.width, info.height);
    check_size(u64::from(width), u64::from(height))?;
    // An animated PNG plays in a browser; LibreOffice is not measured.
    if info.animation_control.is_some() {
        return Err(unsupported("image-animation"));
    }
    let palette = (info.color_type == png::ColorType::Indexed)
        .then(|| info.palette.as_ref().map_or(0, |palette| palette.len() / 3));
    charge(budget, u64::from(width), u64::from(height))?;
    let size = reader.output_buffer_size().ok_or_else(decode)?;
    let mut buffer = vec![0; size];
    let output = reader.next_frame(&mut buffer).map_err(|_| decode())?;
    // Indices past the palette paint differently (a browser's choice,
    // LibreOffice's black), so each one must name an entry.
    if let Some(entries) = palette {
        let depth = output.bit_depth as usize;
        for row in buffer.chunks(output.line_size).take(height as usize) {
            for x in 0..width as usize {
                let bit = x * depth;
                let index = (row[bit / 8] >> (8 - depth - bit % 8)) & ((1u16 << depth) - 1) as u8;
                if usize::from(index) >= entries {
                    return Err(unsupported("image-palette"));
                }
            }
        }
    }
    reader.finish().map_err(|_| decode())
}

fn gif(bytes: &[u8], budget: &mut u64) -> Result<(), String> {
    let decode = || unsupported("image-decode");
    let mut options = gif::DecodeOptions::new();
    options.set_color_output(gif::ColorOutput::Indexed);
    options.set_memory_limit(gif::MemoryLimit::Bytes(
        std::num::NonZeroU64::new(MAX_DECODE_BYTES).expect("non-zero"),
    ));
    options.check_frame_consistency(true);
    let mut decoder = options
        .read_info(Cursor::new(bytes))
        .map_err(|_| decode())?;
    check_size(u64::from(decoder.width()), u64::from(decoder.height()))?;
    let mut frames = 0usize;
    while let Some(frame) = decoder.next_frame_info().map_err(|_| decode())? {
        // Every frame lies within the screen; each is charged before it is
        // decoded.
        charge(budget, u64::from(frame.width), u64::from(frame.height))?;
        let mut buffer = vec![0; decoder.buffer_size()];
        decoder
            .read_into_buffer(&mut buffer)
            .map_err(|_| decode())?;
        // Indices past the colour table paint differently per decoder.
        let entries = decoder.palette().map_err(|_| decode())?.len() / 3;
        if buffer.iter().any(|&index| usize::from(index) >= entries) {
            return Err(unsupported("image-palette"));
        }
        frames += 1;
    }
    if frames == 0 {
        return Err(decode());
    }
    Ok(())
}

fn jpeg(bytes: &[u8], budget: &mut u64) -> Result<(), String> {
    let decode = || unsupported("image-decode");
    let options = zune_core::options::DecoderOptions::default()
        .set_max_width(MAX_SIDE as usize)
        .set_max_height(MAX_SIDE as usize)
        .set_strict_mode(true);
    let mut decoder = zune_jpeg::JpegDecoder::new_with_options(Cursor::new(bytes), options);
    decoder.decode_headers().map_err(|_| decode())?;
    let (width, height) = decoder.dimensions().ok_or_else(decode)?;
    check_size(width as u64, height as u64)?;
    charge(budget, width as u64, height as u64)?;
    decoder.decode().map_err(|_| decode())?;
    Ok(())
}

#[cfg(test)]
pub(crate) mod tests {
    use super::*;

    fn dec(bytes: &[u8]) -> Result<&'static str, String> {
        decode(bytes, &mut MAX_DOCUMENT_PIXELS.clone())
    }

    /// A PNG with `chunks` (type, data) after its IHDR, CRCs computed.
    fn png_chunks(ihdr: &[u8], chunks: &[(&[u8; 4], Vec<u8>)]) -> Vec<u8> {
        let mut out = b"\x89PNG\r\n\x1a\n".to_vec();
        let mut chunk = |kind: &[u8; 4], data: &[u8]| {
            out.extend((data.len() as u32).to_be_bytes());
            let mut crc = crc32fast::Hasher::new();
            crc.update(kind);
            crc.update(data);
            out.extend(kind);
            out.extend(data);
            out.extend(crc.finalize().to_be_bytes());
        };
        chunk(b"IHDR", ihdr);
        for (kind, data) in chunks {
            chunk(kind, data);
        }
        chunk(b"IEND", &[]);
        out
    }

    fn zlib(raw: &[u8]) -> Vec<u8> {
        use std::io::Write;
        let mut encoder =
            flate2::write::ZlibEncoder::new(Vec::new(), flate2::Compression::default());
        encoder.write_all(raw).unwrap();
        encoder.finish().unwrap()
    }

    /// A 2×1 RGB PNG (one red, one blue pixel).
    pub(crate) fn png_bytes() -> Vec<u8> {
        let mut out = Vec::new();
        {
            let mut encoder = png::Encoder::new(&mut out, 2, 1);
            encoder.set_color(png::ColorType::Rgb);
            let mut writer = encoder.write_header().unwrap();
            writer.write_image_data(&[255, 0, 0, 0, 0, 255]).unwrap();
        }
        out
    }

    #[test]
    fn decodes_png_and_rejects_damage() {
        let bytes = png_bytes();
        assert_eq!(dec(&bytes).unwrap(), "image/png");
        // A truncated stream stops inside the image data.
        assert_eq!(
            dec(&bytes[..bytes.len() - 20]).unwrap_err(),
            unsupported("image-decode")
        );
        let mut corrupt = bytes.clone();
        let data = corrupt.len() - 20;
        corrupt[data] ^= 0xff;
        assert_eq!(dec(&corrupt).unwrap_err(), unsupported("image-decode"));
    }

    #[test]
    fn verifies_png_checksums_and_palette_indices() {
        // 1×1 8-bit indexed, one palette entry.
        let ihdr = [0, 0, 0, 1, 0, 0, 0, 1, 8, 3, 0, 0, 0];
        let palette = (b"PLTE", vec![255, 0, 0]);
        let valid = png_chunks(&ihdr, &[palette.clone(), (b"IDAT", zlib(&[0, 0]))]);
        assert_eq!(dec(&valid).unwrap(), "image/png");
        // A pixel index past the palette.
        let outside = png_chunks(&ihdr, &[palette.clone(), (b"IDAT", zlib(&[0, 255]))]);
        assert_eq!(dec(&outside).unwrap_err(), unsupported("image-palette"));
        // A wrong zlib Adler-32 under a correct chunk CRC.
        let mut data = zlib(&[0, 0]);
        let end = data.len() - 1;
        data[end] ^= 1;
        let adler = png_chunks(&ihdr, &[palette, (b"IDAT", data)]);
        assert_eq!(dec(&adler).unwrap_err(), unsupported("image-decode"));
    }

    #[test]
    fn charges_every_decoded_pixel_against_the_document() {
        let bytes = png_bytes();
        let mut budget = 3;
        assert_eq!(decode(&bytes, &mut budget).unwrap(), "image/png");
        assert_eq!(budget, 1);
        assert!(decode(&bytes, &mut budget).is_err());
    }

    #[test]
    fn decodes_gif_frames() {
        let mut out = Vec::new();
        {
            let mut encoder = gif::Encoder::new(&mut out, 2, 2, &[0, 0, 0, 255, 0, 0]).unwrap();
            for _ in 0..2 {
                let frame = gif::Frame {
                    width: 2,
                    height: 2,
                    buffer: std::borrow::Cow::Borrowed(&[0, 1, 1, 0]),
                    ..Default::default()
                };
                encoder.write_frame(&frame).unwrap();
            }
        }
        assert_eq!(dec(&out).unwrap(), "image/gif");
        // Two 2×2 frames: 8 pixels.
        let mut budget = 7;
        assert!(decode(&out, &mut budget).is_err());
        // An index past the two-entry table.
        let mut outside = Vec::new();
        {
            let mut encoder = gif::Encoder::new(&mut outside, 1, 1, &[0, 0, 0, 255, 0, 0]).unwrap();
            let frame = gif::Frame {
                width: 1,
                height: 1,
                buffer: std::borrow::Cow::Borrowed(&[3]),
                ..Default::default()
            };
            encoder.write_frame(&frame).unwrap();
        }
        assert_eq!(dec(&outside).unwrap_err(), unsupported("image-palette"));
        assert_eq!(
            dec(&out[..out.len() - 8]).unwrap_err(),
            unsupported("image-decode")
        );
    }

    #[test]
    fn decodes_jpeg_and_rejects_truncation() {
        // A 1×1 baseline JPEG.
        let bytes: &[u8] = &JPEG_1X1;
        assert_eq!(dec(bytes).unwrap(), "image/jpeg");
        assert_eq!(
            dec(&bytes[..bytes.len() - 6]).unwrap_err(),
            unsupported("image-decode")
        );
    }

    #[test]
    fn rejects_other_formats_and_sizes() {
        assert_eq!(dec(b"BM....").unwrap_err(), unsupported("image-format"));
        assert_eq!(
            dec(b"<svg xmlns='http://www.w3.org/2000/svg'/>").unwrap_err(),
            unsupported("image-format")
        );
        assert_eq!(check_size(0, 1).unwrap_err(), unsupported("image-size"));
        assert_eq!(
            check_size(32_769, 1).unwrap_err(),
            unsupported("image-size")
        );
        assert_eq!(
            check_size(8_000, 8_000).unwrap_err(),
            unsupported("image-size")
        );
        assert!(check_size(6_000, 6_000).is_ok());
    }

    /// 1×1 grey baseline JPEG (libjpeg output).
    const JPEG_1X1: [u8; 159] = [
        0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46, 0x00, 0x01, 0x01, 0x00, 0x00,
        0x01, 0x00, 0x01, 0x00, 0x00, 0xff, 0xdb, 0x00, 0x43, 0x00, 0x10, 0x0b, 0x0c, 0x0e, 0x0c,
        0x0a, 0x10, 0x0e, 0x0d, 0x0e, 0x12, 0x11, 0x10, 0x13, 0x18, 0x28, 0x1a, 0x18, 0x16, 0x16,
        0x18, 0x31, 0x23, 0x25, 0x1d, 0x28, 0x3a, 0x33, 0x3d, 0x3c, 0x39, 0x33, 0x38, 0x37, 0x40,
        0x48, 0x5c, 0x4e, 0x40, 0x44, 0x57, 0x45, 0x37, 0x38, 0x50, 0x6d, 0x51, 0x57, 0x5f, 0x62,
        0x67, 0x68, 0x67, 0x3e, 0x4d, 0x71, 0x79, 0x70, 0x64, 0x78, 0x5c, 0x65, 0x67, 0x63, 0xff,
        0xc0, 0x00, 0x0b, 0x08, 0x00, 0x01, 0x00, 0x01, 0x01, 0x01, 0x11, 0x00, 0xff, 0xc4, 0x00,
        0x14, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0xff, 0xc4, 0x00, 0x14, 0x10, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xff, 0xda, 0x00, 0x08,
        0x01, 0x01, 0x00, 0x00, 0x3f, 0x00, 0x3f, 0xff, 0xd9,
    ];
}
