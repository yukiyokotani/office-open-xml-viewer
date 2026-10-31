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

/// The media type of `bytes` once every frame of it has decoded.
pub(super) fn decode(bytes: &[u8]) -> Result<&'static str, String> {
    if bytes.starts_with(b"\x89PNG\r\n\x1a\n") {
        png(bytes)?;
        Ok("image/png")
    } else if bytes.starts_with(b"GIF87a") || bytes.starts_with(b"GIF89a") {
        gif(bytes)?;
        Ok("image/gif")
    } else if bytes.starts_with(&[0xff, 0xd8, 0xff]) {
        jpeg(bytes)?;
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

fn png(bytes: &[u8]) -> Result<(), String> {
    let decode = || unsupported("image-decode");
    let limits = png::Limits {
        bytes: MAX_DECODE_BYTES as usize,
    };
    let mut reader = png::Decoder::new_with_limits(Cursor::new(bytes), limits)
        .read_info()
        .map_err(|_| decode())?;
    let info = reader.info();
    check_size(u64::from(info.width), u64::from(info.height))?;
    // An animated PNG plays in a browser; LibreOffice is not measured.
    if info.animation_control.is_some() {
        return Err(unsupported("image-animation"));
    }
    let size = reader.output_buffer_size().ok_or_else(decode)?;
    let mut buffer = vec![0; size];
    reader.next_frame(&mut buffer).map_err(|_| decode())?;
    reader.finish().map_err(|_| decode())
}

fn gif(bytes: &[u8]) -> Result<(), String> {
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
    while decoder.read_next_frame().map_err(|_| decode())?.is_some() {
        frames += 1;
        // Every frame is decoded within the screen's own bounds; their count
        // is bounded by the file (each frame takes bytes).
        if frames > bytes.len() {
            return Err(decode());
        }
    }
    if frames == 0 {
        return Err(decode());
    }
    Ok(())
}

fn jpeg(bytes: &[u8]) -> Result<(), String> {
    let decode = || unsupported("image-decode");
    let options = zune_core::options::DecoderOptions::default()
        .set_max_width(MAX_SIDE as usize)
        .set_max_height(MAX_SIDE as usize)
        .set_strict_mode(true);
    let mut decoder = zune_jpeg::JpegDecoder::new_with_options(Cursor::new(bytes), options);
    decoder.decode_headers().map_err(|_| decode())?;
    let (width, height) = decoder.dimensions().ok_or_else(decode)?;
    check_size(width as u64, height as u64)?;
    decoder.decode().map_err(|_| decode())?;
    Ok(())
}

#[cfg(test)]
pub(crate) mod tests {
    use super::*;

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
        assert_eq!(decode(&bytes).unwrap(), "image/png");
        // A truncated stream stops inside the image data.
        assert_eq!(
            decode(&bytes[..bytes.len() - 20]).unwrap_err(),
            unsupported("image-decode")
        );
        let mut corrupt = bytes.clone();
        let data = corrupt.len() - 20;
        corrupt[data] ^= 0xff;
        assert_eq!(decode(&corrupt).unwrap_err(), unsupported("image-decode"));
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
        assert_eq!(decode(&out).unwrap(), "image/gif");
        assert_eq!(
            decode(&out[..out.len() - 8]).unwrap_err(),
            unsupported("image-decode")
        );
    }

    #[test]
    fn decodes_jpeg_and_rejects_truncation() {
        // A 1×1 baseline JPEG.
        let bytes: &[u8] = &JPEG_1X1;
        assert_eq!(decode(bytes).unwrap(), "image/jpeg");
        assert_eq!(
            decode(&bytes[..bytes.len() - 6]).unwrap_err(),
            unsupported("image-decode")
        );
    }

    #[test]
    fn rejects_other_formats_and_sizes() {
        assert_eq!(decode(b"BM....").unwrap_err(), unsupported("image-format"));
        assert_eq!(
            decode(b"<svg xmlns='http://www.w3.org/2000/svg'/>").unwrap_err(),
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
