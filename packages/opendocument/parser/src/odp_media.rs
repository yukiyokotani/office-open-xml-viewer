//! Package media referenced by an ODP document: pictures, fill images,
//! bullet images and object replacement images. Each admitted resource is a
//! package file addressed by its exact ODF item name (Part 2 §3.7); the PPTX
//! renderer fetches it lazily through `extract_image` and owns decoding and
//! its pixel budgets.

use crate::iri::{self, Target};
use crate::package::OdfPackage;
use std::collections::HashMap;

/// Implementation bound on the header bytes read to size an image.
const MAX_HEADER_BYTES: usize = 64 * 1024 * 1024;

#[derive(Clone, Debug)]
pub struct Resource {
    pub mime: &'static str,
    /// Intrinsic size in pixels and resolution (dots per inch), when the
    /// raster header states them.
    pub pixels: Option<(u32, u32)>,
    pub dpi: Option<(f64, f64)>,
    /// Vector frame size in EMU (EMF `rclFrame`, 1/100 mm).
    pub frame: Option<(f64, f64)>,
}

impl Resource {
    /// The natural size in EMU (pixels at the stated resolution, else 96 dpi).
    pub fn natural_emu(&self) -> Option<(f64, f64)> {
        if let Some(frame) = self.frame {
            return Some(frame);
        }
        let (w, h) = self.pixels?;
        let (dx, dy) = self.dpi.unwrap_or((96.0, 96.0));
        Some((w as f64 * 914_400.0 / dx, h as f64 * 914_400.0 / dy))
    }
}

pub struct Media<'p> {
    package: &'p OdfPackage,
    admitted: HashMap<String, Resource>,
    /// StarView metafiles converted to SVG, by package path.
    converted: HashMap<String, Vec<u8>>,
}

impl<'p> Media<'p> {
    pub fn new(package: &'p OdfPackage) -> Self {
        Self {
            package,
            admitted: HashMap::new(),
            converted: HashMap::new(),
        }
    }

    /// Admit the package file an `xlink:href` names (relative to the
    /// containing file). External or missing targets yield `None`: an
    /// external image is never fetched.
    pub fn admit(&mut self, base: &str, href: &str) -> Option<(String, Resource)> {
        let Target::Package { path, .. } = iri::resolve(base, href)? else {
            return None;
        };
        if let Some(resource) = self.admitted.get(&path) {
            return Some((path, resource.clone()));
        }
        if !self.package.contains(&path) {
            return None;
        }
        let bytes = self
            .package
            .session()
            .run_operation("odp-media", |operation| {
                operation.read_bytes_bounded(&path, MAX_HEADER_BYTES)
            })
            .ok()?;
        let resource = if bytes.starts_with(b"VCLMTF") {
            let converted = crate::svm::to_svg(&bytes)?;
            self.converted.insert(path.clone(), converted.svg);
            Resource {
                mime: "image/svg+xml",
                pixels: None,
                dpi: None,
                frame: Some(converted.size_emu),
            }
        } else {
            sniff(&bytes)?
        };
        self.admitted.insert(path.clone(), resource.clone());
        Some((path, resource))
    }

    /// Register generated SVG content (an exact hatch fill) under a
    /// synthetic resource key no package item can take (`odp-generated/`).
    pub fn generated(&mut self, kind: &str, svg: Vec<u8>, size: (f64, f64)) -> (String, Resource) {
        let key = format!("odp-generated/{kind}-{}.svg", self.converted.len());
        let resource = Resource {
            mime: "image/svg+xml",
            pixels: None,
            dpi: None,
            frame: Some(size),
        };
        self.converted.insert(key.clone(), svg);
        self.admitted.insert(key.clone(), resource.clone());
        (key, resource)
    }

    pub fn package(&self) -> &'p OdfPackage {
        self.package
    }

    pub fn into_admitted(self) -> (HashMap<String, Resource>, HashMap<String, Vec<u8>>) {
        (self.admitted, self.converted)
    }
}

/// Identify an image by its signature (file names in ODF packages carry no
/// reliable type) and read its pixel size and resolution.
pub fn sniff(bytes: &[u8]) -> Option<Resource> {
    if bytes.starts_with(b"\x89PNG\r\n\x1a\n") {
        let w = u32::from_be_bytes(bytes.get(16..20)?.try_into().ok()?);
        let h = u32::from_be_bytes(bytes.get(20..24)?.try_into().ok()?);
        // pHYs (pixels per metre, unit 1).
        let mut dpi = None;
        let mut pos = 8usize;
        while pos + 12 <= bytes.len() {
            let len = u32::from_be_bytes(bytes[pos..pos + 4].try_into().ok()?) as usize;
            let kind = &bytes[pos + 4..pos + 8];
            if kind == b"pHYs" && len >= 9 && pos + 8 + 9 <= bytes.len() {
                let d = &bytes[pos + 8..];
                let x = u32::from_be_bytes(d[0..4].try_into().ok()?) as f64;
                let y = u32::from_be_bytes(d[4..8].try_into().ok()?) as f64;
                if d[8] == 1 && x > 0.0 && y > 0.0 {
                    dpi = Some((x * 0.0254, y * 0.0254));
                }
                break;
            }
            if kind == b"IDAT" || kind == b"IEND" {
                break;
            }
            pos = pos.checked_add(12 + len)?;
        }
        return Some(Resource {
            mime: "image/png",
            pixels: Some((w, h)),
            dpi,
            frame: None,
        });
    }
    if bytes.starts_with(&[0xFF, 0xD8]) {
        return Some(jpeg(bytes));
    }
    if bytes.starts_with(b"GIF87a") || bytes.starts_with(b"GIF89a") {
        let w = u16::from_le_bytes(bytes.get(6..8)?.try_into().ok()?) as u32;
        let h = u16::from_le_bytes(bytes.get(8..10)?.try_into().ok()?) as u32;
        return Some(Resource {
            mime: "image/gif",
            pixels: Some((w, h)),
            dpi: None,
            frame: None,
        });
    }
    if bytes.starts_with(b"BM") {
        let w = i32::from_le_bytes(bytes.get(18..22)?.try_into().ok()?).unsigned_abs();
        let h = i32::from_le_bytes(bytes.get(22..26)?.try_into().ok()?).unsigned_abs();
        return Some(Resource {
            mime: "image/bmp",
            pixels: Some((w, h)),
            dpi: None,
            frame: None,
        });
    }
    if bytes.starts_with(b"II*\0") || bytes.starts_with(b"MM\0*") {
        return Some(Resource {
            mime: "image/tiff",
            pixels: None,
            dpi: None,
            frame: None,
        });
    }
    if bytes.starts_with(b"RIFF") && bytes.get(8..12) == Some(b"WEBP") {
        return Some(Resource {
            mime: "image/webp",
            pixels: None,
            dpi: None,
            frame: None,
        });
    }
    // EMF: EMR_HEADER record type 1 with the " EMF" signature at offset 40.
    if bytes.get(0..4) == Some(&[1, 0, 0, 0]) && bytes.get(40..44) == Some(b" EMF") {
        let frame = (|| {
            let v = |o: usize| -> Option<f64> {
                Some(i32::from_le_bytes(bytes.get(o..o + 4)?.try_into().ok()?) as f64)
            };
            let (l, t, r, b) = (v(24)?, v(28)?, v(32)?, v(36)?);
            let (w, h) = ((r - l) * 360.0, (b - t) * 360.0);
            (w > 0.0 && h > 0.0).then_some((w, h))
        })();
        return Some(Resource {
            mime: "image/emf",
            pixels: None,
            dpi: None,
            frame,
        });
    }
    // WMF: placeable header or a standard METAHEADER.
    if bytes.starts_with(&[0xD7, 0xCD, 0xC6, 0x9A])
        || (bytes.get(0..2).is_some_and(|t| t == [1, 0] || t == [2, 0])
            && bytes.get(2..4) == Some(&[9, 0]))
    {
        return Some(Resource {
            mime: "image/wmf",
            pixels: None,
            dpi: None,
            frame: None,
        });
    }
    if bytes.starts_with(b"VCLMTF") {
        // StarView metafiles are converted by `Media::admit`.
        return None;
    }
    let head = std::str::from_utf8(&bytes[..bytes.len().min(1024)]).unwrap_or("");
    if head.contains("<svg") || (head.trim_start().starts_with("<?xml") && head.contains("svg")) {
        return Some(Resource {
            mime: "image/svg+xml",
            pixels: None,
            dpi: None,
            frame: None,
        });
    }
    None
}

fn jpeg(bytes: &[u8]) -> Resource {
    let mut pixels = None;
    let mut dpi = None;
    let mut pos = 2usize;
    while pos + 4 <= bytes.len() {
        if bytes[pos] != 0xFF {
            break;
        }
        let marker = bytes[pos + 1];
        if marker == 0xD8 || (0xD0..=0xD7).contains(&marker) || marker == 0x01 {
            pos += 2;
            continue;
        }
        let len = u16::from_be_bytes([bytes[pos + 2], bytes[pos + 3]]) as usize;
        let body = &bytes[pos + 4..(pos + 2 + len).min(bytes.len())];
        match marker {
            0xE0 if body.starts_with(b"JFIF\0") && body.len() >= 12 => {
                let unit = body[7];
                let x = u16::from_be_bytes([body[8], body[9]]) as f64;
                let y = u16::from_be_bytes([body[10], body[11]]) as f64;
                if x > 0.0 && y > 0.0 {
                    dpi = match unit {
                        1 => Some((x, y)),
                        2 => Some((x * 2.54, y * 2.54)),
                        _ => None,
                    };
                }
            }
            0xC0..=0xCF
                if marker != 0xC4 && marker != 0xC8 && marker != 0xCC && body.len() >= 5 =>
            {
                let h = u16::from_be_bytes([body[1], body[2]]) as u32;
                let w = u16::from_be_bytes([body[3], body[4]]) as u32;
                pixels = Some((w, h));
                break;
            }
            0xDA | 0xD9 => break,
            _ => {}
        }
        pos += 2 + len;
    }
    Resource {
        mime: "image/jpeg",
        pixels,
        dpi,
        frame: None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn sniffs_png_size_and_resolution() {
        let mut png = b"\x89PNG\r\n\x1a\n".to_vec();
        png.extend_from_slice(&13u32.to_be_bytes());
        png.extend_from_slice(b"IHDR");
        png.extend_from_slice(&40u32.to_be_bytes());
        png.extend_from_slice(&20u32.to_be_bytes());
        png.extend_from_slice(&[8, 6, 0, 0, 0, 0, 0, 0, 0]);
        png.extend_from_slice(&9u32.to_be_bytes());
        png.extend_from_slice(b"pHYs");
        png.extend_from_slice(&7874u32.to_be_bytes());
        png.extend_from_slice(&7874u32.to_be_bytes());
        png.extend_from_slice(&[1, 0, 0, 0, 0]);
        let r = sniff(&png).unwrap();
        assert_eq!(r.mime, "image/png");
        assert_eq!(r.pixels, Some((40, 20)));
        let (dx, _) = r.dpi.unwrap();
        assert!((dx - 200.0).abs() < 0.1);
        assert!(sniff(b"VCLMTF....").is_none());
    }
}
