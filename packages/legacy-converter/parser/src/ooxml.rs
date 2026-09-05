use std::io::{self, Cursor, Seek, SeekFrom, Write};

use zip::write::SimpleFileOptions;

pub fn write_package(
    parts: &[(String, String)],
    max_output_bytes: usize,
) -> Result<Vec<u8>, String> {
    write_package_bytes(
        parts
            .iter()
            .map(|(name, body)| (name.as_str(), body.as_bytes())),
        max_output_bytes,
    )
}

/// Borrow binary media without cloning it into UTF-8 or a second part buffer.
pub fn write_package_bytes<'a>(
    parts: impl IntoIterator<Item = (&'a str, &'a [u8])>,
    max_output_bytes: usize,
) -> Result<Vec<u8>, String> {
    let cursor = BoundedCursor::new(max_output_bytes);
    let mut writer = zip::ZipWriter::new(cursor);
    let options = SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated)
        .unix_permissions(0o644);
    for (name, body) in parts {
        writer.start_file(name, options).map_err(zip_error)?;
        writer.write_all(body).map_err(io_error)?;
    }
    let output = writer.finish().map_err(zip_error)?.into_inner();
    if output.len() > max_output_bytes {
        return Err("OUTPUT_TOO_LARGE".into());
    }
    Ok(output)
}

struct BoundedCursor {
    inner: Cursor<Vec<u8>>,
    maximum: usize,
}

impl BoundedCursor {
    fn new(maximum: usize) -> Self {
        Self {
            inner: Cursor::new(Vec::new()),
            maximum,
        }
    }

    fn into_inner(self) -> Vec<u8> {
        self.inner.into_inner()
    }
}

impl Write for BoundedCursor {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        let position = usize::try_from(self.inner.position()).map_err(|_| output_limit_error())?;
        let end = position
            .checked_add(buffer.len())
            .ok_or_else(output_limit_error)?;
        if end.max(self.inner.get_ref().len()) > self.maximum {
            return Err(output_limit_error());
        }
        self.inner.write(buffer)
    }

    fn flush(&mut self) -> io::Result<()> {
        self.inner.flush()
    }
}

impl Seek for BoundedCursor {
    fn seek(&mut self, position: SeekFrom) -> io::Result<u64> {
        let old = self.inner.position();
        let next = self.inner.seek(position)?;
        if next > self.maximum as u64 {
            self.inner.set_position(old);
            return Err(output_limit_error());
        }
        Ok(next)
    }
}

fn output_limit_error() -> io::Error {
    io::Error::new(io::ErrorKind::StorageFull, "OUTPUT_TOO_LARGE")
}

fn zip_error(error: zip::result::ZipError) -> String {
    if error.to_string().contains("OUTPUT_TOO_LARGE") {
        "OUTPUT_TOO_LARGE".into()
    } else {
        "failed to write OOXML package".into()
    }
}

fn io_error(error: io::Error) -> String {
    if error.to_string().contains("OUTPUT_TOO_LARGE") {
        "OUTPUT_TOO_LARGE".into()
    } else {
        "failed to write OOXML package".into()
    }
}

pub fn xml_text(input: &str) -> String {
    let mut output = String::with_capacity(input.len());
    for character in input.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '\r' | '\n' | '\t' => output.push(character),
            '\u{20}'..='\u{d7ff}' | '\u{e000}'..='\u{fffd}' | '\u{10000}'..='\u{10ffff}' => {
                output.push(character)
            }
            _ => output.push('\u{fffd}'),
        }
    }
    output
}

pub fn xml_attr(input: &str) -> String {
    xml_text(input)
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

pub const ROOT_RELS_DOCX: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;

pub const ROOT_RELS_XLSX: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#;

pub const ROOT_RELS_PPTX: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#;

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Read;

    #[test]
    fn binary_parts_round_trip_without_utf8_conversion_and_obey_output_limit() {
        let binary = [0, 255, 128, 1];
        let parts = [
            ("part.xml", b"<part/>".as_slice()),
            ("media.png", binary.as_slice()),
        ];
        let bytes = write_package_bytes(parts, 4096).unwrap();
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes.clone())).unwrap();
        let mut actual = Vec::new();
        archive
            .by_name("media.png")
            .unwrap()
            .read_to_end(&mut actual)
            .unwrap();
        assert_eq!(actual, binary);
        assert_eq!(write_package_bytes(parts, bytes.len()).unwrap(), bytes);
        assert_eq!(
            write_package_bytes(parts, bytes.len() - 1).unwrap_err(),
            "OUTPUT_TOO_LARGE"
        );
    }
}
