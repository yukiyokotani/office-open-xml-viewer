use std::io::{self, Cursor, Seek, SeekFrom, Write};

use zip::write::SimpleFileOptions;

pub fn write_package(
    parts: &[(String, String)],
    max_output_bytes: usize,
) -> Result<Vec<u8>, String> {
    let _projected_uncompressed = parts.iter().try_fold(0usize, |total, (name, body)| {
        total
            .checked_add(name.len())
            .and_then(|value| value.checked_add(body.len()))
            .ok_or_else(|| "generated OOXML size overflow".to_string())
    })?;
    let cursor = BoundedCursor::new(max_output_bytes);
    let mut writer = zip::ZipWriter::new(cursor);
    let options = SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated)
        .unix_permissions(0o644);
    for (name, body) in parts {
        writer.start_file(name, options).map_err(zip_error)?;
        writer.write_all(body.as_bytes()).map_err(io_error)?;
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
