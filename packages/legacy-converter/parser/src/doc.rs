//! Word Binary File (`.doc`) compatibility subset.
//!
//! The reader accepts Word 97-2003 FIB/CLX piece tables and preserves main-story
//! text, paragraphs, tabs, line breaks, page breaks, and displayed field results.
//! Formatting, sections, drawings, revisions, headers/footers, notes, and OLE are
//! deliberately not inferred. See [MS-DOC] 2.5.1 (FIB), 2.8.35 (Clx), and
//! 2.9.177 (PlcPcd). Unsupported/encrypted inputs fail closed.

use crate::cfb::CompoundFile;
use crate::ooxml::{write_package, xml_text, ROOT_RELS_DOCX};

const FIB_IDENT: u16 = 0xa5ec;
const FIB_WORD_97: u16 = 0x00c1;
const FIB_FLAGS_OFFSET: usize = 0x0a;
const CCP_TEXT_OFFSET: usize = 0x4c;
const FC_CLX_OFFSET: usize = 0x01a2;
const LCB_CLX_OFFSET: usize = 0x01a6;
const MAX_PIECES: usize = 1_000_000;

pub struct DocConversion {
    pub bytes: Vec<u8>,
    pub warnings: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum Token {
    Text(String),
    Tab,
    LineBreak,
    PageBreak,
}

pub fn convert(cfb: &CompoundFile<'_>, max_output_bytes: usize) -> Result<DocConversion, String> {
    let word = cfb.stream("WordDocument").map_err(unsupported)?;
    if word.len() < LCB_CLX_OFFSET + 4 {
        return Err(unsupported("truncated Word FIB"));
    }
    if u16_at(&word, 0)? != FIB_IDENT || u16_at(&word, 2)? < FIB_WORD_97 {
        return Err(unsupported(
            "only Word 97-2003 binary documents are supported",
        ));
    }
    let flags = u16_at(&word, FIB_FLAGS_OFFSET)?;
    if (flags & 0x0100) != 0 || (flags & 0x8000) != 0 {
        return Err(unsupported(
            "encrypted Word binary documents are not supported",
        ));
    }
    let table_name = if (flags & 0x0200) != 0 {
        "1Table"
    } else {
        "0Table"
    };
    let table = cfb.stream(table_name).map_err(unsupported)?;
    let ccp_text = usize::try_from(u32_at(&word, CCP_TEXT_OFFSET)?)
        .map_err(|_| unsupported("Word main story is too large"))?;
    let fc_clx = usize::try_from(u32_at(&word, FC_CLX_OFFSET)?)
        .map_err(|_| unsupported("Word CLX offset is too large"))?;
    let lcb_clx = usize::try_from(u32_at(&word, LCB_CLX_OFFSET)?)
        .map_err(|_| unsupported("Word CLX size is too large"))?;
    let clx_end = fc_clx
        .checked_add(lcb_clx)
        .ok_or_else(|| unsupported("Word CLX range overflow"))?;
    let clx = table
        .get(fc_clx..clx_end)
        .ok_or_else(|| unsupported("Word CLX lies outside its table stream"))?;
    let text = decode_piece_table(&word, clx, ccp_text)?;
    let paragraphs = tokenize_story(&text);
    let document_xml = build_document_xml(&paragraphs);
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#;
    let parts = [
        ("[Content_Types].xml".into(), content_types.to_string()),
        ("_rels/.rels".into(), ROOT_RELS_DOCX.to_string()),
        ("word/document.xml".into(), document_xml),
    ];
    Ok(DocConversion {
        bytes: write_package(&parts, max_output_bytes)?,
        warnings: vec![
            "legacy-doc:main-story-plain-text-only".into(),
            "legacy-doc:formatting-and-embedded-objects-omitted".into(),
        ],
    })
}

fn decode_piece_table(word: &[u8], clx: &[u8], ccp_text: usize) -> Result<String, String> {
    let mut offset = 0usize;
    while clx.get(offset) == Some(&0x01) {
        let size = u16_at(clx, offset + 1)? as usize;
        offset = offset
            .checked_add(3 + size)
            .ok_or_else(|| unsupported("Word CLX record overflow"))?;
        if offset > clx.len() {
            return Err(unsupported("truncated Word CLX property record"));
        }
    }
    if clx.get(offset) != Some(&0x02) {
        return Err(unsupported("missing Word piece table"));
    }
    let plc_size = usize::try_from(u32_at(clx, offset + 1)?)
        .map_err(|_| unsupported("Word piece table is too large"))?;
    let plc_start = offset + 5;
    let plc_end = plc_start
        .checked_add(plc_size)
        .ok_or_else(|| unsupported("Word piece table range overflow"))?;
    let plc = clx
        .get(plc_start..plc_end)
        .ok_or_else(|| unsupported("truncated Word piece table"))?;
    if plc_size < 4 || (plc_size - 4) % 12 != 0 {
        return Err(unsupported("invalid Word piece table size"));
    }
    let piece_count = (plc_size - 4) / 12;
    if piece_count == 0 || piece_count > MAX_PIECES {
        return Err(unsupported("invalid Word piece count"));
    }
    let cp_bytes = (piece_count + 1) * 4;
    let pcd_start = cp_bytes;
    let mut output = String::with_capacity(ccp_text.min(1024 * 1024));
    let mut expected_cp = 0usize;
    for index in 0..piece_count {
        let cp_start = usize::try_from(u32_at(plc, index * 4)?)
            .map_err(|_| unsupported("Word character position is too large"))?;
        let cp_end = usize::try_from(u32_at(plc, (index + 1) * 4)?)
            .map_err(|_| unsupported("Word character position is too large"))?;
        if cp_start > cp_end || cp_start != expected_cp {
            return Err(unsupported("non-contiguous Word piece table"));
        }
        expected_cp = cp_end;
        if cp_start >= ccp_text {
            break;
        }
        let chars = cp_end.min(ccp_text) - cp_start;
        let pcd = pcd_start + index * 8;
        let raw_fc = u32_at(plc, pcd + 2)?;
        let compressed = (raw_fc & 0x4000_0000) != 0;
        let file_offset = usize::try_from(raw_fc & 0x3fff_ffff)
            .map_err(|_| unsupported("Word piece offset is too large"))?;
        if compressed {
            let file_offset = file_offset / 2;
            let end = file_offset
                .checked_add(chars)
                .ok_or_else(|| unsupported("Word text piece range overflow"))?;
            let bytes = word
                .get(file_offset..end)
                .ok_or_else(|| unsupported("Word text piece lies outside WordDocument"))?;
            for byte in bytes {
                output.push(decode_windows_1252(*byte));
            }
        } else {
            let byte_count = chars
                .checked_mul(2)
                .ok_or_else(|| unsupported("Word text piece size overflow"))?;
            let end = file_offset
                .checked_add(byte_count)
                .ok_or_else(|| unsupported("Word Unicode piece range overflow"))?;
            let bytes = word
                .get(file_offset..end)
                .ok_or_else(|| unsupported("Word Unicode piece lies outside WordDocument"))?;
            let units = bytes
                .chunks_exact(2)
                .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]));
            output.extend(char::decode_utf16(units).map(|value| value.unwrap_or('\u{fffd}')));
        }
    }
    if expected_cp < ccp_text {
        return Err(unsupported(
            "Word piece table does not cover the main story",
        ));
    }
    Ok(output)
}

fn tokenize_story(text: &str) -> Vec<Vec<Token>> {
    let mut paragraphs = vec![Vec::new()];
    let mut buffered = String::new();
    let mut fields: Vec<bool> = Vec::new();
    let flush = |paragraph: &mut Vec<Token>, buffered: &mut String| {
        if !buffered.is_empty() {
            paragraph.push(Token::Text(std::mem::take(buffered)));
        }
    };
    for character in text.chars() {
        match character {
            '\u{13}' => fields.push(false),
            '\u{14}' if !fields.is_empty() => {
                if let Some(field) = fields.last_mut() {
                    *field = true;
                }
            }
            '\u{15}' if !fields.is_empty() => {
                fields.pop();
            }
            _ if fields.iter().any(|result| !result) => {}
            '\r' => {
                flush(
                    paragraphs.last_mut().expect("opening paragraph"),
                    &mut buffered,
                );
                paragraphs.push(Vec::new());
            }
            '\t' => {
                flush(
                    paragraphs.last_mut().expect("opening paragraph"),
                    &mut buffered,
                );
                paragraphs
                    .last_mut()
                    .expect("opening paragraph")
                    .push(Token::Tab);
            }
            '\u{0b}' => {
                flush(
                    paragraphs.last_mut().expect("opening paragraph"),
                    &mut buffered,
                );
                paragraphs
                    .last_mut()
                    .expect("opening paragraph")
                    .push(Token::LineBreak);
            }
            '\u{0c}' => {
                flush(
                    paragraphs.last_mut().expect("opening paragraph"),
                    &mut buffered,
                );
                paragraphs
                    .last_mut()
                    .expect("opening paragraph")
                    .push(Token::PageBreak);
            }
            '\u{20}'..='\u{10ffff}' => buffered.push(character),
            _ => {}
        }
    }
    flush(
        paragraphs.last_mut().expect("opening paragraph"),
        &mut buffered,
    );
    if paragraphs.len() > 1 && paragraphs.last().is_some_and(Vec::is_empty) {
        paragraphs.pop();
    }
    paragraphs
}

fn build_document_xml(paragraphs: &[Vec<Token>]) -> String {
    let mut xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>"#,
    );
    for paragraph in paragraphs {
        xml.push_str("<w:p>");
        for token in paragraph {
            match token {
                Token::Text(text) => {
                    xml.push_str("<w:r><w:t xml:space=\"preserve\">");
                    xml.push_str(&xml_text(text));
                    xml.push_str("</w:t></w:r>");
                }
                Token::Tab => xml.push_str("<w:r><w:tab/></w:r>"),
                Token::LineBreak => xml.push_str("<w:r><w:br/></w:r>"),
                Token::PageBreak => xml.push_str("<w:r><w:br w:type=\"page\"/></w:r>"),
            }
        }
        xml.push_str("</w:p>");
    }
    xml.push_str("<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\"/></w:sectPr></w:body></w:document>");
    xml
}

fn decode_windows_1252(byte: u8) -> char {
    const SPECIAL: [char; 32] = [
        '\u{20ac}', '\u{0081}', '\u{201a}', '\u{0192}', '\u{201e}', '\u{2026}', '\u{2020}',
        '\u{2021}', '\u{02c6}', '\u{2030}', '\u{0160}', '\u{2039}', '\u{0152}', '\u{008d}',
        '\u{017d}', '\u{008f}', '\u{0090}', '\u{2018}', '\u{2019}', '\u{201c}', '\u{201d}',
        '\u{2022}', '\u{2013}', '\u{2014}', '\u{02dc}', '\u{2122}', '\u{0161}', '\u{203a}',
        '\u{0153}', '\u{009d}', '\u{017e}', '\u{0178}',
    ];
    if (0x80..=0x9f).contains(&byte) {
        SPECIAL[(byte - 0x80) as usize]
    } else {
        char::from(byte)
    }
}

fn unsupported(message: impl Into<String>) -> String {
    format!("UNSUPPORTED:{}", message.into())
}

fn u16_at(bytes: &[u8], offset: usize) -> Result<u16, String> {
    let raw = bytes
        .get(offset..offset + 2)
        .ok_or_else(|| unsupported("truncated Word integer"))?;
    Ok(u16::from_le_bytes([raw[0], raw[1]]))
}

fn u32_at(bytes: &[u8], offset: usize) -> Result<u32, String> {
    let raw = bytes
        .get(offset..offset + 4)
        .ok_or_else(|| unsupported("truncated Word integer"))?;
    Ok(u32::from_le_bytes(raw.try_into().expect("four-byte slice")))
}

#[cfg(test)]
mod tests {
    use super::{decode_piece_table, tokenize_story, Token};

    #[test]
    fn decodes_a_unicode_piece_table() {
        let source = "Hello 日本語\rSecond";
        let units: Vec<u16> = source.encode_utf16().collect();
        let mut word = vec![0u8; 128];
        for (index, unit) in units.iter().enumerate() {
            word[32 + index * 2..34 + index * 2].copy_from_slice(&unit.to_le_bytes());
        }
        let chars = units.len() as u32;
        let mut clx = vec![0x02];
        clx.extend_from_slice(&16u32.to_le_bytes());
        clx.extend_from_slice(&0u32.to_le_bytes());
        clx.extend_from_slice(&chars.to_le_bytes());
        clx.extend_from_slice(&0u16.to_le_bytes());
        clx.extend_from_slice(&32u32.to_le_bytes());
        clx.extend_from_slice(&0u16.to_le_bytes());
        assert_eq!(
            decode_piece_table(&word, &clx, units.len()).unwrap(),
            source
        );
    }

    #[test]
    fn keeps_only_displayed_field_results() {
        let paragraphs = tokenize_story("A\u{13} HYPERLINK hidden\u{14}visible\u{15}\tB\rC");
        assert_eq!(
            paragraphs,
            vec![
                vec![
                    Token::Text("Avisible".into()),
                    Token::Tab,
                    Token::Text("B".into())
                ],
                vec![Token::Text("C".into())],
            ]
        );
    }
}
