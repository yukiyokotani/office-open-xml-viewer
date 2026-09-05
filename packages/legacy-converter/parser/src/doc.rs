//! Word Binary File (`.doc`) compatibility subset.
//!
//! The reader accepts Word 97-2003 FIB/CLX piece tables and preserves main-story
//! text, paragraphs, tabs, line breaks, page breaks, and displayed field results.
//! Section geometry, character formatting and passive inline JPEG/PNG pictures
//! are preserved. Floating drawings, revisions,
//! headers/footers, notes, and OLE are
//! deliberately not inferred. See [MS-DOC] 2.5.1 (FIB), 2.8.35 (Clx), and
//! 2.9.177 (PlcPcd). Unsupported/encrypted inputs fail closed.

use crate::cfb::CompoundFile;
use crate::ooxml::{write_package_bytes, xml_text, ROOT_RELS_DOCX};
mod border;
mod character;
mod fkp;
mod formatting;
mod paragraph;
mod pictures;
mod sections;
mod settings;
mod sprm;
mod table;
mod table_output;
mod tabs;

const FIB_IDENT: u16 = 0xa5ec;
const FIB_WORD_97: u16 = 0x00c1;
const FIB_FLAGS_OFFSET: usize = 0x0a;
const CCP_TEXT_OFFSET: usize = 0x4c;
const FC_CLX_OFFSET: usize = 0x01a2;
const LCB_CLX_OFFSET: usize = 0x01a6;
const MAX_PIECES: usize = 1_000_000;
// Resource policy independent of the caller's compressed output ZIP ceiling.
const MAX_DOCUMENT_XML_BYTES: usize = 256 * 1024 * 1024;
const MAX_MAIN_STORY_UNITS: usize = 64 * 1024 * 1024;
const MAX_STORY_CONTROLS: usize = 1_000_000;

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
    ColumnBreak,
    Picture,
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
    let default_tab_twips = settings::default_tab_twips(&word, &table)?;
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
    let story = read_story(&word, clx, ccp_text)?;
    let sections = sections::read(&word, &table, ccp_text)?;
    let data = if cfb.has_entry("Data") {
        cfb.stream("Data").map_err(unsupported)?
    } else {
        Vec::new()
    };
    let mut formatting = formatting::Formatting::read(&word, &table, &data)?;
    let mut pictures = pictures::Store::new(&data);
    let document_xml = build_formatted_document(
        &story,
        &sections,
        Some(&mut formatting),
        Some(&mut pictures),
        MAX_DOCUMENT_XML_BYTES,
    )?;
    let mut content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>"#.to_string();
    if default_tab_twips.is_some() {
        content_types.push_str(r#"<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>"#);
    }
    let media = pictures.parts();
    if !media.is_empty() {
        content_types.push_str(r#"<Default Extension="png" ContentType="image/png"/><Default Extension="jpg" ContentType="image/jpeg"/>"#);
    }
    content_types.push_str("</Types>");
    let mut parts: Vec<(String, String)> = vec![
        ("[Content_Types].xml".into(), content_types),
        ("_rels/.rels".into(), ROOT_RELS_DOCX.to_string()),
        ("word/document.xml".into(), document_xml),
    ];
    let mut relationships = String::new();
    if let Some(interval) = default_tab_twips {
        parts.push(("word/settings.xml".into(), settings::xml(interval)));
        relationships.push_str(r#"<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>"#);
    }
    relationships.push_str(&pictures.relationships());
    if !relationships.is_empty() {
        parts.push(("word/_rels/document.xml.rels".into(), format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{relationships}</Relationships>"#)));
    }
    let mut warnings = vec![
        "legacy-doc:advanced-table-formatting-and-embedded-objects-omitted".into(),
        "legacy-doc:headers-footers-and-advanced-section-properties-omitted".into(),
    ];
    if pictures.omitted {
        warnings.push("legacy-doc:unsupported-inline-pictures-omitted".into());
    }
    if default_tab_twips.is_none() {
        warnings.push("legacy-doc:missing-document-properties-default-tab-interval".into());
    }
    if formatting.unsupported_character_properties {
        warnings.push("legacy-doc:unsupported-character-properties-omitted".into());
    }
    if formatting.unsupported_paragraph_properties {
        warnings.push("legacy-doc:unsupported-paragraph-properties-omitted".into());
    }
    if formatting.unsupported_piece_properties {
        warnings.push("legacy-doc:unsupported-piece-properties-omitted".into());
    }
    if formatting.unsupported_table_properties {
        warnings.push("legacy-doc:unsupported-table-properties-omitted".into());
    }
    if formatting.missing_tables {
        warnings.push("legacy-doc:missing-formatting-tables-default-character-properties".into());
    }
    if sections.is_empty() {
        warnings.push("legacy-doc:missing-section-table-default-page-geometry".into());
    }
    if sections.iter().any(|s| s.incomplete_margins) {
        warnings.push("legacy-doc:incomplete-section-margin-defaults".into());
    }
    Ok(DocConversion {
        bytes: write_package_bytes(
            parts
                .iter()
                .map(|(name, body)| (name.as_str(), body.as_bytes()))
                .chain(media.iter().map(|(name, bytes)| (name.as_str(), *bytes))),
            max_output_bytes,
        )?,
        warnings,
    })
}

struct Piece {
    start: usize,
    end: usize,
    fc: usize,
    width: usize,
    prm: u16,
}

struct Story<'a> {
    text: String,
    pieces: Vec<Piece>,
    prcs: Vec<&'a [u8]>,
}

impl Story<'_> {
    fn position(&self, cp: usize) -> Option<(usize, usize, &Piece)> {
        let i = self
            .pieces
            .partition_point(|piece| piece.start <= cp)
            .checked_sub(1)?;
        let piece = &self.pieces[i];
        if cp >= piece.end {
            return None;
        }
        Some((i, piece.fc + (cp - piece.start) * piece.width, piece))
    }
}

#[cfg(test)]
fn decode_piece_table(word: &[u8], clx: &[u8], ccp_text: usize) -> Result<String, String> {
    Ok(read_story(word, clx, ccp_text)?.text)
}

fn read_story<'a>(word: &[u8], clx: &'a [u8], ccp_text: usize) -> Result<Story<'a>, String> {
    // Logical pieces may repeatedly reference the same physical bytes. Bound
    // decoded main-story work before copying, independently of input ZIP/CFB size.
    if ccp_text > MAX_MAIN_STORY_UNITS {
        return Err(unsupported("Word main story character budget exceeded"));
    }
    let mut offset = 0usize;
    let mut prcs = Vec::new();
    while clx.get(offset) == Some(&0x01) {
        let size = u16_at(clx, offset + 1)? as usize;
        if size > i16::MAX as usize || prcs.len() >= 32768 {
            return Err(unsupported(
                "invalid or excessive Word CLX property records",
            ));
        }
        prcs.push(
            clx.get(offset + 3..offset + 3 + size)
                .ok_or_else(|| unsupported("truncated Word CLX property record"))?,
        );
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
    let mut pieces = Vec::new();
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
        let prm = u16_at(plc, pcd + 6)?;
        if prm & 1 != 0 && (prm >> 1) as usize >= prcs.len() {
            return Err(unsupported("Word piece property index outside CLX"));
        }
        pieces.push(Piece {
            start: cp_start,
            end: cp_end.min(ccp_text),
            fc: if compressed {
                file_offset / 2
            } else {
                file_offset
            },
            width: if compressed { 1 } else { 2 },
            prm,
        });
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
    Ok(Story {
        text: output,
        pieces,
        prcs,
    })
}

#[cfg(test)]
fn tokenize_story(text: &str) -> Vec<Vec<Token>> {
    tokenize_with_fields(text, &mut Fields::default(), 0, true)
        .into_iter()
        .map(|p| p.tokens.into_iter().map(|(token, _)| token).collect())
        .collect()
}

#[derive(Default)]
struct Fields {
    results: Vec<bool>,
    hidden: usize,
}

struct Paragraph {
    tokens: Vec<(Token, usize)>,
    end_cp: usize,
    mark: char,
}

fn tokenize_with_fields(
    text: &str,
    fields: &mut Fields,
    base_cp: usize,
    trim_final: bool,
) -> Vec<Paragraph> {
    let mut paragraphs = vec![Paragraph {
        tokens: Vec::new(),
        end_cp: base_cp,
        mark: '\0',
    }];
    let mut buffered = String::new();
    let mut buffer_cp = base_cp;
    let mut cp = base_cp;
    let flush = |paragraph: &mut Paragraph, buffered: &mut String, start| {
        if !buffered.is_empty() {
            paragraph
                .tokens
                .push((Token::Text(std::mem::take(buffered)), start));
        }
    };
    for character in text.chars() {
        let paragraph = paragraphs.last_mut().expect("opening paragraph");
        paragraph.end_cp = cp;
        // A field instruction creates a gap in CPs. Flush even when it creates
        // no visible output, so subsequent text retains its physical mapping.
        if character < ' ' {
            flush(paragraph, &mut buffered, buffer_cp);
        }
        match character {
            '\u{13}' => {
                fields.results.push(false);
                fields.hidden += 1;
            }
            '\u{14}' => {
                if let Some(result) = fields.results.last_mut() {
                    if !*result {
                        fields.hidden -= 1;
                        *result = true;
                    }
                }
            }
            '\u{15}' => {
                if fields.results.pop() == Some(false) {
                    fields.hidden -= 1;
                }
            }
            _ if fields.hidden != 0 => {}
            '\r' | '\u{7}' => {
                paragraph.mark = character;
                paragraphs.push(Paragraph {
                    tokens: Vec::new(),
                    end_cp: cp + 1,
                    mark: '\0',
                });
            }
            '\t' => paragraph.tokens.push((Token::Tab, cp)),
            '\u{0b}' => paragraph.tokens.push((Token::LineBreak, cp)),
            '\u{0c}' => paragraph.tokens.push((Token::PageBreak, cp)),
            '\u{0e}' => paragraph.tokens.push((Token::ColumnBreak, cp)),
            '\u{1}' => paragraph.tokens.push((Token::Picture, cp)),
            '\u{20}'..='\u{10ffff}' => {
                if buffered.is_empty() {
                    buffer_cp = cp;
                }
                buffered.push(character);
            }
            _ => {}
        }
        cp += character.len_utf16();
    }
    flush(
        paragraphs.last_mut().expect("opening paragraph"),
        &mut buffered,
        buffer_cp,
    );
    if trim_final && paragraphs.len() > 1 && paragraphs.last().is_some_and(|p| p.tokens.is_empty())
    {
        paragraphs.pop();
    }
    paragraphs
}

#[cfg(test)]
fn build_document_xml(text: &str, sections: &[sections::Section]) -> Result<String, String> {
    build_formatted_document(
        &Story {
            text: text.into(),
            pieces: Vec::new(),
            prcs: Vec::new(),
        },
        sections,
        None,
        None,
        usize::MAX,
    )
}

fn build_formatted_document(
    story: &Story<'_>,
    sections: &[sections::Section],
    mut formatting: Option<&mut formatting::Formatting<'_>>,
    mut pictures: Option<&mut pictures::Store<'_>>,
    max_bytes: usize,
) -> Result<String, String> {
    // Every control can introduce a paragraph, token or field-stack entry.
    // Charge before constructing these arrays, not only after XML expansion.
    if story
        .text
        .bytes()
        .filter(|b| *b < 32)
        .take(MAX_STORY_CONTROLS + 1)
        .count()
        > MAX_STORY_CONTROLS
    {
        return Err(unsupported("Word story structure budget exceeded"));
    }
    let mut xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>"#,
    );
    let chunks = sections::split_story(&story.text, sections)?;
    let mut fields = Fields::default();
    for (section_index, chunk) in chunks.iter().enumerate() {
        let mut table_writer = table_output::Writer::new(max_bytes.saturating_sub(xml.len()));
        let base_cp = if section_index == 0 {
            0
        } else {
            sections[section_index - 1].end
        };
        let mut paragraphs = tokenize_with_fields(
            chunk,
            &mut fields,
            base_cp,
            section_index + 1 == chunks.len(),
        );
        if section_index + 1 < chunks.len() {
            // split_story removed this paragraph's section-break character.
            paragraphs.last_mut().expect("opening paragraph").end_cp =
                sections[section_index].end - 1;
        }
        for (paragraph_index, paragraph) in paragraphs.iter().enumerate() {
            let mut paragraph_xml = String::new();
            let xml = &mut paragraph_xml;
            let style = if let Some(f) = formatting.as_deref() {
                story
                    .position(paragraph.end_cp)
                    .map(|(_, fc, _)| f.paragraph_style(fc))
                    .transpose()?
                    .unwrap_or(0)
            } else {
                0
            };
            xml.push_str("<w:p>");
            // ECMA-376 17.6.17/18: intermediate sectPr is in the final
            // paragraph's pPr; only the last section is a direct body child.
            let section_end =
                section_index + 1 < chunks.len() && paragraph_index + 1 == paragraphs.len();
            if formatting.is_some() || section_end {
                xml.push_str("<w:pPr>");
                // CT_PPr orders paragraph-mark rPr before sectPr. The mark's
                // own character properties determine empty-paragraph metrics.
                if let Some(f) = formatting.as_deref_mut() {
                    if let Some((_, fc, piece)) = story.position(paragraph.end_cp) {
                        xml.push_str(&f.paragraph_xml(style, fc, piece.prm, &story.prcs)?);
                        xml.push_str(&f.run_xml(style, fc, piece.prm, &story.prcs)?);
                    }
                }
                if section_end {
                    xml.push_str(&sections[section_index].xml);
                }
                xml.push_str("</w:pPr>");
            }
            for (token, cp) in &paragraph.tokens {
                match token {
                    Token::Text(text) => {
                        write_text_runs(xml, text, *cp, story, style, &mut formatting, max_bytes)?;
                    }
                    Token::Picture => {
                        if let Some(store) = pictures.as_deref_mut() {
                            let mut drawing = None;
                            if let Some(f) = formatting.as_deref_mut() {
                                if let Some((_, fc, piece)) = story.position(*cp) {
                                    if let Some(offset) = f.inline_picture_location(
                                        style,
                                        fc,
                                        piece.prm,
                                        &story.prcs,
                                    )? {
                                        let content = store.drawing(offset)?;
                                        if !content.is_empty() {
                                            drawing = Some((
                                                f.run_xml(style, fc, piece.prm, &story.prcs)?,
                                                content,
                                            ));
                                        }
                                    }
                                }
                            }
                            if let Some((properties, content)) = drawing {
                                xml.push_str("<w:r>");
                                xml.push_str(&properties);
                                xml.push_str(&content);
                                xml.push_str("</w:r>");
                            } else {
                                store.omitted = true;
                            }
                        }
                    }
                    _ => {
                        xml.push_str("<w:r>");
                        if let Some(f) = formatting.as_deref_mut() {
                            if let Some((_, fc, piece)) = story.position(*cp) {
                                xml.push_str(&f.run_xml(style, fc, piece.prm, &story.prcs)?);
                            }
                        }
                        xml.push_str(match token {
                            Token::Tab => "<w:tab/>",
                            Token::LineBreak => "<w:br/>",
                            Token::PageBreak => "<w:br w:type=\"page\"/>",
                            Token::ColumnBreak => "<w:br w:type=\"column\"/>",
                            Token::Text(_) | Token::Picture => unreachable!(),
                        });
                        xml.push_str("</w:r>");
                    }
                }
                if xml.len() > max_bytes {
                    return Err("OUTPUT_TOO_LARGE".into());
                }
            }
            xml.push_str("</w:p>");
            if xml.len() > max_bytes {
                return Err("OUTPUT_TOO_LARGE".into());
            }
            let mut table_properties = table::Properties::default();
            if let Some(f) = formatting.as_deref_mut() {
                if let Some((_, fc, piece)) = story.position(paragraph.end_cp) {
                    table_properties = f.table_properties(fc, piece.prm, &story.prcs)?;
                }
            }
            if section_end && table_properties.depth()? != 0 {
                return Err(unsupported("Word section break inside table"));
            }
            table_writer.push(table_properties, paragraph.mark, paragraph_xml)?;
        }
        xml.push_str(&table_writer.finish()?);
    }
    if let Some(last) = sections.last() {
        xml.push_str(&last.xml);
    } else {
        // Existing compatibility policy when no section table is available;
        // explicitly warned, not inferred from the file's name or content.
        xml.push_str("<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/></w:sectPr>");
    }
    xml.push_str("</w:body></w:document>");
    if xml.len() > max_bytes {
        return Err("OUTPUT_TOO_LARGE".into());
    }
    Ok(xml)
}

fn write_text_runs(
    xml: &mut String,
    text: &str,
    mut cp: usize,
    story: &Story<'_>,
    style: usize,
    formatting: &mut Option<&mut formatting::Formatting<'_>>,
    max_bytes: usize,
) -> Result<(), String> {
    let mut start = 0;
    let mut key = None;
    let mut properties = String::new();
    let write = |xml: &mut String, text: &str, properties: &str| -> Result<(), String> {
        // Check expansion before allocating an escaped copy of a potentially
        // document-sized run. Also bound the raw part before ZIP construction.
        let escaped_len = text
            .chars()
            .try_fold(0usize, |length, c| {
                length.checked_add(match c {
                    '&' => 5,
                    '<' | '>' => 4,
                    _ => c.len_utf8(),
                })
            })
            .ok_or("OUTPUT_TOO_LARGE")?;
        let tags_len = "<w:r><w:t xml:space=\"preserve\"></w:t></w:r>".len();
        let extra = escaped_len
            .checked_add(properties.len())
            .and_then(|n| n.checked_add(tags_len))
            .ok_or("OUTPUT_TOO_LARGE")?;
        if extra > max_bytes.saturating_sub(xml.len()) {
            return Err("OUTPUT_TOO_LARGE".into());
        }
        xml.push_str("<w:r>");
        xml.push_str(properties);
        xml.push_str("<w:t xml:space=\"preserve\">");
        xml.push_str(&xml_text(text));
        xml.push_str("</w:t></w:r>");
        Ok(())
    };
    for (byte, character) in text.char_indices() {
        if let Some(f) = formatting.as_deref_mut() {
            let (piece_index, fc, piece) = story
                .position(cp)
                .ok_or_else(|| unsupported("Word visible character outside piece table"))?;
            let next_key = Some((piece_index, f.characters.at(fc).map(|(id, _)| id)));
            if key != next_key {
                if byte > start {
                    write(xml, &text[start..byte], &properties)?;
                }
                if xml.len() > max_bytes {
                    return Err("OUTPUT_TOO_LARGE".into());
                }
                properties = f.run_xml(style, fc, piece.prm, &story.prcs)?;
                key = next_key;
                start = byte;
            }
        }
        cp += character.len_utf16();
    }
    if start < text.len() {
        write(xml, &text[start..], &properties)?;
    }
    Ok(())
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

    fn formatted_fixture(
        pieces: &[(&str, usize, bool)],
        ranges: &[u32],
        properties: &[&[u8]],
    ) -> String {
        formatted_fixture_kind(pieces, ranges, properties, false)
    }

    fn formatted_fixture_kind(
        pieces: &[(&str, usize, bool)],
        ranges: &[u32],
        properties: &[&[u8]],
        paragraph: bool,
    ) -> String {
        let mut word = vec![0; 4096];
        let mut clx = vec![2];
        clx.extend(((pieces.len() * 12 + 4) as u32).to_le_bytes());
        let mut cp = 0u32;
        clx.extend(cp.to_le_bytes());
        for (text, _, _) in pieces {
            cp += text.encode_utf16().count() as u32;
            clx.extend(cp.to_le_bytes());
        }
        for (text, fc, compressed) in pieces {
            clx.extend([0, 0]);
            clx.extend(
                (if *compressed {
                    (*fc as u32 * 2) | 0x4000_0000
                } else {
                    *fc as u32
                })
                .to_le_bytes(),
            );
            clx.extend([0, 0]);
            if *compressed {
                word[*fc..*fc + text.len()].copy_from_slice(text.as_bytes());
            } else {
                for (i, unit) in text.encode_utf16().enumerate() {
                    word[*fc + i * 2..*fc + i * 2 + 2].copy_from_slice(&unit.to_le_bytes());
                }
            }
        }
        let mut table = Vec::new();
        table.extend(ranges[0].to_le_bytes());
        table.extend(ranges.last().unwrap().to_le_bytes());
        table.extend(1u32.to_le_bytes());
        let field = if paragraph { 0x106 } else { 0xfe };
        word[field..field + 4].copy_from_slice(&12u32.to_le_bytes());
        let page = &mut word[512..1024];
        for (i, fc) in ranges.iter().enumerate() {
            page[i * 4..i * 4 + 4].copy_from_slice(&fc.to_le_bytes());
        }
        let mut offset = 128;
        for (i, bytes) in properties.iter().enumerate() {
            if bytes.is_empty() {
                continue;
            }
            page[ranges.len() * 4 + i * if paragraph { 13 } else { 1 }] = (offset / 2) as u8;
            let header = if paragraph && bytes.len() % 2 == 0 {
                page[offset + 1] = (bytes.len() / 2) as u8;
                2
            } else {
                page[offset] = if paragraph {
                    bytes.len().div_ceil(2)
                } else {
                    bytes.len()
                } as u8;
                1
            };
            page[offset + header..offset + header + bytes.len()].copy_from_slice(bytes);
            offset += (bytes.len() + header + 1) & !1;
        }
        page[511] = properties.len() as u8;
        let story = super::read_story(&word, &clx, cp as usize).unwrap();
        let mut f = super::formatting::Formatting::read(&word, &table, &[]).unwrap();
        super::build_formatted_document(&story, &[], Some(&mut f), None, usize::MAX).unwrap()
    }

    #[test]
    fn restores_table_using_the_row_marks_physical_papx() {
        let xml = formatted_fixture_kind(
            &[("A\u{7}B\u{7}\u{7}\r", 1500, true)],
            &[1500, 1502, 1504, 1505, 1506],
            &[
                &[0, 0, 0x16, 0x24, 1],
                &[0, 0, 0x16, 0x24, 1],
                &[
                    0, 0, 0x16, 0x24, 1, 0x17, 0x24, 1, 0x21, 0x76, 0, 2, 0xe8, 3,
                ],
                &[],
            ],
            true,
        );
        assert_eq!(xml.matches("<w:tbl>").count(), 1);
        assert_eq!(xml.matches("<w:tc>").count(), 2);
        assert_eq!(xml.matches("<w:p>").count(), 3);
        assert_eq!(xml.matches("<w:gridCol w:w=\"1000\"/>").count(), 2);
        assert!(xml.contains(">A</w:t>"));
        assert!(xml.contains(">B</w:t>"));
    }

    #[test]
    fn formats_reordered_unicode_and_compressed_pieces_by_physical_offsets() {
        let xml = formatted_fixture(
            &[("A\r", 1600, false), ("B\r", 1500, true)],
            &[1500, 1502, 1600, 1604],
            &[&[0x35, 8, 1], &[], &[0x43, 0x4a, 40, 0]],
        );
        assert!(xml.contains("<w:sz w:val=\"40\"/></w:rPr><w:t xml:space=\"preserve\">A</w:t>"));
        assert!(xml.contains(
            "<w:b w:val=\"1\"/><w:sz w:val=\"20\"/></w:rPr><w:t xml:space=\"preserve\">B</w:t>"
        ));
        assert!(xml.find(">A</w:t>").unwrap() < xml.find(">B</w:t>").unwrap());
    }

    #[test]
    fn field_gaps_and_surrogate_pairs_do_not_shift_character_properties() {
        let text = "😀\u{13}HIDDEN\u{14}B\u{15}C\r";
        let end = 1500 + text.encode_utf16().count() as u32 * 2;
        let xml = formatted_fixture(
            &[(text, 1500, false)],
            &[1500, 1520, 1522, end],
            &[&[], &[0x35, 8, 1], &[]],
        );
        assert!(!xml.contains("HIDDEN"));
        assert!(xml.contains(
            "<w:b w:val=\"1\"/><w:sz w:val=\"20\"/></w:rPr><w:t xml:space=\"preserve\">B</w:t>"
        ));
        assert!(
            xml.contains("<w:rPr><w:sz w:val=\"20\"/></w:rPr><w:t xml:space=\"preserve\">C</w:t>")
        );
        assert!(xml.contains('😀'));
    }

    #[test]
    fn empty_paragraphs_keep_the_marks_font_size_without_inventing_text() {
        let xml = formatted_fixture(
            &[("\r", 1500, false)],
            &[1500, 1502],
            &[&[0x43, 0x4a, 36, 0]],
        );
        assert!(xml.contains("<w:rPr><w:sz w:val=\"36\"/></w:rPr></w:pPr></w:p>"));
        assert!(!xml.contains("<w:t"));
    }

    #[test]
    fn limits_escaped_runs_and_empty_paragraph_expansion_before_packaging() {
        assert!(super::read_story(&[], &[], super::MAX_MAIN_STORY_UNITS + 1)
            .err()
            .unwrap()
            .contains("character budget"));
        for text in ["&".repeat(1000), "\r".repeat(1000)] {
            let story = super::Story {
                text,
                pieces: vec![],
                prcs: vec![],
            };
            assert_eq!(
                super::build_formatted_document(&story, &[], None, None, 1024).unwrap_err(),
                "OUTPUT_TOO_LARGE"
            );
        }
        let story = super::Story {
            text: "\t".repeat(super::MAX_STORY_CONTROLS + 1),
            pieces: vec![],
            prcs: vec![],
        };
        assert!(
            super::build_formatted_document(&story, &[], None, None, usize::MAX)
                .unwrap_err()
                .contains("structure budget")
        );
    }

    #[test]
    fn nested_fields_do_not_rescan_the_entire_stack_per_character() {
        let text = format!(
            "{}hidden{}shown",
            "\u{13}".repeat(20_000),
            "\u{15}".repeat(20_000)
        );
        assert_eq!(
            tokenize_story(&text),
            vec![vec![Token::Text("shown".into())]]
        );
    }

    #[test]
    fn keeps_the_empty_paragraph_that_owns_a_section_break() {
        let sections = [
            super::sections::Section {
                end: 3,
                xml: "<w:sectPr/>".into(),
                incomplete_margins: false,
            },
            super::sections::Section {
                end: 5,
                xml: "<w:sectPr/>".into(),
                incomplete_margins: false,
            },
        ];
        let xml = super::build_document_xml("A\r\u{c}B\r", &sections).unwrap();
        assert!(xml.contains("</w:r></w:p><w:p><w:pPr><w:sectPr/></w:pPr></w:p>"));
    }

    #[test]
    fn writes_section_properties_at_their_ooxml_positions_without_extra_page_breaks() {
        let sections = [
            super::sections::Section {
                end: 2,
                xml: "<w:sectPr><w:type w:val=\"continuous\"/></w:sectPr>".into(),
                incomplete_margins: false,
            },
            super::sections::Section {
                end: 6,
                xml: "<w:sectPr/>".into(),
                incomplete_margins: false,
            },
        ];
        let xml = super::build_document_xml("A\u{c}B\u{c}\u{e}C", &sections).unwrap();
        assert!(xml.contains("<w:p><w:pPr><w:sectPr>"));
        assert!(xml.ends_with("</w:p><w:sectPr/></w:body></w:document>"));
        assert_eq!(xml.matches("w:type=\"page\"").count(), 1);
        assert_eq!(xml.matches("w:type=\"column\"").count(), 1);
    }

    #[test]
    fn field_instructions_remain_hidden_across_section_boundaries() {
        let sections = [
            super::sections::Section {
                end: 3,
                xml: "<w:sectPr/>".into(),
                incomplete_margins: false,
            },
            super::sections::Section {
                end: 8,
                xml: "<w:sectPr/>".into(),
                incomplete_margins: false,
            },
        ];
        let xml = super::build_document_xml("\u{13}X\u{c}Y\u{14}OK\u{15}", &sections).unwrap();
        assert!(!xml.contains('X') && !xml.contains('Y'));
        assert!(xml.contains("OK"));
    }

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
                    Token::Text("A".into()),
                    Token::Text("visible".into()),
                    Token::Tab,
                    Token::Text("B".into())
                ],
                vec![Token::Text("C".into())],
            ]
        );
    }
}
