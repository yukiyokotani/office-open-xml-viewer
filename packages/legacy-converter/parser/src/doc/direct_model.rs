//! Internal direct DOC body producer. This is intentionally not a public route:
//! unsupported body owners fail closed while incremental model coverage lands.

use super::{unsupported, AcquiredDoc, Fields, Token};
use docx_model::paragraph_breaks::{visit_para_on_page_breaks, ParaPiece};
use docx_model::{
    BodyElement, BreakType, DocRun, Document, DocumentSettings, DocumentTypographySettingsWire,
};

mod payload;

pub(super) fn build(mut facts: AcquiredDoc<'_>, max_bytes: usize) -> Result<Document, String> {
    if max_bytes == 0 {
        return Err("OUTPUT_TOO_LARGE".into());
    }
    // MS-DOC 2.3.3 / 2.8.22: the first six header-document stories are
    // footnote/endnote separators, not page headers. Only authored nonempty
    // header/footer ranges produce entries; an explicitly blank paragraph does.
    if facts
        .headers
        .as_ref()
        .is_some_and(|headers| !headers.entries.is_empty())
        || facts.note_stories.iter().any(Option::is_some)
    {
        return Err(unsupported(
            "direct DOC model does not yet support headers or notes",
        ));
    }
    if facts.formatting.missing_tables {
        return Err(unsupported(
            "direct DOC model requires complete formatting tables",
        ));
    }

    let even_and_odd = facts
        .document_settings
        .as_ref()
        .is_some_and(|settings| settings.even_and_odd_headers);
    let section = match facts.sections.last() {
        Some(section) => section.project_final(facts.sections.len() - 1, even_and_odd)?,
        None => {
            return Err(unsupported(
                "direct DOC model requires resolved section facts",
            ));
        }
    };
    let settings = facts
        .document_settings
        .as_ref()
        .map(|settings| DocumentSettings {
            default_tab_stop: Some(f64::from(settings.default_tab_twips) / 20.0),
            ..DocumentSettings::default()
        });
    let document_typography_settings = Some(DocumentTypographySettingsWire {
        normal_style_font_size_pt: facts.formatting.direct_normal_style_font_size_pt()?,
    });

    let mut budget = ModelBudget::new(max_bytes);
    budget.charge(std::mem::size_of::<Document>())?;
    budget.charge(payload::section(&section)?)?;
    if facts
        .story
        .text
        .bytes()
        .filter(|byte| *byte < 32)
        .take(super::MAX_STORY_CONTROLS + 1)
        .count()
        > super::MAX_STORY_CONTROLS
    {
        return Err(unsupported("Word story structure budget exceeded"));
    }
    let mut body = Vec::new();
    let chunks = super::sections::split_story(&facts.story.text, &facts.sections)?;
    let mut fields = Fields::default();
    for (section_index, chunk) in chunks.iter().enumerate() {
        let ending = (section_index + 1 < chunks.len())
            .then(|| facts.sections[section_index].project_ending(section_index))
            .transpose()?;
        let base_cp = if section_index == 0 {
            0
        } else {
            facts.sections[section_index - 1].end
        };
        let mut paragraphs = super::tokenize_with_fields(
            chunk,
            &mut fields,
            base_cp,
            section_index + 1 == chunks.len(),
        );
        if section_index + 1 < chunks.len() {
            // split_story consumed the section-break form feed. The paragraph
            // mark formatting remains owned by the preceding physical CP.
            paragraphs.last_mut().expect("opening paragraph").end_cp =
                facts.sections[section_index].end - 1;
        }

        let paragraph_count = paragraphs.len();
        for (paragraph_index, source) in paragraphs.into_iter().enumerate() {
            if source.mark == '\u{7}' {
                return Err(unsupported("direct DOC model does not yet support tables"));
            }
            let (_, mark_fc, mark_piece) = facts
                .story
                .position(source.end_cp)
                .ok_or_else(|| unsupported("Word paragraph mark outside piece table"))?;
            let style = facts.formatting.paragraph_style(mark_fc)?;
            if facts
                .formatting
                .table_properties(mark_fc, mark_piece.prm, &facts.story.prcs)?
                .depth()?
                != 0
            {
                return Err(unsupported("direct DOC model does not yet support tables"));
            }
            let direct = facts.formatting.direct_paragraph(
                style,
                mark_fc,
                mark_piece.prm,
                &facts.story.prcs,
            )?;
            if direct.numbering.is_some() {
                return Err(unsupported(
                    "direct DOC model does not yet support numbered paragraphs",
                ));
            }
            let mut paragraph = direct.paragraph;
            budget.paragraph(&paragraph)?;

            for (token, cp) in source.tokens {
                match token {
                    Token::Text(text) => {
                        super::visit_text_runs(
                            &text,
                            cp,
                            &facts.story,
                            &mut Some(&mut facts.formatting),
                            |formatting, fc, prm| {
                                formatting.direct_text_run(
                                    style,
                                    fc,
                                    prm,
                                    &facts.story.prcs,
                                    String::new(),
                                )
                            },
                            |part, run| {
                                if let Some(mut run) = run.flatten() {
                                    budget.text(&mut paragraph.runs, &mut run, part)?;
                                }
                                Ok(())
                            },
                        )?;
                    }
                    Token::Tab => {
                        push_control_text(
                            &mut paragraph,
                            &facts.story,
                            &mut facts.formatting,
                            style,
                            cp,
                            "\t",
                            &mut budget,
                        )?;
                    }
                    Token::LineBreak => {
                        budget.push(
                            &mut paragraph.runs,
                            DocRun::Break {
                                break_type: BreakType::Line,
                            },
                        )?;
                    }
                    Token::PageBreak | Token::ColumnBreak => {
                        budget.push(
                            &mut paragraph.runs,
                            DocRun::Break {
                                break_type: if matches!(token, Token::PageBreak) {
                                    BreakType::Page
                                } else {
                                    BreakType::Column
                                },
                            },
                        )?;
                    }
                    Token::Picture | Token::FloatingPicture => {
                        return Err(unsupported(
                            "direct DOC model does not yet support pictures",
                        ));
                    }
                    Token::NoteMarker | Token::NoteReference(_) => {
                        return Err(unsupported(
                            "direct DOC model does not yet support note content",
                        ));
                    }
                    Token::FieldBegin(_) | Token::FieldEnd => {
                        return Err(unsupported(
                            "direct DOC model does not retain field structures yet",
                        ));
                    }
                }
            }
            // Match the shared parser's boundary contract on projected runs,
            // after vanish filtering but before chunk visibility normalization.
            // A whitespace/line-break run is not the same as an absent run.
            match paragraph.runs.as_slice() {
                [DocRun::Break {
                    break_type: BreakType::Page,
                }] => {
                    let subsumed = paragraph_index + 1 == paragraph_count
                        && ending.as_ref().is_some_and(|ending| {
                            !matches!(ending.kind.as_str(), "continuous" | "nextColumn")
                        });
                    if !subsumed {
                        budget.push(
                            &mut body,
                            BodyElement::PageBreak {
                                parity: None,
                                same_paragraph_as_previous: None,
                            },
                        )?;
                    }
                }
                [DocRun::Break {
                    break_type: BreakType::Column,
                }] => {
                    budget.push(&mut body, BodyElement::ColumnBreak)?;
                }
                _ => visit_para_on_page_breaks(paragraph, |piece| {
                    let element = match piece {
                        ParaPiece::Para(paragraph) => {
                            // Runs move through the normalizer; their text/font
                            // payload was already charged during acquisition.
                            // Charge output metadata and the new vector backing
                            // before retaining each piece, and stop on failure.
                            budget.normalized_paragraph(&paragraph)?;
                            BodyElement::Paragraph(Box::new(paragraph))
                        }
                        ParaPiece::PageBreak {
                            same_paragraph_as_previous,
                        } => BodyElement::PageBreak {
                            parity: None,
                            same_paragraph_as_previous: same_paragraph_as_previous.then_some(true),
                        },
                        ParaPiece::ColumnBreak => BodyElement::ColumnBreak,
                    };
                    budget.push(&mut body, element)
                })?,
            }
        }

        if let Some(ending) = ending {
            budget.charge(payload::ending_section(
                &ending.kind,
                ending.columns.as_ref(),
                ending.page_num_type.as_ref(),
                &ending.text_direction,
                ending.geom.as_ref(),
                ending.placement.as_ref(),
            )?)?;
            budget.push(
                &mut body,
                BodyElement::SectionBreak {
                    kind: ending.kind,
                    columns: ending.columns,
                    headers: Box::default(),
                    footers: Box::default(),
                    title_page: ending.title_page,
                    geom: Some(ending.geom),
                    page_num_type: ending.page_num_type,
                    text_direction: ending.text_direction,
                    section_placement: ending.placement,
                },
            )?;
        }
    }

    if facts.formatting.unsupported_character_properties
        || facts.formatting.unsupported_paragraph_properties
        || facts.formatting.unsupported_piece_properties
        || facts.formatting.unsupported_table_properties
    {
        return Err(unsupported(
            "direct DOC model encountered unsupported formatting",
        ));
    }
    if facts.pictures.omitted || facts.floating.omitted {
        return Err(unsupported(
            "direct DOC model encountered omitted drawing content",
        ));
    }

    Ok(Document {
        section,
        body,
        settings,
        document_typography_settings,
        ..Document::default()
    })
}

fn push_control_text(
    paragraph: &mut docx_model::DocParagraph,
    story: &super::Story<'_>,
    formatting: &mut super::formatting::Formatting<'_>,
    style: usize,
    cp: usize,
    text: &str,
    budget: &mut ModelBudget,
) -> Result<bool, String> {
    let (_, fc, piece) = story
        .position(cp)
        .ok_or_else(|| unsupported("Word control outside piece table"))?;
    if let Some(mut run) =
        formatting.direct_text_run(style, fc, piece.prm, &story.prcs, String::new())?
    {
        budget.text(&mut paragraph.runs, &mut run, text)?;
        return Ok(true);
    }
    Ok(false)
}

/// Cumulative admitted model storage, not serialized size or an RSS estimate.
/// Count struct storage, owned dynamic payload and vector capacity growth.
/// Temporary paragraph templates are charged without refunds. Acquisition and
/// tokenization retain their separate input/structure limits; this budget stops
/// CHPX fragmentation from multiplying retained model allocations unchecked.
struct ModelBudget {
    remaining_bytes: usize,
}

impl ModelBudget {
    fn new(remaining_bytes: usize) -> Self {
        Self { remaining_bytes }
    }

    fn charge(&mut self, required: usize) -> Result<(), String> {
        self.remaining_bytes = self
            .remaining_bytes
            .checked_sub(required)
            .ok_or_else(|| "OUTPUT_TOO_LARGE".to_string())?;
        Ok(())
    }

    fn paragraph(&mut self, paragraph: &docx_model::DocParagraph) -> Result<(), String> {
        self.charge(std::mem::size_of::<docx_model::DocParagraph>())?;
        self.charge(payload::paragraph(paragraph)?)
    }

    fn normalized_paragraph(&mut self, paragraph: &docx_model::DocParagraph) -> Result<(), String> {
        self.charge(std::mem::size_of::<docx_model::DocParagraph>())?;
        self.charge(payload::paragraph_metadata(paragraph)?)?;
        self.charge(
            paragraph
                .runs
                .capacity()
                .checked_mul(std::mem::size_of::<DocRun>())
                .ok_or("OUTPUT_TOO_LARGE")?,
        )
    }

    fn push<T>(&mut self, values: &mut Vec<T>, value: T) -> Result<(), String> {
        if values.len() == values.capacity() {
            let old = values.capacity();
            let growth = old.max(1);
            self.charge(
                growth
                    .checked_mul(std::mem::size_of::<T>())
                    .ok_or("OUTPUT_TOO_LARGE")?,
            )?;
            values
                .try_reserve_exact(growth)
                .map_err(|_| "OUTPUT_TOO_LARGE".to_string())?;
            // Allocators may reserve more than requested. Charge that excess
            // before retaining another value; failures drop the whole model.
            let expected = old.checked_add(growth).ok_or("OUTPUT_TOO_LARGE")?;
            self.charge(
                values
                    .capacity()
                    .saturating_sub(expected)
                    .checked_mul(std::mem::size_of::<T>())
                    .ok_or("OUTPUT_TOO_LARGE")?,
            )?;
        }
        values.push(value);
        Ok(())
    }

    fn text(
        &mut self,
        runs: &mut Vec<DocRun>,
        run: &mut docx_model::TextRun,
        text: &str,
    ) -> Result<(), String> {
        self.charge(std::mem::size_of::<docx_model::TextRun>())?;
        self.charge(payload::text_run(run)?)?;
        self.charge(text.len())?;
        run.text
            .try_reserve_exact(text.len())
            .map_err(|_| "OUTPUT_TOO_LARGE".to_string())?;
        self.charge(run.text.capacity().saturating_sub(text.len()))?;
        run.text.push_str(text);
        self.push(runs, DocRun::Text(Box::new(std::mem::take(run))))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cfb::{test_support::build_cfb, CompoundFile};

    fn source(text: &str) -> Vec<u8> {
        let units: Vec<u16> = text.encode_utf16().collect();
        source_with_sections_and_header(text, &[(units.len(), 2, 12240, 15840, 1, 720)], None)
    }

    fn source_with_sections(text: &str, sections: &[(usize, u8, u16, u16, u16, u16)]) -> Vec<u8> {
        source_with_sections_and_header(text, sections, None)
    }

    fn source_with_sections_and_header(
        text: &str,
        sections: &[(usize, u8, u16, u16, u16, u16)],
        authored_blank_header: Option<bool>,
    ) -> Vec<u8> {
        source_with_typography(text, sections, authored_blank_header, None, None)
    }

    fn source_with_typography(
        text: &str,
        sections: &[(usize, u8, u16, u16, u16, u16)],
        authored_blank_header: Option<bool>,
        normal_hps: Option<u16>,
        body_hps: Option<u16>,
    ) -> Vec<u8> {
        let main_units = text.encode_utf16().count();
        let header = match authored_blank_header {
            None => "",
            Some(false) => "\r\r",
            Some(true) => "\r\r\r\r",
        };
        let units: Vec<u16> = text.encode_utf16().chain(header.encode_utf16()).collect();
        let text_offset = 0x400usize;
        let mut word = vec![0u8; text_offset + units.len() * 2];
        word[0..2].copy_from_slice(&0xa5ecu16.to_le_bytes());
        word[2..4].copy_from_slice(&0x00c1u16.to_le_bytes());
        word[6..8].copy_from_slice(&1033u16.to_le_bytes());
        word[0x4c..0x50].copy_from_slice(&(main_units as u32).to_le_bytes());
        word[0x54..0x58].copy_from_slice(&(header.encode_utf16().count() as u32).to_le_bytes());
        word[0x1a2..0x1a6].copy_from_slice(&0u32.to_le_bytes());
        word[0x1a6..0x1aa].copy_from_slice(&21u32.to_le_bytes());
        for (index, unit) in units.iter().enumerate() {
            word[text_offset + index * 2..text_offset + index * 2 + 2]
                .copy_from_slice(&unit.to_le_bytes());
        }
        let mut table = vec![0x02];
        table.extend_from_slice(&16u32.to_le_bytes());
        table.extend_from_slice(&0u32.to_le_bytes());
        table.extend_from_slice(&(units.len() as u32).to_le_bytes());
        table.extend_from_slice(&0u16.to_le_bytes());
        table.extend_from_slice(&(text_offset as u32).to_le_bytes());
        table.extend_from_slice(&0u16.to_le_bytes());
        let section_table_offset = table.len();
        let mut sepx_offsets = Vec::new();
        for &(_, kind, width, height, columns, spacing) in sections {
            let mut sepx = Vec::new();
            sepx.extend([0x09, 0x30, kind]);
            for (code, value) in [
                (0x9023u16, 1440u16),
                (0x9024, 1440),
                (0xb021, 1440),
                (0xb022, 1440),
                (0xb01f, width),
                (0xb020, height),
                (0x500b, columns - 1),
                (0x900c, spacing),
            ] {
                sepx.extend(code.to_le_bytes());
                sepx.extend(value.to_le_bytes());
            }
            sepx_offsets.push(word.len());
            word.extend((sepx.len() as u16).to_le_bytes());
            word.extend(sepx);
        }
        let section_table_size = 4 + sections.len() * 16;
        let mut section_table = vec![0; section_table_size];
        for (index, &(end, ..)) in sections.iter().enumerate() {
            section_table[(index + 1) * 4..(index + 2) * 4]
                .copy_from_slice(&(end as u32).to_le_bytes());
            let sed = (sections.len() + 1) * 4 + index * 12;
            section_table[sed + 2..sed + 6]
                .copy_from_slice(&(sepx_offsets[index] as u32).to_le_bytes());
        }
        table.extend(section_table);
        word[0xca..0xce].copy_from_slice(&(section_table_offset as u32).to_le_bytes());
        word[0xce..0xd2].copy_from_slice(&(section_table_size as u32).to_le_bytes());

        if let Some(authored) = authored_blank_header {
            let mut hdd = Vec::new();
            for index in 0..14 {
                let cp: u32 = match index {
                    0 | 13 => 0,
                    1..=6 => 1,
                    _ if authored => 3,
                    _ => 1,
                };
                hdd.extend(cp.to_le_bytes());
            }
            append_table_part(&mut word, &mut table, 0xf2, &hdd);
        }

        let mut font_table = vec![1, 0, 0, 0];
        let mut font = vec![0; 39];
        for unit in "Test Font\0".encode_utf16() {
            font.extend(unit.to_le_bytes());
        }
        font_table.push(font.len() as u8);
        font_table.extend(font);
        append_table_part(&mut word, &mut table, 0x112, &font_table);

        let mut style_header = vec![0; 18];
        style_header[0..2].copy_from_slice(&15u16.to_le_bytes());
        style_header[2..4].copy_from_slice(&10u16.to_le_bytes());
        for offset in [12usize, 14, 16] {
            style_header[offset..offset + 2].copy_from_slice(&0u16.to_le_bytes());
        }
        let mut stylesheet = Vec::new();
        stylesheet.extend(18u16.to_le_bytes());
        stylesheet.extend(style_header);
        let mut normal = vec![0; 14];
        normal[2..4].copy_from_slice(&0xfff1u16.to_le_bytes());
        if let Some(size) = normal_hps {
            normal[4..6].copy_from_slice(&2u16.to_le_bytes());
            normal.extend([2, 0, 0, 0]);
            normal.extend([4, 0, 0x43, 0x4a]);
            normal.extend(size.to_le_bytes());
        }
        stylesheet.extend((normal.len() as u16).to_le_bytes());
        stylesheet.append(&mut normal);
        for _ in 1..15 {
            stylesheet.extend(0u16.to_le_bytes());
        }
        append_table_part(&mut word, &mut table, 0xa2, &stylesheet);

        for fib_offset in [0xfausize, 0x102] {
            while !word.len().is_multiple_of(512) {
                word.push(0);
            }
            let page_number = word.len() / 512;
            let mut page = vec![0; 512];
            page[0..4].copy_from_slice(&(text_offset as u32).to_le_bytes());
            page[4..8].copy_from_slice(&((text_offset + units.len() * 2) as u32).to_le_bytes());
            if fib_offset == 0xfa {
                if let Some(size) = body_hps {
                    page[8] = 32;
                    page[64..69].copy_from_slice(&[4, 0x43, 0x4a, size as u8, (size >> 8) as u8]);
                }
            }
            page[511] = 1;
            word.extend(page);
            let mut bte = Vec::new();
            bte.extend((text_offset as u32).to_le_bytes());
            bte.extend(((text_offset + units.len() * 2) as u32).to_le_bytes());
            bte.extend((page_number as u32).to_le_bytes());
            append_table_part(&mut word, &mut table, fib_offset, &bte);
        }
        build_cfb(&[("WordDocument", word), ("0Table", table)])
    }

    fn append_table_part(word: &mut [u8], table: &mut Vec<u8>, fib_offset: usize, part: &[u8]) {
        word[fib_offset..fib_offset + 4].copy_from_slice(&(table.len() as u32).to_le_bytes());
        word[fib_offset + 4..fib_offset + 8].copy_from_slice(&(part.len() as u32).to_le_bytes());
        table.extend(part);
    }

    fn hide_first_utf16_unit(bytes: &[u8]) -> Vec<u8> {
        let cfb = CompoundFile::open(bytes).unwrap();
        let mut word = cfb.stream("WordDocument").unwrap();
        let table = cfb.stream("0Table").unwrap();
        let bte = u32::from_le_bytes(word[0xfa..0xfe].try_into().unwrap()) as usize;
        let page_number = u32::from_le_bytes(table[bte + 8..bte + 12].try_into().unwrap()) as usize;
        let page = &mut word[page_number * 512..(page_number + 1) * 512];
        let end = u32::from_le_bytes(page[4..8].try_into().unwrap());
        page[0..4].copy_from_slice(&0x400u32.to_le_bytes());
        page[4..8].copy_from_slice(&0x402u32.to_le_bytes());
        page[8..12].copy_from_slice(&end.to_le_bytes());
        page[12] = 32;
        page[13] = 0;
        page[64..68].copy_from_slice(&[3, 0x3c, 0x08, 1]);
        page[511] = 2;
        build_cfb(&[("WordDocument", word), ("0Table", table)])
    }

    fn make_normal_style_self_referential(bytes: &[u8]) -> Vec<u8> {
        let cfb = CompoundFile::open(bytes).unwrap();
        let word = cfb.stream("WordDocument").unwrap();
        let mut table = cfb.stream("0Table").unwrap();
        let styles = u32::from_le_bytes(word[0xa2..0xa6].try_into().unwrap()) as usize;
        // STSHI is prefixed by its 2-byte size. The first STD follows the
        // 18-byte header and its own 2-byte size; offset 2 is sti/base.
        table[styles + 24..styles + 26].copy_from_slice(&1u16.to_le_bytes());
        build_cfb(&[("WordDocument", word), ("0Table", table)])
    }

    #[test]
    fn source_story_projects_directly_with_controls_and_cached_field_result() {
        let bytes = source("A\tB\u{b}C\r\u{13}PAGE\u{14}42\u{15}\r\u{c}\r\u{e}\rA\u{c}B\u{e}C\r");
        let cfb = CompoundFile::open(&bytes).unwrap();
        let direct = super::super::direct_model(&cfb, 1024 * 1024).unwrap();
        let converted = super::super::convert(&cfb, 1024 * 1024).unwrap();
        let expected: serde_json::Value =
            serde_json::from_str(&docx_parser::parse_docx_native(&converted.bytes).unwrap())
                .unwrap();
        let actual = serde_json::to_value(&direct).unwrap();
        let mut actual_body = actual["body"].clone();
        let mut expected_body = expected["body"].clone();
        for body in [&mut actual_body, &mut expected_body] {
            for value in body.as_array_mut().unwrap() {
                if value["type"] == "paragraph" {
                    value.as_object_mut().unwrap().remove("styleId");
                }
            }
        }
        assert_eq!(actual_body, expected_body);
        assert_eq!(actual["section"], expected["section"]);
    }

    #[test]
    fn document_typography_uses_resolved_normal_style_size_not_body_chpx() {
        let sections = [(5, 2, 12_240, 15_840, 1, 720)];
        for (normal_hps, body_hps, expected_pt) in [
            (None, Some(44), 10.0),
            (Some(2), Some(44), 1.0),
            (Some(22), Some(48), 11.0),
            (Some(3276), Some(20), 1638.0),
        ] {
            let bytes = source_with_typography("Body\r", &sections, None, normal_hps, body_hps);
            let document =
                super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1024 * 1024)
                    .unwrap();
            assert_eq!(
                document
                    .document_typography_settings
                    .unwrap()
                    .normal_style_font_size_pt,
                expected_pt,
                "normal={normal_hps:?}, body={body_hps:?}"
            );
            let BodyElement::Paragraph(paragraph) = &document.body[0] else {
                panic!("expected body paragraph");
            };
            let DocRun::Text(run) = &paragraph.runs[0] else {
                panic!("expected independently formatted body text");
            };
            assert_eq!(run.font_size, f64::from(body_hps.unwrap()) / 2.0);
        }

        let malformed = source_with_typography("Body\r", &sections, None, Some(1), Some(44));
        assert!(
            super::super::direct_model(&CompoundFile::open(&malformed).unwrap(), 1024 * 1024,)
                .is_err()
        );

        let valid = source_with_typography("Body\r", &sections, None, Some(22), Some(44));
        let cyclic = make_normal_style_self_referential(&valid);
        assert!(
            super::super::direct_model(&CompoundFile::open(&cyclic).unwrap(), 1024 * 1024,)
                .unwrap_err()
                .contains("cyclic")
        );
    }

    #[test]
    fn separator_only_header_document_is_not_an_authored_page_header() {
        let section = [(5, 2, 12_240, 15_840, 1, 720)];
        let plain = source("Body\r");
        let separators = source_with_sections_and_header("Body\r", &section, Some(false));
        let blank_header = source_with_sections_and_header("Body\r", &section, Some(true));
        let expected =
            super::super::direct_model(&CompoundFile::open(&plain).unwrap(), 1024 * 1024).unwrap();
        let actual =
            super::super::direct_model(&CompoundFile::open(&separators).unwrap(), 1024 * 1024)
                .unwrap();
        assert_eq!(
            serde_json::to_value(actual).unwrap(),
            serde_json::to_value(expected).unwrap()
        );
        assert!(super::super::direct_model(
            &CompoundFile::open(&blank_header).unwrap(),
            1024 * 1024
        )
        .unwrap_err()
        .contains("headers or notes"));
    }

    #[test]
    fn output_budget_and_unimplemented_body_owners_fail_without_a_document() {
        let bytes = source("visible\r");
        let cfb = CompoundFile::open(&bytes).unwrap();
        assert_eq!(
            super::super::direct_model(&cfb, 1).unwrap_err(),
            "OUTPUT_TOO_LARGE"
        );
        for text in ["\u{1}\r", "cell\u{7}"] {
            let bytes = source(text);
            let cfb = CompoundFile::open(&bytes).unwrap();
            assert!(super::super::direct_model(&cfb, 1024 * 1024)
                .unwrap_err()
                .starts_with("UNSUPPORTED:"));
        }
        let bytes = source(&format!("{}\r", "x".repeat(100)));
        let cfb = CompoundFile::open(&bytes).unwrap();
        assert!(super::super::direct_model(&cfb, 64 * 1024).is_ok());
    }

    #[test]
    fn vector_growth_is_amortized_and_rejected_before_retaining_a_value() {
        let mut budget = ModelBudget::new(4 * std::mem::size_of::<u64>());
        let mut values = Vec::new();
        for value in 0..4 {
            budget.push(&mut values, value as u64).unwrap();
        }
        assert_eq!(budget.remaining_bytes, 0);
        assert_eq!(values.capacity(), 4);
        assert_eq!(budget.push(&mut values, 4).unwrap_err(), "OUTPUT_TOO_LARGE");
        assert_eq!(values, [0, 1, 2, 3]);
    }

    #[test]
    fn actual_run_fragmentation_consumes_budget_not_source_character_count() {
        let mut budget = ModelBudget::new(64 * 1024);
        let mut runs = Vec::new();
        budget
            .text(
                &mut runs,
                &mut docx_model::TextRun::default(),
                &"x".repeat(8192),
            )
            .unwrap();
        assert_eq!(runs.len(), 1);

        let mut budget = ModelBudget::new(64 * 1024);
        let mut runs = Vec::new();
        let mut rejected = false;
        for _ in 0..8192 {
            let before = runs.len();
            if budget
                .text(&mut runs, &mut docx_model::TextRun::default(), "x")
                .is_err()
            {
                assert_eq!(runs.len(), before);
                rejected = true;
                break;
            }
        }
        assert!(rejected);
    }

    #[test]
    fn complete_document_has_a_deterministic_admission_boundary() {
        let bytes = source("First\u{c}second\u{e}third\r");
        let cfb = CompoundFile::open(&bytes).unwrap();
        let (mut low, mut high) = (0, 1024 * 1024);
        while low + 1 < high {
            let middle = low + (high - low) / 2;
            match super::super::direct_model(&cfb, middle) {
                Ok(_) => high = middle,
                Err(error) => {
                    assert_eq!(error, "OUTPUT_TOO_LARGE");
                    low = middle;
                }
            }
        }
        assert!(super::super::direct_model(&cfb, high).is_ok());
        assert_eq!(
            super::super::direct_model(&cfb, high - 1).unwrap_err(),
            "OUTPUT_TOO_LARGE"
        );
    }

    #[test]
    fn incomplete_acquisition_never_returns_a_successful_partial_document() {
        let bytes = source("Visible\r");
        let cfb = CompoundFile::open(&bytes).unwrap();
        for case in 0..8 {
            let result = super::super::with_acquired_doc(&cfb, |mut facts| {
                match case {
                    0 => facts.sections.clear(),
                    1 => facts
                        .sections
                        .push(super::super::sections::Section::for_test(8, 0)),
                    2 => facts.formatting.missing_tables = true,
                    3 => facts.formatting.unsupported_character_properties = true,
                    4 => facts.formatting.unsupported_paragraph_properties = true,
                    5 => facts.formatting.unsupported_piece_properties = true,
                    6 => facts.pictures.omitted = true,
                    7 => facts.floating.omitted = true,
                    _ => unreachable!(),
                }
                build(facts, 1024 * 1024)
            });
            assert!(
                result.unwrap_err().starts_with("UNSUPPORTED:"),
                "case {case}"
            );
        }
    }

    #[test]
    fn multiple_sections_preserve_utf16_boundaries_geometry_columns_and_break_kinds() {
        // First section ends after A (1 CP), supplementary 😀 (2 CP), and the
        // consumed section-break form feed (1 CP).
        let bytes = source_with_sections(
            "A😀\u{c}B\r",
            &[
                (4, 0, 11_906, 16_838, 2, 360),
                (6, 4, 15_840, 12_240, 1, 720),
            ],
        );
        let cfb = CompoundFile::open(&bytes).unwrap();
        let direct = super::super::direct_model(&cfb, 1024 * 1024).unwrap();
        let converted = super::super::convert(&cfb, 1024 * 1024).unwrap();
        let expected: serde_json::Value =
            serde_json::from_str(&docx_parser::parse_docx_native(&converted.bytes).unwrap())
                .unwrap();
        let actual = serde_json::to_value(&direct).unwrap();
        let normalize = |body: &serde_json::Value| {
            let mut body = body.clone();
            for value in body.as_array_mut().unwrap() {
                if value["type"] == "paragraph" {
                    value.as_object_mut().unwrap().remove("styleId");
                }
            }
            body
        };
        assert_eq!(normalize(&actual["body"]), normalize(&expected["body"]));
        assert_eq!(actual["section"], expected["section"]);
        assert_eq!(actual["body"][1]["kind"], "continuous");
        assert_eq!(actual["body"][1]["columns"]["count"], 2);
        assert_eq!(actual["section"]["sectionStart"], "oddPage");
    }

    #[test]
    fn lone_page_break_at_section_boundary_matches_section_kind_normalization() {
        for (kind, expected_page_breaks) in [(0, 1), (1, 1), (2, 0), (3, 0), (4, 0)] {
            // The first form feed is an authored page break. The second is the
            // section mark consumed by split_story.
            let bytes = source_with_sections(
                "\u{c}\u{c}B\r",
                &[
                    (2, kind, 12_240, 15_840, 1, 720),
                    (4, 2, 12_240, 15_840, 1, 720),
                ],
            );
            let cfb = CompoundFile::open(&bytes).unwrap();
            let direct = super::super::direct_model(&cfb, 1024 * 1024).unwrap();
            let converted = super::super::convert(&cfb, 1024 * 1024).unwrap();
            let expected: serde_json::Value =
                serde_json::from_str(&docx_parser::parse_docx_native(&converted.bytes).unwrap())
                    .unwrap();
            let actual = serde_json::to_value(&direct).unwrap();
            let normalize = |body: &serde_json::Value| {
                let mut body = body.clone();
                for value in body.as_array_mut().unwrap() {
                    if value["type"] == "paragraph" {
                        value.as_object_mut().unwrap().remove("styleId");
                    }
                }
                body
            };
            assert_eq!(
                normalize(&actual["body"]),
                normalize(&expected["body"]),
                "section kind {kind}"
            );
            assert_eq!(
                actual["body"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .filter(|element| element["type"] == "pageBreak")
                    .count(),
                expected_page_breaks,
                "section kind {kind}"
            );
        }
    }

    #[test]
    fn standalone_page_break_before_section_ending_paragraph_is_never_suppressed() {
        for kind in 0..=4 {
            // The first form feed belongs to its own paragraph. Only the later
            // form feed terminates the section containing paragraph "A".
            let bytes = source_with_sections(
                "\u{c}\rA\u{c}B\r",
                &[
                    (4, kind, 12_240, 15_840, 1, 720),
                    (6, 2, 12_240, 15_840, 1, 720),
                ],
            );
            let cfb = CompoundFile::open(&bytes).unwrap();
            let direct = super::super::direct_model(&cfb, 1024 * 1024).unwrap();
            let converted = super::super::convert(&cfb, 1024 * 1024).unwrap();
            let expected: serde_json::Value =
                serde_json::from_str(&docx_parser::parse_docx_native(&converted.bytes).unwrap())
                    .unwrap();
            let actual = serde_json::to_value(&direct).unwrap();
            let normalize = |body: &serde_json::Value| {
                let mut body = body.clone();
                for value in body.as_array_mut().unwrap() {
                    if value["type"] == "paragraph" {
                        value.as_object_mut().unwrap().remove("styleId");
                    }
                }
                body
            };
            assert_eq!(
                normalize(&actual["body"]),
                normalize(&expected["body"]),
                "section kind {kind}"
            );
            assert_eq!(
                actual["body"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .filter(|element| element["type"] == "pageBreak")
                    .count(),
                1,
                "section kind {kind}"
            );
        }
    }

    #[test]
    fn section_break_subsumes_only_the_sole_projected_page_break() {
        for (kind, expected_page_breaks) in [(0, 1), (1, 1), (2, 0), (3, 0), (4, 0)] {
            let source = source_with_sections(
                "X\u{c}\u{c}B\r",
                &[
                    (3, kind, 12_240, 15_840, 1, 720),
                    (5, 2, 12_240, 15_840, 1, 720),
                ],
            );
            let bytes = hide_first_utf16_unit(&source);
            let cfb = CompoundFile::open(&bytes).unwrap();
            let direct = super::super::direct_model(&cfb, 1024 * 1024).unwrap();
            let converted = super::super::convert(&cfb, 1024 * 1024).unwrap();
            let expected: serde_json::Value =
                serde_json::from_str(&docx_parser::parse_docx_native(&converted.bytes).unwrap())
                    .unwrap();
            let actual = serde_json::to_value(&direct).unwrap();
            let normalize = |body: &serde_json::Value| {
                let mut body = body.clone();
                for value in body.as_array_mut().unwrap() {
                    if value["type"] == "paragraph" {
                        value.as_object_mut().unwrap().remove("styleId");
                    }
                }
                body
            };
            assert_eq!(
                normalize(&actual["body"]),
                normalize(&expected["body"]),
                "section kind {kind}"
            );
            assert_eq!(
                actual["body"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .filter(|element| element["type"] == "pageBreak")
                    .count(),
                expected_page_breaks,
                "section kind {kind}"
            );
        }

        for text in ["X\u{c}\u{c}B\r", "\u{c}\u{c}\u{c}B\r", "\u{b}\u{c}\u{c}B\r"] {
            let first_end = text[..text.len() - "B\r".len()].encode_utf16().count();
            let total = text.encode_utf16().count();
            let bytes = source_with_sections(
                text,
                &[
                    (first_end, 2, 12_240, 15_840, 1, 720),
                    (total, 2, 12_240, 15_840, 1, 720),
                ],
            );
            let cfb = CompoundFile::open(&bytes).unwrap();
            let direct = super::super::direct_model(&cfb, 1024 * 1024).unwrap();
            let converted = super::super::convert(&cfb, 1024 * 1024).unwrap();
            let expected: serde_json::Value =
                serde_json::from_str(&docx_parser::parse_docx_native(&converted.bytes).unwrap())
                    .unwrap();
            let actual = serde_json::to_value(&direct).unwrap();
            let normalize = |body: &serde_json::Value| {
                let mut body = body.clone();
                for value in body.as_array_mut().unwrap() {
                    if value["type"] == "paragraph" {
                        value.as_object_mut().unwrap().remove("styleId");
                    }
                }
                body
            };
            assert_eq!(
                normalize(&actual["body"]),
                normalize(&expected["body"]),
                "source {text:?}"
            );
        }
    }
}
