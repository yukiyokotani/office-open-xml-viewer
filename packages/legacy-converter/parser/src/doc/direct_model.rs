//! Internal direct DOC body producer. This is intentionally not a public route:
//! unsupported body owners fail closed while incremental model coverage lands.

use super::{unsupported, AcquiredDoc, Fields, Token};
use docx_model::{BodyElement, BreakType, DocRun, Document, DocumentSettings};

mod payload;

pub(super) fn build(mut facts: AcquiredDoc<'_>, max_bytes: usize) -> Result<Document, String> {
    if max_bytes == 0 {
        return Err("OUTPUT_TOO_LARGE".into());
    }
    if facts.sections.len() > 1 {
        return Err(unsupported(
            "direct DOC model does not yet support multiple sections",
        ));
    }
    // MS-DOC 2.3.3 / 2.8.22: the first six header-document stories are
    // footnote/endnote separators, not page headers. Only authored nonempty
    // header/footer ranges produce entries; an explicitly blank paragraph DOES
    // produce an entry and must not be discarded. No entries in any section
    // means there is no earlier header/footer to inherit either.
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
    let section = match facts.sections.first() {
        Some(section) => section.project_final(0, even_and_odd)?,
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
    let paragraphs =
        super::tokenize_with_fields(&facts.story.text, &mut Fields::default(), 0, true);
    let mut body = Vec::new();

    for source in paragraphs {
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
        let direct =
            facts
                .formatting
                .direct_paragraph(style, mark_fc, mark_piece.prm, &facts.story.prcs)?;
        if direct.numbering.is_some() {
            return Err(unsupported(
                "direct DOC model does not yet support numbered paragraphs",
            ));
        }
        let mut paragraph = direct.paragraph;
        budget.paragraph(&paragraph)?;
        // The empty template owns its own fonts/tabs/borders. Account for it
        // independently, even though it is released after this source paragraph.
        budget.paragraph(&paragraph)?;
        let base = paragraph.clone();
        let mut emitted_content = false;
        let mut emitted_segment = false;

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
                                emitted_content = true;
                            }
                            Ok(())
                        },
                    )?;
                }
                Token::Tab => {
                    emitted_content |= push_control_text(
                        &mut paragraph,
                        &mut facts,
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
                    emitted_content = true;
                }
                Token::PageBreak | Token::ColumnBreak => {
                    if !paragraph.runs.is_empty() {
                        budget.push(&mut body, BodyElement::Paragraph(Box::new(paragraph)))?;
                        budget.paragraph(&base)?;
                        paragraph = base.clone();
                    }
                    let element = match token {
                        Token::PageBreak => BodyElement::PageBreak {
                            parity: None,
                            same_paragraph_as_previous: emitted_content.then_some(true),
                        },
                        Token::ColumnBreak => BodyElement::ColumnBreak,
                        _ => unreachable!(),
                    };
                    budget.push(&mut body, element)?;
                    emitted_segment = true;
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
        if !paragraph.runs.is_empty() || !emitted_segment {
            budget.push(&mut body, BodyElement::Paragraph(Box::new(paragraph)))?;
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
        ..Document::default()
    })
}

fn push_control_text(
    paragraph: &mut docx_model::DocParagraph,
    facts: &mut AcquiredDoc<'_>,
    style: usize,
    cp: usize,
    text: &str,
    budget: &mut ModelBudget,
) -> Result<bool, String> {
    let (_, fc, piece) = facts
        .story
        .position(cp)
        .ok_or_else(|| unsupported("Word control outside piece table"))?;
    if let Some(mut run) =
        facts
            .formatting
            .direct_text_run(style, fc, piece.prm, &facts.story.prcs, String::new())?
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
        source_with_header_document(text, None)
    }

    fn source_with_header_document(text: &str, authored_blank_header: Option<bool>) -> Vec<u8> {
        let main_units = text.encode_utf16().count();
        // A separator-only header document is not an authored page header.
        // The optional blank header has its own paragraph plus guard mark.
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
        word[0x54..0x58].copy_from_slice(&(header.len() as u32).to_le_bytes());
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
        let sepx_offset = word.len();
        let mut sepx = Vec::new();
        for (code, value) in [
            (0x9023u16, 1440i16),
            (0x9024, 1440),
            (0xb021, 1440),
            (0xb022, 1440),
        ] {
            sepx.extend(code.to_le_bytes());
            sepx.extend(value.to_le_bytes());
        }
        word.extend((sepx.len() as u16).to_le_bytes());
        word.extend(sepx);
        let mut section_table = vec![0; 20];
        section_table[4..8].copy_from_slice(&(main_units as u32).to_le_bytes());
        section_table[10..14].copy_from_slice(&(sepx_offset as u32).to_le_bytes());
        table.extend(section_table);
        word[0xca..0xce].copy_from_slice(&(section_table_offset as u32).to_le_bytes());
        word[0xce..0xd2].copy_from_slice(&20u32.to_le_bytes());

        if let Some(authored) = authored_blank_header {
            let mut hdd = Vec::new();
            for index in 0..14 {
                let cp: u32 = match index {
                    0 | 13 => 0,
                    1..=6 => 1,
                    _ => {
                        if authored {
                            3
                        } else {
                            1
                        }
                    }
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
    fn separator_only_header_document_is_not_an_authored_page_header() {
        let plain = source("Body\r");
        let separators = source_with_header_document("Body\r", Some(false));
        let blank_header = source_with_header_document("Body\r", Some(true));
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
}
