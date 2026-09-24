//! Internal direct DOC body producer. This is intentionally not a public route:
//! unsupported body owners fail closed while incremental model coverage lands.

use super::{unsupported, AcquiredDoc, Fields};
use docx_model::paragraph_breaks::ParaPiece;
use docx_model::{
    BodyElement, CellElement, DocRun, Document, DocumentSettings, DocumentTypographySettingsWire,
    HeadersFooters,
};

pub(in crate::doc) mod fields;
mod headers;
mod notes;
mod payload;
mod story;
#[cfg(test)]
mod table_tests;
mod tables;

#[derive(Debug)]
pub(crate) struct DirectDocResult {
    pub(crate) document: Document,
    pub(crate) resources: Vec<super::pictures::DirectPictureResource>,
}

pub(super) fn build(
    mut facts: AcquiredDoc<'_>,
    max_bytes: usize,
) -> Result<DirectDocResult, String> {
    if max_bytes == 0 {
        return Err("OUTPUT_TOO_LARGE".into());
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
            adjust_line_height_in_table: Some(settings.adjust_line_height_in_table),
            balance_single_byte_double_byte_width: Some(
                settings.balance_single_byte_double_byte_width,
            ),
            character_spacing_control: settings.character_spacing_control.map(str::to_string),
            ..DocumentSettings::default()
        });
    let document_typography_settings = Some(DocumentTypographySettingsWire {
        normal_style_font_size_pt: facts.formatting.direct_normal_style_font_size_pt()?,
    });

    let mut budget = ModelBudget::new(max_bytes);
    budget.charge(std::mem::size_of::<DirectDocResult>())?;
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
    let note_layout_settings = notes::validate(
        &facts.note_stories,
        &facts.note_references,
        facts.effective_nfib,
        facts.document_settings.as_ref(),
        &facts.sections,
        facts.headers.as_ref(),
    )?;
    let main_fields = match &facts.main_fields {
        Ok(table) => fields::StoryFields::analyze(&facts.story.text, table, &[])?,
        Err(error) => return Err(error.clone()),
    };
    // Header drawings and textbox stories are resolved only by this model;
    // malformed tables fail closed before any output is retained.
    facts.floating.load_direct_parts()?;
    // Word stores TIFF data in PNG BLIPs and reads it back as TIFF: its own
    // DOCX of a corpus document writes that BLIP as media/*.tiff. The
    // package writer keeps rejecting such BLIPs; no other format has this
    // evidence. Painting TIFF needs the caller's optional TIFF decoder.
    facts.pictures.raster = crate::officeart::raster::Raster::TiffAware;
    facts.floating.raster = crate::officeart::raster::Raster::TiffAware;
    let mut body = Vec::new();
    let mut header_resolver = headers::Resolver::new(facts.headers.as_ref())?;
    let mut final_headers = None;
    let mut final_footers = None;
    let chunks = super::sections::split_story(&facts.story.text, &facts.sections)?;
    let mut fields = Fields::default();
    let mut table_sequence = 0;
    let mut numbering = super::numbering::direct::Store::default();
    numbering.begin_story()?;
    let mut placed_note_references = 0;
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
        main_fields.apply(base_cp, &mut paragraphs)?;
        notes::restore_references(
            &facts.note_references,
            &mut paragraphs,
            &mut placed_note_references,
        )?;
        if section_index + 1 < chunks.len() {
            // split_story consumed the section-break form feed. The paragraph
            // mark formatting remains owned by the preceding physical CP.
            paragraphs.last_mut().expect("opening paragraph").end_cp =
                facts.sections[section_index].end - 1;
        }

        let vertical_flow = match &ending {
            Some(ending) => ending.text_direction.as_deref(),
            None => section.text_direction.as_deref(),
        } == Some("tbRl");
        let section_start = body.len();
        story::project(
            &facts.story,
            paragraphs,
            &mut facts.formatting,
            &mut numbering,
            &mut facts.pictures,
            Some((&mut facts.floating, super::floating::Part::Main)),
            &mut budget,
            &mut body,
            ending.as_ref().map(|ending| ending.kind.as_str()),
            &mut table_sequence,
        )?;
        if super::character::Properties::unrenderable_east_asian_vertical(
            &body[section_start..],
            vertical_flow,
        ) {
            facts.formatting.unsupported_character_properties = true;
        }

        let (section_headers, section_footers) = header_resolver.project_section(
            section_index,
            &mut facts.formatting,
            &mut facts.pictures,
            &mut facts.floating,
            &mut budget,
            &mut table_sequence,
        )?;

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
                    headers: Box::new(section_headers),
                    footers: Box::new(section_footers),
                    title_page: ending.title_page,
                    geom: Some(ending.geom),
                    page_num_type: ending.page_num_type,
                    text_direction: ending.text_direction,
                    section_placement: ending.placement,
                },
            )?;
        } else {
            final_headers = Some(section_headers);
            final_footers = Some(section_footers);
        }
    }

    notes::check_all_references_placed(&facts.note_references, placed_note_references)?;
    let (mut footnotes, mut endnotes) = (Vec::new(), Vec::new());
    for story in facts.note_stories.iter().flatten() {
        let (fields, output) = match story.kind {
            super::notes::Kind::Footnote => (&facts.note_fields[0], &mut footnotes),
            super::notes::Kind::Endnote => (&facts.note_fields[1], &mut endnotes),
        };
        let fields = fields.as_ref().map_err(Clone::clone)?;
        notes::project(
            story,
            fields,
            &mut facts.formatting,
            &mut facts.pictures,
            &mut budget,
            &mut table_sequence,
            output,
        )?;
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

    let document = Document {
        section,
        body,
        headers: final_headers.unwrap_or_default(),
        footers: final_footers.unwrap_or_default(),
        settings,
        document_typography_settings,
        footnotes,
        endnotes,
        note_layout_settings,
        ..Document::default()
    };
    if !facts.pictures.has_selected_direct_resources()
        && !facts.floating.has_selected_direct_resources()
    {
        return Ok(DirectDocResult {
            document,
            resources: Vec::new(),
        });
    }
    let references = direct_picture_references(&document, &mut budget)?;
    let mut resources = facts
        .pictures
        .finish_referenced_direct_resources(&references, &mut budget.remaining_bytes)?;
    facts.floating.append_referenced_direct_resources(
        &mut resources,
        &references,
        &mut budget.remaining_bytes,
    )?;
    Ok(DirectDocResult {
        document,
        resources,
    })
}

#[derive(Clone, Copy)]
enum RetainedBlock<'a> {
    Body(&'a BodyElement),
    Cell(&'a CellElement),
}

/// Collect direct-DOC picture references from the final retained model after
/// table merge projection has discarded continuation content. This one-time
/// pass is O(retained nodes + references log references), uses no recursion,
/// and charges its O(retained nodes) worst-case traversal stack and reference
/// vector to the document's cumulative model budget. Documents with no selected
/// direct pictures bypass the pass.
fn direct_picture_references<'a>(
    document: &'a Document,
    budget: &mut ModelBudget,
) -> Result<Vec<&'a str>, String> {
    fn push_headers<'a>(
        headers: &'a HeadersFooters,
        pending: &mut Vec<RetainedBlock<'a>>,
        budget: &mut ModelBudget,
    ) -> Result<(), String> {
        for header in [
            headers.default.as_ref(),
            headers.first.as_ref(),
            headers.even.as_ref(),
        ]
        .into_iter()
        .flatten()
        {
            for block in &header.body {
                budget.push(pending, RetainedBlock::Body(block))?;
            }
        }
        Ok(())
    }

    fn push_key<'a>(
        key: &'a str,
        references: &mut Vec<&'a str>,
        budget: &mut ModelBudget,
    ) -> Result<(), String> {
        if !key.starts_with("legacy-doc/") {
            return Ok(());
        }
        if !(key.starts_with("legacy-doc/image/") || key.starts_with("legacy-doc/float/")) {
            return Err(unsupported("unknown direct DOC picture resource namespace"));
        }
        budget.push(references, key)
    }

    fn paragraph<'a>(
        paragraph: &'a docx_model::DocParagraph,
        pending: &mut Vec<RetainedBlock<'a>>,
        references: &mut Vec<&'a str>,
        budget: &mut ModelBudget,
    ) -> Result<(), String> {
        if let Some(numbering) = &paragraph.numbering {
            if let Some(key) = &numbering.pic_bullet_image_path {
                push_key(key, references, budget)?;
            }
        }
        for run in &paragraph.runs {
            match run {
                DocRun::Image(image) => {
                    push_key(&image.image_path, references, budget)?;
                    if let Some(key) = &image.svg_image_path {
                        push_key(key, references, budget)?;
                    }
                }
                // Textbox stories can hold inline pictures; picture fills
                // share the floating picture resources.
                DocRun::Shape(shape) => {
                    if let Some(docx_model::ShapeFill::Image { image_path, .. }) = &shape.fill {
                        push_key(image_path, references, budget)?;
                    }
                    for block in &shape.text_box_content {
                        match block {
                            docx_model::TextBoxBlockWire::Body(block) => {
                                budget.push(pending, RetainedBlock::Body(block))?
                            }
                            docx_model::TextBoxBlockWire::Unsupported { .. } => {
                                return Err(unsupported(
                                    "unexpected unsupported direct DOC textbox block",
                                ))
                            }
                        }
                    }
                }
                _ => {}
            }
        }
        Ok(())
    }

    fn table<'a>(
        table: &'a docx_model::DocTable,
        pending: &mut Vec<RetainedBlock<'a>>,
        budget: &mut ModelBudget,
    ) -> Result<(), String> {
        for row in &table.rows {
            for cell in &row.cells {
                for block in &cell.content {
                    budget.push(pending, RetainedBlock::Cell(block))?;
                }
            }
        }
        Ok(())
    }

    let mut pending = Vec::new();
    for block in &document.body {
        budget.push(&mut pending, RetainedBlock::Body(block))?;
    }
    push_headers(&document.headers, &mut pending, budget)?;
    push_headers(&document.footers, &mut pending, budget)?;
    for note in document.footnotes.iter().chain(&document.endnotes) {
        for block in &note.content {
            budget.push(&mut pending, RetainedBlock::Body(block))?;
        }
    }

    let mut references = Vec::new();
    while let Some(block) = pending.pop() {
        match block {
            RetainedBlock::Body(BodyElement::Paragraph(value))
            | RetainedBlock::Cell(CellElement::Paragraph(value)) => {
                paragraph(value, &mut pending, &mut references, budget)?
            }
            RetainedBlock::Body(BodyElement::Table(value))
            | RetainedBlock::Cell(CellElement::Table(value)) => table(value, &mut pending, budget)?,
            RetainedBlock::Body(BodyElement::SectionBreak {
                headers, footers, ..
            }) => {
                push_headers(headers, &mut pending, budget)?;
                push_headers(footers, &mut pending, budget)?;
            }
            RetainedBlock::Body(BodyElement::PageBreak { .. } | BodyElement::ColumnBreak) => {}
        }
    }
    references.sort_unstable();
    references.dedup();
    Ok(references)
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
    use std::io::{Cursor, Read};

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
        source_with_typography(text, sections, authored_blank_header, None, None, None)
    }

    pub(super) fn source_with_typography(
        text: &str,
        sections: &[(usize, u8, u16, u16, u16, u16)],
        authored_blank_header: Option<bool>,
        normal_hps: Option<u16>,
        body_hps: Option<u16>,
        header_slots: Option<&[Option<&str>]>,
    ) -> Vec<u8> {
        source_with_stories(
            text,
            sections,
            authored_blank_header,
            normal_hps,
            body_hps,
            header_slots,
            None,
        )
    }

    /// Footnote document for `source_with_stories`: note texts (each ending
    /// with its paragraph mark), their main-story reference CPs and FRD
    /// automatic flags, the six separator stories (each including its guard)
    /// and an optional raw DOP.
    pub(super) struct NotesFixture<'a> {
        pub(super) notes: &'a [&'a str],
        pub(super) references: &'a [(u32, bool)],
        pub(super) separators: [&'a str; 6],
        pub(super) dop: Option<Vec<u8>>,
    }

    /// Dop97-sized DOP with Word's default note properties: bottom-of-page
    /// footnotes, continuous numbering from 1, end-of-document endnotes and
    /// Arabic footnote/endnote formats.
    pub(super) fn default_note_dop() -> Vec<u8> {
        let mut dop = vec![0u8; 500];
        dop[0] = 1 << 5;
        dop[2..4].copy_from_slice(&(1u16 << 2).to_le_bytes());
        dop[10..12].copy_from_slice(&720u16.to_le_bytes());
        dop[52..54].copy_from_slice(&(1u16 << 2).to_le_bytes());
        dop[54] = 3;
        dop
    }

    pub(super) fn source_with_stories(
        text: &str,
        sections: &[(usize, u8, u16, u16, u16, u16)],
        authored_blank_header: Option<bool>,
        normal_hps: Option<u16>,
        body_hps: Option<u16>,
        header_slots: Option<&[Option<&str>]>,
        notes: Option<&NotesFixture<'_>>,
    ) -> Vec<u8> {
        let main_units = text.encode_utf16().count();
        let footnote_text = notes
            .map(|fixture| format!("{}\r", fixture.notes.concat()))
            .unwrap_or_default();
        let (header, header_cps) = if let Some(slots) = header_slots {
            assert_eq!(slots.len(), sections.len() * 6);
            let (mut header, mut cps) = if let Some(fixture) = notes {
                let mut cps = vec![0u32];
                for separator in fixture.separators {
                    cps.push(cps.last().unwrap() + separator.encode_utf16().count() as u32);
                }
                (fixture.separators.concat(), cps)
            } else {
                (String::from("\r"), vec![0u32, 1, 1, 1, 1, 1, 1])
            };
            let mut cp = *cps.last().unwrap();
            for slot in slots {
                if let Some(content) = slot {
                    assert!(content.ends_with('\r'));
                    header.push_str(content);
                    header.push('\r');
                    cp += content.encode_utf16().count() as u32 + 1;
                }
                cps.push(cp);
            }
            header.push('\r');
            (header, Some(cps))
        } else {
            (
                match authored_blank_header {
                    None => String::new(),
                    Some(false) => "\r\r".into(),
                    Some(true) => "\r\r\r\r".into(),
                },
                None,
            )
        };
        let units: Vec<u16> = text
            .encode_utf16()
            .chain(footnote_text.encode_utf16())
            .chain(header.encode_utf16())
            .collect();
        let text_offset = 0x400usize;
        let mut word = vec![0u8; text_offset + units.len() * 2];
        super::super::write_minimal_word97_test_header(&mut word);
        word[6..8].copy_from_slice(&1033u16.to_le_bytes());
        word[0x4c..0x50].copy_from_slice(&(main_units as u32).to_le_bytes());
        word[0x50..0x54]
            .copy_from_slice(&(footnote_text.encode_utf16().count() as u32).to_le_bytes());
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

        if let Some(cps) = header_cps {
            let mut hdd = Vec::new();
            for cp in cps {
                hdd.extend(cp.to_le_bytes());
            }
            hdd.extend(u32::MAX.to_le_bytes());
            append_table_part(&mut word, &mut table, 0xf2, &hdd);
        } else if let Some(authored) = authored_blank_header {
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
        if let Some(fixture) = notes {
            let mut references = Vec::new();
            for (cp, _) in fixture.references {
                references.extend(cp.to_le_bytes());
            }
            references.extend((main_units as u32).to_le_bytes());
            for (_, automatic) in fixture.references {
                references.extend(u16::from(*automatic).to_le_bytes());
            }
            append_table_part(&mut word, &mut table, 0xaa, &references);
            let mut boundaries = Vec::new();
            let mut cp = 0u32;
            for note in fixture.notes {
                boundaries.extend(cp.to_le_bytes());
                cp += note.encode_utf16().count() as u32;
            }
            boundaries.extend(cp.to_le_bytes());
            boundaries.extend((cp + 1).to_le_bytes());
            append_table_part(&mut word, &mut table, 0xb2, &boundaries);
            if let Some(dop) = &fixture.dop {
                append_table_part(&mut word, &mut table, 0x192, dop);
            }
        }
        for (story, fib_offset) in [
            (text, 0x11a),
            (header.as_str(), 0x122),
            (footnote_text.as_str(), 0x12a),
        ] {
            if let Some(field_table) = field_plc(story) {
                append_table_part(&mut word, &mut table, fib_offset, &field_table);
            }
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

    /// MS-DOC 2.8.25 Plcfld for every field character in `story`: flt from
    /// the instruction keyword and grffldEnd fHasSep/fNested from structure.
    pub(super) fn field_plc(story: &str) -> Option<Vec<u8>> {
        let units: Vec<u16> = story.encode_utf16().collect();
        let mut entries: Vec<(u32, u8, u8)> = Vec::new();
        let mut open: Vec<(usize, bool)> = Vec::new();
        for (cp, unit) in units.iter().enumerate() {
            match unit {
                0x13 => {
                    let keyword: String = char::decode_utf16(units[cp + 1..].iter().copied())
                        .map(|value| value.unwrap_or(' '))
                        .skip_while(|value| *value == ' ')
                        .take_while(|value| value.is_ascii_alphabetic())
                        .collect();
                    let flt = match keyword.to_ascii_uppercase().as_str() {
                        "PAGE" => 0x21,
                        "NUMPAGES" => 0x1a,
                        "DATE" => 0x1f,
                        "TIME" => 0x20,
                        "REF" => 0x03,
                        "IF" => 0x07,
                        _ => 0x01,
                    };
                    open.push((entries.len(), false));
                    entries.push((cp as u32, 0x13, flt));
                }
                0x14 => {
                    open.last_mut().unwrap().1 = true;
                    entries.push((cp as u32, 0x14, 0));
                }
                0x15 => {
                    let (_, separator) = open.pop().unwrap();
                    let flags =
                        if separator { 0x80 } else { 0 } | if open.is_empty() { 0 } else { 0x40 };
                    entries.push((cp as u32, 0x15, flags));
                }
                _ => {}
            }
        }
        if entries.is_empty() {
            return None;
        }
        let mut plc = Vec::new();
        for (cp, ..) in &entries {
            plc.extend(cp.to_le_bytes());
        }
        plc.extend((units.len() as u32).to_le_bytes());
        for (_, marker, grffld) in entries {
            plc.extend([marker, grffld]);
        }
        Some(plc)
    }

    fn append_table_part(word: &mut [u8], table: &mut Vec<u8>, fib_offset: usize, part: &[u8]) {
        word[fib_offset..fib_offset + 4].copy_from_slice(&(table.len() as u32).to_le_bytes());
        word[fib_offset + 4..fib_offset + 8].copy_from_slice(&(part.len() as u32).to_le_bytes());
        table.extend(part);
    }

    fn numbered_source(text: &str) -> Vec<u8> {
        with_numbering(&source(text))
    }

    pub(super) fn with_numbering(source: &[u8]) -> Vec<u8> {
        let cfb = CompoundFile::open(&source).unwrap();
        let mut word = cfb.stream("WordDocument").unwrap();
        let mut table = cfb.stream("0Table").unwrap();

        // One simple decimal LST/LVL. LVLF.rgbxchNums points at the first
        // UTF-16 unit; the unit value is the referenced zero-based level.
        let list_start = table.len();
        let mut list = vec![0; 28];
        list[..4].copy_from_slice(&42i32.to_le_bytes());
        for style in list[8..26].chunks_exact_mut(2) {
            style.copy_from_slice(&0x0fffu16.to_le_bytes());
        }
        list[26] = 1;
        table.extend(1u16.to_le_bytes());
        table.extend(list);
        word[0x2e2..0x2e6].copy_from_slice(&(list_start as u32).to_le_bytes());
        word[0x2e6..0x2ea].copy_from_slice(&30u32.to_le_bytes());
        let mut level = vec![0; 28];
        level[0] = 1;
        level[6] = 1;
        level.extend([2, 0, 0, 0, b'.', 0]);
        table.extend(level);

        let lfo_start = table.len();
        table.extend(1u32.to_le_bytes());
        let mut lfo = vec![0; 16];
        lfo[..4].copy_from_slice(&42i32.to_le_bytes());
        table.extend(lfo);
        table.extend([0xff; 4]);
        word[0x2ea..0x2ee].copy_from_slice(&(lfo_start as u32).to_le_bytes());
        word[0x2ee..0x2f2].copy_from_slice(&24u32.to_le_bytes());

        // Apply sprmPIlvl=0 and sprmPIlfo=1 to the existing single PAPX run.
        let bte = u32::from_le_bytes(word[0x102..0x106].try_into().unwrap()) as usize;
        let page_number = u32::from_le_bytes(table[bte + 8..bte + 12].try_into().unwrap()) as usize;
        let page = &mut word[page_number * 512..(page_number + 1) * 512];
        page[8] = 32;
        page[64] = 5;
        page[65..74].copy_from_slice(&[0, 0, 0x0a, 0x26, 0, 0x0b, 0x46, 1, 0]);
        build_cfb(&[("WordDocument", word), ("0Table", table)])
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

    fn picture_record(kind: u16, options: u16, body: &[u8]) -> Vec<u8> {
        [
            options.to_le_bytes().as_slice(),
            &kind.to_le_bytes(),
            &(body.len() as u32).to_le_bytes(),
            body,
        ]
        .concat()
    }

    fn picture_source(text: &str, vanish: bool) -> Vec<u8> {
        with_picture_data(&source(text), vanish)
    }

    fn source_with_proofing(text: &str, no_proof: bool) -> Vec<u8> {
        let source = source(text);
        let cfb = CompoundFile::open(&source).unwrap();
        let mut word = cfb.stream("WordDocument").unwrap();
        let table = cfb.stream("0Table").unwrap();
        let bte = u32::from_le_bytes(word[0xfa..0xfe].try_into().unwrap()) as usize;
        let page_number = u32::from_le_bytes(table[bte + 8..bte + 12].try_into().unwrap()) as usize;
        let page = &mut word[page_number * 512..(page_number + 1) * 512];
        let mut chpx = vec![0x35, 0x08, 1];
        if no_proof {
            chpx.extend([0x75, 0x08, 1]);
        }
        page[8] = 32;
        page[64] = chpx.len() as u8;
        page[65..65 + chpx.len()].copy_from_slice(&chpx);
        build_cfb(&[("WordDocument", word), ("0Table", table)])
    }

    pub(super) fn with_picture_data(source: &[u8], vanish: bool) -> Vec<u8> {
        let cfb = CompoundFile::open(source).unwrap();
        let mut word = cfb.stream("WordDocument").unwrap();
        let table = cfb.stream("0Table").unwrap();
        let bte = u32::from_le_bytes(word[0xfa..0xfe].try_into().unwrap()) as usize;
        let page_number = u32::from_le_bytes(table[bte + 8..bte + 12].try_into().unwrap()) as usize;
        let page = &mut word[page_number * 512..(page_number + 1) * 512];
        let mut sprms = vec![0x55, 0x08, 1, 0x03, 0x6a, 0, 0, 0, 0];
        if vanish {
            sprms.extend([0x3c, 0x08, 1]);
        }
        page[8] = 32;
        page[64] = sprms.len() as u8;
        page[65..65 + sprms.len()].copy_from_slice(&sprms);

        let png = {
            let mut bytes = b"\x89PNG\r\n\x1a\n\0\0\0\x0dIHDR".to_vec();
            bytes.extend_from_slice(&2u32.to_be_bytes());
            bytes.extend_from_slice(&3u32.to_be_bytes());
            bytes.extend_from_slice(&[8, 2, 0, 0, 0, 0, 0, 0, 0]);
            bytes
        };
        let raster = picture_record(0xf01e, 0x6e0 << 4, &[vec![0; 17], png].concat());
        let mut options = Vec::new();
        for (key, value) in [
            (0x0104u16, 1u32),
            (0x0100, 8192),
            (0x0101, 16384),
            (0x0102, 24576),
            (0x0103, 32768),
            (4, 90 * 65536),
        ] {
            options.extend(key.to_le_bytes());
            options.extend(value.to_le_bytes());
        }
        let mut shape = picture_record(0xf00a, (75 << 4) | 2, &[1, 0, 0, 0, 0x40, 8, 0, 0]);
        shape.extend(picture_record(0xf00b, (6 << 4) | 3, &options));
        let mut data = vec![0u8; 68];
        data[4..6].copy_from_slice(&68u16.to_le_bytes());
        data[6..8].copy_from_slice(&100u16.to_le_bytes());
        data[28..30].copy_from_slice(&1440u16.to_le_bytes());
        data[30..32].copy_from_slice(&720u16.to_le_bytes());
        data[32..34].copy_from_slice(&500u16.to_le_bytes());
        data[34..36].copy_from_slice(&2000u16.to_le_bytes());
        data.extend(picture_record(0xf004, 15, &shape));
        data.extend(raster);
        let length = data.len() as u32;
        data[..4].copy_from_slice(&length.to_le_bytes());
        data[88..92].copy_from_slice(&0xc0u32.to_le_bytes());
        build_cfb(&[("WordDocument", word), ("0Table", table), ("Data", data)])
    }

    fn floating_picture_source(text: &str, vanish: bool) -> Vec<u8> {
        let source = source(text);
        let cfb = CompoundFile::open(&source).unwrap();
        let mut word = cfb.stream("WordDocument").unwrap();
        let mut table = cfb.stream("0Table").unwrap();
        let bte = u32::from_le_bytes(word[0xfa..0xfe].try_into().unwrap()) as usize;
        let page_number = u32::from_le_bytes(table[bte + 8..bte + 12].try_into().unwrap()) as usize;
        let page = &mut word[page_number * 512..(page_number + 1) * 512];
        let mut chpx = vec![0x55, 0x08, 1];
        if vanish {
            chpx.extend([0x3c, 0x08, 1]);
        }
        page[8] = 32;
        page[64] = chpx.len() as u8;
        page[65..65 + chpx.len()].copy_from_slice(&chpx);

        let main_units = text.encode_utf16().count();
        let anchor_offset = table.len();
        let flags = (1u16 << 1) | (2 << 3) | (2 << 5) | (2 << 9) | (1 << 15);
        let mut anchors = [
            1u32.to_le_bytes().as_slice(),
            &(main_units as u32).to_le_bytes(),
            &1027u32.to_le_bytes(),
            &(-100i32).to_le_bytes(),
            &200i32.to_le_bytes(),
            &300i32.to_le_bytes(),
            &500i32.to_le_bytes(),
            &flags.to_le_bytes(),
            &[0; 4],
        ]
        .concat();
        word[0x1da..0x1de].copy_from_slice(&(anchor_offset as u32).to_le_bytes());
        word[0x1de..0x1e2].copy_from_slice(&(anchors.len() as u32).to_le_bytes());
        table.append(&mut anchors);

        let mut png = b"\x89PNG\r\n\x1a\n\0\0\0\x0dIHDR".to_vec();
        png.extend(2u32.to_be_bytes());
        png.extend(3u32.to_be_bytes());
        png.extend([8, 2, 0, 0, 0, 0, 0, 0, 0]);
        let blip = picture_record(0xf01e, 0x6e0 << 4, &[vec![0; 17], png].concat());
        let delayed = word.len();
        word.extend(&blip);
        let mut bse = vec![0; 36];
        bse[0] = 6;
        bse[1] = 6;
        bse[20..24].copy_from_slice(&(blip.len() as u32).to_le_bytes());
        bse[24] = 1;
        bse[28..32].copy_from_slice(&(delayed as u32).to_le_bytes());
        let group = picture_record(
            0xf000,
            15,
            &picture_record(0xf001, 31, &picture_record(0xf007, 0x62, &bse)),
        );
        let properties = [
            (0x4104u16, 1u32),
            (0x3bf, 0x8200_0000),
            (0x384, 12_700),
            (0x385, 25_400),
            (0x386, 38_100),
            (0x387, 50_800),
            (0x3aa, 77),
            (0x0100, 8192),
            (0x0101, 16384),
            (0x0102, 24576),
            (0x0103, 32768),
        ];
        let mut options = Vec::new();
        for (key, value) in properties {
            options.extend(key.to_le_bytes());
            options.extend(value.to_le_bytes());
        }
        let shape = picture_record(
            0xf004,
            15,
            &[
                picture_record(
                    0xf00a,
                    (75 << 4) | 2,
                    &[1027u32.to_le_bytes(), 0xac0u32.to_le_bytes()].concat(),
                ),
                picture_record(0xf00b, ((properties.len() as u16) << 4) | 3, &options),
                picture_record(0xf010, 0, &0u32.to_le_bytes()),
            ]
            .concat(),
        );
        let art = [
            group,
            vec![0],
            picture_record(0xf002, 15, &picture_record(0xf003, 15, &shape)),
        ]
        .concat();
        word[0x22a..0x22e].copy_from_slice(&(table.len() as u32).to_le_bytes());
        word[0x22e..0x232].copy_from_slice(&(art.len() as u32).to_le_bytes());
        table.extend(art);
        build_cfb(&[("WordDocument", word), ("0Table", table)])
    }

    /// A textbox (msosptTextBox) or rectangle anchored in the main or header
    /// document. The textbox story follows the header story; its FTXBXS and
    /// Tbkd tables name the shape (MS-DOC 2.3.6-2.3.7, 2.9.106, 2.9.312).
    fn drawing_shape_source(textbox: Option<&str>, header: bool) -> Vec<u8> {
        drawing_source(textbox, header, false, [0, 0])
    }

    /// With `grouped`, the anchored shape is an OfficeArt group: FSPGR
    /// 0..1000 x 0..500 holds a filled rectangle member in its top-left
    /// quarter and a nested group (FSPGR -10..10) in the bottom-right quarter
    /// whose only member is the textbox (spid 2052).
    /// `turns` are the FixedPoint rotations of the group and of its first
    /// (rectangle) member.
    fn drawing_source(
        textbox: Option<&str>,
        header: bool,
        grouped: bool,
        turns: [u32; 2],
    ) -> Vec<u8> {
        let main = if header { "B\r" } else { "B\u{8}\r" };
        let main_units = main.encode_utf16().count();
        let textbox = textbox.unwrap_or("");
        let slots = [None, Some("\u{8}\r"), None, None, None, None];
        let bytes = source_with_typography(
            &format!("{main}{textbox}"),
            &[(main_units, 2, 12_240, 15_840, 1, 720)],
            None,
            None,
            None,
            header.then_some(&slots[..]),
        );
        let bytes = passive_special_source(&bytes);
        let cfb = CompoundFile::open(&bytes).unwrap();
        let mut word = cfb.stream("WordDocument").unwrap();
        let mut table = cfb.stream("0Table").unwrap();
        // The builder placed the textbox text after the main story; move it
        // behind the (optional) header story by swapping the story lengths.
        let textbox_units = textbox.encode_utf16().count() as u32;
        if header {
            assert_eq!(textbox_units, 0);
        }
        word[0x4c..0x50].copy_from_slice(&(main_units as u32).to_le_bytes());
        word[0x64..0x68].copy_from_slice(&textbox_units.to_le_bytes());

        // SPA at the anchor character: page/paragraph origin, no wrapping.
        // Both stories place the anchor character at CP 1 of their own
        // document; the PLC's final CP is its story length.
        let anchor_cp = 1u32;
        let anchor_units = if header { 5 } else { main_units as u32 };
        let flags = (1u16 << 1) | (2 << 3) | (3 << 5);
        let anchors = [
            anchor_cp.to_le_bytes().as_slice(),
            &anchor_units.to_le_bytes(),
            &2050u32.to_le_bytes(),
            &100i32.to_le_bytes(),
            &200i32.to_le_bytes(),
            &2100i32.to_le_bytes(),
            &1200i32.to_le_bytes(),
            &flags.to_le_bytes(),
            &[0; 4],
        ]
        .concat();
        append_table_part(
            &mut word,
            &mut table,
            if header { 0x1e2 } else { 0x1da },
            &anchors,
        );

        let mut properties = vec![
            (0x181u16, 0x00ff_0000u32),
            (0x1c0, 0x0000_00ff),
            (0x1cb, 12_700),
        ];
        if !textbox.is_empty() {
            properties.extend([(0x80, 0x10000), (0x81, 0), (0xbf, 0x20002)]);
        }
        let mut options = Vec::new();
        for (key, value) in &properties {
            options.extend(key.to_le_bytes());
            options.extend(value.to_le_bytes());
        }
        let kind: u16 = if textbox.is_empty() { 1 } else { 202 };
        let shape = picture_record(
            0xf004,
            15,
            &[
                picture_record(
                    0xf00a,
                    (kind << 4) | 2,
                    &[2050u32.to_le_bytes(), 0xa00u32.to_le_bytes()].concat(),
                ),
                picture_record(0xf00b, ((properties.len() as u16) << 4) | 3, &options),
                picture_record(0xf010, 0, &0u32.to_le_bytes()),
            ]
            .concat(),
        );
        let shape = if grouped {
            let rect = |values: [i32; 4]| {
                values
                    .iter()
                    .flat_map(|value| value.to_le_bytes())
                    .collect::<Vec<u8>>()
            };
            let fsp = |kind: u16, spid: u32, flags: u32| {
                picture_record(
                    0xf00a,
                    (kind << 4) | 2,
                    &[spid.to_le_bytes(), flags.to_le_bytes()].concat(),
                )
            };
            let member = |kind: u16, spid: u32, anchor: [i32; 4], options: &[u8], count: usize| {
                picture_record(
                    0xf004,
                    15,
                    &[
                        fsp(kind, spid, 0xa02),
                        picture_record(0xf00b, ((count as u16) << 4) | 3, options),
                        picture_record(0xf00f, 0, &rect(anchor)),
                    ]
                    .concat(),
                )
            };
            let mut fill = Vec::new();
            fill.extend(0x181u16.to_le_bytes());
            fill.extend(0x0000_ff00u32.to_le_bytes());
            fill.extend(0x4u16.to_le_bytes());
            fill.extend(turns[1].to_le_bytes());
            let mut turn = Vec::new();
            turn.extend(0x4u16.to_le_bytes());
            turn.extend(turns[0].to_le_bytes());
            let nested = picture_record(
                0xf003,
                15,
                &[
                    picture_record(
                        0xf004,
                        15,
                        &[
                            picture_record(0xf009, 1, &rect([-10, -10, 10, 10])),
                            fsp(0, 2053, 0x203),
                            picture_record(0xf00f, 0, &rect([500, 250, 1000, 500])),
                        ]
                        .concat(),
                    ),
                    member(202, 2052, [-10, -10, 10, 10], &options, properties.len()),
                ]
                .concat(),
            );
            picture_record(
                0xf003,
                15,
                &[
                    picture_record(
                        0xf004,
                        15,
                        &[
                            picture_record(0xf009, 1, &rect([0, 0, 1000, 500])),
                            fsp(0, 2050, 0x201),
                            picture_record(0xf00b, (1 << 4) | 3, &turn),
                            picture_record(0xf010, 0, &0u32.to_le_bytes()),
                        ]
                        .concat(),
                    ),
                    member(1, 2051, [0, 0, 500, 250], &fill, 2),
                    nested,
                ]
                .concat(),
            )
        } else {
            shape
        };
        let art = [
            picture_record(0xf000, 15, &[]),
            vec![u8::from(header)],
            picture_record(0xf002, 15, &picture_record(0xf003, 15, &shape)),
        ]
        .concat();
        append_table_part(&mut word, &mut table, 0x22a, &art);

        if !textbox.is_empty() {
            // FTXBXS: the textbox and the trailing reusable spare structure.
            let mut ftxbxs = [0u32, textbox_units, textbox_units + 1]
                .iter()
                .flat_map(|cp| cp.to_le_bytes())
                .collect::<Vec<_>>();
            let mut actual = vec![0; 22];
            let lid: u32 = if grouped { 2052 } else { 2050 };
            actual[14..18].copy_from_slice(&lid.to_le_bytes());
            let mut spare = vec![0; 22];
            spare[8] = 1;
            ftxbxs.extend(actual);
            ftxbxs.extend(spare);
            append_table_part(&mut word, &mut table, 0x25a, &ftxbxs);
            let mut tbkd = [0u32, textbox_units, textbox_units + 2]
                .iter()
                .flat_map(|cp| cp.to_le_bytes())
                .collect::<Vec<_>>();
            tbkd.extend([0, 0, 0, 0, 0, 0, 0xff, 0xff, 0, 0, 0, 0]);
            append_table_part(&mut word, &mut table, 0x2f2, &tbkd);
        }
        build_cfb(&[("WordDocument", word), ("0Table", table)])
    }

    #[test]
    fn textbox_shape_projects_paint_margins_and_its_story() {
        let bytes = drawing_shape_source(Some("Inside\rbox\r"), false);
        let result =
            super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1024 * 1024).unwrap();
        let BodyElement::Paragraph(paragraph) = &result.document.body[0] else {
            panic!("body paragraph")
        };
        let [DocRun::Text(_), DocRun::AnchorHost(host), DocRun::Shape(shape)] =
            paragraph.runs.as_slice()
        else {
            panic!("text, anchor host, shape: {:?}", paragraph.runs)
        };
        let acquisition = shape.anchor_acquisition.as_ref().unwrap();
        assert_eq!(
            host.anchor_occurrence_id.as_deref(),
            Some(acquisition.occurrence_id.as_str())
        );
        assert_eq!(shape.preset_geometry.as_deref(), Some("rect"));
        assert_eq!((shape.width_pt, shape.height_pt), (100.0, 50.0));
        assert_eq!((shape.anchor_x_pt, shape.anchor_y_pt), (5.0, 10.0));
        assert_eq!(shape.anchor_x_relative_from.as_deref(), Some("page"));
        assert_eq!(shape.anchor_y_relative_from.as_deref(), Some("paragraph"));
        assert!(shape.anchor_y_from_para && !shape.anchor_x_from_margin);
        assert_eq!(shape.wrap_mode.as_deref(), Some("none"));
        assert!(matches!(
            &shape.fill,
            Some(docx_model::ShapeFill::Solid { color }) if color == "0000FF"
        ));
        assert_eq!(shape.stroke.as_deref(), Some("FF0000"));
        assert_eq!(shape.stroke_width, 1.0);
        assert_eq!(
            (
                shape.text_inset_l,
                shape.text_inset_t,
                shape.text_inset_r,
                shape.text_inset_b
            ),
            (0.0, 3.6, 7.2, 3.6)
        );
        assert_eq!(shape.text_autofit.as_deref(), Some("sp"));
        let texts: Vec<String> = shape
            .text_box_content
            .iter()
            .map(|block| match block {
                docx_model::TextBoxBlockWire::Body(BodyElement::Paragraph(paragraph)) => paragraph
                    .runs
                    .iter()
                    .filter_map(|run| match run {
                        DocRun::Text(text) => Some(text.text.as_str()),
                        _ => None,
                    })
                    .collect(),
                other => panic!("unexpected textbox block {other:?}"),
            })
            .collect();
        assert_eq!(texts, ["Inside", "box"]);
        assert!(result.resources.is_empty());
    }

    #[test]
    fn grouped_members_share_one_host_and_map_through_nested_groups() {
        let bytes = drawing_source(Some("Inside\r"), false, true, [0, 0]);
        let result =
            super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1024 * 1024).unwrap();
        let BodyElement::Paragraph(paragraph) = &result.document.body[0] else {
            panic!("body paragraph")
        };
        let [DocRun::Text(_), DocRun::AnchorHost(host), DocRun::Shape(first), DocRun::Shape(second)] =
            paragraph.runs.as_slice()
        else {
            panic!("text, host and two members: {:?}", paragraph.runs)
        };
        // The SPA frame is 100 x 50 pt at (5, 10); FSPGR is 1000 x 500.
        for (shape, index, frame) in [
            (first, 0, (0.0, 0.0, 50.0, 25.0)),
            (second, 1, (50.0, 25.0, 50.0, 25.0)),
        ] {
            let acquisition = shape.anchor_acquisition.as_ref().unwrap();
            assert_eq!(
                host.anchor_occurrence_id.as_deref(),
                Some(acquisition.occurrence_id.as_str())
            );
            assert_eq!(acquisition.extent.width_pt, Some(100.0));
            assert_eq!(acquisition.extent.height_pt, Some(50.0));
            let group = acquisition.group.as_ref().unwrap();
            assert_eq!((group.source_index, group.source_count), (index, 2));
            let child = &group.resolved_child_frame;
            assert_eq!(
                (
                    child.offset_x_pt,
                    child.offset_y_pt,
                    child.width_pt,
                    child.height_pt
                ),
                frame
            );
            assert_eq!((shape.width_pt, shape.height_pt), (frame.2, frame.3));
            assert_eq!(
                (shape.anchor_x_pt, shape.anchor_y_pt),
                (5.0 + frame.0, 10.0 + frame.1)
            );
            assert_eq!(
                (shape.group_width_pt, shape.group_height_pt),
                (Some(100.0), Some(50.0))
            );
        }
        assert!(matches!(
            &first.fill,
            Some(docx_model::ShapeFill::Solid { color }) if color == "00FF00"
        ));
        assert!(first.text_box_content.is_empty());
        assert_eq!(second.text_box_content.len(), 1);
        assert_eq!(second.z_order, first.z_order + 1);
    }

    #[test]
    fn rotated_members_and_half_turned_groups_follow_word() {
        let member_with = |text: Option<&str>, turns: [u32; 2]| {
            let bytes = drawing_source(text, false, true, turns);
            let result =
                super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1024 * 1024);
            result.map(|result| {
                let BodyElement::Paragraph(paragraph) = &result.document.body[0] else {
                    panic!("body paragraph")
                };
                let [_, _, DocRun::Shape(first), DocRun::Shape(second)] = paragraph.runs.as_slice()
                else {
                    panic!("two members")
                };
                let frame = |shape: &docx_model::ShapeRun| {
                    let child = &shape
                        .anchor_acquisition
                        .as_ref()
                        .unwrap()
                        .group
                        .as_ref()
                        .unwrap()
                        .resolved_child_frame;
                    (
                        child.offset_x_pt,
                        child.offset_y_pt,
                        child.width_pt,
                        child.height_pt,
                        child.rotation_deg,
                        shape.rotation,
                    )
                };
                (frame(first), frame(second))
            })
        };
        let member = |turns: [u32; 2]| member_with(Some("Inside\r"), turns);
        // The member's 50 x 25 pt stored frame is its rotated bounds at 90
        // degrees: the unrotated box is 25 x 50 pt about the same centre.
        assert_eq!(
            member([0, 90 << 16]).unwrap().0,
            (12.5, -12.5, 25.0, 50.0, 90.0, 90.0)
        );
        // Outside [45, 135) and [225, 315) the stored frame is unrotated.
        let (first, _) = member([0, 0x002c_3f74]).unwrap();
        assert_eq!((first.0, first.1, first.2, first.3), (0.0, 0.0, 50.0, 25.0));
        assert!((first.4 - 44.248).abs() < 1e-3);
        let (first, _) = member([0, 0xff78_6402]).unwrap();
        assert_eq!((first.0, first.1, first.2, first.3), (0.0, 0.0, 50.0, 25.0));
        assert!((first.4 - 224.391).abs() < 1e-3);
        // A half-turned group mirrors member centres and adds 180 degrees; its
        // textbox member makes this group rotate text, which stays rejected.
        assert!(member([180 << 16, 0])
            .unwrap_err()
            .contains("rotated Word drawing text"));
        let (first, second) = member_with(None, [180 << 16, 90 << 16]).unwrap();
        // 100 x 50 pt frame: the rotated member's box mirrors through the
        // centre (its centre 25, 12.5 becomes 75, 37.5) and turns to 270 degrees.
        assert_eq!(first, (62.5, 12.5, 25.0, 50.0, 270.0, 270.0));
        assert_eq!(second, (0.0, 0.0, 50.0, 25.0, 180.0, 180.0));
        // Other group angles stay unsupported.
        assert!(member([90 << 16, 0])
            .unwrap_err()
            .contains("rotated Word drawing groups"));
    }

    #[test]
    fn rotated_or_flipped_groups_fail_closed() {
        let bytes = drawing_source(Some("Inside\r"), false, true, [0, 0]);
        let cfb = CompoundFile::open(&bytes).unwrap();
        let word = cfb.stream("WordDocument").unwrap();
        let table = cfb.stream("0Table").unwrap();
        let fsp = [
            0x02u8, 0x00, 0x0a, 0xf0, 8, 0, 0, 0, 0x02, 0x08, 0, 0, 0x01, 0x02, 0, 0,
        ];
        let at = table
            .windows(fsp.len())
            .position(|bytes| bytes == fsp)
            .expect("top-level group FSP");
        for (offset, value, reason) in [
            (12, 0x41u8, "flipped"),
            (13, 0x06, "unsupported shape flags"),
        ] {
            let mut table = table.clone();
            table[at + offset] = value;
            let changed = build_cfb(&[("WordDocument", word.clone()), ("0Table", table)]);
            assert!(super::super::direct_model(
                &CompoundFile::open(&changed).unwrap(),
                1024 * 1024
            )
            .unwrap_err()
            .contains(reason));
        }
    }

    #[test]
    fn header_shapes_resolve_only_against_header_anchors() {
        let bytes = drawing_shape_source(None, true);
        let result =
            super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1024 * 1024).unwrap();
        let header = result.document.headers.default.as_ref().unwrap();
        let BodyElement::Paragraph(paragraph) = &header.body[0] else {
            panic!("header paragraph")
        };
        assert!(matches!(
            paragraph.runs.as_slice(),
            [DocRun::AnchorHost(_), DocRun::Shape(shape)]
                if shape.preset_geometry.as_deref() == Some("rect") && shape.text_box_content.is_empty()
        ));
        // The same drawing is not reachable from the main story.
        let body = drawing_shape_source(None, false);
        let cfb = CompoundFile::open(&body).unwrap();
        let word = cfb.stream("WordDocument").unwrap();
        let table = cfb.stream("0Table").unwrap();
        let art = u32::from_le_bytes(word[0x22a..0x22e].try_into().unwrap()) as usize;
        let mut table = table;
        table[art + 8] = 1; // Relabel the drawing container as the header's.
        let relabeled = build_cfb(&[("WordDocument", word), ("0Table", table)]);
        assert!(
            super::super::direct_model(&CompoundFile::open(&relabeled).unwrap(), 1024 * 1024)
                .unwrap_err()
                .contains("omitted drawing content")
        );
    }

    #[test]
    fn foreign_textbox_owner_fails_closed() {
        let bytes = drawing_shape_source(Some("Inside\r"), false);
        let cfb = CompoundFile::open(&bytes).unwrap();
        let word = cfb.stream("WordDocument").unwrap();
        let mut table = cfb.stream("0Table").unwrap();
        // Point the FTXBXS at another shape identifier.
        let ftxbxs = u32::from_le_bytes(word[0x25a..0x25e].try_into().unwrap()) as usize;
        table[ftxbxs + 12 + 14] = 3;
        let foreign = build_cfb(&[("WordDocument", word), ("0Table", table)]);
        assert!(
            super::super::direct_model(&CompoundFile::open(&foreign).unwrap(), 1024 * 1024)
                .unwrap_err()
                .contains("another shape")
        );
    }

    pub(super) fn passive_special_source(source: &[u8]) -> Vec<u8> {
        let cfb = CompoundFile::open(source).unwrap();
        let mut word = cfb.stream("WordDocument").unwrap();
        let table = cfb.stream("0Table").unwrap();
        let bte = u32::from_le_bytes(word[0xfa..0xfe].try_into().unwrap()) as usize;
        let page_number = u32::from_le_bytes(table[bte + 8..bte + 12].try_into().unwrap()) as usize;
        let page = &mut word[page_number * 512..(page_number + 1) * 512];
        page[8] = 32;
        page[64..68].copy_from_slice(&[3, 0x55, 0x08, 1]);
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

    fn mark_header_field_results_private(bytes: &[u8]) -> Vec<u8> {
        let cfb = CompoundFile::open(bytes).unwrap();
        let word = cfb.stream("WordDocument").unwrap();
        let mut table = cfb.stream("0Table").unwrap();
        let offset = u32::from_le_bytes(word[0x122..0x126].try_into().unwrap()) as usize;
        let size = u32::from_le_bytes(word[0x126..0x12a].try_into().unwrap()) as usize;
        let count = (size - 4) / 6;
        let records = offset + (count + 1) * 4;
        for index in 0..count {
            if table[records + index * 2] & 0x1f == 0x15 {
                table[records + index * 2 + 1] |= 0x20;
            }
        }
        build_cfb(&[("WordDocument", word), ("0Table", table)])
    }

    #[test]
    fn source_story_projects_directly_with_controls_and_cached_field_result() {
        let bytes = source("A\tB\u{b}C\r\u{13}REF x\u{14}42\u{15}\r\u{c}\r\u{e}\rA\u{c}B\u{e}C\r");
        let cfb = CompoundFile::open(&bytes).unwrap();
        let direct = super::super::direct_model(&cfb, 1024 * 1024)
            .unwrap()
            .document;
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
            crate::doc::paragraph::byte_adapter_line_spacing_parity(body);
        }
        assert_eq!(actual_body, expected_body);
        assert_eq!(actual["section"], expected["section"]);
    }

    #[test]
    fn numbered_paragraphs_project_once_per_source_paragraph_without_ooxml_round_trip() {
        let bytes = numbered_source("A\rB\r");
        let cfb = CompoundFile::open(&bytes).unwrap();
        let direct = super::super::direct_model(&cfb, 1024 * 1024)
            .unwrap()
            .document;
        let markers: Vec<_> = direct
            .body
            .iter()
            .filter_map(|element| match element {
                BodyElement::Paragraph(paragraph) => paragraph.numbering.as_deref(),
                _ => None,
            })
            .map(|numbering| numbering.text.as_str())
            .collect();
        assert_eq!(markers, ["1.", "2."]);

        let converted = super::super::convert(&cfb, 1024 * 1024).unwrap();
        let expected: serde_json::Value =
            serde_json::from_str(&docx_parser::parse_docx_native(&converted.bytes).unwrap())
                .unwrap();
        let actual = serde_json::to_value(&direct).unwrap();
        for index in 0..2 {
            assert_eq!(
                actual["body"][index]["numbering"],
                expected["body"][index]["numbering"]
            );
        }
    }

    #[test]
    fn body_numbering_survives_section_headers_while_each_header_story_restarts() {
        let sections = [
            (2, 2, 12_240, 15_840, 1, 720),
            (4, 2, 12_240, 15_840, 1, 720),
        ];
        let mut slots = [None; 12];
        slots[1] = Some("H\r");
        let bytes = with_numbering(&source_with_typography(
            "A\u{c}B\r",
            &sections,
            None,
            None,
            None,
            Some(&slots),
        ));
        let direct = super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1024 * 1024)
            .unwrap()
            .document;
        let body_markers: Vec<_> = direct
            .body
            .iter()
            .filter_map(|element| match element {
                BodyElement::Paragraph(paragraph) => paragraph.numbering.as_deref(),
                _ => None,
            })
            .map(|numbering| numbering.text.as_str())
            .collect();
        assert_eq!(body_markers, ["1.", "2."]);
        let header_markers: Vec<_> = direct
            .body
            .iter()
            .filter_map(|element| match element {
                BodyElement::SectionBreak { headers, .. } => headers.default.as_ref(),
                _ => None,
            })
            .chain(direct.headers.default.as_ref())
            .map(|header| {
                let BodyElement::Paragraph(paragraph) = &header.body[0] else {
                    panic!("numbered header paragraph")
                };
                paragraph.numbering.as_ref().unwrap().text.as_str()
            })
            .collect();
        assert_eq!(header_markers, ["1.", "1."]);
    }

    fn image_runs(document: &Document) -> Vec<&docx_model::ImageRun> {
        document
            .body
            .iter()
            .filter_map(|element| match element {
                BodyElement::Paragraph(paragraph) => Some(paragraph.runs.iter()),
                _ => None,
            })
            .flatten()
            .filter_map(|run| match run {
                DocRun::Image(image) => Some(image.as_ref()),
                _ => None,
            })
            .collect()
    }

    #[test]
    fn inline_picture_result_owns_one_resource_and_preserves_model_geometry() {
        let bytes = picture_source("\u{1}\u{1}\r", false);
        let result = {
            let cfb = CompoundFile::open(&bytes).unwrap();
            super::super::direct_model(&cfb, 1024 * 1024).unwrap()
        };
        drop(bytes);
        let images = image_runs(&result.document);
        assert_eq!(images.len(), 2);
        assert_eq!(result.resources.len(), 1);
        assert_eq!(images[0].image_path, result.resources[0].key);
        assert_eq!(images[0].mime_type, "image/png");
        assert_eq!((images[0].width_pt, images[0].height_pt), (36.0, 72.0));
        assert_eq!(images[0].rotation, 90.0);
        assert!(images[0].flip_h && images[0].flip_v);
        assert_eq!(images[0].src_rect.as_ref().unwrap().l, 0.375);
        assert!(result.resources[0].bytes.starts_with(b"\x89PNG\r\n\x1a\n"));

        let converted = super::super::convert(
            &CompoundFile::open(&picture_source("\u{1}\r", false)).unwrap(),
            1024 * 1024,
        )
        .unwrap();
        let mut archive = zip::ZipArchive::new(Cursor::new(&converted.bytes)).unwrap();
        let mut document_xml = String::new();
        archive
            .by_name("word/document.xml")
            .unwrap()
            .read_to_string(&mut document_xml)
            .unwrap();
        assert!(document_xml.contains("rot=\"5400000\" flipH=\"1\" flipV=\"1\""));
        let expected: serde_json::Value =
            serde_json::from_str(&docx_parser::parse_docx_native(&converted.bytes).unwrap())
                .unwrap();
        let actual = serde_json::to_value(&result.document).unwrap();
        let mut normalized = actual["body"][0]["runs"][0].clone();
        normalized["imagePath"] = expected["body"][0]["runs"][0]["imagePath"].clone();
        normalized.as_object_mut().unwrap().remove("rotation");
        normalized.as_object_mut().unwrap().remove("flipH");
        normalized.as_object_mut().unwrap().remove("flipV");
        assert_eq!(normalized, expected["body"][0]["runs"][0]);
        // The byte adapter emits these acquired facts on pic:spPr/a:xfrm, but
        // the current DOCX inline parser does not yet project that transform.
        assert!(expected["body"][0]["runs"][0].get("rotation").is_none());
        assert!(expected["body"][0]["runs"][0].get("flipH").is_none());
        assert!(expected["body"][0]["runs"][0].get("flipV").is_none());

        let sections = [
            (2, 0, 12_240, 15_840, 1, 720),
            (4, 2, 12_240, 15_840, 1, 720),
        ];
        let slots = [
            None,
            Some("\u{1}\r"),
            None,
            Some("\u{1}\r"),
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
        ];
        let header_source =
            source_with_typography("B\u{c}\u{1}\r", &sections, None, None, None, Some(&slots));
        let header_result = super::super::direct_model(
            &CompoundFile::open(&with_picture_data(&header_source, false)).unwrap(),
            1024 * 1024,
        )
        .unwrap();
        assert_eq!(header_result.resources.len(), 1);
        let body_key = image_runs(&header_result.document)[0].image_path.clone();
        let header = header_result.document.headers.default.as_ref().unwrap();
        let BodyElement::Paragraph(paragraph) = &header.body[0] else {
            panic!("header paragraph")
        };
        let DocRun::Image(header_image) = &paragraph.runs[0] else {
            panic!("header image")
        };
        assert_eq!(body_key, header_image.image_path);
        let footer = header_result.document.footers.default.as_ref().unwrap();
        let BodyElement::Paragraph(paragraph) = &footer.body[0] else {
            panic!("footer paragraph")
        };
        let DocRun::Image(footer_image) = &paragraph.runs[0] else {
            panic!("footer image")
        };
        assert_eq!(body_key, footer_image.image_path);
        let (break_headers, break_footers) = header_result
            .document
            .body
            .iter()
            .find_map(|block| match block {
                BodyElement::SectionBreak {
                    headers, footers, ..
                } => Some((headers, footers)),
                _ => None,
            })
            .expect("section break");
        for slot in [
            break_headers.default.as_ref().unwrap(),
            break_footers.default.as_ref().unwrap(),
        ] {
            let BodyElement::Paragraph(paragraph) = &slot.body[0] else {
                panic!("section header or footer paragraph")
            };
            let DocRun::Image(image) = &paragraph.runs[0] else {
                panic!("section header or footer image")
            };
            assert_eq!(body_key, image.image_path);
        }
        assert_eq!(body_key, header_result.resources[0].key);

        let hidden = picture_source("\u{1}\r", true);
        let hidden =
            super::super::direct_model(&CompoundFile::open(&hidden).unwrap(), 1024 * 1024).unwrap();
        assert!(image_runs(&hidden.document).is_empty());
        assert!(hidden.resources.is_empty());

        let bytes = picture_source("\u{1}\r", false);
        let cfb = CompoundFile::open(&bytes).unwrap();
        let mut low = 1usize;
        let sufficient = loop {
            if super::super::direct_model(&cfb, low).is_ok() {
                break low;
            }
            low *= 2;
        };
        let mut left = sufficient / 2;
        let mut right = sufficient;
        while left + 1 < right {
            let middle = left + (right - left) / 2;
            if super::super::direct_model(&cfb, middle).is_ok() {
                right = middle;
            } else {
                left = middle;
            }
        }
        assert_eq!(
            super::super::direct_model(&cfb, right - 1).unwrap_err(),
            "OUTPUT_TOO_LARGE"
        );
        assert_eq!(
            image_runs(&super::super::direct_model(&cfb, right).unwrap().document).len(),
            1
        );
    }

    #[test]
    fn floating_picture_projects_anchor_host_sidecar_and_owned_resource() {
        let bytes = floating_picture_source("B\u{8}\r", false);
        let result = {
            let cfb = CompoundFile::open(&bytes).unwrap();
            super::super::direct_model(&cfb, 1024 * 1024).unwrap()
        };
        drop(bytes);
        assert_eq!(result.resources.len(), 1);
        assert!(result.resources[0].bytes.starts_with(b"\x89PNG"));
        let BodyElement::Paragraph(paragraph) = &result.document.body[0] else {
            panic!("body paragraph")
        };
        let [DocRun::Text(_), DocRun::AnchorHost(host), DocRun::Image(image)] =
            paragraph.runs.as_slice()
        else {
            panic!("text, anchor host, image")
        };
        let acquisition = image.anchor_acquisition.as_ref().unwrap();
        assert_eq!(
            host.anchor_occurrence_id.as_deref(),
            Some(acquisition.occurrence_id.as_str())
        );
        assert_eq!(image.image_path, result.resources[0].key);
        assert!(image.anchor && image.flip_h && image.flip_v);
        assert_eq!((image.anchor_x_pt, image.anchor_y_pt), (-5.0, 10.0));
        assert_eq!((image.width_pt, image.height_pt), (20.0, 15.0));
        assert_eq!(image.wrap_mode.as_deref(), Some("square"));
        assert_eq!(image.wrap_side.as_deref(), Some("right"));
        assert_eq!(
            (
                image.dist_left,
                image.dist_top,
                image.dist_right,
                image.dist_bottom
            ),
            (1.0, 2.0, 3.0, 4.0)
        );
        assert!(!image.allow_overlap);
        assert_eq!(image.anchor_x_relative_from.as_deref(), Some("page"));
        assert_eq!(image.anchor_y_relative_from.as_deref(), Some("paragraph"));
        assert_eq!(acquisition.wrap.authored_kinds, ["wrapSquare"]);
        assert_eq!(acquisition.behavior.relative_height, Some(77));
        assert_eq!(acquisition.behavior.locked, Some(true));
        assert_eq!(acquisition.behavior.layout_in_cell, Some(false));
        assert_eq!(acquisition.behavior.allow_overlap, Some(false));
        assert_eq!(acquisition.anchor_distances.left_pt, Some(1.0));

        // Cross-check the complete host and acquisition contract, not only
        // selected display fields. The byte parser's known picture-flip loss
        // remains explicit; direct projection preserves the authored flips.
        let reference_source = floating_picture_source("B\u{8}\r", false);
        let converted =
            super::super::convert(&CompoundFile::open(&reference_source).unwrap(), 1024 * 1024)
                .unwrap();
        let expected: serde_json::Value =
            serde_json::from_str(&docx_parser::parse_docx_native(&converted.bytes).unwrap())
                .unwrap();
        let actual = serde_json::to_value(&result.document).unwrap();
        let expected_runs = &expected["body"][0]["runs"];
        let mut actual_host = actual["body"][0]["runs"][1].clone();
        actual_host["__anchorOccurrenceId"] = expected_runs[1]["__anchorOccurrenceId"].clone();
        assert_eq!(actual_host, expected_runs[1]);
        let mut actual_image = actual["body"][0]["runs"][2].clone();
        actual_image["imagePath"] = expected_runs[2]["imagePath"].clone();
        actual_image["__anchorAcquisition"]["occurrenceId"] =
            expected_runs[2]["__anchorAcquisition"]["occurrenceId"].clone();
        assert!(expected_runs[2].get("flipH").is_none());
        assert!(expected_runs[2].get("flipV").is_none());
        actual_image.as_object_mut().unwrap().remove("flipH");
        actual_image.as_object_mut().unwrap().remove("flipV");
        assert_eq!(actual_image, expected_runs[2]);

        let hidden_bytes = floating_picture_source("B\u{8}\r", true);
        let hidden =
            super::super::direct_model(&CompoundFile::open(&hidden_bytes).unwrap(), 1024 * 1024)
                .unwrap();
        assert!(image_runs(&hidden.document).is_empty());
        assert!(hidden.resources.is_empty());

        let bytes = floating_picture_source("B\u{8}\r", false);
        let cfb = CompoundFile::open(&bytes).unwrap();
        assert_eq!(
            super::super::direct_model(&cfb, 1).unwrap_err(),
            "OUTPUT_TOO_LARGE"
        );
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
            let bytes =
                source_with_typography("Body\r", &sections, None, normal_hps, body_hps, None);
            let document =
                super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1024 * 1024)
                    .unwrap()
                    .document;
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

        let malformed = source_with_typography("Body\r", &sections, None, Some(1), Some(44), None);
        assert!(
            super::super::direct_model(&CompoundFile::open(&malformed).unwrap(), 1024 * 1024,)
                .is_err()
        );

        let valid = source_with_typography("Body\r", &sections, None, Some(22), Some(44), None);
        let cyclic = make_normal_style_self_referential(&valid);
        assert!(
            super::super::direct_model(&CompoundFile::open(&cyclic).unwrap(), 1024 * 1024,)
                .unwrap_err()
                .contains("cyclic")
        );
    }

    #[test]
    fn proofing_only_chpx_reaches_the_native_document_without_display_state() {
        let bytes = source_with_proofing("Proof\r", true);
        let document =
            super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1024 * 1024)
                .unwrap()
                .document;
        let BodyElement::Paragraph(paragraph) = &document.body[0] else {
            panic!("expected body paragraph");
        };
        let DocRun::Text(run) = &paragraph.runs[0] else {
            panic!("expected text run");
        };
        assert_eq!(run.text, "Proof");
        assert!(run.bold);

        let baseline_bytes = source_with_proofing("Proof\r", false);
        let baseline =
            super::super::direct_model(&CompoundFile::open(&baseline_bytes).unwrap(), 1024 * 1024)
                .unwrap()
                .document;
        let BodyElement::Paragraph(baseline_paragraph) = &baseline.body[0] else {
            panic!("expected baseline paragraph");
        };
        let DocRun::Text(baseline_run) = &baseline_paragraph.runs[0] else {
            panic!("expected baseline text run");
        };
        assert_eq!(
            serde_json::to_value(run.as_ref()).unwrap(),
            serde_json::to_value(baseline_run.as_ref()).unwrap()
        );
    }

    #[test]
    fn separator_only_header_document_is_not_an_authored_page_header() {
        let section = [(5, 2, 12_240, 15_840, 1, 720)];
        let plain = source("Body\r");
        let separators = source_with_sections_and_header("Body\r", &section, Some(false));
        let blank_header = source_with_sections_and_header("Body\r", &section, Some(true));
        let expected =
            super::super::direct_model(&CompoundFile::open(&plain).unwrap(), 1024 * 1024)
                .unwrap()
                .document;
        let actual =
            super::super::direct_model(&CompoundFile::open(&separators).unwrap(), 1024 * 1024)
                .unwrap()
                .document;
        assert_eq!(
            serde_json::to_value(actual).unwrap(),
            serde_json::to_value(expected).unwrap()
        );
        let blank =
            super::super::direct_model(&CompoundFile::open(&blank_header).unwrap(), 1024 * 1024)
                .unwrap()
                .document;
        let authored = blank.headers.even.as_ref().expect("authored even header");
        assert_eq!(authored.body.len(), 1);
        let BodyElement::Paragraph(paragraph) = &authored.body[0] else {
            panic!("authored blank header remains an empty paragraph")
        };
        assert!(paragraph.runs.is_empty());

        let converted =
            super::super::convert(&CompoundFile::open(&blank_header).unwrap(), 1024 * 1024)
                .unwrap();
        let expected: serde_json::Value =
            serde_json::from_str(&docx_parser::parse_docx_native(&converted.bytes).unwrap())
                .unwrap();
        let actual = serde_json::to_value(blank).unwrap();
        assert_eq!(
            actual["headers"]["even"]["body"].as_array().unwrap().len(),
            1
        );
        assert_eq!(
            actual["headers"]["even"]["body"][0]["runs"],
            expected["headers"]["even"]["body"][0]["runs"]
        );
    }

    fn header_text(value: &docx_model::HeaderFooter) -> String {
        value
            .body
            .iter()
            .filter_map(|element| match element {
                BodyElement::Paragraph(paragraph) => Some(
                    paragraph
                        .runs
                        .iter()
                        .filter_map(|run| match run {
                            DocRun::Text(text) => Some(text.text.as_str()),
                            _ => None,
                        })
                        .collect::<String>(),
                ),
                _ => None,
            })
            .collect()
    }

    fn normalized_header_json(value: &serde_json::Value) -> serde_json::Value {
        let mut value = value.clone();
        for slot in ["even", "default", "first"] {
            let Some(body) = value[slot]["body"].as_array_mut() else {
                continue;
            };
            for element in body {
                if element["type"] == "paragraph" {
                    element.as_object_mut().unwrap().remove("styleId");
                }
            }
        }
        crate::doc::paragraph::byte_adapter_line_spacing_parity(&mut value);
        value
    }

    #[test]
    fn all_six_slots_inherit_and_authored_blank_clears_only_its_slot() {
        let sections = [
            (2, 0, 12_240, 15_840, 1, 720),
            (4, 2, 12_240, 15_840, 1, 720),
            (6, 2, 12_240, 15_840, 1, 720),
        ];
        let slots = [
            Some("HE😀\r"),
            Some("HD\r"),
            Some("FE\r"),
            Some("FD\r"),
            Some("HF\r"),
            Some("FF\r"),
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            Some("\r"),
            None,
            None,
            None,
            None,
        ];
        let bytes =
            source_with_typography("A\u{c}B\u{c}C\r", &sections, None, None, None, Some(&slots));
        let document =
            super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1024 * 1024)
                .unwrap()
                .document;
        let converted =
            super::super::convert(&CompoundFile::open(&bytes).unwrap(), 1024 * 1024).unwrap();
        let expected: serde_json::Value =
            serde_json::from_str(&docx_parser::parse_docx_native(&converted.bytes).unwrap())
                .unwrap();
        let actual = serde_json::to_value(&document).unwrap();
        assert_eq!(
            normalized_header_json(&actual["headers"]),
            normalized_header_json(&expected["headers"])
        );
        assert_eq!(
            normalized_header_json(&actual["footers"]),
            normalized_header_json(&expected["footers"])
        );
        let breaks: Vec<_> = document
            .body
            .iter()
            .filter_map(|element| match element {
                BodyElement::SectionBreak {
                    headers, footers, ..
                } => Some((headers.as_ref(), footers.as_ref())),
                _ => None,
            })
            .collect();
        assert_eq!(breaks.len(), 2);
        for (headers, footers) in &breaks {
            assert_eq!(header_text(headers.even.as_ref().unwrap()), "HE😀");
            assert_eq!(header_text(headers.default.as_ref().unwrap()), "HD");
            assert_eq!(header_text(headers.first.as_ref().unwrap()), "HF");
            assert_eq!(header_text(footers.even.as_ref().unwrap()), "FE");
            assert_eq!(header_text(footers.default.as_ref().unwrap()), "FD");
            assert_eq!(header_text(footers.first.as_ref().unwrap()), "FF");
        }
        assert_eq!(header_text(document.headers.even.as_ref().unwrap()), "HE😀");
        assert_eq!(header_text(document.headers.first.as_ref().unwrap()), "HF");
        assert!(document
            .headers
            .default
            .as_ref()
            .is_some_and(|header| header_text(header).is_empty()));
        assert_eq!(header_text(document.footers.even.as_ref().unwrap()), "FE");
        assert_eq!(
            header_text(document.footers.default.as_ref().unwrap()),
            "FD"
        );
        assert_eq!(header_text(document.footers.first.as_ref().unwrap()), "FF");
    }

    #[test]
    fn header_page_fields_project_renderer_evaluated_field_runs() {
        let sections = [(5, 2, 12_240, 15_840, 1, 720)];
        let slots = [
            None,
            Some("P\u{13} PAGE \\* roman \u{14}42\u{15}/\u{13}NUMPAGES\u{15}\r"),
            None,
            None,
            None,
            None,
        ];
        let bytes = source_with_typography("Body\r", &sections, None, None, None, Some(&slots));
        let document =
            super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1024 * 1024)
                .unwrap()
                .document;
        let header = document.headers.default.as_ref().unwrap();
        let BodyElement::Paragraph(paragraph) = &header.body[0] else {
            panic!("header paragraph");
        };
        let runs: Vec<_> = paragraph
            .runs
            .iter()
            .map(|run| match run {
                DocRun::Text(text) => format!("text:{}", text.text),
                DocRun::Field(field) => format!(
                    "{}:{}:{}",
                    field.field_type, field.instruction, field.fallback_text
                ),
                _ => "other".into(),
            })
            .collect();
        assert_eq!(
            runs,
            [
                "text:P",
                "page:PAGE \\* roman:42",
                "text:/",
                "numPages:NUMPAGES:"
            ]
        );
    }

    #[test]
    fn unverified_private_field_results_are_rejected_not_hidden() {
        let sections = [(5, 2, 12_240, 15_840, 1, 720)];
        let slots = [
            None,
            Some("\u{13}IF\u{14}x\u{15}\r"),
            None,
            None,
            None,
            None,
        ];
        let bytes = source_with_typography("Body\r", &sections, None, None, None, Some(&slots));
        assert!(
            super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1024 * 1024).is_ok()
        );
        let bytes = mark_header_field_results_private(&bytes);
        assert!(
            super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1024 * 1024)
                .unwrap_err()
                .contains("private field result")
        );
    }

    #[test]
    fn inherited_hidden_header_work_is_cumulatively_bounded_before_retokenization() {
        let section_count = 8usize;
        let mut story = String::new();
        let mut sections = Vec::new();
        for section in 0..section_count {
            story.push(if section + 1 == section_count {
                '\r'
            } else {
                '\u{c}'
            });
            sections.push((section + 1, 2, 12_240, 15_840, 1, 720));
        }
        let hidden = format!("\u{13}IF {}\u{15}\r", "x".repeat(16 * 1024));
        let mut slots = vec![None; section_count * 6];
        slots[1] = Some(hidden.as_str());
        let bytes = source_with_typography(&story, &sections, None, None, None, Some(&slots));
        let projected =
            super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1024 * 1024)
                .unwrap()
                .document;
        assert!(projected
            .headers
            .default
            .as_ref()
            .is_some_and(|header| header_text(header).is_empty()));
        assert_eq!(
            super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 192 * 1024)
                .unwrap_err(),
            "OUTPUT_TOO_LARGE"
        );
    }

    #[test]
    fn output_budget_and_unimplemented_body_owners_fail_without_a_document() {
        let bytes = source("visible\r");
        let cfb = CompoundFile::open(&bytes).unwrap();
        assert_eq!(
            super::super::direct_model(&cfb, 1).unwrap_err(),
            "OUTPUT_TOO_LARGE"
        );
        for text in ["cell\u{7}"] {
            let bytes = source(text);
            let cfb = CompoundFile::open(&bytes).unwrap();
            assert!(super::super::direct_model(&cfb, 1024 * 1024)
                .unwrap_err()
                .starts_with("UNSUPPORTED:"));
        }
        let missing_location = source("\u{1}\r");
        assert!(super::super::direct_model(
            &CompoundFile::open(&missing_location).unwrap(),
            1024 * 1024,
        )
        .unwrap_err()
        .contains("no inline location"));
        let non_passive_float = source("\u{8}\r");
        assert!(super::super::direct_model(
            &CompoundFile::open(&non_passive_float).unwrap(),
            1024 * 1024,
        )
        .unwrap_err()
        .contains("not passive-special"));
        let sections = [(2, 2, 12_240, 15_840, 1, 720)];
        let slots = [None, Some("\u{8}\r"), None, None, None, None];
        let header_float = source_with_typography("B\r", &sections, None, None, None, Some(&slots));
        let header_float = passive_special_source(&header_float);
        assert!(super::super::direct_model(
            &CompoundFile::open(&header_float).unwrap(),
            1024 * 1024,
        )
        .unwrap_err()
        .contains("omitted drawing content"));
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
            let result = super::super::with_acquired_doc(&cfb, true, |mut facts| {
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
        let direct = super::super::direct_model(&cfb, 1024 * 1024)
            .unwrap()
            .document;
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
            crate::doc::paragraph::byte_adapter_line_spacing_parity(&mut body);
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
            let direct = super::super::direct_model(&cfb, 1024 * 1024)
                .unwrap()
                .document;
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
                crate::doc::paragraph::byte_adapter_line_spacing_parity(&mut body);
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
            let direct = super::super::direct_model(&cfb, 1024 * 1024)
                .unwrap()
                .document;
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
                crate::doc::paragraph::byte_adapter_line_spacing_parity(&mut body);
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
            let direct = super::super::direct_model(&cfb, 1024 * 1024)
                .unwrap()
                .document;
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
                crate::doc::paragraph::byte_adapter_line_spacing_parity(&mut body);
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
            let direct = super::super::direct_model(&cfb, 1024 * 1024)
                .unwrap()
                .document;
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
                crate::doc::paragraph::byte_adapter_line_spacing_parity(&mut body);
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
