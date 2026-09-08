//! Internal direct DOC body producer. This is intentionally not a public route:
//! unsupported body owners fail closed while incremental model coverage lands.

use super::{unsupported, AcquiredDoc, Fields};
use docx_model::paragraph_breaks::ParaPiece;
use docx_model::{BodyElement, DocRun, Document, DocumentSettings, DocumentTypographySettingsWire};

mod headers;
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
    if facts.note_stories.iter().any(Option::is_some) {
        return Err(unsupported("direct DOC model does not yet support notes"));
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
    let mut body = Vec::new();
    let mut header_resolver = headers::Resolver::new(facts.headers.as_ref());
    let mut final_headers = None;
    let mut final_footers = None;
    let chunks = super::sections::split_story(&facts.story.text, &facts.sections)?;
    let mut fields = Fields::default();
    let mut table_sequence = 0;
    let mut numbering = super::numbering::direct::Store::default();
    numbering.begin_story()?;
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

        story::project(
            &facts.story,
            paragraphs,
            &mut facts.formatting,
            &mut numbering,
            &mut facts.pictures,
            Some(&mut facts.floating),
            &mut budget,
            &mut body,
            ending.as_ref().map(|ending| ending.kind.as_str()),
            &mut table_sequence,
        )?;

        let (section_headers, section_footers) = header_resolver.project_section(
            section_index,
            &mut facts.formatting,
            &mut facts.pictures,
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

    let mut resources = facts
        .pictures
        .finish_direct_resources(&mut budget.remaining_bytes)?;
    facts
        .floating
        .append_direct_resources(&mut resources, &mut budget.remaining_bytes)?;
    let document = Document {
        section,
        body,
        headers: final_headers.unwrap_or_default(),
        footers: final_footers.unwrap_or_default(),
        settings,
        document_typography_settings,
        ..Document::default()
    };
    Ok(DirectDocResult {
        document,
        resources,
    })
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
        let main_units = text.encode_utf16().count();
        let (header, header_cps) = if let Some(slots) = header_slots {
            assert_eq!(slots.len(), sections.len() * 6);
            let mut header = String::from("\r");
            let mut cps = vec![0u32, 1, 1, 1, 1, 1, 1];
            let mut cp = 1u32;
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
        let mut field_positions = Vec::new();
        let mut cp = 0u32;
        for character in header.chars() {
            if matches!(character, '\u{13}'..='\u{15}') {
                field_positions.push((cp, character as u8));
            }
            cp += character.len_utf16() as u32;
        }
        if !field_positions.is_empty() {
            let mut field_table = Vec::new();
            for (position, _) in &field_positions {
                field_table.extend(position.to_le_bytes());
            }
            field_table.extend(cp.to_le_bytes());
            for (_, marker) in field_positions {
                field_table.extend([marker, 0]);
            }
            append_table_part(&mut word, &mut table, 0x122, &field_table);
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

    fn numbered_source(text: &str) -> Vec<u8> {
        with_numbering(&source(text))
    }

    fn with_numbering(source: &[u8]) -> Vec<u8> {
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

    fn passive_special_source(source: &[u8]) -> Vec<u8> {
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
        let bytes = source("A\tB\u{b}C\r\u{13}PAGE\u{14}42\u{15}\r\u{c}\r\u{e}\rA\u{c}B\u{e}C\r");
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

        let sections = [(3, 2, 12_240, 15_840, 1, 720)];
        let slots = [None, Some("\u{1}\r"), None, None, None, None];
        let header_source =
            source_with_typography("B\u{1}\r", &sections, None, None, None, Some(&slots));
        let header_result = super::super::direct_model(
            &CompoundFile::open(&with_picture_data(&header_source, false)).unwrap(),
            1024 * 1024,
        )
        .unwrap();
        assert_eq!(header_result.resources.len(), 1);
        let body_key = image_runs(&header_result.document)[0].image_path.clone();
        let header = header_result.document.headers.default.unwrap();
        let BodyElement::Paragraph(paragraph) = &header.body[0] else {
            panic!("header paragraph")
        };
        let DocRun::Image(header_image) = &paragraph.runs[0] else {
            panic!("header image")
        };
        assert_eq!(body_key, header_image.image_path);
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
    fn restored_page_field_is_rejected_instead_of_exposing_cached_digits_as_text() {
        let sections = [(5, 2, 12_240, 15_840, 1, 720)];
        let slots = [
            None,
            Some("\u{13}PAGE\u{14}42\u{15}\r"),
            None,
            None,
            None,
            None,
        ];
        let bytes = source_with_typography("Body\r", &sections, None, None, None, Some(&slots));
        assert!(
            super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1024 * 1024)
                .unwrap_err()
                .contains("field structures")
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
        let hidden = format!("\u{13}IF\u{14}{}\u{15}\r", "x".repeat(16 * 1024));
        let mut slots = vec![None; section_count * 6];
        slots[1] = Some(hidden.as_str());
        let bytes = source_with_typography(&story, &sections, None, None, None, Some(&slots));
        let bytes = mark_header_field_results_private(&bytes);
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
        .contains("header floating pictures"));
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
