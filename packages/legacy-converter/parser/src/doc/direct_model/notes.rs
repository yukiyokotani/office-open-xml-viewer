//! Direct projection of footnotes and endnotes onto the DOCX model's notes.
//!
//! MS-DOC 2.3.2/2.3.5 (note documents), 2.8.16/2.8.17 (PlcffndRef/Txt),
//! 2.8.19/2.8.20 (PlcfendRef/Txt), 2.3.3 (separator stories), 2.7.2/2.7.4
//! (DOP note properties) and 2.6.4 (section note properties); ECMA-376
//! Part 1 17.11.
//!
//! The DOCX parser represents a note as a `DocxNote` whose blocks are parsed
//! like any story, a body `footnoteReference` as a superscript `TextRun`
//! tagged `NoteRef { kind, id }`, and the in-note `footnoteRef` as the same run
//! with an empty id. The direct model produces exactly that shape.
//!
//! The shared renderer numbers notes in first-reference order with the
//! document-wide numbering format and start (ECMA-376 17.11.17/.18/.20), lays
//! footnotes out at the page bottom and endnotes at
//! the end of the document, and draws its fixed default separator (the DOCX
//! parser likewise ignores separator notes). A document is accepted only when
//! Word's effective note properties produce that display: automatic marks,
//! continuous numbering with one format and start per note kind (Arabic,
//! Roman or letters), bottom-of-page footnotes,
//! end-of-document endnotes and the standard separator stories. Everything
//! else stays rejected; custom marks additionally need renderer support for
//! literal marks that do not consume a number.

use super::{fields::StoryFields, story, ModelBudget};
use crate::doc::{
    formatting, header_fields, headers, notes, numbering, pictures, sections, settings,
    tokenize_with_fields, unsupported, Fields, Paragraph, Token,
};
use docx_model::{DocxNote, NoteLayoutSettingsWire};

/// Last effective nFib whose note properties live in the DOP (MS-DOC 2.7.2).
const DOP_NOTE_PROPERTIES_MAX_NFIB: u16 = 0x00d9;

/// Reject note properties the shared renderer cannot display as Word does and
/// return the document-wide note numbering (ECMA-376 17.11.17/.18/.20) the
/// renderer applies to automatic reference marks.
pub(super) fn validate(
    stories: &[Option<notes::Notes<'_>>],
    references: &notes::References,
    effective_nfib: u16,
    dop: Option<&settings::Properties>,
    sections: &[sections::Section],
    headers: Option<&headers::Headers<'_>>,
) -> Result<Option<NoteLayoutSettingsWire>, String> {
    let present = |kind| {
        stories
            .iter()
            .flatten()
            .any(|notes| notes.kind == kind && !notes.entries.is_empty())
    };
    let (footnotes, endnotes) = (
        present(notes::Kind::Footnote),
        present(notes::Kind::Endnote),
    );
    if !footnotes && !endnotes {
        return Ok(None);
    }
    if references.iter().any(notes::Reference::custom) {
        return Err(unsupported(
            "Word custom note reference marks are not supported",
        ));
    }
    let dop = dop
        .map(|dop| dop.notes)
        .ok_or_else(|| unsupported("Word notes require document note properties"))?;
    // (format, start) for each kind: MSONFC and the first automatic number.
    let (footnote_numbering, endnote_numbering);
    if effective_nfib <= DOP_NOTE_PROPERTIES_MAX_NFIB {
        // MS-DOC 2.7.2/2.7.4: these documents keep note numbering and footnote
        // placement in the DOP. Section note SPRMs would be ambiguous.
        let formats = dop
            .formats
            .ok_or_else(|| unsupported("Word note number formats are missing"))?;
        if sections
            .iter()
            .any(|section| section.note_properties() != Default::default())
        {
            return Err(unsupported(
                "Word section note properties in a DOP-scoped document",
            ));
        }
        if footnotes && (dop.footnote_position != 1 || dop.footnote_restart != 0) {
            return Err(unsupported(footnote_message()));
        }
        if endnotes && dop.endnote_restart != 0 {
            return Err(unsupported(endnote_message()));
        }
        footnote_numbering = (formats.0, dop.footnote_start);
        endnote_numbering = (formats.1, dop.endnote_start);
    } else {
        // MS-DOC 2.6.4 defaults: fpcBottomPage, rncCont, no offset, Arabic
        // footnotes and lowercase-Roman endnotes. With continuous numbering
        // sprmSNFtn/sprmSNEdn add (value - 1) to every number of the section;
        // equal values in every section are a document-wide start value.
        let mut footnote_values = None;
        let mut endnote_values = None;
        for section in sections {
            let properties = section.note_properties();
            if footnotes
                && (!matches!(properties.footnote_position, None | Some(1))
                    || !matches!(properties.footnote_restart, None | Some(0)))
            {
                return Err(unsupported(footnote_message()));
            }
            if endnotes && !matches!(properties.endnote_restart, None | Some(0)) {
                return Err(unsupported(endnote_message()));
            }
            let footnote = (
                properties.footnote_format.unwrap_or(0),
                properties.footnote_offset.unwrap_or(1),
            );
            let endnote = (
                properties.endnote_format.unwrap_or(2),
                properties.endnote_offset.unwrap_or(1),
            );
            if *footnote_values.get_or_insert(footnote) != footnote && footnotes {
                return Err(unsupported(footnote_message()));
            }
            if *endnote_values.get_or_insert(endnote) != endnote && endnotes {
                return Err(unsupported(endnote_message()));
            }
        }
        footnote_numbering = footnote_values.unwrap_or((0, 1));
        endnote_numbering = endnote_values.unwrap_or((2, 1));
    }
    let settings = NoteLayoutSettingsWire {
        footnote_number_format: footnotes
            .then(|| number_format(footnote_numbering.0, footnote_message()))
            .transpose()?,
        footnote_number_start: footnotes
            .then(|| number_start(footnote_numbering.1, footnote_message()))
            .transpose()?,
        endnote_number_format: endnotes
            .then(|| number_format(endnote_numbering.0, endnote_message()))
            .transpose()?,
        endnote_number_start: endnotes
            .then(|| number_start(endnote_numbering.1, endnote_message()))
            .transpose()?,
        ..NoteLayoutSettingsWire::default()
    };
    // DopBase.epc is not scoped by nFib. Only end-of-document placement (3)
    // matches the shared layout; sprmSFEndnote is then irrelevant.
    if endnotes && dop.endnote_position != 3 {
        return Err(unsupported(
            "Word end-of-section endnote placement is not supported",
        ));
    }
    // MS-DOC 2.3.3 separator stories. Word writes the separator (U+0003) and
    // continuation separator (U+0004) characters as one paragraph followed by
    // the guard mark; the continuation notice is empty or one empty paragraph.
    // Their paragraph formatting is not retained, as for DOCX separator notes.
    let headers = headers
        .ok_or_else(|| unsupported("Word notes without separator stories are not supported"))?;
    for (present, base) in [(footnotes, 0), (endnotes, 3)] {
        if !present {
            continue;
        }
        if headers.separator_text(base) != "\u{3}\r\r"
            || headers.separator_text(base + 1) != "\u{4}\r\r"
            || !matches!(headers.separator_text(base + 2), "" | "\r\r")
        {
            return Err(unsupported("custom Word note separators are not supported"));
        }
    }
    Ok(Some(settings))
}

/// MS-OSHARED 2.2.1.3 MSONFC values whose ECMA-376 17.18.59 format the shared
/// note renderer formats natively: Arabic, upper/lower Roman and letters.
fn number_format(value: u16, message: &'static str) -> Result<String, String> {
    match value {
        0 => Ok("decimal"),
        1 => Ok("upperRoman"),
        2 => Ok("lowerRoman"),
        3 => Ok("upperLetter"),
        4 => Ok("lowerLetter"),
        _ => Err(unsupported(message)),
    }
    .map(str::to_string)
}

/// A first automatic number of at least 1 (the 14-bit DOP value or the
/// section SPRM value, at most 16383).
fn number_start(value: u16, message: &'static str) -> Result<i64, String> {
    if value == 0 {
        return Err(unsupported(message));
    }
    Ok(i64::from(value))
}

fn footnote_message() -> &'static str {
    "Word footnote numbering or placement other than continuous Roman, letter or Arabic numbering at the page bottom is not supported"
}

fn endnote_message() -> &'static str {
    "Word endnote numbering other than continuous Roman, letter or Arabic numbering is not supported"
}

/// Turn main-story automatic note characters into references. Every
/// character must name a reference and every reference must be displayed.
pub(super) fn restore_references(
    references: &notes::References,
    paragraphs: &mut [Paragraph],
    placed: &mut usize,
) -> Result<(), String> {
    for paragraph in paragraphs {
        for (token, cp) in &mut paragraph.tokens {
            let token = match token {
                Token::Linked(linked) => &mut linked.token,
                token => token,
            };
            if matches!(token, Token::NoteMarker) {
                let reference = references.get(*cp).ok_or_else(|| {
                    unsupported("Word automatic note character without a note reference")
                })?;
                *token = Token::NoteReference(reference.clone());
                *placed += 1;
            }
        }
    }
    Ok(())
}

pub(super) fn check_all_references_placed(
    references: &notes::References,
    placed: usize,
) -> Result<(), String> {
    if placed != references.iter().count() {
        return Err(unsupported("Word note reference is not displayed"));
    }
    Ok(())
}

/// Project every note of one note document, in PLC order, as DOCX notes whose
/// ids match the main-story references (`Reference::id`).
pub(super) fn project(
    notes: &notes::Notes<'_>,
    table: &header_fields::Table,
    formatting: &mut formatting::Formatting<'_>,
    pictures: &mut pictures::Store<'_>,
    budget: &mut ModelBudget,
    table_sequence: &mut usize,
    output: &mut Vec<DocxNote>,
) -> Result<(), String> {
    let mut partitions = Vec::with_capacity(notes.entries.len() + 1);
    let mut end = 0;
    for entry in &notes.entries {
        partitions.push(entry.cp);
        end = entry.cp + notes.story.text[entry.text.clone()].encode_utf16().count();
    }
    partitions.push(end);
    let fields = StoryFields::analyze(&notes.story.text, table, &partitions)?;
    for (index, entry) in notes.entries.iter().enumerate() {
        let text = &notes.story.text[entry.text.clone()];
        budget.charge(text.len())?;
        let mut paragraphs = tokenize_with_fields(text, &mut Fields::default(), entry.cp, true);
        fields.apply(entry.cp, &mut paragraphs)?;
        for paragraph in &mut paragraphs {
            for (token, _) in &mut paragraph.tokens {
                let token = match token {
                    Token::Linked(linked) => &mut linked.token,
                    token => token,
                };
                if matches!(token, Token::NoteMarker) {
                    // `validate` rejected custom marks, so every note is
                    // automatic; each U+0002 shows the enclosing note number.
                    *token = Token::NoteNumber(notes.kind);
                }
            }
        }
        // Lists restart in every note.
        let mut numbering = numbering::direct::Store::default();
        numbering.begin_story()?;
        let mut content = Vec::new();
        story::project(
            &notes.story,
            paragraphs,
            formatting,
            &mut numbering,
            pictures,
            None,
            budget,
            &mut content,
            None,
            table_sequence,
        )?;
        let id = (index + 1).to_string();
        budget.charge(id.capacity())?;
        budget.push(output, DocxNote { id, content })?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::super::tests::{
        default_note_dop, passive_special_source, source_with_stories, NotesFixture,
    };
    use crate::cfb::CompoundFile;
    use docx_model::{BodyElement, DocRun, Document};

    const SEPARATORS: [&str; 6] = ["\u{3}\r\r", "\u{4}\r\r", "", "", "", ""];

    fn document(text: &str, fixture: &NotesFixture<'_>) -> Result<Document, String> {
        let sections = [(text.encode_utf16().count(), 2, 12_240, 15_840, 1, 720)];
        let slots = [None; 6];
        let bytes = source_with_stories(
            text,
            &sections,
            None,
            None,
            None,
            Some(&slots),
            Some(fixture),
        );
        // MS-DOC 2.3.2: note references and numbers are special characters.
        let bytes = passive_special_source(&bytes);
        super::super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1024 * 1024)
            .map(|result| result.document)
    }

    fn runs(element: &BodyElement) -> Vec<String> {
        let BodyElement::Paragraph(paragraph) = element else {
            panic!("paragraph expected");
        };
        paragraph
            .runs
            .iter()
            .map(|run| match run {
                DocRun::Text(text) => match &text.note_ref {
                    Some(note) => format!(
                        "{}:{}:{}:{}",
                        note.kind,
                        note.id,
                        text.text,
                        text.vert_align.as_deref().unwrap_or_default()
                    ),
                    None => text.text.clone(),
                },
                DocRun::Field(field) => format!("field:{}", field.field_type),
                _ => "other".into(),
            })
            .collect()
    }

    fn fixture<'a>(notes: &'a [&'a str], references: &'a [(u32, bool)]) -> NotesFixture<'a> {
        NotesFixture {
            notes,
            references,
            separators: SEPARATORS,
            dop: Some(default_note_dop()),
        }
    }

    #[test]
    fn footnotes_project_onto_the_docx_note_model() {
        let notes = ["\u{2} one\r", "\u{2} two\rsecond\r"];
        let references = [(1, true), (3, true)];
        let document = document("A\u{2}B\u{2}\r", &fixture(&notes, &references)).unwrap();
        assert_eq!(
            runs(&document.body[0]),
            ["A", "footnote:1:1:super", "B", "footnote:2:2:super"]
        );
        assert!(document.endnotes.is_empty());
        let ids: Vec<_> = document
            .footnotes
            .iter()
            .map(|note| note.id.as_str())
            .collect();
        assert_eq!(ids, ["1", "2"]);
        assert_eq!(
            runs(&document.footnotes[0].content[0]),
            ["footnote:::super", " one"]
        );
        assert_eq!(document.footnotes[1].content.len(), 2);
        assert_eq!(runs(&document.footnotes[1].content[1]), ["second"]);
    }

    #[test]
    fn note_fields_use_the_footnote_field_table() {
        let notes = ["\u{2} p\u{13}PAGE\u{14}9\u{15}\r"];
        let references = [(1, true)];
        let document = document("A\u{2}\r", &fixture(&notes, &references)).unwrap();
        assert_eq!(
            runs(&document.footnotes[0].content[0]),
            ["footnote:::super", " p", "field:page"]
        );
    }

    #[test]
    fn unrepresentable_note_properties_are_rejected() {
        let notes = ["\u{2} one\r"];
        let references = [(1, true)];
        let reject = |fixture: NotesFixture<'_>, text: &str| document(text, &fixture).unwrap_err();
        // A literal custom mark would still be numbered by the renderer.
        let custom = [(1, false)];
        assert!(reject(fixture(&notes, &custom), "A*\r").contains("custom note reference"));
        for (offset, value, message) in [
            // Chicago numbering, per-section and per-page restarts, a zero
            // start and beneath-text placement have no shared rendering.
            (492usize, 9u8, "footnote numbering"),
            (2, 1 | (1 << 2), "footnote numbering"),
            (2, 2 | (1 << 2), "footnote numbering"),
            (2, 0, "footnote numbering"),
            (0, 2 << 5, "footnote numbering"),
        ] {
            let mut properties = fixture(&notes, &references);
            let dop = properties.dop.as_mut().unwrap();
            if offset == 2 {
                dop[2..4].copy_from_slice(&u16::from(value).to_le_bytes());
            } else {
                dop[offset] = value;
            }
            assert!(reject(properties, "A\u{2}\r").contains(message), "{offset}");
        }
        let mut missing = fixture(&notes, &references);
        missing.dop = None;
        assert!(reject(missing, "A\u{2}\r").contains("note properties"));
        for separators in [
            ["", "\u{4}\r\r", "", "", "", ""],
            ["\u{3}\r\r\r", "\u{4}\r\r", "", "", "", ""],
            ["\u{3}\r\r", "\u{3}\r\r", "", "", "", ""],
            ["\u{3}\r\r", "\u{4}\r\r", "x\r\r", "", "", ""],
        ] {
            let mut properties = fixture(&notes, &references);
            properties.separators = separators;
            assert!(reject(properties, "A\u{2}\r").contains("separators"));
        }
    }

    #[test]
    fn dop_note_format_and_start_become_document_note_numbering() {
        let notes = ["\u{2} one\r"];
        let references = [(1, true)];
        let mut properties = fixture(&notes, &references);
        let dop = properties.dop.as_mut().unwrap();
        dop[492] = 2; // nfcFtnRef msonfcLCRoman
        dop[2..4].copy_from_slice(&(3u16 << 2).to_le_bytes()); // nFtn 3
        let settings = document("A\u{2}\r", &properties)
            .unwrap()
            .note_layout_settings
            .unwrap();
        assert_eq!(
            settings.footnote_number_format.as_deref(),
            Some("lowerRoman")
        );
        assert_eq!(settings.footnote_number_start, Some(3));
        assert_eq!(settings.endnote_number_format, None);
    }

    #[test]
    fn every_reference_must_be_displayed_exactly_once() {
        let notes = ["\u{2} one\r"];
        let references = [(5, true)];
        // The only U+0002 lies inside a hidden field instruction.
        let text = "A\u{13}IF \u{2}\u{15}\r";
        let error = document(text, &fixture(&notes, &references)).unwrap_err();
        assert!(error.contains("not displayed"), "{error}");
    }
}
