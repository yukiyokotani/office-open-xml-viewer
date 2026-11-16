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
//! footnotes out at the page bottom and endnotes at the end of the document.
//! Native reserved stories retain explicit rule kinds and effective PAPX/CHPX
//! for its metric consumer. Numeric short-rule span remains library policy,
//! without a claim of Word-exact width or placement.
//!
//! A custom-mark reference (aFtnIdx/aEdnIdx == 0) projects as the DOCX
//! `customMarkFollows` reference host followed by its literal mark text; it
//! consumes no automatic number (ECMA-376 17.11.14 / the shared renderer).
//!
//! Word numbers notes per section (format, first number, restart). When the
//! document-wide model does not produce those numbers, `note_numbers` computes
//! them and every mark of that kind is projected as its text, in the
//! custom-mark shape. A hidden reference character displays no mark in the
//! text; its note is still placed and numbered.
//!
//! Footnote and endnote placement (page bottom, beneath text, section end,
//! document end) is carried by the model's position fields, as the DOCX
//! parser carries `w:footnotePr/w:pos` and `w:endnotePr/w:pos`.
//! A separator story with content of its own has no place in the model's
//! reserved-story shapes: that kind keeps the layout's default separator.

use super::{
    fields::{Linked, StoryFields},
    story, ModelBudget,
};
use crate::doc::{
    formatting, header_fields, headers, notes, numbering, pictures, sections, settings,
    tokenize_with_fields, unsupported, Fields, Paragraph, Token,
};
use docx_model::{
    DocxNote, NoteLayoutSettingsWire, ReservedNoteSeparatorStoriesWire, ReservedNoteSeparatorsWire,
};

mod separators;

mod note_numbers;

/// Last effective nFib whose note properties live in the DOP (MS-DOC 2.7.2).
const DOP_NOTE_PROPERTIES_MAX_NFIB: u16 = 0x00d9;

/// Read the note properties, number the notes and retain the numbering,
/// placement and separator facts. ECMA-376 17.11.17/.18/.20 governs automatic
/// numbering; rule geometry and continuation scope follow the library policies
/// documented beside their gates and layout consumer.
// The parameters are independent note, header, section and acquisition facts
// plus the two mutable ownership budgets (formatting and model), passed
// individually like `story::project`'s projections.
#[allow(clippy::too_many_arguments)]
pub(super) fn validate(
    stories: &[Option<notes::Notes<'_>>],
    references: &mut notes::References,
    effective_nfib: u16,
    dop: Option<&settings::Properties>,
    sections: &[sections::Section],
    headers: Option<&headers::Headers<'_>>,
    formatting: &mut formatting::Formatting<'_>,
    budget: &mut ModelBudget,
) -> Result<Option<NoteLayoutSettingsWire>, String> {
    let story = |kind| {
        stories
            .iter()
            .flatten()
            .find(|notes| notes.kind == kind && !notes.entries.is_empty())
    };
    let (footnotes, endnotes) = (story(notes::Kind::Footnote), story(notes::Kind::Endnote));
    if footnotes.is_none() && endnotes.is_none() {
        return Ok(None);
    }
    let dop = dop
        .map(|dop| dop.notes)
        .ok_or_else(|| unsupported("Word notes require document note properties"))?;
    let dop_scoped = effective_nfib <= DOP_NOTE_PROPERTIES_MAX_NFIB;
    if dop_scoped && dop.formats.is_none() {
        return Err(unsupported("Word note number formats are missing"));
    }
    let mut number = |story: Option<&notes::Notes<'_>>| -> Result<_, String> {
        let Some(story) = story else {
            return Ok(None);
        };
        let hidden: Vec<bool> = story
            .entries
            .iter()
            .map(|entry| {
                references
                    .get(entry.reference_cp)
                    .is_some_and(notes::Reference::hidden)
            })
            .collect();
        let numbering = note_numbers::assign(story, &hidden, effective_nfib, &dop, sections)?;
        if let Some(marks) = &numbering.marks {
            budget.charge(marks.iter().flatten().map(String::capacity).sum())?;
            references.set_marks(story.kind, marks);
        }
        Ok(Some(numbering))
    };
    let footnote_numbering = number(footnotes)?;
    let endnote_numbering = number(endnotes)?;
    let mut settings = NoteLayoutSettingsWire {
        footnote_number_format: footnote_numbering
            .as_ref()
            .map(|numbering| numbering.format.to_string()),
        footnote_number_start: footnote_numbering.as_ref().map(|numbering| numbering.start),
        endnote_number_format: endnote_numbering
            .as_ref()
            .map(|numbering| numbering.format.to_string()),
        endnote_number_start: endnote_numbering.as_ref().map(|numbering| numbering.start),
        ..NoteLayoutSettingsWire::default()
    };
    // MS-DOC 2.7.2 DopBase.fpc and 2.6.4 sprmSFpc: 0 end of section, 1 bottom
    // of page, 2 beneath text. The model states the position once for the
    // document (ECMA-376 17.11.21); the first section speaks for a later
    // document.
    if footnotes.is_some() {
        let position = if dop_scoped {
            dop.footnote_position
        } else {
            sections
                .first()
                .and_then(|section| section.note_properties().footnote_position)
                .unwrap_or(1)
        };
        settings.footnote_position = match position {
            1 => None,
            0 => Some("sectEnd"),
            2 => Some("beneathText"),
            _ => crate::rejection::spec_default("invalid Word footnote position", None),
        }
        .map(str::to_string);
        budget.charge(
            settings
                .footnote_position
                .as_ref()
                .map_or(0, String::capacity),
        )?;
    }
    // DopBase.epc is not scoped by nFib: 0 end of section, 3 end of document
    // (ECMA-376 17.11.22).
    if endnotes.is_some() {
        settings.endnote_position = match dop.endnote_position {
            3 => None,
            0 => Some("sectEnd"),
            _ => crate::rejection::spec_default("invalid Word endnote position", None),
        }
        .map(str::to_string);
        budget.charge(
            settings
                .endnote_position
                .as_ref()
                .map_or(0, String::capacity),
        )?;
    }
    let (footnotes, endnotes) = (footnotes.is_some(), endnotes.is_some());
    // MS-DOC 2.3.3/2.8.22 assigns story roles; 2.6.1 sprmCFSpec assigns
    // U+0003/U+0004 short/full semantics independently of those roles. Retain
    // actual PAPX/CHPX for the shared reserved-separator metric consumer.
    // Footnote short/short is supported within its closed metric-only contract.
    // Numeric short-rule span remains the existing library layout policy;
    // admission does not claim Office-exact rule placement or width.
    let mut reserved = ReservedNoteSeparatorsWire::default();
    for (present, base) in [(footnotes, 0), (endnotes, 3)] {
        if !present {
            continue;
        }
        let Some(headers) = headers else {
            // No separator stories: the layout's default separator.
            crate::rejection::approximated(
                "Word notes without separator stories use the default separator",
                (),
            );
            continue;
        };
        let ordinary = separators::acquire(headers, base, formatting, budget)?;
        let continuation = separators::acquire(headers, base + 1, formatting, budget)?;
        // A paragraph-only notice owns real PAPX/CHPX too. Acquiring it does
        // not equate its formatting with DOCX absence or a default template.
        let notice = separators::acquire(headers, base + 2, formatting, budget)?;
        // Only the footnote paginator consumes continuation roles. Preserve
        // the existing endnote envelope until endnote continuation is wired.
        // Apply metric eligibility only to this newly admitted class: existing
        // short/full admission keeps its previous producer behavior.
        let short_continuation = base == 0
            && continuation.standard_rule(separators::RuleKind::Short)
            && [&ordinary, &continuation, &notice]
                .iter()
                .all(|story| story.metric_only());
        // MS-DOC 2.3.3 explicitly permits zero-length reserved stories. The
        // shared native-story consumer already represents empty stories with
        // no ink or flow. Preserve the equal CPs and explicit None mark;
        // never substitute a default rule or discard custom content.
        // A paragraph-only story (one empty paragraph, no rule) reserves
        // that paragraph without ink in the shared reserved-story consumer.
        // Endnote continuation keeps its existing envelope (see above).
        if !(ordinary.standard_rule(separators::RuleKind::Short)
            || ordinary.empty_story()
            || ordinary.paragraph_only())
            || !(continuation.standard_rule(separators::RuleKind::Full)
                || short_continuation
                || continuation.empty_story()
                || (base == 0 && continuation.paragraph_only()))
            || !matches!(headers.separator_text(base + 2), "" | "\r\r")
            || !(ordinary.wire_shape() && continuation.wire_shape() && notice.wire_shape())
        {
            // A separator with content of its own (text, a rule in another
            // role) has no reserved-story shape in the model. The kind keeps
            // the separator the layout draws without one, as for a DOCX
            // whose separator story it does not read.
            crate::rejection::approximated(
                "custom Word note separator drawn as the default separator",
                (),
            );
            continue;
        }
        // A rule retains its explicit kind and effective character metrics.
        // Continuation remains caller opt-in; no document/default option changes.
        let slot = if base == 0 {
            settings.footnote_separator = ordinary.mark();
            settings.footnote_continuation_separator = continuation.mark();
            &mut reserved.footnote
        } else {
            settings.endnote_separator = ordinary.mark();
            &mut reserved.endnote
        };
        *slot = Some(ReservedNoteSeparatorStoriesWire {
            separator: ordinary.into_wire()?,
            continuation_separator: continuation.into_wire()?,
            continuation_notice: notice.into_wire()?,
        });
    }
    // One new retained container per document. Moved payloads keep their
    // cumulative acquisition charges; nothing in it is a second copy.
    budget.charge(std::mem::size_of::<ReservedNoteSeparatorsWire>())?;
    settings.reserved_separators = Some(Box::new(reserved));
    Ok(Some(settings))
}

/// Turn main-story note characters into references. Every automatic U+0002
/// must name a reference; a custom mark (MS-DOC 2.8.19/2.8.17 aFtnIdx/aEdnIdx
/// == 0) is an ordinary displayed character at the reference CP, so a
/// zero-width custom reference host is inserted before it and the character
/// stays text — the shape of DOCX `footnoteReference customMarkFollows`
/// followed by its mark (ECMA-376 17.11.14). `check_all_references_placed`
/// then requires every reference, automatic or custom, to be displayed.
pub(super) fn restore_references(
    references: &notes::References,
    paragraphs: &mut [Paragraph],
    placed: &mut usize,
) -> Result<(), String> {
    for paragraph in paragraphs {
        let source = std::mem::take(&mut paragraph.tokens);
        let mut tokens = Vec::with_capacity(source.len());
        for (mut token, cp) in source {
            let (inner, link) = match &mut token {
                Token::Linked(linked) => (&mut linked.token, Some(&linked.link)),
                token => (token, None),
            };
            match inner {
                Token::NoteMarker => {
                    let reference = references.get(cp).ok_or_else(|| {
                        unsupported("Word automatic note character without a note reference")
                    })?;
                    *inner = Token::NoteReference(reference.clone());
                    *placed += 1;
                }
                Token::Text(text) => {
                    let end = cp + text.encode_utf16().count();
                    let mut marks = references.custom_in(cp..end).peekable();
                    if marks.peek().is_some() {
                        let link = link.cloned();
                        let wrap = |token| match &link {
                            Some(link) => Token::Linked(Box::new(Linked {
                                token,
                                link: link.clone(),
                            })),
                            None => token,
                        };
                        let text = std::mem::take(text);
                        let (mut start_byte, mut start_cp) = (0, cp);
                        let mut chars = text.char_indices();
                        let mut at = cp;
                        for (mark_cp, reference) in marks {
                            // Advance to the mark character's byte offset.
                            while at < mark_cp {
                                let (_, ch) = chars.next().ok_or_else(|| {
                                    unsupported("Word custom note mark outside its text")
                                })?;
                                at += ch.len_utf16();
                            }
                            let byte = chars.offset();
                            if at != mark_cp {
                                return Err(unsupported(
                                    "Word note reference splits Unicode character",
                                ));
                            }
                            if byte > start_byte {
                                tokens.push((
                                    wrap(Token::Text(text[start_byte..byte].to_string())),
                                    start_cp,
                                ));
                            }
                            tokens.push((wrap(Token::NoteReference(reference.clone())), mark_cp));
                            *placed += 1;
                            (start_byte, start_cp) = (byte, mark_cp);
                        }
                        tokens.push((wrap(Token::Text(text[start_byte..].to_string())), start_cp));
                        continue;
                    }
                }
                _ => {}
            }
            tokens.push((token, cp));
        }
        paragraph.tokens = tokens;
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
#[allow(clippy::too_many_arguments)]
pub(super) fn project(
    notes: &notes::Notes<'_>,
    references: &notes::References,
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
                    // Each U+0002 shows the enclosing note's number. A
                    // custom-mark note consumes no number, so the shared
                    // renderer suppresses it there, as for a DOCX `footnoteRef`
                    // inside a `customMarkFollows` note. A note the reader
                    // numbers itself shows that number as text.
                    *token = match references
                        .get(entry.reference_cp)
                        .and_then(notes::Reference::mark)
                    {
                        Some(mark) => Token::NoteMark(mark.to_string()),
                        None => Token::NoteNumber(notes.kind),
                    };
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
            false, // Page-story breaks do not advance body pages.
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
    fn separator_acquisition_resolves_special_and_formatting_before_role_policy() {
        let mut fixture = fixture(&["\u{2} note\r"], &[(1, true)]);
        fixture.separators[1] = "\u{3}\r\r";
        fixture.separators[2] = "\r\r";
        let bytes = source_with_stories(
            "A\u{2}\r",
            &[(3, 2, 12_240, 15_840, 1, 720)],
            None,
            Some(24),
            None,
            Some(&[None; 6]),
            Some(&fixture),
        );
        let bytes = passive_special_source(&bytes);
        crate::doc::with_acquired_doc(&CompoundFile::open(&bytes).unwrap(), |mut facts| {
            let headers = facts.headers.as_ref().unwrap();
            let mut budget = super::super::ModelBudget::new(1024 * 1024);
            for slot in [0, 1] {
                let acquired =
                    super::separators::acquire(headers, slot, &mut facts.formatting, &mut budget)?;
                let rule = acquired.rule.as_ref().unwrap();
                assert_eq!(rule.kind, Some(super::separators::RuleKind::Short));
                assert!(rule.character.picture.passive_special());
                assert_eq!(rule.character.run.as_ref().unwrap().font_size, 12.0);
                let paragraph = acquired.paragraph.as_ref().unwrap();
                assert_eq!(paragraph.paragraph.default_font_size, Some(12.0));
                assert_eq!(
                    paragraph
                        .mark
                        .as_ref()
                        .unwrap()
                        .run
                        .as_ref()
                        .unwrap()
                        .font_size,
                    12.0
                );
                assert_eq!(acquired.source.content, "\u{3}\r");
                assert_eq!(acquired.source.cp, (slot * 3)..(slot * 3 + 3));
            }
            let notice =
                super::separators::acquire(headers, 2, &mut facts.formatting, &mut budget)?;
            assert!(notice.rule.is_none());
            assert_eq!(notice.source.content_cp, 6..7);
            assert_eq!(
                notice
                    .paragraph
                    .as_ref()
                    .unwrap()
                    .mark
                    .as_ref()
                    .unwrap()
                    .run
                    .as_ref()
                    .unwrap()
                    .font_size,
                12.0
            );
            // Admission preserves the acquired short continuation and its
            // formatting rather than substituting the continuation-role default.
            let settings = super::validate(
                &facts.note_stories,
                &mut facts.note_references,
                facts.effective_nfib,
                facts.document_settings.as_ref(),
                &facts.sections,
                Some(headers),
                &mut facts.formatting,
                &mut budget,
            )
            .unwrap()
            .unwrap();
            assert_eq!(
                settings.footnote_continuation_separator,
                Some(docx_model::NoteSeparatorMark::Short)
            );
            let stories = settings.reserved_separators.unwrap().footnote.unwrap();
            let continuation = stories.continuation_separator.rule.unwrap();
            assert_eq!(
                continuation.mark,
                Some(docx_model::NoteSeparatorMark::Short)
            );
            assert_eq!(continuation.control.run.unwrap().font_size, 12.0);
            let notice = stories.continuation_notice.paragraph.unwrap();
            assert_eq!(notice.content_mark.unwrap().run.unwrap().font_size, 12.0);
            Ok(())
        })
        .unwrap();
    }

    /// Production native bytes whose Normal style is 14pt, so every retained
    /// CHPX/PAPX font size comes from the real style cascade, not a default.
    fn styled_source(notice: &'static str) -> Vec<u8> {
        let mut fixture = fixture(&["\u{2} note\r"], &[(1, true)]);
        fixture.separators[2] = notice;
        let bytes = source_with_stories(
            "A\u{2}\r",
            &[(3, 2, 12_240, 15_840, 1, 720)],
            None,
            Some(28),
            None,
            Some(&[None; 6]),
            Some(&fixture),
        );
        passive_special_source(&bytes)
    }

    #[test]
    fn admitted_separator_formatting_moves_into_the_shared_wire() {
        use docx_model::NoteSeparatorMark as Mark;
        for notice in ["", "\r\r"] {
            let bytes = styled_source(notice);
            let cfb = CompoundFile::open(&bytes).unwrap();
            let result = super::super::super::direct_model(&cfb, 1024 * 1024).unwrap();
            let settings = result.document.note_layout_settings.unwrap();
            // Retention leaves the scalar marks that select the drawn rule.
            assert_eq!(settings.footnote_separator, Some(Mark::Short));
            assert_eq!(settings.footnote_continuation_separator, Some(Mark::Full));
            let reserved = settings.reserved_separators.unwrap();
            assert!(reserved.endnote.is_none());
            let stories = reserved.footnote.unwrap();
            for (story, mark) in [
                (&stories.separator, Mark::Short),
                (&stories.continuation_separator, Mark::Full),
            ] {
                let rule = story.rule.as_ref().unwrap();
                assert_eq!(rule.mark, Some(mark));
                assert_eq!(rule.control.run.as_ref().unwrap().font_size, 14.0);
                let paragraph = story.paragraph.as_ref().unwrap();
                assert_eq!(paragraph.paragraph.default_font_size, Some(14.0));
                let content_mark = paragraph.content_mark.as_ref().unwrap();
                assert_eq!(content_mark.run.as_ref().unwrap().font_size, 14.0);
            }
            let notice_story = &stories.continuation_notice;
            assert!(notice_story.rule.is_none());
            if notice.is_empty() {
                assert!(notice_story.paragraph.is_none());
            } else {
                let paragraph = notice_story.paragraph.as_ref().unwrap();
                let content_mark = paragraph.content_mark.as_ref().unwrap();
                assert_eq!(content_mark.run.as_ref().unwrap().font_size, 14.0);
            }
        }
        // The content paragraph mark is its own character, not the control
        // and not the guard: acquisition addresses them separately.
        let bytes = styled_source("\r\r");
        let cfb = CompoundFile::open(&bytes).unwrap();
        crate::doc::with_acquired_doc(&cfb, |mut facts| {
            let headers = facts.headers.as_ref().unwrap();
            let mut budget = super::super::ModelBudget::new(1024 * 1024);
            for (slot, start) in [(0, 0), (1, 3)] {
                let story =
                    super::separators::acquire(headers, slot, &mut facts.formatting, &mut budget)?;
                assert_eq!(story.source.content_cp, start..start + 2);
                assert_eq!(story.source.guard_cp, Some(start + 2));
                assert_eq!(story.rule.as_ref().unwrap().source.cp, start);
                assert_eq!(story.paragraph.as_ref().unwrap().source.cp, start + 1);
            }
            let notice =
                super::separators::acquire(headers, 2, &mut facts.formatting, &mut budget)?;
            assert_eq!(notice.source.guard_cp, Some(7));
            assert_eq!(notice.paragraph.as_ref().unwrap().source.cp, 6);
            Ok(())
        })
        .unwrap();
    }

    #[test]
    fn retained_separator_container_is_charged_once_without_payload_copies() {
        let bytes = styled_source("\r\r");
        let cfb = CompoundFile::open(&bytes).unwrap();
        let limit = 1024 * 1024;
        let acquired = crate::doc::with_acquired_doc(&cfb, |mut facts| {
            let headers = facts.headers.as_ref().unwrap();
            let mut budget = super::super::ModelBudget::new(limit);
            for slot in 0..3 {
                super::separators::acquire(headers, slot, &mut facts.formatting, &mut budget)?;
            }
            Ok(limit - budget.remaining_bytes)
        })
        .unwrap();
        let validate = |limit| {
            crate::doc::with_acquired_doc(&cfb, |mut facts| {
                let mut budget = super::super::ModelBudget::new(limit);
                let result = super::validate(
                    &facts.note_stories,
                    &mut facts.note_references,
                    facts.effective_nfib,
                    facts.document_settings.as_ref(),
                    &facts.sections,
                    facts.headers.as_ref(),
                    &mut facts.formatting,
                    &mut budget,
                );
                Ok((result, limit - budget.remaining_bytes))
            })
            .unwrap()
        };
        let (result, used) = validate(limit);
        assert!(result.unwrap().unwrap().reserved_separators.is_some());
        // Moved paragraph/run payloads were charged by acquisition; only the
        // one boxed container is new.
        let container = std::mem::size_of::<docx_model::ReservedNoteSeparatorsWire>();
        assert_eq!(used, acquired + container);
        assert_eq!(validate(used - 1).0.unwrap_err(), "OUTPUT_TOO_LARGE");
    }

    #[test]
    fn raw_separator_controls_without_special_meaning_are_custom_separators() {
        let fixture = fixture(&["\u{2} note\r"], &[(1, true)]);
        let bytes = source_with_stories(
            "A\u{2}\r",
            &[(3, 2, 12_240, 15_840, 1, 720)],
            None,
            None,
            None,
            Some(&[None; 6]),
            Some(&fixture),
        );
        let bytes = passive_special_source(&bytes);
        let cfb = CompoundFile::open(&bytes).unwrap();
        let mut word = cfb.stream("WordDocument").unwrap();
        let table = cfb.stream("0Table").unwrap();
        let bte = u32::from_le_bytes(word[0xfa..0xfe].try_into().unwrap()) as usize;
        let page_number = u32::from_le_bytes(table[bte + 8..bte + 12].try_into().unwrap()) as usize;
        let header_cp = u32::from_le_bytes(word[0x4c..0x50].try_into().unwrap())
            + u32::from_le_bytes(word[0x50..0x54].try_into().unwrap());
        let page = &mut word[page_number * 512..(page_number + 1) * 512];
        let start_fc = u32::from_le_bytes(page[..4].try_into().unwrap());
        let end_fc = u32::from_le_bytes(page[4..8].try_into().unwrap());
        page[4..8].copy_from_slice(&(start_fc + header_cp * 2).to_le_bytes());
        page[8..12].copy_from_slice(&end_fc.to_le_bytes());
        page[12..14].copy_from_slice(&[32, 34]);
        page[64..68].copy_from_slice(&[3, 0x55, 0x08, 1]);
        page[68..72].copy_from_slice(&[3, 0x55, 0x08, 0]);
        page[511] = 2;
        let bytes = crate::cfb::test_support::build_scoped_cfb(&[
            ("WordDocument", word),
            ("0Table", table),
        ]);
        // Without the special-character property U+0003 is ordinary text:
        // the story is a custom separator, and the kind keeps the layout's
        // default one.
        crate::rejection::take_best_effort();
        let result =
            crate::doc::direct_model(&CompoundFile::open(&bytes).unwrap(), 1024 * 1024).unwrap();
        assert_default_separator(&result.document);
    }

    /// The footnote kind carries no reserved separator story: the layout
    /// draws its default separator.
    fn assert_default_separator(document: &Document) {
        let settings = document.note_layout_settings.as_ref().unwrap();
        assert_eq!(settings.footnote_separator, None);
        assert_eq!(settings.footnote_continuation_separator, None);
        assert!(settings
            .reserved_separators
            .as_ref()
            .unwrap()
            .footnote
            .is_none());
        assert!(crate::rejection::take_best_effort().contains(&(
            crate::rejection::BestEffort::Approximated,
            "custom Word note separator drawn as the default separator"
        )));
    }

    #[test]
    fn footnotes_project_onto_the_docx_note_model() {
        let notes = ["\u{2} one\r", "\u{2} two\rsecond\r"];
        let references = [(1, true), (3, true)];
        let document = document("A\u{2}B\u{2}\r", &fixture(&notes, &references)).unwrap();
        let settings = document.note_layout_settings.as_ref().unwrap();
        assert_eq!(
            settings.footnote_separator,
            Some(docx_model::NoteSeparatorMark::Short)
        );
        assert_eq!(
            settings.footnote_continuation_separator,
            Some(docx_model::NoteSeparatorMark::Full)
        );
        assert_eq!(settings.endnote_separator, None);
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
    fn custom_note_marks_project_as_custom_mark_follows_references() {
        // MS-DOC 2.8.19 aFtnIdx == 0: the literal mark at the reference CP
        // stays ordinary text after a custom reference host, as DOCX
        // `footnoteReference customMarkFollows` + its mark, and the note's own
        // text keeps its literal mark. Automatic notes still number.
        let notes = ["\u{2020} custom\r", "\u{2} auto\r"];
        let references = [(1, false), (4, true)];
        let projected = document("A\u{2020}BC\u{2}\r", &fixture(&notes, &references)).unwrap();
        let BodyElement::Paragraph(paragraph) = &projected.body[0] else {
            panic!("paragraph expected");
        };
        let notes_refs: Vec<_> = paragraph
            .runs
            .iter()
            .filter_map(|run| match run {
                DocRun::Text(text) => text.note_ref.as_ref().map(|note| note.custom_mark_follows),
                _ => None,
            })
            .collect();
        assert_eq!(notes_refs, [true, false]);
        assert_eq!(
            runs(&projected.body[0]),
            ["A", "footnote:1:1:", "\u{2020}BC", "footnote:2:2:super"]
        );
        assert_eq!(
            runs(&projected.footnotes[0].content[0]),
            ["\u{2020} custom"]
        );
        assert_eq!(
            runs(&projected.footnotes[1].content[0]),
            ["footnote:::super", " auto"]
        );
        // A custom mark at the start or end of a text token, and adjacent
        // custom marks, each keep their own host and literal character.
        let notes = ["* one\r", "* two\r"];
        let projected =
            document("*\u{2020}\r", &fixture(&notes, &[(0, false), (1, false)])).unwrap();
        assert_eq!(
            runs(&projected.body[0]),
            ["footnote:1:1:", "*", "footnote:2:2:", "\u{2020}"]
        );
        // A custom mark must still be a displayed character.
        let reject = |fixture: NotesFixture<'_>, text: &str| document(text, &fixture).unwrap_err();
        assert!(
            reject(fixture(&notes[..1], &[(1, false)]), "A\tB\r").contains("custom note marker")
        );
    }

    #[test]
    fn paragraph_only_separator_stories_reserve_their_paragraph_without_a_rule() {
        // MS-DOC 2.3.3: a separator story holding only an empty paragraph.
        let notes = ["\u{2} one\r"];
        let references = [(1, true)];
        let mut properties = fixture(&notes, &references);
        properties.separators[0] = "\r\r";
        properties.separators[1] = "\r\r";
        let projected = document("A\u{2}\r", &properties).unwrap();
        let settings = serde_json::to_value(projected.note_layout_settings.unwrap()).unwrap();
        for role in ["separator", "continuationSeparator"] {
            let story = &settings["reservedSeparators"]["footnote"][role];
            assert!(story.get("rule").is_none());
            assert!(story["paragraph"]["contentMark"]["run"].is_object());
        }
        // A rule-free story with authored text is a custom separator.
        let mut properties = fixture(&notes, &references);
        properties.separators[0] = "x\r\r";
        crate::rejection::take_best_effort();
        assert_default_separator(&document("A\u{2}\r", &properties).unwrap());
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
    fn empty_reserved_separator_stories_preserve_explicit_absence() {
        let notes = ["\u{2} one\r"];
        let references = [(1, true)];
        let mut properties = fixture(&notes, &references);
        properties.separators = [""; 6];
        let document = document("A\u{2}\r", &properties).unwrap();
        assert_eq!(document.footnotes.len(), 1);
        let settings = serde_json::to_value(document.note_layout_settings.unwrap()).unwrap();
        assert_eq!(settings["footnoteSeparator"], "none");
        assert_eq!(settings["footnoteContinuationSeparator"], "none");
        for role in ["separator", "continuationSeparator", "continuationNotice"] {
            assert_eq!(
                settings["reservedSeparators"]["footnote"][role],
                serde_json::json!({})
            );
        }
    }

    fn with_dop(
        notes: &[&str],
        references: &[(u32, bool)],
        text: &str,
        edit: impl FnOnce(&mut Vec<u8>),
    ) -> Document {
        let mut properties = fixture(notes, references);
        edit(properties.dop.as_mut().unwrap());
        document(text, &properties).unwrap()
    }

    fn best_effort() -> Vec<&'static str> {
        crate::rejection::take_best_effort()
            .into_iter()
            .map(|(_, site)| site)
            .collect()
    }

    #[test]
    fn symbol_numbering_is_projected_as_the_text_of_each_mark() {
        // nfcFtnRef 9 (chicago): the shared renderer has no such numbering,
        // so the reader numbers the notes.
        let notes = ["\u{2} one\r", "\u{2} two\r"];
        let references = [(1, true), (3, true)];
        best_effort();
        let document = with_dop(&notes, &references, "A\u{2}B\u{2}\r", |dop| dop[492] = 9);
        assert_eq!(
            runs(&document.body[0]),
            ["A", "footnote:1:1:", "*", "B", "footnote:2:2:", "\u{2020}"]
        );
        // The reference host is the custom-mark shape; the mark follows as
        // superscript text, in the body and inside the note.
        let BodyElement::Paragraph(paragraph) = &document.body[0] else {
            panic!()
        };
        let (DocRun::Text(host), DocRun::Text(mark)) = (&paragraph.runs[1], &paragraph.runs[2])
        else {
            panic!()
        };
        assert!(host.note_ref.as_ref().unwrap().custom_mark_follows);
        assert_eq!(mark.vert_align.as_deref(), Some("super"));
        assert_eq!(
            runs(&document.footnotes[1].content[0]),
            ["\u{2020}", " two"]
        );
        assert!(best_effort().is_empty());
    }

    #[test]
    fn per_page_restart_counts_continuously() {
        let notes = ["\u{2} one\r", "\u{2} two\r"];
        let references = [(1, true), (3, true)];
        best_effort();
        // rncFtn 2 (restart each page), nFtn 1.
        let document = with_dop(&notes, &references, "A\u{2}B\u{2}\r", |dop| {
            dop[2..4].copy_from_slice(&(2u16 | (1 << 2)).to_le_bytes());
        });
        assert_eq!(
            runs(&document.body[0]),
            ["A", "footnote:1:1:super", "B", "footnote:2:2:super"]
        );
        assert_eq!(
            best_effort(),
            ["Word per-page note numbering restart counts continuously"]
        );
    }

    #[test]
    fn a_zero_start_counts_from_one_and_positions_reach_the_model() {
        let notes = ["\u{2} one\r"];
        let references = [(1, true)];
        best_effort();
        let document = with_dop(&notes, &references, "A\u{2}\r", |dop| {
            dop[2..4].copy_from_slice(&0u16.to_le_bytes());
        });
        let settings = document.note_layout_settings.unwrap();
        assert_eq!(settings.footnote_number_start, Some(1));
        assert_eq!(settings.footnote_position, None);
        assert_eq!(best_effort(), ["zero Word note number start"]);
        // DopBase.fpc 2 (beneath text) and 0 (end of section).
        for (fpc, position) in [(2u8, "beneathText"), (0, "sectEnd")] {
            let document = with_dop(&notes, &references, "A\u{2}\r", |dop| {
                dop[0] = (dop[0] & !0x60) | (fpc << 5);
            });
            let settings = document.note_layout_settings.unwrap();
            assert_eq!(settings.footnote_position.as_deref(), Some(position));
        }
        assert!(best_effort().is_empty());
    }

    #[test]
    fn a_restart_in_one_section_is_the_document_wide_numbering() {
        // rncFtn 1 (restart each section) with every note in one section.
        let notes = ["\u{2} one\r", "\u{2} two\r"];
        let references = [(1, true), (3, true)];
        let document = with_dop(&notes, &references, "A\u{2}B\u{2}\r", |dop| {
            dop[2..4].copy_from_slice(&(1u16 | (4 << 2)).to_le_bytes());
        });
        assert_eq!(
            runs(&document.body[0]),
            ["A", "footnote:1:1:super", "B", "footnote:2:2:super"]
        );
        assert_eq!(
            document.note_layout_settings.unwrap().footnote_number_start,
            Some(4)
        );
    }

    #[test]
    fn a_restart_in_each_section_is_numbered_by_the_reader() {
        // rncFtn 1 (restart each section), nFtn 1, two sections with a note
        // each: both notes are number 1, which the document-wide model
        // cannot produce.
        let notes = ["\u{2} one\r", "\u{2} two\r"];
        let references = [(1, true), (5, true)];
        let mut properties = fixture(&notes, &references);
        properties.dop.as_mut().unwrap()[2..4].copy_from_slice(&(1u16 | (1 << 2)).to_le_bytes());
        let text = "A\u{2}\r\u{c}B\u{2}\r";
        let sections = [
            (4, 2, 12_240, 15_840, 1, 720),
            (text.encode_utf16().count(), 2, 12_240, 15_840, 1, 720),
        ];
        let bytes = source_with_stories(
            text,
            &sections,
            None,
            None,
            None,
            Some(&[None; 12]),
            Some(&properties),
        );
        let bytes = passive_special_source(&bytes);
        let document =
            super::super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1024 * 1024)
                .unwrap()
                .document;
        let marks: Vec<Vec<String>> = document
            .body
            .iter()
            .filter(|element| matches!(element, BodyElement::Paragraph(_)))
            .map(runs)
            .collect();
        assert_eq!(marks[0], ["A", "footnote:1:1:", "1"]);
        assert_eq!(marks.last().unwrap(), &["B", "footnote:2:2:", "1"]);
        for note in &document.footnotes {
            assert_eq!(runs(&note.content[0])[0], "1");
        }
    }

    #[test]
    fn missing_properties_and_custom_separators() {
        let notes = ["\u{2} one\r"];
        let references = [(1, true)];
        let mut missing = fixture(&notes, &references);
        missing.dop = None;
        assert!(document("A\u{2}\r", &missing)
            .unwrap_err()
            .contains("note properties"));
        for separators in [
            ["\u{3}\r\r\r", "\u{4}\r\r", "", "", "", ""],
            ["\u{4}\r\r", "\u{3}\r\r", "", "", "", ""],
            ["\u{3}\r\r", "\u{4}\r\r", "x\r\r", "", "", ""],
        ] {
            let mut properties = fixture(&notes, &references);
            properties.separators = separators;
            best_effort();
            assert_default_separator(&document("A\u{2}\r", &properties).unwrap());
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
