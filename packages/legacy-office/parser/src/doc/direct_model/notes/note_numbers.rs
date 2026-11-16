//! Automatic note numbers (MS-DOC 2.7.2 DopBase rncFtn/nFtn/rncEdn/nEdn,
//! 2.7.4 Dop97 nfcFtnRef/nfcEdnRef, 2.6.4 sprmSRncFtn/sprmSNFtn/
//! sprmSNfcFtnRef and their endnote forms; ECMA-376 Part 1 17.11.17-.20).
//!
//! The DOCX model numbers the notes of a kind in reference order with one
//! number format and one start for the whole document. Word numbers them per
//! section: each section states its format, its first number and whether the
//! count continues from the previous section or restarts.
//!
//! This module computes the number Word shows for every automatic note. When
//! the result is what the document-wide model produces, the model keeps
//! numbering the notes. Otherwise every mark of that kind is projected as
//! the text of its number: the shape of a DOCX custom mark (17.11.14
//! `customMarkFollows`), which the renderer draws without numbering it.

use crate::doc::{notes, number_format, sections, settings};

/// Last effective nFib whose note properties live in the DOP (MS-DOC 2.7.2).
const DOP_NOTE_PROPERTIES_MAX_NFIB: u16 = 0x00d9;

/// MS-DOC 2.9.? Rnc.
const RESTART_EACH_SECTION: u8 = 1;
const RESTART_EACH_PAGE: u8 = 2;

/// MS-OSHARED 2.2.1.3 MSONFC values without a displayed number.
const NFC_CHICAGO: u16 = 9;

/// Numbering of one note kind.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct Numbering {
    /// ECMA-376 17.18.59 ST_NumberFormat of the document-wide model.
    pub(super) format: &'static str,
    /// First automatic number of the document-wide model.
    pub(super) start: i64,
    /// The displayed mark of every note, by note index, when the
    /// document-wide model does not produce it; `None` entries are custom
    /// marks.
    pub(super) marks: Option<Vec<Option<String>>>,
}

/// The format, first number and restart rule of one section.
#[derive(Clone, Copy)]
struct Scope {
    format: u16,
    start: u16,
    restart: u8,
}

/// `hidden[i]`: whether the reference mark of note `i` is hidden text.
pub(super) fn assign(
    story: &notes::Notes<'_>,
    hidden: &[bool],
    effective_nfib: u16,
    dop: &settings::NoteProperties,
    sections: &[sections::Section],
) -> Numbering {
    let footnote = story.kind == notes::Kind::Footnote;
    // MS-DOC 2.6.4: without the section property a footnote number is
    // Arabic and an endnote number lower-case Roman, counted from 1.
    let default_format = if footnote { 0 } else { 2 };
    let document = (effective_nfib <= DOP_NOTE_PROPERTIES_MAX_NFIB).then(|| {
        // MS-DOC 2.7.2/2.7.4: these documents keep note numbering in the
        // DOP; a section note property does not apply to them.
        if sections
            .iter()
            .any(|section| section.note_properties() != Default::default())
        {
            crate::rejection::ignored_value(
                "Word section note properties in a DOP-scoped document",
            );
        }
        let formats = dop.formats.unwrap_or((0, 2));
        if footnote {
            Scope {
                format: formats.0,
                start: dop.footnote_start,
                restart: dop.footnote_restart,
            }
        } else {
            Scope {
                format: formats.1,
                start: dop.endnote_start,
                restart: dop.endnote_restart,
            }
        }
    });
    let scope_of = |section: usize| -> Scope {
        if let Some(document) = document {
            return document;
        }
        let properties = sections
            .get(section)
            .map(sections::Section::note_properties)
            .unwrap_or_default();
        if footnote {
            Scope {
                format: properties.footnote_format.unwrap_or(default_format),
                start: properties.footnote_offset.unwrap_or(1),
                restart: properties.footnote_restart.unwrap_or(0),
            }
        } else {
            Scope {
                format: properties.endnote_format.unwrap_or(default_format),
                start: properties.endnote_offset.unwrap_or(1),
                restart: properties.endnote_restart.unwrap_or(0),
            }
        }
    };
    let section_of = |cp: usize| sections.partition_point(|section| section.end <= cp);

    // (format, number) of every automatic note, in reference order.
    let mut numbers: Vec<Option<(u16, i64)>> = Vec::with_capacity(story.entries.len());
    let mut previous: Option<(usize, i64)> = None;
    let mut per_page = false;
    for entry in &story.entries {
        if !entry.automatic {
            // A custom mark consumes no number (ECMA-376 17.11.14).
            numbers.push(None);
            continue;
        }
        let section = section_of(entry.reference_cp);
        let scope = scope_of(section);
        per_page |= scope.restart == RESTART_EACH_PAGE;
        let number = match previous {
            Some((last_section, last))
                if !(scope.restart == RESTART_EACH_SECTION && last_section != section) =>
            {
                last + 1
            }
            _ => first_number(scope.start),
        };
        previous = Some((section, number));
        numbers.push(Some((scope.format, number)));
    }
    if per_page {
        // Pages exist only after layout, and the DOCX model has no per-page
        // restart: the count continues across pages.
        crate::rejection::approximated(
            "Word per-page note numbering restart counts continuously",
            (),
        );
    }

    let mut automatic = numbers.iter().flatten();
    let (format, start) = automatic.next().copied().unwrap_or_else(|| {
        let scope = scope_of(0);
        (scope.format, first_number(scope.start))
    });
    let uniform = automatic
        .zip(start + 1..)
        .all(|((other, number), expected)| *other == format && *number == expected);
    let name = model_format(format, default_format);
    let any_hidden = story
        .entries
        .iter()
        .zip(hidden)
        .any(|(entry, hidden)| entry.automatic && *hidden);
    let marks = (!uniform || name.is_none() || any_hidden).then(|| {
        numbers
            .iter()
            .map(|number| number.map(|(format, number)| mark(format, number)))
            .collect()
    });
    Numbering {
        format: name.unwrap_or("decimal"),
        start,
        marks,
    }
}

/// ECMA-376 17.11.20: the first number is at least 1. MS-DOC 2.7.2 allows
/// nFtn/nEdn 1..=16383; a zero is outside that range and counts from 1.
fn first_number(start: u16) -> i64 {
    if start == 0 {
        crate::rejection::spec_default("zero Word note number start", 1)
    } else {
        i64::from(start)
    }
}

/// The ST_NumberFormat the document-wide model carries for `format`, or
/// `None` when the shared renderer has no numbering of that name (symbols).
fn model_format(format: u16, default: u16) -> Option<&'static str> {
    let name = |value: u16| {
        u8::try_from(value)
            .ok()
            .and_then(|value| number_format::name(value).ok())
    };
    match name(format) {
        // No number at all, or a list bullet: not a note numbering.
        Some("chicago" | "none" | "bullet") => None,
        Some(name) => Some(name),
        None => {
            crate::rejection::ignored_value("invalid Word note number format");
            name(default)
        }
    }
}

/// The text of note number `number` in MSONFC `format`.
fn mark(format: u16, number: i64) -> String {
    let number = number.max(1) as u64;
    match format {
        0 => number.to_string(),
        1 => roman(number),
        2 => roman(number).to_ascii_lowercase(),
        3 => letters(number, b'A'),
        4 => letters(number, b'a'),
        NFC_CHICAGO => {
            // ECMA-376 17.18.59 chicago: *, †, ‡, §, then each doubled,
            // tripled and so on.
            const SYMBOLS: [char; 4] = ['*', '\u{2020}', '\u{2021}', '\u{a7}'];
            let index = (number - 1) as usize;
            std::iter::repeat_n(SYMBOLS[index % 4], (index / 4 + 1).min(64)).collect()
        }
        _ => {
            // The reader formats only the numberings above itself.
            crate::rejection::approximated("Word note number format shown as Arabic", ());
            number.to_string()
        }
    }
}

fn roman(mut number: u64) -> String {
    const PARTS: [(u64, &str); 13] = [
        (1000, "M"),
        (900, "CM"),
        (500, "D"),
        (400, "CD"),
        (100, "C"),
        (90, "XC"),
        (50, "L"),
        (40, "XL"),
        (10, "X"),
        (9, "IX"),
        (5, "V"),
        (4, "IV"),
        (1, "I"),
    ];
    let mut text = String::new();
    for (value, part) in PARTS {
        while number >= value {
            text.push_str(part);
            number -= value;
        }
    }
    text
}

/// ECMA-376 17.18.59 upperLetter/lowerLetter: a..z, aa..zz, aaa..
fn letters(number: u64, first: u8) -> String {
    let index = (number - 1) as usize;
    std::iter::repeat_n(
        char::from(first + (index % 26) as u8),
        (index / 26 + 1).min(64),
    )
    .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn marks_follow_their_number_format() {
        assert_eq!(mark(0, 12), "12");
        assert_eq!(mark(1, 1994), "MCMXCIV");
        assert_eq!(mark(2, 4), "iv");
        assert_eq!(mark(3, 27), "AA");
        assert_eq!(mark(4, 3), "c");
        let chicago: Vec<String> = (1..=6).map(|number| mark(NFC_CHICAGO, number)).collect();
        assert_eq!(
            chicago,
            [
                "*",
                "\u{2020}",
                "\u{2021}",
                "\u{a7}",
                "**",
                "\u{2020}\u{2020}"
            ]
        );
        crate::rejection::take_best_effort();
        assert_eq!(mark(11, 3), "3");
        assert_eq!(
            crate::rejection::take_best_effort(),
            [(
                crate::rejection::BestEffort::Approximated,
                "Word note number format shown as Arabic"
            )]
        );
    }

    fn story(references: &[(usize, bool)]) -> notes::Notes<'static> {
        notes::Notes {
            kind: notes::Kind::Footnote,
            story: crate::doc::Story {
                text: String::new(),
                pieces: Vec::new(),
                prcs: Vec::new(),
            },
            entries: references
                .iter()
                .map(|(reference_cp, automatic)| notes::Entry {
                    reference_cp: *reference_cp,
                    automatic: *automatic,
                    cp: 0,
                    text: 0..0,
                })
                .collect(),
        }
    }

    fn dop(restart: u8, start: u16, format: u16) -> settings::NoteProperties {
        settings::NoteProperties {
            footnote_position: 1,
            footnote_restart: restart,
            footnote_start: start,
            endnote_restart: 0,
            endnote_start: 1,
            endnote_position: 3,
            formats: Some((format, 2)),
        }
    }

    #[test]
    fn continuous_numbering_stays_with_the_document_wide_model() {
        let notes = story(&[(1, true), (5, false), (9, true)]);
        let sections = [
            sections::Section::for_test(4, 2),
            sections::Section::for_test(20, 2),
        ];
        let numbering = assign(&notes, &[false; 3], 0x00c1, &dop(0, 3, 2), &sections);
        assert_eq!(
            numbering,
            Numbering {
                format: "lowerRoman",
                start: 3,
                marks: None
            }
        );
    }

    #[test]
    fn section_restarts_and_hidden_references_are_numbered_by_the_reader() {
        let notes = story(&[(1, true), (2, true), (9, false), (10, true)]);
        let sections = [
            sections::Section::for_test(4, 2),
            sections::Section::for_test(20, 2),
        ];
        // Restart each section from 2, upper-case letters.
        let numbering = assign(&notes, &[false; 4], 0x00c1, &dop(1, 2, 3), &sections);
        assert_eq!(numbering.format, "upperLetter");
        assert_eq!(
            numbering.marks,
            Some(vec![
                Some("B".into()),
                Some("C".into()),
                None,
                Some("B".into())
            ])
        );
        // A hidden reference: the note keeps its number, as text.
        let numbering = assign(
            &notes,
            &[false, true, false, false],
            0x00c1,
            &dop(0, 1, 0),
            &sections,
        );
        assert_eq!(
            numbering.marks,
            Some(vec![
                Some("1".into()),
                Some("2".into()),
                None,
                Some("3".into())
            ])
        );
    }

    #[test]
    fn later_documents_number_by_section_properties() {
        let notes = story(&[(1, true), (9, true)]);
        let mut second = sections::Section::for_test(20, 2);
        second.set_note_properties_for_test(sections::NoteProperties {
            footnote_format: Some(2),
            ..Default::default()
        });
        let sections = [sections::Section::for_test(4, 2), second];
        crate::rejection::take_best_effort();
        // nFib above 0x00D9: the DOP numbering (chicago here) does not apply.
        let numbering = assign(&notes, &[false; 2], 0x0101, &dop(0, 7, 9), &sections);
        assert_eq!(numbering.format, "decimal");
        assert_eq!(
            numbering.marks,
            Some(vec![Some("1".into()), Some("ii".into())])
        );
        assert!(crate::rejection::take_best_effort().is_empty());
        // The same sections in a DOP-scoped document: the DOP governs.
        let numbering = assign(&notes, &[false; 2], 0x00c1, &dop(0, 7, 0), &sections);
        assert_eq!((numbering.start, numbering.marks), (7, None));
        assert_eq!(
            crate::rejection::take_best_effort(),
            [(
                crate::rejection::BestEffort::IgnoredValue,
                "Word section note properties in a DOP-scoped document"
            )]
        );
    }

    #[test]
    fn model_formats_exclude_symbol_numberings() {
        assert_eq!(model_format(0, 0), Some("decimal"));
        assert_eq!(model_format(2, 0), Some("lowerRoman"));
        assert_eq!(model_format(14, 0), Some("decimalFullWidth"));
        assert_eq!(model_format(NFC_CHICAGO, 0), None);
        crate::rejection::take_best_effort();
        assert_eq!(model_format(200, 2), Some("lowerRoman"));
        assert_eq!(crate::rejection::take_best_effort().len(), 1);
    }
}
