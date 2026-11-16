//! Cell comments of the direct XLS model.
//!
//! BIFF8 keeps a comment as a 2.4.179 Note record (NoteSh 2.5.186: the cell,
//! fShow, the object identifier and the author) and a drawing object (Obj
//! ot 25) with that identifier, whose TxO holds the comment text. They become
//! the XLSX model's classic notes, as the XLSX parser reads ECMA-376 18.7.3
//! `comment`: the cell reference, the author and the text. The renderer
//! marks each such cell with the comment indicator.
//!
//! The comment box itself is not a drawing of the model: a note is shown as
//! its indicator, also when Excel keeps the box visible on the sheet (fShow).
//! Excel prints neither by default.

use super::{u16_at, Record};

/// Resource policy for the Note records of one sheet.
const MAX_NOTES: usize = 65_536;

struct Note {
    row: u16,
    column: u16,
    shown: bool,
    object: u16,
    author: Option<String>,
}

/// The Note records of one worksheet substream, in stream order.
#[derive(Default)]
pub(super) struct Notes(Vec<Note>);

fn truncated() -> String {
    crate::rejection::corrupt("truncated BIFF Note record")
}

impl Notes {
    pub(super) fn is_empty(&self) -> bool {
        self.0.is_empty()
    }

    /// NoteSh: row, col, flags (fShow is bit 1), idObj, stAuthor (an
    /// XLUnicodeString) and an optional unused byte.
    pub(super) fn push(&mut self, data: &[u8]) -> Result<(), String> {
        if self.0.len() >= MAX_NOTES {
            return Err(crate::rejection::resource("too many BIFF cell comments"));
        }
        let (row, column) = (u16_at(data, 0)?, u16_at(data, 2)?);
        if column > 255 {
            return Err(crate::rejection::corrupt(
                "BIFF Note column exceeds the BIFF8 limit",
            ));
        }
        let shown = u16_at(data, 4)? & 0x0002 != 0;
        let object = u16_at(data, 6)?;
        let count = usize::from(u16_at(data, 8)?);
        let author = match *data.get(10).ok_or_else(truncated)? & 0x01 {
            0 => data
                .get(11..11 + count)
                .ok_or_else(truncated)?
                .iter()
                .map(|&byte| char::from(byte))
                .collect::<String>(),
            _ => {
                let bytes = data.get(11..11 + count * 2).ok_or_else(truncated)?;
                let units: Vec<u16> = bytes
                    .as_chunks::<2>()
                    .0
                    .iter()
                    .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
                    .collect();
                String::from_utf16_lossy(&units)
            }
        };
        self.0.push(Note {
            row,
            column,
            shown,
            object,
            author: (!author.is_empty()).then_some(author),
        });
        Ok(())
    }
}

/// A1 reference of a zero-based cell.
fn reference(row: u16, column: u16) -> String {
    let mut index = u32::from(column) + 1;
    let mut letters = Vec::new();
    while index > 0 {
        let rem = (index - 1) % 26;
        letters.push(char::from(b'A' + rem as u8));
        index = (index - 1) / 26;
    }
    let column: String = letters.iter().rev().collect();
    format!("{column}{}", u32::from(row) + 1)
}

/// The notes of one sheet as model comments. `text` gives the stream offset
/// of the TxO of a comment object by its identifier; a note without such an
/// object is not a comment. `budget` bounds the retained comment bytes of the
/// workbook.
pub(super) fn project(
    notes: &Notes,
    records: &[Record<'_>],
    text: impl Fn(u16) -> Option<usize>,
    budget: &mut usize,
) -> Result<Vec<xlsx_model::XlsxComment>, String> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(notes.0.len())
        .map_err(|_| crate::rejection::resource("XLS comment allocation failed"))?;
    // Each comment object's TxO is read once, however many Notes name it.
    let mut texts: std::collections::BTreeMap<u16, String> = std::collections::BTreeMap::new();
    for note in &notes.0 {
        // A Note without its comment object is no comment: Excel shows
        // none, and such a record alone never changed the sheet.
        let Some(offset) = text(note.object) else {
            continue;
        };
        if note.shown {
            // The box Excel keeps on the sheet is not drawn.
            crate::rejection::skipped("BIFF cell comments are not projected");
        }
        let text = match texts.get(&note.object) {
            Some(text) => text.clone(),
            None => {
                let text = super::shapes::txo_text(records, offset)?;
                texts.insert(note.object, text.clone());
                text
            }
        };
        let cell_ref = reference(note.row, note.column);
        let bytes = std::mem::size_of::<xlsx_model::XlsxComment>()
            + text.len()
            + cell_ref.len()
            + note.author.as_ref().map_or(0, String::len);
        *budget = budget
            .checked_sub(bytes)
            .ok_or_else(|| crate::rejection::resource("XLS comment retention budget exceeded"))?;
        output.push(xlsx_model::XlsxComment {
            kind: xlsx_model::XlsxCommentKind::Note,
            cell_ref,
            id: None,
            person_id: None,
            author: note.author.clone(),
            date: None,
            root_text: None,
            text,
            resolved: None,
            replies: Vec::new(),
        });
    }
    Ok(output)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::rejection::{take_best_effort, BestEffort};

    fn note(row: u16, column: u16, flags: u16, object: u16, author: &str) -> Vec<u8> {
        let mut data = Vec::new();
        for value in [row, column, flags, object, author.len() as u16] {
            data.extend(value.to_le_bytes());
        }
        data.push(0);
        data.extend(author.bytes());
        data.push(0);
        data
    }

    #[test]
    fn notes_are_classic_comments_with_their_text_and_author() {
        take_best_effort();
        // TxO with cchText 5 and its text and formatting Continue records.
        let mut txo = vec![0u8; 18];
        txo[10] = 5;
        txo[12] = 16;
        let text = [&[0u8][..], b"hello"].concat();
        // One run from character 0, then the last run at cchText.
        let mut runs = [0u8; 16];
        runs[8] = 5;
        let records = [
            Record {
                kind: 0x01b6,
                offset: 40,
                data: &txo,
            },
            Record {
                kind: 0x003c,
                offset: 62,
                data: &text,
            },
            Record {
                kind: 0x003c,
                offset: 72,
                data: &runs,
            },
        ];
        let mut notes = Notes::default();
        notes.push(&note(2, 1, 0, 7, "Ann")).unwrap();
        notes.push(&note(0, 27, 0x0002, 7, "")).unwrap();
        // A Note that names no comment object.
        notes.push(&note(5, 5, 0, 9, "Bob")).unwrap();
        let mut budget = 1 << 20;
        let comments = project(
            &notes,
            &records,
            |object| (object == 7).then_some(40),
            &mut budget,
        )
        .unwrap();
        assert_eq!(
            serde_json::to_value(&comments).unwrap(),
            serde_json::json!([
                {"kind": "note", "cellRef": "B3", "author": "Ann", "text": "hello"},
                {"kind": "note", "cellRef": "AB1", "text": "hello"}
            ])
        );
        // The second note's box is shown on the sheet (fShow) and not drawn.
        assert_eq!(
            take_best_effort(),
            [(BestEffort::Skipped, "BIFF cell comments are not projected")]
        );
        // The TxO's formatting runs are part of it: without their Continue
        // record the comment cannot be read.
        assert!(project(
            &notes,
            &records[..2],
            |object| (object == 7).then_some(40),
            &mut budget
        )
        .unwrap_err()
        .contains("formatting runs"));
        // ... and they must end at cchText.
        let unterminated = [0u8; 16];
        let mut broken = records;
        broken[2].data = &unterminated;
        assert!(project(
            &notes,
            &broken,
            |object| (object == 7).then_some(40),
            &mut budget
        )
        .unwrap_err()
        .contains("invalid BIFF TxO formatting runs"));
        take_best_effort();
        // Retained bytes are bounded.
        let mut small = 8;
        assert!(project(&notes, &records, |_| Some(40), &mut small)
            .unwrap_err()
            .starts_with("LEGACY_RESOURCE_LIMIT:"));
        // A truncated author and a column past IV are malformed.
        let mut short = note(0, 0, 0, 1, "Ann");
        short.truncate(12);
        assert!(Notes::default().push(&short).is_err());
        assert!(Notes::default().push(&note(0, 256, 0, 1, "")).is_err());
    }
}
