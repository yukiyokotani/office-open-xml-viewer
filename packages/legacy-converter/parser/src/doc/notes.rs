//! MS-DOC 2.3.2/5, 2.8.16/17/19/20: note reference and text PLCs.
use super::{formatting, read_story_range, u16_at, u32_at, unsupported, Story};
use std::collections::BTreeMap;
use std::ops::Range;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum Kind {
    Footnote,
    Endnote,
}

impl Kind {
    pub fn tag(self) -> &'static str {
        match self {
            Self::Footnote => "footnote",
            Self::Endnote => "endnote",
        }
    }
    fn offsets(self) -> (usize, usize) {
        match self {
            Self::Footnote => (0x50, 0xaa),
            Self::Endnote => (0x60, 0x20a),
        }
    }
}

pub(super) struct Entry {
    pub reference_cp: usize,
    pub automatic: bool,
    pub cp: usize,
    pub text: Range<usize>,
}

pub(super) struct Notes<'a> {
    pub kind: Kind,
    pub story: Story<'a>,
    pub entries: Vec<Entry>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(super) struct Reference {
    kind: Kind,
    id: usize,
    custom: bool,
}
impl Reference {
    pub(in crate::doc) fn kind(&self) -> Kind {
        self.kind
    }

    /// One-based position in its note document, used as the note id.
    pub(in crate::doc) fn id(&self) -> usize {
        self.id
    }

    /// True for a literal custom mark (FRD.nAuto == 0).
    pub(in crate::doc) fn custom(&self) -> bool {
        self.custom
    }
}

pub(super) struct References {
    by_cp: BTreeMap<usize, Reference>,
}
impl References {
    pub(in crate::doc) fn get(&self, cp: usize) -> Option<&Reference> {
        self.by_cp.get(&cp)
    }

    pub(in crate::doc) fn iter(&self) -> impl Iterator<Item = &Reference> {
        self.by_cp.values()
    }

    pub fn read(
        notes: &[Option<Notes<'_>>],
        main: &Story<'_>,
        formatting: &mut formatting::Formatting<'_>,
    ) -> Result<Self, String> {
        let mut by_cp = BTreeMap::new();
        for notes in notes.iter().flatten() {
            for (i, entry) in notes.entries.iter().enumerate() {
                if by_cp
                    .insert(
                        entry.reference_cp,
                        Reference {
                            kind: notes.kind,
                            id: i + 1,
                            custom: !entry.automatic,
                        },
                    )
                    .is_some()
                {
                    return Err(unsupported("overlapping Word note references"));
                }
            }
        }
        if by_cp.is_empty() {
            return Ok(Self { by_cp });
        }
        let mut references = by_cp.iter().peekable();
        let mut cp = 0;
        for ch in main.text.chars() {
            if let Some((&position, reference)) = references.peek().copied() {
                if position < cp {
                    return Err(unsupported("Word note reference splits Unicode character"));
                }
                if position == cp {
                    if reference.custom {
                        if ch < ' ' {
                            return Err(unsupported("unsupported Word custom note marker"));
                        }
                    } else {
                        if ch != '\u{2}' {
                            return Err(unsupported("invalid Word automatic note marker"));
                        }
                        let (_, fc, piece) = main
                            .position(cp)
                            .ok_or_else(|| unsupported("Word note reference outside story"))?;
                        let style = formatting.paragraph_style(fc)?;
                        if !formatting
                            .passive_special_character(style, fc, piece.prm, &main.prcs)?
                        {
                            return Err(unsupported(
                                "Word note reference lacks special-character property",
                            ));
                        }
                    }
                    references.next();
                }
            }
            cp += ch.len_utf16();
        }
        if references.next().is_some() {
            return Err(unsupported("Word note reference outside text"));
        }
        Ok(Self { by_cp })
    }
}

pub(super) fn read_all<'a>(
    word: &[u8],
    table: &[u8],
    clx: &'a [u8],
) -> Result<Vec<Option<Notes<'a>>>, String> {
    // A shared decoded-character budget prevents two independent note stories
    // from each reserving the maximum before either is projected.
    (u32_at(word, 0x50)? as usize)
        .checked_add(u32_at(word, 0x60)? as usize)
        .filter(|n| *n <= super::MAX_MAIN_STORY_UNITS)
        .ok_or_else(|| unsupported("Word aggregate note character budget exceeded"))?;
    [Kind::Footnote, Kind::Endnote]
        .into_iter()
        .map(|kind| read(word, table, clx, kind))
        .collect()
}

pub(super) fn read<'a>(
    word: &[u8],
    table: &[u8],
    clx: &'a [u8],
    kind: Kind,
) -> Result<Option<Notes<'a>>, String> {
    let (ccp, fc) = kind.offsets();
    let length = u32_at(word, ccp)? as usize;
    let reference_size = u32_at(word, fc + 4)? as usize;
    let text_size = u32_at(word, fc + 12)? as usize;
    if length == 0 && text_size == 0 && matches!(reference_size, 0 | 4) {
        // A PLC with no data still has one final, undefined CP. Both a missing
        // reference PLC and that empty PLC are valid for an absent note story.
        if reference_size == 4 {
            let offset = u32_at(word, fc)? as usize;
            table
                .get(offset..)
                .and_then(|b| b.get(..4))
                .ok_or_else(|| unsupported("Word empty note table outside its stream"))?;
        }
        return Ok(None);
    }
    if length == 0
        || (reference_size != 0 && (reference_size < 4 || !(reference_size - 4).is_multiple_of(6)))
    {
        return Err(unsupported("inconsistent Word note document"));
    }
    let count = reference_size.saturating_sub(4) / 6;
    // Resource policy: retained entries, not a limitation of the file format.
    if count > 65_536 || text_size != (count + 2) * 4 {
        return Err(unsupported("invalid or excessive Word note table length"));
    }
    let read_plc = |offset| -> Result<&[u8], String> {
        let start = u32_at(word, offset)? as usize;
        let size = u32_at(word, offset + 4)? as usize;
        table
            .get(start..)
            .and_then(|b| b.get(..size))
            .ok_or_else(|| unsupported("Word note table outside its stream"))
    };
    let references = if reference_size == 0 {
        &[][..]
    } else {
        read_plc(fc)?
    };
    let boundaries = read_plc(fc + 8)?;
    let main = u32_at(word, 0x4c)? as usize;
    let mut entries = Vec::with_capacity(count);
    let mut previous = None;
    for i in 0..count {
        let cp = u32_at(references, i * 4)? as usize;
        if cp >= main || previous.is_some_and(|p| cp <= p) {
            return Err(unsupported("invalid Word note reference position"));
        }
        previous = Some(cp);
        entries.push(Entry {
            reference_cp: cp,
            automatic: u16_at(references, (count + 1) * 4 + i * 2)? != 0,
            cp: 0,
            text: 0..0,
        });
    }
    let mut cps = Vec::with_capacity(count + 1);
    for i in 0..=count {
        let cp = u32_at(boundaries, i * 4)? as usize;
        if cp >= length || cps.last().is_some_and(|p| cp <= *p) {
            return Err(unsupported("invalid Word note text position"));
        }
        cps.push(cp);
    }
    if cps[count] != length - 1 {
        return Err(unsupported("Word note text coverage mismatch"));
    }
    // MS-DOC 2.3: main, footnotes, headers, comments, then endnotes.
    let mut start = main;
    if kind == Kind::Endnote {
        for offset in [0x50, 0x54, 0x5c] {
            start = start
                .checked_add(u32_at(word, offset)? as usize)
                .filter(|n| *n <= i32::MAX as usize)
                .ok_or_else(|| unsupported("Word note story range overflow"))?;
        }
    }
    let story = read_story_range(word, clx, start, length)?;
    if story
        .text
        .bytes()
        .filter(|b| *b < 32)
        .take(super::MAX_STORY_CONTROLS + 1)
        .count()
        > super::MAX_STORY_CONTROLS
    {
        return Err(unsupported("Word note structure budget exceeded"));
    }
    // One forward UTF-16-to-UTF-8 pass over each aggregate note story.
    // Never decode/copy the CLX once per note or scan from the start per entry.
    let mut chars = story.text.char_indices();
    let (mut cp, mut byte) = (0, 0);
    let mut offsets = Vec::with_capacity(cps.len());
    for target in &cps {
        while cp < *target {
            let (at, ch) = chars
                .next()
                .ok_or_else(|| unsupported("truncated Word note text"))?;
            cp += ch.len_utf16();
            byte = at + ch.len_utf8();
        }
        if cp != *target {
            return Err(unsupported("Word note boundary splits Unicode character"));
        }
        offsets.push(byte);
    }
    for (i, entry) in entries.iter_mut().enumerate() {
        entry.cp = cps[i];
        entry.text = offsets[i]..offsets[i + 1];
        if !story.text[entry.text.clone()].ends_with('\r') {
            return Err(unsupported("Word note lacks final paragraph mark"));
        }
    }
    Ok(Some(Notes {
        kind,
        story,
        entries,
    }))
}

#[cfg(test)]
mod tests {
    use super::*;
    fn fixture(kind: Kind) -> (Vec<u8>, Vec<u8>, Vec<u8>) {
        let mut word = vec![0; 2048];
        let main = "A\u{2}B*\r";
        let notes = "\u{2}😀 note\r* custom\r\r";
        let preceding = if kind == Kind::Endnote {
            "F\rH\rC\r"
        } else {
            ""
        };
        let text = format!("{main}{preceding}{notes}");
        word[0x4c..0x50].copy_from_slice(&(main.len() as u32).to_le_bytes());
        if kind == Kind::Endnote {
            for offset in [0x50, 0x54, 0x5c] {
                word[offset..offset + 4].copy_from_slice(&2u32.to_le_bytes());
            }
        }
        let (ccp, fc) = kind.offsets();
        word[ccp..ccp + 4].copy_from_slice(&(notes.encode_utf16().count() as u32).to_le_bytes());
        word[fc + 4..fc + 8].copy_from_slice(&16u32.to_le_bytes());
        word[fc + 8..fc + 12].copy_from_slice(&16u32.to_le_bytes());
        word[fc + 12..fc + 16].copy_from_slice(&16u32.to_le_bytes());
        let mut table = Vec::new();
        for cp in [1u32, 3, u32::MAX] {
            table.extend(cp.to_le_bytes());
        }
        table.extend(1u16.to_le_bytes());
        table.extend(0u16.to_le_bytes());
        for cp in [0u32, 9, 18, u32::MAX] {
            table.extend(cp.to_le_bytes());
        }
        for (i, unit) in text.encode_utf16().enumerate() {
            word[1024 + i * 2..1026 + i * 2].copy_from_slice(&unit.to_le_bytes());
        }
        let mut clx = vec![2];
        clx.extend(16u32.to_le_bytes());
        clx.extend(0u32.to_le_bytes());
        clx.extend((text.encode_utf16().count() as u32).to_le_bytes());
        clx.extend([0, 0]);
        clx.extend(1024u32.to_le_bytes());
        clx.extend([0, 0]);
        (word, table, clx)
    }
    #[test]
    fn maps_both_note_stories_after_preceding_documents_and_ignores_final_cps() {
        for kind in [Kind::Footnote, Kind::Endnote] {
            let (word, table, clx) = fixture(kind);
            let notes = read(&word, &table, &clx, kind).unwrap().unwrap();
            assert_eq!(notes.entries.len(), 2);
            assert!(notes.entries[0].automatic);
            assert!(!notes.entries[1].automatic);
            assert_eq!(notes.entries[1].reference_cp, 3);
            assert_eq!(
                &notes.story.text[notes.entries[0].text.clone()],
                "\u{2}😀 note\r"
            );
            assert_eq!(
                &notes.story.text[notes.entries[1].text.clone()],
                "* custom\r"
            );
            assert_eq!(
                notes.story.position(0).unwrap().1,
                if kind == Kind::Endnote { 1046 } else { 1034 }
            );
        }
    }
    #[test]
    fn rejects_inconsistent_plcs_ranges_unicode_boundaries_and_missing_paragraphs() {
        for (offset, value) in [(0, 5u32), (4, 1), (16, 19), (20, 2), (24, 17)] {
            let (word, mut table, clx) = fixture(Kind::Footnote);
            table[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
            assert!(read(&word, &table, &clx, Kind::Footnote).is_err());
        }
        for (offset, value) in [
            (0x50, 0u32),
            (0xae, 15),
            (0xb6, 12),
            (0xaa, u32::MAX),
            (0x50, 64 * 1024 * 1024 + 1),
        ] {
            let (mut word, table, clx) = fixture(Kind::Footnote);
            word[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
            assert!(read(&word, &table, &clx, Kind::Footnote).is_err());
        }
        let (mut word, table, clx) = fixture(Kind::Footnote);
        word[1050..1052].copy_from_slice(&u16::from(b'x').to_le_bytes());
        assert!(read(&word, &table, &clx, Kind::Footnote).is_err());
    }

    #[test]
    fn absent_notes_ignore_undefined_offsets() {
        let mut word = vec![0; 1024];
        word[0xaa..0xae].fill(255);
        word[0xb2..0xb6].fill(255);
        assert!(read(&word, &[], &[], Kind::Footnote).unwrap().is_none());
        word[0xaa..0xae].fill(0);
        word[0xae..0xb2].copy_from_slice(&4u32.to_le_bytes());
        assert!(read(&word, &[255; 4], &[], Kind::Footnote)
            .unwrap()
            .is_none());
        assert!(read(&word, &[255; 3], &[], Kind::Footnote).is_err());
    }

    #[test]
    fn rejects_aggregate_note_budget_before_decoding_or_allocating_tables() {
        let mut word = vec![0; 1024];
        for offset in [0x50, 0x60] {
            word[offset..offset + 4].copy_from_slice(&(33u32 * 1024 * 1024).to_le_bytes());
        }
        let error = read_all(&word, &[], &[]).err().unwrap();
        assert!(error.contains("aggregate note character budget"));
        let (mut word, table, clx) = fixture(Kind::Footnote);
        word[0xae..0xb2].copy_from_slice(&(65_537u32 * 6 + 4).to_le_bytes());
        assert!(read(&word, &table, &clx, Kind::Footnote)
            .err()
            .unwrap()
            .contains("excessive Word note table"));
    }
}
