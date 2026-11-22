//! Word textbox stories (MS-DOC 2.3.6-2.3.7, 2.8.23, 2.8.30-2.8.32,
//! 2.9.106-2.9.107 and 2.9.312).
//!
//! The textbox document follows the endnote document; the header textbox
//! document follows it. `PlcftxbxTxt`/`PlcfHdrtxbxTxt` give one FTXBXS per
//! textbox (its `lid` is the owning shape's spid) and the text range that
//! ends in the textbox's final paragraph mark. The last FTXBXS, and every
//! entry flagged fReusable, is a spare structure whose range is ignored.
//! `PlcfTxbxBkd`/`PlcfTxbxHdrBkd` break an FTXBXS range into one range per
//! linked textbox; a range split across several boxes is a linked chain and
//! is not reconstructed here.

use super::Part;
use crate::doc::direct_model::fields::StoryFields;
use crate::doc::{header_fields, read_story_range, u16_at, u32_at, unsupported, Story};
use std::ops::Range;

// Retained table-entry resource policy, independent of MS-DOC PLC syntax.
const MAX_TEXTBOXES: usize = 100_000;

pub(in crate::doc) struct Textboxes<'a> {
    pub story: Story<'a>,
    /// One entry per FTXBXS; `None` for reusable spare structures.
    boxes: Vec<Option<Textbox>>,
    /// MS-DOC 2.8.25 Plcfld of this textbox document (PlcfFldTxbx at FIB
    /// 0x262, PlcffldHdrTxbx at 0x272), validated like any other story.
    pub fields: StoryFields,
}

#[derive(Debug)]
struct Textbox {
    lid: u32,
    cp: usize,
    bytes: Range<usize>,
}

impl<'a> Textboxes<'a> {
    /// Read the textbox story of one drawing part. `Ok(None)` means the part
    /// has no textbox document; inconsistent tables fail closed.
    pub fn read(
        word: &[u8],
        table: &[u8],
        clx: &'a [u8],
        part: Part,
    ) -> Result<Option<Self>, String> {
        let ccp = |offset: usize| -> Result<usize, String> {
            let value = u32_at(word, offset)?;
            if value > i32::MAX as u32 {
                return Err(unsupported("negative Word story length"));
            }
            Ok(value as usize)
        };
        // FibRgLw97: ccpText, ccpFtn, ccpHdd, (reserved3), ccpAtn, ccpEdn,
        // ccpTxbx, ccpHdrTxbx. MS-DOC 2.3.5-2.3.7 order the textbox documents
        // after the endnote document; reserved3 is ignored.
        let mut start = 0usize;
        for offset in [0x4c, 0x50, 0x54, 0x5c, 0x60] {
            start = start
                .checked_add(ccp(offset)?)
                .ok_or_else(|| unsupported("Word textbox story range overflow"))?;
        }
        let main_length = ccp(0x64)?;
        let (length, plc, breaks, field_table) = match part {
            Part::Main => (main_length, 0x25a, 0x2f2, 0x262),
            Part::Header => {
                start = start
                    .checked_add(main_length)
                    .ok_or_else(|| unsupported("Word textbox story range overflow"))?;
                (ccp(0x68)?, 0x26a, 0x2fa, 0x272)
            }
        };
        let plc = plc_bytes(word, table, plc)?;
        let breaks = plc_bytes(word, table, breaks)?;
        // MS-DOC 2.5.6: both tables are present exactly when the story is.
        match (length, plc, breaks) {
            (0, None, None) => Ok(None),
            (1.., Some(plc), Some(breaks)) => {
                let story = read_story_range(word, clx, start, length)?;
                let boxes = boxes(&story.text, length, plc, breaks)?;
                // Every textbox is projected independently: no field may
                // cross a textbox boundary.
                let mut partitions = Vec::new();
                for textbox in boxes.iter().flatten() {
                    partitions.push(textbox.cp);
                    partitions.push(
                        textbox.cp + story.text[textbox.bytes.clone()].encode_utf16().count(),
                    );
                }
                partitions.sort_unstable();
                let table = header_fields::Table::read_at(word, table, field_table, length)?;
                let fields = StoryFields::analyze(&story.text, &table, &partitions)?;
                Ok(Some(Self {
                    story,
                    boxes,
                    fields,
                }))
            }
            _ => Err(unsupported("inconsistent Word textbox story tables")),
        }
    }

    /// The text of the FTXBXS selected by a shape's MSOPSText_lTxid, whose
    /// high word is its one-based PLC index (MS-DOC 2.9.106). The range ends
    /// with the textbox's final paragraph mark.
    pub fn text(&self, index: usize, spid: u32) -> Result<(&str, usize), String> {
        let textbox = index
            .checked_sub(1)
            .and_then(|index| self.boxes.get(index))
            .and_then(Option::as_ref)
            .ok_or_else(|| unsupported("Word shape text references no textbox"))?;
        if textbox.lid != spid {
            return Err(unsupported("Word textbox belongs to another shape"));
        }
        let text = &self.story.text[textbox.bytes.clone()];
        // MS-DOC 2.8.32: each range ends with a 0x0D that separates it from
        // the next range. When the range's content already ends with its
        // final paragraph (or cell) mark, that separator is not a paragraph:
        // Word's own DOCX of the corpus documents has exactly one paragraph
        // fewer than the range's marks in every textbox (26 boxes, five
        // documents). A range whose only final mark is the separator keeps
        // it as its last paragraph mark.
        match text.strip_suffix('\r') {
            Some(content) if content.ends_with(['\r', '\u{7}']) => Ok((content, textbox.cp)),
            _ => Ok((text, textbox.cp)),
        }
    }
}

fn plc_bytes<'t>(word: &[u8], table: &'t [u8], fib: usize) -> Result<Option<&'t [u8]>, String> {
    let size = u32_at(word, fib + 4)? as usize;
    if size == 0 {
        return Ok(None);
    }
    let offset = u32_at(word, fib)? as usize;
    table
        .get(offset..)
        .and_then(|bytes| bytes.get(..size))
        .map(Some)
        .ok_or_else(|| unsupported("Word textbox table out of bounds"))
}

fn boxes(
    text: &str,
    length: usize,
    plc: &[u8],
    breaks: &[u8],
) -> Result<Vec<Option<Textbox>>, String> {
    let count = plc_count(plc, 22)?;
    let mut ranges = Vec::new();
    for index in 0..count {
        let start = u32_at(plc, index * 4)? as usize;
        let end = u32_at(plc, (index + 1) * 4)? as usize;
        let entry = &plc[(count + 1) * 4 + index * 22..][..22];
        // The last FTXBXS is always reusable, regardless of fReusable.
        if index + 1 == count || u16_at(entry, 8)? != 0 {
            ranges.push(None);
            continue;
        }
        // MS-DOC 2.9.106: an actual textbox spans more than one CP and its
        // range ends with its final paragraph mark inside the story.
        if start + 1 >= end || end > length {
            return Err(unsupported("invalid Word textbox text range"));
        }
        ranges.push(Some((u32_at(entry, 14)?, start..end)));
    }
    validate_breaks(breaks, &ranges)?;

    // One forward UTF-16/UTF-8 translation of the aggregate story.
    let mut targets: Vec<usize> = ranges
        .iter()
        .flatten()
        .flat_map(|(_, range)| [range.start, range.end])
        .collect();
    targets.sort_unstable();
    targets.dedup();
    let mut offsets = std::collections::BTreeMap::new();
    let mut chars = text.char_indices();
    let (mut cp, mut byte) = (0usize, 0usize);
    for target in targets {
        while cp < target {
            let (at, character) = chars
                .next()
                .ok_or_else(|| unsupported("truncated Word textbox story"))?;
            cp += character.len_utf16();
            byte = at + character.len_utf8();
        }
        if cp != target {
            return Err(unsupported(
                "Word textbox boundary splits a Unicode character",
            ));
        }
        offsets.insert(target, byte);
    }
    ranges
        .into_iter()
        .map(|range| {
            range
                .map(|(lid, range)| {
                    let bytes = offsets[&range.start]..offsets[&range.end];
                    if !text[bytes.clone()].ends_with('\r') {
                        return Err(unsupported("Word textbox lacks its final paragraph mark"));
                    }
                    Ok(Textbox {
                        lid,
                        cp: range.start,
                        bytes,
                    })
                })
                .transpose()
        })
        .collect()
}

fn plc_count(plc: &[u8], element: usize) -> Result<usize, String> {
    if plc.len() < 4 || !(plc.len() - 4).is_multiple_of(4 + element) {
        return Err(unsupported("invalid Word textbox table size"));
    }
    let count = (plc.len() - 4) / (4 + element);
    if count == 0 {
        return Err(unsupported("invalid Word textbox count"));
    }
    if count > MAX_TEXTBOXES {
        return Err(crate::rejection::resource(
            "Word textbox count budget exceeded",
        ));
    }
    for index in 0..count {
        if u32_at(plc, index * 4)? >= u32_at(plc, (index + 1) * 4)? {
            return Err(unsupported("invalid Word textbox CP order"));
        }
    }
    Ok(count)
}

/// Require exactly one Tbkd per actual textbox, covering its whole range.
/// Several Tbkd ranges for one FTXBXS describe a linked chain (MS-DOC
/// 2.9.312), whose per-box text distribution is not reconstructed.
fn validate_breaks(breaks: &[u8], ranges: &[Option<(u32, Range<usize>)>]) -> Result<(), String> {
    let count = plc_count(breaks, 6)?;
    // The CP up to which the break descriptors of each textbox cover its
    // text. A chain of linked textboxes has one descriptor per box, each
    // continuing where the previous one ends (MS-DOC 2.8.30 PlcfTxbxBkd).
    let mut covered: Vec<Option<usize>> = vec![None; ranges.len()];
    // The final Tbkd is not associated with any FTXBXS.
    for index in 0..count - 1 {
        let start = u32_at(breaks, index * 4)? as usize;
        let end = u32_at(breaks, (index + 1) * 4)? as usize;
        let itxbxs = u16_at(breaks, (count + 1) * 4 + index * 6)? as i16;
        let Some(Some((_, range))) = usize::try_from(itxbxs)
            .ok()
            .and_then(|itxbxs| ranges.get(itxbxs))
        else {
            continue; // Spare structures carry no displayed text.
        };
        let slot = &mut covered[itxbxs as usize];
        if start != slot.unwrap_or(range.start) || end <= start || end > range.end {
            return Err(crate::rejection::corrupt(
                "Word textbox break descriptors do not cover its text in order",
            ));
        }
        *slot = Some(end);
    }
    for (range, covered) in ranges.iter().zip(&covered) {
        match (range, covered) {
            (Some((_, range)), Some(end)) if *end == range.end => {}
            (Some(_), Some(_)) => {
                return Err(crate::rejection::corrupt(
                    "Word textbox break descriptors do not cover its text in order",
                ))
            }
            (Some(_), None) => return Err(unsupported("Word textbox lacks its break descriptor")),
            (None, _) => {}
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn plc(cps: &[u32], elements: &[Vec<u8>]) -> Vec<u8> {
        let mut bytes: Vec<u8> = cps.iter().flat_map(|cp| cp.to_le_bytes()).collect();
        for element in elements {
            bytes.extend(element);
        }
        bytes
    }
    fn ftxbxs(reusable: bool, lid: u32) -> Vec<u8> {
        let mut bytes = vec![0; 22];
        bytes[8] = u8::from(reusable);
        bytes[14..18].copy_from_slice(&lid.to_le_bytes());
        bytes
    }
    fn tbkd(index: i16) -> Vec<u8> {
        let mut bytes = vec![0; 6];
        bytes[..2].copy_from_slice(&index.to_le_bytes());
        bytes
    }

    #[test]
    fn main_and_header_tables_distinguish_empty_counts_from_resource_limits() {
        for part in [Part::Main, Part::Header] {
            for limit_breaks in [false, true] {
                for count in [0, 100_000, 100_001] {
                    let spare_table = |count: usize, element: Vec<u8>| {
                        let cps: Vec<_> = (0..=count as u32).collect();
                        plc(&cps, &vec![element; count])
                    };
                    let boxes = spare_table(if limit_breaks { 1 } else { count }, ftxbxs(true, 0));
                    let breaks = spare_table(if limit_breaks { count } else { 1 }, tbkd(-1));
                    let mut word = vec![0; 0x402];
                    word[0x400..].copy_from_slice(&13u16.to_le_bytes());
                    let (length, table_fib, breaks_fib) = match part {
                        Part::Main => (0x64, 0x25a, 0x2f2),
                        Part::Header => (0x68, 0x26a, 0x2fa),
                    };
                    word[length..length + 4].copy_from_slice(&1u32.to_le_bytes());
                    word[table_fib + 4..table_fib + 8]
                        .copy_from_slice(&(boxes.len() as u32).to_le_bytes());
                    word[breaks_fib..breaks_fib + 4]
                        .copy_from_slice(&(boxes.len() as u32).to_le_bytes());
                    word[breaks_fib + 4..breaks_fib + 8]
                        .copy_from_slice(&(breaks.len() as u32).to_le_bytes());
                    let table = [boxes, breaks].concat();
                    let clx = [
                        vec![2],
                        16u32.to_le_bytes().to_vec(),
                        plc(&[0, 1], &[vec![0, 0, 0, 4, 0, 0, 0, 0]]),
                    ]
                    .concat();
                    let result = Textboxes::read(&word, &table, &clx, part);
                    match count {
                        100_000 => assert_eq!(result.ok().flatten().unwrap().story.text, "\r"),
                        100_001 => {
                            assert!(result.err().unwrap().starts_with("LEGACY_RESOURCE_LIMIT:"))
                        }
                        0 => assert!(result.err().unwrap().starts_with("UNSUPPORTED:")),
                        _ => unreachable!(),
                    }
                }
            }
        }
    }

    #[test]
    fn maps_each_actual_textbox_to_its_shape_and_final_paragraph_mark() {
        let text = "ab\rcé\r\r";
        let table = plc(
            &[0, 3, 6, 9],
            &[ftxbxs(false, 7), ftxbxs(false, 8), ftxbxs(false, 0)],
        );
        let breaks = plc(&[0, 3, 6, 9], &[tbkd(0), tbkd(1), tbkd(-1)]);
        let boxes = boxes(text, 7, &table, &breaks).unwrap();
        let story = Textboxes {
            story: Story {
                text: text.into(),
                pieces: Vec::new(),
                prcs: Vec::new(),
            },
            boxes,
            fields: StoryFields::analyze(text, &header_fields::Table::for_test(&[], 0), &[])
                .unwrap(),
        };
        assert_eq!(story.text(1, 7).unwrap(), ("ab\r", 0));
        assert_eq!(story.text(2, 8).unwrap(), ("cé\r", 3));
        // A range ending with its paragraph mark plus the separator drops
        // the separator.
        let text = "ab\r\rc\r";
        let table = plc(&[0, 4, 7], &[ftxbxs(false, 7), ftxbxs(false, 0)]);
        let breaks = plc(&[0, 4, 7], &[tbkd(0), tbkd(-1)]);
        let story = Textboxes {
            story: Story {
                text: text.into(),
                pieces: Vec::new(),
                prcs: Vec::new(),
            },
            boxes: super::boxes(text, 7, &table, &breaks).unwrap(),
            fields: StoryFields::analyze(text, &header_fields::Table::for_test(&[], 0), &[])
                .unwrap(),
        };
        assert_eq!(story.text(1, 7).unwrap(), ("ab\r", 0));
        // Spare, out-of-range and foreign-shape references fail closed.
        assert!(story.text(3, 0).is_err());
        assert!(story.text(0, 7).is_err());
        assert!(story.text(2, 7).is_err());
    }

    #[test]
    fn reads_linked_chains_and_rejects_missing_breaks_and_unterminated_ranges() {
        let text = "ab\rcd\r\r";
        let plc2 = plc(&[0, 6, 7], &[ftxbxs(false, 7), ftxbxs(false, 0)]);
        // A chain of linked textboxes: one break descriptor per box, each
        // continuing the text of one FTXBXS. The story is one textbox.
        let chain = plc(&[0, 3, 6, 7], &[tbkd(0), tbkd(0), tbkd(-1)]);
        let linked = boxes(text, 7, &plc2, &chain).unwrap();
        assert_eq!(
            linked[0].as_ref().map(|textbox| textbox.bytes.clone()),
            Some(0..6)
        );
        // Descriptors that start late or stop short are malformed.
        for cps in [[1, 3, 6, 7], [0, 3, 5, 7]] {
            let broken = plc(&cps, &[tbkd(0), tbkd(0), tbkd(-1)]);
            assert!(boxes(text, 7, &plc2, &broken).is_err(), "{cps:?}");
        }
        let missing = plc(&[0, 6, 7], &[tbkd(1), tbkd(-1)]);
        assert!(boxes(text, 7, &plc2, &missing).is_err());
        let unterminated = plc(&[0, 2, 7], &[ftxbxs(false, 7), ftxbxs(false, 0)]);
        let breaks = plc(&[0, 2, 7], &[tbkd(0), tbkd(-1)]);
        assert!(boxes(text, 7, &unterminated, &breaks).is_err());
        let outside = plc(&[0, 8, 9], &[ftxbxs(false, 7), ftxbxs(false, 0)]);
        let breaks = plc(&[0, 8, 9], &[tbkd(0), tbkd(-1)]);
        assert!(boxes(text, 7, &outside, &breaks).is_err());
    }

    #[test]
    fn reusable_entries_carry_no_text() {
        let text = "ab\rx\r";
        let table = plc(
            &[0, 3, 4, 9],
            &[ftxbxs(false, 7), ftxbxs(true, 0), ftxbxs(false, 0)],
        );
        let breaks = plc(&[0, 3, 4], &[tbkd(0), tbkd(-1)]);
        let boxes = boxes(text, 5, &table, &breaks).unwrap();
        assert!(boxes[0].is_some());
        assert!(boxes[1].is_none() && boxes[2].is_none());
    }
}
