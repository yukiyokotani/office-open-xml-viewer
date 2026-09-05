//! MS-DOC 2.3.3 / 2.8.22: six separators then six stories per section.
//! Empty CP ranges inherit; nonempty blank paragraphs explicitly clear a header.
use super::{
    build_formatted_story, formatting, header_fields, pictures, read_story_range, sections, u32_at,
    unsupported, Content, Story, StoryParts, MAX_STORY_CONTROLS,
};
use std::ops::Range;

pub(super) struct Headers<'a> {
    pub story: Story<'a>,
    pub entries: Vec<Entry>,
    fields: header_fields::Table,
}

pub(super) struct Entry {
    pub index: usize,
    pub cp: usize,
    pub text: Range<usize>,
}

impl Headers<'_> {
    pub fn attach_references(&self, sections: &mut [sections::Section]) {
        let mut entries = self.entries.iter().peekable();
        for (index, section) in sections.iter_mut().enumerate() {
            let mut references = String::new();
            while entries.peek().is_some_and(|entry| entry.section() == index) {
                references.push_str(&entries.next().unwrap().reference());
            }
            // CT_SectPr: header/footer references precede the page geometry.
            section.xml.insert_str("<w:sectPr>".len(), &references);
        }
    }
    pub fn build_parts(
        &self,
        formatting: &mut formatting::Formatting<'_>,
        pictures: &mut pictures::Store<'_>,
        mut remaining: usize,
    ) -> Result<StoryParts, String> {
        let mut output = StoryParts::default();
        for entry in &self.entries {
            let text = &self.story.text[entry.text.clone()];
            if text.ends_with('\u{7}') {
                // MS-DOC 2.4.3: a depth-one table's terminating paragraph
                // is U+0007, not U+000D. Require its actual row properties.
                let cp = entry.cp + text.encode_utf16().count() - 1;
                let (_, fc, piece) = self
                    .story
                    .position(cp)
                    .ok_or_else(|| unsupported("Word header table mark outside story"))?;
                let properties = formatting.table_properties(fc, piece.prm, &self.story.prcs)?;
                if properties.depth()? != 1 || !properties.row_end {
                    return Err(unsupported(
                        "Word header lacks a final paragraph or table row",
                    ));
                }
            }
            output.omitted_floating |= text.contains('\u{8}');
            pictures.begin_part();
            let xml = build_formatted_story(
                &self.story,
                Content::HeaderFooter {
                    kind: entry.kind(),
                    text,
                    cp: entry.cp,
                    fields: &self.fields,
                },
                Some(formatting),
                Some(pictures),
                None,
                remaining,
            )?;
            remaining = remaining.checked_sub(xml.len()).ok_or("OUTPUT_TOO_LARGE")?;
            let filename = entry.filename();
            output.parts.push((format!("word/{filename}"), xml));
            let relationships = pictures.relationships();
            if !relationships.is_empty() {
                let xml = format!(
                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{relationships}</Relationships>"#
                );
                remaining = remaining.checked_sub(xml.len()).ok_or("OUTPUT_TOO_LARGE")?;
                output
                    .parts
                    .push((format!("word/_rels/{filename}.rels"), xml));
            }
            output.content_types.push_str(&format!(r#"<Override PartName="/word/{filename}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{}+xml"/>"#, entry.kind()));
            output.relationships.push_str(&format!(r#"<Relationship Id="{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{}" Target="{filename}"/>"#, entry.id(), entry.kind()));
        }
        Ok(output)
    }
}

impl Entry {
    pub fn section(&self) -> usize {
        self.index / 6
    }
    pub fn kind(&self) -> &'static str {
        if matches!(self.index % 6, 0 | 1 | 4) {
            "header"
        } else {
            "footer"
        }
    }
    pub fn variant(&self) -> &'static str {
        match self.index % 6 {
            0 | 2 => "even",
            1 | 3 => "default",
            _ => "first",
        }
    }
    pub fn id(&self) -> String {
        format!("rIdHf{}", self.index + 1)
    }
    pub fn filename(&self) -> String {
        format!("{}{}.xml", self.kind(), self.index + 1)
    }
    pub fn reference(&self) -> String {
        format!(
            r#"<w:{}Reference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:type="{}" r:id="{}"/>"#,
            self.kind(),
            self.variant(),
            self.id()
        )
    }
}

pub(super) fn read<'a>(
    word: &[u8],
    table: &[u8],
    clx: &'a [u8],
    sections: usize,
) -> Result<Option<Headers<'a>>, String> {
    let length = u32_at(word, 0x54)? as usize; // FibRgLw97.ccpHdd
    let size = u32_at(word, 0xf6)? as usize; // FibRgFcLcb97.lcbPlcfHdd
    if length == 0 && size == 0 {
        return Ok(None);
    }
    if length == 0 || size == 0 || sections == 0 || sections > 16384 {
        return Err(unsupported("inconsistent Word header document"));
    }
    // CP-only PLC: story starts, final story end, and one undefined CP.
    let count = 6 + sections * 6;
    if size != (count + 2) * 4 {
        return Err(unsupported("invalid Word header table length"));
    }
    let offset = u32_at(word, 0xf2)? as usize;
    let plc = table
        .get(offset..)
        .and_then(|b| b.get(..size))
        .ok_or_else(|| unsupported("Word header table outside its stream"))?;
    let mut cps = Vec::with_capacity(count + 1);
    for i in 0..=count {
        let cp = u32_at(plc, i * 4)? as usize;
        if cp >= length || cps.last().is_some_and(|previous| cp < *previous) {
            return Err(unsupported("invalid Word header character range"));
        }
        cps.push(cp);
    }
    if cps[0] != 0 || cps[count] != length - 1 {
        return Err(unsupported("Word header story coverage mismatch"));
    }
    let main = u32_at(word, 0x4c)? as usize;
    let footnotes = u32_at(word, 0x50)? as usize;
    if main > i32::MAX as usize || footnotes > i32::MAX as usize {
        return Err(unsupported("negative Word story length"));
    }
    let start = main
        .checked_add(footnotes)
        .ok_or_else(|| unsupported("Word header range overflow"))?;
    let story = read_story_range(word, clx, start, length)?;
    if story
        .text
        .bytes()
        .filter(|b| *b < 32)
        .take(MAX_STORY_CONTROLS + 1)
        .count()
        > MAX_STORY_CONTROLS
    {
        return Err(unsupported("Word header structure budget exceeded"));
    }
    let entries = split(&story.text, &cps)?;
    let fields = header_fields::Table::read(word, table, length)?;
    Ok(Some(Headers {
        story,
        entries,
        fields,
    }))
}

fn split(text: &str, cps: &[usize]) -> Result<Vec<Entry>, String> {
    // One forward UTF-16/UTF-8 translation for the aggregate header document:
    // no per-header scanning/cloning of the whole piece/property table.
    let mut chars = text.char_indices();
    let (mut cp, mut byte) = (0, 0);
    let mut offsets = Vec::with_capacity(cps.len());
    for target in cps {
        while cp < *target {
            let (at, ch) = chars
                .next()
                .ok_or_else(|| unsupported("truncated Word header text"))?;
            cp += ch.len_utf16();
            byte = at + ch.len_utf8();
        }
        if cp != *target {
            return Err(unsupported(
                "Word header boundary splits a Unicode character",
            ));
        }
        offsets.push(byte);
    }
    let mut entries = Vec::new();
    for i in 0..cps.len() - 1 {
        if cps[i] == cps[i + 1] {
            continue;
        }
        let content = &text[offsets[i]..offsets[i + 1]];
        if !content.ends_with('\r')
            || (i >= 6 && !content.ends_with("\r\r") && !content.ends_with("\u{7}\r"))
        {
            return Err(unsupported(
                "Word header story lacks its paragraph/guard mark",
            ));
        }
        if i >= 6 {
            // Resource policy: parts and per-part relationships are retained
            // until packaging, independently of the compressed byte ceiling.
            if entries.len() >= 4096 {
                return Err(unsupported("Word header part budget exceeded"));
            }
            entries.push(Entry {
                index: i - 6,
                cp: cps[i],
                text: offsets[i]..offsets[i + 1] - 1,
            });
        }
    }
    Ok(entries)
}

#[cfg(test)]
mod tests {
    use super::*;
    fn fixture() -> (Vec<u8>, Vec<u8>, Vec<u8>) {
        let text = "BODY\rFTN\rA\r\r\r";
        let mut word = vec![0; 1024];
        word[0x4c..0x50].copy_from_slice(&5u32.to_le_bytes());
        word[0x50..0x54].copy_from_slice(&4u32.to_le_bytes());
        word[0x54..0x58].copy_from_slice(&4u32.to_le_bytes());
        word[0xf6..0xfa].copy_from_slice(&56u32.to_le_bytes());
        for (i, unit) in text.encode_utf16().enumerate() {
            word[512 + i * 2..514 + i * 2].copy_from_slice(&unit.to_le_bytes());
        }
        let mut table = vec![0; 32];
        for _ in 0..5 {
            table.extend(3u32.to_le_bytes());
        }
        table.extend(u32::MAX.to_le_bytes());
        let mut clx = vec![2];
        clx.extend(16u32.to_le_bytes());
        clx.extend(0u32.to_le_bytes());
        clx.extend((text.len() as u32).to_le_bytes());
        clx.extend([0, 0]);
        clx.extend(512u32.to_le_bytes());
        clx.extend([0, 0]);
        (word, table, clx)
    }

    #[test]
    fn reads_header_after_footnotes_and_ignores_undefined_final_cp() {
        let (word, table, clx) = fixture();
        let headers = read(&word, &table, &clx, 1).unwrap().unwrap();
        assert_eq!(headers.story.text, "A\r\r\r");
        assert_eq!(headers.entries.len(), 1);
        assert_eq!(headers.entries[0].index, 1);
        assert_eq!(headers.story.position(0).unwrap().1, 530);
        assert_eq!(&headers.story.text[headers.entries[0].text.clone()], "A\r");
    }

    #[test]
    fn rejects_inconsistent_or_unbounded_header_metadata() {
        for (offset, value) in [
            (0x54, 0),
            (0x54, 3),
            (0xf6, 0),
            (0xf6, 55),
            (0xf2, u32::MAX),
            (0x50, u32::MAX),
            (0x54, 64 * 1024 * 1024 + 1),
        ] {
            let (mut word, table, clx) = fixture();
            word[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
            assert!(read(&word, &table, &clx, 1).is_err());
        }
        let (mut word, table, clx) = fixture();
        assert!(read(&word, &table, &clx, 0).is_err());
        assert!(read(&word, &table, &clx, 2).is_err());
        word[0x54..0x58].fill(0);
        word[0xf6..0xfa].fill(0);
        word[0xf2..0xf6].fill(255);
        assert!(read(&word, &[], &clx, 0).unwrap().is_none());
        let text = "\r\r".repeat(4097);
        let mut cps = vec![0; 7];
        cps.extend((1..=4097).map(|n| n * 2));
        assert!(split(&text, &cps).is_err());
    }
    #[test]
    fn distinguishes_inherited_and_explicit_blank_across_sections() {
        // Skip six separators; first default header is text, second is inherited,
        // third is an explicit empty paragraph. An astral scalar counts as 2 CPs.
        let text = "A😀\r\r\r\r\r";
        let mut cps = vec![0; 7];
        cps.extend([0, 5, 5, 5, 5, 5]);
        cps.extend([5; 6]);
        cps.extend([5, 7, 7, 7, 7, 7]);
        let entries = split(text, &cps).unwrap();
        assert_eq!(entries.len(), 2);
        assert_eq!(entries[0].section(), 0);
        assert_eq!(&text[entries[0].text.clone()], "A😀\r");
        assert_eq!(entries[1].section(), 2);
        assert_eq!(&text[entries[1].text.clone()], "\r");
        assert_eq!(entries[1].variant(), "default");
    }
    #[test]
    fn rejects_broken_guards_unicode_boundaries_and_short_text() {
        for (text, end) in [("A\r", 2), ("😀\r\r", 1), ("A\r\r", 5)] {
            let mut cps = vec![0; 7];
            cps.push(end);
            assert!(split(text, &cps).is_err());
        }
    }
}
