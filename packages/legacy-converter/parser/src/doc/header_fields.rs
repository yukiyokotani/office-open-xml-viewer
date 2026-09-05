//! Passive PAGE/NUMPAGES projection only. MS-DOC 2.8.25, 2.9.88/89/110;
//! ECMA-376 17.16.5.42/44 and 17.16.18. No external/calculating fields execute.
use super::{u32_at, unsupported, Paragraph, Token, MAX_STORY_CONTROLS};
use std::collections::{BTreeMap, VecDeque};

pub(super) struct Table(BTreeMap<usize, (u8, u8)>);
impl Table {
    pub fn read(word: &[u8], table: &[u8], length: usize) -> Result<Self, String> {
        let size = u32_at(word, 0x126)? as usize;
        let mut entries = BTreeMap::new();
        if size == 0 {
            return Ok(Self(entries));
        }
        if size < 4 || !(size - 4).is_multiple_of(6) || (size - 4) / 6 > MAX_STORY_CONTROLS {
            return Err(unsupported("invalid Word header field table length"));
        }
        let offset = u32_at(word, 0x122)? as usize;
        let plc = table
            .get(offset..)
            .and_then(|b| b.get(..size))
            .ok_or_else(|| unsupported("Word header field table outside its stream"))?;
        let count = (size - 4) / 6;
        let mut previous = None;
        for i in 0..=count {
            let cp = u32_at(plc, i * 4)? as usize;
            // MS-DOC 2.8.25: the final CP is only an ordering sentinel,
            // not a field location. Its value is otherwise undefined.
            if (i < count && cp >= length) || previous.is_some_and(|p| cp <= p) {
                return Err(unsupported("invalid Word header field position"));
            }
            previous = Some(cp);
            if i < count {
                let offset = (count + 1) * 4 + i * 2;
                let ch = plc[offset] & 0x1f; // fldch reserved bits MUST be ignored.
                if !matches!(ch, 0x13..=0x15) {
                    return Err(unsupported("invalid Word field marker"));
                }
                entries.insert(cp, (ch, plc[offset + 1]));
            }
        }
        Ok(Self(entries))
    }

    pub fn restore(&self, text: &str, base_cp: usize, paragraphs: &mut [Paragraph]) {
        let mut depth = 0usize;
        let mut start = 0;
        let mut instruction = String::new();
        let mut separator = None;
        let mut eligible = false;
        let mut events = VecDeque::new();
        let mut private_ranges = Vec::new();
        let mut starts = Vec::new();
        let mut cp = base_cp;
        for ch in text.chars() {
            match ch {
                '\u{13}' => {
                    starts.push(cp);
                    if depth == 0 {
                        start = cp;
                        instruction.clear();
                        separator = None;
                        eligible = self.0.get(&cp).is_some_and(|(c, _)| *c == 0x13);
                    } else {
                        eligible = false;
                    }
                    depth += 1;
                }
                '\u{14}' if depth == 1 => {
                    if separator.is_some() {
                        eligible = false;
                    }
                    separator = Some(cp);
                    eligible &= self.0.get(&cp).is_some_and(|(c, _)| *c == 0x14);
                }
                '\u{15}' if depth > 0 => {
                    depth -= 1;
                    let private_start = starts.pop().expect("matched field begin");
                    if self
                        .0
                        .get(&cp)
                        .is_some_and(|(c, flags)| *c == 0x15 && flags & 0x20 != 0)
                    {
                        private_ranges.push(private_start..=cp);
                    }
                    if depth == 0 && eligible && separator.is_some()
                        // fLocked keeps its cached display; fPrivateResult is
                        // not turned into a newly visible, recomputed value.
                        && self.0.get(&cp).is_some_and(|(c, flags)| *c == 0x15 && flags & 0x30 == 0)
                        && supported(&instruction)
                    {
                        events.push_back((Token::FieldBegin(instruction.clone()), start + 1));
                        events.push_back((Token::FieldEnd, cp));
                    }
                }
                '\r' | '\u{7}' | '\u{b}' | '\u{c}' if depth > 0 => eligible = false,
                _ if depth == 1 && separator.is_none() && eligible => {
                    if instruction.len() + ch.len_utf8() <= 512 && (ch >= ' ' || ch == '\t') {
                        instruction.push(ch);
                    } else {
                        eligible = false;
                    }
                }
                _ => {}
            }
            cp += ch.len_utf16();
        }
        // The tokenizer flushes at field controls, so no cached text token can
        // straddle an event. Merge once in CP order without rescanning paragraphs.
        private_ranges.sort_unstable_by_key(|range| *range.start());
        let mut private_ranges: VecDeque<_> = private_ranges.into();
        for paragraph in paragraphs {
            let original = std::mem::take(&mut paragraph.tokens);
            for token in original {
                while private_ranges
                    .front()
                    .is_some_and(|range| *range.end() < token.1)
                {
                    private_ranges.pop_front();
                }
                if private_ranges
                    .front()
                    .is_some_and(|range| range.contains(&token.1))
                {
                    continue;
                }
                while events.front().is_some_and(|(_, cp)| *cp <= token.1) {
                    paragraph.tokens.push(events.pop_front().unwrap());
                }
                paragraph.tokens.push(token);
            }
            while events
                .front()
                .is_some_and(|(_, cp)| *cp <= paragraph.end_cp)
            {
                paragraph.tokens.push(events.pop_front().unwrap());
            }
        }
    }
}

fn supported(instruction: &str) -> bool {
    let mut words = instruction.split_ascii_whitespace();
    if !words
        .next()
        .is_some_and(|w| w.eq_ignore_ascii_case("PAGE") || w.eq_ignore_ascii_case("NUMPAGES"))
    {
        return false;
    }
    while let Some(switch) = words.next() {
        if switch != "\\*" {
            return false;
        }
        let Some(format) = words.next() else {
            return false;
        };
        if !matches!(
            format.to_ascii_lowercase().as_str(),
            "mergeformat" | "charformat" | "arabic" | "roman" | "alphabetic"
        ) {
            return false;
        }
    }
    true
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn allows_only_passive_page_field_instructions() {
        for value in ["PAGE", " numpages ", "PAGE \\* roman \\* MERGEFORMAT"] {
            assert!(supported(value));
        }
        for value in [
            "PAGEREF X",
            "INCLUDETEXT file",
            "DATE",
            "PAGE \\*",
            "PAGE \\h",
            "PAGE \\* Unknown",
            "NUMPAGES \\# 0",
        ] {
            assert!(!supported(value));
        }
    }

    fn tokens(text: &str, flags: u8) -> Vec<Token> {
        let mut cp = 7;
        let mut entries = BTreeMap::new();
        for ch in text.chars() {
            if matches!(ch, '\u{13}'..='\u{15}') {
                entries.insert(cp, (ch as u8, if ch == '\u{15}' { flags } else { 0 }));
            }
            cp += ch.len_utf16();
        }
        let mut paragraphs =
            super::super::tokenize_with_fields(text, &mut super::super::Fields::default(), 7, true);
        Table(entries).restore(text, 7, &mut paragraphs);
        paragraphs
            .into_iter()
            .flat_map(|p| p.tokens.into_iter().map(|(t, _)| t))
            .collect()
    }

    #[test]
    fn keeps_field_events_ordered_with_astral_prefix_and_adjacent_fields() {
        let actual = tokens(
            "😀\u{13}PAGE\u{14}99\u{15}\u{13}NUMPAGES\u{14}99\u{15}Z\r",
            0x80,
        );
        assert_eq!(
            actual,
            vec![
                Token::Text("😀".into()),
                Token::FieldBegin("PAGE".into()),
                Token::Text("99".into()),
                Token::FieldEnd,
                Token::FieldBegin("NUMPAGES".into()),
                Token::Text("99".into()),
                Token::FieldEnd,
                Token::Text("Z".into())
            ]
        );
    }
    #[test]
    fn locks_nested_multiline_incomplete_and_oversized_fields_do_not_recompute() {
        for text in [
            "\u{13}PAGE\u{14}99\u{15}".into(),
            "\u{13}PAGE\u{14}\u{13}PAGE\u{14}99\u{15}\u{15}".into(),
            "\u{13}PAGE\r\u{14}99\u{15}".into(),
            "\u{13}PAGE\u{14}99".into(),
            format!("\u{13}PAGE{}\u{14}99\u{15}", " ".repeat(513)),
        ] {
            let flags = if text == "\u{13}PAGE\u{14}99\u{15}" {
                0x90
            } else {
                0x80
            };
            assert!(!tokens(&text, flags)
                .iter()
                .any(|t| matches!(t, Token::FieldBegin(_))));
        }
        assert_eq!(
            tokens("A\u{13}PAGE\u{14}secret\u{15}B", 0xa0),
            vec![Token::Text("A".into()), Token::Text("B".into())]
        );
        assert_eq!(
            tokens("A\u{13}IF\u{14}x\u{13}PAGE\u{14}secret\u{15}y\u{15}B", 0xa0),
            vec![Token::Text("A".into()), Token::Text("B".into())]
        );
    }

    #[test]
    fn validates_field_table_boundaries_and_ignores_reserved_marker_bits() {
        let mut word = vec![0; 0x12a];
        word[0x126..0x12a].copy_from_slice(&10u32.to_le_bytes());
        let mut table = Vec::from(2u32.to_le_bytes());
        table.extend(4u32.to_le_bytes());
        table.extend([0xf3, 0]);
        assert_eq!(Table::read(&word, &table, 4).unwrap().0[&2].0, 0x13);
        table[4..8].copy_from_slice(&u32::MAX.to_le_bytes());
        assert!(Table::read(&word, &table, 4).is_ok());
        for cp in [2u32, 1] {
            table[4..8].copy_from_slice(&cp.to_le_bytes());
            assert!(Table::read(&word, &table, 4).is_err());
        }
        table[0..4].copy_from_slice(&4u32.to_le_bytes());
        table[4..8].copy_from_slice(&6u32.to_le_bytes());
        assert!(Table::read(&word, &table, 4).is_err());
        word[0x126..0x12a].copy_from_slice(&0u32.to_le_bytes());
        word[0x122..0x126].copy_from_slice(&u32::MAX.to_le_bytes());
        assert!(Table::read(&word, &[], 4).unwrap().0.is_empty());
    }
}
