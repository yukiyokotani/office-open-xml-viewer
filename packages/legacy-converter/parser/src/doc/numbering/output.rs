//! Ordinary WordprocessingML numbering (ECMA-376 17.9). Paragraph-context
//! variants share an abstract definition: marker formatting must not create a
//! new list sequence. MS-DOC 2.4.6.4 distinguishes LSID from iLfo; start-at
//! overrides apply once per original LFO/level, not once per formatting variant.

use super::{Level, Reference, Tables};
use crate::ooxml::xml_attr;
use std::collections::{BTreeMap, BTreeSet};

const OPEN: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:numbering xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">";
const CLOSE: &str = "</w:numbering>";

#[derive(Default)]
pub struct Store {
    scope: usize,
    abstracts: BTreeMap<(usize, i32), (usize, String)>,
    instances: BTreeMap<Key, (usize, String)>,
    started: BTreeSet<(usize, usize, u8)>,
    unrepresentable: BTreeSet<(usize, u8)>,
    bytes: usize,
    pub omitted: bool,
}

#[derive(PartialEq, Eq, PartialOrd, Ord)]
struct Key {
    scope: usize,
    lfo: usize,
    level: u8,
    starts: [Option<u16>; 9],
    ppr: String,
    rpr: String,
}

impl Store {
    pub fn begin_story(&mut self) {
        // MS-DOC 2.2.3: main story, each header/footer and each note are
        // separate valid-selection ranges. Share numbering only within one.
        self.scope += 1;
    }

    pub fn activate(
        &mut self,
        tables: &Tables<'_>,
        reference: Reference,
        ppr: String,
        rpr: String,
        limit: usize,
    ) -> Result<Option<usize>, String> {
        let selection = tables.resolve(reference)?;
        if self
            .unrepresentable
            .contains(&(reference.index, reference.level))
        {
            return Ok(None);
        }
        if selection.instance.auto_number_field.is_some() {
            self.omitted = true; // Retain cached field results, never execute.
            return Ok(None);
        }
        let starts = std::array::from_fn(|index| {
            if self
                .started
                .contains(&(self.scope, reference.index, index as u8))
            {
                None
            } else {
                if index == usize::from(reference.level) {
                    selection.start_override
                } else {
                    selection
                        .instance
                        .levels
                        .iter()
                        .find(|o| usize::from(o.index) == index)
                        .and_then(|o| o.start)
                }
            }
        });
        let key = Key {
            scope: self.scope,
            lfo: reference.index,
            level: reference.level,
            starts,
            ppr,
            rpr,
        };
        if let Some((id, _)) = self.instances.get(&key) {
            return Ok(Some(*id));
        }
        let abstract_key = (self.scope, selection.list.id);
        let abstract_id = if let Some((id, _)) = self.abstracts.get(&abstract_key) {
            *id
        } else {
            let id = self.abstracts.len();
            let kind = if selection.list.simple {
                "singleLevel"
            } else if selection.list.hybrid {
                "hybridMultilevel"
            } else {
                "multilevel"
            };
            let mut xml = format!(
                "<w:abstractNum w:abstractNumId=\"{id}\"><w:multiLevelType w:val=\"{kind}\"/>"
            );
            let levels = &selection.list.levels;
            for (index, level) in levels.iter().enumerate() {
                let formats = |n: u8| levels.get(usize::from(n)).map(|l| l.format);
                // Every used level receives a full context override below.
                // The abstract supplies counter defaults; an unrepresentable
                // unused template must not disable unrelated live levels.
                let lvl = level_xml(level, index as u8, formats, "", "", false)?
                    .expect("counter-only definition");
                append(&mut xml, &lvl, limit)?;
            }
            append(&mut xml, "</w:abstractNum>", limit)?;
            self.charge(xml.len(), limit)?;
            self.abstracts.insert(abstract_key, (id, xml));
            id
        };
        let id = self.instances.len() + 1;
        let mut xml = format!("<w:num w:numId=\"{id}\"><w:abstractNumId w:val=\"{abstract_id}\"/>");
        // Only this paragraph's level receives its effective mark/paragraph
        // formatting. Other replacement levels still supply ancestor formats.
        for index in 0..9_u8 {
            // MS-DOC 2.4.6.3 part 2 step 4 / 2.9.150 fLegal reformats
            // every referenced number, including the current level, but keeps
            // msonfcArabicLZ. ECMA-376 17.9.4 isLgl would erase that padding.
            // Express the effective formats in this paragraph's instance;
            // never change the shared source formats or counter identity.
            let legal_context = selection.level.legal && index <= reference.level;
            let replacement = selection
                .instance
                .levels
                .iter()
                .find(|o| o.index == index)
                .and_then(|o| o.formatting.as_ref());
            if index != reference.level
                && replacement.is_none()
                && key.starts[usize::from(index)].is_none()
                && !legal_context
            {
                continue;
            }
            append(
                &mut xml,
                &format!("<w:lvlOverride w:ilvl=\"{index}\">"),
                limit,
            )?;
            if let Some(start) = key.starts[usize::from(index)] {
                append(
                    &mut xml,
                    &format!("<w:startOverride w:val=\"{start}\"/>"),
                    limit,
                )?;
            }
            if index != reference.level && replacement.is_none() && !legal_context {
                append(&mut xml, "</w:lvlOverride>", limit)?;
                continue;
            }
            let mut level = *replacement
                .or_else(|| selection.list.levels.get(usize::from(index)))
                .ok_or_else(|| super::unsupported("missing Word list level"))?;
            if legal_context && level.start.is_some() && level.format != 0x16 {
                level.format = 0; // msonfcArabic; bullet/none have no sequence.
            }
            let formats = |n: u8| {
                selection
                    .instance
                    .levels
                    .iter()
                    .find(|o| o.index == n)
                    .and_then(|o| o.formatting.as_ref())
                    .or_else(|| selection.list.levels.get(usize::from(n)))
                    .map(|l| l.format)
            };
            let (ppr, rpr) = if index == reference.level {
                (key.ppr.as_str(), key.rpr.as_str())
            } else {
                ("", "")
            };
            let Some(lvl) = level_xml(&level, index, formats, ppr, rpr, index == reference.level)?
            else {
                self.omitted = true;
                self.unrepresentable
                    .insert((reference.index, reference.level));
                return Ok(None);
            };
            append(&mut xml, &lvl, limit)?;
            append(&mut xml, "</w:lvlOverride>", limit)?;
        }
        append(&mut xml, "</w:num>", limit)?;
        // Include retained context keys in the independent metadata ceiling;
        // these are bounded caches, not one unbounded entry per paragraph.
        self.charge(xml.len() + key.ppr.len() + key.rpr.len(), limit)?;
        self.instances.insert(key, (id, xml));
        self.started
            .insert((self.scope, reference.index, reference.level));
        Ok(Some(id))
    }

    fn charge(&mut self, extra: usize, limit: usize) -> Result<(), String> {
        self.bytes = self
            .bytes
            .checked_add(extra)
            .filter(|n| *n <= limit)
            .ok_or("OUTPUT_TOO_LARGE")?;
        Ok(())
    }

    pub fn xml(&self, limit: usize) -> Result<Option<String>, String> {
        if self.instances.is_empty() {
            return Ok(None);
        }
        let mut xml = OPEN.to_string();
        for (_, part) in self.abstracts.values() {
            append(&mut xml, part, limit)?;
        }
        for (_, part) in self.instances.values() {
            append(&mut xml, part, limit)?;
        }
        append(&mut xml, CLOSE, limit)?;
        Ok(Some(xml))
    }
}

fn append(xml: &mut String, value: &str, limit: usize) -> Result<(), String> {
    if value.len() > limit.saturating_sub(xml.len()) {
        return Err("OUTPUT_TOO_LARGE".into());
    }
    xml.push_str(value);
    Ok(())
}

fn level_xml(
    level: &Level<'_>,
    index: u8,
    formats: impl Fn(u8) -> Option<u8>,
    ppr: &str,
    rpr: &str,
    live: bool,
) -> Result<Option<String>, String> {
    let text = match template(level, formats)? {
        Some(text) => text,
        None if live => return Ok(None),
        None => String::new(),
    };
    let mut xml = format!(
        "<w:lvl w:ilvl=\"{index}\" w:tentative=\"{}\">",
        u8::from(level.tentative)
    );
    if let Some(start) = level.start {
        xml.push_str(&format!("<w:start w:val=\"{start}\"/>"));
    }
    xml.push_str(&format!(
        "<w:numFmt w:val=\"{}\"/>",
        super::super::number_format::name(level.format)?
    ));
    if let Some(restart) = level.restart {
        xml.push_str(&format!("<w:lvlRestart w:val=\"{restart}\"/>"));
    }
    xml.push_str(&format!(
        "<w:suff w:val=\"{}\"/><w:lvlText w:val=\"{}\"/><w:lvlJc w:val=\"{}\"/>",
        ["tab", "space", "nothing"][usize::from(level.follow)],
        xml_attr(&text),
        ["left", "center", "right"][usize::from(level.justification)]
    ));
    if !ppr.is_empty() {
        xml.push_str("<w:pPr>");
        xml.push_str(ppr);
        xml.push_str("</w:pPr>");
    }
    xml.push_str(rpr);
    xml.push_str("</w:lvl>");
    Ok(Some(xml))
}

fn template(
    level: &Level<'_>,
    formats: impl Fn(u8) -> Option<u8>,
) -> Result<Option<String>, String> {
    if level.format == 0x17
        && level
            .text
            .chunks_exact(2)
            .next()
            .is_some_and(|pair| u16::from_le_bytes([pair[0], pair[1]]) & 0xf000 != 0)
    {
        // MS-DOC 2.4.6.3 masks high-nibble binary bullet characters before
        // applying their marker font. Until that font-aware mapping is
        // represented in OOXML, neither the raw nor masked code point is safe.
        return Ok(None);
    }
    let mut units = Vec::with_capacity(level.text.len() / 2 + 9);
    let mut previous_literal_percent = false;
    for (index, pair) in level.text.chunks_exact(2).enumerate() {
        let unit = u16::from_le_bytes([pair[0], pair[1]]);
        if let Some((_, target)) = level
            .placeholders
            .iter()
            .flatten()
            .find(|(offset, _)| usize::from(*offset) == index + 1)
        {
            if !matches!(formats(*target), Some(0x17 | 0xff) | None) {
                units.extend([b'%' as u16, u16::from(b'1' + *target)]);
                previous_literal_percent = false;
            }
        } else {
            // OOXML lvlText has no escaping convention for literal "%1".
            // Do not silently turn authored text into a dynamic counter.
            if previous_literal_percent && (b'1' as u16..=b'9' as u16).contains(&unit) {
                return Ok(None);
            }
            if unit < 0x20 || matches!(unit, 0xfffe | 0xffff) {
                return Ok(None);
            }
            previous_literal_percent = unit == u16::from(b'%');
            units.push(unit);
        }
    }
    String::from_utf16(&units)
        .map(Some)
        .map_err(|_| super::unsupported("invalid Word numbering text Unicode"))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn legal_context_uses_effective_lfo_formats_without_mutating_shared_levels() {
        // The LVLF-permitted MSONFC values, plus the no-sequence sentinel.
        for format in (0..=59)
            .chain([0xff])
            .filter(|f| !matches!(f, 8 | 9 | 15 | 19))
        {
            for legal in [false, true] {
                let (word, table) = super::super::tests::one();
                let mut tables = Tables::read(&word, &table).unwrap();
                let base = tables.lists[0].levels[0];
                tables.lists[0].simple = false;
                tables.lists[0].levels = vec![base; 9];
                tables.lists[0].levels[1].legal = !legal;
                let mut ancestor = base;
                ancestor.format = format;
                ancestor.start = (!matches!(format, 0x17 | 0xff)).then_some(1);
                let mut current = base;
                current.format = 4;
                current.legal = legal;
                tables.overrides[0].levels = vec![
                    super::super::LevelOverride {
                        index: 0,
                        start: None,
                        formatting: Some(ancestor),
                    },
                    super::super::LevelOverride {
                        index: 1,
                        start: None,
                        formatting: Some(current),
                    },
                ];
                let mut store = Store::default();
                let reference = Reference::new(1, 1).unwrap().unwrap();
                store
                    .activate(&tables, reference, String::new(), String::new(), 20_000)
                    .unwrap();
                let (_, xml) = store.instances.values().next().unwrap();
                let expected = if legal && ancestor.start.is_some() && format != 0x16 {
                    0
                } else {
                    format
                };
                let expected = super::super::super::number_format::name(expected).unwrap();
                let ancestor_xml = xml.split("<w:lvlOverride w:ilvl=\"1\">").next().unwrap();
                assert!(ancestor_xml.contains(&format!("<w:numFmt w:val=\"{expected}\"/>")));
                let current_xml = xml.split("<w:lvlOverride w:ilvl=\"1\">").nth(1).unwrap();
                assert!(current_xml.contains(if legal {
                    "w:val=\"decimal\""
                } else {
                    "w:val=\"lowerLetter\""
                }));
                assert!(!store.xml(20_000).unwrap().unwrap().contains("<w:isLgl"));
                assert_eq!(
                    tables.overrides[0].levels[0]
                        .formatting
                        .as_ref()
                        .unwrap()
                        .format,
                    format
                );
                assert_eq!(
                    tables.overrides[0].levels[1]
                        .formatting
                        .as_ref()
                        .unwrap()
                        .format,
                    4
                );
                assert_eq!(tables.lists[0].levels[1].legal, !legal);
            }
        }
    }

    #[test]
    fn deepest_legal_context_is_bounded_cached_and_does_not_replace_normal_contexts() {
        let (word, table) = super::super::tests::one();
        let mut tables = Tables::read(&word, &table).unwrap();
        let mut base = tables.lists[0].levels[0];
        base.format = 1;
        tables.lists[0].simple = false;
        tables.lists[0].levels = vec![base; 9];
        tables.lists[0].levels[8].legal = true;
        let mut store = Store::default();
        let normal = Reference::new(1, 0).unwrap().unwrap();
        let legal = Reference::new(1, 8).unwrap().unwrap();
        assert_eq!(
            store
                .activate(&tables, normal, String::new(), String::new(), 20_000)
                .unwrap(),
            Some(1)
        );
        let normal_xml = store.instances.values().next().unwrap().1.clone();
        for _ in 0..1000 {
            assert_eq!(
                store
                    .activate(&tables, legal, String::new(), String::new(), 20_000)
                    .unwrap(),
                Some(2)
            );
        }
        assert_eq!(
            store
                .activate(&tables, normal, String::new(), String::new(), 20_000)
                .unwrap(),
            Some(1)
        );
        assert_eq!(store.instances.len(), 2);
        assert_eq!(store.abstracts.len(), 1);
        assert_eq!(
            store.instances.values().find(|(id, _)| *id == 1).unwrap().1,
            normal_xml
        );
        let (_, legal_xml) = store.instances.values().find(|(id, _)| *id == 2).unwrap();
        assert_eq!(legal_xml.matches("<w:lvlOverride ").count(), 9);
        assert_eq!(
            legal_xml.matches("<w:numFmt w:val=\"decimal\"/>").count(),
            9
        );
        let bytes = store.bytes;
        let mut bounded = Store::default();
        bounded
            .activate(&tables, normal, String::new(), String::new(), bytes)
            .unwrap();
        bounded
            .activate(&tables, legal, String::new(), String::new(), bytes)
            .unwrap();
        let mut insufficient = Store::default();
        insufficient
            .activate(&tables, normal, String::new(), String::new(), bytes - 1)
            .unwrap();
        assert_eq!(
            insufficient
                .activate(&tables, legal, String::new(), String::new(), bytes - 1)
                .unwrap_err(),
            "OUTPUT_TOO_LARGE"
        );
    }

    #[test]
    fn caches_repeated_contexts_bounds_retained_xml_and_separates_stories() {
        let (word, table) = super::super::tests::one();
        let tables = Tables::read(&word, &table).unwrap();
        let reference = Reference::new(1, 0).unwrap().unwrap();
        let mut store = Store::default();
        let ppr = "<w:ind w:left=\"720\" w:hanging=\"360\"/>";
        for _ in 0..1000 {
            assert_eq!(
                store
                    .activate(&tables, reference, ppr.into(), "<w:rPr/>".into(), 10_000)
                    .unwrap(),
                Some(1)
            );
        }
        assert_eq!(store.instances.len(), 1);
        assert_eq!(store.abstracts.len(), 1);
        let xml = store.xml(10_000).unwrap().unwrap();
        assert!(store.xml(xml.len() - 1).is_err());
        assert_eq!(store.xml(xml.len()).unwrap().unwrap(), xml);
        store.begin_story();
        assert_eq!(
            store
                .activate(&tables, reference, ppr.into(), "<w:rPr/>".into(), 10_000)
                .unwrap(),
            Some(2)
        );
        assert_eq!(store.abstracts.len(), 2);
        assert!(Store::default()
            .activate(&tables, reference, ppr.into(), String::new(), 10)
            .is_err());
    }

    #[test]
    fn template_offsets_are_utf16_and_do_not_reinterpret_literals_or_missing_sequences() {
        let (word, table) = super::super::tests::one();
        let tables = Tables::read(&word, &table).unwrap();
        let mut level = tables.lists[0].levels[0];
        level.text = &[0x3d, 0xd8, 0x00, 0xde, 0, 0, b'&', 0];
        level.placeholders[0] = Some((3, 0));
        assert_eq!(template(&level, |_| Some(0)).unwrap().unwrap(), "😀%1&");
        let xml = level_xml(&level, 0, |_| Some(0), "", "", true)
            .unwrap()
            .unwrap();
        assert!(xml.contains("😀%1&amp;"));
        level.text = &[b'%', 0, 0, 0, b'1', 0];
        level.placeholders[0] = Some((2, 0));
        assert!(template(&level, |_| Some(0x17)).unwrap().is_none());
        assert_eq!(template(&level, |_| Some(0)).unwrap().unwrap(), "%%11");
        level.text = &[0, 0, b'.', 0];
        level.placeholders[0] = Some((1, 0));
        assert_eq!(template(&level, |_| Some(0xff)).unwrap().unwrap(), ".");
    }

    #[test]
    fn emits_low_nibble_bullets_and_caches_high_nibble_bullet_omission() {
        let (word, table) = super::super::tests::one();
        let mut tables = Tables::read(&word, &table).unwrap();
        let reference = Reference::new(1, 0).unwrap().unwrap();
        let level = &mut tables.lists[0].levels[0];
        level.format = 0x17;
        level.start = None;
        level.restart = None;
        level.placeholders = [None; 9];
        level.text = &[0xb7, 0x00];
        assert_eq!(template(level, |_| Some(0x17)).unwrap(), Some("·".into()));

        for text in [&[0x22, 0x20][..], &[0xb7, 0xf0][..]] {
            tables.lists[0].levels[0].text = text;
            let mut store = Store::default();
            assert_eq!(
                store
                    .activate(&tables, reference, String::new(), String::new(), 10_000)
                    .unwrap(),
                None
            );
            let bytes = store.bytes;
            assert_eq!(
                store
                    .activate(&tables, reference, String::new(), String::new(), 10_000)
                    .unwrap(),
                None
            );
            assert!(store.omitted);
            assert_eq!(store.unrepresentable.len(), 1);
            assert_eq!(store.instances.len(), 0);
            assert_eq!(store.bytes, bytes);
        }
    }
}
