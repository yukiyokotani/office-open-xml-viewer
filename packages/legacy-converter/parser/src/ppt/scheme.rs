//! Slide/master color schemes: MS-PPT 2.4.14.1-2, 2.5.10-15, 2.12.1-2.
use super::*;
use std::collections::BTreeMap;

// Resource policy: each master retains at most eight text types * five levels,
// plus five document-default levels. Bound decoded styles independently of bytes.
const MAX_MASTERS: usize = 10_000;

pub(super) type Scheme = [u32; 8];

#[derive(Clone, Copy, Default)]
struct Entry {
    local: Option<Scheme>,
    parent: Option<u32>,
    main_parent: Option<u32>,
}

fn entry(record: Record<'_>, budget: &mut usize) -> Result<Entry, String> {
    let mut result = Entry::default();
    let mut atom_seen = false;
    for atom in parse_records(record.payload, budget)? {
        match atom.kind {
            1007 => {
                if atom_seen || atom.version != 2 || atom.payload.len() != 24 {
                    return Err(unsupported("invalid PowerPoint SlideAtom for color scheme"));
                }
                atom_seen = true;
                let master = u32_at(atom.payload, 12)?;
                if record.kind != 1016 && master != 0 {
                    result.main_parent = Some(master);
                }
                // Main masters are roots: MS-PPT 2.5.10 requires masterIdRef=0.
                // Their SlideFlags do not identify a parent color scheme.
                if record.kind == 1016 && u32_at(atom.payload, 12)? != 0 {
                    return Err(unsupported("PowerPoint main master has a parent"));
                }
                if record.kind != 1016 && u16_at(atom.payload, 20)? & 2 != 0 {
                    result.parent = Some(u32_at(atom.payload, 12)?);
                }
            }
            2032 if atom.instance == 1 => {
                if result.local.is_some() || atom.version != 0 || atom.payload.len() != 32 {
                    return Err(unsupported("invalid PowerPoint slide color scheme"));
                }
                let mut colors = [0; 8];
                for (i, color) in colors.iter_mut().enumerate() {
                    // ColorStruct's high byte is unused, not an index/flag.
                    *color = u32_at(atom.payload, i * 4)? & 0xffffff;
                }
                result.local = Some(colors);
            }
            // Available-scheme lists (instance 6) are not the active scheme.
            _ => {}
        }
    }
    Ok(result)
}

pub(super) struct Resolver {
    masters: BTreeMap<u32, Entry>,
    cache: BTreeMap<u32, Option<Scheme>>,
    text_styles: BTreeMap<u32, std::rc::Rc<text_style::Master>>,
}
impl Resolver {
    pub fn new(
        document: &[u8],
        children: &[Record<'_>],
        offsets: &BTreeMap<u32, usize>,
        budget: &mut usize,
    ) -> Result<Self, String> {
        let lists: Vec<_> = children
            .iter()
            .filter(|r| r.kind == 4080 && r.instance == 1)
            .collect();
        if lists.len() > 1 {
            return Err(unsupported("duplicate PowerPoint master list"));
        }
        let mut masters = BTreeMap::new();
        let defaults = text_style::document_defaults(children, budget)?;
        let mut text_styles = BTreeMap::new();
        if let Some(list) = lists.first() {
            if list.version != 15 {
                return Err(unsupported("invalid PowerPoint master list"));
            }
            for item in parse_records(list.payload, budget)? {
                if masters.len() >= MAX_MASTERS {
                    return Err(unsupported("too many PowerPoint masters"));
                }
                if item.kind != 1011 || item.version != 0 || item.payload.len() != 20 {
                    return Err(unsupported("invalid PowerPoint master reference"));
                }
                let id = u32_at(item.payload, 12)?;
                if id == 0 || masters.contains_key(&id) {
                    return Err(unsupported("duplicate or zero PowerPoint master ID"));
                }
                let persist = u32_at(item.payload, 0)?;
                let offset = offsets
                    .get(&persist)
                    .ok_or_else(|| unsupported("unresolved PowerPoint master persist ID"))?;
                let record = parse_record_at(document, *offset, budget)?;
                if !matches!(record.kind, 1006 | 1016) || record.version != 15 {
                    return Err(unsupported("invalid PowerPoint master persist object"));
                }
                masters.insert(id, entry(record, budget)?);
                if record.kind == 1016 {
                    let records = parse_records(record.payload, budget)?;
                    text_styles.insert(
                        id,
                        std::rc::Rc::new(text_style::Master::parse(&records, &defaults, budget)?),
                    );
                }
            }
        }
        Ok(Self {
            masters,
            cache: BTreeMap::new(),
            text_styles,
        })
    }
    pub fn text_master(
        &self,
        slide: Record<'_>,
        budget: &mut usize,
    ) -> Result<Option<std::rc::Rc<text_style::Master>>, String> {
        let mut parent = entry(slide, budget)?.main_parent;
        let mut path = Vec::new();
        while let Some(id) = parent {
            *budget = budget
                .checked_sub(1)
                .ok_or_else(|| unsupported("PowerPoint text master work budget exceeded"))?;
            if path.len() >= MAX_DEPTH || path.contains(&id) {
                return Err(unsupported(
                    "cyclic or excessive PowerPoint text master inheritance",
                ));
            }
            path.push(id);
            let e = self
                .masters
                .get(&id)
                .ok_or_else(|| unsupported("unresolved PowerPoint text master"))?;
            if let Some(style) = self.text_styles.get(&id) {
                return Ok(Some(style.clone()));
            }
            parent = e.main_parent;
        }
        Ok(None)
    }
    pub fn slide(
        &mut self,
        record: Record<'_>,
        budget: &mut usize,
    ) -> Result<Option<Scheme>, String> {
        let e = entry(record, budget)?;
        match e.parent {
            Some(id) => self.master(id, &mut Vec::new(), budget),
            None => Ok(e.local),
        }
    }
    fn master(
        &mut self,
        id: u32,
        path: &mut Vec<u32>,
        budget: &mut usize,
    ) -> Result<Option<Scheme>, String> {
        *budget = budget
            .checked_sub(1)
            .ok_or_else(|| unsupported("PowerPoint scheme work budget exceeded"))?;
        if let Some(value) = self.cache.get(&id) {
            return Ok(*value);
        }
        // Implementation resource policy, not a format-defined depth limit.
        if path.len() >= MAX_DEPTH || path.contains(&id) {
            return Err(unsupported(
                "cyclic or excessive PowerPoint scheme inheritance",
            ));
        }
        let e = *self
            .masters
            .get(&id)
            .ok_or_else(|| unsupported("unresolved PowerPoint master scheme"))?;
        path.push(id);
        let value = match e.parent {
            Some(parent) => self.master(parent, path, budget)?,
            None => e.local,
        };
        path.pop();
        self.cache.insert(id, value);
        Ok(value)
    }
}

/// MS-PPT ColorIndexStruct uses the high byte for text's scheme index.
pub(super) fn text(color: u32, scheme: Option<&Scheme>) -> Option<u32> {
    match color >> 24 {
        0..=7 => scheme.map(|s| s[(color >> 24) as usize]),
        0xfe => Some(color & 0xffffff),
        _ => None,
    }
}

/// MS-ODRAW COLORREF uses fSchemeIndex and the low byte, unlike text colors.
pub(super) fn drawing(color: u32, scheme: Option<&Scheme>) -> Option<u32> {
    if color >> 24 == 0xff {
        return None;
    }
    match color & 0x1f000000 {
        0 | 0x04000000 => Some(color & 0xffffff), // fSystemRGB is literal RGB.
        0x08000000 => scheme.and_then(|s| s.get((color & 255) as usize).copied()),
        _ => None, // System/palette indices and combined references unresolved.
    }
}

#[cfg(test)]
mod tests {
    use super::super::persist::tests::record;
    use super::*;

    fn slide(kind: u16, parent: u32, inherit: bool, color: u32) -> Vec<u8> {
        let mut atom = [0; 24];
        atom[12..16].copy_from_slice(&parent.to_le_bytes());
        atom[20] = if inherit { 2 } else { 0 };
        record(
            15,
            kind,
            &[
                record(2, 1007, &atom),
                record(0x10, 2032, &color.to_le_bytes().repeat(8)),
                record(0x60, 2032, &0xaabbccu32.to_le_bytes().repeat(8)),
            ]
            .concat(),
        )
    }
    fn parsed(bytes: &[u8]) -> Record<'_> {
        parse_record_at(bytes, 0, &mut 100).unwrap()
    }

    #[test]
    fn resolves_master_ids_through_persist_offsets_and_respects_local_scheme() {
        let old = slide(1016, 0, false, 0x123456);
        let current = slide(1016, 0, true, 0xab563412);
        let title = slide(1006, 100, true, 0xffffff);
        let document = [&old[..], &current, &title].concat();
        let mut offsets = BTreeMap::new();
        offsets.insert(7, old.len());
        offsets.insert(9, old.len() + current.len());
        let mut refs = Vec::new();
        for (persist, id) in [(7u32, 100u32), (9, 200)] {
            let mut atom = [0; 20];
            atom[..4].copy_from_slice(&persist.to_le_bytes());
            atom[12..16].copy_from_slice(&id.to_le_bytes());
            refs.extend(record(0, 1011, &atom));
        }
        let list = record(0x1f, 4080, &refs);
        let mut r = Resolver::new(&document, &[parsed(&list)], &offsets, &mut 1000).unwrap();
        assert_eq!(
            r.slide(parsed(&slide(1006, 200, true, 0)), &mut 100)
                .unwrap(),
            Some([0x563412; 8])
        );
        assert_eq!(
            r.slide(parsed(&slide(1006, 200, false, 0x987654)), &mut 100)
                .unwrap(),
            Some([0x987654; 8])
        );
        assert_eq!(r.cache.len(), 2);
        let via_title = r
            .text_master(parsed(&slide(1006, 200, false, 0)), &mut 100)
            .unwrap()
            .unwrap();
        let direct = r
            .text_master(parsed(&slide(1006, 100, false, 0)), &mut 100)
            .unwrap()
            .unwrap();
        assert!(std::rc::Rc::ptr_eq(&via_title, &direct));
        // Cached master schemes still charge work on repeated slide references.
        assert!(r.master(200, &mut Vec::new(), &mut 0).is_err());
    }

    #[test]
    fn rejects_cycles_missing_masters_and_excessive_inheritance() {
        let mut r = Resolver {
            masters: BTreeMap::new(),
            cache: BTreeMap::new(),
            text_styles: BTreeMap::new(),
        };
        assert!(r.master(1, &mut Vec::new(), &mut 100).is_err());
        r.masters.insert(
            1,
            Entry {
                parent: Some(2),
                local: Some([0; 8]),
                main_parent: None,
            },
        );
        r.masters.insert(
            2,
            Entry {
                parent: Some(1),
                local: Some([0; 8]),
                main_parent: None,
            },
        );
        assert!(r
            .master(1, &mut Vec::new(), &mut 100)
            .unwrap_err()
            .contains("cyclic"));
        for i in 1..=MAX_DEPTH as u32 + 1 {
            r.masters.insert(
                i,
                Entry {
                    parent: Some(i + 1),
                    local: None,
                    main_parent: None,
                },
            );
        }
        assert!(r
            .master(1, &mut Vec::new(), &mut 1000)
            .unwrap_err()
            .contains("excessive"));
    }

    #[test]
    fn caps_retained_master_styles_before_allocating_another_master() {
        let document = slide(1016, 0, false, 0);
        let mut refs = Vec::new();
        for id in 1..=MAX_MASTERS as u32 + 1 {
            let mut atom = [0; 20];
            atom[..4].copy_from_slice(&1u32.to_le_bytes());
            atom[12..16].copy_from_slice(&id.to_le_bytes());
            refs.extend(record(0, 1011, &atom));
        }
        let list = record(0x1f, 4080, &refs);
        let mut offsets = BTreeMap::new();
        offsets.insert(1, 0);
        let error = Resolver::new(
            &document,
            &[parsed(&list)],
            &offsets,
            &mut MAX_RECORDS.clone(),
        )
        .err()
        .expect("must reject excessive masters");
        assert!(error.contains("too many PowerPoint masters"));
    }

    #[test]
    fn text_master_cycles_and_budgets_are_independent_of_color_inheritance() {
        let mut r = Resolver {
            masters: BTreeMap::new(),
            cache: BTreeMap::new(),
            text_styles: BTreeMap::new(),
        };
        for (id, parent) in [(1, 2), (2, 1)] {
            r.masters.insert(
                id,
                Entry {
                    local: None,
                    parent: None,
                    main_parent: Some(parent),
                },
            );
        }
        let input = slide(1006, 1, false, 0);
        assert!(r
            .text_master(parsed(&input), &mut 100)
            .err()
            .unwrap()
            .contains("cyclic"));
        assert!(r.text_master(parsed(&input), &mut 0).is_err());
        r.masters.remove(&2);
        assert!(r
            .text_master(parsed(&input), &mut 100)
            .err()
            .unwrap()
            .contains("unresolved"));
    }

    #[test]
    fn rejects_duplicate_truncated_schemes_and_does_not_use_available_schemes() {
        let scheme = record(0x10, 2032, &[0; 32]);
        for atoms in [
            record(0x10, 2032, &[0; 31]),
            [scheme.clone(), scheme].concat(),
        ] {
            assert!(entry(parsed(&record(15, 1006, &atoms)), &mut 100).is_err());
        }
        let available = record(15, 1006, &record(0x60, 2032, &[1; 32]));
        assert!(entry(parsed(&available), &mut 100).unwrap().local.is_none());
    }

    #[test]
    fn text_and_drawing_indexes_are_distinct_and_undefined_colors_stay_unresolved() {
        let s = [
            0x010203, 0x112233, 0x223344, 0x334455, 0x445566, 0x556677, 0x667788, 0x778899,
        ];
        for i in 0..8 {
            assert_eq!(text((i << 24) | 0xabcdef, Some(&s)), Some(s[i as usize]));
            assert_eq!(drawing(0x08000000 | i, Some(&s)), Some(s[i as usize]));
        }
        assert_eq!(text(0xfe563412, None), Some(0x563412));
        assert_eq!(drawing(0x04563412, None), Some(0x563412));
        assert_eq!(drawing(0xe0563412, None), Some(0x563412));
        for color in [0x08000008, 0x10000001, 0x01000001, 0xffffffff] {
            assert_eq!(drawing(color, Some(&s)), None);
        }
        assert_eq!(text(0xff123456, Some(&s)), None);
        assert_eq!(text(0x09000000, Some(&s)), None);
        assert_eq!(drawing(0x08000000, None), None);
    }
}
