//! Resolve live slides through the persist directory, never physical stream order.
//! [MS-PPT] 2.1.2, 2.3.3-2.3.6, 2.4.2, 2.4.14.3-2.4.14.5.
use super::*;
use std::collections::{BTreeMap, HashSet};

pub(super) struct Presentation<'a> {
    pub slides: Vec<(Record<'a>, Vec<String>)>,
    pub outline_styles: Vec<Vec<Option<&'a [u8]>>>,
    pub outline_types: Vec<Vec<u16>>,
    pub text_masters: Vec<Option<std::rc::Rc<text_style::Master>>>,
    pub fonts: Vec<String>,
    pub schemes: Vec<Option<scheme::Scheme>>,
    pub image_entries: Vec<Record<'a>>,
    pub size: (u32, u32),
}

pub(super) fn resolve<'a>(
    document: &'a [u8],
    current_edit: usize,
    budget: &mut usize,
) -> Result<Presentation<'a>, String> {
    let mut offsets = BTreeMap::new();
    let mut edit_offset = current_edit;
    let mut document_id = None;
    loop {
        let edit = parse_record_at(document, edit_offset, budget)?;
        if edit.kind != USER_EDIT_ATOM || edit.version != 0 || edit.payload.len() < 28 {
            return Err(unsupported("invalid PowerPoint UserEditAtom"));
        }
        if edit.payload.len() != 28 {
            return Err(unsupported(
                "encrypted or unsupported PowerPoint UserEditAtom",
            ));
        }
        let previous = u32_at(edit.payload, 8)? as usize;
        let directory_offset = u32_at(edit.payload, 12)? as usize;
        document_id.get_or_insert(u32_at(edit.payload, 16)?);
        if (previous != 0 && previous >= edit_offset)
            || directory_offset <= previous
            || directory_offset >= edit_offset
        {
            return Err(unsupported("invalid PowerPoint edit/directory offsets"));
        }
        let directory = parse_record_at(document, directory_offset, budget)?;
        if directory.kind != 0x1772 || directory.version != 0 {
            return Err(unsupported("invalid PowerPoint persist directory"));
        }
        let mut position = 0;
        let mut seen = HashSet::new();
        while position < directory.payload.len() {
            let head = u32_at(directory.payload, position)?;
            let first = head & 0xfffff;
            let count = head >> 20;
            if first == 0 || count == 0 || first + count > 0xfffff {
                return Err(unsupported("invalid PowerPoint persist ID range"));
            }
            position += 4;
            *budget = budget
                .checked_sub(count as usize)
                .ok_or_else(|| unsupported("PowerPoint persist-entry budget exceeded"))?;
            for id in first..first + count {
                if !seen.insert(id) {
                    return Err(unsupported("duplicate PowerPoint persist ID in one edit"));
                }
                let offset = u32_at(directory.payload, position)? as usize;
                if offset < previous || offset >= directory_offset {
                    return Err(unsupported("invalid PowerPoint persist object offset"));
                }
                // Newest edit is visited first; older records cannot replace it.
                offsets.entry(id).or_insert(offset);
                position += 4;
            }
        }
        if previous == 0 {
            break;
        }
        edit_offset = previous;
    }
    let document_id =
        document_id.ok_or_else(|| unsupported("missing PowerPoint current document"))?;
    let offset = offsets
        .get(&document_id)
        .ok_or_else(|| unsupported("unresolved PowerPoint document persist ID"))?;
    let record = parse_record_at(document, *offset, budget)?;
    if record.kind != DOCUMENT_CONTAINER || record.version != 15 {
        return Err(unsupported("invalid PowerPoint document persist object"));
    }
    let children = parse_records(record.payload, budget)?;
    let mut schemes = scheme::Resolver::new(document, &children, &offsets, budget)?;
    let fonts = text_style::fonts(&children, budget)?;
    let atom = children
        .iter()
        .find(|r| r.kind == 1001)
        .ok_or_else(|| unsupported("missing PowerPoint DocumentAtom"))?;
    if atom.payload.len() != 40 || atom.version != 1 {
        return Err(unsupported("invalid PowerPoint DocumentAtom"));
    }
    let size = (
        dimension(u32_at(atom.payload, 0)?)?,
        dimension(u32_at(atom.payload, 4)?)?,
    );
    let lists: Vec<_> = children
        .iter()
        .filter(|r| r.kind == 4080 && r.instance == 0 && r.version == 15)
        .collect();
    if lists.len() != 1 {
        return Err(unsupported("missing or duplicate PowerPoint slide list"));
    }
    let mut slides: Vec<(Record<'a>, Vec<String>)> = Vec::new();
    let mut outline_styles: Vec<Vec<Option<&'a [u8]>>> = Vec::new();
    let mut outline_types: Vec<Vec<u16>> = Vec::new();
    let mut seen = HashSet::new();
    let mut outline_budget = MAX_TEXT_BYTES;
    for item in parse_records(lists[0].payload, budget)? {
        match item.kind {
            1011 => {
                if item.payload.len() != 20 {
                    return Err(unsupported("invalid PowerPoint SlidePersistAtom"));
                }
                let id = u32_at(item.payload, 0)?;
                if !seen.insert(id) || slides.len() >= MAX_SLIDES {
                    return Err(unsupported(
                        "duplicate or excessive PowerPoint slide references",
                    ));
                }
                let offset = offsets
                    .get(&id)
                    .ok_or_else(|| unsupported("unresolved PowerPoint slide persist ID"))?;
                let slide = parse_record_at(document, *offset, budget)?;
                if slide.kind != SLIDE_CONTAINER || slide.version != 15 {
                    return Err(unsupported("invalid PowerPoint slide persist object"));
                }
                slides.push((slide, Vec::new()));
                outline_styles.push(Vec::new());
                outline_types.push(Vec::new());
            }
            3999 => {
                let (_, outline) = slides
                    .last_mut()
                    .ok_or_else(|| unsupported("orphan PowerPoint outline text"))?;
                if outline.len() >= MAX_TEXT_BLOCKS_PER_SLIDE {
                    return Err(unsupported("too many PowerPoint outline text blocks"));
                }
                outline.push(String::new());
                outline_types
                    .last_mut()
                    .expect("slide exists")
                    .push(text_style::text_type(item)?);
                outline_styles.last_mut().expect("slide exists").push(None);
            }
            TEXT_CHARS_ATOM | TEXT_BYTES_ATOM => {
                let text = slides
                    .last_mut()
                    .and_then(|(_, outline)| outline.last_mut())
                    .ok_or_else(|| unsupported("PowerPoint outline text lacks a header"))?;
                let decoded = decode_text(item)?;
                charge_text(&mut outline_budget, decoded.len())?;
                text.push_str(&decoded);
            }
            4001 => {
                let slot = outline_styles
                    .last_mut()
                    .and_then(|styles| styles.last_mut())
                    .ok_or_else(|| unsupported("orphan PowerPoint outline style"))?;
                if item.version != 0 || slot.is_some() {
                    return Err(unsupported("invalid PowerPoint outline style"));
                }
                *slot = Some(item.payload);
            }
            _ => {}
        }
    }
    if slides.is_empty() {
        return Err(unsupported("PowerPoint presentation has no slides"));
    }
    Ok(Presentation {
        image_entries: media::catalog(&children, budget)?,
        text_masters: slides
            .iter()
            .map(|(slide, _)| schemes.text_master(*slide, budget))
            .collect::<Result<_, _>>()?,
        outline_types,
        schemes: slides
            .iter()
            .map(|(slide, _)| schemes.slide(*slide, budget))
            .collect::<Result<_, _>>()?,
        slides,
        outline_styles,
        fonts,
        size,
    })
}

fn dimension(master_units: u32) -> Result<u32, String> {
    if !(576..=32256).contains(&master_units) {
        return Err(unsupported("invalid PowerPoint slide dimensions"));
    }
    // One master unit = 1/576 inch, one inch = 914400 EMU. Round
    // half-EMU dimensions to the nearest integer required by PresentationML.
    Ok(((u64::from(master_units) * 914400 + 288) / 576) as u32)
}

#[cfg(test)]
pub(crate) mod tests {
    use super::*;
    pub(crate) fn record(options: u16, kind: u16, bytes: &[u8]) -> Vec<u8> {
        [
            options.to_le_bytes().as_slice(),
            &kind.to_le_bytes(),
            &(bytes.len() as u32).to_le_bytes(),
            bytes,
        ]
        .concat()
    }
    pub(crate) fn fixture() -> (Vec<u8>, usize) {
        let mut atom = vec![0u8; 40];
        atom[..4].copy_from_slice(&7680u32.to_le_bytes());
        atom[4..8].copy_from_slice(&4320u32.to_le_bytes());
        let mut list = Vec::new();
        for id in [3u32, 2] {
            let mut persist = [0u8; 20];
            persist[..4].copy_from_slice(&id.to_le_bytes());
            list.extend(record(0, 1011, &persist));
            list.extend(record(0, 3999, &[0; 4]));
            list.extend(record(
                0,
                TEXT_BYTES_ATOM,
                if id == 3 { b"second" } else { b"first" },
            ));
        }
        let mut stream = record(
            15,
            DOCUMENT_CONTAINER,
            &[record(1, 1001, &atom), record(15, 4080, &list)].concat(),
        );
        let first = stream.len() as u32;
        stream.extend(record(15, SLIDE_CONTAINER, &record(0, 3998, &[0; 4])));
        let second = stream.len() as u32;
        stream.extend(record(15, SLIDE_CONTAINER, &record(0, 3998, &[0; 4])));
        // Dead slide must not be emitted.
        stream.extend(record(
            15,
            SLIDE_CONTAINER,
            &record(0, TEXT_BYTES_ATOM, b"deleted"),
        ));
        let directory = stream.len() as u32;
        stream.extend(record(
            0,
            0x1772,
            &[
                0x00300001u32.to_le_bytes(),
                0u32.to_le_bytes(),
                first.to_le_bytes(),
                second.to_le_bytes(),
            ]
            .concat(),
        ));
        let edit = stream.len();
        let mut user = vec![0u8; 28];
        user[12..16].copy_from_slice(&directory.to_le_bytes());
        user[16..20].copy_from_slice(&1u32.to_le_bytes());
        stream.extend(record(0, USER_EDIT_ATOM, &user));
        (stream, edit)
    }
    #[test]
    fn resolves_order_outline_text_and_size_without_deleted_slides() {
        let (stream, edit) = fixture();
        let mut budget = MAX_RECORDS;
        let result = resolve(&stream, edit, &mut budget).unwrap();
        assert_eq!(result.size, (12192000, 6858000));
        assert_eq!(result.slides.len(), 2);
        assert_eq!(result.slides[0].1, ["second"]);
        assert_eq!(result.slides[1].1, ["first"]);
    }
    #[test]
    fn latest_edit_replaces_only_the_persist_objects_it_updates() {
        let (mut stream, old_edit) = fixture();
        let replacement = stream.len() as u32;
        stream.extend(record(
            15,
            SLIDE_CONTAINER,
            &record(0, TEXT_BYTES_ATOM, b"updated"),
        ));
        let directory = stream.len() as u32;
        stream.extend(record(
            0,
            0x1772,
            &[0x00100003u32.to_le_bytes(), replacement.to_le_bytes()].concat(),
        ));
        let current = stream.len();
        let mut user = vec![0u8; 28];
        user[8..12].copy_from_slice(&(old_edit as u32).to_le_bytes());
        user[12..16].copy_from_slice(&directory.to_le_bytes());
        user[16..20].copy_from_slice(&1u32.to_le_bytes());
        stream.extend(record(0, USER_EDIT_ATOM, &user));
        let mut budget = MAX_RECORDS;
        let result = resolve(&stream, current, &mut budget).unwrap();
        assert_eq!(result.slides.len(), 2);
        let mut texts = Vec::new();
        collect_text(
            result.slides[0].0.payload,
            0,
            &mut budget,
            &mut texts,
            &result.slides[0].1,
            &mut MAX_TEXT_BYTES.clone(),
        )
        .unwrap();
        assert_eq!(texts, ["updated"]);
    }
    #[test]
    fn rejects_self_referential_edit_chain() {
        let (mut stream, edit) = fixture();
        stream[edit + 16..edit + 20].copy_from_slice(&(edit as u32).to_le_bytes());
        assert!(resolve(&stream, edit, &mut MAX_RECORDS.clone()).is_err());
    }
}
