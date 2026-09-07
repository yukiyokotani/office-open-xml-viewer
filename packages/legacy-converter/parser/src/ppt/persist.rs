//! Resolve live slides through the persist directory, never physical stream order.
//! [MS-PPT] 2.1.2, 2.3.3-2.3.6, 2.4.2, 2.4.14.3-2.4.14.5.
use super::*;
use std::collections::{BTreeMap, HashSet};

pub(super) struct Presentation<'a> {
    pub shape_masters: shape_master::Resolver<'a>,
    pub slides: Vec<(Record<'a>, Vec<String>)>,
    pub outline_styles: Vec<Vec<Option<&'a [u8]>>>,
    pub outline_types: Vec<Vec<u16>>,
    pub outline_slide_numbers: Vec<Vec<Vec<u32>>>,
    pub first_slide_number: u16,
    pub text_masters: Vec<Option<std::rc::Rc<text_style::Master>>>,
    pub fonts: Vec<String>,
    pub schemes: Vec<Option<scheme::Scheme>>,
    pub image_entries: Vec<Record<'a>>,
    pub backgrounds: Vec<Option<paint::Paint>>,
    pub object_masters: Vec<std::rc::Rc<[Record<'a>]>>,
    pub size: (u32, u32),
}

/// Owned edit-chain index, independent of borrowed slide/master views. Direct
/// sessions can retain this once and resolve individual records on demand.
pub(super) struct PersistDirectory {
    pub offsets: BTreeMap<u32, usize>,
    pub document_offset: usize,
}

pub(super) struct SlideListIndex {
    pub slides: Vec<SlideIndexEntry>,
}

pub(super) struct SlideIndexEntry {
    pub record: RecordSpan,
    pub outline: Vec<String>,
    pub outline_styles: Vec<Option<ByteSpan>>,
    pub outline_types: Vec<u16>,
    pub outline_slide_numbers: Vec<Vec<u32>>,
}

fn is_slide_list(record: Record<'_>) -> bool {
    record.kind == 4080 && record.instance == 0 && record.version == 15
}

fn resolve_slide_list(
    document: &[u8],
    children: &[RecordSpan],
    offsets: &BTreeMap<u32, usize>,
    budget: &mut usize,
) -> Result<SlideListIndex, String> {
    let mut lists = Vec::new();
    for span in children {
        let record = span.view(document)?;
        if is_slide_list(record) {
            lists.push(span);
        }
    }
    if lists.len() != 1 {
        return Err(unsupported("missing or duplicate PowerPoint slide list"));
    }
    let items = parse_record_spans(document, lists[0].payload_span(), budget)?;
    let mut slides: Vec<SlideIndexEntry> = Vec::new();
    let mut seen = HashSet::new();
    let mut outline_budget = MAX_TEXT_BYTES;
    for item_span in items {
        let item = item_span.view(document)?;
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
                let (slide, _) = record_span_with_end(document, *offset, budget, "PowerPoint")?;
                let slide_view = slide.view(document)?;
                if slide_view.kind != SLIDE_CONTAINER || slide_view.version != 15 {
                    return Err(unsupported("invalid PowerPoint slide persist object"));
                }
                slides.push(SlideIndexEntry {
                    record: slide,
                    outline: Vec::new(),
                    outline_styles: Vec::new(),
                    outline_types: Vec::new(),
                    outline_slide_numbers: Vec::new(),
                });
            }
            3999 => {
                let slide = slides
                    .last_mut()
                    .ok_or_else(|| unsupported("orphan PowerPoint outline text"))?;
                if slide.outline.len() >= MAX_TEXT_BLOCKS_PER_SLIDE {
                    return Err(unsupported("too many PowerPoint outline text blocks"));
                }
                slide.outline.push(String::new());
                slide.outline_types.push(text_style::text_type(item)?);
                slide.outline_styles.push(None);
                slide.outline_slide_numbers.push(Vec::new());
            }
            TEXT_CHARS_ATOM | TEXT_BYTES_ATOM => {
                let text = slides
                    .last_mut()
                    .and_then(|slide| slide.outline.last_mut())
                    .ok_or_else(|| unsupported("PowerPoint outline text lacks a header"))?;
                let decoded = decode_text(item)?;
                charge_text(&mut outline_budget, decoded.len())?;
                text.push_str(&decoded);
            }
            4001 => {
                let slot = slides
                    .last_mut()
                    .and_then(|slide| slide.outline_styles.last_mut())
                    .ok_or_else(|| unsupported("orphan PowerPoint outline style"))?;
                if item.version != 0 || slot.is_some() {
                    return Err(unsupported("invalid PowerPoint outline style"));
                }
                *slot = Some(item_span.payload_span().clone());
            }
            4056 => slides
                .last_mut()
                .and_then(|slide| slide.outline_slide_numbers.last_mut())
                .ok_or_else(|| unsupported("orphan PowerPoint slide-number atom"))?
                .push(text_style::slide_number_position(item)?),
            _ => {}
        }
    }
    if slides.is_empty() {
        return Err(unsupported("PowerPoint presentation has no slides"));
    }
    Ok(SlideListIndex { slides })
}

pub(super) fn resolve_directory(
    document: &[u8],
    current_edit: usize,
    budget: &mut usize,
) -> Result<PersistDirectory, String> {
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
    let document_offset = *offsets
        .get(&document_id)
        .ok_or_else(|| unsupported("unresolved PowerPoint document persist ID"))?;
    Ok(PersistDirectory {
        offsets,
        document_offset,
    })
}

pub(super) fn resolve<'a>(
    document: &'a [u8],
    current_edit: usize,
    budget: &mut usize,
) -> Result<Presentation<'a>, String> {
    let PersistDirectory {
        offsets,
        document_offset,
    } = resolve_directory(document, current_edit, budget)?;
    let (record_span, _) = record_span_with_end(document, document_offset, budget, "PowerPoint")?;
    let record = record_span.view(document)?;
    if record.kind != DOCUMENT_CONTAINER || record.version != 15 {
        return Err(unsupported("invalid PowerPoint document persist object"));
    }
    let mut children = Vec::new();
    let mut slide_lists = Vec::new();
    // Legacy master consumers still need borrowed children. Retain spans only
    // for slide lists, not a second full vector of document child metadata.
    visit_record_spans(document, record_span.payload_span(), budget, |span| {
        let child = span.view(document)?;
        if is_slide_list(child) {
            slide_lists.push(span);
        }
        children.push(child);
        Ok(())
    })?;
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
    // MS-PPT 2.4.2 DocumentAtom: zero is allowed; the upper bound is exclusive.
    let first_slide_number = u16_at(atom.payload, 32)?;
    if first_slide_number >= 10000 {
        return Err(unsupported("invalid PowerPoint first slide number"));
    }
    let slide_index = resolve_slide_list(document, &slide_lists, &offsets, budget)?;
    let mut slides = Vec::with_capacity(slide_index.slides.len());
    let mut outline_styles = Vec::with_capacity(slide_index.slides.len());
    let mut outline_types = Vec::with_capacity(slide_index.slides.len());
    let mut outline_slide_numbers = Vec::with_capacity(slide_index.slides.len());
    for slide in slide_index.slides {
        slides.push((slide.record.view(document)?, slide.outline));
        outline_styles.push(
            slide
                .outline_styles
                .into_iter()
                .map(|style| style.map(|span| span.view(document)).transpose())
                .collect::<Result<Vec<_>, _>>()?,
        );
        outline_types.push(slide.outline_types);
        outline_slide_numbers.push(slide.outline_slide_numbers);
    }
    Ok(Presentation {
        first_slide_number,
        outline_slide_numbers,
        object_masters: slides
            .iter()
            .map(|(slide, _)| schemes.objects(*slide, budget))
            .collect::<Result<_, _>>()?,
        backgrounds: slides
            .iter()
            .map(|(slide, _)| schemes.background(*slide, budget))
            .collect::<Result<_, _>>()?,
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
        shape_masters: schemes.shape_masters,
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
        fixture_with_styles(false)
    }
    fn fixture_with_styles(styles: bool) -> (Vec<u8>, usize) {
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
            if styles {
                let text = if id == 3 { "second" } else { "first" };
                let count = (text.encode_utf16().count() as u32 + 1).to_le_bytes();
                // One valid, property-free PF run followed by one CF run.
                let style = [
                    count.as_slice(),
                    0u16.to_le_bytes().as_slice(),
                    0u32.to_le_bytes().as_slice(),
                    count.as_slice(),
                    0u32.to_le_bytes().as_slice(),
                ]
                .concat();
                list.extend(record(0, 4001, &style));
            }
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

    fn owned_slide_index(stream: &[u8], edit: usize, budget: &mut usize) -> SlideListIndex {
        let directory = resolve_directory(stream, edit, budget).unwrap();
        let (document, _) =
            record_span_with_end(stream, directory.document_offset, budget, "PowerPoint").unwrap();
        let children = parse_record_spans(stream, document.payload_span(), budget).unwrap();
        resolve_slide_list(stream, &children, &directory.offsets, budget).unwrap()
    }
    #[test]
    fn directory_index_is_owned_and_does_not_parse_slide_or_document_bodies() {
        let index = {
            let (stream, edit) = fixture();
            // One UserEditAtom, one directory record and three persist entries.
            // No budget remains for even one document or slide body record.
            let mut budget = 5;
            let index = resolve_directory(&stream, edit, &mut budget).unwrap();
            assert_eq!(budget, 0);
            assert!(resolve_directory(&stream, edit, &mut 4).is_err());
            index
        };
        // Source lifetime is not part of the retained directory type.
        assert_eq!(index.document_offset, 0);
        assert_eq!(index.offsets.len(), 3);
        assert_eq!(index.offsets[&1], 0);
        assert!(index.offsets[&2] < index.offsets[&3]);
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
    fn slide_list_index_owns_metadata_and_styles_reference_original_backing() {
        let (stream, edit) = fixture_with_styles(true);
        let index = owned_slide_index(&stream, edit, &mut MAX_RECORDS.clone());
        let moved = stream;
        assert_eq!(index.slides.len(), 2);
        assert_eq!(index.slides[0].outline, ["second"]);
        let style = |count: u32| {
            [
                count.to_le_bytes().as_slice(),
                0u16.to_le_bytes().as_slice(),
                0u32.to_le_bytes().as_slice(),
                count.to_le_bytes().as_slice(),
                0u32.to_le_bytes().as_slice(),
            ]
            .concat()
        };
        assert_eq!(
            index.slides[0].outline_styles[0]
                .as_ref()
                .unwrap()
                .view(&moved)
                .unwrap(),
            style(7)
        );
        assert_eq!(
            index.slides[1].outline_styles[0]
                .as_ref()
                .unwrap()
                .view(&moved)
                .unwrap(),
            style(6)
        );
    }

    #[test]
    fn slide_list_children_cannot_escape_into_following_document_record() {
        let mut malformed_child = record(0, 1011, &[0; 2]);
        malformed_child[4..8].copy_from_slice(&4u32.to_le_bytes());
        let list = record(15, 4080, &malformed_child);
        let sibling = record(0, 9, &[8, 8]);
        let document = record(15, DOCUMENT_CONTAINER, &[list, sibling].concat());
        let (document_span, _) = record_span_with_end(&document, 0, &mut 10, "PowerPoint").unwrap();
        let children =
            parse_record_spans(&document, document_span.payload_span(), &mut 10).unwrap();
        let error = resolve_slide_list(&document, &children, &BTreeMap::new(), &mut 10)
            .err()
            .unwrap();
        assert!(error.contains("truncated PowerPoint record"));
    }
    #[test]
    fn document_starting_slide_number_accepts_zero_and_rejects_out_of_range() {
        for first in [0u16, 7, 9999, 10000, u16::MAX] {
            let (mut stream, edit) = fixture();
            // Fixture starts with DocumentContainer, then DocumentAtom.
            stream[48..50].copy_from_slice(&first.to_le_bytes());
            let result = resolve(&stream, edit, &mut MAX_RECORDS.clone());
            if first < 10000 {
                assert_eq!(result.unwrap().first_slide_number, first);
            } else {
                assert!(result.err().unwrap().contains("first slide number"));
            }
        }
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
        let original = resolve_directory(&stream, old_edit, &mut MAX_RECORDS.clone()).unwrap();
        let latest = resolve_directory(&stream, current, &mut MAX_RECORDS.clone()).unwrap();
        assert_eq!(latest.offsets.len(), original.offsets.len());
        assert_eq!(latest.document_offset, original.document_offset);
        assert_eq!(latest.offsets[&2], original.offsets[&2]);
        assert_eq!(latest.offsets[&3], replacement as usize);
        let index = owned_slide_index(&stream, current, &mut MAX_RECORDS.clone());
        let replacement = index.slides[0].record.view(&stream).unwrap();
        let mut indexed_text = Vec::new();
        collect_text(
            replacement.payload,
            0,
            &mut MAX_RECORDS.clone(),
            &mut indexed_text,
            &index.slides[0].outline,
            &mut MAX_TEXT_BYTES.clone(),
        )
        .unwrap();
        assert_eq!(indexed_text, ["updated"]);
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
