use super::*;

fn level(index: u8) -> Vec<u8> {
    let mut bytes = vec![0; 28];
    bytes[0] = 1;
    bytes[6] = 1;
    bytes.extend_from_slice(&[2, 0, index, 0, b'.', 0]);
    bytes
}

fn list(id: i32, simple: bool) -> Vec<u8> {
    let mut bytes = vec![0; 28];
    bytes[..4].copy_from_slice(&id.to_le_bytes());
    for style in bytes[8..26].chunks_exact_mut(2) {
        style.copy_from_slice(&0x0fff_u16.to_le_bytes());
    }
    bytes[26] = u8::from(simple);
    bytes
}

fn lfo(id: i32, count: u8) -> Vec<u8> {
    let mut bytes = vec![0; 16];
    bytes[..4].copy_from_slice(&id.to_le_bytes());
    bytes[12] = count;
    bytes
}

fn tables(
    lists: &[Vec<u8>],
    levels: &[Vec<u8>],
    lfos: &[Vec<u8>],
    data: &[Vec<u8>],
) -> (Vec<u8>, Vec<u8>) {
    let mut word = vec![0; 0x400];
    let mut table = vec![0xee; 7];
    let start = table.len() as u32;
    table.extend_from_slice(&(lists.len() as u16).to_le_bytes());
    table.extend(lists.iter().flatten());
    word[0x2e2..0x2e6].copy_from_slice(&start.to_le_bytes());
    word[0x2e6..0x2ea].copy_from_slice(&(2 + lists.len() as u32 * 28).to_le_bytes());
    table.extend(levels.iter().flatten());
    let start = table.len() as u32;
    table.extend_from_slice(&(lfos.len() as u32).to_le_bytes());
    table.extend(lfos.iter().flatten());
    table.extend(data.iter().flatten());
    word[0x2ea..0x2ee].copy_from_slice(&start.to_le_bytes());
    word[0x2ee..0x2f2].copy_from_slice(&(table.len() as u32 - start).to_le_bytes());
    (word, table)
}

fn one() -> (Vec<u8>, Vec<u8>) {
    tables(
        &[list(42, true)],
        &[level(0)],
        &[lfo(42, 0)],
        &[vec![0xff; 4]],
    )
}

#[test]
fn reads_appended_levels_outside_lcb_plflst_and_keeps_override_identity() {
    let mut levels = vec![level(0)];
    levels.extend((0..9).map(level));
    let (word, table) = tables(
        &[list(42, true), list(-2, false)],
        &levels,
        &[lfo(-2, 0), lfo(42, 0), lfo(-2, 0)],
        &[vec![0xff; 4], vec![0; 4], vec![0xff; 4]],
    );
    let result = Tables::read(&word, &table).unwrap();
    assert_eq!(result.lists.len(), 2);
    assert_eq!(result.lists[0].levels.len(), 1);
    assert!(result.lists[0].simple);
    assert!(!result.lists[1].simple);
    assert_eq!(result.lists[0].styles, [0xfff; 9]);
    assert_eq!(result.lists[1].levels.len(), 9);
    let a = result
        .resolve(Reference::new(1, 8).unwrap().unwrap())
        .unwrap();
    let b = result
        .resolve(Reference::new(3, 8).unwrap().unwrap())
        .unwrap();
    assert_eq!(a.list.id, -2);
    assert!(std::ptr::eq(a.list, b.list));
    assert!(!std::ptr::eq(a.instance, b.instance));
    assert_eq!(a.level.placeholders[0], Some((1, 8)));
    assert_eq!(a.level.text, &[8, 0, b'.', 0]);
    assert_eq!(a.level.start, Some(1));
    assert_eq!(a.level.restart, Some(8));
    assert_eq!(a.instance.first_cp, None);
    assert_eq!(a.instance.auto_number_field, None);
    assert_eq!(a.level.justification, 0);
    assert_eq!(a.level.follow, 0);
}

#[test]
#[ignore = "requires an explicit local-only DOC corpus"]
fn reads_local_office_list_tables() {
    let root = std::path::PathBuf::from(
        std::env::var("OOXML_DOC_LIST_CORPUS").expect("set OOXML_DOC_LIST_CORPUS"),
    );
    let mut pending = vec![root];
    let mut counts = (0, 0, 0, 0);
    while let Some(directory) = pending.pop() {
        for entry in std::fs::read_dir(directory).unwrap() {
            let entry = entry.unwrap();
            let kind = entry.file_type().unwrap();
            let name = entry.file_name();
            let name = name.to_string_lossy();
            if name.starts_with('.') || name.starts_with("~$") {
                continue;
            }
            if kind.is_dir() {
                pending.push(entry.path());
                continue;
            }
            if !kind.is_file() || !name.ends_with(".doc") {
                continue;
            }
            let bytes = std::fs::read(entry.path()).unwrap();
            let compound = crate::cfb::CompoundFile::open(&bytes).unwrap();
            let word = compound.stream("WordDocument").unwrap();
            let table_name = if u16_at(&word, 10).unwrap() & 0x200 != 0 {
                "1Table"
            } else {
                "0Table"
            };
            let table = compound.stream(table_name).unwrap();
            let parsed = Tables::read(&word, &table)
                .unwrap_or_else(|e| panic!("{}: {e}", entry.path().display()));
            counts.0 += 1;
            counts.1 += parsed.lists.len();
            counts.2 += parsed.overrides.len();
            counts.3 += parsed.lists.iter().map(|l| l.levels.len()).sum::<usize>();
        }
    }
    assert!(counts.0 > 0);
    eprintln!(
        "DOC inputs: {}; list definitions: {}; overrides: {}; base levels: {}",
        counts.0, counts.1, counts.2, counts.3
    );
}

#[test]
fn preserves_reference_removal_skip_and_negative_indent_semantics() {
    for value in [0, -2047] {
        // iLvl must be ignored when the paragraph is not in a list.
        assert!(Reference::new(value, 255).unwrap().is_none());
    }
    for value in [1, 2046, -1, -2046] {
        let r = Reference::new(value, 8).unwrap().unwrap();
        assert_eq!(r.index, value.unsigned_abs() as usize - 1);
        assert_eq!(r.preserve_indent, value < 0);
        assert_eq!(r.level, 8);
        assert!(Reference::new(value, 12).unwrap().is_none());
        for invalid in [9, 10, 11, 13, 255] {
            assert!(Reference::new(value, invalid).is_err());
        }
    }
    for value in [2047, i16::MAX, -2048, i16::MIN] {
        assert!(Reference::new(value, 0).is_err());
    }
}

#[test]
fn reads_papx_before_chpx_and_honors_dormant_fields() {
    let mut lvl = level(0);
    lvl[5] = 0x84; // legal, tentative; restart limit is dormant.
    lvl[26] = 255;
    lvl[24] = 3;
    lvl[25] = 4;
    lvl.splice(28..28, [0x0f, 0x84, 0xd0, 0x02, 0x35, 0x08, 1]);
    let (word, table) = tables(&[list(42, true)], &[lvl], &[lfo(42, 0)], &[vec![0; 4]]);
    let parsed = Tables::read(&word, &table).unwrap();
    let lvl = &parsed.lists[0].levels[0];
    assert_eq!(lvl.papx, &[0x0f, 0x84, 0xd0, 0x02]);
    assert_eq!(lvl.chpx, &[0x35, 0x08, 1]);
    assert!(lvl.legal);
    assert!(!lvl.tentative); // fTentative is ignored outside hybrid lists.
    assert_eq!(lvl.restart, Some(0));
    assert_eq!(parsed.overrides[0].first_cp, Some(0));
}

#[test]
fn distinguishes_start_override_from_complete_format_override() {
    for flags in [0, 0x10, 0x20, 0x30] {
        let mut data = vec![0xff; 4];
        data.extend_from_slice(&if flags == 0x10 { 7_i32 } else { -1 }.to_le_bytes());
        data.extend_from_slice(&[flags, 0xee, 0xdd, 0xcc]);
        if flags & 0x20 != 0 {
            let mut lvl = level(0);
            lvl[0] = 11;
            lvl[4] = 1;
            data.extend(lvl);
        }
        let (word, table) = tables(&[list(42, true)], &[level(0)], &[lfo(42, 1)], &[data]);
        let parsed = Tables::read(&word, &table).unwrap();
        let selected = parsed
            .resolve(Reference::new(1, 0).unwrap().unwrap())
            .unwrap();
        assert_eq!(selected.level.format, if flags & 0x20 != 0 { 1 } else { 0 });
        assert_eq!(
            selected.start_override,
            match flags {
                0x10 => Some(7),
                0x30 => Some(11),
                _ => None,
            }
        );
        assert_eq!(
            selected.level.start,
            if flags & 0x20 != 0 { Some(11) } else { Some(1) }
        );
    }
}

#[test]
fn honors_numberless_levels_and_zero_terminated_placeholder_array() {
    for format in [0x17, 0xff] {
        let mut lvl = level(0);
        lvl[..4].copy_from_slice(&(-1_i32).to_le_bytes());
        lvl[4] = format;
        lvl[5] = 8;
        lvl[6] = 0;
        lvl[7..15].fill(255); // Must be ignored after the first zero.
        lvl[26] = 255;
        lvl.truncate(28);
        lvl.extend_from_slice(&[1, 0, 0xb7, 0xf0]);
        let (word, table) = tables(&[list(42, true)], &[lvl], &[lfo(42, 0)], &[vec![0; 4]]);
        let parsed = Tables::read(&word, &table).unwrap();
        let lvl = &parsed.lists[0].levels[0];
        assert_eq!(lvl.start, None);
        assert_eq!(lvl.restart, None);
        assert_eq!(lvl.text, &[0xb7, 0xf0]); // Preserve raw glyph, do not guess a Unicode bullet.
        assert_eq!(lvl.placeholders, [None; 9]);
    }
}

#[test]
fn rejects_invalid_level_metadata() {
    for (offset, value) in [
        (0, 0xff),
        (4, 8),
        (4, 9),
        (4, 15),
        (4, 19),
        (4, 60),
        (5, 3),
        (6, 3),
        (15, 3),
        (30, 1),
    ] {
        let mut lvl = level(0);
        lvl[offset] = value;
        if offset == 0 {
            lvl[3] = 0xff;
        }
        let (word, table) = tables(&[list(42, true)], &[lvl], &[lfo(42, 0)], &[vec![0; 4]]);
        assert!(
            Tables::read(&word, &table).is_err(),
            "offset {offset}, value {value}"
        );
    }
    for patch in [[5, 8, 26, 1], [4, 0x17, 6, 0], [6, 1, 7, 1]] {
        let mut lvl = level(0);
        lvl[patch[0] as usize] = patch[1];
        lvl[patch[2] as usize] = patch[3];
        let (word, table) = tables(&[list(42, true)], &[lvl], &[lfo(42, 0)], &[vec![0; 4]]);
        assert!(Tables::read(&word, &table).is_err());
    }
}

#[test]
fn bounds_every_prefix_and_rejects_ambiguous_or_missing_references() {
    let (word, table) = one();
    for size in 0..table.len() {
        assert!(
            Tables::read(&word, &table[..size]).is_err(),
            "prefix {size}"
        );
    }
    let parsed = Tables::read(&word, &table).unwrap();
    assert!(parsed
        .resolve(Reference::new(2, 0).unwrap().unwrap())
        .is_err());
    assert!(parsed
        .resolve(Reference::new(1, 1).unwrap().unwrap())
        .is_err());
    for lists in [vec![list(42, true), list(42, true)], vec![list(-1, true)]] {
        let (word, table) = tables(&lists, &[level(0), level(0)], &[lfo(42, 0)], &[vec![0; 4]]);
        assert!(Tables::read(&word, &table).is_err());
    }
    let (word, table) = tables(&[list(42, true)], &[level(0)], &[lfo(99, 0)], &[vec![0; 4]]);
    assert!(Tables::read(&word, &table).is_err());
}

#[test]
fn empty_ranges_ignore_offsets_and_counts_cannot_drive_unbounded_allocation() {
    let mut word = vec![0; 0x400];
    word[0x2e2..0x2e6].fill(255);
    word[0x2ea..0x2ee].fill(255);
    let parsed = Tables::read(&word, &[]).unwrap();
    assert!(parsed.lists.is_empty() && parsed.overrides.is_empty());
    let (mut word, mut table) = one();
    table[7..9].copy_from_slice(&0xffff_u16.to_le_bytes());
    assert!(Tables::read(&word, &table).is_err());
    word[0x2e6..0x2ea].fill(0);
    let pos = u32::from_le_bytes(word[0x2ea..0x2ee].try_into().unwrap()) as usize;
    table[pos..pos + 4].fill(255);
    assert!(Tables::read(&word, &table).is_err());
    let mut budget = Budget {
        levels: 0,
        bytes: usize::MAX,
    };
    assert!(read_level(&mut Reader::new(&level(0)), 0, false, &mut budget).is_err());
    let mut budget = Budget {
        levels: 1,
        bytes: 29,
    };
    assert!(read_level(&mut Reader::new(&level(0)), 0, false, &mut budget).is_err());
}

#[test]
fn handles_all_nine_placeholders_restart_boundaries_and_hybrid_metadata() {
    for restart in 0..=8 {
        let mut levels: Vec<_> = (0..9).map(level).collect();
        let last = &mut levels[8];
        last[5] = 0x8e; // right, legal, noRestart, tentative
        last[6..15].copy_from_slice(&[1, 2, 3, 4, 5, 6, 7, 8, 9]);
        last[15] = 2;
        last[26] = restart;
        last.truncate(28);
        last.extend_from_slice(&9_u16.to_le_bytes());
        for index in 0..9_u16 {
            last.extend_from_slice(&index.to_le_bytes());
        }
        let mut lst = list(42, false);
        lst[26] = 0x10;
        lst[8..10].copy_from_slice(&1_u16.to_le_bytes());
        let (word, table) = tables(&[lst], &levels, &[lfo(42, 0)], &[vec![0; 4]]);
        let parsed = Tables::read(&word, &table).unwrap();
        let lvl = &parsed.lists[0].levels[8];
        assert!(lvl.tentative && lvl.legal);
        assert_eq!(lvl.justification, 2);
        assert_eq!(lvl.follow, 2);
        assert_eq!(lvl.restart, Some(restart));
        assert_eq!(lvl.placeholders[8], Some((9, 8)));
        assert_eq!(parsed.lists[0].styles[0], 1);
    }
}

#[test]
fn rejects_ambiguous_overrides_and_keeps_empty_override_storage_small() {
    for indices in [[0, 0], [0, 9], [0, 15]] {
        let mut data = vec![0; 4];
        for index in indices {
            data.extend_from_slice(&[0, 0, 0, 0, index, 0, 0, 0]);
        }
        let (word, table) = tables(&[list(42, true)], &[level(0)], &[lfo(42, 2)], &[data]);
        assert!(Tables::read(&word, &table).is_err());
    }
    let lfos = vec![lfo(42, 0); 2046];
    let data = vec![vec![0xff; 4]; 2046];
    let (word, table) = tables(&[list(42, true)], &[level(0)], &lfos, &data);
    let parsed = Tables::read(&word, &table).unwrap();
    assert_eq!(parsed.overrides.len(), 2046);
    assert!(parsed.overrides.iter().all(|o| o.levels.capacity() == 0));
    let selected = parsed
        .resolve(Reference::new(2046, 0).unwrap().unwrap())
        .unwrap();
    assert_eq!(selected.list.id, 42);
    assert!(parsed
        .resolve(Reference {
            index: 0,
            level: 255,
            preserve_indent: false
        })
        .is_err());
}

#[test]
fn keeps_autonum_field_lists_distinct_from_ordinary_paragraph_numbering() {
    for field in [0xfc, 0xfd, 0xfe] {
        let mut lst = list(42, true);
        lst[26] |= 4;
        let mut instance = lfo(42, 0);
        instance[13] = field;
        let (word, table) = tables(&[lst], &[level(0)], &[instance], &[vec![0xff; 4]]);
        let parsed = Tables::read(&word, &table).unwrap();
        assert!(parsed.lists[0].auto_number);
        assert_eq!(parsed.overrides[0].auto_number_field, Some(field));
    }
    for field in [1, 0xfb, 0xfc] {
        let mut instance = lfo(42, 0);
        instance[13] = field;
        let (word, table) = tables(
            &[list(42, true)],
            &[level(0)],
            &[instance],
            &[vec![0xff; 4]],
        );
        assert!(Tables::read(&word, &table).is_err());
    }
}

#[test]
fn bounds_nested_override_payloads_and_rejects_trailing_lfo_data() {
    let mut data = vec![0xff; 4];
    data.extend_from_slice(&[0xff, 0xff, 0xff, 0xff, 0x30, 0, 0, 0]);
    data.extend(level(0));
    let (mut word, mut table) = tables(&[list(42, true)], &[level(0)], &[lfo(42, 1)], &[data]);
    let pos = u32_at(&word, FC_PLF_LFO).unwrap() as usize;
    for end in pos..table.len() {
        word[FC_PLF_LFO + 4..FC_PLF_LFO + 8].copy_from_slice(&((end - pos) as u32).to_le_bytes());
        if end == pos {
            continue;
        } // An empty declared table is absent.
        assert!(Tables::read(&word, &table[..end]).is_err());
    }
    table.push(0);
    word[FC_PLF_LFO + 4..FC_PLF_LFO + 8]
        .copy_from_slice(&((table.len() - pos) as u32).to_le_bytes());
    assert!(Tables::read(&word, &table).is_err());
}

#[test]
fn checks_each_level_prefix_and_charges_the_exact_borrowed_payload() {
    let bytes = level(0);
    for end in 0..bytes.len() {
        let mut budget = Budget {
            levels: 1,
            bytes: bytes.len(),
        };
        assert!(read_level(&mut Reader::new(&bytes[..end]), 0, false, &mut budget).is_err());
    }
    let mut budget = Budget {
        levels: 1,
        bytes: bytes.len(),
    };
    let mut reader = Reader::new(&bytes);
    let parsed = read_level(&mut reader, 0, false, &mut budget).unwrap();
    assert!(reader.bytes.is_empty());
    assert_eq!(budget.levels, 0);
    assert_eq!(budget.bytes, 0);
    assert!(std::ptr::eq(parsed.text.as_ptr(), bytes[30..].as_ptr()));
    // Even empty lists must not reinterpret a count/header length mismatch.
    let (mut word, table) = one();
    word[FC_PLF_LST + 4..FC_PLF_LST + 8].copy_from_slice(&31_u32.to_le_bytes());
    assert!(Tables::read(&word, &table).is_err());
}
