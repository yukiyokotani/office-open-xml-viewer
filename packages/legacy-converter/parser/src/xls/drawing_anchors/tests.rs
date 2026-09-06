use super::*;

fn art(kind: u16, options: u16, data: &[u8]) -> Vec<u8> {
    [
        options.to_le_bytes().to_vec(),
        kind.to_le_bytes().to_vec(),
        (data.len() as u32).to_le_bytes().to_vec(),
        data.to_vec(),
    ]
    .concat()
}
fn records(data: &[(u16, Vec<u8>)]) -> Vec<Record<'_>> {
    let mut offset = 0;
    data.iter()
        .map(|(kind, data)| {
            let r = Record {
                kind: *kind,
                offset,
                data,
            };
            offset += 4 + data.len();
            r
        })
        .collect()
}
fn flags(id: u32) -> Vec<u8> {
    art(
        0xf00a,
        2,
        &[id.to_le_bytes(), 0x200u32.to_le_bytes()].concat(),
    )
}
fn cmo(id: u16) -> Vec<u8> {
    let mut data = vec![0; 22];
    data[..10].copy_from_slice(&[0x15, 0, 0x12, 0, 8, 0, id as u8, (id >> 8) as u8, 0x11, 0]);
    data
}
fn anchor() -> Vec<u8> {
    [
        3u16,
        2,
        (-17i16) as u16,
        3,
        255,
        256,
        1025,
        65535,
        (-257i16) as u16,
    ]
    .into_iter()
    .flat_map(u16::to_le_bytes)
    .collect()
}
fn fixture() -> Vec<(u16, Vec<u8>)> {
    let before = [flags(101), art(0xf010, 0, &anchor()), art(0xf011, 0, &[])].concat();
    let body = [before.clone(), art(0xf00d, 0, &[])].concat();
    let whole = art(0xf002, 15, &art(0xf004, 15, &body));
    let cut = 16 + before.len();
    vec![
        (BOF, vec![0, 6, 0x10, 0]),
        (0xec, whole[..cut].to_vec()),
        (0x5d, cmo(7)),
        (0xec, whole[cut..].to_vec()),
        (0x1b6, vec![0; 18]),
        (0x3c, vec![0xff; 8]),
        (EOF, vec![]),
    ]
}
fn run(data: &[(u16, Vec<u8>)]) -> Result<Vec<DrawingAnchor>, String> {
    let mut work = 1000;
    let mut remaining = MAX_BYTES;
    let mut drawing = assemble(&records(data), &mut work, &mut remaining)?.unwrap();
    let mut result = Vec::new();
    walk(&mut drawing, 4, &mut work, &mut result)?;
    Ok(result)
}

#[test]
fn preserves_signed_fractional_endpoints_and_exact_object_ownership() {
    let result = run(&fixture()).unwrap();
    assert_eq!(
        result,
        vec![DrawingAnchor {
            sheet: 4,
            shape_id: 101,
            shape_flags: 0x200,
            object_id: 7,
            object_type: 8,
            object_flags: 0x11,
            group_depth: 0,
            behavior: 3,
            picture: None,
            from: CellCorner {
                column: 2,
                row: 3,
                dx: -17,
                dy: 255
            },
            to: CellCorner {
                column: 256,
                row: 65535,
                dx: 1025,
                dy: -257
            }
        }]
    );
}

#[test]
fn only_drawing_continuations_are_joined_at_every_byte_boundary() {
    let source = fixture();
    let expected = run(&source).unwrap();
    for cut in 1..source[1].1.len() {
        for kind in [0x3c, 0xec] {
            let mut data = source.clone();
            let remainder = data[1].1.split_off(cut);
            data.insert(2, (kind, remainder));
            assert_eq!(run(&data).unwrap(), expected);
        }
    }
    // The private bytes of a following Obj/TxO are not part of recLen.
    let mut data = source.clone();
    data.insert(3, (0x3c, vec![0xaa; 100]));
    assert_eq!(run(&data).unwrap(), expected);
    // Embedded chart drawing streams have their own BOF/EOF owner.
    data.splice(
        3..3,
        [
            (BOF, vec![0, 6, 0x20, 0]),
            (0xec, vec![0xff; 40]),
            (EOF, vec![]),
        ],
    );
    assert_eq!(run(&data).unwrap(), expected);
}

#[test]
fn refuses_nearby_decoy_clients_duplicates_and_missing_owners() {
    for index in [2, 4] {
        let mut data = fixture();
        data.insert(index, (0x1234, vec![]));
        assert!(run(&data).is_err());
    }
    let mut data = fixture();
    data[2].0 = 0x1b6;
    assert!(run(&data).is_err());
    for length in 0..22 {
        let mut data = fixture();
        data[2].1.truncate(length);
        assert!(run(&data).is_err());
    }
    for (offset, value) in [(0, 0), (2, 0), (4, 10), (4, 255)] {
        let mut data = fixture();
        data[2].1[offset] = value;
        assert!(run(&data).is_err());
    }
    let mut data = fixture();
    data[1].1.extend_from_slice(&[0, 0, 0, 0]);
    assert!(run(&data).is_err());
}

#[test]
fn bounds_fragments_substreams_and_retained_work() {
    let data = fixture();
    let records = records(&data);
    let mut remaining = MAX_BYTES;
    assert!(assemble(&records, &mut 0, &mut remaining).is_err());
    assert!(assemble(&records, &mut 100, &mut 0).is_err());
    assert!(assemble(&records[..records.len() - 1], &mut 100, &mut remaining).is_err());
    let mut overlong = data.clone();
    overlong[1].1.resize(8225, 0);
    assert!(run(&overlong).is_err());
    let mut nested = data.clone();
    nested.splice(1..1, (0..=MAX_DEPTH).map(|_| (BOF, vec![0, 6, 0x20, 0])));
    assert!(run(&nested).is_err());
    let mut drawing = assemble(&records, &mut 100, &mut remaining)
        .unwrap()
        .unwrap();
    assert!(walk(&mut drawing, 0, &mut 1, &mut Vec::new()).is_err());
}

#[test]
fn traverses_only_owned_groups_in_order_and_bounds_the_stack() {
    let source = fixture();
    let original: Vec<u8> = [source[1].1.as_slice(), source[3].1.as_slice()].concat();
    let shape = &original[8..];
    let before_client_len = source[1].1.len() - 8;
    // Group headers have identities but need no worksheet anchor or Obj.
    let head = |id| art(0xf004, 15, &flags(id));
    let inner = art(0xf003, 15, &[head(2), shape.to_vec()].concat());
    let outer = art(0xf003, 15, &[head(1), inner, head(3)].concat());
    let bytes = art(0xf002, 15, &outer);
    let cut = 8 + 8 + head(1).len() + 8 + head(2).len() + before_client_len;
    let mut data = source.clone();
    data[1].1 = bytes[..cut].to_vec();
    // The client textbox marker ends before the later sibling shape, so give
    // that sibling its own fragment after the TxO/Continue pair.
    data[3].1 = bytes[cut..cut + 8].to_vec();
    data.insert(6, (0xec, bytes[cut + 8..].to_vec()));
    assert_eq!(run(&data).unwrap()[0].group_depth, 2);

    let mut body = head(11);
    for id in 12..(MAX_DEPTH as u32 + 14) {
        body = art(0xf003, 15, &[head(id), body].concat());
    }
    assert!(run(&[
        (BOF, vec![0, 6, 0x10, 0]),
        (0xec, art(0xf002, 15, &body)),
        (EOF, vec![])
    ])
    .is_err());
}

#[test]
fn ignores_opaque_container_decoys_and_rejects_cross_sheet_ranges() {
    let mut data = fixture();
    let decoy = art(
        0x7777,
        15,
        &art(0xf004, 15, &[flags(101), art(0xf011, 0, &[])].concat()),
    );
    let length = u32::from_le_bytes(data[1].1[4..8].try_into().unwrap()) + decoy.len() as u32;
    data[1].1[4..8].copy_from_slice(&length.to_le_bytes());
    data[1].1.splice(8..8, decoy);
    assert_eq!(run(&data).unwrap(), run(&fixture()).unwrap());
    let mut duplicate = fixture();
    // Two identity records in the same owner must not overwrite each other.
    let extra = flags(222);
    for at in [4, 12] {
        let n =
            u32::from_le_bytes(duplicate[1].1[at..at + 4].try_into().unwrap()) + extra.len() as u32;
        duplicate[1].1[at..at + 4].copy_from_slice(&n.to_le_bytes());
    }
    duplicate[1].1.splice(16..16, extra);
    assert!(run(&duplicate).is_err());

    let mut missing = fixture();
    missing.pop();
    missing.extend(fixture());
    assert!(run(&missing).is_err());
}

#[test]
fn validates_anchor_schema_but_ignores_reserved_bits_without_clamping() {
    let payload = anchor();
    for length in 0..payload.len() {
        assert!(corner(&payload[..length], 10).is_err());
    }
    let mut bad = payload.clone();
    bad[2..4].copy_from_slice(&257u16.to_le_bytes());
    assert!(corner(&bad, 2).is_err());
    let mut data = fixture();
    // Dg + Sp + FSP + anchor header = 16 + 16 + 8.
    data[1].1[40..42].copy_from_slice(&0xfff3u16.to_le_bytes());
    assert_eq!(run(&data).unwrap()[0].behavior, 3);
    data[1].1[40..42].copy_from_slice(&1u16.to_le_bytes());
    assert!(run(&data).is_err());
}

#[test]
fn tab_order_comes_from_boundsheet_not_physical_stream_order() {
    let mut data = vec![
        (BOF, vec![0, 6, 5, 0]),
        (BOUNDSHEET8, vec![0, 0, 0, 0, 0, 0, 1, 0, b'B']),
        (BOUNDSHEET8, vec![0, 0, 0, 0, 0, 0, 1, 0, b'A']),
        (EOF, vec![]),
    ];
    let start_a = data.iter().map(|(_, d)| d.len() + 4).sum::<usize>();
    data.extend(fixture());
    let start_b = data.iter().map(|(_, d)| d.len() + 4).sum::<usize>();
    data.extend(fixture());
    data[1].1[..4].copy_from_slice(&(start_b as u32).to_le_bytes());
    data[2].1[..4].copy_from_slice(&(start_a as u32).to_le_bytes());
    let result = workbook(&records(&data)).unwrap();
    assert_eq!(
        result.iter().map(|r| r.sheet).collect::<Vec<_>>(),
        vec![1, 0]
    );
    data[2].1[..4].copy_from_slice(&(start_b as u32).to_le_bytes());
    assert!(workbook(&records(&data)).is_err());
}

#[test]
fn enforces_shape_identity_and_global_retained_anchor_limits() {
    let mut body = Vec::new();
    for id in 0..=MAX_OBJECTS as u32 {
        body.extend_from_slice(&art(0xf004, 15, &flags(id)));
    }
    let mut drawing = Drawing {
        bytes: art(0xf002, 15, &body),
        clients: BTreeMap::new(),
    };
    assert!(walk(&mut drawing, 0, &mut 2_000_000, &mut Vec::new())
        .unwrap_err()
        .contains("excessive BIFF shapes"));

    let source = fixture();
    let mut remaining = MAX_BYTES;
    let mut drawing = assemble(&records(&source), &mut 1000, &mut remaining)
        .unwrap()
        .unwrap();
    let mut retained = vec![run(&source).unwrap()[0]; MAX_OBJECTS];
    assert!(walk(&mut drawing, 0, &mut 1000, &mut retained)
        .unwrap_err()
        .contains("retained anchor budget"));
    assert_eq!(retained.len(), MAX_OBJECTS);
}

#[test]
fn binds_only_the_owning_shapes_picture_property_and_plain_native_object() {
    let prop = art(
        0xf00b,
        0x13,
        &[
            0x4104u16.to_le_bytes().to_vec(),
            7u32.to_le_bytes().to_vec(),
        ]
        .concat(),
    );
    let mut source = fixture();
    source[2]
        .1
        .extend_from_slice(&[7, 0, 2, 0, 255, 255, 8, 0, 2, 0, 0, 0, 0, 0, 0, 0]);
    let mut outer = source.clone();
    let root_size = u32::from_le_bytes(outer[1].1[4..8].try_into().unwrap()) + prop.len() as u32;
    outer[1].1[4..8].copy_from_slice(&root_size.to_le_bytes());
    outer[1].1.splice(8..8, prop.clone());
    assert!(run(&outer).unwrap()[0].picture.is_none());

    for at in [4, 12] {
        let n = u32::from_le_bytes(source[1].1[at..at + 4].try_into().unwrap()) + prop.len() as u32;
        source[1].1[at..at + 4].copy_from_slice(&n.to_le_bytes());
    }
    source[1].1.splice(32..32, prop);
    assert_eq!(
        run(&source).unwrap()[0].picture,
        Some(PictureReference {
            store_index: 7,
            crop: [0; 4],
            rotation: 0,
            clipboard_format: 0xffff,
            auto_picture: false,
        })
    );
    source[2].1[32] = 16; // ActiveX cannot become a passive catalog reference.
    assert!(run(&source).unwrap()[0].picture.is_none());
}
