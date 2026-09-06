use super::*;

fn object() -> Vec<u8> {
    let mut bytes = vec![0; 38];
    for (at, value) in [
        (0, 0x15u16),
        (2, 0x12),
        (4, 8),
        (6, 7),
        (8, 0x11),
        (22, 7),
        (24, 2),
        (26, 0xffff),
        (28, 8),
        (30, 2),
    ] {
        bytes[at..at + 2].copy_from_slice(&value.to_le_bytes());
    }
    bytes
}
fn table(entries: &[(u16, u32)]) -> Vec<u8> {
    entries
        .iter()
        .flat_map(|(id, value)| [id.to_le_bytes().to_vec(), value.to_le_bytes().to_vec()].concat())
        .collect()
}
fn parse(entries: &[(u16, u32)]) -> Result<Properties, String> {
    let bytes = table(entries);
    let mut props = Properties::default();
    props.read(
        ArtRecord {
            kind: 0xf00b,
            version: 3,
            instance: entries.len() as u16,
            payload: &bytes,
        },
        &mut 100,
    )?;
    Ok(props)
}

#[test]
fn keeps_local_index_signed_crop_and_rotation_without_layout_guesses() {
    let props = parse(&[
        (0x4104, 7),
        (4, (-45i32 * 65536) as u32),
        (0x100, i32::MIN as u32),
        (0x101, 65536),
        (0x102, 0),
        (0x103, i32::MAX as u32),
    ])
    .unwrap();
    assert_eq!(
        props.reference(0x200, Some(&object())).unwrap(),
        Some(PictureReference {
            store_index: 7,
            crop: [i32::MIN, 65536, 0, i32::MAX],
            rotation: -45 * 65536,
            clipboard_format: 0xffff,
            auto_picture: false,
        })
    );
    assert!(parse(&[(0x4104, 0)])
        .unwrap()
        .reference(0x200, Some(&object()))
        .unwrap()
        .is_none());
    assert!(props.reference(0x200, None).unwrap().is_none());
    for flag in [1, 4, 8, 16, 1024] {
        assert!(props
            .reference(0x200 | flag, Some(&object()))
            .unwrap()
            .is_none());
    }
    for cf in [2u16, 9, 0xffff] {
        let mut data = object();
        data[26..28].copy_from_slice(&cf.to_le_bytes());
        assert!(passive_object(&data).unwrap());
    }
}

#[test]
fn distinguishes_undefined_bits_from_dde_control_camera_or_dynamic_pictures() {
    for bit in [1, 2, 3, 4, 5, 7, 8, 9] {
        let mut bytes = object();
        bytes[32..34].copy_from_slice(&(1u16 << bit).to_le_bytes());
        assert!(!passive_object(&bytes).unwrap());
    }
    let mut bytes = object();
    bytes[32..34].copy_from_slice(&0xfc41u16.to_le_bytes());
    bytes[34..].fill(255); // reserved values MUST be ignored.
    assert!(passive_object(&bytes).unwrap());
    bytes[32..34].copy_from_slice(&0x12u16.to_le_bytes());
    assert!(passive_object(&bytes).is_err());
    let mut bytes = object();
    bytes[8] |= 4;
    assert!(!passive_object(&bytes).unwrap());
    for extra in [vec![4, 0, 0, 0], vec![9, 0, 4, 0, 1, 2, 3, 4]] {
        assert!(!passive_object(&[object(), extra].concat()).unwrap());
    }
}

#[test]
fn rejects_truncated_or_misowned_picture_fields_instead_of_scanning_for_them() {
    let bytes = object();
    for length in 6..bytes.len() {
        assert!(passive_object(&bytes[..length]).is_err());
    }
    for at in [22, 24, 26, 28, 30] {
        let mut bad = bytes.clone();
        bad[at] = 99;
        assert!(passive_object(&bad).is_err());
    }
    let mut chart = bytes;
    chart[4] = 5;
    assert!(!passive_object(&chart).unwrap());
}

#[test]
fn does_not_turn_complex_linked_hidden_or_script_properties_into_indices() {
    for value in [1, 2, 4, 8, 10, 12, u32::MAX] {
        assert!(parse(&[(0x4104, 7), (0x106, value)])
            .unwrap()
            .reference(0x200, Some(&object()))
            .unwrap()
            .is_none());
    }
    for bit in [1, 7] {
        assert!(
            parse(&[(0x4104, 7), (0x3bf, (1 << bit) | (1 << (bit + 16)))])
                .unwrap()
                .reference(0x200, Some(&object()))
                .unwrap()
                .is_none()
        );
        assert!(parse(&[(0x4104, 7), (0x3bf, 1 << bit)])
            .unwrap()
            .reference(0x200, Some(&object()))
            .unwrap()
            .is_some());
    }
    let complex = [table(&[(0xc104, 3)]), vec![1, 2, 3]].concat();
    let mut props = Properties::default();
    props
        .read(
            ArtRecord {
                kind: 0xf00b,
                version: 3,
                instance: 1,
                payload: &complex,
            },
            &mut 10,
        )
        .unwrap();
    assert!(props.reference(0x200, Some(&object())).unwrap().is_none());
    for entries in [
        vec![(0x104, 7)],
        vec![(0x4104, 7), (0x4104, 8)],
        vec![(0x4106, 0)],
        vec![(0x4104, 7), (0x4100, 0)],
    ] {
        assert!(parse(&entries).is_err());
    }
    let mut props = parse(&[(0x4104, 7)]).unwrap();
    assert!(props
        .read(
            ArtRecord {
                kind: 0xf00b,
                version: 3,
                instance: 0,
                payload: &[]
            },
            &mut 10
        )
        .is_err());
}
