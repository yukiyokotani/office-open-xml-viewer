//! Independent archive-parser oracle for native style projection.
//! XML parsing exists only in this test module, never in the direct producer.

use super::{records, styles::Styles};
use crate::cfb::CompoundFile;

fn native_styles(input: &[u8]) -> serde_json::Value {
    let cfb = CompoundFile::open(input).expect("valid test CFB");
    let stream = cfb
        .stream("Workbook")
        .or_else(|_| cfb.stream("Book"))
        .unwrap();
    let records = records(&stream).unwrap();
    serde_json::to_value(
        Styles::parse(&records)
            .unwrap()
            .resolve()
            .unwrap()
            .into_model(),
    )
    .unwrap()
}

fn archive_styles(input: Vec<u8>) -> serde_json::Value {
    let mut archive = xlsx_parser::XlsxArchive::new(input, None, None, None).unwrap();
    let value: serde_json::Value = serde_json::from_slice(&archive.parse().unwrap()).unwrap();
    assert!(
        value["styles"].is_object(),
        "oracle must return a full style table"
    );
    value["styles"].clone()
}

fn record(kind: u16, payload: &[u8]) -> Vec<u8> {
    [
        kind.to_le_bytes().as_slice(),
        &(payload.len() as u16).to_le_bytes(),
        payload,
    ]
    .concat()
}

#[test]
fn empty_and_font_only_styles_match_archive_defaults() {
    for include_font in [false, true] {
        let mut stream = record(super::BOF, &[0, 6, 5, 0]);
        let bound = stream.len() + 4;
        stream.extend(record(super::BOUNDSHEET8, &[0, 0, 0, 0, 0, 0, 1, 0, b'S']));
        if include_font {
            let mut font = vec![0_u8; 16];
            font[..2].copy_from_slice(&220_u16.to_le_bytes());
            font[4..6].copy_from_slice(&0x7fff_u16.to_le_bytes());
            font[6..8].copy_from_slice(&400_u16.to_le_bytes());
            font[14] = 1;
            font.push(b'F');
            stream.extend(record(0x31, &font));
        }
        stream.extend(record(super::EOF, &[]));
        let offset = stream.len() as u32;
        stream[bound..bound + 4].copy_from_slice(&offset.to_le_bytes());
        stream.extend(record(super::BOF, &[0, 6, 0x10, 0]));
        stream.extend(record(super::EOF, &[]));
        let input = crate::cfb::test_support::build_cfb(&[("Workbook", stream)]);
        let converted = crate::convert_native(&input, crate::LegacyFormat::Xls, 1_000_000).unwrap();
        assert_eq!(native_styles(&input), archive_styles(converted.bytes));
    }
}

#[test]
fn native_style_matrix_matches_existing_archive_parser() {
    let mut stream = record(super::BOF, &[0, 6, 5, 0]);
    let bound = stream.len() + 4;
    stream.extend(record(super::BOUNDSHEET8, &[0, 0, 0, 0, 0, 0, 1, 0, b'S']));
    for weight in [400_u16, 650, 700] {
        let mut font = vec![0_u8; 16];
        font[..2].copy_from_slice(&220_u16.to_le_bytes());
        font[4..6].copy_from_slice(&0x7fff_u16.to_le_bytes());
        font[6..8].copy_from_slice(&weight.to_le_bytes());
        font[14] = 1;
        font.push(b'F');
        stream.extend(record(0x31, &font));
    }
    let mut count = 0_u32;
    for horizontal in 0_u8..8 {
        for vertical in 0_u8..5 {
            for reading in 0_u8..3 {
                for rotation in [0_u8, 1, 90, 91, 180, 255] {
                    let mut xf = [0_u8; 20];
                    xf[..2].copy_from_slice(&((count % 3) as u16).to_le_bytes());
                    xf[2..4].copy_from_slice(&((count % 50) as u16).to_le_bytes());
                    xf[4] = (count % 16) as u8;
                    xf[6] = horizontal
                        | (vertical << 4)
                        | ((count as u8 & 1) << 3)
                        | ((count as u8 & 2) << 6);
                    xf[7] = rotation;
                    xf[8] = (reading << 6) | (count as u8 & 15) | ((count as u8 & 1) << 4);
                    let edge = count % 14;
                    let b1 = edge
                        | (edge << 4)
                        | (edge << 8)
                        | (edge << 12)
                        | ((count % 64) << 16)
                        | ((count % 64) << 23)
                        | ((count % 4) << 30);
                    let b2 = count % 64
                        | ((count % 64) << 7)
                        | ((count % 64) << 14)
                        | (edge << 21)
                        | ((count % 19) << 26);
                    xf[10..14].copy_from_slice(&b1.to_le_bytes());
                    xf[14..18].copy_from_slice(&b2.to_le_bytes());
                    xf[18..20].copy_from_slice(
                        &((count % 64 | ((count % 64) << 7)) as u16).to_le_bytes(),
                    );
                    stream.extend(record(0xe0, &xf));
                    count += 1;
                }
            }
        }
    }
    stream.extend(record(super::EOF, &[]));
    let offset = stream.len() as u32;
    stream[bound..bound + 4].copy_from_slice(&offset.to_le_bytes());
    stream.extend(record(super::BOF, &[0, 6, 0x10, 0]));
    stream.extend(record(super::EOF, &[]));
    let input = crate::cfb::test_support::build_cfb(&[("Workbook", stream)]);
    let converted =
        crate::convert_native(&input, crate::LegacyFormat::Xls, 10 * 1024 * 1024).unwrap();
    let expected = archive_styles(converted.bytes);
    let actual = native_styles(&input);
    assert_eq!(count, 720);
    assert_eq!(actual, expected);
}

#[test]
#[ignore = "requires local binary corpus and immutable pre-change OOXML outputs"]
fn native_styles_match_immutable_corpus_archive_models() {
    let corpus = std::path::PathBuf::from(
        std::env::var_os("LEGACY_XLS_STYLE_CORPUS").expect("set corpus directory"),
    );
    let baseline = std::path::PathBuf::from(
        std::env::var_os("LEGACY_XLS_STYLE_BASELINE").expect("set immutable baseline directory"),
    );
    let mut inputs: Vec<_> = std::fs::read_dir(corpus)
        .unwrap()
        .map(|entry| entry.unwrap().path())
        .filter(|path| path.extension().is_some_and(|extension| extension == "xls"))
        .collect();
    inputs.sort();
    assert!(!inputs.is_empty(), "empty corpus is not coverage");
    for (index, path) in inputs.iter().enumerate() {
        let expected = archive_styles(
            std::fs::read(baseline.join(format!("{index}.xlsx"))).expect("missing baseline"),
        );
        let actual = native_styles(&std::fs::read(path).unwrap());
        // Do not print private paths or document values in test diagnostics.
        assert!(
            actual == expected,
            "native style mismatch at corpus index {index}"
        );
    }
    eprintln!("native style model parity: {} corpus inputs", inputs.len());
}
