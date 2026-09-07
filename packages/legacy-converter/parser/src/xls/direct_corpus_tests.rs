//! Local-only end-to-end oracle for the native XLS renderer-model session.
//! Immutable OOXML baselines are parsed only by the existing XLSX parser here.

use super::direct::DirectSession;
use crate::cfb::CompoundFile;
use serde_json::Value;
use std::path::PathBuf;

fn directory_offset(streams: usize) -> usize {
    // cfb::test_support::build_cfb pads each regular stream to eight sectors.
    512 + streams * 8 * 512
}

fn set_directory_link(bytes: &mut [u8], streams: usize, id: usize, at: usize, target: u32) {
    let offset = directory_offset(streams) + id * 128 + at;
    bytes[offset..offset + 4].copy_from_slice(&target.to_le_bytes());
}

#[test]
fn direct_session_requires_root_workbook_stream() {
    let streams = [("Embedded", Vec::new()), ("Workbook", vec![0; 4096])];
    let mut bytes = crate::cfb::test_support::build_cfb(&streams);
    let storage = directory_offset(streams.len()) + 128;
    bytes[storage + 66] = 1;
    bytes[storage + 116..storage + 120].copy_from_slice(&0xffff_ffffu32.to_le_bytes());
    bytes[storage + 120..storage + 128].fill(0);
    set_directory_link(&mut bytes, streams.len(), 0, 76, 1);
    set_directory_link(&mut bytes, streams.len(), 1, 76, 2);

    let compound = CompoundFile::open(&bytes).expect("structurally valid scoped CFB fixture");
    let scoped = compound
        .scoped_streams()
        .expect("structurally valid scoped directory tree");
    assert!(scoped.has_stream(&["Embedded", "Workbook"]).unwrap());
    assert!(!scoped.has_stream(&["Workbook"]).unwrap());
    let error = DirectSession::new(&compound)
        .err()
        .expect("missing root stream");
    assert!(error.contains("missing CFB path entry: Book"));
}

#[test]
fn direct_retention_preflight_bounds_decoded_string_expansion() {
    let string_header = u16::MAX.to_le_bytes();
    let per_record = usize::from(u16::MAX) * 3;
    let count = (256 * 1024 * 1024) / per_record + 1;
    let records: Vec<_> = (0..count)
        .map(|offset| super::Record {
            kind: super::STRING,
            offset,
            data: &string_header,
        })
        .collect();
    let error = super::validate_direct_retention(&records).unwrap_err();
    assert!(error.contains("neutral retention budget exceeded"));
}

fn parsed_bootstrap(bytes: Vec<u8>) -> Value {
    let mut archive = xlsx_parser::XlsxArchive::new(bytes, None, None, None).unwrap();
    serde_json::from_slice(&archive.parse().unwrap()).unwrap()
}

fn resolve_shared_cells(sheet: &mut Value, shared_strings: &[Value]) {
    let rows = sheet["rows"].as_array_mut().expect("native worksheet rows");
    for row in rows {
        for cell in row["cells"].as_array_mut().expect("native worksheet cells") {
            let value = &mut cell["value"];
            if value["type"] != "shared" {
                continue;
            }
            let index = value["si"].as_u64().expect("native shared-string index") as usize;
            let mut resolved = shared_strings
                .get(index)
                .expect("native shared-string index in bootstrap")
                .clone();
            resolved
                .as_object_mut()
                .expect("native shared string object")
                .insert("type".into(), Value::String("text".into()));
            *value = resolved;
        }
    }
}

fn first_difference_path(actual: &Value, expected: &Value, path: &str) -> Option<String> {
    match (actual, expected) {
        (Value::Array(actual), Value::Array(expected)) => {
            if actual.len() != expected.len() {
                return Some(format!("{path}.length"));
            }
            actual
                .iter()
                .zip(expected)
                .enumerate()
                .find_map(|(index, (actual, expected))| {
                    first_difference_path(actual, expected, &format!("{path}[{index}]"))
                })
        }
        (Value::Object(actual), Value::Object(expected)) => {
            for key in actual.keys().chain(expected.keys()) {
                match (actual.get(key), expected.get(key)) {
                    (Some(actual), Some(expected)) => {
                        if let Some(path) =
                            first_difference_path(actual, expected, &format!("{path}.{key}"))
                        {
                            return Some(path);
                        }
                    }
                    _ => return Some(format!("{path}.{key}")),
                }
            }
            None
        }
        _ if actual == expected => None,
        _ => Some(path.to_owned()),
    }
}

#[test]
#[ignore = "requires local binary corpus and immutable pre-change OOXML outputs"]
fn native_session_matches_immutable_archive_models() {
    let corpus =
        PathBuf::from(std::env::var_os("LEGACY_XLS_STYLE_CORPUS").expect("set corpus directory"));
    let baseline = PathBuf::from(
        std::env::var_os("LEGACY_XLS_STYLE_BASELINE").expect("set immutable baseline directory"),
    );
    let mut inputs: Vec<_> = std::fs::read_dir(corpus)
        .unwrap()
        .map(|entry| entry.unwrap().path())
        .filter(|path| path.extension().is_some_and(|extension| extension == "xls"))
        .collect();
    inputs.sort();
    assert!(!inputs.is_empty(), "empty corpus is not coverage");

    let mut compared_sheets = 0usize;
    let mut mismatches = Vec::new();
    for (document, path) in inputs.iter().enumerate() {
        let source = std::fs::read(path).unwrap();
        let compound = CompoundFile::open(&source)
            .unwrap_or_else(|error| panic!("corpus index {document} CFB error: {error}"));
        let mut direct = DirectSession::new(&compound).unwrap_or_else(|error| {
            panic!("corpus index {document} direct admission error: {error}")
        });
        if direct.requires_measurement_decision() {
            direct.configure_mdw(None).unwrap_or_else(|error| {
                panic!("corpus index {document} measurement decision error: {error}")
            });
        }
        let actual_bootstrap =
            serde_json::to_value(direct.bootstrap().unwrap_or_else(|error| {
                panic!("corpus index {document} bootstrap error: {error}")
            }))
            .unwrap();

        let baseline_bytes = std::fs::read(baseline.join(format!("{document}.xlsx")))
            .unwrap_or_else(|error| panic!("corpus index {document} missing baseline: {error}"));
        let expected_bootstrap = parsed_bootstrap(baseline_bytes.clone());
        if let Some(path) = first_difference_path(
            &actual_bootstrap["workbook"],
            &expected_bootstrap["workbook"],
            "$",
        ) {
            mismatches.push(format!("corpus index {document} workbook: {path}"));
        }
        if let Some(path) = first_difference_path(
            &actual_bootstrap["styles"],
            &expected_bootstrap["styles"],
            "$",
        ) {
            mismatches.push(format!("corpus index {document} styles: {path}"));
        }
        let shared_strings = actual_bootstrap["sharedStrings"]
            .as_array()
            .expect("native bootstrap shared strings");

        let sheet_count = actual_bootstrap["workbook"]["sheets"]
            .as_array()
            .expect("native workbook sheets")
            .len();
        // Reverse order exercises indexed slots instead of an accidental
        // dependency on consuming sheets in workbook order.
        for sheet_index in (0..sheet_count).rev() {
            let sheet_name = actual_bootstrap["workbook"]["sheets"][sheet_index]["name"]
                .as_str()
                .expect("native sheet name");
            let projected = direct.projected_sheet(sheet_index, sheet_name)
                .unwrap_or_else(|error| {
                    panic!("corpus index {document} sheet {sheet_index} error: {error}")
                });
            assert!(projected.worksheet.rows.is_empty());
            let mut actual = serde_json::to_value(projected.worksheet).unwrap();
            actual["rows"] = serde_json::to_value(projected.rows).unwrap();
            resolve_shared_cells(&mut actual, shared_strings);
            // The XLSX parser oracle crosses its JSON byte boundary, while
            // `to_value` retains the model's original f64 bits. Cross the same
            // serde_json boundary before exact comparison so numeric spelling
            // and reparsing are compared symmetrically.
            actual = serde_json::from_slice(&serde_json::to_vec(&actual).unwrap()).unwrap();
            let expected: Value = serde_json::from_str(
                &xlsx_parser::parse_sheet_native(&baseline_bytes, sheet_index as u32, sheet_name)
                    .unwrap_or_else(|error| {
                        panic!(
                            "corpus index {document} baseline sheet {sheet_index} error: {error}"
                        )
                    }),
            )
            .unwrap();
            if let Some(path) = first_difference_path(&actual, &expected, "$") {
                mismatches.push(format!(
                    "corpus index {document}, sheet index {sheet_index}: {path}"
                ));
            }
            compared_sheets += 1;
        }
        // The transitional ownership API can still consume projected slots.
        for _ in 0..sheet_count {
            assert!(direct.next_sheet().unwrap().is_some());
        }
        assert!(direct.next_sheet().unwrap().is_none());
    }
    assert!(compared_sheets > 0, "no worksheets were compared");
    assert!(
        mismatches.is_empty(),
        "native worksheet model mismatches:\n{}",
        mismatches.join("\n")
    );
    eprintln!(
        "native worksheet model parity: {} corpus inputs, {compared_sheets} sheets",
        inputs.len()
    );
}
