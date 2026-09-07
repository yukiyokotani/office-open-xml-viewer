//! Local-only regression oracle for the neutral SST and its native projection.
//! Prior OOXML archives are read by the existing XLSX parser only in tests.

#[test]
#[ignore = "requires local binary corpus and immutable pre-change OOXML outputs"]
fn native_strings_match_immutable_corpus_inline_models() {
    use super::CellValue;
    use std::{collections::BTreeMap, path::PathBuf};

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
    let mut compared = 0;
    let mut mismatches = Vec::new();
    for (document, path) in inputs.iter().enumerate() {
        let bytes = std::fs::read(path).unwrap();
        let compound = crate::cfb::CompoundFile::open(&bytes).unwrap();
        let prepared = super::prepare(&compound, false).unwrap();
        let archive =
            std::fs::read(baseline.join(format!("{document}.xlsx"))).expect("missing baseline");
        let mut budget = super::rich::MAX_MODEL_BYTES;
        let strings: Vec<_> = prepared
            .shared_strings
            .iter()
            .map(|text| {
                serde_json::to_value(text.model(&prepared.styles, &mut budget).unwrap()).unwrap()
            })
            .collect();
        for (sheet_index, (name, sheet)) in prepared.sheets.iter().enumerate() {
            let expected: serde_json::Value = serde_json::from_str(
                &xlsx_parser::parse_sheet_native(&archive, sheet_index as u32, name).unwrap(),
            )
            .unwrap();
            let mut cells = BTreeMap::new();
            for row in expected["rows"].as_array().expect("oracle worksheet rows") {
                for cell in row["cells"].as_array().unwrap() {
                    cells.insert(
                        (cell["row"].as_u64().unwrap(), cell["col"].as_u64().unwrap()),
                        cell,
                    );
                }
            }
            for (&row, columns) in &sheet.rows {
                for (&column, value) in columns {
                    if let CellValue::SharedString(index) = value {
                        let reference = cells
                            .get(&(u64::from(row) + 1, u64::from(column) + 1))
                            .expect("missing oracle cell");
                        let mut expected_value = reference["value"].clone();
                        assert_eq!(expected_value["type"], "text");
                        expected_value.as_object_mut().unwrap().remove("type");
                        if strings[*index] != expected_value {
                            // Content-free coordinates, never private names or text.
                            mismatches.push((document, sheet_index, row, column));
                        }
                        compared += 1;
                    }
                }
            }
        }
    }
    assert!(compared > 0, "no shared-string cells were checked");
    assert!(
        mismatches.is_empty(),
        "native SST mismatches: {mismatches:?}"
    );
    eprintln!(
        "native SST model parity: {} corpus inputs, {compared} shared-string cells",
        inputs.len()
    );
}
