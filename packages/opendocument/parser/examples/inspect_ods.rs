//! Local inspection: read ODS files and print the sheet count, the cells,
//! drawings and charts per sheet, or the rejection.
//! `cargo run --example inspect_ods --features ods -- <files>`; with
//! `--json <dir>` it also writes each workbook's bootstrap and worksheets.
use opendocument_parser::{
    ods,
    package::{OdfPackage, PackageLimits},
};
fn main() {
    let mut args: Vec<String> = std::env::args().skip(1).collect();
    let json_dir = args.iter().position(|a| a == "--json").map(|i| {
        let dir = args.remove(i + 1);
        args.remove(i);
        dir
    });
    for path in args {
        let bytes = std::fs::read(&path).expect("read");
        let result = OdfPackage::open(
            bytes,
            ooxml_common::resource::OoxmlFormat::Xlsx,
            PackageLimits::default(),
            &ods::MEDIA_TYPES,
        )
        .and_then(|package| ods::read(&package));
        match result {
            Ok(mut workbook) => {
                workbook.apply_layout(Some(7.0));
                let summary: Vec<String> = workbook
                    .sheets
                    .iter()
                    .map(|s| {
                        format!(
                            "{}r/{}c/{}i/{}ch/{}sh",
                            s.rows.len(),
                            s.rows.iter().map(|r| r.cells.len()).sum::<usize>(),
                            s.worksheet.images.len(),
                            s.worksheet.charts.len(),
                            s.worksheet.shape_groups.len()
                        )
                    })
                    .collect();
                println!("{path}: ok {}", summary.join(" "));
                if let Some(dir) = &json_dir {
                    let stem = std::path::Path::new(&path)
                        .file_stem()
                        .unwrap()
                        .to_string_lossy()
                        .to_string();
                    let mut sheets = Vec::new();
                    for sheet in &mut workbook.sheets {
                        sheet.worksheet.rows = std::mem::take(&mut sheet.rows);
                        sheets.push(serde_json::to_value(&sheet.worksheet).unwrap());
                    }
                    let boot = ods::bootstrap(&mut workbook);
                    let value = serde_json::json!({ "workbook": boot, "sheets": sheets });
                    std::fs::write(format!("{dir}/{stem}.json"), value.to_string()).unwrap();
                }
            }
            Err(error) => println!("{path}: {error}"),
        }
    }
}
