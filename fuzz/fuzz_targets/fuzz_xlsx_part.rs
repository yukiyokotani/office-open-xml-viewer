#![no_main]
use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    if let Some(package) = ooxml_fuzz::xlsx_part(data) {
        let _ = xlsx_parser::parse_workbook_native(&package);
        let _ = xlsx_parser::parse_sheet_native(&package, 0, "Sheet1");
    }
});
