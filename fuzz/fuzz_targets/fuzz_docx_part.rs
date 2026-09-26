#![no_main]
use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    if let Some(package) = ooxml_fuzz::docx_part(data) {
        let _ = docx_parser::parse_docx_native(&package);
    }
});
