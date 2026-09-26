#![no_main]
use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    if let Some(package) = ooxml_fuzz::pptx_part(data) {
        let _ = pptx_parser::parse_pptx_native(&package);
    }
});
