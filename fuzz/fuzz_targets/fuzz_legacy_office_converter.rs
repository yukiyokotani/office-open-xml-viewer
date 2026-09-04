#![no_main]

use legacy_office_converter::{build_fuzz_input, convert_native, LegacyFormat};
use libfuzzer_sys::fuzz_target;

// Exercise CFB classification, FAT/MiniFAT traversal, each legacy-record
// parser, and the bounded OOXML writer. A small output ceiling keeps accidental
// valid inputs from turning corpus expansion into an unbounded allocation.
fuzz_target!(|data: &[u8]| {
    let bounded = &data[..data.len().min(1024 * 1024)];
    for format in [LegacyFormat::Doc, LegacyFormat::Xls, LegacyFormat::Ppt] {
        let _ = convert_native(data, format, 8 * 1024 * 1024);
        let wrapped = build_fuzz_input(bounded, format);
        let _ = convert_native(&wrapped, format, 8 * 1024 * 1024);
    }
});
