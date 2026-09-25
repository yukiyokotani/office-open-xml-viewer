#![no_main]

use legacy_office_converter::fuzzing;
use libfuzzer_sys::fuzz_target;

// Exercise the direct DOC projection (piece table, formatting, tables, notes,
// headers, fields, pictures) and its document cursor. Raw input reaches the compound-file reader; the wrapped
// input places the bytes behind a valid root-linked container.
fuzz_target!(|data: &[u8]| {
    fuzzing::direct_doc(data);
    fuzzing::direct_doc(&fuzzing::wrapped_doc(&data[..data.len().min(1024 * 1024)]));
});
