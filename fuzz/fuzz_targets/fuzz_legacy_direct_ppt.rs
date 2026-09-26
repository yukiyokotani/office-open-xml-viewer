#![no_main]

use legacy_office_converter::fuzzing;
use libfuzzer_sys::fuzz_target;

// Exercise the direct PPT session and slide cursor, including alternative
// shape XML (metroBlob) resolution on every shape that carries one. Raw input reaches the compound-file reader; the wrapped
// input places the bytes behind a valid root-linked container.
fuzz_target!(|data: &[u8]| {
    fuzzing::direct_ppt(data);
    fuzzing::direct_ppt(&fuzzing::wrapped_ppt(&data[..data.len().min(1024 * 1024)]));
});
