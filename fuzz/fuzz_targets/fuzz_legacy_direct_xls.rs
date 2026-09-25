#![no_main]

use legacy_office_converter::fuzzing;
use libfuzzer_sys::fuzz_target;

// Exercise the direct XLS session (BIFF8 records, styles, drawings, charts)
// through bootstrap and every worksheet projection. Raw input reaches the compound-file reader; the wrapped
// input places the bytes behind a valid root-linked container.
fuzz_target!(|data: &[u8]| {
    fuzzing::direct_xls(data);
    fuzzing::direct_xls(&fuzzing::wrapped_xls(&data[..data.len().min(1024 * 1024)]));
});
