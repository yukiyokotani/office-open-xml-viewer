#![no_main]

use legacy_office_converter::fuzzing;
use libfuzzer_sys::fuzz_target;

// Exercise alternative shape XML (metroBlob) resolution directly: package and
// relationship reading, the standalone shape parse and every comparison that
// decides between adopting it, keeping the binary shape and failing closed.
fuzz_target!(|data: &[u8]| {
    fuzzing::ppt_alternative(&data[..data.len().min(1024 * 1024)]);
});
