#![no_main]

use libfuzzer_sys::fuzz_target;
use std::io::{Cursor, Write};
use zip::write::SimpleFileOptions;

// Treat the input as the shape part's XML so mutations reach the standalone
// parser even when they would not form a complete OPC archive on their own.
fuzz_target!(|data: &[u8]| {
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    if writer
        .start_file("shape.xml", SimpleFileOptions::default())
        .is_ok()
        && writer.write_all(data).is_ok()
    {
        if let Ok(package) = writer.finish() {
            let _ = pptx_parser::parse_standalone_shape_part(
                package.get_ref(),
                "shape.xml",
                "",
                None,
                1_048_576,
                2_097_152,
            );
        }
    }
});
