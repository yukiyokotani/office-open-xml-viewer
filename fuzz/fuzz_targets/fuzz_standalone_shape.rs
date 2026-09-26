#![no_main]

use libfuzzer_sys::fuzz_target;
use std::io::{Cursor, Write};
use zip::write::SimpleFileOptions;

const SHAPE: &str = r#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr><p:style><a:fillRef idx="1"/></p:style></p:sp>"#;
const THEME: &str = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fmtScheme name="fuzz"><a:fillStyleLst><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:fillStyleLst></a:fmtScheme></a:themeElements></a:theme>"#;
const CLR_MAP: &str = r#"<p:clrMap xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" accent1="accent1"/>"#;

// Mutate each input to the direct standalone API while keeping the others
// valid, so the fuzzer reaches shape, theme, and color-map parsing.
fuzz_target!(|data: &[u8]| {
    let (&mode, payload) = data.split_first().unwrap_or((&0, &[]));
    let text = String::from_utf8_lossy(payload);
    let (shape, theme, clr_map) = match mode % 3 {
        0 => (text.as_ref(), THEME, CLR_MAP),
        1 => (SHAPE, text.as_ref(), CLR_MAP),
        _ => (SHAPE, THEME, text.as_ref()),
    };
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    if writer
        .start_file("shape.xml", SimpleFileOptions::default())
        .is_ok()
        && writer.write_all(shape.as_bytes()).is_ok()
    {
        if let Ok(package) = writer.finish() {
            let _ = pptx_parser::parse_standalone_shape_part(
                package.get_ref(),
                "shape.xml",
                theme,
                Some(clr_map),
                1_048_576,
                2_097_152,
            );
        }
    }
});
