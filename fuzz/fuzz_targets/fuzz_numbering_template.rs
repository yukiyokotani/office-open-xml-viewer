#![no_main]

use libfuzzer_sys::fuzz_target;
use ooxml_common::numbering::{NumberingTemplate, TemplatePart};

fuzz_target!(|data: &[u8]| {
    let parts = data
        .chunks(3)
        .take(20)
        .map(|chunk| {
            if chunk[0] & 1 == 0 {
                TemplatePart::Counter(chunk.get(1).copied().unwrap_or(0))
            } else {
                TemplatePart::Literal(String::from_utf8_lossy(chunk).into_owned())
            }
        })
        .collect();
    if let Ok(template) = NumberingTemplate::new(parts) {
        let _ = template.expand(|level| level as u32, |_| "decimal");
    }
    let _ = NumberingTemplate::new(vec![TemplatePart::Literal(
        String::from_utf8_lossy(data).into_owned(),
    )]);
});
