//! Inspect already-extracted alternative shape XML without applying it to text.
//! This uses dev-only dependencies and is not compiled into converter WASM.
#[path = "../src/officeart/metro_text.rs"]
mod metro_text;
use std::io::Read;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = std::env::args().skip(1);
    let path = args.next().ok_or("expected one extracted shape XML path")?;
    if args.next().is_some() {
        return Err("expected exactly one path".into());
    }
    let mut bytes = Vec::new();
    std::fs::File::open(path)?
        .take(1024 * 1024 + 1)
        .read_to_end(&mut bytes)?;
    let mut budget = metro_text::Budget {
        bytes: 1024 * 1024,
        events: 100_000,
        paragraphs: 10_000,
    };
    let Some(paragraphs) = metro_text::read(&bytes, &mut budget)? else {
        println!("Unsupported alternative text structure; no projection");
        return Ok(());
    };
    println!("paragraphs={}", paragraphs.len());
    for (index, p) in paragraphs.iter().enumerate() {
        // Never print document text or treat this count as an identity proof.
        println!("paragraph={index} utf16={} margin_left={:?} margin_right={:?} indent={:?} default_tab_size={:?} level={:?}",
            p.literal.encode_utf16().count(), p.margin_left, p.margin_right, p.indent, p.default_tab_size, p.level);
    }
    Ok(())
}
