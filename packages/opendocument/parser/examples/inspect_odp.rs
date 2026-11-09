//! Local inspection: admit or reject ODP files and print the reason or the
//! slide count. `cargo run --example inspect_odp --features odp -- <files>`
//! With `ODP_DUMP_SLIDE=<n>` the n-th (0-based) slide model is printed as JSON.
use opendocument_parser::{
    odp,
    package::{OdfPackage, PackageLimits},
};
fn main() {
    let dump = std::env::var("ODP_DUMP_SLIDE")
        .ok()
        .and_then(|v| v.parse::<usize>().ok());
    for path in std::env::args().skip(1) {
        let bytes = std::fs::read(&path).expect("read");
        let result = OdfPackage::open(
            bytes,
            ooxml_common::resource::OoxmlFormat::Pptx,
            PackageLimits::default(),
            &odp::MEDIA_TYPES,
        )
        .and_then(|package| odp::read(&package));
        match result {
            Ok(deck) => {
                let hidden: Vec<usize> = deck
                    .slides
                    .iter()
                    .filter(|s| s.hidden)
                    .map(|s| s.index)
                    .collect();
                let elements: usize = deck.slides.iter().map(|s| s.elements.len()).sum();
                println!(
                    "{path}: admitted, {} slides, hidden={hidden:?}, {elements} elements, {} resources",
                    deck.slides.len(),
                    deck.resources.len()
                );
                if let Some(n) = dump {
                    if let Some(slide) = deck.slides.get(n) {
                        println!("{}", serde_json::to_string_pretty(slide).unwrap());
                    }
                }
            }
            Err(error) => println!("{path}: {error}"),
        }
    }
}
