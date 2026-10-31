//! Local inspection: admit or reject ODT files and print the reason or the
//! paragraph count. `cargo run --example inspect_odt --features odt -- <files>`
use opendocument_parser::{
    odt,
    package::{OdfPackage, PackageLimits},
};
fn main() {
    for path in std::env::args().skip(1) {
        let bytes = std::fs::read(&path).expect("read");
        let result = OdfPackage::open(
            bytes,
            ooxml_common::resource::OoxmlFormat::Docx,
            PackageLimits::default(),
            &odt::MEDIA_TYPES,
        )
        .and_then(|package| odt::read(&package));
        match result {
            Ok((document, images)) => println!(
                "{path}: admitted, {} body elements, {} images",
                document.body.len(),
                images.len()
            ),
            Err(error) => println!("{path}: {error}"),
        }
    }
}
