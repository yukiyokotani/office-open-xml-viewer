//! Local inspection: for every embedded object of each ODP file, print its
//! chart class and whether the shared chart reader maps it.
//! `cargo run --example inspect_odp_charts --features odp -- <files>`
use opendocument_parser::{
    odf_chart, odp,
    package::{OdfPackage, PackageLimits},
};
fn main() {
    for path in std::env::args().skip(1) {
        let bytes = std::fs::read(&path).expect("read");
        let package = match OdfPackage::open(
            bytes,
            ooxml_common::resource::OoxmlFormat::Pptx,
            PackageLimits::default(),
            &odp::MEDIA_TYPES,
        ) {
            Ok(p) => p,
            Err(e) => {
                println!("{path}\topen-error\t{e}");
                continue;
            }
        };
        let mut dirs: Vec<String> = package
            .session()
            .entry_paths()
            .into_iter()
            .filter_map(|p| p.strip_suffix("/content.xml").map(str::to_string))
            .filter(|d| !d.contains('/'))
            .collect();
        dirs.sort();
        for dir in dirs {
            let content = package
                .read_text(&format!("{dir}/content.xml"))
                .unwrap_or_default();
            let class = content
                .split("chart:class=\"")
                .nth(1)
                .and_then(|s| s.split('"').next())
                .unwrap_or("-");
            let result = match odf_chart::read_chart_object(&package, &dir, None) {
                Ok(Some(_)) => "mapped".to_string(),
                Ok(None) => "none".to_string(),
                Err(e) => format!("error {e}"),
            };
            println!("{path}\t{dir}\t{class}\t{result}");
        }
    }
}
