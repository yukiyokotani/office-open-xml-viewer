//! Prints the serialize-mode JSON Schema of the PPTX model root (`Presentation`, whose `slides` are the slide payloads),
//! i.e. the shape of the JSON the parser hands to TypeScript.
//!
//! Dev-only (`--features schema`); consumed by
//! `scripts/check-model-wire-drift.mjs` to detect drift between the Rust model
//! and the hand-written TS model types.

use schemars::generate::SchemaSettings;

fn main() {
    let generator = SchemaSettings::draft2020_12()
        .for_serialize()
        .into_generator();
    let schema = generator.into_root_schema_for::<pptx_model::Presentation>();
    println!(
        "{}",
        serde_json::to_string_pretty(&schema).expect("schema serializes")
    );
}
