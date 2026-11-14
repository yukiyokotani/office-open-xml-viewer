//! Prints the serialize-mode JSON Schemas of the XLSX model roots, keyed by
//! root name: the workbook payload (`Workbook`) and each worksheet payload
//! (`Worksheet`), i.e. the shapes of the JSON the parser hands to TypeScript.
//!
//! Dev-only (`--features schema`); consumed by
//! `scripts/check-model-wire-drift.mjs` to detect drift between the Rust model
//! and the hand-written TS model types.

use schemars::generate::SchemaSettings;
use schemars::JsonSchema;

fn root<T: JsonSchema>() -> serde_json::Value {
    let generator = SchemaSettings::draft2020_12()
        .for_serialize()
        .into_generator();
    serde_json::to_value(generator.into_root_schema_for::<T>()).expect("schema serializes")
}

fn main() {
    let roots = serde_json::json!({
        "Workbook": root::<xlsx_model::Workbook>(),
        "Worksheet": root::<xlsx_model::Worksheet>(),
    });
    println!(
        "{}",
        serde_json::to_string_pretty(&roots).expect("schema serializes")
    );
}
