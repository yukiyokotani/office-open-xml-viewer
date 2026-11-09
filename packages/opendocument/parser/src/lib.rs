//! OpenDocument readers projecting packaged ODF documents directly into the
//! OOXML renderer models. Each reader is a cargo feature and its own WASM.

pub mod error;
pub mod iri;
pub mod package;
pub mod script;
pub mod style;
pub mod xml;

#[cfg(feature = "odt")]
pub mod odt;
#[cfg(feature = "odt")]
mod odt_wasm;

#[cfg(feature = "odp")]
pub mod odp;
#[cfg(feature = "odp")]
mod odp_cursor;
#[cfg(feature = "odp")]
pub mod odp_geometry;
#[cfg(feature = "odp")]
mod odp_extrusion;
#[cfg(feature = "odp")]
mod odp_media;
#[cfg(feature = "odp")]
mod odp_paint;
#[cfg(feature = "odp")]
mod odp_style;
#[cfg(feature = "odp")]
mod odp_table;
#[cfg(feature = "odp")]
mod odp_text;
#[cfg(feature = "odp")]
mod odp_units;
#[cfg(feature = "odp")]
mod odp_wasm;
#[cfg(feature = "odf-chart")]
pub mod odf_chart;
#[cfg(feature = "odf-chart")]
pub mod odf_fill;
#[cfg(feature = "odf-chart")]
pub mod odf_props;
#[cfg(feature = "odf-chart")]
pub mod odf_text;
#[cfg(feature = "ods")]
pub mod ods;
#[cfg(feature = "odf-chart")]
pub mod ods_number;
#[cfg(feature = "ods")]
mod ods_wasm;
#[cfg(any(feature = "odt", feature = "odp"))]
mod svm;
