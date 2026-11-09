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
