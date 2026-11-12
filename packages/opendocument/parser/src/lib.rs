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

#[cfg(feature = "odf-chart")]
pub mod budget;
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

/// Every allocation of the readers built on the shared ODF modules is
/// metered ([`budget`]): their resource limits are measured by the allocator,
/// not estimated by the readers.
#[cfg(feature = "odf-chart")]
#[global_allocator]
static METERED: budget::Metered = budget::Metered;
