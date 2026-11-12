//! OpenDocument readers projecting packaged ODF documents directly into the
//! OOXML renderer models. Each reader is a cargo feature and its own WASM.

pub mod budget;
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
#[cfg(feature = "odt")]
mod svm;

/// Every allocation of the reader is metered ([`budget`]): its resource
/// limits are measured by the allocator, not estimated by the reader.
#[cfg(feature = "odt")]
#[global_allocator]
static METERED: budget::Metered = budget::Metered;
