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

/// The bytes one copy of `value` retains: inline plus owned heap
/// ([`ooxml_common::chart::RetainedBytes`] destructures every field, so a
/// new model field fails to compile until it is counted). Readers that copy
/// data by repetition (repeated rows and cells, their links and notes,
/// per-sheet names, per-series chart styles) admit this size for each copy.
#[cfg(feature = "odf-chart")]
pub(crate) fn retained_size<T: ooxml_common::chart::RetainedBytes>(value: &T) -> usize {
    usize::try_from(value.heap_bytes())
        .unwrap_or(usize::MAX)
        .saturating_add(std::mem::size_of::<T>())
}
