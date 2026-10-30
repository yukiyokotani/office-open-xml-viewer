//! OpenDocument readers projecting packaged ODF documents directly into the
//! OOXML renderer models. Each reader is a cargo feature and its own WASM.

pub mod error;
pub mod iri;
pub mod package;
pub mod script;
pub mod xml;
