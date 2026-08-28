//! Shared OOXML parsing helpers used by the docx, pptx and xlsx Rust parsers.
//!
//! This crate is intentionally small — only logic that already had two
//! near-duplicate copies in the wild parsers belongs here, plus format-agnostic
//! grammars like OMML (`math`) that every host schema embeds verbatim. Anything
//! schema-specific (DocParagraph, ShapeRun, Slide, etc.) stays in the
//! consuming crate.

pub mod blip;
#[doc(hidden)]
pub mod bounded_xml;
pub mod chart;
pub mod color;
pub mod content_types;
pub mod custom_geometry;
pub mod depth;
pub mod drawing;
pub mod fill;
pub mod json_measurement;
pub mod line;
pub mod math;
pub mod mce;
pub mod ns;
pub mod package_session;
pub mod pull;
pub mod rels;
pub mod resource;
pub mod text;
pub mod theme;
pub mod units;
pub mod zip;
