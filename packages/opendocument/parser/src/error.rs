//! Error vocabulary shared by the readers. The calling realm decodes the
//! encryption marker in `src/native-rejection.ts`; messages never carry
//! document text. There is no "unsupported content" error: the readers render
//! every feature they meet and reject only corrupt, encrypted or over-limit
//! input.

/// An encrypted package (ODF Part 2 §4.8 `manifest:encryption-data`).
pub const UNSUPPORTED_ENCRYPTION: &str = "ODF_UNSUPPORTED_ENCRYPTION";

/// A structurally invalid package or document (not a fidelity limitation).
pub fn invalid(message: &str) -> String {
    format!("OpenDocument package is invalid: {message}")
}

/// A resource-limit rejection (content-free): input over an implementation
/// bound is rejected, never shortened.
pub fn resource(what: &str) -> String {
    format!("OpenDocument document exceeds the {what} limit")
}

/// Whether `error` is a structural invalidity ([`invalid`]), as opposed to a
/// resource limit or a package failure, which always reach the caller.
pub fn is_invalid(error: &str) -> bool {
    error.starts_with("OpenDocument package is invalid:")
}
