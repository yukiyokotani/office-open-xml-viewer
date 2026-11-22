//! Reader admission policy, separate from document syntax and cursor misuse.
//! These internal markers are consumed only around native calls. They must not
//! be inferred from diagnostic wording in JS: limits and encryption also used
//! to share UNSUPPORTED with representational failures.
//!
//! # Policy
//!
//! The readers render every feature they read. A whole input is rejected only
//! for one of the classes below; everything else is projected to the nearest
//! equivalent of the target model, without misplacing the content around it.
//!
//! | class | constructor | marker today |
//! |---|---|---|
//! | corrupt or unsafe framing | [`corrupt`] | `UNSUPPORTED:` |
//! | resource limit | [`resource`] | `LEGACY_RESOURCE_LIMIT:` |
//! | encryption | [`encryption`] | `LEGACY_UNSUPPORTED_ENCRYPTION:` |
//! | version these readers do not read | [`out_of_scope_version`] | `UNSUPPORTED:` |
//! | not classified yet | [`unsupported`] | `UNSUPPORTED:` |
//!
//! [`unsupported`] is the remainder: sites that still reject although the
//! policy says to render. Converting one means replacing it with exactly one
//! of:
//!
//! - [`corrupt`] or [`out_of_scope_version`], when the site belongs to that
//!   class (no behaviour change);
//! - [`ignored_value`] or [`spec_default`], when a single value inside a
//!   correctly framed record is outside its specified range: the property is
//!   left unset, or takes the default the specification gives it;
//! - [`skipped`], when the feature has no visual equivalent and the content
//!   that is drawn does not depend on it;
//! - [`approximated`], when the target model has a nearest equivalent;
//! - a real projection, when dropping the construct would move the
//!   surrounding content. No helper stands in for that.
//!
//! The site argument of the best-effort helpers is the diagnostic of the
//! rejection they replace, so a conversion stays traceable. It is recorded in
//! unit tests only ([`take_best_effort`]); there is no runtime report.
//!
//! A flag that is set where a construct is met and turned into one rejection
//! at the end of a projection is a [`Deferred`].

const UNSUPPORTED: &str = "UNSUPPORTED:";
const RESOURCE_LIMIT: &str = "LEGACY_RESOURCE_LIMIT:";
const ENCRYPTION: &str = "LEGACY_UNSUPPORTED_ENCRYPTION:";
/// Typed quota of the shared OOXML resource accounting.
const MEASURED_RESOURCE_LIMIT: &str = "OOXML_RESOURCE_LIMIT:";

/// Mark `message` as an `UNSUPPORTED:` rejection unless it already carries a
/// more specific class, so forwarding an inner error never reclassifies it.
fn marked(message: String) -> String {
    if message.starts_with(RESOURCE_LIMIT)
        || message.starts_with(ENCRYPTION)
        || message.starts_with(MEASURED_RESOURCE_LIMIT)
    {
        message
    } else {
        format!("{UNSUPPORTED}{message}")
    }
}

/// A rejection that is not classified yet. New code must not call this: use
/// the constructor of the class, or a best-effort helper (see the module
/// documentation).
pub(crate) fn unsupported(message: impl Into<String>) -> String {
    marked(message.into())
}

/// The structure cannot be read without guessing: truncated or overflowing
/// records, offsets and character positions out of range, broken record
/// framing, duplicate or cyclic ownership, unsafe ZIP or XML.
// The reader features adopt the classes and helpers independently, so each
// may be unused in a single-feature build.
#[allow(dead_code)]
pub(crate) fn corrupt(message: impl Into<String>) -> String {
    marked(message.into())
}

/// A well-formed file of a version these readers do not read (Word 6/95,
/// BIFF5/7, PowerPoint 95).
#[allow(dead_code)]
pub(crate) fn out_of_scope_version(message: impl Into<String>) -> String {
    marked(message.into())
}

pub(crate) fn resource(message: impl Into<String>) -> String {
    format!("{RESOURCE_LIMIT}{}", message.into())
}

/// MS-DOC 2.5.2 FibBase encryption flags, MS-XLS 2.4.117 FilePass,
/// MS-PPT 2.3.2 CurrentUserAtom: these readers do not decrypt binary input.
pub(crate) fn encryption(message: impl Into<String>) -> String {
    format!("{ENCRYPTION}{}", message.into())
}

/// How a construct was rendered instead of rejecting the input.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(dead_code)]
pub(crate) enum BestEffort {
    /// An out-of-range value was ignored; the property is unset or default.
    IgnoredValue,
    /// A feature without a visual equivalent produced no output.
    Skipped,
    /// A construct was mapped to the nearest equivalent of the target model.
    Approximated,
}

#[cfg(test)]
thread_local! {
    static BEST_EFFORT: std::cell::RefCell<Vec<(BestEffort, &'static str)>> =
        const { std::cell::RefCell::new(Vec::new()) };
}

#[inline(always)]
#[allow(dead_code, unused_variables)]
fn note(kind: BestEffort, site: &'static str) {
    #[cfg(test)]
    BEST_EFFORT.with(|notes| notes.borrow_mut().push((kind, site)));
}

/// The best-effort decisions taken on this thread since the last call, in
/// order. Tests assert these together with the projected model.
#[cfg(test)]
pub(crate) fn take_best_effort() -> Vec<(BestEffort, &'static str)> {
    BEST_EFFORT.with(|notes| std::mem::take(&mut *notes.borrow_mut()))
}

/// A value inside a correctly framed record is outside its specified range.
/// The caller leaves the property as it was before this record, as Office
/// does. Framing errors (a truncated operand, a bad length) are [`corrupt`].
#[inline(always)]
#[allow(dead_code)]
pub(crate) fn ignored_value(site: &'static str) {
    note(BestEffort::IgnoredValue, site);
}

/// [`ignored_value`] for a property that must hold a value: returns the
/// default the specification gives it.
#[inline(always)]
#[allow(dead_code)]
pub(crate) fn spec_default<T>(site: &'static str, default: T) -> T {
    note(BestEffort::IgnoredValue, site);
    default
}

/// A feature without a visual equivalent in the target model. The caller
/// emits nothing for it and keeps reading; only the stored content is drawn.
/// Not for constructs whose removal moves the content around them.
#[inline(always)]
#[allow(dead_code)]
pub(crate) fn skipped(site: &'static str) {
    note(BestEffort::Skipped, site);
}

/// A construct the target model cannot hold exactly: returns `nearest`, the
/// equivalent the caller chose.
#[inline(always)]
#[allow(dead_code)]
pub(crate) fn approximated<T>(site: &'static str, nearest: T) -> T {
    note(BestEffort::Approximated, site);
    nearest
}

/// A rejection decided at the end of a projection: [`Deferred::defer`] where
/// the construct is met, [`Deferred::admit`] once after the walk.
///
/// Converting a `defer` site means replacing that call with a best-effort
/// helper or a projection. When no `defer` call is left the field and its
/// `admit` go away.
#[derive(Debug, Default, Clone, Copy, PartialEq, Eq)]
#[allow(dead_code)]
pub(crate) struct Deferred {
    pending: bool,
    /// Why the first `defer` happened. Unit tests only.
    #[cfg(test)]
    reason: Option<&'static str>,
}

#[allow(dead_code)]
impl Deferred {
    /// `reason` names the construct, in the wording of a rejection.
    #[inline(always)]
    #[allow(unused_variables)]
    pub(crate) fn defer(&mut self, reason: &'static str) {
        self.pending = true;
        #[cfg(test)]
        self.reason.get_or_insert(reason);
    }

    pub(crate) fn is_pending(&self) -> bool {
        self.pending
    }

    /// The reason of the first `defer`.
    #[cfg(test)]
    pub(crate) fn reason(&self) -> Option<&'static str> {
        self.reason
    }

    /// `Err` with the not-yet-classified rejection `message` when pending.
    pub(crate) fn admit(&self, message: &'static str) -> Result<(), String> {
        if self.pending {
            Err(unsupported(message))
        } else {
            Ok(())
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn forwarding_preserves_resource_and_encryption_classification() {
        assert_eq!(
            unsupported("unknown structure"),
            "UNSUPPORTED:unknown structure"
        );
        let quota = resource("retained model budget");
        assert_eq!(unsupported(unsupported(quota.clone())), quota);
        let measured = "OOXML_RESOURCE_LIMIT:{\"code\":\"ooxml-resource-limit\"}";
        assert_eq!(unsupported(unsupported(measured)), measured);
        let encrypted = encryption("encrypted input");
        assert_eq!(unsupported(unsupported(encrypted.clone())), encrypted);
    }

    #[test]
    fn class_constructors_keep_the_markers_the_native_boundary_decodes() {
        assert_eq!(corrupt("truncated record"), "UNSUPPORTED:truncated record");
        assert_eq!(
            out_of_scope_version("unsupported Word FIB version"),
            "UNSUPPORTED:unsupported Word FIB version"
        );
        // Forwarding an inner limit or encryption error through a class
        // constructor must not turn it into a corrupt-input rejection.
        for inner in [resource("budget"), encryption("encrypted input")] {
            assert_eq!(corrupt(inner.clone()), inner);
            assert_eq!(out_of_scope_version(inner.clone()), inner);
        }
    }

    #[test]
    fn best_effort_helpers_return_the_callers_value_and_record_the_site() {
        take_best_effort();
        ignored_value("invalid value");
        assert_eq!(spec_default("invalid enum", 3u8), 3);
        skipped("feature without output");
        assert_eq!(approximated("unsupported fill", "solid"), "solid");
        assert_eq!(
            take_best_effort(),
            [
                (BestEffort::IgnoredValue, "invalid value"),
                (BestEffort::IgnoredValue, "invalid enum"),
                (BestEffort::Skipped, "feature without output"),
                (BestEffort::Approximated, "unsupported fill"),
            ]
        );
        assert!(take_best_effort().is_empty());
    }

    #[test]
    fn deferred_rejects_once_pending_and_keeps_the_first_reason() {
        let mut deferred = Deferred::default();
        assert!(!deferred.is_pending());
        assert_eq!(deferred.admit("omitted content"), Ok(()));
        deferred.defer("first construct");
        deferred.defer("second construct");
        assert!(deferred.is_pending());
        assert_eq!(deferred.reason(), Some("first construct"));
        assert_eq!(
            deferred.admit("omitted content"),
            Err("UNSUPPORTED:omitted content".into())
        );
    }
}
