//! Markup Compatibility and Extensibility (MCE) — the `<mc:AlternateContent>`
//! reference processing model shared by the docx, pptx and xlsx parsers.
//!
//! ECMA-376 Part 3 (Markup Compatibility and Extensibility), §9.3 "Step 2:
//! Processing the AlternateContent, Choice and Fallback Elements" defines which
//! branch of an `<mc:AlternateContent>` a consumer selects, verbatim:
//!
//! > A Choice element shall be marked as selected if the following conditions
//! > are satisfied:
//! > 1) Each of the namespaces specified by the Requires attribute of this
//! >    element is included in the given application configuration;
//! > 2) No preceding sibling Choice element is marked as selected; and
//! > 3) The element is not a descendant of an application-defined extension
//! >    element.
//! > A Fallback element shall be marked as selected if the following conditions
//! > are satisfied:
//! > 1) No preceding sibling Choice element is marked as selected; and
//! > 2) The element is not a descendant of an application-defined extension
//! >    element.
//!
//! The "given application configuration" is the set of namespaces the consumer
//! understands — i.e. those it has a handler that produces renderable output
//! for. Each parser passes its own membership test as the `understood`
//! predicate; the algorithm here is format-agnostic. Condition (3) concerns
//! MCE's own nested-extension elements, which none of these host schemas emit
//! around Choice/Fallback, so it needs no special handling: we only ever walk
//! the direct Choice/Fallback children of one AlternateContent.
//!
//! This unifies three previously divergent local behaviours (issue #787): docx
//! already implemented §9.3, pptx guessed via output-emptiness (never inspected
//! `Requires`), and xlsx always took the Choice (never the Fallback), so an
//! un-understood Choice with a renderable Fallback silently dropped content.

use crate::bounded_xml::MCE_NS;
use roxmltree::Node;

/// Namespace support result for one prefix named by `Choice/@Requires`.
#[doc(hidden)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RequiredNamespaceSupport {
    Unresolved,
    Unsupported,
    Understood,
}

/// Format-neutral classification of one `Choice/@Requires` value.
#[doc(hidden)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChoiceRequiresClassification {
    Missing,
    Blank,
    Unresolved,
    Unsupported,
    Understood,
}

/// Classify the §9.3 all-required-namespaces predicate without deciding how a
/// host treats non-conformant missing or blank attributes.
#[doc(hidden)]
pub fn classify_choice_requires(
    requires: Option<&str>,
    mut support_for_prefix: impl FnMut(&str) -> RequiredNamespaceSupport,
) -> ChoiceRequiresClassification {
    let Some(requires) = requires else {
        return ChoiceRequiresClassification::Missing;
    };
    let mut prefixes = requires.split_whitespace().peekable();
    if prefixes.peek().is_none() {
        return ChoiceRequiresClassification::Blank;
    }
    let mut unsupported = false;
    let mut unresolved = false;
    for prefix in prefixes {
        match support_for_prefix(prefix) {
            RequiredNamespaceSupport::Unresolved => unresolved = true,
            RequiredNamespaceSupport::Unsupported => unsupported = true,
            RequiredNamespaceSupport::Understood => {}
        }
    }
    if unresolved {
        ChoiceRequiresClassification::Unresolved
    } else if unsupported {
        ChoiceRequiresClassification::Unsupported
    } else {
        ChoiceRequiresClassification::Understood
    }
}

/// Select the active branch of an `<mc:AlternateContent>` per ECMA-376 Part 3
/// §9.3 (Step 2), returning the selected `<mc:Choice>` / `<mc:Fallback>` element
/// node, or `None` when neither a selectable Choice nor a Fallback exists.
///
/// - A `<mc:Choice>` is selected iff EVERY namespace named by its `Requires`
///   attribute is understood AND no preceding sibling Choice was already
///   selected. `Requires` is a whitespace-delimited list of namespace *prefixes*
///   (Part 3 §7.6), each resolved to a URI against the element's in-scope
///   `xmlns` declarations via `lookup_namespace_uri`; the URI is then tested
///   with `understood`. Resolving prefixes (not raw strings) means both the
///   Transitional and Strict conformance classes work, and an arbitrary
///   producer prefix (`cx`, `cx1`, …) binds to the same URI.
/// - The schema requires `Requires` to list ≥1 prefix (§7.6); a missing or
///   whitespace-only value is non-conformant and can never satisfy §9.3(1)
///   ("*each* of the namespaces … is included"), so such a Choice is never
///   selected.
/// - If no Choice is selected, the `<mc:Fallback>` (if present) is selected.
///
/// `understood(ns_uri) -> bool` is the consumer's membership test for its
/// application configuration: return `true` only for namespaces the caller can
/// actually render, so an un-understood Choice correctly yields the Fallback.
pub fn select_alternate_content<'a, 'i>(
    ac: Node<'a, 'i>,
    understood: &dyn Fn(&str) -> bool,
) -> Option<Node<'a, 'i>> {
    // Part 3 Annex A.1.7 permits ignorable foreign children here. A foreign
    // `Choice` or `Fallback` cannot pre-empt the actual mc branch.
    for choice in ac.children().filter(|n| {
        n.is_element()
            && n.tag_name().namespace() == Some(MCE_NS)
            && n.tag_name().name() == "Choice"
    }) {
        let classification = classify_choice_requires(choice.attribute("Requires"), |prefix| {
            match choice.lookup_namespace_uri(Some(prefix)) {
                None => RequiredNamespaceSupport::Unresolved,
                Some(namespace) if understood(namespace) => RequiredNamespaceSupport::Understood,
                Some(_) => RequiredNamespaceSupport::Unsupported,
            }
        });
        if classification == ChoiceRequiresClassification::Understood {
            // §9.3(2): first matching Choice wins; later Choices are ignored.
            return Some(choice);
        }
    }
    // §9.3: no Choice selected → the Fallback (if any) is selected.
    ac.children().find(|n| {
        n.is_element()
            && n.tag_name().namespace() == Some(MCE_NS)
            && n.tag_name().name() == "Fallback"
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Parse an `<mc:AlternateContent>` document, run §9.3 selection with the
    /// given understood-URI set, and return the selected branch's `id` attribute
    /// (each Choice/Fallback in the fixtures is tagged with a distinct `id`), or
    /// `None` when nothing is selected.
    fn select_id(xml: &str, understood: &[&str]) -> Option<String> {
        let doc = roxmltree::Document::parse(xml).unwrap();
        let ac = doc.root_element();
        let pred = |ns: &str| understood.contains(&ns);
        select_alternate_content(ac, &pred).and_then(|n| n.attribute("id").map(str::to_string))
    }

    // Fixtures bind prefixes to distinct URIs (mixing the Transitional-style and
    // an unrelated "urn:" form to prove URI — not prefix-string — matching).
    const NS: &str = concat!(
        r#"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" "#,
        r#"xmlns:k1="urn:known:one" "#,
        r#"xmlns:k2="urn:known:two" "#,
        r#"xmlns:kalt="urn:known:one" "#, // second prefix, SAME URI as k1
        r#"xmlns:u1="urn:unknown:one""#,
    );

    #[test]
    fn understood_single_choice_selected() {
        let xml = format!(
            r#"<mc:AlternateContent {NS}>
                 <mc:Choice id="c1" Requires="k1"><x/></mc:Choice>
                 <mc:Fallback id="fb"><x/></mc:Fallback>
               </mc:AlternateContent>"#
        );
        assert_eq!(select_id(&xml, &["urn:known:one"]).as_deref(), Some("c1"));
    }

    #[test]
    fn unknown_only_choice_falls_back() {
        let xml = format!(
            r#"<mc:AlternateContent {NS}>
                 <mc:Choice id="c1" Requires="u1"><x/></mc:Choice>
                 <mc:Fallback id="fb"><x/></mc:Fallback>
               </mc:AlternateContent>"#
        );
        assert_eq!(select_id(&xml, &["urn:known:one"]).as_deref(), Some("fb"));
    }

    #[test]
    fn multi_namespace_requires_needs_all_understood() {
        let xml = format!(
            r#"<mc:AlternateContent {NS}>
                 <mc:Choice id="c1" Requires="k1 u1"><x/></mc:Choice>
                 <mc:Fallback id="fb"><x/></mc:Fallback>
               </mc:AlternateContent>"#
        );
        // Only k1 understood → the "k1 u1" Choice is NOT all-understood → Fallback.
        assert_eq!(select_id(&xml, &["urn:known:one"]).as_deref(), Some("fb"));
        // Both understood → the Choice is selected.
        assert_eq!(
            select_id(&xml, &["urn:known:one", "urn:unknown:one"]).as_deref(),
            Some("c1")
        );
    }

    #[test]
    fn second_choice_selected_when_first_not_understood() {
        let xml = format!(
            r#"<mc:AlternateContent {NS}>
                 <mc:Choice id="c1" Requires="u1"><x/></mc:Choice>
                 <mc:Choice id="c2" Requires="k1"><x/></mc:Choice>
                 <mc:Fallback id="fb"><x/></mc:Fallback>
               </mc:AlternateContent>"#
        );
        assert_eq!(select_id(&xml, &["urn:known:one"]).as_deref(), Some("c2"));
    }

    #[test]
    fn foreign_namesakes_do_not_select_choice_or_fallback() {
        let choice = format!(
            r#"<mc:AlternateContent {NS} mc:Ignorable="u1">
                 <u1:Choice id="foreign" Requires="k1"/>
                 <mc:Choice id="effective" Requires="k1"/>
                 <mc:Fallback id="fallback"/>
               </mc:AlternateContent>"#
        );
        assert_eq!(
            select_id(&choice, &["urn:known:one"]).as_deref(),
            Some("effective")
        );
        let fallback = format!(
            r#"<mc:AlternateContent {NS} mc:Ignorable="u1">
                 <mc:Choice id="unsupported" Requires="u1"/>
                 <u1:Fallback id="foreign"/>
                 <mc:Fallback id="effective"/>
               </mc:AlternateContent>"#
        );
        assert_eq!(select_id(&fallback, &[]).as_deref(), Some("effective"));
    }

    #[test]
    fn first_understood_choice_wins_over_later() {
        let xml = format!(
            r#"<mc:AlternateContent {NS}>
                 <mc:Choice id="c1" Requires="k1"><x/></mc:Choice>
                 <mc:Choice id="c2" Requires="k2"><x/></mc:Choice>
                 <mc:Fallback id="fb"><x/></mc:Fallback>
               </mc:AlternateContent>"#
        );
        // §9.3(2): the FIRST understood Choice is selected even when a later
        // Choice is also understood.
        assert_eq!(
            select_id(&xml, &["urn:known:one", "urn:known:two"]).as_deref(),
            Some("c1")
        );
    }

    #[test]
    fn prefix_resolves_by_uri_not_spelling() {
        // `kalt` is a different prefix bound to the SAME URI as `k1`. Selection
        // is by resolved URI, so a Choice `Requires="kalt"` is understood.
        let xml = format!(
            r#"<mc:AlternateContent {NS}>
                 <mc:Choice id="c1" Requires="kalt"><x/></mc:Choice>
                 <mc:Fallback id="fb"><x/></mc:Fallback>
               </mc:AlternateContent>"#
        );
        assert_eq!(select_id(&xml, &["urn:known:one"]).as_deref(), Some("c1"));
    }

    #[test]
    fn missing_requires_is_never_selected() {
        // Non-conformant: no Requires attribute at all → can't be all-understood.
        let xml = format!(
            r#"<mc:AlternateContent {NS}>
                 <mc:Choice id="c1"><x/></mc:Choice>
                 <mc:Fallback id="fb"><x/></mc:Fallback>
               </mc:AlternateContent>"#
        );
        assert_eq!(select_id(&xml, &["urn:known:one"]).as_deref(), Some("fb"));
    }

    #[test]
    fn blank_requires_is_never_selected() {
        // Non-conformant: whitespace-only Requires → empty prefix list → not
        // selectable (an empty list can never be "each … is included").
        let xml = format!(
            r#"<mc:AlternateContent {NS}>
                 <mc:Choice id="c1" Requires="   "><x/></mc:Choice>
                 <mc:Fallback id="fb"><x/></mc:Fallback>
               </mc:AlternateContent>"#
        );
        assert_eq!(select_id(&xml, &["urn:known:one"]).as_deref(), Some("fb"));
    }

    #[test]
    fn unbound_requires_prefix_is_never_selected() {
        let xml = format!(
            r#"<mc:AlternateContent {NS}>
                 <mc:Choice id="c1" Requires="missing"><x/></mc:Choice>
                 <mc:Fallback id="fb"><x/></mc:Fallback>
               </mc:AlternateContent>"#
        );
        assert_eq!(select_id(&xml, &["urn:known:one"]).as_deref(), Some("fb"));
    }

    #[test]
    fn no_selectable_choice_and_no_fallback_returns_none() {
        let xml = format!(
            r#"<mc:AlternateContent {NS}>
                 <mc:Choice id="c1" Requires="u1"><x/></mc:Choice>
               </mc:AlternateContent>"#
        );
        assert_eq!(select_id(&xml, &["urn:known:one"]), None);
    }

    #[test]
    fn understood_choice_with_no_fallback_still_selected() {
        let xml = format!(
            r#"<mc:AlternateContent {NS}>
                 <mc:Choice id="c1" Requires="k1"><x/></mc:Choice>
               </mc:AlternateContent>"#
        );
        assert_eq!(select_id(&xml, &["urn:known:one"]).as_deref(), Some("c1"));
    }

    #[test]
    fn empty_understood_set_always_falls_back() {
        let xml = format!(
            r#"<mc:AlternateContent {NS}>
                 <mc:Choice id="c1" Requires="k1"><x/></mc:Choice>
                 <mc:Choice id="c2" Requires="k2"><x/></mc:Choice>
                 <mc:Fallback id="fb"><x/></mc:Fallback>
               </mc:AlternateContent>"#
        );
        assert_eq!(select_id(&xml, &[]).as_deref(), Some("fb"));
    }

    #[test]
    fn requires_classification_distinguishes_host_mapping_cases() {
        assert_eq!(
            classify_choice_requires(None, |_| RequiredNamespaceSupport::Understood),
            ChoiceRequiresClassification::Missing
        );
        assert_eq!(
            classify_choice_requires(Some("  "), |_| RequiredNamespaceSupport::Understood),
            ChoiceRequiresClassification::Blank
        );
        assert_eq!(
            classify_choice_requires(Some("k missing"), |prefix| {
                if prefix == "missing" {
                    RequiredNamespaceSupport::Unresolved
                } else {
                    RequiredNamespaceSupport::Understood
                }
            }),
            ChoiceRequiresClassification::Unresolved
        );
        assert_eq!(
            classify_choice_requires(Some("k u"), |prefix| {
                if prefix == "u" {
                    RequiredNamespaceSupport::Unsupported
                } else {
                    RequiredNamespaceSupport::Understood
                }
            }),
            ChoiceRequiresClassification::Unsupported
        );
        assert_eq!(
            classify_choice_requires(Some("k1 k2"), |_| { RequiredNamespaceSupport::Understood }),
            ChoiceRequiresClassification::Understood
        );
    }
}
