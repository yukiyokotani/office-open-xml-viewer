//! DOCX resource compatibility policy, applied AFTER ordinary MCE selection.
//! ECMA-376 Part 3 §9.3 still owns Requires and branch ordering. This narrow
//! library policy substitutes only an authored picture fallback for a selected
//! single-drawing ChartEx payload that the package cannot render. It is not a
//! general MCE rule or an inference about Office's choice-selection algorithm.
//! Office's 2015 ChartEx capability namespace is recognized only while testing
//! `Requires` on a Choice that has this exact raw ChartEx drawing shape. It is
//! not added to the document's general MCE application configuration. The
//! resource verdict still rejects an unrenderable ChartEx part and substitutes
//! its authored fallback.
use crate::document_projector::{
    docx_is_application_defined_extension_element, docx_understands_namespace,
};
use ooxml_common::ns::{is_a_ns, is_c_ns, is_w_ns, is_wp_ns};
use ooxml_common::{bounded_xml::MCE_NS, mce::ChoiceRequiresClassification};
use std::collections::hash_map::{Entry, HashMap};
use std::collections::HashSet;
use std::hash::Hash;

// [MS-ODRAWXML] §§2.1.5, 2.24.1.1 and 2.24.3.76 identify the
// ChartEx part family, chart element and its relationship-id attribute.
pub(crate) const CHARTEX_NS: &str = "http://schemas.microsoft.com/office/drawing/2014/chartex";
pub(crate) const CHARTEX_CAPABILITY_NS: &str =
    "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex";
pub(crate) const CHARTEX_REL: &str =
    "http://schemas.microsoft.com/office/2014/relationships/chartEx";

pub(crate) struct PathLevel(fn(Option<&str>) -> bool, &'static [&'static str]);
impl PathLevel {
    pub(crate) fn matches(&self, namespace: Option<&str>, local: &str) -> bool {
        (self.0)(namespace) && self.1.contains(&local)
    }
}
fn chart_namespace(ns: Option<&str>) -> bool {
    is_c_ns(ns) || ns == Some(CHARTEX_NS)
}
// Only raw DIRECT children count; first matching child wins at each level.
// In particular, extension/ignorable payloads and nested ACs are never searched.
pub(crate) const PATH: [PathLevel; 5] = [
    PathLevel(is_w_ns, &["drawing"]),
    PathLevel(is_wp_ns, &["inline", "anchor"]),
    PathLevel(is_a_ns, &["graphic"]),
    PathLevel(is_a_ns, &["graphicData"]),
    PathLevel(chart_namespace, &["chart"]),
];
#[derive(Default)]
pub(crate) struct Facts {
    pub(crate) direct_children: u8,
    pub(crate) drawing: bool,
    pub(crate) chartex_uri: bool,
    pub(crate) renderable_rid: bool,
}
#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) enum Verdict {
    Parent,
    Renderable,
    Unrenderable,
}
pub(crate) fn verdict(facts: &Facts) -> Verdict {
    if facts.direct_children != 1 || !facts.drawing || !facts.chartex_uri {
        Verdict::Parent
    } else if facts.renderable_rid {
        Verdict::Renderable
    } else {
        Verdict::Unrenderable
    }
}

pub(crate) fn native_verdict(choice: roxmltree::Node, rids: &HashSet<String>) -> Verdict {
    let mut facts = Facts::default();
    let mut elements = choice.children().filter(|node| node.is_element());
    let Some(mut node) = elements.next() else {
        return Verdict::Parent;
    };
    facts.direct_children = if elements.next().is_some() { 2 } else { 1 };
    facts.drawing = PATH[0].matches(node.tag_name().namespace(), node.tag_name().name());
    if !facts.drawing || facts.direct_children != 1 {
        return Verdict::Parent;
    }
    for (index, level) in PATH.iter().enumerate().skip(1) {
        let Some(child) = node.children().find(|child| {
            child.is_element()
                && level.matches(child.tag_name().namespace(), child.tag_name().name())
        }) else {
            break;
        };
        node = child;
        if index == 3 {
            facts.chartex_uri = node.attribute("uri") == Some(CHARTEX_NS);
        }
        if index == 4 {
            facts.renderable_rid = ooxml_common::ns::attr_ns(
                &node,
                ooxml_common::ns::relationships::TRANSITIONAL,
                ooxml_common::ns::relationships::STRICT,
                "id",
            )
            .is_some_and(|rid| rids.contains(rid));
        }
    }
    verdict(&facts)
}

/// Select with the parent's application configuration, except that the 2015
/// capability token is accepted for one Choice only after its raw direct-child
/// shape has been classified as the narrow ChartEx drawing path above.
pub(crate) fn select_native_alternate_content<'a, 'i>(
    alternate: roxmltree::Node<'a, 'i>,
    rids: &HashSet<String>,
    understood: &dyn Fn(&str) -> bool,
    mce_understood: &dyn Fn(&str) -> bool,
) -> Option<roxmltree::Node<'a, 'i>> {
    for choice in alternate.children().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(MCE_NS)
            && node.tag_name().name() == "Choice"
    }) {
        let parent_classification =
            ooxml_common::mce::classify_choice_requires(choice.attribute("Requires"), |prefix| {
                match choice.lookup_namespace_uri(Some(prefix)) {
                    None => ooxml_common::mce::RequiredNamespaceSupport::Unresolved,
                    Some(namespace) if understood(namespace) => {
                        ooxml_common::mce::RequiredNamespaceSupport::Understood
                    }
                    Some(_) => ooxml_common::mce::RequiredNamespaceSupport::Unsupported,
                }
            });
        if parent_classification == ChoiceRequiresClassification::Understood {
            return Some(choice);
        }
        if parent_classification != ChoiceRequiresClassification::Unsupported
            || native_verdict(choice, rids) == Verdict::Parent
        {
            continue;
        }
        let local_classification =
            ooxml_common::mce::classify_choice_requires(choice.attribute("Requires"), |prefix| {
                match choice.lookup_namespace_uri(Some(prefix)) {
                    None => ooxml_common::mce::RequiredNamespaceSupport::Unresolved,
                    Some(namespace)
                        if understood(namespace) || namespace == CHARTEX_CAPABILITY_NS =>
                    {
                        ooxml_common::mce::RequiredNamespaceSupport::Understood
                    }
                    Some(_) => ooxml_common::mce::RequiredNamespaceSupport::Unsupported,
                }
            });
        if local_classification == ChoiceRequiresClassification::Understood
            && native_branch_must_understand(choice, mce_understood)
        {
            return Some(choice);
        }
    }
    alternate.children().find(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(MCE_NS)
            && node.tag_name().name() == "Fallback"
    })
}

/// A resource override must not newly select a Fallback that the MCE processor
/// would reject. In that case retaining the parent's selected Choice preserves
/// native/streaming parity and the parent's fail-closed drawing result. The
/// selected-ChartEx preflight below applies the same per-element predicate.
pub(crate) fn native_branch_must_understand(
    branch: roxmltree::Node,
    understood: &dyn Fn(&str) -> bool,
) -> bool {
    branch
        .attribute((MCE_NS, "MustUnderstand"))
        .is_none_or(|prefixes| {
            prefixes.split_whitespace().all(|prefix| {
                branch
                    .lookup_namespace_uri(Some(prefix))
                    .is_some_and(understood)
            })
        })
}

/// Native body preflight for an already selected run-level ChartEx Choice.
///
/// ECMA-376 Part 3 §9.3 selects the branch and §9.4 processes it; §9.1 and the
/// Annex A.2.5 example require a MustUnderstand mismatch on that processed path
/// to be signaled (§9.4 item 5 alone does not spell out that substep). The
/// selected Choice and every effective descendant are checked before the
/// resource override, so an unrenderable part cannot hide the mismatch behind
/// its picture fallback. Inherited `mc:Ignorable`/`mc:ProcessContent`,
/// selected nested branches (including their authored resource substitutions)
/// and opaque extension lists follow the streamed projector, so ignored,
/// unselected and opaque payload is never checked. Run-level selections retain
/// the native drawing arm's namespace configuration, including for a nested
/// run; this is not a general native/streaming MCE configuration unification.
///
/// Scope is library policy: only a run-level AlternateContent selected with the
/// parser arm's configuration into the exact ChartEx shape is a target. Other
/// native MCE MustUnderstand processing and the header/footer/note stories are
/// unchanged. The caller owns how the mismatch is reported.
///
/// One borrowed depth-first pass over the parsed part: the frame stack is
/// bounded by the depth already enforced by `parse_guarded`, and the counted
/// directive maps hold only `&str` slices for the active path.
pub(crate) fn validate_selected_chartex_must_understand(
    root: roxmltree::Node,
    rids: &HashSet<String>,
) -> Result<(), String> {
    let mut directives = ActiveMceDirectives::default();
    let mut frames = Vec::new();
    let mut next = Some((root, false));
    while let Some((node, validate)) = next.take() {
        if let Some(frame) = enter_effective_element(node, validate, rids, &mut directives)? {
            frames.push(frame);
        }
        while let Some(frame) = frames.last_mut() {
            if let Some(child) = frame.next {
                frame.next = if frame.siblings {
                    child.next_sibling()
                } else {
                    frame.resource_fallback.take()
                };
                next = Some((child, frame.validate));
                break;
            }
            let element = frame.element;
            frames.pop();
            directives.update(element, false);
        }
    }
    Ok(())
}

struct EffectiveFrame<'a, 'input> {
    element: roxmltree::Node<'a, 'input>,
    next: Option<roxmltree::Node<'a, 'input>>,
    /// False for AlternateContent, whose only effective child is the selection.
    siblings: bool,
    // The selected Choice must be validated first, before substitution can
    // visit the authored fallback actually processed by the stream projector.
    resource_fallback: Option<roxmltree::Node<'a, 'input>>,
    validate: bool,
}

fn enter_effective_element<'a, 'input>(
    node: roxmltree::Node<'a, 'input>,
    validate: bool,
    rids: &HashSet<String>,
    directives: &mut ActiveMceDirectives<'a>,
) -> Result<Option<EffectiveFrame<'a, 'input>>, String> {
    if !node.is_element() {
        return Ok(None);
    }
    let namespace = node.tag_name().namespace();
    let local = node.tag_name().name();
    // §§8 and 9.1: application-defined extension payload is opaque, including
    // any MCE attributes it carries.
    if docx_is_application_defined_extension_element(namespace, local) {
        return Ok(None);
    }
    directives.update(node, true);
    if namespace.is_some_and(|namespace| {
        !docx_understands_namespace(namespace) && directives.ignores(namespace, local)
    }) {
        directives.update(node, false);
        return Ok(None);
    }
    if validate && !native_branch_must_understand(node, &docx_understands_namespace) {
        // Matches the streamed projector's selected-ChartEx diagnostic.
        return Err(
            "document MCE MustUnderstand namespace is not understood in selected ChartEx Choice"
                .to_string(),
        );
    }
    if namespace != Some(MCE_NS) || local != "AlternateContent" {
        return Ok(Some(EffectiveFrame {
            element: node,
            next: node.first_child(),
            siblings: true,
            resource_fallback: None,
            validate,
        }));
    }
    let run_level = node.parent_element().is_some_and(|parent| {
        is_w_ns(parent.tag_name().namespace()) && parent.tag_name().name() == "r"
    });
    // A run-level AC uses the parser arm's exact selection; elsewhere the
    // document configuration matches the streamed projector's selection.
    let understood: fn(&str) -> bool = if run_level {
        crate::parser::docx_understands_drawing_ns
    } else {
        docx_understands_namespace
    };
    let selected =
        select_native_alternate_content(node, rids, &understood, &docx_understands_namespace);
    let target = run_level
        && selected.is_some_and(|branch| {
            branch.tag_name().name() == "Choice" && native_verdict(branch, rids) != Verdict::Parent
        });
    // The target container itself is on the processed path. Earlier frames
    // outside this bounded ChartEx seam retain the native parser's policy.
    if target && !validate && !native_branch_must_understand(node, &docx_understands_namespace) {
        return Err(
            "document MCE MustUnderstand namespace is not understood in selected ChartEx Choice"
                .to_string(),
        );
    }
    let resource_fallback = selected
        .filter(|branch| {
            branch.tag_name().name() == "Choice"
                && native_verdict(*branch, rids) == Verdict::Unrenderable
        })
        .and_then(|_| {
            node.children().find(|branch| {
                branch.is_element()
                    && branch.tag_name().namespace() == Some(MCE_NS)
                    && branch.tag_name().name() == "Fallback"
                    && native_branch_must_understand(*branch, &docx_understands_namespace)
            })
        });
    Ok(Some(EffectiveFrame {
        element: node,
        next: selected,
        siblings: false,
        resource_fallback,
        validate: validate || target,
    }))
}

/// Ignorable/ProcessContent declarations active on the traversal path. Counts
/// let leaving an element remove exactly what entering it added, so lookups
/// stay constant-time without per-scope set copies or ancestor rescans.
#[derive(Default)]
struct ActiveMceDirectives<'a> {
    ignorable: HashMap<&'a str, usize>,
    process_content: HashMap<(&'a str, &'a str), usize>,
}

impl<'a> ActiveMceDirectives<'a> {
    fn update(&mut self, node: roxmltree::Node<'a, '_>, enter: bool) {
        let tokens = |name: &str| {
            node.attribute((MCE_NS, name))
                .unwrap_or_default()
                .split_whitespace()
        };
        for prefix in tokens("Ignorable") {
            if let Some(namespace) = node.lookup_namespace_uri(Some(prefix)) {
                count_directive(&mut self.ignorable, namespace, enter);
            }
        }
        for name in tokens("ProcessContent") {
            if let Some((prefix, local)) = name.split_once(':') {
                if let Some(namespace) = node.lookup_namespace_uri(Some(prefix)) {
                    count_directive(&mut self.process_content, (namespace, local), enter);
                }
            }
        }
    }

    fn ignores(&self, namespace: &'a str, local: &'a str) -> bool {
        self.ignorable.contains_key(namespace)
            && !self.process_content.contains_key(&(namespace, local))
            && !self.process_content.contains_key(&(namespace, "*"))
    }
}

fn count_directive<K: Eq + Hash>(counts: &mut HashMap<K, usize>, key: K, enter: bool) {
    if enter {
        *counts.entry(key).or_default() += 1;
    } else if let Entry::Occupied(mut entry) = counts.entry(key) {
        *entry.get_mut() -= 1;
        if *entry.get() == 0 {
            entry.remove();
        }
    }
}
