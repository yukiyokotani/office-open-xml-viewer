//! Bounded, namespace-preserving projection of `word/document.xml` into
//! complete logical body blocks.
//!
//! WordprocessingML is not page-random-access: sections, fields, notes, and
//! pagination depend on the complete preceding story. This projector does not
//! pretend otherwise. It removes only the whole-part XML arena: a first pass
//! validates the part and records compact cross-block facts, while a second
//! pass can hand one complete paragraph/table/section block at a time to the
//! existing semantic parser.

use crate::chartex_choice::{self, Facts, Verdict, CHARTEX_NS, PATH};
use std::collections::HashSet;
use std::io::BufRead;
use std::rc::Rc;

use ooxml_common::bounded_xml::{
    self, BoundedXmlError, BoundedXmlReadError, BoundedXmlReader, MceScope as StreamedMceScope,
    NamespaceContext,
};
use ooxml_common::depth::MAX_XML_DEPTH;
use ooxml_common::ns::is_w_ns;
use ooxml_common::package_session::PackageLimitReporter;
use ooxml_common::resource::{HardResourceLimitKind, HARD_MAX_DOCX_BODY_BLOCK_XML_BYTES};
use quick_xml::events::{BytesStart, Event};
use quick_xml::XmlVersion;

const DOCUMENT_PART: &str = "word/document.xml";
/// One lexical event (a start tag with all its attributes, a text node, a
/// comment) may be as large as the logical body block that contains it, whose
/// projection is already bounded by `HARD_MAX_DOCX_BODY_BLOCK_XML_BYTES`. A
/// smaller per-event ceiling would reject blocks that ceiling admits, and Office
/// writes such tokens: a legacy VML shape (ECMA-376 Part 4 §19.1.2.19,
/// CT_Shape) carries the DrawingML equivalent as a base64-encoded package in its
/// `o:gfxdata` attribute, so a single start tag grows with the shape's content
/// and the format sets no bound on it. The event reader still stops one hostile
/// or unterminated token at this ceiling rather than at the part size, and it
/// applies equally to events that are not retained (an unselected
/// `mc:Fallback`, a comment).
pub(crate) const STREAMED_XML_EVENT_BYTES: usize = HARD_MAX_DOCX_BODY_BLOCK_XML_BYTES as usize;
// Library resource policy: only retained namespace declarations and resolved
// MCE state use this budget; ordinary attribute payload uses the event/block
// ceiling, including when Part 3 §9.4 removes physical ancestors.
const STREAMED_XML_CONTEXT_BYTES: usize = 4 * 1024 * 1024;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct ProjectedBodyBlock {
    pub(crate) ordinal: usize,
    /// Byte offset of the source block start in `word/document.xml`. This is a
    /// stable parser-owned identity used by logical-table acquisition facts.
    pub(crate) source_offset: usize,
    pub(crate) local_name: String,
    pub(crate) xml: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct DocumentBodyPlan {
    pub(crate) cover_break_after: Vec<bool>,
}

struct ElementFrame {
    context: Rc<NamespaceContext>,
    mce_scope: Rc<StreamedMceScope>,
    kind: ProcessedElementKind,
    visible: bool,
}

enum ProcessedElementKind {
    Retained {
        namespace: Option<String>,
        local_name: String,
        opaque: bool,
    },
    Unwrapped,
    Ignored,
    AlternateContent {
        selected_branch: bool,
        seen_choice: bool,
        seen_fallback: bool,
        rollback_to: Option<usize>,
        choice_end_len: usize,
    },
    AlternateBranch {
        selected: bool,
        observer: Option<ChoiceObserver>,
        filter_direct_children: bool,
    },
}

impl ElementFrame {
    fn children_are_visible(&self) -> bool {
        self.visible
            && match &self.kind {
                ProcessedElementKind::Ignored
                | ProcessedElementKind::AlternateBranch {
                    selected: false, ..
                } => false,
                ProcessedElementKind::AlternateBranch {
                    observer: Some(observer),
                    ..
                } if observer.discarded => false,
                _ => true,
            }
    }
}

// Fixed-size path state per selected Choice. `level` never retreats, enforcing
// first matching direct child; `open_level` retreats on End, preventing a later
// sibling container from supplying descendants of the first container.
struct ChoiceObserver {
    base_depth: usize,
    capture_checkpoint: usize,
    level: u8,
    open_level: u8,
    facts: Facts,
    rid: Option<String>,
    capability_provisional: bool,
    discarded: bool,
    deferred_must_understand: bool,
}

#[derive(Clone, Copy)]
struct ChoiceSupport {
    understood: bool,
    capability_provisional: bool,
}

#[derive(Default)]
struct SdtFrame {
    open_depth: usize,
    properties_depth: Option<usize>,
    content_depth: Option<usize>,
    cover_pages: bool,
    last_block_ordinal: Option<usize>,
}

struct Capture {
    root_depth: usize,
    source_offset: usize,
    local_name: String,
    xml: Vec<u8>,
}

/// Streaming body projector. `next_block` retains at most one complete logical
/// block and bounded active namespace state. Block-level `w:sdt` wrappers are
/// transparent, matching `element_children_flat`; their cached `w:sdtContent`
/// children become independent blocks.
pub(crate) struct DocumentBodyProjector<R: BufRead> {
    reader: BoundedXmlReader<R>,
    reporter: Option<PackageLimitReporter>,
    renderable_chartex_rids: Rc<HashSet<String>>,
    frames: Vec<ElementFrame>,
    depth: usize,
    body_content_depth: Option<usize>,
    body_seen: bool,
    body_closed: bool,
    sdts: Vec<SdtFrame>,
    capture: Option<Capture>,
    cover_break_after: Vec<bool>,
    discarded_provisional_choices: usize,
    finished: bool,
}

impl<R: BufRead> DocumentBodyProjector<R> {
    pub(crate) fn new(
        source: R,
        reporter: Option<PackageLimitReporter>,
        renderable_chartex_rids: Rc<HashSet<String>>,
    ) -> Self {
        Self {
            reader: BoundedXmlReader::new(source, STREAMED_XML_EVENT_BYTES, "document body"),
            reporter,
            renderable_chartex_rids,
            frames: Vec::new(),
            depth: 0,
            body_content_depth: None,
            body_seen: false,
            body_closed: false,
            sdts: Vec::new(),
            capture: None,
            cover_break_after: Vec::new(),
            discarded_provisional_choices: 0,
            finished: false,
        }
    }

    pub(crate) fn next_block(&mut self) -> Result<Option<ProjectedBodyBlock>, String> {
        if self.finished {
            return Ok(None);
        }
        let mut event_buffer = Vec::new();
        loop {
            let bounded = self
                .reader
                .read_event_into(&mut event_buffer)
                .map_err(|error| self.map_read_error(error))?;
            let source_offset = usize::try_from(bounded.span.start)
                .map_err(|_| "document body source offset exceeds this platform".to_string())?;
            let namespace = bounded.namespace.as_deref();
            match bounded.event {
                Event::Start(start) => {
                    let inherited = self
                        .frames
                        .last()
                        .map(|frame| Rc::clone(&frame.context))
                        .unwrap_or_else(NamespaceContext::root);
                    let context = NamespaceContext::derive(
                        &start,
                        &inherited,
                        STREAMED_XML_CONTEXT_BYTES,
                        "document body",
                    )
                    .map_err(|error| {
                        self.map_bounded_error(error, HardResourceLimitKind::XmlContextBytes)
                    })?;
                    let next_depth = self.depth.saturating_add(1);
                    if next_depth > MAX_XML_DEPTH as usize {
                        return Err(self.report_limit(
                            HardResourceLimitKind::XmlNestingDepth,
                            MAX_XML_DEPTH as usize,
                            next_depth,
                        ));
                    }

                    let mut visible = self
                        .frames
                        .last()
                        .is_none_or(ElementFrame::children_are_visible);
                    if self.observe_choice_start(namespace, &start, false) {
                        visible = false;
                    }
                    let (kind, mce_scope) =
                        self.classify_processed_element(namespace, &start, &context, visible)?;
                    if self.capture.is_some() {
                        if let ProcessedElementKind::Retained { opaque, .. } = &kind {
                            if visible {
                                self.append_retained_element(
                                    &start, &context, &mce_scope, *opaque, false,
                                )?;
                            }
                        }
                    } else if let ProcessedElementKind::Retained { opaque, .. } = &kind {
                        if visible
                            && self.handle_start(namespace, &start, next_depth, source_offset)?
                        {
                            self.append_retained_element(
                                &start, &context, &mce_scope, *opaque, false,
                            )?;
                        }
                    }
                    self.frames.push(ElementFrame {
                        context,
                        mce_scope,
                        kind,
                        visible,
                    });
                    self.depth = next_depth;
                }
                Event::Empty(empty) => {
                    let inherited = self
                        .frames
                        .last()
                        .map(|frame| Rc::clone(&frame.context))
                        .unwrap_or_else(NamespaceContext::root);
                    let context = NamespaceContext::derive(
                        &empty,
                        &inherited,
                        STREAMED_XML_CONTEXT_BYTES,
                        "document body",
                    )
                    .map_err(|error| {
                        self.map_bounded_error(error, HardResourceLimitKind::XmlContextBytes)
                    })?;
                    let mut visible = self
                        .frames
                        .last()
                        .is_none_or(ElementFrame::children_are_visible);
                    if self.observe_choice_start(namespace, &empty, true) {
                        visible = false;
                    }
                    let (kind, mce_scope) =
                        self.classify_processed_element(namespace, &empty, &context, visible)?;
                    self.reject_empty_provisional_choice(&kind);
                    if self.capture.is_some() {
                        match kind {
                            ProcessedElementKind::Retained { opaque, .. } if visible => {
                                self.append_retained_element(
                                    &empty, &context, &mce_scope, opaque, true,
                                )?;
                            }
                            ProcessedElementKind::AlternateContent {
                                seen_choice: false, ..
                            } => {
                                return Err(
                                    "document MCE AlternateContent must contain at least one Choice"
                                        .to_string(),
                                );
                            }
                            _ => {}
                        }
                    } else {
                        match kind {
                            ProcessedElementKind::Retained { opaque, .. } if visible => {
                                if self.handle_empty(namespace, &empty, source_offset)? {
                                    self.append_retained_element(
                                        &empty, &context, &mce_scope, opaque, true,
                                    )?;
                                    if let Some(block) = self.finish_capture()? {
                                        return Ok(Some(block));
                                    }
                                }
                            }
                            ProcessedElementKind::AlternateContent {
                                seen_choice: false, ..
                            } => {
                                return Err(
                                    "document MCE AlternateContent must contain at least one Choice"
                                        .to_string(),
                                );
                            }
                            _ => {}
                        }
                    }
                }
                Event::End(end) => {
                    self.observe_choice_end()?;
                    if let Some(capture) = self.capture.as_ref() {
                        let closes_capture = self.depth == capture.root_depth;
                        let retained = self.discarded_provisional_choices == 0
                            && self.frames.last().is_some_and(|frame| {
                                frame.visible
                                    && matches!(frame.kind, ProcessedElementKind::Retained { .. })
                            });
                        self.handle_end(namespace, end.local_name().as_ref())?;
                        if retained {
                            self.append_capture(Event::End(end))?;
                        }
                        self.depth = self.depth.saturating_sub(1);
                        self.frames.pop();
                        if closes_capture {
                            if let Some(block) = self.finish_capture()? {
                                return Ok(Some(block));
                            }
                        }
                    } else {
                        self.handle_end(namespace, end.local_name().as_ref())?;
                        self.depth = self.depth.saturating_sub(1);
                        self.frames.pop();
                    }
                }
                Event::Eof => {
                    self.finished = true;
                    if self.capture.is_some() {
                        return Err(
                            "word/document.xml ended inside a logical body block".to_string()
                        );
                    }
                    if !self.body_seen {
                        return Err("word/document.xml: no <w:body> element".to_string());
                    }
                    if !self.body_closed {
                        return Err("word/document.xml ended before </w:body>".to_string());
                    }
                    return Ok(None);
                }
                event => {
                    let retained_content = self.frames.last().is_some_and(|frame| {
                        frame.visible
                            && matches!(
                                frame.kind,
                                ProcessedElementKind::Retained { .. }
                                    | ProcessedElementKind::Unwrapped
                                    | ProcessedElementKind::AlternateBranch {
                                        selected: true,
                                        filter_direct_children: false,
                                        ..
                                    }
                            )
                    });
                    if self.capture.is_some() && retained_content {
                        self.append_capture(event)?;
                    }
                }
            }
        }
    }

    fn observe_choice_start(
        &mut self,
        namespace: Option<&str>,
        element: &BytesStart<'_>,
        empty: bool,
    ) -> bool {
        let local = element.local_name();
        let Ok(local) = std::str::from_utf8(local.as_ref()) else {
            return false;
        };
        let mut discard_checkpoint = None;
        let mut discarded_count = 0usize;
        // Only the five raw ancestor slots can participate in the path. No
        // descendants scan, sibling rescan, or retained unselected payload.
        for k in 1..=PATH.len().min(self.frames.len()) {
            let index = self.frames.len() - k;
            let ProcessedElementKind::AlternateBranch {
                observer: Some(observer),
                ..
            } = &mut self.frames[index].kind
            else {
                continue;
            };
            if observer.discarded {
                continue;
            }
            debug_assert_eq!(self.depth + 1 - observer.base_depth, k);
            if k == 1 {
                observer.facts.direct_children = observer.facts.direct_children.saturating_add(1);
            }
            let path_match = observer.level as usize == k - 1
                && observer.open_level as usize == k - 1
                && PATH[k - 1].matches(namespace, local);
            if observer.capability_provisional
                && k == 1
                && (observer.facts.direct_children != 1 || !path_match)
            {
                observer.discarded = true;
                discarded_count += 1;
                discard_checkpoint = Some(
                    discard_checkpoint.map_or(observer.capture_checkpoint, |checkpoint: usize| {
                        checkpoint.min(observer.capture_checkpoint)
                    }),
                );
                continue;
            }
            if !path_match {
                continue;
            }
            observer.level = k as u8;
            if !empty {
                observer.open_level = k as u8;
            }
            if k == 1 {
                observer.facts.drawing = true;
            }
            if k == 4 || k == 5 {
                for attribute in element.attributes() {
                    let Ok(attribute) = attribute else {
                        observer.facts.chartex_uri = false;
                        break;
                    };
                    let (ns, name) = self
                        .reader
                        .reader()
                        .resolver()
                        .resolve_attribute(attribute.key);
                    if (k == 4 && attribute.key.as_ref() == b"uri")
                        || (k == 5
                            && name.as_ref() == b"id"
                            && bounded_xml::resolved_namespace_is(&ns, |ns| {
                                ooxml_common::ns::is_r_ns(Some(ns))
                            }))
                    {
                        let Ok(value) = attribute.normalized_value(XmlVersion::Implicit1_0) else {
                            // Observation adds no XML error path. Leave malformed
                            // attribute bytes to the parent's validation instead
                            // of replacing them with a resource fallback.
                            observer.facts.chartex_uri = false;
                            break;
                        };
                        if k == 4 {
                            observer.facts.chartex_uri = value == CHARTEX_NS;
                        } else {
                            observer.rid = Some(value.into_owned());
                        }
                    }
                }
                if observer.capability_provisional && k == 4 && !observer.facts.chartex_uri {
                    observer.discarded = true;
                    discarded_count += 1;
                    discard_checkpoint = Some(
                        discard_checkpoint
                            .map_or(observer.capture_checkpoint, |checkpoint: usize| {
                                checkpoint.min(observer.capture_checkpoint)
                            }),
                    );
                }
            }
        }
        if let Some(checkpoint) = discard_checkpoint {
            self.discarded_provisional_choices += discarded_count;
            self.capture
                .as_mut()
                .expect("provisional observer requires capture")
                .xml
                .truncate(checkpoint);
            true
        } else {
            false
        }
    }

    fn reject_empty_provisional_choice(&mut self, kind: &ProcessedElementKind) {
        let ProcessedElementKind::AlternateBranch {
            observer:
                Some(ChoiceObserver {
                    capability_provisional: true,
                    capture_checkpoint,
                    ..
                }),
            ..
        } = kind
        else {
            return;
        };
        self.capture
            .as_mut()
            .expect("provisional observer requires capture")
            .xml
            .truncate(*capture_checkpoint);
        if let Some(ElementFrame {
            kind:
                ProcessedElementKind::AlternateContent {
                    selected_branch, ..
                },
            ..
        }) = self.frames.last_mut()
        {
            *selected_branch = false;
        }
    }

    fn observe_choice_end(&mut self) -> Result<(), String> {
        for k in 1..=PATH.len().min(self.frames.len().saturating_sub(1)) {
            let index = self.frames.len() - k - 1;
            if let ProcessedElementKind::AlternateBranch {
                observer: Some(observer),
                ..
            } = &mut self.frames[index].kind
            {
                if observer.open_level as usize == k {
                    observer.open_level -= 1;
                }
            }
        }
        let Some((capability_provisional, discarded, verdict, checkpoint, deferred_error)) = self
            .frames
            .last_mut()
            .and_then(|frame| match &mut frame.kind {
                ProcessedElementKind::AlternateBranch {
                    observer: Some(observer),
                    ..
                } => {
                    observer.facts.renderable_rid = observer
                        .rid
                        .as_ref()
                        .is_some_and(|rid| self.renderable_chartex_rids.contains(rid));
                    let verdict = if observer.discarded {
                        Verdict::Parent
                    } else {
                        chartex_choice::verdict(&observer.facts)
                    };
                    Some((
                        observer.capability_provisional,
                        observer.discarded,
                        verdict,
                        observer.capture_checkpoint,
                        observer.deferred_must_understand,
                    ))
                }
                _ => None,
            })
        else {
            return Ok(());
        };
        if discarded {
            self.discarded_provisional_choices =
                self.discarded_provisional_choices.saturating_sub(1);
        }
        let end_len = self
            .capture
            .as_ref()
            .expect("observer requires capture")
            .xml
            .len();
        let parent = self.frames.len() - 2;
        if let ProcessedElementKind::AlternateContent {
            selected_branch,
            rollback_to,
            choice_end_len,
            ..
        } = &mut self.frames[parent].kind
        {
            if capability_provisional && verdict == Verdict::Parent {
                // The 2015 token is unsupported outside this exact ChartEx
                // shape. Restore the parent's open selection state so a later
                // Choice or the Fallback wins, and retain none of this branch.
                self.capture
                    .as_mut()
                    .expect("observer requires capture")
                    .xml
                    .truncate(checkpoint);
                *selected_branch = false;
            } else if verdict == Verdict::Unrenderable {
                // Keep selected_branch set: later understood Choices remain
                // dropped. Without an authored Fallback, leave the Choice bytes.
                *rollback_to = Some(checkpoint);
                *choice_end_len = end_len;
            }
        }
        // ECMA-376 Part 3 §§9.3–9.4: process MustUnderstand only after
        // Choice selection. A 2015-only Choice is provisional until its exact
        // ChartEx shape is known; a discarded ordinary drawing must not reject
        // the document. Selected ChartEx branches keep the existing mismatch.
        if deferred_error
            && !(capability_provisional && verdict == Verdict::Parent)
            && !Self::defer_must_understand_mismatch(self.frames.iter_mut().rev().skip(1))
        {
            // The diagnostic intentionally omits the namespace text: retaining
            // one authored namespace string per pending Choice could multiply
            // the bounded namespace-context budget by nesting depth.
            return Err(
                "document MCE MustUnderstand namespace is not understood in selected ChartEx Choice"
                    .to_string(),
            );
        }
        Ok(())
    }

    fn validate_processed_must_understand(&mut self, namespaces: &[String]) -> Result<(), String> {
        if namespaces
            .iter()
            .any(|namespace| !docx_understands_namespace(namespace))
            && Self::defer_must_understand_mismatch(self.frames.iter_mut().rev())
        {
            return Ok(());
        }
        bounded_xml::validate_mce_must_understand(
            namespaces,
            &docx_understands_namespace,
            "document",
        )
    }

    fn defer_must_understand_mismatch<'a>(
        frames: impl Iterator<Item = &'a mut ElementFrame>,
    ) -> bool {
        for frame in frames {
            if let ProcessedElementKind::AlternateBranch {
                observer: Some(observer),
                ..
            } = &mut frame.kind
            {
                if observer.capability_provisional {
                    // A discarded Choice is already unselected even while its
                    // descendant frames remain open. Absorb their mismatches
                    // here rather than validating them or deferring outward.
                    if !observer.discarded {
                        // One bit per depth-bounded observer; retain neither
                        // descendant payload nor authored namespace text.
                        observer.deferred_must_understand = true;
                    }
                    return true;
                }
            }
        }
        false
    }

    pub(crate) fn plan(&self) -> Result<DocumentBodyPlan, String> {
        if !self.finished {
            return Err("document body projection plan is not complete".to_string());
        }
        Ok(DocumentBodyPlan {
            cover_break_after: self.cover_break_after.clone(),
        })
    }

    fn classify_processed_element(
        &mut self,
        namespace: Option<&str>,
        element: &BytesStart<'_>,
        context: &Rc<NamespaceContext>,
        visible: bool,
    ) -> Result<(ProcessedElementKind, Rc<StreamedMceScope>), String> {
        let inherited = self
            .frames
            .last()
            .map(|frame| Rc::clone(&frame.mce_scope))
            .unwrap_or_else(StreamedMceScope::root);
        if self.frames.last().is_some_and(|frame| {
            matches!(
                frame.kind,
                ProcessedElementKind::AlternateBranch {
                    filter_direct_children: true,
                    ..
                }
            )
        }) && !(is_w_ns(namespace)
            && matches!(element.local_name().as_ref(), b"drawing" | b"pict"))
        {
            return Ok((ProcessedElementKind::Ignored, inherited));
        }
        if !visible {
            return Ok((ProcessedElementKind::Ignored, inherited));
        }

        if self.frames.last().is_some_and(|frame| {
            matches!(
                frame.kind,
                ProcessedElementKind::Retained { opaque: true, .. }
            )
        }) {
            return Ok((
                ProcessedElementKind::Retained {
                    namespace: namespace.map(str::to_string),
                    local_name: local_name(element)?,
                    opaque: true,
                },
                inherited,
            ));
        }

        let local_name = local_name(element)?;

        // ECMA-376 Part 3 §§8 and 9.1: application-defined extension
        // elements are processed opaquely. MCE directives on the extension
        // payload must not affect either the element or its descendants.
        if docx_is_application_defined_extension_element(namespace, &local_name) {
            return Ok((
                ProcessedElementKind::Retained {
                    namespace: namespace.map(str::to_string),
                    local_name,
                    opaque: true,
                },
                inherited,
            ));
        }

        let attributes = bounded_xml::derive_mce_attributes(
            self.reader.reader(),
            element,
            &inherited,
            context,
            STREAMED_XML_CONTEXT_BYTES,
            "document",
        )
        .map_err(|error| self.map_bounded_error(error, HardResourceLimitKind::XmlContextBytes))?;
        let is_mc = namespace == Some(bounded_xml::MCE_NS);
        if is_mc
            && matches!(
                local_name.as_str(),
                "AlternateContent" | "Choice" | "Fallback"
            )
        {
            bounded_xml::validate_mce_alternate_element_attributes(
                self.reader.reader(),
                element,
                &local_name,
                &attributes.scope,
                "document",
            )?;
        }

        let parent_is_alternate_content = self.frames.last().is_some_and(|frame| {
            matches!(frame.kind, ProcessedElementKind::AlternateContent { .. })
        });
        if parent_is_alternate_content {
            if !is_mc || !matches!(local_name.as_str(), "Choice" | "Fallback") {
                let ignored = namespace.is_some_and(|namespace| {
                    attributes.scope.is_ignorable(namespace)
                        && !docx_understands_namespace(namespace)
                        && !attributes.scope.processes_content(namespace, &local_name)
                });
                if ignored {
                    return Ok((ProcessedElementKind::Ignored, attributes.scope));
                }
                return Err(
                    "document MCE AlternateContent may contain only Choice/Fallback children after Ignorable processing"
                        .to_string(),
                );
            }

            let selection_open = self.frames.last().is_some_and(|frame| {
                matches!(
                    frame.kind,
                    ProcessedElementKind::AlternateContent {
                        selected_branch: false,
                        ..
                    }
                )
            });
            let mut choice_support = if local_name == "Choice" && selection_open {
                Some(self.streamed_choice_is_understood(element, context)?)
            } else {
                None
            };
            if choice_support.is_some_and(|support| support.capability_provisional)
                && attributes
                    .must_understand
                    .iter()
                    .any(|namespace| !docx_understands_namespace(namespace))
            {
                // The parent leaves a 2015-only Choice unselected, so its
                // MustUnderstand must remain inactive until the Choice has both
                // the narrow ChartEx shape and processable MCE requirements.
                choice_support = Some(ChoiceSupport {
                    understood: false,
                    capability_provisional: false,
                });
            }

            let alternate = self
                .frames
                .last_mut()
                .expect("AlternateContent parent checked");
            let ProcessedElementKind::AlternateContent {
                selected_branch,
                seen_choice,
                seen_fallback,
                rollback_to,
                choice_end_len,
            } = &mut alternate.kind
            else {
                unreachable!("AlternateContent parent checked")
            };
            let fallback_override = local_name == "Fallback" && rollback_to.is_some();
            let fallback_understood = attributes
                .must_understand
                .iter()
                .all(|namespace| docx_understands_namespace(namespace));
            let filter_direct_children = fallback_override && fallback_understood;
            let selected = if local_name == "Choice" {
                if *seen_fallback {
                    return Err(
                        "document MCE Choice cannot follow Fallback in AlternateContent"
                            .to_string(),
                    );
                }
                *seen_choice = true;
                !*selected_branch
                    && choice_support
                        .expect("open Choice selection has a classification")
                        .understood
            } else {
                if !*seen_choice {
                    return Err(
                        "document MCE AlternateContent must contain at least one Choice before Fallback"
                            .to_string(),
                    );
                }
                if *seen_fallback {
                    return Err(
                        "document MCE AlternateContent may contain at most one Fallback"
                            .to_string(),
                    );
                }
                *seen_fallback = true;
                if fallback_override && !fallback_understood {
                    // The parent selected Choice remains authoritative when the
                    // resource override's Fallback cannot satisfy §9.4.
                    rollback_to.take();
                    false
                } else if let Some(checkpoint) = rollback_to.take() {
                    let capture = self.capture.as_mut().expect("observer requires capture");
                    debug_assert_eq!(capture.xml.len(), *choice_end_len);
                    capture.xml.truncate(checkpoint);
                    true
                } else {
                    !*selected_branch
                }
            };
            if selected {
                *selected_branch = true;
                self.validate_processed_must_understand(&attributes.must_understand)?;
            }
            return Ok((
                ProcessedElementKind::AlternateBranch {
                    selected,
                    observer: (selected && local_name == "Choice")
                        .then(|| {
                            self.capture.as_ref().map(|capture| ChoiceObserver {
                                base_depth: self.depth + 1,
                                capture_checkpoint: capture.xml.len(),
                                level: 0,
                                open_level: 0,
                                facts: Facts::default(),
                                rid: None,
                                capability_provisional: choice_support
                                    .expect("selected Choice has a classification")
                                    .capability_provisional,
                                discarded: false,
                                deferred_must_understand: false,
                            })
                        })
                        .flatten(),
                    filter_direct_children,
                },
                attributes.scope,
            ));
        }

        if is_mc && local_name == "AlternateContent" {
            self.validate_processed_must_understand(&attributes.must_understand)?;
            return Ok((
                ProcessedElementKind::AlternateContent {
                    selected_branch: false,
                    seen_choice: false,
                    seen_fallback: false,
                    rollback_to: None,
                    choice_end_len: 0,
                },
                attributes.scope,
            ));
        }

        if let Some(namespace) = namespace.filter(|namespace| {
            attributes.scope.is_ignorable(namespace) && !docx_understands_namespace(namespace)
        }) {
            if attributes.scope.processes_content(namespace, &local_name) {
                bounded_xml::validate_mce_process_content_element(
                    self.reader.reader(),
                    element,
                    "document",
                )?;
                self.validate_processed_must_understand(&attributes.must_understand)?;
                return Ok((ProcessedElementKind::Unwrapped, attributes.scope));
            }
            return Ok((ProcessedElementKind::Ignored, attributes.scope));
        }

        self.validate_processed_must_understand(&attributes.must_understand)?;
        Ok((
            ProcessedElementKind::Retained {
                namespace: namespace.map(str::to_string),
                local_name,
                opaque: false,
            },
            attributes.scope,
        ))
    }

    fn append_retained_element(
        &mut self,
        element: &BytesStart<'_>,
        context: &NamespaceContext,
        scope: &StreamedMceScope,
        opaque: bool,
        empty: bool,
    ) -> Result<(), String> {
        let capture = self.capture.as_ref().expect("retained element has capture");
        let block_root = capture.root_depth == self.depth + 1;
        let processed_parent = self
            .frames
            .iter()
            .rev()
            .find_map(|frame| match &frame.kind {
                ProcessedElementKind::Retained { .. } if frame.visible => {
                    Some(frame.context.as_ref())
                }
                _ => None,
            });
        let physical_parent_is_retained = self.frames.last().is_some_and(|frame| {
            frame.visible && matches!(frame.kind, ProcessedElementKind::Retained { .. })
        });
        // Part 3 §9.4 removes selected Choice/ProcessContent ancestors. Repair
        // their namespace bindings together with stripping/quote conversion in
        // one exact preflight against the REMAINING block budget, then write
        // directly to the block. The caller borrows the sole event buffer.
        let repair = if block_root {
            Some((context, None))
        } else if !physical_parent_is_retained {
            Some((context, processed_parent))
        } else {
            None
        };
        let result = bounded_xml::append_projected_start_tag(
            &mut self
                .capture
                .as_mut()
                .expect("retained element has capture")
                .xml,
            self.reader.reader(),
            element,
            bounded_xml::StartTagProjection {
                repair,
                mce: (!opaque).then_some((scope, &docx_understands_namespace)),
                empty,
            },
            HARD_MAX_DOCX_BODY_BLOCK_XML_BYTES as usize,
        );
        if matches!(result, Err(BoundedXmlError::Limit { .. })) && self.discard_overflowing_choice()
        {
            return Ok(());
        }
        result.map_err(|error| {
            self.map_bounded_error(error, HardResourceLimitKind::DocxBodyBlockXmlBytes)
        })
    }

    fn streamed_choice_is_understood(
        &self,
        choice: &BytesStart<'_>,
        context: &NamespaceContext,
    ) -> Result<ChoiceSupport, String> {
        use ooxml_common::mce::ChoiceRequiresClassification;

        let classification = bounded_xml::classify_bounded_mce_choice_requires(
            choice,
            context,
            &docx_understands_namespace,
            STREAMED_XML_CONTEXT_BYTES,
            "document",
        )
        .map_err(|error| self.map_bounded_error(error, HardResourceLimitKind::XmlContextBytes))?;
        match classification {
            ChoiceRequiresClassification::Understood => Ok(ChoiceSupport {
                understood: true,
                capability_provisional: false,
            }),
            ChoiceRequiresClassification::Unsupported => {
                let with_chart_ex_capability = bounded_xml::classify_bounded_mce_choice_requires(
                    choice,
                    context,
                    &|namespace| {
                        docx_understands_namespace(namespace)
                            || namespace == crate::chartex_choice::CHARTEX_CAPABILITY_NS
                    },
                    STREAMED_XML_CONTEXT_BYTES,
                    "document",
                )
                .map_err(|error| {
                    self.map_bounded_error(error, HardResourceLimitKind::XmlContextBytes)
                })?;
                Ok(ChoiceSupport {
                    understood: with_chart_ex_capability
                        == ChoiceRequiresClassification::Understood
                        && self.capture.is_some(),
                    capability_provisional: with_chart_ex_capability
                        == ChoiceRequiresClassification::Understood,
                })
            }
            ChoiceRequiresClassification::Missing => {
                Err("document MCE Choice must have a Requires attribute".to_string())
            }
            ChoiceRequiresClassification::Blank => {
                Err("document MCE Choice must have a non-empty Requires attribute".to_string())
            }
            ChoiceRequiresClassification::Unresolved => {
                Err("document MCE Choice Requires uses an unbound namespace prefix".to_string())
            }
        }
    }

    fn handle_start(
        &mut self,
        namespace: Option<&str>,
        start: &BytesStart<'_>,
        next_depth: usize,
        source_offset: usize,
    ) -> Result<bool, String> {
        let local_name = local_name(start)?;
        let is_word = is_w_ns(namespace);
        if is_word && local_name == "body" && self.body_content_depth.is_none() {
            self.body_seen = true;
            self.body_content_depth = Some(next_depth);
            return Ok(false);
        }

        self.observe_sdt_metadata(is_word, &local_name, start, next_depth)?;
        if self.at_logical_level() && is_word && local_name == "sdt" {
            self.sdts.push(SdtFrame {
                open_depth: next_depth,
                ..SdtFrame::default()
            });
            return Ok(false);
        }
        if self.at_logical_level() && is_word && is_body_block(&local_name) {
            self.capture = Some(Capture {
                root_depth: next_depth,
                source_offset,
                local_name,
                xml: Vec::new(),
            });
            return Ok(true);
        }
        Ok(false)
    }

    fn handle_empty(
        &mut self,
        namespace: Option<&str>,
        empty: &BytesStart<'_>,
        source_offset: usize,
    ) -> Result<bool, String> {
        let local_name = local_name(empty)?;
        let is_word = is_w_ns(namespace);
        self.observe_sdt_metadata(is_word, &local_name, empty, self.depth)?;
        if !(self.at_logical_level() && is_word && is_body_block(&local_name)) {
            return Ok(false);
        }
        self.capture = Some(Capture {
            // Empty elements do not push a physical reader frame, but their
            // logical element depth is still one below the current parent.
            // Keep that identity consistent with Start/End captures so an
            // empty block inside sdtContent can become the cover wrapper's
            // last projected block.
            root_depth: self.depth.saturating_add(1),
            source_offset,
            local_name,
            xml: Vec::new(),
        });
        Ok(true)
    }

    fn handle_end(&mut self, namespace: Option<&str>, raw_local_name: &[u8]) -> Result<(), String> {
        let local_name = std::str::from_utf8(raw_local_name)
            .map_err(|error| format!("document body element name: {error}"))?;
        if is_w_ns(namespace) && local_name == "body" && self.body_content_depth == Some(self.depth)
        {
            self.body_content_depth = None;
            self.body_closed = true;
        }
        if self.frames.last().is_some_and(|frame| {
            matches!(
                frame.kind,
                ProcessedElementKind::AlternateContent {
                    seen_choice: false,
                    ..
                }
            )
        }) {
            return Err(
                "document MCE AlternateContent must contain at least one Choice".to_string(),
            );
        }
        if let Some(sdt) = self.sdts.last_mut() {
            if sdt.properties_depth == Some(self.depth) {
                sdt.properties_depth = None;
            }
            if sdt.content_depth == Some(self.depth) {
                sdt.content_depth = None;
            }
            if is_w_ns(namespace) && local_name == "sdt" && sdt.open_depth == self.depth {
                let completed = self.sdts.pop().expect("matching SDT frame exists");
                if completed.cover_pages {
                    if let Some(ordinal) = completed.last_block_ordinal {
                        if let Some(flag) = self.cover_break_after.get_mut(ordinal) {
                            *flag = true;
                        }
                    }
                }
            }
        }
        Ok(())
    }

    fn observe_sdt_metadata(
        &mut self,
        is_word: bool,
        local_name: &str,
        element: &BytesStart<'_>,
        next_depth: usize,
    ) -> Result<(), String> {
        let parent_is_sdt = self.processed_parent_is_word("sdt");
        let Some(sdt) = self.sdts.last_mut() else {
            return Ok(());
        };
        if is_word && parent_is_sdt {
            if local_name == "sdtPr" {
                sdt.properties_depth = Some(next_depth);
            } else if local_name == "sdtContent" {
                sdt.content_depth = Some(next_depth);
            }
        }
        if is_word
            && local_name == "docPartGallery"
            && sdt
                .properties_depth
                .is_some_and(|depth| self.depth >= depth)
            && attribute_value(self.reader.reader(), element, "val", &|namespace| {
                is_w_ns(Some(namespace))
            })?
            .as_deref()
                == Some("Cover Pages")
        {
            sdt.cover_pages = true;
        }
        Ok(())
    }

    fn at_logical_level(&self) -> bool {
        self.processed_parent_is_word("body") || self.processed_parent_is_word("sdtContent")
    }

    fn processed_parent_is_word(&self, expected_local_name: &str) -> bool {
        self.frames.iter().rev().find_map(|frame| {
            if !frame.visible {
                return None;
            }
            match &frame.kind {
                ProcessedElementKind::Retained {
                    namespace,
                    local_name,
                    ..
                } => Some(
                    is_w_ns(namespace.as_deref()) && local_name.as_str() == expected_local_name,
                ),
                _ => None,
            }
        }) == Some(true)
    }

    fn discard_overflowing_choice(&mut self) -> bool {
        let checkpoint = self.frames.iter_mut().rev().find_map(|frame| {
            let ProcessedElementKind::AlternateBranch {
                observer: Some(observer),
                ..
            } = &mut frame.kind
            else {
                return None;
            };
            if !observer.capability_provisional || observer.discarded {
                return None;
            }
            // The parent never retains a 2015-only Choice. If a provisional
            // candidate cannot fit the parent's block budget before its
            // ChartEx shape is known, fail closed to the parent-selected
            // later Choice/Fallback instead of introducing a new limit.
            observer.discarded = true;
            Some(observer.capture_checkpoint)
        });
        if let Some(checkpoint) = checkpoint {
            self.discarded_provisional_choices += 1;
            self.capture
                .as_mut()
                .expect("provisional observer requires capture")
                .xml
                .truncate(checkpoint);
            true
        } else {
            false
        }
    }

    fn append_capture(&mut self, event: Event<'_>) -> Result<(), String> {
        let projected = bounded_xml::projected_event_bytes(&event);
        let would_overflow = self.capture.as_ref().is_some_and(|capture| {
            capture.xml.len().saturating_add(projected)
                > HARD_MAX_DOCX_BODY_BLOCK_XML_BYTES as usize
        });
        if would_overflow && self.discard_overflowing_choice() {
            return Ok(());
        }
        let capture = self
            .capture
            .as_mut()
            .ok_or_else(|| "document body capture is not active".to_string())?;
        bounded_xml::append_projected_event(
            &mut capture.xml,
            &event,
            0,
            HARD_MAX_DOCX_BODY_BLOCK_XML_BYTES as usize,
            "document body block",
        )
        .map_err(|error| {
            self.map_bounded_error(error, HardResourceLimitKind::DocxBodyBlockXmlBytes)
        })
    }

    /// The captured block, or `None` for a `<w:tbl>` with no row. CT_Tbl
    /// allows none (EG_ContentRowContent is minOccurs="0"); such a table has
    /// nothing to lay out, so it is no body block: it takes no ordinal, cannot
    /// end a Cover Pages building block, and is not seen by the §17.4.37
    /// sequence or REF passes, exactly as if it were absent.
    fn finish_capture(&mut self) -> Result<Option<ProjectedBodyBlock>, String> {
        let capture = self
            .capture
            .take()
            .ok_or_else(|| "document body capture is not active".to_string())?;
        if capture.local_name == "tbl" && !captured_table_has_rows(&capture.xml)? {
            return Ok(None);
        }
        let ordinal = self.cover_break_after.len();
        self.cover_break_after.push(false);
        for sdt in &mut self.sdts {
            if sdt
                .content_depth
                .is_some_and(|depth| capture.root_depth > depth)
            {
                sdt.last_block_ordinal = Some(ordinal);
            }
        }
        Ok(Some(ProjectedBodyBlock {
            ordinal,
            source_offset: capture.source_offset,
            local_name: capture.local_name,
            xml: capture.xml,
        }))
    }

    fn map_read_error(&self, error: BoundedXmlReadError) -> String {
        match error {
            BoundedXmlReadError::Xml(message) => message,
            BoundedXmlReadError::NamespaceDeclarationLimit { limit, observed } => self
                .report_limit(
                    HardResourceLimitKind::XmlNamespaceDeclarations,
                    limit,
                    observed,
                ),
            BoundedXmlReadError::EventLimit { limit, observed } => {
                self.report_limit(HardResourceLimitKind::XmlEventBytes, limit, observed)
            }
        }
    }

    fn map_bounded_error(&self, error: BoundedXmlError, kind: HardResourceLimitKind) -> String {
        match error {
            BoundedXmlError::Xml(message) => message,
            BoundedXmlError::Limit { limit, observed } => self.report_limit(kind, limit, observed),
        }
    }

    fn report_limit(&self, kind: HardResourceLimitKind, limit: usize, observed: usize) -> String {
        match bounded_xml::report_hard_limit(
            self.reporter.as_ref(),
            kind,
            Some(DOCUMENT_PART),
            limit,
            observed,
        ) {
            Err(message) => message,
            Ok(()) => format!("document body projector limit exceeded: {observed} > {limit}"),
        }
    }
}

pub(crate) fn docx_understands_namespace(namespace: &str) -> bool {
    use ooxml_common::ns::{is_a_ns, is_c_ns, is_m_ns, is_pic_ns, is_r_ns, is_wp_ns};

    is_w_ns(Some(namespace))
        || is_a_ns(Some(namespace))
        || is_c_ns(Some(namespace))
        || is_m_ns(Some(namespace))
        || is_pic_ns(Some(namespace))
        || is_r_ns(Some(namespace))
        || is_wp_ns(Some(namespace))
        || crate::parser::docx_understands_drawing_ns(namespace)
}

pub(crate) fn docx_is_application_defined_extension_element(
    namespace: Option<&str>,
    local_name: &str,
) -> bool {
    use ooxml_common::ns::{is_a_ns, is_c_ns, is_w_ns};

    local_name == "extLst" && (is_w_ns(namespace) || is_a_ns(namespace) || is_c_ns(namespace))
}

fn local_name(element: &BytesStart<'_>) -> Result<String, String> {
    std::str::from_utf8(element.local_name().as_ref())
        .map(str::to_string)
        .map_err(|error| format!("document body element name: {error}"))
}

fn is_body_block(local_name: &str) -> bool {
    matches!(local_name, "p" | "tbl" | "sectPr")
}

/// Whether a captured `<w:tbl>` has a `<w:tr>` child, looking through
/// `<w:sdt>`/`<w:sdtContent>` row wrappers exactly as the parser's
/// `children_w_flat` does. Rows follow tblPr/tblGrid, so the scan stops at the
/// first row; only a row-less table is read to its end.
fn captured_table_has_rows(xml: &[u8]) -> Result<bool, String> {
    use quick_xml::name::ResolveResult;
    #[derive(Clone, Copy)]
    enum Frame {
        /// The table, or an sdtContent inside a row wrapper: a row here counts.
        Rows,
        /// A w:sdt row wrapper: only its sdtContent continues the row level.
        Sdt,
        Other,
    }
    let word = |namespace: &ResolveResult<'_>| {
        matches!(namespace, ResolveResult::Bound(uri)
            if std::str::from_utf8(uri.as_ref()).is_ok_and(|uri| is_w_ns(Some(uri))))
    };
    let mut reader = quick_xml::NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut frames: Vec<Frame> = Vec::new();
    loop {
        buffer.clear();
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| format!("document body table block: {error}"))?;
        let (element, empty) = match &event {
            Event::Start(element) => (element, false),
            Event::Empty(element) => (element, true),
            Event::End(_) => {
                frames.pop();
                continue;
            }
            Event::Eof => return Ok(false),
            _ => continue,
        };
        let name = element.local_name();
        let name = name.as_ref();
        let frame = match frames.last() {
            None => Frame::Rows,
            Some(Frame::Rows) if word(&namespace) && name == b"tr" => return Ok(true),
            Some(Frame::Rows) if word(&namespace) && name == b"sdt" => Frame::Sdt,
            Some(Frame::Sdt) if word(&namespace) && name == b"sdtContent" => Frame::Rows,
            Some(_) => Frame::Other,
        };
        if !empty {
            frames.push(frame);
        }
    }
}

fn attribute_value<R>(
    reader: &quick_xml::NsReader<R>,
    element: &BytesStart<'_>,
    local_name: &str,
    accepted_namespace: &dyn Fn(&str) -> bool,
) -> Result<Option<String>, String> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| format!("document body attribute: {error}"))?;
        let key = attribute.key.as_ref();
        let resolved_local = key.rsplit(|byte| *byte == b':').next().unwrap_or(key);
        if resolved_local != local_name.as_bytes() {
            continue;
        }
        let accepted = if key.contains(&b':') {
            let (namespace, _) = reader.resolver().resolve_attribute(attribute.key);
            bounded_xml::resolved_namespace_is(&namespace, accepted_namespace)
        } else {
            true
        };
        if !accepted {
            continue;
        }
        let value = attribute
            .normalized_value(XmlVersion::Implicit1_0)
            .map_err(|error| format!("document body attribute value: {error}"))?;
        return Ok(Some(value.into_owned()));
    }
    Ok(None)
}

#[cfg(test)]
mod tests {
    use std::io::{BufReader, Cursor};

    use super::*;

    const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    const WPS: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

    fn project(xml: &str) -> (Vec<ProjectedBodyBlock>, DocumentBodyPlan) {
        let mut projector = DocumentBodyProjector::new(
            BufReader::new(Cursor::new(xml.as_bytes().to_vec())),
            None,
            Rc::default(),
        );
        let mut blocks = Vec::new();
        while let Some(block) = projector.next_block().unwrap() {
            blocks.push(block);
        }
        let plan = projector.plan().unwrap();
        (blocks, plan)
    }

    #[test]
    fn projects_body_blocks_without_retaining_the_document_wrapper() {
        let xml = format!(
            r#"<w:document xmlns:w="{W}" xmlns:r="urn:rels"><w:body>
              <w:p><w:r><w:t>A</w:t></w:r></w:p>
              <w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>
              <w:sectPr/>
            </w:body></w:document>"#,
        );
        let (blocks, plan) = project(&xml);
        assert_eq!(
            blocks
                .iter()
                .map(|block| block.local_name.as_str())
                .collect::<Vec<_>>(),
            ["p", "tbl", "sectPr"]
        );
        assert_eq!(plan.cover_break_after, [false, false, false]);
        assert_eq!(blocks[0].source_offset, xml.find("<w:p>").unwrap());
        assert_eq!(blocks[1].source_offset, xml.find("<w:tbl>").unwrap());
        for block in blocks {
            let text = String::from_utf8(block.xml).unwrap();
            assert!(
                text.contains("xmlns:w="),
                "inherited namespace must be repaired: {text}"
            );
            roxmltree::Document::parse(&text).unwrap();
        }
    }

    #[test]
    fn flattens_nested_block_sdts_and_marks_only_the_cover_tail() {
        let xml = format!(
            r#"<w:document xmlns:w="{W}"><w:body>
              <w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Cover Pages"/></w:docPartObj></w:sdtPr>
                <w:sdtContent><w:p/><w:sdt><w:sdtContent><w:p/><w:tbl><w:tr/></w:tbl></w:sdtContent></w:sdt></w:sdtContent>
              </w:sdt>
              <w:p/>
            </w:body></w:document>"#,
        );
        let (blocks, plan) = project(&xml);
        assert_eq!(
            blocks
                .iter()
                .map(|block| block.local_name.as_str())
                .collect::<Vec<_>>(),
            ["p", "p", "tbl", "p"]
        );
        assert_eq!(plan.cover_break_after, [false, false, true, false]);
    }

    #[test]
    fn row_less_tables_are_not_blocks_and_never_end_a_cover() {
        // CT_Tbl allows no w:tr; a table with none (directly, or only an empty
        // sdt row wrapper) is no block. A row reached through sdt wrappers is
        // a row, and a w:tr inside tblPr-level or other content is not.
        let xml = format!(
            r#"<w:document xmlns:w="{W}"><w:body>
              <w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Cover Pages"/></w:docPartObj></w:sdtPr>
                <w:sdtContent><w:p/><w:tbl/><w:tbl><w:tblPr/><w:sdt><w:sdtContent/></w:sdt></w:tbl></w:sdtContent>
              </w:sdt>
              <w:tbl><w:tblPr/><w:sdt><w:sdtContent><w:sdt><w:sdtContent><w:tr/></w:sdtContent></w:sdt></w:sdtContent></w:sdt></w:tbl>
              <w:tbl><w:tblPr><w:tr/></w:tblPr><w:customXml><w:tr/></w:customXml></w:tbl>
              <w:p/>
            </w:body></w:document>"#,
        );
        let (blocks, plan) = project(&xml);
        assert_eq!(
            blocks
                .iter()
                .map(|block| block.local_name.as_str())
                .collect::<Vec<_>>(),
            ["p", "tbl", "p"]
        );
        assert_eq!(
            blocks.iter().map(|block| block.ordinal).collect::<Vec<_>>(),
            [0, 1, 2]
        );
        assert_eq!(plan.cover_break_after, [true, false, false]);
    }

    #[test]
    fn ignores_unsupported_top_level_elements_but_validates_the_tail() {
        let xml = format!(
            r#"<w:document xmlns:w="{W}"><w:body><w:altChunk/><w:p/></w:body></w:document>"#,
        );
        let (blocks, _) = project(&xml);
        assert_eq!(blocks.len(), 1);
        assert_eq!(blocks[0].local_name, "p");
    }

    #[test]
    fn applies_alternate_content_selection_before_body_block_classification() {
        let xml = format!(
            r#"<w:document xmlns:w="{W}" xmlns:mc="{MC}" xmlns:wps="{WPS}" xmlns:u="urn:unsupported"><w:body>
              <mc:AlternateContent>
                <mc:Choice Requires="u"><w:p><w:r><w:t>hidden choice</w:t></w:r></w:p></mc:Choice>
                <mc:Choice Requires="wps"><w:p><w:r><w:t>selected choice</w:t></w:r></w:p></mc:Choice>
                <mc:Fallback><w:p><w:r><w:t>hidden fallback</w:t></w:r></w:p></mc:Fallback>
              </mc:AlternateContent>
              <mc:AlternateContent>
                <mc:Choice Requires="u"><w:tbl><w:tr/></w:tbl></mc:Choice>
                <mc:Fallback><w:sectPr/></mc:Fallback>
              </mc:AlternateContent>
            </w:body></w:document>"#,
        );
        let (blocks, _) = project(&xml);
        assert_eq!(
            blocks
                .iter()
                .map(|block| block.local_name.as_str())
                .collect::<Vec<_>>(),
            ["p", "sectPr"]
        );
        let paragraph = String::from_utf8(blocks[0].xml.clone()).unwrap();
        assert!(paragraph.contains("selected choice"), "{paragraph}");
        assert!(!paragraph.contains("hidden choice"), "{paragraph}");
        assert!(!paragraph.contains("hidden fallback"), "{paragraph}");
    }

    #[test]
    fn exposes_process_content_children_and_repairs_moved_namespaces() {
        let xml = format!(
            r#"<w:document xmlns:w="{W}" xmlns:mc="{MC}" xmlns:f="urn:future" mc:Ignorable="f" mc:ProcessContent="f:wrapper"><w:body>
              <f:drop><w:p><w:r><w:t>ignored</w:t></w:r></w:p></f:drop>
              <f:wrapper xmlns:r="urn:relationships"><w:p r:id="r1"><w:r><w:t>kept</w:t></w:r></w:p></f:wrapper>
            </w:body></w:document>"#,
        );
        let (blocks, _) = project(&xml);
        assert_eq!(blocks.len(), 1);
        let paragraph = String::from_utf8(blocks[0].xml.clone()).unwrap();
        assert!(paragraph.contains("kept"), "{paragraph}");
        assert!(!paragraph.contains("ignored"), "{paragraph}");
        assert!(
            paragraph.contains("xmlns:r=\"urn:relationships\""),
            "{paragraph}"
        );
        roxmltree::Document::parse(&paragraph).unwrap();
    }

    #[test]
    fn preprocesses_mce_inside_captured_paragraph_and_table_blocks() {
        let xml = format!(
            r#"<w:document xmlns:w="{W}" xmlns:mc="{MC}" xmlns:f="urn:future" mc:Ignorable="f" mc:ProcessContent="f:rWrap f:pWrap"><w:body>
              <w:p>
                <mc:AlternateContent>
                  <mc:Choice Requires="f"><w:r><w:t>hidden choice</w:t></w:r></mc:Choice>
                  <mc:Fallback><f:rWrap><w:r><w:t>kept fallback</w:t></w:r></f:rWrap></mc:Fallback>
                </mc:AlternateContent>
              </w:p>
              <w:tbl><w:tr><w:tc><f:pWrap><w:p><w:r><w:t>kept cell</w:t></w:r></w:p></f:pWrap></w:tc></w:tr></w:tbl>
            </w:body></w:document>"#,
        );
        let (blocks, _) = project(&xml);
        assert_eq!(blocks.len(), 2);
        let paragraph = String::from_utf8(blocks[0].xml.clone()).unwrap();
        assert!(paragraph.contains("kept fallback"), "{paragraph}");
        assert!(!paragraph.contains("hidden choice"), "{paragraph}");
        assert!(!paragraph.contains("AlternateContent"), "{paragraph}");
        assert!(!paragraph.contains("rWrap"), "{paragraph}");
        let table = String::from_utf8(blocks[1].xml.clone()).unwrap();
        assert!(table.contains("kept cell"), "{table}");
        assert!(!table.contains("pWrap"), "{table}");
    }

    #[test]
    fn captured_must_understand_and_ignored_attributes_follow_part3() {
        let rejected = format!(
            r#"<w:document xmlns:w="{W}" xmlns:mc="{MC}" xmlns:u="urn:unsupported"><w:body><w:p><w:r mc:MustUnderstand="u"><w:t>x</w:t></w:r></w:p></w:body></w:document>"#,
        );
        let mut projector = DocumentBodyProjector::new(
            BufReader::new(Cursor::new(rejected.into_bytes())),
            None,
            Rc::default(),
        );
        assert!(projector
            .next_block()
            .unwrap_err()
            .contains("MustUnderstand namespace is not understood"));

        let stripped = format!(
            r#"<w:document xmlns:w="{W}" xmlns:mc="{MC}" xmlns:f="urn:future" mc:Ignorable="f"><w:body><w:p f:val="ignored" mc:Ignorable="f"><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>"#,
        );
        let (blocks, _) = project(&stripped);
        let paragraph = String::from_utf8(blocks[0].xml.clone()).unwrap();
        assert!(!paragraph.contains("f:val"), "{paragraph}");
        assert!(!paragraph.contains("mc:Ignorable"), "{paragraph}");
    }

    #[test]
    fn application_defined_extension_payload_is_opaque_to_mce_processing() {
        let xml = format!(
            r#"<w:document xmlns:w="{W}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:mc="{MC}" xmlns:f="urn:future"><w:body><w:p><w:r><w:drawing><a:extLst mc:Ignorable="f" mc:MustUnderstand="f"><a:ext><f:payload f:val="kept"/></a:ext></a:extLst></w:drawing></w:r></w:p></w:body></w:document>"#,
        );
        let (blocks, _) = project(&xml);
        let paragraph = String::from_utf8(blocks[0].xml.clone()).unwrap();
        assert!(paragraph.contains("f:payload"), "{paragraph}");
        assert!(paragraph.contains("f:val=\"kept\""), "{paragraph}");
        assert!(paragraph.contains("mc:MustUnderstand=\"f\""), "{paragraph}");
    }

    #[test]
    fn cover_gallery_value_must_be_in_the_wordprocessing_namespace() {
        let xml = format!(
            r#"<w:document xmlns:w="{W}" xmlns:f="urn:foreign"><w:body>
              <w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery f:val="Cover Pages"/></w:docPartObj></w:sdtPr>
                <w:sdtContent><w:p/></w:sdtContent>
              </w:sdt>
            </w:body></w:document>"#,
        );
        let (_, plan) = project(&xml);
        assert_eq!(plan.cover_break_after, [false]);
    }

    #[test]
    fn streamed_alternate_content_matches_shared_materialized_selector() {
        let xml = format!(
            r#"<w:document xmlns:w="{W}" xmlns:mc="{MC}" xmlns:wps="{WPS}" xmlns:u="urn:unsupported"><w:body>
              <mc:AlternateContent>
                <mc:Choice Requires="u"><w:p><w:r><w:t>choice</w:t></w:r></w:p></mc:Choice>
                <mc:Fallback><w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></mc:Fallback>
              </mc:AlternateContent>
            </w:body></w:document>"#,
        );
        let materialized = roxmltree::Document::parse(&xml).unwrap();
        let alternate = materialized
            .descendants()
            .find(|node| node.is_element() && node.tag_name().name() == "AlternateContent")
            .unwrap();
        let selected =
            ooxml_common::mce::select_alternate_content(alternate, &docx_understands_namespace)
                .unwrap();
        let materialized_names = selected
            .children()
            .filter(|node| node.is_element() && is_w_ns(node.tag_name().namespace()))
            .map(|node| node.tag_name().name())
            .collect::<Vec<_>>();

        let (streamed, _) = project(&xml);
        assert_eq!(
            streamed
                .iter()
                .map(|block| block.local_name.as_str())
                .collect::<Vec<_>>(),
            materialized_names
        );
    }

    #[test]
    fn selected_branch_keeps_sdt_cover_semantics_on_processed_parent_chain() {
        let xml = format!(
            r#"<w:document xmlns:w="{W}" xmlns:mc="{MC}" xmlns:u="urn:unsupported"><w:body>
              <mc:AlternateContent>
                <mc:Choice Requires="u"><w:p/></mc:Choice>
                <mc:Fallback><w:sdt>
                  <w:sdtPr><w:docPartObj><w:docPartGallery w:val="Cover Pages"/></w:docPartObj></w:sdtPr>
                  <w:sdtContent><w:p/><w:tbl><w:tr/></w:tbl></w:sdtContent>
                </w:sdt></mc:Fallback>
              </mc:AlternateContent>
              <w:p/>
            </w:body></w:document>"#,
        );
        let (blocks, plan) = project(&xml);
        assert_eq!(
            blocks
                .iter()
                .map(|block| block.local_name.as_str())
                .collect::<Vec<_>>(),
            ["p", "tbl", "p"]
        );
        assert_eq!(plan.cover_break_after, [false, true, false]);
    }

    #[test]
    fn must_understand_is_checked_only_on_the_selected_processed_path() {
        let xml = format!(
            r#"<w:document xmlns:w="{W}" xmlns:mc="{MC}" xmlns:u="urn:unsupported"><w:body>
              <mc:AlternateContent>
                <mc:Choice Requires="u" mc:MustUnderstand="u"><w:tbl><w:tr/></w:tbl></mc:Choice>
                <mc:Fallback><w:p/></mc:Fallback>
              </mc:AlternateContent>
            </w:body></w:document>"#,
        );
        let (blocks, _) = project(&xml);
        assert_eq!(blocks[0].local_name, "p");

        let mismatch = format!(
            r#"<w:document xmlns:w="{W}" xmlns:mc="{MC}" xmlns:u="urn:unsupported" mc:MustUnderstand="u"><w:body><w:p/></w:body></w:document>"#,
        );
        let mut projector = DocumentBodyProjector::new(
            BufReader::new(Cursor::new(mismatch.into_bytes())),
            None,
            Rc::default(),
        );
        assert!(projector
            .next_block()
            .unwrap_err()
            .contains("MustUnderstand namespace is not understood"));
    }

    #[test]
    fn missing_body_is_fatal_for_the_streamed_required_part() {
        let xml = format!(r#"<w:document xmlns:w="{W}"/>"#);
        let mut projector = DocumentBodyProjector::new(
            BufReader::new(Cursor::new(xml.into_bytes())),
            None,
            Rc::default(),
        );
        assert!(projector.next_block().unwrap_err().contains("no <w:body>"));
    }

    #[test]
    fn repaired_large_events_stay_within_two_block_buffers() {
        use ooxml_common::resource::{OoxmlFormat, ResourceGovernor};
        use std::io::Read;

        // Generate the payload through Read, excluding whole-input storage.
        // Count allocator-requested live bytes, including the event buffer,
        // projected block, namespace state, and all rewrite intermediates.
        let mut peaks = Vec::new();
        for byte in *b"A\"" {
            let peak = crate::allocation_tests::peak(|| {
                let governor = ResourceGovernor::from_wasm(OoxmlFormat::Docx, None, None, None);
                let scope = governor.scope("parse");
                let prefix = format!(
                    r#"<w:document xmlns:w="{W}" xmlns:mc="{MC}" xmlns:wps="{WPS}" xmlns:v="urn:schemas-microsoft-com:vml"><w:body><w:p><mc:AlternateContent><mc:Choice Requires="wps" xmlns:o="urn:schemas-microsoft-com:office:office"><v:shape o:gfxdata='"#,
                );
                let suffix = b"'/></mc:Choice></mc:AlternateContent></w:p></w:body></w:document>";
                let input = Cursor::new(prefix.into_bytes())
                    .chain(std::io::repeat(byte).take(31 * 1024 * 1024))
                    .chain(Cursor::new(suffix));
                let mut projector =
                    DocumentBodyProjector::new(BufReader::new(input), None, Rc::default());
                let result = projector.next_block();
                if byte == b'A' {
                    let block = result.unwrap().unwrap();
                    assert!(block.xml.len() > 31 * 1024 * 1024);
                } else {
                    let error = result.unwrap_err();
                    let payload: serde_json::Value = serde_json::from_str(
                        error
                            .strip_prefix("OOXML_RESOURCE_LIMIT:")
                            .expect("typed block limit"),
                    )
                    .unwrap();
                    assert_eq!(
                        payload["details"]["violation"]["resource"],
                        "docx-body-block-xml"
                    );
                    assert_eq!(
                        scope.resource_limit_error().as_deref(),
                        Some(error.as_str())
                    );
                }
            });
            eprintln!(
                "31 MiB byte {byte}: peak {peak} bytes ({:.3} MiB)",
                peak as f64 / 1048576.0
            );
            peaks.push(peak);
        }
        for peak in peaks {
            assert!(
                peak <= 2 * HARD_MAX_DOCX_BODY_BLOCK_XML_BYTES as usize,
                "peak live allocation {peak} exceeds two block buffers"
            );
        }
    }

    #[test]
    fn namespace_declaration_count_exhaustion_is_typed_and_latched() {
        use ooxml_common::resource::{OoxmlFormat, ResourceGovernor};
        let uri = "x".repeat(8 * 1024);
        let declarations = (0..257)
            .map(|n| format!(r#" xmlns:n{n}="{uri}""#))
            .collect::<String>();
        let xml = format!(
            r#"<w:document xmlns:w="{W}"><w:body><w:p{declarations}/></w:body></w:document>"#
        );
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Docx, None, None, None);
        let scope = governor.scope("parse");
        let mut projector = DocumentBodyProjector::new(
            BufReader::new(Cursor::new(xml.into_bytes())),
            None,
            Rc::default(),
        );
        let error = projector.next_block().unwrap_err();
        let payload: serde_json::Value = serde_json::from_str(
            error
                .strip_prefix("OOXML_RESOURCE_LIMIT:")
                .expect("typed namespace limit"),
        )
        .unwrap();
        let violation = &payload["details"]["violation"];
        assert_eq!(violation["resource"], "xml-context");
        assert_eq!(violation["metric"], "declarations");
        assert_eq!(violation["limit"], 256);
        assert_eq!(violation["observed"], 257);
        assert_eq!(
            scope.resource_limit_error().as_deref(),
            Some(error.as_str())
        );
        assert_eq!(projector.next_block().unwrap_err(), error);
    }

    #[test]
    fn admits_large_vml_shape_start_tags_within_the_block_ceiling() {
        // ECMA-376 Part 4 §19.1.2.19 CT_Shape declares o:gfxdata, a
        // base64Binary DrawingML package. The same tag must be admitted after
        // Part 3 §9.4 removes selected MCE/ProcessContent wrappers.
        let gfxdata = "A".repeat(STREAMED_XML_CONTEXT_BYTES + 4);
        let shape = format!(r#"<v:shape o:gfxdata="{gfxdata}"/>"#);
        for content in [
            format!(
                r#"<v:shape xmlns:o="urn:schemas-microsoft-com:office:office" o:gfxdata="{gfxdata}"/>"#
            ),
            // Repair a large Empty event immediately below selected Choice.
            format!(
                r#"<mc:AlternateContent><mc:Choice Requires="wps" xmlns:o="urn:schemas-microsoft-com:office:office">{shape}</mc:Choice></mc:AlternateContent>"#
            ),
            // Repair a Start event through nested selection, including a
            // selected Fallback, inside w:pict within w:object.
            format!(
                r#"<mc:AlternateContent><mc:Choice Requires="wps"><mc:AlternateContent><mc:Choice Requires="f"/><mc:Fallback xmlns:o="urn:schemas-microsoft-com:office:office"><v:shape o:gfxdata="{gfxdata}"><v:path/></v:shape></mc:Fallback></mc:AlternateContent></mc:Choice></mc:AlternateContent>"#
            ),
            // Ignorable attribute stripping must preserve the large retained
            // attribute when ProcessContent removes the physical parent.
            format!(
                r#"<f:wrapper xmlns:o="urn:schemas-microsoft-com:office:office"><v:shape f:drop="ignored" o:gfxdata="{gfxdata}"/></f:wrapper>"#
            ),
        ] {
            let xml = format!(
                r#"<w:document xmlns:w="{W}" xmlns:mc="{MC}" xmlns:wps="{WPS}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:f="urn:future" mc:Ignorable="f" mc:ProcessContent="f:wrapper"><w:body><w:p><w:r><w:object><w:pict>{content}</w:pict></w:object></w:r></w:p></w:body></w:document>"#,
            );
            let (blocks, _) = project(&xml);
            assert_eq!(blocks.len(), 1);
            let projected = std::str::from_utf8(&blocks[0].xml).unwrap();
            let doc = roxmltree::Document::parse(projected).unwrap();
            let shape = doc
                .descendants()
                .find(|node| node.has_tag_name(("urn:schemas-microsoft-com:vml", "shape")))
                .unwrap();
            assert_eq!(
                shape.attribute(("urn:schemas-microsoft-com:office:office", "gfxdata")),
                Some(gfxdata.as_str())
            );
            assert!(!projected.contains("AlternateContent"));
            assert!(!projected.contains("f:wrapper"));
            assert!(!projected.contains("f:drop"));
        }
    }

    #[test]
    fn admits_large_unselected_fallback_tag_without_retaining_it() {
        let gfxdata = "A".repeat(STREAMED_XML_CONTEXT_BYTES + 4);
        let xml = format!(
            r#"<w:document xmlns:w="{W}" xmlns:mc="{MC}" xmlns:wps="{WPS}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><w:body><w:p><w:r><w:pict><mc:AlternateContent><mc:Choice Requires="wps"><v:shape id="selected"/></mc:Choice><mc:Fallback><v:shape o:gfxdata="{gfxdata}"/></mc:Fallback></mc:AlternateContent></w:pict></w:r></w:p></w:body></w:document>"#,
        );
        let (blocks, _) = project(&xml);
        assert_eq!(blocks.len(), 1);
        let projected = std::str::from_utf8(&blocks[0].xml).unwrap();
        assert!(projected.contains(r#"id="selected""#));
        assert!(!projected.contains("gfxdata"));
        assert!(projected.len() < 1024);
    }

    #[test]
    fn block_root_repair_uses_the_block_budget_for_start_and_empty_events() {
        // Synthetic ordinary attribute isolates the block-root repair gates;
        // this is a projection budget test, not a schema-validity claim.
        let data = "A".repeat(STREAMED_XML_CONTEXT_BYTES + 4);
        for tail in ["/>", "></w:p>"] {
            let xml = format!(
                r#"<w:document xmlns:w="{W}"><w:body><w:p data="{data}"{tail}</w:body></w:document>"#,
            );
            let (blocks, _) = project(&xml);
            assert_eq!(blocks.len(), 1);
            let doc =
                roxmltree::Document::parse(std::str::from_utf8(&blocks[0].xml).unwrap()).unwrap();
            assert_eq!(doc.root_element().attribute("data"), Some(data.as_str()));
        }
    }

    #[test]
    fn namespace_declaration_explosion_still_fails_with_typed_context_limit() {
        use ooxml_common::resource::{OoxmlFormat, ResourceGovernor};

        // Five declarations cross the 4 MiB state budget while staying well
        // below the event/block ceiling. Include ignored branches: namespace
        // state is bounded even when their ordinary payload is not retained.
        let uri = "x".repeat(STREAMED_XML_CONTEXT_BYTES / 4);
        let declarations = (0..5)
            .map(|n| format!(r#" xmlns:n{n}="urn:{n}:{uri}""#))
            .collect::<String>();
        for requires in ["wps", "f"] {
            let governor = ResourceGovernor::from_wasm(OoxmlFormat::Docx, None, None, None);
            let scope = governor.scope("parse");
            let xml = format!(
                r#"<w:document xmlns:w="{W}" xmlns:mc="{MC}" xmlns:wps="{WPS}" xmlns:f="urn:future" xmlns:v="urn:schemas-microsoft-com:vml"><w:body><w:p><w:r><w:pict><mc:AlternateContent><mc:Choice Requires="{requires}"><v:shape{declarations}/></mc:Choice><mc:Fallback/></mc:AlternateContent></w:pict></w:r></w:p></w:body></w:document>"#,
            );
            let mut projector = DocumentBodyProjector::new(
                BufReader::new(Cursor::new(xml.into_bytes())),
                None,
                Rc::default(),
            );
            let error = projector.next_block().unwrap_err();
            let payload: serde_json::Value = serde_json::from_str(
                error
                    .strip_prefix("OOXML_RESOURCE_LIMIT:")
                    .expect("typed limit"),
            )
            .unwrap();
            assert_eq!(payload["code"], "ooxml-resource-limit");
            let violation = &payload["details"]["violation"];
            assert_eq!(violation["resource"], "xml-context");
            assert_eq!(violation["limit"], STREAMED_XML_CONTEXT_BYTES);
            assert_eq!(violation["observed"], STREAMED_XML_CONTEXT_BYTES + 1);
            assert_eq!(violation["part"], DOCUMENT_PART);
            assert_eq!(
                scope.resource_limit_error().as_deref(),
                Some(error.as_str())
            );
        }
    }

    #[test]
    fn rejects_one_unretained_event_larger_than_the_block_ceiling() {
        let comment = "x".repeat(HARD_MAX_DOCX_BODY_BLOCK_XML_BYTES as usize);
        let xml = format!(
            r#"<w:document xmlns:w="{W}"><w:body><w:p><!--{comment}--></w:p></w:body></w:document>"#,
        );
        let mut projector = DocumentBodyProjector::new(
            BufReader::new(Cursor::new(xml.into_bytes())),
            None,
            Rc::default(),
        );
        assert!(projector
            .next_block()
            .unwrap_err()
            .contains("limit exceeded"));
    }

    #[test]
    fn exact_block_projection_limit_is_inclusive() {
        fn paragraph_inner(projected_bytes: usize) -> String {
            let root_bytes = format!(r#"<w:p xmlns:w="{W}"></w:p>"#).len();
            let unit_overhead = "<w:r><w:t></w:t></w:r>".len();
            let mut remaining = projected_bytes - root_bytes;
            let mut inner = String::new();
            while remaining > unit_overhead + 256 * 1024 {
                inner.push_str("<w:r><w:t>");
                inner.push_str(&"x".repeat(256 * 1024));
                inner.push_str("</w:t></w:r>");
                remaining -= unit_overhead + 256 * 1024;
            }
            inner.push_str("<w:r><w:t>");
            inner.push_str(&"x".repeat(remaining - unit_overhead));
            inner.push_str("</w:t></w:r>");
            inner
        }
        let exact_inner = paragraph_inner(HARD_MAX_DOCX_BODY_BLOCK_XML_BYTES as usize);
        let exact = format!(
            r#"<w:document xmlns:w="{W}"><w:body><w:p>{exact_inner}</w:p></w:body></w:document>"#,
        );
        let mut projector = DocumentBodyProjector::new(
            BufReader::new(Cursor::new(exact.into_bytes())),
            None,
            Rc::default(),
        );
        let block = projector.next_block().unwrap().unwrap();
        assert_eq!(block.xml.len(), HARD_MAX_DOCX_BODY_BLOCK_XML_BYTES as usize);

        let over_inner = paragraph_inner(HARD_MAX_DOCX_BODY_BLOCK_XML_BYTES as usize + 1);
        let over = format!(
            r#"<w:document xmlns:w="{W}"><w:body><w:p>{over_inner}</w:p></w:body></w:document>"#,
        );
        let mut projector = DocumentBodyProjector::new(
            BufReader::new(Cursor::new(over.into_bytes())),
            None,
            Rc::default(),
        );
        assert!(projector
            .next_block()
            .unwrap_err()
            .contains("limit exceeded"));
    }
}
