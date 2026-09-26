//! Bounded, namespace-preserving projection of `word/document.xml` into
//! complete logical body blocks.
//!
//! WordprocessingML is not page-random-access: sections, fields, notes, and
//! pagination depend on the complete preceding story. This projector does not
//! pretend otherwise. It removes only the whole-part XML arena: a first pass
//! validates the part and records compact cross-block facts, while a second
//! pass can hand one complete paragraph/table/section block at a time to the
//! existing semantic parser.

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
const STREAMED_XML_EVENT_BYTES: usize = 1024 * 1024;
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
    },
    AlternateBranch {
        selected: bool,
    },
}

impl ElementFrame {
    fn children_are_visible(&self) -> bool {
        self.visible
            && !matches!(
                self.kind,
                ProcessedElementKind::Ignored
                    | ProcessedElementKind::AlternateBranch { selected: false }
            )
    }
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
    frames: Vec<ElementFrame>,
    depth: usize,
    body_content_depth: Option<usize>,
    body_seen: bool,
    body_closed: bool,
    sdts: Vec<SdtFrame>,
    capture: Option<Capture>,
    cover_break_after: Vec<bool>,
    finished: bool,
}

impl<R: BufRead> DocumentBodyProjector<R> {
    pub(crate) fn new(source: R, reporter: Option<PackageLimitReporter>) -> Self {
        Self {
            reader: BoundedXmlReader::new(source, STREAMED_XML_EVENT_BYTES, "document body"),
            reporter,
            frames: Vec::new(),
            depth: 0,
            body_content_depth: None,
            body_seen: false,
            body_closed: false,
            sdts: Vec::new(),
            capture: None,
            cover_break_after: Vec::new(),
            finished: false,
        }
    }

    pub(crate) fn next_block(&mut self) -> Result<Option<ProjectedBodyBlock>, String> {
        if self.finished {
            return Ok(None);
        }
        loop {
            let bounded = self
                .reader
                .read_event()
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

                    let visible = self
                        .frames
                        .last()
                        .is_none_or(ElementFrame::children_are_visible);
                    let (kind, mce_scope) =
                        self.classify_processed_element(namespace, &start, &context, visible)?;
                    if self.capture.is_some() {
                        if let ProcessedElementKind::Retained { opaque, .. } = &kind {
                            if visible {
                                let mut retained = start;
                                self.prepare_retained_element(
                                    &mut retained,
                                    &context,
                                    &mce_scope,
                                    *opaque,
                                )?;
                                self.append_capture(Event::Start(retained))?;
                            }
                        }
                    } else if let ProcessedElementKind::Retained { opaque, .. } = &kind {
                        if visible {
                            let mut retained = start;
                            if !*opaque {
                                bounded_xml::strip_processed_mce_attributes(
                                    self.reader.reader(),
                                    &mut retained,
                                    &mce_scope,
                                    &docx_understands_namespace,
                                    "document",
                                )?;
                            }
                            self.handle_start(
                                namespace,
                                retained,
                                &context,
                                next_depth,
                                source_offset,
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
                    let visible = self
                        .frames
                        .last()
                        .is_none_or(ElementFrame::children_are_visible);
                    let (kind, mce_scope) =
                        self.classify_processed_element(namespace, &empty, &context, visible)?;
                    if self.capture.is_some() {
                        match kind {
                            ProcessedElementKind::Retained { opaque, .. } if visible => {
                                let mut retained = empty;
                                self.prepare_retained_element(
                                    &mut retained,
                                    &context,
                                    &mce_scope,
                                    opaque,
                                )?;
                                self.append_capture(Event::Empty(retained))?;
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
                                let mut retained = empty;
                                if !opaque {
                                    bounded_xml::strip_processed_mce_attributes(
                                        self.reader.reader(),
                                        &mut retained,
                                        &mce_scope,
                                        &docx_understands_namespace,
                                        "document",
                                    )?;
                                }
                                if let Some(block) =
                                    self.handle_empty(namespace, retained, &context, source_offset)?
                                {
                                    return Ok(Some(block));
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
                    if let Some(capture) = self.capture.as_ref() {
                        let closes_capture = self.depth == capture.root_depth;
                        let retained = self.frames.last().is_some_and(|frame| {
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
                            return self.finish_capture().map(Some);
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
                                    | ProcessedElementKind::AlternateBranch { selected: true }
                            )
                    });
                    if self.capture.is_some() && retained_content {
                        self.append_capture(event)?;
                    }
                }
            }
        }
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
            let choice_understood = if local_name == "Choice" && selection_open {
                Some(self.streamed_choice_is_understood(element, context)?)
            } else {
                None
            };

            let alternate = self
                .frames
                .last_mut()
                .expect("AlternateContent parent checked");
            let ProcessedElementKind::AlternateContent {
                selected_branch,
                seen_choice,
                seen_fallback,
            } = &mut alternate.kind
            else {
                unreachable!("AlternateContent parent checked")
            };
            let selected = if local_name == "Choice" {
                if *seen_fallback {
                    return Err(
                        "document MCE Choice cannot follow Fallback in AlternateContent"
                            .to_string(),
                    );
                }
                *seen_choice = true;
                !*selected_branch
                    && choice_understood.expect("open Choice selection has a classification")
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
                !*selected_branch
            };
            if selected {
                *selected_branch = true;
                bounded_xml::validate_mce_must_understand(
                    &attributes.must_understand,
                    &docx_understands_namespace,
                    "document",
                )?;
            }
            return Ok((
                ProcessedElementKind::AlternateBranch { selected },
                attributes.scope,
            ));
        }

        if is_mc && local_name == "AlternateContent" {
            bounded_xml::validate_mce_must_understand(
                &attributes.must_understand,
                &docx_understands_namespace,
                "document",
            )?;
            return Ok((
                ProcessedElementKind::AlternateContent {
                    selected_branch: false,
                    seen_choice: false,
                    seen_fallback: false,
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
                bounded_xml::validate_mce_must_understand(
                    &attributes.must_understand,
                    &docx_understands_namespace,
                    "document",
                )?;
                return Ok((ProcessedElementKind::Unwrapped, attributes.scope));
            }
            return Ok((ProcessedElementKind::Ignored, attributes.scope));
        }

        bounded_xml::validate_mce_must_understand(
            &attributes.must_understand,
            &docx_understands_namespace,
            "document",
        )?;
        Ok((
            ProcessedElementKind::Retained {
                namespace: namespace.map(str::to_string),
                local_name,
                opaque: false,
            },
            attributes.scope,
        ))
    }

    fn prepare_retained_element(
        &self,
        element: &mut BytesStart<'static>,
        context: &NamespaceContext,
        scope: &StreamedMceScope,
        opaque: bool,
    ) -> Result<(), String> {
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
        if !physical_parent_is_retained {
            bounded_xml::inject_missing_namespaces(
                element,
                context,
                processed_parent,
                STREAMED_XML_CONTEXT_BYTES,
                "document body block",
            )
            .map_err(|error| {
                self.map_bounded_error(error, HardResourceLimitKind::XmlContextBytes)
            })?;
        }
        if !opaque {
            bounded_xml::strip_processed_mce_attributes(
                self.reader.reader(),
                element,
                scope,
                &docx_understands_namespace,
                "document",
            )?;
        }
        Ok(())
    }

    fn streamed_choice_is_understood(
        &self,
        choice: &BytesStart<'_>,
        context: &NamespaceContext,
    ) -> Result<bool, String> {
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
            ChoiceRequiresClassification::Understood => Ok(true),
            ChoiceRequiresClassification::Unsupported => Ok(false),
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
        start: BytesStart<'static>,
        context: &Rc<NamespaceContext>,
        next_depth: usize,
        source_offset: usize,
    ) -> Result<(), String> {
        let local_name = local_name(&start)?;
        let is_word = is_w_ns(namespace);
        if is_word && local_name == "body" && self.body_content_depth.is_none() {
            self.body_seen = true;
            self.body_content_depth = Some(next_depth);
            return Ok(());
        }

        self.observe_sdt_metadata(is_word, &local_name, &start, next_depth)?;
        if self.at_logical_level() && is_word && local_name == "sdt" {
            self.sdts.push(SdtFrame {
                open_depth: next_depth,
                ..SdtFrame::default()
            });
            return Ok(());
        }
        if self.at_logical_level() && is_word && is_body_block(&local_name) {
            let mut root = start;
            bounded_xml::inject_missing_namespaces(
                &mut root,
                context,
                None,
                STREAMED_XML_CONTEXT_BYTES,
                "document body block",
            )
            .map_err(|error| {
                self.map_bounded_error(error, HardResourceLimitKind::XmlContextBytes)
            })?;
            self.capture = Some(Capture {
                root_depth: next_depth,
                source_offset,
                local_name,
                xml: Vec::new(),
            });
            self.append_capture(Event::Start(root))?;
        }
        Ok(())
    }

    fn handle_empty(
        &mut self,
        namespace: Option<&str>,
        empty: BytesStart<'static>,
        context: &Rc<NamespaceContext>,
        source_offset: usize,
    ) -> Result<Option<ProjectedBodyBlock>, String> {
        let local_name = local_name(&empty)?;
        let is_word = is_w_ns(namespace);
        self.observe_sdt_metadata(is_word, &local_name, &empty, self.depth)?;
        if !(self.at_logical_level() && is_word && is_body_block(&local_name)) {
            return Ok(None);
        }
        let mut root = empty;
        bounded_xml::inject_missing_namespaces(
            &mut root,
            context,
            None,
            STREAMED_XML_CONTEXT_BYTES,
            "document body block",
        )
        .map_err(|error| self.map_bounded_error(error, HardResourceLimitKind::XmlContextBytes))?;
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
        self.append_capture(Event::Empty(root))?;
        self.finish_capture().map(Some)
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

    fn append_capture(&mut self, event: Event<'static>) -> Result<(), String> {
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

    fn finish_capture(&mut self) -> Result<ProjectedBodyBlock, String> {
        let capture = self
            .capture
            .take()
            .ok_or_else(|| "document body capture is not active".to_string())?;
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
        Ok(ProjectedBodyBlock {
            ordinal,
            source_offset: capture.source_offset,
            local_name: capture.local_name,
            xml: capture.xml,
        })
    }

    fn map_read_error(&self, error: BoundedXmlReadError) -> String {
        match error {
            BoundedXmlReadError::Xml(message) => message,
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
        let mut projector =
            DocumentBodyProjector::new(BufReader::new(Cursor::new(xml.as_bytes().to_vec())), None);
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
                <w:sdtContent><w:p/><w:sdt><w:sdtContent><w:p/><w:tbl/></w:sdtContent></w:sdt></w:sdtContent>
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
                <mc:Choice Requires="u"><w:tbl/></mc:Choice>
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
        let mut projector =
            DocumentBodyProjector::new(BufReader::new(Cursor::new(rejected.into_bytes())), None);
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
                  <w:sdtContent><w:p/><w:tbl/></w:sdtContent>
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
                <mc:Choice Requires="u" mc:MustUnderstand="u"><w:tbl/></mc:Choice>
                <mc:Fallback><w:p/></mc:Fallback>
              </mc:AlternateContent>
            </w:body></w:document>"#,
        );
        let (blocks, _) = project(&xml);
        assert_eq!(blocks[0].local_name, "p");

        let mismatch = format!(
            r#"<w:document xmlns:w="{W}" xmlns:mc="{MC}" xmlns:u="urn:unsupported" mc:MustUnderstand="u"><w:body><w:p/></w:body></w:document>"#,
        );
        let mut projector =
            DocumentBodyProjector::new(BufReader::new(Cursor::new(mismatch.into_bytes())), None);
        assert!(projector
            .next_block()
            .unwrap_err()
            .contains("MustUnderstand namespace is not understood"));
    }

    #[test]
    fn missing_body_is_fatal_for_the_streamed_required_part() {
        let xml = format!(r#"<w:document xmlns:w="{W}"/>"#);
        let mut projector =
            DocumentBodyProjector::new(BufReader::new(Cursor::new(xml.into_bytes())), None);
        assert!(projector.next_block().unwrap_err().contains("no <w:body>"));
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
        let mut projector =
            DocumentBodyProjector::new(BufReader::new(Cursor::new(exact.into_bytes())), None);
        let block = projector.next_block().unwrap().unwrap();
        assert_eq!(block.xml.len(), HARD_MAX_DOCX_BODY_BLOCK_XML_BYTES as usize);

        let over_inner = paragraph_inner(HARD_MAX_DOCX_BODY_BLOCK_XML_BYTES as usize + 1);
        let over = format!(
            r#"<w:document xmlns:w="{W}"><w:body><w:p>{over_inner}</w:p></w:body></w:document>"#,
        );
        let mut projector =
            DocumentBodyProjector::new(BufReader::new(Cursor::new(over.into_bytes())), None);
        assert!(projector
            .next_block()
            .unwrap_err()
            .contains("limit exceeded"));
    }
}
