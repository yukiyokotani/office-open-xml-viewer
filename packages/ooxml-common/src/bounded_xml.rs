//! Format-neutral bounded XML streaming primitives.
//!
//! Host-schema state machines remain in the consuming format crates. This
//! module owns only lexical event bounds, source positions, inherited namespace
//! context, bounded event projection, and the namespace-based core of MCE
//! `Choice` selection.

use crate::mce::ChoiceRequiresClassification;
use crate::package_session::PackageLimitReporter;
use crate::resource::{observe_hard_limit, HardResourceLimitKind};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{QName, ResolveResult};
use quick_xml::{NsReader, Writer, XmlVersion};
use std::collections::HashSet;
use std::io::{BufRead, Read};
use std::rc::Rc;

pub const MCE_NS: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XmlSourceSpan {
    pub start: u64,
    pub end: u64,
}

#[derive(Debug)]
pub struct BoundedXmlEvent {
    pub namespace: Option<String>,
    pub event: Event<'static>,
    pub span: XmlSourceSpan,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum BoundedXmlReadError {
    Xml(String),
    EventLimit { limit: usize, observed: usize },
}

/// A `BufRead` adapter whose allowance is reset before every quick-xml event.
/// It prevents one unterminated or hostile token from growing quick-xml's
/// caller-owned event buffer to the complete part size.
pub struct EventBudgetReader<R> {
    inner: R,
    limit: usize,
    remaining: usize,
    exceeded: bool,
    current_event_bytes: usize,
    max_event_bytes: usize,
}

impl<R> EventBudgetReader<R> {
    pub fn new(inner: R, limit: usize) -> Self {
        Self {
            inner,
            limit,
            remaining: limit,
            exceeded: false,
            current_event_bytes: 0,
            max_event_bytes: 0,
        }
    }

    fn reset_event_budget(&mut self) {
        self.remaining = self.limit;
        self.exceeded = false;
        self.current_event_bytes = 0;
    }

    pub fn max_event_bytes(&self) -> usize {
        self.max_event_bytes
    }
}

impl<R: BufRead> Read for EventBudgetReader<R> {
    fn read(&mut self, output: &mut [u8]) -> std::io::Result<usize> {
        let available = self.fill_buf()?;
        let count = available.len().min(output.len());
        output[..count].copy_from_slice(&available[..count]);
        self.consume(count);
        Ok(count)
    }
}

impl<R: BufRead> BufRead for EventBudgetReader<R> {
    fn fill_buf(&mut self) -> std::io::Result<&[u8]> {
        if self.remaining == 0 {
            if self.inner.fill_buf()?.is_empty() {
                return Ok(&[]);
            }
            self.exceeded = true;
            return Err(std::io::Error::other(
                "XML event exceeded its hidden byte ceiling",
            ));
        }
        let available = self.inner.fill_buf()?;
        Ok(&available[..available.len().min(self.remaining)])
    }

    fn consume(&mut self, amount: usize) {
        let consumed = amount.min(self.remaining);
        self.inner.consume(consumed);
        self.remaining -= consumed;
        self.current_event_bytes = self.current_event_bytes.saturating_add(consumed);
        self.max_event_bytes = self.max_event_bytes.max(self.current_event_bytes);
    }
}

/// Namespace-aware, source-positioned event reader over one bounded XML source.
pub struct BoundedXmlReader<R: BufRead> {
    reader: NsReader<EventBudgetReader<R>>,
    event_buf: Vec<u8>,
    label: &'static str,
    terminal: Option<BoundedXmlReadError>,
}

impl<R: BufRead> BoundedXmlReader<R> {
    pub fn new(source: R, event_limit: usize, label: &'static str) -> Self {
        let mut reader = NsReader::from_reader(EventBudgetReader::new(source, event_limit));
        reader.config_mut().trim_text(false);
        Self {
            reader,
            event_buf: Vec::with_capacity(event_limit),
            label,
            terminal: None,
        }
    }

    pub fn reader(&self) -> &NsReader<EventBudgetReader<R>> {
        &self.reader
    }

    pub fn max_event_bytes(&self) -> usize {
        self.reader.get_ref().max_event_bytes()
    }

    pub fn read_event(&mut self) -> Result<BoundedXmlEvent, BoundedXmlReadError> {
        if let Some(error) = &self.terminal {
            return Err(error.clone());
        }
        self.reader.get_mut().reset_event_budget();
        let start = self.reader.buffer_position();
        let read = self.reader.read_resolved_event_into(&mut self.event_buf);
        let (namespace, event) = match read {
            Ok(read) => read,
            Err(error) => {
                let terminal = if self.reader.get_ref().exceeded {
                    let limit = self.reader.get_ref().limit;
                    BoundedXmlReadError::EventLimit {
                        limit,
                        observed: limit.saturating_add(1),
                    }
                } else {
                    BoundedXmlReadError::Xml(format!("{} XML stream: {error}", self.label))
                };
                self.terminal = Some(terminal.clone());
                return Err(terminal);
            }
        };
        let namespace = match resolved_namespace_uri(&namespace, self.label) {
            Ok(namespace) => namespace,
            Err(error) => {
                let terminal = BoundedXmlReadError::Xml(error);
                self.terminal = Some(terminal.clone());
                return Err(terminal);
            }
        };
        let event = event.into_owned();
        let end = self.reader.buffer_position();
        self.event_buf.clear();
        Ok(BoundedXmlEvent {
            namespace,
            event,
            span: XmlSourceSpan { start, end },
        })
    }
}

pub fn resolved_namespace_is(
    namespace: &ResolveResult<'_>,
    predicate: impl FnOnce(&str) -> bool,
) -> bool {
    match namespace {
        ResolveResult::Bound(namespace) => std::str::from_utf8(namespace.as_ref())
            .ok()
            .map(predicate)
            .unwrap_or(false),
        ResolveResult::Unbound | ResolveResult::Unknown(_) => false,
    }
}

fn resolved_namespace_uri(
    namespace: &ResolveResult<'_>,
    label: &str,
) -> Result<Option<String>, String> {
    match namespace {
        ResolveResult::Bound(namespace) => {
            let namespace = std::str::from_utf8(namespace.as_ref()).map_err(|e| e.to_string())?;
            quick_xml::escape::unescape(namespace)
                .map(|namespace| Some(namespace.into_owned()))
                .map_err(|e| format!("{label} namespace URI: {e}"))
        }
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Unknown(prefix) => Err(format!(
            "{label} XML uses unbound namespace prefix {}",
            String::from_utf8_lossy(prefix.as_ref())
        )),
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum BoundedXmlError {
    Xml(String),
    Limit { limit: usize, observed: usize },
}

impl From<String> for BoundedXmlError {
    fn from(error: String) -> Self {
        Self::Xml(error)
    }
}

/// Persistent in-scope namespace declarations. A derived context stores only
/// declarations introduced by one element and shares its parent by `Rc`.
#[derive(Default)]
pub struct NamespaceContext {
    parent: Option<Rc<NamespaceContext>>,
    declarations: Vec<(Option<String>, String)>,
    active_bytes: usize,
}

impl NamespaceContext {
    pub fn root() -> Rc<Self> {
        Rc::new(Self::default())
    }

    pub fn derive(
        element: &BytesStart<'_>,
        inherited: &Rc<Self>,
        context_limit: usize,
        label: &str,
    ) -> Result<Rc<Self>, BoundedXmlError> {
        let mut local_context_bytes = 0usize;
        for attribute in element.attributes() {
            let attribute = attribute.map_err(|error| {
                BoundedXmlError::Xml(format!("{label} namespace attribute: {error}"))
            })?;
            let key = attribute.key.as_ref();
            if key == b"xmlns" || key.starts_with(b"xmlns:") {
                local_context_bytes = local_context_bytes
                    .checked_add(key.len())
                    .and_then(|total| total.checked_add(attribute.value.len()))
                    .unwrap_or(usize::MAX);
            }
        }
        if local_context_bytes == 0 {
            return Ok(Rc::clone(inherited));
        }
        let active_bytes = inherited.active_bytes.saturating_add(local_context_bytes);
        if active_bytes > context_limit {
            return Err(BoundedXmlError::Limit {
                limit: context_limit,
                observed: active_bytes,
            });
        }
        let mut declarations = Vec::new();
        for attribute in element.attributes() {
            let attribute = attribute.map_err(|error| {
                BoundedXmlError::Xml(format!("{label} namespace attribute: {error}"))
            })?;
            let key = attribute.key.as_ref();
            let prefix = if key == b"xmlns" {
                Some(None)
            } else if let Some(prefix) = key.strip_prefix(b"xmlns:") {
                Some(Some(
                    std::str::from_utf8(prefix)
                        .map_err(|error| BoundedXmlError::Xml(error.to_string()))?
                        .to_string(),
                ))
            } else {
                None
            };
            if let Some(prefix) = prefix {
                let namespace = attribute
                    .normalized_value(XmlVersion::Implicit1_0)
                    .map_err(|error| BoundedXmlError::Xml(error.to_string()))?
                    .into_owned();
                declarations.push((prefix, namespace));
            }
        }
        Ok(Rc::new(Self {
            parent: Some(Rc::clone(inherited)),
            declarations,
            active_bytes,
        }))
    }

    pub fn active_bytes(&self) -> usize {
        self.active_bytes
    }

    pub fn namespace_for_prefix(&self, prefix: &str) -> Option<&str> {
        self.declarations
            .iter()
            .rev()
            .find_map(|(candidate, namespace)| {
                (candidate.as_deref() == Some(prefix)).then_some(namespace.as_str())
            })
            .or_else(|| {
                self.parent
                    .as_deref()
                    .and_then(|parent| parent.namespace_for_prefix(prefix))
            })
    }

    pub fn effective_bindings(&self) -> Vec<(Option<String>, String)> {
        let mut chain = Vec::new();
        let mut context = Some(self);
        while let Some(current) = context {
            chain.push(current);
            context = current.parent.as_deref();
        }
        let mut bindings: Vec<(Option<String>, String)> = Vec::new();
        for current in chain.into_iter().rev() {
            for (prefix, namespace) in &current.declarations {
                if let Some(existing) = bindings
                    .iter_mut()
                    .find(|(candidate, _)| candidate == prefix)
                {
                    existing.1.clone_from(namespace);
                } else {
                    bindings.push((prefix.clone(), namespace.clone()));
                }
            }
        }
        bindings
    }
}

/// Persistent namespace-resolved MCE state. Host-schema state machines and the
/// set of namespaces each format understands remain consuming-format concerns.
#[derive(Default)]
pub struct MceScope {
    parent: Option<Rc<MceScope>>,
    ignorable: HashSet<String>,
    process_content: HashSet<(String, String)>,
    active_bytes: usize,
}

/// Namespace-resolved compatibility-control attributes for one physical XML
/// element. This is the format-neutral Part 3 §§7.2–7.4 input to a host's MCE
/// state machine.
pub struct MceAttributes {
    pub scope: Rc<MceScope>,
    pub must_understand: Vec<String>,
}

/// Parse `mc:Ignorable`, `mc:ProcessContent`, and `mc:MustUnderstand`, resolve
/// their prefixes/QNames against the active namespace context, validate the
/// cross-attribute constraints, and derive the persistent scope for children.
pub fn derive_mce_attributes<R>(
    reader: &NsReader<R>,
    element: &BytesStart<'_>,
    inherited: &Rc<MceScope>,
    context: &NamespaceContext,
    context_limit: usize,
    label: &str,
) -> Result<MceAttributes, BoundedXmlError> {
    let mut ignorable = HashSet::new();
    let mut process_content = HashSet::new();
    let mut must_understand = Vec::new();
    let mut declared_process_content = Vec::new();
    let base_context_bytes = context
        .active_bytes()
        .saturating_add(inherited.active_bytes());
    let mut derived_bytes = 0usize;
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| BoundedXmlError::Xml(format!("{label} MCE attribute: {error}")))?;
        let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
        if !resolved_namespace_is(&namespace, |namespace| namespace == MCE_NS) {
            continue;
        }
        let value = attribute
            .normalized_value(XmlVersion::Implicit1_0)
            .map_err(|error| BoundedXmlError::Xml(format!("{label} MCE attribute value: {error}")))?
            .into_owned();
        match local_name.as_ref() {
            b"Ignorable" => {
                ignorable.extend(resolve_mce_prefixes(
                    &value,
                    context,
                    "Ignorable",
                    base_context_bytes,
                    &mut derived_bytes,
                    context_limit,
                    label,
                )?);
            }
            b"ProcessContent" => {
                declared_process_content.extend(resolve_mce_qnames(
                    &value,
                    context,
                    "ProcessContent",
                    base_context_bytes,
                    &mut derived_bytes,
                    context_limit,
                    label,
                )?);
            }
            b"MustUnderstand" => {
                must_understand = resolve_mce_prefixes(
                    &value,
                    context,
                    "MustUnderstand",
                    base_context_bytes,
                    &mut derived_bytes,
                    context_limit,
                    label,
                )?;
            }
            _ => {}
        }
    }
    for (namespace, local_name) in declared_process_content {
        if namespace == MCE_NS
            || !(ignorable.contains(&namespace) || inherited.is_ignorable(&namespace))
        {
            return Err(BoundedXmlError::Xml(format!(
                "{label} MCE ProcessContent namespace must be declared Ignorable: {namespace}"
            )));
        }
        process_content.insert((namespace, local_name));
    }
    Ok(MceAttributes {
        scope: MceScope::derive(inherited, ignorable, process_content, derived_bytes),
        must_understand,
    })
}

fn resolve_mce_prefixes(
    value: &str,
    context: &NamespaceContext,
    attribute_name: &str,
    base_context_bytes: usize,
    derived_bytes: &mut usize,
    context_limit: usize,
    label: &str,
) -> Result<Vec<String>, BoundedXmlError> {
    value
        .split_whitespace()
        .map(|prefix| {
            let namespace = context.namespace_for_prefix(prefix).ok_or_else(|| {
                BoundedXmlError::Xml(format!(
                    "{label} MCE {attribute_name} uses unbound namespace prefix: {prefix}"
                ))
            })?;
            if namespace == MCE_NS {
                return Err(BoundedXmlError::Xml(format!(
                    "{label} MCE {attribute_name} must not name the Markup Compatibility namespace"
                )));
            }
            charge_mce_context(
                base_context_bytes,
                derived_bytes,
                namespace.len(),
                context_limit,
            )?;
            Ok(namespace.to_string())
        })
        .collect()
}

fn resolve_mce_qnames(
    value: &str,
    context: &NamespaceContext,
    attribute_name: &str,
    base_context_bytes: usize,
    derived_bytes: &mut usize,
    context_limit: usize,
    label: &str,
) -> Result<Vec<(String, String)>, BoundedXmlError> {
    value
        .split_whitespace()
        .map(|name| {
            let (prefix, local_name) = name.split_once(':').ok_or_else(|| {
                BoundedXmlError::Xml(format!(
                    "{label} MCE {attribute_name} name must be namespace-qualified: {name}"
                ))
            })?;
            if prefix.is_empty() || local_name.is_empty() || local_name.contains(':') {
                return Err(BoundedXmlError::Xml(format!(
                    "{label} MCE {attribute_name} contains invalid QName: {name}"
                )));
            }
            let namespace = context.namespace_for_prefix(prefix).ok_or_else(|| {
                BoundedXmlError::Xml(format!(
                    "{label} MCE {attribute_name} uses unbound namespace prefix: {prefix}"
                ))
            })?;
            charge_mce_context(
                base_context_bytes,
                derived_bytes,
                namespace.len().saturating_add(local_name.len()),
                context_limit,
            )?;
            Ok((namespace.to_string(), local_name.to_string()))
        })
        .collect()
}

fn charge_mce_context(
    base_context_bytes: usize,
    derived_bytes: &mut usize,
    additional_bytes: usize,
    context_limit: usize,
) -> Result<(), BoundedXmlError> {
    let next_derived = derived_bytes
        .checked_add(additional_bytes)
        .unwrap_or(usize::MAX);
    let observed = base_context_bytes.saturating_add(next_derived);
    if observed > context_limit {
        return Err(BoundedXmlError::Limit {
            limit: context_limit,
            observed,
        });
    }
    *derived_bytes = next_derived;
    Ok(())
}

/// Apply Part 3 §9.4 MustUnderstand processing with a host-supplied application
/// configuration.
pub fn validate_mce_must_understand(
    namespaces: &[String],
    understands: &dyn Fn(&str) -> bool,
    label: &str,
) -> Result<(), String> {
    if let Some(namespace) = namespaces.iter().find(|namespace| !understands(namespace)) {
        Err(format!(
            "{label} MCE MustUnderstand namespace is not understood: {namespace}"
        ))
    } else {
        Ok(())
    }
}

/// Validate the Part 3 §§7.5–7.7 attribute grammar shared by
/// AlternateContent, Choice, and Fallback.
pub fn validate_mce_alternate_element_attributes<R>(
    reader: &NsReader<R>,
    element: &BytesStart<'_>,
    local_name: &str,
    scope: &MceScope,
    label: &str,
) -> Result<(), String> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| format!("{label} MCE attribute: {error}"))?;
        let raw_name = attribute.key.as_ref();
        if raw_name == b"xmlns" || raw_name.starts_with(b"xmlns:") {
            continue;
        }
        if !raw_name.contains(&b':') {
            if local_name == "Choice" && raw_name == b"Requires" {
                continue;
            }
            return Err(format!(
                "{label} MCE {local_name} has a forbidden unqualified attribute"
            ));
        }
        let (namespace, _) = reader.resolver().resolve_attribute(attribute.key);
        let allowed = match namespace {
            ResolveResult::Bound(namespace) => {
                let namespace =
                    std::str::from_utf8(namespace.as_ref()).map_err(|error| error.to_string())?;
                namespace == MCE_NS || scope.is_ignorable(namespace)
            }
            _ => false,
        };
        if !allowed {
            return Err(format!(
                "{label} MCE {local_name} qualified attribute namespace is not MCE or ignorable"
            ));
        }
    }
    Ok(())
}

/// Part 3 §9.2 forbids XML context attributes on an element unwrapped through
/// ProcessContent because removing that element would change their semantics.
pub fn validate_mce_process_content_element<R>(
    reader: &NsReader<R>,
    element: &BytesStart<'_>,
    label: &str,
) -> Result<(), String> {
    const XML_NS: &str = "http://www.w3.org/XML/1998/namespace";
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| format!("{label} MCE attribute: {error}"))?;
        let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
        if resolved_namespace_is(&namespace, |namespace| namespace == XML_NS)
            && matches!(local_name.as_ref(), b"base" | b"lang" | b"space")
        {
            return Err(format!(
                "{label} MCE ProcessContent element cannot carry xml:base, xml:lang, or xml:space"
            ));
        }
    }
    Ok(())
}

/// Produce the Part 3 §9.4 step-5 attribute set for an ordinary retained
/// element. Compatibility-control attributes and attributes in unsupported
/// ignorable namespaces do not belong to the processed infoset. Host-owned
/// application-defined extension elements must bypass this function because
/// their complete subtree is opaque to MCE processing.
pub fn strip_processed_mce_attributes<R>(
    reader: &NsReader<R>,
    element: &mut BytesStart<'static>,
    scope: &MceScope,
    understands: &dyn Fn(&str) -> bool,
    label: &str,
) -> Result<(), String> {
    let name = std::str::from_utf8(element.name().as_ref())
        .map_err(|error| error.to_string())?
        .to_string();
    let mut retained = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| format!("{label} MCE attribute: {error}"))?;
        let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
        let compatibility_control = resolved_namespace_is(&namespace, |namespace| {
            namespace == MCE_NS
                && matches!(
                    local_name.as_ref(),
                    b"Ignorable" | b"ProcessContent" | b"MustUnderstand"
                )
        });
        let unsupported_ignorable = match namespace {
            ResolveResult::Bound(namespace) => {
                let namespace =
                    std::str::from_utf8(namespace.as_ref()).map_err(|error| error.to_string())?;
                scope.is_ignorable(namespace) && !understands(namespace)
            }
            _ => false,
        };
        if !compatibility_control && !unsupported_ignorable {
            retained.push(attribute.to_owned());
        }
    }
    let mut processed = BytesStart::new(name);
    for mut attribute in retained {
        // XML 1.0 §2.3 permits a literal double quote inside a single-quoted
        // attribute. quick-xml's Attribute retains the raw value but
        // push_attribute always writes double-quoted delimiters. Escape only
        // that newly conflicting delimiter: decoding/re-escaping the value
        // would alter existing entities and character-reference whitespace.
        // Existing element/context and projection gates charge the expanded
        // bytes; their limits remain authoritative.
        let quote_count = attribute.value.iter().filter(|&&byte| byte == b'"').count();
        if quote_count != 0 {
            let escaped_len = quote_count
                .checked_mul(5)
                .and_then(|extra| attribute.value.len().checked_add(extra))
                .ok_or_else(|| format!("{label} XML attribute escaped length overflow"))?;
            let mut value = Vec::with_capacity(escaped_len);
            for &byte in attribute.value.as_ref() {
                if byte == b'"' {
                    value.extend_from_slice(b"&quot;");
                } else {
                    value.push(byte);
                }
            }
            attribute.value = std::borrow::Cow::Owned(value);
        }
        processed.push_attribute(attribute);
    }
    *element = processed;
    Ok(())
}

impl MceScope {
    pub fn root() -> Rc<Self> {
        Rc::new(Self::default())
    }

    pub fn derive(
        inherited: &Rc<Self>,
        ignorable: HashSet<String>,
        process_content: HashSet<(String, String)>,
        derived_bytes: usize,
    ) -> Rc<Self> {
        if ignorable.is_empty() && process_content.is_empty() {
            return Rc::clone(inherited);
        }
        Rc::new(Self {
            parent: Some(Rc::clone(inherited)),
            ignorable,
            process_content,
            active_bytes: inherited.active_bytes.saturating_add(derived_bytes),
        })
    }

    pub fn active_bytes(&self) -> usize {
        self.active_bytes
    }

    pub fn is_ignorable(&self, namespace: &str) -> bool {
        self.ignorable.contains(namespace)
            || self
                .parent
                .as_deref()
                .is_some_and(|parent| parent.is_ignorable(namespace))
    }

    pub fn processes_content(&self, namespace: &str, local_name: &str) -> bool {
        self.process_content
            .iter()
            .any(|(candidate_namespace, candidate_local)| {
                candidate_namespace == namespace
                    && (candidate_local == local_name || candidate_local == "*")
            })
            || self
                .parent
                .as_deref()
                .is_some_and(|parent| parent.processes_content(namespace, local_name))
    }
}

/// Repair namespace declarations when preprocessing removes physical XML
/// ancestors but retains and serializes one of their descendants.
pub fn inject_missing_namespaces(
    element: &mut BytesStart<'static>,
    element_context: &NamespaceContext,
    processed_parent_context: Option<&NamespaceContext>,
    context_limit: usize,
    label: &str,
) -> Result<(), BoundedXmlError> {
    let mut locally_declared = HashSet::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            BoundedXmlError::Xml(format!("{label} namespace attribute: {error}"))
        })?;
        let key = attribute.key.as_ref();
        if key == b"xmlns" {
            locally_declared.insert(None);
        } else if let Some(prefix) = key.strip_prefix(b"xmlns:") {
            locally_declared.insert(Some(
                std::str::from_utf8(prefix)
                    .map_err(|error| BoundedXmlError::Xml(error.to_string()))?
                    .to_string(),
            ));
        }
    }
    let element_namespaces = element_context.effective_bindings();
    let processed_parent_namespaces = processed_parent_context
        .map(NamespaceContext::effective_bindings)
        .unwrap_or_default();
    let mut missing = Vec::new();
    let mut injected_bytes = 0usize;
    for (prefix, namespace) in element_namespaces {
        let parent_preserves_binding =
            processed_parent_namespaces
                .iter()
                .any(|(parent_prefix, parent_namespace)| {
                    parent_prefix == &prefix && parent_namespace == &namespace
                });
        if locally_declared.contains(&prefix) || parent_preserves_binding {
            continue;
        }
        let escaped_namespace_bytes = quick_xml::escape::escape(&namespace).len();
        injected_bytes = injected_bytes
            .checked_add(escaped_namespace_bytes)
            .and_then(|total| total.checked_add(prefix.as_ref().map_or(9, |p| p.len() + 10)))
            .unwrap_or(usize::MAX);
        missing.push((prefix, namespace));
    }
    let observed = element.len().saturating_add(injected_bytes);
    if observed > context_limit {
        return Err(BoundedXmlError::Limit {
            limit: context_limit,
            observed,
        });
    }
    // Preflight every declaration before mutating the start tag, so a limit
    // failure cannot leave a partially repaired fragment.
    for (prefix, namespace) in &missing {
        match prefix {
            Some(prefix) => {
                let key = format!("xmlns:{prefix}");
                element.push_attribute((key.as_str(), namespace.as_str()));
            }
            None => element.push_attribute(("xmlns", namespace.as_str())),
        }
    }
    Ok(())
}

pub fn projected_event_bytes(event: &Event<'_>) -> usize {
    match event {
        Event::Start(event) => event.len().saturating_add(2),
        Event::End(event) => event.len().saturating_add(3),
        Event::Empty(event) => event.len().saturating_add(3),
        Event::Text(event) => event.len(),
        Event::Comment(event) => event.len().saturating_add(7),
        Event::CData(event) => event.len().saturating_add(12),
        Event::Decl(event) => event.len().saturating_add(4),
        Event::PI(event) => event.len().saturating_add(4),
        Event::DocType(event) => event.len().saturating_add(11),
        Event::GeneralRef(event) => event.len().saturating_add(2),
        Event::Eof => 0,
    }
}

pub fn append_projected_event(
    target: &mut Vec<u8>,
    event: &Event<'_>,
    retained_overhead: usize,
    limit: usize,
    label: &str,
) -> Result<(), BoundedXmlError> {
    let event_bytes = projected_event_bytes(event);
    let observed = retained_overhead
        .checked_add(target.len())
        .and_then(|current| current.checked_add(event_bytes))
        .unwrap_or(usize::MAX);
    if observed > limit {
        return Err(BoundedXmlError::Limit { limit, observed });
    }
    target
        .try_reserve_exact(event_bytes)
        .map_err(|error| BoundedXmlError::Xml(error.to_string()))?;
    Writer::new(target)
        .write_event(event.borrow())
        .map_err(|error| BoundedXmlError::Xml(format!("{label} XML projection: {error}")))
}

/// Parse and neutrally classify the namespace-only core of MCE `Choice`
/// selection. The host decides whether missing/blank non-conformance is fatal.
pub fn classify_mce_choice_requires(
    choice: &BytesStart<'_>,
    context: &NamespaceContext,
    understands: &dyn Fn(&str) -> bool,
    label: &str,
) -> Result<ChoiceRequiresClassification, String> {
    classify_bounded_mce_choice_requires(choice, context, understands, usize::MAX, label).map_err(
        |error| match error {
            BoundedXmlError::Xml(message) => message,
            BoundedXmlError::Limit { limit, observed } => {
                format!("{label} MCE context limit exceeded: {observed} > {limit}")
            }
        },
    )
}

/// Bounded variant of [`classify_mce_choice_requires`] for streaming hosts.
/// Prefix resolution and the host application-configuration predicate are
/// shared, while the caller supplies the active context ceiling.
pub fn classify_bounded_mce_choice_requires(
    choice: &BytesStart<'_>,
    context: &NamespaceContext,
    understands: &dyn Fn(&str) -> bool,
    context_limit: usize,
    label: &str,
) -> Result<ChoiceRequiresClassification, BoundedXmlError> {
    let mut requires = None;
    for attribute in choice.attributes() {
        let attribute = attribute.map_err(|error| {
            BoundedXmlError::Xml(format!("{label} MCE Choice attribute: {error}"))
        })?;
        if attribute.key == QName(b"Requires") {
            requires = Some(
                attribute
                    .normalized_value(XmlVersion::Implicit1_0)
                    .map_err(|error| {
                        BoundedXmlError::Xml(format!("{label} MCE Requires: {error}"))
                    })?
                    .into_owned(),
            );
        }
    }
    let Some(requires) = requires.as_deref() else {
        return Ok(ChoiceRequiresClassification::Missing);
    };
    if requires.split_whitespace().next().is_none() {
        return Ok(ChoiceRequiresClassification::Blank);
    }

    let mut derived_bytes = 0usize;
    let mut unresolved = false;
    let mut unsupported = false;
    for prefix in requires.split_whitespace() {
        let Some(namespace) = context.namespace_for_prefix(prefix) else {
            unresolved = true;
            continue;
        };
        if namespace == MCE_NS {
            return Err(BoundedXmlError::Xml(format!(
                "{label} MCE Choice Requires must not name the Markup Compatibility namespace"
            )));
        }
        charge_mce_context(
            context.active_bytes(),
            &mut derived_bytes,
            namespace.len(),
            context_limit,
        )?;
        unsupported |= !understands(namespace);
    }
    Ok(if unresolved {
        ChoiceRequiresClassification::Unresolved
    } else if unsupported {
        ChoiceRequiresClassification::Unsupported
    } else {
        ChoiceRequiresClassification::Understood
    })
}

/// Attribute a format-owned hard XML ceiling to its active package operation.
pub fn report_hard_limit(
    reporter: Option<&PackageLimitReporter>,
    kind: HardResourceLimitKind,
    part: Option<&str>,
    limit: usize,
    observed: usize,
) -> Result<(), String> {
    let limit = u64::try_from(limit).unwrap_or(u64::MAX);
    let observed = u64::try_from(observed).unwrap_or(u64::MAX);
    if let Some(reporter) = reporter {
        reporter.observe_hard_limit(kind, part, limit, observed)
    } else {
        observe_hard_limit(kind, part, limit, observed)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use quick_xml::events::Event;
    use std::io::{BufReader, Cursor};

    #[test]
    fn event_reader_bounds_one_token_and_reports_source_spans() {
        let xml = br#"<r xmlns:a="urn:a"><a:x>v</a:x></r>"#;
        let source = BufReader::with_capacity(1, Cursor::new(xml.as_slice()));
        let mut reader = BoundedXmlReader::new(source, 64, "test");

        let root = reader.read_event().unwrap();
        assert!(matches!(root.event, Event::Start(_)));
        assert_eq!(root.span.start, 0);
        assert!(root.span.end > root.span.start);
        assert!(reader.event_buf.is_empty());

        let child = reader.read_event().unwrap();
        assert_eq!(child.namespace.as_deref(), Some("urn:a"));
        assert!(child.span.start >= root.span.end);

        let over = BufReader::new(Cursor::new(b"<r>12345</r>".as_slice()));
        let mut over = BoundedXmlReader::new(over, 4, "test");
        assert!(matches!(over.read_event().unwrap().event, Event::Start(_)));
        let terminal = over.read_event().unwrap_err();
        assert!(matches!(
            terminal,
            BoundedXmlReadError::EventLimit {
                limit: 4,
                observed: 5
            }
        ));
        let buffered = over.event_buf.clone();
        let position = over.reader.buffer_position();
        assert!(!buffered.is_empty());
        for _ in 0..3 {
            assert_eq!(over.read_event().unwrap_err(), terminal);
            assert_eq!(over.event_buf, buffered);
            assert_eq!(over.reader.buffer_position(), position);
            assert_eq!(over.max_event_bytes(), 4);
        }
    }

    #[test]
    fn xml_error_is_terminal_and_does_not_reuse_the_event_buffer() {
        let source = BufReader::new(Cursor::new(b"<u:x/>tail".as_slice()));
        let mut reader = BoundedXmlReader::new(source, 64, "test");

        let terminal = reader.read_event().unwrap_err();
        assert!(matches!(terminal, BoundedXmlReadError::Xml(_)));
        let buffered = reader.event_buf.clone();
        let position = reader.reader.buffer_position();
        assert!(!buffered.is_empty());
        for _ in 0..3 {
            assert_eq!(reader.read_event().unwrap_err(), terminal);
            assert_eq!(reader.event_buf, buffered);
            assert_eq!(reader.reader.buffer_position(), position);
        }

        let source = BufReader::new(Cursor::new(b"<r><".as_slice()));
        let mut malformed = BoundedXmlReader::new(source, 64, "test");
        assert!(matches!(
            malformed.read_event().unwrap().event,
            Event::Start(_)
        ));
        let terminal = malformed.read_event().unwrap_err();
        let buffered = malformed.event_buf.clone();
        let position = malformed.reader.buffer_position();
        assert!(matches!(terminal, BoundedXmlReadError::Xml(_)));
        assert_eq!(malformed.read_event().unwrap_err(), terminal);
        assert_eq!(malformed.event_buf, buffered);
        assert_eq!(malformed.reader.buffer_position(), position);
    }

    #[test]
    fn namespace_context_is_persistent_and_rebinding_is_effective() {
        let root = NamespaceContext::root();
        let first = quick_xml::events::BytesStart::from_content(
            r#"r xmlns="urn:default" xmlns:a="urn:a""#,
            1,
        );
        let first = NamespaceContext::derive(&first, &root, 1024, "test").unwrap();
        let second = quick_xml::events::BytesStart::from_content(r#"a:x xmlns:a="urn:other""#, 3);
        let second = NamespaceContext::derive(&second, &first, 1024, "test").unwrap();

        assert_eq!(first.namespace_for_prefix("a"), Some("urn:a"));
        assert_eq!(second.namespace_for_prefix("a"), Some("urn:other"));
        assert_eq!(
            second.effective_bindings(),
            vec![
                (None, "urn:default".to_string()),
                (Some("a".to_string()), "urn:other".to_string())
            ]
        );
    }

    #[test]
    fn event_projection_preflights_the_complete_write() {
        let event = Event::Text(quick_xml::events::BytesText::new("abc"));
        let mut exact = Vec::new();
        append_projected_event(&mut exact, &event, 2, 5, "test").unwrap();
        assert_eq!(exact, b"abc");

        let mut over = Vec::new();
        assert_eq!(
            append_projected_event(&mut over, &event, 3, 5, "test"),
            Err(BoundedXmlError::Limit {
                limit: 5,
                observed: 6
            })
        );
        assert!(over.is_empty());
    }

    #[test]
    fn mce_choice_uses_inherited_prefixes_and_the_host_predicate() {
        let root = NamespaceContext::root();
        let declarations = quick_xml::events::BytesStart::from_content(
            r#"r xmlns:k="urn:known" xmlns:u="urn:unknown""#,
            1,
        );
        let context = NamespaceContext::derive(&declarations, &root, 1024, "test").unwrap();
        let known = quick_xml::events::BytesStart::from_content(r#"mc:Choice Requires="k""#, 9);
        let mixed = quick_xml::events::BytesStart::from_content(r#"mc:Choice Requires="k u""#, 9);
        let missing = quick_xml::events::BytesStart::new("mc:Choice");
        let blank = quick_xml::events::BytesStart::from_content(r#"mc:Choice Requires="   ""#, 9);
        let unbound =
            quick_xml::events::BytesStart::from_content(r#"mc:Choice Requires="missing""#, 9);
        let understands = |namespace: &str| namespace == "urn:known";

        assert_eq!(
            classify_mce_choice_requires(&known, &context, &understands, "test").unwrap(),
            ChoiceRequiresClassification::Understood
        );
        assert_eq!(
            classify_mce_choice_requires(&mixed, &context, &understands, "test").unwrap(),
            ChoiceRequiresClassification::Unsupported
        );
        assert_eq!(
            classify_mce_choice_requires(&missing, &context, &understands, "test").unwrap(),
            ChoiceRequiresClassification::Missing
        );
        assert_eq!(
            classify_mce_choice_requires(&blank, &context, &understands, "test").unwrap(),
            ChoiceRequiresClassification::Blank
        );
        assert_eq!(
            classify_mce_choice_requires(&unbound, &context, &understands, "test").unwrap(),
            ChoiceRequiresClassification::Unresolved
        );

        let bad_escape =
            quick_xml::events::BytesStart::from_content(r#"mc:Choice Requires="&bogus;""#, 9);
        assert!(classify_mce_choice_requires(&bad_escape, &context, &understands, "test").is_err());
        let malformed_attribute =
            quick_xml::events::BytesStart::from_content(r#"mc:Choice Requires="k" broken"#, 9);
        assert!(
            classify_mce_choice_requires(&malformed_attribute, &context, &understands, "test")
                .is_err()
        );

        let limit = context
            .active_bytes()
            .saturating_add("urn:known".len())
            .saturating_sub(1);
        assert!(matches!(
            classify_bounded_mce_choice_requires(&known, &context, &understands, limit, "test"),
            Err(BoundedXmlError::Limit { .. })
        ));
    }

    #[test]
    fn mce_attributes_share_namespace_resolution_and_validation() {
        let xml = format!(
            r#"<r xmlns:mc="{MCE_NS}" xmlns:f="urn:future" mc:Ignorable="f" mc:ProcessContent="f:wrap" mc:MustUnderstand="f"><child/></r>"#,
        );
        let mut reader = NsReader::from_str(&xml);
        let mut buffer = Vec::new();
        let (_, event) = reader.read_resolved_event_into(&mut buffer).unwrap();
        let Event::Start(root) = event else {
            panic!("fixture root is a start element")
        };
        let namespace_root = NamespaceContext::root();
        let context = NamespaceContext::derive(&root, &namespace_root, 1024, "test").unwrap();
        let attributes =
            derive_mce_attributes(&reader, &root, &MceScope::root(), &context, 1024, "test")
                .unwrap();

        assert!(attributes.scope.is_ignorable("urn:future"));
        assert!(attributes.scope.processes_content("urn:future", "wrap"));
        assert_eq!(attributes.must_understand, ["urn:future"]);
        assert!(validate_mce_must_understand(
            &attributes.must_understand,
            &|namespace| namespace == "urn:future",
            "test"
        )
        .is_ok());
        assert!(
            validate_mce_must_understand(&attributes.must_understand, &|_| false, "test")
                .unwrap_err()
                .contains("MustUnderstand")
        );
    }

    #[test]
    fn processed_attributes_preserve_quoted_infoset() {
        let xml = format!(
            "<r xmlns:mc=\"{MCE_NS}\" xmlns:f=\"urn:future\" mc:Ignorable=\"f\" f:drop=\"removed\" name='A \"quote\" &quot; &amp; &#x9;\tline\nend'/>"
        );
        let mut reader = NsReader::from_str(&xml);
        let mut buffer = Vec::new();
        let (_, event) = reader.read_resolved_event_into(&mut buffer).unwrap();
        let Event::Empty(element) = event else {
            panic!("fixture root is an empty element")
        };
        let context =
            NamespaceContext::derive(&element, &NamespaceContext::root(), 1024, "test").unwrap();
        let attributes =
            derive_mce_attributes(&reader, &element, &MceScope::root(), &context, 1024, "test")
                .unwrap();
        let mut element = element.into_owned();
        strip_processed_mce_attributes(
            &reader,
            &mut element,
            &attributes.scope,
            &|_| false,
            "test",
        )
        .unwrap();

        let event = Event::Empty(element);
        let mut serialized = Vec::new();
        append_projected_event(&mut serialized, &event, 0, 1024, "test").unwrap();
        let projected = std::str::from_utf8(&serialized).unwrap();
        let doc = roxmltree::Document::parse(projected).unwrap();
        let root = doc.root_element();
        assert_eq!(root.attribute("name"), Some("A \"quote\" \" & \t line end"));
        assert_eq!(root.attribute((MCE_NS, "Ignorable")), None);
        assert_eq!(root.attribute(("urn:future", "drop")), None);
        assert_eq!(
            root.namespaces()
                .find(|ns| ns.name() == Some("f"))
                .unwrap()
                .uri(),
            "urn:future"
        );
    }

    #[test]
    fn mce_scope_persists_inherited_ignorable_and_process_content_state() {
        let root = MceScope::root();
        let parent = MceScope::derive(
            &root,
            HashSet::from(["urn:future".to_string()]),
            HashSet::new(),
            10,
        );
        let child = MceScope::derive(
            &parent,
            HashSet::new(),
            HashSet::from([("urn:future".to_string(), "keep".to_string())]),
            7,
        );

        assert!(child.is_ignorable("urn:future"));
        assert!(child.processes_content("urn:future", "keep"));
        assert!(!child.processes_content("urn:future", "drop"));
        assert_eq!(child.active_bytes(), 17);
    }

    #[test]
    fn fragment_namespace_repair_preflights_escaped_serialized_bytes() {
        let root = NamespaceContext::root();
        let declaration = BytesStart::from_content(r#"r xmlns:f="urn:a&amp;b""#, 1);
        let context = NamespaceContext::derive(&declaration, &root, 1024, "test").unwrap();
        let mut element = BytesStart::new("f:item");
        let injected_bytes = 10 + "f".len() + "urn:a&amp;b".len();
        let exact_limit = element.len() + injected_bytes;

        inject_missing_namespaces(&mut element, &context, None, exact_limit, "test").unwrap();
        assert_eq!(element.len(), exact_limit);
        assert!(String::from_utf8_lossy(element.as_ref()).contains(r#"xmlns:f="urn:a&amp;b""#));

        let mut over = BytesStart::new("f:item");
        let original_len = over.len();
        assert_eq!(
            inject_missing_namespaces(&mut over, &context, None, exact_limit - 1, "test"),
            Err(BoundedXmlError::Limit {
                limit: exact_limit - 1,
                observed: exact_limit,
            })
        );
        assert_eq!(over.len(), original_len);
    }
}
