//! Persistent, format-neutral resource policy and accounting for one OOXML
//! package session. ZIP readers report observations here; they do not interpret
//! public options or own poisoning/error serialization.

use serde::Serialize;
use std::cell::RefCell;
use std::collections::HashMap;
use std::rc::Rc;

// The generated hard central-directory ceiling is a conservative browser-safety
// budget for footer/index input plus retained filename bytes; it is deliberately
// not an exact accounting of allocator heap use.
include!("resource-policy.generated.rs");

#[derive(Clone, Copy, Debug, Serialize)]
#[serde(rename_all = "lowercase")]
pub enum OoxmlFormat {
    Docx,
    Xlsx,
    Pptx,
}

#[derive(Clone)]
pub struct ResourceGovernor(Rc<RefCell<GovernorState>>);

#[derive(Clone, Copy)]
struct ResourcePolicy {
    public_entry: Option<u64>,
    public_total: Option<u64>,
    public_entries: Option<u64>,
}

struct GovernorState {
    format: OoxmlFormat,
    policy: ResourcePolicy,
    operation: String,
    usage: ResourceUsage,
    // Central-directory identity, not a display path: duplicate ZIP names must
    // not accidentally share accounting state.
    max_actual_by_part: HashMap<usize, u64>,
    next_operation_id: u64,
    operations: HashMap<u64, LogicalOperationState>,
    first_error: Option<String>,
}

struct LogicalOperationState {
    name: String,
    inflated_bytes: u64,
}

/// Opaque identity for a logical operation whose accounting survives multiple
/// bounded protocol pulls.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct ResourceOperation(u64);

#[derive(Clone, Copy, Debug, Default, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ResourceUsage {
    pub archive_entry_count: u64,
    pub declared_inflated_bytes: u64,
    pub largest_inflated_entry_bytes: u64,
    pub distinct_inflated_bytes: u64,
    pub operation_inflated_bytes: u64,
}

/// Closed vocabulary for non-configurable parser/model safety ceilings.
/// Format parsers choose a semantic kind; this shared layer owns its stable
/// wire discriminants so stage/resource/metric strings cannot drift.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum HardResourceLimitKind {
    XmlEventBytes,
    XmlContextBytes,
    XmlNestingDepth,
    DocxBodyBlockXmlBytes,
    DocxBodyChunkJsonBytes,
    DocxBootstrapJsonBytes,
    DocxRetainedModelJsonBytes,
    DocxMarkdownBytes,
    PptxSlideXmlBytes,
    PptxSlideJsonBytes,
    PptxSharedDependencyXmlBytes,
    XmlDomComplexity,
    PptxSharedDependencyProjectionBytes,
    PptxSharedCacheEntries,
    PptxSharedCacheProjectionBytes,
    PptxBootstrapSlides,
    PptxBootstrapProjectionBytes,
    PptxBootstrapJsonBytes,
    PptxMarkdownBytes,
    PptxMaterializedSlideJsonBytes,
    WorksheetRowProjectionBytes,
    WorksheetShellProjectionBytes,
    WorksheetModelRows,
    WorksheetModelCells,
    WorksheetCellContentOwnedUtf8Bytes,
    WorksheetJsonBytes,
    WorksheetCacheRows,
    WorksheetCacheCells,
}

impl HardResourceLimitKind {
    const fn wire_fields(self) -> (&'static str, &'static str, &'static str) {
        match self {
            Self::XmlEventBytes => ("parsing", "xml-event", "bytes"),
            Self::XmlContextBytes => ("parsing", "xml-context", "bytes"),
            Self::XmlNestingDepth => ("parsing", "xml-tree", "depth"),
            Self::DocxBodyBlockXmlBytes => ("parsing", "docx-body-block-xml", "bytes"),
            Self::DocxBodyChunkJsonBytes => ("serialization", "docx-body-chunk-json", "bytes"),
            Self::DocxBootstrapJsonBytes => ("serialization", "docx-bootstrap-json", "bytes"),
            Self::DocxRetainedModelJsonBytes => {
                ("serialization", "docx-retained-model-json", "bytes")
            }
            Self::DocxMarkdownBytes => ("serialization", "docx-markdown", "bytes"),
            Self::PptxSlideXmlBytes => ("parsing", "pptx-slide-xml", "bytes"),
            Self::PptxSlideJsonBytes => ("serialization", "pptx-slide-json", "bytes"),
            Self::PptxSharedDependencyXmlBytes => {
                ("parsing", "pptx-shared-dependency-xml", "bytes")
            }
            Self::XmlDomComplexity => ("parsing", "xml-dom", "complexity-units"),
            Self::PptxSharedDependencyProjectionBytes => {
                ("parsing", "pptx-shared-dependency", "projected-bytes")
            }
            Self::PptxSharedCacheEntries => ("parsing", "pptx-shared-cache", "entries"),
            Self::PptxSharedCacheProjectionBytes => {
                ("parsing", "pptx-shared-cache", "projected-bytes")
            }
            Self::PptxBootstrapSlides => ("parsing", "pptx-bootstrap", "slides"),
            Self::PptxBootstrapProjectionBytes => ("parsing", "pptx-bootstrap", "projected-bytes"),
            Self::PptxBootstrapJsonBytes => ("serialization", "pptx-bootstrap-json", "bytes"),
            Self::PptxMarkdownBytes => ("serialization", "pptx-markdown", "bytes"),
            Self::PptxMaterializedSlideJsonBytes => (
                "serialization",
                "pptx-materialized-slides",
                "projected-json-bytes",
            ),
            Self::WorksheetRowProjectionBytes => ("parsing", "worksheet-row", "projected-bytes"),
            Self::WorksheetShellProjectionBytes => {
                ("parsing", "worksheet-shell", "projected-bytes")
            }
            Self::WorksheetModelRows => ("parsing", "worksheet-model", "rows"),
            Self::WorksheetModelCells => ("parsing", "worksheet-model", "cells"),
            Self::WorksheetCellContentOwnedUtf8Bytes => {
                ("parsing", "worksheet-cell-content", "owned-utf8-bytes")
            }
            Self::WorksheetJsonBytes => ("serialization", "worksheet-json", "bytes"),
            Self::WorksheetCacheRows => ("parsing", "worksheet-cache", "rows"),
            Self::WorksheetCacheCells => ("parsing", "worksheet-cache", "cells"),
        }
    }
}

impl ResourceUsage {
    fn for_wire(self) -> Self {
        const JS_MAX_SAFE_INTEGER: u64 = 9_007_199_254_740_991;
        Self {
            archive_entry_count: self.archive_entry_count.min(JS_MAX_SAFE_INTEGER),
            declared_inflated_bytes: self.declared_inflated_bytes.min(JS_MAX_SAFE_INTEGER),
            largest_inflated_entry_bytes: self
                .largest_inflated_entry_bytes
                .min(JS_MAX_SAFE_INTEGER),
            distinct_inflated_bytes: self.distinct_inflated_bytes.min(JS_MAX_SAFE_INTEGER),
            operation_inflated_bytes: self.operation_inflated_bytes.min(JS_MAX_SAFE_INTEGER),
        }
    }
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct ResourceLimitEnvelope<'a> {
    code: &'static str,
    details: ResourceLimitDetails<'a>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct ResourceLimitDetails<'a> {
    stage: &'static str,
    violation: ResourceViolation<'a>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct ResourceViolation<'a> {
    format: OoxmlFormat,
    operation: &'a str,
    resource: &'static str,
    metric: &'static str,
    #[serde(skip_serializing_if = "Option::is_none")]
    part: Option<&'a str>,
    limit: u64,
    observed: u64,
    configurable: bool,
    usage: ResourceUsage,
}

struct LimitCrossing<'a> {
    stage: &'static str,
    resource: &'static str,
    metric: &'static str,
    part: Option<&'a str>,
    limit: u64,
    observed: u64,
    configurable: bool,
}

thread_local! {
    static ACTIVE_GOVERNOR: RefCell<Option<ResourceGovernor>> = const { RefCell::new(None) };
}

/// Dynamic routing scope for legacy ZIP helper call sites. The durable state is
/// owned by [`ResourceGovernor`], not this thread-local slot.
#[must_use = "binding the scope keeps the package governor active"]
pub struct ResourceScope {
    governor: ResourceGovernor,
    previous_active: Option<ResourceGovernor>,
    previous_operation: String,
    previous_operation_bytes: u64,
    restore_parent_operation: bool,
    logical_operation_id: Option<u64>,
}

impl Drop for ResourceScope {
    fn drop(&mut self) {
        if self.logical_operation_id.is_some() || self.restore_parent_operation {
            let mut state = self.governor.0.borrow_mut();
            if let Some(id) = self.logical_operation_id {
                let inflated_bytes = state.usage.operation_inflated_bytes;
                if let Some(operation) = state.operations.get_mut(&id) {
                    operation.inflated_bytes = inflated_bytes;
                }
            }
            if self.restore_parent_operation {
                state.operation = std::mem::take(&mut self.previous_operation);
                state.usage.operation_inflated_bytes = self.previous_operation_bytes;
            }
        }
        ACTIVE_GOVERNOR.with(|active| {
            *active.borrow_mut() = self.previous_active.take();
        });
    }
}

impl ResourceScope {
    pub fn resource_limit_error(&self) -> Option<String> {
        self.governor.first_error()
    }
}

impl ResourceGovernor {
    /// Construct a browser policy from the normalized WASM scalar ABI.
    ///
    /// `None` means an older/native caller omitted the value and receives the
    /// standard default. `Some(0)` is the explicit public `null` sentinel and
    /// disables that configurable limit while retaining hard safety ceilings.
    pub fn from_wasm(
        format: OoxmlFormat,
        max_archive_entry_bytes: Option<u64>,
        max_total_inflated_bytes: Option<u64>,
        max_archive_entries: Option<u64>,
    ) -> Self {
        fn decode(value: Option<u64>, standard: u64) -> Option<u64> {
            match value {
                None => Some(standard),
                Some(0) => None,
                Some(value) => Some(value),
            }
        }
        Self(Rc::new(RefCell::new(GovernorState {
            format,
            policy: ResourcePolicy {
                public_entry: decode(max_archive_entry_bytes, STANDARD_MAX_ARCHIVE_ENTRY_BYTES),
                public_total: decode(max_total_inflated_bytes, STANDARD_MAX_TOTAL_INFLATED_BYTES),
                public_entries: decode(max_archive_entries, STANDARD_MAX_ARCHIVE_ENTRIES),
            },
            operation: "open".to_string(),
            usage: ResourceUsage::default(),
            max_actual_by_part: HashMap::new(),
            next_operation_id: 1,
            operations: HashMap::new(),
            first_error: None,
        })))
    }

    pub fn scope(&self, operation: impl Into<String>) -> ResourceScope {
        self.enter_scope(operation.into(), 0, None)
    }

    pub fn begin_operation(&self, name: impl Into<String>) -> ResourceOperation {
        let mut state = self.0.borrow_mut();
        let id = state.next_operation_id;
        state.next_operation_id = state.next_operation_id.saturating_add(1);
        state.operations.insert(
            id,
            LogicalOperationState {
                name: name.into(),
                inflated_bytes: 0,
            },
        );
        ResourceOperation(id)
    }

    pub fn scope_operation(&self, operation: ResourceOperation) -> Result<ResourceScope, String> {
        let state = self.0.borrow();
        let logical = state
            .operations
            .get(&operation.0)
            .ok_or_else(|| "resource operation does not belong to this governor".to_string())?;
        let name = logical.name.clone();
        let inflated_bytes = logical.inflated_bytes;
        drop(state);
        Ok(self.enter_scope(name, inflated_bytes, Some(operation.0)))
    }

    fn enter_scope(
        &self,
        operation: String,
        operation_inflated_bytes: u64,
        logical_operation_id: Option<u64>,
    ) -> ResourceScope {
        let mut state = self.0.borrow_mut();
        let previous_operation = std::mem::replace(&mut state.operation, operation);
        let previous_operation_bytes = state.usage.operation_inflated_bytes;
        state.usage.operation_inflated_bytes = operation_inflated_bytes;
        drop(state);
        let previous_active =
            ACTIVE_GOVERNOR.with(|active| active.borrow_mut().replace(self.clone()));
        let restore_parent_operation = previous_active
            .as_ref()
            .is_some_and(|active| Rc::ptr_eq(&active.0, &self.0));
        ResourceScope {
            governor: self.clone(),
            previous_active,
            previous_operation,
            previous_operation_bytes,
            restore_parent_operation,
            logical_operation_id,
        }
    }

    pub fn usage_for_operation(&self, operation: ResourceOperation) -> Option<ResourceUsage> {
        let state = self.0.borrow();
        let operation = state.operations.get(&operation.0)?;
        let mut usage = state.usage;
        usage.operation_inflated_bytes = operation.inflated_bytes;
        Some(usage)
    }

    /// Remove a completed logical operation from the live governor ledger.
    /// Callers that need a final checkpoint must capture it first.
    pub fn end_operation(&self, operation: ResourceOperation) {
        self.0.borrow_mut().operations.remove(&operation.0);
    }

    pub fn clear_operations(&self) {
        self.0.borrow_mut().operations.clear();
    }

    pub fn first_error(&self) -> Option<String> {
        self.0.borrow().first_error.clone()
    }

    pub fn usage(&self) -> ResourceUsage {
        self.0.borrow().usage
    }
}

pub(crate) fn active_governor() -> Option<ResourceGovernor> {
    ACTIVE_GOVERNOR.with(|active| active.borrow().clone())
}

fn effective_limit(public: Option<u64>, hard: u64) -> (u64, bool) {
    match public {
        Some(limit) if limit <= hard => (limit, true),
        _ => (hard, false),
    }
}

fn safe_part(path: &str) -> &str {
    let bytes = path.as_bytes();
    let windows_drive = bytes.len() >= 2 && bytes[0].is_ascii_alphabetic() && bytes[1] == b':';
    let safe = !path.is_empty()
        && path.len() <= 1024
        && !path.starts_with('/')
        && !path.starts_with('\\')
        && !path.contains('\\')
        && !path.contains('?')
        && !path.contains('#')
        && !path.contains("://")
        && !windows_drive
        && !path
            .split('/')
            .any(|segment| segment.is_empty() || segment == "." || segment == "..")
        && !path.chars().any(char::is_control);
    if safe {
        path
    } else {
        // Keep the required part discriminator without reflecting an unsafe
        // attacker-controlled address. This token satisfies the shared wire
        // contract's relative OPC-address grammar.
        "untrusted-archive-entry"
    }
}

impl GovernorState {
    fn fail(&mut self, crossing: LimitCrossing<'_>) -> String {
        if let Some(error) = &self.first_error {
            return error.clone();
        }
        let part = crossing.part.map(safe_part);
        let operation = self.operation.clone();
        let envelope = ResourceLimitEnvelope {
            code: "ooxml-resource-limit",
            details: ResourceLimitDetails {
                stage: crossing.stage,
                violation: ResourceViolation {
                    format: self.format,
                    operation: &operation,
                    resource: crossing.resource,
                    metric: crossing.metric,
                    part,
                    limit: crossing.limit,
                    // Keep every wire number exactly representable in JS and
                    // freeze known crossings at the first byte beyond the cap.
                    observed: crossing.observed.min(crossing.limit.saturating_add(1)),
                    configurable: crossing.configurable,
                    usage: self.usage.for_wire(),
                },
            },
        };
        let error = format!(
            "OOXML_RESOURCE_LIMIT:{}",
            serde_json::to_string(&envelope).expect("resource-limit details serialize")
        );
        self.first_error = Some(error.clone());
        error
    }

    fn assert_healthy(&self) -> Result<(), String> {
        match &self.first_error {
            Some(error) => Err(error.clone()),
            None => Ok(()),
        }
    }
}

pub(crate) fn assert_healthy() -> Result<(), String> {
    let Some(governor) = active_governor() else {
        return Ok(());
    };
    let result = governor.0.borrow().assert_healthy();
    result
}

pub(crate) fn observe_archive_metadata(
    entry_count: u64,
    declared_inflated_bytes: u64,
) -> Result<(), String> {
    let Some(governor) = active_governor() else {
        return Ok(());
    };
    let mut state = governor.0.borrow_mut();
    state.assert_healthy()?;
    state.usage.archive_entry_count = state.usage.archive_entry_count.max(entry_count);
    state.usage.declared_inflated_bytes = declared_inflated_bytes;
    if let Some(limit) = state.policy.public_entries {
        if entry_count > limit {
            return Err(state.fail(LimitCrossing {
                stage: "container",
                resource: "archive",
                metric: "entry-count",
                part: None,
                limit,
                observed: entry_count,
                configurable: true,
            }));
        }
    }
    if entry_count > HARD_MAX_ARCHIVE_ENTRIES {
        return Err(state.fail(LimitCrossing {
            stage: "container",
            resource: "archive",
            metric: "entry-count",
            part: None,
            limit: HARD_MAX_ARCHIVE_ENTRIES,
            observed: entry_count,
            configurable: false,
        }));
    }
    Ok(())
}

/// Enforce the non-configurable aggregate ZIP-index metadata ceiling.
pub(crate) fn observe_archive_central_directory_bytes(observed: u64) -> Result<(), String> {
    let Some(governor) = active_governor() else {
        return Ok(());
    };
    let mut state = governor.0.borrow_mut();
    state.assert_healthy()?;
    if observed > HARD_MAX_CENTRAL_DIRECTORY_BYTES {
        return Err(state.fail(LimitCrossing {
            stage: "container",
            resource: "archive",
            metric: "central-directory-bytes",
            part: None,
            limit: HARD_MAX_CENTRAL_DIRECTORY_BYTES,
            observed,
            configurable: false,
        }));
    }
    Ok(())
}

/// Latch a proven non-configurable parser/model ceiling crossing. Without an
/// active package scope this remains a no-op so pure native parser helpers can
/// still return their local structured error.
pub fn observe_hard_limit(
    kind: HardResourceLimitKind,
    part: Option<&str>,
    limit: u64,
    observed: u64,
) -> Result<(), String> {
    if observed <= limit {
        return Ok(());
    }
    let Some(governor) = active_governor() else {
        return Ok(());
    };
    let mut state = governor.0.borrow_mut();
    state.assert_healthy()?;
    let (stage, resource, metric) = kind.wire_fields();
    Err(state.fail(LimitCrossing {
        stage,
        resource,
        metric,
        part,
        limit,
        observed,
        configurable: false,
    }))
}

pub(crate) fn read_allowance(
    part_id: usize,
    path: &str,
    declared_size: u64,
) -> Result<u64, String> {
    let Some(governor) = active_governor() else {
        return Ok(HARD_MAX_ARCHIVE_ENTRY_BYTES);
    };
    let mut state = governor.0.borrow_mut();
    state.assert_healthy()?;
    let (entry_limit, entry_configurable) =
        effective_limit(state.policy.public_entry, HARD_MAX_ARCHIVE_ENTRY_BYTES);
    if declared_size > entry_limit {
        return Err(state.fail(LimitCrossing {
            stage: "container",
            resource: "archive-entry",
            metric: "declared-inflated-bytes",
            part: Some(path),
            limit: entry_limit,
            observed: declared_size,
            configurable: entry_configurable,
        }));
    }
    let old = state.max_actual_by_part.get(&part_id).copied().unwrap_or(0);
    let (total_limit, _) =
        effective_limit(state.policy.public_total, HARD_MAX_TOTAL_INFLATED_BYTES);
    let total_allowance =
        old.saturating_add(total_limit.saturating_sub(state.usage.distinct_inflated_bytes));
    Ok(entry_limit.min(total_allowance))
}

pub(crate) fn observe_inflated(
    part_id: usize,
    path: &str,
    cumulative_for_read: u64,
    delivered_increment: u64,
) -> Result<(), String> {
    let Some(governor) = active_governor() else {
        return Ok(());
    };
    let mut state = governor.0.borrow_mut();
    state.assert_healthy()?;
    state.usage.operation_inflated_bytes = state
        .usage
        .operation_inflated_bytes
        .saturating_add(delivered_increment);

    let old = state.max_actual_by_part.get(&part_id).copied().unwrap_or(0);
    let next = old.max(cumulative_for_read);
    state.usage.largest_inflated_entry_bytes = state
        .usage
        .largest_inflated_entry_bytes
        .max(cumulative_for_read);
    let next_total = state
        .usage
        .distinct_inflated_bytes
        .saturating_add(next.saturating_sub(old));
    state.max_actual_by_part.insert(part_id, next);
    state.usage.distinct_inflated_bytes = next_total;

    let (entry_limit, entry_configurable) =
        effective_limit(state.policy.public_entry, HARD_MAX_ARCHIVE_ENTRY_BYTES);
    if cumulative_for_read > entry_limit {
        return Err(state.fail(LimitCrossing {
            stage: "decompression",
            resource: "archive-entry",
            metric: "actual-inflated-bytes",
            part: Some(path),
            limit: entry_limit,
            observed: cumulative_for_read,
            configurable: entry_configurable,
        }));
    }
    let (total_limit, total_configurable) =
        effective_limit(state.policy.public_total, HARD_MAX_TOTAL_INFLATED_BYTES);
    if next_total > total_limit {
        return Err(state.fail(LimitCrossing {
            stage: "decompression",
            resource: "archive",
            metric: "distinct-inflated-bytes",
            part: Some(path),
            limit: total_limit,
            observed: next_total,
            configurable: total_configurable,
        }));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn zero_wire_value_disables_only_the_public_limit() {
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Docx, Some(0), Some(0), None);
        let state = governor.0.borrow();
        assert_eq!(state.policy.public_entry, None);
        assert_eq!(state.policy.public_total, None);
        assert_eq!(
            state.policy.public_entries,
            Some(STANDARD_MAX_ARCHIVE_ENTRIES)
        );
        assert_eq!(effective_limit(state.policy.public_entry, 10), (10, false));
    }

    #[test]
    fn entry_count_public_limit_is_typed_configurable_and_inclusive() {
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Docx, Some(0), Some(0), Some(2));
        let _scope = governor.scope("open");
        observe_archive_metadata(2, 0).unwrap();
        let error = observe_archive_metadata(3, 0).unwrap_err();
        let envelope: serde_json::Value =
            serde_json::from_str(error.strip_prefix("OOXML_RESOURCE_LIMIT:").unwrap()).unwrap();
        let violation = &envelope["details"]["violation"];
        assert_eq!(violation["metric"], "entry-count");
        assert_eq!(violation["limit"], 2);
        assert_eq!(violation["observed"], 3);
        assert_eq!(violation["configurable"], true);
        assert!(violation.get("part").is_none());
    }

    #[test]
    fn entry_count_reports_the_public_policy_before_the_hard_ceiling() {
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Docx, Some(0), Some(0), Some(2));
        let _scope = governor.scope("open");
        let error = observe_archive_metadata(HARD_MAX_ARCHIVE_ENTRIES + 1, 0).unwrap_err();
        let envelope: serde_json::Value =
            serde_json::from_str(error.strip_prefix("OOXML_RESOURCE_LIMIT:").unwrap()).unwrap();
        let violation = &envelope["details"]["violation"];
        assert_eq!(violation["limit"], 2);
        assert_eq!(violation["configurable"], true);
    }

    #[test]
    fn zero_wire_entry_count_disables_only_public_limit() {
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Xlsx, Some(0), Some(0), Some(0));
        let _scope = governor.scope("open");
        observe_archive_metadata(STANDARD_MAX_ARCHIVE_ENTRIES + 1, 0).unwrap();
        assert_eq!(governor.0.borrow().policy.public_entries, None);
    }

    #[test]
    fn scope_restores_outer_operation_but_retains_session_usage() {
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Xlsx, Some(100), Some(100), None);
        {
            let _outer = governor.scope("parse");
            observe_inflated(0, "xl/a.xml", 4, 4).unwrap();
            {
                let _inner = governor.scope("parse-sheet");
                observe_inflated(1, "xl/b.xml", 4, 4).unwrap();
            }
            assert_eq!(governor.0.borrow().operation, "parse");
        }
        assert_eq!(governor.usage().distinct_inflated_bytes, 8);
        assert_eq!(governor.usage().largest_inflated_entry_bytes, 4);
        assert_eq!(governor.usage().operation_inflated_bytes, 4);
    }

    #[test]
    fn completed_top_level_scope_retains_its_operation_usage() {
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Docx, Some(100), Some(100), None);
        {
            let _scope = governor.scope("parse");
            observe_inflated(0, "word/document.xml", 7, 7).unwrap();
        }
        assert_eq!(governor.0.borrow().operation, "parse");
        assert_eq!(governor.usage().operation_inflated_bytes, 7);

        {
            let _scope = governor.scope("markdown");
            observe_inflated(0, "word/document.xml", 7, 7).unwrap();
        }
        assert_eq!(governor.0.borrow().operation, "markdown");
        assert_eq!(governor.usage().operation_inflated_bytes, 7);
    }

    #[test]
    fn first_violation_poison_is_stable() {
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Pptx, Some(4), Some(20), None);
        let _scope = governor.scope("extract-image");
        let first = observe_inflated(0, "ppt/media/a.png", 5, 5).unwrap_err();
        let later = read_allowance(1, "ppt/media/b.png", 1).unwrap_err();
        assert_eq!(later, first);
        assert!(first.starts_with("OOXML_RESOURCE_LIMIT:"));
    }

    #[test]
    fn unsafe_part_names_use_the_wire_safe_redaction_token() {
        for unsafe_path in [
            "",
            "/word/document.xml",
            "../word/document.xml",
            "word/../document.xml",
            "word/./document.xml",
            "word//document.xml",
            "word\\document.xml",
            "C:/word/document.xml",
            "https://example.invalid/document.xml",
            "word/document.xml?query",
            "word/document.xml#fragment",
            "word/document.xml\nsecret",
        ] {
            assert_eq!(safe_part(unsafe_path), "untrusted-archive-entry");
        }
        let oversized = "a".repeat(1025);
        assert_eq!(safe_part(&oversized), "untrusted-archive-entry");
        assert_eq!(safe_part("word/document.xml"), "word/document.xml");

        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Docx, Some(4), Some(20), None);
        let _scope = governor.scope("parse");
        let first = observe_inflated(0, "../../private/document.xml", 5, 5).unwrap_err();
        let replay = read_allowance(1, "word/document.xml", 1).unwrap_err();
        assert_eq!(replay, first);
        let envelope: serde_json::Value =
            serde_json::from_str(first.strip_prefix("OOXML_RESOURCE_LIMIT:").unwrap()).unwrap();
        assert_eq!(
            envelope["details"]["violation"]["part"],
            "untrusted-archive-entry"
        );
    }

    #[test]
    fn hard_parser_limit_uses_closed_wire_discriminants_and_poisons() {
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Xlsx, None, None, None);
        let _scope = governor.scope("parse-sheet");
        observe_hard_limit(
            HardResourceLimitKind::WorksheetRowProjectionBytes,
            Some("xl/worksheets/sheet1.xml"),
            8,
            8,
        )
        .unwrap();
        let error = observe_hard_limit(
            HardResourceLimitKind::WorksheetRowProjectionBytes,
            Some("xl/worksheets/sheet1.xml"),
            8,
            9,
        )
        .unwrap_err();
        let json: serde_json::Value = serde_json::from_str(
            error
                .strip_prefix("OOXML_RESOURCE_LIMIT:")
                .expect("typed envelope prefix"),
        )
        .unwrap();
        let violation = &json["details"]["violation"];
        assert_eq!(json["details"]["stage"], "parsing");
        assert_eq!(violation["resource"], "worksheet-row");
        assert_eq!(violation["metric"], "projected-bytes");
        assert_eq!(violation["part"], "xl/worksheets/sheet1.xml");
        assert_eq!(violation["configurable"], false);
        assert_eq!(
            observe_hard_limit(HardResourceLimitKind::XmlEventBytes, None, 1, 2).unwrap_err(),
            error
        );
    }

    #[test]
    fn worksheet_json_limit_is_inclusive_and_crosses_at_plus_one() {
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Xlsx, Some(0), Some(0), None);
        let _scope = governor.scope("parse-sheet");
        observe_hard_limit(
            HardResourceLimitKind::WorksheetJsonBytes,
            Some("xl/worksheets/sheet1.xml"),
            64,
            64,
        )
        .unwrap();
        let error = observe_hard_limit(
            HardResourceLimitKind::WorksheetJsonBytes,
            Some("xl/worksheets/sheet1.xml"),
            64,
            65,
        )
        .unwrap_err();
        let envelope: serde_json::Value =
            serde_json::from_str(error.strip_prefix("OOXML_RESOURCE_LIMIT:").unwrap()).unwrap();
        assert_eq!(envelope["details"]["stage"], "serialization");
        assert_eq!(
            envelope["details"]["violation"]["resource"],
            "worksheet-json"
        );
        assert_eq!(envelope["details"]["violation"]["observed"], 65);
        assert_eq!(envelope["details"]["violation"]["configurable"], false);
    }
}
