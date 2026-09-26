//! Shared OOXML chart model and parsers used by DOCX, XLSX, and PPTX.
//!
//! The format parsers read a chart `<c:chartSpace>` (and the modern
//! `<cx:chartSpace>` for waterfall / treemap / box-and-whisker etc.) but
//! historically did so with two near-identical bodies sitting in
//! `packages/xlsx/parser/src/lib.rs` and `packages/pptx/parser/src/lib.rs`.
//! The result was that fields added on one side stayed missing on the other
//! until a host-specific regression exposed the drift (for example, one
//! adapter once discarded `legendPos` while another preserved it).
//!
//! This module owns the shared chart wire model and XML parsing. Package
//! adapters supply theme, image, style, and formula-resolution resources
//! through [`ChartParseContext`]; host-specific frame geometry stays local.
//!
//! ## Namespace handling
//!
//! All helpers match elements by local name only. Real chart documents put
//! everything under either the `c:` (chart 2006) or `cx:` (chartEx 2014)
//! namespace and never mix non-chart elements at these paths, so the strict
//! `tag_name().namespace() == Some(c_ns)` check in xlsx adds nothing in
//! practice — this module drops it for symmetry with the pptx side and to
//! keep the API simple. If a future format wedges a non-chart element into
//! `<c:plotArea>` the caller can pre-filter before delegating here.
//!
//! All field references are to ECMA-376 / ISO-29500 part 1 §21.2 (DrawingML
//! Charts) unless stated otherwise.

use roxmltree::Node;
use serde::{Deserialize, Serialize};
use std::collections::{BTreeMap, BTreeSet};

use crate::text::{parse_body_pr, BodyPrDefaults};
use crate::units::coordinate32_to_emu;

mod classic_style;

/// Resource ceiling for the expanded Chart Colors total set. Typical Office
/// parts contain 6 base colors × at most 9 variations; this bound prevents an
/// adversarial colors×variations product from amplifying a bounded XML tree.
const MAX_CHART_COLOR_STYLE_ENTRIES: usize = 4096;
/// A legend is a bounded UI list even when its backing series cache is much
/// larger. Bound authored entry overrides before allocating the wire model.
const MAX_CHART_LEGEND_ENTRIES: usize = 4096;
/// Aggregate structured-fill components retained after expanding one Chart
/// Style role across the linked Chart Colors palette. A gradient contributes
/// one component per stop; solid and pattern fills contribute one. This bounds
/// the otherwise multiplicative `palette entries × gradient stops` wire model.
const MAX_CHART_STYLE_PAINT_COMPONENTS: usize = 1_048_576;
/// A single structured paint recipe can be replayed for every visible chart
/// datum. Keep it bounded before `parse_grad_fill` allocates/sorts its stop
/// list; the aggregate budgets below then bound all retained recipes.
const MAX_CHART_PAINT_RECIPE_COMPONENTS: usize = 4096;
const MAX_CHART_MARKER_GRADIENT_STOPS: usize = MAX_CHART_PAINT_RECIPE_COMPONENTS;
/// Aggregate direct marker-paint components retained by one chart. This is an
/// availability boundary, not a visual compatibility rule.
const MAX_CHART_MARKER_PAINT_COMPONENTS: usize = MAX_CHART_STYLE_PAINT_COMPONENTS;
/// Direct data/trendline label shapes can replay one gradient for every label.
/// Bound both a single recipe and the aggregate retained wire components before
/// parsing stop lists. These are availability ceilings, not visual rules.
const MAX_CHART_LABEL_GRADIENT_STOPS: usize = MAX_CHART_PAINT_RECIPE_COMPONENTS;
const MAX_CHART_LABEL_PAINT_COMPONENTS: usize = MAX_CHART_STYLE_PAINT_COMPONENTS;
/// Aggregate role×palette slots retained from one linked Chart Style. Typical
/// Office parts use 30 paint roles × 6–54 colors. Reject the entire fallback
/// table before expansion rather than retaining an arbitrary role or palette
/// prefix when a hostile colors part exceeds this bounded wire budget.
const MAX_CHART_STYLE_ROLE_SLOTS: usize = 8_192;
/// Maximum cache width accepted from `<c:ptCount>` / `<cx:lvl ptCount>`.
/// Chart data originates in worksheet ranges, whose largest single dimension
/// is 1,048,576 rows. Rejecting wider sparse caches prevents an XML attribute
/// or point index from requesting an unbounded WASM allocation.
const MAX_CHART_CACHE_POINTS: usize = 1_048_576;
/// Classic plot groups are renderer work items even when they contain no
/// series. Keep public/wire planning aligned with the synchronous Canvas
/// availability ceiling before allocating the ordered group vector.
const MAX_CHART_PLOT_GROUPS: usize = 10_000;
const MAX_CHART_PLOT_SERIES: usize = 10_000;
const MAX_CHART_PLOT_AXES: usize = 4;
const MAX_CHART_GROUP_AXIS_IDS: usize = 3;

fn claim_plot_group_axis_slot(
    axis_id: Option<&String>,
    known_axis_ids: &BTreeMap<String, String>,
    primary: &mut Option<String>,
    secondary: &mut Option<String>,
) -> String {
    let slot = match axis_id {
        None => "unresolved",
        Some(axis_id) if !known_axis_ids.contains_key(axis_id) => "unresolved",
        Some(axis_id) if primary.as_ref().is_none_or(|current| current == axis_id) => {
            primary.get_or_insert_with(|| axis_id.clone());
            "primary"
        }
        Some(axis_id) if secondary.as_ref().is_none_or(|current| current == axis_id) => {
            secondary.get_or_insert_with(|| axis_id.clone());
            "secondary"
        }
        Some(_) => "unresolved",
    };
    slot.to_string()
}

mod axis;
mod cache;
mod chartex;
mod classic;
mod labels;
mod model;
mod style;

pub use axis::*;
pub use cache::*;
use chartex::*;
use classic::*;
pub use labels::*;
pub use model::*;
use style::*;

#[cfg(test)]
mod tests;

/// Package-owned sidecars and lookup hooks for one chart part. All fields are
/// optional so callers only supply resources present in the host package.
/// A color resolver is required to parse; omission returns `None`.
/// Formula lookup uses a cell because the host may memoize range resolution
/// while the parse entry point accepts a shared context reference.
#[derive(Default)]
pub struct ChartParseContext<'a> {
    pub color_resolver: Option<&'a dyn ColorResolver>,
    pub style_xml: Option<&'a str>,
    pub color_style_xml: Option<&'a str>,
    pub images: Option<&'a dyn ChartImageResolver>,
    pub references: std::cell::Cell<Option<&'a mut dyn ChartReferenceResolver>>,
}

impl<'a> ChartParseContext<'a> {
    /// Construct a context from the host's optional linked parts and resources.
    #[inline(never)]
    pub fn new(
        color_resolver: &'a dyn ColorResolver,
        style_xml: Option<&'a str>,
        color_style_xml: Option<&'a str>,
        images: Option<&'a dyn ChartImageResolver>,
        references: Option<&'a mut dyn ChartReferenceResolver>,
    ) -> Self {
        Self {
            color_resolver: Some(color_resolver),
            style_xml,
            color_style_xml,
            images,
            references: std::cell::Cell::new(references),
        }
    }
}

/// Parse an ECMA-376 DrawingML chart part into the shared wire model.
pub fn parse_chart_part(root: Node, context: &ChartParseContext<'_>) -> Option<ChartModel> {
    parse_part(root, context, parse_classic_impl)
}

/// Parse a Microsoft chartEx part into the shared wire model.
pub fn parse_chartex_part(root: Node, context: &ChartParseContext<'_>) -> Option<ChartModel> {
    parse_part(root, context, parse_chartex_impl)
}

type ChartPartParser = fn(
    Node<'_, '_>,
    &dyn ColorResolver,
    Option<&str>,
    Option<&str>,
    &mut dyn ChartReferenceResolver,
    &dyn ChartImageResolver,
) -> Option<ChartModel>;

fn parse_part(
    root: Node,
    context: &ChartParseContext<'_>,
    parser: ChartPartParser,
) -> Option<ChartModel> {
    let color_resolver = context.color_resolver?;
    let images = context.images.unwrap_or(&EmptyChartImageResolver);
    let mut references = context.references.take();
    let mut empty_references = EmptyChartReferenceResolver;
    let active_references = references.as_deref_mut().unwrap_or(&mut empty_references);
    let chart = parser(
        root,
        color_resolver,
        context.style_xml,
        context.color_style_xml,
        active_references,
        images,
    );
    context.references.set(references);
    chart
}
