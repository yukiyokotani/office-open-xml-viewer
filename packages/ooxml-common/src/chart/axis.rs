use super::*;

/// `<c:catAx|valAx><c:numFmt formatCode>` — the value-axis tick label
/// number format (ECMA-376 §21.2.2.21). Caller passes the already-located
/// `<c:catAx>` / `<c:valAx>` node.
#[cfg(test)]
pub(super) fn extract_axis_format_code(axis_node: Node) -> Option<String> {
    child(axis_node, "numFmt")
        .and_then(|n| n.attribute("formatCode"))
        .map(|s| s.to_string())
        .filter(|s| !s.is_empty() && s != "General")
}

/// Authored `<c:numFmt>` (§21.2.2.121), kept apart from the effective tick
/// format projected into the renderer's `*formatCode` fields. `None` for
/// `sourceLinked` preserves omission, whose effective meaning is true.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartAxisNumberFormat {
    pub authored_code: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub source_linked: Option<bool>,
}

pub(super) fn axis_number_format(axis: Node) -> Option<ChartAxisNumberFormat> {
    let number_format = child(axis, "numFmt")?;
    let authored_code = number_format.attribute("formatCode")?.to_string();
    let source_linked = number_format
        .attribute("sourceLinked")
        .and_then(|value| match value {
            "1" | "true" => Some(true),
            "0" | "false" => Some(false),
            _ => None,
        });
    Some(ChartAxisNumberFormat {
        authored_code,
        source_linked,
    })
}

pub(super) enum AxisNumberFormatSource {
    Formula(String),
    Literal(String),
    Unavailable,
}

pub(super) fn effective_axis_format_code(
    format: &ChartAxisNumberFormat,
    source: Option<&AxisNumberFormatSource>,
    references: &mut dyn ChartReferenceResolver,
) -> Option<String> {
    let code = if format.source_linked.unwrap_or(true) {
        match source {
            Some(AxisNumberFormatSource::Formula(formula)) => references
                .resolve_number_format(formula)
                .unwrap_or_else(|| format.authored_code.clone()),
            Some(AxisNumberFormatSource::Literal(code)) => code.clone(),
            _ => format.authored_code.clone(),
        }
    } else {
        format.authored_code.clone()
    };
    (!code.is_empty() && code != "General").then_some(code)
}

/// `<c:catAx|valAx><c:scaling>` — read explicit `<c:min val>` / `<c:max val>`.
/// Returns `(min, max)`; either can be `None` when the file leaves Excel to
/// pick the auto bound.
pub(super) fn extract_axis_min_max(axis_node: Node) -> (Option<f64>, Option<f64>) {
    let Some(scaling) = child(axis_node, "scaling") else {
        return (None, None);
    };
    let mn = child(scaling, "min")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<f64>().ok());
    let mx = child(scaling, "max")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<f64>().ok());
    (mn, mx)
}

/// `<c:catAx|valAx><c:crosses val>` and `<c:crossesAt val>` (ECMA-376
/// §21.2.2.33/§21.2.2.34). `crosses` is `autoZero` | `min` | `max`; `crossesAt`
/// is an explicit numeric override. Returns `(crosses, crosses_at)`.
pub(super) fn extract_axis_crosses(axis_node: Node) -> (Option<String>, Option<f64>) {
    let crosses = child(axis_node, "crosses")
        .and_then(|n| n.attribute("val"))
        .map(|s| s.to_string());
    let crosses_at = child(axis_node, "crossesAt")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<f64>().ok());
    (crosses, crosses_at)
}

/// `<c:radarChart><c:radarStyle val>` (ECMA-376 §21.2.3.10): `standard` (line
/// only), `marker` (line + markers), or `filled` (closed area). `None` when
/// the chart is not a radar chart or omits the element.
pub(super) fn extract_radar_style(root: Node) -> Option<String> {
    root.descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "radarStyle")
        .and_then(|n| n.attribute("val"))
        .map(|s| s.to_string())
}

/// Parse a `<c:layout><c:manualLayout>` node into a [`ChartManualLayout`]
/// (ECMA-376 §21.2.2.88). `layout_node` is the `<c:layout>` element; returns
/// `None` when it carries no `<c:manualLayout>` child. `layoutTarget` defaults
/// to `"outer"` per CT_LayoutTarget; `x`/`y` default to 0; `w`/`h` stay `None`
/// when absent.
pub(super) fn extract_manual_layout(layout_node: Node) -> Option<ChartManualLayout> {
    let manual = child(layout_node, "manualLayout")?;
    // CT_LayoutMode@val defaults to factor in both Strict and Transitional
    // dml-chart.xsd. The element may also be present with no val attribute.
    let mut x_mode = "factor".to_string();
    let mut y_mode = "factor".to_string();
    let mut w_mode = "factor".to_string();
    let mut h_mode = "factor".to_string();
    let mut layout_target = Some("outer".to_string());
    let mut x = 0.0_f64;
    let mut y = 0.0_f64;
    let mut w: Option<f64> = None;
    let mut h: Option<f64> = None;
    for ch in manual.children().filter(|n| n.is_element()) {
        let val_str = attr(&ch, "val");
        match ch.tag_name().name() {
            "xMode" => {
                if let Some(v) = val_str {
                    x_mode = v;
                }
            }
            "yMode" => {
                if let Some(v) = val_str {
                    y_mode = v;
                }
            }
            "wMode" => {
                if let Some(v) = val_str {
                    w_mode = v;
                }
            }
            "hMode" => {
                if let Some(v) = val_str {
                    h_mode = v;
                }
            }
            "layoutTarget" => {
                layout_target = val_str;
            }
            "x" => {
                if let Some(v) = val_str.and_then(|s| s.parse::<f64>().ok()) {
                    x = v;
                }
            }
            "y" => {
                if let Some(v) = val_str.and_then(|s| s.parse::<f64>().ok()) {
                    y = v;
                }
            }
            "w" => {
                w = val_str.and_then(|s| s.parse::<f64>().ok());
            }
            "h" => {
                h = val_str.and_then(|s| s.parse::<f64>().ok());
            }
            _ => {}
        }
    }
    Some(ChartManualLayout {
        x_mode,
        y_mode,
        w_mode,
        h_mode,
        layout_target,
        x,
        y,
        w,
        h,
    })
}

/// `<c:legend><c:layout><c:manualLayout>` (ECMA-376 §21.2.2.31) → a
/// [`LegendManualLayout`]. Unlike the plot/title layout, the legend variant has
/// no `layoutTarget`; optional `w`/`h` retain their authored presence.
/// `legend_node` is the `<c:legend>` element. `None` when it has no manual layout.
pub(super) fn extract_legend_manual_layout(legend_node: Node) -> Option<LegendManualLayout> {
    let layout = child(legend_node, "layout")?;
    let manual = extract_manual_layout(layout)?;
    Some(LegendManualLayout {
        x_mode: manual.x_mode,
        y_mode: manual.y_mode,
        w_mode: manual.w_mode,
        h_mode: manual.h_mode,
        x: manual.x,
        y: manual.y,
        w: manual.w,
        h: manual.h,
    })
}

/// `<c:catAx|valAx><c:delete val="1"/>` — true when the axis (labels, ticks
/// and line) should be hidden. ECMA-376 §21.2.2.40. `<c:delete>` is a
/// `CT_Boolean` (dml-chart.xsd `val` default `true`), so a bare `<c:delete/>`
/// means the axis IS deleted; an absent element leaves the axis shown.
pub(super) fn axis_is_deleted(axis_node: Node) -> bool {
    bool_child(axis_node, "delete").unwrap_or(false)
}

/// `<c:catAx|valAx><c:majorTickMark val>` / `<c:minorTickMark val>`. Values
/// are the ECMA-376 §21.2.3.48 ST_TickMark enum: `none` | `out` | `in` |
/// `cross`. `CT_TickMark@val` itself defaults to `cross`, so a present bare
/// element is distinct from an omitted element and resolves to `cross`.
pub(super) fn extract_axis_tick_mark(axis_node: Node, name: &str) -> Option<String> {
    child(axis_node, name).map(|node| node.attribute("val").unwrap_or("cross").to_string())
}

/// Like [`extract_axis_tick_mark`] but applies Office's application default
/// `"out"` when the entire major-tick element is absent. This is intentionally
/// separate from the schema default `cross` for a present element whose `val`
/// attribute is omitted.
pub(super) fn extract_axis_tick_mark_or_default(axis_node: Node, name: &str) -> String {
    extract_axis_tick_mark(axis_node, name).unwrap_or_else(|| "out".to_string())
}

/// First `<a:defRPr@sz>` or `<a:rPr@sz>` found inside the axis's `<c:txPr>`.
/// Sizes are OOXML hundredths of a point (e.g. 1200 = 12 pt).
pub(super) fn extract_axis_tick_label_size(axis_node: Node) -> Option<i32> {
    let txpr = child(axis_node, "txPr")?;
    txpr.descendants().find_map(|n| {
        if !n.is_element() {
            return None;
        }
        let tag = n.tag_name().name();
        if tag != "defRPr" && tag != "rPr" {
            return None;
        }
        n.attribute("sz").and_then(parse_text_font_size_hpt)
    })
}

/// First `<a:defRPr>` / `<a:rPr>` bold state inside the axis's `<c:txPr>`
/// (ECMA-376 §21.2.2.17). Visual direct formatting with omitted `b` is false;
/// empty or metadata-only properties remain eligible for style inheritance.
pub(super) fn extract_axis_tick_label_bold(axis_node: Node) -> Option<bool> {
    let txpr = child(axis_node, "txPr")?;
    chart_text_bool_from_present_props(first_chart_text_character_props(txpr), "b")
}

/// First `<a:defRPr@i>` / `<a:rPr@i>` italic flag inside the axis's
/// `<c:txPr>`. DrawingML character properties keep italic independent from
/// bold, so preserve it separately through the chart model.
pub(super) fn extract_axis_tick_label_italic(axis_node: Node) -> Option<bool> {
    let txpr = child(axis_node, "txPr")?;
    chart_text_bool_from_present_props(first_chart_text_character_props(txpr), "i")
}

/// Plain text of `node`'s direct-child `<c:title>` (ECMA-376 §21.2.2.210
/// `CT_Title`). Works for the `<c:chart>` element (chart title) or a
/// `<c:catAx>` / `<c:valAx>` (axis title). Walks `<a:t>` (rich text runs) and
/// `<c:v>` (string-ref cache) descendants and concatenates their text.
/// Returns `None` when there is no `<c:title>` child or it carries no text.
pub(super) fn extract_chart_title_text(node: Node) -> Option<String> {
    let title = child(node, "title")?;
    let mut text = String::new();
    for d in title.descendants().filter(|n| n.is_element()) {
        match d.tag_name().name() {
            "t" | "v" => {
                if let Some(t) = d.text() {
                    text.push_str(t);
                }
            }
            _ => {}
        }
    }
    if text.is_empty() {
        None
    } else {
        Some(text)
    }
}

pub(super) fn parse_text_font_size_hpt(value: &str) -> Option<i32> {
    value
        .parse::<i32>()
        .ok()
        .filter(|size| (100..=400_000).contains(size))
}

/// Title text-property nodes in DrawingML cascade order. The rich text body is
/// direct authoring and precedes title-level `txPr`; within each scope an
/// explicit run property precedes its default run property.
pub(super) fn title_text_property_nodes<'a, 'input>(
    title: Node<'a, 'input>,
) -> Vec<Node<'a, 'input>> {
    let rich = child(title, "tx").and_then(|tx| child(tx, "rich"));
    let tx_pr = child(title, "txPr");
    let mut nodes = Vec::new();
    for (scope, wanted) in [
        (rich, "rPr"),
        (rich, "defRPr"),
        (tx_pr, "rPr"),
        (tx_pr, "defRPr"),
    ] {
        if let Some(scope) = scope {
            nodes.extend(
                scope
                    .descendants()
                    .filter(|node| node.is_element() && node.tag_name().name() == wanted),
            );
        }
    }
    // Preserve compatibility with simplified producers/tests that place the
    // DrawingML paragraph directly under `<c:title>` without the CT_Tx wrapper.
    if nodes.is_empty() {
        for wanted in ["rPr", "defRPr"] {
            nodes.extend(
                title
                    .descendants()
                    .filter(|node| node.is_element() && node.tag_name().name() == wanted),
            );
        }
    }
    nodes
}

/// First `<a:defRPr@sz>` / `<a:rPr@sz>` (hundredths of a point) inside `node`'s
/// direct-child `<c:title>`. `None` when absent.
#[cfg(test)]
pub(super) fn extract_chart_title_size(node: Node) -> Option<i32> {
    let title = child(node, "title")?;
    title.descendants().find_map(|n| {
        if !n.is_element() {
            return None;
        }
        let tag = n.tag_name().name();
        if tag != "defRPr" && tag != "rPr" {
            return None;
        }
        n.attribute("sz").and_then(parse_text_font_size_hpt)
    })
}

/// chartEx (`<cx:chartSpace>`) title font size in hundredths of a point.
///
/// Unlike the legacy chart, whose `<c:title>` is a direct child of the chart
/// node, a chartEx title lives at `<cx:chart><cx:title>` (a grandchild of the
/// part root), so this walks all descendants to find the first `<cx:title>` and
/// reads its first `<a:defRPr@sz>` / `<a:rPr@sz>`. `None` when the title carries
/// no explicit size — which is the common case (see
/// [`extract_chartex_style_title_size`]).
pub(super) fn extract_chartex_title_size(root: Node) -> Option<i32> {
    let title = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "title")?;
    title.descendants().find_map(|n| {
        if !n.is_element() {
            return None;
        }
        let tag = n.tag_name().name();
        if tag != "defRPr" && tag != "rPr" {
            return None;
        }
        n.attribute("sz").and_then(parse_text_font_size_hpt)
    })
}

/// Whether a chart-part relationship targets its chartStyle sidecar. Office
/// packages in the wild use the 2011 URI, while MS-ODRAWXML §2.1.2 specifies
/// the 2012 URI; all three host parsers must accept both exact revisions.
pub fn is_chart_style_relationship_type(value: &str) -> bool {
    matches!(
        value,
        "http://schemas.microsoft.com/office/2011/relationships/chartStyle"
            | "http://schemas.microsoft.com/office/2012/relationships/chartStyle"
    )
}
/// Accepts both Office's 2011 and 2012 relationship namespace revisions.
pub const CHART_COLOR_STYLE_REL_TYPE_SUFFIX: &str = "relationships/chartColorStyle";

/// Title font size (hundredths of a point) declared by the chart's associated
/// chartStyle part (`<cs:chartStyle><cs:title><cs:defRPr@sz>`).
///
/// A chartEx part almost never inlines the title size on its own `<cx:title>`;
/// instead the size lives in the sibling `styleN.xml` reached via the chart
/// part's Office 2011 / MS-ODRAWXML 2012 `chartStyle` relationship. Word's default
/// modern chart style writes `<cs:title><cs:defRPr sz="1400">` (14pt). `None`
/// when `style_xml` is absent, malformed, or declares no `<cs:title>` size; the
/// renderer then uses its shared deterministic fallback.
#[cfg(test)]
pub(super) fn extract_chartex_style_title_size(style_xml: &str) -> Option<i32> {
    let doc = crate::depth::parse_guarded(style_xml).ok()?;
    let title = doc
        .root_element()
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "title")?;
    // `<cs:title>`'s size sits on its direct-child `<cs:defRPr@sz>`; scan
    // descendants so a nested `<a:defRPr>`/`<a:rPr>` (if any) is also honored.
    title.descendants().find_map(|n| {
        if !n.is_element() {
            return None;
        }
        let tag = n.tag_name().name();
        if tag != "defRPr" && tag != "rPr" {
            return None;
        }
        n.attribute("sz").and_then(parse_text_font_size_hpt)
    })
}

/// Effective bold flag of the title's direct character-property cascade.
/// Visual direct formatting with omitted `b` is regular. Empty or language-only
/// carriers remain `None`, matching Office-authored titles that inherit bold
/// from the numeric or linked chart style.
pub(super) fn extract_chart_title_bold(node: Node) -> Option<bool> {
    let title = child(node, "title")?;
    chart_text_bool_from_property_cascade(
        title_text_property_nodes(title).into_iter().map(Some),
        "b",
    )
}

/// Effective italic flag of the title's direct character-property cascade.
/// See [`extract_chart_title_bold`] for the visual-format ownership boundary.
pub(super) fn extract_chart_title_italic(node: Node) -> Option<bool> {
    let title = child(node, "title")?;
    chart_text_bool_from_property_cascade(
        title_text_property_nodes(title).into_iter().map(Some),
        "i",
    )
}

/// First `<a:solidFill>/<a:srgbClr@val>` (hex without `#`) inside `node`'s
/// direct-child `<c:title>`. Only an `<a:srgbClr>` that is a direct child of a
/// `<a:solidFill>` is honored — this skips gradient stops and other non-fill
/// color nodes. `<a:schemeClr>` is left unresolved here (the theme palette is
/// not wired through to chart title/border parsing yet — a known limitation
/// shared by both parsers). `None` = renderer default.
#[cfg(test)]
pub(super) fn extract_chart_title_srgb(node: Node) -> Option<String> {
    let title = child(node, "title")?;
    title.descendants().find_map(|n| {
        if !n.is_element() || n.tag_name().name() != "srgbClr" {
            return None;
        }
        // Skip srgbClr nodes that aren't inside a solidFill (e.g. a gradient stop).
        let parent_is_solid = n
            .parent()
            .map(|p| p.tag_name().name() == "solidFill")
            .unwrap_or(false);
        if !parent_is_solid {
            return None;
        }
        n.attribute("val").map(|s| s.to_string())
    })
}

/// Theme-aware chart-title text color from `node`'s direct-child `<c:title>`,
/// resolved to a hex string (no leading `#`) via the caller's `ColorResolver`.
///
/// Unlike [`extract_chart_title_srgb`] (srgb-only, a historical limitation),
/// this resolves BOTH `<a:srgbClr>` and `<a:schemeClr>` (e.g. `tx2` → the
/// theme's dark-2 slot) plus the surrounding lumMod/lumOff/tint/shade
/// transforms, because chart parts now thread a `&dyn ColorResolver` through
/// `parse_chart_part`. Works for the `<c:chart>` element (chart title) or a
/// `<c:catAx>` / `<c:valAx>` (axis title) since both scope to the node's
/// direct-child `<c:title>`.
///
/// The search is restricted to a `<a:solidFill>` that is a run-property fill
/// (its ancestor chain includes `<a:defRPr>` or `<a:rPr>`), so a title-frame
/// `<c:spPr><a:solidFill>` background fill can never shadow the text color.
/// `None` when there is no `<c:title>`, no run-property solid fill, or the
/// resolver cannot map the contained color (renderer default applies).
pub(super) fn extract_chart_title_color(
    node: Node,
    resolver: &dyn ColorResolver,
) -> Option<String> {
    let title = child(node, "title")?;
    title.descendants().find_map(|n| {
        if !n.is_element() || n.tag_name().name() != "solidFill" {
            return None;
        }
        // Only honor a solidFill that is a text run-property fill — its ancestor
        // chain must pass through a `<a:defRPr>` / `<a:rPr>`. This excludes a
        // `<c:title><c:spPr><a:solidFill>` frame fill.
        let is_run_prop = n
            .ancestors()
            .any(|a| matches!(a.tag_name().name(), "defRPr" | "rPr"));
        if !is_run_prop {
            return None;
        }
        resolver.resolve_solid_fill(n)
    })
}

pub(super) fn extract_axis_title_size(axis_node: Node) -> Option<i32> {
    let title = child(axis_node, "title")?;
    title_text_property_nodes(title)
        .into_iter()
        .find_map(|props| props.attribute("sz").and_then(parse_text_font_size_hpt))
}

pub(super) fn extract_axis_title_bold(axis_node: Node) -> Option<bool> {
    let title = child(axis_node, "title")?;
    chart_text_bool_from_property_cascade(
        title_text_property_nodes(title).into_iter().map(Some),
        "b",
    )
}

pub(super) fn extract_axis_title_italic(axis_node: Node) -> Option<bool> {
    let title = child(axis_node, "title")?;
    chart_text_bool_from_property_cascade(
        title_text_property_nodes(title).into_iter().map(Some),
        "i",
    )
}

#[cfg(test)]
fn extract_axis_title_srgb(axis_node: Node) -> Option<String> {
    let title = child(axis_node, "title")?;
    title_text_property_nodes(title)
        .into_iter()
        .find_map(|props| {
            child(props, "solidFill")
                .and_then(|fill| child(fill, "srgbClr"))
                .and_then(|color| color.attribute("val"))
                .map(ToOwned::to_owned)
        })
}

pub(super) fn extract_axis_title_color(
    axis_node: Node,
    resolver: &dyn ColorResolver,
) -> Option<String> {
    let title = child(axis_node, "title")?;
    title_text_property_nodes(title)
        .into_iter()
        .find_map(|props| {
            child(props, "solidFill").and_then(|fill| resolver.resolve_solid_fill(fill))
        })
}

/// Axis title text + run props from a `<c:catAx>` / `<c:valAx>` node. Rich
/// `rPr` wins over rich `defRPr`, then title-level `txPr`, independently for
/// each property. Run props are resolved only when title text is present.
///
/// NOTE: the color here is srgb-only. Prefer
/// [`extract_axis_title_with_props_resolved`] when a `ColorResolver` is in hand
/// so a `<a:schemeClr>` axis-title color resolves too; this srgb-only variant is
/// kept for callers without a resolver.
#[cfg(test)]
pub(super) fn extract_axis_title_with_props(
    axis_node: Node,
) -> (Option<String>, Option<i32>, Option<bool>, Option<String>) {
    match extract_chart_title_text(axis_node) {
        None => (None, None, None, None),
        Some(text) => (
            Some(text),
            extract_axis_title_size(axis_node),
            extract_axis_title_bold(axis_node),
            extract_axis_title_srgb(axis_node),
        ),
    }
}

/// Like [`extract_axis_title_with_props`] but resolves the axis-title color via
/// the caller's `ColorResolver`, so a `<a:schemeClr>` (theme) axis-title color
/// resolves in addition to a literal `<a:srgbClr>`. All other fields
/// (text/size/bold) are identical. Returns `(text, size_hpt, bold, color_hex)`.
pub(super) fn extract_axis_title_with_props_resolved(
    axis_node: Node,
    resolver: &dyn ColorResolver,
) -> (Option<String>, Option<i32>, Option<bool>, Option<String>) {
    match extract_chart_title_text(axis_node) {
        None => (None, None, None, None),
        Some(text) => (
            Some(text),
            extract_axis_title_size(axis_node),
            extract_axis_title_bold(axis_node),
            // The axis `txPr` is the inherited text default for both its tick
            // labels and an axis title whose own rich-text runs omit color.
            // Direct title run/default properties remain more specific.
            extract_axis_title_color(axis_node, resolver)
                .or_else(|| extract_axis_tick_label_color(axis_node, resolver)),
        ),
    }
}

pub(super) fn axis_title_body_properties<'a, 'input>(
    axis_node: Node<'a, 'input>,
) -> Option<(Option<Node<'a, 'input>>, Option<Node<'a, 'input>>)> {
    let title = child(axis_node, "title")?;
    let rich_body_pr = child(title, "tx")
        .and_then(|tx| child(tx, "rich"))
        .and_then(|rich| child(rich, "bodyPr"));
    let txpr_body_pr = child(title, "txPr").and_then(|txpr| child(txpr, "bodyPr"));
    Some((rich_body_pr, txpr_body_pr))
}

/// Effective local top/bottom text insets for an axis-title text body.
///
/// The rich-text bodyPr wins property-by-property over the title txPr bodyPr;
/// the remaining omissions use the CT_TextBodyProperties defaults. Keeping
/// only the resolved sum in the chart wire model is enough for the renderer's
/// cross-axis layout band without adding a second complete text-body model.
pub(super) fn extract_axis_title_vertical_inset(axis_node: Node) -> Option<i64> {
    let (rich_body_pr, txpr_body_pr) = axis_title_body_properties(axis_node)?;
    let spec = BodyPrDefaults::spec();
    let txpr = txpr_body_pr.map(|body_pr| parse_body_pr(body_pr, &spec));
    let rich_defaults = txpr.as_ref().map_or_else(
        || spec.clone(),
        |body| BodyPrDefaults {
            anchor: body.anchor.clone(),
            wrap: body.wrap.clone(),
            vert: body.vert.clone(),
            l_ins: body.l_ins,
            t_ins: body.t_ins,
            r_ins: body.r_ins,
            b_ins: body.b_ins,
            auto_fit: body.auto_fit.clone(),
        },
    );
    let effective = rich_body_pr
        .map(|body_pr| parse_body_pr(body_pr, &rich_defaults))
        .or(txpr)?;
    Some(effective.t_ins.saturating_add(effective.b_ins))
}

/// Authored DrawingML `bodyPr@rot` for an axis title in raw `ST_Angle` units.
/// Rich-text body properties win over the title-level `txPr` fallback for this
/// property only. `vert` is retained independently by
/// [`extract_axis_title_vertical_mode`] because the two attributes describe
/// different transforms and may legally coexist.
pub(super) fn extract_axis_title_rotation(axis_node: Node) -> Option<i32> {
    let (rich_body_pr, txpr_body_pr) = axis_title_body_properties(axis_node)?;
    rich_body_pr
        .into_iter()
        .chain(txpr_body_pr)
        .find_map(|body_pr| {
            body_pr
                .attribute("rot")
                .and_then(|value| value.parse::<i32>().ok())
        })
}

/// Authored DrawingML `bodyPr@vert` for an axis title. Preserve every schema
/// mode so the renderer can distinguish horizontal, rigid vertical,
/// East-Asian/Mongolian vertical, and WordArt stacking. The current canvas
/// painter explicitly approximates non-rigid vertical modes as vertical flow;
/// retaining the token avoids silently treating them as horizontal.
pub(super) fn extract_axis_title_vertical_mode(axis_node: Node) -> Option<String> {
    let (rich_body_pr, txpr_body_pr) = axis_title_body_properties(axis_node)?;
    rich_body_pr
        .into_iter()
        .chain(txpr_body_pr)
        .find_map(|body_pr| {
            body_pr.attribute("vert").and_then(|vertical| {
                matches!(
                    vertical,
                    "horz"
                        | "vert"
                        | "vert270"
                        | "wordArtVert"
                        | "eaVert"
                        | "mongolianVert"
                        | "wordArtVertRtl"
                )
                .then(|| vertical.to_string())
            })
        })
}

/// `<c|cx:axis><c|cx:title><c|cx:layout><c|cx:manualLayout>` using the same
/// CT_ManualLayout resolver as chart/plot titles. Namespace-local names keep
/// classic charts and ChartEx on one parser path.
pub(super) fn extract_axis_title_manual_layout(axis_node: Node) -> Option<ChartManualLayout> {
    let title = child(axis_node, "title")?;
    let layout = child(title, "layout")?;
    extract_manual_layout(layout)
}

// ============================================================================
// Chart text font faces (CH10) — `<c:txPr>` / `<c:title>` → `<a:latin@typeface>`
// ============================================================================

/// First `<a:latin typeface>` (DrawingML §20.1.4.2.24) descendant of `container`.
/// Empty typefaces are dropped; a theme reference like `+mn-lt` / `+mj-lt` is
/// returned verbatim so the caller can resolve it against the font scheme.
pub(super) fn first_latin_typeface(container: Node) -> Option<String> {
    container.descendants().find_map(|n| {
        if !n.is_element() || n.tag_name().name() != "latin" {
            return None;
        }
        n.attribute("typeface")
            .filter(|s| !s.is_empty())
            .map(|s| s.to_string())
    })
}

/// Resolve a title typeface property-by-property: an authored run face wins
/// over any paragraph/title default regardless of document order.
pub(super) fn title_latin_typeface(container: Node) -> Option<String> {
    title_text_property_nodes(container)
        .into_iter()
        .find_map(first_latin_typeface)
}

/// `<c:catAx|valAx><c:txPr>…<a:latin typeface>` — the axis tick-label font face.
/// Scoped to the axis's `<c:txPr>` so an axis *title* face (under `<c:title>`)
/// is not misread as the tick face. `None` when absent (renderer falls back to
/// the theme body font, then sans-serif).
pub(super) fn extract_axis_tick_label_face(axis_node: Node) -> Option<String> {
    first_latin_typeface(child(axis_node, "txPr")?)
}

/// `<c:catAx|valAx><c:title>…<a:latin typeface>` — the axis-title font face.
/// Scoped to the axis's direct-child `<c:title>`. `None` when absent.
pub(super) fn extract_axis_title_face(axis_node: Node) -> Option<String> {
    title_latin_typeface(child(axis_node, "title")?)
}

/// First chart-group `<c:dLbls><c:txPr>…<a:latin typeface>` — the chart-wide
/// data-label font face. Series-local faces remain on their series.
pub(super) fn extract_data_label_face(root: Node) -> Option<String> {
    root.descendants()
        .filter(|n| is_chart_group_data_labels(*n))
        .find_map(|dlbls| first_latin_typeface(child(dlbls, "txPr")?))
}

/// `<c:legend><c:txPr>` text properties (CH10). Returns
/// `(face, size_hpt, bold, italic)` from the first character-property carrier.
/// Color is resolved separately via [`extract_legend_font_color`] (needs the
/// theme resolver). All `None` when the legend has no `<c:txPr>`.
pub(super) fn extract_legend_text_props(
    root: Node,
) -> (Option<String>, Option<i32>, Option<bool>, Option<bool>) {
    let Some(legend) = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "legend")
    else {
        return (None, None, None, None);
    };
    let Some(txpr) = child(legend, "txPr") else {
        return (None, None, None, None);
    };
    let face = first_latin_typeface(txpr);
    let size = txpr.descendants().find_map(|n| {
        let tag = n.tag_name().name();
        if n.is_element() && (tag == "defRPr" || tag == "rPr") {
            n.attribute("sz").and_then(parse_text_font_size_hpt)
        } else {
            None
        }
    });
    let props = first_chart_text_character_props(txpr);
    let bold = chart_text_bool_from_present_props(props, "b");
    let italic = chart_text_bool_from_present_props(props, "i");
    (face, size, bold, italic)
}

/// `<c:legend><c:txPr>…<a:solidFill>` legend text color, resolved to a hex
/// string (no `#`) via the caller's `ColorResolver`. Scoped to the legend's
/// `<c:txPr>` so a legend-frame `<c:spPr>` fill doesn't leak. `None` when absent.
pub(super) fn extract_legend_font_color(
    root: Node,
    resolver: &dyn ColorResolver,
) -> Option<String> {
    let legend = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "legend")?;
    let txpr = child(legend, "txPr")?;
    txpr.descendants().find_map(|n| {
        if n.is_element() && n.tag_name().name() == "solidFill" {
            resolver.resolve_solid_fill(n)
        } else {
            None
        }
    })
}

/// Parse the optional classic-chart data table (`CT_DTable`). Each border/key
/// child is a CT_Boolean: a present bare element means true, while omission
/// means that feature is not requested. Text and line properties use the same
/// DrawingML grammar as the surrounding chart. The direct resolvable solid
/// fill is retained separately from the table-grid line style; its observed
/// Excel paint extent is an application compatibility behavior, not defined by
/// `CT_DTable` itself.
pub(super) fn extract_chart_data_table(
    plot_area: Node,
    resolver: &dyn ColorResolver,
) -> Option<ChartDataTable> {
    let table = child(plot_area, "dTable")?;
    if !chart_data_table_paint_within_limit(table) {
        return None;
    }
    let txpr = child(table, "txPr");
    let text_props = txpr.and_then(|body| {
        body.descendants()
            .find(|node| node.is_element() && matches!(node.tag_name().name(), "defRPr" | "rPr"))
    });
    let font_color = txpr.and_then(|body| {
        body.descendants().find_map(|node| {
            (node.is_element() && node.tag_name().name() == "solidFill")
                .then(|| resolver.resolve_solid_fill(node))
                .flatten()
        })
    });
    let text_paint = chart_text_body_paint(txpr, resolver);
    let direct_fill = extract_direct_shape_fill(child(table, "spPr"), resolver);
    let direct_line = extract_direct_shape_line(table, resolver);
    Some(ChartDataTable {
        style: parse_direct_chart_effect_style(table, resolver),
        show_horizontal_border: bool_child(table, "showHorzBorder").unwrap_or(false),
        show_vertical_border: bool_child(table, "showVertBorder").unwrap_or(false),
        show_outline: bool_child(table, "showOutline").unwrap_or(false),
        show_keys: bool_child(table, "showKeys").unwrap_or(false),
        font_size_hpt: text_props
            .and_then(|props| props.attribute("sz"))
            .and_then(parse_text_font_size_hpt),
        font_face: txpr.and_then(first_latin_typeface),
        font_color,
        font_paint_authored: text_paint.authored.then_some(true),
        font_hidden: text_paint.hidden.then_some(true),
        font_bold: chart_text_bool_from_present_props(text_props, "b"),
        font_italic: chart_text_bool_from_present_props(text_props, "i"),
        fill_color: direct_fill.color,
        fill: direct_fill.fill,
        fill_hidden: direct_fill.hidden,
        fill_paint_authored: direct_fill.paint_authored,
        line_color: direct_line.color,
        line_width_emu: direct_line.width_emu,
        line_dash: direct_line.dash,
        line_hidden: direct_line.hidden,
        line_paint_authored: direct_line.paint_authored,
    })
}

/// Preflight the direct data-table fill before gradient-stop expansion and
/// sorting. CT_DTable reuses DrawingML shape properties, so it shares the same
/// per-recipe ceiling as every other retained chart paint.
pub(super) fn chart_data_table_paint_within_limit(table: Node) -> bool {
    child(table, "spPr")
        .and_then(chart_style_paint_component_count)
        .unwrap_or(0)
        <= MAX_CHART_PAINT_RECIPE_COMPONENTS
}

// ============================================================================
// Axis scale model (CH6) — gridlines / units / logBase / orientation / labels
// ============================================================================
//
// All helpers take the already-located `<c:catAx>` / `<c:valAx>` node (per
// EG_AxShared, ECMA-376 §21.2.2). `<c:majorGridlines>` / `<c:minorGridlines>`
// are direct children of the axis; `<c:logBase>` / `<c:orientation>` live under
// `<c:scaling>`; `<c:majorUnit>` / `<c:minorUnit>` are direct children of a
// `<c:valAx>` (after `<c:crossBetween>`).

/// `<c:catAx|valAx><c:majorGridlines>` presence (ECMA-376 §21.2.2.100,
/// `CT_ChartLines`). The element carries only an optional `<c:spPr>` line
/// style; its mere PRESENCE requests gridlines. Returns `true` when the axis
/// declares `<c:majorGridlines>`. Office writes it on the value axis by default
/// and omits it on the category axis, so this maps directly to "draw them".
pub(super) fn axis_has_major_gridlines(axis_node: Node) -> bool {
    child(axis_node, "majorGridlines").is_some()
}

/// Whether declared major gridlines have a paintable line. DrawingML
/// `<a:noFill>` suppresses the stroke even though `<c:majorGridlines>` remains
/// present in the chart model. Keep this distinct from
/// [`axis_has_major_gridlines`], which intentionally reports XML presence.
pub(super) fn axis_major_gridlines_visible(axis_node: Node) -> bool {
    let Some(gridlines) = child(axis_node, "majorGridlines") else {
        return false;
    };
    let no_fill = child(gridlines, "spPr")
        .and_then(|sp_pr| child(sp_pr, "ln"))
        .is_some_and(|ln| child(ln, "noFill").is_some());
    !no_fill
}

/// `<c:catAx|valAx><c:majorGridlines><c:spPr><a:ln>` gridline style (ECMA-376
/// §21.2.2.100, `CT_ChartLines` → DrawingML §20.1.2.2.24). The `<c:spPr>` on the
/// gridlines element styles the gridline stroke exactly like `<c:spPr>` on an
/// axis styles the axis rule, so this reuses the same `<a:ln>` resolver. Returns
/// `(color, width_emu, dash)`: the resolved hex (no `#`) when the line carries a
/// `<a:solidFill>` (e.g. `accent3`), and the `<a:ln w>` width in EMU when
/// present, plus the DrawingML preset dash name. All fields are `None` when the
/// axis omits `<c:majorGridlines>` or the
/// element carries no `<c:spPr><a:ln>` — the renderer then keeps its faint
/// default gridline. Visibility is modeled separately by
/// [`axis_major_gridlines_visible`] so `<a:noFill>` suppresses the stroke rather
/// than falling through to a default colour.
pub(super) fn extract_gridline_style_named(
    axis_node: Node,
    resolver: &dyn ColorResolver,
    element_name: &str,
) -> (Option<String>, Option<u32>, Option<String>) {
    let Some(gridlines) = child(axis_node, element_name) else {
        return (None, None, None);
    };
    let (color, width, _no_fill) = extract_sp_pr_ln_style(gridlines, resolver);
    let dash = child(gridlines, "spPr")
        .and_then(|shape| child(shape, "ln"))
        .and_then(|line| child(line, "prstDash"))
        .and_then(|preset| attr(&preset, "val"));
    (color, width, dash)
}

pub(super) fn extract_gridline_style(
    axis_node: Node,
    resolver: &dyn ColorResolver,
) -> (Option<String>, Option<u32>, Option<String>) {
    extract_gridline_style_named(axis_node, resolver, "majorGridlines")
}

pub(super) fn extract_minor_gridline_style(
    axis_node: Node,
    resolver: &dyn ColorResolver,
) -> (Option<String>, Option<u32>, Option<String>) {
    extract_gridline_style_named(axis_node, resolver, "minorGridlines")
}

/// `<c:catAx|valAx><c:minorGridlines>` presence (ECMA-376 §21.2.2.109). Same
/// presence-only semantics as [`axis_has_major_gridlines`]. Minor gridlines
/// require a minor unit to place them; the renderer only draws them when both a
/// `<c:minorGridlines>` element and a resolvable minor step exist.
pub(super) fn axis_has_minor_gridlines(axis_node: Node) -> bool {
    child(axis_node, "minorGridlines").is_some()
}

/// Whether declared minor gridlines have a paintable line. This mirrors
/// [`axis_major_gridlines_visible`]: the element requests minor gridlines, but
/// an authored DrawingML `<a:noFill/>` on their line suppresses the stroke.
pub(super) fn axis_minor_gridlines_visible(axis_node: Node) -> bool {
    let Some(gridlines) = child(axis_node, "minorGridlines") else {
        return false;
    };
    !child(gridlines, "spPr")
        .and_then(|shape| child(shape, "ln"))
        .is_some_and(|line| child(line, "noFill").is_some())
}

/// `<c:valAx><c:majorUnit val>` (ECMA-376 §21.2.2.103, `ST_AxisUnit`
/// §21.2.3.1) — an explicit distance between major ticks/gridlines. Must be a
/// positive floating-point number; non-positive values are rejected so they
/// can't wedge the renderer into an infinite gridline loop. `None` when absent
/// (the renderer keeps its Excel-style auto "nice" step).
pub(super) fn extract_axis_major_unit(axis_node: Node) -> Option<f64> {
    child(axis_node, "majorUnit")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<f64>().ok())
        .filter(|v| v.is_finite() && *v > 0.0)
}

/// `<c:valAx><c:minorUnit val>` (ECMA-376 §21.2.2.112) — explicit distance
/// between minor ticks/gridlines. Positive floating-point; `None` when absent.
pub(super) fn extract_axis_minor_unit(axis_node: Node) -> Option<f64> {
    child(axis_node, "minorUnit")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<f64>().ok())
        .filter(|v| v.is_finite() && *v > 0.0)
}

/// `<c:catAx|valAx><c:scaling><c:logBase val>` (ECMA-376 §21.2.2.98,
/// `ST_LogBase` §21.2.3.25) — the base of a logarithmic value axis. Per the
/// spec the base shall be `>= 2`; smaller/invalid values are rejected. `None`
/// when the axis is linear (the common case).
pub(super) fn extract_axis_log_base(axis_node: Node) -> Option<f64> {
    let scaling = child(axis_node, "scaling")?;
    child(scaling, "logBase")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<f64>().ok())
        .filter(|v| v.is_finite() && *v >= 2.0)
}

/// `<c:catAx|valAx><c:scaling><c:orientation val>` (ECMA-376 §21.2.2.130,
/// `ST_Orientation` §21.2.3.30) — axis direction. Returns the raw enum string
/// `"minMax"` (normal, the default) or `"maxMin"` (reversed). `None` when the
/// element is absent (the renderer treats absent and `"minMax"` identically, so
/// omitting it is byte-stable).
pub(super) fn extract_axis_orientation(axis_node: Node) -> Option<String> {
    let scaling = child(axis_node, "scaling")?;
    child(scaling, "orientation")
        .and_then(|n| n.attribute("val"))
        .map(|s| s.to_string())
}

/// `<c:catAx|valAx><c:tickLblPos val>` (ECMA-376 §21.2.2.207, `ST_TickLblPos`
/// §21.2.3.47) — where the tick labels sit: `"high"` | `"low"` | `"nextTo"`
/// (default) | `"none"` (labels not drawn). Returns the raw enum string; `None`
/// when absent (renderer treats absent as `"nextTo"`, byte-stable).
pub(super) fn extract_axis_tick_label_pos(axis_node: Node) -> Option<String> {
    child(axis_node, "tickLblPos")
        .and_then(|n| n.attribute("val"))
        .map(|s| s.to_string())
}

/// `<c:catAx|valAx><c:txPr><a:bodyPr rot>` (DrawingML `ST_Angle`, 60000ths of a
/// degree — §20.1.10.3) — tick-label rotation. Scoped to the axis's `<c:txPr>`
/// body properties so a title's rotation isn't misread. Returns the raw
/// 60000ths-degree integer; `None` when absent or 0 is not written (renderer
/// treats absent as 0, byte-stable). A value like `-2700000` = -45°.
pub(super) fn extract_axis_tick_label_rotation(axis_node: Node) -> Option<i32> {
    let txpr = child(axis_node, "txPr")?;
    let body_pr = child(txpr, "bodyPr")?;
    body_pr.attribute("rot").and_then(|v| v.parse::<i32>().ok())
}

/// chartEx (`<cx:chartSpace>`) axis visibility. ChartEx encodes the
/// scale type via a `<cx:catScaling>` / `<cx:valScaling>` child rather
/// than separate `<c:catAx>` / `<c:valAx>` elements, so callers can't just
/// reuse `axis_is_deleted` — this helper walks `<cx:axis hidden="1">` and
/// pairs each one with its scaling kind.
///
/// Returns `(cat_hidden, val_hidden)`. Defaults to `(false, false)` when no
/// `<cx:axis>` declares `hidden`.
pub(super) fn extract_chartex_axis_hidden(root: Node) -> (bool, bool) {
    let mut cat_hidden = false;
    let mut val_hidden = false;
    for ax in root
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "axis")
    {
        let hidden = ax.attribute("hidden").map(|v| v == "1").unwrap_or(false);
        if !hidden {
            continue;
        }
        let is_val = ax
            .children()
            .any(|c| c.is_element() && c.tag_name().name() == "valScaling");
        let is_cat = ax
            .children()
            .any(|c| c.is_element() && c.tag_name().name() == "catScaling");
        if is_val {
            val_hidden = true;
        }
        if is_cat {
            cat_hidden = true;
        }
    }
    (cat_hidden, val_hidden)
}

/// ChartEx `<cx:axis><cx:majorTickMarks|minorTickMarks type>`
/// (MS-ODRAWXML §2.24.3.89 CT_TickMarks). Unlike the classic chart axis,
/// ChartEx has no schema default that creates tick marks: an omitted element
/// or omitted `type` therefore resolves to `none`.
pub(super) fn extract_chartex_axis_tick_mark(axis: Option<Node>, name: &str) -> String {
    axis.and_then(|axis| child(axis, name))
        .and_then(|tick_marks| tick_marks.attribute("type"))
        .unwrap_or("none")
        .to_string()
}

/// Text saved by chartEx in either DrawingML-rich or compact `txData/v` form.
pub(super) fn chartex_text(container: Node) -> Option<String> {
    if let Some(rich) = container
        .descendants()
        .find(|node| node.is_element() && node.tag_name().name() == "rich")
    {
        let text = flatten_rich_text(rich, None);
        if !text.is_empty() {
            return Some(text);
        }
    }
    container
        .descendants()
        .find(|node| node.is_element() && node.tag_name().name() == "txData")
        .and_then(|tx_data| child(tx_data, "v"))
        .and_then(|value| value.text())
        .map(str::trim)
        .filter(|text| !text.is_empty())
        .map(ToOwned::to_owned)
}
