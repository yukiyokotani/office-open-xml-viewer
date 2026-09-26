use super::*;

#[derive(Default)]
pub(super) struct DirectShapeFill {
    pub(super) color: Option<String>,
    pub(super) fill: Option<ChartStyleFill>,
    pub(super) hidden: Option<bool>,
    pub(super) paint_authored: Option<bool>,
}

#[derive(Default)]
pub(super) struct DirectShapeLine {
    pub(super) color: Option<String>,
    pub(super) fill: Option<ChartStyleFill>,
    pub(super) width_emu: Option<u32>,
    pub(super) dash: Option<String>,
    pub(super) dash_authored: Option<bool>,
    pub(super) custom_dash: Option<Vec<ChartLineDashSegment>>,
    pub(super) cap: Option<String>,
    pub(super) join: Option<String>,
    pub(super) compound: Option<String>,
    pub(super) hidden: Option<bool>,
    pub(super) paint_authored: Option<bool>,
}

pub(super) fn extract_direct_shape_fill(
    shape: Option<Node>,
    resolver: &dyn ColorResolver,
) -> DirectShapeFill {
    extract_direct_shape_fill_with_images(
        shape,
        resolver,
        &EmptyChartImageResolver,
        ChartImageSource::Chart,
    )
}

pub(super) fn extract_direct_shape_fill_with_images(
    shape: Option<Node>,
    resolver: &dyn ColorResolver,
    image_resolver: &dyn ChartImageResolver,
    image_source: ChartImageSource,
) -> DirectShapeFill {
    let fill_paint = shape.and_then(|shape| {
        parse_chart_style_paint(shape, resolver, None, image_resolver, image_source)
    });
    let paint_authored = shape
        .and_then(|shape| {
            shape.children().find(|node| {
                node.is_element()
                    && matches!(
                        node.tag_name().name(),
                        "noFill" | "solidFill" | "gradFill" | "pattFill" | "blipFill" | "grpFill"
                    )
            })
        })
        .map(|_| true);
    // Keep the established shape-level resolver fallback: host resolvers may
    // resolve a transformed theme fill as an effective solid even when the
    // lower-level color node cannot be projected independently.
    let resolved_shape_fill = shape.and_then(|shape| {
        (child(shape, "noFill").is_none())
            .then(|| resolver.resolve_shape_fill(shape))
            .flatten()
    });
    let (color, fill, hidden) = match fill_paint {
        Some(ChartStylePaint::NoFill) => (None, None, Some(true)),
        Some(ChartStylePaint::Fill(fill)) => {
            let fill = *fill;
            let color = match &fill {
                ChartStyleFill::Solid { color } => Some(color.clone()),
                _ => resolved_shape_fill.clone(),
            };
            (color, Some(fill), None)
        }
        Some(ChartStylePaint::Unresolved) => (
            resolved_shape_fill.clone(),
            resolved_shape_fill.map(|color| ChartStyleFill::Solid { color }),
            None,
        ),
        None => (None, None, None),
    };
    DirectShapeFill {
        color,
        fill,
        hidden,
        paint_authored,
    }
}

pub(super) fn extract_direct_shape_line_from_sp_pr(
    shape: Option<Node>,
    resolver: &dyn ColorResolver,
) -> DirectShapeLine {
    use crate::line::LineDash;

    let line = shape.and_then(|shape| child(shape, "ln"));
    let parsed_line = line.map(|line| parse_chart_style_line(line, resolver, None));
    let paint_authored = line
        .and_then(|line| {
            line.children().find(|child| {
                child.is_element()
                    && matches!(
                        child.tag_name().name(),
                        "noFill" | "solidFill" | "gradFill" | "pattFill"
                    )
            })
        })
        .map(|_| true);
    let color = line.and_then(|line| resolver.resolve_shape_fill(line));
    let width_emu = line
        .and_then(|line| line.attribute("w"))
        .and_then(|value| value.parse::<u32>().ok());
    let no_fill = line.is_some_and(|line| child(line, "noFill").is_some());
    let fill = parsed_line
        .as_ref()
        .and_then(|line| line.paint.as_ref())
        .and_then(chart_style_fill_from_line_paint);
    let dash = parsed_line
        .as_ref()
        .and_then(|line| match line.dash.as_ref() {
            Some(LineDash::Preset(value)) => value.clone(),
            _ => None,
        });
    let custom_dash = parsed_line
        .as_ref()
        .and_then(|line| match line.dash.as_ref() {
            Some(LineDash::Custom(stops)) => Some(
                stops
                    .iter()
                    .map(|stop| ChartLineDashSegment {
                        dash: stop.dash / 100_000.0,
                        space: stop.space / 100_000.0,
                    })
                    .collect(),
            ),
            _ => None,
        });
    DirectShapeLine {
        color: if no_fill { None } else { color },
        fill,
        width_emu,
        dash,
        dash_authored: parsed_line
            .as_ref()
            .and_then(|line| line.dash.as_ref())
            .map(|_| true),
        custom_dash,
        cap: parsed_line.as_ref().and_then(|line| line.cap.clone()),
        join: parsed_line
            .as_ref()
            .and_then(|line| match line.join.as_ref() {
                Some(crate::line::LineJoin::Round) => Some("round".to_owned()),
                Some(crate::line::LineJoin::Bevel) => Some("bevel".to_owned()),
                Some(crate::line::LineJoin::Miter { .. }) => Some("miter".to_owned()),
                None => None,
            }),
        compound: parsed_line.and_then(|line| line.compound),
        hidden: no_fill.then_some(true),
        paint_authored,
    }
}

pub(super) fn extract_direct_shape_line(
    node: Node,
    resolver: &dyn ColorResolver,
) -> DirectShapeLine {
    extract_direct_shape_line_from_sp_pr(child(node, "spPr"), resolver)
}

#[derive(Default)]
pub(super) struct LegendFrameStyle {
    pub(super) fill_color: Option<String>,
    pub(super) fill: Option<ChartStyleFill>,
    pub(super) fill_hidden: Option<bool>,
    pub(super) fill_paint_authored: Option<bool>,
    pub(super) line_color: Option<String>,
    pub(super) line_fill: Option<ChartStyleFill>,
    pub(super) line_width_emu: Option<u32>,
    pub(super) line_dash: Option<String>,
    pub(super) line_dash_authored: Option<bool>,
    pub(super) line_custom_dash: Option<Vec<ChartLineDashSegment>>,
    pub(super) line_cap: Option<String>,
    pub(super) line_join: Option<String>,
    pub(super) line_compound: Option<String>,
    pub(super) line_hidden: Option<bool>,
    pub(super) line_paint_authored: Option<bool>,
}

/// `<c:legend><c:spPr>` frame paint. Direct fill and line paint are preserved
/// independently, including `noFill` and unresolved authored paint, so a
/// linked Chart Style can supply only genuinely omitted properties.
pub(super) fn extract_legend_frame_style(
    root: Node,
    resolver: &dyn ColorResolver,
) -> LegendFrameStyle {
    let Some(legend) = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "legend")
    else {
        return LegendFrameStyle::default();
    };
    let shape = child(legend, "spPr");
    let direct_fill = extract_direct_shape_fill(shape, resolver);
    let direct_line = extract_direct_shape_line(legend, resolver);
    LegendFrameStyle {
        fill_color: direct_fill.color,
        fill: direct_fill.fill,
        fill_hidden: direct_fill.hidden,
        fill_paint_authored: direct_fill.paint_authored,
        line_color: direct_line.color,
        line_fill: direct_line.fill,
        line_width_emu: direct_line.width_emu,
        line_dash: direct_line.dash,
        line_dash_authored: direct_line.dash_authored,
        line_custom_dash: direct_line.custom_dash,
        line_cap: direct_line.cap,
        line_join: direct_line.join,
        line_compound: direct_line.compound,
        line_hidden: direct_line.hidden,
        line_paint_authored: direct_line.paint_authored,
    }
}

// ============================================================================
// Pie / doughnut geometry (CH8)
// ============================================================================

/** Parse the OOXML percentage unions that accept either an unsigned integer
 * or the Strict/Transitional percentage lexical form (`"100%"`). */
pub(super) fn parse_unsigned_percent(value: &str) -> Option<u32> {
    value.strip_suffix('%').unwrap_or(value).parse::<u32>().ok()
}

/// `<c:doughnutChart><c:holeSize val>` (§21.2.2.82) — hole diameter percentage
/// (1–90). Clamped to the ECMA range. `None` when absent. `root` is the chart
/// space (or `<c:chart>`); the search is scoped to a `<c:doughnutChart>` so a
/// hole size only ever comes from a doughnut plot.
pub(super) fn extract_hole_size(root: Node) -> Option<u32> {
    let doughnut = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "doughnutChart")?;
    child(doughnut, "holeSize")
        .and_then(|n| n.attribute("val"))
        .and_then(parse_unsigned_percent)
        .map(|v| v.clamp(1, 90))
}

/// `<c:pieChart|doughnutChart><c:firstSliceAng val>` (§21.2.2.52) — start angle
/// in degrees (0–360, clockwise from 12 o'clock). Clamped to the ECMA range.
/// `None` when absent (renderer defaults to 0).
pub(super) fn extract_first_slice_angle(root: Node) -> Option<u32> {
    root.descendants()
        .find(|n| {
            n.is_element()
                && (n.tag_name().name() == "pieChart" || n.tag_name().name() == "doughnutChart")
        })
        .and_then(|pie| child(pie, "firstSliceAng"))
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<u32>().ok())
        .map(|v| v.min(360))
}

/// `<c:dPt><c:explosion val>` (§21.2.2.61) — pie/doughnut slice pull-out
/// amount, parsed as the unbounded `xsd:unsignedInt` the schema (`CT_UnsignedInt`)
/// actually specifies (no 0–100 clamp here; see `ChartDataPointOverride::explosion`
/// for how renderers interpret the value). Caller passes a `<c:dPt>` node.
/// `None` when absent.
pub(super) fn extract_dpt_explosion(dpt_node: Node) -> Option<u32> {
    child(dpt_node, "explosion")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<u32>().ok())
}

/// Explicit chart-frame border from `<c:chartSpace><c:spPr><a:ln>` (ECMA-376
/// §21.2.2.5 / DrawingML §20.1.2.2.24). `chart_space_root` is the
/// `<c:chartSpace>` element. Returns `(srgb_color, width_emu)` under the locked
/// policy shared by both parsers: a border is drawn ONLY when the XML explicitly
/// declares a paintable line.
///
///  - no `<a:ln>` (or no `<c:spPr>`) → `(None, None)` — no default border;
///  - `<a:ln><a:noFill/>` → border explicitly off → color `None` (width still
///    reported when `@w` is present);
///  - `<a:ln><a:solidFill><a:srgbClr@val>` → `(Some(hex), width)`.
///
/// `@w` (EMU) is captured as `u32` regardless of the fill. `<a:schemeClr>` is
/// intentionally left unresolved here (theme not wired through to chart border
/// parsing yet).
/// `<c:date1904>` (ECMA-376 §21.2.2.38) as a direct child of `<c:chartSpace>`.
/// The element is a `CT_Boolean`: `val` defaults to `true` when the element is
/// present but the attribute is omitted, so `<c:date1904/>` alone means
/// date1904=true. `val="0"` / `"false"` disable it. Absent element ⇒ false (the
/// default 1900 date system, §18.17.4.1).
pub(super) fn extract_chart_date1904(chart_space_root: Node) -> bool {
    match child(chart_space_root, "date1904") {
        Some(n) => match n.attribute("val") {
            None => true, // element present, val implied true
            Some(v) => v == "1" || v.eq_ignore_ascii_case("true"),
        },
        None => false,
    }
}

/// `<c:ser><c:smooth val>` (ECMA-376 §21.2.2.194) — line/area series smoothing
/// flag. `ser_node` is the `<c:ser>` element. Returns `Some(true/false)` when
/// the element is present (CT_Boolean: `val` implied true when omitted),
/// `None` when the series has no `<c:smooth>` (straight-polyline default). Shared
/// so the pptx and xlsx parsers honor the flag identically.
pub(super) fn extract_series_smooth(ser_node: Node) -> Option<bool> {
    child(ser_node, "smooth").map(|n| match n.attribute("val") {
        None => true, // element present, val implied true
        Some(v) => v == "1" || v.eq_ignore_ascii_case("true"),
    })
}

/// Parse `bool_val`: a `CT_Boolean` child's `val` where an absent attribute
/// implies true (the OOXML default when the element is present).
pub(super) fn bool_child(parent: Node, name: &str) -> Option<bool> {
    child(parent, name).map(|n| match n.attribute("val") {
        None => true,
        Some(v) => v == "1" || v.eq_ignore_ascii_case("true"),
    })
}

/// Read one `CT_Boolean` child without accepting producer-specific lexical
/// aliases. ECMA-376 defines xsd:boolean and an implied `true` value when the
/// element is present without `val`. A present invalid value remains authored
/// but fails closed to `false` instead of reviving a less-specific `true`.
pub(super) fn strict_boolean_child(parent: Node, name: &str) -> Option<bool> {
    child(parent, name).map(|node| {
        node.attribute("val")
            .map(|value| crate::drawing::parse_xsd_bool(value).unwrap_or(false))
            .unwrap_or(true)
    })
}

/// `<c:ser><c:trendline>` (ECMA-376 §21.2.2.211, `CT_Trendline`) — every
/// trendline declared on `ser_node` (0..N). Each carries a required
/// `<c:trendlineType>` plus optional order/period/forward/backward/intercept,
/// the `<c:dispRSqr>` / `<c:dispEq>` label flags, optional
/// `<c:trendlineLbl>` layout/text properties, and an `<c:spPr><a:ln>` line style
/// (color resolved via `resolver`, width in EMU). Returns `None` when the series
/// declares no trendline (byte-stable); otherwise the parsed vec. Shared so pptx
/// and xlsx honor trendlines identically.
#[cfg(test)]
pub(super) fn extract_series_trendlines(
    ser_node: Node,
    resolver: &dyn ColorResolver,
) -> Option<Vec<ChartTrendline>> {
    let label_shapes = ser_node
        .children()
        .filter(|node| node.is_element() && node.tag_name().name() == "trendline")
        .filter_map(|trendline| child(trendline, "trendlineLbl"))
        .filter_map(|label| child(label, "spPr"));
    let allow_label_paints = label_paint_recipes_within_budget(label_shapes);
    extract_series_trendlines_with_paint_policy(ser_node, resolver, allow_label_paints)
}

pub(super) fn extract_series_trendlines_with_paint_policy(
    ser_node: Node,
    resolver: &dyn ColorResolver,
    allow_label_paints: bool,
) -> Option<Vec<ChartTrendline>> {
    let mut out = Vec::new();
    for tl in ser_node
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "trendline")
    {
        // trendlineType is required per the schema; skip a malformed trendline
        // that somehow lacks it rather than emitting an empty type.
        let Some(trendline_type) = child(tl, "trendlineType").and_then(|n| n.attribute("val"))
        else {
            continue;
        };
        let u32_val = |name: &str| -> Option<u32> {
            child(tl, name)
                .and_then(|n| n.attribute("val"))
                .and_then(|v| v.parse::<u32>().ok())
        };
        let f64_val = |name: &str| -> Option<f64> {
            child(tl, name)
                .and_then(|n| n.attribute("val"))
                .and_then(|v| v.parse::<f64>().ok())
        };
        // Preserve the paint component independently from its resolved colour:
        // an unsupported direct paint still outranks linked/numeric styles.
        let direct_line = extract_direct_shape_line(tl, resolver);
        let label = child(tl, "trendlineLbl");
        let label_txpr = label.and_then(|node| child(node, "txPr"));
        let label_tx = label.and_then(|node| child(node, "tx"));
        let label_rich = label_tx.and_then(|tx| child(tx, "rich"));
        // Label-wide fallback comes only from paragraph defaults. Individual
        // rich-run rPr stays on `label_rich_runs`; promoting it here would make
        // one run's noFill/colour/typography leak into unformatted siblings.
        let txpr_default_prop = label_txpr.and_then(first_paragraph_default_run_props);
        let run_props = [txpr_default_prop];
        let label_text = label_tx
            .and_then(|tx| {
                child(tx, "rich")
                    .map(|rich| flatten_rich_text(rich, None))
                    .or_else(|| {
                        child(tx, "strRef")
                            .and_then(|reference| child(reference, "strCache"))
                            .map(|cache| {
                                cache
                                    .children()
                                    .filter(|node| {
                                        node.is_element() && node.tag_name().name() == "pt"
                                    })
                                    .filter_map(|point| child(point, "v"))
                                    .filter_map(|value| value.text())
                                    .collect::<Vec<_>>()
                                    .join("\n")
                            })
                    })
            })
            .filter(|text| !text.is_empty());
        let label_rich_runs =
            label.and_then(|node| parse_data_label_rich_runs(node, resolver, None));
        let label_manual_layout = label
            .and_then(|node| child(node, "layout"))
            .and_then(extract_manual_layout);
        let label_num_fmt = label.and_then(|node| child(node, "numFmt"));
        let label_format_code = label_num_fmt
            .and_then(|node| node.attribute("formatCode"))
            .map(str::to_string);
        let label_format_source_linked = label_num_fmt
            .and_then(|node| node.attribute("sourceLinked"))
            .map(|value| value == "1" || value.eq_ignore_ascii_case("true"));
        let label_font_size_hpt = run_props
            .iter()
            .flatten()
            .find_map(|node| attr(node, "sz").and_then(|value| parse_text_font_size_hpt(&value)));
        let label_font_bold =
            chart_text_bool_from_present_props(run_props.iter().flatten().next().copied(), "b");
        let label_font_italic =
            chart_text_bool_from_present_props(run_props.iter().flatten().next().copied(), "i");
        let label_text_paint = chart_text_paint(run_props, resolver);
        let label_font_color = label_text_paint.color.clone();
        let label_font_face = run_props
            .iter()
            .flatten()
            .find_map(|node| first_latin_typeface(*node));
        let label_font_language = run_props
            .iter()
            .flatten()
            .find_map(|node| node.attribute("lang").map(ToOwned::to_owned));
        let label_font_baseline = run_props.iter().flatten().find_map(|node| {
            node.attribute("baseline")
                .and_then(parse_chart_text_percentage)
        });
        let label_body = merge_chart_label_body_styles(
            chart_label_body_style(label_rich.and_then(|rich| child(rich, "bodyPr"))),
            chart_label_body_style(label_txpr.and_then(|txpr| child(txpr, "bodyPr"))),
        );
        let label_text_align = label_txpr
            .and_then(first_paragraph_properties)
            .and_then(|node| attr(&node, "algn"));
        out.push(ChartTrendline {
            style: parse_direct_chart_effect_style(tl, resolver),
            name: child(tl, "name")
                .and_then(|node| node.text())
                .map(str::to_string)
                .filter(|name| !name.is_empty()),
            trendline_type: trendline_type.to_string(),
            order: u32_val("order"),
            period: u32_val("period"),
            forward: f64_val("forward"),
            backward: f64_val("backward"),
            intercept: f64_val("intercept"),
            disp_r_sqr: bool_child(tl, "dispRSqr"),
            disp_eq: bool_child(tl, "dispEq"),
            label_manual_layout,
            label_text,
            label_rich_runs,
            label_format_code,
            label_format_source_linked,
            label_font_size_hpt,
            label_font_bold,
            label_font_italic,
            label_font_color,
            label_font_paint_authored: label_text_paint.authored.then_some(true),
            label_font_hidden: label_text_paint.hidden.then_some(true),
            label_font_face,
            label_font_language,
            label_font_baseline,
            label_text_rotation: label_body.rotation,
            label_text_wrap: label_body.wrap,
            label_text_vertical_anchor: label_body.anchor,
            label_text_vertical_mode: label_body.vertical_mode,
            label_text_l_ins_emu: label_body.left_inset,
            label_text_t_ins_emu: label_body.top_inset,
            label_text_r_ins_emu: label_body.right_inset,
            label_text_b_ins_emu: label_body.bottom_inset,
            label_text_body_authored: label_body.authored.then_some(true),
            label_box: label.and_then(|node| {
                parse_label_box_with_policy(
                    node.children()
                        .find(|child| child.is_element() && child.tag_name().name() == "spPr"),
                    resolver,
                    allow_label_paints,
                )
            }),
            label_text_align,
            line_color: direct_line.color,
            line_width_emu: direct_line.width_emu,
            line_dash: direct_line.dash,
            line_hidden: direct_line.hidden,
            line_paint_authored: direct_line.paint_authored,
        });
    }
    if out.is_empty() {
        None
    } else {
        Some(out)
    }
}

/// `<c:chart><c:dispBlanksAs val>` (ECMA-376 §21.2.2.42, `ST_DispBlanksAs`
/// §21.2.3.10) — how blank cells are plotted ("gap" | "zero" | "span").
/// `root` may be the `<c:chartSpace>` or `<c:chart>` node; the single
/// `<c:dispBlanksAs>` is found by descendant walk either way. Returns `None`
/// when the element is absent (the renderer defaults to "gap"). Per the XSD the
/// `@val` default is "zero" (applies only when `<c:dispBlanksAs/>` is present
/// but the attribute is omitted). Shared so pptx and xlsx behave identically.
pub(super) fn extract_disp_blanks_as(root: Node) -> Option<String> {
    root.descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "dispBlanksAs")
        .map(|n| n.attribute("val").unwrap_or("zero").to_string())
}

/// `<c:chart><c:showDLblsOverMax>` (ECMA-376 §21.2.2.180). The element is a
/// `CT_Boolean`, so a bare element implies true. Absence remains `None` to
/// preserve the authored chart-level state; its effective schema behavior is
/// false in the renderer.
pub(super) fn extract_show_data_labels_over_max(root: Node) -> Option<bool> {
    root.descendants()
        .find(|node| node.is_element() && node.tag_name().name() == "chart")
        .and_then(|chart| bool_child(chart, "showDLblsOverMax"))
}

/// `<c:chartSpace><c:roundedCorners>` (§21.2.2.159). The element is optional;
/// when present without `val`, CT_Boolean's attribute default is true.
pub(super) fn extract_chart_space_rounded_corners(root: Node) -> Option<bool> {
    child(root, "roundedCorners").map(|node| {
        node.attribute("val")
            .map(|value| value == "1" || value.eq_ignore_ascii_case("true"))
            .unwrap_or(true)
    })
}

#[cfg(test)]
pub(super) fn extract_chart_space_border(chart_space_root: Node) -> (Option<String>, Option<u32>) {
    let Some(ln) = child(chart_space_root, "spPr").and_then(|sp| child(sp, "ln")) else {
        return (None, None);
    };
    let width = ln.attribute("w").and_then(|v| v.parse::<u32>().ok());
    // An explicit `<a:noFill/>` turns the border off → no color.
    if child(ln, "noFill").is_some() {
        return (None, width);
    }
    // Only an srgbClr inside a direct `<a:solidFill>` is honored.
    let color = child(ln, "solidFill")
        .and_then(|sf| child(sf, "srgbClr"))
        .and_then(|srgb| srgb.attribute("val"))
        .map(|s| s.to_string());
    (color, width)
}

/// First chart-group `<c:dLbls><c:txPr>` font size (hpt). Series-local sizes
/// stay on `ChartSeriesDataLabels` and must not become a sibling fallback.
pub(super) fn extract_data_label_font_size(root: Node) -> Option<i32> {
    root.descendants()
        .filter(|n| is_chart_group_data_labels(*n))
        .find_map(|dl| {
            child(dl, "txPr").and_then(|tx| {
                tx.descendants().find_map(|n| {
                    if !n.is_element() {
                        return None;
                    }
                    let tag = n.tag_name().name();
                    if tag != "defRPr" && tag != "rPr" {
                        return None;
                    }
                    n.attribute("sz").and_then(parse_text_font_size_hpt)
                })
            })
        })
}

/// First chart-group data-label bold state from `<c:dLbls><c:txPr>`. A present
/// character-property node with omitted `b` is false.
pub(super) fn extract_data_label_font_bold(root: Node) -> Option<bool> {
    root.descendants()
        .filter(|n| is_chart_group_data_labels(*n))
        .find_map(|labels| {
            child(labels, "txPr").and_then(|tx| {
                chart_text_bool_from_present_props(first_chart_text_character_props(tx), "b")
            })
        })
}

/// First chart-group data-label italic flag from `<c:dLbls><c:txPr>`.
pub(super) fn extract_data_label_font_italic(root: Node) -> Option<bool> {
    root.descendants()
        .filter(|n| is_chart_group_data_labels(*n))
        .find_map(|labels| {
            child(labels, "txPr").and_then(|tx| {
                chart_text_bool_from_present_props(first_chart_text_character_props(tx), "i")
            })
        })
}

/// First chart-group `<c:dLbls><c:txPr>...<a:solidFill>` resolved to a hex
/// color. Series-local text paint remains on its series.
///
/// Note we deliberately scope the search to inside `<c:txPr>` so a
/// sibling `<c:dLbls><c:spPr><a:solidFill>` (the label *background*
/// fill, distinct from the text color) can't shadow the answer.
pub(super) fn extract_data_label_font_color(
    root: Node,
    resolver: &dyn ColorResolver,
) -> Option<String> {
    for dlbls in root
        .descendants()
        .filter(|n| is_chart_group_data_labels(*n))
    {
        let Some(txpr) = child(dlbls, "txPr") else {
            continue;
        };
        for desc in txpr.descendants().filter(|n| n.is_element()) {
            if desc.tag_name().name() != "solidFill" {
                continue;
            }
            if let Some(c) = resolver.resolve_solid_fill(desc) {
                return Some(c);
            }
        }
    }
    None
}

/// `<c:catAx|valAx><c:txPr>` tick-label text color, resolved to a hex string
/// (no leading `#`). Walks the axis's `<c:txPr>` for the first descendant
/// `<a:solidFill>` the resolver can map — this is the `<a:defRPr><a:solidFill>`
/// that ECMA-376 §21.2.2.* / §21.1.2.2.* uses to color the axis tick labels
/// (e.g. PowerPoint's "category labels in gray"). Scoped to `<c:txPr>` so the
/// sibling `<c:spPr>` axis-line fill can't shadow the answer.
pub(super) fn extract_axis_tick_label_color(
    axis_node: Node,
    resolver: &dyn ColorResolver,
) -> Option<String> {
    let txpr = child(axis_node, "txPr")?;
    for desc in txpr.descendants().filter(|n| n.is_element()) {
        if desc.tag_name().name() != "solidFill" {
            continue;
        }
        if let Some(c) = resolver.resolve_solid_fill(desc) {
            return Some(c);
        }
    }
    None
}

/// `<c:catAx|valAx><c:spPr><a:ln>` axis-line style (ECMA-376 §21.2.2.* line
/// properties via DrawingML §20.1.2.2.24). Returns `(color, width_emu, no_fill)`:
///
///  - `color`: resolved hex (no `#`) when the line carries a `<a:solidFill>`.
///  - `width_emu`: the `<a:ln w>` width in EMU when present.
///  - `no_fill`: true when the line is explicitly `<a:noFill>`. The shared
///    model records this separately from axis deletion; Office suppresses the
///    axis rule and its tick marks while retaining labels and gridlines.
///
/// When the axis has no `<c:spPr><a:ln>` at all the tuple is
/// `(None, None, false)` and the caller falls back to its default rule.
pub(super) fn extract_axis_line_style(
    axis_node: Node,
    resolver: &dyn ColorResolver,
) -> (Option<String>, Option<u32>, bool) {
    extract_sp_pr_ln_style(axis_node, resolver)
}

/// `<c:catAx|dateAx|valAx|serAx><c:spPr><a:ln><a:prstDash val>`.
/// The dash token is kept separate from [`extract_axis_line_style`] so the
/// existing color/width/noFill contract remains stable for host callers.
pub(super) fn extract_axis_line_dash(axis_node: Node) -> Option<String> {
    child(axis_node, "spPr")
        .and_then(|shape| child(shape, "ln"))
        .and_then(|line| child(line, "prstDash"))
        .and_then(|dash| dash.attribute("val"))
        .map(ToOwned::to_owned)
}

/// `<…><c:spPr><a:ln>` line style for any node that carries a `<c:spPr>` shape
/// property (an axis, a `<c:majorGridlines>` element, etc.). Returns
/// `(color, width_emu, no_fill)` with the same contract as
/// [`extract_axis_line_style`]:
///
///  - `color`: resolved hex (no `#`) when the line carries a `<a:solidFill>`.
///  - `width_emu`: the `<a:ln w>` width in EMU when present.
///  - `no_fill`: true when the line is explicitly `<a:noFill>`.
///
/// `(None, None, false)` when the node has no `<c:spPr><a:ln>`.
pub(super) fn extract_sp_pr_ln_style(
    node: Node,
    resolver: &dyn ColorResolver,
) -> (Option<String>, Option<u32>, bool) {
    let Some(sp_pr) = child(node, "spPr") else {
        return (None, None, false);
    };
    let Some(ln) = child(sp_pr, "ln") else {
        return (None, None, false);
    };
    let width = ln.attribute("w").and_then(|v| v.parse::<u32>().ok());
    let no_fill = child(ln, "noFill").is_some();
    // The rule is a SHAPE stroke, so resolve it through `resolve_shape_fill`
    // (full DrawingML grammar incl. lumMod/lumOff tints). xlsx keeps its lighter
    // transform-free `resolve_solid_fill` for series/legend/title fills, so a
    // scheme-color line (e.g. a `bg1 lumMod 65%` light-gray rule, or an
    // `accent3` gridline) must go through the shape path to render at the right
    // strength rather than its untransformed base color.
    let color = resolver.resolve_shape_fill(ln);
    (color, width, no_fill)
}

pub(super) fn parse_chart_decoration_line_style(
    node: Node,
    resolver: &dyn ColorResolver,
) -> ChartDecorationLineStyle {
    let direct = extract_direct_shape_line(node, resolver);
    ChartDecorationLineStyle {
        style: parse_direct_chart_effect_style(node, resolver),
        color: direct.color,
        fill: direct.fill,
        paint_authored: direct.paint_authored,
        width_emu: direct.width_emu,
        dash: direct.dash,
        cap: direct.cap,
        join: direct.join,
        hidden: direct.hidden,
    }
}

pub(super) fn parse_chart_up_down_bar_paint(
    bar: Option<Node>,
    resolver: &dyn ColorResolver,
) -> ChartStockBarPaint {
    let Some(bar) = bar else {
        return ChartStockBarPaint::default();
    };
    let sp_pr = child(bar, "spPr");
    let direct_fill = extract_direct_shape_fill(sp_pr, resolver);
    let direct_line = extract_direct_shape_line(bar, resolver);
    ChartStockBarPaint {
        style: parse_direct_chart_effect_style(bar, resolver),
        fill_color: direct_fill.color,
        fill: direct_fill.fill,
        fill_paint_authored: direct_fill.paint_authored,
        fill_hidden: direct_fill.hidden,
        line_color: direct_line.color,
        line_paint_authored: direct_line.paint_authored,
        line_width_emu: direct_line.width_emu,
        line_dash: direct_line.dash,
        line_cap: direct_line.cap,
        line_join: direct_line.join,
        line_hidden: direct_line.hidden,
    }
}

pub(super) fn parse_chart_up_down_bar_style(
    node: Node,
    resolver: &dyn ColorResolver,
) -> ChartStockUpDownBarStyle {
    ChartStockUpDownBarStyle {
        gap_width_percent: child(node, "gapWidth")
            .and_then(|value| value.attribute("val"))
            .and_then(|value| value.trim_end_matches('%').parse::<f64>().ok())
            .filter(|value| value.is_finite() && *value >= 0.0)
            .unwrap_or(150.0),
        up: parse_chart_up_down_bar_paint(child(node, "upBars"), resolver),
        down: parse_chart_up_down_bar_paint(child(node, "downBars"), resolver),
    }
}

/// Resolve a color inside a chart-style recipe. `phClr` is the branch/series
/// accent supplied by the style reference; fixed scheme colors keep resolving
/// through the host theme. DrawingML color transforms remain on the authored
/// color node and are therefore applied by the shared resolver.
pub(super) fn resolve_chart_style_color(
    color_container: Node,
    resolver: &dyn ColorResolver,
    placeholder: Option<&str>,
) -> Option<String> {
    let adapter = ColorResolverThemeAdapter(resolver);
    let style_resolver = crate::color::StyleMatrixColorResolver::new(&adapter, placeholder);
    crate::color::parse_color_node(color_container, &style_resolver, resolver.tint_mode())
}

#[derive(Debug, Clone, PartialEq)]
pub(super) enum ChartStylePaint {
    NoFill,
    Unresolved,
    Fill(Box<ChartStyleFill>),
}

pub(super) fn chart_style_reference_color(
    reference: Node,
    resolver: &dyn ColorResolver,
    accent: Option<&str>,
    palette: Option<&[Option<String>]>,
    color_style_method: Option<&str>,
) -> Result<Option<String>, ()> {
    const DRAWINGML_COLORS: &[&str] = &[
        "scrgbClr",
        "srgbClr",
        "hslClr",
        "sysClr",
        "schemeClr",
        "prstClr",
    ];
    // CT_StyleReference accepts a normal DrawingML color choice as an
    // alternative to CT_StyleColor. It is a fixed reference color, independent
    // of the linked Chart Colors part.
    if reference
        .children()
        .any(|node| node.is_element() && DRAWINGML_COLORS.contains(&node.tag_name().name()))
    {
        return resolve_chart_style_color(reference, resolver, None)
            .map(Some)
            .ok_or(());
    }
    let Some(style_color) = child(reference, "styleClr") else {
        return Ok(None);
    };
    let value = style_color.attribute("val").unwrap_or("auto");
    // MS-ODRAWXML §2.8.4.6 ST_StyleColorVal: unsigned integers are fixed
    // zero-based indexes, `auto` is the relative object index, and every other
    // string maps to index zero. It is not an RGB or theme-scheme value.
    let selected = if value == "auto" {
        accent.map(str::to_owned)
    } else {
        let index = value.parse::<usize>().unwrap_or(0);
        palette
            .and_then(|colors| {
                chart_color_style_base_index(color_style_method, index, colors.len())
                    .and_then(|mapped| colors.get(mapped))
            })
            .and_then(Clone::clone)
    }
    .ok_or(())?;
    // CT_StyleColor is itself a DrawingML color-transform container.
    let ignore_transforms = reference.attribute("mods").is_some_and(|mods| {
        mods.split_ascii_whitespace()
            .any(|modifier| modifier == "ignoreCSTransforms")
    });
    Ok(Some(if ignore_transforms {
        selected
    } else {
        crate::color::apply_color_transforms(&selected, style_color, resolver.tint_mode())
    }))
}

pub(super) fn chart_style_placeholder(
    reference: Option<Node>,
    resolver: &dyn ColorResolver,
    accent: Option<&str>,
    palette: Option<&[Option<String>]>,
    color_style_method: Option<&str>,
) -> Option<String> {
    match reference.map(|reference| {
        chart_style_reference_color(reference, resolver, accent, palette, color_style_method)
    }) {
        Some(Ok(Some(color))) => Some(color),
        Some(Err(())) => None,
        Some(Ok(None)) | None => accent.map(str::to_owned),
    }
}

/// MS-ODRAWXML §2.8.4.2 base-color selection. The linear brightness operation
/// is intentionally not performed because the specification does not define
/// its color space/range; this function only applies the normative index map
/// before CT_StyleColor and style-matrix transforms.
pub(super) fn chart_color_style_base_index(
    method: Option<&str>,
    index: usize,
    color_count: usize,
) -> Option<usize> {
    if color_count == 0 {
        return None;
    }
    match method {
        Some("withinLinear" | "withinLinearReversed") => Some(0),
        _ => Some(index % color_count),
    }
}

pub(super) fn chart_style_reference_index(reference: Option<Node>) -> Option<usize> {
    let style_color = reference.and_then(|reference| child(reference, "styleClr"))?;
    match style_color.attribute("val").unwrap_or("auto") {
        "auto" => None,
        value => Some(value.parse::<usize>().unwrap_or(0)),
    }
}

pub(super) fn parse_chart_image_fill(
    blip_fill: Node,
    color_resolver: &dyn ColorResolver,
    image_resolver: &dyn ChartImageResolver,
    source: ChartImageSource,
) -> Option<ChartStyleFill> {
    let blip = child(blip_fill, "blip")?;
    let mut duotone_count = 0usize;
    let has_unsupported_effect = blip.children().any(|node| {
        if !node.is_element() {
            return false;
        }
        match node.tag_name().name() {
            "alphaModFix" | "extLst" => false,
            "duotone" => {
                duotone_count += 1;
                duotone_count > 1
            }
            _ => true,
        }
    });
    if has_unsupported_effect {
        return None;
    }
    let alpha_effects_valid = blip
        .children()
        .filter(|node| node.is_element() && node.tag_name().name() == "alphaModFix")
        .all(|node| {
            node.attribute("amt").is_none_or(|value| {
                crate::units::drawingml_percentage_to_fraction(value)
                    .is_some_and(|fraction| fraction >= 0.0)
            })
        });
    if !alpha_effects_valid {
        return None;
    }
    let resolved_svg = crate::blip::svg_blip_rid(blip)
        .and_then(|relationship_id| image_resolver.resolve_image(source, &relationship_id));
    let resolved_raster = {
        crate::blip::blip_embed_rid(&blip)
            .and_then(|relationship_id| image_resolver.resolve_image(source, &relationship_id))
    };
    let (image_path, mime_type, svg_image_path) = match (resolved_raster, resolved_svg) {
        (Some((path, mime)), Some((svg_path, _))) => (path, mime, Some(svg_path)),
        (Some((path, mime)), None) => (path, mime, None),
        (None, Some((path, mime))) => (path, mime, None),
        (None, None) => return None,
    };
    let adapter = ColorResolverThemeAdapter(color_resolver);
    let src_rect = crate::blip::parse_src_rect(blip_fill);
    let alpha = crate::blip::parse_blip_alpha(blip_fill);
    // CT_Blip owns the effect sequence. A sibling <a:duotone> directly under
    // blipFill is not schema-valid and must not silently become a compatibility
    // fallback for chart markers.
    let duotone = child(blip, "duotone").and_then(|_| {
        crate::blip::parse_blip_duotone(blip_fill, &adapter, color_resolver.tint_mode())
    });
    if duotone_count == 1 && duotone.is_none() {
        return None;
    }
    let tile_node = child(blip_fill, "tile");
    let stretch_node = child(blip_fill, "stretch");
    // EG_FillModeProperties is optional and has no schema default. When absent
    // (or malformed with both choices), retain authored provenance outside this
    // recipe but do not invent stretch/tile semantics.
    if tile_node.is_some() == stretch_node.is_some() {
        return None;
    }
    let tile = tile_node.map(crate::fill::parse_tile);
    let stretch = stretch_node.is_some();
    let fill_rect = stretch_node.and_then(crate::fill::parse_fill_rect);
    let dpi = blip_fill
        .attribute("dpi")
        .and_then(|value| value.parse::<u32>().ok());
    let rot_with_shape = blip_fill
        .attribute("rotWithShape")
        .and_then(|value| match value {
            "1" | "true" => Some(true),
            "0" | "false" => Some(false),
            _ => None,
        });
    Some(ChartStyleFill::Image {
        image_path,
        mime_type,
        svg_image_path,
        dpi,
        rot_with_shape,
        src_rect,
        fill_rect,
        stretch,
        tile,
        alpha,
        duotone,
    })
}

pub(super) fn parse_chart_style_paint(
    container: Node,
    resolver: &dyn ColorResolver,
    placeholder: Option<&str>,
    image_resolver: &dyn ChartImageResolver,
    image_source: ChartImageSource,
) -> Option<ChartStylePaint> {
    let adapter = ColorResolverThemeAdapter(resolver);
    let style_resolver = crate::color::StyleMatrixColorResolver::new(&adapter, placeholder);
    if child(container, "noFill").is_some() {
        return Some(ChartStylePaint::NoFill);
    }
    if let Some(fill) = child(container, "solidFill") {
        return Some(
            resolve_chart_style_color(fill, resolver, placeholder)
                .map(|color| ChartStylePaint::Fill(Box::new(ChartStyleFill::Solid { color })))
                .unwrap_or(ChartStylePaint::Unresolved),
        );
    }
    if let Some(fill) = child(container, "gradFill") {
        return Some(
            crate::fill::parse_grad_fill(fill, &style_resolver, resolver.tint_mode())
                .map(|gradient| {
                    ChartStylePaint::Fill(Box::new(ChartStyleFill::Gradient {
                        stops: gradient.stops,
                        angle: gradient.angle,
                        grad_type: gradient.grad_type,
                        scaled: gradient.scaled,
                        path: gradient.path,
                        fill_to_rect: gradient.fill_to_rect,
                        tile_rect: gradient.tile_rect,
                        flip: gradient.flip,
                        rot_with_shape: gradient.rot_with_shape,
                    }))
                })
                .unwrap_or(ChartStylePaint::Unresolved),
        );
    }
    if let Some(fill) = child(container, "blipFill") {
        return Some(
            parse_chart_image_fill(fill, resolver, image_resolver, image_source)
                .map(|fill| ChartStylePaint::Fill(Box::new(fill)))
                .unwrap_or(ChartStylePaint::Unresolved),
        );
    }
    child(container, "pattFill").map(|fill| {
        let pattern = crate::fill::parse_patt_fill(fill, &style_resolver, resolver.tint_mode());
        ChartStylePaint::Fill(Box::new(ChartStyleFill::Pattern {
            fg: pattern.fg,
            bg: pattern.bg,
            preset: pattern.preset,
        }))
    })
}

pub(super) fn chart_style_fill_from_line_paint(
    paint: &crate::line::LinePaint,
) -> Option<ChartStyleFill> {
    match paint {
        crate::line::LinePaint::NoFill => None,
        crate::line::LinePaint::Solid { color } => {
            color.clone().map(|color| ChartStyleFill::Solid { color })
        }
        crate::line::LinePaint::Gradient(Some(gradient)) => Some(ChartStyleFill::Gradient {
            stops: gradient.stops.clone(),
            angle: gradient.angle,
            grad_type: gradient.grad_type.clone(),
            scaled: gradient.scaled,
            path: gradient.path.clone(),
            fill_to_rect: gradient.fill_to_rect.clone(),
            tile_rect: gradient.tile_rect.clone(),
            flip: gradient.flip.clone(),
            rot_with_shape: gradient.rot_with_shape,
        }),
        crate::line::LinePaint::Gradient(None) => None,
        crate::line::LinePaint::Pattern(pattern) => Some(ChartStyleFill::Pattern {
            fg: pattern.fg.clone(),
            bg: pattern.bg.clone(),
            preset: pattern.preset.clone(),
        }),
    }
}

/// Returns the structured-fill component count without resolving colors or
/// sorting gradient stops. Chart Style expands one authored recipe over every
/// Chart Colors entry, so this preflight must run before the palette loop.
pub(super) fn chart_style_paint_component_count(container: Node) -> Option<usize> {
    if child(container, "noFill").is_some() {
        return Some(0);
    }
    if child(container, "solidFill").is_some() {
        return Some(1);
    }
    if let Some(gradient) = child(container, "gradFill") {
        return Some(
            child(gradient, "gsLst")
                .map(|list| {
                    list.children()
                        .filter(|node| node.is_element() && node.tag_name().name() == "gs")
                        .count()
                })
                .unwrap_or(0),
        );
    }
    if child(container, "blipFill").is_some() || child(container, "pattFill").is_some() {
        return Some(1);
    }
    None
}

/// Parse one classic 3-D series shape only after its direct fill and outline
/// recipes fit the same per-recipe and aggregate availability ceilings as dPt
/// paints. This runs before gradient stop expansion; an oversized recipe makes
/// the enclosing chart fail atomically rather than retaining a paint prefix.
pub(super) fn parse_three_d_series_style_with_budget(
    series: Node,
    resolver: &dyn ColorResolver,
    image_resolver: &dyn ChartImageResolver,
    component_budget: &mut usize,
    budget_exceeded: &mut bool,
) -> Option<ChartExElementStyle> {
    let sp_pr = child(series, "spPr")?;
    let fill_components = chart_style_paint_component_count(sp_pr).unwrap_or(0);
    let line_components = child(sp_pr, "ln")
        .and_then(chart_style_paint_component_count)
        .unwrap_or(0);
    let components = fill_components.saturating_add(line_components);
    if fill_components > MAX_CHART_PAINT_RECIPE_COMPONENTS
        || line_components > MAX_CHART_PAINT_RECIPE_COMPONENTS
        || components > *component_budget
    {
        *budget_exceeded = true;
        return None;
    }
    *component_budget -= components;
    Some(parse_chartex_element_style(
        series,
        resolver,
        None,
        None,
        image_resolver,
        ChartImageSource::Chart,
    ))
}

pub(super) fn chart_style_paint_entry_limit(
    component_count: Option<usize>,
    palette_entries: usize,
    component_budget: usize,
) -> usize {
    let palette_entries = palette_entries.min(MAX_CHART_COLOR_STYLE_ENTRIES);
    match component_count {
        Some(components) if components > 0 => palette_entries.min(component_budget / components),
        _ => palette_entries,
    }
}

pub(super) enum ChartStyleMatrixRecipe {
    NoStyle,
    Missing,
    Xml(String),
}

pub(super) fn chart_style_fill_ref_xml(
    fill_ref: Node,
    resolver: &dyn ColorResolver,
) -> ChartStyleMatrixRecipe {
    use crate::theme::StyleMatrixLookup;

    let index = fill_ref
        .attribute("idx")
        .and_then(|value| value.parse::<usize>().ok())
        .unwrap_or(usize::MAX);
    if index == usize::MAX {
        return ChartStyleMatrixRecipe::Missing;
    }
    let Some(format_scheme) = resolver.theme_format_scheme() else {
        return if index == 0 {
            ChartStyleMatrixRecipe::NoStyle
        } else {
            ChartStyleMatrixRecipe::Missing
        };
    };
    let entry = match format_scheme.lookup_fill_ref(index) {
        StyleMatrixLookup::NoStyle => return ChartStyleMatrixRecipe::NoStyle,
        StyleMatrixLookup::Missing => return ChartStyleMatrixRecipe::Missing,
        StyleMatrixLookup::Entry(entry) => entry,
    };
    ChartStyleMatrixRecipe::Xml(entry.to_xml())
}

pub(super) fn parse_chart_style_line(
    line: Node,
    resolver: &dyn ColorResolver,
    placeholder: Option<&str>,
) -> crate::line::LineProperties {
    let adapter = ColorResolverThemeAdapter(resolver);
    let style_resolver = crate::color::StyleMatrixColorResolver::new(&adapter, placeholder);
    crate::line::parse_line_properties(line, &style_resolver, resolver.tint_mode())
}

pub(super) fn parse_chart_style_line_paint(
    line: Node,
    resolver: &dyn ColorResolver,
    placeholder: Option<&str>,
) -> Option<crate::line::LinePaint> {
    let adapter = ColorResolverThemeAdapter(resolver);
    let style_resolver = crate::color::StyleMatrixColorResolver::new(&adapter, placeholder);
    crate::line::parse_line_paint(line, &style_resolver, resolver.tint_mode())
}

pub(super) fn chart_style_line_ref_xml(
    line_ref: Node,
    resolver: &dyn ColorResolver,
) -> ChartStyleMatrixRecipe {
    use crate::theme::StyleMatrixLookup;

    let index = line_ref
        .attribute("idx")
        .and_then(|value| value.parse::<usize>().ok())
        .unwrap_or(usize::MAX);
    if index == usize::MAX {
        return ChartStyleMatrixRecipe::Missing;
    }
    let Some(format_scheme) = resolver.theme_format_scheme() else {
        return if index == 0 {
            ChartStyleMatrixRecipe::NoStyle
        } else {
            ChartStyleMatrixRecipe::Missing
        };
    };
    let entry = match format_scheme.lookup_line_ref(index) {
        StyleMatrixLookup::NoStyle => return ChartStyleMatrixRecipe::NoStyle,
        StyleMatrixLookup::Missing => return ChartStyleMatrixRecipe::Missing,
        StyleMatrixLookup::Entry(entry) => entry,
    };
    ChartStyleMatrixRecipe::Xml(entry.to_xml())
}

/// Resolution state for one effectRef. Unlike fill/line compatibility, a
/// broken concrete effect reference must stay distinct from omission: falling
/// back to a numeric style effect would invent paint the authored document did
/// not request.
pub(super) enum ChartStyleEffectRecipe {
    NoStyle,
    Missing,
    Xml(String),
}

pub(super) type ParsedChartStyleEffects = (
    Option<Vec<Option<crate::effect::Shadow>>>,
    Option<Vec<Option<crate::effect::Shadow>>>,
    Option<Vec<Option<crate::effect::Glow>>>,
    Option<Vec<Option<crate::effect::SoftEdge>>>,
    Option<Vec<Option<crate::effect::Reflection>>>,
    Option<bool>,
    Option<bool>,
    Option<bool>,
);

pub(super) fn chart_style_effect_ref_xml(
    effect_ref: Node,
    resolver: &dyn ColorResolver,
) -> ChartStyleEffectRecipe {
    use crate::theme::StyleMatrixLookup;

    let Some(index) = effect_ref
        .attribute("idx")
        .and_then(|value| value.parse::<usize>().ok())
    else {
        return ChartStyleEffectRecipe::Missing;
    };
    if index == 0 {
        return ChartStyleEffectRecipe::NoStyle;
    }
    let Some(format_scheme) = resolver.theme_format_scheme() else {
        return ChartStyleEffectRecipe::Missing;
    };
    match format_scheme.lookup_effect_ref(index) {
        StyleMatrixLookup::NoStyle => ChartStyleEffectRecipe::NoStyle,
        StyleMatrixLookup::Missing => ChartStyleEffectRecipe::Missing,
        StyleMatrixLookup::Entry(entry) => ChartStyleEffectRecipe::Xml(entry.to_xml()),
    }
}

pub(super) fn parse_chart_style_effects(
    style_node: Node,
    local_sp_pr: Option<Node>,
    resolver: &dyn ColorResolver,
    placeholders: &[Option<&str>],
    accents: Option<&[Option<String>]>,
    color_style_method: Option<&str>,
) -> ParsedChartStyleEffects {
    let direct_effect_list = local_sp_pr.and_then(|sp_pr| child(sp_pr, "effectLst"));
    let direct_effect_dag = local_sp_pr.and_then(|sp_pr| child(sp_pr, "effectDag"));
    let effect_ref = child(style_node, "effectRef");
    let effect_recipe = (direct_effect_list.is_none() && direct_effect_dag.is_none())
        .then(|| effect_ref.map(|reference| chart_style_effect_ref_xml(reference, resolver)))
        .flatten();
    let effect_recipe_xml = match effect_recipe.as_ref() {
        Some(ChartStyleEffectRecipe::Xml(xml)) => Some(xml),
        _ => None,
    };
    let recipe_doc = effect_recipe_xml.and_then(|xml| roxmltree::Document::parse(xml).ok());
    let (effect_authored, effect_no_style, mut effect_unsupported) =
        if direct_effect_list.is_some() || direct_effect_dag.is_some() {
            // CT_ShapeProperties carries one effect choice. Even an empty list
            // replaces the referenced component; effectDag is concrete but is
            // not representable by the current five-effect wire model.
            (Some(true), None, direct_effect_dag.map(|_| true))
        } else {
            match effect_recipe.as_ref() {
                Some(ChartStyleEffectRecipe::NoStyle) => (None, Some(true), None),
                Some(ChartStyleEffectRecipe::Missing) => (Some(true), None, Some(true)),
                Some(ChartStyleEffectRecipe::Xml(_)) if recipe_doc.is_some() => {
                    (Some(true), None, None)
                }
                Some(ChartStyleEffectRecipe::Xml(_)) => (Some(true), None, Some(true)),
                None => (None, None, None),
            }
        };

    if effect_authored.is_none() {
        return (None, None, None, None, None, None, effect_no_style, None);
    }

    let referenced_effect_style = recipe_doc.as_ref().and_then(|document| {
        let root = document.root_element();
        root.descendants()
            .find(|node| node.is_element() && node.tag_name().name() == "effectStyle")
    });
    let referenced_effect_choice = referenced_effect_style.and_then(|effect_style| {
        child(effect_style, "effectLst")
            .or_else(|| child(effect_style, "effectDag"))
            .map(|choice| (choice.tag_name().name().to_owned(), choice))
    });
    if recipe_doc.is_some() && referenced_effect_choice.is_none() {
        effect_unsupported = Some(true);
    }
    if referenced_effect_choice
        .as_ref()
        .is_some_and(|(name, _)| name == "effectDag")
    {
        effect_unsupported = Some(true);
    }
    if referenced_effect_style
        .is_some_and(|style| child(style, "scene3d").is_some() || child(style, "sp3d").is_some())
    {
        effect_unsupported = Some(true);
    }

    let mut shadows = Vec::with_capacity(placeholders.len());
    let mut inner_shadows = Vec::with_capacity(placeholders.len());
    let mut glows = Vec::with_capacity(placeholders.len());
    let mut soft_edges = Vec::with_capacity(placeholders.len());
    let mut reflections = Vec::with_capacity(placeholders.len());

    for accent in placeholders.iter().take(MAX_CHART_COLOR_STYLE_ENTRIES) {
        let placeholder = effect_ref.and_then(|reference| {
            chart_style_placeholder(
                Some(reference),
                resolver,
                *accent,
                accents,
                color_style_method,
            )
        });
        let adapter = ColorResolverThemeAdapter(resolver);
        let style_resolver =
            crate::color::StyleMatrixColorResolver::new(&adapter, placeholder.as_deref());
        let effect_list = direct_effect_list.or_else(|| {
            referenced_effect_choice
                .as_ref()
                .and_then(|(name, node)| (name == "effectLst").then_some(*node))
        });
        let parsed = effect_list.map(|node| {
            crate::effect::parse_effect_list(node, &style_resolver, resolver.tint_mode())
        });
        if parsed.as_ref().is_some_and(|effects| effects.unsupported) {
            effect_unsupported = Some(true);
        }
        shadows.push(parsed.as_ref().and_then(|effects| effects.shadow.clone()));
        inner_shadows.push(
            parsed
                .as_ref()
                .and_then(|effects| effects.inner_shadow.clone()),
        );
        glows.push(parsed.as_ref().and_then(|effects| effects.glow.clone()));
        soft_edges.push(
            parsed
                .as_ref()
                .and_then(|effects| effects.soft_edge.clone()),
        );
        reflections.push(
            parsed
                .as_ref()
                .and_then(|effects| effects.reflection.clone()),
        );
    }

    let shadows = shadows.iter().any(Option::is_some).then_some(shadows);
    let inner_shadows = inner_shadows
        .iter()
        .any(Option::is_some)
        .then_some(inner_shadows);
    let glows = glows.iter().any(Option::is_some).then_some(glows);
    let soft_edges = soft_edges.iter().any(Option::is_some).then_some(soft_edges);
    let reflections = reflections
        .iter()
        .any(Option::is_some)
        .then_some(reflections);
    (
        shadows,
        inner_shadows,
        glows,
        soft_edges,
        reflections,
        effect_authored,
        effect_no_style,
        effect_unsupported,
    )
}

/// Preserve only the direct DrawingML effect component from a classic chart
/// carrier such as `<c:marker>` or `<c:upBars>`. Their established fill/line
/// fields remain authoritative; this narrow adapter avoids reparsing those
/// paints or bypassing their existing gradient/image budgets.
pub(super) fn parse_direct_chart_effect_style(
    owner: Node,
    resolver: &dyn ColorResolver,
) -> Option<ChartExElementStyle> {
    let sp_pr = child(owner, "spPr")?;
    Some(parse_direct_chart_effect_style_from_sp_pr(
        owner, sp_pr, resolver,
    ))
}

pub(super) fn parse_direct_chart_effect_style_from_sp_pr(
    owner: Node,
    sp_pr: Node,
    resolver: &dyn ColorResolver,
) -> ChartExElementStyle {
    // Keep the direct shape as one generic carrier. The legacy scalar fields
    // used by established consumers remain in place, while newer/flat chart
    // roles (title, axisTitle, gridlines, etc.) can resolve the same bounded
    // fill, line, geometry and effects without adding one boolean or struct per
    // role. Relationship-backed picture fills require the caller's image
    // resolver and therefore remain on the dedicated image-aware paths.
    let fill = extract_direct_shape_fill(Some(sp_pr), resolver);
    let line = extract_direct_shape_line_from_sp_pr(Some(sp_pr), resolver);
    let placeholders = [None];
    let (
        shadows,
        inner_shadows,
        glows,
        soft_edges,
        reflections,
        effect_authored,
        effect_no_style,
        effect_unsupported,
    ) = parse_chart_style_effects(owner, Some(sp_pr), resolver, &placeholders, None, None);
    ChartExElementStyle {
        shape_properties_present: Some(true),
        fill_paints: fill.fill.clone().map(|paint| vec![Some(paint)]),
        fill_colors: fill.color.clone().map(|color| vec![Some(color)]),
        fill_hidden: fill.hidden,
        fill_paint_authored: fill.paint_authored,
        line_paints: line.fill.clone().map(|paint| vec![Some(paint)]),
        line_colors: line.color.clone().map(|color| vec![Some(color)]),
        line_paint_authored: line.paint_authored,
        line_width_emu: line.width_emu,
        line_dash: line.dash,
        line_dash_authored: line.dash_authored,
        line_custom_dash: line.custom_dash,
        line_cap: line.cap,
        line_join: line.join,
        line_compound: line.compound,
        line_hidden: line.hidden,
        shadows,
        inner_shadows,
        glows,
        soft_edges,
        reflections,
        effect_authored,
        effect_no_style,
        effect_unsupported,
        ..ChartExElementStyle::default()
    }
}

/// Width inherited by a classic Style 2 axis overlay from the first theme
/// line-style entry. Classic chart axes do not carry an explicit `lnRef`, but
/// Office's Style 2 recipe applies `lnStyleLst[0]` before overlaying the local
/// `<c:*Ax><c:spPr><a:ln>` properties. Consequently a local line that authors
/// only its color retains the theme width; an absent local line remains absent.
/// The Office vector boundary set establishes this rule for primary category,
/// date and value axes, while secondary axes and other numbered built-in styles
/// remain unresolved.
pub(super) fn classic_style_two_axis_line_width_emu(resolver: &dyn ColorResolver) -> Option<u32> {
    use crate::theme::StyleMatrixLookup;

    let entry = match resolver.theme_format_scheme()?.lookup_line_ref(1) {
        StyleMatrixLookup::Entry(entry) => entry,
        StyleMatrixLookup::NoStyle | StyleMatrixLookup::Missing => return None,
    };
    let xml = entry.to_xml();
    let document = crate::depth::parse_guarded(&xml).ok()?;
    let line = document
        .descendants()
        .find(|node| node.is_element() && node.tag_name().name() == "ln")?;
    parse_chart_style_line(line, resolver, None)
        .width
        .and_then(|width| u32::try_from(width).ok())
}

/// Whether a shape property block authors any DrawingML fill choice.  This is
/// deliberately broader than the structured fills currently representable by
/// `ChartStylePaint`: an unsupported local fill still overrides `fillRef` and
/// must fail closed instead of reviving an inherited theme fill.
pub(super) fn shape_has_fill_choice(sp_pr: Node) -> bool {
    [
        "noFill",
        "solidFill",
        "gradFill",
        "pattFill",
        "blipFill",
        "grpFill",
    ]
    .iter()
    .any(|name| child(sp_pr, name).is_some())
}

pub(super) fn parse_chartex_element_style(
    style_node: Node,
    resolver: &dyn ColorResolver,
    accents: Option<&[Option<String>]>,
    color_style_method: Option<&str>,
    image_resolver: &dyn ChartImageResolver,
    image_source: ChartImageSource,
) -> ChartExElementStyle {
    use crate::line::{LineDash, LineJoin, LinePaint, LineProperties};

    let placeholders: Vec<Option<&str>> = accents
        .map(|values| {
            (0..values.len())
                .map(|index| {
                    chart_color_style_base_index(color_style_method, index, values.len())
                        .and_then(|mapped| values.get(mapped))
                        .and_then(Option::as_deref)
                })
                .collect()
        })
        .unwrap_or_else(|| vec![None]);
    let fill_ref = child(style_node, "fillRef");
    let line_ref = child(style_node, "lnRef");
    // Serialize and parse each referenced theme recipe once per style role.
    // Placeholder substitution and color transforms are then the only work in
    // the palette loop (rather than reparsing a DOM for every palette entry).
    let fill_recipe = fill_ref.map(|reference| chart_style_fill_ref_xml(reference, resolver));
    let fill_recipe_xml = fill_recipe.as_ref().and_then(|recipe| match recipe {
        ChartStyleMatrixRecipe::Xml(xml) => Some(xml.as_str()),
        _ => None,
    });
    let fill_recipe_doc = fill_recipe_xml.and_then(|xml| roxmltree::Document::parse(xml).ok());
    let local_sp_pr = child(style_node, "spPr");
    let style_modifiers = style_node.attribute("mods").unwrap_or_default();
    let has_modifier = |name: &str| {
        style_modifiers
            .split_ascii_whitespace()
            .any(|modifier| modifier == name)
            .then_some(true)
    };
    let local_fill_authored = local_sp_pr.is_some_and(shape_has_fill_choice);
    let fill_paint_authored = (local_fill_authored
        || matches!(
            fill_recipe.as_ref(),
            Some(ChartStyleMatrixRecipe::Xml(_) | ChartStyleMatrixRecipe::Missing)
        ))
    .then_some(true);
    let fill_no_style = (matches!(fill_recipe.as_ref(), Some(ChartStyleMatrixRecipe::NoStyle))
        && !local_fill_authored)
        .then_some(true);
    let line_recipe = line_ref.map(|reference| chart_style_line_ref_xml(reference, resolver));
    let line_recipe_xml = line_recipe.as_ref().and_then(|recipe| match recipe {
        ChartStyleMatrixRecipe::Xml(xml) => Some(xml.as_str()),
        _ => None,
    });
    let line_recipe_doc = line_recipe_xml.and_then(|xml| roxmltree::Document::parse(xml).ok());
    let fill_component_count = if local_fill_authored {
        local_sp_pr
            .and_then(chart_style_paint_component_count)
            .or(Some(0))
    } else {
        match fill_recipe.as_ref() {
            Some(ChartStyleMatrixRecipe::NoStyle | ChartStyleMatrixRecipe::Missing) => Some(0),
            Some(ChartStyleMatrixRecipe::Xml(_)) => fill_recipe_doc
                .as_ref()
                .and_then(|document| chart_style_paint_component_count(document.root_element())),
            None => None,
        }
    };
    let fill_entry_count = placeholders.len().min(MAX_CHART_COLOR_STYLE_ENTRIES);
    let parsed_fill_entries = chart_style_paint_entry_limit(
        fill_component_count,
        fill_entry_count,
        MAX_CHART_STYLE_PAINT_COMPONENTS,
    );
    let mut fills = Vec::with_capacity(fill_entry_count);
    for (index, accent) in placeholders
        .iter()
        .take(MAX_CHART_COLOR_STYLE_ENTRIES)
        .enumerate()
    {
        if index >= parsed_fill_entries {
            fills.push(None);
            continue;
        }
        let placeholder =
            chart_style_placeholder(fill_ref, resolver, *accent, accents, color_style_method);
        let local_paint = local_sp_pr.and_then(|sp_pr| {
            parse_chart_style_paint(
                sp_pr,
                resolver,
                placeholder.as_deref(),
                image_resolver,
                image_source,
            )
        });
        let paint = if local_fill_authored {
            local_paint
        } else {
            local_paint.or_else(|| match fill_recipe.as_ref() {
                Some(ChartStyleMatrixRecipe::NoStyle) => Some(ChartStylePaint::NoFill),
                Some(ChartStyleMatrixRecipe::Missing) => Some(ChartStylePaint::Unresolved),
                Some(ChartStyleMatrixRecipe::Xml(_)) => {
                    fill_recipe_doc.as_ref().and_then(|document| {
                        parse_chart_style_paint(
                            document.root_element(),
                            resolver,
                            placeholder.as_deref(),
                            image_resolver,
                            ChartImageSource::Theme,
                        )
                    })
                }
                None => None,
            })
        };
        fills.push(paint);
    }
    let fill_hidden = (fills.iter().any(Option::is_some)
        && fills
            .iter()
            .all(|paint| matches!(paint, Some(ChartStylePaint::NoFill))))
    .then_some(true);
    let fill_paints = (!fill_hidden.unwrap_or(false))
        .then(|| {
            fills
                .iter()
                .map(|paint| match paint {
                    Some(ChartStylePaint::Fill(fill)) => Some((**fill).clone()),
                    _ => None,
                })
                .collect::<Vec<_>>()
        })
        .filter(|paints| paints.iter().any(Option::is_some));
    let fill_colors = (!fill_hidden.unwrap_or(false))
        .then(|| {
            fills
                .iter()
                .map(|paint| match paint {
                    Some(ChartStylePaint::Fill(fill)) => match fill.as_ref() {
                        ChartStyleFill::Solid { color } => Some(color.clone()),
                        _ => None,
                    },
                    _ => None,
                })
                .collect::<Vec<_>>()
        })
        .filter(|colors| colors.iter().any(Option::is_some));

    let local_line = local_sp_pr.and_then(|sp_pr| child(sp_pr, "ln"));
    let inherited_line = line_recipe_doc
        .as_ref()
        .and_then(|document| child(document.root_element(), "ln"));
    let local_line_fill_authored = local_line.is_some_and(shape_has_fill_choice);
    // lnRef idx=0 is paint fall-through, not an instruction to discard a
    // locally authored width/dash/cap/join. Preserve the sentinel whenever
    // local a:ln contributes geometry but no paint; the effective cascade can
    // then take numeric paint and overlay this linked geometry.
    let line_no_style = (matches!(line_recipe.as_ref(), Some(ChartStyleMatrixRecipe::NoStyle))
        && !local_line_fill_authored)
        .then_some(true);
    let line_paint_authored = (local_line_fill_authored
        || inherited_line.is_some_and(shape_has_fill_choice)
        || matches!(line_recipe.as_ref(), Some(ChartStyleMatrixRecipe::Missing)))
    .then_some(true);
    let line_component_count = if local_line_fill_authored {
        local_line
            .and_then(chart_style_paint_component_count)
            .or(Some(0))
    } else {
        match line_recipe.as_ref() {
            Some(ChartStyleMatrixRecipe::NoStyle | ChartStyleMatrixRecipe::Missing) => Some(0),
            Some(ChartStyleMatrixRecipe::Xml(_)) => inherited_line
                .and_then(chart_style_paint_component_count)
                .or(Some(0)),
            None => None,
        }
    };
    let line_entry_count = placeholders.len().min(MAX_CHART_COLOR_STYLE_ENTRIES);
    let parsed_line_entries = chart_style_paint_entry_limit(
        line_component_count,
        line_entry_count,
        MAX_CHART_STYLE_PAINT_COMPONENTS,
    );

    // Width/dash/cap/join/compound are invariant across Chart Colors palette
    // entries. Parse that geometry once; only line paint needs placeholder
    // substitution for every color. This prevents an authored unbounded-style
    // custom dash from becoming palette-size × dash-size work.
    let first_placeholder = placeholders.first().and_then(|accent| {
        chart_style_placeholder(line_ref, resolver, *accent, accents, color_style_method)
    });
    let inherited_geometry = match line_recipe.as_ref() {
        Some(ChartStyleMatrixRecipe::NoStyle) => Some(LineProperties {
            paint: Some(LinePaint::NoFill),
            ..LineProperties::default()
        }),
        Some(ChartStyleMatrixRecipe::Missing) => None,
        Some(ChartStyleMatrixRecipe::Xml(_)) => inherited_line
            .map(|line| parse_chart_style_line(line, resolver, first_placeholder.as_deref())),
        None => None,
    };
    let local_geometry =
        local_line.map(|line| parse_chart_style_line(line, resolver, first_placeholder.as_deref()));
    let first_line = match (local_geometry, inherited_geometry) {
        (Some(local), Some(inherited)) => Some(local.with_fallback(&inherited)),
        (Some(local), None) => Some(local),
        (None, inherited) => inherited,
    };

    let resolved_line_paints = placeholders
        .iter()
        .take(MAX_CHART_COLOR_STYLE_ENTRIES)
        .enumerate()
        .map(|(index, accent)| {
            if index >= parsed_line_entries {
                return None;
            }
            let placeholder =
                chart_style_placeholder(line_ref, resolver, *accent, accents, color_style_method);
            let inherited = match line_recipe.as_ref() {
                Some(ChartStyleMatrixRecipe::NoStyle) => Some(LinePaint::NoFill),
                Some(ChartStyleMatrixRecipe::Missing) => None,
                Some(ChartStyleMatrixRecipe::Xml(_)) => inherited_line.and_then(|line| {
                    parse_chart_style_line_paint(line, resolver, placeholder.as_deref())
                }),
                None => None,
            };
            let local = local_line.and_then(|line| {
                parse_chart_style_line_paint(line, resolver, placeholder.as_deref())
            });
            if local_line_fill_authored {
                local
            } else {
                local.or(inherited)
            }
        })
        .collect::<Vec<_>>();
    let line_hidden = (resolved_line_paints.iter().any(Option::is_some)
        && resolved_line_paints
            .iter()
            .all(|paint| matches!(paint, Some(LinePaint::NoFill))))
    .then_some(true);
    let line_colors = (!line_hidden.unwrap_or(false))
        .then(|| {
            resolved_line_paints
                .iter()
                .map(|paint| match paint {
                    Some(LinePaint::Solid { color }) => color.clone(),
                    _ => None,
                })
                .collect::<Vec<_>>()
        })
        .filter(|colors| colors.iter().any(Option::is_some));
    let line_paints = (!line_hidden.unwrap_or(false))
        .then(|| {
            resolved_line_paints
                .iter()
                .map(|paint| paint.as_ref().and_then(chart_style_fill_from_line_paint))
                .collect::<Vec<_>>()
        })
        .filter(|paints| paints.iter().any(Option::is_some));
    let line_width_emu = first_line
        .as_ref()
        .and_then(|line| line.width)
        .and_then(|width| u32::try_from(width).ok());
    let line_dash = first_line
        .as_ref()
        .and_then(|line| match line.dash.as_ref() {
            Some(LineDash::Preset(value)) => value.clone(),
            _ => None,
        });
    let line_custom_dash = first_line
        .as_ref()
        .and_then(|line| match line.dash.as_ref() {
            Some(LineDash::Custom(stops)) => Some(
                stops
                    .iter()
                    .map(|stop| ChartLineDashSegment {
                        dash: stop.dash / 100_000.0,
                        space: stop.space / 100_000.0,
                    })
                    .collect(),
            ),
            _ => None,
        });
    let line_join = first_line
        .as_ref()
        .and_then(|line| match line.join.as_ref() {
            Some(LineJoin::Round) => Some("round".to_owned()),
            Some(LineJoin::Bevel) => Some("bevel".to_owned()),
            Some(LineJoin::Miter { .. }) => Some("miter".to_owned()),
            None => None,
        });
    let (font_size_hpt, font_bold, mut font_color, font_face) =
        extract_chartex_style_text_props(Some(style_node), resolver);
    let direct_font_paint = chart_text_paint([child(style_node, "defRPr")], resolver);
    let font_ref_paint_authored = child(style_node, "fontRef").is_some_and(|font_ref| {
        font_ref.children().any(|node| {
            node.is_element()
                && matches!(
                    node.tag_name().name(),
                    "styleClr"
                        | "srgbClr"
                        | "schemeClr"
                        | "sysClr"
                        | "prstClr"
                        | "scrgbClr"
                        | "hslClr"
                )
        })
    });
    let font_ref = child(style_node, "fontRef");
    let font_colors = (!direct_font_paint.authored && font_ref_paint_authored).then(|| {
        placeholders
            .iter()
            .take(MAX_CHART_COLOR_STYLE_ENTRIES)
            .map(|accent| {
                font_ref.and_then(|reference| {
                    chart_style_reference_color(
                        reference,
                        resolver,
                        *accent,
                        accents,
                        color_style_method,
                    )
                    .ok()
                    .flatten()
                })
            })
            .collect::<Vec<_>>()
    });
    if !direct_font_paint.authored {
        font_color = font_colors
            .as_ref()
            .and_then(|colors| colors.first())
            .and_then(Clone::clone)
            .or(font_color);
    }
    let font_paint_authored =
        (direct_font_paint.authored || font_ref_paint_authored || font_color.is_some())
            .then_some(true);
    let font_hidden = direct_font_paint.hidden.then_some(true);
    let font_italic = extract_chartex_style_text_italic(Some(style_node));
    let (
        font_language,
        font_baseline,
        text_rotation,
        text_wrap,
        text_vertical_anchor,
        text_vertical_mode,
        text_l_ins_emu,
        text_t_ins_emu,
        text_r_ins_emu,
        text_b_ins_emu,
    ) = extract_chartex_style_text_body(style_node);
    let text_body_authored = child(style_node, "bodyPr").map(|_| true);
    let (
        shadows,
        inner_shadows,
        glows,
        soft_edges,
        reflections,
        effect_authored,
        effect_no_style,
        effect_unsupported,
    ) = parse_chart_style_effects(
        style_node,
        local_sp_pr,
        resolver,
        &placeholders,
        accents,
        color_style_method,
    );

    ChartExElementStyle {
        shape_properties_present: local_sp_pr.map(|_| true),
        allow_no_fill_override: has_modifier("allowNoFillOverride"),
        allow_no_line_override: has_modifier("allowNoLineOverride"),
        font_size_hpt,
        font_bold,
        font_italic,
        font_color,
        font_colors,
        font_color_index: chart_style_reference_index(font_ref),
        font_formatting_indices: None,
        font_paint_authored,
        font_hidden,
        font_face,
        font_language,
        font_baseline,
        text_rotation,
        text_wrap,
        text_vertical_anchor,
        text_vertical_mode,
        text_l_ins_emu,
        text_t_ins_emu,
        text_r_ins_emu,
        text_b_ins_emu,
        text_body_authored,
        fill_paints,
        fill_colors,
        fill_hidden,
        fill_paint_authored,
        fill_no_style,
        line_colors,
        line_paints,
        line_paint_authored,
        line_width_emu,
        line_hidden,
        line_no_style,
        line_dash,
        line_dash_authored: first_line
            .as_ref()
            .and_then(|line| line.dash.as_ref())
            .map(|_| true),
        line_custom_dash,
        line_cap: first_line.as_ref().and_then(|line| line.cap.clone()),
        line_join,
        line_compound: first_line.as_ref().and_then(|line| line.compound.clone()),
        shadows,
        inner_shadows,
        glows,
        soft_edges,
        reflections,
        effect_authored,
        effect_no_style,
        effect_unsupported,
        fill_color_index: chart_style_reference_index(fill_ref),
        fill_formatting_indices: None,
        fill_semantic_fallback_indices: None,
        line_color_index: chart_style_reference_index(line_ref),
        line_formatting_indices: None,
        line_semantic_fallback_indices: None,
        effect_formatting_indices: None,
        effect_color_index: chart_style_reference_index(child(style_node, "effectRef")),
    }
}

/// Resolve the total color set defined by a linked Chart Colors part. Per
/// MS-ODRAWXML §2.8.3.2, every contained color is repeated for every
/// `<cs:variation>`, with the variation transforms appended to the base color.
pub(super) fn parse_chart_color_style(
    xml: &str,
    resolver: &dyn ColorResolver,
) -> Option<(String, Vec<Option<String>>)> {
    let document = crate::depth::parse_guarded(xml).ok()?;
    let root = document.root_element();
    let method = root.attribute("meth").unwrap_or("cycle").to_owned();
    let adapter = ColorResolverThemeAdapter(resolver);
    let colors = root
        .children()
        .filter(|node| {
            node.is_element()
                && matches!(
                    node.tag_name().name(),
                    "srgbClr" | "schemeClr" | "sysClr" | "prstClr" | "scrgbClr" | "hslClr"
                )
        })
        .map(|node| {
            crate::color::color_source_from_element(node).and_then(|source| {
                crate::color::resolve_color_source(source, &adapter, resolver.tint_mode())
            })
        })
        .take(MAX_CHART_COLOR_STYLE_ENTRIES + 1)
        .collect::<Vec<_>>();
    if colors.is_empty()
        || colors.len() > MAX_CHART_COLOR_STYLE_ENTRIES
        || colors.iter().all(Option::is_none)
    {
        return None;
    }
    let variations = root
        .children()
        .filter(|node| node.is_element() && node.tag_name().name() == "variation")
        .take(MAX_CHART_COLOR_STYLE_ENTRIES + 1)
        .collect::<Vec<_>>();
    if variations.len() > MAX_CHART_COLOR_STYLE_ENTRIES {
        return None;
    }
    if variations.is_empty() {
        return Some((method, colors));
    }
    if variations
        .len()
        .checked_mul(colors.len())
        .is_none_or(|entries| entries > MAX_CHART_COLOR_STYLE_ENTRIES)
    {
        return None;
    }
    let palette = variations
        .iter()
        .flat_map(|variation| {
            colors.iter().map(move |color| {
                color.as_deref().map(|color| {
                    crate::color::apply_color_transforms(color, *variation, resolver.tint_mode())
                })
            })
        })
        .collect::<Vec<_>>();
    Some((method, palette))
}

/// Paint-bearing CT_ChartStyle children in schema order (MS-ODRAWXML
/// §2.8.3.1). `dataPointMarkerLayout` and `extLst` have different grammars and
/// remain outside the style-entry table.
pub(super) const CHART_STYLE_ROLE_NAMES: [&str; 30] = [
    "axisTitle",
    "categoryAxis",
    "chartArea",
    "dataLabel",
    "dataLabelCallout",
    "dataPoint",
    "dataPoint3D",
    "dataPointLine",
    "dataPointMarker",
    "dataPointWireframe",
    "dataTable",
    "downBar",
    "dropLine",
    "errorBar",
    "floor",
    "gridlineMajor",
    "gridlineMinor",
    "hiLoLine",
    "leaderLine",
    "legend",
    "plotArea",
    "plotArea3D",
    "seriesAxis",
    "seriesLine",
    "title",
    "trendline",
    "trendlineLabel",
    "upBar",
    "valueAxis",
    "wall",
];

/// Number of structured-fill components that one Chart Style role expands for
/// each Chart Colors entry. A local `spPr` fill overrides `fillRef`; otherwise
/// the referenced theme recipe contributes its gradient stops/pattern/solid
/// component. This intentionally runs before any role×palette expansion so a
/// theme gradient cannot bypass the aggregate resource budget merely because
/// it lives outside `styleN.xml`.
pub(super) fn chart_style_role_fill_component_count(
    role: Node,
    resolver: &dyn ColorResolver,
) -> Option<usize> {
    if let Some(sp_pr) = child(role, "spPr") {
        if shape_has_fill_choice(sp_pr) {
            return Some(chart_style_paint_component_count(sp_pr).unwrap_or(0));
        }
    }
    let Some(fill_ref) = child(role, "fillRef") else {
        return Some(0);
    };
    match chart_style_fill_ref_xml(fill_ref, resolver) {
        ChartStyleMatrixRecipe::Xml(xml) => {
            let document = crate::depth::parse_guarded(&xml).ok()?;
            Some(chart_style_paint_component_count(document.root_element()).unwrap_or(0))
        }
        ChartStyleMatrixRecipe::NoStyle | ChartStyleMatrixRecipe::Missing => Some(0),
    }
}

pub(super) fn chart_style_role_line_component_count(
    role: Node,
    resolver: &dyn ColorResolver,
) -> Option<usize> {
    if let Some(line) = child(role, "spPr").and_then(|sp_pr| child(sp_pr, "ln")) {
        if shape_has_fill_choice(line) {
            return Some(chart_style_paint_component_count(line).unwrap_or(0));
        }
    }
    let Some(line_ref) = child(role, "lnRef") else {
        return Some(0);
    };
    match chart_style_line_ref_xml(line_ref, resolver) {
        ChartStyleMatrixRecipe::Xml(xml) => {
            let document = crate::depth::parse_guarded(&xml).ok()?;
            let root = document.root_element();
            let line =
                child(root, "ln").or_else(|| (root.tag_name().name() == "ln").then_some(root));
            Some(
                line.and_then(chart_style_paint_component_count)
                    .unwrap_or(0),
            )
        }
        ChartStyleMatrixRecipe::NoStyle | ChartStyleMatrixRecipe::Missing => Some(0),
    }
}

pub(super) fn parse_chart_style_role_table(
    style_root: Node,
    resolver: &dyn ColorResolver,
    palette: Option<&[Option<String>]>,
    color_style_method: Option<&str>,
    image_resolver: &dyn ChartImageResolver,
) -> Option<BTreeMap<String, ChartExElementStyle>> {
    let role_nodes = style_root
        .children()
        .filter(|node| {
            node.is_element() && CHART_STYLE_ROLE_NAMES.contains(&node.tag_name().name())
        })
        .collect::<Vec<_>>();
    if role_nodes.is_empty() {
        return None;
    }

    let palette_entries = palette.map_or(1, |colors| colors.len().max(1));
    let total_slots = role_nodes.len().checked_mul(palette_entries)?;
    if total_slots > MAX_CHART_STYLE_ROLE_SLOTS {
        return None;
    }

    // Preflight both local and referenced-theme structured fills before any
    // role×palette expansion. The per-role guard remains a second line of
    // defence; this aggregate guard bounds all roles together.
    let fill_components = role_nodes.iter().try_fold(0usize, |total, role| {
        let fill_components = chart_style_role_fill_component_count(*role, resolver)?;
        let line_components = chart_style_role_line_component_count(*role, resolver)?;
        // Palette expansion happens after this point. Refuse an oversized
        // source recipe before parsing it even when the aggregate expanded
        // table would still fit its larger chart-wide ceiling.
        if fill_components > MAX_CHART_PAINT_RECIPE_COMPONENTS
            || line_components > MAX_CHART_PAINT_RECIPE_COMPONENTS
        {
            return None;
        }
        let components = fill_components.checked_add(line_components)?;
        total.checked_add(components.checked_mul(palette_entries)?)
    })?;
    if fill_components > MAX_CHART_STYLE_PAINT_COMPONENTS {
        return None;
    }

    Some(
        role_nodes
            .into_iter()
            .map(|role| {
                let name = role.tag_name().name().to_owned();
                let style = parse_chartex_element_style(
                    role,
                    resolver,
                    palette,
                    color_style_method,
                    image_resolver,
                    ChartImageSource::Style,
                );
                (name, style)
            })
            .collect(),
    )
}

/// Preserve component ownership when a linked style part is structurally
/// valid but exceeds a renderer resource budget. Treating that part as absent
/// would let the lower numeric style repaint an explicitly authored component.
/// The sentinel is intentionally cheap: it records only source-level ownership
/// and lets omitted components continue to inherit normally.
pub(super) fn unresolved_chart_style_role_table(
    style_root: Node,
    resolver: &dyn ColorResolver,
    image_resolver: &dyn ChartImageResolver,
) -> Option<BTreeMap<String, ChartExElementStyle>> {
    let roles = style_root
        .children()
        .filter(|node| {
            node.is_element() && CHART_STYLE_ROLE_NAMES.contains(&node.tag_name().name())
        })
        .map(|role| {
            let sp_pr = child(role, "spPr");
            let local_fill = sp_pr.is_some_and(shape_has_fill_choice);
            let fill_ref = child(role, "fillRef");
            let fill_index = fill_ref
                .and_then(|reference| reference.attribute("idx"))
                .and_then(|value| value.parse::<usize>().ok());
            let local_line = sp_pr.and_then(|shape| child(shape, "ln"));
            let local_line_paint = local_line.is_some_and(shape_has_fill_choice);
            let line_ref = child(role, "lnRef");
            let line_index = line_ref
                .and_then(|reference| reference.attribute("idx"))
                .and_then(|value| value.parse::<usize>().ok());
            let direct_font = child(role, "defRPr");
            let font_ref = child(role, "fontRef");
            let effect_ref = child(role, "effectRef");
            let local_effect = sp_pr.is_some_and(|shape| {
                child(shape, "effectLst").is_some() || child(shape, "effectDag").is_some()
            });
            // Parse one bounded representative entry so cheap typography,
            // line geometry and explicit indices survive a palette/aggregate
            // rejection. Paint ownership below is then restored even when the
            // representative recipe itself exceeded its per-recipe budget.
            let mut style = parse_chartex_element_style(
                role,
                resolver,
                None,
                None,
                image_resolver,
                ChartImageSource::Style,
            );
            if local_fill || fill_ref.is_some_and(|_| fill_index != Some(0)) {
                style.fill_paint_authored = Some(true);
            } else if fill_index == Some(0) {
                style.fill_no_style = Some(true);
            }
            if local_line_paint || line_ref.is_some_and(|_| line_index != Some(0)) {
                style.line_paint_authored = Some(true);
            } else if line_index == Some(0) {
                style.line_no_style = Some(true);
            }
            if direct_font.is_some_and(shape_has_fill_choice) || font_ref.is_some() {
                style.font_paint_authored = Some(true);
            }
            if local_effect
                || effect_ref.is_some_and(|reference| {
                    reference
                        .attribute("idx")
                        .and_then(|value| value.parse::<usize>().ok())
                        != Some(0)
                })
            {
                style.effect_authored = Some(true);
                style.effect_unsupported = Some(true);
            } else if effect_ref.is_some() {
                style.effect_no_style = Some(true);
            }
            (role.tag_name().name().to_owned(), style)
        })
        .collect::<BTreeMap<_, _>>();
    (!roles.is_empty()).then_some(roles)
}

/// An authored Chart Style relationship whose target cannot be read or parsed
/// is not equivalent to an absent relationship. Its component ownership is
/// unknowable, so every standardized role fails closed instead of allowing the
/// less-specific numeric style to repaint it.
pub(super) fn unreadable_chart_style_role_table() -> BTreeMap<String, ChartExElementStyle> {
    CHART_STYLE_ROLE_NAMES
        .iter()
        .map(|role| {
            (
                (*role).to_owned(),
                ChartExElementStyle {
                    font_paint_authored: Some(true),
                    font_hidden: Some(true),
                    fill_paint_authored: Some(true),
                    fill_hidden: Some(true),
                    line_paint_authored: Some(true),
                    line_hidden: Some(true),
                    effect_authored: Some(true),
                    effect_unsupported: Some(true),
                    ..Default::default()
                },
            )
        })
        .collect()
}
