use super::*;

pub(super) fn extract_chartex_style_text_props(
    style_node: Option<Node>,
    resolver: &dyn ColorResolver,
) -> (Option<i32>, Option<bool>, Option<String>, Option<String>) {
    let Some(style_node) = style_node else {
        return (None, None, None, None);
    };
    let def_r_pr = child(style_node, "defRPr");
    let font_ref = child(style_node, "fontRef");
    let size = def_r_pr
        .and_then(|props| props.attribute("sz"))
        .and_then(parse_text_font_size_hpt);
    let bold = chart_text_bool_from_present_props(def_r_pr, "b");
    // A directly-authored DrawingML fill choice owns the text paint even when
    // that choice is noFill or a paint kind the current wire model cannot
    // resolve. Falling through to fontRef in that case would let inherited
    // formatting revive text that the direct formatting intentionally hides.
    let direct_paint_authored = def_r_pr.is_some_and(shape_has_fill_choice);
    let direct_color = def_r_pr
        .and_then(|props| child(props, "solidFill"))
        .and_then(|fill| resolver.resolve_solid_fill(fill));
    let inherited_color = (!direct_paint_authored)
        .then_some(font_ref)
        .flatten()
        .and_then(|font| {
            font.children()
                .find(|node| {
                    node.is_element()
                        && matches!(
                            node.tag_name().name(),
                            "srgbClr" | "schemeClr" | "sysClr" | "prstClr" | "scrgbClr" | "hslClr"
                        )
                })
                .and_then(|color_node| {
                    crate::color::color_source_from_element(color_node).and_then(|source| {
                        crate::color::resolve_color_source(
                            source,
                            &ColorResolverThemeAdapter(resolver),
                            resolver.tint_mode(),
                        )
                    })
                })
        });
    let color = direct_color.or(inherited_color);
    let face = def_r_pr.and_then(first_latin_typeface).or_else(|| {
        match font_ref.and_then(|font| font.attribute("idx")) {
            Some("major") => resolver.theme_major_font_latin(),
            Some("minor") => resolver.theme_minor_font_latin(),
            _ => None,
        }
    });
    (size, bold, color, face)
}

pub(super) fn extract_chartex_style_text_italic(style_node: Option<Node>) -> Option<bool> {
    chart_text_bool_from_present_props(child(style_node?, "defRPr"), "i")
}

pub(super) fn extract_chart_space_text_style(
    chart_space: Node,
    resolver: &dyn ColorResolver,
) -> Option<ChartExElementStyle> {
    let tx_pr = child(chart_space, "txPr")?;
    let paint = chart_text_body_paint(Some(tx_pr), resolver);
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
    ) = extract_chartex_style_text_body(tx_pr);
    Some(ChartExElementStyle {
        font_size_hpt: extract_axis_tick_label_size(chart_space),
        font_bold: extract_axis_tick_label_bold(chart_space),
        font_italic: extract_axis_tick_label_italic(chart_space),
        font_color: paint.color,
        font_paint_authored: paint.authored.then_some(true),
        font_hidden: paint.hidden.then_some(true),
        font_face: extract_axis_tick_label_face(chart_space),
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
        text_body_authored: child(tx_pr, "bodyPr").map(|_| true),
        ..Default::default()
    })
}

/// Normalize DrawingML `ST_Percentage` to a fraction. Strict packages use a
/// percent lexical form (`30%`); Transitional packages may use an integer
/// (`30000`, thousandths of a percent).
pub(super) fn parse_chart_text_percentage(value: &str) -> Option<f64> {
    if let Some(percent) = value.strip_suffix('%') {
        if percent.is_empty() || percent.starts_with('+') || percent.matches('.').count() > 1 {
            return None;
        }
        let digits = percent.strip_prefix('-').unwrap_or(percent);
        if digits.is_empty()
            || !digits.chars().all(|ch| ch.is_ascii_digit() || ch == '.')
            || digits.starts_with('.')
            || digits.ends_with('.')
        {
            return None;
        }
        let parsed = percent.parse::<f64>().ok()?;
        return parsed.is_finite().then_some(parsed / 100.0);
    }
    value
        .parse::<i32>()
        .ok()
        .map(|raw| f64::from(raw) / 100_000.0)
}

/// Parse the chart-specific percentage unions used by `CT_HPercent`,
/// `CT_DepthPercent`, `CT_GapAmount`, and `CT_Thickness`. Unlike DrawingML
/// `ST_Percentage`, the Transitional integer branch is already expressed in
/// whole percent, while the Strict branch accepts only an integer followed by
/// `%`. Keep that lexical distinction here so camera rotations and other
/// scalar chart types are never accidentally interpreted as percentages.
pub(super) fn parse_chart_integer_percent(
    value: &str,
    strict: bool,
    min: u32,
    max: u32,
) -> Option<f64> {
    let raw = if let Some(percent) = value.strip_suffix('%') {
        if percent.is_empty() || !percent.bytes().all(|byte| byte.is_ascii_digit()) {
            return None;
        }
        percent.parse::<u32>().ok()?
    } else {
        if strict {
            return None;
        }
        value.trim().parse::<u32>().ok()?
    };
    (min..=max).contains(&raw).then_some(f64::from(raw))
}

#[allow(clippy::type_complexity)]
pub(super) fn extract_chartex_style_text_body(
    style_node: Node,
) -> (
    Option<String>,
    Option<f64>,
    Option<i32>,
    Option<String>,
    Option<String>,
    Option<String>,
    Option<i64>,
    Option<i64>,
    Option<i64>,
    Option<i64>,
) {
    let def_r_pr = child(style_node, "defRPr");
    let body = chart_label_body_style(child(style_node, "bodyPr"));
    (
        def_r_pr
            .and_then(|props| props.attribute("lang"))
            .map(ToOwned::to_owned),
        def_r_pr
            .and_then(|props| props.attribute("baseline"))
            .and_then(parse_chart_text_percentage),
        body.rotation,
        body.wrap,
        body.anchor,
        body.vertical_mode,
        body.left_inset,
        body.top_inset,
        body.right_inset,
        body.bottom_inset,
    )
}

pub(super) fn parse_chartex_data_point_overrides(
    series: Node,
    resolver: &dyn ColorResolver,
    image_resolver: &dyn ChartImageResolver,
    paint_budget: &mut usize,
    paint_budget_exceeded: &mut bool,
) -> Vec<ChartDataPointOverride> {
    series
        .children()
        .filter(|node| node.is_element() && node.tag_name().name() == "dataPt")
        .filter_map(|point| {
            let idx = attr(&point, "idx")?.parse::<u32>().ok()?;
            let (color, fill_hidden, line_color, line_width_emu, line_dash, line_hidden) =
                parse_data_point_shape(point, resolver);
            let shape = child(point, "spPr");
            let fill_components = shape
                .and_then(chart_style_paint_component_count)
                .unwrap_or(0);
            let line_components = shape
                .and_then(|sp_pr| child(sp_pr, "ln"))
                .and_then(chart_style_paint_component_count)
                .unwrap_or(0);
            let components = fill_components.saturating_add(line_components);
            let within_limit = fill_components <= MAX_CHART_MARKER_GRADIENT_STOPS
                && line_components <= MAX_CHART_MARKER_GRADIENT_STOPS
                && components <= *paint_budget;
            let chartex_style = if within_limit {
                *paint_budget -= components;
                shape.map(|_| {
                    parse_chartex_element_style(
                        point,
                        resolver,
                        None,
                        None,
                        image_resolver,
                        ChartImageSource::Chart,
                    )
                })
            } else {
                *paint_budget_exceeded = true;
                None
            };
            Some(ChartDataPointOverride {
                idx,
                color,
                fill_hidden,
                chartex_style,
                line_color,
                line_width_emu,
                line_dash,
                line_hidden,
                marker_symbol: None,
                marker_size: None,
                marker_fill: None,
                marker_fill_paint: None,
                marker_fill_paint_authored: None,
                marker_style: None,
                marker_line: None,
                marker_line_width_emu: None,
                marker_line_paint_authored: None,
                bubble_3d: None,
                explosion: None,
            })
        })
        .collect()
}

pub(super) fn parse_chartex_histogram_binning(series: Node) -> Option<ChartexHistogramBinning> {
    let binning = child(child(series, "layoutPr")?, "binning")?;
    let finite_text = |name: &str| {
        child(binning, name)
            .and_then(|node| node.text())
            .and_then(|text| text.trim().parse::<f64>().ok())
            .filter(|value| value.is_finite())
    };
    let finite_attr = |name: &str| {
        attr(&binning, name)
            .and_then(|text| text.parse::<f64>().ok())
            .filter(|value| value.is_finite())
    };
    Some(ChartexHistogramBinning {
        bin_size: finite_text("binSize").filter(|value| *value > 0.0),
        bin_count: child(binning, "binCount")
            .and_then(|node| node.text())
            .and_then(|text| text.trim().parse::<u32>().ok())
            .filter(|value| *value > 0),
        interval_closed: attr(&binning, "intervalClosed")
            .filter(|value| value == "l" || value == "r"),
        underflow: finite_attr("underflow"),
        overflow: finite_attr("overflow"),
    })
}

pub(super) type DataPointShape = (
    Option<String>,
    Option<bool>,
    Option<String>,
    Option<u32>,
    Option<String>,
    Option<bool>,
);

pub(super) fn parse_data_point_shape(point: Node, resolver: &dyn ColorResolver) -> DataPointShape {
    let shape = child(point, "spPr");
    let color = shape.and_then(|shape| resolver.resolve_shape_fill(shape));
    let fill_hidden = shape
        .and_then(|shape| child(shape, "noFill"))
        .map(|_| true)
        .or_else(|| color.as_ref().map(|_| false));
    let (line_color, line_width_emu, line_no_fill) = extract_sp_pr_ln_style(point, resolver);
    let line_dash = shape
        .and_then(|shape| child(shape, "ln"))
        .and_then(|line| child(line, "prstDash"))
        .and_then(|preset| attr(&preset, "val"));
    let line_hidden = if line_no_fill {
        Some(true)
    } else if line_color.is_some() || line_width_emu.is_some() || line_dash.is_some() {
        Some(false)
    } else {
        None
    };
    (
        color,
        fill_hidden,
        line_color,
        line_width_emu,
        line_dash,
        line_hidden,
    )
}

pub(super) type ChartexSeriesLabels = (
    Option<Vec<Option<String>>>,
    Option<Vec<ChartDataLabelOverride>>,
    Option<ChartSeriesDataLabels>,
);

pub(super) fn first_paragraph_properties<'a, 'input: 'a>(
    text_body: Node<'a, 'input>,
) -> Option<Node<'a, 'input>> {
    text_body
        .children()
        .find(|node| node.is_element() && node.tag_name().name() == "p")
        .and_then(|paragraph| child(paragraph, "pPr"))
}

pub(super) fn first_paragraph_default_run_props<'a, 'input: 'a>(
    text_body: Node<'a, 'input>,
) -> Option<Node<'a, 'input>> {
    first_paragraph_properties(text_body).and_then(|properties| child(properties, "defRPr"))
}

pub(super) fn parse_chartex_series_labels(
    series: Node,
    value_count: usize,
    resolver: &dyn ColorResolver,
    allow_label_paints: bool,
) -> ChartexSeriesLabels {
    let Some(labels) = child(series, "dataLabels") else {
        return (None, None, None);
    };
    let visibility = child(labels, "visibility");
    let bool_value = |name: &str| {
        visibility
            .and_then(|node| chart_text_bool_attr(node, name))
            .unwrap_or(false)
    };
    let series_tx_pr = child(labels, "txPr");
    // MS-ODRAWXML CT_DataLabels permits one paragraph without runs here.
    // Only that paragraph's default run properties participate in the
    // collection-level label style; invalid runs/later paragraphs are ignored.
    let series_run_props = series_tx_pr.and_then(first_paragraph_default_run_props);
    let series_text_paint = chart_text_paint([series_run_props], resolver);
    let series_body = chart_label_body_style(series_tx_pr.and_then(|tx| child(tx, "bodyPr")));
    let defaults = ChartSeriesDataLabels {
        deleted: None,
        show_val: bool_value("value"),
        show_cat_name: bool_value("categoryName"),
        show_ser_name: bool_value("seriesName"),
        show_percent: false,
        show_bubble_size: false,
        show_legend_key: false,
        position: attr(&labels, "pos"),
        font_color: series_text_paint.color.clone(),
        font_paint_authored: series_text_paint.authored.then_some(true),
        font_hidden: series_text_paint.hidden.then_some(true),
        format_code: child(labels, "numFmt").and_then(|node| attr(&node, "formatCode")),
        separator: child(labels, "separator").map(|node| node.text().unwrap_or("").to_owned()),
        font_bold: chart_text_bool_from_present_props(series_run_props, "b"),
        font_italic: chart_text_bool_from_present_props(series_run_props, "i"),
        font_language: series_run_props
            .and_then(|props| props.attribute("lang"))
            .map(ToOwned::to_owned),
        font_baseline: series_run_props
            .and_then(|props| props.attribute("baseline"))
            .and_then(parse_chart_text_percentage),
        font_size_hpt: series_run_props
            .and_then(|props| props.attribute("sz"))
            .and_then(parse_text_font_size_hpt),
        font_face: series_run_props.and_then(first_latin_typeface),
        text_rotation: series_body.rotation,
        text_wrap: series_body.wrap,
        text_vertical_anchor: series_body.anchor,
        text_vertical_mode: series_body.vertical_mode,
        text_l_ins_emu: series_body.left_inset,
        text_t_ins_emu: series_body.top_inset,
        text_r_ins_emu: series_body.right_inset,
        text_b_ins_emu: series_body.bottom_inset,
        text_body_authored: series_body.authored.then_some(true),
        text_align: series_tx_pr
            .and_then(first_paragraph_properties)
            .and_then(|props| attr(&props, "algn")),
        label_box: parse_label_box_with_policy(child(labels, "spPr"), resolver, allow_label_paints),
        show_leader_lines: false,
        leader_line_color: None,
        leader_line_paint_authored: None,
        leader_line_width_emu: None,
        leader_line_hidden: None,
        leader_line_dash: None,
        leader_line_style: None,
    };
    let mut colors = vec![None; value_count];
    let mut has_color = false;
    let mut overrides = Vec::new();
    for label in labels
        .children()
        .filter(|node| node.is_element() && node.tag_name().name() == "dataLabel")
    {
        let Some(index) = attr(&label, "idx")
            .and_then(|value| value.parse::<u32>().ok())
            .and_then(|value| usize::try_from(value).ok())
        else {
            continue;
        };
        // CT_DataLabel@idx addresses either a bounded series cache or a bounded
        // pre-order hierarchy node. It may legitimately exceed `value_count`
        // for treemap/sunburst, but it must never resize a sparse color vector
        // from an XML attribute (e.g. idx=4294967295).
        if index >= MAX_CHART_CACHE_POINTS {
            continue;
        }
        let tx_pr = child(label, "txPr");
        let tx_default_props = tx_pr.and_then(first_paragraph_default_run_props);
        // The rich-text parser already applies each paragraph's defRPr to its
        // own runs. Passing the first run as a body default would incorrectly
        // leak direct formatting into later sibling runs.
        let rich_runs = tx_pr.and_then(|tx| parse_data_label_rich_body(tx, None, resolver, None));
        // Paragraph defRPr is already applied to the runs in that paragraph.
        // It is a point-wide fallback only for a style-only txPr with no rich
        // text; otherwise it would leak the first paragraph into later ones.
        let point_default_props = rich_runs.is_none().then_some(tx_default_props).flatten();
        let text_paint = chart_text_paint([point_default_props], resolver);
        let font_color = text_paint.color.clone();
        if let Some(color) = font_color.clone().filter(|_| index < colors.len()) {
            colors[index] = Some(color);
            has_color = true;
        }
        let body_style = chart_label_body_style(tx_pr.and_then(|tx| child(tx, "bodyPr")));
        let label_visibility = child(label, "visibility");
        overrides.push(ChartDataLabelOverride {
            idx: index as u32,
            text: rich_runs
                .as_ref()
                .map(|runs| runs.iter().map(|run| run.text.as_str()).collect())
                .or_else(|| tx_pr.map(|node| flatten_rich_text(node, None)))
                .unwrap_or_default(),
            position: attr(&label, "pos"),
            font_color,
            font_paint_authored: text_paint.authored.then_some(true),
            font_hidden: text_paint.hidden.then_some(true),
            font_size_hpt: point_default_props
                .and_then(|props| props.attribute("sz"))
                .and_then(parse_text_font_size_hpt),
            font_face: point_default_props.and_then(first_latin_typeface),
            font_bold: chart_text_bool_from_present_props(point_default_props, "b"),
            font_italic: chart_text_bool_from_present_props(point_default_props, "i"),
            font_language: point_default_props
                .and_then(|props| props.attribute("lang"))
                .map(ToOwned::to_owned),
            font_baseline: point_default_props
                .and_then(|props| props.attribute("baseline"))
                .and_then(parse_chart_text_percentage),
            text_rotation: body_style.rotation,
            text_wrap: body_style.wrap,
            text_vertical_anchor: body_style.anchor,
            text_vertical_mode: body_style.vertical_mode,
            text_l_ins_emu: body_style.left_inset,
            text_t_ins_emu: body_style.top_inset,
            text_r_ins_emu: body_style.right_inset,
            text_b_ins_emu: body_style.bottom_inset,
            text_body_authored: body_style.authored.then_some(true),
            text_align: if rich_runs.is_none() {
                tx_pr
                    .and_then(first_paragraph_properties)
                    .and_then(|props| attr(&props, "algn"))
            } else {
                None
            },
            format_code: child(label, "numFmt").and_then(|node| attr(&node, "formatCode")),
            separator: child(label, "separator").map(|node| node.text().unwrap_or("").to_owned()),
            manual_layout: None,
            label_box: parse_label_box_with_policy(
                child(label, "spPr"),
                resolver,
                allow_label_paints,
            ),
            show_val: label_visibility.and_then(|node| chart_text_bool_attr(node, "value")),
            show_cat_name: label_visibility
                .and_then(|node| chart_text_bool_attr(node, "categoryName")),
            show_ser_name: label_visibility
                .and_then(|node| chart_text_bool_attr(node, "seriesName")),
            show_percent: None,
            show_bubble_size: None,
            show_legend_key: None,
            rich_runs,
            deleted: None,
        });
    }
    let mut override_positions: std::collections::HashMap<u32, usize> = overrides
        .iter()
        .enumerate()
        .map(|(position, override_)| (override_.idx, position))
        .collect();
    for hidden in labels
        .children()
        .filter(|node| node.is_element() && node.tag_name().name() == "dataLabelHidden")
    {
        let Some(idx) = attr(&hidden, "idx").and_then(|value| value.parse::<u32>().ok()) else {
            continue;
        };
        if usize::try_from(idx)
            .ok()
            .is_none_or(|index| index >= MAX_CHART_CACHE_POINTS)
        {
            continue;
        }
        if let Some(position) = override_positions.get(&idx).copied() {
            overrides[position].deleted = Some(true);
        } else {
            overrides.push(ChartDataLabelOverride {
                idx,
                text: String::new(),
                rich_runs: None,
                position: None,
                font_color: None,
                font_paint_authored: None,
                font_hidden: None,
                font_size_hpt: None,
                font_face: None,
                font_bold: None,
                font_italic: None,
                font_language: None,
                font_baseline: None,
                text_rotation: None,
                text_wrap: None,
                text_vertical_anchor: None,
                text_vertical_mode: None,
                text_l_ins_emu: None,
                text_t_ins_emu: None,
                text_r_ins_emu: None,
                text_b_ins_emu: None,
                text_body_authored: None,
                text_align: None,
                format_code: None,
                separator: None,
                manual_layout: None,
                label_box: None,
                show_val: None,
                show_cat_name: None,
                show_ser_name: None,
                show_percent: None,
                show_bubble_size: None,
                show_legend_key: None,
                deleted: Some(true),
            });
            override_positions.insert(idx, overrides.len() - 1);
        }
    }
    (
        has_color.then_some(colors),
        (!overrides.is_empty()).then_some(overrides),
        Some(defaults),
    )
}

pub(super) fn parse_chartex_impl(
    chartspace_root: Node,
    resolver: &dyn ColorResolver,
    style_xml: Option<&str>,
    color_style_xml: Option<&str>,
    references: &mut dyn ChartReferenceResolver,
    image_resolver: &dyn ChartImageResolver,
) -> Option<ChartModel> {
    let root = chartspace_root;
    let chart_node = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "chart")?;
    let style_doc = style_xml.and_then(|xml| crate::depth::parse_guarded(xml).ok());
    let color_style = color_style_xml.and_then(|xml| parse_chart_color_style(xml, resolver));
    let unresolved_color_style_palette =
        (color_style_xml.is_some() && color_style.is_none()).then(|| vec![None]);
    let style_element = |name: &str| {
        style_doc.as_ref().and_then(|doc| {
            doc.root_element()
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == name)
        })
    };

    // CT_PlotAreaRegion may contain several series. `hidden` is an authored
    // series visibility flag; a hidden leading series must not select the
    // chart layout or data used by the visible plot.
    let all_series_nodes: Vec<Node> = root
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "series")
        .collect();
    let series_nodes: Vec<Node> = all_series_nodes
        .iter()
        .copied()
        .filter(|node| {
            !attr(node, "hidden")
                .is_some_and(|value| value == "1" || value.eq_ignore_ascii_case("true"))
        })
        .collect();
    let series_format_index = |series: Node| -> u32 {
        attr(&series, "formatIdx")
            .and_then(|value| value.parse::<u32>().ok())
            .or_else(|| {
                all_series_nodes
                    .iter()
                    .position(|candidate| *candidate == series)
                    .and_then(|index| u32::try_from(index).ok())
            })
            .unwrap_or(0)
    };
    // MS-ODRAWXML §2.24.4.19 defines this closed set.  Preserve a future
    // identifier verbatim, but let it own the chart-wide fail-closed result:
    // otherwise a preceding known series or a trailing `paretoLine` could
    // silently normalize the chart to a visually similar implemented layout.
    const KNOWN_SERIES_LAYOUTS: [&str; 8] = [
        "boxWhisker",
        "clusteredColumn",
        "funnel",
        "paretoLine",
        "regionMap",
        "sunburst",
        "treemap",
        "waterfall",
    ];
    let unknown_series = series_nodes.iter().copied().find(|node| {
        !attr(node, "layoutId")
            .is_some_and(|layout| KNOWN_SERIES_LAYOUTS.contains(&layout.as_str()))
    });
    // A Pareto plot is represented by an ordinary owner series plus an
    // auxiliary `paretoLine` whose `ownerIdx` names the owner's original
    // document-order series index (CT_Series@ownerIdx, [MS-ODRAWXML]
    // 2.24.3.77). This is independent of `formatIdx`; select the linked owner
    // even when a hidden or auxiliary series appears first.
    let pareto_pair = if unknown_series.is_none() {
        series_nodes.iter().copied().find_map(|pareto| {
            if attr(&pareto, "layoutId").as_deref() != Some("paretoLine") {
                return None;
            }
            let owner_idx = attr(&pareto, "ownerIdx")?.parse::<usize>().ok()?;
            let owner = *all_series_nodes.get(owner_idx)?;
            let owner_is_hidden = attr(&owner, "hidden")
                .is_some_and(|value| value == "1" || value.eq_ignore_ascii_case("true"));
            (owner != pareto
                && !owner_is_hidden
                && attr(&owner, "layoutId").as_deref() != Some("paretoLine"))
            .then_some((owner, pareto))
        })
    } else {
        None
    };
    let (series_node, pareto_series_node) = if let Some(unknown) = unknown_series {
        (unknown, None)
    } else {
        pareto_pair
            .map(|(owner, pareto)| (owner, Some(pareto)))
            .unwrap_or((*series_nodes.first()?, None))
    };
    let layout_id = attr(&series_node, "layoutId").unwrap_or_default();
    // [MS-ODRAWXML] represents a histogram as a clusteredColumn series with a
    // CT_Binning child; `histogram` is not an ST_SeriesLayout enumeration.
    // Normalize the semantic family here so raw observations cannot reach the
    // ordinary clustered-column renderer.
    let chartex_histogram_binning = (pareto_series_node.is_none()
        && layout_id == "clusteredColumn")
        .then(|| parse_chartex_histogram_binning(series_node))
        .flatten();
    let chart_type = if pareto_series_node.is_some() {
        "pareto".to_string()
    } else if chartex_histogram_binning.is_some() {
        "histogram".to_string()
    } else {
        layout_id
    };
    let data_by_id: std::collections::HashMap<String, Node> = root
        .descendants()
        .filter(|node| node.is_element() && node.tag_name().name() == "data")
        .filter_map(|data| attr(&data, "id").map(|id| (id, data)))
        .collect();
    let data_for_series = |series: Node| -> Option<Node> {
        let data_id = child(series, "dataId").and_then(|node| attr(&node, "val"));
        match data_id {
            Some(id) => data_by_id.get(&id).copied(),
            // Retain compatibility with early ChartEx producers that placed a
            // single data block in chartSpace but omitted CT_Series.dataId.
            None => Some(root),
        }
    };
    let primary_data = data_for_series(series_node)?;
    // Label paint work is bounded per chart, not per CT_Series. Preflight every
    // visible series that this parser will retain before resolving any of its
    // gradient/pattern recipes, so several individually valid series cannot
    // multiply the chart-wide allocation ceiling.
    let retained_extra_label_series = (chart_type == "clusteredColumn")
        .then(|| {
            series_nodes
                .iter()
                .copied()
                .skip(1)
                .filter(|node| attr(node, "layoutId").as_deref() == Some(chart_type.as_str()))
                .filter(|node| data_for_series(*node).is_some())
        })
        .into_iter()
        .flatten();
    let allow_chartex_label_paints = chartex_label_paint_recipes_within_limit(
        std::iter::once(series_node)
            .chain(pareto_series_node)
            .chain(retained_extra_label_series),
        MAX_CHART_LABEL_PAINT_COMPONENTS,
    );
    let mut chartex_point_paint_budget = MAX_CHART_MARKER_PAINT_COMPONENTS;
    let mut chartex_point_paint_budget_exceeded = false;

    // ── chartEx title (MS 2014 chartex ext) ──────────────────────────────────
    // Office may save either DrawingML rich text or the compact
    // `<cx:txData><cx:v>` form used by Excel-authored XLSX chartEx parts.
    let chartex_title = child(chart_node, "title").and_then(chartex_text);
    let chartex_title_present = child(chart_node, "title").is_some();

    // Keep chart-local text separate from the associated style part. Core
    // performs the single direct > linked > numeric cascade after both source
    // layers have been parsed; flattening them here loses authored ownership
    // and prevents relative fontRef/styleClr colors from using object indexes.
    let chartex_title_font_size_hpt = extract_chartex_title_size(root);
    let chartex_title_font_bold = extract_chart_title_bold(chart_node);
    let chartex_title_font_italic = extract_chart_title_italic(chart_node);
    let chartex_title_text_paint = chart_title_text_paint(Some(chart_node), resolver);
    let chartex_title_font_color = if chartex_title_text_paint.authored {
        chartex_title_text_paint.color.clone()
    } else {
        None
    };
    let chartex_title_font_face = child(chart_node, "title").and_then(first_latin_typeface);

    // ── chartEx theme accent palette ─────────────────────────────────────────
    // boxWhisker series and hierarchy/region-map branches color off the theme
    // accents (`accent[(idx % 6) + 1]`, the same cycle Office draws). Resolve
    // accent1..6 once here; `None` when the resolver owns no default palette
    // (pptx), letting the renderer fall back to its own `CHART_PALETTE`.
    let theme_accents: Option<Vec<String>> = if matches!(
        chart_type.as_str(),
        "waterfall" | "boxWhisker" | "sunburst" | "treemap" | "regionMap"
    ) {
        let accents: Vec<String> = (0..6)
            .filter_map(|i| resolver.resolve_series_accent(i))
            .collect();
        if accents.len() == 6 {
            Some(accents)
        } else {
            None
        }
    } else {
        None
    };
    let chartex_color_style_method = color_style.as_ref().map(|(method, _)| method.clone());
    let chartex_color_palette = color_style
        .as_ref()
        .map(|(_, palette)| palette.clone())
        .or_else(|| unresolved_color_style_palette.clone());
    // Linked Chart Style roles apply to every classic/ChartEx family, not only
    // the branch-colored layouts that expose `chartexAccents`. Resolve the
    // ordinary six-color theme palette independently whenever a style part is
    // present so `phClr` recipes are not discarded for other layouts.
    let theme_style_palette = style_doc.as_ref().and_then(|_| {
        let colors = (0..6)
            .map(|index| resolver.resolve_series_accent(index))
            .collect::<Vec<_>>();
        colors.iter().any(Option::is_some).then_some(colors)
    });
    let style_palette = chartex_color_palette
        .as_deref()
        .or(theme_style_palette.as_deref());
    let chart_style_roles = if style_xml.is_some() && style_doc.is_none() {
        Some(unreadable_chart_style_role_table())
    } else {
        style_doc.as_ref().and_then(|document| {
            parse_chart_style_role_table(
                document.root_element(),
                resolver,
                style_palette,
                chartex_color_style_method.as_deref(),
                image_resolver,
            )
            .or_else(|| {
                unresolved_chart_style_role_table(document.root_element(), resolver, image_resolver)
            })
        })
    };
    let chartex_data_point_style = chart_style_roles
        .as_ref()
        .and_then(|roles| roles.get("dataPoint"))
        .cloned();
    let chartex_data_point_line_style = chart_style_roles
        .as_ref()
        .and_then(|roles| roles.get("dataPointLine"))
        .cloned();
    let chartex_series_line_style = chart_style_roles
        .as_ref()
        .and_then(|roles| roles.get("seriesLine"))
        .cloned();
    let chartex_data_point_marker_style = chart_style_roles
        .as_ref()
        .and_then(|roles| roles.get("dataPointMarker"))
        .cloned();
    let marker_layout = style_element("dataPointMarkerLayout");
    let chartex_marker_size_pt = marker_layout
        .and_then(|node| node.attribute("size"))
        .and_then(|value| value.parse::<u8>().ok())
        .filter(|value| (2..=72).contains(value));
    let chartex_marker_symbol = marker_layout
        .and_then(|node| node.attribute("symbol"))
        .map(ToOwned::to_owned);
    let chart_style_marker_size_pt = chartex_marker_size_pt;
    let chart_style_marker_symbol = chartex_marker_symbol.clone();
    // [MS-ODRAWXML] CT_SeriesElementVisibilities: an authored false value
    // suppresses waterfall connector lines. Keep omission as `None` so the
    // renderer can preserve the layout's ordinary default without fabricating
    // an XML value.
    let chartex_connector_lines = child(series_node, "layoutPr")
        .and_then(|layout| child(layout, "visibility"))
        .and_then(|visibility| chart_text_bool_attr(visibility, "connectorLines"));
    // Keep the raw theme palette separate from effective style-role paint.
    let chartex_accents = theme_accents;

    // ── chartEx box-and-whisker structured parse ─────────────────────────────
    let chartex_box = if chart_type == "boxWhisker" {
        parse_chartex_boxwhisker(root, resolver, references, image_resolver)
    } else {
        None
    };

    // ── chartEx sunburst structured parse ────────────────────────────────────
    let chartex_sunburst = if chart_type == "sunburst" {
        parse_chartex_sunburst(primary_data, references)
    } else {
        None
    };

    // ── chartEx treemap structured parse ────────────────────────────────────
    let chartex_treemap = if chart_type == "treemap" {
        parse_chartex_treemap(primary_data, series_node, references)
    } else {
        None
    };

    // ── chartEx Region Map structured parse ────────────────────────────────
    // MS-ODRAWXML §2.24.3.77/79: source identities and color values are data;
    // provider/geocoding/geometry remain a renderer/host responsibility.
    let chartex_region_map = if chart_type == "regionMap" {
        parse_chartex_region_map(primary_data, series_node, resolver, references)
    } else {
        None
    };

    // The flat compatibility fields use the deepest hierarchy labels/sizes.
    // Formula-only chartEx dimensions are resolved through the package host.
    let hierarchy_rows = chartex_treemap
        .as_ref()
        .map(|data| data.rows.as_slice())
        .or_else(|| chartex_sunburst.as_ref().map(|data| data.rows.as_slice()));
    let categories: Vec<String> = hierarchy_rows
        .map(|rows| {
            rows.iter()
                .map(|row| row.path.last().cloned().unwrap_or_default())
                .collect()
        })
        .or_else(|| chartex_box.as_ref().map(|data| data.categories.clone()))
        .or_else(|| {
            chartex_region_map
                .as_ref()
                .map(|data| data.rows.iter().map(|row| row.label.clone()).collect())
        })
        .or_else(|| {
            chartex_string_levels(primary_data, references)
                .and_then(|levels| levels.into_iter().next())
        })
        .unwrap_or_default();

    let pt_count = categories.len().max(1);

    let raw_values: Vec<Option<f64>> = hierarchy_rows
        .map(|rows| rows.iter().map(|row| Some(row.size)).collect())
        .or_else(|| {
            if chartex_box.is_some() {
                None
            } else if let Some(region_map) = chartex_region_map.as_ref() {
                Some(region_map.rows.iter().map(|row| row.value).collect())
            } else {
                chartex_number_values(primary_data, &["val"], references)
            }
        })
        .unwrap_or_else(|| vec![None; pt_count]);
    let source_number_format =
        chartex_number_format(primary_data, &["size", "val", "colorVal"], references);

    let series_name_for = |node: Node, references: &mut dyn ChartReferenceResolver| {
        node.descendants()
            .find(|child_node| child_node.is_element() && child_node.tag_name().name() == "txData")
            .and_then(|tx_data| {
                child(tx_data, "v")
                    .and_then(|value| value.text())
                    .map(str::trim)
                    .filter(|value| !value.is_empty())
                    .map(ToOwned::to_owned)
                    .or_else(|| {
                        child(tx_data, "f")
                            .and_then(|formula| formula.text())
                            .map(str::trim)
                            .filter(|formula| !formula.is_empty())
                            .and_then(|formula| references.resolve_strings(formula))
                            .and_then(|values| {
                                values.into_iter().find(|value| !value.trim().is_empty())
                            })
                    })
            })
            .unwrap_or_default()
    };
    let series_name = series_name_for(series_node, references);

    // `<cx:subtotals><cx:idx val>` identifies only points explicitly marked as
    // totals. The first waterfall point starts at zero geometrically, but it is
    // still an ordinary increase/decrease point unless index 0 is present.
    let mut subtotal_indices: Vec<u32> = Vec::new();
    if let Some(subtotals_node) = series_node
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "subtotals")
    {
        for idx_node in subtotals_node
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "idx")
        {
            if let Some(v) = attr(&idx_node, "val").and_then(|v| v.parse::<u32>().ok()) {
                if !subtotal_indices.contains(&v) {
                    subtotal_indices.push(v);
                }
            }
        }
    }

    // Series shape properties are local formatting on CT_Series
    // ([MS-ODRAWXML] 2.24.3.77). Preserve both fill and outline instead of
    // letting the linked Chart Style silently replace an authored `spPr`.
    let color = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "spPr")
        .and_then(|sp| {
            sp.children()
                .find(|n| n.is_element() && n.tag_name().name() == "solidFill")
        })
        .and_then(|fill| resolver.resolve_solid_fill(fill));
    let (line_color, line_width_emu, line_no_fill) = extract_sp_pr_ln_style(series_node, resolver);
    let chartex_style = child(series_node, "spPr").map(|_| {
        parse_chartex_element_style(
            series_node,
            resolver,
            None,
            None,
            image_resolver,
            ChartImageSource::Chart,
        )
    });
    let line_hidden = if line_no_fill {
        Some(true)
    } else if line_color.is_some() || line_width_emu.is_some() {
        Some(false)
    } else {
        None
    };

    let (data_label_colors, data_label_overrides, series_data_labels) = parse_chartex_series_labels(
        series_node,
        raw_values.len(),
        resolver,
        allow_chartex_label_paints,
    );

    let mut series = vec![ChartSeries {
        name: series_name,
        chartex_format_idx: Some(series_format_index(series_node)),
        values: raw_values,
        color,
        fill_pattern: None,
        invert_if_negative: None,
        automatic_negative_style: None,
        inverted_fill: None,
        inverted_fill_hidden: None,
        inverted_fill_authored: None,
        inverted_line_color: None,
        inverted_line_width_emu: None,
        inverted_line_hidden: None,
        inverted_line_authored: None,
        chartex_style,
        line_color,
        line_width_emu,
        three_d_shape: None,
        source_hidden: None,
        data_point_colors: None,
        explosion: None,
        data_label_colors,
        categories: None,
        bubble_x_source_is_string: None,
        bubble_sizes: None,
        bubble_3d_group_default: None,
        bubble_3d: None,
        val_format_code: source_number_format,
        cat_format_code: None,
        cat_format_builtin_id: None,
        cat_format_codes: None,
        label_color: None,
        series_type: None,
        line_group_index: None,
        area_group_index: None,
        bar_group_index: None,
        bar_group_direction: None,
        bar_group_grouping: None,
        bar_group_gap_width: None,
        bar_group_overlap: None,
        use_secondary_axis: None,
        show_marker: None,
        marker_symbol: None,
        automatic_marker_symbol: None,
        marker_size: None,
        marker_fill: None,
        marker_fill_paint: None,
        marker_fill_paint_authored: None,
        marker_style: None,
        marker_line: None,
        marker_line_width_emu: None,
        marker_line_paint_authored: None,
        data_point_overrides: {
            let overrides = parse_chartex_data_point_overrides(
                series_node,
                resolver,
                image_resolver,
                &mut chartex_point_paint_budget,
                &mut chartex_point_paint_budget_exceeded,
            );
            (!overrides.is_empty()).then_some(overrides)
        },
        data_label_overrides,
        series_data_labels,
        err_bars: None,
        // chartEx (waterfall) has no `<c:smooth>` concept.
        smooth: None,
        // chartEx series carry no classic `<c:trendline>`.
        trend_lines: None,
        // chartEx has no scatter connecting line to suppress.
        line_hidden,
    }];

    // Preserve the authored Pareto line as a style carrier. Its cached values
    // are retained on the wire for diagnostics, while the core derives stable
    // cumulative fractions from the owner values so invalid/negative filtering
    // and source-identity remapping happen exactly once.
    if let Some(pareto_node) = pareto_series_node {
        let pareto_data = data_for_series(pareto_node);
        let pareto_values = pareto_data
            .and_then(|data| chartex_number_values(data, &["val"], references))
            .unwrap_or_default();
        let pareto_color =
            child(pareto_node, "spPr").and_then(|shape| resolver.resolve_shape_fill(shape));
        let (pareto_line_color, pareto_line_width_emu, pareto_line_no_fill) =
            extract_sp_pr_ln_style(pareto_node, resolver);
        let pareto_chartex_style = child(pareto_node, "spPr").map(|_| {
            parse_chartex_element_style(
                pareto_node,
                resolver,
                None,
                None,
                image_resolver,
                ChartImageSource::Chart,
            )
        });
        let mut pareto_series = series[0].clone();
        pareto_series.name = series_name_for(pareto_node, references);
        pareto_series.chartex_format_idx = Some(series_format_index(pareto_node));
        pareto_series.values = pareto_values;
        pareto_series.categories = None;
        pareto_series.color = pareto_line_color.clone().or(pareto_color);
        pareto_series.chartex_style = pareto_chartex_style;
        pareto_series.line_color = pareto_line_color;
        pareto_series.line_width_emu = pareto_line_width_emu;
        pareto_series.line_hidden = if pareto_line_no_fill {
            Some(true)
        } else if pareto_series.line_color.is_some() || pareto_series.line_width_emu.is_some() {
            Some(false)
        } else {
            None
        };
        pareto_series.series_type = Some("line".to_string());
        pareto_series.use_secondary_axis = Some(true);
        pareto_series.show_marker = Some(false);
        pareto_series.data_point_overrides = {
            let overrides = parse_chartex_data_point_overrides(
                pareto_node,
                resolver,
                image_resolver,
                &mut chartex_point_paint_budget,
                &mut chartex_point_paint_budget_exceeded,
            );
            (!overrides.is_empty()).then_some(overrides)
        };
        let (label_colors, label_overrides, label_defaults) = parse_chartex_series_labels(
            pareto_node,
            pareto_series.values.len(),
            resolver,
            allow_chartex_label_paints,
        );
        pareto_series.data_label_colors = label_colors;
        pareto_series.data_label_overrides = label_overrides;
        pareto_series.series_data_labels = label_defaults;
        series.push(pareto_series);
    }

    // A flat clustered-column ChartEx plot may contain several CT_Series, each
    // selecting its own CT_Data through dataId. Preserve every visible series
    // rather than silently collapsing the plot to the first one.
    if chart_type == "clusteredColumn" {
        for extra_node in series_nodes
            .iter()
            .copied()
            .skip(1)
            .filter(|node| attr(node, "layoutId").as_deref() == Some(chart_type.as_str()))
        {
            let Some(extra_data) = data_for_series(extra_node) else {
                continue;
            };
            let extra_values =
                chartex_number_values(extra_data, &["val"], references).unwrap_or_default();
            let extra_categories = chartex_string_levels(extra_data, references)
                .and_then(|levels| levels.into_iter().next())
                .unwrap_or_default();
            let extra_name = series_name_for(extra_node, references);
            let extra_color =
                child(extra_node, "spPr").and_then(|shape| resolver.resolve_shape_fill(shape));
            let (extra_line_color, extra_line_width_emu, extra_line_no_fill) =
                extract_sp_pr_ln_style(extra_node, resolver);
            let extra_chartex_style = child(extra_node, "spPr").map(|_| {
                parse_chartex_element_style(
                    extra_node,
                    resolver,
                    None,
                    None,
                    image_resolver,
                    ChartImageSource::Chart,
                )
            });
            let mut extra = series[0].clone();
            extra.name = extra_name;
            extra.chartex_format_idx = Some(series_format_index(extra_node));
            extra.values = extra_values;
            extra.categories = (!extra_categories.is_empty()).then_some(extra_categories);
            extra.val_format_code = chartex_number_format(extra_data, &["val"], references);
            extra.color = extra_color;
            extra.chartex_style = extra_chartex_style;
            extra.line_color = extra_line_color;
            extra.line_width_emu = extra_line_width_emu;
            extra.line_hidden = if extra_line_no_fill {
                Some(true)
            } else if extra.line_color.is_some() || extra.line_width_emu.is_some() {
                Some(false)
            } else {
                None
            };
            extra.data_point_overrides = {
                let overrides = parse_chartex_data_point_overrides(
                    extra_node,
                    resolver,
                    image_resolver,
                    &mut chartex_point_paint_budget,
                    &mut chartex_point_paint_budget_exceeded,
                );
                (!overrides.is_empty()).then_some(overrides)
            };
            let (label_colors, label_overrides, label_defaults) = parse_chartex_series_labels(
                extra_node,
                extra.values.len(),
                resolver,
                allow_chartex_label_paints,
            );
            extra.data_label_colors = label_colors;
            extra.data_label_overrides = label_overrides;
            extra.series_data_labels = label_defaults;
            series.push(extra);
        }
    }

    // ChartEx axis visibility — shared helper that pairs each `<cx:axis hidden>`
    // with its `<cx:catScaling>` / `<cx:valScaling>` child to disambiguate cat
    // vs. val (chartEx doesn't declare axis kind via the `id` attribute).
    let (cat_axis_hidden, val_axis_hidden) = extract_chartex_axis_hidden(root);
    let cat_axis = root.descendants().find(|axis| {
        axis.is_element()
            && axis.tag_name().name() == "axis"
            && axis
                .children()
                .any(|child| child.is_element() && child.tag_name().name() == "catScaling")
    });
    let val_axis = root.descendants().find(|axis| {
        axis.is_element()
            && axis.tag_name().name() == "axis"
            && axis
                .children()
                .any(|child| child.is_element() && child.tag_name().name() == "valScaling")
    });
    let cat_axis_major_tick_mark = extract_chartex_axis_tick_mark(cat_axis, "majorTickMarks");
    let val_axis_major_tick_mark = extract_chartex_axis_tick_mark(val_axis, "majorTickMarks");
    let val_axis_minor_tick_mark = extract_chartex_axis_tick_mark(val_axis, "minorTickMarks");
    let val_scaling = val_axis.and_then(|axis| child(axis, "valScaling"));
    let val_min = val_scaling
        .and_then(|scaling| attr(&scaling, "min"))
        .and_then(|value| value.parse::<f64>().ok());
    let val_max = val_scaling
        .and_then(|scaling| attr(&scaling, "max"))
        .and_then(|value| value.parse::<f64>().ok());
    // MS-ODRAWXML §2.24.3.90 CT_ValueAxisScaling stores the ChartEx major
    // interval as an attribute on `<cx:valScaling>` (unlike the classic
    // `<c:valAx><c:majorUnit val>` child). `auto` and an omitted attribute both
    // remain `None`; a positive finite number is an authored override.
    let val_axis_major_unit = val_scaling
        .and_then(|scaling| attr(&scaling, "majorUnit"))
        .and_then(|value| value.parse::<f64>().ok())
        .filter(|value| value.is_finite() && *value > 0.0);
    let val_axis_minor_unit = val_scaling
        .and_then(|scaling| attr(&scaling, "minorUnit"))
        .and_then(|value| value.parse::<f64>().ok())
        .filter(|value| value.is_finite() && *value > 0.0);
    let cat_axis_title = cat_axis
        .and_then(|axis| child(axis, "title"))
        .and_then(chartex_text);
    let val_axis_title = val_axis
        .and_then(|axis| child(axis, "title"))
        .and_then(chartex_text);
    let inline_cat_title_size = cat_axis.and_then(extract_axis_title_size);
    let inline_cat_title_bold = cat_axis.and_then(extract_axis_title_bold);
    let inline_cat_title_italic = cat_axis.and_then(extract_axis_title_italic);
    let inline_cat_title_color = cat_axis.and_then(|axis| extract_axis_title_color(axis, resolver));
    let inline_cat_title_face = cat_axis.and_then(extract_axis_title_face);
    let cat_axis_title_rotation = cat_axis.and_then(extract_axis_title_rotation);
    let cat_axis_title_vertical_mode = cat_axis.and_then(extract_axis_title_vertical_mode);
    let cat_axis_title_manual_layout = cat_axis.and_then(extract_axis_title_manual_layout);
    let cat_axis_title_text_vertical_inset_emu =
        cat_axis.and_then(extract_axis_title_vertical_inset);
    let inline_val_title_size = val_axis.and_then(extract_axis_title_size);
    let inline_val_title_bold = val_axis.and_then(extract_axis_title_bold);
    let inline_val_title_italic = val_axis.and_then(extract_axis_title_italic);
    let inline_val_title_color = val_axis.and_then(|axis| extract_axis_title_color(axis, resolver));
    let inline_val_title_face = val_axis.and_then(extract_axis_title_face);
    let val_axis_title_rotation = val_axis.and_then(extract_axis_title_rotation);
    let val_axis_title_vertical_mode = val_axis.and_then(extract_axis_title_vertical_mode);
    let val_axis_title_manual_layout = val_axis.and_then(extract_axis_title_manual_layout);
    let val_axis_title_text_vertical_inset_emu =
        val_axis.and_then(extract_axis_title_vertical_inset);
    let cat_axis_title_font_size_hpt = inline_cat_title_size;
    let cat_axis_title_font_bold = inline_cat_title_bold;
    let cat_axis_title_font_italic = inline_cat_title_italic;
    let cat_axis_title_text_paint = chart_title_text_paint(cat_axis, resolver);
    let cat_axis_title_font_color = if cat_axis_title_text_paint.authored {
        cat_axis_title_text_paint.color.clone()
    } else {
        inline_cat_title_color
    };
    let cat_axis_title_font_face = inline_cat_title_face;
    let val_axis_title_font_size_hpt = inline_val_title_size;
    let val_axis_title_font_bold = inline_val_title_bold;
    let val_axis_title_font_italic = inline_val_title_italic;
    let val_axis_title_text_paint = chart_title_text_paint(val_axis, resolver);
    let val_axis_title_font_color = if val_axis_title_text_paint.authored {
        val_axis_title_text_paint.color.clone()
    } else {
        inline_val_title_color
    };
    let val_axis_title_font_face = inline_val_title_face;
    let cat_axis_font_size_hpt = cat_axis.and_then(extract_axis_tick_label_size);
    let cat_axis_font_bold = cat_axis.and_then(extract_axis_tick_label_bold);
    let cat_axis_font_italic = cat_axis.and_then(extract_axis_tick_label_italic);
    let cat_axis_text_paint = cat_axis
        .map(|axis| chart_text_body_paint(child(axis, "txPr"), resolver))
        .unwrap_or_default();
    let cat_axis_font_color = if cat_axis_text_paint.authored {
        cat_axis_text_paint.color.clone()
    } else {
        cat_axis.and_then(|axis| extract_axis_tick_label_color(axis, resolver))
    };
    let cat_axis_font_face = cat_axis.and_then(extract_axis_tick_label_face);
    let val_axis_font_size_hpt = val_axis.and_then(extract_axis_tick_label_size);
    let val_axis_font_bold = val_axis.and_then(extract_axis_tick_label_bold);
    let val_axis_font_italic = val_axis.and_then(extract_axis_tick_label_italic);
    let val_axis_text_paint = val_axis
        .map(|axis| chart_text_body_paint(child(axis, "txPr"), resolver))
        .unwrap_or_default();
    let val_axis_font_color = if val_axis_text_paint.authored {
        val_axis_text_paint.color.clone()
    } else {
        val_axis.and_then(|axis| extract_axis_tick_label_color(axis, resolver))
    };
    let val_axis_font_face = val_axis.and_then(extract_axis_tick_label_face);
    let data_labels = child(series_node, "dataLabels");
    let data_label_font_size_hpt = data_labels.and_then(extract_axis_tick_label_size);
    let data_label_font_bold = data_labels.and_then(extract_axis_tick_label_bold);
    let data_label_font_italic = data_labels.and_then(extract_axis_tick_label_italic);
    let data_label_text_paint = chart_text_body_paint(data_labels, resolver);
    let data_label_font_color = if data_label_text_paint.authored {
        data_label_text_paint.color.clone()
    } else {
        None
    };
    let data_label_font_face = data_labels.and_then(extract_axis_tick_label_face);
    let data_label_position = data_labels.and_then(|labels| attr(&labels, "pos"));
    let (cat_axis_line_color, cat_axis_line_width_emu, cat_axis_line_hidden) = cat_axis
        .map(|axis| extract_axis_line_style(axis, resolver))
        .unwrap_or((None, None, false));
    let cat_axis_line_dash = cat_axis.and_then(extract_axis_line_dash);
    let cat_axis_line_paint_authored =
        cat_axis.and_then(|axis| extract_direct_shape_line(axis, resolver).paint_authored);
    let (val_axis_line_color, val_axis_line_width_emu, val_axis_line_hidden) = val_axis
        .map(|axis| extract_axis_line_style(axis, resolver))
        .unwrap_or((None, None, false));
    let val_axis_line_dash = val_axis.and_then(extract_axis_line_dash);
    let val_axis_line_paint_authored =
        val_axis.and_then(|axis| extract_direct_shape_line(axis, resolver).paint_authored);
    let val_axis_major_gridlines = val_axis.map(axis_major_gridlines_visible);
    let val_axis_minor_gridlines = val_axis.map(axis_has_minor_gridlines);
    let (
        val_axis_minor_gridline_color,
        val_axis_minor_gridline_width_emu,
        val_axis_minor_gridline_dash,
    ) = val_axis
        .map(|axis| extract_minor_gridline_style(axis, resolver))
        .unwrap_or((None, None, None));
    let cat_axis_minor_gridlines = cat_axis.map(axis_has_minor_gridlines);
    let (
        cat_axis_minor_gridline_color,
        cat_axis_minor_gridline_width_emu,
        cat_axis_minor_gridline_dash,
    ) = cat_axis
        .map(|axis| extract_minor_gridline_style(axis, resolver))
        .unwrap_or((None, None, None));
    let (val_axis_gridline_color, val_axis_gridline_width_emu, val_axis_gridline_dash) = val_axis
        .map(|axis| extract_gridline_style(axis, resolver))
        .unwrap_or((None, None, None));
    let val_axis_format_code = val_axis
        .and_then(|axis| child(axis, "numFmt"))
        .and_then(|format| attr(&format, "formatCode"));
    let legend = root
        .descendants()
        .find(|node| node.is_element() && node.tag_name().name() == "legend");
    let show_legend = legend.is_some();
    let legend_pos = legend.and_then(|node| attr(&node, "pos"));
    let (legend_font_face, legend_font_size_hpt, legend_font_bold, legend_font_italic) =
        extract_legend_text_props(root);
    let legend_font_color = extract_legend_font_color(root, resolver);
    let legend_frame = extract_legend_frame_style(root, resolver);
    let chart_line_style = extract_direct_shape_line(root, resolver);

    // `<cx:catScaling gapWidth>` (chartEx) — same semantics as legacy
    // `<c:gapWidth>` but stored as a *fraction* (e.g. 0.8 ≡ 80%) instead of
    // an integer percentage. Convert to the legacy percentage form so the
    // shared renderer's `barW = catGap / (1 + gapWidth/100)` formula works
    // uniformly across chart types. Omission stays `None`: ChartEx has no
    // schema default, so the renderer owns the shared ordinal-layout policy.
    let bar_gap_width = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "catScaling")
        .and_then(|n| attr(&n, "gapWidth"))
        .and_then(|v| v.parse::<f64>().ok())
        .map(|frac| (frac * 100.0).round() as i32);
    let chart_space_sp_pr = root
        .children()
        .find(|node| node.is_element() && node.tag_name().name() == "spPr");
    let chart_fill_style = extract_direct_shape_fill_with_images(
        chart_space_sp_pr,
        resolver,
        image_resolver,
        ChartImageSource::Chart,
    );
    // DrawingML shape properties own fill and outline independently. The
    // Chart Style `allowNo*Override` modifiers permit an *explicit* local
    // noFill choice to replace the corresponding style component; they do not
    // turn an omitted sibling component into noFill. This matters for the
    // common ChartEx chart-space form that authors only `<a:ln><a:noFill/>`:
    // its linked/default chart-area fill must remain eligible.
    let chart_bg = if chart_fill_style.paint_authored == Some(true) {
        chart_fill_style.color.clone()
    } else {
        resolver.default_chart_bg()
    };

    if chartex_point_paint_budget_exceeded {
        return None;
    }

    Some(ChartModel {
        chart_type,
        title: chartex_title,
        title_rich_runs: None,
        title_present: chartex_title_present,
        // chartEx data lives in its structured fields; not asserted here.
        authored_without_series: false,
        categories,
        category_source_hidden: None,
        category_levels: None,
        series,
        plot_groups: None,
        // chartEx layouts color by branch/series index, not §21.2.2.227
        // varyColors (a `<c:>` chart-group element that chartEx has no analog
        // of), so the flag never applies here.
        vary_colors: None,
        chart_text_boxes: None,
        chart_text_style: extract_chart_space_text_style(root, resolver),
        chart_area_style: parse_direct_chart_effect_style(root, resolver),
        plot_area_style: root
            .descendants()
            .find(|node| node.is_element() && node.tag_name().name() == "plotArea")
            .and_then(|node| parse_direct_chart_effect_style(node, resolver)),
        legend_style: legend.and_then(|node| parse_direct_chart_effect_style(node, resolver)),
        title_style: child(chart_node, "title")
            .and_then(|node| parse_direct_chart_effect_style(node, resolver)),
        cat_axis_style: cat_axis.and_then(|node| parse_direct_chart_effect_style(node, resolver)),
        val_axis_style: val_axis.and_then(|node| parse_direct_chart_effect_style(node, resolver)),
        cat_axis_title_style: cat_axis
            .and_then(|axis| child(axis, "title"))
            .and_then(|node| parse_direct_chart_effect_style(node, resolver)),
        val_axis_title_style: val_axis
            .and_then(|axis| child(axis, "title"))
            .and_then(|node| parse_direct_chart_effect_style(node, resolver)),
        cat_axis_major_gridline_style: cat_axis
            .and_then(|axis| child(axis, "majorGridlines"))
            .and_then(|node| parse_direct_chart_effect_style(node, resolver)),
        cat_axis_minor_gridline_style: cat_axis
            .and_then(|axis| child(axis, "minorGridlines"))
            .and_then(|node| parse_direct_chart_effect_style(node, resolver)),
        val_axis_major_gridline_style: val_axis
            .and_then(|axis| child(axis, "majorGridlines"))
            .and_then(|node| parse_direct_chart_effect_style(node, resolver)),
        val_axis_minor_gridline_style: val_axis
            .and_then(|axis| child(axis, "minorGridlines"))
            .and_then(|node| parse_direct_chart_effect_style(node, resolver)),
        val_max,
        val_min,
        subtotal_indices,
        show_data_labels: false,
        cat_axis_hidden,
        val_axis_hidden,
        plot_area_bg: None,
        plot_area_fill: None,
        plot_area_fill_hidden: None,
        plot_area_fill_paint_authored: None,
        plot_area_fill_automatic: None,
        plot_area_line_color: None,
        plot_area_line_fill: None,
        plot_area_line_width_emu: None,
        plot_area_line_dash: None,
        plot_area_line_dash_authored: None,
        plot_area_line_custom_dash: None,
        plot_area_line_cap: None,
        plot_area_line_join: None,
        plot_area_line_compound: None,
        plot_area_line_hidden: None,
        plot_area_line_paint_authored: None,
        chart_bg,
        chart_fill: chart_fill_style.fill,
        chart_fill_hidden: chart_fill_style.hidden,
        chart_fill_paint_authored: chart_fill_style.paint_authored,
        rounded_corners: None,
        plot_visible_only: None,
        show_legend,
        data_table: None,
        cat_axis_cross_between: "between".to_string(),
        val_axis_major_tick_mark,
        cat_axis_major_tick_mark,
        title_font_size_hpt: chartex_title_font_size_hpt,
        title_font_color: chartex_title_font_color,
        title_font_paint_authored: chartex_title_text_paint.authored.then_some(true),
        title_font_face: chartex_title_font_face,
        cat_axis_font_size_hpt,
        val_axis_font_size_hpt,
        cat_axis_font_color,
        cat_axis_font_paint_authored: cat_axis_text_paint.authored.then_some(true),
        val_axis_font_color,
        val_axis_font_paint_authored: val_axis_text_paint.authored.then_some(true),
        cat_axis_line_color,
        cat_axis_line_width_emu,
        cat_axis_line_dash,
        cat_axis_line_paint_authored,
        cat_axis_line_hidden,
        val_axis_line_color,
        val_axis_line_width_emu,
        val_axis_line_dash,
        val_axis_line_paint_authored,
        val_axis_line_hidden,
        data_label_font_size_hpt,
        legend_pos,
        bar_gap_width,
        bar_overlap: None,
        data_label_position,
        data_label_font_color,
        data_label_font_paint_authored: data_label_text_paint.authored.then_some(true),
        data_label_format_code: None,
        data_label_font_bold,
        data_label_font_italic,
        data_label_font_language: None,
        data_label_font_baseline: None,
        val_axis_format_code,
        val_axis_number_format: None,
        val_axis_display_units: None,
        cat_axis_display_units: None,
        plot_area_manual_layout: None,
        cartesian_auto_layout_profile: None,
        scatter_style: None,
        bubble_scale: None,
        bubble_size_represents: None,
        show_negative_bubbles: None,
        // chartEx (waterfall/treemap/etc.) has its own axis model. Axis-title
        // text and orientation are shared with the classic renderer model;
        // an explicit chartSpace border remains unwired here.
        cat_axis_title,
        val_axis_title,
        cat_axis_title_font_size_hpt,
        cat_axis_title_font_bold,
        cat_axis_title_font_italic,
        cat_axis_title_font_color,
        cat_axis_title_font_paint_authored: cat_axis_title_text_paint.authored.then_some(true),
        cat_axis_title_rotation,
        cat_axis_title_vertical_mode,
        cat_axis_title_manual_layout,
        cat_axis_title_text_vertical_inset_emu,
        val_axis_title_font_size_hpt,
        val_axis_title_font_bold,
        val_axis_title_font_italic,
        val_axis_title_font_color,
        val_axis_title_font_paint_authored: val_axis_title_text_paint.authored.then_some(true),
        val_axis_title_rotation,
        val_axis_title_vertical_mode,
        val_axis_title_manual_layout,
        val_axis_title_text_vertical_inset_emu,
        title_font_bold: chartex_title_font_bold,
        title_font_italic: chartex_title_font_italic,
        title_font_language: None,
        title_font_baseline: None,
        cat_axis_font_bold,
        cat_axis_font_italic,
        val_axis_font_bold,
        val_axis_font_italic,
        chart_border_color: chart_line_style.color,
        chart_border_line_fill: chart_line_style.fill,
        chart_border_width_emu: chart_line_style.width_emu,
        chart_border_dash: chart_line_style.dash,
        chart_border_dash_authored: chart_line_style.dash_authored,
        chart_border_custom_dash: chart_line_style.custom_dash,
        chart_border_cap: chart_line_style.cap,
        chart_border_join: chart_line_style.join,
        chart_border_compound: chart_line_style.compound,
        chart_border_hidden: chart_line_style.hidden,
        chart_border_paint_authored: chart_line_style.paint_authored,
        secondary_val_axis: None,
        secondary_cat_axis: None,
        // chartEx charts (waterfall/treemap/etc.) are not pie/doughnut and
        // don't carry `<c:txPr>` axis/legend faces; only the theme fallback
        // fonts are threaded so their data labels can pick up the body font.
        hole_size: None,
        first_slice_angle: None,
        cat_axis_font_face,
        val_axis_font_face,
        cat_axis_title_font_face,
        val_axis_title_font_face,
        data_label_font_face,
        legend_font_face,
        legend_font_color,
        legend_font_paint_authored: legend
            .map(|legend| chart_text_body_paint(child(legend, "txPr"), resolver))
            .is_some_and(|paint| paint.authored)
            .then_some(true),
        legend_font_size_hpt,
        legend_font_bold,
        legend_font_italic,
        legend_font_language: None,
        legend_font_baseline: None,
        legend_fill_color: legend_frame.fill_color,
        legend_fill: legend_frame.fill,
        legend_fill_hidden: legend_frame.fill_hidden,
        legend_fill_paint_authored: legend_frame.fill_paint_authored,
        legend_line_color: legend_frame.line_color,
        legend_line_fill: legend_frame.line_fill,
        legend_line_width_emu: legend_frame.line_width_emu,
        legend_line_dash: legend_frame.line_dash,
        legend_line_dash_authored: legend_frame.line_dash_authored,
        legend_line_custom_dash: legend_frame.line_custom_dash,
        legend_line_cap: legend_frame.line_cap,
        legend_line_join: legend_frame.line_join,
        legend_line_compound: legend_frame.line_compound,
        legend_line_hidden: legend_frame.line_hidden,
        legend_line_paint_authored: legend_frame.line_paint_authored,
        theme_major_font_latin: resolver.theme_major_font_latin(),
        theme_minor_font_latin: resolver.theme_minor_font_latin(),
        val_axis_minor_tick_mark: Some(val_axis_minor_tick_mark),
        cat_axis_minor_tick_mark: None,
        legend_manual_layout: None,
        legend_overlay: None,
        legend_entries: None,
        title_manual_layout: None,
        cat_axis_crosses: None,
        cat_axis_crosses_at: None,
        val_axis_crosses: None,
        val_axis_crosses_at: None,
        cat_axis_format_code: None,
        cat_axis_number_format: None,
        cat_axis_min: None,
        cat_axis_max: None,
        radar_style: None,
        // chartEx (cx: namespace) has its own date-axis model; the legacy
        // `<c:date1904>` element does not apply here, so keep the 1900
        // default until/unless a chartEx date system is wired.
        date1904: false,
        // chartEx waterfall has no line/area blanks to display.
        disp_blanks_as: None,
        show_data_labels_over_max: None,
        // chartEx (cx:) has its own axis model (`<cx:axis>`). Shared fields are
        // populated only where CT_ValueAxisScaling has the same semantics as
        // the classic value-axis contract.
        val_axis_major_gridlines,
        cat_axis_major_gridlines: None,
        val_axis_gridline_color,
        val_axis_gridline_width_emu,
        val_axis_gridline_dash,
        val_axis_gridline_paint_authored: None,
        cat_axis_gridline_color: None,
        cat_axis_gridline_width_emu: None,
        cat_axis_gridline_dash: None,
        cat_axis_gridline_paint_authored: None,
        val_axis_minor_gridlines,
        val_axis_minor_gridline_color,
        val_axis_minor_gridline_width_emu,
        val_axis_minor_gridline_dash,
        val_axis_minor_gridline_paint_authored: None,
        cat_axis_minor_gridlines,
        cat_axis_minor_gridline_color,
        cat_axis_minor_gridline_width_emu,
        cat_axis_minor_gridline_dash,
        cat_axis_minor_gridline_paint_authored: None,
        val_axis_major_unit,
        val_axis_minor_unit,
        cat_axis_major_unit: None,
        cat_axis_minor_unit: None,
        cat_axis_is_date: None,
        cat_axis_base_time_unit: None,
        cat_axis_major_time_unit: None,
        cat_axis_minor_time_unit: None,
        cat_axis_no_multi_level_labels: None,
        val_axis_log_base: None,
        cat_axis_log_base: None,
        val_axis_orientation: None,
        cat_axis_orientation: None,
        cat_axis_tick_label_pos: None,
        cat_axis_tick_label_skip: None,
        cat_axis_tick_mark_skip: None,
        cat_axis_label_alignment: None,
        cat_axis_label_offset_percent: None,
        val_axis_tick_label_pos: None,
        cat_axis_label_rotation: None,
        line_group_decorations: None,
        area_group_decorations: None,
        bar_group_decorations: None,
        stock_drop_lines: None,
        stock_hi_low_line_style: None,
        stock_hi_low_lines: None,
        stock_hi_low_line_color: None,
        stock_up_down_bars: None,
        stock_up_down_bar_style: None,
        stock_automatic_style: None,
        surface_wireframe: None,
        surface_band_formats: None,
        classic_surface_band_styles: None,
        legacy_chart_style: None,
        theme_accent_colors: None,
        of_pie: None,
        three_d: None,
        chartex_box,
        chartex_sunburst,
        chartex_treemap,
        chartex_region_map,
        chartex_histogram_binning,
        chartex_accents,
        chart_style_roles,
        classic_chart_style_roles: None,
        classic_varying_point_chart_style_roles: None,
        classic_varying_point_chart_style_roles_by_group: None,
        chart_style_color_palette: chartex_color_palette.clone(),
        chart_style_color_method: chartex_color_style_method.clone(),
        chart_style_marker_size_pt,
        chart_style_marker_symbol,
        chartex_color_palette,
        chartex_color_style_method,
        chartex_data_point_style,
        chartex_data_point_line_style,
        chartex_series_line_style,
        chartex_data_point_marker_style,
        chartex_marker_size_pt,
        chartex_marker_symbol,
        chartex_connector_lines,
    })
}

/// Parse the structured box-and-whisker data of a chartEx `boxWhisker`.
///
/// A box-and-whisker chart has one `<cx:series layoutId="boxWhisker">` per data
/// column; each series' `<cx:dataId val="N">` selects a `<cx:data id="N">`
/// carrying RAW sample points (a `<cx:strDim type="cat">` of per-point category
/// labels and a `<cx:numDim type="val">` of the sample values). This groups
/// each series' points by the unique categories (taken in first-seen order from
/// the first series' data) and threads the `<cx:layoutPr>` visibility /
/// statistics flags. Quartiles / mean / whiskers / outliers are the renderer's
/// job. Returns `None` when there is no plottable series.
pub(super) fn parse_chartex_boxwhisker(
    root: Node,
    resolver: &dyn ColorResolver,
    references: &mut dyn ChartReferenceResolver,
    image_resolver: &dyn ChartImageResolver,
) -> Option<ChartexBoxWhisker> {
    // Build id -> <cx:data> lookup.
    let data_by_id: std::collections::HashMap<String, Node> = root
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "data")
        .filter_map(|d| attr(&d, "id").map(|id| (id, d)))
        .collect();

    // Series nodes, in document order (one column each).
    let all_series_nodes: Vec<Node> = root
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "series")
        .collect();
    let series_nodes: Vec<Node> = all_series_nodes
        .iter()
        .copied()
        .filter(|node| {
            !attr(node, "hidden")
                .is_some_and(|value| value == "1" || value.eq_ignore_ascii_case("true"))
        })
        .collect();
    if series_nodes.is_empty() {
        return None;
    }

    // Per-series raw (category-label, value) points, resolving each series' own
    // <cx:dataId> -> <cx:data>.
    let per_series_points: Vec<Vec<(Option<String>, f64)>> = series_nodes
        .iter()
        .map(|s| {
            let data_id = s
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "dataId")
                .and_then(|n| attr(&n, "val"));
            let data = data_id.as_ref().and_then(|id| data_by_id.get(id).copied());
            match data {
                Some(d) => chartex_data_cat_val_points(d, references),
                None => Vec::new(),
            }
        })
        .collect();

    let series_names: Vec<String> = series_nodes
        .iter()
        .enumerate()
        .map(|(index, series)| {
            series
                .descendants()
                .find(|node| node.is_element() && node.tag_name().name() == "txData")
                .and_then(|tx_data| {
                    child(tx_data, "v")
                        .and_then(|value| value.text())
                        .map(str::trim)
                        .filter(|name| !name.is_empty())
                        .map(ToOwned::to_owned)
                        .or_else(|| {
                            child(tx_data, "f")
                                .and_then(|formula| formula.text())
                                .and_then(|formula| references.resolve_strings(formula))
                                .and_then(|values| {
                                    values.into_iter().find(|value| !value.trim().is_empty())
                                })
                        })
                })
                .unwrap_or_else(|| format!("Series {}", index + 1))
        })
        .collect();

    // Unique categories in first-seen order across all series (first series'
    // order dominates; later series only contribute unseen labels).
    let mut categories: Vec<String> = Vec::new();
    for pts in &per_series_points {
        for (cat, _) in pts {
            if let Some(cat) = cat {
                if !categories.iter().any(|existing| existing == cat) {
                    categories.push(cat.clone());
                }
            }
        }
    }
    if per_series_points.iter().all(Vec::is_empty) {
        return None;
    }
    // Excel's common XLSX form stores one formula-only numeric dimension per
    // named series and no category dimension. Each series is then one box; use
    // the series names as category labels and place its values on the diagonal.
    let one_box_per_series = categories.is_empty();
    if one_box_per_series {
        categories.clone_from(&series_names);
    }
    let cat_index: std::collections::HashMap<&str, usize> = categories
        .iter()
        .enumerate()
        .map(|(i, c)| (c.as_str(), i))
        .collect();

    let series: Vec<ChartexBoxSeries> = series_nodes
        .iter()
        .enumerate()
        .map(|(si, s)| {
            let name = series_names[si].clone();

            // Bin this series' raw points into the shared category order.
            let mut values_by_category: Vec<Vec<f64>> = vec![Vec::new(); categories.len()];
            for (cat, v) in &per_series_points[si] {
                if one_box_per_series {
                    values_by_category[si].push(*v);
                } else if let Some(cat) = cat {
                    if let Some(&ci) = cat_index.get(cat.as_str()) {
                        values_by_category[ci].push(*v);
                    }
                }
            }

            // `<cx:layoutPr><cx:visibility …>` flags; Office defaults when omitted:
            // meanMarker on, meanLine off, outliers and nonoutliers on.
            let vis = s
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "visibility");
            let bool_attr = |name: &str, dflt: bool| {
                vis.and_then(|v| attr(&v, name))
                    .map(|s| s == "1" || s == "true")
                    .unwrap_or(dflt)
            };
            let quartile_method = s
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "statistics")
                .and_then(|st| attr(&st, "quartileMethod"))
                .unwrap_or_else(|| "exclusive".to_string());

            ChartexBoxSeries {
                name,
                chartex_format_idx: attr(s, "formatIdx")
                    .and_then(|value| value.parse::<u32>().ok())
                    .or_else(|| {
                        all_series_nodes
                            .iter()
                            .position(|candidate| candidate == s)
                            .and_then(|index| u32::try_from(index).ok())
                    }),
                color: child(*s, "spPr").and_then(|shape| resolver.resolve_shape_fill(shape)),
                line_color: extract_sp_pr_ln_style(*s, resolver).0,
                line_width_emu: extract_sp_pr_ln_style(*s, resolver).1,
                chartex_style: child(*s, "spPr").map(|_| {
                    parse_chartex_element_style(
                        *s,
                        resolver,
                        None,
                        None,
                        image_resolver,
                        ChartImageSource::Chart,
                    )
                }),
                values_by_category,
                mean_marker: bool_attr("meanMarker", true),
                mean_line: bool_attr("meanLine", false),
                show_outliers: bool_attr("outliers", true),
                show_nonoutliers: bool_attr("nonoutliers", true),
                quartile_method,
            }
        })
        .collect();

    Some(ChartexBoxWhisker {
        one_box_per_series,
        categories,
        series,
    })
}

/// Collect a chartEx data part's aligned category/value samples. Authored
/// `<cx:lvl>` caches win; formula-only dimensions use the package resolver.
/// A missing category dimension is preserved as `None` because Excel uses that
/// form for one-box-per-series charts.
pub(super) fn chartex_data_cat_val_points(
    data: Node,
    references: &mut dyn ChartReferenceResolver,
) -> Vec<(Option<String>, f64)> {
    let categories =
        chartex_string_levels(data, references).and_then(|levels| levels.into_iter().next());
    let values = chartex_number_values(data, &["val", "size"], references).unwrap_or_default();
    values
        .into_iter()
        .enumerate()
        .filter_map(|(index, value)| {
            let value = value?;
            if !value.is_finite() {
                return None;
            }
            let category = categories
                .as_ref()
                .and_then(|items| items.get(index))
                .map(|item| item.trim())
                .filter(|item| !item.is_empty())
                .map(ToOwned::to_owned);
            Some((category, value))
        })
        .collect()
}

/// Parse the structured hierarchy of a chartEx `sunburst`.
///
/// A sunburst's single `<cx:data>` carries a `<cx:strDim type="cat">` with
/// several `<cx:lvl>` (lvl[0] = deepest / Leaf, subsequent lvls step toward the
/// root, last lvl = Branch) and one `<cx:numDim type="size">`. Each data-point
/// `idx` yields a root→leaf `path` (Branch, …, Leaf) with empty trailing
/// segments trimmed — a node that is itself a leaf terminates before the
/// deepest level — and the `size` value at that `idx`. Returns `None` when
/// there is no size dimension or no rows.
pub(super) fn bounded_chartex_point_count(level: Node) -> Option<usize> {
    if let Some(declared) = attr(&level, "ptCount") {
        let count = declared.parse::<usize>().ok()?;
        return (count <= MAX_CHART_CACHE_POINTS).then_some(count);
    }
    let mut count = 0usize;
    for point in level
        .children()
        .filter(|node| node.is_element() && node.tag_name().name() == "pt")
    {
        let index = attr(&point, "idx")?.parse::<usize>().ok()?;
        let required = index.checked_add(1)?;
        if required > MAX_CHART_CACHE_POINTS {
            return None;
        }
        count = count.max(required);
    }
    Some(count)
}

pub(super) fn chartex_string_levels_for_types(
    root: Node,
    dimension_types: &[&str],
    references: &mut dyn ChartReferenceResolver,
) -> Option<Vec<Vec<String>>> {
    let cat_dim = root.descendants().find(|n| {
        n.is_element()
            && n.tag_name().name() == "strDim"
            && attr(n, "type").is_some_and(|kind| dimension_types.contains(&kind.as_str()))
    })?;
    // Levels in document order: lvl[0] = Leaf (deepest), last = Branch (root).
    let levels: Vec<Node> = cat_dim
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "lvl")
        // Hierarchy levels are siblings in XML but become recursive nodes in
        // Sunburst/Treemap layout. Reuse the shared OOXML depth ceiling so a
        // wide sequence of `<cx:lvl>` cannot bypass the parser's stack bound.
        .take(crate::depth::MAX_XML_DEPTH as usize)
        .collect();
    if !levels.is_empty() {
        // Preflight the aggregate slot budget before allocating any level.
        // A per-level cap alone still permits MAX_XML_DEPTH full-width levels.
        let level_counts = levels
            .iter()
            .map(|level| bounded_chartex_point_count(*level))
            .collect::<Option<Vec<_>>>()?;
        let total_slots = level_counts.iter().try_fold(0usize, |total, count| {
            total
                .checked_add(*count)
                .filter(|sum| *sum <= MAX_CHART_CACHE_POINTS)
        })?;
        debug_assert!(total_slots <= MAX_CHART_CACHE_POINTS);
        return Some(
            levels
                .into_iter()
                .zip(level_counts)
                .map(|(level, point_count)| {
                    let mut values = vec![String::new(); point_count];
                    for point in level
                        .children()
                        .filter(|node| node.is_element() && node.tag_name().name() == "pt")
                    {
                        let Some(index) =
                            attr(&point, "idx").and_then(|value| value.parse::<usize>().ok())
                        else {
                            continue;
                        };
                        if index < values.len() {
                            values[index] = point.text().unwrap_or("").replace('\n', " ");
                        }
                    }
                    values
                })
                .collect(),
        );
    }
    let formula = cat_dim
        .children()
        .find(|node| node.is_element() && node.tag_name().name() == "f")
        .and_then(|node| node.text())
        .map(str::trim)
        .filter(|formula| !formula.is_empty())?;
    references.resolve_string_levels(formula).map(|mut levels| {
        levels.truncate(crate::depth::MAX_XML_DEPTH as usize);
        levels
    })
}

pub(super) fn chartex_string_levels(
    root: Node,
    references: &mut dyn ChartReferenceResolver,
) -> Option<Vec<Vec<String>>> {
    chartex_string_levels_for_types(root, &["cat"], references)
}

pub(super) fn chartex_number_values(
    root: Node,
    dimension_types: &[&str],
    references: &mut dyn ChartReferenceResolver,
) -> Option<Vec<Option<f64>>> {
    let dimension = root.descendants().find(|n| {
        n.is_element()
            && n.tag_name().name() == "numDim"
            && attr(n, "type").is_some_and(|kind| dimension_types.contains(&kind.as_str()))
    })?;
    if let Some(level) = dimension
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "lvl")
    {
        let point_count = bounded_chartex_point_count(level)?;
        let mut values = vec![None; point_count];
        for point in level
            .children()
            .filter(|node| node.is_element() && node.tag_name().name() == "pt")
        {
            let Some(index) = attr(&point, "idx").and_then(|value| value.parse::<usize>().ok())
            else {
                continue;
            };
            if index < values.len() {
                values[index] = point.text().and_then(|text| text.parse::<f64>().ok());
            }
        }
        return Some(values);
    }
    let formula = dimension
        .children()
        .find(|node| node.is_element() && node.tag_name().name() == "f")
        .and_then(|node| node.text())
        .map(str::trim)
        .filter(|formula| !formula.is_empty())?;
    references.resolve_numbers(formula)
}

pub(super) fn chartex_number_format(
    root: Node,
    dimension_types: &[&str],
    references: &mut dyn ChartReferenceResolver,
) -> Option<String> {
    let dimension = root.descendants().find(|n| {
        n.is_element()
            && n.tag_name().name() == "numDim"
            && attr(n, "type").is_some_and(|kind| dimension_types.contains(&kind.as_str()))
    })?;
    let formula = dimension
        .children()
        .find(|node| node.is_element() && node.tag_name().name() == "f")
        .and_then(|node| node.text())
        .map(str::trim)
        .filter(|formula| !formula.is_empty())?;
    references.resolve_number_format(formula)
}

pub(super) fn parse_chartex_hierarchy_rows(
    root: Node,
    references: &mut dyn ChartReferenceResolver,
) -> Option<Vec<ChartexSunburstRow>> {
    let levels = chartex_string_levels(root, references)?;
    let sizes = chartex_number_values(root, &["size", "val"], references)?;
    let n = sizes
        .len()
        .max(levels.iter().map(Vec::len).max().unwrap_or(0));

    let mut rows: Vec<ChartexSunburstRow> = Vec::new();
    for idx in 0..n {
        let size = sizes.get(idx).copied().flatten().unwrap_or(0.0);
        // Build path root→leaf: iterate levels from LAST (Branch/root) to FIRST
        // (Leaf/deepest). Trailing empty leaf cells are trimmed so a node that is
        // itself a leaf terminates early.
        let mut path: Vec<String> = Vec::new();
        for level in levels.iter().rev() {
            let label = level.get(idx).cloned().unwrap_or_default();
            if label.is_empty() {
                break;
            }
            path.push(label);
        }
        if path.is_empty() {
            continue;
        }
        rows.push(ChartexSunburstRow { path, size });
    }
    if rows.is_empty() {
        return None;
    }
    Some(rows)
}

pub(super) fn parse_chartex_value_color_stop(
    parent: Node,
    name: &str,
) -> Option<ChartexValueColorStop> {
    let stop = child(parent, name)?;
    for kind in ["extremeValue", "number", "percent"] {
        if let Some(value_node) = child(stop, kind) {
            let value = (kind != "extremeValue")
                .then(|| attr(&value_node, "val").and_then(|value| value.parse::<f64>().ok()))
                .flatten()
                .filter(|value| value.is_finite());
            return Some(ChartexValueColorStop {
                kind: kind.to_string(),
                value,
            });
        }
    }
    None
}

pub(super) fn parse_chartex_region_map(
    data: Node,
    series: Node,
    resolver: &dyn ColorResolver,
    references: &mut dyn ChartReferenceResolver,
) -> Option<ChartexRegionMap> {
    let labels = chartex_string_levels_for_types(data, &["cat"], references)
        .and_then(|levels| levels.into_iter().next())
        .unwrap_or_default();
    let entity_ids = chartex_string_levels_for_types(data, &["entityId"], references)
        .and_then(|levels| levels.into_iter().next())
        .unwrap_or_default();
    let values = chartex_number_values(data, &["colorVal"], references).unwrap_or_default();
    let row_count = labels.len().max(entity_ids.len()).max(values.len());
    if row_count == 0 || row_count > MAX_CHART_CACHE_POINTS {
        return None;
    }
    let rows = (0..row_count)
        .map(|index| ChartexRegionMapRow {
            label: labels.get(index).cloned().unwrap_or_default(),
            entity_id: entity_ids
                .get(index)
                .filter(|value| !value.trim().is_empty())
                .cloned(),
            value: values
                .get(index)
                .copied()
                .flatten()
                .filter(|value| value.is_finite()),
        })
        .collect();

    let layout = child(series, "layoutPr");
    let region_label_layout = layout
        .and_then(|node| child(node, "regionLabelLayout"))
        .and_then(|node| attr(&node, "val"))
        .filter(|value| matches!(value.as_str(), "none" | "bestFitOnly" | "showAll"));
    let geography = layout
        .and_then(|node| child(node, "geography"))
        .map(|node| {
            let cache = child(node, "geoCache");
            ChartexGeography {
                projection_type: attr(&node, "projectionType").filter(|value| {
                    matches!(
                        value.as_str(),
                        "mercator" | "miller" | "robinson" | "albers"
                    )
                }),
                viewed_region_type: attr(&node, "viewedRegionType").filter(|value| {
                    matches!(
                        value.as_str(),
                        "dataOnly"
                            | "postalCode"
                            | "county"
                            | "state"
                            | "countryRegion"
                            | "countryRegionList"
                            | "world"
                    )
                }),
                culture_language: attr(&node, "cultureLanguage"),
                culture_region: attr(&node, "cultureRegion"),
                attribution: attr(&node, "attribution"),
                cache_provider: cache.and_then(|value| attr(&value, "provider")),
                cache_present: cache.is_some(),
            }
        });

    let value_colors = child(series, "valueColors");
    let value_positions = child(series, "valueColorPositions");
    let colors = if value_colors.is_some() || value_positions.is_some() {
        Some(ChartexRegionMapColors {
            stop_count: value_positions
                .and_then(|node| attr(&node, "count"))
                .and_then(|value| value.parse::<u8>().ok())
                .filter(|value| matches!(value, 2 | 3)),
            min_color: value_colors
                .and_then(|node| child(node, "minColor"))
                .and_then(|node| resolver.resolve_solid_fill(node)),
            mid_color: value_colors
                .and_then(|node| child(node, "midColor"))
                .and_then(|node| resolver.resolve_solid_fill(node)),
            max_color: value_colors
                .and_then(|node| child(node, "maxColor"))
                .and_then(|node| resolver.resolve_solid_fill(node)),
            min_position: value_positions
                .and_then(|node| parse_chartex_value_color_stop(node, "min")),
            mid_position: value_positions
                .and_then(|node| parse_chartex_value_color_stop(node, "mid")),
            max_position: value_positions
                .and_then(|node| parse_chartex_value_color_stop(node, "max")),
        })
    } else {
        None
    };

    Some(ChartexRegionMap {
        rows,
        region_label_layout,
        geography,
        colors,
    })
}

pub(super) fn parse_chartex_sunburst(
    data: Node,
    references: &mut dyn ChartReferenceResolver,
) -> Option<ChartexSunburst> {
    parse_chartex_hierarchy_rows(data, references).map(|rows| ChartexSunburst { rows })
}

/// Parse a chartEx treemap. Its category and size dimensions use the same
/// hierarchy representation as sunburst; `parentLabelLayout` only affects how
/// parent captions are painted.
pub(super) fn parse_chartex_treemap(
    data: Node,
    series: Node,
    references: &mut dyn ChartReferenceResolver,
) -> Option<ChartexTreemap> {
    let rows = parse_chartex_hierarchy_rows(data, references)?;
    let parent_label_layout = series
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "parentLabelLayout")
        .and_then(|layout| attr(&layout, "val"));
    Some(ChartexTreemap {
        rows,
        parent_label_layout,
    })
}
