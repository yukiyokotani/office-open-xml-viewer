use super::*;

pub(super) fn parse_classic_impl(
    chart_root: Node,
    color_resolver: &dyn ColorResolver,
    style_xml: Option<&str>,
    color_style_xml: Option<&str>,
    references: &mut dyn ChartReferenceResolver,
    image_resolver: &dyn ChartImageResolver,
) -> Option<ChartModel> {
    let root = chart_root;
    let plot_visible_only = child(root, "chart").and_then(|chart| bool_child(chart, "plotVisOnly"));
    let legacy_chart_style = child(root, "style")
        .and_then(|style| style.attribute("val"))
        .and_then(|value| value.parse::<u8>().ok())
        .filter(|value| (1..=48).contains(value));
    let mapped_resolver =
        ChartColorMapping::from_chart_space(chart_root).map(|mapping| ChartMappedColorResolver {
            base: color_resolver,
            mapping,
        });
    let color_resolver: &dyn ColorResolver = mapped_resolver
        .as_ref()
        .map(|resolver| resolver as &dyn ColorResolver)
        .unwrap_or(color_resolver);
    let theme_accent_colors = {
        let colors = (0..6)
            .filter_map(|index| color_resolver.resolve_series_accent(index))
            .collect::<Vec<_>>();
        (colors.len() == 6).then_some(colors)
    };
    let style_doc = style_xml.and_then(|xml| crate::depth::parse_guarded(xml).ok());
    let color_style = color_style_xml.and_then(|xml| parse_chart_color_style(xml, color_resolver));
    let unresolved_color_style_palette =
        (color_style_xml.is_some() && color_style.is_none()).then(|| vec![None]);
    let chart_style_color_method = color_style.as_ref().map(|(method, _)| method.clone());
    let chart_style_color_palette = color_style
        .as_ref()
        .map(|(_, palette)| palette.clone())
        .or_else(|| unresolved_color_style_palette.clone());
    let theme_style_palette = style_doc.as_ref().and_then(|_| {
        let colors = (0..6)
            .map(|index| color_resolver.resolve_series_accent(index))
            .collect::<Vec<_>>();
        colors.iter().any(Option::is_some).then_some(colors)
    });
    let style_palette = chart_style_color_palette
        .as_deref()
        .or(theme_style_palette.as_deref());
    let chart_style_roles = if style_xml.is_some() && style_doc.is_none() {
        Some(unreadable_chart_style_role_table())
    } else {
        style_doc.as_ref().and_then(|document| {
            parse_chart_style_role_table(
                document.root_element(),
                color_resolver,
                style_palette,
                chart_style_color_method.as_deref(),
                image_resolver,
            )
            .or_else(|| {
                unresolved_chart_style_role_table(
                    document.root_element(),
                    color_resolver,
                    image_resolver,
                )
            })
        })
    };
    let marker_layout = style_doc
        .as_ref()
        .and_then(|document| child(document.root_element(), "dataPointMarkerLayout"));
    let chart_style_marker_size_pt = marker_layout
        .and_then(|node| node.attribute("size"))
        .and_then(|value| value.parse::<u8>().ok())
        .filter(|value| (2..=72).contains(value));
    let chart_style_marker_symbol = marker_layout
        .and_then(|node| node.attribute("symbol"))
        .map(ToOwned::to_owned);

    // Determine chart type by finding the first recognized chart element
    let find_chart = |name: &str| {
        root.descendants()
            .find(|n| n.is_element() && n.tag_name().name() == name)
    };
    let is_classic_group_name = |name: &str| {
        matches!(
            name,
            "areaChart"
                | "area3DChart"
                | "lineChart"
                | "line3DChart"
                | "stockChart"
                | "radarChart"
                | "scatterChart"
                | "pieChart"
                | "pie3DChart"
                | "doughnutChart"
                | "barChart"
                | "bar3DChart"
                | "ofPieChart"
                | "surfaceChart"
                | "surface3DChart"
                | "bubbleChart"
        )
    };
    let has_nonempty_classic_group = root.descendants().any(|node| {
        node.is_element()
            && is_classic_group_name(node.tag_name().name())
            && child(node, "ser").is_some()
    });
    // Empty CT_PlotArea groups are schema-valid and remain in `plot_groups`,
    // but they do not own the flattened compatibility series. Prefer a group
    // that actually has a `ser` when selecting the legacy chart-family
    // projection so an empty earlier family cannot reinterpret the sole
    // visible group. All-empty charts retain the historical first-match
    // projection and render as no data.
    let find_render_chart = |name: &str| {
        root.descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == name)
            .find(|group| child(*group, "ser").is_some())
            .or_else(|| {
                (!has_nonempty_classic_group)
                    .then(|| find_chart(name))
                    .flatten()
            })
    };

    // ECMA-376 3D chart types (§21.2.2.15 bar3DChart, §21.2.2.96 line3DChart,
    // §21.2.2.4 area3DChart, §21.2.2.140 pie3DChart) share the ordinary 2D
    // SERIES model (`<c:ser>`/`<c:cat>`/`<c:val>`/grouping/labels). Keep those
    // canonical renderer family names, while `ChartThreeD` retains view/depth
    // controls for a bounded compatibility projection. Direct CT_Surface
    // floor/wall paint is retained below; scene lighting remains a renderer
    // concern rather than being guessed in the parser.
    // Surface charts retain their matrix semantics in a dedicated shared
    // renderer; unlike ordinary 3-D chart families they cannot be flattened to
    // a line/area family without losing the value-band topology.
    let read_grouping = |group: &Node, default: &str| -> String {
        group
            .children()
            .find(|c| c.is_element() && c.tag_name().name() == "grouping")
            .and_then(|n| attr(&n, "val"))
            .unwrap_or_else(|| default.into())
    };
    let chart_type = if let Some(bc) =
        find_render_chart("barChart").or_else(|| find_render_chart("bar3DChart"))
    {
        // §21.2.2.17 barDir + §21.2.2.77 grouping (Bar Grouping). bar3DChart shares both
        // and retains its extra `<c:gapDepth>` in the shared 3-D scene model.
        // `clustered` is the 2D default;
        // `standard` (the bar3DChart default) folds to clustered as well since
        // `canonical_chart_type` treats any non-stacked grouping as clustered.
        let grouping = read_grouping(&bc, "clustered");
        let bar_dir = bc
            .children()
            .find(|c| c.is_element() && c.tag_name().name() == "barDir")
            .and_then(|n| attr(&n, "val"))
            .unwrap_or_else(|| "col".into());
        canonical_chart_type("bar", &bar_dir, &grouping)
    } else if let Some(ac) =
        find_render_chart("areaChart").or_else(|| find_render_chart("area3DChart"))
    {
        // An area+line combination must dispatch through the area renderer so
        // the fill-bearing group is stacked before the line overlay. Selecting
        // line merely because a `<c:lineChart>` is also present discards every
        // authored area fill. Per-series `series_type` keeps both groups.
        let grouping = read_grouping(&ac, "standard");
        canonical_chart_type("area", "col", &grouping)
    } else if let Some(lc) =
        find_render_chart("lineChart").or_else(|| find_render_chart("line3DChart"))
    {
        let grouping = read_grouping(&lc, "standard");
        canonical_chart_type("line", "col", &grouping)
    } else if find_render_chart("pieChart").is_some() || find_render_chart("pie3DChart").is_some() {
        "pie".to_string()
    } else if find_render_chart("ofPieChart").is_some() {
        // Keep the family distinct: §21.2.2.126 assigns a secondary pie/bar and
        // connector geometry that cannot be represented by a plain pie.
        "ofPie".to_string()
    } else if find_render_chart("doughnutChart").is_some() {
        "doughnut".to_string()
    } else if find_render_chart("scatterChart").is_some() {
        "scatter".to_string()
    } else if find_render_chart("bubbleChart").is_some() {
        "bubble".to_string()
    } else if find_render_chart("radarChart").is_some() {
        "radar".to_string()
    } else if find_render_chart("stockChart").is_some() {
        // §21.2.2.198 stockChart — high/low/close[/open] series drawn as
        // per-category hi-lo lines + close ticks by the core stock renderer.
        "stock".to_string()
    } else if find_render_chart("surfaceChart").is_some() {
        "surface".to_string()
    } else if find_render_chart("surface3DChart").is_some() {
        // A 3-D surface is not a top-down contour chart. Keep it distinct so
        // the 2-D renderer cannot silently flatten its camera/depth geometry.
        "surface3D".to_string()
    } else {
        "unknown".to_string()
    };

    // §21.2.2.198 stockChart decorations: `<c:dropLines>`, `<c:hiLowLines>`
    // (§21.2.2.80), and `<c:upDownBars>` (§21.2.2.218). All are direct
    // children of `<c:stockChart>`.
    // The hi-lo line spans each category's low↔high; its `<c:spPr><a:ln>` fill is
    // resolved so the renderer strokes it in the file's color (else a gray
    // default). Up/down gap and direct bar paint are retained for the stock
    // renderer. Every field stays `None` for non-stock charts (byte-stable wire).
    let (
        stock_drop_lines,
        stock_hi_low_line_style,
        stock_hi_low_lines,
        stock_hi_low_line_color,
        stock_up_down_bars,
        stock_up_down_bar_style,
    ) = if let Some(stock) = find_chart("stockChart") {
        let drop_lines = child(stock, "dropLines")
            .map(|node| parse_chart_decoration_line_style(node, color_resolver));
        let hi_low = child(stock, "hiLowLines");
        let hi_low_style =
            hi_low.map(|node| parse_chart_decoration_line_style(node, color_resolver));
        let hi_low_color = hi_low_style.as_ref().and_then(|style| style.color.clone());
        let up_down_node = child(stock, "upDownBars");
        let up_down = up_down_node.is_some();
        let up_down_style =
            up_down_node.map(|node| parse_chart_up_down_bar_style(node, color_resolver));
        (
            drop_lines,
            hi_low_style,
            Some(hi_low.is_some()),
            hi_low_color,
            if up_down { Some(true) } else { None },
            up_down_style,
        )
    } else {
        (None, None, None, None, None, None)
    };
    // Empty stock-decoration paint is application-defined. Retain only the
    // bounded Office rules observed across 3/4-series stock charts, omitted
    // and explicit legacy styles, direct/noFill controls, and a substituted
    // dark-1 theme. Styles outside this observed set stay unresolved instead
    // of receiving a guessed palette.
    let stock_has_decoration = stock_drop_lines.is_some()
        || stock_hi_low_lines == Some(true)
        || stock_up_down_bars == Some(true);
    let stock_automatic_style = if find_chart("stockChart").is_some() && stock_has_decoration {
        let tint_amounts = match legacy_chart_style.unwrap_or(2) {
            1 => Some((0.75, 0.15)),
            2 | 10 => Some((0.95, 0.05)),
            _ => None,
        };
        tint_amounts.and_then(|(up_tint, down_tint)| {
            let dark = color_resolver.resolve_scheme_color("dk1")?;
            Some(ChartStockAutomaticStyle {
                line_color: dark.clone(),
                line_width_emu: 12_700,
                up_fill_color: crate::color::apply_signed_tint_or_shade(
                    &dark,
                    up_tint,
                    color_resolver.tint_mode(),
                ),
                down_fill_color: crate::color::apply_signed_tint_or_shade(
                    &dark,
                    down_tint,
                    color_resolver.tint_mode(),
                ),
            })
        })
    } else {
        None
    };
    let surface_group = find_chart("surfaceChart").or_else(|| find_chart("surface3DChart"));
    let surface_wireframe = surface_group.and_then(|surface| bool_child(surface, "wireframe"));
    let surface_band_format_nodes = surface_group
        .and_then(|surface| child(surface, "bandFmts"))
        .map(|formats| {
            formats
                .children()
                .filter(|node| node.is_element() && node.tag_name().name() == "bandFmt")
                .take(MAX_CHART_COLOR_STYLE_ENTRIES + 1)
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    if surface_band_format_nodes.len() > MAX_CHART_COLOR_STYLE_ENTRIES {
        return None;
    }
    // A band is one logical formatted datum even though the renderer clips it
    // into many polygons. Bound each direct fill/outline recipe and their
    // chart-wide aggregate before parsing any gradient stop list.
    let mut surface_paint_components = 0usize;
    for format in &surface_band_format_nodes {
        let Some(sp_pr) = child(*format, "spPr") else {
            continue;
        };
        for components in [
            chart_style_paint_component_count(sp_pr).unwrap_or(0),
            child(sp_pr, "ln")
                .and_then(chart_style_paint_component_count)
                .unwrap_or(0),
        ] {
            if components > MAX_CHART_MARKER_GRADIENT_STOPS
                || components > MAX_CHART_MARKER_PAINT_COMPONENTS - surface_paint_components
            {
                return None;
            }
            surface_paint_components += components;
        }
    }
    let surface_band_formats = (!surface_band_format_nodes.is_empty())
        .then(|| {
            surface_band_format_nodes
                .iter()
                .filter_map(|format| {
                    let idx = child(*format, "idx")
                        .and_then(|node| node.attribute("val"))
                        .and_then(|value| value.parse::<u32>().ok())?;
                    let shape = child(*format, "spPr");
                    let mut style = shape.map(|_| {
                        parse_chartex_element_style(
                            *format,
                            color_resolver,
                            None,
                            None,
                            image_resolver,
                            ChartImageSource::Chart,
                        )
                    });
                    // The legacy `fill` field remains the single authoritative
                    // carrier for direct band paint. Derive it from the one
                    // parsed style entry, then retain only provenance in the
                    // full style so a gradient is neither expanded nor sent
                    // over the wire twice.
                    let (fill, fill_hidden) = style
                        .as_ref()
                        .map(|style| {
                            let fill = style
                                .fill_paints
                                .as_deref()
                                .and_then(|paints| paints.first())
                                .and_then(Option::as_ref)
                                .cloned();
                            (fill, (style.fill_hidden == Some(true)).then_some(true))
                        })
                        .unwrap_or((None, None));
                    if let Some(style) = style.as_mut() {
                        style.fill_paints = None;
                        style.fill_colors = None;
                        style.fill_hidden = None;
                        style.fill_no_style = None;
                    }
                    let (line_color, line_width_emu, line_hidden) =
                        extract_sp_pr_ln_style(*format, color_resolver);
                    Some(ChartSurfaceBandFormat {
                        idx,
                        style,
                        fill,
                        fill_hidden,
                        line_color,
                        line_width_emu,
                        line_hidden: line_hidden.then_some(true),
                    })
                })
                .collect::<Vec<_>>()
        })
        .filter(|formats| !formats.is_empty());

    let of_pie = find_chart("ofPieChart").map(|node| {
        let percent = |name: &str, default: f64| {
            child(node, name)
                .and_then(|value| value.attribute("val"))
                .and_then(|value| value.trim_end_matches('%').parse::<f64>().ok())
                .filter(|value| value.is_finite())
                .unwrap_or(default)
        };
        let split_type_node = child(node, "splitType");
        let split_type = split_type_node
            .and_then(|value| value.attribute("val"))
            .filter(|value| matches!(*value, "auto" | "cust" | "percent" | "pos" | "val"))
            .unwrap_or("auto")
            .to_string();
        let split_pos_node = child(node, "splitPos");
        ChartOfPie {
            r#type: child(node, "ofPieType")
                .and_then(|value| value.attribute("val"))
                .filter(|value| matches!(*value, "pie" | "bar"))
                .unwrap_or("pie")
                .to_string(),
            split_type,
            split_type_authored: split_type_node.is_some(),
            split_pos: split_pos_node
                .and_then(|value| value.attribute("val"))
                .and_then(|value| value.parse::<f64>().ok())
                .filter(|value| value.is_finite()),
            split_pos_authored: split_pos_node.is_some(),
            custom_split_indices: parse_of_pie_custom_split_with_limit(
                node,
                MAX_CHART_CACHE_POINTS,
            ),
            second_pie_size_percent: percent("secondPieSize", 75.0).clamp(5.0, 200.0),
            gap_width_percent: percent("gapWidth", 150.0).max(0.0),
            series_lines: child(node, "serLines").is_some(),
            series_line_style: child(node, "serLines")
                .map(|line| parse_chart_decoration_line_style(line, color_resolver)),
        }
    });

    let three_d_group = find_chart("bar3DChart")
        .or_else(|| find_chart("line3DChart"))
        .or_else(|| find_chart("area3DChart"))
        .or_else(|| find_chart("pie3DChart"))
        .or_else(|| find_chart("surfaceChart"))
        .or_else(|| find_chart("surface3DChart"));
    let three_d_series_axis = root
        .descendants()
        .find(|node| node.is_element() && node.tag_name().name() == "serAx")
        .map(|axis| {
            let (title, title_font_size_hpt, title_font_bold, title_font_color) =
                extract_axis_title_with_props_resolved(axis, color_resolver);
            let (line_color, line_width_emu, line_hidden) =
                extract_axis_line_style(axis, color_resolver);
            let line_dash = extract_axis_line_dash(axis);
            let positive_u32 = |name: &str| {
                child(axis, name)
                    .and_then(|node| node.attribute("val"))
                    .and_then(|value| value.parse::<u32>().ok())
                    .filter(|value| *value > 0)
            };
            ChartThreeDSeriesAxis {
                style: parse_direct_chart_effect_style(axis, color_resolver),
                title_style: child(axis, "title")
                    .and_then(|node| parse_direct_chart_effect_style(node, color_resolver)),
                major_gridline_style: child(axis, "majorGridlines")
                    .and_then(|node| parse_direct_chart_effect_style(node, color_resolver)),
                minor_gridline_style: child(axis, "minorGridlines")
                    .and_then(|node| parse_direct_chart_effect_style(node, color_resolver)),
                title,
                hidden: axis_is_deleted(axis),
                orientation: extract_axis_orientation(axis),
                tick_label_pos: extract_axis_tick_label_pos(axis),
                tick_label_skip: positive_u32("tickLblSkip"),
                tick_mark_skip: positive_u32("tickMarkSkip"),
                major_tick_mark: extract_axis_tick_mark_or_default(axis, "majorTickMark"),
                minor_tick_mark: extract_axis_tick_mark(axis, "minorTickMark"),
                font_color: extract_axis_tick_label_color(axis, color_resolver),
                font_paint_authored: chart_text_body_paint(child(axis, "txPr"), color_resolver)
                    .authored
                    .then_some(true),
                font_size_hpt: extract_axis_tick_label_size(axis),
                font_bold: extract_axis_tick_label_bold(axis),
                font_italic: extract_axis_tick_label_italic(axis),
                font_face: extract_axis_tick_label_face(axis),
                line_color,
                line_width_emu,
                line_dash,
                line_paint_authored: extract_direct_shape_line(axis, color_resolver).paint_authored,
                line_hidden,
                title_font_size_hpt,
                title_font_bold,
                title_font_italic: extract_axis_title_italic(axis),
                title_font_color,
                title_font_paint_authored: chart_title_text_paint(Some(axis), color_resolver)
                    .authored
                    .then_some(true),
                title_font_face: extract_axis_title_face(axis),
                title_rotation: extract_axis_title_rotation(axis),
                title_vertical_mode: extract_axis_title_vertical_mode(axis),
                title_manual_layout: extract_axis_title_manual_layout(axis),
            }
        });
    // ECMA-376 §21.2.2.69/.191/.11 and CT_Surface retain independent
    // DrawingML paint for the floor, side wall and back wall. These faces are
    // part of the chart's authored 3-D box; dropping them forces the renderer
    // to invent fills and loses wall boundary rules present in the package.
    let chart_container = root
        .descendants()
        .find(|node| node.is_element() && node.tag_name().name() == "chart");
    // CT_Surface has only three instances, but each DrawingML gradient can
    // contain an unbounded gsLst. Preflight every direct floor/wall fill and
    // outline before expanding any recipe, using the same per-recipe and
    // chart-wide ceilings as other repeated chart paints.
    for surface in chart_container.into_iter().flat_map(|chart| {
        chart.children().filter(|node| {
            node.is_element() && matches!(node.tag_name().name(), "floor" | "sideWall" | "backWall")
        })
    }) {
        let Some(sp_pr) = child(surface, "spPr") else {
            continue;
        };
        for components in [
            chart_style_paint_component_count(sp_pr).unwrap_or(0),
            child(sp_pr, "ln")
                .and_then(chart_style_paint_component_count)
                .unwrap_or(0),
        ] {
            if components > MAX_CHART_MARKER_GRADIENT_STOPS
                || components > MAX_CHART_MARKER_PAINT_COMPONENTS - surface_paint_components
            {
                return None;
            }
            surface_paint_components += components;
        }
    }
    let parse_three_d_surface = |name: &str| -> Option<ChartThreeDSurface> {
        let surface = chart_container.and_then(|chart| child(chart, name))?;
        let sp_pr = child(surface, "spPr");
        let style = sp_pr.map(|_| {
            parse_chartex_element_style(
                surface,
                color_resolver,
                None,
                None,
                image_resolver,
                ChartImageSource::Chart,
            )
        });
        let fill_color = sp_pr.and_then(|shape| color_resolver.resolve_shape_fill(shape));
        let fill_hidden = sp_pr
            .and_then(|shape| child(shape, "noFill"))
            .map(|_| true)
            .or_else(|| fill_color.as_ref().map(|_| false));
        let (line_color, line_width_emu, line_no_fill) =
            extract_sp_pr_ln_style(surface, color_resolver);
        let line_dash = sp_pr
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
        let thickness_percent = child(surface, "thickness")
            .and_then(|node| node.attribute("val"))
            .and_then(|value| {
                parse_chart_integer_percent(
                    value,
                    surface.tag_name().namespace() == Some(crate::ns::chart::STRICT),
                    0,
                    u32::MAX,
                )
            });
        let picture_options =
            child(surface, "pictureOptions").map(|options| ChartThreeDPictureOptions {
                apply_to_front: strict_boolean_child(options, "applyToFront"),
                apply_to_sides: strict_boolean_child(options, "applyToSides"),
                apply_to_end: strict_boolean_child(options, "applyToEnd"),
                picture_format: child(options, "pictureFormat")
                    .and_then(|node| node.attribute("val"))
                    .filter(|value| matches!(*value, "stretch" | "stack" | "stackScale"))
                    .map(str::to_string),
                picture_format_authored: child(options, "pictureFormat").map(|_| true),
                picture_stack_unit: child(options, "pictureStackUnit")
                    .and_then(|node| node.attribute("val"))
                    .and_then(|value| value.trim().parse::<f64>().ok())
                    .filter(|value| value.is_finite() && *value > 0.0),
                picture_stack_unit_authored: child(options, "pictureStackUnit").map(|_| true),
            });
        Some(ChartThreeDSurface {
            style,
            fill_color,
            fill_hidden,
            line_color,
            line_width_emu,
            line_dash,
            line_hidden,
            thickness_percent,
            picture_options,
        })
    };
    let three_d = three_d_group.map(|group| {
        let view = root
            .descendants()
            .find(|node| node.is_element() && node.tag_name().name() == "view3D");
        let strict_chart = root.tag_name().namespace() == Some(crate::ns::chart::STRICT);
        let scalar_with_schema_default =
            |parent: Option<Node>, name: &str, default: i32, min: i32, max: i32| {
                let node = parent.and_then(|parent| child(parent, name))?;
                match node.attribute("val") {
                    Some(value) => value
                        .trim()
                        .parse::<i32>()
                        .ok()
                        .filter(|value| (min..=max).contains(value)),
                    None => Some(default),
                }
            };
        let percent_with_schema_default =
            |parent: Option<Node>, name: &str, default: f64, min: u32, max: u32| {
                let node = parent.and_then(|parent| child(parent, name))?;
                match node.attribute("val") {
                    Some(value) => parse_chart_integer_percent(value, strict_chart, min, max),
                    None => Some(default),
                }
            };
        ChartThreeD {
            view_3d_present: Some(view.is_some()),
            // CT_View3D child values have schema defaults even when their
            // element is present without @val. Preserve that authored/bare
            // distinction instead of replacing it with a renderer default.
            rotation_x: scalar_with_schema_default(view, "rotX", 0, -90, 90),
            rotation_x_authored: Some(view.is_some_and(|node| child(node, "rotX").is_some())),
            rotation_y: scalar_with_schema_default(view, "rotY", 0, 0, 360)
                .and_then(|value| u32::try_from(value).ok()),
            rotation_y_authored: Some(view.is_some_and(|node| child(node, "rotY").is_some())),
            height_percent: percent_with_schema_default(view, "hPercent", 100.0, 5, 500),
            height_percent_authored: Some(
                view.is_some_and(|node| child(node, "hPercent").is_some()),
            ),
            depth_percent: percent_with_schema_default(view, "depthPercent", 100.0, 20, 2000),
            depth_percent_authored: Some(
                view.is_some_and(|node| child(node, "depthPercent").is_some()),
            ),
            perspective: scalar_with_schema_default(view, "perspective", 30, 0, 240)
                .and_then(|value| u32::try_from(value).ok()),
            perspective_authored: Some(
                view.is_some_and(|node| child(node, "perspective").is_some()),
            ),
            right_angle_axes: view.and_then(|node| bool_child(node, "rAngAx")),
            right_angle_axes_authored: Some(
                view.is_some_and(|node| child(node, "rAngAx").is_some()),
            ),
            gap_depth_percent: percent_with_schema_default(Some(group), "gapDepth", 150.0, 0, 500),
            gap_depth_percent_authored: Some(child(group, "gapDepth").is_some()),
            shape: child(group, "shape")
                .and_then(|node| node.attribute("val"))
                .map(str::to_string),
            bar_grouping: (group.tag_name().name() == "bar3DChart").then(|| {
                child(group, "grouping")
                    .and_then(|node| node.attribute("val"))
                    .unwrap_or("standard")
                    .to_string()
            }),
            series_axis: three_d_series_axis,
            floor: parse_three_d_surface("floor"),
            side_wall: parse_three_d_surface("sideWall"),
            back_wall: parse_three_d_surface("backWall"),
        }
    });

    // Title text. The CHART title is the direct-child `<c:title>` of `<c:chart>`
    // (ECMA-376 §21.2.2.210) — NOT any `<c:title>` descendant. A `descendants()`
    // search would pick up the first AXIS title (which lives inside `<c:plotArea>`
    // → `<c:valAx>`/`<c:catAx>`) on a chart that has axis titles but no chart
    // title, wrongly promoting it to the chart title. Scope strictly to the
    // `<c:chart>` element's own `<c:title>` child.
    let chart_node = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "chart");
    let title_node_opt = chart_node.and_then(|c| child(c, "title"));
    // CT_Tx allows either DrawingML rich text (`<a:t>`) or a string-reference
    // cache (`<c:strRef><c:strCache><c:v>`). Reuse the scoped title helper so
    // both legal forms are honored; looking only for `<a:t>` incorrectly made
    // a cached authored title disappear and triggered the series-name auto
    // title below.
    let mut title = chart_node.and_then(extract_chart_title_text);
    let title_rich_runs = title_node_opt
        .and_then(|title_node| parse_chart_title_rich_runs(title_node, color_resolver));
    // Title font size in hundredths of a point — taken from the first
    // defRPr@sz or rPr@sz we find inside the title. ECMA-376 uses hpt for size.
    let title_font_size_hpt = title_node_opt.and_then(|t| {
        t.descendants().find_map(|n| {
            if !n.is_element() {
                return None;
            }
            let tag = n.tag_name().name();
            if tag != "defRPr" && tag != "rPr" {
                return None;
            }
            attr(&n, "sz").and_then(|v| parse_text_font_size_hpt(&v))
        })
    });
    // Title font color — resolved via the `ColorResolver` so a `<a:schemeClr>`
    // (e.g. `tx2` → the theme dark-2 slot) resolves in addition to a literal
    // `<a:srgbClr>`. `extract_chart_title_color` scopes to the direct-child
    // `<c:title>` of the node it's given, so pass `title_node_opt`'s parent (the
    // element that holds `<c:title>`). Previously hardcoded `None` (the srgb was
    // never threaded into the wire model); resolving it fixes titles that use a
    // theme scheme color, which Office decks commonly do.
    let title_font_color = title_node_opt
        .and_then(|t| t.parent())
        .and_then(|parent| extract_chart_title_color(parent, color_resolver));

    // val axis max / min and visibility — shared helpers in ooxml-common
    // so xlsx & pptx stay in sync (`<c:scaling><c:min|max val>` §21.2.2.160
    // scaling and `<c:delete val>` ECMA-376 §21.2.2.40 delete).
    // Combo charts (bar + line) declare TWO `<c:valAx>`: a PRIMARY (axPos="l",
    // `<c:crosses val="autoZero">`) and a SECONDARY (axPos="r",
    // `<c:crosses val="max">`). Collect both so series can be mapped to the
    // right scale and the right-hand axis drawn. The primary axis keeps driving
    // every existing axis read below; only the secondary is new.
    let val_ax_nodes: Vec<roxmltree::Node> = root
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "valAx")
        .take(MAX_CHART_PLOT_AXES + 1)
        .collect();
    if val_ax_nodes.len() > MAX_CHART_PLOT_AXES {
        return None;
    }
    let ax_pos = |n: &roxmltree::Node| -> Option<String> {
        n.children()
            .find(|c| c.is_element() && c.tag_name().name() == "axPos")
            .and_then(|c| attr(&c, "val"))
    };
    let ax_id_of = |n: &roxmltree::Node| -> Option<String> {
        n.children()
            .find(|c| c.is_element() && c.tag_name().name() == "axId")
            .and_then(|c| attr(&c, "val"))
            .and_then(|value| parse_chart_axis_id_value(&value))
    };
    // The category axis is normally `<c:catAx>` or, for a date/time-series X
    // axis, `<c:dateAx>` (§21.2.2.39) — same child grammar, so every cat-axis
    // read below treats them identically. A SCATTER / BUBBLE chart has NO catAx:
    // it declares two `<c:valAx>` and the *horizontal* one (`axPos` b/t) plays
    // the category-axis role, while the *vertical* one (`axPos` l/r) is the
    // value axis. Detect this and route the horizontal valAx into `cat_ax` so
    // its tick-label / line / format / crossing properties land in the cat-axis
    // fields, exactly as Excel presents them.
    let real_cat_ax_nodes: Vec<_> = root
        .descendants()
        .filter(|n| n.is_element() && matches!(n.tag_name().name(), "catAx" | "dateAx"))
        .take(MAX_CHART_PLOT_AXES + 1)
        .collect();
    if real_cat_ax_nodes.len() + val_ax_nodes.len() > MAX_CHART_PLOT_AXES {
        return None;
    }
    let real_cat_ax = real_cat_ax_nodes
        .iter()
        .find(|axis| !matches!(ax_pos(axis).as_deref(), Some("t") | Some("r")))
        .or_else(|| real_cat_ax_nodes.first())
        .copied();
    let is_scatter_axes = real_cat_ax.is_none() && val_ax_nodes.len() >= 2;
    let scatter_x_val_ax = if is_scatter_axes {
        val_ax_nodes
            .iter()
            .find(|n| matches!(ax_pos(n).as_deref(), Some("b") | Some("t")))
            .copied()
    } else {
        None
    };
    let cat_ax = real_cat_ax.or(scatter_x_val_ax);

    // Primary value axis. Normally the first value axis that isn't on the right.
    // For scatter it's the VERTICAL (l/r) axis — never the horizontal one, which
    // is the category axis above. Secondary (combo charts) = a right-edge valAx.
    let val_ax = if is_scatter_axes {
        val_ax_nodes
            .iter()
            .find(|n| matches!(ax_pos(n).as_deref(), Some("l") | Some("r")))
            .or_else(|| val_ax_nodes.first())
            .copied()
    } else {
        val_ax_nodes
            .iter()
            .find(|n| ax_pos(n).as_deref() != Some("r"))
            .or_else(|| val_ax_nodes.first())
            .copied()
    };
    // A scatter/bubble group overlaid on a bar/line/area chart references two
    // additional numeric axes. The group's first `axId` is its horizontal X
    // axis and the second is its vertical Y axis (CT_ScatterChart sequence).
    // Resolve by ID instead of by `axPos`: both the primary bar value axis and
    // the scatter X axis commonly sit at `b`, so position alone is ambiguous.
    let scatter_axis_groups: Vec<Vec<String>> = root
        .descendants()
        .filter(|node| {
            node.is_element() && matches!(node.tag_name().name(), "scatterChart" | "bubbleChart")
        })
        .map(|group| {
            group
                .children()
                .filter(|node| node.is_element() && node.tag_name().name() == "axId")
                .filter_map(|node| attr(&node, "val"))
                .filter_map(|value| parse_chart_axis_id_value(&value))
                .take(MAX_CHART_GROUP_AXIS_IDS + 1)
                .collect()
        })
        .take(MAX_CHART_PLOT_GROUPS + 1)
        .collect();
    if scatter_axis_groups.len() > MAX_CHART_PLOT_GROUPS
        || scatter_axis_groups
            .iter()
            .any(|ids| ids.len() > MAX_CHART_GROUP_AXIS_IDS)
    {
        return None;
    }
    let combo_scatter_axis_ids = if !is_scatter_axes {
        find_chart("scatterChart")
            .or_else(|| find_chart("bubbleChart"))
            .map(|group| {
                group
                    .children()
                    .filter(|node| node.is_element() && node.tag_name().name() == "axId")
                    .filter_map(|node| attr(&node, "val"))
                    .filter_map(|value| parse_chart_axis_id_value(&value))
                    .collect::<Vec<_>>()
            })
            .unwrap_or_default()
    } else {
        Vec::new()
    };
    let axis_by_id = |id: Option<&String>| {
        id.and_then(|wanted| {
            val_ax_nodes
                .iter()
                .find(|node| ax_id_of(node).as_ref() == Some(wanted))
                .copied()
        })
    };
    let secondary_real_cat_ax = real_cat_ax_nodes
        .iter()
        .find(|axis| Some(**axis) != real_cat_ax)
        .copied();
    // A pure scatter chart can contain more than one CT_ScatterChart group,
    // each paired with its own X/Y value axes. The second group's horizontal
    // axis is the top secondary X axis and its vertical axis is the right
    // secondary Y axis. Resolving these by group-local axId is essential:
    // position alone cannot distinguish two bottom/top or left/right numeric
    // pairs, and dropping the pair also loses the series-to-axis binding.
    let pure_scatter_secondary_ids = is_scatter_axes
        .then(|| scatter_axis_groups.get(1))
        .flatten();
    let secondary_cat_ax = axis_by_id(
        pure_scatter_secondary_ids
            .and_then(|ids| ids.first())
            .or_else(|| combo_scatter_axis_ids.first()),
    )
    .or(secondary_real_cat_ax);
    let secondary_val_ax = axis_by_id(
        pure_scatter_secondary_ids
            .and_then(|ids| ids.get(1))
            .or_else(|| combo_scatter_axis_ids.get(1)),
    )
    .or_else(|| {
        if !is_scatter_axes && val_ax_nodes.len() >= 2 {
            val_ax_nodes
                .iter()
                .find(|n| ax_pos(n).as_deref() == Some("r"))
                .copied()
        } else {
            None
        }
    });
    let secondary_ax_id = secondary_val_ax.as_ref().and_then(ax_id_of);
    let (val_min, val_max) = val_ax.map(extract_axis_min_max).unwrap_or((None, None));
    let val_axis_hidden = val_ax.map(axis_is_deleted).unwrap_or(false);
    let cat_axis_hidden = cat_ax.map(axis_is_deleted).unwrap_or(false);

    // Series
    let plot_area = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "plotArea")?;

    let plot_area_fill_style = extract_direct_shape_fill_with_images(
        child(plot_area, "spPr"),
        color_resolver,
        image_resolver,
        ChartImageSource::Chart,
    );
    let plot_area_bg = if plot_area_fill_style.paint_authored == Some(true) {
        plot_area_fill_style.color.clone()
    } else {
        color_resolver.default_plot_area_bg()
    };
    let plot_area_fill_automatic = (plot_area_fill_style.paint_authored != Some(true)
        && plot_area_bg.is_some())
    .then_some(true);
    let plot_area_line_style = extract_direct_shape_line(plot_area, color_resolver);
    if child(plot_area, "dTable").is_some_and(|table| !chart_data_table_paint_within_limit(table)) {
        return None;
    }
    let data_table = extract_chart_data_table(plot_area, color_resolver);

    let classic_group_kind = |name: &str| -> Option<&'static str> {
        match name {
            "areaChart" => Some("area"),
            "area3DChart" => Some("area3D"),
            "lineChart" => Some("line"),
            "line3DChart" => Some("line3D"),
            "stockChart" => Some("stock"),
            "radarChart" => Some("radar"),
            "scatterChart" => Some("scatter"),
            "pieChart" => Some("pie"),
            "pie3DChart" => Some("pie3D"),
            "doughnutChart" => Some("doughnut"),
            "barChart" => Some("bar"),
            "bar3DChart" => Some("bar3D"),
            "ofPieChart" => Some("ofPie"),
            "surfaceChart" => Some("surface"),
            "surface3DChart" => Some("surface3D"),
            "bubbleChart" => Some("bubble"),
            _ => None,
        }
    };
    // CT_PlotArea's chart-group choice is ordered and unbounded. Preserve
    // every recognized direct child, including an empty group, before any
    // compatibility projection chooses the legacy top-level `chartType`.
    let chart_group_nodes: Vec<_> = plot_area
        .children()
        .filter(|node| node.is_element() && classic_group_kind(node.tag_name().name()).is_some())
        .take(MAX_CHART_PLOT_GROUPS + 1)
        .collect();
    if chart_group_nodes.len() > MAX_CHART_PLOT_GROUPS {
        return None;
    }
    let ser_nodes: Vec<_> = chart_group_nodes
        .iter()
        .flat_map(|group| {
            group
                .children()
                .filter(|node| node.is_element() && node.tag_name().name() == "ser")
        })
        .take(MAX_CHART_PLOT_SERIES + 1)
        .collect();
    if ser_nodes.len() > MAX_CHART_PLOT_SERIES {
        return None;
    }
    let plot_axis_nodes = plot_area
        .children()
        .filter(|node| {
            node.is_element()
                && matches!(
                    node.tag_name().name(),
                    "catAx" | "dateAx" | "valAx" | "serAx"
                )
        })
        .take(MAX_CHART_PLOT_AXES + 1)
        .collect::<Vec<_>>();
    if plot_axis_nodes.len() > MAX_CHART_PLOT_AXES {
        return None;
    }
    let known_axis_ids: BTreeMap<String, String> = plot_axis_nodes
        .iter()
        .copied()
        .filter_map(|axis| {
            child(axis, "axId")
                .and_then(|node| attr(&node, "val"))
                .and_then(|id| parse_chart_axis_id_value(&id))
                .map(|id| (id, axis.tag_name().name().to_string()))
        })
        .collect();
    if chart_group_nodes.iter().any(|group| {
        group
            .children()
            .filter(|node| node.is_element() && node.tag_name().name() == "axId")
            .take(MAX_CHART_GROUP_AXIS_IDS + 1)
            .count()
            > MAX_CHART_GROUP_AXIS_IDS
    }) {
        return None;
    }
    // Axis position is not a normative primary/secondary tie-breaker. If two
    // axes of the same schema kind occupy the same side, preserve their IDs
    // but leave group ownership unresolved until an Office-observed rule is
    // available. Assigning them by XML/group encounter order silently routes
    // otherwise valid data through an arbitrary scale.
    let mut axes_by_kind_and_position: BTreeMap<(String, Option<String>), Vec<String>> =
        BTreeMap::new();
    for axis in &plot_axis_nodes {
        let Some(id) = ax_id_of(axis) else {
            continue;
        };
        axes_by_kind_and_position
            .entry((axis.tag_name().name().to_string(), ax_pos(axis)))
            .or_default()
            .push(id);
    }
    let mut ambiguous_axis_ids: BTreeSet<String> = axes_by_kind_and_position
        .into_values()
        .filter(|ids| ids.len() > 1)
        .flatten()
        .collect();
    let mut axis_id_counts = BTreeMap::<String, usize>::new();
    for axis in &plot_axis_nodes {
        if let Some(id) = ax_id_of(axis) {
            *axis_id_counts.entry(id).or_default() += 1;
        }
    }
    ambiguous_axis_ids.extend(
        axis_id_counts
            .into_iter()
            .filter_map(|(id, count)| (count > 1).then_some(id)),
    );
    // Seed group ownership from the axes already selected for the public
    // primary/secondary slots. This keeps an explicitly right-positioned
    // value axis secondary even when no group references the left axis (for
    // example a stock group bound only to the right axis). Same-kind axes that
    // share a position were marked ambiguous above and never reach the claim
    // helper; no source-order tie-breaker is invented for them.
    let mut primary_category_axis_id = cat_ax.as_ref().and_then(ax_id_of);
    let mut secondary_category_axis_id = secondary_cat_ax
        .as_ref()
        .and_then(ax_id_of)
        .filter(|id| Some(id) != primary_category_axis_id.as_ref());
    let mut primary_value_axis_id = val_ax.as_ref().and_then(ax_id_of);
    let mut secondary_value_axis_id = secondary_val_ax
        .as_ref()
        .and_then(ax_id_of)
        .filter(|id| Some(id) != primary_value_axis_id.as_ref());
    let mut primary_series_axis_id = None;
    let mut secondary_series_axis_id = None;
    let mut next_series_start = 0usize;
    let plot_groups = chart_group_nodes
        .iter()
        .map(|group| {
            let kind = classic_group_kind(group.tag_name().name()).expect("filtered group");
            let series_count = group
                .children()
                .filter(|node| node.is_element() && node.tag_name().name() == "ser")
                .count();
            let series_start = next_series_start;
            next_series_start = next_series_start.saturating_add(series_count);
            let axis_ids: Vec<String> = group
                .children()
                .filter(|node| node.is_element() && node.tag_name().name() == "axId")
                .map(|node| attr(&node, "val").and_then(|value| parse_chart_axis_id_value(&value)))
                .collect::<Option<Vec<_>>>()?;
            if axis_ids
                .iter()
                .enumerate()
                .any(|(index, id)| axis_ids[..index].contains(id))
            {
                return None;
            }
            let axis_id_with_kind = |axis_kind: &str| {
                axis_ids.iter().find(|id| {
                    known_axis_ids
                        .get(*id)
                        .is_some_and(|kind| kind == axis_kind)
                })
            };
            let no_axes = matches!(kind, "pie" | "pie3D" | "doughnut" | "ofPie");
            let numeric_axes = matches!(kind, "scatter" | "bubble");
            let category_axis_id = if numeric_axes {
                axis_ids.first()
            } else {
                axis_id_with_kind("catAx").or_else(|| axis_id_with_kind("dateAx"))
            };
            let value_axis_id = if numeric_axes {
                axis_ids.get(1)
            } else {
                axis_id_with_kind("valAx")
            };
            let series_axis_id = axis_id_with_kind("serAx");
            let category_axis_id = category_axis_id.filter(|id| !ambiguous_axis_ids.contains(*id));
            let value_axis_id = value_axis_id.filter(|id| !ambiguous_axis_ids.contains(*id));
            let series_axis_id = series_axis_id.filter(|id| !ambiguous_axis_ids.contains(*id));
            let (category_axis, value_axis, series_axis) = if no_axes {
                ("none".to_string(), "none".to_string(), "none".to_string())
            } else {
                (
                    claim_plot_group_axis_slot(
                        category_axis_id,
                        &known_axis_ids,
                        &mut primary_category_axis_id,
                        &mut secondary_category_axis_id,
                    ),
                    claim_plot_group_axis_slot(
                        value_axis_id,
                        &known_axis_ids,
                        &mut primary_value_axis_id,
                        &mut secondary_value_axis_id,
                    ),
                    if series_axis_id.is_some() {
                        claim_plot_group_axis_slot(
                            series_axis_id,
                            &known_axis_ids,
                            &mut primary_series_axis_id,
                            &mut secondary_series_axis_id,
                        )
                    } else {
                        "none".to_string()
                    },
                )
            };
            let (gap_width, overlap) = if matches!(kind, "bar" | "bar3D") {
                extract_bar_gap_overlap(*group)
            } else {
                (None, None)
            };
            Some(ChartPlotGroup {
                kind: kind.to_string(),
                series_start,
                series_count,
                category_axis,
                value_axis,
                series_axis,
                axis_ids: (!axis_ids.is_empty()).then_some(axis_ids),
                grouping: child(*group, "grouping").and_then(|node| attr(&node, "val")),
                bar_direction: child(*group, "barDir").and_then(|node| attr(&node, "val")),
                scatter_style: child(*group, "scatterStyle").and_then(|node| attr(&node, "val")),
                radar_style: child(*group, "radarStyle").and_then(|node| attr(&node, "val")),
                vary_colors: child(*group, "varyColors")
                    .map(|_| bool_child(*group, "varyColors").unwrap_or(true)),
                gap_width,
                overlap,
                bubble_scale: child(*group, "bubbleScale")
                    .and_then(|node| attr(&node, "val"))
                    .as_deref()
                    .and_then(parse_unsigned_percent)
                    .filter(|value| *value <= 300)
                    .map(f64::from),
                bubble_size_represents: child(*group, "sizeRepresents")
                    .and_then(|node| attr(&node, "val"))
                    .filter(|value| value == "area" || value == "w"),
                show_negative_bubbles: bool_child(*group, "showNegBubbles"),
            })
        })
        .collect::<Option<Vec<_>>>()?;
    let plot_group_uses_secondary_axis = chart_group_nodes
        .iter()
        .zip(plot_groups.iter())
        .map(|(node, group)| (node.id(), group.value_axis == "secondary"))
        .collect::<std::collections::HashMap<_, _>>();
    let mut direct_label_shapes = Vec::new();
    for series in &ser_nodes {
        if let Some(labels) = child(*series, "dLbls") {
            direct_label_shapes.extend(label_shape_nodes(labels));
        }
        direct_label_shapes.extend(
            series
                .children()
                .filter(|node| node.is_element() && node.tag_name().name() == "trendline")
                .filter_map(|trendline| child(trendline, "trendlineLbl"))
                .filter_map(|label| child(label, "spPr")),
        );
    }
    let allow_direct_label_paints =
        label_paint_recipes_within_budget(direct_label_shapes.iter().copied());
    let mut parsed_group_data_labels = chart_group_nodes
        .iter()
        .filter_map(|group| {
            let labels = child(*group, "dLbls")?;
            let defaults = parse_chart_group_data_labels(labels);
            let owned_series = group
                .children()
                .filter(|node| node.is_element() && node.tag_name().name() == "ser")
                .count();
            Some((group.id(), defaults, owned_series))
        })
        .collect::<Vec<_>>();
    // Projection into series is intentionally bounded chart-wide. A separator
    // is arbitrary text, so multiplying it across many groups/series must not
    // amplify a small XML source into unbounded wire/JSON memory. Refuse every
    // group separator atomically when their combined projection exceeds the
    // shared chart-point resource ceiling; semantic show flags remain intact.
    let separators_fit =
        chart_group_separator_projection_within_budget(parsed_group_data_labels.iter().map(
            |(_, defaults, owned_series)| {
                (
                    defaults
                        .as_ref()
                        .and_then(|item| item.separator.as_ref())
                        .map(|separator| separator.chars().count())
                        .unwrap_or(0),
                    *owned_series,
                )
            },
        ));
    if !separators_fit {
        for (_, defaults, _) in &mut parsed_group_data_labels {
            if let Some(defaults) = defaults.as_mut() {
                defaults.separator = None;
            }
        }
    }
    let group_data_labels = parsed_group_data_labels
        .into_iter()
        .map(|(id, defaults, _)| (id, defaults))
        .collect::<std::collections::HashMap<_, _>>();
    let has_group_data_labels = !group_data_labels.is_empty();

    let bar_group_nodes: Vec<_> = plot_area
        .children()
        .filter(|node| {
            node.is_element() && matches!(node.tag_name().name(), "barChart" | "bar3DChart")
        })
        .collect();
    let line_group_nodes: Vec<_> = plot_area
        .children()
        .filter(|node| {
            node.is_element() && matches!(node.tag_name().name(), "lineChart" | "line3DChart")
        })
        .collect();
    let area_group_nodes: Vec<_> = plot_area
        .children()
        .filter(|node| {
            node.is_element() && matches!(node.tag_name().name(), "areaChart" | "area3DChart")
        })
        .collect();
    let line_group_indices: std::collections::HashMap<_, _> = line_group_nodes
        .iter()
        .enumerate()
        .map(|(index, node)| (node.id(), index as u32))
        .collect();
    let area_group_indices: std::collections::HashMap<_, _> = area_group_nodes
        .iter()
        .enumerate()
        .map(|(index, node)| (node.id(), index as u32))
        .collect();
    let bar_group_indices: std::collections::HashMap<_, _> = bar_group_nodes
        .iter()
        .enumerate()
        .map(|(index, node)| (node.id(), index as u32))
        .collect();
    let parsed_line_group_decorations: Vec<_> = line_group_nodes
        .iter()
        .enumerate()
        .filter_map(|(group_index, group)| {
            let drop_lines = child(*group, "dropLines")
                .map(|node| parse_chart_decoration_line_style(node, color_resolver));
            let hi_low_lines = child(*group, "hiLowLines")
                .map(|node| parse_chart_decoration_line_style(node, color_resolver));
            let up_down_bars = child(*group, "upDownBars")
                .map(|node| parse_chart_up_down_bar_style(node, color_resolver));
            (drop_lines.is_some() || hi_low_lines.is_some() || up_down_bars.is_some()).then_some(
                ChartLineGroupDecorations {
                    group_index: group_index as u32,
                    drop_lines,
                    hi_low_lines,
                    up_down_bars,
                },
            )
        })
        .collect();
    let line_group_decorations =
        (!parsed_line_group_decorations.is_empty()).then_some(parsed_line_group_decorations);
    let parsed_area_group_decorations: Vec<_> = area_group_nodes
        .iter()
        .enumerate()
        .filter_map(|(group_index, group)| {
            child(*group, "dropLines").map(|node| ChartAreaGroupDecorations {
                group_index: group_index as u32,
                drop_lines: Some(parse_chart_decoration_line_style(node, color_resolver)),
            })
        })
        .collect();
    let area_group_decorations =
        (!parsed_area_group_decorations.is_empty()).then_some(parsed_area_group_decorations);
    let parsed_bar_group_decorations: Vec<_> = bar_group_nodes
        .iter()
        .enumerate()
        .filter_map(|(group_index, group)| {
            let series_lines: Vec<_> = group
                .children()
                .filter(|node| node.is_element() && node.tag_name().name() == "serLines")
                .map(|node| parse_chart_decoration_line_style(node, color_resolver))
                .collect();
            (!series_lines.is_empty()).then_some(ChartBarGroupDecorations {
                group_index: group_index as u32,
                series_lines: Some(series_lines),
            })
        })
        .collect();
    let bar_group_decorations =
        (!parsed_bar_group_decorations.is_empty()).then_some(parsed_bar_group_decorations);

    // Chart-level category labels from the first series, using the POSITIONAL
    // collector so a sparse cache (labels that start at `idx=1`, or a hole in
    // the middle) keeps its true length and per-index alignment. The old
    // document-order collector collapsed such caches, truncating every series
    // and mis-registering data (issue: cat-less line 11→1, idx=1 radar 11→10).
    // Scatter/bubble carry their X labels in `<c:xVal>` (there is no `<c:cat>`),
    // so read that instead — the shared category list mirrors the first series'
    // X data, matching how Excel drives the horizontal-axis labels.
    let chart_uses_xval = chart_type == "scatter" || chart_type == "bubble";
    let category_tag = if chart_uses_xval { "xVal" } else { "cat" };
    let first_series = ser_nodes.first().copied();
    let shared_category_formula =
        first_series.and_then(|series| external_reference_formula(series, category_tag));
    let resolved_category_levels = first_series
        .and_then(|series| collect_multi_level_str_cache(series, category_tag))
        .or_else(|| {
            shared_category_formula
                .as_deref()
                .and_then(|formula| references.resolve_string_levels(formula))
        });
    let categories: Vec<String> = resolved_category_levels
        .as_ref()
        .and_then(|levels| levels.first().cloned())
        .or_else(|| {
            first_series.and_then(|series| collect_string_source(series, category_tag, references))
        })
        .unwrap_or_default();
    let category_levels = resolved_category_levels.filter(|levels| levels.len() > 1);
    let category_source_hidden =
        first_series.and_then(|series| collect_source_hidden(series, category_tag, references));

    // Map a chart-group element name to the per-series `seriesType` string the
    // renderer dispatches on (mixed bar+line charts key line vs. non-line off
    // this field). Mirrors the xlsx `type_map`; `bubbleChart` folds to
    // `scatter` like everything else.
    // 3D groups fold to the same series type as their 2D equivalent (they are
    // flattened above); `stockChart`/`ofPieChart` have no combo-mixing role so
    // map to a plain type too.
    let group_series_type = |group_name: &str| -> Option<String> {
        match group_name {
            "barChart" | "bar3DChart" => Some("bar"),
            "lineChart" | "line3DChart" => Some("line"),
            "areaChart" | "area3DChart" => Some("area"),
            "pieChart" | "pie3DChart" | "ofPieChart" => Some("pie"),
            "doughnutChart" => Some("doughnut"),
            "radarChart" => Some("radar"),
            "scatterChart" | "bubbleChart" => Some("scatter"),
            "stockChart" => Some("stock"),
            "surfaceChart" => Some("surface"),
            "surface3DChart" => Some("surface3D"),
            _ => None,
        }
        .map(|s| s.to_string())
    };

    // Total series in the plot. §21.2.2.227 varyColors on a NON-pie chart
    // varies each data point by color only for a single-series plot (Office
    // keeps per-series colors when several series share the axes); captured
    // here so the per-series closure can gate the accent fill on it.
    let series_count = ser_nodes.len();
    // Numeric classic styles and linked Chart Styles both address series by
    // the source formatting index. Preserve sparse `<c:ser><c:idx>` values on
    // the shared series carrier before any host filters or renderer grouping
    // can compact them to a visible/document-order ordinal.
    let source_series_formatting_indices = ser_nodes
        .iter()
        .enumerate()
        .map(|(position, series)| {
            child(*series, "idx")
                .and_then(|node| node.attribute("val"))
                .and_then(|value| value.parse::<u32>().ok())
                .map(|value| value as usize)
                .unwrap_or(position)
        })
        .collect::<Vec<_>>();
    // Direct marker gradients are replayed for data points by the Canvas
    // renderer. Share one component budget across every series and point in
    // this chart so many individually-valid recipes cannot amplify the wire
    // model without bound.
    let mut marker_paint_budget = MAX_CHART_MARKER_PAINT_COMPONENTS;
    let mut marker_paint_budget_exceeded = false;

    let series: Vec<ChartSeries> = ser_nodes
        .iter()
        .enumerate()
        .map(|(series_position, ser)| {
            // Each `<c:ser>` is a direct child of its chart-group element
            // (`<c:barChart>`/`<c:lineChart>`/…). `series_type` carries that
            // group's type so the renderer can draw line-group series as a line
            // over the columns in a combo chart (ECMA-376 §21.2.2.97); we also
            // flag series whose group references the secondary value axis so they
            // plot against the right-hand scale.
            let group = ser.parent();
            let is_classic_three_d_series = group.is_some_and(|owner| {
                matches!(
                    owner.tag_name().name(),
                    "bar3DChart" | "line3DChart" | "area3DChart" | "pie3DChart"
                )
            });
            let series_type = group
                .map(|p| p.tag_name().name())
                .and_then(group_series_type);
            let bubble_group = group.filter(|node| node.tag_name().name() == "bubbleChart");
            let line_group_index = group
                .and_then(|owner| line_group_indices.get(&owner.id()))
                .copied();
            let area_group_index = group
                .and_then(|owner| area_group_indices.get(&owner.id()))
                .copied();
            let bar_group_index = group
                .and_then(|owner| bar_group_indices.get(&owner.id()))
                .copied();
            let (bar_group_gap_width, bar_group_overlap) = group
                .filter(|owner| matches!(owner.tag_name().name(), "barChart" | "bar3DChart"))
                .map(extract_bar_gap_overlap)
                .unwrap_or((None, None));
            let bar_group_direction = group
                .filter(|owner| matches!(owner.tag_name().name(), "barChart" | "bar3DChart"))
                .and_then(|owner| child(owner, "barDir"))
                .and_then(|node| node.attribute("val"))
                .map(str::to_string);
            let bar_group_grouping = group
                .filter(|owner| matches!(owner.tag_name().name(), "barChart" | "bar3DChart"))
                .and_then(|owner| child(owner, "grouping"))
                .and_then(|node| node.attribute("val"))
                .map(str::to_string);
            let series_is_scatter_like = matches!(series_type.as_deref(), Some("scatter"));
            let own_category_tag = if series_is_scatter_like {
                "xVal"
            } else {
                "cat"
            };
            let use_secondary_axis = group
                .and_then(|owner| plot_group_uses_secondary_axis.get(&owner.id()))
                .copied()
                .unwrap_or_else(|| match (group, secondary_ax_id.as_deref()) {
                    (Some(g), Some(sec)) => g
                        .children()
                        .filter(|c| c.is_element() && c.tag_name().name() == "axId")
                        .any(|c| attr(&c, "val").as_deref() == Some(sec)),
                    _ => false,
                });

            // Series name from <c:tx>  (can be strRef/strCache, strLit, or a bare <c:v>)
            let name = collect_string_source(*ser, "tx", references)
                .and_then(|values| values.into_iter().next())
                .or_else(|| {
                    ser.children()
                        .find(|n| n.is_element() && n.tag_name().name() == "tx")
                        .and_then(|tx| {
                            tx.children()
                                .find(|n| n.is_element() && n.tag_name().name() == "v")
                                .and_then(|v| v.text().map(|t| t.to_string()))
                        })
                })
                .unwrap_or_default();

            // `<c:idx val>` (ECMA-376 §21.2.2.84) — the canonical series index
            // Office uses for default-palette color selection. `<c:order>` is a
            // separate display-order field and must NOT drive coloring.
            let series_idx: usize = child(*ser, "idx")
                .and_then(|n| n.attribute("val"))
                .and_then(|v| v.parse::<usize>().ok())
                .unwrap_or(0);

            // Per-series category labels. Scatter/bubble put numeric X data in
            // `<c:xVal>` (ECMA-376 §21.2.2.43); every other type reads the
            // series' own `<c:cat>`. The first series is already represented by
            // chart-level `categories`; repeated live formulas also use that
            // canonical vector instead of resolving and retaining duplicate
            // copies. `ChartSeries.categories = None` explicitly means to fall
            // back to `ChartModel.categories` (the shared TS contract). Authored
            // caches/literals and genuinely distinct formulas remain per-series.
            let series_categories: Option<Vec<String>> = {
                let own_formula = external_reference_formula(*ser, own_category_tag);
                let shares_chart_categories = own_category_tag == category_tag
                    && (series_position == 0
                        || (own_formula.is_some()
                            && own_formula.as_deref() == shared_category_formula.as_deref()));
                if shares_chart_categories {
                    None
                } else {
                    let has_own_source = ser.children().any(|node| {
                        node.is_element() && node.tag_name().name() == own_category_tag
                    });
                    match collect_string_source(*ser, own_category_tag, references) {
                        Some(own) if series_is_scatter_like || !own.is_empty() => Some(own),
                        // A distinct scatter/bubble X source that cannot be
                        // resolved must not inherit the first series' X values.
                        // `Some([])` explicitly suppresses that fallback.
                        None if series_is_scatter_like && has_own_source => Some(Vec::new()),
                        _ => None,
                    }
                }
            };
            let bubble_x_source_is_string = bubble_group
                .and_then(|_| child(*ser, "xVal"))
                .and_then(|source| {
                    if child(source, "strRef").is_some()
                        || child(source, "strLit").is_some()
                        || child(source, "multiLvlStrRef").is_some()
                    {
                        Some(true)
                    } else if child(source, "numRef").is_some() || child(source, "numLit").is_some()
                    {
                        Some(false)
                    } else {
                        None
                    }
                });

            // Y values (scatter/bubble → `<c:yVal>`, else `<c:val>`), collected
            // POSITIONALLY. The series' own cache `<c:ptCount>` sizes the vector
            // and each `<c:pt idx>` lands at its index — the length no longer
            // rides on the category count, so a value series with more points
            // than there are cat labels (cat-less line, sparse radar) keeps all
            // of its data.
            let val_tag = if series_is_scatter_like {
                "yVal"
            } else {
                "val"
            };
            let values: Vec<Option<f64>> =
                collect_number_source(*ser, val_tag, references).unwrap_or_default();
            let series_pt_count = values.len().max(1);
            // Value-cache node for the series-value number format (`<c:formatCode>`).
            let val_cache = ser
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == val_tag)
                .and_then(|v| {
                    v.descendants().find(|n| {
                        n.is_element()
                            && (n.tag_name().name() == "numCache"
                                || n.tag_name().name() == "numLit")
                    })
                });

            // Bubble per-point sizes (ECMA-376 §21.2.2.4 `<c:bubbleSize>`).
            // Only meaningful for bubble charts; scatter / others ignore.
            let bubble_sizes: Option<Vec<Option<f64>>> = if group
                .map(|node| node.tag_name().name() == "bubbleChart")
                .unwrap_or(false)
            {
                collect_number_source(*ser, "bubbleSize", references).map(|mut sizes| {
                    sizes.resize(sizes.len().max(series_pt_count), None);
                    sizes
                })
            } else {
                None
            };
            let bubble_3d_group_default =
                bubble_group.and_then(|owner| strict_boolean_child(owner, "bubble3D"));
            let bubble_3d = bubble_group.and_then(|_| strict_boolean_child(*ser, "bubble3D"));
            let mut source_hidden = collect_source_hidden(*ser, val_tag, references);
            if series_is_scatter_like {
                merge_source_hidden(
                    &mut source_hidden,
                    collect_source_hidden(*ser, own_category_tag, references),
                );
            }
            if group
                .map(|node| node.tag_name().name() == "bubbleChart")
                .unwrap_or(false)
            {
                merge_source_hidden(
                    &mut source_hidden,
                    collect_source_hidden(*ser, "bubbleSize", references),
                );
            }

            // Series color from spPr > solidFill (bar/area/pie) or spPr > ln >
            // solidFill (line/scatter/radar carry their color on the stroke).
            // When neither is present, fall back to the theme accent for this
            // series index (`theme.accent[(idx % 6) + 1]`) via the resolver, so
            // the default Office palette renders without theme access. Resolvers
            // whose renderer owns its own palette (pptx) return `None` here and
            // keep `color` unset.
            let color = ser
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "spPr")
                .and_then(|sp| {
                    if sp
                        .children()
                        .any(|n| n.is_element() && n.tag_name().name() == "noFill")
                    {
                        // An explicit shape-level `<a:noFill/>` suppresses the
                        // series fill but does not remove the series from the
                        // data model (notably, an invisible first series can be
                        // the baseline of a stacked area chart). CSS accepts
                        // 8-digit hex, so preserve that authored transparency
                        // without conflating it with `None` (theme fallback).
                        Some("00000000".to_string())
                    } else if sp
                        .children()
                        .any(|n| n.is_element() && n.tag_name().name() == "solidFill")
                    {
                        color_resolver.resolve_shape_fill(sp)
                    } else {
                        sp.children()
                            .find(|n| n.is_element() && n.tag_name().name() == "ln")
                            .and_then(|ln| color_resolver.resolve_shape_fill(ln))
                    }
                })
                .or_else(|| color_resolver.resolve_series_accent(series_idx));

            // §21.2.2.198 series-level `<c:spPr><a:ln>`: an explicit `<a:noFill/>`
            // turns the connecting line OFF, overriding the chart-group
            // `<c:scatterStyle>` (§21.2.2.42) / line-group default. Excel draws a
            // markers-only scatter when the series line is `<a:noFill/>` even
            // though `<c:scatterStyle val="lineMarker">` sets the group default to
            // connect points. Only the noFill flag matters here (color/width ride
            // on `color` above), so discard the other two tuple fields.
            let (line_color, line_width_emu, line_no_fill) =
                extract_sp_pr_ln_style(*ser, color_resolver);

            // Per-data-point colors from <c:dPt> (§21.2.2.52; important for
            // pie charts). The point index is the CHILD element `<c:idx val>`
            // (ECMA-376 §21.2.2.84, CT_UnsignedInt), not an attribute on
            // `<c:dPt>` — the old `attr(dpt, "idx")` always returned None, so
            // every slice fell
            // back to the series colour. The fill is `<c:spPr><a:solidFill>`;
            // restrict to spPr's direct child so a border `<a:ln><a:solidFill>`
            // can't be mistaken for the slice fill.
            // Index `<c:dPt>` once. Re-scanning every dPt for every cache slot
            // is quadratic for a fully-authored series and can monopolize the
            // WASM thread even though the cache itself is bounded. Preserve the
            // document-first duplicate behavior of the former `.find()` path.
            let mut data_points_by_index: std::collections::HashMap<usize, Node> =
                std::collections::HashMap::new();
            for dpt in ser
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "dPt")
            {
                let Some(index) = dpt
                    .children()
                    .find(|n| n.is_element() && n.tag_name().name() == "idx")
                    .and_then(|n| attr(&n, "val"))
                    .and_then(|v| v.parse::<usize>().ok())
                    .filter(|index| *index < series_pt_count)
                else {
                    continue;
                };
                data_points_by_index.entry(index).or_insert(dpt);
            }
            let data_point_colors: Vec<Option<String>> = (0..series_pt_count)
                .map(|i| {
                    data_points_by_index
                        .get(&i)
                        .and_then(|dpt| {
                            dpt.children()
                                .find(|n| n.is_element() && n.tag_name().name() == "spPr")
                        })
                        .and_then(|sp| {
                            if child(sp, "noFill").is_some() {
                                // A direct point noFill is more specific than
                                // both series formatting and varyColors.
                                Some("00000000".to_string())
                            } else {
                                sp.children()
                                    .find(|n| n.is_element() && n.tag_name().name() == "solidFill")
                                    .and_then(|fill| color_resolver.resolve_solid_fill(fill))
                            }
                        })
                })
                .collect();

            // `data_point_colors` is direct `<c:dPt>` formatting only. The
            // effective classic `dataPoint` Chart Style owns automatic
            // varyColors paint, keeping direct point formatting distinguishable
            // and authoritative in every renderer family.
            let has_dpt_colors = data_point_colors.iter().any(|c| c.is_some());

            // Per-point `<c:dPt>` overrides (§21.2.2.39): marker (symbol/size/
            // fill/line/line width) and `<c:explosion>` (pie/doughnut pull-out). Plain
            // Direct point fill is retained in both the compatibility color
            // vector and the structured override. Renderers use the override
            // for direct-over-style precedence; older consumers retain their
            // established indexed-color contract.
            let data_point_overrides: Vec<ChartDataPointOverride> =
                parse_data_point_overrides_with_budget(
                    *ser,
                    color_resolver,
                    image_resolver,
                    true,
                    &mut marker_paint_budget,
                    &mut marker_paint_budget_exceeded,
                )
                .into_iter()
                .filter(|o| {
                    o.color.is_some()
                        || o.fill_hidden.is_some()
                        || o.chartex_style.is_some()
                        || o.line_color.is_some()
                        || o.line_width_emu.is_some()
                        || o.line_dash.is_some()
                        || o.line_hidden.is_some()
                        || o.marker_symbol.is_some()
                        || o.marker_size.is_some()
                        || o.marker_fill.is_some()
                        || o.marker_fill_paint.is_some()
                        || o.marker_fill_paint_authored.is_some()
                        || o.marker_line.is_some()
                        || o.marker_line_width_emu.is_some()
                        || o.marker_line_paint_authored.is_some()
                        || o.bubble_3d.is_some()
                        || o.explosion.is_some()
                })
                .collect();

            // Series value number format from `<c:val>…<c:numCache><c:formatCode>`.
            // Used for data labels when `<c:dLbls>` carries no explicit `<c:numFmt>`
            // (ECMA-376 §21.2.2.121). "General" means "no format" → drop it so the
            // renderer's default integer/decimal formatter takes over.
            let val_format_code = val_cache
                .and_then(|cache| {
                    cache
                        .children()
                        .find(|n| n.is_element() && n.tag_name().name() == "formatCode")
                        .and_then(|fc| fc.text().map(|t| t.to_string()))
                })
                .filter(|s| !s.is_empty() && s != "General");
            let cat_format_code = ser
                .children()
                .find(|node| node.is_element() && node.tag_name().name() == own_category_tag)
                .and_then(|source| {
                    source.descendants().find(|node| {
                        node.is_element() && matches!(node.tag_name().name(), "numCache" | "numLit")
                    })
                })
                .and_then(|cache| child(cache, "formatCode"))
                .and_then(|format| format.text().map(str::to_string))
                .filter(|format| !format.is_empty() && format != "General");
            let cat_format_builtin_id = ser
                .children()
                .find(|node| node.is_element() && node.tag_name().name() == own_category_tag)
                .and_then(|source| child(source, "numRef").or_else(|| child(source, "strRef")))
                .and_then(|reference| child(reference, "f"))
                .and_then(|formula| formula.text())
                .and_then(|formula| references.resolve_number_format_id(formula));
            let cat_format_codes = ser
                .children()
                .find(|node| node.is_element() && node.tag_name().name() == own_category_tag)
                .and_then(|source| {
                    source.descendants().find(|node| {
                        node.is_element() && matches!(node.tag_name().name(), "numCache" | "numLit")
                    })
                })
                .and_then(|cache| {
                    let point_count = child(cache, "ptCount")
                        .and_then(|count| attr(&count, "val"))
                        .and_then(|count| count.parse::<usize>().ok())
                        .unwrap_or_else(|| {
                            cache
                                .children()
                                .filter(|node| node.is_element() && node.tag_name().name() == "pt")
                                .filter_map(|point| attr(&point, "idx"))
                                .filter_map(|idx| idx.parse::<usize>().ok())
                                .max()
                                .map(|idx| idx + 1)
                                .unwrap_or(0)
                        });
                    let mut formats = vec![None; point_count];
                    for point in cache
                        .children()
                        .filter(|node| node.is_element() && node.tag_name().name() == "pt")
                    {
                        let Some(idx) =
                            attr(&point, "idx").and_then(|idx| idx.parse::<usize>().ok())
                        else {
                            continue;
                        };
                        if idx >= formats.len() {
                            formats.resize(idx + 1, None);
                        }
                        formats[idx] = attr(&point, "formatCode")
                            .filter(|format| !format.is_empty() && format != "General");
                    }
                    formats.iter().any(Option::is_some).then_some(formats)
                });

            // Series-level data-label text colour from `<c:dLbls><c:txPr>…solidFill`.
            // Scoped to this `<c:ser>` (not chart-root) so stacked-bar segments keep
            // their independent label colours (white on dark fill, black on light).
            let label_color = ser
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "dLbls")
                .and_then(|dlbls| {
                    dlbls
                        .children()
                        .find(|n| n.is_element() && n.tag_name().name() == "txPr")
                })
                .and_then(|txpr| {
                    txpr.descendants()
                        .find(|n| n.is_element() && n.tag_name().name() == "solidFill")
                })
                .and_then(|fill| color_resolver.resolve_solid_fill(fill));

            // Marker styling (ECMA-376 §21.2.2.32/§21.2.2.34). A per-series
            // `<c:marker>` gives the symbol/size/fill/line; when the symbol is
            // absent the chart-type-level `<c:lineChart><c:marker val>` default
            // (§21.2.2.33) governs visibility. Scatter defaults to visible
            // markers even without an explicit flag.
            // §21.2.2.33 chart-type-level `<c:lineChart><c:marker>` — CT_Boolean,
            // so a bare `<c:marker/>` enables markers (val default true); absent
            // ⇒ false (line series draw no markers unless opted in).
            let chart_marker_default = group.and_then(|g| bool_child(g, "marker")).unwrap_or(false);
            let chart_smooth_default = group
                .filter(|owner| owner.tag_name().name() == "lineChart")
                .and_then(|owner| bool_child(owner, "smooth"));
            let marker_node = child(*ser, "marker");
            let (
                (
                    marker_symbol,
                    marker_size,
                    marker_fill,
                    marker_fill_paint,
                    marker_fill_paint_authored,
                    marker_line,
                    marker_line_width_emu,
                    marker_line_paint_authored,
                ),
                marker_style,
            ) = parse_marker_model_style_with_budget(
                marker_node,
                color_resolver,
                image_resolver,
                &mut marker_paint_budget,
                &mut marker_paint_budget_exceeded,
            );
            let show_marker = match (&marker_symbol, series_is_scatter_like) {
                (Some(sym), _) => sym != "none",
                (None, true) => true,
                _ => chart_marker_default,
            };
            let automatic_marker_symbol = (chart_type == "scatter" && marker_symbol.is_none())
                .then(|| {
                    const SYMBOLS: [&str; 6] =
                        ["diamond", "square", "triangle", "x", "star", "circle"];
                    SYMBOLS[series_idx % SYMBOLS.len()].to_string()
                });

            // Series-level `<c:dLbls>` defaults + per-idx custom labels, and
            // error bars (§21.2.2.20, resolved to absolute plus/minus arrays).
            let dlbl_range_cache = collect_dlbl_range_cache(*ser);
            let (direct_series_data_labels, direct_data_label_overrides) =
                parse_series_data_labels_with_paint_policy(
                    *ser,
                    color_resolver,
                    &dlbl_range_cache,
                    allow_direct_label_paints,
                );
            let direct_labels_node = child(*ser, "dLbls");
            let group_defaults = group
                .and_then(|owner| group_data_labels.get(&owner.id()))
                .cloned()
                // `show_data_labels` is a legacy chart-wide projection. Once
                // any group block is materialized per owned series, an absent
                // block on a sibling group means all flags are false; preserve
                // that explicit effective default so labels cannot leak across
                // combo-chart groups through the legacy fallback.
                .or_else(|| has_group_data_labels.then_some(Some(Default::default())))
                .flatten();
            let series_data_labels = merge_chart_series_data_labels(
                group_defaults,
                direct_series_data_labels,
                direct_labels_node,
            );
            let data_label_overrides = direct_data_label_overrides;
            let visible_error_bar_values = (plot_visible_only == Some(true)).then(|| {
                values
                    .iter()
                    .enumerate()
                    .map(|(index, value)| {
                        let hidden_category = !series_is_scatter_like
                            && category_source_hidden
                                .as_ref()
                                .is_some_and(|hidden| hidden.get(index) == Some(&true));
                        let hidden_value = source_hidden
                            .as_ref()
                            .is_some_and(|hidden| hidden.get(index) == Some(&true));
                        if hidden_category || hidden_value {
                            None
                        } else {
                            *value
                        }
                    })
                    .collect::<Vec<_>>()
            });
            let err_bars = parse_error_bars(
                *ser,
                visible_error_bar_values.as_deref().unwrap_or(&values),
                color_resolver,
            );
            let invert_if_negative = bool_child(*ser, "invertIfNegative");
            let mut automatic_negative_style = None;
            let inverted_format = ser
                .descendants()
                .find(|node| node.is_element() && node.tag_name().name() == "invertSolidFillFmt");
            let inverted_shape = inverted_format.and_then(|format| child(format, "spPr"));
            let inverted_paint = inverted_shape.and_then(|shape| {
                parse_chart_style_paint(
                    shape,
                    color_resolver,
                    None,
                    &EmptyChartImageResolver,
                    ChartImageSource::Chart,
                )
            });
            let inverted_fill_authored =
                matches!(inverted_paint.as_ref(), Some(ChartStylePaint::Fill(_)));
            let (inverted_fill, inverted_fill_hidden) = match inverted_paint {
                Some(ChartStylePaint::NoFill) => (None, Some(true)),
                Some(ChartStylePaint::Fill(fill)) => (Some(*fill), None),
                Some(ChartStylePaint::Unresolved) => (None, None),
                None => (None, None),
            };
            let (inverted_line_color, inverted_line_width_emu, inverted_line_no_fill) =
                inverted_format
                    .map(|format| extract_sp_pr_ln_style(format, color_resolver))
                    .unwrap_or((None, None, false));
            let inverted_line_authored = inverted_format.map(|_| {
                inverted_shape
                    .and_then(|shape| child(shape, "ln"))
                    .is_some()
            });

            // Excel's implicit classic-chart style gives an entirely-negative,
            // otherwise unformatted bar/column series an outline-only marker.
            // A positive/all-negative mirror pair isolates this application
            // default: the positive chart takes accent1, while the negative chart has
            // no fill and a 0.75pt black outline.  OOXML does not encode that
            // application-generated paint, so keep the compatibility rule at
            // the narrow observed boundary.  Any authored chart style, series
            // formatting, point formatting, or invertIfNegative value wins.
            let observed_single_clustered_column = group.is_some_and(|owner| {
                owner.tag_name().name() == "barChart"
                    && child(owner, "barDir").and_then(|node| node.attribute("val")) == Some("col")
                    && child(owner, "grouping")
                        .and_then(|node| node.attribute("val"))
                        .is_none_or(|grouping| grouping == "clustered")
                    && bool_child(owner, "varyColors") == Some(false)
            }) && series_count == 1;
            let implicit_outline_only_negative_series = color_resolver
                .implicit_outline_only_negative_column_style()
                && observed_single_clustered_column
                && legacy_chart_style.is_none()
                && invert_if_negative.is_none()
                && inverted_format.is_none()
                && child(*ser, "spPr").is_none()
                && !ser.children().any(|node| {
                    node.is_element()
                        && node.tag_name().name() == "dPt"
                        && child(node, "spPr").is_some()
                })
                && values.iter().flatten().any(|value| *value < 0.0)
                && values.iter().flatten().all(|value| *value < 0.0);
            if implicit_outline_only_negative_series {
                automatic_negative_style = Some(true);
            }

            let chartex_style = if is_classic_three_d_series {
                parse_three_d_series_style_with_budget(
                    *ser,
                    color_resolver,
                    image_resolver,
                    &mut marker_paint_budget,
                    &mut marker_paint_budget_exceeded,
                )
            } else {
                child(*ser, "spPr").map(|_| {
                    parse_chartex_element_style(
                        *ser,
                        color_resolver,
                        None,
                        None,
                        image_resolver,
                        ChartImageSource::Chart,
                    )
                })
            };

            ChartSeries {
                name,
                chartex_format_idx: source_series_formatting_indices
                    .get(series_position)
                    .copied()
                    .and_then(|index| u32::try_from(index).ok()),
                values,
                source_hidden,
                color,
                fill_pattern: parse_series_pattern_fill(*ser, color_resolver),
                invert_if_negative,
                automatic_negative_style,
                inverted_fill,
                inverted_fill_hidden,
                inverted_fill_authored: inverted_format.map(|_| inverted_fill_authored),
                inverted_line_color,
                inverted_line_width_emu,
                inverted_line_hidden: inverted_line_no_fill.then_some(true),
                inverted_line_authored,
                // DrawingML line properties share the same local `spPr`
                // grammar as ChartEx. Reuse the bounded element-style carrier
                // so classic 3-D line/area rendering does not discard authored
                // dash/cap/join while color/width keep their legacy fields.
                chartex_style,
                line_color,
                line_width_emu,
                three_d_shape: child(*ser, "shape")
                    .and_then(|node| node.attribute("val"))
                    .filter(|value| {
                        matches!(
                            *value,
                            "box" | "cylinder" | "cone" | "coneToMax" | "pyramid" | "pyramidToMax"
                        )
                    })
                    .map(str::to_string),
                data_point_colors: if has_dpt_colors {
                    Some(data_point_colors)
                } else {
                    None
                },
                explosion: child(*ser, "explosion")
                    .and_then(|node| node.attribute("val"))
                    .and_then(|value| value.parse::<u32>().ok()),
                // Legacy `<c:chart>` per-point label colors are extracted via
                // `<c:dLbls><c:dLbl idx>` — not yet wired here; chartEx is the only
                // path that currently consumes this separate color array.
                data_label_colors: None,
                categories: series_categories,
                bubble_x_source_is_string,
                bubble_sizes,
                bubble_3d_group_default,
                bubble_3d,
                val_format_code,
                cat_format_code,
                cat_format_builtin_id,
                cat_format_codes,
                label_color,
                series_type,
                line_group_index,
                area_group_index,
                bar_group_index,
                bar_group_direction,
                bar_group_grouping,
                bar_group_gap_width,
                bar_group_overlap,
                // Shared `ChartSeries.use_secondary_axis` is `Option<bool>`; the
                // legacy default (false) is expressed as `None` so it drops off
                // the wire exactly as the old `skip_serializing_if = "Not::not"`
                // did.
                use_secondary_axis: if use_secondary_axis { Some(true) } else { None },
                // Marker styling / per-series data labels / error bars, now
                // populated by the shared extractors so both pptx and xlsx get
                // markers, dLbls and errBars from the one parse path.
                show_marker: Some(show_marker),
                marker_symbol,
                automatic_marker_symbol,
                marker_size,
                marker_fill,
                marker_fill_paint,
                marker_fill_paint_authored,
                marker_style,
                marker_line,
                marker_line_width_emu,
                marker_line_paint_authored,
                data_point_overrides: if data_point_overrides.is_empty() {
                    None
                } else {
                    Some(data_point_overrides)
                },
                data_label_overrides: if data_label_overrides.is_empty() {
                    None
                } else {
                    Some(data_label_overrides)
                },
                series_data_labels,
                err_bars: if err_bars.is_empty() {
                    None
                } else {
                    Some(err_bars)
                },
                // `<c:ser><c:smooth>` (§21.2.2.194) — line/area spline flag.
                // Shared with the xlsx parser via ooxml-common so both honor the
                // CT_Boolean implied-true semantics.
                smooth: extract_series_smooth(*ser).or(chart_smooth_default),
                // `<c:ser><c:trendline>` (§21.2.2.211) — regression lines. Shared
                // extractor; the line color resolves through the pptx theme.
                trend_lines: extract_series_trendlines_with_paint_policy(
                    *ser,
                    color_resolver,
                    allow_direct_label_paints,
                ),
                // §21.2.2.198 `<c:spPr><a:ln><a:noFill/>` — series connecting line
                // explicitly off. Only serialized when set, so byte-stable for
                // series that carry no line-off (the common case).
                line_hidden: if line_no_fill { Some(true) } else { None },
            }
        })
        .collect();

    // Structured marker paints are replayed once per visible point. Refuse the
    // chart atomically when their aggregate recipe budget is exceeded instead
    // of retaining an input-order-dependent prefix and silently making the
    // remaining authored markers transparent.
    if marker_paint_budget_exceeded {
        return None;
    }

    // Auto-title compatibility (ECMA-376 §21.2.2.7
    // `<c:autoTitleDeleted>`). The normative rule only says that a true value
    // suppresses the chart title; a false or absent value does not itself
    // create `<c:title>` (§21.2.2.210). Across the tested Office-produced
    // boundary cases, Word synthesizes text for an empty `<c:title>` frame but
    // leaves a named single-series chart untitled when that element is absent.
    // For a frame that carries `<c:txPr>` but no `<c:tx>` text, Word's observed
    // rule is:
    //   * exactly ONE series  → the auto title is that single series' name
    //   * two or more series   → NO auto title (a lone series name would be
    //                            misleading, so Word shows none)
    // We adopt only the single-series case; multi-series charts stay untitled,
    // matching Word. The title's `<a:defRPr cap="all">` would uppercase the
    // rendered glyphs ("PRODUCTION IN 2017"); chart-title `cap` is a display
    // transform we do not yet apply, so the model carries the series name
    // VERBATIM ("Production in 2017"). Making the title APPEAR is the goal; the
    // caps transform is a separate, tracked rendering-layer limitation.
    if title.is_none() && title_node_opt.is_some() {
        // §21.2.2.7 `<c:autoTitleDeleted>` — CT_Boolean, so a bare element ⇒ true
        // (the auto title is deleted, suppressing the single-series fallback
        // title); absent ⇒ false (the existing title frame remains eligible).
        let auto_title_deleted = chart_node
            .and_then(|c| bool_child(c, "autoTitleDeleted"))
            .unwrap_or(false);
        if !auto_title_deleted && series.len() == 1 {
            let ser_name = series[0].name.trim();
            if !ser_name.is_empty() {
                title = Some(ser_name.to_string());
            }
        }
    }

    // Chart-group data labels are on when a chart-level `<c:dLbls>` enables
    // `<c:showVal>` or `<c:showPercent>` (§21.2.2.189 / §21.2.2.187).
    // Series-local `<c:ser><c:dLbls>` is retained on that ChartSeries and must
    // not turn labels on for unlabelled sibling series.
    let show_data_labels = root
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "dLbls")
        .filter(|n| {
            n.parent()
                .map(|parent| parent.tag_name().name() != "ser")
                .unwrap_or(true)
        })
        .any(|d_lbls| {
            // `<c:showVal>` / `<c:showPercent>` are CT_Boolean: a present element
            // is ON unless `val` explicitly disables it (bare element ⇒ true), so
            // a bare `<c:showVal/>` enables data labels while `val="0"` does not.
            ["showVal", "showPercent"]
                .iter()
                .any(|name| bool_child(d_lbls, name).unwrap_or(false))
        });

    // Outer chartSpace spPr: we want the child of chartSpace (not plotArea).
    // When the `<c:spPr>` is PRESENT we honor whatever it resolves to (a
    // `<a:solidFill>` hex or, for `<a:noFill>` / an spPr with no fill child,
    // `None`). When it is ABSENT the file relies on the host default chart area
    // — Excel's opaque white vs. PowerPoint's transparent composite — supplied
    // by the resolver via `default_chart_bg`.
    let chart_sp_pr = root
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "spPr");
    let chart_fill_style = extract_direct_shape_fill_with_images(
        chart_sp_pr,
        color_resolver,
        image_resolver,
        ChartImageSource::Chart,
    );
    // `CT_ShapeProperties` carries a fill choice. Merely authoring another
    // property (commonly `<a:ln><a:noFill/></a:ln>`) does not override the
    // host application's default chart-area fill. A direct but unsupported
    // paint is still authored and therefore must not fall back to that default
    // or to a linked Chart Style paint.
    let chart_bg = if chart_fill_style.paint_authored == Some(true) {
        chart_fill_style.color
    } else {
        color_resolver.default_chart_bg()
    };
    let chart_fill = chart_fill_style.fill;
    let rounded_corners = extract_chart_space_rounded_corners(root);

    // <c:legend> + <c:legendPos val> — shared helper.
    let (show_legend, legend_pos) = extract_legend(root);

    // ECMA-376 §21.2.2.35: `<c:crossBetween>` lives on the VALUE axis (not cat),
    // and describes whether value gridlines land between or on category ticks.
    // Default is "between" (categories inset by half a step each side).
    let cat_axis_cross_between = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "valAx")
        .and_then(|ax| {
            ax.children()
                .find(|n| n.is_element() && n.tag_name().name() == "crossBetween")
        })
        .and_then(|n| attr(&n, "val"))
        .unwrap_or_else(|| "between".to_string());

    // Major tick marks (ECMA-376 §21.2.2.49 ST_TickMark, default "cross").
    // Schema default is `out` (ST_TickMark §21.2.3.48), shared with xlsx via
    // the ooxml-common helper so the two parsers don't diverge on the default.
    let read_major_tick_mark = |ax: Option<roxmltree::Node>| -> String {
        ax.map(|n| extract_axis_tick_mark_or_default(n, "majorTickMark"))
            .unwrap_or_else(|| "out".to_string())
    };
    let val_axis_major_tick_mark = read_major_tick_mark(val_ax);
    let cat_axis_major_tick_mark = read_major_tick_mark(cat_ax);

    // Axis-local text properties override the chart-wide `<c:chartSpace><c:txPr>`
    // defaults. Microsoft documents the latter as the OfficeArt text properties
    // for the entire chart, so an axis without its own `<c:txPr>` must inherit
    // these values rather than fall back to renderer constants.
    let chart_text_font_size_hpt = extract_axis_tick_label_size(root);
    let chart_text_font_color = extract_axis_tick_label_color(root, color_resolver);
    let chart_text_paint = chart_text_body_paint(child(root, "txPr"), color_resolver);
    let chart_text_font_bold = extract_axis_tick_label_bold(root);
    let chart_text_font_italic = extract_axis_tick_label_italic(root);
    let chart_text_font_face = extract_axis_tick_label_face(root);
    let cat_axis_font_size_hpt = cat_ax
        .and_then(extract_axis_tick_label_size)
        .or(chart_text_font_size_hpt);
    let val_axis_font_size_hpt = val_ax
        .and_then(extract_axis_tick_label_size)
        .or(chart_text_font_size_hpt);

    // Data-label font size — first `<c:dLbls><c:txPr>` defRPr/rPr@sz we find.
    let data_label_font_size_hpt = extract_data_label_font_size(root);
    let data_label_font_bold = extract_data_label_font_bold(root);
    let data_label_font_italic = extract_data_label_font_italic(root);

    // Bar gap / overlap, dLblPos and numFmt — all shared helpers so any new
    // chart property added to the xlsx side stays applied to pptx without
    // a manual host-specific port.
    let (bar_gap_width, bar_overlap) = extract_bar_gap_overlap(root);
    let data_label_position = extract_data_label_position(root);
    let data_label_format_code = extract_data_label_format_code(root);

    // Data-label font color uses the shared helper too — pptx supplies a
    // ColorResolver wrapper around `parse_color_node` so the
    // ECMA-376 §21.2.2.16 dLbls > txPr > solidFill walk lives in one place.
    let data_label_font_color = extract_data_label_font_color(root, color_resolver);
    let data_label_text_paint = root
        .descendants()
        .find(|node| node.is_element() && node.tag_name().name() == "dLbls")
        .map(|labels| chart_text_body_paint(child(labels, "txPr"), color_resolver))
        .unwrap_or_default();

    // Axis tick-label text color + axis-line style (color / width / noFill).
    // ECMA-376 §21.2.2.* — `<c:catAx|valAx><c:txPr>…<a:solidFill>` colors the
    // tick labels and `<c:spPr><a:ln>` styles the axis rule. Shared helpers so
    // category-label paint and the category-axis line resolve the same way in
    // all three hosts.
    // `CT_ChartSpace.style` is optional and has no schema default. Office
    // nevertheless renders the omitted form with style-2-compatible automatic
    // axis/mark paints; that observed compatibility layer is materialized below
    // as a role table. Keep the flat axis fields source-authored only so the
    // compatibility fallback cannot outrank direct formatting.
    let cat_axis_text_paint = cat_ax
        .map(|axis| chart_text_body_paint(child(axis, "txPr"), color_resolver))
        .unwrap_or_default();
    let val_axis_text_paint = val_ax
        .map(|axis| chart_text_body_paint(child(axis, "txPr"), color_resolver))
        .unwrap_or_default();
    let cat_axis_font_color = cat_ax
        .and_then(|n| extract_axis_tick_label_color(n, color_resolver))
        .or_else(|| {
            (!cat_axis_text_paint.authored)
                .then(|| chart_text_font_color.clone())
                .flatten()
        });
    let val_axis_font_color = val_ax
        .and_then(|n| extract_axis_tick_label_color(n, color_resolver))
        .or_else(|| {
            (!val_axis_text_paint.authored)
                .then(|| chart_text_font_color.clone())
                .flatten()
        });
    let (cat_axis_line_color, mut cat_axis_line_width_emu, cat_axis_line_hidden) = cat_ax
        .map(|n| extract_axis_line_style(n, color_resolver))
        .unwrap_or((None, None, false));
    let cat_axis_line_dash = cat_ax.and_then(extract_axis_line_dash);
    let cat_axis_line_paint_authored =
        cat_ax.and_then(|axis| extract_direct_shape_line(axis, color_resolver).paint_authored);
    let (val_axis_line_color, mut val_axis_line_width_emu, val_axis_line_hidden) = val_ax
        .map(|n| extract_axis_line_style(n, color_resolver))
        .unwrap_or((None, None, false));
    let val_axis_line_dash = val_ax.and_then(extract_axis_line_dash);
    let val_axis_line_paint_authored =
        val_ax.and_then(|axis| extract_direct_shape_line(axis, color_resolver).paint_authored);
    if legacy_chart_style == Some(2) {
        let cat_needs_theme_width = cat_ax.is_some_and(|axis| {
            !cat_axis_line_hidden
                && cat_axis_line_width_emu.is_none()
                && cat_axis_line_color.is_some()
                && child(axis, "spPr")
                    .and_then(|shape| child(shape, "ln"))
                    .is_some()
        });
        let val_needs_theme_width = val_ax.is_some_and(|axis| {
            !val_axis_line_hidden
                && val_axis_line_width_emu.is_none()
                && val_axis_line_color.is_some()
                && child(axis, "spPr")
                    .and_then(|shape| child(shape, "ln"))
                    .is_some()
        });
        if cat_needs_theme_width || val_needs_theme_width {
            let inherited_width = classic_style_two_axis_line_width_emu(color_resolver);
            if cat_needs_theme_width {
                cat_axis_line_width_emu = inherited_width;
            }
            if val_needs_theme_width {
                val_axis_line_width_emu = inherited_width;
            }
        }
    }

    // Plot groups carry the axis IDs and series ranges. A linked axis reads
    // its source format from the first series bound to that axis, in XML
    // order. A numRef's chart cache must not substitute for worksheet style;
    // an actual numLit formatCode is the source format for literal data.
    // Excel uses General when the first numLit omits formatCode; it does not
    // advance to a later source-backed series or use the authored axis code.
    // Observed in controlled workbooks rendered by Excel: one/two-series line
    // charts, line/bar mixes (including reversed c:order with unchanged ser
    // order), a literal-first series followed by a worksheet reference,
    // primary/secondary value axes, numeric category/date axes, and scatter/
    // bubble horizontal/vertical axes, with differing source styles and
    // explicit/omitted sourceLinked. A scatter series without xVal/yVal source
    // nodes retains the authored code. DOCX/PPTX have no worksheet
    // resolver, so only literal source formats can replace authored codes.
    let axis_source = |axis_role: &str, category: bool| {
        plot_groups
            .iter()
            .filter(|group| {
                (if category {
                    &group.category_axis
                } else {
                    &group.value_axis
                }) == axis_role
            })
            .find_map(|group| {
                let series = (group.series_count > 0)
                    .then(|| ser_nodes.get(group.series_start))
                    .flatten()?;
                let source_tag = if matches!(group.kind.as_str(), "scatter" | "bubble") {
                    if category {
                        "xVal"
                    } else {
                        "yVal"
                    }
                } else if category {
                    "cat"
                } else {
                    "val"
                };
                let source = child(*series, source_tag);
                Some(match source {
                    Some(source) if child(source, "numLit").is_some() => child(source, "numLit")
                        .and_then(|literal| child(literal, "formatCode"))
                        .and_then(|code| code.text())
                        .map(|code| AxisNumberFormatSource::Literal(code.to_string()))
                        .unwrap_or_else(|| AxisNumberFormatSource::Literal("General".to_string())),
                    Some(source) => reference_formula(source)
                        .map(AxisNumberFormatSource::Formula)
                        .unwrap_or(AxisNumberFormatSource::Unavailable),
                    None => AxisNumberFormatSource::Unavailable,
                })
            })
    };
    let primary_value_source = axis_source("primary", false);
    let primary_category_source = axis_source("primary", true);
    let secondary_value_source = axis_source("secondary", false);
    let secondary_category_source = axis_source("secondary", true);
    let val_axis_number_format = val_ax.and_then(axis_number_format);
    let cat_axis_number_format = cat_ax.and_then(axis_number_format);
    let val_axis_format_code = val_axis_number_format.as_ref().and_then(|format| {
        effective_axis_format_code(format, primary_value_source.as_ref(), references)
    });
    let cat_axis_format_code = cat_axis_number_format.as_ref().and_then(|format| {
        effective_axis_format_code(format, primary_category_source.as_ref(), references)
    });
    let val_axis_display_units =
        val_ax.and_then(|axis| parse_axis_display_units(axis, color_resolver));
    let cat_axis_display_units =
        cat_ax.and_then(|axis| parse_axis_display_units(axis, color_resolver));

    // Secondary value axis (combo charts) — parse the right-hand `<c:valAx>`
    // into a self-contained spec using the same shared helpers as the primary
    // axis. None for the common single value-axis case.
    let mut parse_auxiliary_value_axis = |ax, source: Option<&AxisNumberFormatSource>| {
        let number_format = axis_number_format(ax);
        let format_code = number_format
            .as_ref()
            .and_then(|format| effective_axis_format_code(format, source, references));
        let (min, max) = extract_axis_min_max(ax);
        let (t, title_size, title_bold, title_color) =
            extract_axis_title_with_props_resolved(ax, color_resolver);
        let resolved_title_bold = title_bold;
        let (line_color, line_width_emu, line_hidden) = extract_axis_line_style(ax, color_resolver);
        let line_dash = extract_axis_line_dash(ax);
        let line_paint_authored = extract_direct_shape_line(ax, color_resolver).paint_authored;
        let (minor_gridline_color, minor_gridline_width_emu, minor_gridline_dash) =
            extract_minor_gridline_style(ax, color_resolver);
        let (major_gridline_color, major_gridline_width_emu, major_gridline_dash) =
            extract_gridline_style(ax, color_resolver);
        let minor_gridline_paint_authored = child(ax, "minorGridlines")
            .and_then(|lines| extract_direct_shape_line(lines, color_resolver).paint_authored);
        let major_gridline_paint_authored = child(ax, "majorGridlines")
            .and_then(|lines| extract_direct_shape_line(lines, color_resolver).paint_authored);
        let (crosses, crosses_at) = extract_axis_crosses(ax);
        SecondaryValueAxis {
            style: parse_direct_chart_effect_style(ax, color_resolver),
            title_style: child(ax, "title")
                .and_then(|node| parse_direct_chart_effect_style(node, color_resolver)),
            major_gridline_style: child(ax, "majorGridlines")
                .and_then(|node| parse_direct_chart_effect_style(node, color_resolver)),
            minor_gridline_style: child(ax, "minorGridlines")
                .and_then(|node| parse_direct_chart_effect_style(node, color_resolver)),
            min,
            max,
            title: t,
            hidden: axis_is_deleted(ax),
            format_code,
            number_format,
            display_units: parse_axis_display_units(ax, color_resolver),
            font_color: extract_axis_tick_label_color(ax, color_resolver)
                .or_else(|| chart_text_font_color.clone()),
            font_paint_authored: {
                let direct = chart_text_body_paint(child(ax, "txPr"), color_resolver);
                let inherited = chart_text_body_paint(child(root, "txPr"), color_resolver);
                (direct.authored || inherited.authored).then_some(true)
            },
            font_size_hpt: extract_axis_tick_label_size(ax).or(chart_text_font_size_hpt),
            font_italic: extract_axis_tick_label_italic(ax).or(chart_text_font_italic),
            font_bold: extract_axis_tick_label_bold(ax).or(chart_text_font_bold),
            font_face: extract_axis_tick_label_face(ax),
            line_color,
            line_width_emu,
            line_dash,
            line_paint_authored,
            line_hidden,
            major_tick_mark: extract_axis_tick_mark_or_default(ax, "majorTickMark"),
            minor_tick_mark: extract_axis_tick_mark(ax, "minorTickMark"),
            minor_gridlines: axis_minor_gridlines_visible(ax),
            minor_gridline_color,
            minor_gridline_width_emu,
            minor_gridline_dash,
            minor_gridline_paint_authored,
            major_gridlines: axis_has_major_gridlines(ax),
            major_gridline_color,
            major_gridline_width_emu,
            major_gridline_dash,
            major_gridline_paint_authored,
            major_unit: extract_axis_major_unit(ax),
            minor_unit: extract_axis_minor_unit(ax),
            log_base: extract_axis_log_base(ax),
            orientation: extract_axis_orientation(ax),
            tick_label_pos: extract_axis_tick_label_pos(ax),
            label_alignment: (ax.tag_name().name() == "catAx")
                .then(|| child(ax, "lblAlgn"))
                .flatten()
                .and_then(|node| attr(&node, "val"))
                .filter(|value| matches!(value.as_str(), "l" | "ctr" | "r")),
            label_offset_percent: matches!(ax.tag_name().name(), "catAx" | "dateAx")
                .then(|| child(ax, "lblOffset"))
                .flatten()
                .map(|node| {
                    attr(&node, "val")
                        .as_deref()
                        .and_then(parse_unsigned_percent)
                        .unwrap_or(100)
                })
                .filter(|value| *value <= 1_000),
            tick_label_skip: child(ax, "tickLblSkip")
                .and_then(|node| node.attribute("val"))
                .and_then(|value| value.parse::<u32>().ok())
                .filter(|value| *value > 0),
            tick_mark_skip: child(ax, "tickMarkSkip")
                .and_then(|node| node.attribute("val"))
                .and_then(|value| value.parse::<u32>().ok())
                .filter(|value| *value > 0),
            crosses,
            crosses_at,
            title_font_size_hpt: title_size,
            title_font_bold: resolved_title_bold,
            title_font_italic: extract_axis_title_italic(ax),
            title_font_color: title_color,
            title_font_paint_authored: chart_title_text_paint(Some(ax), color_resolver)
                .authored
                .then_some(true),
            title_font_face: extract_axis_title_face(ax),
            title_rotation: extract_axis_title_rotation(ax),
            title_vertical_mode: extract_axis_title_vertical_mode(ax),
            title_manual_layout: extract_axis_title_manual_layout(ax),
        }
    };
    let secondary_val_axis = secondary_val_ax
        .map(|axis| parse_auxiliary_value_axis(axis, secondary_value_source.as_ref()));
    let secondary_cat_axis = secondary_cat_ax
        .map(|axis| parse_auxiliary_value_axis(axis, secondary_category_source.as_ref()));

    // `<c:plotArea><c:layout><c:manualLayout>` — use the shared parser so
    // schema defaults and all four layout modes cannot diverge by host format.
    let plot_area_manual_layout = plot_area
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "layout")
        .and_then(extract_manual_layout);

    // `<c:scatterChart><c:scatterStyle val>` — ECMA-376 §21.2.2.42. Lives
    // directly under scatterChart, so a plot_area descendant walk is enough.
    let scatter_style = if chart_type == "scatter" {
        plot_area
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "scatterStyle")
            .and_then(|n| attr(&n, "val"))
    } else {
        None
    };
    let bubble_scale = if chart_type == "bubble" {
        plot_area
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "bubbleScale")
            .and_then(|n| attr(&n, "val"))
            .and_then(|value| parse_unsigned_percent(&value))
            .filter(|value| *value <= 300)
    } else {
        None
    };
    let bubble_size_represents = if chart_type == "bubble" {
        plot_area
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "bubbleChart")
            .and_then(|group| child(group, "sizeRepresents"))
            .and_then(|n| attr(&n, "val"))
            .filter(|value| value == "area" || value == "w")
    } else {
        None
    };
    let show_negative_bubbles = if chart_type == "bubble" {
        plot_area
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "bubbleChart")
            .and_then(|group| bool_child(group, "showNegBubbles"))
    } else {
        None
    };

    // Axis titles + run props (ECMA-376 §21.2.2.210 `CT_Title`). Iterate every
    // `<c:catAx>`/`<c:valAx>` so the scatter case — two `<c:valAx>`, no
    // `<c:catAx>` — resolves correctly: a `<c:valAx>` whose `<c:axPos val>` is
    // `b`/`t` is the horizontal (X) axis → cat-axis title; `l`/`r` is the
    // vertical (Y) axis → val-axis title. A real `<c:catAx>` always feeds the
    // cat-axis title. First title wins for each axis (matches the xlsx parser).
    let mut cat_axis_title: Option<String> = None;
    let mut cat_axis_title_size: Option<i32> = None;
    let mut cat_axis_title_bold: Option<bool> = None;
    let mut cat_axis_title_italic: Option<bool> = None;
    let mut cat_axis_title_color: Option<String> = None;
    let mut cat_axis_title_face: Option<String> = None;
    let mut cat_axis_title_rotation: Option<i32> = None;
    let mut cat_axis_title_vertical_mode: Option<String> = None;
    let mut cat_axis_title_manual_layout: Option<ChartManualLayout> = None;
    let mut cat_axis_title_text_vertical_inset_emu: Option<i64> = None;
    let mut cat_axis_title_font_paint_authored: Option<bool> = None;
    let mut val_axis_title: Option<String> = None;
    let mut val_axis_title_size: Option<i32> = None;
    let mut val_axis_title_bold: Option<bool> = None;
    let mut val_axis_title_italic: Option<bool> = None;
    let mut val_axis_title_color: Option<String> = None;
    let mut val_axis_title_face: Option<String> = None;
    let mut val_axis_title_rotation: Option<i32> = None;
    let mut val_axis_title_vertical_mode: Option<String> = None;
    let mut val_axis_title_manual_layout: Option<ChartManualLayout> = None;
    let mut val_axis_title_text_vertical_inset_emu: Option<i64> = None;
    let mut val_axis_title_font_paint_authored: Option<bool> = None;
    for ax in plot_area
        .children()
        .filter(|n| n.is_element() && matches!(n.tag_name().name(), "catAx" | "dateAx" | "valAx"))
    {
        let is_cat = if matches!(ax.tag_name().name(), "catAx" | "dateAx") {
            true
        } else {
            // valAx: disambiguate by axPos (b/t → X/cat, l/r → Y/val).
            let ax_pos = ax
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "axPos")
                .and_then(|n| attr(&n, "val"))
                .unwrap_or_default();
            matches!(ax_pos.as_str(), "b" | "t")
        };
        if is_cat {
            if cat_axis_title.is_none() {
                let (t, sz, b, col) = extract_axis_title_with_props_resolved(ax, color_resolver);
                if t.is_some() {
                    cat_axis_title = t;
                    cat_axis_title_size = sz;
                    cat_axis_title_bold = b;
                    cat_axis_title_italic = extract_axis_title_italic(ax);
                    cat_axis_title_color = col;
                    cat_axis_title_face = extract_axis_title_face(ax);
                    cat_axis_title_rotation = extract_axis_title_rotation(ax);
                    cat_axis_title_vertical_mode = extract_axis_title_vertical_mode(ax);
                    cat_axis_title_manual_layout = extract_axis_title_manual_layout(ax);
                    cat_axis_title_text_vertical_inset_emu = extract_axis_title_vertical_inset(ax);
                    cat_axis_title_font_paint_authored =
                        chart_title_text_paint(Some(ax), color_resolver)
                            .authored
                            .then_some(true);
                }
            }
        } else if val_axis_title.is_none() {
            let (t, sz, b, col) = extract_axis_title_with_props_resolved(ax, color_resolver);
            if t.is_some() {
                val_axis_title = t;
                val_axis_title_size = sz;
                val_axis_title_bold = b;
                val_axis_title_italic = extract_axis_title_italic(ax);
                val_axis_title_color = col;
                val_axis_title_face = extract_axis_title_face(ax);
                val_axis_title_rotation = extract_axis_title_rotation(ax);
                val_axis_title_vertical_mode = extract_axis_title_vertical_mode(ax);
                val_axis_title_manual_layout = extract_axis_title_manual_layout(ax);
                val_axis_title_text_vertical_inset_emu = extract_axis_title_vertical_inset(ax);
                val_axis_title_font_paint_authored =
                    chart_title_text_paint(Some(ax), color_resolver)
                        .authored
                        .then_some(true);
            }
        }
    }

    // Axis tick-label bold flags (`<c:txPr>…defRPr@b`) and the chart-title bold
    // flag (`<c:title>…defRPr@b`). These were never serialized before; wiring
    // them through reaches parity with the xlsx parser so the renderer's
    // ST_Style bold handling applies uniformly. All three come from the shared
    // ooxml-common helpers so the two parsers stay in lockstep. The chart-title
    // bold helper expects the `<c:title>`'s parent, so pass `title_node_opt`'s
    // parent (the element that holds it as a direct child).
    let cat_axis_font_bold = cat_ax
        .and_then(extract_axis_tick_label_bold)
        .or(chart_text_font_bold);
    let cat_axis_font_italic = cat_ax
        .and_then(extract_axis_tick_label_italic)
        .or(chart_text_font_italic);
    let val_axis_font_bold = val_ax
        .and_then(extract_axis_tick_label_bold)
        .or(chart_text_font_bold);
    let val_axis_font_italic = val_ax
        .and_then(extract_axis_tick_label_italic)
        .or(chart_text_font_italic);
    let title_font_bold = title_node_opt
        .and_then(|t| t.parent())
        .and_then(extract_chart_title_bold);
    let title_font_italic = title_node_opt
        .and_then(|t| t.parent())
        .and_then(extract_chart_title_italic);

    // Explicit chartSpace border from `<c:chartSpace><c:spPr><a:ln>` (ECMA-376
    // §21.2.2.5 / DrawingML §20.1.2.2.24). Resolve the complete DrawingML line
    // color grammar through the package theme: chart borders commonly use
    // `<a:schemeClr val="tx1">` (often with luminance transforms), not only a
    // literal srgb color. `<a:noFill/>` remains an explicit invisible border.
    let chart_line_style = extract_direct_shape_line(root, color_resolver);

    // `<c:date1904>` (ECMA-376 §21.2.2.38) — direct child of `<c:chartSpace>`
    // (`root`). Shared with the xlsx parser via ooxml-common so both honor the
    // CT_Boolean implied-true semantics.
    let date1904 = extract_chart_date1904(root);

    // `<c:chart><c:dispBlanksAs>` (ECMA-376 §21.2.2.42) — null-cell plotting for
    // line/area. Shared with the xlsx parser via ooxml-common.
    let disp_blanks_as = extract_disp_blanks_as(root);
    let show_data_labels_over_max = extract_show_data_labels_over_max(root);

    // ── Chart text font faces (CH10) ────────────────────────────────────────
    // Tick-label faces (`<c:catAx|valAx><c:txPr>…<a:latin>`), data-label face
    // (`<c:dLbls><c:txPr>…<a:latin>`) and legend text props, all via the shared
    // ooxml-common extractors so pptx/xlsx stay in lockstep. Absent faces stay
    // None; the renderer falls back to the theme body/heading font.
    let cat_axis_font_face = cat_ax
        .and_then(extract_axis_tick_label_face)
        .or_else(|| chart_text_font_face.clone());
    let val_axis_font_face = val_ax
        .and_then(extract_axis_tick_label_face)
        .or_else(|| chart_text_font_face.clone());
    let data_label_font_face = extract_data_label_face(root);
    let (legend_font_face, legend_font_size_hpt, legend_font_bold, legend_font_italic) =
        extract_legend_text_props(root);
    let legend_font_color = { extract_legend_font_color(root, color_resolver) };
    let legend_font_paint_authored = root
        .descendants()
        .find(|node| node.is_element() && node.tag_name().name() == "legend")
        .map(|legend| chart_text_body_paint(child(legend, "txPr"), color_resolver))
        .is_some_and(|paint| paint.authored)
        .then_some(true);
    let legend_frame = extract_legend_frame_style(root, color_resolver);
    // Theme fallback fonts: the resolver supplies the theme's major/minor Latin
    // faces (pptx keys them `+mj-lt` / `+mn-lt` in its color+font map). None
    // when the theme lacks a fontScheme. The renderer uses these when a chart
    // text run carries no explicit face.
    let theme_major_font_latin = color_resolver.theme_major_font_latin();
    let theme_minor_font_latin = color_resolver.theme_minor_font_latin();

    // ── Pie / doughnut geometry (CH8) ───────────────────────────────────────
    // holeSize (doughnut) / firstSliceAng (pie + doughnut), shared extractors.
    let hole_size = extract_hole_size(root);
    let first_slice_angle = extract_first_slice_angle(root);

    // ── Axis scale model (CH6) ──────────────────────────────────────────────
    // Gridline presence, manual major/minor units, log scale and orientation —
    // all via the shared ooxml-common extractors on the primary val/cat axes.
    // `<c:majorGridlines>` presence: Office writes it on the value axis by
    // default (renderer keeps its historical always-on when the field is None),
    // so we only emit `Some(false)` when a value axis EXISTS without the element.
    let val_axis_major_gridlines = val_ax.map(axis_major_gridlines_visible);
    let cat_axis_major_gridlines = cat_ax.map(axis_major_gridlines_visible);
    // `<c:majorGridlines><c:spPr><a:ln>` colour/width — the explicit gridline
    // style (for example `accent3` with a 0.25 pt value-axis line).
    // `(None, None)` when absent, so the renderer keeps its faint default.
    let (val_axis_gridline_color, val_axis_gridline_width_emu, val_axis_gridline_dash) = val_ax
        .map(|ax| extract_gridline_style(ax, color_resolver))
        .unwrap_or((None, None, None));
    let val_axis_gridline_paint_authored = val_ax
        .and_then(|axis| child(axis, "majorGridlines"))
        .and_then(|lines| extract_direct_shape_line(lines, color_resolver).paint_authored);
    let (cat_axis_gridline_color, cat_axis_gridline_width_emu, cat_axis_gridline_dash) = cat_ax
        .map(|ax| extract_gridline_style(ax, color_resolver))
        .unwrap_or((None, None, None));
    let cat_axis_gridline_paint_authored = cat_ax
        .and_then(|axis| child(axis, "majorGridlines"))
        .and_then(|lines| extract_direct_shape_line(lines, color_resolver).paint_authored);
    let val_axis_minor_gridlines = val_ax.map(axis_minor_gridlines_visible);
    let (
        val_axis_minor_gridline_color,
        val_axis_minor_gridline_width_emu,
        val_axis_minor_gridline_dash,
    ) = val_ax
        .map(|axis| extract_minor_gridline_style(axis, color_resolver))
        .unwrap_or((None, None, None));
    let val_axis_minor_gridline_paint_authored = val_ax
        .and_then(|axis| child(axis, "minorGridlines"))
        .and_then(|lines| extract_direct_shape_line(lines, color_resolver).paint_authored);
    let cat_axis_minor_gridlines = cat_ax.map(axis_minor_gridlines_visible);
    let (
        cat_axis_minor_gridline_color,
        cat_axis_minor_gridline_width_emu,
        cat_axis_minor_gridline_dash,
    ) = cat_ax
        .map(|axis| extract_minor_gridline_style(axis, color_resolver))
        .unwrap_or((None, None, None));
    let cat_axis_minor_gridline_paint_authored = cat_ax
        .and_then(|axis| child(axis, "minorGridlines"))
        .and_then(|lines| extract_direct_shape_line(lines, color_resolver).paint_authored);
    let val_axis_major_unit = val_ax.and_then(extract_axis_major_unit);
    let val_axis_minor_unit = val_ax.and_then(extract_axis_minor_unit);
    let cat_axis_major_unit = cat_ax.and_then(extract_axis_major_unit);
    let cat_axis_minor_unit = cat_ax.and_then(extract_axis_minor_unit);
    let cat_axis_is_date = real_cat_ax
        .filter(|axis| axis.tag_name().name() == "dateAx")
        .map(|_| true);
    let date_axis_time_unit = |name: &str| {
        real_cat_ax
            .filter(|axis| axis.tag_name().name() == "dateAx")
            .and_then(|axis| child(axis, name))
            // CT_TimeUnit@val defaults to `days` when the child is present
            // without an explicit value.
            .map(|unit| attr(&unit, "val").unwrap_or_else(|| "days".to_string()))
    };
    let cat_axis_base_time_unit = date_axis_time_unit("baseTimeUnit");
    let cat_axis_major_time_unit = date_axis_time_unit("majorTimeUnit");
    let cat_axis_minor_time_unit = date_axis_time_unit("minorTimeUnit");
    let cat_axis_no_multi_level_labels = cat_ax.and_then(|axis| bool_child(axis, "noMultiLvlLbl"));
    let val_axis_log_base = val_ax.and_then(extract_axis_log_base);
    let cat_axis_log_base = cat_ax.and_then(extract_axis_log_base);
    let val_axis_orientation = val_ax.and_then(extract_axis_orientation);
    let cat_axis_orientation = cat_ax.and_then(extract_axis_orientation);
    let cat_axis_tick_label_skip = cat_ax
        .and_then(|axis| child(axis, "tickLblSkip"))
        .and_then(|skip| attr(&skip, "val"))
        .and_then(|skip| skip.parse::<u32>().ok())
        .filter(|skip| *skip > 0);
    let cat_axis_tick_mark_skip = cat_ax
        .and_then(|axis| child(axis, "tickMarkSkip"))
        .and_then(|skip| attr(&skip, "val"))
        .and_then(|skip| skip.parse::<u32>().ok())
        .filter(|skip| *skip > 0);
    let cat_axis_label_alignment = real_cat_ax
        .filter(|axis| axis.tag_name().name() == "catAx")
        .and_then(|axis| child(axis, "lblAlgn"))
        .and_then(|alignment| attr(&alignment, "val"))
        .filter(|alignment| matches!(alignment.as_str(), "l" | "ctr" | "r"));
    let cat_axis_label_offset_percent = real_cat_ax
        .and_then(|axis| child(axis, "lblOffset"))
        .map(|offset| {
            attr(&offset, "val")
                .as_deref()
                .and_then(parse_unsigned_percent)
                // CT_LblOffset@val defaults to 100 when the element is present.
                .unwrap_or(100)
        })
        .filter(|offset| *offset <= 1_000);
    let cat_axis_tick_label_pos = cat_ax.and_then(extract_axis_tick_label_pos);
    let val_axis_tick_label_pos = val_ax.and_then(extract_axis_tick_label_pos);
    let cat_axis_label_rotation = cat_ax.and_then(extract_axis_tick_label_rotation);

    // Chart title font face (`<c:title>…<a:latin>`) — parity with xlsx, which
    // already extracts it. `extract_axis_title_face` scopes to a node's
    // direct-child `<c:title>`, so pass the title's parent (`<c:chart>`).
    let title_font_face = title_node_opt
        .and_then(|t| t.parent())
        .and_then(extract_axis_title_face);

    // Minor tick marks (ECMA-376 §21.2.2.115) — raw ST_TickMark string, `None`
    // when the axis omits `<c:minorTickMark>` (renderer default applies).
    let cat_axis_minor_tick_mark = cat_ax.and_then(|n| extract_axis_tick_mark(n, "minorTickMark"));
    let val_axis_minor_tick_mark = val_ax.and_then(|n| extract_axis_tick_mark(n, "minorTickMark"));

    // Axis crossing (`<c:crosses>` / `<c:crossesAt>`, ECMA-376 §21.2.2.33/.34).
    let (cat_axis_crosses, cat_axis_crosses_at) =
        cat_ax.map(extract_axis_crosses).unwrap_or((None, None));
    let (val_axis_crosses, val_axis_crosses_at) =
        val_ax.map(extract_axis_crosses).unwrap_or((None, None));

    // Category-axis explicit scaling bounds (`<c:scaling><c:min|max>`).
    let (cat_axis_min, cat_axis_max) = cat_ax.map(extract_axis_min_max).unwrap_or((None, None));

    // `<c:radarChart><c:radarStyle>` (ECMA-376 §21.2.3.10).
    let radar_style = extract_radar_style(root);

    // Legend `<c:layout><c:manualLayout>` (ECMA-376 §21.2.2.31).
    let legend_manual_layout = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "legend")
        .and_then(extract_legend_manual_layout);
    let (legend_overlay, legend_entries) = extract_legend_overrides(root, color_resolver);

    // Chart-title `<c:title><c:layout><c:manualLayout>` (ECMA-376 §21.2.2.88).
    let title_manual_layout = title_node_opt
        .and_then(|t| child(t, "layout"))
        .and_then(extract_manual_layout);

    // §21.2.2.227 varyColors chart-level flag. Pie/doughnut preserve the
    // effective boolean because an explicit false makes unspecified points
    // inherit the single series fill. Lone bar/column charts expose only an
    // explicitly enabled case; §21.2.2.227 defaults the `val` attribute to
    // true when the element exists, not the element itself when absent.
    let vary_colors = {
        let is_pie_family = matches!(chart_type.as_str(), "pie" | "doughnut" | "ofPie");
        let is_bar_family = chart_type.contains("Bar");
        if is_pie_family && first_series.is_some() {
            Some(
                first_series
                    .and_then(|series| series.parent())
                    .and_then(|group| bool_child(group, "varyColors"))
                    .unwrap_or(true),
            )
        } else if is_bar_family && series_count == 1 {
            let vary = first_series
                .and_then(|series| series.parent())
                .and_then(|g| bool_child(g, "varyColors"));
            if vary == Some(true) {
                Some(true)
            } else {
                None
            }
        } else {
            None
        }
    };

    let title_present = title_node_opt.is_some() || title.is_some();
    // ECMA-376 §21.2.3.46 applies Tables 5–6 by the formatting index of the
    // painted object. An effectively varying pie/doughnut (including each ring
    // of a multi-series doughnut) or lone bar series has a point-index domain;
    // series-owned roles retain the source `c:ser@idx` domain. Do not combine
    // the two: Table 5's Fade recipe depends on the highest formatting index,
    // so an unused sparse series index must not recolour otherwise identical
    // varying points. The multi-ring behaviour is observed in Excel-produced
    // classic doughnut charts with two series and six points per ring.
    let group_varies_by_point = |group: &ChartPlotGroup| match group.kind.as_str() {
        "pie" | "pie3D" | "doughnut" | "ofPie" => group.vary_colors.unwrap_or(true),
        "bubble" => group.series_count == 1 && group.vary_colors.unwrap_or(true),
        "bar" | "bar3D" | "line" | "scatter" => {
            group.series_count == 1 && group.vary_colors == Some(true)
        }
        // Observed Office behaviour: lone standard/marker radar groups use
        // point formatting when c:varyColors=1, whereas filled radar remains
        // one series-owned polygon. Keep the compatibility rule no broader
        // than the Office-produced true/false counterexamples establish.
        "radar" => {
            group.radar_style.as_deref() != Some("filled")
                && group.series_count == 1
                && group.vary_colors == Some(true)
        }
        _ => false,
    };
    // Office hosts and the existing compatibility modules treat an omitted
    // `c:style` as built-in style 2. Keep `legacy_chart_style` itself optional
    // for source provenance while materializing the effective fallback table.
    let mut classic_chart_style_roles =
        classic_style::resolve_classic_chart_style_roles_with_images(
            legacy_chart_style.unwrap_or(2),
            color_resolver,
            image_resolver,
            chart_text_font_size_hpt,
            &source_series_formatting_indices,
            None,
        );
    if legacy_chart_style.is_none() {
        // `c:style` is optional and has no schema default. Office retains its
        // style-2-compatible automatic mark/axis paints when it is omitted,
        // but does not apply Table 1's title typography. Legacy line and stock
        // counterexamples remain regular, so retain a regular-weight default
        // below the direct and chart-space text layers instead of falling into
        // the renderer's API-level historical bold fallback.
        if let Some(title) = classic_chart_style_roles
            .as_mut()
            .and_then(|roles| roles.get_mut("title"))
        {
            title.font_size_hpt = None;
            title.font_bold = Some(false);
            title.font_italic = None;
            title.font_color = None;
            title.font_colors = None;
            title.font_color_index = None;
            title.font_formatting_indices = None;
            title.font_paint_authored = None;
            title.font_hidden = None;
            title.font_face = None;
            title.font_language = None;
            title.font_baseline = None;
        }
    }
    if let (Some(style), Some(roles)) = (legacy_chart_style, classic_chart_style_roles.as_mut()) {
        let title_has_paragraph_default_run =
            title_node_opt.is_some_and(title_rich_text_has_observed_paragraph_default_run);
        classic_style::apply_office_dark_text_contrast(
            style,
            title_has_paragraph_default_run,
            color_resolver,
            roles,
        );
    }
    // Surface value bands are neither source series nor source points. Their
    // count becomes final only after the Canvas renderer plans the value axis,
    // so preserve a bounded band-domain numeric role here instead of replaying
    // a series-domain palette by modulo. Pattern roles are stable by index;
    // Table 5 Fade roles depend on the final highest band index and therefore
    // retain the complete evidenced 1..48 count lattice. Unsupported Pattern-2 sets
    // return to semantic automatic paint at each host's evidence boundary
    // (after six objects for Word/PowerPoint, or 48 for Excel).
    let effective_classic_style = legacy_chart_style.unwrap_or(2);
    let classic_surface_band_styles = matches!(chart_type.as_str(), "surface" | "surface3D")
        .then(|| {
            let resolve = |band_count| {
                if surface_wireframe == Some(true) {
                    classic_style::resolve_classic_surface_wireframe_band_style_with_images(
                        effective_classic_style,
                        color_resolver,
                        image_resolver,
                        band_count,
                    )
                } else {
                    classic_style::resolve_classic_surface_band_style_with_images(
                        effective_classic_style,
                        color_resolver,
                        image_resolver,
                        band_count,
                    )
                }
            };
            if classic_style::surface_band_palette_depends_on_count(effective_classic_style) {
                (1..=48)
                    .map(resolve)
                    .collect::<Option<Vec<_>>>()
                    .map(|by_band_count| ChartClassicSurfaceBandStyles {
                        fixed: None,
                        by_band_count: Some(by_band_count),
                    })
            } else {
                resolve(48).map(|fixed| ChartClassicSurfaceBandStyles {
                    fixed: Some(fixed),
                    by_band_count: None,
                })
            }
        })
        .flatten();
    let varying_point_counts = plot_groups
        .iter()
        .map(|group| {
            if !group_varies_by_point(group) {
                return None;
            }
            let point_count = series
                .get(group.series_start..group.series_start.saturating_add(group.series_count))
                .unwrap_or_default()
                .iter()
                .map(|item| {
                    item.values
                        .len()
                        .max(item.categories.as_ref().map_or(0, Vec::len))
                })
                .max()
                .unwrap_or(0)
                .max(1);
            Some(point_count)
        })
        .collect::<Vec<_>>();
    // Store the most common bounded point domain once in the singular field.
    // Group slots inherit it with `None`; only exceptional domain sizes carry
    // their own table. This keeps 10,000 identical vary-by-point groups
    // O(groups + palette), rather than cloning the full palette 10,000 times.
    let mut point_count_frequencies = BTreeMap::<usize, usize>::new();
    for point_count in varying_point_counts.iter().flatten().copied() {
        if point_count <= MAX_CHART_COLOR_STYLE_ENTRIES {
            *point_count_frequencies.entry(point_count).or_default() += 1;
        }
    }
    let common_point_count = point_count_frequencies
        .into_iter()
        .max_by(
            |(left_count, left_frequency), (right_count, right_frequency)| {
                left_frequency
                    .cmp(right_frequency)
                    .then_with(|| right_count.cmp(left_count))
            },
        )
        .map(|(point_count, _)| point_count);
    let resolve_varying_roles = |point_count: usize| {
        let point_indices = (0..point_count).collect::<Vec<_>>();
        classic_style::resolve_classic_varying_point_roles_with_images(
            legacy_chart_style.unwrap_or(2),
            color_resolver,
            image_resolver,
            chart_text_font_size_hpt,
            &source_series_formatting_indices,
            &point_indices,
        )
        .unwrap_or_default()
    };
    let classic_varying_point_chart_style_roles = common_point_count.map(resolve_varying_roles);
    let mut varying_roles_by_point_count = BTreeMap::new();
    let classic_varying_point_chart_style_roles_by_group = varying_point_counts
        .into_iter()
        .map(|point_count| match point_count {
            None => None,
            Some(point_count) if point_count > MAX_CHART_COLOR_STYLE_ENTRIES => {
                // An empty map is the resource-refusal sentinel. It prevents
                // fallback to the common or series-domain numeric palette.
                Some(BTreeMap::new())
            }
            Some(point_count) if Some(point_count) == common_point_count => None,
            Some(point_count) => Some(
                varying_roles_by_point_count
                    .entry(point_count)
                    .or_insert_with(|| resolve_varying_roles(point_count))
                    .clone(),
            ),
        })
        .collect::<Vec<_>>();
    let classic_varying_point_chart_style_roles_by_group =
        classic_varying_point_chart_style_roles_by_group
            .iter()
            .any(Option::is_some)
            .then_some(classic_varying_point_chart_style_roles_by_group);
    Some(ChartModel {
        chart_type,
        title,
        title_rich_runs,
        title_present,
        authored_without_series: !has_nonempty_classic_group,
        categories,
        category_source_hidden,
        category_levels,
        series,
        plot_groups: Some(plot_groups),
        vary_colors,
        chart_text_boxes: None,
        chart_text_style: extract_chart_space_text_style(root, color_resolver),
        chart_area_style: parse_direct_chart_effect_style(root, color_resolver),
        plot_area_style: parse_direct_chart_effect_style(plot_area, color_resolver),
        legend_style: root
            .descendants()
            .find(|node| node.is_element() && node.tag_name().name() == "legend")
            .and_then(|node| parse_direct_chart_effect_style(node, color_resolver)),
        title_style: root
            .descendants()
            .find(|node| node.is_element() && node.tag_name().name() == "title")
            .and_then(|node| parse_direct_chart_effect_style(node, color_resolver)),
        cat_axis_style: cat_ax
            .and_then(|node| parse_direct_chart_effect_style(node, color_resolver)),
        val_axis_style: val_ax
            .and_then(|node| parse_direct_chart_effect_style(node, color_resolver)),
        cat_axis_title_style: cat_ax
            .and_then(|axis| child(axis, "title"))
            .and_then(|node| parse_direct_chart_effect_style(node, color_resolver)),
        val_axis_title_style: val_ax
            .and_then(|axis| child(axis, "title"))
            .and_then(|node| parse_direct_chart_effect_style(node, color_resolver)),
        cat_axis_major_gridline_style: cat_ax
            .and_then(|axis| child(axis, "majorGridlines"))
            .and_then(|node| parse_direct_chart_effect_style(node, color_resolver)),
        cat_axis_minor_gridline_style: cat_ax
            .and_then(|axis| child(axis, "minorGridlines"))
            .and_then(|node| parse_direct_chart_effect_style(node, color_resolver)),
        val_axis_major_gridline_style: val_ax
            .and_then(|axis| child(axis, "majorGridlines"))
            .and_then(|node| parse_direct_chart_effect_style(node, color_resolver)),
        val_axis_minor_gridline_style: val_ax
            .and_then(|axis| child(axis, "minorGridlines"))
            .and_then(|node| parse_direct_chart_effect_style(node, color_resolver)),
        val_max,
        val_min,
        subtotal_indices: vec![],
        show_data_labels,
        cat_axis_hidden,
        val_axis_hidden,
        plot_area_bg,
        plot_area_fill: plot_area_fill_style.fill,
        plot_area_fill_hidden: plot_area_fill_style.hidden,
        plot_area_fill_paint_authored: plot_area_fill_style.paint_authored,
        plot_area_fill_automatic,
        plot_area_line_color: plot_area_line_style.color,
        plot_area_line_fill: plot_area_line_style.fill,
        plot_area_line_width_emu: plot_area_line_style.width_emu,
        plot_area_line_dash: plot_area_line_style.dash,
        plot_area_line_dash_authored: plot_area_line_style.dash_authored,
        plot_area_line_custom_dash: plot_area_line_style.custom_dash,
        plot_area_line_cap: plot_area_line_style.cap,
        plot_area_line_join: plot_area_line_style.join,
        plot_area_line_compound: plot_area_line_style.compound,
        plot_area_line_hidden: plot_area_line_style.hidden,
        plot_area_line_paint_authored: plot_area_line_style.paint_authored,
        chart_bg,
        chart_fill,
        chart_fill_hidden: chart_fill_style.hidden,
        chart_fill_paint_authored: chart_fill_style.paint_authored,
        rounded_corners,
        plot_visible_only,
        show_legend,
        data_table,
        cat_axis_cross_between,
        val_axis_major_tick_mark,
        cat_axis_major_tick_mark,
        title_font_size_hpt,
        title_font_color,
        title_font_paint_authored: chart_title_text_paint(chart_node, color_resolver)
            .authored
            .then_some(true),
        title_font_face,
        cat_axis_font_size_hpt,
        val_axis_font_size_hpt,
        cat_axis_font_color,
        cat_axis_font_paint_authored: (cat_axis_text_paint.authored || chart_text_paint.authored)
            .then_some(true),
        val_axis_font_color,
        val_axis_font_paint_authored: (val_axis_text_paint.authored || chart_text_paint.authored)
            .then_some(true),
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
        bar_overlap,
        data_label_position,
        data_label_font_color,
        data_label_font_paint_authored: data_label_text_paint.authored.then_some(true),
        data_label_format_code,
        data_label_font_bold,
        data_label_font_italic,
        data_label_font_language: None,
        data_label_font_baseline: None,
        val_axis_format_code,
        val_axis_number_format,
        val_axis_display_units,
        cat_axis_display_units,
        plot_area_manual_layout,
        cartesian_auto_layout_profile: None,
        scatter_style,
        bubble_scale,
        bubble_size_represents,
        show_negative_bubbles,
        cat_axis_title,
        val_axis_title,
        // TS `ChartElement` renamed the axis-title run-prop fields to the
        // core `ChartModel` names (`…TitleFontSizeHpt/Bold/Color`); the
        // parser locals keep the shorter legacy names.
        cat_axis_title_font_size_hpt: cat_axis_title_size,
        cat_axis_title_font_bold: cat_axis_title_bold,
        cat_axis_title_font_italic: cat_axis_title_italic,
        cat_axis_title_font_color: cat_axis_title_color,
        cat_axis_title_font_paint_authored,
        cat_axis_title_rotation,
        cat_axis_title_vertical_mode,
        cat_axis_title_manual_layout,
        cat_axis_title_text_vertical_inset_emu,
        val_axis_title_font_size_hpt: val_axis_title_size,
        val_axis_title_font_bold: val_axis_title_bold,
        val_axis_title_font_italic: val_axis_title_italic,
        val_axis_title_font_color: val_axis_title_color,
        val_axis_title_font_paint_authored,
        val_axis_title_rotation,
        val_axis_title_vertical_mode,
        val_axis_title_manual_layout,
        val_axis_title_text_vertical_inset_emu,
        title_font_bold,
        title_font_italic,
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
        secondary_val_axis,
        secondary_cat_axis,
        // Pie/doughnut geometry (CH8) + chart text font faces (CH10).
        hole_size,
        first_slice_angle,
        cat_axis_font_face,
        val_axis_font_face,
        cat_axis_title_font_face: cat_axis_title_face,
        val_axis_title_font_face: val_axis_title_face,
        data_label_font_face,
        legend_font_face,
        legend_font_color,
        legend_font_paint_authored,
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
        theme_major_font_latin,
        theme_minor_font_latin,
        // ChartModel fields the legacy pptx `<c:chart>` path leaves unset
        // (they were never in the pptx `ChartElement` copy, so they defaulted
        // to `undefined` on the TS side and stay absent on the wire).
        val_axis_minor_tick_mark,
        cat_axis_minor_tick_mark,
        legend_manual_layout,
        legend_overlay,
        legend_entries,
        title_manual_layout,
        cat_axis_crosses,
        cat_axis_crosses_at,
        val_axis_crosses,
        val_axis_crosses_at,
        cat_axis_format_code,
        cat_axis_number_format,
        cat_axis_min,
        cat_axis_max,
        radar_style,
        date1904,
        disp_blanks_as,
        show_data_labels_over_max,
        // ── Axis scale model (CH6) ──────────────────────────────────────
        val_axis_major_gridlines,
        cat_axis_major_gridlines,
        val_axis_gridline_color,
        val_axis_gridline_width_emu,
        val_axis_gridline_dash,
        val_axis_gridline_paint_authored,
        cat_axis_gridline_color,
        cat_axis_gridline_width_emu,
        cat_axis_gridline_dash,
        cat_axis_gridline_paint_authored,
        val_axis_minor_gridlines,
        val_axis_minor_gridline_color,
        val_axis_minor_gridline_width_emu,
        val_axis_minor_gridline_dash,
        val_axis_minor_gridline_paint_authored,
        cat_axis_minor_gridlines,
        cat_axis_minor_gridline_color,
        cat_axis_minor_gridline_width_emu,
        cat_axis_minor_gridline_dash,
        cat_axis_minor_gridline_paint_authored,
        val_axis_major_unit,
        val_axis_minor_unit,
        cat_axis_major_unit,
        cat_axis_minor_unit,
        cat_axis_is_date,
        cat_axis_base_time_unit,
        cat_axis_major_time_unit,
        cat_axis_minor_time_unit,
        cat_axis_no_multi_level_labels,
        val_axis_log_base,
        cat_axis_log_base,
        val_axis_orientation,
        cat_axis_orientation,
        cat_axis_tick_label_pos,
        cat_axis_tick_label_skip,
        cat_axis_tick_mark_skip,
        cat_axis_label_alignment,
        cat_axis_label_offset_percent,
        val_axis_tick_label_pos,
        cat_axis_label_rotation,
        line_group_decorations,
        area_group_decorations,
        bar_group_decorations,
        stock_drop_lines,
        stock_hi_low_line_style,
        stock_hi_low_lines,
        stock_hi_low_line_color,
        stock_up_down_bars,
        stock_up_down_bar_style,
        stock_automatic_style,
        surface_wireframe,
        surface_band_formats,
        classic_surface_band_styles,
        legacy_chart_style,
        theme_accent_colors,
        of_pie,
        three_d,
        // Legacy `c:` charts never carry the chartEx structured models.
        chartex_box: None,
        chartex_sunburst: None,
        chartex_treemap: None,
        chartex_region_map: None,
        chartex_histogram_binning: None,
        chartex_accents: None,
        chart_style_roles,
        classic_chart_style_roles,
        classic_varying_point_chart_style_roles,
        classic_varying_point_chart_style_roles_by_group,
        chart_style_color_palette,
        chart_style_color_method,
        chart_style_marker_size_pt,
        chart_style_marker_symbol,
        chartex_color_palette: None,
        chartex_color_style_method: None,
        chartex_data_point_style: None,
        chartex_data_point_line_style: None,
        chartex_series_line_style: None,
        chartex_data_point_marker_style: None,
        chartex_marker_size_pt: None,
        chartex_marker_symbol: None,
        chartex_connector_lines: None,
    })
}
