use super::*;

pub(super) fn chart_text_bool_attr(node: Node, name: &str) -> Option<bool> {
    node.attribute(name)
        .map(|value| matches!(value, "1" | "true" | "on"))
}

/** Resolve one character-property carrier at the boundary with chart styles.
 * `b` and `i` have no XML-schema default. Observed Office behavior distinguishes
 * authored visual formatting from the empty/language-only carriers that Office
 * routinely writes on otherwise style-owned chart text: the former resolves an
 * omitted boolean to regular, while the latter leaves the chart-style role
 * eligible. */
pub(super) fn chart_text_bool_from_present_props(
    node: Option<Node<'_, '_>>,
    name: &str,
) -> Option<bool> {
    chart_text_bool_from_property_cascade([node], name)
}

pub(super) fn chart_text_props_own_visual_formatting(props: Node<'_, '_>) -> bool {
    // CT_TextCharacterProperties metadata does not alter glyph appearance and
    // must not accidentally suppress a numeric/linked chart style. All other
    // attributes are visual or metric character formatting. Hyperlinks and an
    // extension list likewise do not establish the run's typography.
    props.attributes().any(|attribute| {
        !matches!(
            attribute.name(),
            "lang" | "altLang" | "noProof" | "dirty" | "err" | "smtClean" | "smtId" | "bmk"
        )
    }) || props.children().any(|child| {
        child.is_element()
            && !matches!(
                child.tag_name().name(),
                "hlinkClick" | "hlinkMouseOver" | "extLst"
            )
    })
}

pub(super) fn chart_text_bool_from_property_cascade<'a, 'input: 'a>(
    nodes: impl IntoIterator<Item = Option<Node<'a, 'input>>>,
    name: &str,
) -> Option<bool> {
    let mut visual_formatting_present = false;
    for node in nodes.into_iter().flatten() {
        if let Some(value) = chart_text_bool_attr(node, name) {
            return Some(value);
        }
        visual_formatting_present |= chart_text_props_own_visual_formatting(node);
    }
    visual_formatting_present.then_some(false)
}

/// Select the first character-property carrier. Whether it owns a requested
/// boolean is decided separately, so empty and language-only carriers can remain
/// style-inheritable while visually formatted carriers resolve omission to false.
pub(super) fn first_chart_text_character_props<'a, 'input>(
    node: Node<'a, 'input>,
) -> Option<Node<'a, 'input>> {
    node.descendants().find(|candidate| {
        candidate.is_element() && matches!(candidate.tag_name().name(), "defRPr" | "rPr")
    })
}

#[derive(Debug, Clone, Default)]
pub(super) struct ChartTextPaint {
    pub(super) color: Option<String>,
    pub(super) authored: bool,
    pub(super) hidden: bool,
}

/// Resolve the first authored DrawingML text fill in the character-property
/// cascade. An unsupported or unresolved direct paint is still authoritative:
/// it blocks lower defaults without inventing a replacement colour.
pub(super) fn chart_text_paint<'a, 'input: 'a>(
    nodes: impl IntoIterator<Item = Option<Node<'a, 'input>>>,
    resolver: &dyn ColorResolver,
) -> ChartTextPaint {
    for props in nodes.into_iter().flatten() {
        if !shape_has_fill_choice(props) {
            continue;
        }
        return ChartTextPaint {
            color: resolver.resolve_shape_fill(props),
            authored: true,
            hidden: child(props, "noFill").is_some(),
        };
    }
    ChartTextPaint::default()
}

pub(super) fn chart_text_body_paint(
    body: Option<Node<'_, '_>>,
    resolver: &dyn ColorResolver,
) -> ChartTextPaint {
    let nodes = body.into_iter().flat_map(|body| {
        body.descendants()
            .filter(|node| node.is_element() && matches!(node.tag_name().name(), "rPr" | "defRPr"))
    });
    chart_text_paint(nodes.map(Some), resolver)
}

pub(super) fn chart_title_text_paint(
    owner: Option<Node<'_, '_>>,
    resolver: &dyn ColorResolver,
) -> ChartTextPaint {
    let Some(title) = owner.and_then(|owner| child(owner, "title")) else {
        return ChartTextPaint::default();
    };
    chart_text_paint(
        title_text_property_nodes(title).into_iter().map(Some),
        resolver,
    )
}

#[derive(Debug, Clone, Default)]
pub(super) struct ChartLabelBodyStyle {
    pub(super) authored: bool,
    pub(super) rotation: Option<i32>,
    pub(super) wrap: Option<String>,
    pub(super) anchor: Option<String>,
    pub(super) vertical_mode: Option<String>,
    pub(super) left_inset: Option<i64>,
    pub(super) top_inset: Option<i64>,
    pub(super) right_inset: Option<i64>,
    pub(super) bottom_inset: Option<i64>,
}

pub(super) fn chart_label_body_style(node: Option<Node<'_, '_>>) -> ChartLabelBodyStyle {
    let Some(body) = node else {
        return ChartLabelBodyStyle::default();
    };
    let signed_coordinate = |name: &str| body.attribute(name).and_then(coordinate32_to_emu);
    ChartLabelBodyStyle {
        authored: true,
        rotation: body
            .attribute("rot")
            .and_then(|value| value.parse::<i32>().ok()),
        wrap: body.attribute("wrap").map(ToOwned::to_owned),
        anchor: body.attribute("anchor").map(ToOwned::to_owned),
        vertical_mode: body.attribute("vert").map(ToOwned::to_owned),
        left_inset: signed_coordinate("lIns"),
        top_inset: signed_coordinate("tIns"),
        right_inset: signed_coordinate("rIns"),
        bottom_inset: signed_coordinate("bIns"),
    }
}

pub(super) fn merge_chart_label_body_styles(
    direct: ChartLabelBodyStyle,
    fallback: ChartLabelBodyStyle,
) -> ChartLabelBodyStyle {
    ChartLabelBodyStyle {
        authored: direct.authored || fallback.authored,
        rotation: direct.rotation.or(fallback.rotation),
        wrap: direct.wrap.or(fallback.wrap),
        anchor: direct.anchor.or(fallback.anchor),
        vertical_mode: direct.vertical_mode.or(fallback.vertical_mode),
        left_inset: direct.left_inset.or(fallback.left_inset),
        top_inset: direct.top_inset.or(fallback.top_inset),
        right_inset: direct.right_inset.or(fallback.right_inset),
        bottom_inset: direct.bottom_inset.or(fallback.bottom_inset),
    }
}

pub(super) fn chart_text_run_from_node(
    run: Node,
    paragraph_default: Option<Node>,
    text_body_default: Option<Node>,
    resolver: &dyn ColorResolver,
) -> Option<ChartTextRun> {
    let text = child(run, "t")?.text().unwrap_or_default().to_string();
    let run_props = child(run, "rPr");
    // ST_TextFontSize is 100..400000 hundredths of a point. Resolve size
    // property-by-property: an invalid direct value is ignored and the next
    // paragraph/text-body default remains eligible, just like the other text
    // properties' independent cascade.
    let font_size_hpt = [run_props, paragraph_default, text_body_default]
        .into_iter()
        .flatten()
        .find_map(|node| node.attribute("sz").and_then(parse_text_font_size_hpt));
    let text_paint = chart_text_paint([run_props, paragraph_default, text_body_default], resolver);
    let font_face = run_props
        .and_then(|node| child(node, "latin"))
        .and_then(|latin| latin.attribute("typeface"))
        .or_else(|| {
            paragraph_default
                .and_then(|node| child(node, "latin"))
                .and_then(|latin| latin.attribute("typeface"))
        })
        .or_else(|| {
            text_body_default
                .and_then(|node| child(node, "latin"))
                .and_then(|latin| latin.attribute("typeface"))
        })
        .map(str::to_string)
        .filter(|face| !face.is_empty());

    Some(ChartTextRun {
        text,
        font_size_hpt,
        bold: chart_text_bool_from_property_cascade(
            [run_props, paragraph_default, text_body_default],
            "b",
        ),
        italic: chart_text_bool_from_property_cascade(
            [run_props, paragraph_default, text_body_default],
            "i",
        ),
        color: text_paint.color,
        color_paint_authored: text_paint.authored.then_some(true),
        color_hidden: text_paint.hidden.then_some(true),
        font_face,
        language: [run_props, paragraph_default, text_body_default]
            .into_iter()
            .flatten()
            .find_map(|node| node.attribute("lang").map(ToOwned::to_owned)),
        baseline: [run_props, paragraph_default, text_body_default]
            .into_iter()
            .flatten()
            .find_map(|node| {
                node.attribute("baseline")
                    .and_then(parse_chart_text_percentage)
            }),
        paragraph_align: None,
    })
}

pub(super) const MAX_CHART_TITLE_RICH_SCALARS: usize = 16_384;
pub(super) const MAX_CHART_TITLE_RICH_PARAGRAPHS: usize = 64;

/// Preserve the DrawingML run cascade of a legacy `<c:title><c:tx><c:rich>`.
/// ECMA-376 §21.2.2.214 permits independently formatted runs; flattening them
/// loses subtitles and authored line breaks. The bounded wire representation
/// is shared by DOCX/XLSX/PPTX chart hosts.
pub(super) fn parse_chart_title_rich_runs(
    title: Node,
    resolver: &dyn ColorResolver,
) -> Option<Vec<ChartTextRun>> {
    let rich = child(title, "tx").and_then(|tx| child(tx, "rich"))?;
    let text_body_default = child(rich, "lstStyle").and_then(|style| {
        style
            .descendants()
            .find(|node| node.is_element() && matches!(node.tag_name().name(), "rPr" | "defRPr"))
    });
    let mut runs = Vec::new();
    let mut scalar_count = 0usize;
    for (paragraph_index, paragraph) in rich
        .children()
        .filter(|node| node.is_element() && node.tag_name().name() == "p")
        .take(MAX_CHART_TITLE_RICH_PARAGRAPHS)
        .enumerate()
    {
        if paragraph_index > 0 && scalar_count < MAX_CHART_TITLE_RICH_SCALARS {
            runs.push(ChartTextRun {
                text: "\n".to_string(),
                font_size_hpt: None,
                bold: None,
                italic: None,
                color: None,
                color_paint_authored: None,
                color_hidden: None,
                font_face: None,
                language: None,
                baseline: None,
                paragraph_align: None,
            });
            scalar_count += 1;
        }
        let paragraph_default = child(paragraph, "pPr").and_then(|props| child(props, "defRPr"));
        for run_node in paragraph
            .children()
            .filter(|node| node.is_element() && matches!(node.tag_name().name(), "r" | "fld"))
        {
            if scalar_count >= MAX_CHART_TITLE_RICH_SCALARS {
                break;
            }
            let Some(mut run) =
                chart_text_run_from_node(run_node, paragraph_default, text_body_default, resolver)
            else {
                continue;
            };
            let remaining = MAX_CHART_TITLE_RICH_SCALARS - scalar_count;
            run.text = run.text.chars().take(remaining).collect();
            scalar_count += run.text.chars().count();
            if !run.text.is_empty() {
                runs.push(run);
            }
        }
    }
    (!runs.is_empty()).then_some(runs)
}

/// Detect the source-shape boundary observed for Office's automatic light
/// title paint in classic styles 41–48. Only paragraph defaults on the rich
/// paragraphs that actually supply title text participate. A sibling
/// `<c:title><c:txPr>` and empty formatting-only paragraphs are different
/// sources and must not enable the compatibility rule.
///
/// The Office-produced probes established the single-textual-paragraph case.
/// Multi-paragraph titles remain on the normative numeric default until a
/// per-paragraph Office rule is established and representable in the model.
pub(super) fn title_rich_text_has_observed_paragraph_default_run(title: Node) -> bool {
    let Some(rich) = child(title, "tx").and_then(|tx| child(tx, "rich")) else {
        return false;
    };
    let mut textual_paragraphs = rich
        .children()
        .filter(|node| node.is_element() && node.tag_name().name() == "p")
        .filter(|paragraph| {
            paragraph
                .children()
                .filter(|node| node.is_element() && matches!(node.tag_name().name(), "r" | "fld"))
                .any(|run| {
                    child(run, "t")
                        .and_then(|text| text.text())
                        .is_some_and(|text| !text.is_empty())
                })
        });
    let Some(paragraph) = textual_paragraphs.next() else {
        return false;
    };
    if textual_paragraphs.next().is_some() {
        return false;
    }
    child(paragraph, "pPr")
        .and_then(|properties| child(properties, "defRPr"))
        .is_some()
}

/// Parse the Chart Drawing part referenced by `<c:userShapes r:id>`.
///
/// ECMA-376 `dml-chartDrawing.xsd` defines each `cdr:relSizeAnchor` as `from`
/// and `to` markers in the inclusive 0..1 chart-space coordinate system,
/// followed by a DrawingML object. This first shared implementation retains
/// text shapes (`cdr:sp/cdr:txBody`) losslessly enough for chart titles,
/// subtitles, notes and source footers. Other object choices remain available
/// for later model extensions instead of being guessed into canvas primitives.
pub(super) fn parse_chart_user_shapes(
    root: Node,
    resolver: &dyn ColorResolver,
) -> Vec<ChartTextBox> {
    root.children()
        .filter(|node| node.is_element() && node.tag_name().name() == "relSizeAnchor")
        .filter_map(|anchor| {
            let from = child(anchor, "from")?;
            let to = child(anchor, "to")?;
            let marker = |node: Node, axis: &str| {
                child(node, axis)
                    .and_then(|value| value.text())
                    .and_then(|value| value.parse::<f64>().ok())
                    .filter(|value| value.is_finite() && (0.0..=1.0).contains(value))
            };
            let x = marker(from, "x")?;
            let y = marker(from, "y")?;
            let x2 = marker(to, "x")?;
            let y2 = marker(to, "y")?;
            if x2 < x || y2 < y {
                return None;
            }

            let shape = child(anchor, "sp")?;
            let text_body = child(shape, "txBody")?;
            let body_pr = child(text_body, "bodyPr");
            let vertical_anchor = body_pr
                .and_then(|body| body.attribute("anchor"))
                .map(str::to_string);
            let wrap = body_pr
                .and_then(|body| body.attribute("wrap"))
                .map(str::to_string);
            let body_defaults = BodyPrDefaults::spec();
            let parsed_body = body_pr.map(|body| parse_body_pr(body, &body_defaults));
            let (l_ins, t_ins, r_ins, b_ins) = parsed_body
                .map(|body| (body.l_ins, body.t_ins, body.r_ins, body.b_ins))
                .unwrap_or((
                    body_defaults.l_ins,
                    body_defaults.t_ins,
                    body_defaults.r_ins,
                    body_defaults.b_ins,
                ));
            let paragraphs = text_body
                .children()
                .filter(|node| node.is_element() && node.tag_name().name() == "p")
                .map(|paragraph| {
                    let p_pr = child(paragraph, "pPr");
                    let paragraph_default = p_pr.and_then(|props| child(props, "defRPr"));
                    let align = p_pr
                        .and_then(|props| props.attribute("algn"))
                        .map(str::to_string);
                    let runs = paragraph
                        .children()
                        .filter(|node| {
                            node.is_element() && matches!(node.tag_name().name(), "r" | "fld")
                        })
                        .filter_map(|run| {
                            chart_text_run_from_node(run, paragraph_default, None, resolver)
                        })
                        .collect();
                    ChartTextParagraph { runs, align }
                })
                .collect::<Vec<_>>();
            if paragraphs.iter().all(|paragraph| paragraph.runs.is_empty()) {
                return None;
            }

            Some(ChartTextBox {
                x,
                y,
                w: x2 - x,
                h: y2 - y,
                paragraphs,
                vertical_anchor,
                wrap,
                l_ins,
                t_ins,
                r_ins,
                b_ins,
            })
        })
        .collect()
}

/// Parse chart user-shape text with the owning chart's optional color-map
/// override applied to DrawingML scheme colors.
pub fn parse_chart_user_shapes_for_chart(
    chart_root: Node,
    user_shapes_root: Node,
    resolver: &dyn ColorResolver,
) -> Vec<ChartTextBox> {
    if let Some(mapping) = ChartColorMapping::from_chart_space(chart_root) {
        let mapped = ChartMappedColorResolver {
            base: resolver,
            mapping,
        };
        parse_chart_user_shapes(user_shapes_root, &mapped)
    } else {
        parse_chart_user_shapes(user_shapes_root, resolver)
    }
}

/// `<c:legend>` presence + `<c:legendPos val>` (ECMA-376 §21.2.2.10).
///
/// `(show_legend, legend_pos)`. When the chart omits `<c:legend>` Office
/// hides the legend even if a default position would otherwise apply, so
/// `show_legend = false` is the authoritative "no legend" signal.
pub(super) fn extract_legend(root: Node) -> (bool, Option<String>) {
    // The legend can sit anywhere inside `<c:chart>` but in practice it's a
    // direct child of `<c:chart>`. Use descendants to be tolerant of either
    // structure — there's only one `<c:legend>` element per chart.
    let legend = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "legend");
    let show = legend.is_some();
    let pos = legend.and_then(|ln| {
        child(ln, "legendPos")
            .and_then(|p| p.attribute("val"))
            .map(|s| s.to_string())
    });
    (show, pos)
}

/// Preserve `<c:legend><c:overlay>` and source-ordered indexed
/// `<c:legendEntry>` overrides (ECMA-376 §21.2.2.94, §21.2.2.132).
///
/// Entry-local text properties are partial: omitted properties inherit the
/// legend-level `<c:txPr>` in the renderer. The entry count is bounded before
/// constructing the wire vector so an authored legend cannot amplify a small
/// chart part into an unbounded UI model.
pub(super) fn extract_legend_overrides(
    root: Node,
    resolver: &dyn ColorResolver,
) -> (Option<bool>, Option<Vec<ChartLegendEntryOverride>>) {
    let Some(legend) = root
        .descendants()
        .find(|node| node.is_element() && node.tag_name().name() == "legend")
    else {
        return (None, None);
    };
    let overlay = bool_child(legend, "overlay");
    let entries = legend
        .children()
        .filter(|node| node.is_element() && node.tag_name().name() == "legendEntry")
        .take(MAX_CHART_LEGEND_ENTRIES)
        .filter_map(|entry| {
            let idx = child(entry, "idx")?.attribute("val")?.parse::<u32>().ok()?;
            let txpr = child(entry, "txPr");
            let run_props = txpr.and_then(first_chart_text_character_props);
            let font_color = txpr.and_then(|body| {
                body.descendants().find_map(|node| {
                    (node.is_element() && node.tag_name().name() == "solidFill")
                        .then(|| resolver.resolve_solid_fill(node))
                        .flatten()
                })
            });
            Some(ChartLegendEntryOverride {
                idx,
                deleted: bool_child(entry, "delete"),
                font_face: txpr.and_then(first_latin_typeface),
                font_color,
                font_size_hpt: run_props
                    .and_then(|props| props.attribute("sz"))
                    .and_then(parse_text_font_size_hpt),
                font_bold: chart_text_bool_from_present_props(run_props, "b"),
                font_italic: chart_text_bool_from_present_props(run_props, "i"),
            })
        })
        .collect::<Vec<_>>();
    (overlay, (!entries.is_empty()).then_some(entries))
}

/// `<c:barChart><c:gapWidth val>` / `<c:overlap val>` (ECMA-376 §21.2.2.13,
/// §21.2.2.25). Returns `(gap%, overlap%)`. Defaults to (None, None) when
/// the file relies on Office's defaults (gap 150, overlap 0).
pub(super) fn extract_bar_gap_overlap(root: Node) -> (Option<i32>, Option<i32>) {
    let gap = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "gapWidth")
        .and_then(|n| n.attribute("val").and_then(|v| v.parse::<i32>().ok()))
        .filter(|value| (0..=500).contains(value));
    let ov = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "overlap")
        .and_then(|n| n.attribute("val").and_then(|v| v.parse::<i32>().ok()))
        .filter(|value| (-100..=100).contains(value));
    (gap, ov)
}

pub(super) fn is_chart_group_data_labels(node: Node) -> bool {
    node.is_element()
        && matches!(node.tag_name().name(), "dLbls" | "dataLabels")
        && node
            .parent()
            .map(|parent| !matches!(parent.tag_name().name(), "ser" | "series"))
            .unwrap_or(true)
}

/// First chart-group-level `<c:dLbls><c:dLblPos val>` in the chart.
/// Series-level positions are retained on `ChartSeriesDataLabels` and must not
/// leak into sibling series as a chart-wide fallback. ECMA-376 §21.2.2.49.
pub(super) fn extract_data_label_position(root: Node) -> Option<String> {
    root.descendants()
        .filter(|n| is_chart_group_data_labels(*n))
        .find_map(|dlbls| {
            child(dlbls, "dLblPos")
                .and_then(|n| n.attribute("val"))
                .map(|s| s.to_string())
        })
}

/// First non-`General` chart-group `<c:dLbls><c:numFmt formatCode>`. Series
/// format codes stay on `ChartSeriesDataLabels` and do not become a sibling
/// series fallback.
/// ECMA-376 §21.2.2.37.
pub(super) fn extract_data_label_format_code(root: Node) -> Option<String> {
    root.descendants()
        .filter(|n| is_chart_group_data_labels(*n))
        .find_map(|dlbls| {
            child(dlbls, "numFmt")
                .and_then(|n| n.attribute("formatCode"))
                .map(|s| s.to_string())
                .filter(|s| !s.is_empty() && s != "General")
        })
}
