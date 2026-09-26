use super::*;

/// Find a direct child of `parent` whose local name is `name`.
pub(super) fn child<'a, 'i>(parent: Node<'a, 'i>, name: &str) -> Option<Node<'a, 'i>> {
    parent
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == name)
}

/// Retain one custom of-pie split atomically. The limit bounds the owned wire
/// vector before allocation/serialization; an oversized list is unresolved,
/// never a retained prefix with different split semantics.
pub(super) fn parse_of_pie_custom_split_with_limit(
    chart: Node<'_, '_>,
    max_points: usize,
) -> Option<Vec<usize>> {
    let split = child(chart, "custSplit")?;
    let mut indices = Vec::new();
    let mut point_count = 0usize;
    for point in split
        .children()
        .filter(|node| node.is_element() && node.tag_name().name() == "secondPiePt")
    {
        point_count = point_count.checked_add(1)?;
        if point_count > max_points {
            return None;
        }
        if let Some(index) = point
            .attribute("val")
            .and_then(|value| value.parse::<usize>().ok())
        {
            indices.push(index);
        }
    }
    Some(indices)
}

/// Read a no-namespace attribute `local` off `node` as an owned `String`.
///
/// Mirrors the pptx/xlsx crate-local `attr` helper exactly (matches only
/// attributes with no namespace, the shape every chart attribute uses), so the
/// chart-structure parse moved into [`parse_chart_part`] stays byte-identical
/// to the per-crate bodies it replaces.
pub(super) fn attr(node: &Node, local: &str) -> Option<String> {
    node.attributes()
        .find(|a| a.name() == local && a.namespace().is_none())
        .map(|a| a.value().to_owned())
}

pub(super) fn parse_chart_axis_id_value(value: &str) -> Option<String> {
    let lexical = value.trim();
    lexical
        .parse::<u32>()
        .map(|id| id.to_string())
        .or_else(|_| {
            // CT_UnsignedInt is the normative type, but Office-compatible
            // producers also serialize the same opaque 32-bit cross-reference
            // IDs through a signed decimal view. Preserve that bounded lexical
            // identity rather than dropping the complete chart; it is never
            // interpreted as a coordinate or allocation size.
            lexical.parse::<i32>().map(|id| id.to_string())
        })
        .ok()
}

/// Theme-aware color resolution for chart text-color helpers.
///
/// pptx and xlsx store their theme palettes in different shapes
/// (`HashMap<String, String>` vs. `&[String]`) and apply DrawingML
/// transforms with different `tint` formulas (Word-literal vs. linear
/// sRGB lerp), so each crate keeps its own resolver. The shared chart
/// helpers take a `&dyn ColorResolver` instead of either concrete type so
/// fields like `<c:dLbls><c:txPr>...<a:solidFill>` can be extracted once.
pub trait ColorResolver {
    /// Resolve an `<a:solidFill>` node to a hex string (no leading `#`),
    /// or `None` when the contained color child can't be mapped to a
    /// concrete RGB value (for example a `<a:schemeClr val="phClr"/>`
    /// that the implementation chooses not to substitute).
    ///
    /// The node passed in is the `<a:solidFill>` element itself; the
    /// implementation reads its direct children for the actual color
    /// (`<a:srgbClr>` / `<a:schemeClr>` / `<a:sysClr>` / `<a:prstClr>`)
    /// and applies the surrounding lumMod/lumOff/tint/shade transforms.
    fn resolve_solid_fill(&self, node: Node) -> Option<String>;

    /// Resolve a theme scheme slot to its base color (no leading `#`). Chart
    /// parts can carry their own `<c:clrMapOvr>` (§21.2.2.30), whose
    /// `CT_ColorMapping` attributes remap logical names such as `accent1` to a
    /// scheme slot such as `accent2`. The shared chart parser performs that
    /// logical-name remapping; the host still owns the theme storage lookup.
    ///
    /// Resolvers that do not expose a theme palette may keep the default. The
    /// chart-local wrapper then falls back to their ordinary fill resolver.
    fn resolve_scheme_color(&self, _name: &str) -> Option<String> {
        None
    }

    /// Color-transform behavior used after resolving a chart-local mapped
    /// scheme color. Charts use the PowerPoint/Excel DrawingML behavior by
    /// default; a host with different semantics can override it.
    fn tint_mode(&self) -> crate::color::TintMode {
        crate::color::TintMode::PowerPointLinear
    }

    /// Resolve the first `<a:solidFill>` among `parent`'s **direct children** to
    /// a hex string (no leading `#`) using the full DrawingML color grammar,
    /// including `lumMod`/`lumOff`/`tint`/`shade` transforms.
    ///
    /// This is the resolver used for chart *shape* fills that sit one level
    /// below their container — series fills/lines (`<c:ser><c:spPr>` /
    /// `…<a:ln>`), marker fill/line (`<c:marker><c:spPr>` / `…<a:ln>`),
    /// per-point fills (`<c:dPt><c:spPr>`) and error-bar strokes
    /// (`<c:errBars><c:spPr>` / `…<a:ln>`). It is intentionally distinct from
    /// [`ColorResolver::resolve_solid_fill`] so resolvers can route every shape
    /// through the complete DrawingML color-transform grammar.
    ///
    /// The default implementation finds the direct-child `<a:solidFill>` and
    /// delegates to [`ColorResolver::resolve_solid_fill`], which is correct for
    /// resolvers whose `resolve_solid_fill` already applies the full grammar
    /// (pptx). xlsx overrides it to route through its DrawingML color path.
    fn resolve_shape_fill(&self, parent: Node) -> Option<String> {
        parent
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "solidFill")
            .and_then(|fill| self.resolve_solid_fill(fill))
    }

    /// Theme major (heading) Latin typeface name, or `None` when the theme
    /// declares no `fontScheme`. Used as the chart-text fallback face when a run
    /// carries no explicit `<a:latin>`. Defaults to `None` so resolvers that do
    /// not carry a theme font map need not override it.
    fn theme_major_font_latin(&self) -> Option<String> {
        None
    }

    /// Theme minor (body) Latin typeface name, or `None` when the theme declares
    /// no `fontScheme`. Companion to [`ColorResolver::theme_major_font_latin`].
    fn theme_minor_font_latin(&self) -> Option<String> {
        None
    }

    /// Default series fill for a series with no explicit `<c:spPr>` fill, keyed
    /// by its `<c:idx>` (ECMA-376 §21.2.2.84). Office cycles the theme accents:
    /// `theme.accent[(idx % 6) + 1]`. Returning the resolved accent hex here
    /// (no leading `#`) lets the renderer draw the correct default palette
    /// without needing theme access.
    ///
    /// Defaults to `None` for callers that do not carry a theme palette.
    fn resolve_series_accent(&self, _idx: usize) -> Option<String> {
        None
    }

    /// DrawingML `a:fmtScheme` used by ChartEx chart-style `fillRef`/`lnRef`.
    /// The shared parser owns reference inheritance; hosts only expose the
    /// already-parsed theme sidecar.
    fn theme_format_scheme(&self) -> Option<&crate::theme::ThemeFormatScheme> {
        None
    }

    /// Chart-area background to use when the `<c:chartSpace>` carries **no**
    /// `<c:spPr>` at all. Excel relies on its default opaque-white chart area in
    /// that case, so the xlsx resolver returns `Some("FFFFFF")`; PowerPoint
    /// composites the chart transparently over the slide, so pptx returns `None`.
    /// (When `<c:spPr>` *is* present the parser honours whatever it resolves to —
    /// a solid hex or `noFill` → `None` — regardless of this default.)
    fn default_chart_bg(&self) -> Option<String> {
        None
    }

    /// Plot-area background to use when `<c:plotArea><c:spPr>` omits a fill
    /// choice. Excel supplies an automatic opaque-white plot-area fill while
    /// PowerPoint/Word hosts keep their existing transparent behavior.
    fn default_plot_area_bg(&self) -> Option<String> {
        None
    }

    /// Whether this host applies Excel's observed implicit outline-only style
    /// to a narrowly-defined, otherwise-unformatted negative column series.
    /// This is not an OOXML default: hosts opt in only after their own Office
    /// output establishes the application behavior.
    fn implicit_outline_only_negative_column_style(&self) -> bool {
        false
    }

    /// Host-scoped compatibility boundary for automatic light text in dark
    /// classic chart styles. Cross-host evidence currently covers style 41;
    /// Word additionally proves the complete 41..=48 range. A host must opt in
    /// to the broader range instead of inheriting it from the shared parser.
    fn office_dark_text_contrast_applies(&self, _style: u8) -> bool {
        false
    }

    /// Host-scoped title-carrier boundary for dark classic styles. Word has a
    /// measured rich-title matrix; other hosts must opt in separately instead
    /// of inheriting that source-shape rule from ordinary chart text.
    fn office_dark_title_contrast_applies(&self, _style: u8) -> bool {
        false
    }

    /// Application-defined tint/shade for a repeated six-accent Pattern 2 set
    /// (ECMA-376 §21.2.3.46 Table 6). The first set is normative and unmodified;
    /// later sets remain unresolved unless the host has measured evidence.
    fn classic_pattern2_set_transform(&self, set_index: usize) -> Option<f64> {
        (set_index == 0).then_some(0.0)
    }
}

/// OPC relationship scope for a chart picture fill. A chart part and its
/// linked Chart Style part each own an independent `.rels` file, so an `rId`
/// is meaningful only together with this source part.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartImageSource {
    Chart,
    Style,
    Theme,
}

/// Host-owned relationship resolver for DrawingML images embedded in charts.
/// The shared parser retains only a normalized package path and MIME type; it
/// never reads or copies image bytes into the chart wire model.
pub trait ChartImageResolver {
    fn resolve_image(
        &self,
        source: ChartImageSource,
        relationship_id: &str,
    ) -> Option<(String, String)>;
}

pub(super) struct EmptyChartImageResolver;

impl ChartImageResolver for EmptyChartImageResolver {
    fn resolve_image(
        &self,
        _source: ChartImageSource,
        _relationship_id: &str,
    ) -> Option<(String, String)> {
        None
    }
}

/// Bounded, path-only relationship index shared by the three package parsers.
/// It deliberately stores no media bytes; the renderer fetches the selected
/// paths through the document's existing lazy media callback.
#[derive(Debug, Clone, Default)]
pub struct ChartImageRelationships {
    chart: BTreeMap<String, (String, String)>,
    style: BTreeMap<String, (String, String)>,
    theme: BTreeMap<String, (String, String)>,
}

impl ChartImageRelationships {
    pub fn insert_parsed_relationships(
        &mut self,
        source: ChartImageSource,
        source_part_path: &str,
        relationships: &BTreeMap<String, crate::rels::RelTarget>,
    ) {
        let target = match source {
            ChartImageSource::Chart => &mut self.chart,
            ChartImageSource::Style => &mut self.style,
            ChartImageSource::Theme => &mut self.theme,
        };
        for (relationship_id, relationship) in relationships {
            if relationship.mode != crate::rels::TargetMode::Internal
                || !relationship
                    .relationship_type
                    .as_deref()
                    .is_some_and(|kind| {
                        kind.strip_suffix("/image").is_some_and(|base| {
                            base == crate::ns::relationships::TRANSITIONAL
                                || base == crate::ns::relationships::STRICT
                        })
                    })
            {
                continue;
            }
            // ECMA-376 Part 2 §6.5.2.3: resolve against the source part; a
            // target naming no package part is treated as an absent image.
            let Some(path) = relationship.resolve_part(source_part_path) else {
                continue;
            };
            let mime = crate::blip::mime_from_ext(&path).to_owned();
            target.insert(relationship_id.clone(), (path, mime));
        }
    }

    pub fn insert_part_relationships(
        &mut self,
        source: ChartImageSource,
        source_part_path: &str,
        relationships_xml: &str,
    ) {
        let relationships = crate::rels::parse_rels(relationships_xml);
        self.insert_parsed_relationships(source, source_part_path, &relationships);
    }
}

impl ChartImageResolver for ChartImageRelationships {
    fn resolve_image(
        &self,
        source: ChartImageSource,
        relationship_id: &str,
    ) -> Option<(String, String)> {
        let relationships = match source {
            ChartImageSource::Chart => &self.chart,
            ChartImageSource::Style => &self.style,
            ChartImageSource::Theme => &self.theme,
        };
        relationships.get(relationship_id).cloned()
    }
}

/// Borrowed resolver composition used by package hosts: chart/style
/// relationships are local to one chart, while theme relationships are shared
/// by every chart in the document. Keeping them as two borrowed indexes avoids
/// cloning a potentially large theme relationship table per chart.
pub struct ChartImageResolverChain<'a> {
    primary: &'a dyn ChartImageResolver,
    fallback: &'a dyn ChartImageResolver,
}

impl<'a> ChartImageResolverChain<'a> {
    pub fn new(primary: &'a dyn ChartImageResolver, fallback: &'a dyn ChartImageResolver) -> Self {
        Self { primary, fallback }
    }
}

impl ChartImageResolver for ChartImageResolverChain<'_> {
    fn resolve_image(
        &self,
        source: ChartImageSource,
        relationship_id: &str,
    ) -> Option<(String, String)> {
        self.primary
            .resolve_image(source, relationship_id)
            .or_else(|| self.fallback.resolve_image(source, relationship_id))
    }
}

/// The direct `CT_ColorMapping` carried by `<c:chartSpace><c:clrMapOvr>`
/// (ECMA-376 §21.2.2.30; `dml-chart.xsd::CT_ChartSpace`). Unlike PresentationML
/// `<p:clrMapOvr>`, this element is not a `CT_ColorMappingOverride` choice: its
/// twelve logical-to-scheme attributes live directly on the chart element.
#[derive(Debug)]
pub(super) struct ChartColorMapping {
    entries: Vec<(String, String)>,
}

impl ChartColorMapping {
    pub(super) fn from_chart_space(chart_root: Node) -> Option<Self> {
        let node = child(chart_root, "clrMapOvr")?;
        let entries = crate::color::SCHEME_DEFAULT_SLOTS
            .iter()
            .filter_map(|(logical, _)| {
                node.attribute(*logical)
                    .map(|slot| ((*logical).to_owned(), slot.to_owned()))
            })
            .collect();
        Some(Self { entries })
    }

    fn map<'a>(&'a self, logical: &'a str) -> &'a str {
        self.entries
            .iter()
            .find(|(name, _)| name == logical)
            .map(|(_, slot)| slot.as_str())
            .unwrap_or(logical)
    }
}

/// Chart-scoped resolver that applies `c:clrMapOvr` before delegating theme
/// slot lookup to the pptx/xlsx host resolver. Keeping this wrapper here makes
/// the mapping behavior identical for both package formats.
pub(super) struct ChartMappedColorResolver<'a> {
    pub(super) base: &'a dyn ColorResolver,
    pub(super) mapping: ChartColorMapping,
}

pub(super) struct ColorResolverThemeAdapter<'a>(pub(super) &'a dyn ColorResolver);

impl crate::color::ThemeResolver for ColorResolverThemeAdapter<'_> {
    fn resolve_scheme_color(&self, name: &str) -> Option<String> {
        self.0.resolve_scheme_color(name)
    }
}

impl crate::color::ThemeResolver for ChartMappedColorResolver<'_> {
    fn resolve_scheme_color(&self, logical: &str) -> Option<String> {
        let mapped = self.mapping.map(logical);
        self.base.resolve_scheme_color(mapped)
    }
}

impl ColorResolver for ChartMappedColorResolver<'_> {
    fn resolve_solid_fill(&self, node: Node) -> Option<String> {
        crate::color::parse_color_node(node, self, self.base.tint_mode())
            .or_else(|| self.base.resolve_solid_fill(node))
    }

    fn resolve_scheme_color(&self, name: &str) -> Option<String> {
        crate::color::ThemeResolver::resolve_scheme_color(self, name)
    }

    fn tint_mode(&self) -> crate::color::TintMode {
        self.base.tint_mode()
    }

    fn resolve_shape_fill(&self, parent: Node) -> Option<String> {
        parent
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "solidFill")
            .and_then(|fill| self.resolve_solid_fill(fill))
            .or_else(|| self.base.resolve_shape_fill(parent))
    }

    fn theme_major_font_latin(&self) -> Option<String> {
        self.base.theme_major_font_latin()
    }

    fn theme_minor_font_latin(&self) -> Option<String> {
        self.base.theme_minor_font_latin()
    }

    fn resolve_series_accent(&self, idx: usize) -> Option<String> {
        let logical = format!("accent{}", idx % 6 + 1);
        let mapped = self.mapping.map(&logical);
        self.base
            .resolve_scheme_color(mapped)
            .or_else(|| self.base.resolve_series_accent(idx))
    }

    fn theme_format_scheme(&self) -> Option<&crate::theme::ThemeFormatScheme> {
        self.base.theme_format_scheme()
    }

    fn default_chart_bg(&self) -> Option<String> {
        self.base.default_chart_bg()
    }

    fn default_plot_area_bg(&self) -> Option<String> {
        self.base.default_plot_area_bg()
    }

    fn implicit_outline_only_negative_column_style(&self) -> bool {
        self.base.implicit_outline_only_negative_column_style()
    }

    fn office_dark_text_contrast_applies(&self, style: u8) -> bool {
        self.base.office_dark_text_contrast_applies(style)
    }

    fn office_dark_title_contrast_applies(&self, style: u8) -> bool {
        self.base.office_dark_title_contrast_applies(style)
    }

    fn classic_pattern2_set_transform(&self, set_index: usize) -> Option<f64> {
        self.base.classic_pattern2_set_transform(set_index)
    }
}

// ============================================================================
// Series-detail extractors (markers, per-point overrides, data labels, error
// bars) — moved verbatim from the xlsx crate so `parse_chart_part` populates
// the rich per-series fields for both pptx and xlsx.
// ============================================================================

/// Parse `<c:marker>` into `(symbol, size, fill, fill_paint,
/// fill_paint_authored, line, line_width_emu)`. ECMA-376 §21.2.2.32 /
/// §21.2.2.34 use the full DrawingML
/// shape-property fill grammar for marker paint, so the shared structured fill
/// is retained in addition to the legacy resolved solid color. `size` is the
/// point value parsed as an integer (matching Excel's `<c:size val>`
/// unsignedByte) then widened to `f64` for the shared model.
pub(super) type ParsedMarkerBlock = (
    Option<String>,
    Option<f64>,
    Option<String>,
    Option<ChartStyleFill>,
    Option<bool>,
    Option<String>,
    Option<u32>,
    Option<bool>,
);

#[cfg(test)]
pub(super) fn parse_marker_block(
    marker_node: Option<Node>,
    resolver: &dyn ColorResolver,
) -> ParsedMarkerBlock {
    parse_marker_block_with_images(marker_node, resolver, &EmptyChartImageResolver)
}

#[cfg(test)]
pub(super) fn parse_marker_block_with_images(
    marker_node: Option<Node>,
    resolver: &dyn ColorResolver,
    image_resolver: &dyn ChartImageResolver,
) -> ParsedMarkerBlock {
    let mut paint_budget = MAX_CHART_MARKER_PAINT_COMPONENTS;
    let mut paint_budget_exceeded = false;
    parse_marker_block_with_budget(
        marker_node,
        resolver,
        image_resolver,
        &mut paint_budget,
        &mut paint_budget_exceeded,
    )
}

#[cfg(test)]
pub(super) fn parse_marker_block_with_budget(
    marker_node: Option<Node>,
    resolver: &dyn ColorResolver,
    image_resolver: &dyn ChartImageResolver,
    paint_budget: &mut usize,
    paint_budget_exceeded: &mut bool,
) -> ParsedMarkerBlock {
    let Some(mk) = marker_node else {
        return (None, None, None, None, None, None, None, None);
    };
    let symbol = child(mk, "symbol")
        .and_then(|n| n.attribute("val"))
        .map(|s| s.to_string());
    let size = child(mk, "size")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<u32>().ok())
        .map(|v| v as f64);
    let sp_pr = child(mk, "spPr");
    // Count stops before resolving colors or collecting/sorting the gradient.
    // An over-budget direct paint remains authored for precedence purposes but
    // is not expanded into the wire model.
    let component_count = sp_pr
        .and_then(chart_style_paint_component_count)
        .unwrap_or(0);
    let within_recipe_limit = component_count <= MAX_CHART_MARKER_GRADIENT_STOPS;
    let within_chart_budget = component_count <= *paint_budget;
    let direct_fill = if within_recipe_limit && within_chart_budget {
        *paint_budget -= component_count;
        extract_direct_shape_fill(sp_pr, resolver)
    } else {
        *paint_budget_exceeded = true;
        DirectShapeFill {
            paint_authored: sp_pr
                .and_then(|shape| {
                    shape.children().find(|node| {
                        node.is_element()
                            && matches!(
                                node.tag_name().name(),
                                "noFill"
                                    | "solidFill"
                                    | "gradFill"
                                    | "pattFill"
                                    | "blipFill"
                                    | "grpFill"
                            )
                    })
                })
                .map(|_| true),
            ..Default::default()
        }
    };
    let fill = if direct_fill.hidden == Some(true) {
        // The renderer accepts 8-digit RRGGBBAA, so preserve an explicit
        // marker noFill as a transparent paint instead of collapsing it into
        // the same None used for an unspecified, inherited fill.
        Some("00000000".to_string())
    } else {
        direct_fill.color
    };
    // A resolved solid already has the compact legacy `markerFill` wire field.
    // Retain the structured union only when it carries geometry/pattern data;
    // this keeps existing solid-marker output byte-stable.
    let fill_paint = if !(within_recipe_limit && within_chart_budget) {
        None
    } else if let Some(blip_fill) = sp_pr.and_then(|shape| child(shape, "blipFill")) {
        parse_chart_image_fill(blip_fill, resolver, image_resolver, ChartImageSource::Chart)
    } else {
        direct_fill
            .fill
            .filter(|fill| !matches!(fill, ChartStyleFill::Solid { .. }))
    };
    let line = sp_pr.and_then(|p| child(p, "ln")).and_then(|ln| {
        if child(ln, "noFill").is_some() {
            // Keep direct marker-outline noFill distinct from an omitted line
            // so linked Chart Style fallback cannot repaint it later.
            Some("00000000".to_string())
        } else {
            resolver.resolve_shape_fill(ln)
        }
    });
    let line_width_emu = sp_pr
        .and_then(|p| child(p, "ln"))
        .and_then(|ln| ln.attribute("w"))
        .and_then(|value| value.parse::<u32>().ok());
    let line_paint_authored = sp_pr
        .and_then(|shape| child(shape, "ln"))
        .and_then(|line| {
            line.children().find(|node| {
                node.is_element()
                    && matches!(
                        node.tag_name().name(),
                        "noFill" | "solidFill" | "gradFill" | "pattFill"
                    )
            })
        })
        .map(|_| true);
    (
        symbol,
        size,
        fill,
        fill_paint,
        direct_fill.paint_authored,
        line,
        line_width_emu,
        line_paint_authored,
    )
}

/// Parse the marker shape once for model retention. The public compatibility
/// tuple above predates `ChartExElementStyle`; parser-backed charts derive its
/// legacy scalar/structured projections from the same bounded parse retained
/// on `marker_style`. Keeping the bounded structured fill in both wire fields
/// is intentional backward compatibility: existing consumers read
/// `markerFillPaint`, while newer consumers need `markerStyle` for geometry and
/// effects. The recipe is parsed and charged to the aggregate budget only once.
pub(super) fn parse_marker_model_style_with_budget(
    marker_node: Option<Node>,
    resolver: &dyn ColorResolver,
    image_resolver: &dyn ChartImageResolver,
    paint_budget: &mut usize,
    paint_budget_exceeded: &mut bool,
) -> (ParsedMarkerBlock, Option<ChartExElementStyle>) {
    let Some(marker) = marker_node else {
        return ((None, None, None, None, None, None, None, None), None);
    };
    let symbol = child(marker, "symbol")
        .and_then(|node| node.attribute("val"))
        .map(str::to_owned);
    let size = child(marker, "size")
        .and_then(|node| node.attribute("val"))
        .and_then(|value| value.parse::<u32>().ok())
        .map(|value| value as f64);
    let sp_pr = child(marker, "spPr");
    let component_count = sp_pr
        .and_then(chart_style_paint_component_count)
        .unwrap_or(0);
    if component_count > MAX_CHART_MARKER_GRADIENT_STOPS || component_count > *paint_budget {
        *paint_budget_exceeded = true;
        let fill_authored = sp_pr.is_some_and(shape_has_fill_choice).then_some(true);
        let line_node = sp_pr.and_then(|shape| child(shape, "ln"));
        let line_authored = line_node.is_some_and(shape_has_fill_choice).then_some(true);
        return (
            (
                symbol,
                size,
                None,
                None,
                fill_authored,
                None,
                line_node
                    .and_then(|line| line.attribute("w"))
                    .and_then(|value| value.parse::<u32>().ok()),
                line_authored,
            ),
            None,
        );
    }
    *paint_budget -= component_count;
    let style = parse_chartex_element_style(
        marker,
        resolver,
        None,
        None,
        image_resolver,
        ChartImageSource::Chart,
    );
    let fill = if style.fill_hidden == Some(true) {
        Some("00000000".to_owned())
    } else {
        style
            .fill_colors
            .as_ref()
            .and_then(|colors| colors.first())
            .and_then(Clone::clone)
    };
    let line = if style.line_hidden == Some(true) {
        Some("00000000".to_owned())
    } else {
        style
            .line_colors
            .as_ref()
            .and_then(|colors| colors.first())
            .and_then(Clone::clone)
    };
    let fill_paint = style
        .fill_paints
        .as_ref()
        .and_then(|paints| paints.first())
        .and_then(Clone::clone)
        // A resolved solid already has the compact legacy `markerFill` field.
        // Retain only structured paints in the compatibility union.
        .filter(|paint| !matches!(paint, ChartStyleFill::Solid { .. }));
    let legacy = (
        symbol,
        size,
        fill,
        fill_paint,
        style.fill_paint_authored,
        line,
        style.line_width_emu,
        style.line_paint_authored,
    );
    (legacy, Some(style))
}

pub(super) fn parse_series_pattern_fill(
    ser_node: Node,
    resolver: &dyn ColorResolver,
) -> Option<ChartPatternFill> {
    let patt_fill = child(ser_node, "spPr").and_then(|shape| child(shape, "pattFill"))?;
    let adapter = ColorResolverThemeAdapter(resolver);
    let pattern = crate::fill::parse_patt_fill(patt_fill, &adapter, resolver.tint_mode());
    Some(ChartPatternFill {
        fill_type: "pattern".to_string(),
        fg: pattern.fg,
        bg: pattern.bg,
        preset: pattern.preset,
    })
}

/// Walk every `<c:dPt>` direct child of the series and collect per-point
/// overrides. Multiple `<c:dPt>` per series is normal; each targets one
/// `<c:idx>` (ECMA-376 §21.2.2.39). Fill from `<c:spPr>`, marker from a nested
/// `<c:marker>`, and `<c:explosion>` (pie/doughnut pull-out) are captured.
#[cfg(test)]
pub(super) fn parse_data_point_overrides(
    ser_node: Node,
    resolver: &dyn ColorResolver,
) -> Vec<ChartDataPointOverride> {
    parse_data_point_overrides_with_images(ser_node, resolver, &EmptyChartImageResolver)
}

#[cfg(test)]
pub(super) fn parse_data_point_overrides_with_images(
    ser_node: Node,
    resolver: &dyn ColorResolver,
    image_resolver: &dyn ChartImageResolver,
) -> Vec<ChartDataPointOverride> {
    let mut paint_budget = MAX_CHART_MARKER_PAINT_COMPONENTS;
    let mut paint_budget_exceeded = false;
    parse_data_point_overrides_with_budget(
        ser_node,
        resolver,
        image_resolver,
        true,
        &mut paint_budget,
        &mut paint_budget_exceeded,
    )
}

pub(super) fn parse_data_point_overrides_with_budget(
    ser_node: Node,
    resolver: &dyn ColorResolver,
    image_resolver: &dyn ChartImageResolver,
    include_shape_style: bool,
    paint_budget: &mut usize,
    paint_budget_exceeded: &mut bool,
) -> Vec<ChartDataPointOverride> {
    let mut result = Vec::new();
    for dpt in ser_node
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "dPt")
    {
        let idx = child(dpt, "idx")
            .and_then(|n| n.attribute("val"))
            .and_then(|v| v.parse::<u32>().ok())
            .unwrap_or(0);
        let (color, fill_hidden, line_color, line_width_emu, line_dash, line_hidden) =
            parse_data_point_shape(dpt, resolver);
        let shape = child(dpt, "spPr");
        let shape_fill_component_count = if include_shape_style {
            shape
                .and_then(chart_style_paint_component_count)
                .unwrap_or(0)
        } else {
            0
        };
        let shape_line_component_count = if include_shape_style {
            shape
                .and_then(|sp_pr| child(sp_pr, "ln"))
                .and_then(chart_style_paint_component_count)
                .unwrap_or(0)
        } else {
            0
        };
        let shape_component_count =
            shape_fill_component_count.saturating_add(shape_line_component_count);
        let shape_style = if include_shape_style
            && shape_fill_component_count <= MAX_CHART_MARKER_GRADIENT_STOPS
            && shape_line_component_count <= MAX_CHART_MARKER_GRADIENT_STOPS
            && shape_component_count <= *paint_budget
        {
            *paint_budget -= shape_component_count;
            shape.map(|_| {
                parse_chartex_element_style(
                    dpt,
                    resolver,
                    None,
                    None,
                    image_resolver,
                    ChartImageSource::Chart,
                )
            })
        } else if !include_shape_style {
            // Compatibility-only callers can still request the effect atom
            // without expanding a structured fill/line recipe.
            parse_direct_chart_effect_style(dpt, resolver)
        } else {
            if include_shape_style && shape_component_count > 0 {
                *paint_budget_exceeded = true;
            }
            None
        };
        let mk = child(dpt, "marker");
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
            mk,
            resolver,
            image_resolver,
            paint_budget,
            paint_budget_exceeded,
        );
        let explosion = extract_dpt_explosion(dpt);
        let bubble_3d = strict_boolean_child(dpt, "bubble3D");
        result.push(ChartDataPointOverride {
            idx,
            color,
            fill_hidden,
            chartex_style: shape_style,
            line_color,
            line_width_emu,
            line_dash,
            line_hidden,
            marker_symbol,
            marker_size,
            marker_fill,
            marker_fill_paint,
            marker_fill_paint_authored,
            marker_style,
            marker_line,
            marker_line_width_emu,
            marker_line_paint_authored,
            bubble_3d,
            explosion,
        });
    }
    result
}

/// Resolve `<c:ser><c:extLst><c:ext><c15:datalabelsRange>` cache: index → label
/// text. Used to substitute `<a:fld type="CELLRANGE">` placeholders. Missing
/// entries stay absent from the map.
pub(super) fn collect_dlbl_range_cache(ser_node: Node) -> std::collections::HashMap<u32, String> {
    let mut map: std::collections::HashMap<u32, String> = std::collections::HashMap::new();
    let Some(ext_lst) = child(ser_node, "extLst") else {
        return map;
    };
    for ext in ext_lst
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "ext")
    {
        for range in ext
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "datalabelsRange")
        {
            for cache in range
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "dlblRangeCache")
            {
                for pt in cache
                    .children()
                    .filter(|n| n.is_element() && n.tag_name().name() == "pt")
                {
                    let Some(idx) = pt.attribute("idx").and_then(|v| v.parse::<u32>().ok()) else {
                        continue;
                    };
                    let v = child(pt, "v")
                        .and_then(|n| n.text())
                        .unwrap_or("")
                        .to_string();
                    map.insert(idx, v);
                }
            }
        }
    }
    map
}

/// Walk a `<c:tx><c:rich>` (or any DrawingML rich-text root) and reduce it to
/// plain text. `<a:fld type="CELLRANGE">` placeholders are substituted from
/// `cellrange_cache`. Other field types and runs are concatenated; newlines
/// come from paragraph breaks.
pub(super) fn flatten_rich_text(rich_root: Node, cellrange_cache: Option<&str>) -> String {
    let mut out = String::new();
    let mut first_para = true;
    for p in rich_root
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "p")
    {
        if !first_para {
            out.push('\n');
        }
        first_para = false;
        for c in p.children().filter(|n| n.is_element()) {
            match c.tag_name().name() {
                "r" => {
                    if let Some(t) = c.children().find(|n| n.tag_name().name() == "t") {
                        if let Some(s) = t.text() {
                            out.push_str(s);
                        }
                    }
                }
                "fld" => {
                    let typ = c.attribute("type").unwrap_or("");
                    if typ == "CELLRANGE" {
                        if let Some(s) = cellrange_cache {
                            out.push_str(s);
                        }
                    } else if let Some(t) = c.children().find(|n| n.tag_name().name() == "t") {
                        if let Some(s) = t.text() {
                            out.push_str(s);
                        }
                    }
                }
                _ => {}
            }
        }
    }
    out
}

pub(super) const MAX_DATA_LABEL_RICH_SCALARS: usize = 4096;
pub(super) const MAX_DATA_LABEL_RICH_LINES: usize = 4;

/// Preserve the effective DrawingML character properties of a custom chart
/// data label without allowing an untrusted rich-text body to grow the public
/// model without bound. Runs in the same paragraph stay inline; paragraph
/// boundaries become a newline run so the canvas renderer can keep authored
/// line breaks while measuring the complete label as one object.
pub(super) fn parse_data_label_rich_runs(
    label: Node,
    resolver: &dyn ColorResolver,
    cellrange_cache: Option<&str>,
) -> Option<Vec<ChartTextRun>> {
    let rich = child(label, "tx").and_then(|tx| child(tx, "rich"))?;
    let txpr_default = child(label, "txPr").and_then(first_paragraph_default_run_props);
    parse_data_label_rich_body(rich, txpr_default, resolver, cellrange_cache)
}

/// Parse the bounded DrawingML paragraphs carried directly by ChartEx
/// `<cx:dataLabel><cx:txPr>` as well as classic `<c:tx><c:rich>`. The caller
/// supplies the effective txPr default because classic labels may keep it in a
/// sibling element while ChartEx keeps it in this same rich body.
pub(super) fn parse_data_label_rich_body(
    rich: Node,
    txpr_default: Option<Node>,
    resolver: &dyn ColorResolver,
    cellrange_cache: Option<&str>,
) -> Option<Vec<ChartTextRun>> {
    let mut runs = Vec::new();
    let mut scalar_count = 0usize;

    for (paragraph_index, paragraph) in rich
        .children()
        .filter(|node| node.is_element() && node.tag_name().name() == "p")
        .take(MAX_DATA_LABEL_RICH_LINES)
        .enumerate()
    {
        let paragraph_align = child(paragraph, "pPr").and_then(|props| attr(&props, "algn"));
        if paragraph_index > 0 {
            if scalar_count >= MAX_DATA_LABEL_RICH_SCALARS {
                break;
            }
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
        for run_node in paragraph.children().filter(|node| {
            node.is_element() && matches!(node.tag_name().name(), "r" | "fld" | "br")
        }) {
            if scalar_count >= MAX_DATA_LABEL_RICH_SCALARS {
                break;
            }
            if run_node.tag_name().name() == "br" {
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
                    paragraph_align: paragraph_align.clone(),
                });
                scalar_count += 1;
                continue;
            }
            let Some(mut run) =
                chart_text_run_from_node(run_node, paragraph_default, txpr_default, resolver)
            else {
                continue;
            };
            if run_node.tag_name().name() == "fld"
                && run_node.attribute("type") == Some("CELLRANGE")
            {
                run.text = cellrange_cache.unwrap_or_default().to_string();
            }
            let remaining = MAX_DATA_LABEL_RICH_SCALARS - scalar_count;
            let bounded = run.text.chars().take(remaining).collect::<String>();
            scalar_count += bounded.chars().count();
            if !bounded.is_empty() {
                run.text = bounded;
                run.paragraph_align = paragraph_align.clone();
                runs.push(run);
            }
        }
    }
    (!runs.is_empty()).then_some(runs)
}

/// Parse a data-label `<c:spPr>` (§21.2.2.197) into a callout [`ChartLabelBox`]
/// (fill + border). Returns `None` when the shape node is absent OR carries
/// neither a resolvable fill nor a border — i.e. nothing that would draw a box.
/// The direct-child `<a:solidFill>` is the box fill; the `<a:ln>` solidFill and
/// its `w` attribute are the border. Colors resolve through
/// [`ColorResolver::resolve_shape_fill`] so a `<a:sysClr>`/`<a:schemeClr>` picks
/// up its transforms (Office writes the default white box as
/// `<a:sysClr val="window">`).
pub(super) fn label_box_paint_component_counts(sp_pr: Node) -> [usize; 2] {
    [
        chart_style_paint_component_count(sp_pr).unwrap_or(0),
        child(sp_pr, "ln")
            .and_then(chart_style_paint_component_count)
            .unwrap_or(0),
    ]
}

pub(super) fn label_shape_nodes<'a, 'input: 'a>(
    d_lbls: Node<'a, 'input>,
) -> impl Iterator<Item = Node<'a, 'input>> + 'a {
    d_lbls
        .children()
        .filter(|node| {
            node.is_element() && matches!(node.tag_name().name(), "spPr" | "dLbl" | "dataLabel")
        })
        .flat_map(|node| {
            if node.tag_name().name() == "spPr" {
                Some(node)
            } else {
                child(node, "spPr")
            }
        })
}

pub(super) fn label_paint_recipes_within_budget<'a, 'input: 'a>(
    shapes: impl Iterator<Item = Node<'a, 'input>>,
) -> bool {
    label_paint_recipes_within_limit(shapes, MAX_CHART_LABEL_PAINT_COMPONENTS)
}

pub(super) fn label_paint_recipes_within_limit<'a, 'input: 'a>(
    shapes: impl Iterator<Item = Node<'a, 'input>>,
    aggregate_limit: usize,
) -> bool {
    let mut total = 0usize;
    for shape in shapes {
        for components in label_box_paint_component_counts(shape) {
            if components > MAX_CHART_LABEL_GRADIENT_STOPS {
                return false;
            }
            let Some(next) = total.checked_add(components) else {
                return false;
            };
            if next > aggregate_limit {
                return false;
            }
            total = next;
        }
    }
    true
}

pub(super) fn chartex_label_paint_recipes_within_limit<'a, 'input: 'a>(
    series: impl Iterator<Item = Node<'a, 'input>>,
    aggregate_limit: usize,
) -> bool {
    label_paint_recipes_within_limit(
        series
            .filter_map(|series| child(series, "dataLabels"))
            .flat_map(label_shape_nodes),
        aggregate_limit,
    )
}

pub(super) fn parse_label_box_with_policy(
    sp_pr: Option<Node>,
    resolver: &dyn ColorResolver,
    allow_structured_paint: bool,
) -> Option<ChartLabelBox> {
    let sp = sp_pr?;
    if !allow_structured_paint {
        let fill_choice = sp.children().find(|node| {
            node.is_element()
                && matches!(
                    node.tag_name().name(),
                    "noFill" | "solidFill" | "gradFill" | "pattFill" | "blipFill" | "grpFill"
                )
        });
        let line = child(sp, "ln");
        let line_paint = line.and_then(|line| {
            line.children().find(|node| {
                node.is_element()
                    && matches!(
                        node.tag_name().name(),
                        "noFill" | "solidFill" | "gradFill" | "pattFill"
                    )
            })
        });
        return Some(ChartLabelBox {
            style: Some(parse_direct_chart_effect_style_from_sp_pr(sp, sp, resolver)),
            fill: None,
            fill_paint: None,
            fill_hidden: fill_choice
                .is_some_and(|paint| paint.tag_name().name() == "noFill")
                .then_some(true),
            fill_paint_authored: fill_choice.map(|_| true),
            border_color: None,
            border_fill: None,
            border_width_emu: line
                .and_then(|line| line.attribute("w"))
                .and_then(|value| value.parse::<u32>().ok()),
            border_hidden: line_paint
                .is_some_and(|paint| paint.tag_name().name() == "noFill")
                .then_some(true),
            border_paint_authored: line_paint.map(|_| true),
            border_dash: line
                .and_then(|line| child(line, "prstDash"))
                .and_then(|dash| dash.attribute("val"))
                .map(ToOwned::to_owned),
            border_dash_authored: line
                .and_then(|line| {
                    line.children().find(|node| {
                        node.is_element()
                            && matches!(node.tag_name().name(), "prstDash" | "custDash")
                    })
                })
                .map(|_| true),
            border_custom_dash: None,
            border_cap: line
                .and_then(|line| line.attribute("cap"))
                .map(ToOwned::to_owned),
            border_join: line.and_then(|line| {
                line.children().find_map(|node| {
                    node.is_element().then(|| match node.tag_name().name() {
                        "round" => Some("round".to_owned()),
                        "bevel" => Some("bevel".to_owned()),
                        "miter" => Some("miter".to_owned()),
                        _ => None,
                    })?
                })
            }),
            border_compound: line
                .and_then(|line| line.attribute("cmpd"))
                .map(ToOwned::to_owned),
        });
    }
    let fill = extract_direct_shape_fill(Some(sp), resolver);
    let border = extract_direct_shape_line_from_sp_pr(Some(sp), resolver);
    Some(ChartLabelBox {
        style: Some(parse_direct_chart_effect_style_from_sp_pr(sp, sp, resolver)),
        fill: fill.color,
        fill_paint: fill.fill,
        fill_hidden: fill.hidden,
        fill_paint_authored: fill.paint_authored,
        border_color: border.color,
        border_fill: border.fill,
        border_width_emu: border.width_emu,
        border_hidden: border.hidden,
        border_paint_authored: border.paint_authored,
        border_dash: border.dash,
        border_dash_authored: border.dash_authored,
        border_custom_dash: border.custom_dash,
        border_cap: border.cap,
        border_join: border.join,
        border_compound: border.compound,
    })
}

pub(super) fn parse_label_box(
    sp_pr: Option<Node>,
    resolver: &dyn ColorResolver,
) -> Option<ChartLabelBox> {
    let allow = sp_pr
        .map(|shape| label_paint_recipes_within_budget(std::iter::once(shape)))
        .unwrap_or(true);
    parse_label_box_with_policy(sp_pr, resolver, allow)
}

/// Parse ECMA-376 §21.2.2.45 `<c:dispUnits>`. The specification defines each
/// built-in token as a divisor; a custom unit is the divisor itself. Invalid,
/// zero, negative and non-finite public XML values are rejected rather than
/// reaching tick formatting as NaN/Infinity.
pub(super) fn parse_axis_display_units(
    axis: Node,
    resolver: &dyn ColorResolver,
) -> Option<ChartDisplayUnits> {
    let units = child(axis, "dispUnits")?;
    let (divisor, built_in_unit) = if let Some(custom) = child(units, "custUnit") {
        let divisor = custom.attribute("val")?.parse::<f64>().ok()?;
        if !divisor.is_finite() || divisor <= 0.0 {
            return None;
        }
        (divisor, None)
    } else {
        let built_in = child(units, "builtInUnit")?;
        // CT_BuiltInUnit@val defaults to `thousands` when the element exists.
        let token = built_in.attribute("val").unwrap_or("thousands");
        let divisor = match token {
            "hundreds" => 100.0,
            "thousands" => 1_000.0,
            "tenThousands" => 10_000.0,
            "hundredThousands" => 100_000.0,
            "millions" => 1_000_000.0,
            "tenMillions" => 10_000_000.0,
            "hundredMillions" => 100_000_000.0,
            "billions" => 1_000_000_000.0,
            "trillions" => 1_000_000_000_000.0,
            _ => return None,
        };
        (divisor, Some(token.to_string()))
    };

    let label = child(units, "dispUnitsLbl").map(|label| {
        let tx_pr = child(label, "txPr");
        let run_props = tx_pr.and_then(|tx| {
            tx.descendants().find(|node| {
                node.is_element() && matches!(node.tag_name().name(), "rPr" | "defRPr")
            })
        });
        let explicit_text = child(label, "tx")
            .map(|tx| flatten_rich_text(tx, None))
            .filter(|text| !text.is_empty());
        let text_paint = chart_text_paint([run_props], resolver);
        ChartDisplayUnitsLabel {
            text: explicit_text,
            manual_layout: child(label, "layout").and_then(extract_manual_layout),
            font_size_hpt: run_props
                .and_then(|props| props.attribute("sz"))
                .and_then(parse_text_font_size_hpt),
            font_bold: chart_text_bool_from_present_props(run_props, "b"),
            font_italic: chart_text_bool_from_present_props(run_props, "i"),
            font_color: text_paint.color,
            font_paint_authored: text_paint.authored.then_some(true),
            font_hidden: text_paint.hidden.then_some(true),
            font_face: tx_pr.and_then(first_latin_typeface),
            rotation: tx_pr
                .and_then(|tx| child(tx, "bodyPr"))
                .and_then(|body| body.attribute("rot"))
                .and_then(|value| value.parse::<i32>().ok()),
            box_style: parse_label_box(
                label
                    .children()
                    .find(|node| node.is_element() && node.tag_name().name() == "spPr"),
                resolver,
            ),
        }
    });

    Some(ChartDisplayUnits {
        divisor,
        built_in_unit,
        label,
    })
}

/// Parse `<c:dLbls><c:leaderLines>` into the authored visibility and line
/// properties. Missing fields remain absent so a linked Chart Style can supply
/// them without overriding local formatting.
/// `show` comes from the sibling `<c:showLeaderLines val>` (§21.2.2.183); the
/// stroke style comes from `<c:leaderLines>` (§21.2.2.92) `<c:spPr><a:ln>`.
pub(super) type ParsedLeaderLines = (
    bool,
    Option<String>,
    Option<u32>,
    Option<bool>,
    Option<String>,
    Option<bool>,
);

pub(super) fn parse_leader_lines(d_lbls: Node, resolver: &dyn ColorResolver) -> ParsedLeaderLines {
    // §21.2.2.183 `<c:showLeaderLines>` — CT_Boolean, so a bare element ⇒ true;
    // absent ⇒ false (no leader lines by default).
    let show = bool_child(d_lbls, "showLeaderLines").unwrap_or(false);
    let direct = child(d_lbls, "leaderLines")
        .map(|lines| extract_direct_shape_line(lines, resolver))
        .unwrap_or_default();
    (
        show,
        direct.color,
        direct.width_emu,
        direct.hidden,
        direct.dash,
        direct.paint_authored,
    )
}

/// Parse a series-level `<c:dLbls>` into `(series_defaults, per_idx_overrides)`.
/// ECMA-376 §21.2.2.47. Colors resolve through [`ColorResolver::resolve_shape_fill`]
/// so a scheme-color label text picks up its lumMod/lumOff transforms.
#[cfg(test)]
pub(super) fn parse_series_data_labels(
    ser_node: Node,
    resolver: &dyn ColorResolver,
    cellrange_cache: &std::collections::HashMap<u32, String>,
) -> (Option<ChartSeriesDataLabels>, Vec<ChartDataLabelOverride>) {
    let allow_label_paints = child(ser_node, "dLbls")
        .map(|labels| label_paint_recipes_within_budget(label_shape_nodes(labels)))
        .unwrap_or(true);
    parse_series_data_labels_with_paint_policy(
        ser_node,
        resolver,
        cellrange_cache,
        allow_label_paints,
    )
}

pub(super) fn parse_series_data_labels_with_paint_policy(
    ser_node: Node,
    resolver: &dyn ColorResolver,
    cellrange_cache: &std::collections::HashMap<u32, String>,
    chart_allows_label_paints: bool,
) -> (Option<ChartSeriesDataLabels>, Vec<ChartDataLabelOverride>) {
    let Some(d_lbls) = child(ser_node, "dLbls") else {
        return (None, Vec::new());
    };
    parse_data_labels_node_with_paint_policy(
        d_lbls,
        resolver,
        cellrange_cache,
        chart_allows_label_paints,
    )
}

pub(super) fn parse_data_labels_node_with_paint_policy(
    d_lbls: Node,
    resolver: &dyn ColorResolver,
    cellrange_cache: &std::collections::HashMap<u32, String>,
    chart_allows_label_paints: bool,
) -> (Option<ChartSeriesDataLabels>, Vec<ChartDataLabelOverride>) {
    let allow_label_paints =
        chart_allows_label_paints && label_paint_recipes_within_budget(label_shape_nodes(d_lbls));

    // CT_Boolean show-flag: element present ⇒ true unless `val` explicitly
    // disables it (§21.2.2, dml-chart.xsd `val` default `true`); element absent
    // ⇒ false (the flag defaults off when the deck names no show-flag element).
    let bool_attr = |n: Node, name: &str| bool_child(n, name).unwrap_or(false);

    let position = child(d_lbls, "dLblPos")
        .and_then(|n| n.attribute("val"))
        .map(|s| s.to_string());
    let format_code = child(d_lbls, "numFmt")
        .and_then(|n| n.attribute("formatCode"))
        .map(|s| s.to_string());
    let separator = child(d_lbls, "separator").map(|n| n.text().unwrap_or("").to_string());
    // defRPr fill / bold / size come from the dLbls-level `<c:txPr>`.
    let txpr = child(d_lbls, "txPr");
    let default_run_props = txpr.and_then(first_paragraph_default_run_props);
    let font_bold_default = chart_text_bool_from_present_props(default_run_props, "b");
    let font_size_default = default_run_props
        .and_then(|n| n.attribute("sz"))
        .and_then(parse_text_font_size_hpt);
    let font_face_default = default_run_props.and_then(first_latin_typeface);
    let default_text_paint = chart_text_paint([default_run_props], resolver);
    let font_color = default_text_paint.color.clone();
    let font_italic_default = chart_text_bool_from_present_props(default_run_props, "i");
    let font_language_default = default_run_props
        .and_then(|props| props.attribute("lang"))
        .map(ToOwned::to_owned);
    let font_baseline_default = default_run_props
        .and_then(|props| props.attribute("baseline"))
        .and_then(parse_chart_text_percentage);
    let body_pr = txpr.and_then(|tx| child(tx, "bodyPr"));
    let body_style = chart_label_body_style(body_pr);

    // §21.2.2.197 series-level callout-box shape (`<c:dLbls><c:spPr>`) and
    // §21.2.2.183/§21.2.2.92 leader-line style. `<c:spPr>` may appear both as a
    // direct child of `<c:dLbls>` (the series default) and inside each
    // `<c:dLbl>` (per-point) — pick the direct child here.
    let label_box = parse_label_box_with_policy(
        d_lbls
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "spPr"),
        resolver,
        allow_label_paints,
    );
    let (
        show_leader_lines,
        leader_line_color,
        leader_line_width_emu,
        leader_line_hidden,
        leader_line_dash,
        leader_line_paint_authored,
    ) = parse_leader_lines(d_lbls, resolver);
    let leader_line_style = child(d_lbls, "leaderLines")
        .and_then(|lines| parse_direct_chart_effect_style(lines, resolver));

    let series_defaults = ChartSeriesDataLabels {
        deleted: bool_child(d_lbls, "delete"),
        show_val: bool_attr(d_lbls, "showVal"),
        show_cat_name: bool_attr(d_lbls, "showCatName"),
        show_ser_name: bool_attr(d_lbls, "showSerName"),
        show_percent: bool_attr(d_lbls, "showPercent"),
        show_bubble_size: bool_attr(d_lbls, "showBubbleSize"),
        show_legend_key: bool_attr(d_lbls, "showLegendKey"),
        position: position.clone(),
        font_color: font_color.clone(),
        font_paint_authored: default_text_paint.authored.then_some(true),
        font_hidden: default_text_paint.hidden.then_some(true),
        format_code,
        separator,
        font_bold: font_bold_default,
        font_italic: font_italic_default,
        font_language: font_language_default,
        font_baseline: font_baseline_default,
        font_size_hpt: font_size_default,
        font_face: font_face_default,
        text_rotation: body_style.rotation,
        text_wrap: body_style.wrap,
        text_vertical_anchor: body_style.anchor,
        text_vertical_mode: body_style.vertical_mode,
        text_l_ins_emu: body_style.left_inset,
        text_t_ins_emu: body_style.top_inset,
        text_r_ins_emu: body_style.right_inset,
        text_b_ins_emu: body_style.bottom_inset,
        text_body_authored: body_style.authored.then_some(true),
        text_align: txpr
            .and_then(first_paragraph_properties)
            .and_then(|props| attr(&props, "algn")),
        label_box,
        show_leader_lines,
        leader_line_color,
        leader_line_width_emu,
        leader_line_hidden,
        leader_line_dash,
        leader_line_paint_authored,
        leader_line_style,
    };

    let mut overrides = Vec::new();
    for dl in d_lbls
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "dLbl")
    {
        let idx = child(dl, "idx")
            .and_then(|n| n.attribute("val"))
            .and_then(|v| v.parse::<u32>().ok())
            .unwrap_or(0);
        // §21.2.2.43 per-point `<c:delete>` — CT_Boolean, so a bare
        // `<c:delete/>` removes this point's label (val default true).
        let deleted = bool_child(dl, "delete");
        let pos = child(dl, "dLblPos")
            .and_then(|n| n.attribute("val"))
            .map(|s| s.to_string());
        let cache_for_idx = cellrange_cache.get(&idx).map(|s| s.as_str());
        let rich_runs = if deleted == Some(true) {
            None
        } else {
            parse_data_label_rich_runs(dl, resolver, cache_for_idx)
        };
        let text = if deleted == Some(true) {
            String::new()
        } else if let Some(runs) = rich_runs.as_ref() {
            runs.iter().map(|run| run.text.as_str()).collect()
        } else {
            match child(dl, "tx") {
                Some(tx_node) => flatten_rich_text(tx_node, cache_for_idx),
                None => cache_for_idx.unwrap_or("").to_string(),
            }
        };
        // Paragraph default character properties apply to all runs. Individual
        // `<a:rPr>` values remain on `rich_runs` and must never be promoted to
        // the label-wide fallback used by sibling runs.
        let default_run_props = child(dl, "txPr").and_then(first_paragraph_default_run_props);
        let text_paint = chart_text_paint([default_run_props], resolver);
        let font_color = text_paint.color.clone();
        let font_size_hpt = default_run_props
            .and_then(|run| run.attribute("sz"))
            .and_then(parse_text_font_size_hpt);
        let font_face = default_run_props.and_then(first_latin_typeface);
        let font_bold = chart_text_bool_from_present_props(default_run_props, "b");
        let font_italic = chart_text_bool_from_present_props(default_run_props, "i");
        let font_language = default_run_props
            .and_then(|props| props.attribute("lang"))
            .map(ToOwned::to_owned);
        let font_baseline = default_run_props
            .and_then(|props| props.attribute("baseline"))
            .and_then(parse_chart_text_percentage);
        let rich_body_pr = child(dl, "tx")
            .and_then(|tx| child(tx, "rich"))
            .and_then(|rich| child(rich, "bodyPr"));
        let direct_body_pr = child(dl, "txPr").and_then(|tx| child(tx, "bodyPr"));
        let label_body = merge_chart_label_body_styles(
            chart_label_body_style(rich_body_pr),
            chart_label_body_style(direct_body_pr),
        );
        let text_align = child(dl, "txPr")
            .and_then(first_paragraph_properties)
            .and_then(|props| attr(&props, "algn"));
        // Per-point callout box (`<c:dLbl>` §21.2.2.47 `<c:spPr>` §21.2.2.197):
        // direct child spPr overrides the series-default box for this one point.
        let label_box = parse_label_box_with_policy(
            dl.children()
                .find(|n| n.is_element() && n.tag_name().name() == "spPr"),
            resolver,
            allow_label_paints,
        );
        // Per-point show-flags (§21.2.2.47 CT_DLbl carries the same show-flag
        // group as CT_DLbls). Read as `Option` so an absent flag falls through
        // to the series default; a present flag overrides it for this point.
        // CT_Boolean: a present-but-`val`-omitted flag is `Some(true)`.
        let opt_bool_flag = |name: &str| -> Option<bool> { bool_child(dl, name) };
        overrides.push(ChartDataLabelOverride {
            idx,
            text,
            rich_runs,
            position: pos,
            font_color,
            font_paint_authored: text_paint.authored.then_some(true),
            font_hidden: text_paint.hidden.then_some(true),
            font_size_hpt,
            font_face,
            font_bold,
            font_italic,
            font_language,
            font_baseline,
            text_rotation: label_body.rotation,
            text_wrap: label_body.wrap,
            text_vertical_anchor: label_body.anchor,
            text_vertical_mode: label_body.vertical_mode,
            text_l_ins_emu: label_body.left_inset,
            text_t_ins_emu: label_body.top_inset,
            text_r_ins_emu: label_body.right_inset,
            text_b_ins_emu: label_body.bottom_inset,
            text_body_authored: label_body.authored.then_some(true),
            text_align,
            format_code: child(dl, "numFmt").and_then(|node| attr(&node, "formatCode")),
            separator: child(dl, "separator").map(|node| node.text().unwrap_or("").to_owned()),
            manual_layout: child(dl, "layout").and_then(extract_manual_layout),
            label_box,
            show_val: opt_bool_flag("showVal"),
            show_cat_name: opt_bool_flag("showCatName"),
            show_ser_name: opt_bool_flag("showSerName"),
            show_percent: opt_bool_flag("showPercent"),
            show_bubble_size: opt_bool_flag("showBubbleSize"),
            show_legend_key: opt_bool_flag("showLegendKey"),
            // §21.2.2.43 `<c:delete>` — record genuine deletes distinctly from a
            // style-only `<c:dLbl>` so the renderer never mistakes an empty tx
            // (compose-from-flags) for a removed label.
            deleted,
        });
    }

    let any_authored_boolean = [
        "showVal",
        "showCatName",
        "showSerName",
        "showPercent",
        "showBubbleSize",
        "showLegendKey",
        "showLeaderLines",
        "delete",
    ]
    .iter()
    .any(|name| child(d_lbls, name).is_some());
    let any_default = any_authored_boolean
        || series_defaults.deleted.is_some()
        || series_defaults.show_val
        || series_defaults.show_cat_name
        || series_defaults.show_ser_name
        || series_defaults.show_percent
        || series_defaults.show_bubble_size
        || series_defaults.show_legend_key
        || series_defaults.position.is_some()
        || series_defaults.font_color.is_some()
        || series_defaults.font_paint_authored.is_some()
        || series_defaults.font_hidden.is_some()
        || series_defaults.format_code.is_some()
        || series_defaults.font_bold.is_some()
        || series_defaults.font_italic.is_some()
        || series_defaults.font_language.is_some()
        || series_defaults.font_baseline.is_some()
        || series_defaults.font_size_hpt.is_some()
        || series_defaults.font_face.is_some()
        || series_defaults.text_rotation.is_some()
        || series_defaults.text_wrap.is_some()
        || series_defaults.text_vertical_anchor.is_some()
        || series_defaults.text_vertical_mode.is_some()
        || series_defaults.text_l_ins_emu.is_some()
        || series_defaults.text_t_ins_emu.is_some()
        || series_defaults.text_r_ins_emu.is_some()
        || series_defaults.text_b_ins_emu.is_some()
        || series_defaults.text_body_authored.is_some()
        || series_defaults.text_align.is_some()
        || series_defaults.label_box.is_some()
        || series_defaults.show_leader_lines
        || series_defaults.leader_line_color.is_some()
        || series_defaults.leader_line_width_emu.is_some()
        || series_defaults.leader_line_hidden.is_some()
        || series_defaults.leader_line_dash.is_some()
        || series_defaults.leader_line_style.is_some();
    let series_out = if any_default {
        Some(series_defaults)
    } else {
        None
    };
    (series_out, overrides)
}

/// Parse the subset of CT_DLbls that Office permits on a chart-group parent.
/// MS-OE376 §2.1.1476 excludes collection delete, point dLbl, leader lines,
/// number/text/shape formatting and txPr at this level. Keeping this parser
/// separate prevents invalid group content from being deep-parsed or cloned
/// into every owned series.
pub(super) fn parse_chart_group_data_labels(d_lbls: Node) -> Option<ChartSeriesDataLabels> {
    let position_node = child(d_lbls, "dLblPos");
    let separator_node = child(d_lbls, "separator");
    let flag_names = [
        "showVal",
        "showCatName",
        "showSerName",
        "showPercent",
        "showBubbleSize",
        "showLegendKey",
    ];
    let has_authored_semantics = position_node.is_some()
        || separator_node.is_some()
        || flag_names.iter().any(|name| child(d_lbls, name).is_some());
    if !has_authored_semantics {
        return None;
    }
    Some(ChartSeriesDataLabels {
        show_val: bool_child(d_lbls, "showVal").unwrap_or(false),
        show_cat_name: bool_child(d_lbls, "showCatName").unwrap_or(false),
        show_ser_name: bool_child(d_lbls, "showSerName").unwrap_or(false),
        show_percent: bool_child(d_lbls, "showPercent").unwrap_or(false),
        show_bubble_size: bool_child(d_lbls, "showBubbleSize").unwrap_or(false),
        show_legend_key: bool_child(d_lbls, "showLegendKey").unwrap_or(false),
        position: position_node.and_then(|node| attr(&node, "val")),
        separator: separator_node.map(|node| {
            node.text()
                .unwrap_or("")
                .chars()
                .take(MAX_DATA_LABEL_RICH_SCALARS)
                .collect()
        }),
        ..ChartSeriesDataLabels::default()
    })
}

pub(super) fn chart_group_separator_projection_within_budget(
    entries: impl IntoIterator<Item = (usize, usize)>,
) -> bool {
    entries
        .into_iter()
        .try_fold(0usize, |total, (separator_scalars, owned_series)| {
            separator_scalars
                .checked_mul(owned_series)
                .and_then(|projected| total.checked_add(projected))
                .filter(|next| *next <= MAX_CHART_CACHE_POINTS)
        })
        .is_some()
}

pub(super) fn merge_chart_label_boxes(
    lower: Option<ChartLabelBox>,
    higher: Option<ChartLabelBox>,
) -> Option<ChartLabelBox> {
    let Some(higher) = higher else {
        return lower;
    };
    let mut merged = lower.unwrap_or_default();
    // `style` retains direct spPr presence and effects that the scalar legacy
    // fields cannot represent. A higher-precedence series/point label shape
    // owns that complete carrier even when its spPr is intentionally empty:
    // chart-style allowNo*Override modifiers consult this provenance.
    if higher.style.is_some() {
        merged.style = higher.style.clone();
    }
    let higher_fill_authored = higher.fill_paint_authored == Some(true)
        || higher.fill.is_some()
        || higher.fill_paint.is_some()
        || higher.fill_hidden.is_some();
    if higher_fill_authored {
        merged.fill = higher.fill;
        merged.fill_paint = higher.fill_paint;
        merged.fill_hidden = higher.fill_hidden;
        merged.fill_paint_authored = higher.fill_paint_authored;
    }
    let higher_border_authored = higher.border_paint_authored == Some(true)
        || higher.border_color.is_some()
        || higher.border_fill.is_some()
        || higher.border_hidden.is_some();
    if higher_border_authored {
        merged.border_color = higher.border_color;
        merged.border_fill = higher.border_fill;
        merged.border_hidden = higher.border_hidden;
        merged.border_paint_authored = higher.border_paint_authored;
    }
    if higher.border_width_emu.is_some() {
        merged.border_width_emu = higher.border_width_emu;
    }
    if higher.border_dash_authored.is_some() {
        merged.border_dash = higher.border_dash;
        merged.border_custom_dash = higher.border_custom_dash;
        merged.border_dash_authored = higher.border_dash_authored;
    }
    if higher.border_cap.is_some() {
        merged.border_cap = higher.border_cap;
    }
    if higher.border_join.is_some() {
        merged.border_join = higher.border_join;
    }
    if higher.border_compound.is_some() {
        merged.border_compound = higher.border_compound;
    }
    Some(merged)
}

pub(super) fn merge_chart_series_data_labels(
    lower: Option<ChartSeriesDataLabels>,
    higher: Option<ChartSeriesDataLabels>,
    higher_node: Option<Node>,
) -> Option<ChartSeriesDataLabels> {
    if higher.is_none() {
        return lower;
    }
    let mut merged = lower.unwrap_or_default();
    let higher = higher.expect("checked above");
    let authored_bool = |name: &str| higher_node.is_some_and(|node| child(node, name).is_some());
    if authored_bool("showVal") {
        merged.show_val = higher.show_val;
    }
    if authored_bool("showCatName") {
        merged.show_cat_name = higher.show_cat_name;
    }
    if authored_bool("showSerName") {
        merged.show_ser_name = higher.show_ser_name;
    }
    if authored_bool("showPercent") {
        merged.show_percent = higher.show_percent;
    }
    if authored_bool("showBubbleSize") {
        merged.show_bubble_size = higher.show_bubble_size;
    }
    if authored_bool("showLegendKey") {
        merged.show_legend_key = higher.show_legend_key;
    }
    if authored_bool("showLeaderLines") {
        merged.show_leader_lines = higher.show_leader_lines;
    }
    if authored_bool("delete") {
        merged.deleted = higher.deleted;
    }

    macro_rules! overlay_option {
        ($field:ident) => {
            if higher.$field.is_some() {
                merged.$field = higher.$field;
            }
        };
    }
    overlay_option!(position);
    overlay_option!(format_code);
    overlay_option!(separator);
    overlay_option!(font_bold);
    overlay_option!(font_italic);
    overlay_option!(font_language);
    overlay_option!(font_baseline);
    overlay_option!(font_size_hpt);
    overlay_option!(font_face);
    overlay_option!(text_rotation);
    overlay_option!(text_wrap);
    overlay_option!(text_vertical_anchor);
    overlay_option!(text_vertical_mode);
    overlay_option!(text_l_ins_emu);
    overlay_option!(text_t_ins_emu);
    overlay_option!(text_r_ins_emu);
    overlay_option!(text_b_ins_emu);
    overlay_option!(text_body_authored);
    overlay_option!(text_align);
    overlay_option!(leader_line_color);
    overlay_option!(leader_line_width_emu);
    overlay_option!(leader_line_hidden);
    overlay_option!(leader_line_dash);
    overlay_option!(leader_line_style);
    if higher.font_paint_authored == Some(true) {
        merged.font_color = higher.font_color;
        merged.font_hidden = higher.font_hidden;
        merged.font_paint_authored = higher.font_paint_authored;
    }
    merged.label_box = merge_chart_label_boxes(merged.label_box, higher.label_box);
    Some(merged)
}

/// Read a `<c:numRef><c:numCache>` or `<c:numLit>` block under `parent` and
/// return per-point values keyed by `<c:pt idx>`. Length is at least
/// `expected_len` (padded with `None`).
pub(super) fn extract_num_block(parent: Node, expected_len: usize) -> Vec<Option<f64>> {
    let cache = parent.descendants().find(|n| {
        n.is_element() && (n.tag_name().name() == "numCache" || n.tag_name().name() == "numLit")
    });
    let Some(cache) = cache else {
        return Vec::new();
    };
    let pt_count: usize = child(cache, "ptCount")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<usize>().ok())
        .unwrap_or(expected_len);
    let len = pt_count.max(expected_len);
    let mut values: Vec<Option<f64>> = vec![None; len];
    for pt in cache
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "pt")
    {
        let Some(idx) = pt.attribute("idx").and_then(|v| v.parse::<usize>().ok()) else {
            continue;
        };
        let v = child(pt, "v")
            .and_then(|n| n.text())
            .and_then(|s| s.trim().parse::<f64>().ok());
        if idx < values.len() {
            values[idx] = v;
        }
    }
    values
}

/// Parse all `<c:errBars>` direct children of a series and resolve per-point
/// plus / minus deltas to absolute numbers. Each errBars block fixes a
/// direction (x|y); a series can have at most one of each direction.
/// ECMA-376 §21.2.2.20.
pub(super) fn parse_error_bars(
    ser_node: Node,
    series_values: &[Option<f64>],
    resolver: &dyn ColorResolver,
) -> Vec<ChartErrBars> {
    let mut result = Vec::new();
    for eb in ser_node
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "errBars")
    {
        let dir = child(eb, "errDir")
            .and_then(|n| n.attribute("val"))
            .unwrap_or("y")
            .to_string();
        let bar_type = child(eb, "errBarType")
            .and_then(|n| n.attribute("val"))
            .unwrap_or("both")
            .to_string();
        let val_type = child(eb, "errValType")
            .and_then(|n| n.attribute("val"))
            .unwrap_or("fixedVal")
            .to_string();
        // §21.2.2.117 `<c:noEndCap>` — CT_Boolean, so a bare element ⇒ true (no
        // I-beam end caps); absent ⇒ false (draw end caps by default).
        let no_end_cap = bool_child(eb, "noEndCap").unwrap_or(false);

        let n_points = series_values.len();
        let mut plus: Vec<Option<f64>> = vec![None; n_points];
        let mut minus: Vec<Option<f64>> = vec![None; n_points];

        match val_type.as_str() {
            "cust" => {
                for (slot, target) in [("plus", &mut plus), ("minus", &mut minus)] {
                    let Some(side) = child(eb, slot) else {
                        continue;
                    };
                    let vals = extract_num_block(side, n_points);
                    if !vals.is_empty() {
                        let len = vals.len().min(target.len());
                        target[..len].copy_from_slice(&vals[..len]);
                    }
                }
            }
            "fixedVal" => {
                let v = child(eb, "val")
                    .and_then(|n| n.attribute("val"))
                    .and_then(|s| s.parse::<f64>().ok())
                    .unwrap_or(0.0);
                for i in 0..n_points {
                    plus[i] = Some(v);
                    minus[i] = Some(v);
                }
            }
            "percentage" => {
                let pct = child(eb, "val")
                    .and_then(|n| n.attribute("val"))
                    .and_then(|s| s.parse::<f64>().ok())
                    .unwrap_or(0.0);
                for (i, v) in series_values.iter().enumerate() {
                    if let Some(val) = v {
                        let d = val.abs() * pct / 100.0;
                        plus[i] = Some(d);
                        minus[i] = Some(d);
                    }
                }
            }
            "stdErr" | "stdDev" => {
                let nums: Vec<f64> = series_values.iter().filter_map(|v| *v).collect();
                if !nums.is_empty() {
                    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
                    let var =
                        nums.iter().map(|v| (v - mean).powi(2)).sum::<f64>() / nums.len() as f64;
                    let std = var.sqrt();
                    let mult = child(eb, "val")
                        .and_then(|n| n.attribute("val"))
                        .and_then(|s| s.parse::<f64>().ok())
                        .unwrap_or(1.0);
                    let sample = if val_type == "stdErr" {
                        std / (nums.len() as f64).sqrt()
                    } else {
                        std
                    };
                    let delta = sample * mult;
                    for i in 0..n_points {
                        plus[i] = Some(delta);
                        minus[i] = Some(delta);
                    }
                }
            }
            _ => {}
        }

        let direct_line = extract_direct_shape_line(eb, resolver);

        result.push(ChartErrBars {
            style: parse_direct_chart_effect_style(eb, resolver),
            dir,
            bar_type,
            plus,
            minus,
            no_end_cap,
            color: direct_line.color,
            line_width_emu: direct_line.width_emu,
            dash: direct_line.dash,
            hidden: direct_line.hidden,
            line_paint_authored: direct_line.paint_authored,
        });
    }
    result
}

/// Positional string-cache collector for `<c:cat>` / `<c:xVal>`. Reads
/// `<c:ptCount>` to size the result, then places each `<c:pt idx>` string at its
/// index (multi-level caches use the innermost `<c:lvl>`). Unlike a naive
/// document-order collector this preserves gaps (sparse caches) so a category
/// list that starts at `idx=1`, or a value series with a hole, keeps its true
/// length and alignment (ECMA-376 §21.2.2.20/.75/.181).
pub(super) fn collect_str_cache_positional(ser_node: Node, child_tag: &str) -> Vec<String> {
    let Some(container) = ser_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == child_tag)
    else {
        return Vec::new();
    };

    // Multi-level categories: use only the first (innermost) lvl.
    if let Some(multi_cache) = container
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "multiLvlStrCache")
    {
        let pt_count: usize = child(multi_cache, "ptCount")
            .and_then(|n| n.attribute("val"))
            .and_then(|v| v.parse().ok())
            .unwrap_or(0);
        if let Some(first_lvl) = child(multi_cache, "lvl") {
            let mut pts: Vec<(usize, String)> = Vec::new();
            for pt in first_lvl
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "pt")
            {
                let idx: usize = pt
                    .attribute("idx")
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0);
                let val = child(pt, "v")
                    .and_then(|n| n.text())
                    .unwrap_or("")
                    .to_string();
                pts.push((idx, val));
            }
            let len = pt_count.max(pts.iter().map(|(i, _)| i + 1).max().unwrap_or(0));
            let mut result = vec![String::new(); len];
            for (idx, val) in pts {
                if idx < result.len() {
                    result[idx] = val;
                }
            }
            return result;
        }
    }

    // Standard strRef/strCache or numRef/numCache.
    let mut pt_count: usize = 0;
    let mut pts: Vec<(usize, String)> = Vec::new();
    for desc in container.descendants() {
        match desc.tag_name().name() {
            "ptCount" => {
                pt_count = desc
                    .attribute("val")
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0);
            }
            "pt" => {
                let idx: usize = desc
                    .attribute("idx")
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0);
                let val = child(desc, "v")
                    .and_then(|n| n.text())
                    .unwrap_or("")
                    .to_string();
                pts.push((idx, val));
            }
            _ => {}
        }
    }
    if pt_count == 0 {
        pt_count = pts.len();
    }
    let mut result = vec![String::new(); pt_count];
    for (idx, val) in pts {
        if idx < result.len() {
            result[idx] = val;
        }
    }
    result
}

/// Preserve every authored `<c:multiLvlStrCache><c:lvl>` in document order
/// (leaf/deepest level first). Empty slots are significant: a non-empty outer
/// label starts a span that continues until the next non-empty slot.
pub(super) fn collect_multi_level_str_cache(
    ser_node: Node<'_, '_>,
    child_tag: &str,
) -> Option<Vec<Vec<String>>> {
    let container = ser_node
        .children()
        .find(|node| node.is_element() && node.tag_name().name() == child_tag)?;
    let cache = container
        .descendants()
        .find(|node| node.is_element() && node.tag_name().name() == "multiLvlStrCache")?;
    let levels: Vec<_> = cache
        .children()
        .filter(|node| node.is_element() && node.tag_name().name() == "lvl")
        .collect();
    if levels.len() < 2 {
        return None;
    }

    let authored_count = child(cache, "ptCount")
        .and_then(|node| node.attribute("val"))
        .and_then(|value| value.parse::<usize>().ok())
        .unwrap_or(0);
    let max_index = levels
        .iter()
        .flat_map(|level| {
            level
                .children()
                .filter(|node| node.is_element() && node.tag_name().name() == "pt")
        })
        .filter_map(|point| point.attribute("idx"))
        .filter_map(|value| value.parse::<usize>().ok())
        .max()
        .map(|index| index.saturating_add(1))
        .unwrap_or(0);
    let point_count = authored_count.max(max_index);
    let total_slots = point_count.checked_mul(levels.len())?;
    if point_count == 0
        || point_count > MAX_CHART_CACHE_POINTS
        || total_slots > MAX_CHART_CACHE_POINTS
    {
        return None;
    }

    Some(
        levels
            .into_iter()
            .map(|level| {
                let mut values = vec![String::new(); point_count];
                for point in level
                    .children()
                    .filter(|node| node.is_element() && node.tag_name().name() == "pt")
                {
                    let Some(index) = point
                        .attribute("idx")
                        .and_then(|value| value.parse::<usize>().ok())
                        .filter(|index| *index < point_count)
                    else {
                        continue;
                    };
                    values[index] = child(point, "v")
                        .and_then(|node| node.text())
                        .unwrap_or("")
                        .to_string();
                }
                values
            })
            .collect(),
    )
}

/// Positional numeric-cache collector for `<c:val>` / `<c:yVal>`. Reads
/// `<c:ptCount>` to size the result, then places each `<c:pt idx>` value at its
/// index (padding gaps with `None`). Sparse-safe companion to
/// [`collect_str_cache_positional`].
pub(super) fn collect_num_cache_positional(ser_node: Node, child_tag: &str) -> Vec<Option<f64>> {
    let Some(container) = ser_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == child_tag)
    else {
        return Vec::new();
    };

    let mut pt_count: usize = 0;
    let mut pts: Vec<(usize, f64)> = Vec::new();
    for desc in container.descendants() {
        match desc.tag_name().name() {
            "ptCount" => {
                pt_count = desc
                    .attribute("val")
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0);
            }
            "pt" => {
                let idx: usize = desc
                    .attribute("idx")
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0);
                if let Some(v) = child(desc, "v")
                    .and_then(|n| n.text())
                    .and_then(|t| t.parse::<f64>().ok())
                {
                    pts.push((idx, v));
                }
            }
            _ => {}
        }
    }
    if pt_count == 0 {
        pt_count = pts.len();
    }
    let mut result: Vec<Option<f64>> = vec![None; pt_count];
    for (idx, val) in pts {
        if idx < result.len() {
            result[idx] = Some(val);
        }
    }
    result
}

/// Package-specific resolver for legacy chart formulas whose authored cache or
/// literal is absent. DrawingML owns the series-field walk; a host package such
/// as XLSX supplies only the external data lookup. DOCX and PPTX use the
/// no-op resolver through [`parse_chart_part`].
pub trait ChartReferenceResolver {
    /// `None` means the host could not resolve the formula. `Some` with an
    /// empty or all-gap vector is a successfully resolved empty range.
    fn resolve_strings(&mut self, formula: &str) -> Option<Vec<String>>;
    fn resolve_numbers(&mut self, formula: &str) -> Option<Vec<Option<f64>>>;

    /// Resolve whether each source cell is in a hidden host row or column.
    /// The formula's cached chart values remain authoritative; this method
    /// supplies provenance only, so package-specific lookup stays outside the
    /// shared DrawingML parser.
    fn resolve_hidden(&mut self, _formula: &str) -> Option<Vec<bool>> {
        None
    }

    /// Resolve the source-linked number format of a numeric reference. ChartEx
    /// dimensions frequently omit caches and `<cx:numFmt>` while data labels
    /// remain linked to the first source cell's worksheet number format.
    fn resolve_number_format(&mut self, _formula: &str) -> Option<String> {
        None
    }

    /// Preserve the worksheet built-in format identity separately from its
    /// textual expansion. Built-in 14 is application-locale-sensitive.
    fn resolve_number_format_id(&mut self, _formula: &str) -> Option<u32> {
        None
    }

    /// Resolve a rectangular hierarchy source into chartEx level vectors in
    /// document order (deepest level first, root level last). Legacy callers
    /// and one-column sources naturally fall back to a single level.
    fn resolve_string_levels(&mut self, formula: &str) -> Option<Vec<Vec<String>>> {
        self.resolve_strings(formula).map(|values| vec![values])
    }
}

pub(super) struct EmptyChartReferenceResolver;

impl ChartReferenceResolver for EmptyChartReferenceResolver {
    fn resolve_strings(&mut self, _formula: &str) -> Option<Vec<String>> {
        None
    }

    fn resolve_numbers(&mut self, _formula: &str) -> Option<Vec<Option<f64>>> {
        None
    }
}

pub(super) fn has_authored_reference_data(container: Node<'_, '_>) -> bool {
    container.descendants().any(|node| {
        node.is_element()
            && matches!(
                node.tag_name().name(),
                "strCache" | "numCache" | "multiLvlStrCache" | "strLit" | "numLit"
            )
    })
}

pub(super) fn reference_formula(container: Node<'_, '_>) -> Option<String> {
    container
        .descendants()
        .find(|node| node.is_element() && node.tag_name().name() == "f")
        .and_then(|node| node.text())
        .map(str::trim)
        .filter(|formula| !formula.is_empty())
        .map(str::to_owned)
}

pub(super) fn collect_string_source(
    ser_node: Node<'_, '_>,
    child_tag: &str,
    references: &mut dyn ChartReferenceResolver,
) -> Option<Vec<String>> {
    let container = ser_node
        .children()
        .find(|node| node.is_element() && node.tag_name().name() == child_tag)?;
    if has_authored_reference_data(container) {
        let authored = collect_str_cache_positional(ser_node, child_tag);
        if !authored.is_empty() {
            return Some(authored);
        }
        // A zero-point cache carries no usable snapshot. When the host can
        // resolve the accompanying formula (notably XLSX worksheet cells),
        // use that live source just as Office does. Keep an unresolved empty
        // cache explicit so non-XLSX hosts never inherit unrelated data.
        return reference_formula(container)
            .and_then(|formula| references.resolve_strings(&formula))
            .or(Some(authored));
    }
    reference_formula(container).and_then(|formula| references.resolve_strings(&formula))
}

pub(super) fn collect_number_source(
    ser_node: Node<'_, '_>,
    child_tag: &str,
    references: &mut dyn ChartReferenceResolver,
) -> Option<Vec<Option<f64>>> {
    let container = ser_node
        .children()
        .find(|node| node.is_element() && node.tag_name().name() == child_tag)?;
    if has_authored_reference_data(container) {
        let authored = collect_num_cache_positional(ser_node, child_tag);
        if !authored.is_empty() {
            return Some(authored);
        }
        return reference_formula(container)
            .and_then(|formula| references.resolve_numbers(&formula))
            .or(Some(authored));
    }
    reference_formula(container).and_then(|formula| references.resolve_numbers(&formula))
}

pub(super) fn collect_source_hidden(
    ser_node: Node<'_, '_>,
    child_tag: &str,
    references: &mut dyn ChartReferenceResolver,
) -> Option<Vec<bool>> {
    let container = ser_node
        .children()
        .find(|node| node.is_element() && node.tag_name().name() == child_tag)?;
    reference_formula(container).and_then(|formula| references.resolve_hidden(&formula))
}

pub(super) fn merge_source_hidden(target: &mut Option<Vec<bool>>, incoming: Option<Vec<bool>>) {
    let Some(incoming) = incoming else {
        return;
    };
    match target {
        Some(existing) => {
            existing.resize(existing.len().max(incoming.len()), false);
            for (index, hidden) in incoming.into_iter().enumerate() {
                existing[index] |= hidden;
            }
        }
        None => *target = Some(incoming),
    }
}

/// Formula identity for a source that genuinely needs the host resolver.
/// Authored caches/literals deliberately return `None`: even when their `<f>`
/// text matches another series, their authored point data remains authoritative.
pub(super) fn external_reference_formula(
    ser_node: Node<'_, '_>,
    child_tag: &str,
) -> Option<String> {
    let container = ser_node
        .children()
        .find(|node| node.is_element() && node.tag_name().name() == child_tag)?;
    (!has_authored_reference_data(container))
        .then(|| reference_formula(container))
        .flatten()
}
