use super::*;

// ============================================================================
// Shared chart data model
// ============================================================================
//
// These structs are the Rust mirror of the TypeScript `ChartModel` in
// `packages/core/src/types/chart.ts`. Both the pptx and xlsx Rust parsers build
// a `ChartModel` and emit it as a single nested `chart` object, so the TS
// renderer (`@silurus/ooxml-core`'s `renderChart`) receives a value that is
// already `ChartModel`-shaped and needs no per-field adapter.
//
// Field-for-field parity with the TS interface is the contract. Serde
// `rename_all = "camelCase"` matches the TS key names. The REQUIRED TS fields
// (no `?`) are serialized unconditionally so the wire object always carries
// them — an `Option<T>` REQUIRED field emits `null` when `None` (matching
// `T | null`), and a `bool`/`Vec` REQUIRED field emits `false`/`[]`. The
// OPTIONAL TS fields (`field?: …`) keep `skip_serializing_if` so they drop off
// the wire when unset; the renderer treats a missing key and an explicit `null`
// identically (every read is `?? default` / `!= null`), so this is
// render-equivalent to emitting `null`.
//
// All field references are ECMA-376 / ISO-29500 part 1 §21.2 (DrawingML Charts)
// as documented on the TS side; see that file for the per-field spec citations.

/// One DrawingML custom-dash atom, normalized to line-width multipliers.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartLineDashSegment {
    /// Dash length as a multiplier of the rendered line width.
    pub dash: f64,
    /// Following space as a multiplier of the rendered line width.
    pub space: f64,
}

/// Effective paint for one role in an Office 2013+ Chart Style part.
#[derive(Serialize, Deserialize, Debug, Clone, Default, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartExElementStyle {
    /// The source element carried a local `spPr`. On a direct formatting
    /// carrier this distinguishes an omitted shape from an authored empty or
    /// opposite-component shape for CT_StyleEntry override modifiers.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shape_properties_present: Option<bool>,
    /// MS-ODRAWXML CT_StyleEntry modifiers. When a direct `spPr` is present,
    /// these allow an omitted paint component to override the linked fill or
    /// line with no paint; geometry and other style components still inherit.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub allow_no_fill_override: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub allow_no_line_override: Option<bool>,
    /// Linked Chart Style text defaults (`fontRef` + `defRPr`). Direct chart
    /// text properties remain authoritative when the renderer composes them.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_size_hpt: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_italic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_color: Option<String>,
    /// Per-color-style-index text colors resolved from `fontRef/styleClr`.
    /// A scalar `font_color` remains as the index-zero compatibility value.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_colors: Option<Vec<Option<String>>>,
    /// Fixed zero-based CT_ColorStyle index from `fontRef/styleClr`; omission
    /// means the text-bearing object supplies its relative formatting index.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_color_index: Option<usize>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_formatting_indices: Option<Vec<usize>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_paint_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_hidden: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_face: Option<String>,
    /// `<cs:*><cs:defRPr lang>` (MS-ODRAWXML CT_StyleEntry). Preserved as an
    /// authored language tag; Canvas does not infer language from label text.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_language: Option<String>,
    /// `<cs:*><cs:defRPr baseline>` as a normalized fraction (`0.3` = 30%).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_baseline: Option<f64>,
    /// Linked `<cs:*><cs:bodyPr>` text-body defaults. Omission remains distinct
    /// from an authored zero/empty value so direct chart text stays dominant.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_rotation: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_wrap: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_vertical_anchor: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_vertical_mode: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_l_ins_emu: Option<i64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_t_ins_emu: Option<i64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_r_ins_emu: Option<i64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_b_ins_emu: Option<i64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_body_authored: Option<bool>,
    /// Per-color-style-index DrawingML fill recipes after `phClr`
    /// substitution. Solid fills remain duplicated in `fill_colors` for wire
    /// compatibility; gradient and pattern fills are retained only here.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_paints: Option<Vec<Option<ChartStyleFill>>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_colors: Option<Vec<Option<String>>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_hidden: Option<bool>,
    /// A linked fill recipe was authored, even when its DrawingML paint is not
    /// representable by the current shared fill model (for example blipFill or
    /// grpFill). Consumers must not replace it with an automatic color.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_paint_authored: Option<bool>,
    /// The linked Chart Style selected the `NoStyle` fill recipe rather than
    /// an authored `<a:noFill>`. Semantic chart marks may supply their default
    /// fill in this case; an explicit no-fill must remain transparent.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_no_style: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_colors: Option<Vec<Option<String>>>,
    /// Per-color-style-index DrawingML outline recipes after `phClr`
    /// substitution. Solid paints remain duplicated in `line_colors` for wire
    /// compatibility; gradient and pattern outlines are retained here.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_paints: Option<Vec<Option<ChartStyleFill>>>,
    /// A linked outline recipe was authored even when its paint is not
    /// representable by the current shared fill model.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_paint_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_dash: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_dash_authored: Option<bool>,
    /// DrawingML `<a:custDash>` atoms. Presence (including an empty list) is
    /// distinct from omission and overrides any inherited preset dash.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_custom_dash: Option<Vec<ChartLineDashSegment>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_cap: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_join: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_compound: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_hidden: Option<bool>,
    /// The linked Chart Style selected the `NoStyle` line recipe rather than
    /// an authored `<a:noFill>`. Semantic chart marks may supply their default
    /// outline in this case; an explicit no-fill must remain suppressed.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_no_style: Option<bool>,
    /// Per-color-style-index DrawingML effects after `phClr` substitution.
    /// Effects share the Chart Colors index with the role's fill and line.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shadows: Option<Vec<Option<crate::effect::Shadow>>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inner_shadows: Option<Vec<Option<crate::effect::Shadow>>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub glows: Option<Vec<Option<crate::effect::Glow>>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub soft_edges: Option<Vec<Option<crate::effect::SoftEdge>>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub reflections: Option<Vec<Option<crate::effect::Reflection>>>,
    /// A concrete local or referenced effect component was authored. An empty
    /// effect list is an explicit clear and therefore remains authored.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub effect_authored: Option<bool>,
    /// The linked style selected `effectRef idx=0`; this sentinel alone falls
    /// through to the lower-precedence numeric classic-style effect.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub effect_no_style: Option<bool>,
    /// Some concrete effect recipe could not be represented or resolved.
    /// Consumers must fail closed rather than inventing a lower-layer effect.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub effect_unsupported: Option<bool>,
    /// Fixed zero-based CT_ColorStyle index from `<cs:styleClr val>`. `None`
    /// means `auto`, so the renderer uses the relative object index.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_color_index: Option<usize>,
    /// Source formatting indexes corresponding to compact fill palette slots.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_formatting_indices: Option<Vec<usize>>,
    /// Source formatting indexes where a bounded numeric style intentionally
    /// delegates to the chart mark's semantic automatic paint. This differs
    /// from an authored-but-unresolved palette slot, which must fail closed.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_semantic_fallback_indices: Option<Vec<usize>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_color_index: Option<usize>,
    /// Source formatting indexes corresponding to compact line palette slots.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_formatting_indices: Option<Vec<usize>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_semantic_fallback_indices: Option<Vec<usize>>,
    /// Source formatting indexes corresponding to compact effect palette slots.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub effect_formatting_indices: Option<Vec<usize>>,
    /// Fixed zero-based CT_ColorStyle index for an effect styleClr reference.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub effect_color_index: Option<usize>,
}

/// DrawingML fill recipe retained from a Chart Style role. Chart Style parts
/// use the same CT_GradientFillProperties and CT_PatternFillProperties grammar
/// as shapes, so this wire shape intentionally mirrors core's shared `Fill`
/// discriminated union instead of introducing chart-specific paint semantics.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(tag = "fillType", rename_all = "camelCase")]
pub enum ChartStyleFill {
    #[serde(rename = "solid")]
    Solid { color: String },
    #[serde(rename = "gradient")]
    Gradient {
        stops: Vec<crate::fill::GradStop>,
        angle: f64,
        #[serde(rename = "gradType")]
        grad_type: String,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        scaled: Option<bool>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        path: Option<String>,
        #[serde(
            rename = "fillToRect",
            default,
            skip_serializing_if = "Option::is_none"
        )]
        fill_to_rect: Option<crate::fill::FillRect>,
        #[serde(rename = "tileRect", default, skip_serializing_if = "Option::is_none")]
        tile_rect: Option<crate::fill::FillRect>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        flip: Option<String>,
        #[serde(
            rename = "rotWithShape",
            default,
            skip_serializing_if = "Option::is_none"
        )]
        rot_with_shape: Option<bool>,
    },
    #[serde(rename = "pattern")]
    Pattern {
        fg: String,
        bg: String,
        preset: String,
    },
    /// DrawingML picture fill retained by relationship path. Image bytes stay
    /// in the owning OPC package and are fetched/decoded lazily by the host.
    #[serde(rename = "image", rename_all = "camelCase")]
    Image {
        image_path: String,
        mime_type: String,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        svg_image_path: Option<String>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        dpi: Option<u32>,
        #[serde(
            rename = "rotWithShape",
            default,
            skip_serializing_if = "Option::is_none"
        )]
        rot_with_shape: Option<bool>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        src_rect: Option<crate::blip::SrcRect>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        fill_rect: Option<crate::fill::FillRect>,
        /// Presence of the normative `<a:stretch>` fill-mode choice. Kept
        /// separately because an empty stretch has no fillRect but is distinct
        /// from an omitted, default-less EG_FillModeProperties choice.
        #[serde(default, skip_serializing_if = "std::ops::Not::not")]
        stretch: bool,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        tile: Option<crate::fill::TileInfo>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        alpha: Option<f64>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        duotone: Option<crate::blip::Duotone>,
    },
}

/// One authored `<c:surfaceChart|surface3DChart><c:bandFmts><c:bandFmt>`.
/// Bands are indexed low-to-high (§21.2.2.14); paint uses the shared DrawingML
/// fill model so pattern/gradient behavior stays host-independent.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartSurfaceBandFormat {
    pub idx: u32,
    /// Direct outline geometry/paint and fill-authorship provenance. The
    /// sibling `fill` field is the one authoritative direct fill recipe, so a
    /// bounded gradient is neither expanded nor serialized twice.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<ChartStyleFill>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_hidden: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_hidden: Option<bool>,
}

/// Built-in `c:style` dataPoint3D roles resolved in the semantic Surface-band
/// formatting-index domain. Pattern palettes are index-stable and use one
/// complete role; Table 5 Fade palettes retain one bounded role per final
/// 1..48 band count because their endpoint depends on the domain maximum.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Default)]
#[serde(rename_all = "camelCase")]
pub struct ChartClassicSurfaceBandStyles {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fixed: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub by_band_count: Option<Vec<ChartExElementStyle>>,
}

/// `<c:plotArea><c:dTable>` (`CT_DTable`) for classic DrawingML charts.
/// It is common chart content shared by DOCX/XLSX/PPTX hosts.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartDataTable {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<ChartExElementStyle>,
    pub show_horizontal_border: bool,
    pub show_vertical_border: bool,
    pub show_outline: bool,
    pub show_keys: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_size_hpt: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_face: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_paint_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_hidden: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_italic: Option<bool>,
    /// Resolved solid compatibility projection of `<c:dTable><c:spPr>`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_color: Option<String>,
    /// Direct DrawingML fill recipe. It is preserved even when the Office-
    /// defined visual extent for this data-table paint is not yet rendered.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<ChartStyleFill>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_hidden: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_paint_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_dash: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_hidden: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_paint_authored: Option<bool>,
}

/// Source-order metadata for one direct classic chart-group child of
/// `<c:plotArea>`. The range indexes the single flattened `series` vector.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartPlotGroup {
    pub kind: String,
    pub series_start: usize,
    pub series_count: usize,
    pub category_axis: String,
    pub value_axis: String,
    pub series_axis: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub axis_ids: Option<Vec<String>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub grouping: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bar_direction: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub scatter_style: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub radar_style: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vary_colors: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub gap_width: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub overlap: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bubble_scale: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bubble_size_represents: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_negative_bubbles: Option<bool>,
}

/// Closed host-specific policies for classic cartesian automatic layout.
///
/// ECMA-376 specifies the manual-layout coordinate model but leaves automatic
/// placement to the consumer. A format adapter may select one of these bounded
/// Office-observed policies; arbitrary numeric layout tuning is deliberately
/// not part of the wire model.
#[derive(Serialize, Deserialize, Debug, Clone, Copy, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub enum ChartCartesianAutoLayoutProfile {
    WordClassicColumn,
}

/// Mirror of TS `ChartModel`. Built by each parser and emitted as the single
/// `chart` object consumed by the core chart renderer. `Default` is an empty
/// starting point for binary-source projections; the XML parser still fills
/// every field explicitly.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Default)]
#[serde(rename_all = "camelCase")]
pub struct ChartModel {
    // ── Required (always serialized) ────────────────────────────────────────
    pub chart_type: String,
    pub title: Option<String>,
    /// Effective DrawingML runs from a legacy chart title. Keeping the run
    /// boundaries preserves authored line breaks and per-run typography while
    /// `title` remains the plain-text compatibility field.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_rich_runs: Option<Vec<ChartTextRun>>,
    /// A direct `<c:title>` / `<cx:title>` exists even when its text is empty.
    /// Empty title placeholders still reserve their authored layout band.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub title_present: bool,
    /// The source chart declares no series at all: every classic CT_PlotArea
    /// chart group is empty (ECMA-376 §21.2.2.145 allows `ser` 0..n), or a
    /// BIFF chart has no Series record. Excel draws such a chart as its empty
    /// chart area (its PDF export of a series-less chart sheet shows only the
    /// chart-area border), unlike a chart whose series could not be read.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub authored_without_series: bool,
    pub categories: Vec<String>,
    /// Host-resolved visibility of the shared category reference. Authored
    /// chart caches remain authoritative for text/value content.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub category_source_hidden: Option<Vec<bool>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub category_levels: Option<Vec<Vec<String>>>,
    pub series: Vec<ChartSeries>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot_groups: Option<Vec<ChartPlotGroup>>,
    pub show_data_labels: bool,
    pub val_min: Option<f64>,
    pub val_max: Option<f64>,
    pub cat_axis_title: Option<String>,
    pub val_axis_title: Option<String>,
    pub cat_axis_hidden: bool,
    pub val_axis_hidden: bool,
    pub cat_axis_line_hidden: bool,
    pub val_axis_line_hidden: bool,
    pub plot_area_bg: Option<String>,
    /// Structured `<c:plotArea><c:spPr>` fill. Solid fills are also mirrored
    /// in `plot_area_bg` for wire compatibility.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot_area_fill: Option<ChartStyleFill>,
    /// Explicit plot-area `noFill`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot_area_fill_hidden: Option<bool>,
    /// A direct plot-area fill paint was authored, even when unresolved.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot_area_fill_paint_authored: Option<bool>,
    /// The resolved plot-area color is a host automatic fallback rather than
    /// direct formatting. Linked chart-style paint therefore has precedence.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot_area_fill_automatic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot_area_line_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot_area_line_fill: Option<ChartStyleFill>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot_area_line_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot_area_line_dash: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot_area_line_dash_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot_area_line_custom_dash: Option<Vec<ChartLineDashSegment>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot_area_line_cap: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot_area_line_join: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot_area_line_compound: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot_area_line_hidden: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot_area_line_paint_authored: Option<bool>,
    pub chart_bg: Option<String>,
    /// Structured non-solid chart-space fill. Solid fills remain represented
    /// by `chart_bg` for wire compatibility.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_fill: Option<ChartStyleFill>,
    /// Explicit chart-area `noFill`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_fill_hidden: Option<bool>,
    /// A direct chart-area fill paint was authored, even when unresolved.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_fill_paint_authored: Option<bool>,
    /// `<c:chartSpace><c:roundedCorners>` (§21.2.2.159). A bare CT_Boolean is
    /// true; omission is preserved as `None`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rounded_corners: Option<bool>,
    /// `<c:chart><c:plotVisOnly>` (§21.2.2.146). A bare CT_Boolean is true;
    /// omission remains `None` because the element itself has no schema default.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot_visible_only: Option<bool>,
    pub show_legend: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_table: Option<ChartDataTable>,
    pub legend_pos: Option<String>,
    pub cat_axis_cross_between: String,
    pub val_axis_major_tick_mark: String,
    pub cat_axis_major_tick_mark: String,
    pub title_font_size_hpt: Option<i32>,
    pub title_font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_font_paint_authored: Option<bool>,
    pub title_font_face: Option<String>,
    pub cat_axis_font_size_hpt: Option<i32>,
    pub val_axis_font_size_hpt: Option<i32>,
    pub data_label_font_size_hpt: Option<i32>,
    pub subtotal_indices: Vec<u32>,
    // ── Optional (skipped when unset) ───────────────────────────────────────
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_minor_tick_mark: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_minor_tick_mark: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_font_paint_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_font_paint_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_manual_layout: Option<LegendManualLayout>,
    /// `<c:legend><c:overlay>` (§21.2.2.132). Absent is preserved as `None`;
    /// a present element without `val` is true per CT_Boolean.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_overlay: Option<bool>,
    /// Source-ordered indexed `<c:legendEntry>` overrides (§21.2.2.94).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_entries: Option<Vec<ChartLegendEntryOverride>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_format_code: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_number_format: Option<ChartAxisNumberFormat>,
    /// `<c:valAx><c:dispUnits>` (§21.2.2.45) scales displayed axis-associated
    /// values (ticks and Office-generated `showVal` data-label text); geometry
    /// and the underlying series values remain unscaled.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_display_units: Option<ChartDisplayUnits>,
    /// The same display-unit contract for a numeric horizontal axis (scatter /
    /// bubble store that axis in `<c:valAx>` but expose it through cat-axis
    /// fields in the shared renderer model).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_display_units: Option<ChartDisplayUnits>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bar_gap_width: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bar_overlap: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_label_position: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_label_font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_label_font_paint_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_label_format_code: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_label_font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_label_font_italic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_label_font_language: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_label_font_baseline: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_font_italic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_font_language: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_font_baseline: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_font_italic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_font_italic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_title_font_size_hpt: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_title_font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_title_font_italic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_title_font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_title_font_paint_authored: Option<bool>,
    /// Authored `<c:catAx><c:title>` `bodyPr@rot` in raw `ST_Angle` units.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_title_rotation: Option<i32>,
    /// Authored `<c:catAx><c:title>` `bodyPr@vert` mode.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_title_vertical_mode: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_title_manual_layout: Option<ChartManualLayout>,
    /// Effective sum of the DrawingML top/bottom text insets for the
    /// category-axis title, including CT_TextBodyProperties defaults.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_title_text_vertical_inset_emu: Option<i64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_title_font_size_hpt: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_title_font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_title_font_italic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_title_font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_title_font_paint_authored: Option<bool>,
    /// Authored `<c:valAx><c:title>` `bodyPr@rot` in raw `ST_Angle` units.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_title_rotation: Option<i32>,
    /// Authored `<c:valAx><c:title>` `bodyPr@vert` mode.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_title_vertical_mode: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_title_manual_layout: Option<ChartManualLayout>,
    /// Effective sum of the DrawingML top/bottom text insets for the
    /// value-axis title.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_title_text_vertical_inset_emu: Option<i64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_border_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_border_line_fill: Option<ChartStyleFill>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_border_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_border_dash: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_border_dash_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_border_custom_dash: Option<Vec<ChartLineDashSegment>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_border_cap: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_border_join: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_border_compound: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_border_hidden: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_border_paint_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_crosses: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_crosses_at: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_crosses: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_crosses_at: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_line_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_line_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_line_dash: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_line_paint_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_line_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_line_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_line_dash: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_line_paint_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_format_code: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_number_format: Option<ChartAxisNumberFormat>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_min: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_max: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_manual_layout: Option<ChartManualLayout>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot_area_manual_layout: Option<ChartManualLayout>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cartesian_auto_layout_profile: Option<ChartCartesianAutoLayoutProfile>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub scatter_style: Option<String>,
    /// `<c:bubbleChart><c:bubbleScale val>` (§21.2.2.21) — bubble diameter
    /// scale as a percentage of the renderer's default bubble size (0–300).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bubble_scale: Option<u32>,
    /// `<c:bubbleChart><c:sizeRepresents val>` (§21.2.2.193,
    /// ST_SizeRepresents §21.2.3.43). Absent means the schema default `area`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bubble_size_represents: Option<String>,
    /// `<c:bubbleChart><c:showNegBubbles val>` (§21.2.2.179). Absent defaults
    /// to false; a bare CT_Boolean element implies true.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_negative_bubbles: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub radar_style: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub secondary_val_axis: Option<SecondaryValueAxis>,
    /// Numeric horizontal axis referenced by a scatter/bubble group overlaid
    /// on a non-scatter primary chart. OOXML represents both scatter axes as
    /// `<c:valAx>`; this keeps the second horizontal axis distinct from the
    /// primary bar/column value axis.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub secondary_cat_axis: Option<SecondaryValueAxis>,
    // ── Pie / doughnut geometry (CH8) ───────────────────────────────────────
    /// `<c:doughnutChart><c:holeSize val>` (§21.2.2.82, `ST_HoleSizePercent`
    /// §21.2.3.55) — hole diameter as 1–90% of the outer diameter. `None` when
    /// absent; the renderer defaults an absent doughnut hole to 50%.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hole_size: Option<u32>,
    /// `<c:pieChart | doughnutChart><c:firstSliceAng val>` (§21.2.2.52,
    /// `ST_FirstSliceAng` §21.2.3.15) — start angle 0–360° clockwise from 12
    /// o'clock. `None` = 0 (byte-stable).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_slice_angle: Option<u32>,
    // ── Chart text font faces (CH10) ────────────────────────────────────────
    /// `<c:catAx><c:txPr>…<a:latin typeface>` tick-label font.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_font_face: Option<String>,
    /// `<c:valAx><c:txPr>…<a:latin typeface>` tick-label font.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_font_face: Option<String>,
    /// `<c:catAx><c:title>…<a:latin typeface>` axis-title font.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_title_font_face: Option<String>,
    /// `<c:valAx><c:title>…<a:latin typeface>` axis-title font.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_title_font_face: Option<String>,
    /// `<c:dLbls><c:txPr>…<a:latin typeface>` data-label font.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_label_font_face: Option<String>,
    /// `<c:legend><c:txPr>…<a:latin typeface>` legend font.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_font_face: Option<String>,
    /// `<c:legend><c:txPr>…<a:solidFill>` legend text color (hex, no `#`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_font_paint_authored: Option<bool>,
    /// `<c:legend><c:txPr>` legend font size (OOXML hundredths of a point).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_font_size_hpt: Option<i32>,
    /// `<c:legend><c:txPr>…defRPr@b` legend bold flag.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_font_italic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_font_language: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_font_baseline: Option<f64>,
    /// `<c:legend><c:spPr>` explicit frame fill (hex, no `#`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_fill_color: Option<String>,
    /// Structured `<c:legend><c:spPr>` fill. Solid fills are also mirrored in
    /// `legend_fill_color` for wire compatibility.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_fill: Option<ChartStyleFill>,
    /// Explicit `<c:legend><c:spPr><a:noFill/>`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_fill_hidden: Option<bool>,
    /// A direct legend fill paint was authored, even when unresolved.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_fill_paint_authored: Option<bool>,
    /// `<c:legend><c:spPr><a:ln>` explicit frame stroke (hex, no `#`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_line_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_line_fill: Option<ChartStyleFill>,
    /// `<c:legend><c:spPr><a:ln@w>` frame stroke width in EMU.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_line_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_line_dash: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_line_dash_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_line_custom_dash: Option<Vec<ChartLineDashSegment>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_line_cap: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_line_join: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_line_compound: Option<String>,
    /// Explicit `<c:legend><c:spPr><a:ln><a:noFill/>`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_line_hidden: Option<bool>,
    /// A direct legend line paint was authored, even when unresolved.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_line_paint_authored: Option<bool>,
    /// Theme heading (majorFont) Latin face — fallback for chart title / axis
    /// titles when their `<c:txPr>` supplies no `<a:latin>`. `None` when the
    /// theme is not threaded (renderer keeps sans-serif; byte-stable).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_major_font_latin: Option<String>,
    /// Theme body (minorFont) Latin face — fallback for tick labels / data
    /// labels / legend. `None` when the theme is not threaded.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_minor_font_latin: Option<String>,
    /// `<c:date1904>` (ECMA-376 §21.2.2.38). `true` = the chart's serial dates
    /// resolve against the 1904 date system. Omitted from JSON when false (the
    /// default 1900 system) for wire parity.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub date1904: bool,
    /// `<c:chart><c:dispBlanksAs val>` (ECMA-376 §21.2.2.42) — how blank cells
    /// are plotted on line/area charts ("gap" | "zero" | "span"). `None` when
    /// the element is absent (the renderer defaults to "gap"); only serialized
    /// when the file sets it.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub disp_blanks_as: Option<String>,
    /// `<c:chart><c:showDLblsOverMax>` (ECMA-376 §21.2.2.180). Absent and
    /// explicit false both suppress labels above the effective axis maximum;
    /// preserving `Option` retains the authored chart-level boolean on wire.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_data_labels_over_max: Option<bool>,
    // ── Axis scale model (CH6) ──────────────────────────────────────────────
    /// `<c:valAx><c:majorGridlines>` presence (§21.2.2.100). `Some(false)` when
    /// the value axis exists but omits the element — Office suppresses the value
    /// gridlines then. `None` when there is no value axis (or the parser path
    /// doesn't model it); the renderer keeps its historical always-on value
    /// gridlines, so a `None`/absent field is byte-stable.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_major_gridlines: Option<bool>,
    /// `<c:catAx><c:majorGridlines>` presence (§21.2.2.100). `Some(true)` turns
    /// on category-axis gridlines (Office omits them by default). `None`/absent
    /// keeps the renderer's historical no-category-gridlines behavior.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_major_gridlines: Option<bool>,
    /// `<c:valAx><c:majorGridlines><c:spPr><a:ln><a:solidFill>` resolved gridline
    /// colour (hex, no `#`) — §21.2.2.100. `None` when the value axis omits the
    /// element or gives it no explicit colour; the renderer then keeps its faint
    /// default gridline (byte-stable).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_gridline_color: Option<String>,
    /// `<c:valAx><c:majorGridlines><c:spPr><a:ln w>` gridline width in EMU.
    /// `None` = the renderer's default hairline (byte-stable).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_gridline_width_emu: Option<u32>,
    /// `<c:valAx><c:majorGridlines>...<a:prstDash val>`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_gridline_dash: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_gridline_paint_authored: Option<bool>,
    /// `<c:catAx><c:majorGridlines><c:spPr><a:ln><a:solidFill>` resolved gridline
    /// colour (hex, no `#`). Only meaningful when `cat_axis_major_gridlines` is
    /// on. `None` keeps the faint default.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_gridline_color: Option<String>,
    /// `<c:catAx><c:majorGridlines><c:spPr><a:ln w>` gridline width in EMU.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_gridline_width_emu: Option<u32>,
    /// `<c:catAx><c:majorGridlines>...<a:prstDash val>`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_gridline_dash: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_gridline_paint_authored: Option<bool>,
    /// `<c:valAx><c:minorGridlines>` presence (§21.2.2.109).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_minor_gridlines: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_minor_gridline_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_minor_gridline_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_minor_gridline_dash: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_minor_gridline_paint_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_minor_gridlines: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_minor_gridline_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_minor_gridline_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_minor_gridline_dash: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_minor_gridline_paint_authored: Option<bool>,
    /// `<c:valAx><c:majorUnit val>` (§21.2.2.103) — explicit major gridline
    /// step, overriding the auto "nice" step. `None` = auto (byte-stable).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_major_unit: Option<f64>,
    /// `<c:valAx><c:minorUnit val>` (§21.2.2.112) — explicit minor gridline step.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_minor_unit: Option<f64>,
    /// Numeric horizontal-axis units for scatter/bubble, whose X axis is a
    /// second `<c:valAx>` even though it occupies the category-axis slot.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_major_unit: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_minor_unit: Option<f64>,
    /// The category axis is `<c:dateAx>` rather than `<c:catAx>`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_is_date: Option<bool>,
    /// `<c:dateAx><c:baseTimeUnit val>` (`days` when omitted).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_base_time_unit: Option<String>,
    /// `<c:dateAx><c:majorTimeUnit val>` (`days` when omitted).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_major_time_unit: Option<String>,
    /// `<c:dateAx><c:minorTimeUnit val>` (`days` when omitted).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_minor_time_unit: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_no_multi_level_labels: Option<bool>,
    /// `<c:valAx><c:scaling><c:logBase val>` (§21.2.2.98) — logarithmic value
    /// axis base (>= 2). `None` = linear (byte-stable).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_log_base: Option<f64>,
    /// Numeric horizontal-axis log base for scatter/bubble's second valAx.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_log_base: Option<f64>,
    /// `<c:valAx><c:scaling><c:orientation val>` (§21.2.2.130) — `"minMax"`
    /// (normal) | `"maxMin"` (reversed). `None`/`"minMax"` = normal (byte-stable).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_orientation: Option<String>,
    /// `<c:catAx><c:scaling><c:orientation val>` — reverses the category axis
    /// left↔right when `"maxMin"`. `None`/`"minMax"` = normal.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_orientation: Option<String>,
    /// `<c:catAx><c:tickLblPos val>` (§21.2.2.207) — `"nextTo"` (default) |
    /// `"low"` | `"high"` | `"none"` (labels hidden). `None` = nextTo.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_tick_label_pos: Option<String>,
    /// `<c:catAx><c:tickLblSkip val>` (§21.2.2.205), 1-based interval.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_tick_label_skip: Option<u32>,
    /// `<c:catAx><c:tickMarkSkip val>` (§21.2.2.206), 1-based interval.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_tick_mark_skip: Option<u32>,
    /// `<c:catAx><c:lblAlgn val>` (§21.2.2.90) — `"l"`, `"ctr"`, or `"r"`.
    /// `None` preserves omission so the renderer can apply its centered default.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_label_alignment: Option<String>,
    /// `<c:catAx|dateAx><c:lblOffset val>` (§21.2.2.91), normalized to the
    /// schema's 0–1000 percentage value. Strict serializes the value with `%`;
    /// Transitional also permits the legacy unsigned-integer lexical form.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_label_offset_percent: Option<u32>,
    /// `<c:valAx><c:tickLblPos val>` (§21.2.2.207).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_tick_label_pos: Option<String>,
    /// `<c:catAx><c:txPr><a:bodyPr rot>` (60000ths of a degree) — category
    /// tick-label rotation. `None`/0 = horizontal (byte-stable).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_label_rotation: Option<i32>,
    /// Group-owned decorations for each classic `<c:lineChart>` /
    /// `<c:line3DChart>` in plot-area document order. Keeping the owner group
    /// prevents combo charts or multiple line groups from sharing decorations.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_group_decorations: Option<Vec<ChartLineGroupDecorations>>,
    /// Group-owned drop lines for each classic `<c:areaChart>` /
    /// `<c:area3DChart>` in plot-area document order.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub area_group_decorations: Option<Vec<ChartAreaGroupDecorations>>,
    /// Group-owned series lines for each classic `<c:barChart>` /
    /// `<c:bar3DChart>` in plot-area document order.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bar_group_decorations: Option<Vec<ChartBarGroupDecorations>>,
    // ── Stock chart (CH13, §21.2.2.198) ──────────────────────────────────────
    /// `<c:stockChart><c:dropLines>` direct DrawingML line paint. Object
    /// presence means the stock chart carries drop-line geometry.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub stock_drop_lines: Option<ChartDecorationLineStyle>,
    /// `<c:stockChart><c:hiLowLines>` direct DrawingML line paint.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub stock_hi_low_line_style: Option<ChartDecorationLineStyle>,
    /// `<c:stockChart><c:hiLowLines>` (§21.2.2.80) presence. When `Some(true)`
    /// the stock renderer draws a vertical line spanning each category's
    /// low↔high value. Only emitted for a stock chart (`chart_type == "stock"`);
    /// `None` on every other chart type keeps the wire byte-stable.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub stock_hi_low_lines: Option<bool>,
    /// `<c:hiLowLines><c:spPr><a:ln><a:solidFill>` resolved color (hex, no `#`).
    /// `None` leaves direct paint omitted. Only meaningful with
    /// `stock_hi_low_lines == Some(true)`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub stock_hi_low_line_color: Option<String>,
    /// `<c:stockChart><c:upDownBars>` (§21.2.2.218) presence. Parsed so a file
    /// that carries open-close up/down bars draws them between Open and Close.
    /// `None` when absent.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub stock_up_down_bars: Option<bool>,
    /// Parsed `<c:upDownBars>` gap and direct up/down bar paint.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub stock_up_down_bar_style: Option<ChartStockUpDownBarStyle>,
    /// Bounded Office automatic paint for empty stock decorations. Kept
    /// separate from direct/linked paint so renderer precedence is explicit.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub stock_automatic_style: Option<ChartStockAutomaticStyle>,
    /// `<c:surfaceChart|surface3DChart><c:wireframe>` effective boolean.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub surface_wireframe: Option<bool>,
    /// Authored low-to-high surface band formatting (§21.2.2.13/14).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub surface_band_formats: Option<Vec<ChartSurfaceBandFormat>>,
    /// Numeric style materialized in the Surface value-band domain. Kept
    /// separate from source-series roles so band indexes cannot alias series.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub classic_surface_band_styles: Option<ChartClassicSurfaceBandStyles>,
    /// Legacy `<c:style@val>` (1..48), used by automatic classic-chart
    /// palettes independently of ChartEx sidecars.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legacy_chart_style: Option<u8>,
    /// Resolved theme accent1..6 palette for renderer-generated classic-chart
    /// objects such as automatic surface value bands.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_accent_colors: Option<Vec<String>>,
    /// Pie-of-pie / bar-of-pie secondary-plot contract (§21.2.2.126).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub of_pie: Option<ChartOfPie>,
    /// Authored 3D chart-space view and group depth controls.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub three_d: Option<ChartThreeD>,
    // ── chartEx structured layouts (CH15, MS 2014 chartex ext) ───────────────
    /// Structured box-and-whisker data (`chart_type == "boxWhisker"`). `None`
    /// for every other chart type — the field is populated ONLY by
    /// `parse_chartex_part` when the series `layoutId` is `boxWhisker`, so the
    /// flat `categories`/`series` model (which waterfall/treemap consume) is
    /// unchanged and the wire stays byte-stable for those.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_box: Option<ChartexBoxWhisker>,
    /// Structured sunburst hierarchy (`chart_type == "sunburst"`). `None`
    /// otherwise (byte-stable for the flat-model chartEx charts).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_sunburst: Option<ChartexSunburst>,
    /// Structured treemap hierarchy (`chart_type == "treemap"`) and its
    /// parent-label layout. `None` otherwise.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_treemap: Option<ChartexTreemap>,
    /// Structured geospatial rows and authored Region Map presentation
    /// (`chart_type == "regionMap"`). Geometry resolution deliberately remains
    /// a renderer concern; the parser preserves source labels/entity ids.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_region_map: Option<ChartexRegionMap>,
    /// ChartEx histogram `CT_Binning` controls. Raw observations remain in
    /// `series[0].values` so the renderer can derive a bounded frequency plan.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_histogram_binning: Option<ChartexHistogramBinning>,
    /// Theme accent palette (`accent1..6` resolved to hex, no `#`) for chartEx
    /// charts that color by branch/series index (boxWhisker series and
    /// sunburst/treemap branches). `None` when the resolver supplies no default palette (pptx);
    /// the renderer then falls back to its own `CHART_PALETTE`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_accents: Option<Vec<String>>,
    /// Total color set from the linked Chart Colors part: contained colors
    /// repeated for every authored variation (MS-ODRAWXML §2.8.3.2).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_color_palette: Option<Vec<Option<String>>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_color_style_method: Option<String>,
    /// Effective paint recipes for every paint-bearing CT_ChartStyle role
    /// (MS-ODRAWXML §2.8.3.1). Direct formatting remains in its authored model
    /// fields; the renderer consults this table only as a linked fallback.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_style_roles: Option<BTreeMap<String, ChartExElementStyle>>,
    /// ECMA-376 §21.2.3.46 built-in style defaults for classic `c:` charts.
    /// Kept separate from Office 2013+ linked Chart Style roles so consumers
    /// can enforce `direct > linked styleN.xml > numeric c:style` precedence.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub classic_chart_style_roles: Option<BTreeMap<String, ChartExElementStyle>>,
    /// Shared point-domain numeric roles for the most common bounded
    /// varyColors domain. Table 5 Fade depends on that domain's highest index.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub classic_varying_point_chart_style_roles: Option<BTreeMap<String, ChartExElementStyle>>,
    /// Point-domain numeric exceptions aligned with `plot_groups`. A `None`
    /// varying slot inherits the shared table; an empty map is refusal.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub classic_varying_point_chart_style_roles_by_group:
        Option<Vec<Option<BTreeMap<String, ChartExElementStyle>>>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_style_color_palette: Option<Vec<Option<String>>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_style_color_method: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_style_marker_size_pt: Option<u8>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_style_marker_symbol: Option<String>,
    /// Chart-space `txPr` inherited by every chart text object. Kept separate
    /// from element-local text and linked/numeric roles so core can apply the
    /// normative element > chart-space > style cascade without flattening
    /// authored paint ownership in the parser.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_text_style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_area_style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot_area_style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_title_style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_title_style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_major_gridline_style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_minor_gridline_style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_major_gridline_style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_minor_gridline_style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_data_point_style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_data_point_line_style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_series_line_style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_data_point_marker_style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_marker_size_pt: Option<u8>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_marker_symbol: Option<String>,
    /// `<cx:series><cx:layoutPr><cx:visibility connectorLines>`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_connector_lines: Option<bool>,
    /// §21.2.2.227 `<c:varyColors val="1"/>` on a SINGLE-series bar/column
    /// chart: color each data point (bar) from the theme/palette sequence and
    /// list one legend entry per point (like a pie). `Some(true)` only for that
    /// non-pie, single-series case the core renderer consumes. The pie family
    /// varies by point by chart semantics; `data_point_colors` is reserved for
    /// direct `<c:dPt>` formatting so it remains distinguishable from styles.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vary_colors: Option<bool>,
    /// Text boxes stored in the chart drawing reached through
    /// `<c:userShapes r:id>` (ECMA-376 chartDrawing `CT_Drawing`). Coordinates
    /// are fractions of the chart space from `<cdr:relSizeAnchor>`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_text_boxes: Option<Vec<ChartTextBox>>,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Default)]
#[serde(rename_all = "camelCase")]
pub struct ChartStockBarPaint {
    /// Direct `<c:upBars|downBars><c:spPr>` DrawingML effects. Fill/line stay
    /// in their established scalar/structured fields on this same carrier.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<ChartStyleFill>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_paint_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_hidden: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_paint_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_dash: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_cap: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_join: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_hidden: Option<bool>,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartStockUpDownBarStyle {
    pub gap_width_percent: f64,
    pub up: ChartStockBarPaint,
    pub down: ChartStockBarPaint,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartStockAutomaticStyle {
    pub line_color: String,
    pub line_width_emu: u32,
    pub up_fill_color: String,
    pub down_fill_color: String,
}

/// Direct DrawingML line paint for group-owned chart decoration geometry.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Default)]
#[serde(rename_all = "camelCase")]
pub struct ChartDecorationLineStyle {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<ChartStyleFill>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub paint_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dash: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cap: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub join: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hidden: Option<bool>,
}

/// Decorations authored on one classic line-chart group. `group_index` is the
/// zero-based document-order index among line groups, matching
/// `ChartSeries::line_group_index`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartLineGroupDecorations {
    pub group_index: u32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub drop_lines: Option<ChartDecorationLineStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hi_low_lines: Option<ChartDecorationLineStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub up_down_bars: Option<ChartStockUpDownBarStyle>,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartAreaGroupDecorations {
    pub group_index: u32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub drop_lines: Option<ChartDecorationLineStyle>,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartBarGroupDecorations {
    pub group_index: u32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub series_lines: Option<Vec<ChartDecorationLineStyle>>,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartOfPie {
    pub r#type: String,
    pub split_type: String,
    pub split_type_authored: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub split_pos: Option<f64>,
    pub split_pos_authored: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub custom_split_indices: Option<Vec<usize>>,
    pub second_pie_size_percent: f64,
    pub gap_width_percent: f64,
    pub series_lines: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub series_line_style: Option<ChartDecorationLineStyle>,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Default)]
#[serde(rename_all = "camelCase")]
pub struct ChartThreeDSeriesAxis {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_gridline_style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_gridline_style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,
    pub hidden: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub orientation: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tick_label_pos: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tick_label_skip: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tick_mark_skip: Option<u32>,
    pub major_tick_mark: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_tick_mark: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_paint_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_size_hpt: Option<i32>,
    pub font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_italic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_face: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_dash: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_paint_authored: Option<bool>,
    pub line_hidden: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_font_size_hpt: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_font_italic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_font_paint_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_font_face: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_rotation: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_vertical_mode: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_manual_layout: Option<ChartManualLayout>,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Default)]
#[serde(rename_all = "camelCase")]
pub struct ChartThreeDSurface {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_hidden: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_dash: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_hidden: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub thickness_percent: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub picture_options: Option<ChartThreeDPictureOptions>,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Default)]
#[serde(rename_all = "camelCase")]
pub struct ChartThreeDPictureOptions {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub apply_to_front: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub apply_to_sides: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub apply_to_end: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub picture_format: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub picture_format_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub picture_stack_unit: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub picture_stack_unit_authored: Option<bool>,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Default)]
#[serde(rename_all = "camelCase")]
pub struct ChartThreeD {
    /// Whether `<c:view3D>` itself is present. Child values retain their schema
    /// defaults separately so compatibility rules can distinguish omission.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub view_3d_present: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rotation_x: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rotation_x_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rotation_y: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rotation_y_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub height_percent: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub height_percent_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub depth_percent: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub depth_percent_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub perspective: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub perspective_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub right_angle_axes: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub right_angle_axes_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub gap_depth_percent: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub gap_depth_percent_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shape: Option<String>,
    /// `<c:bar3DChart><c:grouping val>` (§21.2.2.77). `standard` places
    /// series along the series/depth axis; `clustered` places them beside one
    /// another on the category axis. The canonical 2-D chart family alone
    /// cannot preserve this 3-D distinction.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bar_grouping: Option<String>,
    /// `<c:serAx>` (§21.2.2.175), used by standard 3-D bar/column charts to
    /// label and style the series/depth dimension.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub series_axis: Option<ChartThreeDSeriesAxis>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub floor: Option<ChartThreeDSurface>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub side_wall: Option<ChartThreeDSurface>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub back_wall: Option<ChartThreeDSurface>,
}

/// One formatted run in a chart-drawing text box.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartTextRun {
    pub text: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_size_hpt: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_paint_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_hidden: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_face: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub language: Option<String>,
    /// DrawingML baseline shift normalized to a fraction of the font size.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub baseline: Option<f64>,
    /// Effective `<a:pPr algn>` for the paragraph that owns this run. Keeping
    /// it on each bounded run preserves paragraph-specific alignment without
    /// introducing another nested wire representation for data labels.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub paragraph_align: Option<String>,
}

/// One DrawingML paragraph in a chart-drawing text box.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartTextParagraph {
    pub runs: Vec<ChartTextRun>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub align: Option<String>,
}

/// A text shape anchored relative to the chart space.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartTextBox {
    pub x: f64,
    pub y: f64,
    pub w: f64,
    pub h: f64,
    pub paragraphs: Vec<ChartTextParagraph>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vertical_anchor: Option<String>,
    /// `<a:bodyPr wrap>` (`ST_TextWrappingType`). `None` retains DrawingML's
    /// application-default square wrapping; only the explicit `none` value
    /// disables wrapping in the renderer.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub wrap: Option<String>,
    /// `<a:bodyPr lIns>` — left text inset in EMU. The parser resolves the
    /// ECMA-376 §21.1.2.1.1 default when the attribute is omitted.
    pub l_ins: i64,
    /// `<a:bodyPr tIns>` — top text inset in EMU.
    pub t_ins: i64,
    /// `<a:bodyPr rIns>` — right text inset in EMU.
    pub r_ins: i64,
    /// `<a:bodyPr bIns>` — bottom text inset in EMU.
    pub b_ins: i64,
}

/// Pattern-only chart-series fill descriptor matching TS `PatternFill`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartPatternFill {
    pub fill_type: String,
    pub fg: String,
    pub bg: String,
    pub preset: String,
}

/// Mirror of TS `ChartSeries`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Default)]
#[serde(rename_all = "camelCase")]
pub struct ChartSeries {
    pub name: String,
    /// Effective source formatting index: ChartEx `CT_Series@formatIdx`
    /// ([MS-ODRAWXML] 2.24.3.77), or classic `<c:ser><c:idx>`. When omitted
    /// this is the original document-order index, before series filtering.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_format_idx: Option<u32>,
    pub color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_pattern: Option<ChartPatternFill>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub invert_if_negative: Option<bool>,
    /// Application-generated classic-chart negative style. Kept distinct from
    /// authored `<c:invertIfNegative>` so the shared wire model never reports
    /// an element that was absent from the package.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub automatic_negative_style: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inverted_fill: Option<ChartStyleFill>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inverted_fill_hidden: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inverted_fill_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inverted_line_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inverted_line_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inverted_line_hidden: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inverted_line_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_width_emu: Option<u32>,
    /// Per-series `CT_BarSer/c:shape`, overriding the bar3DChart group shape.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub three_d_shape: Option<String>,
    pub values: Vec<Option<f64>>,
    /// Host-resolved visibility provenance aligned with `values`. A true slot
    /// means at least one required source cell is in a hidden row or column.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub source_hidden: Option<Vec<bool>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_point_colors: Option<Vec<Option<String>>>,
    /// `<c:pieChart|doughnutChart><c:ser><c:explosion val>` default pull-out
    /// amount. A point-level `dPt/explosion` overrides it.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub explosion: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_label_colors: Option<Vec<Option<String>>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub series_type: Option<String>,
    /// Document-order index of the owning classic line-chart group. Group
    /// decorations are resolved through `ChartModel::line_group_decorations`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_group_index: Option<u32>,
    /// Document-order index of the owning classic area-chart group. Group
    /// drop lines resolve through `ChartModel::area_group_decorations`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub area_group_index: Option<u32>,
    /// Document-order index of the owning classic bar-chart group. This keeps
    /// separately authored overlay groups distinct after the shared model
    /// flattens plot-area series.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bar_group_index: Option<u32>,
    /// Direct `<c:barDir val>` on the owning bar-chart group.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bar_group_direction: Option<String>,
    /// Direct `<c:grouping val>` on the owning bar-chart group.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bar_group_grouping: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bar_group_gap_width: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bar_group_overlap: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub use_secondary_axis: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub categories: Option<Vec<String>>,
    /// Bubble-only source provenance. A string-backed `<c:xVal>` has distinct
    /// Office legend semantics from a numeric X source.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bubble_x_source_is_string: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_marker: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_format_code: Option<String>,
    /// Number format of the series category/X source (`<c:cat|xVal>` cache).
    /// Scatter data labels with `showCatName` use this for the displayed X
    /// value (for example `0.15` authored as `15%`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_format_code: Option<String>,
    /// Built-in worksheet number-format ID for the category source. Unlike a
    /// literal cache format, ID 14 is localized by the consuming application.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_format_builtin_id: Option<u32>,
    /// Per-point category/X number formats from `<c:pt@formatCode>`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_format_codes: Option<Vec<Option<String>>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_symbol: Option<String>,
    /// Application automatic scatter marker when the series omits a direct
    /// `<c:marker><c:symbol>`. Kept separate from `marker_symbol` so source
    /// filtering can distinguish automatic style from authored formatting.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub automatic_marker_symbol: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_size: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_fill: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_fill_paint: Option<ChartStyleFill>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_fill_paint_authored: Option<bool>,
    /// Direct `<c:marker><c:spPr>` effect component.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_line: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_line_paint_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_line_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_point_overrides: Option<Vec<ChartDataPointOverride>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_label_overrides: Option<Vec<ChartDataLabelOverride>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub series_data_labels: Option<ChartSeriesDataLabels>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub err_bars: Option<Vec<ChartErrBars>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bubble_sizes: Option<Vec<Option<f64>>>,
    /// Effective default authored on the owning `<c:bubbleChart>` group.
    /// Kept per series because one plot area can contain multiple bubble groups.
    #[serde(
        rename = "bubble3DGroupDefault",
        default,
        skip_serializing_if = "Option::is_none"
    )]
    pub bubble_3d_group_default: Option<bool>,
    /// Direct `<c:bubbleChart><c:ser><c:bubble3D>` override.
    #[serde(rename = "bubble3D", default, skip_serializing_if = "Option::is_none")]
    pub bubble_3d: Option<bool>,
    /// `<c:ser><c:smooth val>` (ECMA-376 §21.2.2.194) — line/area series flag
    /// requesting a smoothed (spline) curve. `None` (omitted) = straight
    /// polyline (the default); only serialized when the file sets it.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub smooth: Option<bool>,
    /// `<c:ser><c:trendline>` per-series trendlines (§21.2.2.211). `None`/empty
    /// when the series declares none (byte-stable).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub trend_lines: Option<Vec<ChartTrendline>>,
    /// `<c:ser><c:spPr><a:ln><a:noFill/>` (§21.2.2.198 CT_ShapeProperties →
    /// DrawingML §20.1.2.2.24 CT_LineProperties). `Some(true)` when the series'
    /// connecting line is explicitly turned OFF. For a scatter/line series this
    /// overrides the chart-group `<c:scatterStyle>` (§21.2.2.42) / line default:
    /// Excel/PowerPoint draw NO connecting line when the series line is
    /// `<a:noFill/>`, even if the group style is `lineMarker`. `None` (omitted)
    /// = the series carries no explicit line-off, so the group default governs.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_hidden: Option<bool>,
}

/// Mirror of TS `ChartTrendline` — `<c:ser><c:trendline>` (§21.2.2.211).
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartTrendline {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<ChartExElementStyle>,
    /// Optional authored `<c:name>` used for the legend entry.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    /// `<c:trendlineType val>` (§21.2.2.213) — linear|exp|log|power|poly|movingAvg.
    pub trendline_type: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub order: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub period: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub forward: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub backward: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub intercept: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub disp_r_sqr: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub disp_eq: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_manual_layout: Option<ChartManualLayout>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_text: Option<String>,
    /** Bounded formatted runs from `<c:trendlineLbl><c:tx><c:rich>`. */
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_rich_runs: Option<Vec<ChartTextRun>>,
    /// `<c:trendlineLbl><c:numFmt formatCode>` — authored formatting for
    /// generated equation / R² values.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_format_code: Option<String>,
    /// `<c:trendlineLbl><c:numFmt sourceLinked>`; preserved independently from
    /// the code so callers can distinguish an authored local format.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_format_source_linked: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_font_size_hpt: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_font_italic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_font_paint_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_font_hidden: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_font_face: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_font_language: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_font_baseline: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_text_rotation: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_text_wrap: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_text_vertical_anchor: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_text_vertical_mode: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_text_l_ins_emu: Option<i64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_text_t_ins_emu: Option<i64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_text_r_ins_emu: Option<i64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_text_b_ins_emu: Option<i64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_text_body_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_box: Option<ChartLabelBox>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_text_align: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_dash: Option<String>,
    /// `<c:spPr><a:ln><a:noFill/>` — the trendline stroke is explicitly absent.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_hidden: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_paint_authored: Option<bool>,
}

/// Mirror of TS `ChartDataPointOverride`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartDataPointOverride {
    pub idx: u32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_hidden: Option<bool>,
    /// Direct classic `<c:dPt><c:spPr>` shape paint. All classic families use
    /// this bounded carrier for structured and unresolved paint precedence;
    /// scalar fields remain compatibility projections.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_dash: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_hidden: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_symbol: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_size: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_fill: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_fill_paint: Option<ChartStyleFill>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_fill_paint_authored: Option<bool>,
    /// Direct `<c:dPt><c:marker><c:spPr>` effect component.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_line: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_line_paint_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_line_width_emu: Option<u32>,
    /// Direct `<c:dPt><c:bubble3D>` override. CT_DPt is shared across classic
    /// chart families; only the bubble renderer consumes this effect.
    #[serde(rename = "bubble3D", default, skip_serializing_if = "Option::is_none")]
    pub bubble_3d: Option<bool>,
    /// `<c:dPt><c:explosion val>` (§21.2.2.61) — pie/doughnut slice pull-out
    /// amount. The schema type is `CT_UnsignedInt` (unbounded `xsd:unsignedInt`);
    /// the spec text itself doesn't define a 0–100 range or "percentage" unit,
    /// only "the amount the data point shall be moved from the center of the
    /// pie". Renderers interpret it as a de-facto percentage of the outer
    /// radius (0–100 typical), matching Office's Point Explosion UI slider
    /// rather than a spec-mandated bound. `None`/absent = 0 (byte-stable).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub explosion: Option<u32>,
}

/// Mirror of TS `ChartDataLabelOverride`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartDataLabelOverride {
    pub idx: u32,
    pub text: String,
    /// Effective DrawingML runs from a custom `<c:dLbl><c:tx><c:rich>` body.
    /// Paragraph breaks are stored as `\n` runs. The parser bounds the payload
    /// to 4096 Unicode scalars and four lines before it crosses the WASM wire.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rich_runs: Option<Vec<ChartTextRun>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub position: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_paint_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_hidden: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_size_hpt: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_face: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_italic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_language: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_baseline: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_rotation: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_wrap: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_vertical_anchor: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_vertical_mode: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_l_ins_emu: Option<i64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_t_ins_emu: Option<i64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_r_ins_emu: Option<i64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_b_ins_emu: Option<i64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_body_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_align: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub format_code: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub separator: Option<String>,
    /// Per-point `<c:dLbl><c:layout><c:manualLayout>` (§21.2.2.47/§21.2.2.88).
    /// The application chooses automatic label geometry when this is absent;
    /// when authored, preserve it so the shared renderer can resolve it against
    /// the same bounded chart rectangle as the automatic anchor.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub manual_layout: Option<ChartManualLayout>,
    /// Per-point label callout box style (`<c:dLbl>` §21.2.2.47 `<c:spPr>`
    /// §21.2.2.197): background fill / border, mirroring the series-level
    /// defaults. Present only when the point's `<c:spPr>` overrides the shape
    /// (e.g. a differently tinted callout for one slice). See [`ChartLabelBox`].
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_box: Option<ChartLabelBox>,
    /// Per-point label-content flags (`<c:dLbl>` §21.2.2.47 carries the full
    /// `CT_DLbl` show-flag group: §21.2.2.189 `<c:showVal>`, §21.2.2.177
    /// `<c:showCatName>`, §21.2.2.180 `<c:showSerName>`, §21.2.2.187
    /// `<c:showPercent>`). When a `<c:dLbl>` sets these they OVERRIDE the
    /// series-level `<c:dLbls>` defaults (§21.2.2.49) for that one point. A
    /// point can therefore disable `showCatName` and enable `showPercent` even
    /// when the series does the reverse. `None` means the point declared no
    /// such flag, so the series default
    /// governs (byte-stable for points that carry none).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_val: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_cat_name: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_ser_name: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_percent: Option<bool>,
    /// `<c:dLbl><c:showBubbleSize>`; absent inherits the series default.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_bubble_size: Option<bool>,
    /// `<c:dLbl><c:showLegendKey>`; absent inherits the series default.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_legend_key: Option<bool>,
    /// `<c:dLbl><c:delete val="1"/>` (§21.2.2.43) — this point's label is
    /// removed. Distinguishes a genuinely deleted label from a `<c:dLbl>` that
    /// merely carries style/flag overrides with no `<c:tx>` (both formerly
    /// collapsed to `text == ""`). `Some(true)` = skip the label entirely;
    /// `None`/absent = not deleted (compose from flags / text). Byte-stable for
    /// points that carry no `<c:delete>`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub deleted: Option<bool>,
}

/// Callout-box style for a pie/doughnut data label — the white (or themed)
/// rounded rectangle with a thin border that Word draws around a `bestFit`
/// label placed outside its slice. Parsed from the label's `<c:spPr>`
/// (§21.2.2.197, the shape properties of a `<c:dLbl>` §21.2.2.47 /
/// `<c:dLbls>` §21.2.2.49): the direct `<a:solidFill>` is the box fill and the
/// `<a:ln>` is its border.
///
/// A `None` on `ChartSeriesDataLabels::label_box` means the file wrote no box
/// shape, so the renderer keeps the historical plain-text label (no callout).
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Default)]
#[serde(rename_all = "camelCase")]
pub struct ChartLabelBox {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<ChartExElementStyle>,
    /// `<c:spPr><a:solidFill>` resolved hex (no `#`). The box background;
    /// `<a:noFill>`/absent leaves this `None` (transparent box).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_paint: Option<ChartStyleFill>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_hidden: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_paint_authored: Option<bool>,
    /// `<c:spPr><a:ln><a:solidFill>` resolved hex (no `#`) — border stroke.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border_fill: Option<ChartStyleFill>,
    /// `<c:spPr><a:ln w>` border width in EMU (12700 EMU = 1 pt).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border_hidden: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border_paint_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border_dash: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border_dash_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border_custom_dash: Option<Vec<ChartLineDashSegment>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border_cap: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border_join: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border_compound: Option<String>,
}

/// Mirror of TS `ChartSeriesDataLabels`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Default)]
#[serde(rename_all = "camelCase")]
pub struct ChartSeriesDataLabels {
    /// Series-level `<c:dLbls><c:delete>` suppresses the collection. A
    /// per-point `<c:dLbl><c:delete val="0">` may override it for that point.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub deleted: Option<bool>,
    pub show_val: bool,
    pub show_cat_name: bool,
    pub show_ser_name: bool,
    pub show_percent: bool,
    /// `<c:dLbls><c:showBubbleSize>` for bubble-chart data labels.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub show_bubble_size: bool,
    /// `<c:dLbls><c:showLegendKey>` for classic data labels.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub show_legend_key: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub position: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_paint_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_hidden: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub format_code: Option<String>,
    /// `<c:dLbls><c:separator>` (§21.2.2.170) inserted between enabled label
    /// components. Office commonly stores a line break here for pie labels.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub separator: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_italic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_language: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_baseline: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_size_hpt: Option<i32>,
    /// `<c:dLbls><c:txPr>…<a:latin typeface>` series-default label face.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_face: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_rotation: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_wrap: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_vertical_anchor: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_vertical_mode: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_l_ins_emu: Option<i64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_t_ins_emu: Option<i64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_r_ins_emu: Option<i64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_b_ins_emu: Option<i64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_body_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_align: Option<String>,
    /// Series-default callout-box style (`<c:dLbls>` §21.2.2.49 `<c:spPr>`
    /// §21.2.2.197) — the box drawn around each pie/doughnut label. When present
    /// the pie renderer switches from plain outer-ring text to Word's boxed
    /// callout layout (box + optional leader line); `None` keeps the plain
    /// labels.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_box: Option<ChartLabelBox>,
    /// `<c:dLbls><c:showLeaderLines val>` (§21.2.2.183) — whether leader lines
    /// connect a label pulled away from its slice back to the slice. Absent =
    /// `false` (Office omits the element when leader lines are off). Only
    /// consulted by the pie/doughnut callout renderer.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub show_leader_lines: bool,
    /// `<c:dLbls><c:leaderLines>` (§21.2.2.92) `<c:spPr><a:ln><a:solidFill>`
    /// resolved hex (no `#`) — the leader-line stroke color. `None` falls back
    /// to a neutral grey in the renderer.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub leader_line_color: Option<String>,
    /// `<c:dLbls><c:leaderLines><c:spPr><a:ln w>` leader-line width in EMU.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub leader_line_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub leader_line_hidden: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub leader_line_dash: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub leader_line_paint_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub leader_line_style: Option<ChartExElementStyle>,
}

/// Mirror of TS `ChartErrBars`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartErrBars {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<ChartExElementStyle>,
    pub dir: String,
    pub bar_type: String,
    pub plus: Vec<Option<f64>>,
    pub minus: Vec<Option<f64>>,
    pub no_end_cap: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dash: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hidden: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_paint_authored: Option<bool>,
}

/// Mirror of TS `SecondaryValueAxis`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct SecondaryValueAxis {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_gridline_style: Option<ChartExElementStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_gridline_style: Option<ChartExElementStyle>,
    pub min: Option<f64>,
    pub max: Option<f64>,
    pub title: Option<String>,
    pub hidden: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub format_code: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub number_format: Option<ChartAxisNumberFormat>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub display_units: Option<ChartDisplayUnits>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_paint_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_size_hpt: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_italic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_face: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_dash: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_paint_authored: Option<bool>,
    pub line_hidden: bool,
    pub major_tick_mark: String,
    /// `<c:valAx><c:minorTickMark val>` (§21.2.2.115). Omission is retained so
    /// the renderer can apply the host's chart-family-specific default.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_tick_mark: Option<String>,
    /// `<c:valAx><c:minorGridlines>` presence and authored line paint.
    #[serde(default)]
    pub minor_gridlines: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_gridline_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_gridline_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_gridline_dash: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_gridline_paint_authored: Option<bool>,
    #[serde(default)]
    pub major_gridlines: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_gridline_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_gridline_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_gridline_dash: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_gridline_paint_authored: Option<bool>,
    /// `<c:valAx><c:majorUnit val>` (§21.2.2.103) — explicit major-unit step on
    /// this secondary axis, overriding the auto "nice" step. `None` ⇒ auto.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_unit: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minor_unit: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub log_base: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub orientation: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tick_label_pos: Option<String>,
    /// Category-axis interval fields are also carried here when this object is
    /// used as `secondaryCatAxis` for a top/right `<c:catAx>`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_alignment: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_offset_percent: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tick_label_skip: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tick_mark_skip: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crosses: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crosses_at: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_font_size_hpt: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_font_italic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_font_paint_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_font_face: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_rotation: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_vertical_mode: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_manual_layout: Option<ChartManualLayout>,
}

/// ECMA-376 §21.2.2.45 display-unit scaling plus the optional §21.2.2.46
/// label. The divisor is kept explicitly so custom and built-in units share
/// one renderer rule. Label formatting is separate from the axis title.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartDisplayUnits {
    pub divisor: f64,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub built_in_unit: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label: Option<ChartDisplayUnitsLabel>,
}

/// Authored display-unit label properties. `text` is absent for Office's
/// automatic label generated from the unit token/divisor.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartDisplayUnitsLabel {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub manual_layout: Option<ChartManualLayout>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_size_hpt: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_italic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_color: Option<String>,
    /// An authored DrawingML text fill owns this cascade slot even when the
    /// paint is `noFill` or cannot be resolved by the host. Keep that
    /// provenance separate from the optional resolved colour so lower axis or
    /// chart defaults cannot revive the label text.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_paint_authored: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_hidden: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_face: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rotation: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub box_style: Option<ChartLabelBox>,
}

/// One box-and-whisker series (chartEx `boxWhisker`, MS 2014 chartex ext).
///
/// A chartEx box-and-whisker chart carries one `<cx:series layoutId="boxWhisker">`
/// per data column, each referencing its own `<cx:data>` (via `<cx:dataId>`) of
/// RAW sample points grouped by category. Statistics (quartiles / mean /
/// whiskers / outliers) are computed by the renderer per the
/// `<cx:layoutPr><cx:statistics quartileMethod>` and `<cx:visibility>` flags;
/// the parser only groups the raw points by category and threads the flags.
/// Mirror of TS `ChartexBoxSeries`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartexBoxSeries {
    /// Series display name (`<cx:tx><cx:txData><cx:v>`), e.g. "Series1".
    pub name: String,
    /// Effective ChartEx `CT_Series@formatIdx`; omitted authoring resolves to
    /// the original document-order series index.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_format_idx: Option<u32>,
    /// Explicit `<cx:series><cx:spPr>` fill (hex, no `#`). Absent authoring is
    /// kept as `None` so the shared renderer can apply Chart Style / linked
    /// Chart Colors before falling back to theme accents.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    /// Explicit `<cx:series><cx:spPr><a:ln>` outline color (hex, no `#`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_color: Option<String>,
    /// Explicit series outline width from `<a:ln@w>` (EMU).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_width_emu: Option<u32>,
    /// Lossless series-local `<cx:spPr>` paint.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chartex_style: Option<ChartExElementStyle>,
    /// Raw sample values grouped by category, parallel to
    /// `ChartexBoxWhisker::categories`. Outer index = category, inner = the
    /// sample points that fell in that category (source order preserved).
    pub values_by_category: Vec<Vec<f64>>,
    /// `<cx:layoutPr><cx:visibility meanMarker>` — draw the mean `×` marker.
    pub mean_marker: bool,
    /// `<cx:layoutPr><cx:visibility meanLine>` — draw a mean connector line.
    pub mean_line: bool,
    /// `<cx:layoutPr><cx:visibility outliers>` — draw outlier points.
    pub show_outliers: bool,
    /// `<cx:layoutPr><cx:visibility nonoutliers>` — draw the non-outlier
    /// (interior) points as dots in addition to the box.
    pub show_nonoutliers: bool,
    /// `<cx:layoutPr><cx:statistics quartileMethod>` — `"exclusive"` (Excel
    /// default, median excluded when splitting halves) or `"inclusive"`.
    pub quartile_method: String,
}

/// A chartEx box-and-whisker chart: the unique categories plus one series per
/// data column. Mirror of TS `ChartexBoxWhisker`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartexBoxWhisker {
    /// The source omitted its category dimension, so every series represents
    /// one category/box rather than a peer inside every category group.
    #[serde(default)]
    pub one_box_per_series: bool,
    /// Unique category labels in first-seen order (the box groups on the
    /// category axis). Each series bins its raw points into these.
    pub categories: Vec<String>,
    /// One entry per `<cx:series>`.
    pub series: Vec<ChartexBoxSeries>,
}

/// One row of a chartEx `sunburst` (MS 2014 chartex ext). A sunburst encodes
/// its hierarchy as one `<cx:strDim type="cat">` with several `<cx:lvl>`
/// (lvl[0] = deepest / Leaf, last lvl = root / Branch) and a single
/// `<cx:numDim type="size">`. Each row's `path` is the branch→…→leaf label
/// chain with empty trailing segments trimmed (a node that is itself a leaf
/// terminates early); `size` is that row's size value. Mirror of TS
/// `ChartexSunburstRow`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartexSunburstRow {
    /// Label chain root→leaf (Branch, Stem, …, Leaf), empty tail trimmed.
    pub path: Vec<String>,
    /// `<cx:numDim type="size">` value for this row (attaches to the deepest
    /// node in `path`).
    pub size: f64,
}

/// A chartEx sunburst: the flat rows the renderer folds into a ring tree, plus
/// the theme accent palette to color branches. Mirror of TS `ChartexSunburst`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartexSunburst {
    /// One row per deepest-level data point.
    pub rows: Vec<ChartexSunburstRow>,
}

/// A chartEx treemap: hierarchy rows plus the requested parent-label layout.
/// The row encoding is identical to sunburst because both layouts consume the
/// same deepest→root `<cx:strDim type="cat">` levels and numeric size values.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartexTreemap {
    pub rows: Vec<ChartexSunburstRow>,
    /// `<cx:layoutPr><cx:parentLabelLayout val>` (`banner`, `overlapping`, or
    /// `none`). Absent stays `None`; the renderer uses its neutral default.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub parent_label_layout: Option<String>,
}

/// One geospatial data row from a chartEx `regionMap` series.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartexRegionMapRow {
    /// Authored category label (usually a country/region name).
    pub label: String,
    /// Optional stable geographic entity identifier from `strDim@type=entityId`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub entity_id: Option<String>,
    /// `numDim@type=colorVal`; missing/non-finite source values remain absent.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<f64>,
}

/// Authored chartEx geography metadata. Opaque provider cache payloads are not
/// copied onto the public wire; their presence/provider are retained so a host
/// can distinguish authored geography from an offline fallback.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartexGeography {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub projection_type: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub viewed_region_type: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub culture_language: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub culture_region: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub attribution: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cache_provider: Option<String>,
    #[serde(default)]
    pub cache_present: bool,
}

/// A minimum/middle/maximum stop position in a Region Map value ramp.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartexValueColorStop {
    /// `extremeValue`, `number`, or `percent`.
    pub kind: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<f64>,
}

/// Optional authored `valueColors`/`valueColorPositions` gradient contract.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartexRegionMapColors {
    /// `CT_ValueColorPositions@count`; omitted defaults to two stops.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub stop_count: Option<u8>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub min_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mid_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub min_position: Option<ChartexValueColorStop>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mid_position: Option<ChartexValueColorStop>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max_position: Option<ChartexValueColorStop>,
}

/// ChartEx Region Map data-only model. Country lookup and geometry stay in the
/// core renderer so parser hosts remain deterministic and network-free.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartexRegionMap {
    pub rows: Vec<ChartexRegionMapRow>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub region_label_layout: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub geography: Option<ChartexGeography>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub colors: Option<ChartexRegionMapColors>,
}

/// ChartEx `CT_Binning` controls ([MS-ODRAWXML] 2.24.3.7). Numeric
/// underflow/overflow bounds are retained; the schema's `auto` token remains
/// `None`, as does omission.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartexHistogramBinning {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bin_size: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bin_count: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub interval_closed: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub underflow: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub overflow: Option<f64>,
}

/// Mirror of TS `ChartManualLayout`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartManualLayout {
    pub x_mode: String,
    pub y_mode: String,
    pub w_mode: String,
    pub h_mode: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub layout_target: Option<String>,
    pub x: f64,
    pub y: f64,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub w: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub h: Option<f64>,
}

/// Mirror of TS `LegendManualLayout`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct LegendManualLayout {
    pub x_mode: String,
    pub y_mode: String,
    pub w_mode: String,
    pub h_mode: String,
    pub x: f64,
    pub y: f64,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub w: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub h: Option<f64>,
}

/// Indexed classic-chart legend entry override (`CT_LegendEntry`).
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartLegendEntryOverride {
    pub idx: u32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub deleted: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_face: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_size_hpt: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_italic: Option<bool>,
}

/// Combine a chart-type family (`bar` / `line` / `area`) with its bar direction
/// and grouping into the canonical `ChartModel.chart_type` vocabulary the core
/// renderer dispatches on.
///
/// This is the Rust home of the logic the xlsx TS renderer used to run in
/// `canonicalChartType` (pptx already emitted the canonical string). `bar_dir`
/// is ECMA-376 §21.2.3.4 `ST_BarDir`: `"bar"` = horizontal, `"col"` (or any
/// other value) = vertical. `grouping` is §21.2.3.17 `ST_Grouping`. Non-bar /
/// non-line / non-area families are returned unchanged.
pub(super) fn canonical_chart_type(chart_type: &str, bar_dir: &str, grouping: &str) -> String {
    let canonical = match chart_type {
        "bar" => {
            let is_h = bar_dir == "bar";
            match (grouping, is_h) {
                ("stacked", true) => "stackedBarH",
                ("stacked", false) => "stackedBar",
                ("percentStacked", true) => "stackedBarHPct",
                ("percentStacked", false) => "stackedBarPct",
                (_, true) => "clusteredBarH",
                (_, false) => "clusteredBar",
            }
        }
        "line" => match grouping {
            "stacked" => "stackedLine",
            "percentStacked" => "stackedLinePct",
            _ => "line",
        },
        "area" => match grouping {
            "stacked" => "stackedArea",
            "percentStacked" => "stackedAreaPct",
            _ => "area",
        },
        other => other,
    };
    canonical.to_string()
}
