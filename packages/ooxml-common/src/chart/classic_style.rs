//! Built-in classic chart styles (`c:style`, values 1..48).
//!
//! ECMA-376 Part 1 §21.2.3.46 Tables 1–6 define these defaults.  Keep the
//! compact table algebra here instead of growing the already-large chart XML
//! parser.  The returned role table deliberately remains separate from linked
//! Office 2013+ `styleN.xml` roles: direct formatting wins first, a linked
//! style is the next fallback, and this numeric built-in style is last.
//!
//! Effect-style tiers are defined by the same tables but are materialized by
//! the shared chart-effect layer.  This module owns the line, fill, and font
//! portions and reuses the ordinary theme style-matrix parser so gradients,
//! patterns, and line geometry do not acquire a second implementation.

use std::collections::BTreeMap;
use std::sync::OnceLock;

#[cfg(test)]
use super::EmptyChartImageResolver;
use super::{
    parse_chartex_element_style, ChartExElementStyle, ChartImageResolver, ChartImageSource,
    ColorResolver, MAX_CHART_COLOR_STYLE_ENTRIES,
};

const DRAWINGML_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const CHART_STYLE_NS: &str = "http://schemas.microsoft.com/office/drawing/2012/chartStyle";

/// Office's application default when an OOXML package omits its optional
/// theme part. This is deliberately *not* used for a present but malformed or
/// incomplete theme: authored theme data continues to own its unresolved
/// paint and therefore fails closed. Word, Excel, and PowerPoint use the same
/// three-entry DrawingML style matrix for theme-less classic charts; keeping
/// it here lets ECMA-376 §21.2.3.46's numeric style tables resolve through the
/// ordinary matrix parser instead of duplicating fill/line/effect semantics.
fn office_default_format_scheme() -> &'static crate::theme::ThemeFormatScheme {
    static SCHEME: OnceLock<crate::theme::ThemeFormatScheme> = OnceLock::new();
    SCHEME.get_or_init(|| {
        crate::theme::ThemeFormatScheme::parse(&format!(
            r#"<a:theme xmlns:a="{DRAWINGML_NS}"><a:themeElements>
              <a:fmtScheme name="Office">
                <a:fillStyleLst>
                  <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
                  <a:gradFill rotWithShape="1"><a:gsLst>
                    <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
                    <a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
                    <a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
                  </a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill>
                  <a:gradFill rotWithShape="1"><a:gsLst>
                    <a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs>
                    <a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs>
                    <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs>
                  </a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill>
                </a:fillStyleLst>
                <a:lnStyleLst>
                  <a:ln w="9525" cap="flat"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln>
                  <a:ln w="25400" cap="flat"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
                  <a:ln w="38100" cap="flat"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
                </a:lnStyleLst>
                <a:effectStyleLst>
                  <a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>
                  <a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>
                  <a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>
                </a:effectStyleLst>
                <a:bgFillStyleLst>
                  <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
                  <a:gradFill rotWithShape="1"><a:gsLst>
                    <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
                    <a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
                    <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs>
                  </a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill>
                  <a:gradFill rotWithShape="1"><a:gsLst>
                    <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
                    <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs>
                  </a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>
                </a:bgFillStyleLst>
              </a:fmtScheme>
            </a:themeElements></a:theme>"#,
        ))
    })
}

fn office_default_scheme_color(name: &str) -> Option<String> {
    Some(
        match name {
            "dk1" | "tx1" => "000000",
            "lt1" | "bg1" => "FFFFFF",
            "dk2" | "tx2" => "44546A",
            "lt2" | "bg2" => "E7E6E6",
            "accent1" => "4472C4",
            "accent2" => "ED7D31",
            "accent3" => "A5A5A5",
            "accent4" => "FFC000",
            "accent5" => "5B9BD5",
            "accent6" => "70AD47",
            "hlink" => "0563C1",
            "folHlink" => "954F72",
            _ => return None,
        }
        .to_owned(),
    )
}

/// Adds only the built-in color defaults used by classic chart-style recipes
/// when the optional theme part is absent. Font faces are never invented: an
/// absent theme carries no concrete font resource identity.
struct OfficeDefaultClassicResolver<'a> {
    source: &'a dyn ColorResolver,
}

impl ColorResolver for OfficeDefaultClassicResolver<'_> {
    fn resolve_solid_fill(&self, node: roxmltree::Node) -> Option<String> {
        self.source.resolve_solid_fill(node)
    }

    fn resolve_scheme_color(&self, name: &str) -> Option<String> {
        self.source
            .resolve_scheme_color(name)
            .or_else(|| office_default_scheme_color(name))
    }

    fn tint_mode(&self) -> crate::color::TintMode {
        self.source.tint_mode()
    }

    fn resolve_shape_fill(&self, parent: roxmltree::Node) -> Option<String> {
        self.source.resolve_shape_fill(parent)
    }

    fn theme_major_font_latin(&self) -> Option<String> {
        self.source.theme_major_font_latin()
    }

    fn theme_minor_font_latin(&self) -> Option<String> {
        self.source.theme_minor_font_latin()
    }

    fn resolve_series_accent(&self, idx: usize) -> Option<String> {
        self.source
            .resolve_series_accent(idx)
            .or_else(|| office_default_scheme_color(&format!("accent{}", idx % 6 + 1)))
    }

    fn classic_pattern2_set_transform(&self, set_index: usize) -> Option<f64> {
        self.source.classic_pattern2_set_transform(set_index)
    }

    fn theme_format_scheme(&self) -> Option<&crate::theme::ThemeFormatScheme> {
        Some(office_default_format_scheme())
    }

    fn default_chart_bg(&self) -> Option<String> {
        self.source.default_chart_bg()
    }

    fn default_plot_area_bg(&self) -> Option<String> {
        self.source.default_plot_area_bg()
    }

    fn implicit_outline_only_negative_column_style(&self) -> bool {
        self.source.implicit_outline_only_negative_column_style()
    }

    fn office_dark_text_contrast_applies(&self, style: u8) -> bool {
        self.source.office_dark_text_contrast_applies(style)
    }

    fn office_dark_title_contrast_applies(&self, style: u8) -> bool {
        self.source.office_dark_title_contrast_applies(style)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ThemeTier {
    None,
    Subtle,
    Moderate,
    Intense,
}

impl ThemeTier {
    fn matrix_index(self) -> Option<u8> {
        match self {
            Self::None => None,
            Self::Subtle => Some(1),
            Self::Moderate => Some(2),
            Self::Intense => Some(3),
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq)]
enum PaletteRecipe {
    Scheme(&'static str, f64),
    Pattern(u8),
    Fade(&'static str),
}

#[derive(Clone, Copy, Debug, PartialEq)]
struct DataPointRecipe {
    fill_2d: ThemeTier,
    fill_3d: ThemeTier,
    line: ThemeTier,
    line_width_multiplier: u32,
    pattern: PaletteRecipe,
    outline: Option<PaletteRecipe>,
}

#[derive(Clone, Copy, Debug, PartialEq)]
struct UpDownRecipe {
    fill: ThemeTier,
    line: ThemeTier,
    up: PaletteRecipe,
    down: PaletteRecipe,
    line_color: Option<PaletteRecipe>,
}

fn scheme(name: &'static str) -> PaletteRecipe {
    PaletteRecipe::Scheme(name, 0.0)
}

fn tint(name: &'static str, retained: f64) -> PaletteRecipe {
    // The table's percentage is DrawingML `<a:tint@val>`: the retained share
    // of the source color. `apply_signed_tint_or_shade` instead accepts the
    // distance toward white, hence the inversion.
    PaletteRecipe::Scheme(name, 1.0 - retained)
}

fn shade(name: &'static str, retained: f64) -> PaletteRecipe {
    // DrawingML `<a:shade@val>` is the retained share of the source colour.
    // The generated-colour helper instead accepts distance toward black.
    PaletteRecipe::Scheme(name, -(1.0 - retained))
}

fn style_accent(style: u8, first: u8) -> &'static str {
    const ACCENTS: [&str; 6] = [
        "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    ];
    ACCENTS[usize::from(style - first)]
}

fn data_point_recipe(style: u8) -> DataPointRecipe {
    let offset = (style - 1) % 8;
    let pattern = match offset {
        0 => PaletteRecipe::Pattern(if style == 41 { 4 } else { 1 }),
        1 => PaletteRecipe::Pattern(2),
        _ => PaletteRecipe::Fade(style_accent(style, style - offset + 2)),
    };
    let (fill_2d, fill_3d, line, line_width_multiplier, outline) = match style {
        1..=8 => (
            ThemeTier::Subtle,
            ThemeTier::Subtle,
            ThemeTier::None,
            3,
            None,
        ),
        9..=16 => (
            ThemeTier::Subtle,
            ThemeTier::Subtle,
            ThemeTier::Subtle,
            5,
            Some(scheme("lt1")),
        ),
        17..=24 => (
            ThemeTier::Intense,
            if style == 18 {
                ThemeTier::Intense
            } else {
                ThemeTier::Subtle
            },
            ThemeTier::None,
            5,
            None,
        ),
        25..=32 => (
            ThemeTier::Intense,
            if style == 26 {
                ThemeTier::Intense
            } else {
                ThemeTier::Subtle
            },
            ThemeTier::None,
            7,
            None,
        ),
        33 => (
            ThemeTier::Subtle,
            ThemeTier::Subtle,
            ThemeTier::Subtle,
            5,
            Some(shade("dk1", 0.5)),
        ),
        34 => (
            ThemeTier::Subtle,
            ThemeTier::Subtle,
            ThemeTier::Subtle,
            5,
            Some(PaletteRecipe::Pattern(3)),
        ),
        35..=40 => (
            ThemeTier::Subtle,
            ThemeTier::Subtle,
            ThemeTier::Subtle,
            5,
            Some(shade(style_accent(style, 35), 0.5)),
        ),
        41..=48 => (
            ThemeTier::Intense,
            if style == 42 {
                ThemeTier::Intense
            } else {
                ThemeTier::Subtle
            },
            ThemeTier::None,
            5,
            None,
        ),
        _ => unreachable!("caller validates ST_Style"),
    };
    DataPointRecipe {
        fill_2d,
        fill_3d,
        line,
        line_width_multiplier,
        pattern,
        outline,
    }
}

/// Whether Table 5's data-point palette depends on the highest formatting
/// index in the semantic object domain. Pattern palettes are stable by index;
/// Fade palettes use the complete domain to place their endpoints.
pub(super) fn surface_band_palette_depends_on_count(style: u8) -> bool {
    (1..=48).contains(&style) && matches!(data_point_recipe(style).pattern, PaletteRecipe::Fade(_))
}

/// Resolve the numeric `dataPoint3D` role in a Surface value-band domain.
///
/// A Surface band index is not a source series index. The final band count is
/// known only after the renderer plans the value axis, so callers retain a
/// bounded set of count-specific Fade roles while Pattern roles may retain one
/// complete 0..47 role. Reusing the ordinary resolver here is intentional: it
/// keeps theme-matrix gradients, patterns, effects, and `phClr` substitution
/// in the same implementation as every other classic data mark.
pub(super) fn resolve_classic_surface_band_style_with_images(
    style: u8,
    resolver: &dyn ColorResolver,
    image_resolver: &dyn ChartImageResolver,
    band_count: usize,
) -> Option<ChartExElementStyle> {
    if band_count == 0 || band_count > 48 {
        return None;
    }
    let band_indices = (0..band_count).collect::<Vec<_>>();
    resolve_classic_chart_style_roles_selected(
        style,
        resolver,
        image_resolver,
        None,
        &band_indices,
        Some(&band_indices),
        true,
    )?
    .remove("dataPoint3D")
}

/// Resolve the numeric Surface wireframe role in the value-band domain.
/// The source series supplies the height matrix, but mesh colour changes at
/// value-band boundaries, so the formatting index is the band index too.
pub(super) fn resolve_classic_surface_wireframe_band_style_with_images(
    style: u8,
    resolver: &dyn ColorResolver,
    image_resolver: &dyn ChartImageResolver,
    band_count: usize,
) -> Option<ChartExElementStyle> {
    if band_count == 0 || band_count > 48 {
        return None;
    }
    let band_indices = (0..band_count).collect::<Vec<_>>();
    resolve_classic_chart_style_roles_selected(
        style,
        resolver,
        image_resolver,
        None,
        &band_indices,
        None,
        false,
    )?
    .remove("dataPointWireframe")
}

fn up_down_recipe(style: u8) -> UpDownRecipe {
    let accent = match style {
        3..=8 => Some(style_accent(style, 3)),
        11..=16 => Some(style_accent(style, 11)),
        19..=24 => Some(style_accent(style, 19)),
        27..=32 => Some(style_accent(style, 27)),
        35..=40 => Some(style_accent(style, 35)),
        43..=48 => Some(style_accent(style, 43)),
        _ => None,
    };
    let (up, down) = match style {
        1 | 9 | 17 | 25 | 41 => (tint("dk1", 0.25), tint("dk1", 0.85)),
        2 | 10 | 18 | 26 => (tint("dk1", 0.05), tint("dk1", 0.95)),
        33 => (scheme("lt1"), tint("dk1", 0.85)),
        34 => (scheme("lt1"), tint("dk1", 0.95)),
        42 => (scheme("lt1"), scheme("dk1")),
        _ => {
            let accent = accent.expect("accent style range");
            (tint(accent, 0.25), shade(accent, 0.25))
        }
    };
    let (fill, line, line_color) = match style {
        1..=16 => (ThemeTier::Subtle, ThemeTier::Subtle, Some(scheme("tx1"))),
        17..=32 => (ThemeTier::Intense, ThemeTier::None, None),
        33 | 34 => (ThemeTier::Subtle, ThemeTier::Subtle, Some(scheme("dk1"))),
        35..=40 => (
            ThemeTier::Subtle,
            ThemeTier::Subtle,
            Some(shade(accent.expect("accent style range"), 0.25)),
        ),
        41..=48 => (ThemeTier::Intense, ThemeTier::None, None),
        _ => unreachable!("caller validates ST_Style"),
    };
    UpDownRecipe {
        fill,
        line,
        up,
        down,
        line_color,
    }
}

fn resolve_scheme_recipe(
    recipe: PaletteRecipe,
    formatting_indices: &[usize],
    highest_formatting_index: usize,
    resolver: &dyn ColorResolver,
) -> Vec<Option<String>> {
    let apply = |name: &str, amount: f64| {
        resolver.resolve_scheme_color(name).map(|color| {
            crate::color::apply_signed_tint_or_shade(&color, amount, resolver.tint_mode())
        })
    };
    let pattern_color = |pattern: u8, index: usize| -> Option<String> {
        const P1: [f64; 6] = [0.885, 0.55, 0.78, 0.925, 0.70, 0.30];
        const P4: [f64; 6] = [0.05, 0.55, 0.78, 0.15, 0.70, 0.30];
        let slot = index % 6;
        match pattern {
            1 => apply("dk1", 1.0 - P1[slot]),
            2 => {
                let set_index = index / 6;
                // ECMA-376 §21.2.3.46 Table 6 fixes the first set at the six
                // unmodified accents and requires changed tint/shade in each
                // repeated set, but leaves those transforms host-defined.
                // Excel's host resolver supplies its measured eight-set cycle.
                // Other hosts keep repeated sets unresolved rather than
                // inheriting Excel's application-defined endpoints.
                resolver
                    .classic_pattern2_set_transform(set_index)
                    .and_then(|amount| apply(style_accent(slot as u8 + 1, 1), amount))
            }
            3 => apply(style_accent(slot as u8 + 1, 1), -0.5),
            4 => apply("dk1", 1.0 - P4[slot]),
            _ => None,
        }
    };
    formatting_indices
        .iter()
        .map(|index| match recipe {
            PaletteRecipe::Scheme(name, amount) => apply(name, amount),
            PaletteRecipe::Pattern(pattern) => pattern_color(pattern, *index),
            PaletteRecipe::Fade(name) => {
                // Table 5 makes the fade endpoints application-defined.  This
                // is ECMA's published deterministic suggestion, not a claim of
                // cross-version Office pixel equivalence.  The source `c:ser`
                // indexes and the global maximum are supplied by the caller.
                let amount =
                    -0.70 + 1.40 * (*index as f64 / (highest_formatting_index as f64 + 1.0));
                apply(name, amount)
            }
        })
        .collect()
}

#[derive(Clone, Copy)]
enum Component {
    Fill,
    Line,
    Effect,
}

fn materialize_component(
    component: Component,
    tier: ThemeTier,
    colors: Vec<Option<String>>,
    resolver: &dyn ColorResolver,
    image_resolver: &dyn ChartImageResolver,
) -> ChartExElementStyle {
    if tier == ThemeTier::None {
        return match component {
            Component::Fill => ChartExElementStyle {
                fill_hidden: Some(true),
                fill_paint_authored: Some(true),
                ..ChartExElementStyle::default()
            },
            Component::Line => ChartExElementStyle {
                line_hidden: Some(true),
                line_paint_authored: Some(true),
                ..ChartExElementStyle::default()
            },
            Component::Effect => ChartExElementStyle {
                // At the lowest-precedence numeric layer this represents the
                // table's concrete "No Effect", not linked `NoStyle` (which
                // would continue falling through the cascade).
                effect_authored: Some(true),
                ..ChartExElementStyle::default()
            },
        };
    }
    // A theme part is optional in every OOXML host. When it is genuinely
    // absent there is no authored style-matrix paint to own this numeric
    // component, so retain the renderer's semantic automatic fallback. A
    // present-but-incomplete matrix is different: parsing below records an
    // authored unresolved component and therefore fails closed.
    if resolver.theme_format_scheme().is_none() {
        return ChartExElementStyle::default();
    }
    let tag = match component {
        Component::Fill => "fillRef",
        Component::Line => "lnRef",
        Component::Effect => "effectRef",
    };
    let xml = format!(
        r#"<cs:role xmlns:cs="{CHART_STYLE_NS}" xmlns:a="{DRAWINGML_NS}"><cs:{tag} idx="{}"><cs:styleClr val="auto"/></cs:{tag}></cs:role>"#,
        tier.matrix_index().expect("non-none theme tier")
    );
    let Ok(document) = crate::depth::parse_guarded(&xml) else {
        return ChartExElementStyle::default();
    };
    parse_chartex_element_style(
        document.root_element(),
        resolver,
        Some(&colors),
        Some("cycle"),
        image_resolver,
        ChartImageSource::Theme,
    )
}

fn apply_fill(target: &mut ChartExElementStyle, source: ChartExElementStyle) {
    target.fill_paints = source.fill_paints;
    target.fill_colors = source.fill_colors;
    target.fill_hidden = source.fill_hidden;
    target.fill_paint_authored = source.fill_paint_authored;
    target.fill_no_style = source.fill_no_style;
    target.fill_color_index = source.fill_color_index;
    target.fill_semantic_fallback_indices = source.fill_semantic_fallback_indices;
}

fn apply_line(target: &mut ChartExElementStyle, source: ChartExElementStyle) {
    target.line_colors = source.line_colors;
    target.line_paints = source.line_paints;
    target.line_paint_authored = source.line_paint_authored;
    target.line_width_emu = source.line_width_emu;
    target.line_dash = source.line_dash;
    target.line_dash_authored = source.line_dash_authored;
    target.line_custom_dash = source.line_custom_dash;
    target.line_cap = source.line_cap;
    target.line_join = source.line_join;
    target.line_compound = source.line_compound;
    target.line_hidden = source.line_hidden;
    target.line_no_style = source.line_no_style;
    target.line_color_index = source.line_color_index;
    target.line_semantic_fallback_indices = source.line_semantic_fallback_indices;
}

fn apply_effect(target: &mut ChartExElementStyle, source: ChartExElementStyle) {
    target.shadows = source.shadows;
    target.inner_shadows = source.inner_shadows;
    target.glows = source.glows;
    target.soft_edges = source.soft_edges;
    target.reflections = source.reflections;
    target.effect_authored = source.effect_authored;
    target.effect_no_style = source.effect_no_style;
    target.effect_unsupported = source.effect_unsupported;
}

fn data_effect_tier(style: u8) -> ThemeTier {
    match style {
        1..=8 | 33..=40 => ThemeTier::None,
        9..=16 => ThemeTier::Subtle,
        17..=24 => ThemeTier::Moderate,
        25..=32 | 41..=48 => ThemeTier::Intense,
        _ => unreachable!("caller validates ST_Style"),
    }
}

fn role_style(
    fill: Option<(ThemeTier, Vec<Option<String>>)>,
    line: Option<(ThemeTier, Vec<Option<String>>)>,
    resolver: &dyn ColorResolver,
    image_resolver: &dyn ChartImageResolver,
) -> ChartExElementStyle {
    let mut style = ChartExElementStyle::default();
    if let Some((tier, colors)) = fill {
        apply_fill(
            &mut style,
            materialize_component(Component::Fill, tier, colors, resolver, image_resolver),
        );
    }
    if let Some((tier, colors)) = line {
        apply_line(
            &mut style,
            materialize_component(Component::Line, tier, colors, resolver, image_resolver),
        );
    }
    style
}

fn constant_palette(recipe: PaletteRecipe, resolver: &dyn ColorResolver) -> Vec<Option<String>> {
    resolve_scheme_recipe(recipe, &[0], 0, resolver)
}

fn table_two(style: u8) -> (PaletteRecipe, PaletteRecipe, PaletteRecipe, ThemeTier) {
    match style {
        1..=32 => (
            scheme("tx1"),
            tint("tx1", 0.5),
            tint("tx1", 0.75),
            ThemeTier::Subtle,
        ),
        33..=34 => (
            scheme("dk1"),
            tint("tx1", 0.5),
            tint("dk1", 0.75),
            ThemeTier::Subtle,
        ),
        35..=40 => (
            scheme("dk1"),
            tint("tx1", 0.5),
            tint("dk1", 0.75),
            ThemeTier::Subtle,
        ),
        41..=48 => (
            scheme("dk1"),
            tint("tx1", 0.9),
            scheme("lt1"),
            ThemeTier::None,
        ),
        _ => unreachable!("caller validates ST_Style"),
    }
}

fn other_line_color(style: u8) -> PaletteRecipe {
    match style {
        1..=32 => scheme("tx1"),
        33..=34 => scheme("dk1"),
        35..=40 => shade("dk1", 0.25),
        41..=48 => scheme("lt1"),
        _ => unreachable!("caller validates ST_Style"),
    }
}

fn table_three(style: u8) -> (PaletteRecipe, PaletteRecipe, ThemeTier) {
    match style {
        1..=32 => (scheme("bg1"), scheme("bg1"), ThemeTier::None),
        33..=34 => (scheme("lt1"), tint("dk1", 0.20), ThemeTier::Subtle),
        35..=40 => (
            scheme("lt1"),
            scheme(style_accent(style, 35)),
            ThemeTier::Subtle,
        ),
        41..=48 => (scheme("dk1"), tint("dk1", 0.95), ThemeTier::Subtle),
        _ => unreachable!("caller validates ST_Style"),
    }
}

fn set_font_defaults(
    style: &mut ChartExElementStyle,
    font_size_hpt: i32,
    bold: bool,
    font_color: Option<&String>,
    resolver: &dyn ColorResolver,
) {
    style.font_size_hpt = Some(font_size_hpt);
    style.font_bold = Some(bold);
    style.font_italic = Some(false);
    style.font_color = font_color.cloned();
    // A present theme owns numeric text paint even when its scheme colour is
    // malformed or unresolved. A genuinely absent optional theme does not;
    // preserve the renderer's semantic automatic text colour in that case.
    style.font_paint_authored = resolver.theme_format_scheme().is_some().then_some(true);
    style.font_face = resolver.theme_minor_font_latin();
}

/// Apply host-evidenced automatic light text for dark built-in chart styles.
/// ECMA-376 Part 1 §21.2.3.46 says that chart text follows the Axis & Major
/// Gridlines colour, which is `dk1` for styles 41–48. Cross-host evidence covers
/// style 41; Word's host resolver opts the complete 41..=48 matrix into the
/// same `lt1` projection.
///
/// Chart titles have one observed source-shape boundary: Office uses the light
/// colour when the rich paragraph carries `a:pPr/a:defRPr`, including an empty
/// carrier, but retains the ECMA dark colour when that carrier is absent.
/// Word-produced style matrices across 41–48 and empty/language/size/bold
/// run-property counterexamples establish its broader scope. Direct and linked
/// text paint remain higher-precedence layers; this changes only the numeric role.
pub(super) fn apply_office_dark_text_contrast(
    style: u8,
    title_has_paragraph_default_run: bool,
    resolver: &dyn ColorResolver,
    roles: &mut BTreeMap<String, ChartExElementStyle>,
) {
    let ordinary_text = resolver.office_dark_text_contrast_applies(style);
    let title_text = resolver.office_dark_title_contrast_applies(style);
    if !ordinary_text && !title_text {
        return;
    }

    // Resolve the observed text token itself. The leader-line role also starts
    // from Table 2's `lt1`, but its completed line style may add transforms or
    // replace solid paint with a gradient/pattern. Reusing that role would make
    // text depend on an unrelated theme format-list entry.
    let default_resolver = OfficeDefaultClassicResolver { source: resolver };
    let resolver: &dyn ColorResolver = if resolver.theme_format_scheme().is_none() {
        &default_resolver
    } else {
        resolver
    };
    let light_color = resolver.resolve_scheme_color("lt1");
    // A numeric style owns its automatic text paint even when a present theme
    // has a malformed/unresolvable lt1 slot. This matches set_font_defaults:
    // the renderer must not silently substitute semantic black in that case.
    let light_paint_authored = Some(true);

    if ordinary_text {
        for role in [
            "categoryAxis",
            "seriesAxis",
            "valueAxis",
            "axisTitle",
            "dataLabel",
            "dataLabelCallout",
            "dataTable",
            "legend",
            "trendlineLabel",
        ] {
            if let Some(role_style) = roles.get_mut(role) {
                role_style.font_color = light_color.clone();
                role_style.font_paint_authored = light_paint_authored;
            }
        }
    }
    if title_text && title_has_paragraph_default_run {
        if let Some(title) = roles.get_mut("title") {
            title.font_color = light_color;
            title.font_paint_authored = light_paint_authored;
        }
    }
}

/// Resolve one complete numeric style. `formatting_indices` is compact source
/// order while each entry retains the index used by a renderer consumer. It is
/// either the source `c:ser@idx` domain or, for an effectively varying chart,
/// the data-point domain. Keeping those domains distinct preserves
/// Table 5's highest-index semantics without allocating through a hostile
/// `u32` index.
#[cfg(test)]
pub(super) fn resolve_classic_chart_style_roles(
    style: u8,
    resolver: &dyn ColorResolver,
    chart_font_size_hpt: Option<i32>,
    series_formatting_indices: &[usize],
    varying_point_formatting_indices: Option<&[usize]>,
) -> Option<BTreeMap<String, ChartExElementStyle>> {
    resolve_classic_chart_style_roles_with_images(
        style,
        resolver,
        &EmptyChartImageResolver,
        chart_font_size_hpt,
        series_formatting_indices,
        varying_point_formatting_indices,
    )
}

pub(super) fn resolve_classic_chart_style_roles_with_images(
    style: u8,
    resolver: &dyn ColorResolver,
    image_resolver: &dyn ChartImageResolver,
    chart_font_size_hpt: Option<i32>,
    series_formatting_indices: &[usize],
    varying_point_formatting_indices: Option<&[usize]>,
) -> Option<BTreeMap<String, ChartExElementStyle>> {
    resolve_classic_chart_style_roles_selected(
        style,
        resolver,
        image_resolver,
        chart_font_size_hpt,
        series_formatting_indices,
        varying_point_formatting_indices,
        false,
    )
}

/// Materialize only the four point-domain roles consumed by a varying classic
/// chart group. A document can contain thousands of bounded one-series groups;
/// rebuilding the other 26 chart-wide roles for every group would turn a
/// compact point-domain table into avoidable whole-document work.
pub(super) fn resolve_classic_varying_point_roles_with_images(
    style: u8,
    resolver: &dyn ColorResolver,
    image_resolver: &dyn ChartImageResolver,
    chart_font_size_hpt: Option<i32>,
    series_formatting_indices: &[usize],
    varying_point_formatting_indices: &[usize],
) -> Option<BTreeMap<String, ChartExElementStyle>> {
    resolve_classic_chart_style_roles_selected(
        style,
        resolver,
        image_resolver,
        chart_font_size_hpt,
        series_formatting_indices,
        Some(varying_point_formatting_indices),
        true,
    )
}

fn resolve_classic_chart_style_roles_selected(
    style: u8,
    resolver: &dyn ColorResolver,
    image_resolver: &dyn ChartImageResolver,
    chart_font_size_hpt: Option<i32>,
    series_formatting_indices: &[usize],
    varying_point_formatting_indices: Option<&[usize]>,
    varying_roles_only: bool,
) -> Option<BTreeMap<String, ChartExElementStyle>> {
    if !(1..=48).contains(&style) {
        return None;
    }
    // Select the absent-theme adapter in this frame. Avoiding a recursive
    // re-entry matters for parsers already operating near the fixed WASM/native
    // test-thread stack limit on deeply nested slide trees.
    let default_resolver = OfficeDefaultClassicResolver { source: resolver };
    let resolver: &dyn ColorResolver = if resolver.theme_format_scheme().is_none() {
        &default_resolver
    } else {
        resolver
    };
    // The palette is replayed per series/point by the renderer. Reuse the
    // linked Chart Colors ceiling so a sparse or exceptionally wide chart
    // cannot amplify one small `<c:style>` scalar into unbounded wire memory.
    if series_formatting_indices.len() > MAX_CHART_COLOR_STYLE_ENTRIES
        || varying_point_formatting_indices
            .is_some_and(|indices| indices.len() > MAX_CHART_COLOR_STYLE_ENTRIES)
    {
        return None;
    }
    let series_indices = if series_formatting_indices.is_empty() {
        &[0][..]
    } else {
        series_formatting_indices
    };
    let varying_point_domain =
        varying_point_formatting_indices.is_some_and(|indices| !indices.is_empty());
    let point_indices = varying_point_formatting_indices
        .filter(|indices| !indices.is_empty())
        .unwrap_or(series_indices);
    let (axis_color, minor_color, chart_line_color, floor_chart_line) = table_two(style);
    let other_color = other_line_color(style);
    let (chart_fill_color, floor_fill_color, floor_wall_fill) = table_three(style);
    let data = data_point_recipe(style);
    let bars = up_down_recipe(style);

    let fixed = |recipe| constant_palette(recipe, resolver);
    let palette = |recipe, indices: &[usize]| {
        let highest = indices.iter().copied().max().unwrap_or(0);
        resolve_scheme_recipe(recipe, indices, highest, resolver)
    };
    let data_palette = |recipe| palette(recipe, point_indices);
    let series_palette = |recipe| palette(recipe, series_indices);
    let mut roles = BTreeMap::new();
    let no_fill = || (ThemeTier::None, Vec::new());
    let no_line = || (ThemeTier::None, Vec::new());
    let make_role_style = |fill, line| role_style(fill, line, resolver, image_resolver);

    if !varying_roles_only {
        for role in ["categoryAxis", "seriesAxis", "valueAxis"] {
            roles.insert(
                role.to_owned(),
                make_role_style(
                    Some(no_fill()),
                    Some((ThemeTier::Subtle, fixed(axis_color))),
                ),
            );
        }
        roles.insert(
            "gridlineMajor".to_owned(),
            make_role_style(None, Some((ThemeTier::Subtle, fixed(axis_color)))),
        );
        roles.insert(
            "gridlineMinor".to_owned(),
            make_role_style(None, Some((ThemeTier::Subtle, fixed(minor_color)))),
        );
        roles.insert(
            "chartArea".to_owned(),
            make_role_style(
                Some((ThemeTier::Subtle, fixed(chart_fill_color))),
                Some((floor_chart_line, fixed(chart_line_color))),
            ),
        );
        roles.insert(
            "dataTable".to_owned(),
            make_role_style(
                Some(no_fill()),
                Some((ThemeTier::Subtle, fixed(chart_line_color))),
            ),
        );
        roles.insert(
            "floor".to_owned(),
            make_role_style(
                Some((floor_wall_fill, fixed(floor_fill_color))),
                Some((floor_chart_line, fixed(chart_line_color))),
            ),
        );
        roles.insert(
            "wall".to_owned(),
            make_role_style(
                Some((floor_wall_fill, fixed(floor_fill_color))),
                Some(no_line()),
            ),
        );
        roles.insert(
            "plotArea".to_owned(),
            make_role_style(
                // ECMA-376 §21.2.3.46 Table 1 fixes the 2-D Plot Area at a Subtle
                // themed fill. Table 3 supplies its colour but its final Themed
                // Fill column applies only to Floor & Walls.
                Some((ThemeTier::Subtle, fixed(floor_fill_color))),
                Some(no_line()),
            ),
        );
        roles.insert(
            "plotArea3D".to_owned(),
            make_role_style(Some(no_fill()), Some(no_line())),
        );
    }

    let point_colors = data_palette(data.pattern);
    let outline_colors = data.outline.map(data_palette).unwrap_or_default();
    roles.insert(
        "dataPoint".to_owned(),
        make_role_style(
            Some((data.fill_2d, point_colors.clone())),
            Some((data.line, outline_colors.clone())),
        ),
    );
    roles.insert(
        "dataPoint3D".to_owned(),
        make_role_style(
            Some((data.fill_3d, point_colors)),
            Some((data.line, outline_colors)),
        ),
    );
    let marker_fill_colors = if varying_point_domain {
        data_palette(data.pattern)
    } else {
        series_palette(data.pattern)
    };
    let marker_line_colors = if varying_point_domain {
        data_palette(data.pattern)
    } else {
        series_palette(data.pattern)
    };
    roles.insert(
        "dataPointMarker".to_owned(),
        make_role_style(
            Some((data.fill_2d, marker_fill_colors)),
            Some((ThemeTier::Subtle, marker_line_colors)),
        ),
    );
    let mut data_line = make_role_style(
        None,
        Some((
            ThemeTier::Subtle,
            if varying_point_domain {
                data_palette(data.pattern)
            } else {
                series_palette(data.pattern)
            },
        )),
    );
    data_line.line_width_emu = data_line
        .line_width_emu
        .map(|width| width.saturating_mul(data.line_width_multiplier));
    roles.insert("dataPointLine".to_owned(), data_line);
    // Surface wireframe is a 3-D mark, not one of ECMA-376 §21.2.3.46's
    // 2-D "Lines for Data Points" families. Excel-produced style-2 output
    // uses the Data Point (3-D) fill palette as the mesh colour without the
    // 2-D line-width multiplier; materialize that renderer-facing role here.
    if !varying_roles_only {
        roles.insert(
            "dataPointWireframe".to_owned(),
            make_role_style(None, Some((data.fill_3d, series_palette(data.pattern)))),
        );
    }

    // The palette vectors stay compact even when source c:ser@idx values are
    // sparse or large. Preserve the source-index correspondence alongside the
    // affected component so renderers can select the compact slot without an
    // allocation proportional to the largest untrusted index.
    let point_indices = point_indices.to_vec();
    for role in ["dataPoint", "dataPoint3D"] {
        if let Some(role_style) = roles.get_mut(role) {
            role_style.fill_formatting_indices = Some(point_indices.clone());
            role_style.line_formatting_indices = Some(point_indices.clone());
        }
    }
    let series_indices = series_indices.to_vec();
    let marker_indices = if varying_point_domain {
        point_indices.clone()
    } else {
        series_indices.clone()
    };
    if let Some(role_style) = roles.get_mut("dataPointMarker") {
        role_style.fill_formatting_indices = Some(marker_indices.clone());
    }
    for role in ["dataPointMarker", "dataPointLine"] {
        if let Some(role_style) = roles.get_mut(role) {
            role_style.line_formatting_indices = Some(marker_indices.clone());
        }
    }
    if let Some(role_style) = roles.get_mut("dataPointWireframe") {
        role_style.line_formatting_indices = Some(series_indices.clone());
    }

    // Table 6 requires a changed transform for every repeated Pattern 2 set,
    // but does not define that host transform. Excel evidence establishes the
    // complete eight-set (48 object) cycle in its host resolver; hosts without
    // repeated-set evidence support only the first six accents. Beyond each
    // host's supported boundary the
    // numeric style deliberately delegates to the mark's semantic automatic
    // paint. Keep those slots distinct from a theme colour that was authored
    // inside the observed range but could not be resolved: the latter remains
    // fail-closed and must not expose a lower automatic colour.
    let point_fallbacks = point_indices
        .iter()
        .copied()
        .filter(|index| {
            resolver
                .classic_pattern2_set_transform(*index / 6)
                .is_none()
        })
        .collect::<Vec<_>>();
    let series_fallbacks = series_indices
        .iter()
        .copied()
        .filter(|index| {
            resolver
                .classic_pattern2_set_transform(*index / 6)
                .is_none()
        })
        .collect::<Vec<_>>();
    if data.pattern == PaletteRecipe::Pattern(2) {
        for role in ["dataPoint", "dataPoint3D"] {
            if let Some(role_style) = roles.get_mut(role) {
                role_style.fill_semantic_fallback_indices =
                    (!point_fallbacks.is_empty()).then(|| point_fallbacks.clone());
            }
        }
        let marker_fallbacks = if varying_point_domain {
            &point_fallbacks
        } else {
            &series_fallbacks
        };
        if let Some(role_style) = roles.get_mut("dataPointMarker") {
            role_style.fill_semantic_fallback_indices =
                (!marker_fallbacks.is_empty()).then(|| marker_fallbacks.clone());
        }
        for role in ["dataPointMarker", "dataPointLine"] {
            if let Some(role_style) = roles.get_mut(role) {
                role_style.line_semantic_fallback_indices =
                    (!marker_fallbacks.is_empty()).then(|| marker_fallbacks.clone());
            }
        }
        if let Some(role_style) = roles.get_mut("dataPointWireframe") {
            role_style.line_semantic_fallback_indices =
                (!series_fallbacks.is_empty()).then(|| series_fallbacks.clone());
        }
    }
    // The Table 5 body and outline recipes are independent. Styles 10 and 34
    // use Pattern 2 only for the body; their fixed-lt1 / Pattern 3 outlines
    // remain valid beyond the host-supported Pattern 2 body boundary.
    if data.outline == Some(PaletteRecipe::Pattern(2)) {
        for role in ["dataPoint", "dataPoint3D"] {
            if let Some(role_style) = roles.get_mut(role) {
                role_style.line_semantic_fallback_indices =
                    (!point_fallbacks.is_empty()).then(|| point_fallbacks.clone());
            }
        }
    }

    if !varying_roles_only {
        for role in [
            "dropLine",
            "errorBar",
            "hiLoLine",
            "leaderLine",
            "seriesLine",
            "trendline",
        ] {
            roles.insert(
                role.to_owned(),
                make_role_style(None, Some((ThemeTier::Subtle, fixed(other_color)))),
            );
        }

        for (role, color) in [("upBar", bars.up), ("downBar", bars.down)] {
            roles.insert(
                role.to_owned(),
                make_role_style(
                    Some((bars.fill, fixed(color))),
                    Some((bars.line, bars.line_color.map(fixed).unwrap_or_default())),
                ),
            );
        }

        for role in [
            "axisTitle",
            "dataLabel",
            "dataLabelCallout",
            "legend",
            "title",
            "trendlineLabel",
        ] {
            roles.insert(
                role.to_owned(),
                make_role_style(Some(no_fill()), Some(no_line())),
            );
        }
    }

    // Table 1 assigns No Effect to every ordinary chart role. Tables 4–5
    // override only data points, markers, and up/down bars with the repeating
    // eight-style tier sequence; their effect color is always dk1.
    let no_effect = materialize_component(
        Component::Effect,
        ThemeTier::None,
        Vec::new(),
        resolver,
        image_resolver,
    );
    for role in roles.values_mut() {
        apply_effect(role, no_effect.clone());
    }
    let themed_effect = materialize_component(
        Component::Effect,
        data_effect_tier(style),
        fixed(scheme("dk1")),
        resolver,
        image_resolver,
    );
    for role in [
        "dataPoint",
        "dataPoint3D",
        "dataPointMarker",
        "upBar",
        "downBar",
    ] {
        if let Some(role_style) = roles.get_mut(role) {
            apply_effect(role_style, themed_effect.clone());
        }
    }

    // DrawingML `sz` uses hundredths of a point despite the historical `hpt`
    // field suffix. ECMA's absent chart-font default is therefore 1000 = 10pt.
    let font_size = chart_font_size_hpt.unwrap_or(1_000);
    let font_color = fixed(axis_color).into_iter().next().flatten();
    for role in [
        "categoryAxis",
        "seriesAxis",
        "valueAxis",
        "axisTitle",
        "dataLabel",
        "dataLabelCallout",
        "dataTable",
        "legend",
        "title",
        "trendlineLabel",
    ] {
        if let Some(role_style) = roles.get_mut(role) {
            let is_title = matches!(role, "axisTitle" | "title");
            set_font_defaults(
                role_style,
                if role == "title" {
                    font_size.saturating_mul(120) / 100
                } else {
                    font_size
                },
                is_title,
                font_color.as_ref(),
                resolver,
            );
        }
    }

    debug_assert_eq!(roles.len(), if varying_roles_only { 4 } else { 30 });
    Some(roles)
}

#[cfg(test)]
mod tests {
    use super::*;

    struct MatrixResolver {
        format_scheme: crate::theme::ThemeFormatScheme,
    }

    impl MatrixResolver {
        fn new() -> Self {
            let theme = format!(
                r#"<a:theme xmlns:a="{DRAWINGML_NS}"><a:themeElements>
                  <a:fmtScheme name="classic-style-test">
                    <a:fillStyleLst>
                      <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
                      <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
                      <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
                    </a:fillStyleLst>
                    <a:lnStyleLst>
                      <a:ln w="1000"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
                      <a:ln w="2000"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
                      <a:ln w="3000"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
                    </a:lnStyleLst>
                    <a:effectStyleLst>
                      <a:effectStyle><a:effectLst><a:outerShdw blurRad="1000" dist="1000" dir="0"><a:schemeClr val="phClr"/></a:outerShdw></a:effectLst></a:effectStyle>
                      <a:effectStyle><a:effectLst><a:glow rad="2000"><a:schemeClr val="phClr"/></a:glow></a:effectLst></a:effectStyle>
                      <a:effectStyle><a:effectLst><a:softEdge rad="3000"/></a:effectLst></a:effectStyle>
                    </a:effectStyleLst>
                    <a:bgFillStyleLst/>
                  </a:fmtScheme>
                </a:themeElements></a:theme>"#,
            );
            Self {
                format_scheme: crate::theme::ThemeFormatScheme::parse(&theme),
            }
        }
    }

    impl ColorResolver for MatrixResolver {
        fn resolve_solid_fill(&self, _: roxmltree::Node) -> Option<String> {
            None
        }

        fn resolve_scheme_color(&self, name: &str) -> Option<String> {
            Some(
                match name {
                    "accent1" => "4472C4",
                    "accent2" => "ED7D31",
                    "accent3" => "A5A5A5",
                    "accent4" => "FFC000",
                    "accent5" => "5B9BD5",
                    "accent6" => "70AD47",
                    "lt1" | "bg1" => "FFFFFF",
                    _ => "000000",
                }
                .to_owned(),
            )
        }

        fn theme_minor_font_latin(&self) -> Option<String> {
            Some("Minor Theme".to_owned())
        }

        fn theme_format_scheme(&self) -> Option<&crate::theme::ThemeFormatScheme> {
            Some(&self.format_scheme)
        }

        fn classic_pattern2_set_transform(&self, set_index: usize) -> Option<f64> {
            [0.0, -0.4, 0.2, -0.2, 0.4, -0.5, 0.3, -0.3]
                .get(set_index)
                .copied()
        }

        fn office_dark_text_contrast_applies(&self, style: u8) -> bool {
            (41..=48).contains(&style)
        }

        fn office_dark_title_contrast_applies(&self, style: u8) -> bool {
            (41..=48).contains(&style)
        }
    }

    #[test]
    fn every_st_style_value_materializes_all_classic_roles() {
        struct Resolver;
        impl ColorResolver for Resolver {
            fn resolve_solid_fill(&self, _: roxmltree::Node) -> Option<String> {
                None
            }
            fn resolve_scheme_color(&self, name: &str) -> Option<String> {
                Some(
                    match name {
                        "accent1" => "111111",
                        "accent2" => "222222",
                        "accent3" => "333333",
                        "accent4" => "444444",
                        "accent5" => "555555",
                        "accent6" => "666666",
                        "lt1" | "bg1" => "FFFFFF",
                        _ => "000000",
                    }
                    .to_owned(),
                )
            }
            fn theme_minor_font_latin(&self) -> Option<String> {
                Some("Minor Theme".to_owned())
            }
        }

        // 1/2 are Table 6 Patterns 1/2; 11..16 identify fade accent1..6.
        // Style 41 alone selects Pattern 4. Keeping the complete literal
        // sequence catches a shifted eight-style block or missing endpoint.
        const EXPECTED_PATTERNS: [u8; 48] = [
            1, 2, 11, 12, 13, 14, 15, 16, 1, 2, 11, 12, 13, 14, 15, 16, 1, 2, 11, 12, 13, 14, 15,
            16, 1, 2, 11, 12, 13, 14, 15, 16, 1, 2, 11, 12, 13, 14, 15, 16, 4, 2, 11, 12, 13, 14,
            15, 16,
        ];
        let pattern_code = |recipe| match recipe {
            PaletteRecipe::Pattern(pattern) => pattern,
            PaletteRecipe::Fade(accent) => 10 + accent[6..].parse::<u8>().unwrap(),
            PaletteRecipe::Scheme(_, _) => 0,
        };

        for style in 1..=48 {
            let roles = resolve_classic_chart_style_roles(
                style,
                &Resolver,
                None,
                &[0, 1, 2, 3, 4, 5],
                None,
            )
            .expect("valid style");
            assert_eq!(roles.len(), 30, "style {style}");
            assert_eq!(roles["title"].font_size_hpt, Some(1_200), "style {style}");
            assert_eq!(roles["title"].font_bold, Some(true), "style {style}");
            assert_eq!(roles["legend"].font_size_hpt, Some(1_000), "style {style}");
            assert_eq!(roles["legend"].font_face.as_deref(), Some("Minor Theme"));
            assert_eq!(
                pattern_code(data_point_recipe(style).pattern),
                EXPECTED_PATTERNS[usize::from(style - 1)],
                "style {style} Table 5/6 pattern",
            );
        }
        assert!(resolve_classic_chart_style_roles(0, &Resolver, None, &[0], None).is_none());
        assert!(resolve_classic_chart_style_roles(49, &Resolver, None, &[0], None).is_none());
    }

    #[test]
    fn varying_group_resolver_materializes_only_the_four_point_roles() {
        let resolver = MatrixResolver::new();
        let point_indices = [0, 1, 2, 47, 48];
        let full = resolve_classic_chart_style_roles_with_images(
            2,
            &resolver,
            &EmptyChartImageResolver,
            None,
            &[7],
            Some(&point_indices),
        )
        .expect("full style");
        let varying = resolve_classic_varying_point_roles_with_images(
            2,
            &resolver,
            &EmptyChartImageResolver,
            None,
            &[7],
            &point_indices,
        )
        .expect("varying roles");

        assert_eq!(
            varying.keys().map(String::as_str).collect::<Vec<_>>(),
            [
                "dataPoint",
                "dataPoint3D",
                "dataPointLine",
                "dataPointMarker"
            ],
        );
        for (role, actual) in varying {
            assert_eq!(actual, full[&role], "role {role}");
        }
    }

    #[test]
    fn numeric_style_resolves_theme_picture_fills_in_theme_relationship_scope() {
        struct PictureTheme {
            format_scheme: crate::theme::ThemeFormatScheme,
        }
        impl ColorResolver for PictureTheme {
            fn resolve_solid_fill(&self, _: roxmltree::Node) -> Option<String> {
                None
            }
            fn resolve_scheme_color(&self, _: &str) -> Option<String> {
                Some("4472C4".to_owned())
            }
            fn theme_format_scheme(&self) -> Option<&crate::theme::ThemeFormatScheme> {
                Some(&self.format_scheme)
            }
        }
        struct Images;
        impl ChartImageResolver for Images {
            fn resolve_image(
                &self,
                source: ChartImageSource,
                relationship_id: &str,
            ) -> Option<(String, String)> {
                (source == ChartImageSource::Theme && relationship_id == "rIdThemeImage")
                    .then(|| ("theme/media/picture.png".to_owned(), "image/png".to_owned()))
            }
        }
        let theme = format!(
            r#"<a:theme xmlns:a="{DRAWINGML_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:themeElements>
              <a:fmtScheme name="picture-theme">
                <a:fillStyleLst>
                  <a:blipFill><a:blip r:embed="rIdThemeImage"/><a:stretch/></a:blipFill>
                  <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
                  <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
                </a:fillStyleLst>
                <a:lnStyleLst>
                  <a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
                  <a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
                  <a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
                </a:lnStyleLst>
                <a:effectStyleLst><a:effectStyle/><a:effectStyle/><a:effectStyle/></a:effectStyleLst>
                <a:bgFillStyleLst/>
              </a:fmtScheme>
            </a:themeElements></a:theme>"#,
        );
        let resolver = PictureTheme {
            format_scheme: crate::theme::ThemeFormatScheme::parse(&theme),
        };

        let roles =
            resolve_classic_chart_style_roles_with_images(1, &resolver, &Images, None, &[0], None)
                .expect("numeric style");
        assert!(matches!(
            roles["dataPoint3D"].fill_paints.as_deref(),
            Some([Some(crate::chart::ChartStyleFill::Image { image_path, .. })])
                if image_path == "theme/media/picture.png"
        ));
    }

    #[test]
    fn unresolved_numeric_text_color_retains_paint_ownership() {
        struct UnresolvedTheme {
            format_scheme: crate::theme::ThemeFormatScheme,
        }
        impl ColorResolver for UnresolvedTheme {
            fn resolve_solid_fill(&self, _: roxmltree::Node) -> Option<String> {
                None
            }
            fn resolve_scheme_color(&self, _: &str) -> Option<String> {
                None
            }
            fn theme_format_scheme(&self) -> Option<&crate::theme::ThemeFormatScheme> {
                Some(&self.format_scheme)
            }
        }

        let resolver = UnresolvedTheme {
            format_scheme: crate::theme::ThemeFormatScheme::default(),
        };
        let roles =
            resolve_classic_chart_style_roles(2, &resolver, None, &[0], None).expect("valid style");
        for role in [
            "categoryAxis",
            "seriesAxis",
            "valueAxis",
            "axisTitle",
            "dataLabel",
            "dataLabelCallout",
            "dataTable",
            "legend",
            "title",
            "trendlineLabel",
        ] {
            assert_eq!(roles[role].font_color, None, "{role}");
            assert_eq!(roles[role].font_paint_authored, Some(true), "{role}");
        }
    }

    #[test]
    fn absent_theme_materializes_the_office_application_default() {
        struct NoTheme;
        impl ColorResolver for NoTheme {
            fn resolve_solid_fill(&self, _: roxmltree::Node) -> Option<String> {
                None
            }
        }

        let roles =
            resolve_classic_chart_style_roles(2, &NoTheme, None, &[0], None).expect("valid style");
        for role in ["categoryAxis", "valueAxis", "gridlineMajor"] {
            assert_eq!(roles[role].line_paint_authored, Some(true), "{role}");
            assert_eq!(roles[role].line_width_emu, Some(9_525), "{role}");
        }
        assert_eq!(roles["dataPoint"].line_paint_authored, Some(true));
        assert_eq!(roles["dataPoint"].line_hidden, Some(true));
        assert_eq!(roles["dataPoint"].fill_paint_authored, Some(true));
        assert_eq!(
            roles["dataPoint"].fill_colors.as_ref().unwrap()[0].as_deref(),
            Some("4472C4")
        );
        assert_eq!(roles["chartArea"].fill_paint_authored, Some(true));
        assert_eq!(roles["plotArea"].fill_paint_authored, Some(true));
        for role in ["categoryAxis", "title", "legend", "dataLabel"] {
            assert_eq!(roles[role].font_paint_authored, Some(true), "{role}");
            assert_eq!(roles[role].font_face, None, "{role}");
        }
        let moderate = resolve_classic_chart_style_roles(17, &NoTheme, None, &[0], None)
            .expect("moderate style");
        assert!(matches!(
            moderate["dataPoint"].fill_paints.as_deref(),
            Some([Some(crate::chart::ChartStyleFill::Gradient { stops, .. })]) if stops.len() == 3
        ));
        let intense = resolve_classic_chart_style_roles(25, &NoTheme, None, &[0], None)
            .expect("intense style");
        assert!(intense["dataPoint3D"]
            .shadows
            .as_ref()
            .and_then(|shadows| shadows.first())
            .and_then(Option::as_ref)
            .is_some());
    }

    #[test]
    fn table_six_first_cycles_and_fade_use_source_formatting_indexes() {
        struct Resolver;
        impl ColorResolver for Resolver {
            fn resolve_solid_fill(&self, _: roxmltree::Node) -> Option<String> {
                None
            }
            fn resolve_scheme_color(&self, name: &str) -> Option<String> {
                Some(
                    match name {
                        "dk1" => "000000",
                        "accent1" => "808080",
                        "accent2" => "222222",
                        "accent3" => "333333",
                        "accent4" => "444444",
                        "accent5" => "555555",
                        "accent6" => "666666",
                        _ => "FFFFFF",
                    }
                    .to_owned(),
                )
            }
            fn classic_pattern2_set_transform(&self, set_index: usize) -> Option<f64> {
                [0.0, -0.4, 0.2, -0.2, 0.4, -0.5, 0.3, -0.3]
                    .get(set_index)
                    .copied()
            }
        }
        let pattern2 = resolve_scheme_recipe(
            PaletteRecipe::Pattern(2),
            &(0..48).collect::<Vec<_>>(),
            47,
            &Resolver,
        );
        let cycle_amounts = [0.0, -0.4, 0.2, -0.2, 0.4, -0.5, 0.3, -0.3];
        for (index, actual) in pattern2.iter().enumerate() {
            let accent = match index % 6 {
                0 => "808080",
                1 => "222222",
                2 => "333333",
                3 => "444444",
                4 => "555555",
                _ => "666666",
            };
            assert_eq!(
                actual.as_deref(),
                Some(
                    crate::color::apply_signed_tint_or_shade(
                        accent,
                        cycle_amounts[index / 6],
                        Resolver.tint_mode(),
                    )
                    .as_str(),
                ),
                "Pattern 2 index {index}",
            );
        }
        assert_eq!(
            resolve_scheme_recipe(PaletteRecipe::Pattern(2), &[48], 48, &Resolver),
            vec![None],
            "the first unobserved ninth-set index remains unresolved",
        );

        // The formatting extent must never recolour existing points. Office
        // keeps indexes 0..5 unchanged and index 6 at its first repeated-set
        // shade in 7-, 12-, and 13-point controls alike.
        let repeated_accent1 =
            crate::color::apply_signed_tint_or_shade("808080", -0.4, Resolver.tint_mode());
        for highest in [6, 11, 12, 47] {
            assert_eq!(
                resolve_scheme_recipe(PaletteRecipe::Pattern(2), &[0, 5, 6], highest, &Resolver),
                vec![
                    Some("808080".to_owned()),
                    Some("666666".to_owned()),
                    Some(repeated_accent1.clone()),
                ],
                "highest formatting index {highest}",
            );
        }

        let fade = resolve_scheme_recipe(PaletteRecipe::Fade("accent1"), &[2, 8], 9, &Resolver);
        assert_ne!(fade[0], fade[1]);
        assert_eq!(fade.len(), 2, "sparse c:ser indexes stay compact");

        let roles =
            resolve_classic_chart_style_roles(3, &MatrixResolver::new(), None, &[2, 8], None)
                .unwrap();
        assert_eq!(
            roles["dataPoint"].fill_formatting_indices.as_deref(),
            Some(&[2, 8][..]),
        );
        assert_eq!(roles["dataPoint"].fill_colors.as_ref().unwrap().len(), 2);
    }

    #[test]
    fn surface_band_domain_uses_band_count_for_fade_but_not_pattern_cycles() {
        let resolver = MatrixResolver::new();
        let pattern_six = resolve_classic_surface_band_style_with_images(
            2,
            &resolver,
            &EmptyChartImageResolver,
            6,
        )
        .unwrap();
        let pattern_thirteen = resolve_classic_surface_band_style_with_images(
            2,
            &resolver,
            &EmptyChartImageResolver,
            13,
        )
        .unwrap();
        assert!(!surface_band_palette_depends_on_count(2));
        assert_eq!(
            &pattern_six.fill_colors.as_ref().unwrap()[..6],
            &pattern_thirteen.fill_colors.as_ref().unwrap()[..6],
            "Pattern 2 colors are stable by band index",
        );

        let fade_six = resolve_classic_surface_band_style_with_images(
            3,
            &resolver,
            &EmptyChartImageResolver,
            6,
        )
        .unwrap();
        let fade_thirteen = resolve_classic_surface_band_style_with_images(
            3,
            &resolver,
            &EmptyChartImageResolver,
            13,
        )
        .unwrap();
        assert!(surface_band_palette_depends_on_count(3));
        assert_ne!(
            fade_six.fill_colors.as_ref().unwrap()[5],
            fade_thirteen.fill_colors.as_ref().unwrap()[5],
            "Table 5 Fade uses the final semantic band extent",
        );
        assert!(resolve_classic_surface_band_style_with_images(
            2,
            &resolver,
            &EmptyChartImageResolver,
            49,
        )
        .is_none());

        let wireframe = resolve_classic_surface_wireframe_band_style_with_images(
            2,
            &resolver,
            &EmptyChartImageResolver,
            5,
        )
        .unwrap();
        assert_eq!(
            wireframe.line_formatting_indices.as_deref(),
            Some(&[0, 1, 2, 3, 4][..]),
        );
        assert_eq!(wireframe.line_colors.as_ref().map(Vec::len), Some(5));
    }

    #[test]
    fn every_classic_style_resolves_surface_band_boundaries_in_its_own_domain() {
        let resolver = MatrixResolver::new();
        for style in 1..=48 {
            let count_dependent = (style - 1) % 8 >= 2;
            assert_eq!(
                surface_band_palette_depends_on_count(style),
                count_dependent,
                "style {style}",
            );
            for band_count in [1, 6, 7, 12, 13, 48] {
                let role = resolve_classic_surface_band_style_with_images(
                    style,
                    &resolver,
                    &EmptyChartImageResolver,
                    band_count,
                )
                .unwrap();
                assert_eq!(
                    role.fill_formatting_indices.as_ref().map(Vec::len),
                    Some(band_count),
                    "style {style}, {band_count} bands",
                );
                assert_eq!(
                    role.fill_colors.as_ref().map(Vec::len),
                    Some(band_count),
                    "style {style}, {band_count} bands",
                );
            }
        }
    }

    #[test]
    fn varying_palette_materialization_has_an_exact_resource_boundary() {
        struct NoTheme;
        impl ColorResolver for NoTheme {
            fn resolve_solid_fill(&self, _: roxmltree::Node) -> Option<String> {
                None
            }
        }
        let exact = (0..MAX_CHART_COLOR_STYLE_ENTRIES).collect::<Vec<_>>();
        let oversized = (0..=MAX_CHART_COLOR_STYLE_ENTRIES).collect::<Vec<_>>();
        assert!(
            resolve_classic_chart_style_roles(11, &NoTheme, None, &[0], Some(&exact)).is_some()
        );
        assert!(
            resolve_classic_chart_style_roles(11, &NoTheme, None, &[0], Some(&oversized)).is_none()
        );
    }

    #[test]
    fn host_unresolved_pattern_two_sets_delegate_without_erasing_outlines() {
        struct FirstSetOnly(MatrixResolver);
        impl ColorResolver for FirstSetOnly {
            fn resolve_solid_fill(&self, _: roxmltree::Node) -> Option<String> {
                None
            }
            fn resolve_scheme_color(&self, name: &str) -> Option<String> {
                self.0.resolve_scheme_color(name)
            }
            fn theme_format_scheme(&self) -> Option<&crate::theme::ThemeFormatScheme> {
                self.0.theme_format_scheme()
            }
        }
        let resolver = FirstSetOnly(MatrixResolver::new());
        for style in [2, 10, 34] {
            let roles =
                resolve_classic_chart_style_roles(style, &resolver, None, &[5, 6], Some(&[5, 6]))
                    .unwrap();
            for role in ["dataPoint", "dataPoint3D"] {
                assert_eq!(
                    roles[role].fill_semantic_fallback_indices.as_deref(),
                    Some(&[6][..])
                );
                assert!(roles[role].fill_colors.as_ref().unwrap()[0].is_some());
                assert!(roles[role].fill_colors.as_ref().unwrap()[1].is_none());
                if style != 2 {
                    assert_eq!(roles[role].line_semantic_fallback_indices, None);
                    assert!(roles[role].line_colors.as_ref().unwrap()[1].is_some());
                }
            }
            assert_eq!(
                roles["dataPointLine"]
                    .line_semantic_fallback_indices
                    .as_deref(),
                Some(&[6][..])
            );
        }
    }

    #[test]
    fn pattern_two_body_fallback_does_not_erase_independent_outlines() {
        let resolver = MatrixResolver::new();
        for style in [10, 34] {
            let roles =
                resolve_classic_chart_style_roles(style, &resolver, None, &[48], Some(&[48]))
                    .expect("valid style");
            for role in ["dataPoint", "dataPoint3D"] {
                assert_eq!(
                    roles[role].fill_semantic_fallback_indices.as_deref(),
                    Some(&[48][..]),
                    "style {style} {role} body",
                );
                assert_eq!(
                    roles[role].line_semantic_fallback_indices, None,
                    "style {style} {role} outline",
                );
                assert!(
                    roles[role]
                        .line_colors
                        .as_ref()
                        .is_some_and(|colors| colors.first().is_some_and(Option::is_some)),
                    "style {style} {role} keeps its resolved outline",
                );
            }
        }
    }

    #[test]
    fn office_dark_text_contrast_covers_41_through_48_and_keeps_title_gate() {
        let resolver = MatrixResolver::new();
        let text_roles = [
            "categoryAxis",
            "seriesAxis",
            "valueAxis",
            "axisTitle",
            "dataLabel",
            "dataLabelCallout",
            "dataTable",
            "legend",
            "trendlineLabel",
        ];

        let mut light = resolve_classic_chart_style_roles(40, &resolver, None, &[0], None).unwrap();
        apply_office_dark_text_contrast(40, true, &resolver, &mut light);
        assert_eq!(light["title"].font_color.as_deref(), Some("000000"));
        assert_eq!(light["categoryAxis"].font_color.as_deref(), Some("000000"));

        for style in 41..=48 {
            let mut without_title_carrier =
                resolve_classic_chart_style_roles(style, &resolver, None, &[0], None).unwrap();
            apply_office_dark_text_contrast(style, false, &resolver, &mut without_title_carrier);
            for role in text_roles {
                assert_eq!(
                    without_title_carrier[role].font_color.as_deref(),
                    Some("FFFFFF"),
                    "style {style} {role}",
                );
            }
            assert_eq!(
                without_title_carrier["title"].font_color.as_deref(),
                Some("000000"),
                "style {style} title without paragraph default run",
            );

            let mut with_title_carrier =
                resolve_classic_chart_style_roles(style, &resolver, None, &[0], None).unwrap();
            apply_office_dark_text_contrast(style, true, &resolver, &mut with_title_carrier);
            assert_eq!(
                with_title_carrier["title"].font_color.as_deref(),
                Some("FFFFFF"),
                "style {style} title with paragraph default run",
            );
        }
    }

    #[test]
    fn office_dark_text_resolves_lt1_independently_of_theme_line_paint() {
        struct CustomTheme {
            format_scheme: crate::theme::ThemeFormatScheme,
        }
        impl ColorResolver for CustomTheme {
            fn office_dark_text_contrast_applies(&self, style: u8) -> bool {
                (41..=48).contains(&style)
            }
            fn office_dark_title_contrast_applies(&self, style: u8) -> bool {
                (41..=48).contains(&style)
            }
            fn resolve_solid_fill(&self, _: roxmltree::Node) -> Option<String> {
                None
            }
            fn resolve_scheme_color(&self, name: &str) -> Option<String> {
                (name == "lt1").then(|| "A1B2C3".to_owned())
            }
            fn theme_format_scheme(&self) -> Option<&crate::theme::ThemeFormatScheme> {
                Some(&self.format_scheme)
            }
        }

        let resolver = CustomTheme {
            // An empty/malformed format list stands in for a gradient, pattern,
            // or otherwise unresolved line style. Text still resolves from the
            // independent theme scheme slot rather than that completed line.
            format_scheme: crate::theme::ThemeFormatScheme::default(),
        };
        let mut roles = resolve_classic_chart_style_roles(41, &resolver, None, &[0], None)
            .expect("valid style");
        assert_eq!(roles["leaderLine"].line_colors, None);

        apply_office_dark_text_contrast(41, true, &resolver, &mut roles);
        assert_eq!(roles["categoryAxis"].font_color.as_deref(), Some("A1B2C3"));
        assert_eq!(roles["title"].font_color.as_deref(), Some("A1B2C3"));
        assert_eq!(roles["categoryAxis"].font_paint_authored, Some(true));
    }

    #[test]
    fn office_dark_text_preserves_absent_and_unresolved_theme_boundaries() {
        struct NoTheme;
        impl ColorResolver for NoTheme {
            fn office_dark_text_contrast_applies(&self, style: u8) -> bool {
                (41..=48).contains(&style)
            }
            fn office_dark_title_contrast_applies(&self, style: u8) -> bool {
                (41..=48).contains(&style)
            }
            fn resolve_solid_fill(&self, _: roxmltree::Node) -> Option<String> {
                None
            }
        }
        let mut absent =
            resolve_classic_chart_style_roles(41, &NoTheme, None, &[0], None).expect("valid style");
        apply_office_dark_text_contrast(41, true, &NoTheme, &mut absent);
        assert_eq!(absent["title"].font_color.as_deref(), Some("FFFFFF"));
        assert_eq!(absent["title"].font_paint_authored, Some(true));

        struct UnresolvedTheme {
            format_scheme: crate::theme::ThemeFormatScheme,
        }
        impl ColorResolver for UnresolvedTheme {
            fn office_dark_text_contrast_applies(&self, style: u8) -> bool {
                (41..=48).contains(&style)
            }
            fn office_dark_title_contrast_applies(&self, style: u8) -> bool {
                (41..=48).contains(&style)
            }
            fn resolve_solid_fill(&self, _: roxmltree::Node) -> Option<String> {
                None
            }
            fn theme_format_scheme(&self) -> Option<&crate::theme::ThemeFormatScheme> {
                Some(&self.format_scheme)
            }
        }
        let resolver = UnresolvedTheme {
            format_scheme: crate::theme::ThemeFormatScheme::default(),
        };
        let mut unresolved = resolve_classic_chart_style_roles(41, &resolver, None, &[0], None)
            .expect("valid style");
        apply_office_dark_text_contrast(41, true, &resolver, &mut unresolved);
        assert_eq!(unresolved["title"].font_color, None);
        assert_eq!(unresolved["title"].font_paint_authored, Some(true));
    }

    #[test]
    fn table_boundaries_select_the_normative_recipes() {
        let resolver = MatrixResolver::new();
        assert_eq!(data_point_recipe(1).line_width_multiplier, 3);
        assert_eq!(data_point_recipe(9).line, ThemeTier::Subtle);
        assert_eq!(data_point_recipe(17).fill_2d, ThemeTier::Intense);
        assert_eq!(data_point_recipe(25).line_width_multiplier, 7);
        assert_eq!(data_point_recipe(33).outline, Some(shade("dk1", 0.5)));
        assert_eq!(
            shade("accent1", 0.25),
            PaletteRecipe::Scheme("accent1", -0.75)
        );
        assert_eq!(
            data_point_recipe(34).outline,
            Some(PaletteRecipe::Pattern(3))
        );
        assert_eq!(data_point_recipe(41).pattern, PaletteRecipe::Pattern(4));
        assert_eq!(data_point_recipe(42).fill_3d, ThemeTier::Intense);
        assert_eq!(table_two(40).3, ThemeTier::Subtle);
        assert_eq!(table_two(41).3, ThemeTier::None);
        assert_eq!(table_three(32).2, ThemeTier::None);
        assert_eq!(table_three(33).2, ThemeTier::Subtle);
        assert_eq!(up_down_recipe(16).line, ThemeTier::Subtle);

        let light = resolve_classic_chart_style_roles(1, &resolver, None, &[0], None).unwrap();
        // Table 1 fixes the 2-D Plot Area themed fill at Subtle. Table 3's
        // Themed Fill column applies only to Floor & Walls; for styles 1-32 it
        // must not turn the 2-D Plot Area into No Fill.
        assert_ne!(light["plotArea"].fill_hidden, Some(true));
        assert_eq!(light["floor"].fill_hidden, Some(true));
        assert_eq!(light["wall"].fill_hidden, Some(true));
        let dark = resolve_classic_chart_style_roles(33, &resolver, None, &[0], None).unwrap();
        assert_ne!(dark["plotArea"].fill_hidden, Some(true));
        assert_eq!(up_down_recipe(17).line, ThemeTier::None);
        assert_eq!(up_down_recipe(40).fill, ThemeTier::Subtle);
        assert_eq!(up_down_recipe(41).fill, ThemeTier::Intense);
    }

    #[test]
    fn all_48_styles_resolve_theme_matrix_geometry_and_effect_tiers() {
        let resolver = MatrixResolver::new();
        for style in 1..=48 {
            let roles = resolve_classic_chart_style_roles(
                style,
                &resolver,
                Some(180),
                &[0, 1, 2, 3, 4, 5],
                None,
            )
            .expect("valid style");
            let expected_width = match style {
                1..=8 => 3_000,
                25..=32 => 7_000,
                _ => 5_000,
            };
            assert_eq!(
                roles["dataPointLine"].line_width_emu,
                Some(expected_width),
                "style {style} Table 5 line multiplier",
            );
            assert_eq!(
                roles["dataPointWireframe"].line_width_emu,
                Some(match data_point_recipe(style).fill_3d {
                    ThemeTier::Subtle => 1_000,
                    ThemeTier::Moderate => 2_000,
                    ThemeTier::Intense => 3_000,
                    ThemeTier::None => unreachable!("3-D fill is always themed"),
                }),
                "style {style} wireframe uses the unmultiplied 3-D fill tier",
            );
            assert_eq!(
                roles["dataPointWireframe"].line_colors, roles["dataPoint3D"].fill_colors,
                "style {style} wireframe reuses the 3-D fill palette",
            );
            assert_eq!(
                roles["dataPoint"].fill_colors.as_ref().map(Vec::len),
                Some(6),
                "style {style} expanded Table 6 palette",
            );
            match style {
                1..=8 | 33..=40 => {
                    assert_eq!(roles["dataPoint"].effect_authored, Some(true));
                    assert!(roles["dataPoint"].shadows.is_none());
                    assert!(roles["dataPoint"].glows.is_none());
                    assert!(roles["dataPoint"].soft_edges.is_none());
                }
                9..=16 => assert!(roles["dataPoint"].shadows.is_some(), "style {style}"),
                17..=24 => assert!(roles["dataPoint"].glows.is_some(), "style {style}"),
                25..=32 | 41..=48 => {
                    assert!(roles["dataPoint"].soft_edges.is_some(), "style {style}")
                }
                _ => unreachable!(),
            }
            assert_eq!(roles["title"].font_size_hpt, Some(216));
        }
    }
}
