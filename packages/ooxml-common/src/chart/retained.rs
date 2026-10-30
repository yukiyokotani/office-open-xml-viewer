//! Structural retained-heap accounting for the shared chart wire model.
//!
//! Every owned field reachable from [`ChartModel`] is listed in an exhaustive
//! destructure. Adding a model field therefore fails compilation until its
//! retained storage is included here. Counts saturate so hostile capacities
//! cannot wrap into an accepted value.

use super::model::*;
use std::collections::{BTreeMap, HashMap};
use std::hash::BuildHasher;

/// Heap storage retained by an owned value, excluding the value's inline size.
pub trait RetainedBytes {
    fn heap_bytes(&self) -> u64;
}

macro_rules! zero_retained {
    ($($ty:ty),+ $(,)?) => {
        $(impl RetainedBytes for $ty {
            fn heap_bytes(&self) -> u64 { 0 }
        })+
    };
}

zero_retained!(
    (),
    bool,
    char,
    u8,
    u16,
    u32,
    u64,
    u128,
    usize,
    i8,
    i16,
    i32,
    i64,
    i128,
    isize,
    f32,
    f64,
);

impl RetainedBytes for String {
    fn heap_bytes(&self) -> u64 {
        u64::try_from(self.capacity()).unwrap_or(u64::MAX)
    }
}

impl<T: RetainedBytes> RetainedBytes for Vec<T> {
    fn heap_bytes(&self) -> u64 {
        let inline = u64::try_from(self.capacity())
            .unwrap_or(u64::MAX)
            .saturating_mul(u64::try_from(std::mem::size_of::<T>()).unwrap_or(u64::MAX));
        self.iter().fold(inline, |total, value| {
            total.saturating_add(value.heap_bytes())
        })
    }
}

impl<T: RetainedBytes> RetainedBytes for Box<T> {
    fn heap_bytes(&self) -> u64 {
        u64::try_from(std::mem::size_of::<T>())
            .unwrap_or(u64::MAX)
            .saturating_add((**self).heap_bytes())
    }
}

impl<T: RetainedBytes> RetainedBytes for Option<T> {
    fn heap_bytes(&self) -> u64 {
        self.as_ref().map_or(0, RetainedBytes::heap_bytes)
    }
}

impl<T: RetainedBytes, const N: usize> RetainedBytes for [T; N] {
    fn heap_bytes(&self) -> u64 {
        self.iter().fold(0u64, |total, value| {
            total.saturating_add(value.heap_bytes())
        })
    }
}

impl<A: RetainedBytes, B: RetainedBytes> RetainedBytes for (A, B) {
    fn heap_bytes(&self) -> u64 {
        self.0.heap_bytes().saturating_add(self.1.heap_bytes())
    }
}

impl<K: RetainedBytes, V: RetainedBytes, S: BuildHasher> RetainedBytes for HashMap<K, V, S> {
    fn heap_bytes(&self) -> u64 {
        let slots = u64::try_from(self.capacity())
            .unwrap_or(u64::MAX)
            .saturating_mul(u64::try_from(std::mem::size_of::<(K, V)>()).unwrap_or(u64::MAX));
        self.iter().fold(slots, |total, (key, value)| {
            total
                .saturating_add(key.heap_bytes())
                .saturating_add(value.heap_bytes())
        })
    }
}

impl<K: RetainedBytes, V: RetainedBytes> RetainedBytes for BTreeMap<K, V> {
    fn heap_bytes(&self) -> u64 {
        let entries = u64::try_from(self.len())
            .unwrap_or(u64::MAX)
            .saturating_mul(u64::try_from(std::mem::size_of::<(K, V)>()).unwrap_or(u64::MAX));
        self.iter().fold(entries, |total, (key, value)| {
            total
                .saturating_add(key.heap_bytes())
                .saturating_add(value.heap_bytes())
        })
    }
}

macro_rules! impl_retained_struct {
    ($ty:ident { $($field:ident),+ $(,)? }) => {
        impl RetainedBytes for $ty {
            fn heap_bytes(&self) -> u64 {
                let Self { $($field),+ } = self;
                0u64 $(.saturating_add($field.heap_bytes()))+
            }
        }
    };
}

macro_rules! impl_retained_foreign_struct {
    ($ty:path, $pattern:path { $($field:ident),+ $(,)? }) => {
        impl RetainedBytes for $ty {
            fn heap_bytes(&self) -> u64 {
                let $pattern { $($field),+ } = self;
                0u64 $(.saturating_add($field.heap_bytes()))+
            }
        }
    };
}

impl_retained_foreign_struct!(
    crate::fill::GradStop,
    crate::fill::GradStop { position, color }
);
impl_retained_foreign_struct!(crate::fill::FillRect, crate::fill::FillRect { l, t, r, b });
impl_retained_foreign_struct!(
    crate::fill::TileInfo,
    crate::fill::TileInfo {
        tx,
        ty,
        sx,
        sy,
        flip,
        algn
    }
);
impl_retained_foreign_struct!(crate::blip::SrcRect, crate::blip::SrcRect { l, t, r, b });
impl_retained_foreign_struct!(crate::blip::Duotone, crate::blip::Duotone { clr1, clr2 });

// CT_Blip effect recipes are bounded by the parser's shared effect-step
// ledger, so this structural recursion has a fixed depth.
// CT_Blip effect recipes are bounded by the parser's shared effect-step
// ledger, so this structural recursion has a fixed depth.
impl RetainedBytes for crate::blip::BlipEffect {
    fn heap_bytes(&self) -> u64 {
        use crate::blip::BlipEffect;
        match self {
            BlipEffect::Grayscale => 0,
            BlipEffect::BiLevel { thresh } => 0u64.saturating_add(thresh.heap_bytes()),
            BlipEffect::ColorChange {
                from,
                from_alpha,
                to,
                to_alpha,
                use_alpha,
            } => 0u64
                .saturating_add(from.heap_bytes())
                .saturating_add(from_alpha.heap_bytes())
                .saturating_add(to.heap_bytes())
                .saturating_add(to_alpha.heap_bytes())
                .saturating_add(use_alpha.heap_bytes()),
            BlipEffect::Luminance { bright, contrast } => 0u64
                .saturating_add(bright.heap_bytes())
                .saturating_add(contrast.heap_bytes()),
            BlipEffect::Duotone { clr1, clr2 } => 0u64
                .saturating_add(clr1.heap_bytes())
                .saturating_add(clr2.heap_bytes()),
            BlipEffect::AlphaBiLevel { thresh } => 0u64.saturating_add(thresh.heap_bytes()),
            BlipEffect::AlphaCeiling => 0,
            BlipEffect::AlphaFloor => 0,
            BlipEffect::AlphaInv { color } => 0u64.saturating_add(color.heap_bytes()),
            BlipEffect::AlphaMod { container } => 0u64.saturating_add(container.heap_bytes()),
            BlipEffect::Blur { radius, grow } => 0u64
                .saturating_add(radius.heap_bytes())
                .saturating_add(grow.heap_bytes()),
            BlipEffect::Hsl { hue, sat, lum } => 0u64
                .saturating_add(hue.heap_bytes())
                .saturating_add(sat.heap_bytes())
                .saturating_add(lum.heap_bytes()),
            BlipEffect::Tint { hue, amt } => 0u64
                .saturating_add(hue.heap_bytes())
                .saturating_add(amt.heap_bytes()),
            BlipEffect::FillOverlay { fill, blend } => 0u64
                .saturating_add(fill.heap_bytes())
                .saturating_add(blend.heap_bytes()),
            BlipEffect::Container { container } => 0u64.saturating_add(container.heap_bytes()),
            BlipEffect::Reference { reference } => 0u64.saturating_add(reference.heap_bytes()),
            BlipEffect::Definitions { containers, work } => 0u64
                .saturating_add(containers.heap_bytes())
                .saturating_add(work.heap_bytes()),
            BlipEffect::EffectCountLimit { observed } => 0u64.saturating_add(observed.heap_bytes()),
            BlipEffect::AlphaOutset { radius } => 0u64.saturating_add(radius.heap_bytes()),
            BlipEffect::Blend { container, blend } => 0u64
                .saturating_add(container.heap_bytes())
                .saturating_add(blend.heap_bytes()),
            BlipEffect::Fill { fill } => 0u64.saturating_add(fill.heap_bytes()),
            BlipEffect::Glow { glow } => 0u64.saturating_add(glow.heap_bytes()),
            BlipEffect::Shadow { shadow, inner } => 0u64
                .saturating_add(shadow.heap_bytes())
                .saturating_add(inner.heap_bytes()),
            BlipEffect::SoftEdge { radius } => 0u64.saturating_add(radius.heap_bytes()),
            BlipEffect::Reflection { reflection } => 0u64.saturating_add(reflection.heap_bytes()),
            BlipEffect::Transform {
                sx,
                sy,
                kx,
                ky,
                tx,
                ty,
            } => 0u64
                .saturating_add(sx.heap_bytes())
                .saturating_add(sy.heap_bytes())
                .saturating_add(kx.heap_bytes())
                .saturating_add(ky.heap_bytes())
                .saturating_add(tx.heap_bytes())
                .saturating_add(ty.heap_bytes()),
            BlipEffect::RelativeOffset { tx, ty } => 0u64
                .saturating_add(tx.heap_bytes())
                .saturating_add(ty.heap_bytes()),
            BlipEffect::PresetShadow { shadows } => 0u64.saturating_add(shadows.heap_bytes()),
            BlipEffect::AlphaModFix { amt } => 0u64.saturating_add(amt.heap_bytes()),
            BlipEffect::AlphaRepl { alpha } => 0u64.saturating_add(alpha.heap_bytes()),
            BlipEffect::ColorReplace { color } => 0u64.saturating_add(color.heap_bytes()),
            BlipEffect::Unsupported { effect } => 0u64.saturating_add(effect.heap_bytes()),
        }
    }
}
impl_retained_foreign_struct!(
    crate::blip::EffectContainer,
    crate::blip::EffectContainer {
        kind,
        name,
        effects
    }
);
impl_retained_foreign_struct!(
    crate::blip::BlipReflection,
    crate::blip::BlipReflection {
        base,
        kx,
        ky,
        algn,
        fade_dir,
        rot_with_shape
    }
);
impl RetainedBytes for crate::blip::EffectFill {
    fn heap_bytes(&self) -> u64 {
        use crate::blip::EffectFill;
        match self {
            EffectFill::None => 0,
            EffectFill::Solid { color } => color.heap_bytes(),
            EffectFill::Gradient {
                stops,
                angle,
                grad_type,
                scaled,
                path,
                fill_to_rect,
                tile_rect,
                flip,
                rot_with_shape,
            } => 0u64
                .saturating_add(stops.heap_bytes())
                .saturating_add(angle.heap_bytes())
                .saturating_add(grad_type.heap_bytes())
                .saturating_add(scaled.heap_bytes())
                .saturating_add(path.heap_bytes())
                .saturating_add(fill_to_rect.heap_bytes())
                .saturating_add(tile_rect.heap_bytes())
                .saturating_add(flip.heap_bytes())
                .saturating_add(rot_with_shape.heap_bytes()),
            EffectFill::Pattern { fg, bg, preset } => 0u64
                .saturating_add(fg.heap_bytes())
                .saturating_add(bg.heap_bytes())
                .saturating_add(preset.heap_bytes()),
            EffectFill::Image {
                image_path,
                mime_type,
                svg_image_path,
                src_rect,
                fill_rect,
                tile,
                dpi,
                stretch,
                rot_with_shape,
                blip_effects,
            } => 0u64
                .saturating_add(image_path.heap_bytes())
                .saturating_add(mime_type.heap_bytes())
                .saturating_add(svg_image_path.heap_bytes())
                .saturating_add(src_rect.heap_bytes())
                .saturating_add(fill_rect.heap_bytes())
                .saturating_add(tile.heap_bytes())
                .saturating_add(dpi.heap_bytes())
                .saturating_add(stretch.heap_bytes())
                .saturating_add(rot_with_shape.heap_bytes())
                .saturating_add(blip_effects.heap_bytes()),
        }
    }
}
impl_retained_foreign_struct!(
    crate::effect::Shadow,
    crate::effect::Shadow {
        color,
        alpha,
        blur,
        dist,
        dir,
        sx,
        sy,
        kx,
        ky,
        algn,
        rot_with_shape
    }
);
impl_retained_foreign_struct!(
    crate::effect::Glow,
    crate::effect::Glow {
        color,
        alpha,
        radius
    }
);
impl_retained_foreign_struct!(crate::effect::SoftEdge, crate::effect::SoftEdge { radius });
impl_retained_foreign_struct!(
    crate::effect::Reflection,
    crate::effect::Reflection {
        blur,
        dist,
        dir,
        st_a,
        st_pos,
        end_a,
        end_pos,
        sx,
        sy
    }
);
impl_retained_foreign_struct!(
    super::ChartAxisNumberFormat,
    super::ChartAxisNumberFormat {
        authored_code,
        source_linked,
    }
);

impl RetainedBytes for ChartStyleFill {
    fn heap_bytes(&self) -> u64 {
        match self {
            Self::Solid { color } => color.heap_bytes(),
            Self::Gradient {
                stops,
                angle,
                grad_type,
                scaled,
                path,
                fill_to_rect,
                tile_rect,
                flip,
                rot_with_shape,
            } => {
                let _ = (angle, scaled, rot_with_shape);
                stops
                    .heap_bytes()
                    .saturating_add(grad_type.heap_bytes())
                    .saturating_add(path.heap_bytes())
                    .saturating_add(fill_to_rect.heap_bytes())
                    .saturating_add(tile_rect.heap_bytes())
                    .saturating_add(flip.heap_bytes())
            }
            Self::Pattern { fg, bg, preset } => fg
                .heap_bytes()
                .saturating_add(bg.heap_bytes())
                .saturating_add(preset.heap_bytes()),
            Self::Image {
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
                blip_effects,
            } => {
                let _ = (dpi, rot_with_shape, src_rect, fill_rect, stretch, alpha);
                image_path
                    .heap_bytes()
                    .saturating_add(mime_type.heap_bytes())
                    .saturating_add(svg_image_path.heap_bytes())
                    .saturating_add(tile.heap_bytes())
                    .saturating_add(duotone.heap_bytes())
                    .saturating_add(blip_effects.heap_bytes())
            }
        }
    }
}

impl RetainedBytes for ChartCartesianAutoLayoutProfile {
    fn heap_bytes(&self) -> u64 {
        match self {
            Self::WordClassicColumn => 0,
        }
    }
}

impl_retained_struct!(ChartLineDashSegment { dash, space });

impl_retained_struct!(ChartExElementStyle {
    shape_properties_present,
    allow_no_fill_override,
    allow_no_line_override,
    font_size_hpt,
    font_bold,
    font_italic,
    font_color,
    font_colors,
    font_color_index,
    font_formatting_indices,
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
    line_dash,
    line_dash_authored,
    line_custom_dash,
    line_cap,
    line_join,
    line_compound,
    line_hidden,
    line_no_style,
    shadows,
    inner_shadows,
    glows,
    soft_edges,
    reflections,
    effect_authored,
    effect_no_style,
    effect_unsupported,
    fill_color_index,
    fill_formatting_indices,
    fill_semantic_fallback_indices,
    line_color_index,
    line_formatting_indices,
    line_semantic_fallback_indices,
    effect_formatting_indices,
    effect_color_index,
});

impl_retained_struct!(ChartSurfaceBandFormat {
    idx,
    style,
    fill,
    fill_hidden,
    line_color,
    line_width_emu,
    line_hidden,
});

impl_retained_struct!(ChartClassicSurfaceBandStyles {
    fixed,
    by_band_count,
});

impl_retained_struct!(ChartDataTable {
    style,
    show_horizontal_border,
    show_vertical_border,
    show_outline,
    show_keys,
    font_size_hpt,
    font_face,
    font_color,
    font_paint_authored,
    font_hidden,
    font_bold,
    font_italic,
    fill_color,
    fill,
    fill_hidden,
    fill_paint_authored,
    line_color,
    line_width_emu,
    line_dash,
    line_hidden,
    line_paint_authored,
});

impl_retained_struct!(ChartPlotGroup {
    kind,
    series_start,
    series_count,
    category_axis,
    value_axis,
    series_axis,
    axis_ids,
    grouping,
    bar_direction,
    scatter_style,
    radar_style,
    vary_colors,
    gap_width,
    overlap,
    bubble_scale,
    bubble_size_represents,
    show_negative_bubbles,
});

impl_retained_struct!(ChartModel {
    chart_type,
    title,
    title_rich_runs,
    title_present,
    authored_without_series,
    categories,
    category_source_hidden,
    category_levels,
    series,
    plot_groups,
    show_data_labels,
    val_min,
    val_max,
    cat_axis_title,
    val_axis_title,
    cat_axis_hidden,
    val_axis_hidden,
    cat_axis_line_hidden,
    val_axis_line_hidden,
    plot_area_bg,
    plot_area_fill,
    plot_area_fill_hidden,
    plot_area_fill_paint_authored,
    plot_area_fill_automatic,
    plot_area_line_color,
    plot_area_line_fill,
    plot_area_line_width_emu,
    plot_area_line_dash,
    plot_area_line_dash_authored,
    plot_area_line_custom_dash,
    plot_area_line_cap,
    plot_area_line_join,
    plot_area_line_compound,
    plot_area_line_hidden,
    plot_area_line_paint_authored,
    chart_bg,
    chart_fill,
    chart_fill_hidden,
    chart_fill_paint_authored,
    rounded_corners,
    plot_visible_only,
    show_legend,
    data_table,
    legend_pos,
    cat_axis_cross_between,
    val_axis_major_tick_mark,
    cat_axis_major_tick_mark,
    title_font_size_hpt,
    title_font_color,
    title_font_paint_authored,
    title_font_face,
    cat_axis_font_size_hpt,
    val_axis_font_size_hpt,
    data_label_font_size_hpt,
    subtotal_indices,
    val_axis_minor_tick_mark,
    cat_axis_minor_tick_mark,
    cat_axis_font_color,
    cat_axis_font_paint_authored,
    val_axis_font_color,
    val_axis_font_paint_authored,
    legend_manual_layout,
    legend_overlay,
    legend_entries,
    val_axis_format_code,
    val_axis_number_format,
    val_axis_display_units,
    cat_axis_display_units,
    bar_gap_width,
    bar_overlap,
    data_label_position,
    data_label_font_color,
    data_label_font_paint_authored,
    data_label_format_code,
    data_label_font_bold,
    data_label_font_italic,
    data_label_font_language,
    data_label_font_baseline,
    title_font_bold,
    title_font_italic,
    title_font_language,
    title_font_baseline,
    cat_axis_font_bold,
    cat_axis_font_italic,
    val_axis_font_bold,
    val_axis_font_italic,
    cat_axis_title_font_size_hpt,
    cat_axis_title_font_bold,
    cat_axis_title_font_italic,
    cat_axis_title_font_color,
    cat_axis_title_font_paint_authored,
    cat_axis_title_rotation,
    cat_axis_title_vertical_mode,
    cat_axis_title_manual_layout,
    cat_axis_title_text_vertical_inset_emu,
    val_axis_title_font_size_hpt,
    val_axis_title_font_bold,
    val_axis_title_font_italic,
    val_axis_title_font_color,
    val_axis_title_font_paint_authored,
    val_axis_title_rotation,
    val_axis_title_vertical_mode,
    val_axis_title_manual_layout,
    val_axis_title_text_vertical_inset_emu,
    chart_border_color,
    chart_border_line_fill,
    chart_border_width_emu,
    chart_border_dash,
    chart_border_dash_authored,
    chart_border_custom_dash,
    chart_border_cap,
    chart_border_join,
    chart_border_compound,
    chart_border_hidden,
    chart_border_paint_authored,
    cat_axis_crosses,
    cat_axis_crosses_at,
    val_axis_crosses,
    val_axis_crosses_at,
    cat_axis_line_color,
    cat_axis_line_width_emu,
    cat_axis_line_dash,
    cat_axis_line_paint_authored,
    val_axis_line_color,
    val_axis_line_width_emu,
    val_axis_line_dash,
    val_axis_line_paint_authored,
    cat_axis_format_code,
    cat_axis_number_format,
    cat_axis_min,
    cat_axis_max,
    title_manual_layout,
    plot_area_manual_layout,
    cartesian_auto_layout_profile,
    scatter_style,
    bubble_scale,
    bubble_size_represents,
    show_negative_bubbles,
    radar_style,
    secondary_val_axis,
    secondary_cat_axis,
    hole_size,
    first_slice_angle,
    cat_axis_font_face,
    val_axis_font_face,
    cat_axis_title_font_face,
    val_axis_title_font_face,
    data_label_font_face,
    legend_font_face,
    legend_font_color,
    legend_font_paint_authored,
    legend_font_size_hpt,
    legend_font_bold,
    legend_font_italic,
    legend_font_language,
    legend_font_baseline,
    legend_fill_color,
    legend_fill,
    legend_fill_hidden,
    legend_fill_paint_authored,
    legend_line_color,
    legend_line_fill,
    legend_line_width_emu,
    legend_line_dash,
    legend_line_dash_authored,
    legend_line_custom_dash,
    legend_line_cap,
    legend_line_join,
    legend_line_compound,
    legend_line_hidden,
    legend_line_paint_authored,
    theme_major_font_latin,
    theme_minor_font_latin,
    date1904,
    disp_blanks_as,
    show_data_labels_over_max,
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
    chartex_box,
    chartex_sunburst,
    chartex_treemap,
    chartex_region_map,
    chartex_histogram_binning,
    chartex_pareto_owner_index,
    chartex_pareto_sort_descending,
    chartex_pareto_flat_endpoint,
    chartex_suppress_geometry,
    chartex_pareto_outline_owner,
    chartex_primary_axis_right,
    chartex_show_unpaired_percentage_axis,
    chartex_accents,
    chartex_color_palette,
    chartex_color_style_method,
    chart_style_roles,
    classic_chart_style_roles,
    classic_varying_point_chart_style_roles,
    classic_varying_point_chart_style_roles_by_group,
    chart_style_color_palette,
    chart_style_color_method,
    chart_style_marker_size_pt,
    chart_style_marker_symbol,
    chart_text_style,
    chart_area_style,
    plot_area_style,
    legend_style,
    title_style,
    cat_axis_style,
    val_axis_style,
    cat_axis_title_style,
    val_axis_title_style,
    cat_axis_major_gridline_style,
    cat_axis_minor_gridline_style,
    val_axis_major_gridline_style,
    val_axis_minor_gridline_style,
    chartex_data_point_style,
    chartex_data_point_line_style,
    chartex_series_line_style,
    chartex_data_point_marker_style,
    chartex_marker_size_pt,
    chartex_marker_symbol,
    chartex_connector_lines,
    vary_colors,
    chart_text_boxes,
});

impl_retained_struct!(ChartStockBarPaint {
    style,
    fill_color,
    fill,
    fill_paint_authored,
    fill_hidden,
    line_color,
    line_paint_authored,
    line_width_emu,
    line_dash,
    line_cap,
    line_join,
    line_hidden,
});

impl_retained_struct!(ChartStockUpDownBarStyle {
    gap_width_percent,
    up,
    down,
});

impl_retained_struct!(ChartStockAutomaticStyle {
    line_color,
    line_width_emu,
    up_fill_color,
    down_fill_color,
});

impl_retained_struct!(ChartDecorationLineStyle {
    style,
    color,
    fill,
    paint_authored,
    width_emu,
    dash,
    cap,
    join,
    hidden,
});

impl_retained_struct!(ChartLineGroupDecorations {
    group_index,
    drop_lines,
    hi_low_lines,
    up_down_bars,
});

impl_retained_struct!(ChartAreaGroupDecorations {
    group_index,
    drop_lines,
});

impl_retained_struct!(ChartBarGroupDecorations {
    group_index,
    series_lines,
});

impl_retained_struct!(ChartOfPie {
    r#type,
    split_type,
    split_type_authored,
    split_pos,
    split_pos_authored,
    custom_split_indices,
    second_pie_size_percent,
    gap_width_percent,
    series_lines,
    series_line_style,
});

impl_retained_struct!(ChartThreeDSeriesAxis {
    style,
    title_style,
    major_gridline_style,
    minor_gridline_style,
    title,
    hidden,
    orientation,
    tick_label_pos,
    tick_label_skip,
    tick_mark_skip,
    major_tick_mark,
    minor_tick_mark,
    font_color,
    font_paint_authored,
    font_size_hpt,
    font_bold,
    font_italic,
    font_face,
    line_color,
    line_width_emu,
    line_dash,
    line_paint_authored,
    line_hidden,
    title_font_size_hpt,
    title_font_bold,
    title_font_italic,
    title_font_color,
    title_font_paint_authored,
    title_font_face,
    title_rotation,
    title_vertical_mode,
    title_manual_layout,
});

impl_retained_struct!(ChartThreeDSurface {
    style,
    fill_color,
    fill_hidden,
    line_color,
    line_width_emu,
    line_dash,
    line_hidden,
    thickness_percent,
    picture_options,
});

impl_retained_struct!(ChartThreeDPictureOptions {
    apply_to_front,
    apply_to_sides,
    apply_to_end,
    picture_format,
    picture_format_authored,
    picture_stack_unit,
    picture_stack_unit_authored,
});

impl_retained_struct!(ChartThreeD {
    view_3d_present,
    rotation_x,
    rotation_x_authored,
    rotation_y,
    rotation_y_authored,
    height_percent,
    height_percent_authored,
    depth_percent,
    depth_percent_authored,
    perspective,
    perspective_authored,
    right_angle_axes,
    right_angle_axes_authored,
    gap_depth_percent,
    gap_depth_percent_authored,
    shape,
    bar_grouping,
    series_axis,
    floor,
    side_wall,
    back_wall,
});

impl_retained_struct!(ChartTextRun {
    text,
    font_size_hpt,
    bold,
    italic,
    color,
    color_paint_authored,
    color_hidden,
    font_face,
    language,
    baseline,
    paragraph_align,
});

impl_retained_struct!(ChartTextParagraph { runs, align });

impl_retained_struct!(ChartTextBox {
    x,
    y,
    w,
    h,
    paragraphs,
    vertical_anchor,
    wrap,
    l_ins,
    t_ins,
    r_ins,
    b_ins,
});

impl_retained_struct!(ChartPatternFill {
    fill_type,
    fg,
    bg,
    preset,
});

impl_retained_struct!(ChartSeries {
    name,
    chartex_format_idx,
    color,
    fill_pattern,
    invert_if_negative,
    automatic_negative_style,
    inverted_fill,
    inverted_fill_hidden,
    inverted_fill_authored,
    inverted_line_color,
    inverted_line_width_emu,
    inverted_line_hidden,
    inverted_line_authored,
    chartex_style,
    line_color,
    line_width_emu,
    three_d_shape,
    values,
    source_hidden,
    data_point_colors,
    explosion,
    data_label_colors,
    label_color,
    series_type,
    line_group_index,
    area_group_index,
    bar_group_index,
    bar_group_direction,
    bar_group_grouping,
    bar_group_gap_width,
    bar_group_overlap,
    use_secondary_axis,
    categories,
    bubble_x_source_is_string,
    show_marker,
    val_format_code,
    cat_format_code,
    cat_format_builtin_id,
    cat_format_codes,
    marker_symbol,
    automatic_marker_symbol,
    marker_size,
    marker_fill,
    marker_fill_paint,
    marker_fill_paint_authored,
    marker_style,
    marker_line,
    marker_line_paint_authored,
    marker_line_width_emu,
    data_point_overrides,
    data_label_overrides,
    series_data_labels,
    err_bars,
    bubble_sizes,
    bubble_3d_group_default,
    bubble_3d,
    smooth,
    trend_lines,
    line_hidden,
});

impl_retained_struct!(ChartTrendline {
    style,
    name,
    trendline_type,
    order,
    period,
    forward,
    backward,
    intercept,
    disp_r_sqr,
    disp_eq,
    label_manual_layout,
    label_text,
    label_rich_runs,
    label_format_code,
    label_format_source_linked,
    label_font_size_hpt,
    label_font_bold,
    label_font_italic,
    label_font_color,
    label_font_paint_authored,
    label_font_hidden,
    label_font_face,
    label_font_language,
    label_font_baseline,
    label_text_rotation,
    label_text_wrap,
    label_text_vertical_anchor,
    label_text_vertical_mode,
    label_text_l_ins_emu,
    label_text_t_ins_emu,
    label_text_r_ins_emu,
    label_text_b_ins_emu,
    label_text_body_authored,
    label_box,
    label_text_align,
    line_color,
    line_width_emu,
    line_dash,
    line_hidden,
    line_paint_authored,
});

impl_retained_struct!(ChartDataPointOverride {
    idx,
    color,
    fill_hidden,
    chartex_style,
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
    marker_line_paint_authored,
    marker_line_width_emu,
    bubble_3d,
    explosion,
});

impl_retained_struct!(ChartDataLabelOverride {
    idx,
    text,
    rich_runs,
    position,
    font_color,
    font_paint_authored,
    font_hidden,
    font_size_hpt,
    font_face,
    font_bold,
    font_italic,
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
    text_align,
    format_code,
    separator,
    manual_layout,
    label_box,
    show_val,
    show_cat_name,
    show_ser_name,
    show_percent,
    show_bubble_size,
    show_legend_key,
    deleted,
});

impl_retained_struct!(ChartLabelBox {
    style,
    fill,
    fill_paint,
    fill_hidden,
    fill_paint_authored,
    border_color,
    border_fill,
    border_width_emu,
    border_hidden,
    border_paint_authored,
    border_dash,
    border_dash_authored,
    border_custom_dash,
    border_cap,
    border_join,
    border_compound,
});

impl_retained_struct!(ChartSeriesDataLabels {
    deleted,
    show_val,
    show_cat_name,
    show_ser_name,
    show_percent,
    show_bubble_size,
    show_legend_key,
    position,
    font_color,
    font_paint_authored,
    font_hidden,
    format_code,
    separator,
    font_bold,
    font_italic,
    font_language,
    font_baseline,
    font_size_hpt,
    font_face,
    text_rotation,
    text_wrap,
    text_vertical_anchor,
    text_vertical_mode,
    text_l_ins_emu,
    text_t_ins_emu,
    text_r_ins_emu,
    text_b_ins_emu,
    text_body_authored,
    text_align,
    label_box,
    show_leader_lines,
    leader_line_color,
    leader_line_width_emu,
    leader_line_hidden,
    leader_line_dash,
    leader_line_paint_authored,
    leader_line_style,
});

impl_retained_struct!(ChartErrBars {
    style,
    dir,
    bar_type,
    plus,
    minus,
    no_end_cap,
    color,
    line_width_emu,
    dash,
    hidden,
    line_paint_authored,
});

impl_retained_struct!(SecondaryValueAxis {
    style,
    title_style,
    major_gridline_style,
    minor_gridline_style,
    min,
    max,
    title,
    hidden,
    format_code,
    number_format,
    display_units,
    font_color,
    font_paint_authored,
    font_size_hpt,
    font_italic,
    font_bold,
    font_face,
    line_color,
    line_width_emu,
    line_dash,
    line_paint_authored,
    line_hidden,
    major_tick_mark,
    minor_tick_mark,
    minor_gridlines,
    minor_gridline_color,
    minor_gridline_width_emu,
    minor_gridline_dash,
    minor_gridline_paint_authored,
    major_gridlines,
    major_gridline_color,
    major_gridline_width_emu,
    major_gridline_dash,
    major_gridline_paint_authored,
    major_unit,
    minor_unit,
    log_base,
    orientation,
    tick_label_pos,
    label_alignment,
    label_offset_percent,
    tick_label_skip,
    tick_mark_skip,
    date_axis,
    crosses,
    crosses_at,
    title_font_size_hpt,
    title_font_bold,
    title_font_italic,
    title_font_color,
    title_font_paint_authored,
    title_font_face,
    title_rotation,
    title_vertical_mode,
    title_manual_layout,
});

impl_retained_struct!(ChartDisplayUnits {
    divisor,
    built_in_unit,
    label,
});

impl_retained_struct!(ChartDisplayUnitsLabel {
    text,
    manual_layout,
    font_size_hpt,
    font_bold,
    font_italic,
    font_color,
    font_paint_authored,
    font_hidden,
    font_face,
    rotation,
    box_style,
});

impl_retained_struct!(ChartexBoxSeries {
    name,
    chartex_format_idx,
    color,
    line_color,
    line_width_emu,
    chartex_style,
    values_by_category,
    mean_marker,
    mean_line,
    show_outliers,
    show_nonoutliers,
    quartile_method,
});

impl_retained_struct!(ChartexBoxWhisker {
    one_box_per_series,
    categories,
    series,
});

impl_retained_struct!(ChartexSunburstRow { path, size });

impl_retained_struct!(ChartexSunburst { rows });

impl_retained_struct!(ChartexTreemap {
    rows,
    parent_label_layout,
});

impl_retained_struct!(ChartexRegionMapRow {
    label,
    entity_id,
    value,
});

impl_retained_struct!(ChartexGeography {
    projection_type,
    viewed_region_type,
    culture_language,
    culture_region,
    attribution,
    cache_provider,
    cache_present,
});

impl_retained_struct!(ChartexValueColorStop { kind, value });

impl_retained_struct!(ChartexRegionMapColors {
    stop_count,
    min_color,
    mid_color,
    max_color,
    min_position,
    mid_position,
    max_position,
});

impl_retained_struct!(ChartexRegionMap {
    rows,
    region_label_layout,
    geography,
    colors,
});

impl_retained_struct!(ChartexHistogramBinning {
    bin_size,
    bin_count,
    interval_closed,
    underflow,
    overflow,
    edge_format_code,
});

impl_retained_struct!(ChartManualLayout {
    x_mode,
    y_mode,
    w_mode,
    h_mode,
    layout_target,
    x,
    y,
    w,
    h,
});

impl_retained_struct!(LegendManualLayout {
    x_mode,
    y_mode,
    w_mode,
    h_mode,
    x,
    y,
    w,
    h,
});

impl_retained_struct!(ChartLegendEntryOverride {
    idx,
    deleted,
    font_face,
    font_color,
    font_size_hpt,
    font_bold,
    font_italic,
});

impl RetainedBytes for crate::text::SpaceLine {
    fn heap_bytes(&self) -> u64 {
        match self {
            Self::Pct { val: _ } | Self::Pts { val: _ } => 0,
        }
    }
}
