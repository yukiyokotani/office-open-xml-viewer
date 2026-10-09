//! Structural retained storage for worksheet ancillary models copied by a
//! cursor. Exhaustive field patterns fail compilation when a new field needs
//! accounting. Capacities and sums saturate rather than wrap at admission.
use super::*;
use ooxml_common::chart::RetainedBytes;

macro_rules! retained_struct {
    ($ty:ident { $($field:ident),* $(,)? }) => {
        impl RetainedBytes for $ty {
            fn heap_bytes(&self) -> u64 {
                let Self { $($field),* } = self;
                0u64 $(.saturating_add($field.heap_bytes()))*
            }
        }
    };
}

retained_struct!(Border {
    left,
    right,
    top,
    bottom,
    diagonal_up,
    diagonal_down,
    horizontal,
    vertical
});

retained_struct!(BorderEdge { style, color });

retained_struct!(CellRange {
    top,
    left,
    bottom,
    right
});

retained_struct!(CfIcon { icon_set, icon_id });

impl RetainedBytes for CfRule {
    fn heap_bytes(&self) -> u64 {
        match self {
            Self::CellIs {
                operator,
                formulas,
                dxf_id,
                priority,
                stop_if_true,
            } => 0u64
                .saturating_add(operator.heap_bytes())
                .saturating_add(formulas.heap_bytes())
                .saturating_add(dxf_id.heap_bytes())
                .saturating_add(priority.heap_bytes())
                .saturating_add(stop_if_true.heap_bytes()),
            Self::Expression {
                formula,
                dxf_id,
                priority,
                stop_if_true,
            } => 0u64
                .saturating_add(formula.heap_bytes())
                .saturating_add(dxf_id.heap_bytes())
                .saturating_add(priority.heap_bytes())
                .saturating_add(stop_if_true.heap_bytes()),
            Self::ColorScale {
                stops,
                priority,
                active_formula,
                stop_if_true,
            } => 0u64
                .saturating_add(stops.heap_bytes())
                .saturating_add(priority.heap_bytes())
                .saturating_add(active_formula.heap_bytes())
                .saturating_add(stop_if_true.heap_bytes()),
            Self::DataBar {
                color,
                min,
                max,
                priority,
                gradient,
                active_formula,
                ext_threshold_formula,
                stop_if_true,
            } => 0u64
                .saturating_add(color.heap_bytes())
                .saturating_add(min.heap_bytes())
                .saturating_add(max.heap_bytes())
                .saturating_add(priority.heap_bytes())
                .saturating_add(gradient.heap_bytes())
                .saturating_add(active_formula.heap_bytes())
                .saturating_add(ext_threshold_formula.heap_bytes())
                .saturating_add(stop_if_true.heap_bytes()),
            Self::Top10 {
                top,
                percent,
                rank,
                dxf_id,
                priority,
                stop_if_true,
            } => 0u64
                .saturating_add(top.heap_bytes())
                .saturating_add(percent.heap_bytes())
                .saturating_add(rank.heap_bytes())
                .saturating_add(dxf_id.heap_bytes())
                .saturating_add(priority.heap_bytes())
                .saturating_add(stop_if_true.heap_bytes()),
            Self::AboveAverage {
                above_average,
                equal_average,
                std_dev,
                dxf_id,
                priority,
                stop_if_true,
            } => 0u64
                .saturating_add(above_average.heap_bytes())
                .saturating_add(equal_average.heap_bytes())
                .saturating_add(std_dev.heap_bytes())
                .saturating_add(dxf_id.heap_bytes())
                .saturating_add(priority.heap_bytes())
                .saturating_add(stop_if_true.heap_bytes()),
            Self::IconSet {
                icon_set,
                cfvos,
                reverse,
                priority,
                custom_icons,
                active_formula,
                stop_if_true,
            } => 0u64
                .saturating_add(icon_set.heap_bytes())
                .saturating_add(cfvos.heap_bytes())
                .saturating_add(reverse.heap_bytes())
                .saturating_add(priority.heap_bytes())
                .saturating_add(custom_icons.heap_bytes())
                .saturating_add(active_formula.heap_bytes())
                .saturating_add(stop_if_true.heap_bytes()),
            Self::Other {
                kind,
                priority,
                unsupported_formula_phases,
                stop_if_true,
            } => 0u64
                .saturating_add(kind.heap_bytes())
                .saturating_add(priority.heap_bytes())
                .saturating_add(unsupported_formula_phases.heap_bytes())
                .saturating_add(stop_if_true.heap_bytes()),
        }
    }
}

retained_struct!(CfStop { kind, value, color });

retained_struct!(CfValue { kind, value });

retained_struct!(ChartAnchor {
    z_order,
    from_col,
    from_col_off,
    from_row,
    from_row_off,
    to_col,
    to_col_off,
    to_row,
    to_row_off,
    chart
});

retained_struct!(ConditionalFormat { sqref, rules });

retained_struct!(DataValidation {
    sqref,
    validation_type,
    operator,
    formula1,
    formula2,
    allow_blank,
    prompt_title,
    prompt,
    error_title,
    error_message
});

retained_struct!(DefinedName { name, formula });

retained_struct!(Dxf {
    font,
    fill,
    border,
    num_fmt,
    font_toggles
});

retained_struct!(DxfFontToggles {
    bold,
    italic,
    underline,
    strike
});

retained_struct!(Fill {
    pattern_type,
    fg_color,
    bg_color,
    gradient
});

retained_struct!(Font {
    bold,
    italic,
    underline,
    strike,
    size,
    color,
    name,
    scheme,
    charset,
    underline_style,
    vert_align
});

retained_struct!(GradientFillSpec {
    gradient_type,
    degree,
    left,
    right,
    top,
    bottom,
    stops
});

retained_struct!(GradientStopSpec { position, color });

retained_struct!(Hyperlink {
    col,
    row,
    url,
    location,
    display
});

retained_struct!(ImageAnchor {
    z_order,
    from_col,
    from_col_off,
    from_row,
    from_row_off,
    to_col,
    to_col_off,
    to_row,
    to_row_off,
    edit_as,
    anchor_tag,
    native_ext_cx,
    native_ext_cy,
    rotation,
    flip_h,
    flip_v,
    image_path,
    mime_type,
    svg_image_path,
    src_rect,
    alpha,
    duotone
});

retained_struct!(NumFmt {
    num_fmt_id,
    format_code
});

impl RetainedBytes for PathCmd {
    fn heap_bytes(&self) -> u64 {
        match self {
            Self::MoveTo { x, y } => 0u64
                .saturating_add(x.heap_bytes())
                .saturating_add(y.heap_bytes()),
            Self::LineTo { x, y } => 0u64
                .saturating_add(x.heap_bytes())
                .saturating_add(y.heap_bytes()),
            Self::CubicBezTo {
                x1,
                y1,
                x2,
                y2,
                x3,
                y3,
            } => 0u64
                .saturating_add(x1.heap_bytes())
                .saturating_add(y1.heap_bytes())
                .saturating_add(x2.heap_bytes())
                .saturating_add(y2.heap_bytes())
                .saturating_add(x3.heap_bytes())
                .saturating_add(y3.heap_bytes()),
            Self::QuadBezTo { x1, y1, x2, y2 } => 0u64
                .saturating_add(x1.heap_bytes())
                .saturating_add(y1.heap_bytes())
                .saturating_add(x2.heap_bytes())
                .saturating_add(y2.heap_bytes()),
            Self::ArcTo {
                wr,
                hr,
                st_ang,
                sw_ang,
            } => 0u64
                .saturating_add(wr.heap_bytes())
                .saturating_add(hr.heap_bytes())
                .saturating_add(st_ang.heap_bytes())
                .saturating_add(sw_ang.heap_bytes()),
            Self::Close => 0,
        }
    }
}

retained_struct!(PathInfo {
    w,
    h,
    fill,
    stroke,
    commands
});

retained_struct!(PivotAxisItem { kind, depth });

impl RetainedBytes for PivotCacheSource {
    fn heap_bytes(&self) -> u64 {
        match self {
            Self::Worksheet {
                sheet,
                reference,
                name,
                relationship_id,
            } => 0u64
                .saturating_add(sheet.heap_bytes())
                .saturating_add(reference.heap_bytes())
                .saturating_add(name.heap_bytes())
                .saturating_add(relationship_id.heap_bytes()),
            Self::External => 0,
            Self::Consolidation => 0,
            Self::Scenario => 0,
        }
    }
}

retained_struct!(PivotDataField {
    field,
    subtotal,
    raw_subtotal,
    name
});

retained_struct!(PivotLocation {
    range,
    first_header_row,
    first_data_row,
    first_data_col
});

impl RetainedBytes for PivotMetadataStatus {
    fn heap_bytes(&self) -> u64 {
        match self {
            Self::Complete => 0,
            Self::Partial { reasons } => 0u64.saturating_add(reasons.heap_bytes()),
        }
    }
}

retained_struct!(PivotPageField { field, item, name });

impl RetainedBytes for PivotPartialReason {
    fn heap_bytes(&self) -> u64 {
        match self {
            Self::MissingCacheRelationship => 0,
            Self::MalformedCacheRelationships => 0,
            Self::UnreadableCacheRelationships => 0,
            Self::ExternalCacheRelationship => 0,
            Self::AmbiguousCacheRelationship => 0,
            Self::UnreadableCacheDefinition => 0,
            Self::MalformedCacheDefinition => 0,
            Self::MalformedField { field } => 0u64.saturating_add(field.heap_bytes()),
            Self::UnsupportedCacheSource { source_type } => {
                0u64.saturating_add(source_type.heap_bytes())
            }
            Self::UnresolvedWorksheetSourceRelationship => 0,
            Self::UnsupportedSemanticFeature { feature } => {
                0u64.saturating_add(feature.heap_bytes())
            }
        }
    }
}

retained_struct!(PivotTableMetadata {
    name,
    cache_id,
    location,
    row_fields,
    column_fields,
    page_fields,
    data_fields,
    refresh_on_load,
    cache_invalid,
    cache_definition_part,
    cache_source,
    status,
    extension_uris,
    style,
    row_items,
    column_items
});

retained_struct!(PivotTableStyle {
    name,
    show_row_headers,
    show_column_headers,
    show_row_stripes,
    show_column_stripes,
    show_last_column,
    elements
});

retained_struct!(PivotTableStyleElement { kind, size, dxf });

retained_struct!(ShapeAnchor {
    from_col,
    from_col_off,
    from_row,
    from_row_off,
    to_col,
    to_col_off,
    to_row,
    to_row_off,
    edit_as,
    native_ext_cx,
    native_ext_cy,
    anchor_tag,
    anchor_ext_cx,
    anchor_ext_cy,
    shapes
});

/// Fieldless `Copy` enum: no heap storage.
impl RetainedBytes for DrawingAnchorTag {
    fn heap_bytes(&self) -> u64 {
        0
    }
}

impl RetainedBytes for ShapeFill {
    fn heap_bytes(&self) -> u64 {
        match self {
            Self::Solid { color } => 0u64.saturating_add(color.heap_bytes()),
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
            Self::Pattern { fg, bg, preset } => 0u64
                .saturating_add(fg.heap_bytes())
                .saturating_add(bg.heap_bytes())
                .saturating_add(preset.heap_bytes()),
        }
    }
}

impl RetainedBytes for ShapeGeom {
    fn heap_bytes(&self) -> u64 {
        match self {
            Self::Preset { name, adj } => 0u64
                .saturating_add(name.heap_bytes())
                .saturating_add(adj.heap_bytes()),
            Self::Custom { paths } => 0u64.saturating_add(paths.heap_bytes()),
            Self::Image {
                image_path,
                mime_type,
                svg_image_path,
                src_rect,
                alpha,
                duotone,
            } => 0u64
                .saturating_add(image_path.heap_bytes())
                .saturating_add(mime_type.heap_bytes())
                .saturating_add(svg_image_path.heap_bytes())
                .saturating_add(src_rect.heap_bytes())
                .saturating_add(alpha.heap_bytes())
                .saturating_add(duotone.heap_bytes()),
        }
    }
}

retained_struct!(ShapeInfo {
    z_order,
    x,
    y,
    w,
    h,
    rot,
    flip_h,
    flip_v,
    fill_color,
    fill,
    stroke_color,
    stroke_width,
    stroke_fill,
    stroke_dash_style,
    stroke_custom_dash,
    stroke_line_cap,
    stroke_line_join,
    stroke_miter_limit,
    stroke_alignment,
    stroke_cmpd,
    stroke_head_end,
    stroke_tail_end,
    geom,
    text
});

retained_struct!(ShapeLineDashSegment { dash, space });

retained_struct!(ShapeLineEnd { r#type, w, len });

retained_struct!(ShapeParagraph {
    align,
    rtl,
    mar_l,
    mar_r,
    indent,
    def_tab_sz,
    tab_stops,
    space_line,
    space_before,
    space_after,
    runs
});

impl RetainedBytes for ShapeStrokeFill {
    fn heap_bytes(&self) -> u64 {
        match self {
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
            Self::Pattern { fg, bg, preset } => 0u64
                .saturating_add(fg.heap_bytes())
                .saturating_add(bg.heap_bytes())
                .saturating_add(preset.heap_bytes()),
        }
    }
}

retained_struct!(ShapeTabStop { pos, algn });

retained_struct!(ShapeText {
    vert,
    anchor_ctr,
    spc_first_last_para,
    anchor,
    wrap,
    auto_fit,
    font_scale,
    ln_spc_reduction,
    l_ins,
    t_ins,
    r_ins,
    b_ins,
    paragraphs
});

impl RetainedBytes for ShapeTextRun {
    fn heap_bytes(&self) -> u64 {
        match self {
            Self::Text {
                text,
                bold,
                italic,
                size,
                spacing,
                color,
                font_face,
                font_face_ea,
                font_face_cs,
            } => 0u64
                .saturating_add(text.heap_bytes())
                .saturating_add(bold.heap_bytes())
                .saturating_add(italic.heap_bytes())
                .saturating_add(size.heap_bytes())
                .saturating_add(spacing.heap_bytes())
                .saturating_add(color.heap_bytes())
                .saturating_add(font_face.heap_bytes())
                .saturating_add(font_face_ea.heap_bytes())
                .saturating_add(font_face_cs.heap_bytes()),
            Self::Break => 0,
            Self::Math {
                nodes,
                color,
                display: _,
                font_size: _,
            } => {
                // Nonempty OMML trees cannot receive a bounded-copy admission
                // until shared MathNode storage accounting exists. Saturate the
                // estimate so callers cannot admit unmeasured storage. This is
                // RetainedBytes resource policy; parsing, rendering and
                // serialization do not use this estimate to reject math.
                if nodes.is_empty() {
                    color.heap_bytes()
                } else {
                    u64::MAX
                }
            }
        }
    }
}

retained_struct!(TableColumnInfo {
    data_dxf_id,
    header_row_dxf_id,
    totals_row_dxf_id
});

retained_struct!(TableInfo {
    range,
    style_name,
    header_row_count,
    totals_row_count,
    show_row_stripes,
    show_column_stripes,
    show_first_column,
    show_last_column,
    accent_color,
    is_custom,
    whole_table_dxf,
    header_row_dxf,
    total_row_dxf,
    first_column_dxf,
    last_column_dxf,
    band1_horizontal_dxf,
    band2_horizontal_dxf,
    columns
});

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn chart_copies_count_numeric_slots_even_when_json_is_compact() {
        let chart = ooxml_common::chart::ChartModel {
            series: vec![ooxml_common::chart::ChartSeries {
                values: vec![Some(0.0); 100],
                ..Default::default()
            }],
            ..Default::default()
        };
        let lower = 100 * std::mem::size_of::<Option<f64>>()
            + std::mem::size_of::<ooxml_common::chart::ChartSeries>();
        let anchor = ChartAnchor {
            z_order: 0,
            from_col: 0,
            from_col_off: 0,
            from_row: 0,
            from_row_off: 0,
            to_col: 1,
            to_col_off: 0,
            to_row: 1,
            to_row_off: 0,
            chart,
        };
        assert!(anchor.heap_bytes() >= lower as u64);
    }
}
