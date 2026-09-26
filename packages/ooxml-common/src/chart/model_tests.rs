#[cfg(test)]
mod tests {
    use super::super::*;

    #[test]
    fn canonical_chart_type_bar_matrix() {
        // Mirrors the TS `canonicalChartType` bar branch (ST_BarDir "bar" =
        // horizontal). Every (grouping, dir) pair must map to the same string
        // the renderer dispatches on.
        assert_eq!(
            canonical_chart_type("bar", "col", "clustered"),
            "clusteredBar"
        );
        assert_eq!(
            canonical_chart_type("bar", "bar", "clustered"),
            "clusteredBarH"
        );
        assert_eq!(canonical_chart_type("bar", "col", "stacked"), "stackedBar");
        assert_eq!(canonical_chart_type("bar", "bar", "stacked"), "stackedBarH");
        assert_eq!(
            canonical_chart_type("bar", "col", "percentStacked"),
            "stackedBarPct"
        );
        assert_eq!(
            canonical_chart_type("bar", "bar", "percentStacked"),
            "stackedBarHPct"
        );
        // Unknown grouping → clustered fallback (matches the TS default arm).
        assert_eq!(
            canonical_chart_type("bar", "col", "standard"),
            "clusteredBar"
        );
    }

    #[test]
    fn canonical_chart_type_line_area_and_passthrough() {
        assert_eq!(canonical_chart_type("line", "col", "standard"), "line");
        assert_eq!(
            canonical_chart_type("line", "col", "stacked"),
            "stackedLine"
        );
        assert_eq!(
            canonical_chart_type("line", "col", "percentStacked"),
            "stackedLinePct"
        );
        assert_eq!(canonical_chart_type("area", "col", "standard"), "area");
        assert_eq!(
            canonical_chart_type("area", "col", "stacked"),
            "stackedArea"
        );
        assert_eq!(
            canonical_chart_type("area", "col", "percentStacked"),
            "stackedAreaPct"
        );
        // Families the renderer already names canonically pass through verbatim.
        for t in ["pie", "doughnut", "scatter", "bubble", "radar", "waterfall"] {
            assert_eq!(canonical_chart_type(t, "col", "clustered"), t);
        }
    }

    /// The wire contract: a `ChartModel` must serialize with the same camelCase
    /// keys the TS `ChartModel` declares, REQUIRED fields present even when
    /// `None`/`false`/empty, OPTIONAL fields dropped when unset. This is the
    /// Rust-side oracle that pins the emitted JSON shape.
    #[test]
    fn chart_model_serializes_canonical_shape() {
        let m = ChartModel {
            chart_type: "clusteredBar".to_string(),
            title: None,
            title_rich_runs: None,
            title_present: false,
            authored_without_series: false,
            categories: vec!["A".to_string(), "B".to_string()],
            category_source_hidden: None,
            category_levels: None,
            series: vec![ChartSeries {
                name: "S1".to_string(),
                chartex_format_idx: None,
                color: Some("FF0000".to_string()),
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
                chartex_style: None,
                line_color: None,
                line_width_emu: None,
                three_d_shape: None,
                values: vec![Some(1.0), None, Some(3.0)],
                source_hidden: None,
                data_point_colors: None,
                explosion: None,
                data_label_colors: None,
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
                categories: None,
                bubble_x_source_is_string: None,
                show_marker: None,
                val_format_code: None,
                cat_format_code: None,
                cat_format_builtin_id: None,
                cat_format_codes: None,
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
                data_point_overrides: None,
                data_label_overrides: None,
                series_data_labels: None,
                err_bars: None,
                bubble_sizes: None,
                bubble_3d_group_default: None,
                bubble_3d: None,
                smooth: None,
                trend_lines: None,
                line_hidden: None,
            }],
            plot_groups: None,
            vary_colors: None,
            chart_text_boxes: None,
            chart_text_style: None,
            chart_area_style: None,
            plot_area_style: None,
            legend_style: None,
            title_style: None,
            cat_axis_style: None,
            val_axis_style: None,
            cat_axis_title_style: None,
            val_axis_title_style: None,
            cat_axis_major_gridline_style: None,
            cat_axis_minor_gridline_style: None,
            val_axis_major_gridline_style: None,
            val_axis_minor_gridline_style: None,
            show_data_labels: false,
            val_min: None,
            val_max: None,
            cat_axis_title: None,
            val_axis_title: None,
            cat_axis_hidden: false,
            val_axis_hidden: false,
            cat_axis_line_hidden: false,
            val_axis_line_hidden: false,
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
            chart_bg: Some("FFFFFF".to_string()),
            chart_fill: None,
            chart_fill_hidden: None,
            chart_fill_paint_authored: None,
            rounded_corners: None,
            plot_visible_only: None,
            show_legend: false,
            data_table: None,
            legend_pos: None,
            cat_axis_cross_between: "between".to_string(),
            val_axis_major_tick_mark: "out".to_string(),
            cat_axis_major_tick_mark: "out".to_string(),
            title_font_size_hpt: None,
            title_font_color: None,
            title_font_paint_authored: None,
            title_font_face: None,
            cat_axis_font_size_hpt: None,
            val_axis_font_size_hpt: None,
            data_label_font_size_hpt: None,
            subtotal_indices: vec![],
            val_axis_minor_tick_mark: None,
            cat_axis_minor_tick_mark: None,
            cat_axis_font_color: None,
            cat_axis_font_paint_authored: None,
            val_axis_font_color: None,
            val_axis_font_paint_authored: None,
            legend_manual_layout: None,
            legend_overlay: None,
            legend_entries: None,
            val_axis_format_code: None,
            val_axis_number_format: None,
            val_axis_display_units: None,
            cat_axis_display_units: None,
            bar_gap_width: None,
            bar_overlap: None,
            data_label_position: None,
            data_label_font_color: None,
            data_label_font_paint_authored: None,
            data_label_format_code: None,
            data_label_font_bold: None,
            data_label_font_italic: None,
            data_label_font_language: None,
            data_label_font_baseline: None,
            title_font_bold: None,
            title_font_italic: None,
            title_font_language: None,
            title_font_baseline: None,
            cat_axis_font_bold: None,
            cat_axis_font_italic: None,
            val_axis_font_bold: None,
            val_axis_font_italic: None,
            cat_axis_title_font_size_hpt: None,
            cat_axis_title_font_bold: None,
            cat_axis_title_font_italic: None,
            cat_axis_title_font_color: None,
            cat_axis_title_font_paint_authored: None,
            cat_axis_title_rotation: None,
            cat_axis_title_vertical_mode: None,
            cat_axis_title_manual_layout: None,
            cat_axis_title_text_vertical_inset_emu: None,
            val_axis_title_font_size_hpt: None,
            val_axis_title_font_bold: None,
            val_axis_title_font_italic: None,
            val_axis_title_font_color: None,
            val_axis_title_font_paint_authored: None,
            val_axis_title_rotation: None,
            val_axis_title_vertical_mode: None,
            val_axis_title_manual_layout: None,
            val_axis_title_text_vertical_inset_emu: None,
            chart_border_color: None,
            chart_border_line_fill: None,
            chart_border_width_emu: None,
            chart_border_dash: None,
            chart_border_dash_authored: None,
            chart_border_custom_dash: None,
            chart_border_cap: None,
            chart_border_join: None,
            chart_border_compound: None,
            chart_border_hidden: None,
            chart_border_paint_authored: None,
            cat_axis_crosses: None,
            cat_axis_crosses_at: None,
            val_axis_crosses: None,
            val_axis_crosses_at: None,
            cat_axis_line_color: None,
            cat_axis_line_width_emu: None,
            cat_axis_line_dash: None,
            cat_axis_line_paint_authored: None,
            val_axis_line_color: None,
            val_axis_line_width_emu: None,
            val_axis_line_dash: None,
            val_axis_line_paint_authored: None,
            cat_axis_format_code: None,
            cat_axis_number_format: None,
            cat_axis_min: None,
            cat_axis_max: None,
            title_manual_layout: None,
            plot_area_manual_layout: None,
            cartesian_auto_layout_profile: None,
            scatter_style: None,
            bubble_scale: None,
            bubble_size_represents: None,
            show_negative_bubbles: None,
            radar_style: None,
            secondary_val_axis: None,
            secondary_cat_axis: None,
            hole_size: None,
            first_slice_angle: None,
            cat_axis_font_face: None,
            val_axis_font_face: None,
            cat_axis_title_font_face: None,
            val_axis_title_font_face: None,
            data_label_font_face: None,
            legend_font_face: None,
            legend_font_color: None,
            legend_font_paint_authored: None,
            legend_font_size_hpt: None,
            legend_font_bold: None,
            legend_font_italic: None,
            legend_font_language: None,
            legend_font_baseline: None,
            legend_fill_color: None,
            legend_fill: None,
            legend_fill_hidden: None,
            legend_fill_paint_authored: None,
            legend_line_color: None,
            legend_line_fill: None,
            legend_line_width_emu: None,
            legend_line_dash: None,
            legend_line_dash_authored: None,
            legend_line_custom_dash: None,
            legend_line_cap: None,
            legend_line_join: None,
            legend_line_compound: None,
            legend_line_hidden: None,
            legend_line_paint_authored: None,
            theme_major_font_latin: None,
            theme_minor_font_latin: None,
            date1904: false,
            disp_blanks_as: None,
            show_data_labels_over_max: None,
            val_axis_major_gridlines: None,
            cat_axis_major_gridlines: None,
            val_axis_gridline_color: None,
            val_axis_gridline_width_emu: None,
            val_axis_gridline_dash: None,
            val_axis_gridline_paint_authored: None,
            cat_axis_gridline_color: None,
            cat_axis_gridline_width_emu: None,
            cat_axis_gridline_dash: None,
            cat_axis_gridline_paint_authored: None,
            val_axis_minor_gridlines: None,
            val_axis_minor_gridline_color: None,
            val_axis_minor_gridline_width_emu: None,
            val_axis_minor_gridline_dash: None,
            val_axis_minor_gridline_paint_authored: None,
            cat_axis_minor_gridlines: None,
            cat_axis_minor_gridline_color: None,
            cat_axis_minor_gridline_width_emu: None,
            cat_axis_minor_gridline_dash: None,
            cat_axis_minor_gridline_paint_authored: None,
            val_axis_major_unit: None,
            val_axis_minor_unit: None,
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
            chartex_box: None,
            chartex_sunburst: None,
            chartex_treemap: None,
            chartex_region_map: None,
            chartex_histogram_binning: None,
            chartex_accents: None,
            chart_style_roles: None,
            classic_chart_style_roles: None,
            classic_varying_point_chart_style_roles: None,
            classic_varying_point_chart_style_roles_by_group: None,
            chart_style_color_palette: None,
            chart_style_color_method: None,
            chart_style_marker_size_pt: None,
            chart_style_marker_symbol: None,
            chartex_color_palette: None,
            chartex_color_style_method: None,
            chartex_data_point_style: None,
            chartex_data_point_line_style: None,
            chartex_series_line_style: None,
            chartex_data_point_marker_style: None,
            chartex_marker_size_pt: None,
            chartex_marker_symbol: None,
            chartex_connector_lines: None,
        };
        let v = serde_json::to_value(&m).unwrap();
        let obj = v.as_object().unwrap();
        // Required scalar keys present with camelCase names, even when None/false.
        assert_eq!(obj["chartType"], "clusteredBar");
        assert!(obj["title"].is_null());
        assert_eq!(obj["showDataLabels"], false);
        assert_eq!(obj["catAxisHidden"], false);
        assert_eq!(obj["catAxisCrossBetween"], "between");
        assert_eq!(obj["valAxisMajorTickMark"], "out");
        assert!(obj["plotAreaBg"].is_null());
        assert_eq!(obj["chartBg"], "FFFFFF");
        assert_eq!(obj["subtotalIndices"], serde_json::json!([]));
        // Optional unset keys dropped from the wire.
        assert!(!obj.contains_key("barGapWidth"));
        assert!(!obj.contains_key("secondaryValAxis"));
        assert!(!obj.contains_key("catAxisFontColor"));
        // date1904 is dropped from the wire when false (default 1900 system).
        assert!(!obj.contains_key("date1904"));
        // Series: required present, optional dropped; array null preserved.
        let s0 = &obj["series"][0];
        assert_eq!(s0["name"], "S1");
        assert_eq!(s0["color"], "FF0000");
        assert_eq!(s0["values"], serde_json::json!([1.0, null, 3.0]));
        assert!(!s0.as_object().unwrap().contains_key("showMarker"));
        // Round-trips back to an equal model (Deserialize parity).
        let back: ChartModel = serde_json::from_value(v).unwrap();
        assert_eq!(back, m);
    }

    #[test]
    fn parse_chart_part_wires_rounded_chart_space_gradient_to_the_shared_model() {
        let document = root_of(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:roundedCorners/><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart><c:spPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="112233"/></a:gs><a:gs pos="100000"><a:srgbClr val="AABBCC"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill></c:spPr></c:chartSpace>"#,
        );
        let model = parse_chart_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("classic chart parses");
        assert_eq!(model.rounded_corners, Some(true));
        assert!(matches!(
            model.chart_fill.as_ref(),
            Some(ChartStyleFill::Gradient { .. })
        ));
        assert_eq!(model.chart_fill_hidden, None);
        assert_eq!(model.chart_fill_paint_authored, Some(true));
    }
}
