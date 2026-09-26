#[cfg(test)]
mod tests {
    use super::super::*;

    #[test]
    fn bar_gap_overlap_default_to_none() {
        let xml =
            r#"<c:barChart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#;
        let d = root_of(xml);
        assert_eq!(extract_bar_gap_overlap(d.root_element()), (None, None));
    }

    #[test]
    fn bar_gap_overlap_explicit() {
        let xml = r#"<c:barChart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:gapWidth val="50"/>
            <c:overlap val="100"/>
        </c:barChart>"#;
        let d = root_of(xml);
        assert_eq!(
            extract_bar_gap_overlap(d.root_element()),
            (Some(50), Some(100))
        );
    }

    #[test]
    fn bar_gap_overlap_reject_schema_out_of_range_values() {
        let xml = r#"<c:barChart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:gapWidth val="501"/>
            <c:overlap val="101"/>
        </c:barChart>"#;
        let d = root_of(xml);
        assert_eq!(extract_bar_gap_overlap(d.root_element()), (None, None));
    }

    #[test]
    fn parse_chart_part_preserves_each_bar_groups_geometry() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:chart><c:plotArea>
              <c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
                <c:ser><c:idx val="0"/><c:tx><c:v>Primary</c:v></c:tx>
                  <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:pt idx="0"><c:v>10</c:v></c:pt></c:numLit></c:val></c:ser>
                <c:gapWidth val="150"/><c:axId val="1"/><c:axId val="2"/>
              </c:barChart>
              <c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
                <c:ser><c:idx val="1"/><c:tx><c:v>Overlay</c:v></c:tx>
                  <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>
                <c:gapWidth val="0"/><c:overlap val="100"/><c:axId val="1"/><c:axId val="3"/>
              </c:barChart>
              <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
              <c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
              <c:valAx><c:axId val="3"/><c:axPos val="r"/><c:crossAx val="1"/></c:valAx>
            </c:plotArea></c:chart></c:chartSpace>"#
        );
        let document = chart_space_of(&xml);
        let model = parse_chart_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("chart");

        assert_eq!(model.series[0].bar_group_index, Some(0));
        assert_eq!(model.series[0].bar_group_direction.as_deref(), Some("col"));
        assert_eq!(
            model.series[0].bar_group_grouping.as_deref(),
            Some("clustered")
        );
        assert_eq!(model.series[0].bar_group_gap_width, Some(150));
        assert_eq!(model.series[0].bar_group_overlap, None);
        assert_eq!(model.series[1].bar_group_index, Some(1));
        assert_eq!(model.series[1].bar_group_direction.as_deref(), Some("col"));
        assert_eq!(
            model.series[1].bar_group_grouping.as_deref(),
            Some("clustered")
        );
        assert_eq!(model.series[1].bar_group_gap_width, Some(0));
        assert_eq!(model.series[1].bar_group_overlap, Some(100));
    }

    #[test]
    fn series_smooth_present_and_absent() {
        // No `<c:smooth>` → None (straight-polyline default).
        let none =
            root_of(r#"<c:ser xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#);
        assert_eq!(extract_series_smooth(none.root_element()), None);
        // `<c:smooth val="1"/>` → Some(true).
        let on = root_of(
            r#"<c:ser xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:smooth val="1"/></c:ser>"#,
        );
        assert_eq!(extract_series_smooth(on.root_element()), Some(true));
        // `<c:smooth val="0"/>` → Some(false) (explicit off).
        let off = root_of(
            r#"<c:ser xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:smooth val="0"/></c:ser>"#,
        );
        assert_eq!(extract_series_smooth(off.root_element()), Some(false));
        // Bare `<c:smooth/>` → Some(true) (CT_Boolean implied-true).
        let bare = root_of(
            r#"<c:ser xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:smooth/></c:ser>"#,
        );
        assert_eq!(extract_series_smooth(bare.root_element()), Some(true));
    }

    #[test]
    fn disp_blanks_as_variants() {
        // Absent element → None (renderer defaults to "gap").
        let absent = root_of(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart/></c:chartSpace>"#,
        );
        assert_eq!(extract_disp_blanks_as(absent.root_element()), None);
        // Explicit values pass through.
        for want in ["gap", "zero", "span"] {
            let xml = format!(
                r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:dispBlanksAs val="{want}"/></c:chart></c:chartSpace>"#,
            );
            assert_eq!(
                extract_disp_blanks_as(root_of(&xml).root_element()).as_deref(),
                Some(want)
            );
        }
        // Bare `<c:dispBlanksAs/>` → XSD @val default "zero".
        let bare = root_of(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:dispBlanksAs/></c:chart></c:chartSpace>"#,
        );
        assert_eq!(
            extract_disp_blanks_as(bare.root_element()).as_deref(),
            Some("zero")
        );
    }

    #[test]
    fn parse_chart_user_shapes_applies_drawingml_text_inset_defaults() {
        let doc = roxmltree::Document::parse(
            r#"<c:userShapes xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                 xmlns:cdr="http://schemas.openxmlformats.org/drawingml/2006/chartDrawing"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <cdr:relSizeAnchor>
                <cdr:from><cdr:x>0</cdr:x><cdr:y>0</cdr:y></cdr:from>
                <cdr:to><cdr:x>1</cdr:x><cdr:y>0.1</cdr:y></cdr:to>
                <cdr:sp><cdr:txBody><a:bodyPr/><a:p><a:r><a:t>Title</a:t></a:r></a:p></cdr:txBody></cdr:sp>
              </cdr:relSizeAnchor>
            </c:userShapes>"#,
        )
        .unwrap();

        let boxes = parse_chart_user_shapes(doc.root_element(), &StubResolver);
        assert_eq!(boxes.len(), 1);
        assert_eq!(boxes[0].l_ins, crate::text::DEFAULT_INS_LR_EMU);
        assert_eq!(boxes[0].r_ins, crate::text::DEFAULT_INS_LR_EMU);
        assert_eq!(boxes[0].t_ins, crate::text::DEFAULT_INS_TB_EMU);
        assert_eq!(boxes[0].b_ins, crate::text::DEFAULT_INS_TB_EMU);
    }

    #[test]
    fn chart_text_boolean_carriers_preserve_inheritable_metadata_and_explicit_states() {
        let xml = format!(
            r#"<c:chart xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="en" altLang="ja-JP"/><a:t>T</a:t></a:r></a:p></c:rich></c:tx></c:title>
              <c:txPr><a:p><a:pPr><a:defRPr lang="en"/></a:pPr></a:p></c:txPr>
              <c:legend><c:txPr><a:p><a:pPr><a:defRPr altLang="ja-JP"/></a:pPr></a:p></c:txPr></c:legend>
              <c:plotArea><c:barChart><c:dLbls><c:txPr><a:p><a:pPr><a:defRPr dirty="0"/></a:pPr></a:p></c:txPr></c:dLbls></c:barChart></c:plotArea>
            </c:chart>"#
        );
        let document = root_of(&xml);
        let root = document.root_element();
        assert_eq!(extract_chart_title_bold(root), None);
        assert_eq!(extract_chart_title_italic(root), None);
        assert_eq!(extract_axis_tick_label_bold(root), None);
        assert_eq!(extract_axis_tick_label_italic(root), None);
        assert_eq!(extract_legend_text_props(root).2, None);
        assert_eq!(extract_legend_text_props(root).3, None);
        assert_eq!(extract_data_label_font_bold(root), None);
        assert_eq!(extract_data_label_font_italic(root), None);

        let title = child(root, "title").unwrap();
        let runs = parse_chart_title_rich_runs(title, &StubResolver).unwrap();
        assert_eq!((runs[0].bold, runs[0].italic), (None, None));

        let absent = format!(
            r#"<c:chart xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:title><c:tx><c:rich><a:p><a:r><a:t>T</a:t></a:r></a:p></c:rich></c:tx></c:title>
              <c:txPr><a:p/></c:txPr><c:legend><c:txPr><a:p/></c:txPr></c:legend>
              <c:plotArea><c:barChart><c:dLbls><c:txPr><a:p/></c:txPr></c:dLbls></c:barChart></c:plotArea>
            </c:chart>"#
        );
        let document = root_of(&absent);
        let root = document.root_element();
        assert_eq!(extract_chart_title_bold(root), None);
        assert_eq!(extract_chart_title_italic(root), None);
        assert_eq!(extract_axis_tick_label_bold(root), None);
        assert_eq!(extract_axis_tick_label_italic(root), None);
        assert_eq!(extract_legend_text_props(root).2, None);
        assert_eq!(extract_legend_text_props(root).3, None);
        assert_eq!(extract_data_label_font_bold(root), None);
        assert_eq!(extract_data_label_font_italic(root), None);

        let explicit_xml = format!(
            r#"<c:valAx xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:txPr><a:p><a:pPr><a:defRPr b="1" i="1"/></a:pPr></a:p></c:txPr></c:valAx>"#
        );
        let explicit = root_of(&explicit_xml);
        assert_eq!(
            extract_axis_tick_label_bold(explicit.root_element()),
            Some(true)
        );
        assert_eq!(
            extract_axis_tick_label_italic(explicit.root_element()),
            Some(true)
        );
    }

    #[test]
    fn chart_text_font_sizes_enforce_st_text_font_size_boundaries() {
        for (size, expected) in [
            ("99", None),
            ("100", Some(100)),
            ("400000", Some(400_000)),
            ("400001", None),
        ] {
            let xml = format!(
                r#"<c:chart xmlns:c="{C_NS}" xmlns:a="{A_NS}">
                  <c:title><c:tx><c:rich><a:p><a:r><a:rPr sz="{size}"/><a:t>T</a:t></a:r></a:p></c:rich></c:tx></c:title>
                  <c:txPr><a:p><a:pPr><a:defRPr sz="{size}"/></a:pPr></a:p></c:txPr>
                  <c:legend><c:txPr><a:p><a:pPr><a:defRPr sz="{size}"/></a:pPr></a:p></c:txPr></c:legend>
                </c:chart>"#
            );
            let document = root_of(&xml);
            let root = document.root_element();
            assert_eq!(extract_chart_title_size(root), expected, "title {size}");
            assert_eq!(extract_axis_tick_label_size(root), expected, "axis {size}");
            assert_eq!(extract_legend_text_props(root).1, expected, "legend {size}");
        }
    }

    #[test]
    fn chart_space_border_solid() {
        let xml = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="1B4332"/></a:solidFill></a:ln></c:spPr>
        </c:chartSpace>"#;
        let d = root_of(xml);
        assert_eq!(
            extract_chart_space_border(d.root_element()),
            (Some("1B4332".to_string()), Some(19050))
        );
    }

    #[test]
    fn chart_space_border_absent() {
        let xml =
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#;
        let d = root_of(xml);
        assert_eq!(extract_chart_space_border(d.root_element()), (None, None));
    }

    #[test]
    fn chart_space_rounded_corners_preserve_omission_bare_and_explicit_values() {
        let absent = root_of(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#,
        );
        assert_eq!(
            extract_chart_space_rounded_corners(absent.root_element()),
            None
        );
        let bare = root_of(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:roundedCorners/></c:chartSpace>"#,
        );
        assert_eq!(
            extract_chart_space_rounded_corners(bare.root_element()),
            Some(true)
        );
        let explicit_false = root_of(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:roundedCorners val="false"/></c:chartSpace>"#,
        );
        assert_eq!(
            extract_chart_space_rounded_corners(explicit_false.root_element()),
            Some(false)
        );
    }

    #[test]
    fn direct_chart_frame_line_retains_every_compound_kind() {
        for compound in ["sng", "dbl", "thinThick", "thickThin", "tri"] {
            let xml = format!(
                r#"<c:plotArea xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:spPr><a:ln cmpd="{compound}"><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:ln></c:spPr></c:plotArea>"#,
            );
            let document = root_of(&xml);
            let line = extract_direct_shape_line(document.root_element(), &FixtureResolver);
            assert_eq!(line.compound.as_deref(), Some(compound));
        }
    }

    #[test]
    fn chart_date1904_variants() {
        // §21.2.2.38: CT_Boolean. Element present + val omitted ⇒ true.
        let ns = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        let bare = format!(r#"<c:chartSpace xmlns:c="{ns}"><c:date1904/></c:chartSpace>"#);
        assert!(extract_chart_date1904(root_of(&bare).root_element()));

        let one = format!(r#"<c:chartSpace xmlns:c="{ns}"><c:date1904 val="1"/></c:chartSpace>"#);
        assert!(extract_chart_date1904(root_of(&one).root_element()));

        let word =
            format!(r#"<c:chartSpace xmlns:c="{ns}"><c:date1904 val="true"/></c:chartSpace>"#);
        assert!(extract_chart_date1904(root_of(&word).root_element()));

        let zero = format!(r#"<c:chartSpace xmlns:c="{ns}"><c:date1904 val="0"/></c:chartSpace>"#);
        assert!(!extract_chart_date1904(root_of(&zero).root_element()));

        // Word form of the falsey value: `val="false"` also disables the 1904
        // system (CT_Boolean accepts both "0" and "false").
        let false_word =
            format!(r#"<c:chartSpace xmlns:c="{ns}"><c:date1904 val="false"/></c:chartSpace>"#);
        assert!(!extract_chart_date1904(root_of(&false_word).root_element()));

        // Absent element ⇒ false (default 1900 system).
        let absent = format!(r#"<c:chartSpace xmlns:c="{ns}"/>"#);
        assert!(!extract_chart_date1904(root_of(&absent).root_element()));
    }

    #[test]
    fn hole_size_from_doughnut() {
        let xml = format!(
            r#"<c:chart xmlns:c="{C_NS}"><c:plotArea><c:doughnutChart><c:holeSize val="60"/></c:doughnutChart></c:plotArea></c:chart>"#
        );
        assert_eq!(extract_hole_size(root_of(&xml).root_element()), Some(60));
        // Clamped to the ECMA 1–90 range.
        let hi = format!(
            r#"<c:chart xmlns:c="{C_NS}"><c:doughnutChart><c:holeSize val="200"/></c:doughnutChart></c:chart>"#
        );
        assert_eq!(extract_hole_size(root_of(&hi).root_element()), Some(90));
        // A pie chart has no hole → None even if a stray holeSize appears elsewhere.
        let pie = format!(r#"<c:chart xmlns:c="{C_NS}"><c:pieChart/></c:chart>"#);
        assert_eq!(extract_hole_size(root_of(&pie).root_element()), None);
    }

    #[test]
    fn first_slice_angle_from_pie_or_doughnut() {
        let pie = format!(
            r#"<c:chart xmlns:c="{C_NS}"><c:pieChart><c:firstSliceAng val="90"/></c:pieChart></c:chart>"#
        );
        assert_eq!(
            extract_first_slice_angle(root_of(&pie).root_element()),
            Some(90)
        );
        let dn = format!(
            r#"<c:chart xmlns:c="{C_NS}"><c:doughnutChart><c:firstSliceAng val="270"/></c:doughnutChart></c:chart>"#
        );
        assert_eq!(
            extract_first_slice_angle(root_of(&dn).root_element()),
            Some(270)
        );
        // Absent ⇒ None (renderer defaults to 0).
        let none = format!(r#"<c:chart xmlns:c="{C_NS}"><c:pieChart/></c:chart>"#);
        assert_eq!(
            extract_first_slice_angle(root_of(&none).root_element()),
            None
        );
    }

    #[test]
    fn dpt_explosion() {
        let with =
            format!(r#"<c:dPt xmlns:c="{C_NS}"><c:idx val="1"/><c:explosion val="25"/></c:dPt>"#);
        assert_eq!(
            extract_dpt_explosion(root_of(&with).root_element()),
            Some(25)
        );
        let without = format!(r#"<c:dPt xmlns:c="{C_NS}"><c:idx val="1"/></c:dPt>"#);
        assert_eq!(
            extract_dpt_explosion(root_of(&without).root_element()),
            None
        );
    }

    #[test]
    fn pie_series_explosion_is_preserved_as_the_point_default() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:chart><c:plotArea><c:pieChart><c:varyColors val="1"/>
                <c:ser><c:idx val="0"/><c:order val="0"/><c:explosion val="35"/>
                  <c:val><c:numLit><c:ptCount val="2"/>
                    <c:pt idx="0"><c:v>10</c:v></c:pt>
                    <c:pt idx="1"><c:v>20</c:v></c:pt>
                  </c:numLit></c:val>
                </c:ser>
              </c:pieChart></c:plotArea></c:chart>
            </c:chartSpace>"#
        );
        let model = parse_chart_part(
            chart_space_of(&xml).root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("pie chart parses");
        assert_eq!(model.series[0].explosion, Some(35));
    }

    #[test]
    fn chart_group_separator_projection_uses_one_chart_wide_budget() {
        assert!(chart_group_separator_projection_within_budget([
            (MAX_DATA_LABEL_RICH_SCALARS, 128),
            (MAX_DATA_LABEL_RICH_SCALARS, 128),
        ]));
        assert!(!chart_group_separator_projection_within_budget([
            (MAX_DATA_LABEL_RICH_SCALARS, 128),
            (MAX_DATA_LABEL_RICH_SCALARS, 128),
            (1, 1),
        ]));
        assert!(!chart_group_separator_projection_within_budget([
            (MAX_CHART_CACHE_POINTS, 1),
            (1, 1),
        ]));
    }

    #[test]
    fn chart_group_separator_projection_fails_closed_atomically_in_parser() {
        let separator = "s".repeat(MAX_DATA_LABEL_RICH_SCALARS);
        let group = |start: usize, count: usize, separator: &str| {
            let series = (start..start + count)
                .map(|index| {
                    format!(
                        r#"<c:ser><c:idx val="{index}"/><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>"#,
                    )
                })
                .collect::<String>();
            format!(
                r#"<c:barChart><c:barDir val="col"/>{series}<c:dLbls><c:showVal/><c:separator>{separator}</c:separator></c:dLbls></c:barChart>"#,
            )
        };
        let exact_groups = format!(
            "{}{}",
            group(0, 128, &separator),
            group(128, 128, &separator),
        );
        let exact_xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea>{exact_groups}</c:plotArea></c:chart></c:chartSpace>"#,
        );
        let exact = parse_chart_part(
            chart_space_of(&exact_xml).root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("exact projection parses");
        assert_eq!(
            exact.series[0]
                .series_data_labels
                .as_ref()
                .and_then(|labels| labels.separator.as_ref())
                .map(|value| value.chars().count()),
            Some(MAX_DATA_LABEL_RICH_SCALARS),
        );
        assert_eq!(
            exact.series[255]
                .series_data_labels
                .as_ref()
                .and_then(|labels| labels.separator.as_ref())
                .map(|value| value.chars().count()),
            Some(MAX_DATA_LABEL_RICH_SCALARS),
        );

        let overflow_groups = format!("{exact_groups}{}", group(256, 1, "x"));
        let overflow_xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea>{overflow_groups}</c:plotArea></c:chart></c:chartSpace>"#,
        );
        let overflow = parse_chart_part(
            chart_space_of(&overflow_xml).root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("overflow projection parses");
        assert!(overflow.series.iter().all(|series| series
            .series_data_labels
            .as_ref()
            .is_some_and(|labels| labels.separator.is_none())));
    }

    #[test]
    fn chart_level_smooth_is_the_line_series_default() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:chart>
                <c:plotArea><c:lineChart>
                  <c:grouping val="standard"/><c:smooth/>
                  <c:ser><c:idx val="0"/><c:order val="0"/>
                    <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
                  </c:ser>
                  <c:ser><c:idx val="1"/><c:order val="1"/><c:smooth val="0"/>
                    <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val>
                  </c:ser>
                </c:lineChart></c:plotArea>
              </c:chart></c:chartSpace>"#
        );
        let model = parse_chart_part(
            chart_space_of(&xml).root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("line chart parses");
        assert_eq!(model.series[0].smooth, Some(true));
        assert_eq!(model.series[1].smooth, Some(false));
    }

    /// (a) Bar chart with the full decoration set: title (size/bold/color),
    /// legend, styled category + value axes, gap/overlap, chartSpace border,
    /// and value-axis major gridlines. Every field asserted here is a distinct
    /// probe `parse_chart_part` wires up; a regression in any one shows here
    /// without needing a full-document golden diff.
    #[test]
    fn parse_chart_part_bar_full_decoration() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:chart>
                <c:title><c:tx><c:rich>
                  <a:p><a:pPr><a:defRPr sz="1800" b="1"><a:solidFill><a:srgbClr val="1b4332"/></a:solidFill></a:defRPr></a:pPr>
                  <a:r><a:t>Quarterly Revenue</a:t></a:r></a:p>
                </c:rich></c:tx></c:title>
                <c:plotArea>
                  <c:barChart>
                    <c:barDir val="col"/>
                    <c:grouping val="clustered"/>
                    <c:gapWidth val="80"/>
                    <c:overlap val="-10"/>
                    <c:ser>
                      <c:idx val="0"/>
                      <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx>
                      <c:spPr><a:solidFill><a:srgbClr val="2d6a4f"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="595959"/></a:solidFill></a:ln></c:spPr>
                      <c:cat><c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:cat>
                      <c:val><c:numCache><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numCache></c:val>
                    </c:ser>
                    <c:axId val="1"/>
                    <c:axId val="2"/>
                  </c:barChart>
                  <c:catAx>
                    <c:axId val="1"/>
                    <c:axPos val="b"/>
                    <c:title><c:tx><c:rich><a:bodyPr vert="horz"/><a:p><a:r><a:t>Quarter</a:t></a:r></a:p></c:rich></c:tx></c:title>
                    <c:spPr><a:ln><a:solidFill><a:srgbClr val="808080"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr>
                  </c:catAx>
                  <c:valAx>
                    <c:axId val="2"/>
                    <c:axPos val="l"/>
                    <c:title><c:tx><c:rich><a:bodyPr rot="-1800000"/><a:p><a:r><a:t>Revenue</a:t></a:r></a:p></c:rich></c:tx>
                      <c:layout><c:manualLayout><c:xMode val="edge"/><c:yMode val="edge"/><c:x val="0.2"/><c:y val="0.1"/></c:manualLayout></c:layout>
                    </c:title>
                    <c:majorGridlines><c:spPr><a:ln w="3175"><a:solidFill><a:schemeClr val="accent3"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>
                    <c:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="404040"/></a:solidFill><a:prstDash val="dot"/></a:ln></c:spPr>
                    <c:scaling><c:min val="0"/><c:max val="30"/></c:scaling>
                  </c:valAx>
                </c:plotArea>
                <c:legend><c:legendPos val="b"/></c:legend>
              </c:chart>
              <c:spPr><a:ln w="19050" cap="rnd"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:prstDash val="lgDashDotDot"/><a:miter/></a:ln></c:spPr>
            </c:chartSpace>"#
        );
        let doc = chart_space_of(&xml);
        let m = parse_chart_part(
            doc.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("bar chart parses");

        assert_eq!(m.chart_type, "clusteredBar");
        assert_eq!(m.title.as_deref(), Some("Quarterly Revenue"));
        assert_eq!(m.title_font_size_hpt, Some(1800));
        assert_eq!(m.title_font_bold, Some(true));
        // `parse_chart_part` now resolves the title's `<a:solidFill>` via the
        // `ColorResolver` (schemeClr resolution was added — see
        // `extract_chart_title_color`). The fixture's title carries
        // `<a:srgbClr val="1b4332">`, which the resolver returns uppercased.
        // (This assertion previously pinned `None` as a known limitation; the
        // limitation is now fixed, so the expected value flips to the resolved
        // hex — a deliberate, visible contract change.)
        assert_eq!(m.title_font_color.as_deref(), Some("1B4332"));
        assert_eq!(m.categories, vec!["Q1".to_string(), "Q2".to_string()]);
        assert_eq!(m.series.len(), 1);
        assert_eq!(m.series[0].name, "Revenue");
        assert_eq!(m.series[0].values, vec![Some(10.0), Some(20.0)]);
        assert_eq!(m.series[0].color.as_deref(), Some("2D6A4F"));
        assert_eq!(m.series[0].line_color.as_deref(), Some("595959"));
        assert_eq!(m.series[0].line_width_emu, Some(12700));
        assert!(m.show_legend);
        assert_eq!(m.legend_pos.as_deref(), Some("b"));
        assert_eq!(m.bar_gap_width, Some(80));
        assert_eq!(m.bar_overlap, Some(-10));
        assert_eq!(m.val_min, Some(0.0));
        assert_eq!(m.val_max, Some(30.0));
        assert_eq!(m.val_axis_major_gridlines, Some(true));
        assert_eq!(m.cat_axis_major_gridlines, Some(false));
        // The value-axis `<c:majorGridlines><c:spPr><a:ln>` carries an explicit
        // `accent3` colour (resolver → A5A5A5) and a 3175 EMU (0.25 pt) width.
        assert_eq!(m.val_axis_gridline_color.as_deref(), Some("A5A5A5"));
        assert_eq!(m.val_axis_gridline_width_emu, Some(3175));
        // The category axis has no gridlines element → no gridline style.
        assert_eq!(m.cat_axis_gridline_color, None);
        assert_eq!(m.cat_axis_gridline_width_emu, None);
        assert_eq!(m.cat_axis_line_color.as_deref(), Some("808080"));
        assert_eq!(m.cat_axis_line_dash.as_deref(), Some("dash"));
        assert_eq!(m.cat_axis_line_paint_authored, Some(true));
        assert_eq!(m.val_axis_line_color.as_deref(), Some("404040"));
        assert_eq!(m.val_axis_line_width_emu, Some(12700));
        assert_eq!(m.val_axis_line_dash.as_deref(), Some("dot"));
        assert_eq!(m.val_axis_line_paint_authored, Some(true));
        // The chartSpace border is theme-aware: scheme tx1 resolves through
        // the same color resolver as other DrawingML lines.
        assert_eq!(m.chart_border_color.as_deref(), Some("000000"));
        assert_eq!(m.chart_border_width_emu, Some(19050));
        assert_eq!(m.chart_border_dash.as_deref(), Some("lgDashDotDot"));
        assert_eq!(m.chart_border_cap.as_deref(), Some("rnd"));
        assert_eq!(m.chart_border_join.as_deref(), Some("miter"));
        assert!(!m.cat_axis_hidden);
        assert!(!m.val_axis_hidden);
        assert_eq!(m.cat_axis_title_rotation, None);
        assert_eq!(m.cat_axis_title_vertical_mode.as_deref(), Some("horz"));
        assert_eq!(m.cat_axis_title_text_vertical_inset_emu, Some(91_440));
        assert_eq!(m.val_axis_title_rotation, Some(-1_800_000));
        assert_eq!(m.val_axis_title_text_vertical_inset_emu, Some(91_440));
        let val_title_layout = m
            .val_axis_title_manual_layout
            .expect("value-axis title manual layout");
        assert_eq!(val_title_layout.x, 0.2);
        assert_eq!(val_title_layout.y, 0.1);
    }

    #[test]
    fn parse_chart_part_two_scatter_groups_preserves_secondary_xy_axes() {
        let scatter = |idx: usize, x_axis: u32, y_axis: u32, x: f64, y: f64| {
            format!(
                r#"<c:scatterChart><c:scatterStyle val="marker"/>
                  <c:ser><c:idx val="{idx}"/>
                    <c:xVal><c:numLit><c:pt idx="0"><c:v>{x}</c:v></c:pt></c:numLit></c:xVal>
                    <c:yVal><c:numLit><c:pt idx="0"><c:v>{y}</c:v></c:pt></c:numLit></c:yVal>
                  </c:ser><c:axId val="{x_axis}"/><c:axId val="{y_axis}"/>
                </c:scatterChart>"#,
            )
        };
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:chart><c:plotArea>
              {}{}
              <c:valAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/>
                <c:dispUnits><c:builtInUnit val="thousands"/></c:dispUnits></c:valAx>
              <c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/>
                <c:dispUnits><c:builtInUnit val="hundreds"/></c:dispUnits></c:valAx>
              <c:valAx><c:axId val="3"/><c:axPos val="t"/><c:crossAx val="4"/>
                <c:dispUnits><c:custUnit val="1000"/></c:dispUnits></c:valAx>
              <c:valAx><c:axId val="4"/><c:axPos val="r"/><c:crossAx val="3"/>
                <c:dispUnits><c:custUnit val="10"/></c:dispUnits></c:valAx>
            </c:plotArea></c:chart></c:chartSpace>"#,
            scatter(0, 1, 2, 1000.0, 100.0),
            scatter(1, 3, 4, 2000.0, 20.0),
        );
        let doc = chart_space_of(&xml);
        let model = parse_chart_part(
            doc.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("dual scatter groups parse");

        assert_eq!(model.series[0].use_secondary_axis, None);
        assert_eq!(model.series[1].use_secondary_axis, Some(true));
        assert_eq!(
            model
                .secondary_cat_axis
                .as_ref()
                .and_then(|axis| axis.display_units.as_ref())
                .map(|units| units.divisor),
            Some(1000.0),
        );
        assert_eq!(
            model
                .secondary_val_axis
                .as_ref()
                .and_then(|axis| axis.display_units.as_ref())
                .map(|units| units.divisor),
            Some(10.0),
        );
    }

    /// ECMA-376 Part 1 `c:dTable` (`CT_DTable`): the data table is
    /// authored on `c:plotArea`, independently from the legend and axes. Its
    /// four CT_Boolean border/key switches and DrawingML text defaults must
    /// survive the shared parser so DOCX/XLSX/PPTX charts use one model.
    #[test]
    fn parse_chart_part_data_table_properties() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:chart><c:plotArea>
                <c:barChart>
                  <c:barDir val="col"/><c:grouping val="clustered"/>
                  <c:ser><c:idx val="0"/>
                    <c:tx><c:v>Sales</c:v></c:tx>
                    <c:cat><c:strCache><c:pt idx="0"><c:v>Jan</c:v></c:pt></c:strCache></c:cat>
                    <c:val><c:numCache><c:pt idx="0"><c:v>10</c:v></c:pt></c:numCache></c:val>
                  </c:ser>
                </c:barChart>
                <c:catAx><c:axPos val="b"/></c:catAx>
                <c:valAx><c:axPos val="l"/></c:valAx>
                <c:dTable>
                  <c:showHorzBorder/>
                  <c:showVertBorder val="0"/>
                  <c:showOutline val="1"/>
                  <c:showKeys val="1"/>
                  <c:spPr><a:solidFill><a:srgbClr val="FFF2CC"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="445566"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr>
                  <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000" b="1" i="1"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="Aptos"/></a:defRPr></a:pPr></a:p></c:txPr>
                </c:dTable>
              </c:plotArea></c:chart>
            </c:chartSpace>"#
        );
        let model = parse_chart_part(
            chart_space_of(&xml).root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("data-table chart parses");
        let table = model.data_table.expect("c:dTable preserved");
        assert!(table.show_horizontal_border);
        assert!(!table.show_vertical_border);
        assert!(table.show_outline);
        assert!(table.show_keys);
        assert_eq!(table.font_size_hpt, Some(1000));
        assert_eq!(table.font_face.as_deref(), Some("Aptos"));
        assert_eq!(table.font_color.as_deref(), Some("112233"));
        assert_eq!(table.font_bold, Some(true));
        assert_eq!(table.font_italic, Some(true));
        assert_eq!(table.fill_color.as_deref(), Some("FFF2CC"));
        assert!(matches!(
            table.fill,
            Some(ChartStyleFill::Solid { ref color }) if color == "FFF2CC"
        ));
        assert_eq!(table.fill_hidden, None);
        assert_eq!(table.fill_paint_authored, Some(true));
        assert_eq!(table.line_color.as_deref(), Some("445566"));
        assert_eq!(table.line_width_emu, Some(12700));
        assert_eq!(table.line_dash.as_deref(), Some("dash"));
    }

    #[test]
    fn parse_chart_data_table_bounds_gradient_before_expansion() {
        let chart = |stop_count: usize, trailing_fill: &str| {
            let stops = (0..stop_count)
                .map(|index| format!(r#"<a:gs pos="{index}"><a:srgbClr val="112233"/></a:gs>"#,))
                .collect::<String>();
            format!(
                r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
                  <c:chart><c:plotArea>
                    <c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/>
                      <c:cat><c:strCache><c:pt idx="0"><c:v>A</c:v></c:pt></c:strCache></c:cat>
                      <c:val><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:val>
                    </c:ser></c:barChart>
                    <c:dTable><c:spPr><a:gradFill><a:gsLst>{stops}</a:gsLst></a:gradFill>{trailing_fill}</c:spPr></c:dTable>
                  </c:plotArea></c:chart>
                </c:chartSpace>"#,
            )
        };

        let exact_xml = chart(MAX_CHART_PAINT_RECIPE_COMPONENTS, "");
        let exact = parse_chart_part(
            chart_space_of(&exact_xml).root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("paint recipe at the shared ceiling parses");
        assert!(matches!(
            exact.data_table.and_then(|table| table.fill),
            Some(ChartStyleFill::Gradient { ref stops, .. })
                if stops.len() == MAX_CHART_PAINT_RECIPE_COMPONENTS
        ));

        for trailing_fill in ["", r#"<a:pattFill prst="diagCross"/>"#, "<a:blipFill/>"] {
            let oversized_xml = chart(MAX_CHART_PAINT_RECIPE_COMPONENTS + 1, trailing_fill);
            assert!(parse_chart_part(
                chart_space_of(&oversized_xml).root_element(),
                &ChartParseContext {
                    color_resolver: Some(&FixtureResolver),
                    ..Default::default()
                }
            )
            .is_none());
        }
    }

    /// Excel's implicit classic style is not serialized when a bare chart has
    /// no `<c:style>` or series/point formatting.  The Office-observed sign
    /// mirror pair shows that the all-negative chart becomes outline-only, while the
    /// positive mirror remains an ordinary accent-filled column.  Preserve the
    /// observed effective paint without extending it to mixed-sign or authored
    /// series, whose Office defaults were not part of that boundary set.
    #[test]
    fn parse_chart_part_applies_bounded_outline_only_negative_default() {
        fn parsed_series(value: i32, authored_fill: bool) -> ChartSeries {
            let sp_pr = if authored_fill {
                r#"<c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr>"#
            } else {
                ""
            };
            let xml = format!(
                r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
                  <c:chart><c:plotArea><c:barChart>
                    <c:barDir val="col"/><c:varyColors val="0"/>
                    <c:ser><c:idx val="0"/>{sp_pr}
                      <c:val><c:numCache><c:pt idx="0"><c:v>{value}</c:v></c:pt></c:numCache></c:val>
                    </c:ser>
                  </c:barChart></c:plotArea></c:chart>
                </c:chartSpace>"#
            );
            parse_chart_part(
                chart_space_of(&xml).root_element(),
                &ChartParseContext {
                    color_resolver: Some(&XlsxCompatibilityFixtureResolver),
                    ..Default::default()
                },
            )
            .expect("bare column chart parses")
            .series
            .into_iter()
            .next()
            .expect("one series")
        }

        let negative = parsed_series(-24_000, false);
        assert_eq!(negative.invert_if_negative, None);
        assert_eq!(negative.automatic_negative_style, Some(true));
        assert_eq!(negative.inverted_fill_hidden, None);
        assert_eq!(negative.inverted_line_color, None);
        assert_eq!(negative.inverted_line_width_emu, None);

        let positive = parsed_series(24_000, false);
        assert_eq!(positive.invert_if_negative, None);
        assert_eq!(positive.automatic_negative_style, None);
        assert_eq!(positive.inverted_fill_hidden, None);

        let authored = parsed_series(-24_000, true);
        assert_eq!(authored.invert_if_negative, None);
        assert_eq!(authored.automatic_negative_style, None);
        assert_eq!(authored.inverted_fill_hidden, None);

        fn automatic_styles(
            style: &str,
            plot_markup: &str,
            resolver: &dyn ColorResolver,
        ) -> Vec<Option<bool>> {
            let xml = format!(
                r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart">{style}
                  <c:chart><c:plotArea>{plot_markup}</c:plotArea></c:chart>
                </c:chartSpace>"#
            );
            parse_chart_part(
                chart_space_of(&xml).root_element(),
                &ChartParseContext {
                    color_resolver: Some(resolver),
                    ..Default::default()
                },
            )
            .expect("scope-boundary chart parses")
            .series
            .into_iter()
            .map(|series| series.automatic_negative_style)
            .collect()
        }
        let negative_series = r#"<c:ser><c:idx val="0"/>
          <c:val><c:numCache><c:pt idx="0"><c:v>-1</c:v></c:pt></c:numCache></c:val>
        </c:ser>"#;
        let observed_plot = format!(
            r#"<c:barChart><c:barDir val="col"/><c:varyColors val="0"/>{negative_series}</c:barChart>"#
        );

        // Host opt-in is required; the shared DOCX/PPTX default stays unset.
        assert_eq!(
            automatic_styles("", &observed_plot, &FixtureResolver),
            vec![None]
        );
        assert_eq!(
            automatic_styles(
                "",
                &format!(r#"<c:barChart><c:barDir val="col"/>{negative_series}</c:barChart>"#),
                &XlsxCompatibilityFixtureResolver,
            ),
            vec![None]
        );
        assert_eq!(
            automatic_styles(
                "",
                &format!(
                    r#"<c:barChart><c:barDir val="col"/><c:varyColors val="1"/>{negative_series}</c:barChart>"#
                ),
                &XlsxCompatibilityFixtureResolver,
            ),
            vec![None]
        );
        // Every authored/structural gate outside the observed sign-mirror boundary remains
        // unset until a corresponding Office boundary is available.
        assert_eq!(
            automatic_styles(
                r#"<c:style val="2"/>"#,
                &observed_plot,
                &XlsxCompatibilityFixtureResolver,
            ),
            vec![None]
        );
        assert_eq!(
            automatic_styles(
                "",
                &format!(
                    r#"<c:barChart><c:barDir val="bar"/><c:varyColors val="0"/>{negative_series}</c:barChart>"#
                ),
                &XlsxCompatibilityFixtureResolver,
            ),
            vec![None]
        );
        assert_eq!(
            automatic_styles(
                "",
                &format!(
                    r#"<c:barChart><c:barDir val="col"/><c:grouping val="stacked"/><c:varyColors val="0"/>{negative_series}</c:barChart>"#
                ),
                &XlsxCompatibilityFixtureResolver,
            ),
            vec![None]
        );
        assert_eq!(
            automatic_styles(
                "",
                &format!(
                    r#"<c:bar3DChart><c:barDir val="col"/><c:varyColors val="0"/>{negative_series}</c:bar3DChart>"#
                ),
                &XlsxCompatibilityFixtureResolver,
            ),
            vec![None]
        );
        assert_eq!(
            automatic_styles(
                "",
                &format!(
                    r#"<c:barChart><c:barDir val="col"/><c:varyColors val="0"/>{negative_series}<c:ser><c:idx val="1"/><c:val><c:numCache><c:pt idx="0"><c:v>-2</c:v></c:pt></c:numCache></c:val></c:ser></c:barChart>"#
                ),
                &XlsxCompatibilityFixtureResolver,
            ),
            vec![None, None]
        );
        assert_eq!(
            automatic_styles(
                "",
                r#"<c:barChart><c:barDir val="col"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></c:spPr></c:dPt><c:val><c:numCache><c:pt idx="0"><c:v>-1</c:v></c:pt></c:numCache></c:val></c:ser></c:barChart>"#,
                &XlsxCompatibilityFixtureResolver,
            ),
            vec![None]
        );
        assert_eq!(
            automatic_styles(
                "",
                r#"<c:barChart><c:barDir val="col"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:invertIfNegative val="0"/><c:val><c:numCache><c:pt idx="0"><c:v>-1</c:v></c:pt></c:numCache></c:val></c:ser></c:barChart>"#,
                &XlsxCompatibilityFixtureResolver,
            ),
            vec![None]
        );
        assert_eq!(
            automatic_styles(
                "",
                r#"<c:barChart><c:barDir val="col"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:val><c:numCache><c:pt idx="0"><c:v>-1</c:v></c:pt></c:numCache></c:val><c:extLst><c:ext uri="invert"><c14:invertSolidFillFmt><c14:spPr><a:solidFill><a:srgbClr val="FF00FF"/></a:solidFill></c14:spPr></c14:invertSolidFillFmt></c:ext></c:extLst></c:ser></c:barChart>"#,
                &XlsxCompatibilityFixtureResolver,
            ),
            vec![None]
        );
        assert_eq!(
            automatic_styles(
                "",
                r#"<c:barChart><c:barDir val="col"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:val><c:numCache><c:pt idx="0"><c:v>-1</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt></c:numCache></c:val></c:ser></c:barChart>"#,
                &XlsxCompatibilityFixtureResolver,
            ),
            vec![None]
        );
        assert_eq!(
            automatic_styles(
                "",
                &format!(
                    r#"<c:barChart><c:barDir val="col"/><c:varyColors val="0"/>{negative_series}</c:barChart><c:lineChart><c:ser><c:idx val="1"/><c:val><c:numCache><c:pt idx="0"><c:v>-2</c:v></c:pt></c:numCache></c:val></c:ser></c:lineChart>"#
                ),
                &XlsxCompatibilityFixtureResolver,
            ),
            vec![None, None]
        );
    }

    /// (e) `chart_type` normalization for stacked/percentStacked. As of CH13,
    /// `parse_chart_part` routes bar/line/area type detection through the shared
    /// `canonical_chart_type` helper (previously an inline match duplicated the
    /// logic and — as a latent bug — folded a percentStacked BAR down to plain
    /// `stackedBar`, so the renderer's `stackedBarPct` 100%-normalization never
    /// fired for a parsed chart). It now distinguishes the percent variant for
    /// BAR (`stackedBarPct` / `stackedBarHPct`) and AREA (`stackedAreaPct`),
    /// matching the LINE behavior and the standalone helper's own matrix test.
    #[test]
    fn parse_chart_part_stacked_percent_stacked_chart_type() {
        fn bar_chart_type(grouping: &str, bar_dir: &str) -> String {
            let xml = format!(
                r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea>
                  <c:barChart>
                    <c:barDir val="{bar_dir}"/>
                    <c:grouping val="{grouping}"/>
                    <c:ser><c:idx val="0"/>
                      <c:cat><c:strCache><c:pt idx="0"><c:v>A</c:v></c:pt></c:strCache></c:cat>
                      <c:val><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:val>
                    </c:ser>
                  </c:barChart>
                </c:plotArea></c:chart></c:chartSpace>"#
            );
            let doc = chart_space_of(&xml);
            parse_chart_part(
                doc.root_element(),
                &ChartParseContext {
                    color_resolver: Some(&FixtureResolver),
                    ..Default::default()
                },
            )
            .unwrap()
            .chart_type
        }
        fn line_chart_type(grouping: &str) -> String {
            let xml = format!(
                r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea>
                  <c:lineChart>
                    <c:grouping val="{grouping}"/>
                    <c:ser><c:idx val="0"/>
                      <c:cat><c:strCache><c:pt idx="0"><c:v>A</c:v></c:pt></c:strCache></c:cat>
                      <c:val><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:val>
                    </c:ser>
                  </c:lineChart>
                </c:plotArea></c:chart></c:chartSpace>"#
            );
            let doc = chart_space_of(&xml);
            parse_chart_part(
                doc.root_element(),
                &ChartParseContext {
                    color_resolver: Some(&FixtureResolver),
                    ..Default::default()
                },
            )
            .unwrap()
            .chart_type
        }

        assert_eq!(bar_chart_type("stacked", "col"), "stackedBar");
        // CH13: percentStacked now maps to the Pct canonical variant (the
        // renderer normalizes those to 100%), fixing the prior fold-to-stacked.
        assert_eq!(bar_chart_type("percentStacked", "col"), "stackedBarPct");
        assert_eq!(bar_chart_type("percentStacked", "bar"), "stackedBarHPct");
        // Line + area also distinguish percentStacked.
        assert_eq!(line_chart_type("percentStacked"), "stackedLinePct");
    }

    /// A missing plot area is not a chart. An authored empty chart group is
    /// schema-valid and remains represented so its source-order slot is not
    /// lost when a later group gains data.
    #[test]
    fn parse_chart_part_rejects_missing_plot_area_but_retains_empty_group() {
        let no_plot_area = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:title/></c:chart></c:chartSpace>"#
        );
        assert!(parse_chart_part(
            chart_space_of(&no_plot_area).root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            }
        )
        .is_none());

        let empty_series = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea>
                <c:barChart><c:barDir val="col"/><c:grouping val="clustered"/></c:barChart>
              </c:plotArea></c:chart></c:chartSpace>"#
        );
        let model = parse_chart_part(
            chart_space_of(&empty_series).root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("empty group remains represented");
        assert!(model.series.is_empty());
        let groups = model.plot_groups.expect("ordered group");
        assert_eq!(groups.len(), 1);
        assert_eq!(
            (groups[0].kind.as_str(), groups[0].series_count),
            ("bar", 0)
        );
    }

    /// §21.2.2.47: a per-point `<c:dLbl>` carries its own show-flag group and
    /// text style, overriding the series-level `<c:dLbls>` (§21.2.2.49) for that
    /// point. Office can set `showCatName=0 showPercent=1` plus white text per
    /// slice while the series default is `showCatName=1` black. The parser
    /// must surface both the series default AND the per-point flag / color
    /// overrides, and mark a genuine `<c:delete>` distinctly from a style-only
    /// `<c:dLbl>` (which has an empty `text`).
    #[test]
    fn parse_chart_part_pie_per_point_dlbl_overrides_series_defaults() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:chart><c:plotArea>
                <c:pieChart>
                  <c:ser>
                    <c:idx val="0"/>
                    <c:dLbls>
                      <c:dLbl>
                        <c:idx val="0"/>
                        <c:txPr><a:p><a:pPr><a:defRPr b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>
                        <c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="1"/>
                      </c:dLbl>
                      <c:dLbl>
                        <c:idx val="1"/>
                        <c:delete val="1"/>
                      </c:dLbl>
                      <c:txPr><a:p><a:pPr><a:defRPr b="1"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>
                      <c:showVal val="0"/><c:showCatName val="1"/><c:showSerName val="0"/><c:showPercent val="1"/>
                    </c:dLbls>
                    <c:cat><c:strCache><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strCache></c:cat>
                    <c:val><c:numCache><c:pt idx="0"><c:v>60</c:v></c:pt><c:pt idx="1"><c:v>40</c:v></c:pt></c:numCache></c:val>
                  </c:ser>
                </c:pieChart>
              </c:plotArea></c:chart>
            </c:chartSpace>"#
        );
        let m = parse_chart_part(
            chart_space_of(&xml).root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("pie chart parses");
        assert_eq!(m.chart_type, "pie");
        let def = m.series[0]
            .series_data_labels
            .as_ref()
            .expect("series-level dLbls present");
        // Series default: category name ON, black text.
        assert!(def.show_cat_name, "series default shows category name");
        assert!(def.show_percent);
        assert_eq!(def.font_color.as_deref(), Some("000000"));

        let ovs = m.series[0]
            .data_label_overrides
            .as_ref()
            .expect("per-point overrides present");
        let ov0 = ovs.iter().find(|o| o.idx == 0).expect("idx 0 override");
        // Per-point idx 0: category name OFF, percent ON, white — overrides the
        // series default so this slice renders as white percent-only.
        assert_eq!(ov0.show_cat_name, Some(false));
        assert_eq!(ov0.show_percent, Some(true));
        assert_eq!(ov0.font_color.as_deref(), Some("FFFFFF"));
        assert_ne!(ov0.deleted, Some(true), "a style-only dLbl is not a delete");

        let ov1 = ovs.iter().find(|o| o.idx == 1).expect("idx 1 override");
        // Per-point idx 1: a genuine `<c:delete>` → flagged so the renderer skips it.
        assert_eq!(ov1.deleted, Some(true));
    }

    #[test]
    fn ct_boolean_series_dlbl_show_flags_bare_are_true() {
        // §21.2.2.187/.180/.185/.183 series-level show* flags. A bare
        // <c:showVal/> etc. ⇒ true; the shared parser must not collapse it to
        // false. `<c:showSerName val="0"/>` stays false (explicit override).
        let cache = std::collections::HashMap::new();
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:dLbls>
                <c:showVal/>
                <c:showCatName/>
                <c:showSerName val="0"/>
                <c:showPercent/>
                <c:showBubbleSize/>
                <c:showLegendKey/>
              </c:dLbls>
            </c:ser>"#
        );
        let d = root_of(&xml);
        let (defaults, _) = parse_series_data_labels(d.root_element(), &FixtureResolver, &cache);
        let defaults = defaults.expect("dLbls present");
        assert!(defaults.show_val, "bare <c:showVal/> ⇒ true");
        assert!(defaults.show_cat_name, "bare <c:showCatName/> ⇒ true");
        assert!(
            !defaults.show_ser_name,
            "<c:showSerName val=\"0\"/> ⇒ false"
        );
        assert!(defaults.show_percent, "bare <c:showPercent/> ⇒ true");
        assert!(defaults.show_bubble_size, "bare <c:showBubbleSize/> ⇒ true");
        assert!(defaults.show_legend_key, "bare <c:showLegendKey/> ⇒ true");
    }

    #[test]
    fn ct_boolean_series_show_leader_lines_bare_is_true() {
        // §21.2.2.183 `<c:showLeaderLines/>` ⇒ true (val default).
        let cache = std::collections::HashMap::new();
        let bare =
            format!(r#"<c:ser xmlns:c="{C_NS}"><c:dLbls><c:showLeaderLines/></c:dLbls></c:ser>"#);
        let d = root_of(&bare);
        let (defaults, _) = parse_series_data_labels(d.root_element(), &FixtureResolver, &cache);
        assert!(
            defaults.expect("dLbls present").show_leader_lines,
            "bare <c:showLeaderLines/> ⇒ true"
        );
    }

    #[test]
    fn ct_boolean_per_point_dlbl_bare_delete_and_flags_are_true() {
        // §21.2.2.43 per-point `<c:delete/>` ⇒ that point's label removed.
        // §21.2.2.47 per-point show* ⇒ Some(true) for a bare flag (overrides the
        // series default for that point), Some(false) for val="0".
        let cache = std::collections::HashMap::new();
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:dLbls>
                <c:dLbl>
                  <c:idx val="0"/>
                  <c:delete/>
                </c:dLbl>
                <c:dLbl>
                  <c:idx val="2"/>
                  <c:showPercent/>
                  <c:showCatName val="0"/>
                  <c:showBubbleSize val="0"/>
                  <c:showLegendKey/>
                </c:dLbl>
                <c:showVal val="1"/>
              </c:dLbls>
            </c:ser>"#
        );
        let d = root_of(&xml);
        let (_, overrides) = parse_series_data_labels(d.root_element(), &FixtureResolver, &cache);
        let del = overrides
            .iter()
            .find(|o| o.idx == 0)
            .expect("idx 0 override");
        assert_eq!(
            del.deleted,
            Some(true),
            "bare per-point <c:delete/> ⇒ deleted"
        );
        let flags = overrides
            .iter()
            .find(|o| o.idx == 2)
            .expect("idx 2 override");
        assert_eq!(
            flags.show_percent,
            Some(true),
            "bare <c:showPercent/> ⇒ Some(true)"
        );
        assert_eq!(flags.show_cat_name, Some(false), "val=\"0\" ⇒ Some(false)");
        assert_eq!(flags.show_bubble_size, Some(false));
        assert_eq!(flags.show_legend_key, Some(true));
    }

    #[test]
    fn ct_boolean_err_bars_no_end_cap_bare_is_true() {
        // §21.2.2.117 `<c:noEndCap/>` ⇒ true (no I-beam end caps).
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:errBars>
                <c:errBarType val="both"/>
                <c:errValType val="fixedVal"/>
                <c:noEndCap/>
                <c:val val="1"/>
              </c:errBars>
            </c:ser>"#
        );
        let d = root_of(&xml);
        let bars = parse_error_bars(d.root_element(), &[Some(1.0)], &FixtureResolver);
        assert_eq!(bars.len(), 1);
        assert!(bars[0].no_end_cap, "bare <c:noEndCap/> ⇒ true");
    }

    #[test]
    fn classic_single_series_uses_distinct_series_and_varying_point_index_domains() {
        let parse = |group: &str| {
            let xml = format!(
                r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:style val="3"/>
                  <c:chart><c:plotArea>{group}</c:plotArea></c:chart></c:chartSpace>"#,
            );
            let document = chart_space_of(&xml);
            parse_chart_part(
                document.root_element(),
                &ChartParseContext {
                    color_resolver: Some(&FixtureResolver),
                    ..Default::default()
                },
            )
            .expect("single-series classic chart parses")
        };
        let series = r#"<c:ser><c:idx val="8"/><c:order val="0"/>
          <c:cat><c:strLit><c:ptCount val="3"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strLit></c:cat>
          <c:val><c:numLit><c:ptCount val="3"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numLit></c:val></c:ser>"#;

        let line = parse(&format!(r#"<c:lineChart>{series}</c:lineChart>"#));
        assert_eq!(line.series[0].chartex_format_idx, Some(8));
        let line_roles = line.classic_chart_style_roles.expect("numeric line roles");
        assert_eq!(
            line_roles["dataPointLine"]
                .line_formatting_indices
                .as_deref(),
            Some(&[8][..]),
        );

        let bar = parse(&format!(
            r#"<c:barChart><c:barDir val="col"/><c:varyColors val="1"/>{series}</c:barChart>"#,
        ));
        assert_eq!(bar.series[0].chartex_format_idx, Some(8));
        let bar_roles = bar
            .classic_varying_point_chart_style_roles
            .expect("numeric varying-point bar roles");
        assert_eq!(
            bar_roles["dataPoint"].fill_formatting_indices.as_deref(),
            Some(&[0, 1, 2][..]),
        );
        assert_eq!(
            bar_roles["dataPoint"].line_formatting_indices.as_deref(),
            Some(&[0, 1, 2][..]),
        );
        let bar_series_roles = bar
            .classic_chart_style_roles
            .expect("numeric bar series roles");
        assert_eq!(
            bar_series_roles["dataPointMarker"]
                .fill_formatting_indices
                .as_deref(),
            Some(&[8][..]),
        );
        assert_eq!(
            bar_series_roles["dataPointLine"]
                .line_formatting_indices
                .as_deref(),
            Some(&[8][..]),
        );
    }

    #[test]
    fn classic_varying_point_domains_are_group_local() {
        let series = |idx: usize, count: usize| {
            let categories = (0..count)
                .map(|point| format!(r#"<c:pt idx="{point}"><c:v>C{point}</c:v></c:pt>"#))
                .collect::<String>();
            let values = (0..count)
                .map(|point| format!(r#"<c:pt idx="{point}"><c:v>{point}</c:v></c:pt>"#))
                .collect::<String>();
            format!(
                r#"<c:ser><c:idx val="{idx}"/><c:order val="{idx}"/>
                  <c:cat><c:strLit><c:ptCount val="{count}"/>{categories}</c:strLit></c:cat>
                  <c:val><c:numLit><c:ptCount val="{count}"/>{values}</c:numLit></c:val>
                </c:ser>"#,
            )
        };
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:style val="11"/>
              <c:chart><c:plotArea>
                <c:areaChart>{}</c:areaChart>
                <c:lineChart><c:varyColors val="1"/>{}</c:lineChart>
                <c:radarChart><c:radarStyle val="marker"/><c:varyColors val="1"/>{}</c:radarChart>
              </c:plotArea></c:chart></c:chartSpace>"#,
            series(0, 100),
            series(1, 3),
            series(2, 5),
        );
        let document = chart_space_of(&xml);
        let model = parse_chart_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("mixed varying chart parses");
        let groups = model
            .classic_varying_point_chart_style_roles_by_group
            .expect("group domains");
        assert_eq!(groups.len(), 3);
        assert!(groups[0].is_none());
        assert!(groups[1].is_none(), "common point domain is stored once");
        assert_eq!(
            model
                .classic_varying_point_chart_style_roles
                .as_ref()
                .unwrap()["dataPoint"]
                .fill_formatting_indices
                .as_deref(),
            Some(&[0, 1, 2][..]),
        );
        assert_eq!(
            groups[2].as_ref().unwrap()["dataPoint"]
                .fill_formatting_indices
                .as_deref(),
            Some(&[0, 1, 2, 3, 4][..]),
        );
    }

    /// §21.2.2.140 pie3DChart retains the canonical `pie` family plus the
    /// shared view3D model; series/cat/val continue through the ordinary model.
    #[test]
    fn parse_chart_part_pie3d_flattens_to_pie() {
        let group = format!(r#"<c:pie3DChart><c:varyColors val="1"/>{CH13_SER}</c:pie3DChart>"#);
        let xml_p = chart_space_with_group(&group);
        let d = chart_space_of(&xml_p);
        let m = parse_chart_part(
            d.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("pie3D parses");
        assert_eq!(m.chart_type, "pie");
        assert_eq!(m.series.len(), 1);
        assert_eq!(m.series[0].values, vec![Some(3.0), Some(7.0)]);
        assert_eq!(m.categories, vec!["A".to_string(), "B".to_string()]);
        assert!(m.three_d.is_some());
    }

    /// §21.2.2.15 bar3DChart uses the shared stacked-bar data model while its
    /// view/depth controls remain available to the simplified 3D painter.
    #[test]
    fn parse_chart_part_bar3d_flattens_by_grouping_and_dir() {
        let series = CH13_SER.replace(
            "<c:idx val=\"0\"/>",
            "<c:idx val=\"0\"/><c:shape val=\"pyramid\"/>",
        );
        let group = format!(
            r#"<c:bar3DChart><c:barDir val="col"/><c:grouping val="stacked"/><c:gapDepth val="150"/><c:shape val="cylinder"/>{series}</c:bar3DChart>"#
        );
        let xml = chart_space_with_group(&group).replace(
            "<c:chart><c:plotArea>",
            r#"<c:chart><c:view3D><c:rotX val="25"/><c:hPercent val="120"/><c:rotY val="330"/><c:depthPercent val="250"/><c:rAngAx val="0"/><c:perspective val="45"/></c:view3D>
              <c:floor><c:spPr><a:noFill/><a:ln w="3175"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="solid"/></a:ln></c:spPr></c:floor>
              <c:sideWall><c:spPr><a:noFill/><a:ln w="12700"><a:solidFill><a:srgbClr val="808080"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr></c:sideWall>
              <c:backWall><c:thickness val="25%"/><c:spPr><a:gradFill><a:gsLst>
                <a:gs pos="0"><a:srgbClr val="F2F2F2"/></a:gs>
                <a:gs pos="100000"><a:srgbClr val="808080"/></a:gs>
              </a:gsLst><a:lin ang="0"/></a:gradFill><a:ln><a:noFill/></a:ln></c:spPr>
                <c:pictureOptions><c:applyToFront/><c:applyToSides val="0"/><c:applyToEnd val="1"/>
                  <c:pictureFormat val="stackScale"/><c:pictureStackUnit val="2.5"/>
                </c:pictureOptions></c:backWall>
              <c:plotArea>"#,
        );
        let d = chart_space_of(&xml);
        let m = parse_chart_part(
            d.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("bar3D parses");
        assert_eq!(m.chart_type, "stackedBar");
        let view = m.three_d.expect("3D view");
        assert_eq!(view.rotation_x, Some(25));
        assert_eq!(view.rotation_y, Some(330));
        assert_eq!(view.height_percent, Some(120.0));
        assert_eq!(view.depth_percent, Some(250.0));
        assert_eq!(view.perspective, Some(45));
        assert_eq!(view.right_angle_axes, Some(false));
        assert_eq!(view.gap_depth_percent, Some(150.0));
        assert_eq!(view.shape.as_deref(), Some("cylinder"));
        assert_eq!(view.bar_grouping.as_deref(), Some("stacked"));
        let floor = view.floor.as_ref().expect("floor surface");
        assert_eq!(
            floor.style.as_ref().and_then(|style| style.fill_hidden),
            Some(true)
        );
        assert_eq!(floor.fill_hidden, Some(true));
        assert_eq!(floor.line_color.as_deref(), Some("000000"));
        assert_eq!(floor.line_width_emu, Some(3175));
        assert_eq!(floor.line_dash.as_deref(), Some("solid"));
        assert_eq!(floor.line_hidden, Some(false));
        let side = view.side_wall.as_ref().expect("side wall surface");
        assert_eq!(side.fill_hidden, Some(true));
        assert_eq!(side.line_color.as_deref(), Some("808080"));
        assert_eq!(side.line_width_emu, Some(12700));
        assert_eq!(side.line_dash.as_deref(), Some("dash"));
        let back = view.back_wall.as_ref().expect("back wall surface");
        assert_eq!(back.fill_color, None);
        let back_style = back.style.as_ref().expect("structured back-wall style");
        assert!(matches!(
            back_style.fill_paints.as_deref().and_then(|paints| paints.first()).and_then(Option::as_ref),
            Some(ChartStyleFill::Gradient { stops, .. }) if stops.len() == 2
        ));
        assert_eq!(back.line_hidden, Some(true));
        assert_eq!(back.thickness_percent, Some(25.0));
        let picture = back.picture_options.as_ref().expect("picture options");
        assert_eq!(picture.apply_to_front, Some(true));
        assert_eq!(picture.apply_to_sides, Some(false));
        assert_eq!(picture.apply_to_end, Some(true));
        assert_eq!(picture.picture_format.as_deref(), Some("stackScale"));
        assert_eq!(picture.picture_format_authored, Some(true));
        assert_eq!(picture.picture_stack_unit, Some(2.5));
        assert_eq!(picture.picture_stack_unit_authored, Some(true));
        assert_eq!(m.series[0].three_d_shape.as_deref(), Some("pyramid"));

        // barDir=bar (horizontal) + clustered → clusteredBarH.
        let group_h = format!(
            r#"<c:bar3DChart><c:barDir val="bar"/><c:grouping val="clustered"/>{CH13_SER}</c:bar3DChart>"#
        );
        let xml_h = chart_space_with_group(&group_h);
        let d2 = chart_space_of(&xml_h);
        let m2 = parse_chart_part(
            d2.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("bar3D-h parses");
        assert_eq!(m2.chart_type, "clusteredBarH");
        assert_eq!(
            m2.three_d
                .as_ref()
                .and_then(|view| view.bar_grouping.as_deref()),
            Some("clustered")
        );

        // CT_Grouping@val defaults to `standard`. 3-D standard bars occupy the
        // series axis, while clustered bars share one depth plane; the
        // canonical 2-D family name cannot carry that distinction by itself.
        let group_standard =
            format!(r#"<c:bar3DChart><c:barDir val="col"/>{CH13_SER}</c:bar3DChart>"#);
        let xml_standard = chart_space_with_group(&group_standard);
        let d3 = chart_space_of(&xml_standard);
        let m3 = parse_chart_part(
            d3.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("bar3D standard parses");
        assert_eq!(m3.chart_type, "clusteredBar");
        assert_eq!(
            m3.three_d
                .as_ref()
                .and_then(|view| view.bar_grouping.as_deref()),
            Some("standard")
        );

        // `<c:serAx>` is a real third coordinate axis for standard 3-D bars,
        // not a legend surrogate. Preserve its authored title, tick interval,
        // text paint and rule through the shared chart wire model.
        let xml_with_series_axis = xml_standard.replace(
            "</c:plotArea>",
            r#"<c:serAx>
              <c:axId val="3"/><c:scaling><c:orientation val="maxMin"/></c:scaling>
              <c:title><c:layout><c:manualLayout><c:x val="0.7"/><c:y val="0.68"/></c:manualLayout></c:layout>
                <c:tx><c:rich><a:bodyPr rot="-5400000"/><a:p><a:pPr><a:defRPr sz="800" b="1" i="1">
                  <a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="Arial"/>
                </a:defRPr></a:pPr><a:r><a:t>Series Axis</a:t></a:r></a:p></c:rich></c:tx>
              </c:title>
              <c:majorTickMark val="cross"/><c:minorTickMark val="cross"/><c:tickLblPos val="low"/>
              <c:tickLblSkip val="2"/><c:tickMarkSkip val="3"/>
              <c:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:ln></c:spPr>
              <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="900" b="0">
                <a:solidFill><a:srgbClr val="778899"/></a:solidFill><a:latin typeface="Calibri"/>
              </a:defRPr></a:pPr></a:p></c:txPr>
            </c:serAx></c:plotArea>"#,
        );
        let series_axis_doc = chart_space_of(&xml_with_series_axis);
        let series_axis_model = parse_chart_part(
            series_axis_doc.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("bar3D series axis parses");
        let series_axis = series_axis_model
            .three_d
            .as_ref()
            .and_then(|view| view.series_axis.as_ref())
            .expect("series axis");
        assert_eq!(series_axis.title.as_deref(), Some("Series Axis"));
        assert_eq!(series_axis.orientation.as_deref(), Some("maxMin"));
        assert_eq!(series_axis.tick_label_pos.as_deref(), Some("low"));
        assert_eq!(series_axis.tick_label_skip, Some(2));
        assert_eq!(series_axis.tick_mark_skip, Some(3));
        assert_eq!(series_axis.major_tick_mark, "cross");
        assert_eq!(series_axis.minor_tick_mark.as_deref(), Some("cross"));
        assert_eq!(series_axis.font_color.as_deref(), Some("778899"));
        assert_eq!(series_axis.font_size_hpt, Some(900));
        assert_eq!(series_axis.line_color.as_deref(), Some("445566"));
        assert_eq!(series_axis.line_width_emu, Some(12_700));
        assert_eq!(series_axis.line_paint_authored, Some(true));
        assert_eq!(series_axis.title_font_size_hpt, Some(800));
        assert_eq!(series_axis.title_font_bold, Some(true));
        assert_eq!(series_axis.title_font_italic, Some(true));
        assert_eq!(series_axis.title_font_color.as_deref(), Some("112233"));
        assert_eq!(series_axis.title_rotation, Some(-5_400_000));
        assert_eq!(
            series_axis
                .title_manual_layout
                .as_ref()
                .map(|layout| (layout.x, layout.y)),
            Some((0.7, 0.68))
        );
    }

    #[test]
    fn parse_chart_part_keeps_invalid_picture_option_presence_fail_closed() {
        let group = format!(r#"<c:bar3DChart><c:barDir val="col"/>{CH13_SER}</c:bar3DChart>"#);
        let xml = chart_space_with_group(&group).replace(
            "<c:chart><c:plotArea>",
            r#"<c:chart><c:backWall><c:pictureOptions>
              <c:applyToFront val="TRUE"/><c:pictureFormat val="future"/>
              <c:pictureStackUnit val="NaN"/>
            </c:pictureOptions></c:backWall><c:plotArea>"#,
        );
        let document = chart_space_of(&xml);
        let model = parse_chart_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("3-D picture option provenance parses");
        let options = model
            .three_d
            .as_ref()
            .and_then(|view| view.back_wall.as_ref())
            .and_then(|surface| surface.picture_options.as_ref())
            .expect("picture options");
        assert_eq!(options.apply_to_front, Some(false));
        assert_eq!(options.picture_format, None);
        assert_eq!(options.picture_format_authored, Some(true));
        assert_eq!(options.picture_stack_unit, None);
        assert_eq!(options.picture_stack_unit_authored, Some(true));
    }

    #[test]
    fn parse_chart_part_bounds_surface_band_format_count_atomically() {
        let bands = |count: usize| {
            (0..count)
                .map(|index| format!(r#"<c:bandFmt><c:idx val="{index}"/></c:bandFmt>"#))
                .collect::<String>()
        };
        let surface = |count: usize| {
            format!(
                r#"<c:surface3DChart>{CH13_SER}<c:bandFmts>{}</c:bandFmts></c:surface3DChart>"#,
                bands(count),
            )
        };

        let exact_xml = chart_space_with_group(&surface(MAX_CHART_COLOR_STYLE_ENTRIES));
        let exact_document = chart_space_of(&exact_xml);
        let exact = parse_chart_part(
            exact_document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("exact band-format boundary");
        assert_eq!(
            exact.surface_band_formats.as_ref().map(Vec::len),
            Some(MAX_CHART_COLOR_STYLE_ENTRIES),
        );

        let oversized_xml = chart_space_with_group(&surface(MAX_CHART_COLOR_STYLE_ENTRIES + 1));
        let oversized_document = chart_space_of(&oversized_xml);
        assert!(parse_chart_part(
            oversized_document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            }
        )
        .is_none());
    }

    /// §21.2.2.96 line3DChart → `line`; §21.2.2.4 area3DChart(stacked) →
    /// `stackedArea`.
    #[test]
    fn parse_chart_part_line3d_area3d_flatten() {
        let line = format!(r#"<c:line3DChart>{CH13_SER}</c:line3DChart>"#);
        let xml_l = chart_space_with_group(&line);
        let dl = chart_space_of(&xml_l);
        assert_eq!(
            parse_chart_part(
                dl.root_element(),
                &ChartParseContext {
                    color_resolver: Some(&FixtureResolver),
                    ..Default::default()
                }
            )
            .unwrap()
            .chart_type,
            "line"
        );
        let area =
            format!(r#"<c:area3DChart><c:grouping val="stacked"/>{CH13_SER}</c:area3DChart>"#);
        let xml_a = chart_space_with_group(&area);
        let da = chart_space_of(&xml_a);
        assert_eq!(
            parse_chart_part(
                da.root_element(),
                &ChartParseContext {
                    color_resolver: Some(&FixtureResolver),
                    ..Default::default()
                }
            )
            .unwrap()
            .chart_type,
            "stackedArea"
        );

        let styled_series = r#"<c:ser><c:idx val="0"/>
          <c:spPr><a:ln cap="rnd"><a:solidFill><a:srgbClr val="123456"/></a:solidFill>
            <a:prstDash val="dash"/><a:round/></a:ln></c:spPr>
          <c:marker><c:symbol val="circle"/><c:spPr><a:pattFill prst="pct20">
            <a:fgClr><a:srgbClr val="112233"/></a:fgClr><a:bgClr><a:srgbClr val="DDEEFF"/></a:bgClr>
          </a:pattFill></c:spPr></c:marker>
          <c:dPt><c:idx val="0"/><c:spPr><a:ln w="25400"><a:solidFill>
            <a:srgbClr val="ABCDEF"/></a:solidFill><a:prstDash val="dot"/></a:ln></c:spPr>
            <c:marker><c:symbol val="diamond"/><c:spPr><a:pattFill prst="pct30">
              <a:fgClr><a:srgbClr val="445566"/></a:fgClr><a:bgClr><a:srgbClr val="AABBCC"/></a:bgClr>
            </a:pattFill></c:spPr></c:marker></c:dPt>
          <c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>A</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>"#;
        let styled_group = format!(r#"<c:line3DChart>{styled_series}</c:line3DChart>"#);
        let styled_xml = chart_space_with_group(&styled_group);
        let styled_doc = chart_space_of(&styled_xml);
        let styled = parse_chart_part(
            styled_doc.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .unwrap();
        let style = styled.series[0]
            .chartex_style
            .as_ref()
            .expect("classic direct line style");
        assert_eq!(style.line_dash.as_deref(), Some("dash"));
        assert_eq!(style.line_cap.as_deref(), Some("rnd"));
        assert_eq!(style.line_join.as_deref(), Some("round"));
        assert_eq!(
            styled.series[0]
                .marker_style
                .as_ref()
                .and_then(|style| style.fill_paints.as_ref())
                .and_then(|paints| paints.first())
                .cloned()
                .flatten(),
            Some(ChartStyleFill::Pattern {
                fg: "112233".to_string(),
                bg: "DDEEFF".to_string(),
                preset: "pct20".to_string(),
            })
        );
        assert_eq!(styled.series[0].marker_fill_paint_authored, Some(true));
        let point = &styled.series[0].data_point_overrides.as_ref().unwrap()[0];
        assert_eq!(point.line_color.as_deref(), Some("ABCDEF"));
        assert_eq!(point.line_width_emu, Some(25400));
        assert_eq!(point.line_dash.as_deref(), Some("dot"));
        assert_eq!(
            point
                .marker_style
                .as_ref()
                .and_then(|style| style.fill_paints.as_ref())
                .and_then(|paints| paints.first())
                .cloned()
                .flatten(),
            Some(ChartStyleFill::Pattern {
                fg: "445566".to_string(),
                bg: "AABBCC".to_string(),
                preset: "pct30".to_string(),
            })
        );
        assert_eq!(point.marker_fill_paint_authored, Some(true));
    }

    /// §21.2.2.198 stockChart → `stock` (its high/low/close series flow through
    /// the shared collectors unchanged).
    #[test]
    fn parse_chart_part_stock_detected() {
        let hi = r#"<c:ser><c:idx val="0"/><c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>High</c:v></c:pt></c:strCache></c:strRef></c:tx>
            <c:cat><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:cat>
            <c:val><c:numCache><c:pt idx="0"><c:v>55</c:v></c:pt></c:numCache></c:val></c:ser>"#;
        let group = format!(
            r#"<c:stockChart>{hi}
              <c:dropLines><c:spPr><a:ln w="12700" cap="sq"><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:prstDash val="dashDot"/><a:round/></a:ln></c:spPr></c:dropLines>
              <c:hiLowLines><c:spPr><a:ln w="25400" cap="rnd"><a:solidFill><a:srgbClr val="808080"/></a:solidFill><a:prstDash val="dot"/><a:bevel/></a:ln></c:spPr></c:hiLowLines>
            </c:stockChart>"#
        );
        let xml = chart_space_with_group(&group);
        let d = chart_space_of(&xml);
        let m = parse_chart_part(
            d.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("stock parses");
        assert_eq!(m.chart_type, "stock");
        assert_eq!(m.series.len(), 1);
        assert_eq!(m.series[0].name, "High");
        assert_eq!(m.series[0].values, vec![Some(55.0)]);
        let drop_lines = m.stock_drop_lines.expect("stock drop lines");
        assert_eq!(drop_lines.color.as_deref(), Some("123456"));
        assert_eq!(drop_lines.width_emu, Some(12_700));
        assert_eq!(drop_lines.dash.as_deref(), Some("dashDot"));
        assert_eq!(drop_lines.cap.as_deref(), Some("sq"));
        assert_eq!(drop_lines.join.as_deref(), Some("round"));
        assert_eq!(drop_lines.hidden, None);
        assert_eq!(drop_lines.paint_authored, Some(true));
        // hiLowLines present + its resolved line color; no upDownBars in fixture.
        assert_eq!(m.stock_hi_low_lines, Some(true));
        assert_eq!(m.stock_hi_low_line_color.as_deref(), Some("808080"));
        let hi_low = m.stock_hi_low_line_style.expect("stock high-low style");
        assert_eq!(hi_low.color.as_deref(), Some("808080"));
        assert_eq!(hi_low.width_emu, Some(25_400));
        assert_eq!(hi_low.dash.as_deref(), Some("dot"));
        assert_eq!(hi_low.cap.as_deref(), Some("rnd"));
        assert_eq!(hi_low.join.as_deref(), Some("bevel"));
        assert_eq!(hi_low.hidden, None);
        assert_eq!(hi_low.paint_authored, Some(true));
        assert_eq!(m.stock_up_down_bars, None);
    }

    /// A stock chart WITHOUT `<c:hiLowLines>` but WITH `<c:upDownBars>`: the
    /// hi-lo flag is `Some(false)` (element absent), and the gap plus direct
    /// up/down paint are retained for the stock renderer.
    #[test]
    fn parse_chart_part_stock_up_down_bars_recognized() {
        let ser = r#"<c:ser><c:idx val="0"/>
            <c:cat><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:cat>
            <c:val><c:numCache><c:pt idx="0"><c:v>5</c:v></c:pt></c:numCache></c:val></c:ser>"#;
        let group = format!(
            r#"<c:stockChart>{ser}
            <c:dropLines><c:spPr><a:ln><a:noFill/></a:ln></c:spPr></c:dropLines>
            <c:upDownBars>
            <c:gapWidth val="80"/>
            <c:upBars><c:spPr><a:solidFill><a:srgbClr val="00AA00"/></a:solidFill>
              <a:ln w="12700" cap="sq"><a:solidFill><a:srgbClr val="006600"/></a:solidFill><a:prstDash val="dash"/><a:round/></a:ln>
            </c:spPr></c:upBars>
            <c:downBars><c:spPr><a:noFill/><a:ln w="25400"><a:noFill/><a:bevel/></a:ln></c:spPr></c:downBars>
          </c:upDownBars></c:stockChart>"#
        );
        let xml = chart_space_with_group(&group);
        let d = chart_space_of(&xml);
        let m = parse_chart_part(
            d.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("stock parses");
        let drop_lines = m.stock_drop_lines.expect("stock drop lines");
        assert_eq!(drop_lines.hidden, Some(true));
        assert_eq!(drop_lines.paint_authored, Some(true));
        assert_eq!(m.stock_hi_low_lines, Some(false));
        assert_eq!(m.stock_hi_low_line_color, None);
        assert_eq!(m.stock_up_down_bars, Some(true));
        let style = m.stock_up_down_bar_style.expect("up/down style");
        assert_eq!(style.gap_width_percent, 80.0);
        assert_eq!(style.up.fill_color.as_deref(), Some("00AA00"));
        assert_eq!(style.up.fill_paint_authored, Some(true));
        assert_eq!(style.up.line_color.as_deref(), Some("006600"));
        assert_eq!(style.up.line_paint_authored, Some(true));
        assert_eq!(style.up.line_width_emu, Some(12700));
        assert_eq!(style.up.line_dash.as_deref(), Some("dash"));
        assert_eq!(style.up.line_cap.as_deref(), Some("sq"));
        assert_eq!(style.up.line_join.as_deref(), Some("round"));
        assert_eq!(style.down.fill_hidden, Some(true));
        assert_eq!(style.down.fill_paint_authored, Some(true));
        assert_eq!(style.down.line_hidden, Some(true));
        assert_eq!(style.down.line_paint_authored, Some(true));
        assert_eq!(style.down.line_width_emu, Some(25400));
        assert_eq!(style.down.line_join.as_deref(), Some("bevel"));
    }

    #[test]
    fn parse_chart_part_line_group_preserves_drop_hi_low_and_up_down_geometry() {
        let series = |index: u32, name: &str, value: i32| {
            format!(
                r#"<c:ser><c:idx val="{index}"/><c:order val="{index}"/><c:tx><c:v>{name}</c:v></c:tx>
              <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
              <c:val><c:numLit><c:pt idx="0"><c:v>{value}</c:v></c:pt></c:numLit></c:val></c:ser>"#
            )
        };
        let first = format!(
            r#"<c:lineChart><c:grouping val="standard"/>{}{}
              <c:dropLines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="111111"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr></c:dropLines>
              <c:hiLowLines><c:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="222222"/></a:solidFill></a:ln></c:spPr></c:hiLowLines>
              <c:upDownBars><c:gapWidth val="80"/><c:upBars><c:spPr><a:solidFill><a:srgbClr val="EEEEEE"/></a:solidFill></c:spPr></c:upBars><c:downBars><c:spPr><a:solidFill><a:srgbClr val="333333"/></a:solidFill></c:spPr></c:downBars></c:upDownBars>
            </c:lineChart>"#,
            series(0, "First", 10),
            series(1, "Last", 20),
        );
        let second = format!(
            r#"<c:lineChart><c:grouping val="standard"/>{}</c:lineChart>"#,
            series(2, "Second group", 30),
        );
        let xml = chart_space_with_group(&format!("{first}{second}"));
        let document = chart_space_of(&xml);
        let model = parse_chart_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("line chart");

        assert_eq!(model.series[0].line_group_index, Some(0));
        assert_eq!(model.series[1].line_group_index, Some(0));
        assert_eq!(model.series[2].line_group_index, Some(1));
        let groups = model.line_group_decorations.expect("line decorations");
        assert_eq!(groups.len(), 1);
        let group = &groups[0];
        assert_eq!(group.group_index, 0);
        let drop = group.drop_lines.as_ref().expect("drop lines");
        assert_eq!(drop.color.as_deref(), Some("111111"));
        assert_eq!(drop.width_emu, Some(9525));
        assert_eq!(drop.dash.as_deref(), Some("dash"));
        let hi_low = group.hi_low_lines.as_ref().expect("hi-low lines");
        assert_eq!(hi_low.color.as_deref(), Some("222222"));
        assert_eq!(hi_low.width_emu, Some(12700));
        let up_down = group.up_down_bars.as_ref().expect("up/down bars");
        assert_eq!(up_down.gap_width_percent, 80.0);
        assert_eq!(up_down.up.fill_color.as_deref(), Some("EEEEEE"));
        assert_eq!(up_down.down.fill_color.as_deref(), Some("333333"));
    }

    #[test]
    fn parse_chart_part_preserves_area_group_drop_lines_and_ownership() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:chart><c:plotArea><c:areaChart>
                <c:grouping val="standard"/>
                <c:ser><c:idx val="0"/><c:order val="0"/>
                  <c:cat><c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numLit></c:val>
                </c:ser>
                <c:dropLines><c:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:ln></c:spPr></c:dropLines>
                <c:axId val="1"/><c:axId val="2"/>
              </c:areaChart></c:plotArea></c:chart>
            </c:chartSpace>"#,
        );
        let document = chart_space_of(&xml);
        let model = parse_chart_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("area chart parses");
        let wire = serde_json::to_value(model).expect("serializes");

        assert_eq!(wire["series"][0]["areaGroupIndex"], 0);
        assert_eq!(wire["areaGroupDecorations"][0]["groupIndex"], 0);
        assert_eq!(
            wire["areaGroupDecorations"][0]["dropLines"]["color"],
            "123456"
        );
        assert_eq!(
            wire["areaGroupDecorations"][0]["dropLines"]["widthEmu"],
            12700
        );
    }

    #[test]
    fn parse_chart_part_preserves_bar_group_series_lines_and_ownership() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:chart><c:plotArea><c:barChart>
                <c:barDir val="col"/><c:grouping val="stacked"/>
                <c:ser><c:idx val="0"/><c:order val="0"/>
                  <c:cat><c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numLit></c:val>
                </c:ser>
                <c:serLines><c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="234567"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr></c:serLines>
                <c:serLines><c:spPr><a:ln><a:noFill/></a:ln></c:spPr></c:serLines>
                <c:axId val="1"/><c:axId val="2"/>
              </c:barChart></c:plotArea></c:chart>
            </c:chartSpace>"#,
        );
        let document = chart_space_of(&xml);
        let model = parse_chart_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("bar chart parses");
        let wire = serde_json::to_value(model).expect("serializes");

        assert_eq!(wire["series"][0]["barGroupIndex"], 0);
        assert_eq!(wire["barGroupDecorations"][0]["groupIndex"], 0);
        assert_eq!(
            wire["barGroupDecorations"][0]["seriesLines"]
                .as_array()
                .map(Vec::len),
            Some(2)
        );
        assert_eq!(
            wire["barGroupDecorations"][0]["seriesLines"][0]["color"],
            "234567"
        );
        assert_eq!(
            wire["barGroupDecorations"][0]["seriesLines"][0]["widthEmu"],
            19050
        );
        assert_eq!(
            wire["barGroupDecorations"][0]["seriesLines"][0]["dash"],
            "dash"
        );
        assert_eq!(
            wire["barGroupDecorations"][0]["seriesLines"][1]["hidden"],
            true
        );
    }

    #[test]
    fn parse_chart_part_surface_preserves_matrix_role_and_wireframe() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:style val="10"/><c:chart>
              <c:view3D><c:rotX val="90"/></c:view3D><c:plotArea><c:surfaceChart>
                <c:wireframe val="0"/>
                <c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Y1</c:v></c:tx>
                  <c:cat><c:strLit><c:pt idx="0"><c:v>X1</c:v></c:pt><c:pt idx="1"><c:v>X2</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numLit></c:val></c:ser>
                <c:ser><c:idx val="1"/><c:order val="1"/><c:tx><c:v>Y2</c:v></c:tx>
                  <c:cat><c:strLit><c:pt idx="0"><c:v>X1</c:v></c:pt><c:pt idx="1"><c:v>X2</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:pt idx="0"><c:v>20</c:v></c:pt><c:pt idx="1"><c:v>40</c:v></c:pt></c:numLit></c:val></c:ser>
                <c:bandFmts><c:bandFmt><c:idx val="0"/><c:spPr>
                  <a:pattFill prst="pct20"><a:fgClr><a:srgbClr val="112233"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill>
                  <a:ln w="12700"><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:ln>
                </c:spPr></c:bandFmt><c:bandFmt><c:idx val="1"/><c:spPr>
                  <a:grpFill/>
                </c:spPr></c:bandFmt></c:bandFmts>
              </c:surfaceChart><c:serAx><c:axId val="3"/></c:serAx></c:plotArea>
            </c:chart></c:chartSpace>"#
        );
        let document = chart_space_of(&xml);
        let model = parse_chart_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("surface");
        assert_eq!(model.chart_type, "surface");
        assert_eq!(model.surface_wireframe, Some(false));
        assert_eq!(model.legacy_chart_style, Some(10));
        assert_eq!(model.theme_accent_colors.as_ref().map(Vec::len), Some(6));
        assert_eq!(model.categories, vec!["X1", "X2"]);
        assert_eq!(model.series.len(), 2);
        assert_eq!(model.series[0].series_type.as_deref(), Some("surface"));
        assert_eq!(
            model.classic_chart_style_roles.as_ref().unwrap()["dataPoint3D"]
                .fill_colors
                .as_ref()
                .map(Vec::len),
            Some(2),
            "ordinary numeric role retains the source-series domain",
        );
        let surface_numeric = model
            .classic_surface_band_styles
            .as_ref()
            .expect("filled Surface has a band-domain numeric role");
        assert!(surface_numeric.by_band_count.is_none());
        assert_eq!(
            surface_numeric
                .fixed
                .as_ref()
                .and_then(|style| style.fill_colors.as_ref())
                .map(Vec::len),
            Some(48),
        );
        let bands = model.surface_band_formats.as_ref().expect("band formats");
        let band = &bands[0];
        assert_eq!(band.idx, 0);
        assert!(matches!(band.fill, Some(ChartStyleFill::Pattern { .. })));
        let band_style = band.style.as_ref().expect("full band style");
        assert!(band_style.fill_paints.is_none());
        assert_eq!(band_style.fill_paint_authored, Some(true));
        assert!(matches!(
            band_style
                .line_paints
                .as_deref()
                .and_then(|paints| paints.first())
                .and_then(Option::as_ref),
            Some(ChartStyleFill::Solid { color }) if color == "445566"
        ));
        assert_eq!(band.line_color.as_deref(), Some("445566"));
        assert_eq!(band.line_width_emu, Some(12700));
        let unresolved = &bands[1];
        assert!(unresolved.fill.is_none());
        assert_eq!(unresolved.fill_hidden, None);
        assert_eq!(
            unresolved
                .style
                .as_ref()
                .and_then(|style| style.fill_paint_authored),
            Some(true),
        );
        assert!(model
            .three_d
            .as_ref()
            .and_then(|scene| scene.series_axis.as_ref())
            .is_some());

        let wireframe_xml = xml.replace(r#"<c:wireframe val="0"/>"#, r#"<c:wireframe val="1"/>"#);
        let wireframe_document = chart_space_of(&wireframe_xml);
        let wireframe = parse_chart_part(
            wireframe_document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("wireframe surface");
        let wireframe_numeric = wireframe
            .classic_surface_band_styles
            .as_ref()
            .and_then(|styles| styles.fixed.as_ref())
            .expect("wireframe has a band-domain numeric role");
        assert!(wireframe_numeric.fill_colors.is_none());
        assert_eq!(
            wireframe_numeric
                .line_formatting_indices
                .as_ref()
                .map(Vec::len),
            Some(48),
        );
    }

    #[test]
    fn parse_chart_part_preserves_bare_view3d_schema_defaults() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:chart>
              <c:view3D><c:rotX/><c:hPercent/><c:rotY/><c:depthPercent/><c:rAngAx/><c:perspective/></c:view3D><c:plotArea>
                <c:line3DChart><c:grouping val="standard"/><c:ser><c:idx val="0"/><c:order val="0"/>
                  <c:cat><c:strLit><c:pt idx="0"><c:v>X1</c:v></c:pt><c:pt idx="1"><c:v>X2</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val>
                </c:ser><c:gapDepth/></c:line3DChart>
              </c:plotArea></c:chart></c:chartSpace>"#,
        );
        let document = chart_space_of(&xml);
        let model = parse_chart_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("line3D");
        let view = model.three_d.expect("3-D view");
        assert_eq!(view.view_3d_present, Some(true));
        assert_eq!(view.rotation_x, Some(0));
        assert_eq!(view.rotation_x_authored, Some(true));
        assert_eq!(view.rotation_y, Some(0));
        assert_eq!(view.rotation_y_authored, Some(true));
        assert_eq!(view.height_percent, Some(100.0));
        assert_eq!(view.height_percent_authored, Some(true));
        assert_eq!(view.depth_percent, Some(100.0));
        assert_eq!(view.depth_percent_authored, Some(true));
        assert_eq!(view.perspective, Some(30));
        assert_eq!(view.perspective_authored, Some(true));
        assert_eq!(view.right_angle_axes, Some(true));
        assert_eq!(view.right_angle_axes_authored, Some(true));
        assert_eq!(view.gap_depth_percent, Some(150.0));
        assert_eq!(view.gap_depth_percent_authored, Some(true));
    }

    #[test]
    fn parse_chart_part_parses_view3d_scalars_and_chart_percentages_by_schema_type() {
        for (namespace, height, depth, gap) in [
            (C_NS, " 125 ", " 250 ", " 175 "),
            (
                "http://purl.oclc.org/ooxml/drawingml/chart",
                "125%",
                "250%",
                "175%",
            ),
        ] {
            let xml = format!(
                r#"<c:chartSpace xmlns:c="{namespace}"><c:chart>
                  <c:view3D><c:rotX val=" -30 "/><c:hPercent val="{height}"/>
                    <c:rotY val=" 270 "/><c:depthPercent val="{depth}"/>
                    <c:perspective val=" 40 "/></c:view3D><c:plotArea>
                    <c:line3DChart><c:grouping val="standard"/><c:ser>
                      <c:idx val="0"/><c:order val="0"/>
                      <c:cat><c:strLit><c:pt idx="0"><c:v>X</c:v></c:pt></c:strLit></c:cat>
                      <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
                    </c:ser><c:gapDepth val="{gap}"/></c:line3DChart>
                  </c:plotArea></c:chart></c:chartSpace>"#,
            );
            let document = chart_space_of(&xml);
            let view = parse_chart_part(
                document.root_element(),
                &ChartParseContext {
                    color_resolver: Some(&FixtureResolver),
                    ..Default::default()
                },
            )
            .expect("line3D")
            .three_d
            .expect("3-D view");
            assert_eq!(view.rotation_x, Some(-30));
            assert_eq!(view.height_percent, Some(125.0));
            assert_eq!(view.rotation_y, Some(270));
            assert_eq!(view.depth_percent, Some(250.0));
            assert_eq!(view.perspective, Some(40));
            assert_eq!(view.gap_depth_percent, Some(175.0));
        }
    }

    #[test]
    fn parse_chart_part_rejects_cross_type_and_non_lexical_strict_view3d_values() {
        let namespace = "http://purl.oclc.org/ooxml/drawingml/chart";
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{namespace}"><c:chart>
              <c:view3D><c:rotX val="-30%"/><c:hPercent val="125"/>
                <c:rotY val="270%"/><c:depthPercent val="250.5%"/>
                <c:perspective val="40%"/></c:view3D><c:plotArea>
                <c:line3DChart><c:grouping val="standard"/><c:ser>
                  <c:idx val="0"/><c:order val="0"/>
                  <c:cat><c:strLit><c:pt idx="0"><c:v>X</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
                </c:ser><c:gapDepth val="175"/></c:line3DChart>
              </c:plotArea></c:chart></c:chartSpace>"#,
        );
        let document = chart_space_of(&xml);
        let view = parse_chart_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("line3D")
        .three_d
        .expect("3-D view");
        assert_eq!(view.rotation_x, None);
        assert_eq!(view.height_percent, None);
        assert_eq!(view.rotation_y, None);
        assert_eq!(view.depth_percent, None);
        assert_eq!(view.perspective, None);
        assert_eq!(view.gap_depth_percent, None);
        assert_eq!(view.rotation_x_authored, Some(true));
        assert_eq!(view.gap_depth_percent_authored, Some(true));
    }

    #[test]
    fn parse_chart_part_distinguishes_omitted_view3d_children_from_bare_defaults() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea>
              <c:line3DChart><c:grouping val="standard"/><c:ser>
                <c:idx val="0"/><c:order val="0"/>
                <c:cat><c:strLit><c:pt idx="0"><c:v>X1</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
              </c:ser></c:line3DChart>
            </c:plotArea></c:chart></c:chartSpace>"#,
        );
        let document = chart_space_of(&xml);
        let model = parse_chart_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("line3D");
        let view = model.three_d.expect("3-D model");
        assert_eq!(view.view_3d_present, Some(false));
        assert_eq!(view.rotation_x, None);
        assert_eq!(view.rotation_x_authored, Some(false));
        assert_eq!(view.rotation_y, None);
        assert_eq!(view.rotation_y_authored, Some(false));
        assert_eq!(view.height_percent, None);
        assert_eq!(view.height_percent_authored, Some(false));
        assert_eq!(view.depth_percent, None);
        assert_eq!(view.depth_percent_authored, Some(false));
        assert_eq!(view.perspective, None);
        assert_eq!(view.perspective_authored, Some(false));
        assert_eq!(view.right_angle_axes, None);
        assert_eq!(view.right_angle_axes_authored, Some(false));
        assert_eq!(view.gap_depth_percent, None);
        assert_eq!(view.gap_depth_percent_authored, Some(false));
    }

    #[test]
    fn parse_chart_part_keeps_surface_3d_distinct_from_contour_surface() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:chart>
              <c:view3D><c:rotX val="30"/></c:view3D><c:plotArea><c:surface3DChart>
                <c:ser><c:idx val="0"/><c:order val="0"/>
                  <c:cat><c:strLit><c:pt idx="0"><c:v>X1</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:pt idx="0"><c:v>10</c:v></c:pt></c:numLit></c:val>
                </c:ser><c:axId val="1"/><c:axId val="2"/><c:axId val="3"/>
              </c:surface3DChart><c:serAx><c:axId val="3"/></c:serAx></c:plotArea>
            </c:chart></c:chartSpace>"#,
        );
        let document = chart_space_of(&xml);
        let model = parse_chart_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("surface3D");
        assert_eq!(model.chart_type, "surface3D");
        assert_eq!(model.series[0].series_type.as_deref(), Some("surface3D"));
    }

    /// §21.2.2.126 ofPieChart remains a distinct renderer family.
    #[test]
    fn parse_chart_part_ofpie_keeps_family_and_defaults() {
        let group = format!(
            r#"<c:ofPieChart><c:ofPieType val="pie"/><c:varyColors val="1"/>{CH13_SER}</c:ofPieChart>"#
        );
        let xml = chart_space_with_group(&group);
        let d = chart_space_of(&xml);
        let m = parse_chart_part(
            d.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("ofPie parses");
        assert_eq!(m.chart_type, "ofPie");
        assert_eq!(m.series[0].values, vec![Some(3.0), Some(7.0)]);
        let of_pie = m.of_pie.expect("ofPie contract");
        assert_eq!(of_pie.r#type, "pie");
        assert_eq!(of_pie.split_type, "auto");
        assert!(!of_pie.split_type_authored);
        assert!(!of_pie.split_pos_authored);
        assert_eq!(of_pie.second_pie_size_percent, 75.0);
    }

    #[test]
    fn parse_chart_part_ofpie_retains_invalid_split_pos_pair_for_fail_closed_rendering() {
        let group = format!(
            r#"<c:ofPieChart><c:ofPieType val="pie"/><c:varyColors val="1"/>{CH13_SER}<c:splitPos val="2"/></c:ofPieChart>"#
        );
        let xml = chart_space_with_group(&group);
        let document = chart_space_of(&xml);
        let model = parse_chart_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("ofPie parses");
        let of_pie = model.of_pie.expect("ofPie contract");
        assert_eq!(of_pie.split_type, "auto");
        assert!(!of_pie.split_type_authored);
        assert_eq!(of_pie.split_pos, Some(2.0));
        assert!(of_pie.split_pos_authored);
    }

    #[test]
    fn parse_chart_part_ofpie_retains_invalid_split_pos_presence() {
        for (split_type, split_pos) in [
            ("", r#"<c:splitPos/>"#),
            (
                r#"<c:splitType val="cust"/>"#,
                r#"<c:splitPos val="NaN"/><c:custSplit><c:secondPiePt val="1"/></c:custSplit>"#,
            ),
        ] {
            let group = format!(
                r#"<c:ofPieChart><c:ofPieType val="pie"/><c:varyColors val="1"/>{CH13_SER}{split_type}{split_pos}</c:ofPieChart>"#
            );
            let xml = chart_space_with_group(&group);
            let document = chart_space_of(&xml);
            let model = parse_chart_part(
                document.root_element(),
                &ChartParseContext {
                    color_resolver: Some(&FixtureResolver),
                    ..Default::default()
                },
            )
            .expect("ofPie parses");
            let of_pie = model.of_pie.expect("ofPie contract");
            assert_eq!(of_pie.split_pos, None);
            assert!(of_pie.split_pos_authored);
        }
    }

    /// Full ofPieChart contract retains secondary-plot controls without
    /// disturbing the source-index palette carried by the series.
    #[test]
    fn parse_chart_part_ofpie_full_contract_keeps_split() {
        let ser = r#"<c:ser><c:idx val="0"/>
            <c:cat><c:strCache>
              <c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt>
              <c:pt idx="2"><c:v>C</c:v></c:pt><c:pt idx="3"><c:v>D</c:v></c:pt>
            </c:strCache></c:cat>
            <c:val><c:numCache>
              <c:pt idx="0"><c:v>40</c:v></c:pt><c:pt idx="1"><c:v>30</c:v></c:pt>
              <c:pt idx="2"><c:v>20</c:v></c:pt><c:pt idx="3"><c:v>10</c:v></c:pt>
            </c:numCache></c:val></c:ser>"#;
        // bar-of-pie with a custom split of the last two points into the bar,
        // plus directly formatted connector series-lines.
        let group = format!(
            r#"<c:ofPieChart>
                <c:ofPieType val="bar"/>
                <c:varyColors val="1"/>
                {ser}
                <c:gapWidth val="100"/>
                <c:splitType val="pos"/>
                <c:splitPos val="2"/>
                <c:secondPieSize val="75"/>
                <c:serLines><c:spPr><a:ln w="25400"><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr></c:serLines>
            </c:ofPieChart>"#
        );
        let xml = chart_space_with_group(&group);
        let d = chart_space_of(&xml);
        let m = parse_chart_part(
            d.root_element(),
            &ChartParseContext {
                color_resolver: Some(&AccentResolver),
                ..Default::default()
            },
        )
        .expect("ofPie parses");
        assert_eq!(m.chart_type, "ofPie");
        assert_eq!(
            m.series[0].values,
            vec![Some(40.0), Some(30.0), Some(20.0), Some(10.0)]
        );
        // Automatic varyColors paint belongs to the effective dataPoint style;
        // this field is reserved for direct dPt formatting.
        assert!(m.series[0].data_point_colors.is_none());
        let of_pie = m.of_pie.expect("ofPie contract");
        assert_eq!(of_pie.r#type, "bar");
        assert_eq!(of_pie.split_type, "pos");
        assert!(of_pie.split_type_authored);
        assert_eq!(of_pie.split_pos, Some(2.0));
        assert!(of_pie.split_pos_authored);
        assert_eq!(of_pie.second_pie_size_percent, 75.0);
        assert_eq!(of_pie.gap_width_percent, 100.0);
        assert!(of_pie.series_lines);
        let series_line = of_pie.series_line_style.expect("connector style");
        assert_eq!(series_line.color.as_deref(), Some("123456"));
        assert_eq!(series_line.width_emu, Some(25400));
        assert_eq!(series_line.dash.as_deref(), Some("dash"));
    }

    #[test]
    fn parse_chart_part_preserves_bubble3d_group_series_and_point_provenance() {
        for namespace in [C_NS, "http://purl.oclc.org/ooxml/drawingml/chart"] {
            let xml = format!(
                r#"<c:chartSpace xmlns:c="{namespace}"><c:chart><c:plotArea>
                  <c:bubbleChart>
                    <c:ser><c:idx val="0"/><c:order val="0"/>
                      <c:dPt><c:idx val="0"/><c:bubble3D/></c:dPt>
                      <c:xVal><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:xVal>
                      <c:yVal><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:yVal>
                      <c:bubbleSize><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>5</c:v></c:pt><c:pt idx="1"><c:v>6</c:v></c:pt></c:numLit></c:bubbleSize>
                      <c:bubble3D val="0"/>
                    </c:ser>
                    <c:bubble3D/>
                    <c:axId val="1"/><c:axId val="2"/>
                  </c:bubbleChart>
                  <c:bubbleChart>
                    <c:ser><c:idx val="1"/><c:order val="1"/>
                      <c:dPt><c:idx val="0"/><c:bubble3D/></c:dPt>
                      <c:dPt><c:idx val="1"/><c:bubble3D val="TRUE"/></c:dPt>
                      <c:xVal><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:xVal>
                      <c:yVal><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>7</c:v></c:pt><c:pt idx="1"><c:v>8</c:v></c:pt></c:numLit></c:yVal>
                      <c:bubbleSize><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>5</c:v></c:pt><c:pt idx="1"><c:v>6</c:v></c:pt></c:numLit></c:bubbleSize>
                    </c:ser>
                    <c:bubble3D val="false"/>
                    <c:axId val="1"/><c:axId val="2"/>
                  </c:bubbleChart>
                </c:plotArea></c:chart></c:chartSpace>"#
            );
            let document = root_of(&xml);
            let chart = parse_chart_part(
                document.root_element(),
                &ChartParseContext {
                    color_resolver: Some(&FixtureResolver),
                    ..Default::default()
                },
            )
            .expect("bubble chart parses");

            assert_eq!(chart.series.len(), 2);
            assert_eq!(chart.series[0].bubble_3d_group_default, Some(true));
            assert_eq!(chart.series[0].bubble_3d, Some(false));
            let first_points = chart.series[0]
                .data_point_overrides
                .as_ref()
                .expect("bubble3D-only point is retained");
            assert_eq!(first_points.len(), 1);
            assert_eq!(first_points[0].bubble_3d, Some(true));

            assert_eq!(chart.series[1].bubble_3d_group_default, Some(false));
            assert_eq!(chart.series[1].bubble_3d, None);
            let second_points = chart.series[1]
                .data_point_overrides
                .as_ref()
                .expect("both point booleans are retained");
            assert_eq!(second_points.len(), 2);
            assert_eq!(second_points[0].bubble_3d, Some(true));
            assert_eq!(
                second_points[1].bubble_3d,
                Some(false),
                "invalid xsd:boolean stays authored and fails closed"
            );

            let wire = serde_json::to_value(&chart).expect("chart serializes");
            assert_eq!(wire["series"][0]["bubble3DGroupDefault"], true);
            assert_eq!(wire["series"][0]["bubble3D"], false);
            assert_eq!(
                wire["series"][0]["dataPointOverrides"][0]["bubble3D"], true,
                "the Rust-to-TypeScript boundary preserves DrawingML's 3D token"
            );
            assert!(wire["series"][0].get("bubble3d").is_none());
            assert!(wire["series"][0]["dataPointOverrides"][0]
                .get("bubble3d")
                .is_none());
        }
    }

    #[test]
    fn parse_chart_part_ignores_legacy_series_bubble3d_outside_bubble_family() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:chart><c:plotArea><c:lineChart>
              <c:ser><c:idx val="0"/><c:order val="0"/>
                <c:dPt><c:idx val="0"/><c:spPr><a:gradFill><a:gsLst>
                  <a:gs pos="0"><a:srgbClr val="112233"/></a:gs>
                  <a:gs pos="100000"><a:srgbClr val="DDEEFF"/></a:gs>
                </a:gsLst></a:gradFill></c:spPr><c:bubble3D/></c:dPt>
                <c:cat><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
                <c:bubble3D/>
              </c:ser>
            </c:lineChart></c:plotArea></c:chart></c:chartSpace>"#
        );
        let document = root_of(&xml);
        let chart = parse_chart_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("line chart parses");
        assert_eq!(chart.series[0].bubble_3d_group_default, None);
        assert_eq!(chart.series[0].bubble_3d, None);
        let point = &chart.series[0].data_point_overrides.as_ref().unwrap()[0];
        assert_eq!(
            point.bubble_3d,
            Some(true),
            "CT_DPt is shared and retains the point value without applying it to lines"
        );
        assert!(
            point.chartex_style.is_some(),
            "shared dPt shape paint is retained"
        );
    }

    #[test]
    fn parse_chart_part_resolves_shared_categories_once_without_series_clones() {
        let first = r#"<c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:strRef><c:f>Cats</c:f></c:strRef></c:cat><c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>"#;
        let repeated = r#"<c:ser><c:idx val="1"/><c:order val="1"/><c:cat><c:strRef><c:f>Cats</c:f></c:strRef></c:cat><c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>"#;
        let category_less: String = (2..12)
            .map(|idx| format!(r#"<c:ser><c:idx val="{idx}"/><c:order val="{idx}"/><c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>{idx}</c:v></c:pt></c:numLit></c:val></c:ser>"#))
            .collect();
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea><c:lineChart>{first}{repeated}{category_less}</c:lineChart></c:plotArea></c:chart></c:chartSpace>"#
        );
        let doc = root_of(&xml);
        let mut references = CountingCategoryResolver { string_calls: 0 };
        let chart = parse_chart_part(
            doc.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                references: std::cell::Cell::new(Some(&mut references)),
                ..Default::default()
            },
        )
        .expect("multi-series chart parses");

        assert_eq!(references.string_calls, 1);
        assert_eq!(chart.categories, vec!["A", "B", "C"]);
        assert_eq!(chart.series.len(), 12);
        assert!(chart
            .series
            .iter()
            .all(|series| series.categories.is_none()));
    }

    #[test]
    fn parse_chart_part_does_not_reuse_first_x_for_unresolved_distinct_scatter_x() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea><c:scatterChart>
              <c:ser><c:idx val="0"/><c:order val="0"/><c:xVal><c:numRef><c:f>X</c:f></c:numRef></c:xVal><c:yVal><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:yVal></c:ser>
              <c:ser><c:idx val="1"/><c:order val="1"/><c:xVal><c:numRef><c:f>UnavailableX</c:f></c:numRef></c:xVal><c:yVal><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:yVal></c:ser>
              <c:ser><c:idx val="2"/><c:order val="2"/><c:xVal><c:numRef><c:f>X</c:f><c:numCache><c:ptCount val="0"/></c:numCache></c:numRef></c:xVal><c:yVal><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>5</c:v></c:pt></c:numLit></c:yVal></c:ser>
            </c:scatterChart></c:plotArea></c:chart></c:chartSpace>"#
        );
        let doc = root_of(&xml);
        let mut references = FormulaResolver;
        let chart = parse_chart_part(
            doc.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                references: std::cell::Cell::new(Some(&mut references)),
                ..Default::default()
            },
        )
        .expect("two-series scatter parses");

        assert_eq!(chart.categories, vec!["1", "2"]);
        assert_eq!(chart.series[0].categories, None);
        assert_eq!(chart.series[1].categories, Some(Vec::new()));
        assert_eq!(
            chart.series[2].categories,
            Some(vec!["1".into(), "2".into()])
        );
    }

    #[test]
    fn classic_plot_groups_keep_scatter_bubble_distinct_and_fail_closed_on_ambiguous_axes() {
        let series = |idx: usize, bubble: bool| {
            format!(
                r#"<c:ser><c:idx val="{idx}"/><c:order val="{idx}"/><c:xVal><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:xVal><c:yVal><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:yVal>{}</c:ser>"#,
                if bubble {
                    r#"<c:bubbleSize><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>3</c:v></c:pt></c:numLit></c:bubbleSize>"#
                } else {
                    ""
                },
            )
        };
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea>
              <c:bubbleChart>{}<c:axId val="30"/><c:axId val="40"/></c:bubbleChart>
              <c:scatterChart><c:scatterStyle val="lineMarker"/>{}<c:axId val="10"/><c:axId val="20"/></c:scatterChart>
              <c:valAx><c:axId val="10"/><c:axPos val="b"/><c:crossAx val="20"/></c:valAx>
              <c:valAx><c:axId val="20"/><c:axPos val="l"/><c:crossAx val="10"/></c:valAx>
              <c:valAx><c:axId val="30"/><c:axPos val="b"/><c:crossAx val="40"/></c:valAx>
              <c:valAx><c:axId val="40"/><c:axPos val="l"/><c:crossAx val="30"/></c:valAx>
            </c:plotArea></c:chart></c:chartSpace>"#,
            series(0, true),
            series(1, false),
        );
        let document = root_of(&xml);
        let model = parse_chart_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("numeric mixed groups");
        let groups = model.plot_groups.expect("ordered groups");
        assert_eq!(
            groups
                .iter()
                .map(|group| group.kind.as_str())
                .collect::<Vec<_>>(),
            vec!["bubble", "scatter"]
        );
        assert_eq!(
            (
                groups[0].category_axis.as_str(),
                groups[0].value_axis.as_str()
            ),
            ("unresolved", "unresolved")
        );
        assert_eq!(
            (
                groups[1].category_axis.as_str(),
                groups[1].value_axis.as_str()
            ),
            ("unresolved", "unresolved")
        );
        assert_eq!(groups[1].scatter_style.as_deref(), Some("lineMarker"));
    }

    #[test]
    fn stock_group_decorations_survive_a_later_line_group() {
        let series = |index: usize, value: usize| {
            format!(
                r#"<c:ser><c:idx val="{index}"/><c:order val="{index}"/><c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>{value}</c:v></c:pt></c:numLit></c:val></c:ser>"#,
            )
        };
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:chart><c:plotArea>
              <c:stockChart>{}{}{}<c:hiLowLines><c:spPr><a:ln><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:ln></c:spPr></c:hiLowLines></c:stockChart>
              <c:lineChart>{}</c:lineChart>
            </c:plotArea></c:chart></c:chartSpace>"#,
            series(0, 5),
            series(1, 1),
            series(2, 3),
            series(3, 4),
        );
        let document = root_of(&xml);
        let model = parse_chart_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("stock and line combo");

        assert_eq!(model.chart_type, "line");
        assert_eq!(
            model
                .plot_groups
                .as_ref()
                .expect("groups")
                .iter()
                .map(|group| group.kind.as_str())
                .collect::<Vec<_>>(),
            vec!["stock", "line"],
        );
        assert_eq!(model.stock_hi_low_lines, Some(true));
        assert_eq!(model.stock_hi_low_line_color.as_deref(), Some("123456"));
    }

    #[test]
    fn classic_plot_group_count_is_bounded_atomically() {
        let parse_group_count = |count: usize| {
            let groups = "<c:lineChart/>".repeat(count);
            let xml = format!(
                r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea>{groups}</c:plotArea></c:chart></c:chartSpace>"#,
            );
            let document = root_of(&xml);
            parse_chart_part(
                document.root_element(),
                &ChartParseContext {
                    color_resolver: Some(&FixtureResolver),
                    ..Default::default()
                },
            )
        };

        let exact = parse_group_count(MAX_CHART_PLOT_GROUPS).expect("exact group limit");
        assert_eq!(
            exact.plot_groups.expect("groups").len(),
            MAX_CHART_PLOT_GROUPS
        );
        assert!(parse_group_count(MAX_CHART_PLOT_GROUPS + 1).is_none());
    }

    #[test]
    fn empty_group_does_not_select_the_visible_legacy_family() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea>
              <c:barChart><c:barDir val="col"/></c:barChart>
              <c:pieChart><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:pieChart>
            </c:plotArea></c:chart></c:chartSpace>"#,
        );
        let document = root_of(&xml);
        let model = parse_chart_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("empty plus visible group");
        assert_eq!(model.chart_type, "pie");
        assert_eq!(model.series.len(), 1);
    }

    #[test]
    fn classic_plot_series_and_axes_are_bounded_before_projection() {
        let parse_series_count = |count: usize| {
            let series = "<c:ser/>".repeat(count);
            let xml = format!(
                r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea><c:lineChart><c:grouping val="standard"/>{series}</c:lineChart></c:plotArea></c:chart></c:chartSpace>"#,
            );
            let document = root_of(&xml);
            parse_chart_part(
                document.root_element(),
                &ChartParseContext {
                    color_resolver: Some(&FixtureResolver),
                    ..Default::default()
                },
            )
        };
        assert_eq!(
            parse_series_count(MAX_CHART_PLOT_SERIES)
                .expect("exact series limit")
                .series
                .len(),
            MAX_CHART_PLOT_SERIES,
        );
        assert!(parse_series_count(MAX_CHART_PLOT_SERIES + 1).is_none());

        let parse_axis_count = |count: usize| {
            let axes = [
                r#"<c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>"#,
                r#"<c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>"#,
                r#"<c:catAx><c:axId val="3"/><c:axPos val="t"/><c:crossAx val="4"/></c:catAx>"#,
                r#"<c:valAx><c:axId val="4"/><c:axPos val="r"/><c:crossAx val="3"/></c:valAx>"#,
                r#"<c:serAx><c:axId val="5"/><c:axPos val="r"/></c:serAx>"#,
            ][..count]
                .join("");
            let xml = format!(
                r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea><c:lineChart/>{axes}</c:plotArea></c:chart></c:chartSpace>"#,
            );
            let document = root_of(&xml);
            parse_chart_part(
                document.root_element(),
                &ChartParseContext {
                    color_resolver: Some(&FixtureResolver),
                    ..Default::default()
                },
            )
        };
        assert!(parse_axis_count(MAX_CHART_PLOT_AXES).is_some());
        assert!(parse_axis_count(MAX_CHART_PLOT_AXES + 1).is_none());
    }
}
