#[cfg(test)]
mod tests {
    use super::super::*;

    #[test]
    fn theme_reference_typeface_passes_through() {
        // A `+mn-lt` theme reference is returned verbatim (the renderer resolves
        // it against the theme font scheme).
        let xml = format!(
            r#"<c:valAx xmlns:c="{C_NS}" xmlns:a="{A_NS}">
                 <c:txPr><a:p><a:pPr><a:defRPr><a:latin typeface="+mn-lt"/></a:defRPr></a:pPr></a:p></c:txPr>
               </c:valAx>"#
        );
        assert_eq!(
            extract_axis_tick_label_face(root_of(&xml).root_element()).as_deref(),
            Some("+mn-lt")
        );
    }

    #[test]
    fn series_trendlines_parse() {
        // No trendline → None (byte-stable).
        let none = format!(r#"<c:ser xmlns:c="{C_NS}"/>"#);
        assert_eq!(
            extract_series_trendlines(root_of(&none).root_element(), &StubResolver),
            None
        );
        // A linear fit + a period-3 moving average, the linear one with a red line.
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}" xmlns:a="{A_NS}">
                 <c:trendline>
                   <c:name>Authored trend name</c:name>
                   <c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="ff0000"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr>
                   <c:trendlineType val="linear"/>
                   <c:dispEq val="1"/>
                   <c:dispRSqr val="1"/>
                   <c:trendlineLbl>
                     <c:numFmt formatCode="0.00" sourceLinked="0"/>
                     <c:layout><c:manualLayout><c:xMode val="edge"/><c:yMode val="edge"/><c:x val="0.1"/><c:y val="0.2"/></c:manualLayout></c:layout>
                     <c:tx><c:rich><a:bodyPr rot="1800000" wrap="none" anchor="b" lIns="12700"/><a:p><a:pPr algn="r"><a:defRPr sz="1800" b="1"><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:latin typeface="Georgia"/></a:defRPr></a:pPr><a:r><a:rPr sz="2000" b="0" i="1" lang="en-US" baseline="25000"><a:solidFill><a:srgbClr val="654321"/></a:solidFill></a:rPr><a:t>Authored</a:t></a:r></a:p><a:p><a:r><a:t>fit</a:t></a:r></a:p></c:rich></c:tx>
                     <c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="25400"><a:solidFill><a:srgbClr val="808080"/></a:solidFill></a:ln></c:spPr>
                     <c:txPr><a:bodyPr/><a:p><a:pPr algn="ctr"><a:defRPr sz="1800" b="1"><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:latin typeface="Georgia"/></a:defRPr></a:pPr></a:p></c:txPr>
                   </c:trendlineLbl>
                 </c:trendline>
                 <c:trendline>
                   <c:spPr><a:ln><a:noFill/></a:ln></c:spPr>
                   <c:trendlineType val="movingAvg"/>
                   <c:period val="3"/>
                 </c:trendline>
                 <c:trendline>
                   <c:trendlineType val="linear"/>
                   <c:trendlineLbl><c:tx><c:strRef><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Cached fit</c:v></c:pt></c:strCache></c:strRef></c:tx></c:trendlineLbl>
                 </c:trendline>
               </c:ser>"#
        );
        let got = extract_series_trendlines(root_of(&xml).root_element(), &StubResolver).unwrap();
        assert_eq!(got.len(), 3);
        assert_eq!(got[0].name.as_deref(), Some("Authored trend name"));
        assert_eq!(got[0].trendline_type, "linear");
        assert_eq!(got[0].line_color.as_deref(), Some("FF0000"));
        assert_eq!(got[0].line_width_emu, Some(19050));
        assert_eq!(got[0].line_dash.as_deref(), Some("dash"));
        assert_eq!(got[0].disp_eq, Some(true));
        assert_eq!(got[0].disp_r_sqr, Some(true));
        assert_eq!(got[0].label_text.as_deref(), Some("Authored\nfit"));
        let rich_runs = got[0]
            .label_rich_runs
            .as_ref()
            .expect("trendline rich runs");
        assert_eq!(
            rich_runs
                .iter()
                .map(|run| run.text.as_str())
                .collect::<String>(),
            "Authored\nfit"
        );
        assert_eq!(rich_runs[0].paragraph_align.as_deref(), Some("r"));
        assert_eq!(rich_runs[0].italic, Some(true));
        assert_eq!(rich_runs[2].paragraph_align, None);
        assert_eq!(got[0].label_format_code.as_deref(), Some("0.00"));
        assert_eq!(got[0].label_format_source_linked, Some(false));
        assert_eq!(got[0].label_font_size_hpt, Some(1800));
        assert_eq!(got[0].label_font_bold, Some(true));
        assert_eq!(got[0].label_font_italic, Some(false));
        assert_eq!(got[0].label_font_color.as_deref(), Some("123456"));
        assert_eq!(got[0].label_font_face.as_deref(), Some("Georgia"));
        assert_eq!(got[0].label_font_language, None);
        assert_eq!(got[0].label_font_baseline, None);
        assert_eq!(got[0].label_text_rotation, Some(1_800_000));
        assert_eq!(got[0].label_text_wrap.as_deref(), Some("none"));
        assert_eq!(got[0].label_text_vertical_anchor.as_deref(), Some("b"));
        assert_eq!(got[0].label_text_l_ins_emu, Some(12_700));
        assert_eq!(got[0].label_text_align.as_deref(), Some("ctr"));
        let label_box = got[0].label_box.as_ref().expect("trendline label box");
        assert_eq!(label_box.fill.as_deref(), Some("FFFFFF"));
        assert_eq!(label_box.border_color.as_deref(), Some("808080"));
        assert_eq!(label_box.border_width_emu, Some(25400));
        let manual = got[0]
            .label_manual_layout
            .as_ref()
            .expect("trendline label manual layout");
        assert_eq!(manual.x_mode, "edge");
        assert_eq!(manual.y_mode, "edge");
        assert_eq!(manual.x, 0.1);
        assert_eq!(manual.y, 0.2);
        assert_eq!(got[1].trendline_type, "movingAvg");
        assert_eq!(got[1].period, Some(3));
        assert_eq!(got[1].line_color, None);
        assert_eq!(got[0].line_hidden, None);
        assert_eq!(got[1].line_hidden, Some(true));
        assert_eq!(got[2].label_text.as_deref(), Some("Cached fit"));
    }

    #[test]
    fn ct_boolean_chart_level_marker_bare_enables_markers() {
        // §21.2.2.33 `<c:lineChart><c:marker/>` ⇒ true ⇒ line series show markers
        // even without a per-series `<c:marker>`. A bare element must read true.
        let bare = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:chart>
                <c:plotArea><c:lineChart>
                  <c:marker/>
                  <c:ser><c:idx val="0"/>
                    <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>
                </c:lineChart></c:plotArea>
              </c:chart></c:chartSpace>"#
        );
        let d = chart_space_of(&bare);
        let m = parse_chart_part(
            d.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("chart");
        assert_eq!(
            m.series[0].show_marker,
            Some(true),
            "bare <c:marker/> ⇒ markers enabled"
        );

        // Control: `<c:marker val="0"/>` ⇒ markers off for a line series.
        let off = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:chart>
                <c:plotArea><c:lineChart>
                  <c:marker val="0"/>
                  <c:ser><c:idx val="0"/>
                    <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>
                </c:lineChart></c:plotArea>
              </c:chart></c:chartSpace>"#
        );
        let d2 = chart_space_of(&off);
        let m2 = parse_chart_part(
            d2.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("chart");
        assert_eq!(
            m2.series[0].show_marker,
            Some(false),
            "<c:marker val=\"0\"/> ⇒ markers off"
        );
    }

    #[test]
    fn parse_chart_part_bar_scatter_combo_keeps_xy_sources_and_both_numeric_axes() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:chart><c:plotArea>
                <c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/>
                  <c:ser><c:idx val="0"/>
                    <c:cat><c:strLit><c:pt idx="0"><c:v>Top</c:v></c:pt><c:pt idx="1"><c:v>Bottom</c:v></c:pt></c:strLit></c:cat>
                    <c:val><c:numLit><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>0</c:v></c:pt></c:numLit></c:val>
                  </c:ser><c:axId val="1"/><c:axId val="2"/>
                </c:barChart>
                <c:scatterChart><c:scatterStyle val="marker"/>
                  <c:ser><c:idx val="1"/>
                    <c:marker><c:symbol val="circle"/></c:marker>
                    <c:xVal><c:numLit><c:formatCode>0%</c:formatCode><c:pt idx="0" formatCode="0%"><c:v>0.15</c:v></c:pt><c:pt idx="1" formatCode="0.0%"><c:v>0.83</c:v></c:pt></c:numLit></c:xVal>
                    <c:yVal><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt></c:numLit></c:yVal>
                  </c:ser><c:axId val="3"/><c:axId val="4"/>
                </c:scatterChart>
                <c:catAx><c:axId val="1"/><c:axPos val="l"/><c:tickLblSkip val="2"/><c:tickMarkSkip val="3"/></c:catAx>
                <c:valAx><c:axId val="2"/><c:axPos val="b"/><c:scaling><c:max val="1.4"/></c:scaling></c:valAx>
                <c:valAx><c:axId val="3"/><c:axPos val="b"/><c:delete val="1"/></c:valAx>
                <c:valAx><c:axId val="4"/><c:axPos val="r"/><c:delete val="1"/><c:scaling><c:min val="0"/><c:max val="2"/></c:scaling></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>"#
        );
        let doc = chart_space_of(&xml);
        let model = parse_chart_part(
            doc.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("bar/scatter combo parses");

        assert_eq!(model.chart_type, "clusteredBarH");
        assert_eq!(model.cat_axis_tick_label_skip, Some(2));
        assert_eq!(model.cat_axis_tick_mark_skip, Some(3));
        let scatter = &model.series[1];
        assert_eq!(scatter.series_type.as_deref(), Some("scatter"));
        assert_eq!(
            scatter.categories.as_deref(),
            Some(&["0.15".into(), "0.83".into()][..])
        );
        assert_eq!(scatter.values, vec![Some(2.0), Some(1.0)]);
        assert_eq!(scatter.use_secondary_axis, Some(true));
        assert_eq!(scatter.show_marker, Some(true));
        assert_eq!(scatter.cat_format_code.as_deref(), Some("0%"));
        assert_eq!(
            scatter.cat_format_codes.as_deref(),
            Some(&[Some("0%".into()), Some("0.0%".into())][..])
        );
        assert_eq!(
            model.secondary_cat_axis.as_ref().and_then(|axis| axis.max),
            None
        );
        let y_axis = model.secondary_val_axis.expect("scatter Y axis parsed");
        assert_eq!((y_axis.min, y_axis.max), (Some(0.0), Some(2.0)));
    }

    #[test]
    fn parse_marker_block_none_node_returns_all_none() {
        assert_eq!(
            parse_marker_block(None, &FixtureResolver),
            (None, None, None, None, None, None, None, None)
        );
    }

    #[test]
    fn parse_marker_block_symbol_none_no_sppr() {
        let xml = format!(r#"<c:marker xmlns:c="{C_NS}"><c:symbol val="none"/></c:marker>"#);
        let d = root_of(&xml);
        let (symbol, size, fill, fill_paint, fill_authored, line, line_width_emu, line_authored) =
            parse_marker_block(Some(d.root_element()), &FixtureResolver);
        assert_eq!(symbol.as_deref(), Some("none"));
        assert_eq!(size, None);
        assert_eq!(fill, None);
        assert_eq!(fill_paint, None);
        assert_eq!(fill_authored, None);
        assert_eq!(line, None);
        assert_eq!(line_width_emu, None);
        assert_eq!(line_authored, None);
    }

    #[test]
    fn parse_marker_block_ignores_schema_invalid_sibling_duotone() {
        struct Images;
        impl ChartImageResolver for Images {
            fn resolve_image(
                &self,
                source: ChartImageSource,
                relationship_id: &str,
            ) -> Option<(String, String)> {
                (source == ChartImageSource::Chart && relationship_id == "rId1")
                    .then(|| ("xl/media/marker.png".to_owned(), "image/png".to_owned()))
            }
        }
        let xml = format!(
            r#"<c:marker xmlns:c="{C_NS}" xmlns:a="{A_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <c:symbol val="picture"/><c:spPr><a:blipFill><a:blip r:embed="rId1"/>
                <a:duotone><a:srgbClr val="000000"/><a:srgbClr val="FFFFFF"/></a:duotone>
                <a:stretch/></a:blipFill></c:spPr>
            </c:marker>"#
        );
        let document = root_of(&xml);
        let (_, _, _, paint, _, _, _, _) = parse_marker_block_with_images(
            Some(document.root_element()),
            &FixtureResolver,
            &Images,
        );
        let Some(ChartStyleFill::Image { duotone, .. }) = paint else {
            panic!("expected resolved image fill");
        };
        assert_eq!(duotone, None);
    }

    #[test]
    fn parse_marker_block_rejects_gradient_beyond_resource_ceiling_before_expansion() {
        let stops = (0..=MAX_CHART_MARKER_GRADIENT_STOPS)
            .map(|index| format!(r#"<a:gs pos="{}"><a:srgbClr val="112233"/></a:gs>"#, index))
            .collect::<String>();
        let xml = format!(
            r#"<c:marker xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:symbol val="circle"/>
              <c:spPr><a:gradFill><a:gsLst>{stops}</a:gsLst></a:gradFill></c:spPr>
            </c:marker>"#
        );
        let d = root_of(&xml);
        let (_, _, fill, fill_paint, fill_authored, _, _, _) =
            parse_marker_block(Some(d.root_element()), &FixtureResolver);
        assert_eq!(fill, None);
        assert_eq!(fill_paint, None);
        assert_eq!(fill_authored, Some(true));
    }

    #[test]
    fn parse_error_bars_fixed_val_both_directions() {
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:errBars>
                <c:errDir val="y"/>
                <c:errBarType val="both"/>
                <c:errValType val="fixedVal"/>
                <c:val val="2.5"/>
                <c:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="333333"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr>
              </c:errBars>
            </c:ser>"#
        );
        let d = root_of(&xml);
        let values = vec![Some(10.0), Some(20.0), None];
        let bars = parse_error_bars(d.root_element(), &values, &FixtureResolver);
        assert_eq!(bars.len(), 1);
        let b = &bars[0];
        assert_eq!(b.dir, "y");
        assert_eq!(b.bar_type, "both");
        assert_eq!(b.plus, vec![Some(2.5), Some(2.5), Some(2.5)]);
        assert_eq!(b.minus, vec![Some(2.5), Some(2.5), Some(2.5)]);
        assert!(!b.no_end_cap);
        assert_eq!(b.color.as_deref(), Some("333333"));
        assert_eq!(b.line_width_emu, Some(12700));
        assert_eq!(b.dash.as_deref(), Some("dash"));
    }

    #[test]
    fn parse_error_bars_percentage_scales_per_point() {
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}">
              <c:errBars>
                <c:errDir val="x"/>
                <c:errBarType val="plus"/>
                <c:errValType val="percentage"/>
                <c:val val="10"/>
                <c:noEndCap val="1"/>
              </c:errBars>
            </c:ser>"#
        );
        let d = root_of(&xml);
        let values = vec![Some(100.0), Some(-50.0), None];
        let bars = parse_error_bars(d.root_element(), &values, &FixtureResolver);
        assert_eq!(bars.len(), 1);
        let b = &bars[0];
        assert_eq!(b.dir, "x");
        assert!(b.no_end_cap);
        // 10% of |value|; the None slot stays None (nothing to scale).
        assert_eq!(b.plus, vec![Some(10.0), Some(5.0), None]);
        assert_eq!(b.minus, vec![Some(10.0), Some(5.0), None]);
    }

    #[test]
    fn parse_error_bars_absent_returns_empty() {
        let xml = format!(r#"<c:ser xmlns:c="{C_NS}"><c:val/></c:ser>"#);
        let d = root_of(&xml);
        assert!(parse_error_bars(d.root_element(), &[], &FixtureResolver).is_empty());
    }

    /// Sparse `<c:pt idx>` cache: `ptCount=11` but only two points are present
    /// (`idx=1` and `idx=9`). The result must be sized to the declared
    /// `ptCount`, not to the number of `<c:pt>` elements present, and every
    /// unlisted index must stay the empty-string placeholder (not shifted).
    #[test]
    fn collect_str_cache_positional_sparse_ptcount_and_gaps() {
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}">
              <c:cat><c:strCache>
                <c:ptCount val="11"/>
                <c:pt idx="1"><c:v>Feb</c:v></c:pt>
                <c:pt idx="9"><c:v>Oct</c:v></c:pt>
              </c:strCache></c:cat>
            </c:ser>"#
        );
        let d = root_of(&xml);
        let cats = collect_str_cache_positional(d.root_element(), "cat");
        assert_eq!(cats.len(), 11);
        assert_eq!(cats[0], "");
        assert_eq!(cats[1], "Feb");
        assert_eq!(cats[9], "Oct");
        assert_eq!(cats[10], "");
    }

    #[test]
    fn collect_str_cache_positional_missing_container_is_empty() {
        let xml = format!(r#"<c:ser xmlns:c="{C_NS}"></c:ser>"#);
        let d = root_of(&xml);
        assert!(collect_str_cache_positional(d.root_element(), "cat").is_empty());
    }

    /// Companion numeric collector: same sparse/idx=1-start shape, but with
    /// `None` gaps instead of empty strings, and one genuinely missing `<c:v>`
    /// (idx present, value absent) which must also collapse to `None`.
    #[test]
    fn collect_num_cache_positional_sparse_ptcount_and_gaps() {
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}">
              <c:val><c:numCache>
                <c:ptCount val="11"/>
                <c:pt idx="1"><c:v>42</c:v></c:pt>
                <c:pt idx="9"><c:v>7</c:v></c:pt>
              </c:numCache></c:val>
            </c:ser>"#
        );
        let d = root_of(&xml);
        let vals = collect_num_cache_positional(d.root_element(), "val");
        assert_eq!(vals.len(), 11);
        assert_eq!(vals[0], None);
        assert_eq!(vals[1], Some(42.0));
        assert_eq!(vals[9], Some(7.0));
        assert_eq!(vals[10], None);
    }

    #[test]
    fn collect_num_cache_positional_missing_container_is_empty() {
        let xml = format!(r#"<c:ser xmlns:c="{C_NS}"></c:ser>"#);
        let d = root_of(&xml);
        assert!(collect_num_cache_positional(d.root_element(), "val").is_empty());
    }

    #[test]
    fn ofpie_custom_split_is_atomic_at_the_parser_resource_boundary() {
        let exact = roxmltree::Document::parse(
            r#"<ofPieChart><custSplit><secondPiePt val="0"/><secondPiePt val="1"/></custSplit></ofPieChart>"#,
        )
        .expect("valid XML");
        assert_eq!(
            parse_of_pie_custom_split_with_limit(exact.root_element(), 2),
            Some(vec![0, 1])
        );

        let plus_one = roxmltree::Document::parse(
            r#"<ofPieChart><custSplit><secondPiePt val="0"/><secondPiePt val="1"/><secondPiePt val="2"/></custSplit></ofPieChart>"#,
        )
        .expect("valid XML");
        assert_eq!(
            parse_of_pie_custom_split_with_limit(plus_one.root_element(), 2),
            None
        );
    }

    #[test]
    fn parse_chart_part_resolves_cacheless_bubble_fields() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea><c:bubbleChart>
              <c:ser><c:idx val="0"/><c:order val="0"/>
                <c:tx><c:strRef><c:f>Name</c:f></c:strRef></c:tx>
                <c:xVal><c:numRef><c:f>X</c:f></c:numRef></c:xVal>
                <c:yVal><c:numRef><c:f>Y</c:f></c:numRef></c:yVal>
                <c:bubbleSize><c:numRef><c:f>Size</c:f></c:numRef></c:bubbleSize>
              </c:ser>
              <c:bubbleScale val="40"/>
              <c:sizeRepresents val="w"/>
              <c:showNegBubbles/>
            </c:bubbleChart></c:plotArea></c:chart></c:chartSpace>"#
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
        .expect("bubble chart parses");

        assert_eq!(chart.categories, vec!["1", "2"]);
        assert_eq!(chart.series[0].name, "Resolved series");
        assert_eq!(chart.series[0].categories, None);
        assert_eq!(chart.series[0].values, vec![Some(10.0), Some(20.0)]);
        assert_eq!(
            chart.series[0].bubble_sizes,
            Some(vec![Some(3.0), Some(5.0)])
        );
        assert_eq!(chart.bubble_scale, Some(40));
        assert_eq!(chart.bubble_size_represents.as_deref(), Some("w"));
        assert_eq!(chart.show_negative_bubbles, Some(true));

        // Strict OOXML uses the percentage lexical form; Transitional accepts
        // both this and the integer form above.
        let strict_xml = xml.replace(
            "<c:bubbleScale val=\"40\"/>",
            "<c:bubbleScale val=\"40%\"/>",
        );
        let strict_doc = root_of(&strict_xml);
        let mut strict_references = FormulaResolver;
        let strict_chart = parse_chart_part(
            strict_doc.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                references: std::cell::Cell::new(Some(&mut strict_references)),
                ..Default::default()
            },
        )
        .expect("strict bubble scale parses");
        assert_eq!(strict_chart.bubble_scale, Some(40));

        let disabled_xml = xml.replace("<c:showNegBubbles/>", "<c:showNegBubbles val=\"false\"/>");
        let disabled_doc = root_of(&disabled_xml);
        let mut disabled_references = FormulaResolver;
        let disabled_chart = parse_chart_part(
            disabled_doc.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                references: std::cell::Cell::new(Some(&mut disabled_references)),
                ..Default::default()
            },
        )
        .expect("explicit false showNegBubbles parses");
        assert_eq!(disabled_chart.show_negative_bubbles, Some(false));
    }

    #[test]
    fn parse_chart_part_keeps_unresolved_cacheless_bubble_sizes_absent() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea><c:bubbleChart>
              <c:ser><c:idx val="0"/><c:order val="0"/>
                <c:xVal><c:numRef><c:f>X</c:f></c:numRef></c:xVal>
                <c:yVal><c:numRef><c:f>Y</c:f></c:numRef></c:yVal>
                <c:bubbleSize><c:numRef><c:f>Size</c:f></c:numRef></c:bubbleSize>
              </c:ser>
            </c:bubbleChart></c:plotArea></c:chart></c:chartSpace>"#
        );
        let doc = root_of(&xml);
        let chart = parse_chart_part(
            doc.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("cacheless bubble chart still parses");

        assert_eq!(chart.categories, Vec::<String>::new());
        assert_eq!(chart.series[0].categories, None);
        assert_eq!(chart.series[0].values, Vec::<Option<f64>>::new());
        assert_eq!(chart.series[0].bubble_sizes, None);
        assert_eq!(chart.show_negative_bubbles, None);
    }

    #[test]
    fn parse_chart_part_preserves_authored_multilevel_cache() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea><c:barChart>
              <c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/>
                <c:cat><c:multiLvlStrRef><c:f>CachedCats</c:f><c:multiLvlStrCache>
                  <c:ptCount val="2"/><c:lvl><c:pt idx="0"><c:v>Authored A</c:v></c:pt><c:pt idx="1"><c:v>Authored B</c:v></c:pt></c:lvl>
                  <c:lvl><c:pt idx="0"><c:v>Outer</c:v></c:pt></c:lvl>
                </c:multiLvlStrCache></c:multiLvlStrRef></c:cat>
                <c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val>
              </c:ser>
            </c:barChart><c:catAx><c:axPos val="b"/><c:noMultiLvlLbl val="0"/></c:catAx>
            </c:plotArea></c:chart></c:chartSpace>"#
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
        .expect("bar chart parses");

        assert_eq!(chart.categories, vec!["Authored A", "Authored B"]);
        assert_eq!(
            chart.category_levels,
            Some(vec![
                vec!["Authored A".to_string(), "Authored B".to_string()],
                vec!["Outer".to_string(), String::new()],
            ])
        );
        assert_eq!(chart.cat_axis_no_multi_level_labels, Some(false));
        assert_eq!(chart.series[0].categories, None);
    }

    #[test]
    fn plot_visible_only_preserves_source_masks_and_precedes_statistical_error_bars() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea><c:lineChart>
              <c:ser><c:idx val="0"/><c:order val="0"/>
                <c:cat><c:strRef><c:f>Cats</c:f><c:strCache><c:ptCount val="4"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt><c:pt idx="3"><c:v>D</c:v></c:pt></c:strCache></c:strRef></c:cat>
                <c:val><c:numRef><c:f>Values</c:f><c:numCache><c:ptCount val="4"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>100</c:v></c:pt><c:pt idx="3"><c:v>40</c:v></c:pt></c:numCache></c:numRef></c:val>
                <c:errBars><c:errDir val="y"/><c:errBarType val="both"/><c:errValType val="stdDev"/><c:val val="1"/></c:errBars>
              </c:ser>
            </c:lineChart></c:plotArea><c:plotVisOnly/></c:chart></c:chartSpace>"#
        );
        let doc = root_of(&xml);
        let mut references = HiddenSourceResolver;
        let chart = parse_chart_part(
            doc.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                references: std::cell::Cell::new(Some(&mut references)),
                ..Default::default()
            },
        )
        .expect("line chart parses");

        assert_eq!(chart.plot_visible_only, Some(true));
        assert_eq!(
            chart.category_source_hidden,
            Some(vec![false, true, false, false])
        );
        assert_eq!(
            chart.series[0].source_hidden,
            Some(vec![false, false, true, false])
        );
        let errors = chart.series[0].err_bars.as_ref().expect("error bars");
        assert_eq!(errors[0].plus, vec![Some(15.0); 4]);
        assert_eq!(errors[0].minus, vec![Some(15.0); 4]);
    }

    #[test]
    fn multi_level_category_cache_rejects_unbounded_aggregate_slots() {
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}"><c:cat><c:multiLvlStrRef>
              <c:multiLvlStrCache><c:ptCount val="1048576"/><c:lvl/><c:lvl/></c:multiLvlStrCache>
            </c:multiLvlStrRef></c:cat></c:ser>"#
        );
        let document = root_of(&xml);
        assert_eq!(
            collect_multi_level_str_cache(document.root_element(), "cat"),
            None,
        );
    }

    #[test]
    fn classic_plot_groups_retain_exact_kind_source_order_and_empty_ranges() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea>
              <c:areaChart><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:areaChart>
              <c:area3DChart/><c:lineChart/><c:line3DChart/><c:stockChart/>
              <c:radarChart/><c:scatterChart/><c:pieChart/><c:pie3DChart/>
              <c:doughnutChart/><c:barChart/><c:bar3DChart/><c:ofPieChart/>
              <c:surfaceChart/><c:surface3DChart/><c:bubbleChart/>
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
        .expect("classic chart groups");
        let groups = model.plot_groups.expect("ordered groups");
        assert_eq!(
            groups
                .iter()
                .map(|group| group.kind.as_str())
                .collect::<Vec<_>>(),
            vec![
                "area",
                "area3D",
                "line",
                "line3D",
                "stock",
                "radar",
                "scatter",
                "pie",
                "pie3D",
                "doughnut",
                "bar",
                "bar3D",
                "ofPie",
                "surface",
                "surface3D",
                "bubble",
            ],
        );
        assert_eq!((groups[0].series_start, groups[0].series_count), (0, 1));
        assert!(groups[1..]
            .iter()
            .all(|group| (group.series_start, group.series_count) == (1, 0)));
    }
}
