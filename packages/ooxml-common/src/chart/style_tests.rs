#[cfg(test)]
mod tests {
    use super::super::*;

    #[test]
    fn chart_style_relationship_accepts_office_and_documented_revisions_only() {
        assert!(is_chart_style_relationship_type(
            "http://schemas.microsoft.com/office/2011/relationships/chartStyle"
        ));
        assert!(is_chart_style_relationship_type(
            "http://schemas.microsoft.com/office/2012/relationships/chartStyle"
        ));
        assert!(!is_chart_style_relationship_type(
            "https://example.invalid/office/2012/relationships/chartStyle"
        ));
    }

    #[test]
    fn secondary_chart_text_carriers_leave_empty_properties_style_inheritable() {
        let table_xml = format!(
            r#"<c:plotArea xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:dTable><c:txPr><a:p><a:pPr><a:defRPr/></a:pPr></a:p></c:txPr></c:dTable></c:plotArea>"#
        );
        let table_document = root_of(&table_xml);
        let table = extract_chart_data_table(table_document.root_element(), &StubResolver).unwrap();
        assert_eq!((table.font_bold, table.font_italic), (None, None));

        let units_xml = format!(
            r#"<c:valAx xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:dispUnits><c:builtInUnit val="thousands"/><c:dispUnitsLbl><c:txPr><a:p><a:pPr><a:defRPr/></a:pPr></a:p></c:txPr></c:dispUnitsLbl></c:dispUnits></c:valAx>"#
        );
        let units_document = root_of(&units_xml);
        let units = parse_axis_display_units(units_document.root_element(), &StubResolver).unwrap();
        let label = units.label.unwrap();
        assert_eq!((label.font_bold, label.font_italic), (None, None));

        let trendline_xml = format!(
            r#"<c:ser xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:trendline><c:trendlineType val="linear"/><c:trendlineLbl><c:txPr><a:p><a:pPr><a:defRPr/></a:pPr></a:p></c:txPr></c:trendlineLbl></c:trendline></c:ser>"#
        );
        let trendline_document = root_of(&trendline_xml);
        let trendline = extract_series_trendlines(trendline_document.root_element(), &StubResolver)
            .unwrap()
            .remove(0);
        assert_eq!(
            (trendline.label_font_bold, trendline.label_font_italic),
            (None, None)
        );

        let legend_xml = format!(
            r#"<c:chart xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:legend><c:legendEntry><c:idx val="0"/><c:txPr><a:p><a:pPr><a:defRPr/></a:pPr></a:p></c:txPr></c:legendEntry></c:legend></c:chart>"#
        );
        let legend_document = root_of(&legend_xml);
        let entry = extract_legend_overrides(legend_document.root_element(), &StubResolver)
            .1
            .unwrap()
            .remove(0);
        assert_eq!((entry.font_bold, entry.font_italic), (None, None));

        let chartex_xml = format!(
            r#"<cx:series xmlns:cx="{CX_NS}" xmlns:a="{A_NS}"><cx:dataLabels><cx:txPr><a:p><a:pPr><a:defRPr/></a:pPr></a:p></cx:txPr></cx:dataLabels></cx:series>"#
        );
        let chartex_document = root_of(&chartex_xml);
        let defaults =
            parse_chartex_series_labels(chartex_document.root_element(), 1, &StubResolver, true)
                .2
                .unwrap();
        assert_eq!((defaults.font_bold, defaults.font_italic), (None, None));
    }

    #[test]
    fn chart_space_border_nofill_color_none_width_kept() {
        let xml = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:spPr><a:ln w="12700"><a:noFill/></a:ln></c:spPr>
        </c:chartSpace>"#;
        let d = root_of(xml);
        // noFill turns the border off → color None, but @w is still reported.
        assert_eq!(
            extract_chart_space_border(d.root_element()),
            (None, Some(12700))
        );
    }

    #[test]
    fn chart_space_structured_fill_preserves_gradient_and_pattern_recipes() {
        let gradient = root_of(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:spPr><a:gradFill rotWithShape="0"><a:gsLst><a:gs pos="0"><a:srgbClr val="112233"/></a:gs><a:gs pos="100000"><a:srgbClr val="AABBCC"/></a:gs></a:gsLst><a:lin ang="2700000" scaled="1"/></a:gradFill></c:spPr></c:chartSpace>"#,
        );
        assert!(matches!(
            extract_direct_shape_fill(child(gradient.root_element(), "spPr"), &StubResolver).fill,
            Some(ChartStyleFill::Gradient {
                angle,
                rot_with_shape: Some(false),
                ..
            }) if (angle - 45.0).abs() < 1e-9
        ));

        let pattern = root_of(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:spPr><a:pattFill prst="diagCross"><a:fgClr><a:srgbClr val="112233"/></a:fgClr><a:bgClr><a:srgbClr val="AABBCC"/></a:bgClr></a:pattFill></c:spPr></c:chartSpace>"#,
        );
        assert_eq!(
            extract_direct_shape_fill(child(pattern.root_element(), "spPr"), &StubResolver).fill,
            Some(ChartStyleFill::Pattern {
                fg: "112233".to_string(),
                bg: "AABBCC".to_string(),
                preset: "diagCross".to_string(),
            })
        );
    }

    #[test]
    fn direct_line_preserves_structured_paint_and_bare_dash_choice() {
        let document = root_of(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:spPr><a:ln><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="112233"/></a:gs><a:gs pos="100000"><a:srgbClr val="AABBCC"/></a:gs></a:gsLst><a:lin ang="2700000"/></a:gradFill><a:prstDash/></a:ln></c:spPr></c:chartSpace>"#,
        );
        let line = extract_direct_shape_line(document.root_element(), &StubResolver);
        assert!(matches!(
            line.fill,
            Some(ChartStyleFill::Gradient { angle, .. }) if (angle - 45.0).abs() < 1e-9
        ));
        assert_eq!(line.color, None);
        assert_eq!(line.dash, None);
        assert_eq!(line.dash_authored, Some(true));
    }

    #[test]
    fn parse_chart_part_resolves_direct_chart_and_plot_picture_fills() {
        struct Images;
        impl ChartImageResolver for Images {
            fn resolve_image(
                &self,
                source: ChartImageSource,
                relationship_id: &str,
            ) -> Option<(String, String)> {
                if source != ChartImageSource::Chart {
                    return None;
                }
                match relationship_id {
                    "rChart" => Some(("xl/media/chart.png".to_owned(), "image/png".to_owned())),
                    "rPlot" => Some(("xl/media/plot.png".to_owned(), "image/png".to_owned())),
                    _ => None,
                }
            }
        }
        let document = root_of(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <c:chart><c:plotArea>
                <c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart>
                <c:spPr><a:blipFill rotWithShape="1"><a:blip r:embed="rPlot"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></c:spPr>
              </c:plotArea></c:chart>
              <c:spPr><a:blipFill dpi="0" rotWithShape="0"><a:blip r:embed="rChart"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill></c:spPr>
            </c:chartSpace>"#,
        );
        let model = parse_chart_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: None,
                color_style_xml: None,
                images: Some(&Images),
                ..Default::default()
            },
        )
        .expect("classic chart parses");

        assert!(matches!(
            model.chart_fill,
            Some(ChartStyleFill::Image { ref image_path, stretch: true, .. })
                if image_path == "xl/media/chart.png"
        ));
        assert!(matches!(
            model.plot_area_fill,
            Some(ChartStyleFill::Image { ref image_path, stretch: true, .. })
                if image_path == "xl/media/plot.png"
        ));
    }

    #[test]
    fn parse_chart_part_preserves_direct_plot_area_fill_provenance() {
        let gradient = root_of(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart><c:spPr><a:gradFill rotWithShape="0"><a:gsLst><a:gs pos="0"><a:srgbClr val="112233"/></a:gs><a:gs pos="100000"><a:srgbClr val="AABBCC"/></a:gs></a:gsLst><a:lin ang="2700000"/></a:gradFill><a:ln w="12700" cap="rnd" cmpd="thinThick"><a:solidFill><a:srgbClr val="445566"/></a:solidFill><a:custDash><a:ds d="125000" sp="75000"/></a:custDash><a:bevel/></a:ln></c:spPr></c:plotArea></c:chart></c:chartSpace>"#,
        );
        let model = parse_chart_part(
            gradient.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("classic chart parses");
        assert!(matches!(
            model.plot_area_fill,
            Some(ChartStyleFill::Gradient {
                angle,
                rot_with_shape: Some(false),
                ..
            }) if (angle - 45.0).abs() < 1e-9
        ));
        assert_eq!(model.plot_area_fill_hidden, None);
        assert_eq!(model.plot_area_fill_paint_authored, Some(true));
        assert_eq!(model.plot_area_line_color.as_deref(), Some("445566"));
        assert_eq!(model.plot_area_line_width_emu, Some(12700));
        assert_eq!(model.plot_area_line_dash, None);
        assert_eq!(
            model.plot_area_line_custom_dash.as_deref(),
            Some(
                &[ChartLineDashSegment {
                    dash: 1.25,
                    space: 0.75,
                }][..]
            ),
        );
        assert_eq!(model.plot_area_line_cap.as_deref(), Some("rnd"));
        assert_eq!(model.plot_area_line_join.as_deref(), Some("bevel"));
        assert_eq!(model.plot_area_line_compound.as_deref(), Some("thinThick"));
        assert_eq!(model.plot_area_line_hidden, None);
        assert_eq!(model.plot_area_line_paint_authored, Some(true));

        let no_fill = root_of(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea></c:chart></c:chartSpace>"#,
        );
        let model = parse_chart_part(
            no_fill.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("classic chart parses");
        assert_eq!(model.plot_area_bg, None);
        assert_eq!(model.plot_area_fill, None);
        assert_eq!(model.plot_area_fill_hidden, Some(true));
        assert_eq!(model.plot_area_fill_paint_authored, Some(true));
        assert_eq!(model.plot_area_line_color, None);
        assert_eq!(model.plot_area_line_hidden, Some(true));
        assert_eq!(model.plot_area_line_paint_authored, Some(true));
    }

    /// §21.2.2.30 / `CT_ChartSpace`: chart-local `clrMapOvr` is a direct
    /// `CT_ColorMapping`. It replaces the application's logical color mapping,
    /// so an authored `schemeClr accent1` resolves through the declared
    /// `accent1=accent2` slot mapping for both pptx and xlsx callers.
    #[test]
    fn chart_color_map_override_remaps_explicit_and_default_series_accents() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:clrMapOvr bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"
                accent1="accent2" accent2="accent2" accent3="accent3"
                accent4="accent4" accent5="accent5" accent6="accent6"
                hlink="hlink" folHlink="folHlink"/>
              <c:chart><c:plotArea><c:barChart>
                <c:barDir val="col"/><c:grouping val="clustered"/>
                <c:ser><c:idx val="0"/><c:order val="0"/>
                  <c:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></c:spPr>
                  <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
                </c:ser>
                <c:ser><c:idx val="6"/><c:order val="1"/>
                  <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val>
                </c:ser>
              </c:barChart></c:plotArea></c:chart>
            </c:chartSpace>"#
        );
        let doc = chart_space_of(&xml);
        let chart = parse_chart_part(
            doc.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("chart parses");

        assert_eq!(chart.series[0].color.as_deref(), Some("ED7D31"));
        assert_eq!(chart.series[1].color.as_deref(), Some("ED7D31"));
    }

    #[test]
    fn area_series_no_fill_stays_transparent_but_remains_in_the_stack() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:chart><c:plotArea><c:areaChart>
                <c:grouping val="stacked"/>
                <c:ser><c:idx val="0"/><c:order val="0"/>
                  <c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>
                  <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:pt idx="0"><c:v>80</c:v></c:pt></c:numLit></c:val>
                </c:ser>
              </c:areaChart>
              <c:lineChart><c:grouping val="standard"/>
                <c:ser><c:idx val="1"/><c:order val="1"/>
                  <c:spPr><a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></c:spPr>
                  <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:pt idx="0"><c:v>85</c:v></c:pt></c:numLit></c:val>
                </c:ser>
              </c:lineChart></c:plotArea></c:chart>
            </c:chartSpace>"#
        );
        let doc = chart_space_of(&xml);
        let chart = parse_chart_part(
            doc.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("chart parses");

        assert_eq!(chart.chart_type, "stackedArea");
        assert_eq!(chart.series[0].series_type.as_deref(), Some("area"));
        assert_eq!(chart.series[0].color.as_deref(), Some("00000000"));
        assert_eq!(chart.series[0].line_hidden, Some(true));
        assert_eq!(chart.series[0].values, vec![Some(80.0)]);
        assert_eq!(chart.series[1].series_type.as_deref(), Some("line"));
        assert_eq!(chart.series[1].line_color.as_deref(), Some("000000"));
    }

    #[test]
    fn chart_space_shape_properties_without_a_fill_keep_the_host_default() {
        let chart_xml = |shape_properties: &str| {
            format!(
                r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
                  <c:chart><c:plotArea><c:barChart>
                    <c:barDir val="col"/><c:grouping val="clustered"/>
                    <c:ser><c:idx val="0"/><c:order val="0"/>
                      <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                      <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
                    </c:ser>
                  </c:barChart></c:plotArea></c:chart>
                  {shape_properties}
                </c:chartSpace>"#
            )
        };

        // `spPr` is often present only to suppress the chart border. Since it
        // carries no fill choice, it must not also make the chart area
        // transparent.
        let line_only = chart_xml("<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>");
        let doc = chart_space_of(&line_only);
        let parsed = parse_chart_part(
            doc.root_element(),
            &ChartParseContext {
                color_resolver: Some(&WhiteChartFixtureResolver),
                ..Default::default()
            },
        )
        .expect("line-only chart parses");
        assert_eq!(parsed.chart_bg.as_deref(), Some("FFFFFF"));
        assert_eq!(parsed.chart_fill_paint_authored, None);
        assert_eq!(parsed.chart_border_hidden, Some(true));
        assert_eq!(parsed.chart_border_paint_authored, Some(true));

        // An explicit fill choice remains authoritative.
        let no_fill = chart_xml("<c:spPr><a:noFill/></c:spPr>");
        let doc = chart_space_of(&no_fill);
        let parsed = parse_chart_part(
            doc.root_element(),
            &ChartParseContext {
                color_resolver: Some(&WhiteChartFixtureResolver),
                ..Default::default()
            },
        )
        .expect("noFill chart parses");
        assert_eq!(parsed.chart_bg, None);
        assert_eq!(parsed.chart_fill_hidden, Some(true));
        assert_eq!(parsed.chart_fill_paint_authored, Some(true));
    }

    #[test]
    fn plot_area_omitted_fill_uses_the_host_default_but_no_fill_remains_explicit() {
        let chart_xml = |plot_shape_properties: &str| {
            format!(
                r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:chart><c:plotArea>
                <c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/>
                  <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
                </c:ser></c:barChart>{plot_shape_properties}
              </c:plotArea></c:chart>
            </c:chartSpace>"#
            )
        };
        let omitted = chart_xml("");
        let parsed = parse_chart_part(
            chart_space_of(&omitted).root_element(),
            &ChartParseContext {
                color_resolver: Some(&WhiteChartFixtureResolver),
                ..Default::default()
            },
        )
        .expect("plot-area chart parses");
        assert_eq!(parsed.plot_area_bg.as_deref(), Some("FFFFFF"));
        assert_eq!(parsed.plot_area_fill_automatic, Some(true));

        let no_fill = chart_xml("<c:spPr><a:noFill/></c:spPr>");
        let parsed = parse_chart_part(
            chart_space_of(&no_fill).root_element(),
            &ChartParseContext {
                color_resolver: Some(&WhiteChartFixtureResolver),
                ..Default::default()
            },
        )
        .expect("noFill plot-area chart parses");
        assert_eq!(parsed.plot_area_bg, None);
        assert_eq!(parsed.plot_area_fill_hidden, Some(true));
        assert_eq!(parsed.plot_area_fill_automatic, None);
    }

    /// (c) Doughnut chart with per-point `<c:dPt>` colors, `showPercent`, and
    /// `holeSize`/`firstSliceAng`. Doughnut (not pie) is used because
    /// `extract_hole_size` only ever matches a `<c:doughnutChart>` — a pie
    /// fixture would leave `hole_size` permanently `None`.
    #[test]
    fn parse_chart_part_doughnut_dpt_colors_and_geometry() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:chart><c:plotArea>
                <c:doughnutChart>
                  <c:holeSize val="45"/>
                  <c:firstSliceAng val="90"/>
                  <c:ser>
                    <c:idx val="0"/>
                    <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Share</c:v></c:pt></c:strCache></c:strRef></c:tx>
                    <c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:srgbClr val="ff0000"/></a:solidFill></c:spPr></c:dPt>
                    <c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="00ff00"/></a:solidFill></c:spPr></c:dPt>
                    <c:cat><c:strCache><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strCache></c:cat>
                    <c:val><c:numCache><c:pt idx="0"><c:v>60</c:v></c:pt><c:pt idx="1"><c:v>40</c:v></c:pt></c:numCache></c:val>
                    <c:dLbls><c:showPercent val="1"/></c:dLbls>
                  </c:ser>
                </c:doughnutChart>
              </c:plotArea></c:chart>
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
        .expect("doughnut chart parses");

        assert_eq!(m.chart_type, "doughnut");
        assert_eq!(m.hole_size, Some(45));
        assert_eq!(m.first_slice_angle, Some(90));
        assert!(!m.show_data_labels);
        assert_eq!(
            m.series[0]
                .series_data_labels
                .as_ref()
                .map(|labels| labels.show_percent),
            Some(true)
        );

        let colors = m.series[0]
            .data_point_colors
            .as_ref()
            .expect("dPt colors populated");
        assert_eq!(colors[0].as_deref(), Some("FF0000"));
        assert_eq!(colors[1].as_deref(), Some("00FF00"));
    }

    /// DrawingML `spPr` is not a solid-color-only grammar. Preserve the
    /// authored recipe/noFill provenance in the common model even though the
    /// desktop-Excel text-box extent is currently rendered only for the
    /// separately verified solid subset.
    #[test]
    fn parse_chart_data_table_preserves_non_solid_fill_provenance() {
        let parse_table = |shape: &str| {
            let xml = format!(
                r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
                  <c:chart><c:plotArea>
                    <c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/>
                      <c:cat><c:strCache><c:pt idx="0"><c:v>A</c:v></c:pt></c:strCache></c:cat>
                      <c:val><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:val>
                    </c:ser></c:barChart>
                    <c:dTable><c:spPr>{shape}</c:spPr></c:dTable>
                  </c:plotArea></c:chart>
                </c:chartSpace>"#
            );
            parse_chart_part(
                chart_space_of(&xml).root_element(),
                &ChartParseContext {
                    color_resolver: Some(&FixtureResolver),
                    ..Default::default()
                },
            )
            .expect("data-table chart parses")
            .data_table
            .expect("c:dTable preserved")
        };

        let gradient = parse_table(
            r#"<a:gradFill><a:gsLst>
              <a:gs pos="0"><a:srgbClr val="112233"/></a:gs>
              <a:gs pos="100000"><a:srgbClr val="AABBCC"/></a:gs>
            </a:gsLst><a:lin ang="5400000"/></a:gradFill>"#,
        );
        assert!(matches!(
            gradient.fill,
            Some(ChartStyleFill::Gradient { .. })
        ));
        assert_eq!(gradient.fill_hidden, None);
        assert_eq!(gradient.fill_paint_authored, Some(true));

        let no_fill = parse_table("<a:noFill/>");
        assert_eq!(no_fill.fill, None);
        assert_eq!(no_fill.fill_color, None);
        assert_eq!(no_fill.fill_hidden, Some(true));
        assert_eq!(no_fill.fill_paint_authored, Some(true));
    }

    /// `c:invertIfNegative` and the Office 2010 alternate fill extension are
    /// series formatting, not XLSX viewer policy. Preserve them in the shared
    /// chart model so every DrawingML chart host paints negative bars alike.
    #[test]
    fn parse_chart_part_negative_bar_alternate_fill() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"
                 xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart">
              <c:chart><c:plotArea><c:barChart>
                <c:barDir val="col"/><c:grouping val="clustered"/>
                <c:ser><c:idx val="0"/><c:tx><c:v>Profit</c:v></c:tx>
                  <c:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></c:spPr>
                  <c:invertIfNegative/>
                  <c:val><c:numCache><c:pt idx="0"><c:v>-3</c:v></c:pt></c:numCache></c:val>
                  <c:extLst><c:ext uri="invert"><c14:invertSolidFillFmt>
                    <c14:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></c14:spPr>
                  </c14:invertSolidFillFmt></c:ext></c:extLst>
                </c:ser>
              </c:barChart></c:plotArea></c:chart>
            </c:chartSpace>"#
        );
        let model = parse_chart_part(
            chart_space_of(&xml).root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("negative-fill chart parses");
        let series = &model.series[0];
        assert_eq!(series.invert_if_negative, Some(true));
        assert_eq!(
            series.inverted_fill,
            Some(ChartStyleFill::Solid {
                color: "FFFFFF".to_string(),
            })
        );
        assert_eq!(series.inverted_fill_hidden, None);
        assert_eq!(series.inverted_fill_authored, Some(true));
        assert_eq!(series.inverted_line_color, None);
        assert_eq!(series.inverted_line_width_emu, None);
        assert_eq!(series.inverted_line_hidden, None);
        assert_eq!(series.inverted_line_authored, Some(false));
    }

    /// §21.2.2.198: a scatter series whose `<c:spPr><a:ln>` is `<a:noFill/>`
    /// has its connecting line turned OFF, overriding the group-level
    /// `<c:scatterStyle val="lineMarker">` (§21.2.2.42). The parser must set
    /// `line_hidden = Some(true)` so the renderer draws markers only. A series with a paintable line leaves
    /// `line_hidden = None`.
    #[test]
    fn parse_chart_part_scatter_series_line_nofill_sets_line_hidden() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:chart><c:plotArea>
                <c:scatterChart>
                  <c:scatterStyle val="lineMarker"/>
                  <c:ser>
                    <c:idx val="0"/>
                    <c:spPr><a:ln w="25400"><a:noFill/></a:ln></c:spPr>
                    <c:xVal><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:xVal>
                    <c:yVal><c:numCache><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numCache></c:yVal>
                  </c:ser>
                  <c:ser>
                    <c:idx val="1"/>
                    <c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill></a:ln></c:spPr>
                    <c:xVal><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:xVal>
                    <c:yVal><c:numCache><c:pt idx="0"><c:v>5</c:v></c:pt><c:pt idx="1"><c:v>8</c:v></c:pt></c:numCache></c:yVal>
                  </c:ser>
                  <c:axId val="1"/><c:axId val="2"/>
                </c:scatterChart>
                <c:valAx><c:axId val="1"/><c:axPos val="b"/><c:scaling><c:logBase val="2"/><c:orientation val="maxMin"/></c:scaling><c:majorUnit val="0.5"/><c:minorUnit val="0.1"/><c:crossAx val="2"/></c:valAx>
                <c:valAx><c:axId val="2"/><c:axPos val="l"/><c:scaling><c:logBase val="10"/><c:orientation val="maxMin"/></c:scaling><c:crossAx val="1"/></c:valAx>
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
        .expect("scatter chart parses");
        assert_eq!(m.chart_type, "scatter");
        assert_eq!(m.scatter_style.as_deref(), Some("lineMarker"));
        assert_eq!(m.cat_axis_major_unit, Some(0.5));
        assert_eq!(m.cat_axis_minor_unit, Some(0.1));
        assert_eq!(m.cat_axis_log_base, Some(2.0));
        assert_eq!(m.cat_axis_orientation.as_deref(), Some("maxMin"));
        assert_eq!(m.val_axis_log_base, Some(10.0));
        assert_eq!(m.val_axis_orientation.as_deref(), Some("maxMin"));
        // Series 0: explicit `<a:noFill/>` line → line_hidden set.
        assert_eq!(
            m.series[0].line_hidden,
            Some(true),
            "a `<a:ln><a:noFill/>` series line must record line_hidden"
        );
        // Series 1: a paintable line → line_hidden stays None (group style governs).
        assert_eq!(
            m.series[1].line_hidden, None,
            "a series with a solid line must NOT set line_hidden"
        );
        assert_eq!(m.series[0].marker_symbol, None);
        assert_eq!(
            m.series[0].automatic_marker_symbol.as_deref(),
            Some("diamond")
        );
        assert_eq!(
            m.series[1].automatic_marker_symbol.as_deref(),
            Some("square")
        );
    }

    #[test]
    fn parse_marker_block_symbol_size_fill_line() {
        let xml = format!(
            r#"<c:marker xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:symbol val="circle"/>
              <c:size val="6"/>
              <c:spPr>
                <a:solidFill><a:srgbClr val="ff0000"/></a:solidFill>
                <a:ln w="25400"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:ln>
              </c:spPr>
            </c:marker>"#
        );
        let d = root_of(&xml);
        let (symbol, size, fill, fill_paint, fill_authored, line, line_width_emu, line_authored) =
            parse_marker_block(Some(d.root_element()), &FixtureResolver);
        assert_eq!(symbol.as_deref(), Some("circle"));
        assert_eq!(size, Some(6.0));
        assert_eq!(fill.as_deref(), Some("FF0000"));
        assert_eq!(fill_paint, None);
        assert_eq!(fill_authored, Some(true));
        assert_eq!(line.as_deref(), Some("4472C4"));
        assert_eq!(line_width_emu, Some(25400));
        assert_eq!(line_authored, Some(true));
    }

    #[test]
    fn parse_marker_block_preserves_explicit_no_fill() {
        let xml = format!(
            r#"<c:marker xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:symbol val="circle"/>
              <c:spPr><a:noFill/><a:ln><a:solidFill><a:srgbClr val="777777"/></a:solidFill></a:ln></c:spPr>
            </c:marker>"#
        );
        let d = root_of(&xml);
        let (_, _, fill, fill_paint, fill_authored, line, line_width_emu, line_authored) =
            parse_marker_block(Some(d.root_element()), &FixtureResolver);
        assert_eq!(fill.as_deref(), Some("00000000"));
        assert_eq!(fill_paint, None);
        assert_eq!(fill_authored, Some(true));
        assert_eq!(line.as_deref(), Some("777777"));
        assert_eq!(line_width_emu, None);
        assert_eq!(line_authored, Some(true));

        let line_no_fill = format!(
            r#"<c:marker xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:symbol val="circle"/><c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr></c:marker>"#
        );
        let d = root_of(&line_no_fill);
        let (_, _, fill, fill_paint, fill_authored, line, _, line_authored) =
            parse_marker_block(Some(d.root_element()), &FixtureResolver);
        assert_eq!(fill.as_deref(), Some("FFFFFF"));
        assert_eq!(fill_paint, None);
        assert_eq!(fill_authored, Some(true));
        assert_eq!(line.as_deref(), Some("00000000"));
        assert_eq!(line_authored, Some(true));
    }

    #[test]
    fn parse_marker_block_preserves_structured_pattern_fill() {
        let xml = format!(
            r#"<c:marker xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:symbol val="diamond"/>
              <c:spPr><a:pattFill prst="pct30">
                <a:fgClr><a:srgbClr val="112233"/></a:fgClr>
                <a:bgClr><a:srgbClr val="DDEEFF"/></a:bgClr>
              </a:pattFill></c:spPr>
            </c:marker>"#
        );
        let d = root_of(&xml);
        let (_, _, fill, fill_paint, fill_authored, _, _, _) =
            parse_marker_block(Some(d.root_element()), &FixtureResolver);
        assert_eq!(fill, None);
        assert_eq!(fill_authored, Some(true));
        assert_eq!(
            fill_paint,
            Some(ChartStyleFill::Pattern {
                fg: "112233".to_string(),
                bg: "DDEEFF".to_string(),
                preset: "pct30".to_string(),
            })
        );
    }

    #[test]
    fn parse_marker_block_keeps_unresolved_picture_fill_authored() {
        let xml = format!(
            r#"<c:marker xmlns:c="{C_NS}" xmlns:a="{A_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <c:symbol val="picture"/>
              <c:spPr><a:blipFill><a:blip r:embed="rId1"/><a:stretch/></a:blipFill></c:spPr>
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
    fn parse_marker_block_retains_resolved_picture_fill_geometry() {
        struct Images;
        impl ChartImageResolver for Images {
            fn resolve_image(
                &self,
                source: ChartImageSource,
                relationship_id: &str,
            ) -> Option<(String, String)> {
                if source != ChartImageSource::Chart {
                    return None;
                }
                match relationship_id {
                    "rId1" => Some(("xl/media/marker.png".to_owned(), "image/png".to_owned())),
                    "rIdSvg" => {
                        Some(("xl/media/marker.svg".to_owned(), "image/svg+xml".to_owned()))
                    }
                    _ => None,
                }
            }
        }
        let xml = format!(
            r#"<c:marker xmlns:c="{C_NS}" xmlns:a="{A_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main">
              <c:symbol val="picture"/><c:size val="9"/>
              <c:spPr><a:blipFill dpi="192" rotWithShape="0"><a:blip r:embed="rId1"><a:alphaModFix amt="50000"/><a:alphaModFix amt="50000"/><a:extLst><a:ext uri="{{96DAC541-7B7A-43D3-8B79-37D633B846F1}}"><asvg:svgBlip r:embed="rIdSvg"/></a:ext></a:extLst></a:blip>
                <a:srcRect l="10000" t="20000" r="30000" b="40000"/>
                <a:stretch><a:fillRect l="-5000"/></a:stretch>
              </a:blipFill></c:spPr>
            </c:marker>"#
        );
        let d = root_of(&xml);
        let (_, _, fill, fill_paint, fill_authored, _, _, _) =
            parse_marker_block_with_images(Some(d.root_element()), &FixtureResolver, &Images);
        assert_eq!(fill, None);
        assert_eq!(fill_authored, Some(true));
        assert_eq!(
            fill_paint,
            Some(ChartStyleFill::Image {
                image_path: "xl/media/marker.png".to_owned(),
                mime_type: "image/png".to_owned(),
                svg_image_path: Some("xl/media/marker.svg".to_owned()),
                dpi: Some(192),
                rot_with_shape: Some(false),
                src_rect: Some(crate::blip::SrcRect {
                    l: 0.1,
                    t: 0.2,
                    r: 0.3,
                    b: 0.4,
                }),
                fill_rect: Some(crate::fill::FillRect {
                    l: -0.05,
                    ..Default::default()
                }),
                stretch: true,
                tile: None,
                alpha: Some(0.25),
                duotone: None,
            })
        );
    }

    #[test]
    fn parse_marker_block_does_not_invent_an_omitted_picture_fill_mode() {
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
        for fill_mode in ["", "<a:stretch/><a:tile/>"] {
            let xml = format!(
                r#"<c:marker xmlns:c="{C_NS}" xmlns:a="{A_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <c:symbol val="picture"/><c:spPr><a:blipFill><a:blip r:embed="rId1"/>{fill_mode}</a:blipFill></c:spPr>
                </c:marker>"#
            );
            let document = root_of(&xml);
            let (_, _, _, paint, authored, _, _, _) = parse_marker_block_with_images(
                Some(document.root_element()),
                &FixtureResolver,
                &Images,
            );
            assert_eq!(paint, None);
            assert_eq!(authored, Some(true));
        }
    }

    #[test]
    fn parse_marker_block_fails_closed_for_multiple_duotone_effects() {
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
              <c:symbol val="picture"/><c:spPr><a:blipFill><a:blip r:embed="rId1">
                <a:duotone><a:srgbClr val="000000"/><a:srgbClr val="FFFFFF"/></a:duotone>
                <a:duotone><a:srgbClr val="112233"/><a:srgbClr val="DDEEFF"/></a:duotone>
              </a:blip><a:stretch/></a:blipFill></c:spPr>
            </c:marker>"#
        );
        let document = root_of(&xml);
        let (_, _, _, paint, authored, _, _, _) = parse_marker_block_with_images(
            Some(document.root_element()),
            &FixtureResolver,
            &Images,
        );
        assert_eq!(paint, None);
        assert_eq!(authored, Some(true));
    }

    #[test]
    fn parse_marker_block_fails_closed_for_malformed_supported_effects() {
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
        for effect in [
            "<a:alphaModFix amt=\"not-a-number\"/>",
            "<a:duotone><a:srgbClr val=\"000000\"/></a:duotone>",
        ] {
            let xml = format!(
                r#"<c:marker xmlns:c="{C_NS}" xmlns:a="{A_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <c:symbol val="picture"/><c:spPr><a:blipFill><a:blip r:embed="rId1">{effect}</a:blip><a:stretch/></a:blipFill></c:spPr>
                </c:marker>"#
            );
            let document = root_of(&xml);
            let (_, _, _, paint, authored, _, _, _) = parse_marker_block_with_images(
                Some(document.root_element()),
                &FixtureResolver,
                &Images,
            );
            assert_eq!(paint, None, "effect must fail closed: {effect}");
            assert_eq!(authored, Some(true));
        }
    }

    #[test]
    fn parse_marker_block_charges_picture_fill_to_the_aggregate_budget() {
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
                <a:stretch/></a:blipFill></c:spPr>
            </c:marker>"#
        );
        let document = root_of(&xml);
        let mut budget = 0;
        let mut exceeded = false;
        let (_, _, _, paint, authored, _, _, _) = parse_marker_block_with_budget(
            Some(document.root_element()),
            &FixtureResolver,
            &Images,
            &mut budget,
            &mut exceeded,
        );
        assert!(exceeded);
        assert_eq!(paint, None);
        assert_eq!(authored, Some(true));
    }

    #[test]
    fn point_marker_paint_budget_reports_aggregate_overflow() {
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:dPt><c:idx val="0"/><c:marker><c:symbol val="circle"/><c:spPr>
                <a:pattFill prst="pct20"><a:fgClr><a:srgbClr val="112233"/></a:fgClr>
                  <a:bgClr><a:srgbClr val="DDEEFF"/></a:bgClr></a:pattFill>
              </c:spPr></c:marker></c:dPt>
              <c:dPt><c:idx val="1"/><c:marker><c:symbol val="circle"/><c:spPr>
                <a:pattFill prst="pct30"><a:fgClr><a:srgbClr val="445566"/></a:fgClr>
                  <a:bgClr><a:srgbClr val="AABBCC"/></a:bgClr></a:pattFill>
              </c:spPr></c:marker></c:dPt>
            </c:ser>"#
        );
        let document = root_of(&xml);
        let mut component_budget = 1;
        let mut paint_budget_exceeded = false;
        let _points = parse_data_point_overrides_with_budget(
            document.root_element(),
            &FixtureResolver,
            &EmptyChartImageResolver,
            false,
            &mut component_budget,
            &mut paint_budget_exceeded,
        );
        assert!(paint_budget_exceeded);
        assert_eq!(component_budget, 0);
    }

    #[test]
    fn parse_chart_part_atomically_rejects_marker_paint_budget_overflow() {
        let stops = (0..=MAX_CHART_MARKER_GRADIENT_STOPS)
            .map(|index| format!(r#"<a:gs pos="{}"><a:srgbClr val="112233"/></a:gs>"#, index))
            .collect::<String>();
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:chart><c:plotArea>
              <c:lineChart><c:ser><c:idx val="0"/><c:marker><c:symbol val="circle"/>
                <c:spPr><a:gradFill><a:gsLst>{stops}</a:gsLst></a:gradFill></c:spPr>
              </c:marker><c:cat><c:strCache><c:pt idx="0"><c:v>A</c:v></c:pt></c:strCache></c:cat>
              <c:val><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:val>
              </c:ser></c:lineChart>
            </c:plotArea></c:chart></c:chartSpace>"#
        );
        let document = chart_space_of(&xml);
        assert!(parse_chart_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            }
        )
        .is_none());
    }

    #[test]
    fn parse_series_pattern_fill_preserves_preset_and_colors() {
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:spPr>
              <a:pattFill prst="pct30">
                <a:fgClr><a:schemeClr val="accent3"/></a:fgClr>
                <a:bgClr><a:schemeClr val="bg1"/></a:bgClr>
              </a:pattFill>
            </c:spPr></c:ser>"#
        );
        let d = root_of(&xml);
        let fill = parse_series_pattern_fill(d.root_element(), &FixtureResolver)
            .expect("pattern fill parses");
        assert_eq!(fill.fill_type, "pattern");
        assert_eq!(fill.preset, "pct30");
        assert_eq!(fill.fg, "A5A5A5");
        assert_eq!(fill.bg, "FFFFFF");
    }

    #[test]
    fn error_bar_and_leader_line_preserve_no_fill_and_dash() {
        let error_xml = format!(
            r#"<c:ser xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:errBars>
              <c:errBarType val="both"/><c:errValType val="fixedVal"/><c:val val="1"/>
              <c:spPr><a:ln><a:noFill/><a:prstDash val="dashDot"/></a:ln></c:spPr>
            </c:errBars></c:ser>"#,
        );
        let error_doc = root_of(&error_xml);
        let bars = parse_error_bars(error_doc.root_element(), &[Some(1.0)], &FixtureResolver);
        assert_eq!(bars[0].hidden, Some(true));
        assert_eq!(bars[0].dash.as_deref(), Some("dashDot"));

        let cache = std::collections::HashMap::new();
        let label_xml = format!(
            r#"<c:ser xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:dLbls>
              <c:showLeaderLines val="1"/><c:leaderLines><c:spPr>
                <a:ln><a:noFill/><a:prstDash val="sysDot"/></a:ln>
              </c:spPr></c:leaderLines>
            </c:dLbls></c:ser>"#,
        );
        let label_doc = root_of(&label_xml);
        let (defaults, _) =
            parse_series_data_labels(label_doc.root_element(), &FixtureResolver, &cache);
        let defaults = defaults.expect("leader-line defaults");
        assert_eq!(defaults.leader_line_hidden, Some(true));
        assert_eq!(defaults.leader_line_dash.as_deref(), Some("sysDot"));
    }

    #[test]
    fn extract_radar_style_present_and_absent() {
        let xml = format!(
            r#"<c:radarChart xmlns:c="{C_NS}"><c:radarStyle val="marker"/></c:radarChart>"#
        );
        let d = root_of(&xml);
        assert_eq!(
            extract_radar_style(d.root_element()).as_deref(),
            Some("marker")
        );

        let none_xml = format!(r#"<c:barChart xmlns:c="{C_NS}"></c:barChart>"#);
        let d2 = root_of(&none_xml);
        assert!(extract_radar_style(d2.root_element()).is_none());
    }

    #[test]
    fn chart_style_fill_preflight_limits_palette_expansion_before_gradient_parse() {
        let document = roxmltree::Document::parse(
            r#"<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:gradFill><a:gsLst>
                <a:gs pos="0"><a:srgbClr val="000000"/></a:gs>
                <a:gs pos="50000"><a:srgbClr val="808080"/></a:gs>
                <a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs>
              </a:gsLst></a:gradFill>
            </a:spPr>"#,
        )
        .expect("gradient style XML");
        assert_eq!(
            chart_style_paint_component_count(document.root_element()),
            Some(3)
        );
        assert_eq!(chart_style_paint_entry_limit(Some(3), 4, 6), 2);
        assert_eq!(chart_style_paint_entry_limit(Some(7), 4, 6), 0);
        assert_eq!(chart_style_paint_entry_limit(Some(0), 4, 0), 4);
        assert_eq!(chart_style_paint_entry_limit(None, 4, 0), 4);
    }

    #[test]
    fn classic_chart_preserves_linked_chart_style_role_table() {
        let chart_xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}"><c:style val="2"/><c:chart><c:plotArea>
              <c:lineChart><c:ser><c:idx val="0"/><c:order val="0"/>
                <c:cat><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
              </c:ser></c:lineChart>
            </c:plotArea></c:chart></c:chartSpace>"#,
        );
        let style_xml = format!(
            r#"<cs:chartStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}">
              <cs:dataPointMarkerLayout symbol="diamond" size="9"/>
              <cs:chartArea><cs:spPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:ln cmpd="dbl"><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="334455"/></a:gs><a:gs pos="100000"><a:srgbClr val="DDEEFF"/></a:gs></a:gsLst><a:lin ang="2700000"/></a:gradFill><a:custDash><a:ds d="125000" sp="75000"/></a:custDash></a:ln></cs:spPr></cs:chartArea>
              <cs:plotArea><cs:spPr><a:ln cmpd="thickThin"><a:solidFill><a:srgbClr val="556677"/></a:solidFill><a:custDash/></a:ln></cs:spPr></cs:plotArea>
              <cs:legend><cs:spPr><a:ln><a:solidFill><a:srgbClr val="667788"/></a:solidFill><a:custDash><a:ds d="200000" sp="50000"/></a:custDash></a:ln></cs:spPr></cs:legend>
              <cs:categoryAxis><cs:defRPr sz="700" b="1" i="1"><a:solidFill><a:srgbClr val="778899"/></a:solidFill><a:latin typeface="Axis Face"/></cs:defRPr></cs:categoryAxis>
              <cs:dataLabelCallout><cs:defRPr sz="850" b="1" i="1" lang="ja-JP" baseline="25000"><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:latin typeface="Callout Face"/></cs:defRPr><cs:bodyPr rot="-2700000" wrap="none" anchor="b" vert="vert" lIns="12700" tIns="25400" rIns="38100" bIns="50800"/></cs:dataLabelCallout>
              <cs:trendlineLabel><cs:fontRef idx="minor"><a:srgbClr val="595959"/></cs:fontRef><cs:defRPr lang="en-US" baseline="-12.5%"><a:blipFill/></cs:defRPr><cs:bodyPr rot="1800000" wrap="square" anchor="ctr"/></cs:trendlineLabel>
              <cs:dropLine><cs:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:ln></cs:spPr></cs:dropLine>
              <cs:seriesLine><cs:lnRef idx="1"><cs:styleClr val="auto"/></cs:lnRef></cs:seriesLine>
              <cs:gridlineMinor><cs:spPr><a:ln><a:noFill/></a:ln></cs:spPr></cs:gridlineMinor>
            </cs:chartStyle>"#,
        );
        let colors_xml = format!(
            r#"<cs:colorStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}" meth="cycle">
              <a:srgbClr val="AA0000"/><a:srgbClr val="00AA00"/>
            </cs:colorStyle>"#,
        );
        let theme = format!(
            r#"<a:theme xmlns:a="{A_NS}"><a:themeElements>
              <a:fmtScheme name="geometry-only line recipe">
                <a:fillStyleLst/>
                <a:lnStyleLst><a:ln w="25400"><a:prstDash val="dash"/></a:ln></a:lnStyleLst>
                <a:effectStyleLst/><a:bgFillStyleLst/>
              </a:fmtScheme>
            </a:themeElements></a:theme>"#,
        );
        let resolver = FormatSchemeFixtureResolver {
            format_scheme: crate::theme::ThemeFormatScheme::parse(&theme),
        };
        let chart_doc = root_of(&chart_xml);
        let model = parse_chart_part(
            chart_doc.root_element(),
            &ChartParseContext {
                color_resolver: Some(&resolver),
                style_xml: Some(&style_xml),
                color_style_xml: Some(&colors_xml),
                ..Default::default()
            },
        )
        .expect("classic chart parses");

        assert_eq!(model.chart_style_color_method.as_deref(), Some("cycle"));
        assert_eq!(model.chart_style_marker_size_pt, Some(9));
        assert_eq!(model.chart_style_marker_symbol.as_deref(), Some("diamond"));
        assert_eq!(
            model.chart_style_color_palette.as_ref().map(Vec::len),
            Some(2)
        );
        let roles = model.chart_style_roles.expect("linked role table");
        assert_eq!(roles.len(), 9);
        let classic_roles = model
            .classic_chart_style_roles
            .expect("numeric built-in role table");
        assert_eq!(classic_roles.len(), 30);
        assert!(classic_roles.contains_key("title"));
        assert!(!roles.contains_key("title"));
        assert_eq!(classic_roles["title"].font_size_hpt, Some(1_200));
        assert_eq!(
            roles["chartArea"].fill_colors.as_deref(),
            Some(&[Some("112233".to_string()), Some("112233".to_string())][..]),
        );
        assert_eq!(roles["dropLine"].line_width_emu, Some(12_700));
        assert_eq!(roles["dropLine"].line_paint_authored, Some(true));
        assert_eq!(roles["seriesLine"].line_width_emu, Some(25_400));
        assert_eq!(roles["seriesLine"].line_dash.as_deref(), Some("dash"));
        assert_eq!(roles["seriesLine"].line_paint_authored, None);
        assert_eq!(
            roles["chartArea"].line_custom_dash.as_deref(),
            Some(
                &[ChartLineDashSegment {
                    dash: 1.25,
                    space: 0.75,
                }][..]
            ),
        );
        assert_eq!(roles["chartArea"].line_compound.as_deref(), Some("dbl"));
        assert!(matches!(
            roles["chartArea"]
                .line_paints
                .as_ref()
                .and_then(|paints| paints.first())
                .and_then(Option::as_ref),
            Some(ChartStyleFill::Gradient { angle, .. }) if (*angle - 45.0).abs() < 1e-9
        ));
        assert_eq!(roles["plotArea"].line_custom_dash.as_deref(), Some(&[][..]));
        assert_eq!(
            roles["plotArea"].line_compound.as_deref(),
            Some("thickThin")
        );
        assert_eq!(
            roles["legend"].line_custom_dash.as_deref(),
            Some(
                &[ChartLineDashSegment {
                    dash: 2.0,
                    space: 0.5,
                }][..]
            ),
        );
        assert_eq!(
            roles["dropLine"]
                .line_colors
                .as_ref()
                .and_then(|colors| colors[0].as_deref()),
            Some("445566"),
        );
        assert_eq!(roles["gridlineMinor"].line_hidden, Some(true));
        assert_eq!(roles["categoryAxis"].font_size_hpt, Some(700));
        assert_eq!(roles["categoryAxis"].font_bold, Some(true));
        assert_eq!(roles["categoryAxis"].font_italic, Some(true));
        assert_eq!(roles["categoryAxis"].font_color.as_deref(), Some("778899"));
        assert_eq!(
            roles["categoryAxis"].font_face.as_deref(),
            Some("Axis Face")
        );
        let callout = &roles["dataLabelCallout"];
        assert_eq!(callout.font_language.as_deref(), Some("ja-JP"));
        assert_eq!(callout.font_baseline, Some(0.25));
        assert_eq!(callout.text_rotation, Some(-2_700_000));
        assert_eq!(callout.text_wrap.as_deref(), Some("none"));
        assert_eq!(callout.text_vertical_anchor.as_deref(), Some("b"));
        assert_eq!(callout.text_vertical_mode.as_deref(), Some("vert"));
        assert_eq!(callout.text_l_ins_emu, Some(12_700));
        assert_eq!(callout.text_t_ins_emu, Some(25_400));
        assert_eq!(callout.text_r_ins_emu, Some(38_100));
        assert_eq!(callout.text_b_ins_emu, Some(50_800));
        let trendline_label = &roles["trendlineLabel"];
        assert_eq!(trendline_label.font_language.as_deref(), Some("en-US"));
        assert_eq!(trendline_label.font_baseline, Some(-0.125));
        assert_eq!(trendline_label.font_color, None);
        assert_eq!(trendline_label.font_paint_authored, Some(true));
        assert_eq!(trendline_label.text_rotation, Some(1_800_000));
    }

    #[test]
    fn authored_unreadable_chart_style_fails_closed_for_every_role() {
        let chart_xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}"><c:style val="3"/><c:chart><c:plotArea>
              <c:barChart><c:barDir val="col"/><c:ser><c:idx val="7"/><c:order val="0"/>
                <c:cat><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
              </c:ser></c:barChart>
            </c:plotArea></c:chart></c:chartSpace>"#,
        );
        let document = chart_space_of(&chart_xml);
        for invalid_style in ["\0", "<cs:chartStyle"] {
            let model = parse_chart_part(
                document.root_element(),
                &ChartParseContext {
                    color_resolver: Some(&FixtureResolver),
                    style_xml: Some(invalid_style),
                    color_style_xml: None,
                    ..Default::default()
                },
            )
            .expect("the chart itself remains renderable");
            let roles = model
                .chart_style_roles
                .expect("authored unresolved role table");
            assert_eq!(roles.len(), CHART_STYLE_ROLE_NAMES.len());
            for role in CHART_STYLE_ROLE_NAMES {
                let style = &roles[role];
                assert_eq!(style.fill_paint_authored, Some(true));
                assert_eq!(style.fill_hidden, Some(true));
                assert_eq!(style.line_paint_authored, Some(true));
                assert_eq!(style.line_hidden, Some(true));
                assert_eq!(style.font_paint_authored, Some(true));
                assert_eq!(style.font_hidden, Some(true));
                assert_eq!(style.effect_authored, Some(true));
                assert_eq!(style.effect_unsupported, Some(true));
            }
        }
    }

    #[test]
    fn identical_varying_groups_share_one_numeric_point_palette() {
        let groups = (0..MAX_CHART_PLOT_GROUPS)
            .map(|idx| {
                format!(
                    r#"<c:barChart><c:barDir val="col"/><c:varyColors val="1"/><c:ser>
                  <c:idx val="{idx}"/><c:order val="{idx}"/><c:val><c:numLit>
                  <c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt>
                  </c:numLit></c:val></c:ser></c:barChart>"#,
                )
            })
            .collect::<String>();
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:style val="2"/>
              <c:chart><c:plotArea>{groups}</c:plotArea></c:chart></c:chartSpace>"#,
        );
        let document = chart_space_of(&xml);
        let model = parse_chart_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("bounded repeated varying groups parse");

        assert_eq!(
            model.plot_groups.as_ref().map(Vec::len),
            Some(MAX_CHART_PLOT_GROUPS)
        );
        assert!(model.classic_varying_point_chart_style_roles.is_some());
        assert!(model
            .classic_varying_point_chart_style_roles_by_group
            .is_none());
    }

    #[test]
    fn classic_multi_series_doughnut_replays_the_point_palette_for_each_ring() {
        let series = |idx: u8, values: &[u8]| {
            let points = values
                .iter()
                .enumerate()
                .map(|(point_idx, value)| {
                    format!(r#"<c:pt idx="{point_idx}"><c:v>{value}</c:v></c:pt>"#)
                })
                .collect::<String>();
            format!(
                r#"<c:ser><c:idx val="{idx}"/><c:order val="{idx}"/>
                  <c:cat><c:strLit><c:ptCount val="3"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:ptCount val="3"/>{points}</c:numLit></c:val>
                </c:ser>"#,
            )
        };
        let group = format!(
            r#"<c:doughnutChart><c:varyColors val="1"/>{}{}</c:doughnutChart>"#,
            series(8, &[1, 2, 3]),
            series(9, &[3, 2, 1]),
        );
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:style val="3"/>
              <c:chart><c:plotArea>{group}</c:plotArea></c:chart></c:chartSpace>"#,
        );
        let document = chart_space_of(&xml);
        let model = parse_chart_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("multi-series doughnut parses");

        assert_eq!(model.vary_colors, Some(true));
        let roles = model
            .classic_varying_point_chart_style_roles
            .expect("numeric varying-point doughnut roles");
        for role in ["dataPoint", "dataPoint3D"] {
            assert_eq!(
                roles[role].fill_formatting_indices.as_deref(),
                Some(&[0, 1, 2][..]),
                "{role} follows the point index within every ring",
            );
        }
        let series_roles = model
            .classic_chart_style_roles
            .expect("numeric series-owned doughnut roles");
        for role in ["dataPointMarker", "dataPointLine", "dataPointWireframe"] {
            assert_eq!(
                series_roles[role].line_formatting_indices.as_deref(),
                Some(&[8, 9][..]),
                "{role} remains series-owned",
            );
        }

        let group_off = group.replacen(
            r#"<c:varyColors val="1"/>"#,
            r#"<c:varyColors val="0"/>"#,
            1,
        );
        let xml_off = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:style val="3"/>
              <c:chart><c:plotArea>{group_off}</c:plotArea></c:chart></c:chartSpace>"#,
        );
        let document_off = chart_space_of(&xml_off);
        let model_off = parse_chart_part(
            document_off.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("non-varying multi-series doughnut parses");
        assert_eq!(model_off.vary_colors, Some(false));
        let roles_off = model_off
            .classic_chart_style_roles
            .expect("numeric non-varying doughnut roles");
        assert_eq!(
            roles_off["dataPoint"].fill_formatting_indices.as_deref(),
            Some(&[8, 9][..]),
            "an explicitly non-varying doughnut is coloured by series/ring",
        );
    }

    #[test]
    fn linked_marker_style_preserves_unsupported_fill_provenance() {
        let chart_xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea>
              <c:lineChart><c:ser><c:idx val="0"/><c:order val="0"/>
                <c:cat><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
              </c:ser></c:lineChart>
            </c:plotArea></c:chart></c:chartSpace>"#,
        );
        let unsupported_style_xml = format!(
            r#"<cs:chartStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}">
              <cs:dataPointMarker>
                <cs:fillRef idx="1"><cs:styleClr val="auto"/></cs:fillRef>
                <cs:spPr><a:blipFill/></cs:spPr>
              </cs:dataPointMarker>
            </cs:chartStyle>"#,
        );
        let inherited_style_xml = format!(
            r#"<cs:chartStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}">
              <cs:dataPointMarker>
                <cs:fillRef idx="1"><cs:styleClr val="auto"/></cs:fillRef>
              </cs:dataPointMarker>
            </cs:chartStyle>"#,
        );
        let theme_xml = format!(
            r#"<a:theme xmlns:a="{A_NS}"><a:themeElements>
              <a:fmtScheme name="Office">
                <a:fillStyleLst>
                  <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
                </a:fillStyleLst>
                <a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/>
              </a:fmtScheme>
            </a:themeElements></a:theme>"#,
        );
        let resolver = FormatSchemeFixtureResolver {
            format_scheme: crate::theme::ThemeFormatScheme::parse(&theme_xml),
        };
        let document = root_of(&chart_xml);
        let model = parse_chart_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&resolver),
                style_xml: Some(&unsupported_style_xml),
                color_style_xml: None,
                ..Default::default()
            },
        )
        .expect("classic chart parses");
        let role = &model.chart_style_roles.expect("linked roles")["dataPointMarker"];
        assert_eq!(role.fill_paint_authored, Some(true));
        assert_eq!(role.fill_hidden, None);
        assert_eq!(role.fill_colors, None);
        assert_eq!(role.fill_paints, None);

        let inherited_model = parse_chart_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&resolver),
                style_xml: Some(&inherited_style_xml),
                color_style_xml: None,
                ..Default::default()
            },
        )
        .expect("classic chart with inherited marker fill parses");
        let inherited_role =
            &inherited_model.chart_style_roles.expect("linked roles")["dataPointMarker"];
        assert_eq!(inherited_role.fill_paint_authored, Some(true));
        assert!(inherited_role
            .fill_colors
            .as_ref()
            .is_some_and(|colors| colors.iter().any(Option::is_some)));
    }

    #[test]
    fn linked_chart_effects_resolve_palette_and_preserve_precedence_provenance() {
        let chart_xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea>
              <c:lineChart><c:ser><c:idx val="0"/><c:order val="0"/>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
              </c:ser></c:lineChart>
            </c:plotArea></c:chart></c:chartSpace>"#,
        );
        let style_xml = format!(
            r#"<cs:chartStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}">
              <cs:dataPoint><cs:effectRef idx="1"><cs:styleClr val="1"/></cs:effectRef></cs:dataPoint>
              <cs:legend><cs:effectRef idx="1"><cs:styleClr val="auto"/></cs:effectRef><cs:spPr><a:effectLst/></cs:spPr></cs:legend>
              <cs:title><cs:effectRef idx="1"><cs:styleClr val="auto"/></cs:effectRef><cs:spPr><a:effectDag/></cs:spPr></cs:title>
              <cs:plotArea><cs:effectRef idx="0"/></cs:plotArea>
              <cs:chartArea><cs:effectRef idx="99"/></cs:chartArea>
            </cs:chartStyle>"#,
        );
        let colors_xml = format!(
            r#"<cs:colorStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}" meth="cycle">
              <a:srgbClr val="AA0000"/><a:srgbClr val="00AA00"/>
            </cs:colorStyle>"#,
        );
        let theme_xml = format!(
            r#"<a:theme xmlns:a="{A_NS}"><a:themeElements><a:fmtScheme name="effects">
              <a:fillStyleLst/><a:lnStyleLst/>
              <a:effectStyleLst><a:effectStyle><a:effectLst>
                <a:outerShdw blurRad="12700"><a:schemeClr val="phClr"><a:alpha val="50000"/></a:schemeClr></a:outerShdw>
              </a:effectLst></a:effectStyle></a:effectStyleLst>
              <a:bgFillStyleLst/>
            </a:fmtScheme></a:themeElements></a:theme>"#,
        );
        let resolver = FormatSchemeFixtureResolver {
            format_scheme: crate::theme::ThemeFormatScheme::parse(&theme_xml),
        };
        let document = root_of(&chart_xml);
        let model = parse_chart_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&resolver),
                style_xml: Some(&style_xml),
                color_style_xml: Some(&colors_xml),
                ..Default::default()
            },
        )
        .expect("classic chart parses");
        let roles = model.chart_style_roles.expect("linked roles");

        let shadows = roles["dataPoint"].shadows.as_ref().expect("shadows");
        assert_eq!(shadows.len(), 2);
        // A numeric ST_StyleColorVal is a fixed zero-based Chart Colors index,
        // not a relative object index. Every expanded effect entry therefore
        // resolves through the selected green placeholder; the retained index
        // lets compact downstream palettes preserve the same semantics.
        assert_eq!(shadows[0].as_ref().unwrap().color, "00AA00");
        assert_eq!(shadows[1].as_ref().unwrap().color, "00AA00");
        assert!((shadows[0].as_ref().unwrap().alpha - 0.5).abs() < 0.01);
        assert_eq!(roles["dataPoint"].effect_authored, Some(true));
        assert_eq!(roles["dataPoint"].effect_unsupported, None);
        assert_eq!(roles["dataPoint"].effect_color_index, Some(1));

        let direct_empty = &roles["legend"];
        assert_eq!(direct_empty.effect_authored, Some(true));
        assert_eq!(direct_empty.shadows, None);
        assert_eq!(direct_empty.effect_unsupported, None);

        assert_eq!(roles["title"].effect_authored, Some(true));
        assert_eq!(roles["title"].effect_unsupported, Some(true));
        assert_eq!(roles["plotArea"].effect_no_style, Some(true));
        assert_eq!(roles["plotArea"].effect_authored, None);
        assert_eq!(roles["chartArea"].effect_authored, Some(true));
        assert_eq!(roles["chartArea"].effect_unsupported, Some(true));
    }

    #[test]
    fn direct_point_marker_and_up_down_bar_effects_use_generic_style_carriers() {
        let series_xml = format!(
            r#"<c:ser xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:dPt><c:idx val="2"/><c:spPr><a:effectLst>
                <a:outerShdw blurRad="12700"><a:srgbClr val="112233"/></a:outerShdw>
              </a:effectLst></c:spPr><c:marker><c:spPr><a:effectLst>
                <a:glow rad="25400"><a:srgbClr val="445566"/></a:glow>
              </a:effectLst></c:spPr></c:marker></c:dPt>
            </c:ser>"#,
        );
        let document = root_of(&series_xml);
        let points = parse_data_point_overrides(document.root_element(), &FixtureResolver);
        let point = &points[0];
        assert_eq!(
            point
                .chartex_style
                .as_ref()
                .and_then(|style| style.shadows.as_ref())
                .and_then(|effects| effects[0].as_ref())
                .map(|shadow| shadow.color.as_str()),
            Some("112233"),
        );
        assert_eq!(
            point
                .marker_style
                .as_ref()
                .and_then(|style| style.glows.as_ref())
                .and_then(|effects| effects[0].as_ref())
                .map(|glow| (glow.color.as_str(), glow.radius)),
            Some(("445566", 25_400)),
        );

        let bars_xml = format!(
            r#"<c:upDownBars xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:upBars><c:spPr><a:effectLst/></c:spPr></c:upBars>
              <c:downBars><c:spPr><a:effectDag/></c:spPr></c:downBars>
            </c:upDownBars>"#,
        );
        let bars_document = root_of(&bars_xml);
        let bars = parse_chart_up_down_bar_style(bars_document.root_element(), &FixtureResolver);
        assert_eq!(bars.up.style.as_ref().unwrap().effect_authored, Some(true));
        assert_eq!(bars.up.style.as_ref().unwrap().effect_unsupported, None);
        assert_eq!(
            bars.down.style.as_ref().unwrap().effect_unsupported,
            Some(true)
        );
    }

    #[test]
    fn linked_marker_style_retains_picture_relationship_from_style_part() {
        struct Images;
        impl ChartImageResolver for Images {
            fn resolve_image(
                &self,
                source: ChartImageSource,
                relationship_id: &str,
            ) -> Option<(String, String)> {
                (source == ChartImageSource::Style && relationship_id == "rIdPic").then(|| {
                    (
                        "word/media/style-marker.png".to_owned(),
                        "image/png".to_owned(),
                    )
                })
            }
        }
        let chart_xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea><c:lineChart>
              <c:ser><c:idx val="0"/><c:order val="0"/>
                <c:cat><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
              </c:ser></c:lineChart></c:plotArea></c:chart></c:chartSpace>"#
        );
        let style_xml = format!(
            r#"<cs:chartStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <cs:dataPointMarker><cs:spPr><a:blipFill><a:blip r:embed="rIdPic"/>
                <a:tile tx="12700" ty="25400" sx="50000" sy="50000" flip="xy" algn="ctr"/>
              </a:blipFill></cs:spPr></cs:dataPointMarker>
            </cs:chartStyle>"#
        );
        let document = root_of(&chart_xml);
        let model = parse_chart_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: Some(&style_xml),
                color_style_xml: None,
                images: Some(&Images),
                ..Default::default()
            },
        )
        .expect("classic chart parses");
        let role = &model.chart_style_roles.expect("linked roles")["dataPointMarker"];
        assert_eq!(role.fill_paint_authored, Some(true));
        assert!(matches!(
            role.fill_paints.as_ref().and_then(|paints| paints[0].as_ref()),
            Some(ChartStyleFill::Image { image_path, tile: Some(tile), .. })
                if image_path == "word/media/style-marker.png"
                    && tile.tx == Some(12_700)
                    && tile.ty == Some(25_400)
                    && tile.sx.is_some_and(|value| (value - 0.5).abs() < 1e-9)
                    && tile.flip.as_deref() == Some("xy")
                    && tile.algn.as_deref() == Some("ctr")
        ));
    }

    #[test]
    fn chart_image_relationships_accept_only_normative_image_types() {
        let transitional = crate::ns::relationships::TRANSITIONAL;
        let strict = crate::ns::relationships::STRICT;
        let rels = format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rTransitional" Type="{transitional}/image" Target="../media/a.png"/>
              <Relationship Id="rStrict" Type="{strict}/image" Target="../media/b.svg"/>
              <Relationship Id="rVendor" Type="urn:vendor/image" Target="../media/c.png"/>
              <Relationship Id="rRemote" Type="https://attacker.invalid/image" Target="../media/d.png"/>
              <Relationship Id="rExternal" Type="{transitional}/image" Target="https://example.invalid/e.png" TargetMode="External"/>
              <Relationship Id="rMissing" Target="../media/f.png"/>
            </Relationships>"#
        );
        let mut images = ChartImageRelationships::default();
        images.insert_part_relationships(ChartImageSource::Chart, "xl/charts/chart1.xml", &rels);
        assert_eq!(
            images.resolve_image(ChartImageSource::Chart, "rTransitional"),
            Some(("xl/media/a.png".to_owned(), "image/png".to_owned())),
        );
        assert_eq!(
            images.resolve_image(ChartImageSource::Chart, "rStrict"),
            Some(("xl/media/b.svg".to_owned(), "image/svg+xml".to_owned())),
        );
        for id in ["rVendor", "rRemote", "rExternal", "rMissing"] {
            assert_eq!(images.resolve_image(ChartImageSource::Chart, id), None);
        }
    }

    #[test]
    fn linked_chart_style_role_table_fails_closed_before_oversized_expansion() {
        let roles = CHART_STYLE_ROLE_NAMES
            .iter()
            .map(|name| format!("<cs:{name}><cs:spPr><a:solidFill><a:srgbClr val=\"112233\"/></a:solidFill></cs:spPr></cs:{name}>"))
            .collect::<String>();
        let style_xml = format!(
            r#"<cs:chartStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}">{roles}</cs:chartStyle>"#,
        );
        let colors = (0..300)
            .map(|index| format!("<a:srgbClr val=\"{:06X}\"/>", index))
            .collect::<String>();
        let colors_xml = format!(
            r#"<cs:colorStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}" meth="cycle">{colors}</cs:colorStyle>"#,
        );
        let style_doc = root_of(&style_xml);
        let (_, palette) = parse_chart_color_style(&colors_xml, &FixtureResolver)
            .expect("bounded color style parses");
        assert!(parse_chart_style_role_table(
            style_doc.root_element(),
            &FixtureResolver,
            Some(&palette),
            Some("cycle"),
            &EmptyChartImageResolver,
        )
        .is_none());
    }

    #[test]
    fn linked_chart_style_role_table_counts_referenced_theme_gradients() {
        let roles = CHART_STYLE_ROLE_NAMES
            .iter()
            .map(|name| {
                format!(
                    "<cs:{name}><cs:fillRef idx=\"1\"><cs:styleClr val=\"auto\"/></cs:fillRef></cs:{name}>"
                )
            })
            .collect::<String>();
        let style_xml = format!(
            r#"<cs:chartStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}">{roles}</cs:chartStyle>"#,
        );
        let colors = (0..273)
            .map(|index| format!("<a:srgbClr val=\"{:06X}\"/>", index))
            .collect::<String>();
        let colors_xml = format!(
            r#"<cs:colorStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}" meth="cycle">{colors}</cs:colorStyle>"#,
        );
        let stops = (0..129)
            .map(|index| {
                format!(
                    "<a:gs pos=\"{}\"><a:schemeClr val=\"phClr\"/></a:gs>",
                    index * 100_000 / 128
                )
            })
            .collect::<String>();
        let theme_xml = format!(
            r#"<a:theme xmlns:a="{A_NS}"><a:themeElements><a:fmtScheme name="bounded"><a:fillStyleLst><a:gradFill><a:gsLst>{stops}</a:gsLst><a:lin ang="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme></a:themeElements></a:theme>"#,
        );
        let resolver = FormatSchemeFixtureResolver {
            format_scheme: crate::theme::ThemeFormatScheme::parse(&theme_xml),
        };
        let style_doc = root_of(&style_xml);
        let (_, palette) =
            parse_chart_color_style(&colors_xml, &resolver).expect("color style parses");

        assert!(parse_chart_style_role_table(
            style_doc.root_element(),
            &resolver,
            Some(&palette),
            Some("cycle"),
            &EmptyChartImageResolver,
        )
        .is_none());

        // A local fill choice has higher precedence than fillRef even when the
        // local grammar is not representable. It therefore contributes no
        // inherited gradient work to the aggregate preflight.
        let locally_overridden_roles = CHART_STYLE_ROLE_NAMES
            .iter()
            .map(|name| {
                format!(
                    "<cs:{name}><cs:fillRef idx=\"1\"><cs:styleClr val=\"auto\"/></cs:fillRef><cs:spPr><a:blipFill/></cs:spPr></cs:{name}>"
                )
            })
            .collect::<String>();
        let locally_overridden_style_xml = format!(
            r#"<cs:chartStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}">{locally_overridden_roles}</cs:chartStyle>"#,
        );
        let locally_overridden_style_doc = root_of(&locally_overridden_style_xml);
        let table = parse_chart_style_role_table(
            locally_overridden_style_doc.root_element(),
            &resolver,
            Some(&palette),
            Some("cycle"),
            &EmptyChartImageResolver,
        )
        .expect("unsupported local fills suppress inherited gradient work");
        assert!(table
            .values()
            .all(|style| style.fill_paint_authored == Some(true)
                && style.fill_colors.is_none()
                && style.fill_paints.is_none()));
    }

    #[test]
    fn linked_chart_style_role_table_bounds_each_source_paint_recipe() {
        let gradient = |count: usize| {
            (0..count)
                .map(|index| format!(r#"<a:gs pos="{index}"><a:srgbClr val="112233"/></a:gs>"#))
                .collect::<String>()
        };
        let style = |count: usize, outline: bool| {
            let paint = format!(
                r#"<a:gradFill><a:gsLst>{}</a:gsLst><a:lin ang="0"/></a:gradFill>"#,
                gradient(count),
            );
            let sp_pr = if outline {
                format!(r#"<a:noFill/><a:ln>{paint}</a:ln>"#)
            } else {
                paint
            };
            format!(
                r#"<cs:chartStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}"><cs:dataPoint3D><cs:spPr>{sp_pr}</cs:spPr></cs:dataPoint3D></cs:chartStyle>"#,
            )
        };

        for outline in [false, true] {
            let exact_xml = style(MAX_CHART_PAINT_RECIPE_COMPONENTS, outline);
            let exact = root_of(&exact_xml);
            assert!(parse_chart_style_role_table(
                exact.root_element(),
                &FixtureResolver,
                None,
                None,
                &EmptyChartImageResolver,
            )
            .is_some());

            let oversized_xml = style(MAX_CHART_PAINT_RECIPE_COMPONENTS + 1, outline);
            let oversized = root_of(&oversized_xml);
            assert!(parse_chart_style_role_table(
                oversized.root_element(),
                &FixtureResolver,
                None,
                None,
                &EmptyChartImageResolver,
            )
            .is_none());
        }
    }

    #[test]
    fn parse_chart_part_retains_structured_point_paint_for_classic_3d_series() {
        let series = CH13_SER.replace(
            "<c:idx val=\"0\"/>",
            r#"<c:idx val="0"/><c:dPt><c:idx val="1"/><c:spPr><a:gradFill><a:gsLst>
              <a:gs pos="0"><a:srgbClr val="112233"/></a:gs>
              <a:gs pos="100000"><a:srgbClr val="DDEEFF"/></a:gs>
            </a:gsLst><a:lin ang="0"/></a:gradFill><a:ln w="12700"><a:pattFill prst="pct10">
              <a:fgClr><a:srgbClr val="FF0000"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr>
            </a:pattFill></a:ln></c:spPr></c:dPt>"#,
        );
        let group = format!(r#"<c:pie3DChart><c:varyColors val="1"/>{series}</c:pie3DChart>"#);
        let xml = chart_space_with_group(&group);
        let document = chart_space_of(&xml);
        let model = parse_chart_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("pie3D");
        let style = model.series[0]
            .data_point_overrides
            .as_ref()
            .expect("point override")[0]
            .chartex_style
            .as_ref()
            .expect("point style");
        assert!(matches!(
            style.fill_paints.as_deref().and_then(|paints| paints.first()).and_then(Option::as_ref),
            Some(ChartStyleFill::Gradient { stops, .. }) if stops.len() == 2
        ));
        assert!(matches!(
            style
                .line_paints
                .as_deref()
                .and_then(|paints| paints.first())
                .and_then(Option::as_ref),
            Some(ChartStyleFill::Pattern { .. })
        ));
    }

    #[test]
    fn parse_chart_part_rejects_oversized_three_d_surface_paint_before_expansion() {
        let stops = (0..=MAX_CHART_MARKER_GRADIENT_STOPS)
            .map(|index| format!(r#"<a:gs pos="{}"><a:srgbClr val="112233"/></a:gs>"#, index))
            .collect::<String>();
        let group = format!(r#"<c:bar3DChart><c:barDir val="col"/>{CH13_SER}</c:bar3DChart>"#);
        let xml = chart_space_with_group(&group).replace(
            "<c:chart><c:plotArea>",
            &format!(
                r#"<c:chart><c:floor><c:spPr><a:gradFill><a:gsLst>{stops}</a:gsLst><a:lin ang="0"/></a:gradFill></c:spPr></c:floor><c:plotArea>"#,
            ),
        );
        let document = chart_space_of(&xml);
        assert!(parse_chart_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            }
        )
        .is_none());

        let surface = format!(
            r#"<c:surface3DChart>{CH13_SER}<c:bandFmts><c:bandFmt><c:idx val="0"/>
              <c:spPr><a:gradFill><a:gsLst>{stops}</a:gsLst><a:lin ang="0"/></a:gradFill></c:spPr>
            </c:bandFmt></c:bandFmts></c:surface3DChart>"#,
        );
        let surface_xml = chart_space_with_group(&surface);
        let surface_document = chart_space_of(&surface_xml);
        assert!(parse_chart_part(
            surface_document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            }
        )
        .is_none());
    }

    #[test]
    fn parse_chart_part_bounds_direct_three_d_series_fill_and_line_recipes() {
        let gradient = |count: usize| {
            (0..count)
                .map(|index| format!(r#"<a:gs pos="{index}"><a:srgbClr val="112233"/></a:gs>"#))
                .collect::<String>()
        };
        let group = |count: usize, outline: bool| {
            let paint = format!(
                r#"<a:gradFill><a:gsLst>{}</a:gsLst><a:lin ang="0"/></a:gradFill>"#,
                gradient(count),
            );
            let sp_pr = if outline {
                format!(r#"<c:spPr><a:noFill/><a:ln>{paint}</a:ln></c:spPr>"#)
            } else {
                format!(r#"<c:spPr>{paint}</c:spPr>"#)
            };
            let series = CH13_SER.replace(
                r#"<c:idx val="0"/>"#,
                &format!(r#"<c:idx val="0"/>{sp_pr}"#),
            );
            format!(r#"<c:bar3DChart><c:barDir val="col"/>{series}</c:bar3DChart>"#)
        };

        for outline in [false, true] {
            let exact_xml =
                chart_space_with_group(&group(MAX_CHART_PAINT_RECIPE_COMPONENTS, outline));
            let exact_document = chart_space_of(&exact_xml);
            assert!(parse_chart_part(
                exact_document.root_element(),
                &ChartParseContext {
                    color_resolver: Some(&FixtureResolver),
                    ..Default::default()
                }
            )
            .is_some());

            let oversized_xml =
                chart_space_with_group(&group(MAX_CHART_PAINT_RECIPE_COMPONENTS + 1, outline));
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
    }

    #[test]
    fn parse_chart_part_stock_hi_low_no_fill_stays_hidden() {
        let group = r#"<c:stockChart>
          <c:ser><c:idx val="0"/><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>5</c:v></c:pt></c:numLit></c:val></c:ser>
          <c:hiLowLines><c:spPr><a:ln w="12700"><a:noFill/></a:ln></c:spPr></c:hiLowLines>
        </c:stockChart>"#;
        let xml = chart_space_with_group(group);
        let document = chart_space_of(&xml);
        let model = parse_chart_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("stock");
        let style = model.stock_hi_low_line_style.expect("high-low style");
        assert_eq!(style.hidden, Some(true));
        assert_eq!(style.width_emu, Some(12_700));
    }

    #[test]
    fn stock_automatic_paint_is_theme_aware_and_limited_to_observed_styles() {
        let group = format!(
            r#"<c:stockChart>{CH13_SER}<c:hiLowLines/><c:upDownBars><c:upBars/><c:downBars/></c:upDownBars></c:stockChart>"#,
        );
        let parse = |style: Option<u8>, resolver: &dyn ColorResolver| {
            let xml = chart_space_with_group(&group);
            let xml = match style {
                Some(style) => xml.replacen(
                    "<c:chart>",
                    &format!(r#"<c:style val="{style}"/><c:chart>"#),
                    1,
                ),
                None => xml,
            };
            let document = chart_space_of(&xml);
            parse_chart_part(
                document.root_element(),
                &ChartParseContext {
                    color_resolver: Some(resolver),
                    ..Default::default()
                },
            )
            .expect("stock")
        };

        let default_style = parse(None, &FixtureResolver)
            .stock_automatic_style
            .expect("omitted style uses the observed default stock recipe");
        assert_eq!(default_style.line_color, "000000");
        assert_eq!(default_style.line_width_emu, 12_700);
        assert_eq!(default_style.up_fill_color, "F9F9F9");
        assert_eq!(default_style.down_fill_color, "3F3F3F");

        let explicit_style_two = parse(Some(2), &FixtureResolver)
            .stock_automatic_style
            .expect("explicit style 2 uses the observed default stock recipe");
        assert_eq!(explicit_style_two, default_style);

        let style_one = parse(Some(1), &FixtureResolver)
            .stock_automatic_style
            .expect("style 1 recipe");
        assert_eq!(style_one.up_fill_color, "E1E1E1");
        assert_eq!(style_one.down_fill_color, "6C6C6C");

        let themed = parse(Some(10), &DarkRedFixtureResolver)
            .stock_automatic_style
            .expect("style 10 recipe");
        assert_eq!(themed.line_color, "240000");
        assert_eq!(themed.up_fill_color, "F9F9F9");
        assert_eq!(themed.down_fill_color, "493F3F");

        assert!(parse(Some(48), &FixtureResolver)
            .stock_automatic_style
            .is_none());
    }

    #[test]
    fn parse_chart_part_stock_up_down_bars_preserve_structured_fills() {
        let group = r#"<c:stockChart>
          <c:ser><c:idx val="0"/><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>5</c:v></c:pt></c:numLit></c:val></c:ser>
          <c:upDownBars>
            <c:upBars><c:spPr><a:pattFill prst="diagCross"><a:fgClr><a:srgbClr val="112233"/></a:fgClr><a:bgClr><a:srgbClr val="DDEEFF"/></a:bgClr></a:pattFill></c:spPr></c:upBars>
            <c:downBars><c:spPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="000000"/></a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill></c:spPr></c:downBars>
          </c:upDownBars>
        </c:stockChart>"#;
        let xml = chart_space_with_group(group);
        let document = chart_space_of(&xml);
        let model = parse_chart_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("stock");
        let style = model.stock_up_down_bar_style.expect("up/down style");
        assert!(matches!(
            style.up.fill,
            Some(ChartStyleFill::Pattern { ref fg, ref bg, ref preset })
                if fg == "112233" && bg == "DDEEFF" && preset == "diagCross"
        ));
        assert!(matches!(
            style.down.fill,
            Some(ChartStyleFill::Gradient { ref stops, angle, .. })
                if stops.len() == 2 && angle == 90.0
        ));
    }

    /// `data_point_colors` retains only direct `<c:dPt>` fill while automatic
    /// varyColors paint remains in the effective classic dataPoint style.
    #[test]
    fn parse_chart_part_pie_vary_colors_fills_accents() {
        let ser = r#"<c:ser><c:idx val="0"/>
            <c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill></c:spPr></c:dPt>
            <c:cat><c:strCache>
              <c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt>
            </c:strCache></c:cat>
            <c:val><c:numCache>
              <c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt>
            </c:numCache></c:val></c:ser>"#;
        // No <c:varyColors> element → defaults to ON for the pie family.
        let group = format!(r#"<c:pieChart>{ser}</c:pieChart>"#);
        let xml = chart_space_with_group(&group);
        let d = chart_space_of(&xml);
        let m = parse_chart_part(
            d.root_element(),
            &ChartParseContext {
                color_resolver: Some(&AccentResolver),
                ..Default::default()
            },
        )
        .expect("pie parses");
        let colors = m.series[0]
            .data_point_colors
            .as_ref()
            .expect("direct dPt palette");
        assert_eq!(colors[0].as_deref(), Some("112233"));
        assert_eq!(colors[1], None);
        assert_eq!(colors[2], None);

        // varyColors="0" disables the per-slice accent fill: only the explicit
        // dPt color remains, the rest fall back to None (renderer palette).
        let group_off = format!(r#"<c:pieChart><c:varyColors val="0"/>{ser}</c:pieChart>"#);
        let xml_off = chart_space_with_group(&group_off);
        let d2 = chart_space_of(&xml_off);
        let m2 = parse_chart_part(
            d2.root_element(),
            &ChartParseContext {
                color_resolver: Some(&AccentResolver),
                ..Default::default()
            },
        )
        .expect("pie parses");
        let colors2 = m2.series[0].data_point_colors.as_ref().unwrap();
        assert_eq!(colors2[0].as_deref(), Some("112233"));
        assert_eq!(colors2[1], None);
        assert_eq!(colors2[2], None);
    }

    /// A SINGLE-series bar chart with `<c:varyColors>` ABSENT keeps its
    /// per-series paint. §21.2.2.227 defaults only the `val` attribute when the
    /// element is present; it does not make an absent element effective.
    #[test]
    fn parse_chart_part_bar_single_series_without_vary_colors_keeps_series_color() {
        let group = format!(
            r#"<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>{CH13_SER}</c:barChart>"#
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
        .expect("bar parses");
        assert_eq!(m.vary_colors, None);
        assert!(m.series[0].data_point_colors.is_none());
    }

    /// A SINGLE-series bar chart with an explicit `<c:varyColors val="0"/>`
    /// keeps its one per-series color (Office records the forced-single-color
    /// choice this way). No per-point fill,
    /// no chart-level flag.
    #[test]
    fn parse_chart_part_bar_single_series_vary_off_keeps_series_color() {
        let group = format!(
            r#"<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>{CH13_SER}</c:barChart>"#
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
        .expect("bar parses");
        assert_eq!(m.vary_colors, None);
        assert!(m.series[0].data_point_colors.is_none());
    }

    /// §21.2.2.227 varyColors on a SINGLE-series bar/column chart sets the
    /// chart-level contract while retaining only direct `<c:dPt>` paint in the
    /// per-point field. Automatic palette paint belongs to the classic style.
    #[test]
    fn parse_chart_part_bar_vary_colors_single_series_fills_accents() {
        let ser = r#"<c:ser><c:idx val="0"/>
            <c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill></c:spPr></c:dPt>
            <c:cat><c:strCache>
              <c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt>
              <c:pt idx="2"><c:v>C</c:v></c:pt><c:pt idx="3"><c:v>D</c:v></c:pt>
            </c:strCache></c:cat>
            <c:val><c:numCache>
              <c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt>
              <c:pt idx="2"><c:v>3</c:v></c:pt><c:pt idx="3"><c:v>4</c:v></c:pt>
            </c:numCache></c:val></c:ser>"#;
        let group = format!(
            r#"<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="1"/>{ser}</c:barChart>"#
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
        .expect("bar parses");
        assert_eq!(m.vary_colors, Some(true));
        let colors = m.series[0]
            .data_point_colors
            .as_ref()
            .expect("direct dPt palette");
        assert_eq!(colors[0].as_deref(), Some("112233"));
        assert_eq!(colors[1], None);
        assert_eq!(colors[2], None);
        assert_eq!(colors[3], None);
    }

    /// §21.2.2.227 varyColors on a MULTI-series bar chart is a no-op for the
    /// per-point fill: Office keeps per-series colors when several series share
    /// the axes, so neither the per-point palette nor the chart-level flag is
    /// emitted. Only the single-series case (issue #931) varies by point.
    #[test]
    fn parse_chart_part_bar_vary_colors_multi_series_keeps_series_colors() {
        let ser0 = r#"<c:ser><c:idx val="0"/>
            <c:cat><c:strCache><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strCache></c:cat>
            <c:val><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:val></c:ser>"#;
        let ser1 = r#"<c:ser><c:idx val="1"/>
            <c:cat><c:strCache><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strCache></c:cat>
            <c:val><c:numCache><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numCache></c:val></c:ser>"#;
        let group = format!(
            r#"<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="1"/>{ser0}{ser1}</c:barChart>"#
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
        .expect("bar parses");
        assert_eq!(m.vary_colors, None);
        assert!(m.series[0].data_point_colors.is_none());
        assert!(m.series[1].data_point_colors.is_none());
    }

    #[test]
    fn parse_chart_part_keeps_bounded_bubble_point_structured_fill() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:chart><c:plotArea>
              <c:bubbleChart><c:ser><c:idx val="0"/><c:order val="0"/>
                <c:dPt><c:idx val="0"/><c:spPr><a:gradFill><a:gsLst>
                  <a:gs pos="0"><a:srgbClr val="112233"/></a:gs>
                  <a:gs pos="100000"><a:srgbClr val="DDEEFF"/></a:gs>
                </a:gsLst></a:gradFill></c:spPr></c:dPt>
                <c:xVal><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:xVal>
                <c:yVal><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:yVal>
                <c:bubbleSize><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>3</c:v></c:pt></c:numLit></c:bubbleSize>
              </c:ser></c:bubbleChart>
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
        let point = &chart.series[0].data_point_overrides.as_ref().unwrap()[0];
        let style = point
            .chartex_style
            .as_ref()
            .expect("bubble dPt keeps shape paint");
        assert_eq!(style.fill_paint_authored, Some(true));
        assert!(matches!(
            style.fill_paints.as_deref(),
            Some([Some(ChartStyleFill::Gradient { stops, .. })]) if stops.len() == 2
        ));
    }

    #[test]
    fn bubble_point_shape_paints_share_the_marker_aggregate_budget() {
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:dPt><c:idx val="0"/><c:spPr><a:pattFill prst="pct20">
                <a:fgClr><a:srgbClr val="112233"/></a:fgClr>
                <a:bgClr><a:srgbClr val="DDEEFF"/></a:bgClr>
              </a:pattFill></c:spPr></c:dPt>
              <c:dPt><c:idx val="1"/><c:spPr><a:ln><a:pattFill prst="pct30">
                <a:fgClr><a:srgbClr val="445566"/></a:fgClr>
                <a:bgClr><a:srgbClr val="AABBCC"/></a:bgClr>
              </a:pattFill></a:ln></c:spPr></c:dPt>
            </c:ser>"#
        );
        let document = root_of(&xml);
        let mut component_budget = 1;
        let mut paint_budget_exceeded = false;
        let _ = parse_data_point_overrides_with_budget(
            document.root_element(),
            &FixtureResolver,
            &EmptyChartImageResolver,
            true,
            &mut component_budget,
            &mut paint_budget_exceeded,
        );
        assert!(paint_budget_exceeded);
        assert_eq!(component_budget, 0);
    }

    /// Office varies a lone bubble series by point when `<c:varyColors>` is
    /// omitted, just as it does for a lone bar series. Automatic colors remain
    /// in the effective style while direct dPt fills stay in the point field.
    #[test]
    fn parse_chart_part_bubble_vary_colors_defaults_on_and_honors_false() {
        let ser = r#"<c:ser><c:idx val="0"/><c:order val="0"/>
            <c:xVal><c:numLit><c:ptCount val="3"/>
              <c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt>
            </c:numLit></c:xVal>
            <c:yVal><c:numLit><c:ptCount val="3"/>
              <c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt>
            </c:numLit></c:yVal>
            <c:bubbleSize><c:numLit><c:ptCount val="3"/>
              <c:pt idx="0"><c:v>100</c:v></c:pt><c:pt idx="1"><c:v>400</c:v></c:pt><c:pt idx="2"><c:v>900</c:v></c:pt>
            </c:numLit></c:bubbleSize>
          </c:ser>"#;
        let default_xml = chart_space_with_group(&format!("<c:bubbleChart>{ser}</c:bubbleChart>"));
        let default_doc = chart_space_of(&default_xml);
        let default_chart = parse_chart_part(
            default_doc.root_element(),
            &ChartParseContext {
                color_resolver: Some(&AccentResolver),
                ..Default::default()
            },
        )
        .expect("bubble chart parses");
        assert!(default_chart.series[0].data_point_colors.is_none());

        let off_xml = chart_space_with_group(&format!(
            "<c:bubbleChart><c:varyColors val=\"0\"/>{ser}</c:bubbleChart>"
        ));
        let off_doc = chart_space_of(&off_xml);
        let off_chart = parse_chart_part(
            off_doc.root_element(),
            &ChartParseContext {
                color_resolver: Some(&AccentResolver),
                ..Default::default()
            },
        )
        .expect("bubble chart parses");
        assert!(off_chart.series[0].data_point_colors.is_none());

        let point_formats = r#"<c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr></c:dPt>
            <c:dPt><c:idx val="2"/><c:spPr><a:noFill/></c:spPr></c:dPt>"#;
        let no_fill_ser = ser.replace(
            "<c:xVal>",
            &format!(r#"<c:spPr><a:noFill/></c:spPr>{point_formats}<c:xVal>"#),
        );
        let no_fill_xml = chart_space_with_group(&format!(
            "<c:bubbleChart><c:varyColors/>{no_fill_ser}</c:bubbleChart>"
        ));
        let no_fill_doc = chart_space_of(&no_fill_xml);
        let no_fill_chart = parse_chart_part(
            no_fill_doc.root_element(),
            &ChartParseContext {
                color_resolver: Some(&AccentResolver),
                ..Default::default()
            },
        )
        .expect("noFill bubble parses");
        assert_eq!(no_fill_chart.series[0].color.as_deref(), Some("00000000"));
        assert_eq!(
            no_fill_chart.series[0].data_point_colors,
            Some(vec![
                None,
                Some("FF0000".to_string()),
                Some("00000000".to_string())
            ]),
        );

        let solid_ser = ser.replace(
            "<c:xVal>",
            &format!(r#"<c:spPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></c:spPr>{point_formats}<c:xVal>"#),
        );
        let solid_xml = chart_space_with_group(&format!(
            "<c:bubbleChart><c:varyColors/>{solid_ser}</c:bubbleChart>"
        ));
        let solid_doc = chart_space_of(&solid_xml);
        let solid_chart = parse_chart_part(
            solid_doc.root_element(),
            &ChartParseContext {
                color_resolver: Some(&AccentResolver),
                ..Default::default()
            },
        )
        .expect("solid bubble parses");
        assert_eq!(
            solid_chart.series[0].data_point_colors,
            Some(vec![
                None,
                Some("FF0000".to_string()),
                Some("00000000".to_string()),
            ]),
        );
    }
}
