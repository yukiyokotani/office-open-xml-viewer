#[cfg(test)]
mod tests {
    use super::super::*;

    #[test]
    fn chartex_axis_hidden_value_only() {
        let xml = r#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
            <cx:axis id="0"><cx:catScaling/></cx:axis>
            <cx:axis id="1" hidden="1"><cx:valScaling/></cx:axis>
        </cx:chartSpace>"#;
        let d = root_of(xml);
        assert_eq!(extract_chartex_axis_hidden(d.root_element()), (false, true));
    }

    #[test]
    fn chartex_axis_tick_marks_require_an_authored_type() {
        let xml = r#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
            <cx:axis id="0"><cx:catScaling/><cx:majorTickMarks type="out"/></cx:axis>
            <cx:axis id="1"><cx:valScaling/></cx:axis>
        </cx:chartSpace>"#;
        let d = root_of(xml);
        let cat_axis = d
            .root_element()
            .descendants()
            .find(|node| node.is_element() && child(*node, "catScaling").is_some());
        let val_axis = d
            .root_element()
            .descendants()
            .find(|node| node.is_element() && child(*node, "valScaling").is_some());
        assert_eq!(
            extract_chartex_axis_tick_mark(cat_axis, "majorTickMarks"),
            "out"
        );
        assert_eq!(
            extract_chartex_axis_tick_mark(val_axis, "majorTickMarks"),
            "none"
        );
    }

    /// (a) Waterfall with the full decoration set: a category dimension, a value
    /// dimension with negatives, `<cx:subtotals>` (idx 0 is implicit, idx 5 is
    /// explicit), a series `<cx:spPr>` fill, per-idx `<cx:dataLabel>` colours
    /// (positives → tx1, negatives → accent1), a hidden value axis, and a
    /// `<cx:catScaling gapWidth="0.8">` fraction (→ legacy 80%).
    #[test]
    fn parse_chartex_part_waterfall_full_contract() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chartData>
                <cx:data id="0">
                  <cx:strDim type="cat">
                    <cx:lvl ptCount="4">
                      <cx:pt idx="0">Start</cx:pt>
                      <cx:pt idx="1">Up</cx:pt>
                      <cx:pt idx="2">Down</cx:pt>
                      <cx:pt idx="3">End</cx:pt>
                    </cx:lvl>
                  </cx:strDim>
                  <cx:numDim type="val">
                    <cx:lvl ptCount="4">
                      <cx:pt idx="0">100</cx:pt>
                      <cx:pt idx="1">40</cx:pt>
                      <cx:pt idx="2">-30</cx:pt>
                      <cx:pt idx="3">110</cx:pt>
                    </cx:lvl>
                  </cx:numDim>
                </cx:data>
              </cx:chartData>
              <cx:chart>
                <cx:plotArea>
                  <cx:plotAreaRegion>
                    <cx:series layoutId="waterfall">
                      <cx:spPr><a:solidFill><a:srgbClr val="196eca"/></a:solidFill></cx:spPr>
                      <cx:layoutPr><cx:visibility connectorLines="0"/></cx:layoutPr>
                      <cx:dataLabels pos="outEnd">
                        <cx:dataLabel idx="0">
                          <cx:txPr><a:p><a:pPr><a:defRPr><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:defRPr></a:pPr></a:p></cx:txPr>
                        </cx:dataLabel>
                        <cx:dataLabel idx="2">
                          <cx:visibility categoryName="1" value="0"/>
                          <cx:numFmt formatCode="0.0"/>
                          <cx:separator>|</cx:separator>
                          <cx:txPr><a:p><a:pPr><a:defRPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:defRPr></a:pPr></a:p></cx:txPr>
                        </cx:dataLabel>
                        <cx:dataLabelHidden idx="1"/>
                      </cx:dataLabels>
                      <cx:subtotals>
                        <cx:idx val="0"/>
                        <cx:idx val="3"/>
                      </cx:subtotals>
                    </cx:series>
                  </cx:plotAreaRegion>
                  <cx:axis id="0"><cx:catScaling gapWidth="0.8"/></cx:axis>
                  <cx:axis id="1" hidden="1"><cx:valScaling/>
                    <cx:title><cx:tx><cx:rich><a:bodyPr rot="-1800000"/><a:p><a:r><a:t>Value</a:t></a:r></a:p></cx:rich></cx:tx>
                      <cx:layout><cx:manualLayout><cx:xMode val="edge"/><cx:yMode val="edge"/><cx:x val="0.3"/><cx:y val="0.4"/></cx:manualLayout></cx:layout>
                    </cx:title>
                  </cx:axis>
                </cx:plotArea>
              </cx:chart>
            </cx:chartSpace>"#
        );
        let d = chart_space_of(&xml);
        let m = parse_chartex_part(
            d.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: None,
                ..Default::default()
            },
        )
        .expect("waterfall parses");

        assert_eq!(m.chart_type, "waterfall");
        assert_eq!(m.categories, vec!["Start", "Up", "Down", "End"]);
        assert_eq!(m.series.len(), 1);
        assert_eq!(
            m.series[0].values,
            vec![Some(100.0), Some(40.0), Some(-30.0), Some(110.0)]
        );
        // Series fill resolved through the resolver (srgbClr uppercased).
        assert_eq!(m.series[0].color.as_deref(), Some("196ECA"));
        // Per-idx label colours: idx0/idx3 unset (None), idx0 tx1→000000,
        // idx2 accent1→4472C4. Presence of any override materializes the vec.
        let dl = m.series[0]
            .data_label_colors
            .as_ref()
            .expect("per-label colours present");
        assert_eq!(dl.len(), 4);
        assert_eq!(dl[0].as_deref(), Some("000000"));
        assert_eq!(dl[1], None);
        assert_eq!(dl[2].as_deref(), Some("4472C4"));
        assert_eq!(dl[3], None);
        let overrides = m.series[0]
            .data_label_overrides
            .as_ref()
            .expect("ChartEx point-label overrides present");
        let hidden = overrides
            .iter()
            .find(|override_| override_.idx == 1)
            .unwrap();
        assert_eq!(hidden.deleted, Some(true));
        let visible_parts = overrides
            .iter()
            .find(|override_| override_.idx == 2)
            .unwrap();
        assert_eq!(visible_parts.show_cat_name, Some(true));
        assert_eq!(visible_parts.show_val, Some(false));
        assert_eq!(visible_parts.format_code.as_deref(), Some("0.0"));
        assert_eq!(visible_parts.separator.as_deref(), Some("|"));
        // Both totals were explicitly authored; duplicates are de-duplicated.
        assert_eq!(m.subtotal_indices, vec![0, 3]);
        // gapWidth fraction 0.8 → legacy percent 80.
        assert_eq!(m.bar_gap_width, Some(80));
        // Hidden value axis, visible category axis.
        assert!(!m.cat_axis_hidden);
        assert!(m.val_axis_hidden);
        assert_eq!(m.val_axis_title.as_deref(), Some("Value"));
        assert_eq!(m.val_axis_title_rotation, Some(-1_800_000));
        assert_eq!(
            m.val_axis_title_manual_layout
                .as_ref()
                .map(|layout| layout.x),
            Some(0.3)
        );
        assert_eq!(m.chartex_connector_lines, Some(false));
        // Theme fallback faces threaded from the resolver (NIT-2: not a direct
        // `theme.get("+mj-lt")`).
        assert_eq!(m.theme_major_font_latin.as_deref(), Some("Calibri Light"));
        assert_eq!(m.theme_minor_font_latin.as_deref(), Some("Calibri"));
    }

    #[test]
    fn parse_chartex_axis_title_runs_override_chart_style_property_by_property() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chart><cx:plotArea>
                <cx:plotAreaRegion><cx:series layoutId="waterfall"/></cx:plotAreaRegion>
                <cx:axis id="0"><cx:catScaling/><cx:title><cx:tx><cx:rich>
                  <a:bodyPr/><a:p><a:r><a:rPr i="1"/><a:t>Category</a:t></a:r></a:p>
                </cx:rich></cx:tx></cx:title></cx:axis>
                <cx:axis id="1"><cx:valScaling/><cx:title><cx:tx><cx:rich>
                  <a:bodyPr/><a:p><a:pPr><a:defRPr><a:solidFill><a:srgbClr val="778899"/></a:solidFill></a:defRPr></a:pPr><a:r><a:rPr sz="1200" b="0" i="1"><a:latin typeface="Inline Val"/></a:rPr><a:t>Value</a:t></a:r></a:p>
                </cx:rich></cx:tx></cx:title><cx:txPr><a:p><a:pPr><a:defRPr sz="900" i="1"><a:latin typeface="Inline Tick Val"/></a:defRPr></a:pPr></a:p></cx:txPr></cx:axis>
              </cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let style = format!(
            r#"<cs:chartStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}">
              <cs:axisTitle><cs:defRPr sz="1000" b="1" i="0"><a:solidFill><a:srgbClr val="445566"/></a:solidFill><a:latin typeface="Style Axis"/></cs:defRPr></cs:axisTitle>
              <cs:categoryAxis><cs:defRPr sz="700" b="0" i="0"><a:latin typeface="Tick Cat"/></cs:defRPr></cs:categoryAxis>
              <cs:valueAxis><cs:defRPr sz="800" b="0" i="0"><a:latin typeface="Tick Val"/></cs:defRPr></cs:valueAxis>
            </cs:chartStyle>"#
        );
        let document = chart_space_of(&xml);
        let model = parse_chartex_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: Some(&style),
                ..Default::default()
            },
        )
        .expect("ChartEx axis titles parse");

        assert_eq!(model.cat_axis_title_font_size_hpt, None);
        assert_eq!(model.cat_axis_title_font_bold, Some(false));
        assert_eq!(model.cat_axis_title_font_italic, Some(true));
        assert_eq!(model.cat_axis_title_font_color, None);
        assert_eq!(model.cat_axis_title_text_vertical_inset_emu, Some(91_440));
        assert_eq!(model.cat_axis_title_font_face.as_deref(), None);
        // Inline run and default-run properties win independently over style.
        assert_eq!(model.val_axis_title_font_size_hpt, Some(1200));
        assert_eq!(model.val_axis_title_font_bold, Some(false));
        assert_eq!(model.val_axis_title_font_italic, Some(true));
        assert_eq!(model.val_axis_title_font_color.as_deref(), Some("778899"));
        assert_eq!(model.val_axis_title_text_vertical_inset_emu, Some(91_440));
        assert_eq!(
            model.val_axis_title_font_face.as_deref(),
            Some("Inline Val")
        );
        // The parser retains direct text separately from the linked role so
        // core can apply one direct > linked > numeric cascade.
        assert_eq!(model.cat_axis_font_size_hpt, None);
        assert_eq!(model.cat_axis_font_italic, None);
        assert_eq!(model.val_axis_font_size_hpt, Some(900));
        assert_eq!(model.val_axis_font_italic, Some(true));
        assert_eq!(model.val_axis_font_face.as_deref(), Some("Inline Tick Val"));
        let roles = model
            .chart_style_roles
            .as_ref()
            .expect("linked style roles");
        assert_eq!(roles["axisTitle"].font_size_hpt, Some(1000));
        assert_eq!(roles["categoryAxis"].font_size_hpt, Some(700));
        assert_eq!(roles["valueAxis"].font_size_hpt, Some(800));
    }

    #[test]
    fn parse_chartex_data_point_style_inherits_theme_fill_and_line_refs() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chart><cx:plotArea><cx:plotAreaRegion>
                <cx:series layoutId="waterfall"/>
              </cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let style = format!(
            r#"<cs:chartStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}">
              <cs:dataPoint>
                <cs:fillRef idx="1"><cs:styleClr val="auto"/></cs:fillRef>
                <cs:lnRef idx="1"><cs:styleClr val="auto"/></cs:lnRef>
                <cs:spPr><a:ln w="19050"/></cs:spPr>
              </cs:dataPoint>
            </cs:chartStyle>"#
        );
        let theme = format!(
            r#"<a:theme xmlns:a="{A_NS}"><a:themeElements>
              <a:fmtScheme name="Office">
                <a:fillStyleLst>
                  <a:solidFill><a:schemeClr val="phClr"><a:lumMod val="50000"/></a:schemeClr></a:solidFill>
                </a:fillStyleLst>
                <a:lnStyleLst>
                  <a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"><a:lumMod val="75000"/></a:schemeClr></a:solidFill></a:ln>
                </a:lnStyleLst>
                <a:effectStyleLst/><a:bgFillStyleLst/>
              </a:fmtScheme>
            </a:themeElements></a:theme>"#
        );
        let resolver = FormatSchemeFixtureResolver {
            format_scheme: crate::theme::ThemeFormatScheme::parse(&theme),
        };
        let document = chart_space_of(&xml);
        let model = parse_chartex_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&resolver),
                style_xml: Some(&style),
                ..Default::default()
            },
        )
        .expect("ChartEx style refs parse");

        let accents = model.chartex_accents.expect("raw theme palette");
        let style = model.chartex_data_point_style.expect("dataPoint style");
        let fills = style.fill_colors.expect("fillRef palette");
        let lines = style.line_colors.expect("lnRef palette");
        assert_eq!(fills.len(), 6);
        assert_eq!(lines.len(), 6);
        assert_eq!(accents[0], "5B9BD5");
        assert_ne!(fills[0].as_deref(), Some(accents[0].as_str()));
        assert_ne!(lines[0], fills[0]);
        assert_ne!(lines[0], lines[1]);
        // Local width overlays only that property; the theme recipe still
        // supplies paint through lnRef.
        assert_eq!(style.line_width_emu, Some(19050));
        assert_eq!(style.line_hidden, None);
        assert_eq!(style.fill_hidden, None);
    }

    #[test]
    fn parse_chartex_style_distinguishes_no_style_from_explicit_no_fill() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}"><cx:chart><cx:plotArea><cx:plotAreaRegion>
              <cx:series layoutId="boxWhisker"/>
            </cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>"#
        );
        let no_style = format!(
            r#"<cs:chartStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}">
              <cs:dataPoint>
                <cs:fillRef idx="0"><cs:styleClr val="auto"/></cs:fillRef>
                <cs:lnRef idx="0"><cs:styleClr val="auto"/></cs:lnRef>
              </cs:dataPoint>
            </cs:chartStyle>"#
        );
        let document = chart_space_of(&xml);
        let model = parse_chartex_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: Some(&no_style),
                ..Default::default()
            },
        )
        .expect("NoStyle line parses");
        let role = model.chartex_data_point_style.expect("dataPoint role");
        assert_eq!(role.fill_hidden, Some(true));
        assert_eq!(role.fill_no_style, Some(true));
        assert_eq!(role.line_hidden, Some(true));
        assert_eq!(role.line_no_style, Some(true));

        let explicit_no_fill = no_style.replace(
            "</cs:dataPoint>",
            "<cs:spPr><a:noFill/><a:ln><a:noFill/></a:ln></cs:spPr></cs:dataPoint>",
        );
        let model = parse_chartex_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: Some(&explicit_no_fill),
                ..Default::default()
            },
        )
        .expect("explicit noFill line parses");
        let role = model.chartex_data_point_style.expect("dataPoint role");
        assert_eq!(role.fill_hidden, Some(true));
        assert_eq!(role.fill_no_style, None);
        assert_eq!(role.line_hidden, Some(true));
        assert_eq!(role.line_no_style, None);
    }

    #[test]
    fn parse_chartex_no_style_line_keeps_local_geometry_for_numeric_fallback() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}"><cx:chart><cx:plotArea><cx:plotAreaRegion>
              <cx:series layoutId="boxWhisker"/>
            </cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>"#
        );
        let style = format!(
            r#"<cs:chartStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}">
              <cs:dataPoint>
                <cs:lnRef idx="0"><cs:styleClr val="auto"/></cs:lnRef>
                <cs:spPr><a:ln w="25400" cap="rnd"><a:prstDash val="dash"/><a:bevel/></a:ln></cs:spPr>
              </cs:dataPoint>
            </cs:chartStyle>"#
        );
        let document = chart_space_of(&xml);
        let model = parse_chartex_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: Some(&style),
                ..Default::default()
            },
        )
        .expect("NoStyle line geometry parses");
        let role = model.chartex_data_point_style.expect("dataPoint role");
        assert_eq!(role.line_no_style, Some(true));
        assert_eq!(role.line_hidden, Some(true));
        assert_eq!(role.line_width_emu, Some(25_400));
        assert_eq!(role.line_dash.as_deref(), Some("dash"));
        assert_eq!(role.line_dash_authored, Some(true));
        assert_eq!(role.line_cap.as_deref(), Some("rnd"));
        assert_eq!(role.line_join.as_deref(), Some("bevel"));
    }

    #[test]
    fn parse_chartex_linked_color_style_and_role_specific_paints() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chart><cx:plotArea><cx:plotAreaRegion>
                <cx:series layoutId="boxWhisker"/>
              </cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let style = format!(
            r#"<cs:chartStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}">
              <cs:dataPoint><cs:spPr><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></cs:spPr></cs:dataPoint>
              <cs:dataPointLine><cs:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="dash"/><a:round/></a:ln></cs:spPr></cs:dataPointLine>
              <cs:seriesLine><cs:spPr><a:ln w="38100"><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></a:ln></cs:spPr></cs:seriesLine>
              <cs:dataPointMarker><cs:spPr><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:ln w="9525"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></cs:spPr></cs:dataPointMarker>
            </cs:chartStyle>"#
        );
        let colors = format!(
            r#"<cs:colorStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}" meth="cycle">
              <a:schemeClr val="accent1"/><a:schemeClr val="accent2"/>
              <cs:variation/><cs:variation><a:lumMod val="50000"/></cs:variation>
            </cs:colorStyle>"#
        );
        let document = chart_space_of(&xml);
        let model = parse_chartex_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: Some(&style),
                color_style_xml: Some(&colors),
                ..Default::default()
            },
        )
        .expect("linked chart color/style parts parse");

        assert_eq!(model.chartex_color_style_method.as_deref(), Some("cycle"));
        let palette = model.chartex_color_palette.clone().expect("color palette");
        assert_eq!(
            model.chartex_color_palette.as_deref(),
            Some(
                &[
                    Some("4472C4".to_string()),
                    Some("ED7D31".to_string()),
                    Some("203864".to_string()),
                    Some("843C0B".to_string()),
                ][..]
            )
        );
        let point = model.chartex_data_point_style.expect("point role");
        assert_eq!(point.fill_colors, Some(palette));
        let line = model.chartex_data_point_line_style.expect("line role");
        assert_eq!(line.line_width_emu, Some(28575));
        assert_eq!(line.line_cap.as_deref(), Some("rnd"));
        assert_eq!(line.line_dash.as_deref(), Some("dash"));
        assert_eq!(line.line_join.as_deref(), Some("round"));
        let series_line = model.chartex_series_line_style.expect("seriesLine role");
        assert_eq!(series_line.line_width_emu, Some(38100));
        assert_eq!(
            series_line.line_colors,
            Some(vec![Some("ED7D31".to_string()); 4])
        );
        let marker = model.chartex_data_point_marker_style.expect("marker role");
        assert_eq!(marker.line_width_emu, Some(9525));
        assert_eq!(
            marker.line_colors,
            Some(vec![Some("FFFFFF".to_string()); 4])
        );

        let no_fill_style = style.replace(
            "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></cs:spPr></cs:dataPoint>",
            "<a:noFill/></cs:spPr></cs:dataPoint>",
        );
        let no_fill_model = parse_chartex_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: Some(&no_fill_style),
                color_style_xml: Some(&colors),
                ..Default::default()
            },
        )
        .expect("dataPoint noFill parses");
        assert_eq!(
            no_fill_model
                .chartex_data_point_style
                .expect("point role")
                .fill_hidden,
            Some(true)
        );
    }

    #[test]
    fn parse_chartex_chart_space_shape_components_override_independently() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chart><cx:plotArea><cx:plotAreaRegion>
                <cx:series layoutId="boxWhisker"/>
              </cx:plotAreaRegion></cx:plotArea></cx:chart>
              <cx:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></cx:spPr>
            </cx:chartSpace>"#
        );
        let style = format!(
            r#"<cs:chartStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}">
              <cs:chartArea mods="allowNoFillOverride allowNoLineOverride">
                <cs:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></cs:spPr>
              </cs:chartArea>
            </cs:chartStyle>"#
        );
        let document = chart_space_of(&xml);
        let model = parse_chartex_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: Some(&style),
                color_style_xml: None,
                ..Default::default()
            },
        )
        .expect("ChartEx chart area parses");

        assert_eq!(model.chart_fill_paint_authored, Some(true));
        assert_eq!(model.chart_border_hidden, None);
        assert_eq!(model.chart_border_paint_authored, None);

        // The modifier changes whether an explicit noFill may override the
        // linked component; it does not make an omitted line authoritative.
        let unmodified = style.replace(" mods=\"allowNoFillOverride allowNoLineOverride\"", "");
        let control = parse_chartex_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: Some(&unmodified),
                color_style_xml: None,
                ..Default::default()
            },
        )
        .expect("control ChartEx chart area parses");
        assert_eq!(control.chart_border_hidden, None);
        assert_eq!(control.chart_border_paint_authored, None);

        let line_only_xml = xml.replace(
            "<a:solidFill><a:schemeClr val=\"bg1\"/></a:solidFill>",
            "<a:ln><a:solidFill><a:srgbClr val=\"112233\"/></a:solidFill></a:ln>",
        );
        let fill_style = style.replace(
            "<cs:spPr><a:ln w=\"9525\"><a:solidFill><a:srgbClr val=\"D9D9D9\"/></a:solidFill></a:ln></cs:spPr>",
            "<cs:spPr><a:solidFill><a:srgbClr val=\"FFFFFF\"/></a:solidFill></cs:spPr>",
        );
        let line_only_document = chart_space_of(&line_only_xml);
        let line_only = parse_chartex_part(
            line_only_document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: Some(&fill_style),
                color_style_xml: None,
                ..Default::default()
            },
        )
        .expect("ChartEx chart-area line-only override parses");
        assert_eq!(line_only.chart_fill_hidden, None);
        assert_eq!(line_only.chart_fill_paint_authored, None);
        assert_eq!(line_only.chart_border_color.as_deref(), Some("112233"));
        assert_eq!(line_only.chart_border_paint_authored, Some(true));
    }

    #[test]
    fn parse_chartex_font_ref_style_colors_keep_palette_and_ownership() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}"><cx:chart><cx:plotArea>
              <cx:plotAreaRegion><cx:series layoutId="boxWhisker"/></cx:plotAreaRegion>
            </cx:plotArea></cx:chart></cx:chartSpace>"#,
        );
        let style = format!(
            r#"<cs:chartStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}">
              <cs:dataLabel><cs:fontRef idx="minor"><cs:styleClr val="auto"/></cs:fontRef></cs:dataLabel>
              <cs:title><cs:fontRef idx="major"><cs:styleClr val="1"/></cs:fontRef></cs:title>
              <cs:legend><cs:fontRef idx="minor"><cs:styleClr val="missing-name"/></cs:fontRef></cs:legend>
              <cs:trendlineLabel><cs:fontRef idx="minor"><cs:styleClr val="auto"/></cs:fontRef><cs:defRPr><a:noFill/></cs:defRPr></cs:trendlineLabel>
            </cs:chartStyle>"#,
        );
        let colors = format!(
            r#"<cs:colorStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}" meth="cycle">
              <a:srgbClr val="112233"/><a:srgbClr val="445566"/>
            </cs:colorStyle>"#,
        );
        let document = chart_space_of(&xml);
        let model = parse_chartex_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: Some(&style),
                color_style_xml: Some(&colors),
                ..Default::default()
            },
        )
        .expect("fontRef style colors parse");
        let roles = model.chart_style_roles.expect("linked roles");
        assert_eq!(
            roles["dataLabel"].font_colors.as_deref(),
            Some(&[Some("112233".to_string()), Some("445566".to_string())][..]),
        );
        assert_eq!(roles["dataLabel"].font_color_index, None);
        assert_eq!(roles["title"].font_color_index, Some(1));
        assert_eq!(roles["title"].font_color.as_deref(), Some("445566"));
        assert_eq!(roles["legend"].font_color_index, Some(0));
        assert_eq!(roles["legend"].font_color.as_deref(), Some("112233"));
        assert_eq!(roles["trendlineLabel"].font_colors, None);
        assert_eq!(roles["trendlineLabel"].font_paint_authored, Some(true));
        assert_eq!(roles["trendlineLabel"].font_hidden, Some(true));
    }

    #[test]
    fn parse_chartex_style_retains_shared_gradient_and_pattern_fill_recipes() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chart><cx:plotArea><cx:plotAreaRegion>
                <cx:series layoutId="boxWhisker"/>
              </cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let gradient_style = format!(
            r#"<cs:chartStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}">
              <cs:dataPoint>
                <cs:fillRef idx="1"><cs:styleClr val="auto"/></cs:fillRef>
                <cs:spPr><a:gradFill rotWithShape="0">
                  <a:gsLst>
                    <a:gs pos="0"><a:schemeClr val="phClr"/></a:gs>
                    <a:gs pos="100000"><a:schemeClr val="lt1"/></a:gs>
                  </a:gsLst>
                  <a:lin ang="5400000" scaled="1"/>
                </a:gradFill></cs:spPr>
              </cs:dataPoint>
            </cs:chartStyle>"#
        );
        let document = chart_space_of(&xml);
        let gradient_model = parse_chartex_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: Some(&gradient_style),
                ..Default::default()
            },
        )
        .expect("gradient Chart Style parses");
        let gradient_paints = gradient_model
            .chartex_data_point_style
            .expect("point style")
            .fill_paints
            .expect("gradient paints");
        assert_eq!(gradient_paints.len(), 6);
        assert!(matches!(
            gradient_paints[0].as_ref(),
            Some(ChartStyleFill::Gradient {
                stops,
                angle,
                grad_type,
                scaled: Some(true),
                rot_with_shape: Some(false),
                ..
            }) if stops[0].color == "5B9BD5"
                && stops[1].color == "FFFFFF"
                && (*angle - 90.0).abs() < 1e-9
                && grad_type == "linear"
        ));

        let pattern_style = gradient_style.replace(
            r#"<a:gradFill rotWithShape="0">
                  <a:gsLst>
                    <a:gs pos="0"><a:schemeClr val="phClr"/></a:gs>
                    <a:gs pos="100000"><a:schemeClr val="lt1"/></a:gs>
                  </a:gsLst>
                  <a:lin ang="5400000" scaled="1"/>
                </a:gradFill>"#,
            r#"<a:pattFill prst="diagCross">
                  <a:fgClr><a:schemeClr val="phClr"/></a:fgClr>
                  <a:bgClr><a:schemeClr val="lt1"/></a:bgClr>
                </a:pattFill>"#,
        );
        let pattern_model = parse_chartex_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: Some(&pattern_style),
                ..Default::default()
            },
        )
        .expect("pattern Chart Style parses");
        let pattern_paints = pattern_model
            .chartex_data_point_style
            .expect("point style")
            .fill_paints
            .expect("pattern paints");
        assert!(matches!(
            pattern_paints[0].as_ref(),
            Some(ChartStyleFill::Pattern { fg, bg, preset })
                if fg == "5B9BD5" && bg == "FFFFFF" && preset == "diagCross"
        ));
    }

    #[test]
    fn parse_chartex_style_color_uses_index_semantics_and_transforms() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chart><cx:plotArea><cx:plotAreaRegion>
                <cx:series layoutId="boxWhisker"/>
              </cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let style = format!(
            r#"<cs:chartStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}">
              <cs:dataPoint>
                <cs:fillRef idx="1"><cs:styleClr val="2"><a:lumMod val="50000"/></cs:styleClr></cs:fillRef>
                <cs:spPr><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></cs:spPr>
              </cs:dataPoint>
              <cs:dataPointLine>
                <cs:lnRef idx="1"><cs:styleClr val="named-extension-value"/></cs:lnRef>
                <cs:spPr><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></cs:spPr>
              </cs:dataPointLine>
              <cs:dataPointMarker>
                <cs:fillRef idx="1"><a:srgbClr val="AA5500"/></cs:fillRef>
                <cs:lnRef idx="1" mods="ignoreCSTransforms"><cs:styleClr val="2"><a:lumMod val="10000"/></cs:styleClr></cs:lnRef>
                <cs:spPr>
                  <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
                  <a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
                </cs:spPr>
              </cs:dataPointMarker>
            </cs:chartStyle>"#
        );
        let colors = format!(
            r#"<cs:colorStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}" meth="acrossLinear">
              <a:schemeClr val="accent1"/><a:schemeClr val="missingSlot"/><a:schemeClr val="accent2"/>
            </cs:colorStyle>"#
        );
        let document = chart_space_of(&xml);
        let model = parse_chartex_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: Some(&style),
                color_style_xml: Some(&colors),
                ..Default::default()
            },
        )
        .expect("styleClr values parse");

        assert_eq!(
            model.chartex_color_style_method.as_deref(),
            Some("acrossLinear")
        );
        assert_eq!(
            model.chartex_color_palette.as_deref(),
            Some(&[Some("4472C4".to_string()), None, Some("ED7D31".to_string()),][..]),
        );
        let point = model.chartex_data_point_style.expect("point role");
        assert_eq!(point.fill_color_index, Some(2));
        let fixed_fills = point.fill_colors.expect("fixed transformed fill");
        assert!(fixed_fills.iter().all(|color| color == &fixed_fills[0]));
        assert_ne!(fixed_fills[0].as_deref(), Some("ED7D31"));

        let within_colors = colors.replace("meth=\"acrossLinear\"", "meth=\"withinLinear\"");
        let within = parse_chartex_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: Some(&style),
                color_style_xml: Some(&within_colors),
                ..Default::default()
            },
        )
        .expect("withinLinear fixed index parses");
        let within_fills = within
            .chartex_data_point_style
            .expect("within point role")
            .fill_colors
            .expect("within fills");
        assert!(within_fills.iter().all(|color| color == &within_fills[0]));
        assert_ne!(within_fills[0], fixed_fills[0]);

        let wrapped_style = style.replace(
            "<cs:styleClr val=\"2\"><a:lumMod val=\"50000\"/>",
            "<cs:styleClr val=\"8\"><a:lumMod val=\"50000\"/>",
        );
        let cycle_colors = colors.replace("meth=\"acrossLinear\"", "meth=\"cycle\"");
        let wrapped = parse_chartex_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: Some(&wrapped_style),
                color_style_xml: Some(&cycle_colors),
                ..Default::default()
            },
        )
        .expect("cycle fixed index wraps");
        assert_eq!(
            wrapped
                .chartex_data_point_style
                .expect("wrapped point role")
                .fill_colors,
            Some(fixed_fills.clone()),
        );

        let line = model.chartex_data_point_line_style.expect("line role");
        assert_eq!(line.line_color_index, Some(0));
        assert_eq!(line.line_colors, Some(vec![Some("4472C4".to_string()); 3]));

        let marker = model.chartex_data_point_marker_style.expect("marker role");
        assert_eq!(
            marker.fill_colors,
            Some(vec![Some("AA5500".to_string()); 3])
        );
        assert_eq!(
            marker.line_colors,
            Some(vec![Some("ED7D31".to_string()); 3])
        );

        let unresolved_style = style.replace(
            "mods=\"ignoreCSTransforms\"><cs:styleClr val=\"2\"",
            "mods=\"ignoreCSTransforms\"><cs:styleClr val=\"1\"",
        );
        let unresolved = parse_chartex_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: Some(&unresolved_style),
                color_style_xml: Some(&colors),
                ..Default::default()
            },
        )
        .expect("unresolved fixed slot remains unresolved");
        assert_eq!(
            unresolved
                .chartex_data_point_marker_style
                .expect("marker role")
                .line_colors,
            None,
        );
    }

    #[test]
    fn chartex_cache_dimensions_reject_unbounded_counts_and_sparse_indices() {
        let huge_count = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chartData><cx:data>
                <cx:strDim type="cat"><cx:lvl ptCount="4294967295"/></cx:strDim>
                <cx:numDim type="size"><cx:lvl ptCount="4294967295"/></cx:numDim>
              </cx:data></cx:chartData>
            </cx:chartSpace>"#,
        );
        let count_doc = chart_space_of(&huge_count);
        let mut references = EmptyChartReferenceResolver;
        assert!(chartex_string_levels(count_doc.root_element(), &mut references).is_none());
        assert!(
            chartex_number_values(count_doc.root_element(), &["size"], &mut references,).is_none()
        );

        let huge_index = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}">
              <cx:chartData><cx:data>
                <cx:strDim type="cat"><cx:lvl><cx:pt idx="4294967295">x</cx:pt></cx:lvl></cx:strDim>
                <cx:numDim type="size"><cx:lvl><cx:pt idx="4294967295">1</cx:pt></cx:lvl></cx:numDim>
              </cx:data></cx:chartData>
            </cx:chartSpace>"#,
        );
        let index_doc = chart_space_of(&huge_index);
        assert!(chartex_string_levels(index_doc.root_element(), &mut references).is_none());
        assert!(
            chartex_number_values(index_doc.root_element(), &["size"], &mut references,).is_none()
        );

        let aggregate = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}">
              <cx:chartData><cx:data><cx:strDim type="cat">
                <cx:lvl ptCount="524289"/><cx:lvl ptCount="524288"/>
              </cx:strDim></cx:data></cx:chartData>
            </cx:chartSpace>"#,
        );
        let aggregate_doc = chart_space_of(&aggregate);
        assert!(chartex_string_levels(aggregate_doc.root_element(), &mut references).is_none());
    }

    /// (b) Treemap: the same deepest→root category levels as sunburst, plus the
    /// parent-label layout. The structured model preserves the full hierarchy;
    /// the legacy flat fields remain populated for compatibility.
    #[test]
    fn parse_chartex_part_treemap_hierarchy_values() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chartData>
                <cx:data id="0">
                  <cx:strDim type="cat">
                    <cx:lvl ptCount="3">
                      <cx:pt idx="0">North</cx:pt>
                      <cx:pt idx="1">South</cx:pt>
                      <cx:pt idx="2">East</cx:pt>
                    </cx:lvl>
                    <cx:lvl ptCount="3">
                      <cx:pt idx="0">Americas</cx:pt>
                      <cx:pt idx="1">Americas</cx:pt>
                      <cx:pt idx="2">Asia</cx:pt>
                    </cx:lvl>
                  </cx:strDim>
                  <cx:numDim type="val">
                    <cx:lvl ptCount="3">
                      <cx:pt idx="0">50</cx:pt>
                      <cx:pt idx="1">30</cx:pt>
                      <cx:pt idx="2">20</cx:pt>
                    </cx:lvl>
                  </cx:numDim>
                </cx:data>
              </cx:chartData>
              <cx:chart>
                <cx:plotArea>
                  <cx:plotAreaRegion>
                    <cx:series layoutId="treemap">
                      <cx:dataLabels pos="inEnd">
                        <cx:visibility seriesName="0" categoryName="1" value="1"/>
                        <cx:separator>&#10;</cx:separator>
                      </cx:dataLabels>
                      <cx:layoutPr><cx:parentLabelLayout val="banner"/></cx:layoutPr>
                    </cx:series>
                  </cx:plotAreaRegion>
                </cx:plotArea>
              </cx:chart>
            </cx:chartSpace>"#
        );
        let d = chart_space_of(&xml);
        let m = parse_chartex_part(
            d.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: None,
                ..Default::default()
            },
        )
        .expect("treemap parses");

        assert_eq!(m.chart_type, "treemap");
        assert_eq!(m.categories, vec!["North", "South", "East"]);
        assert_eq!(m.series[0].values, vec![Some(50.0), Some(30.0), Some(20.0)]);
        assert_eq!(m.series[0].color, None);
        assert_eq!(m.series[0].data_label_colors, None);
        let labels = m.series[0]
            .series_data_labels
            .as_ref()
            .expect("data labels");
        assert!(labels.show_cat_name);
        assert!(labels.show_val);
        assert!(!labels.show_ser_name);
        assert_eq!(labels.position.as_deref(), Some("inEnd"));
        assert_eq!(labels.separator.as_deref(), Some("\n"));
        let tm = m.chartex_treemap.expect("treemap data present");
        assert_eq!(tm.parent_label_layout.as_deref(), Some("banner"));
        assert_eq!(tm.rows.len(), 3);
        assert_eq!(tm.rows[0].path, vec!["Americas", "North"]);
        assert_eq!(tm.rows[0].size, 50.0);
        assert_eq!(tm.rows[2].path, vec!["Asia", "East"]);
        assert_eq!(tm.rows[2].size, 20.0);
        // No `<cx:subtotals>` → no total points. Index 0 still begins at the
        // zero baseline but keeps the ordinary positive/negative formatting.
        assert!(m.subtotal_indices.is_empty());
        // No `<cx:catScaling gapWidth>` → unset (renderer default applies).
        assert_eq!(m.bar_gap_width, None);
        assert!(!m.cat_axis_hidden);
        assert!(!m.val_axis_hidden);
    }

    #[test]
    fn parse_chartex_region_map_preserves_color_values_identity_and_geography() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chartData><cx:data id="0">
                <cx:strDim type="cat"><cx:lvl ptCount="3">
                  <cx:pt idx="0">United States</cx:pt><cx:pt idx="1">Japan</cx:pt><cx:pt idx="2">Unknown</cx:pt>
                </cx:lvl></cx:strDim>
                <cx:strDim type="entityId"><cx:lvl ptCount="3">
                  <cx:pt idx="0">US</cx:pt><cx:pt idx="1">JP</cx:pt>
                </cx:lvl></cx:strDim>
                <cx:numDim type="colorVal"><cx:lvl ptCount="3">
                  <cx:pt idx="0">850</cx:pt><cx:pt idx="1">640</cx:pt><cx:pt idx="2">NaN</cx:pt>
                </cx:lvl></cx:numDim>
              </cx:data></cx:chartData>
              <cx:chart><cx:plotArea><cx:plotAreaRegion>
                <cx:series layoutId="regionMap">
                  <cx:valueColors>
                    <cx:minColor><a:srgbClr val="C1E5F5"/></cx:minColor>
                    <cx:midColor><a:srgbClr val="5B9BD5"/></cx:midColor>
                    <cx:maxColor><a:srgbClr val="104862"/></cx:maxColor>
                  </cx:valueColors>
                  <cx:valueColorPositions count="3">
                    <cx:min><cx:extremeValue/></cx:min>
                    <cx:mid><cx:percent val="40"/></cx:mid>
                    <cx:max><cx:number val="1000"/></cx:max>
                  </cx:valueColorPositions>
                  <cx:dataId val="0"/>
                  <cx:layoutPr>
                    <cx:regionLabelLayout val="bestFitOnly"/>
                    <cx:geography projectionType="robinson" viewedRegionType="world"
                      cultureLanguage="en-US" cultureRegion="US" attribution="Provider">
                      <cx:geoCache provider="Example"><cx:binary>AA==</cx:binary></cx:geoCache>
                    </cx:geography>
                  </cx:layoutPr>
                </cx:series>
              </cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>"#,
        );
        let document = chart_space_of(&xml);
        let model = parse_chartex_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: None,
                ..Default::default()
            },
        )
        .expect("region map parses");

        assert_eq!(model.chart_type, "regionMap");
        assert_eq!(model.categories, vec!["United States", "Japan", "Unknown"]);
        assert_eq!(model.series[0].values, vec![Some(850.0), Some(640.0), None]);
        let map = model.chartex_region_map.expect("structured region map");
        assert_eq!(map.rows[0].entity_id.as_deref(), Some("US"));
        assert_eq!(map.rows[1].entity_id.as_deref(), Some("JP"));
        assert_eq!(map.rows[2].entity_id, None);
        assert_eq!(map.region_label_layout.as_deref(), Some("bestFitOnly"));
        let geography = map.geography.expect("geography");
        assert_eq!(geography.projection_type.as_deref(), Some("robinson"));
        assert_eq!(geography.viewed_region_type.as_deref(), Some("world"));
        assert_eq!(geography.cache_provider.as_deref(), Some("Example"));
        assert!(geography.cache_present);
        let colors = map.colors.expect("authored ramp");
        assert_eq!(colors.stop_count, Some(3));
        assert_eq!(colors.min_color.as_deref(), Some("C1E5F5"));
        assert_eq!(colors.mid_color.as_deref(), Some("5B9BD5"));
        assert_eq!(colors.max_color.as_deref(), Some("104862"));
        assert_eq!(colors.min_position.as_ref().unwrap().kind, "extremeValue");
        assert_eq!(colors.mid_position.as_ref().unwrap().value, Some(40.0));
        assert_eq!(colors.max_position.as_ref().unwrap().value, Some(1000.0));
    }

    #[test]
    fn parse_chartex_treemap_retains_hierarchy_node_label_overrides() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chartData><cx:data id="0">
                <cx:strDim type="cat">
                  <cx:lvl ptCount="2"><cx:pt idx="0">Leaf A</cx:pt><cx:pt idx="1">Leaf B</cx:pt></cx:lvl>
                  <cx:lvl ptCount="2"><cx:pt idx="0">Group A</cx:pt><cx:pt idx="1">Group B</cx:pt></cx:lvl>
                </cx:strDim>
                <cx:numDim type="size"><cx:lvl ptCount="2"><cx:pt idx="0">10</cx:pt><cx:pt idx="1">20</cx:pt></cx:lvl></cx:numDim>
              </cx:data></cx:chartData>
              <cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="treemap">
                <cx:dataLabels pos="inEnd">
                  <cx:visibility categoryName="1" value="1"/>
                  <cx:spPr><a:solidFill><a:srgbClr val="DDEEFF"/></a:solidFill></cx:spPr>
                  <cx:txPr><a:bodyPr lIns="1pt" anchor="b"/><a:p><a:pPr algn="l"><a:defRPr><a:solidFill><a:srgbClr val="008000"/></a:solidFill><a:latin typeface="Primary"/></a:defRPr></a:pPr><a:r><a:rPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:latin typeface="Ignored"/></a:rPr><a:t>ignored</a:t></a:r></a:p><a:p><a:pPr algn="r"><a:defRPr><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:defRPr></a:pPr></a:p></cx:txPr>
                  <cx:dataLabel idx="3"><cx:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:ln></cx:spPr><cx:txPr><a:bodyPr rIns="-0.5pt"/><a:p><a:r><a:rPr sz="900" i="1" lang="en-US" baseline="12.5%"><a:noFill/></a:rPr><a:t>Custom Leaf&#10;20</a:t></a:r><a:r><a:t> plain</a:t></a:r></a:p></cx:txPr></cx:dataLabel>
                </cx:dataLabels>
              </cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let d = chart_space_of(&xml);
        let m = parse_chartex_part(
            d.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: None,
                ..Default::default()
            },
        )
        .expect("treemap parses");
        let defaults = m.series[0].series_data_labels.as_ref().expect("defaults");
        assert_eq!(defaults.font_color.as_deref(), Some("008000"));
        assert_eq!(defaults.font_paint_authored, Some(true));
        assert_eq!(defaults.font_hidden, None);
        assert_eq!(defaults.font_face.as_deref(), Some("Primary"));
        assert_eq!(defaults.text_align.as_deref(), Some("l"));
        assert_eq!(defaults.text_l_ins_emu, Some(12_700));
        assert_eq!(defaults.text_vertical_anchor.as_deref(), Some("b"));
        assert_eq!(defaults.text_body_authored, Some(true));
        assert_eq!(
            defaults
                .label_box
                .as_ref()
                .and_then(|box_| box_.fill.as_deref()),
            Some("DDEEFF")
        );
        let overrides = m.series[0].data_label_overrides.as_ref().expect("override");
        assert_eq!(overrides.len(), 1);
        assert_eq!(overrides[0].idx, 3);
        assert_eq!(overrides[0].text, "Custom Leaf\n20 plain");
        assert_eq!(overrides[0].font_color, None);
        assert_eq!(overrides[0].font_paint_authored, None);
        assert_eq!(overrides[0].font_size_hpt, None);
        assert_eq!(overrides[0].font_italic, None);
        assert_eq!(overrides[0].font_language, None);
        assert_eq!(overrides[0].font_baseline, None);
        assert_eq!(overrides[0].text_r_ins_emu, Some(-6_350));
        assert_eq!(overrides[0].text_body_authored, Some(true));
        let rich_runs = overrides[0].rich_runs.as_ref().expect("rich runs");
        assert_eq!(rich_runs.len(), 2);
        assert_eq!(rich_runs[0].color, None);
        assert_eq!(rich_runs[0].color_paint_authored, Some(true));
        assert_eq!(rich_runs[0].color_hidden, Some(true));
        assert_eq!(rich_runs[1].color, None);
        assert_eq!(rich_runs[1].color_paint_authored, None);
        assert_eq!(
            overrides[0]
                .label_box
                .as_ref()
                .and_then(|box_| box_.border_color.as_deref()),
            Some("445566")
        );
        assert_eq!(m.series[0].data_label_colors, None);
    }

    #[test]
    fn parse_chartex_data_label_index_is_bounded_without_sparse_allocation() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:lvl ptCount="1"><cx:pt idx="0">1</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData>
              <cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="waterfall">
                <cx:dataLabels><cx:dataLabel idx="4294967295"><cx:txPr><a:p><a:r><a:t>x</a:t></a:r></a:p></cx:txPr></cx:dataLabel></cx:dataLabels>
                <cx:dataId val="0"/>
              </cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let document = chart_space_of(&xml);
        let model = parse_chartex_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: None,
                ..Default::default()
            },
        )
        .expect("bounded chart parses");
        assert_eq!(model.series[0].data_label_colors, None);
        assert_eq!(model.series[0].data_label_overrides, None);
    }

    #[test]
    fn chartex_point_label_paint_budget_is_atomic() {
        let stops = (0..=MAX_CHART_LABEL_GRADIENT_STOPS)
            .map(|index| format!(r#"<a:gs pos="{index}"><a:srgbClr val="112233"/></a:gs>"#))
            .collect::<String>();
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:lvl ptCount="1"><cx:pt idx="0">1</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData>
              <cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="waterfall">
                <cx:dataLabels>
                  <cx:spPr><a:solidFill><a:srgbClr val="445566"/></a:solidFill></cx:spPr>
                  <cx:dataLabel idx="0"><cx:spPr><a:gradFill><a:gsLst>{stops}</a:gsLst></a:gradFill></cx:spPr></cx:dataLabel>
                </cx:dataLabels><cx:dataId val="0"/>
              </cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let document = chart_space_of(&xml);
        let model = parse_chartex_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: None,
                ..Default::default()
            },
        )
        .expect("bounded chart parses");
        let defaults = model.series[0]
            .series_data_labels
            .as_ref()
            .expect("defaults");
        let series_box = defaults.label_box.as_ref().expect("series box");
        assert_eq!(series_box.fill_paint_authored, Some(true));
        assert!(series_box.fill.is_none() && series_box.fill_paint.is_none());
        let point_box = model.series[0].data_label_overrides.as_ref().unwrap()[0]
            .label_box
            .as_ref()
            .expect("point box");
        assert_eq!(point_box.fill_paint_authored, Some(true));
        assert!(point_box.fill.is_none() && point_box.fill_paint.is_none());
    }

    #[test]
    fn chartex_series_share_one_label_paint_component_budget() {
        let label_shape = r#"<cx:dataLabels><cx:spPr><a:gradFill><a:gsLst>
          <a:gs pos="0"><a:srgbClr val="112233"/></a:gs>
          <a:gs pos="100000"><a:srgbClr val="445566"/></a:gs>
        </a:gsLst></a:gradFill></cx:spPr></cx:dataLabels>"#;
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chart><cx:plotArea><cx:plotAreaRegion>
                <cx:series layoutId="clusteredColumn">{label_shape}</cx:series>
                <cx:series layoutId="clusteredColumn">{label_shape}</cx:series>
              </cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let document = chart_space_of(&xml);
        let series = document
            .root_element()
            .descendants()
            .filter(|node| node.is_element() && node.tag_name().name() == "series")
            .collect::<Vec<_>>();

        assert!(chartex_label_paint_recipes_within_limit(
            std::iter::once(series[0]),
            3
        ));
        assert!(chartex_label_paint_recipes_within_limit(
            std::iter::once(series[1]),
            3
        ));
        assert!(!chartex_label_paint_recipes_within_limit(
            series.iter().copied(),
            3
        ));

        // The chart-level decision is applied atomically to every retained
        // series: authored provenance remains, but no structured recipe is
        // expanded for an arbitrary prefix of the chart.
        for series in series {
            let (_, _, defaults) = parse_chartex_series_labels(series, 1, &FixtureResolver, false);
            let label_box = defaults
                .as_ref()
                .and_then(|labels| labels.label_box.as_ref())
                .expect("authored label box");
            assert_eq!(label_box.fill_paint_authored, Some(true));
            assert!(label_box.fill.is_none() && label_box.fill_paint.is_none());
        }
    }

    #[test]
    fn chartex_empty_run_keeps_point_default_for_composed_label_text() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:lvl ptCount="1"><cx:pt idx="0">1</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData>
              <cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="waterfall">
                <cx:dataLabels><cx:dataLabel idx="0"><cx:visibility value="1"/><cx:txPr><a:p><a:pPr><a:defRPr><a:solidFill><a:srgbClr val="008000"/></a:solidFill></a:defRPr></a:pPr><a:r><a:t/></a:r></a:p></cx:txPr></cx:dataLabel></cx:dataLabels>
                <cx:dataId val="0"/>
              </cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let document = chart_space_of(&xml);
        let model = parse_chartex_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: None,
                ..Default::default()
            },
        )
        .expect("chart parses");
        let label = &model.series[0].data_label_overrides.as_ref().unwrap()[0];
        assert_eq!(label.text, "");
        assert_eq!(label.rich_runs, None);
        assert_eq!(label.font_color.as_deref(), Some("008000"));
        assert_eq!(label.font_paint_authored, Some(true));
        assert_eq!(label.show_val, Some(true));
    }

    /// Excel may omit every `<cx:lvl>` cache and reference hidden workbook
    /// names (`_xlchart.v1.*`) instead. The package
    /// resolver must populate both the structured treemap and flat fallback.
    #[test]
    fn parse_chartex_part_treemap_resolves_formula_only_dimensions() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}">
              <cx:chartData><cx:data id="0">
                <cx:strDim type="cat"><cx:f>_xlchart.v1.0</cx:f></cx:strDim>
                <cx:numDim type="size"><cx:f>_xlchart.v1.2</cx:f></cx:numDim>
              </cx:data></cx:chartData>
              <cx:chart><cx:plotArea><cx:plotAreaRegion>
                <cx:series layoutId="treemap"><cx:dataId val="0"/><cx:layoutPr><cx:parentLabelLayout val="overlapping"/></cx:layoutPr></cx:series>
              </cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let document = chart_space_of(&xml);
        let mut references = FormulaOnlyTreemapResolver;
        let model = parse_chartex_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: None,
                references: std::cell::Cell::new(Some(&mut references)),
                ..Default::default()
            },
        )
        .expect("formula-only treemap parses");

        let treemap = model.chartex_treemap.expect("structured treemap data");
        assert_eq!(treemap.rows[0].path, vec!["Americas", "North"]);
        assert_eq!(treemap.rows[2].path, vec!["Asia", "East"]);
        assert_eq!(treemap.rows[2].size, 20.0);
        assert_eq!(model.categories, vec!["North", "South", "East"]);
        assert_eq!(
            model.series[0].values,
            vec![Some(50.0), Some(30.0), Some(20.0)]
        );
        assert_eq!(model.series[0].val_format_code.as_deref(), Some("#,##0"));
    }

    #[test]
    fn parse_chartex_formula_only_flat_layouts_do_not_invent_missing_categories() {
        for layout in ["waterfall", "clusteredColumn", "funnel", "paretoLine"] {
            let xml = format!(
                r#"<cx:chartSpace xmlns:cx="{CX_NS}">
                  <cx:chartData><cx:data id="0">
                    <cx:strDim type="cat"><cx:f>_xlchart.missing</cx:f></cx:strDim>
                    <cx:numDim type="val"><cx:f>_xlchart.values</cx:f></cx:numDim>
                  </cx:data></cx:chartData>
                  <cx:chart><cx:plotArea><cx:plotAreaRegion>
                    <cx:series layoutId="{layout}"><cx:tx><cx:txData><cx:f>_xlchart.name</cx:f></cx:txData></cx:tx><cx:dataId val="0"/></cx:series>
                  </cx:plotAreaRegion></cx:plotArea></cx:chart>
                </cx:chartSpace>"#
            );
            let document = chart_space_of(&xml);
            let mut references = FormulaOnlyFlatResolver;
            let model = parse_chartex_part(
                document.root_element(),
                &ChartParseContext {
                    color_resolver: Some(&FixtureResolver),
                    style_xml: None,
                    references: std::cell::Cell::new(Some(&mut references)),
                    ..Default::default()
                },
            )
            .expect("formula-only flat ChartEx parses");
            assert_eq!(model.series[0].name, "Authored series");
            assert_eq!(
                model.series[0].values,
                vec![Some(3.0), Some(2.0), Some(1.0)]
            );
            assert!(model.categories.is_empty());
        }
    }

    #[test]
    fn parse_chartex_preserves_an_unknown_future_layout_without_guessing() {
        for series_xml in [
            r#"<cx:series layoutId="futureLayout"><cx:dataId val="0"/></cx:series>"#,
            r#"<cx:series layoutId="clusteredColumn"><cx:dataId val="0"/></cx:series>
               <cx:series layoutId="futureLayout"><cx:dataId val="0"/></cx:series>"#,
            r#"<cx:series layoutId="futureLayout"><cx:dataId val="0"/></cx:series>
               <cx:series layoutId="paretoLine" ownerIdx="0"><cx:dataId val="0"/></cx:series>"#,
        ] {
            let xml = format!(
                r#"<cx:chartSpace xmlns:cx="{CX_NS}">
                  <cx:chartData><cx:data id="0">
                    <cx:strDim type="cat"><cx:lvl ptCount="2"><cx:pt idx="0">A</cx:pt><cx:pt idx="1">B</cx:pt></cx:lvl></cx:strDim>
                    <cx:numDim type="val"><cx:lvl ptCount="2"><cx:pt idx="0">2</cx:pt><cx:pt idx="1">-1</cx:pt></cx:lvl></cx:numDim>
                  </cx:data></cx:chartData>
                  <cx:chart><cx:plotArea><cx:plotAreaRegion>{series_xml}</cx:plotAreaRegion></cx:plotArea></cx:chart>
                </cx:chartSpace>"#
            );
            let document = chart_space_of(&xml);
            let model = parse_chartex_part(
                document.root_element(),
                &ChartParseContext {
                    color_resolver: Some(&FixtureResolver),
                    style_xml: None,
                    ..Default::default()
                },
            )
            .expect("future ChartEx layout remains inspectable");

            assert_eq!(model.chart_type, "futureLayout");
            assert_eq!(model.categories, vec!["A", "B"]);
            assert_eq!(model.series[0].values, vec![Some(2.0), Some(-1.0)]);
            assert!(model.chartex_histogram_binning.is_none());
            assert!(model.chartex_box.is_none());
            assert!(model.chartex_sunburst.is_none());
            assert!(model.chartex_treemap.is_none());
            assert!(model.chartex_region_map.is_none());
        }
    }

    #[test]
    fn parse_chartex_flat_series_resolve_data_id_and_skip_hidden_series() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chartData>
                <cx:data id="0">
                  <cx:strDim type="cat"><cx:lvl ptCount="1"><cx:pt idx="0">Unused</cx:pt></cx:lvl></cx:strDim>
                  <cx:numDim type="val"><cx:lvl ptCount="1"><cx:pt idx="0">999</cx:pt></cx:lvl></cx:numDim>
                </cx:data>
                <cx:data id="1">
                  <cx:strDim type="cat"><cx:lvl ptCount="2"><cx:pt idx="0">A</cx:pt><cx:pt idx="1">B</cx:pt></cx:lvl></cx:strDim>
                  <cx:numDim type="val"><cx:lvl ptCount="2"><cx:pt idx="0">1</cx:pt><cx:pt idx="1">2</cx:pt></cx:lvl></cx:numDim>
                </cx:data>
                <cx:data id="2">
                  <cx:strDim type="cat"><cx:lvl ptCount="2"><cx:pt idx="0">A</cx:pt><cx:pt idx="1">B</cx:pt></cx:lvl></cx:strDim>
                  <cx:numDim type="val"><cx:lvl ptCount="2"><cx:pt idx="0">3</cx:pt><cx:pt idx="1">4</cx:pt></cx:lvl></cx:numDim>
                </cx:data>
                <cx:data id="3"><cx:numDim type="val"><cx:lvl ptCount="1"><cx:pt idx="0">5</cx:pt></cx:lvl></cx:numDim></cx:data>
              </cx:chartData>
              <cx:chart><cx:plotArea><cx:plotAreaRegion>
                <cx:series layoutId="waterfall" hidden="1"><cx:dataId val="0"/></cx:series>
                <cx:series layoutId="clusteredColumn" formatIdx="0"><cx:tx><cx:txData><cx:v>First</cx:v></cx:txData></cx:tx><cx:dataId val="1"/></cx:series>
                <cx:series layoutId="clusteredColumn"><cx:tx><cx:txData><cx:v>Second</cx:v></cx:txData></cx:tx>
                  <cx:dataPt idx="0"><cx:spPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:ln w="25400"><a:solidFill><a:srgbClr val="778899"/></a:solidFill><a:prstDash val="dash"/></a:ln></cx:spPr></cx:dataPt>
                  <cx:dataPt idx="1"><cx:spPr><a:noFill/><a:ln><a:noFill/></a:ln></cx:spPr></cx:dataPt>
                  <cx:dataLabels pos="outEnd"><cx:visibility value="1"/><cx:numFmt formatCode="0.0"/>
                    <cx:dataLabel idx="0"><cx:txPr><a:p><a:pPr><a:defRPr><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:defRPr></a:pPr></a:p></cx:txPr></cx:dataLabel>
                    <cx:dataLabelHidden idx="1"/>
                  </cx:dataLabels><cx:dataId val="2"/>
                </cx:series>
                <cx:series layoutId="paretoLine" ownerIdx="99"><cx:dataId val="3"/></cx:series>
              </cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let document = chart_space_of(&xml);
        let model = parse_chartex_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: None,
                ..Default::default()
            },
        )
        .expect("visible ChartEx series parse");

        assert_eq!(model.chart_type, "clusteredColumn");
        assert_eq!(model.categories, vec!["A", "B"]);
        assert_eq!(model.series.len(), 2);
        assert_eq!(model.series[0].name, "First");
        assert_eq!(model.series[0].chartex_format_idx, Some(0));
        assert_eq!(model.series[0].values, vec![Some(1.0), Some(2.0)]);
        assert_eq!(model.series[1].name, "Second");
        assert_eq!(model.series[1].chartex_format_idx, Some(2));
        assert_eq!(model.series[1].values, vec![Some(3.0), Some(4.0)]);
        assert_eq!(
            model.series[1].categories.as_deref(),
            Some(&["A".into(), "B".into()][..])
        );
        let labels = model.series[1]
            .series_data_labels
            .as_ref()
            .expect("series-local labels");
        assert!(labels.show_val);
        assert_eq!(labels.position.as_deref(), Some("outEnd"));
        assert_eq!(labels.format_code.as_deref(), Some("0.0"));
        let label_overrides = model.series[1].data_label_overrides.as_ref().unwrap();
        assert_eq!(
            label_overrides
                .iter()
                .find(|item| item.idx == 1)
                .unwrap()
                .deleted,
            Some(true)
        );
        assert_eq!(
            model.series[1].data_label_colors.as_ref().unwrap()[0].as_deref(),
            Some("445566")
        );
        let point = &model.series[1].data_point_overrides.as_ref().unwrap()[0];
        assert_eq!(point.idx, 0);
        assert_eq!(point.color.as_deref(), Some("112233"));
        assert_eq!(point.fill_hidden, Some(false));
        assert_eq!(point.line_color.as_deref(), Some("778899"));
        assert_eq!(point.line_width_emu, Some(25400));
        assert_eq!(point.line_dash.as_deref(), Some("dash"));
        assert_eq!(point.line_hidden, Some(false));
        let hidden_point = &model.series[1].data_point_overrides.as_ref().unwrap()[1];
        assert_eq!(hidden_point.idx, 1);
        assert_eq!(hidden_point.fill_hidden, Some(true));
        assert_eq!(hidden_point.line_hidden, Some(true));
    }

    #[test]
    fn parse_chartex_invalid_pareto_owners_remain_bounded() {
        let ordinary_series = (0..512)
            .map(|_| r#"<cx:series layoutId="clusteredColumn"><cx:dataId val="0"/></cx:series>"#)
            .collect::<String>();
        let invalid_pareto_series = (0..512)
            .map(|_| {
                r#"<cx:series layoutId="paretoLine" ownerIdx="999999"><cx:dataId val="0"/></cx:series>"#
            })
            .collect::<String>();
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}">
              <cx:chartData><cx:data id="0">
                <cx:strDim type="cat"><cx:lvl ptCount="1"><cx:pt idx="0">A</cx:pt></cx:lvl></cx:strDim>
                <cx:numDim type="val"><cx:lvl ptCount="1"><cx:pt idx="0">1</cx:pt></cx:lvl></cx:numDim>
              </cx:data></cx:chartData>
              <cx:chart><cx:plotArea><cx:plotAreaRegion>
                {ordinary_series}{invalid_pareto_series}
              </cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let document = chart_space_of(&xml);
        let model = parse_chartex_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: None,
                ..Default::default()
            },
        )
        .expect("ordinary series remain selectable");

        assert_eq!(model.chart_type, "clusteredColumn");
        assert_eq!(model.series.len(), 512);
    }

    #[test]
    fn parse_chartex_pareto_line_keeps_its_owner_and_direct_line_style() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chartData>
                <cx:data id="0">
                  <cx:strDim type="cat"><cx:lvl ptCount="3">
                    <cx:pt idx="0">Five</cx:pt><cx:pt idx="1">Twenty</cx:pt><cx:pt idx="2">Ten</cx:pt>
                  </cx:lvl></cx:strDim>
                  <cx:numDim type="val"><cx:lvl ptCount="3">
                    <cx:pt idx="0">5</cx:pt><cx:pt idx="1">20</cx:pt><cx:pt idx="2">10</cx:pt>
                  </cx:lvl></cx:numDim>
                </cx:data>
                <cx:data id="1"><cx:numDim type="val"><cx:lvl ptCount="3">
                  <cx:pt idx="0">0.142857</cx:pt><cx:pt idx="1">0.714286</cx:pt><cx:pt idx="2">1</cx:pt>
                </cx:lvl></cx:numDim></cx:data>
              </cx:chartData>
              <cx:chart><cx:plotArea><cx:plotAreaRegion>
                <cx:series layoutId="waterfall" hidden="1" formatIdx="1"/>
                <cx:series layoutId="clusteredColumn" formatIdx="7">
                  <cx:tx><cx:txData><cx:v>Frequency</cx:v></cx:txData></cx:tx>
                  <cx:dataPt idx="1"><cx:spPr><a:solidFill><a:srgbClr val="00AA00"/></a:solidFill></cx:spPr></cx:dataPt>
                  <cx:dataId val="0"/>
                </cx:series>
                <cx:series layoutId="paretoLine" ownerIdx="1" formatIdx="8">
                  <cx:tx><cx:txData><cx:v>Cumulative %</cx:v></cx:txData></cx:tx>
                  <cx:spPr><a:ln w="25400"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:ln></cx:spPr>
                  <cx:dataId val="1"/>
                </cx:series>
              </cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let document = chart_space_of(&xml);
        let model = parse_chartex_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: None,
                ..Default::default()
            },
        )
        .expect("owner-backed Pareto parse");

        assert_eq!(model.chart_type, "pareto");
        assert_eq!(model.categories, vec!["Five", "Twenty", "Ten"]);
        assert_eq!(model.series.len(), 2);
        assert_eq!(model.series[0].name, "Frequency");
        assert_eq!(model.series[0].chartex_format_idx, Some(7));
        assert_eq!(
            model.series[0].values,
            vec![Some(5.0), Some(20.0), Some(10.0)]
        );
        assert_eq!(model.series[1].name, "Cumulative %");
        assert_eq!(model.series[1].chartex_format_idx, Some(8));
        assert_eq!(model.series[1].series_type.as_deref(), Some("line"));
        assert_eq!(model.series[1].use_secondary_axis, Some(true));
        assert_eq!(model.series[1].color.as_deref(), Some("333333"));
        assert_eq!(model.series[1].line_width_emu, Some(25400));
    }

    #[test]
    fn parse_chartex_histogram_retains_binning_contract() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}">
              <cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:lvl ptCount="5">
                <cx:pt idx="0">0</cx:pt><cx:pt idx="1">1</cx:pt><cx:pt idx="2">2</cx:pt>
                <cx:pt idx="3">3</cx:pt><cx:pt idx="4">4</cx:pt>
              </cx:lvl></cx:numDim></cx:data></cx:chartData>
              <cx:chart><cx:plotArea><cx:plotAreaRegion>
                <cx:series layoutId="clusteredColumn"><cx:dataId val="0"/><cx:layoutPr>
                  <cx:binning intervalClosed="r" underflow="0" overflow="4"><cx:binCount>2</cx:binCount></cx:binning>
                </cx:layoutPr></cx:series>
              </cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let document = chart_space_of(&xml);
        let model = parse_chartex_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: None,
                ..Default::default()
            },
        )
        .expect("histogram parses");
        let binning = model
            .chartex_histogram_binning
            .expect("histogram binning contract");
        assert_eq!(model.chart_type, "histogram");
        assert_eq!(binning.bin_count, Some(2));
        assert_eq!(binning.bin_size, None);
        assert_eq!(binning.interval_closed.as_deref(), Some("r"));
        assert_eq!(binning.underflow, Some(0.0));
        assert_eq!(binning.overflow, Some(4.0));
        assert_eq!(
            model.series[0].values,
            vec![Some(0.0), Some(1.0), Some(2.0), Some(3.0), Some(4.0)]
        );
    }

    #[test]
    fn parse_chartex_histogram_binning_accepts_size_and_keeps_auto_unset() {
        let size_xml = r#"<cx:series xmlns:cx="urn:cx"><cx:layoutPr><cx:binning intervalClosed="l" underflow="auto"><cx:binSize>0.5</cx:binSize></cx:binning></cx:layoutPr></cx:series>"#;
        let size_document = root_of(size_xml);
        let size =
            parse_chartex_histogram_binning(size_document.root_element()).expect("bin size parses");
        assert_eq!(size.bin_size, Some(0.5));
        assert_eq!(size.bin_count, None);
        assert_eq!(size.interval_closed.as_deref(), Some("l"));
        assert_eq!(size.underflow, None);

        let invalid_xml = r#"<cx:series xmlns:cx="urn:cx"><cx:layoutPr><cx:binning intervalClosed="x" overflow="NaN"><cx:binSize>-1</cx:binSize></cx:binning></cx:layoutPr></cx:series>"#;
        let invalid_document = root_of(invalid_xml);
        let invalid = parse_chartex_histogram_binning(invalid_document.root_element())
            .expect("empty automatic contract remains present");
        assert_eq!(
            invalid,
            ChartexHistogramBinning {
                bin_size: None,
                bin_count: None,
                interval_closed: None,
                underflow: None,
                overflow: None,
            }
        );
    }

    /// (c) A `<cx:chartSpace>` with no `<cx:series>` is not a chartEx chart —
    /// `parse_chartex_part` returns `None` rather than an empty model.
    #[test]
    fn parse_chartex_part_returns_none_without_series() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}"><cx:chart><cx:plotArea/></cx:chart></cx:chartSpace>"#
        );
        let d = chart_space_of(&xml);
        assert!(parse_chartex_part(
            d.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: None,
                ..Default::default()
            }
        )
        .is_none());
    }

    /// (d) Newlines inside a category `<cx:pt>` are flattened to spaces (Office
    /// writes multi-line axis labels this way; the renderer wants a single
    /// line).
    #[test]
    fn parse_chartex_part_category_newline_flattened() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}">
              <cx:chartData><cx:data id="0">
                <cx:strDim type="cat"><cx:lvl ptCount="1"><cx:pt idx="0">FY2024
1Q</cx:pt></cx:lvl></cx:strDim>
                <cx:numDim type="val"><cx:lvl ptCount="1"><cx:pt idx="0">5</cx:pt></cx:lvl></cx:numDim>
              </cx:data></cx:chartData>
              <cx:chart><cx:plotArea><cx:plotAreaRegion>
                <cx:series layoutId="waterfall"/>
              </cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let d = chart_space_of(&xml);
        let m = parse_chartex_part(
            d.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: None,
                ..Default::default()
            },
        )
        .expect("parses");
        assert_eq!(m.categories, vec!["FY2024 1Q"]);
    }

    /// A box-and-whisker chart with two series, each referencing its own
    /// `<cx:data>` (via `<cx:dataId>`) of RAW sample points grouped across two
    /// categories. Verifies: (a) categories unique-in-order, (b) each series'
    /// points binned by category, (c) absent explicit series fills preserved,
    /// (d) `<cx:visibility>` / `<cx:statistics>` flags threaded, (e) the title
    /// is parsed and the accent palette exposed.
    #[test]
    fn parse_chartex_part_boxwhisker_two_series_binned_by_category() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chartData>
                <cx:data id="0">
                  <cx:strDim type="cat"><cx:lvl ptCount="3">
                    <cx:pt idx="0">Cat A</cx:pt><cx:pt idx="1">Cat A</cx:pt><cx:pt idx="2">Cat B</cx:pt>
                  </cx:lvl></cx:strDim>
                  <cx:numDim type="val"><cx:lvl ptCount="3">
                    <cx:pt idx="0">1</cx:pt><cx:pt idx="1">3</cx:pt><cx:pt idx="2">10</cx:pt>
                  </cx:lvl></cx:numDim>
                </cx:data>
                <cx:data id="1">
                  <cx:strDim type="cat"><cx:lvl ptCount="3">
                    <cx:pt idx="0">Cat A</cx:pt><cx:pt idx="1">Cat B</cx:pt><cx:pt idx="2">Cat B</cx:pt>
                  </cx:lvl></cx:strDim>
                  <cx:numDim type="val"><cx:lvl ptCount="3">
                    <cx:pt idx="0">5</cx:pt><cx:pt idx="1">7</cx:pt><cx:pt idx="2">9</cx:pt>
                  </cx:lvl></cx:numDim>
                </cx:data>
              </cx:chartData>
              <cx:chart>
                <cx:title><cx:tx><cx:rich><a:p><a:r><a:t>My box chart</a:t></a:r></a:p></cx:rich></cx:tx></cx:title>
                <cx:plotArea><cx:plotAreaRegion>
                  <cx:series layoutId="boxWhisker">
                    <cx:tx><cx:txData><cx:v>Series1</cx:v></cx:txData></cx:tx>
                    <cx:dataId val="0"/>
                    <cx:layoutPr>
                      <cx:visibility meanLine="0" meanMarker="1" nonoutliers="0" outliers="1"/>
                      <cx:statistics quartileMethod="exclusive"/>
                    </cx:layoutPr>
                  </cx:series>
                  <cx:series layoutId="boxWhisker">
                    <cx:tx><cx:txData><cx:v>Series2</cx:v></cx:txData></cx:tx>
                    <cx:dataId val="1"/>
                    <cx:layoutPr>
                      <cx:visibility meanLine="1" meanMarker="0" nonoutliers="1" outliers="0"/>
                      <cx:statistics quartileMethod="inclusive"/>
                    </cx:layoutPr>
                  </cx:series>
                </cx:plotAreaRegion></cx:plotArea>
              </cx:chart>
            </cx:chartSpace>"#
        );
        let d = chart_space_of(&xml);
        let m = parse_chartex_part(
            d.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: None,
                ..Default::default()
            },
        )
        .expect("boxWhisker parses");
        assert_eq!(m.chart_type, "boxWhisker");
        assert_eq!(m.title.as_deref(), Some("My box chart"));
        assert_eq!(
            m.chartex_accents.as_deref(),
            Some(
                &["5B9BD5", "ED7D31", "A5A5A5", "FFC000", "4472C4", "70AD47"].map(String::from)[..]
            )
        );
        let box_data = m.chartex_box.expect("box data present");
        assert_eq!(box_data.categories, vec!["Cat A", "Cat B"]);
        assert_eq!(box_data.series.len(), 2);

        let s0 = &box_data.series[0];
        assert_eq!(s0.name, "Series1");
        assert_eq!(s0.color, None); // shared renderer applies style/theme fallback
                                    // Series1: Cat A got points 1 & 3, Cat B got 10.
        assert_eq!(s0.values_by_category, vec![vec![1.0, 3.0], vec![10.0]]);
        assert!(s0.mean_marker && !s0.mean_line && s0.show_outliers && !s0.show_nonoutliers);
        assert_eq!(s0.quartile_method, "exclusive");

        let s1 = &box_data.series[1];
        assert_eq!(s1.name, "Series2");
        assert_eq!(s1.color, None); // shared renderer applies style/theme fallback
                                    // Series2: Cat A got 5, Cat B got 7 & 9.
        assert_eq!(s1.values_by_category, vec![vec![5.0], vec![7.0, 9.0]]);
        assert!(!s1.mean_marker && s1.mean_line && !s1.show_outliers && s1.show_nonoutliers);
        assert_eq!(s1.quartile_method, "inclusive");
    }

    #[test]
    fn parse_chartex_boxwhisker_retains_direct_picture_marker_relationship() {
        struct Images;
        impl ChartImageResolver for Images {
            fn resolve_image(
                &self,
                source: ChartImageSource,
                relationship_id: &str,
            ) -> Option<(String, String)> {
                (source == ChartImageSource::Chart && relationship_id == "rIdBox")
                    .then(|| ("xl/media/box.svg".to_owned(), "image/svg+xml".to_owned()))
            }
        }
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <cx:chartData><cx:data id="0"><cx:strDim type="cat"><cx:lvl ptCount="1"><cx:pt idx="0">A</cx:pt></cx:lvl></cx:strDim>
                <cx:numDim type="val"><cx:lvl ptCount="1"><cx:pt idx="0">1</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData>
              <cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="boxWhisker">
                <cx:tx><cx:txData><cx:v>S</cx:v></cx:txData></cx:tx><cx:dataId val="0"/>
                <cx:spPr><a:blipFill rotWithShape="1"><a:blip><a:extLst><a:ext uri="{{96DAC541-7B7A-43D3-8B79-37D633B846F1}}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rIdBox"/></a:ext></a:extLst></a:blip><a:stretch/></a:blipFill></cx:spPr>
                <cx:layoutPr><cx:visibility nonoutliers="1" outliers="1"/></cx:layoutPr>
              </cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let document = chart_space_of(&xml);
        let model = parse_chartex_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: None,
                color_style_xml: None,
                images: Some(&Images),
                ..Default::default()
            },
        )
        .expect("box chart parses");
        assert!(matches!(
            model.chartex_box.as_ref().and_then(|box_chart| box_chart.series[0]
                .chartex_style.as_ref())
                .and_then(|style| style.fill_paints.as_ref())
                .and_then(|paints| paints.first())
                .and_then(Option::as_ref),
            Some(ChartStyleFill::Image { image_path, rot_with_shape: Some(true), .. })
                if image_path == "xl/media/box.svg"
        ));
    }

    #[test]
    fn parse_chartex_part_boxwhisker_discards_non_finite_values_and_keeps_repeats() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chartData><cx:data id="0">
                <cx:strDim type="cat"><cx:lvl ptCount="5">
                  <cx:pt idx="0">Drop</cx:pt><cx:pt idx="1">Keep</cx:pt>
                  <cx:pt idx="2">Keep</cx:pt><cx:pt idx="3">Drop</cx:pt>
                  <cx:pt idx="4">Drop</cx:pt>
                </cx:lvl></cx:strDim>
                <cx:numDim type="val"><cx:lvl ptCount="5">
                  <cx:pt idx="0">NaN</cx:pt><cx:pt idx="1">5</cx:pt>
                  <cx:pt idx="2">5</cx:pt><cx:pt idx="3">inf</cx:pt>
                  <cx:pt idx="4">-inf</cx:pt>
                </cx:lvl></cx:numDim>
              </cx:data></cx:chartData>
              <cx:chart><cx:plotArea><cx:plotAreaRegion>
                <cx:series layoutId="boxWhisker"><cx:dataId val="0"/></cx:series>
              </cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let document = chart_space_of(&xml);
        let model = parse_chartex_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: None,
                ..Default::default()
            },
        )
        .expect("finite repeated samples remain plottable");
        let box_data = model.chartex_box.expect("box data");
        assert_eq!(box_data.categories, vec!["Keep"]);
        assert_eq!(box_data.series[0].values_by_category, vec![vec![5.0, 5.0]]);
    }

    /// Excel-authored XLSX box-and-whisker charts may omit every cache and
    /// store one formula-only numeric dimension per named series. In that form
    /// each series is one category/box rather than a shared categorized grid.
    #[test]
    fn parse_chartex_part_boxwhisker_resolves_formula_only_series() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:spPr><a:ln w="9525"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln></cx:spPr>
              <cx:chartData>
                <cx:data id="0"><cx:numDim type="val"><cx:f>_xlchart.v1.1</cx:f></cx:numDim></cx:data>
                <cx:data id="1"><cx:numDim type="val"><cx:f>_xlchart.v1.3</cx:f></cx:numDim></cx:data>
              </cx:chartData>
              <cx:chart>
                <cx:title><cx:tx><cx:rich><a:p><a:r>
                  <a:rPr sz="1100" b="1"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr>
                  <a:t>Readiness</a:t>
                </a:r></a:p></cx:rich></cx:tx></cx:title>
                <cx:plotArea><cx:plotAreaRegion>
                  <cx:series layoutId="boxWhisker">
                    <cx:tx><cx:txData><cx:v>Foundations</cx:v></cx:txData></cx:tx>
                    <cx:spPr><a:ln w="6350"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln></cx:spPr>
                    <cx:dataId val="0"/>
                  </cx:series>
                  <cx:series layoutId="boxWhisker">
                    <cx:tx><cx:txData><cx:f>_xlchart.name</cx:f></cx:txData></cx:tx>
                    <cx:dataId val="1"/>
                  </cx:series>
                </cx:plotAreaRegion>
                <cx:axis id="0" hidden="1"><cx:catScaling/></cx:axis>
                <cx:axis id="1">
                  <cx:valScaling min="1" max="6" majorUnit="0.25" minorUnit="0.05"/>
                  <cx:minorTickMarks type="in"/>
                  <cx:title><cx:tx><cx:rich><a:p><a:r>
                    <a:rPr sz="900" b="0"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr>
                    <a:t>A&amp;R readiness score</a:t>
                  </a:r></a:p></cx:rich></cx:tx></cx:title>
                  <cx:majorGridlines/>
                  <cx:numFmt formatCode="0.0" sourceLinked="0"/>
                  <cx:txPr><a:p><a:pPr><a:defRPr sz="1200" b="0"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="Calibri"/></a:defRPr></a:pPr></a:p></cx:txPr>
                </cx:axis>
                </cx:plotArea>
                <cx:legend pos="r"><cx:txPr><a:p><a:pPr><a:defRPr sz="900" b="0"><a:solidFill><a:srgbClr val="445566"/></a:solidFill><a:latin typeface="Calibri"/></a:defRPr></a:pPr></a:p></cx:txPr></cx:legend>
              </cx:chart>
            </cx:chartSpace>"#
        );
        let style = format!(
            r#"<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="{A_NS}">
              <cs:dataPointMarkerLayout symbol="circle" size="5"/>
              <cs:gridlineMajor><cs:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></cs:spPr></cs:gridlineMajor>
              <cs:dataPoint><cs:spPr><a:ln w="19050"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></cs:spPr></cs:dataPoint>
              <cs:valueAxis>
                <cs:fontRef idx="minor"><a:srgbClr val="595959"/></cs:fontRef>
                <cs:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill><a:prstDash val="dashDot"/></a:ln></cs:spPr>
                <cs:defRPr sz="900" b="0"/>
              </cs:valueAxis>
            </cs:chartStyle>"#
        );
        let d = chart_space_of(&xml);
        let mut references = FormulaOnlyBoxResolver;
        let m = parse_chartex_part(
            d.root_element(),
            &ChartParseContext {
                color_resolver: Some(&WhiteChartFixtureResolver),
                style_xml: Some(&style),
                references: std::cell::Cell::new(Some(&mut references)),
                ..Default::default()
            },
        )
        .expect("boxWhisker parses");

        assert_eq!(m.title.as_deref(), Some("Readiness"));
        assert_eq!(m.title_font_size_hpt, Some(1100));
        assert_eq!(m.title_font_bold, Some(true));
        assert_eq!(m.title_font_color.as_deref(), Some("000000"));
        assert_eq!(m.title_font_face.as_deref(), Some("Calibri"));
        assert_eq!(m.chart_bg.as_deref(), Some("FFFFFF"));
        assert_eq!(m.chart_border_color.as_deref(), Some("000000"));
        assert_eq!(m.chart_border_width_emu, Some(9525));
        assert_eq!((m.val_min, m.val_max), (Some(1.0), Some(6.0)));
        assert_eq!(m.val_axis_major_unit, Some(0.25));
        assert_eq!(m.val_axis_minor_unit, Some(0.05));
        assert_eq!(m.val_axis_minor_tick_mark.as_deref(), Some("in"));
        assert_eq!(m.val_axis_title.as_deref(), Some("A&R readiness score"));
        assert_eq!(m.val_axis_title_font_size_hpt, Some(900));
        assert_eq!(m.val_axis_title_font_bold, Some(false));
        assert_eq!(m.val_axis_title_font_color.as_deref(), Some("000000"));
        assert_eq!(m.val_axis_title_font_face.as_deref(), Some("Calibri"));
        assert_eq!(m.val_axis_format_code.as_deref(), Some("0.0"));
        assert_eq!(m.val_axis_major_gridlines, Some(true));
        // Axis-local text paint is direct formatting and remains authoritative
        // over the associated Chart Style valueAxis role. Other typography
        // properties still inherit independently.
        assert_eq!(m.val_axis_font_size_hpt, Some(1200));
        assert_eq!(m.val_axis_font_bold, Some(false));
        assert_eq!(m.val_axis_font_color.as_deref(), Some("112233"));
        assert_eq!(m.val_axis_font_face.as_deref(), Some("Calibri"));
        assert_eq!(m.val_axis_line_color, None);
        assert_eq!(m.val_axis_line_width_emu, None);
        assert_eq!(m.val_axis_line_dash, None);
        assert!(!m.val_axis_line_hidden);
        assert_eq!(m.val_axis_gridline_color, None);
        assert_eq!(m.val_axis_gridline_width_emu, None);
        let roles = m.chart_style_roles.as_ref().expect("linked style roles");
        assert_eq!(
            roles["valueAxis"].line_colors.as_ref().unwrap()[0].as_deref(),
            Some("BFBFBF")
        );
        assert_eq!(
            roles["gridlineMajor"].line_colors.as_ref().unwrap()[0].as_deref(),
            Some("D9D9D9")
        );
        assert_eq!(m.chart_style_marker_size_pt, Some(5));
        assert_eq!(m.chart_style_marker_symbol.as_deref(), Some("circle"));
        assert!(
            m.chartex_box
                .as_ref()
                .expect("structured box data")
                .one_box_per_series
        );
        let point_style = m
            .chartex_data_point_style
            .as_ref()
            .expect("dataPoint style");
        assert_eq!(
            point_style.line_colors.as_deref(),
            Some(&vec![Some("FFFFFF".to_string()); 6][..])
        );
        assert_eq!(point_style.line_width_emu, Some(19050));
        assert_eq!(point_style.line_hidden, None);
        // ChartEx does not inherit the classic chart-axis tick default. With
        // no `<cx:majorTickMarks>` element, Excel draws no tick marks.
        assert_eq!(m.val_axis_major_tick_mark, "none");
        assert_eq!(m.cat_axis_major_tick_mark, "none");
        assert_eq!(m.chartex_marker_size_pt, Some(5));
        assert_eq!(m.chartex_marker_symbol.as_deref(), Some("circle"));
        assert!(m.show_legend);
        assert_eq!(m.legend_pos.as_deref(), Some("r"));
        assert_eq!(m.legend_font_size_hpt, Some(900));
        assert_eq!(m.legend_font_bold, Some(false));
        assert_eq!(m.legend_font_color.as_deref(), Some("445566"));
        assert_eq!(m.legend_font_face.as_deref(), Some("Calibri"));
        let no_fill_style = style.replace(
            "<a:solidFill><a:schemeClr val=\"lt1\"/></a:solidFill>",
            "<a:noFill/>",
        );
        let mut no_fill_references = FormulaOnlyBoxResolver;
        let no_fill_model = parse_chartex_part(
            d.root_element(),
            &ChartParseContext {
                color_resolver: Some(&WhiteChartFixtureResolver),
                style_xml: Some(&no_fill_style),
                references: std::cell::Cell::new(Some(&mut no_fill_references)),
                ..Default::default()
            },
        )
        .expect("boxWhisker with noFill data-point style parses");
        let no_fill_point = no_fill_model
            .chartex_data_point_style
            .expect("dataPoint style");
        assert_eq!(no_fill_point.line_colors, None);
        assert_eq!(no_fill_point.line_width_emu, Some(19050));
        assert_eq!(no_fill_point.line_hidden, Some(true));
        let placeholder_style = format!(
            r#"<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="{A_NS}">
              <cs:dataPoint><cs:spPr>
                <a:solidFill><a:schemeClr val="phClr"><a:lumMod val="50000"/></a:schemeClr></a:solidFill>
                <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"><a:lumMod val="50000"/></a:schemeClr></a:solidFill></a:ln>
              </cs:spPr></cs:dataPoint>
            </cs:chartStyle>"#
        );
        let mut placeholder_references = FormulaOnlyBoxResolver;
        let placeholder_model = parse_chartex_part(
            d.root_element(),
            &ChartParseContext {
                color_resolver: Some(&WhiteChartFixtureResolver),
                style_xml: Some(&placeholder_style),
                references: std::cell::Cell::new(Some(&mut placeholder_references)),
                ..Default::default()
            },
        )
        .expect("boxWhisker with phClr data-point style parses");
        let raw_accents = placeholder_model.chartex_accents.expect("raw palette");
        let placeholder_style = placeholder_model
            .chartex_data_point_style
            .expect("dataPoint style");
        let placeholder_fills = placeholder_style.fill_colors.expect("phClr fills");
        let placeholder_lines = placeholder_style.line_colors.expect("phClr lines");
        assert_eq!(placeholder_lines, placeholder_fills.to_vec());
        assert_ne!(
            placeholder_fills[0].as_deref(),
            Some(raw_accents[0].as_str())
        );
        assert_ne!(placeholder_fills[0], placeholder_fills[1]);
        let box_data = m.chartex_box.expect("box data present");
        assert_eq!(box_data.categories, vec!["Foundations", "Adaptation"]);
        assert_eq!(
            box_data.series[0].values_by_category,
            vec![vec![1.0, 2.0, 3.0], vec![]]
        );
        assert_eq!(box_data.series[0].line_color.as_deref(), Some("000000"));
        assert_eq!(box_data.series[0].line_width_emu, Some(6350));
        assert_eq!(
            box_data.series[1].values_by_category,
            vec![vec![], vec![4.0, 5.0, 6.0]]
        );
        assert!(box_data.series.iter().all(|series| {
            series.mean_marker && series.show_outliers && series.show_nonoutliers
        }));
    }

    /// A sunburst with three `<cx:lvl>` (Leaf / Stem / Branch, in document
    /// order) and a `<cx:numDim type="size">`. Verifies each row's path is built
    /// root→leaf (Branch first) with empty trailing (leaf) cells trimmed so a
    /// node that is itself a leaf terminates early, and that sizes attach by idx.
    #[test]
    fn parse_chartex_part_sunburst_hierarchy_paths_trim_empty_tail() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
              <cx:chartData><cx:data id="0">
                <cx:strDim type="cat">
                  <cx:lvl ptCount="3">
                    <cx:pt idx="0">Leaf 1</cx:pt><cx:pt idx="1"/><cx:pt idx="2">Leaf 3</cx:pt>
                  </cx:lvl>
                  <cx:lvl ptCount="3">
                    <cx:pt idx="0">Stem 1</cx:pt><cx:pt idx="1">Leaf 2</cx:pt><cx:pt idx="2">Stem 2</cx:pt>
                  </cx:lvl>
                  <cx:lvl ptCount="3">
                    <cx:pt idx="0">Branch 1</cx:pt><cx:pt idx="1">Branch 1</cx:pt><cx:pt idx="2">Branch 2</cx:pt>
                  </cx:lvl>
                </cx:strDim>
                <cx:numDim type="size"><cx:lvl ptCount="3">
                  <cx:pt idx="0">22</cx:pt><cx:pt idx="1">17</cx:pt><cx:pt idx="2">18</cx:pt>
                </cx:lvl></cx:numDim>
              </cx:data></cx:chartData>
              <cx:chart>
                <cx:title><cx:tx><cx:rich><a:p><a:r><a:t>My sunburst</a:t></a:r></a:p></cx:rich></cx:tx></cx:title>
                <cx:plotArea><cx:plotAreaRegion>
                  <cx:series layoutId="sunburst">
                    <cx:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></cx:spPr>
                    <cx:dataId val="0"/>
                  </cx:series>
                </cx:plotAreaRegion></cx:plotArea>
              </cx:chart>
            </cx:chartSpace>"#
        );
        let d = chart_space_of(&xml);
        let m = parse_chartex_part(
            d.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: None,
                ..Default::default()
            },
        )
        .expect("sunburst parses");
        assert_eq!(m.chart_type, "sunburst");
        assert_eq!(m.title.as_deref(), Some("My sunburst"));
        assert_eq!(m.series[0].line_color.as_deref(), Some("FFFFFF"));
        assert_eq!(m.series[0].line_width_emu, Some(12700));
        assert_eq!(m.series[0].line_hidden, Some(false));
        let sb = m.chartex_sunburst.expect("sunburst data present");
        assert_eq!(sb.rows.len(), 3);
        // Row 0: full Branch→Stem→Leaf chain.
        assert_eq!(sb.rows[0].path, vec!["Branch 1", "Stem 1", "Leaf 1"]);
        assert_eq!(sb.rows[0].size, 22.0);
        // Row 1: empty Leaf cell → path terminates at Stem ("Leaf 2" is itself a leaf).
        assert_eq!(sb.rows[1].path, vec!["Branch 1", "Leaf 2"]);
        assert_eq!(sb.rows[1].size, 17.0);
        // Row 2: full chain under a different branch.
        assert_eq!(sb.rows[2].path, vec!["Branch 2", "Stem 2", "Leaf 3"]);
        assert_eq!(sb.rows[2].size, 18.0);

        let no_fill_xml = xml.replace(
            r#"<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>"#,
            "<a:noFill/>",
        );
        let no_fill_document = chart_space_of(&no_fill_xml);
        let no_fill_model = parse_chartex_part(
            no_fill_document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: None,
                ..Default::default()
            },
        )
        .expect("sunburst with an explicit noFill outline parses");
        assert_eq!(no_fill_model.series[0].line_color, None);
        assert_eq!(no_fill_model.series[0].line_width_emu, Some(12700));
        assert_eq!(no_fill_model.series[0].line_hidden, Some(true));
    }

    /// A waterfall chart must not get hierarchy/box structured fields. It does
    /// retain the shared ChartEx accent palette because positive, negative and
    /// subtotal columns use the chart theme's first three accents.
    #[test]
    fn parse_chartex_part_waterfall_leaves_structured_fields_none() {
        let xml = format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}">
              <cx:chartData><cx:data id="0">
                <cx:strDim type="cat"><cx:lvl ptCount="1"><cx:pt idx="0">A</cx:pt></cx:lvl></cx:strDim>
                <cx:numDim type="val"><cx:lvl ptCount="1"><cx:pt idx="0">5</cx:pt></cx:lvl></cx:numDim>
              </cx:data></cx:chartData>
              <cx:chart><cx:plotArea><cx:plotAreaRegion>
                <cx:series layoutId="waterfall"/>
              </cx:plotAreaRegion></cx:plotArea></cx:chart>
            </cx:chartSpace>"#
        );
        let d = chart_space_of(&xml);
        let m = parse_chartex_part(
            d.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: None,
                ..Default::default()
            },
        )
        .expect("waterfall parses");
        assert!(m.chartex_box.is_none());
        assert!(m.chartex_sunburst.is_none());
        assert!(m.chartex_treemap.is_none());
        assert_eq!(
            m.chartex_accents.as_ref().expect("waterfall accents"),
            &vec![
                "5B9BD5".to_string(),
                "ED7D31".to_string(),
                "A5A5A5".to_string(),
                "FFC000".to_string(),
                "4472C4".to_string(),
                "70AD47".to_string(),
            ]
        );
        assert!(m.title.is_none());
    }

    /// `<cs:title><cs:defRPr sz>` in a chartStyle part extracts the title size
    /// (hpt). Word's default modern chart style writes 1400 (14pt).
    #[test]
    fn extract_chartex_style_title_size_reads_cs_defrpr() {
        let style = format!(
            r#"<cs:chartStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}">
              <cs:title><cs:defRPr sz="1400" b="0"/></cs:title>
            </cs:chartStyle>"#
        );
        assert_eq!(extract_chartex_style_title_size(&style), Some(1400));
        // No <cs:title> / no sz → None.
        assert!(extract_chartex_style_title_size(&format!(
            r#"<cs:chartStyle xmlns:cs="{CS_NS}"><cs:dataPoint/></cs:chartStyle>"#
        ))
        .is_none());
        // Malformed XML → None (not a panic).
        assert!(extract_chartex_style_title_size("<not xml").is_none());
    }

    /// A chartEx title keeps direct text separate from its linked style role;
    /// core later applies that role, while an inline `sz` remains direct.
    #[test]
    fn parse_chartex_part_title_size_resolves_from_style_part() {
        let chart_xml = |title_rpr: &str| {
            format!(
                r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="{A_NS}">
                  <cx:chartData><cx:data id="0">
                    <cx:strDim type="cat"><cx:lvl ptCount="1"><cx:pt idx="0">Leaf</cx:pt></cx:lvl></cx:strDim>
                    <cx:numDim type="size"><cx:lvl ptCount="1"><cx:pt idx="0">1</cx:pt></cx:lvl></cx:numDim>
                  </cx:data></cx:chartData>
                  <cx:chart>
                    <cx:title><cx:tx><cx:rich><a:p><a:pPr>{title_rpr}</a:pPr>
                      <a:r><a:t>T</a:t></a:r></a:p></cx:rich></cx:tx></cx:title>
                    <cx:plotArea><cx:plotAreaRegion>
                      <cx:series layoutId="sunburst"><cx:dataId val="0"/></cx:series>
                    </cx:plotAreaRegion></cx:plotArea>
                  </cx:chart>
                </cx:chartSpace>"#
            )
        };
        let style = format!(
            r#"<cs:chartStyle xmlns:cs="{CS_NS}" xmlns:a="{A_NS}"><cs:title><cs:fontRef idx="major"><a:srgbClr val="445566"/></cs:fontRef><cs:defRPr sz="1400" b="0"/></cs:title></cs:chartStyle>"#
        );

        // An empty inline carrier stays non-owning; the linked role owns the
        // typography, including its explicit regular bold/italic state.
        let x0 = chart_xml("<a:defRPr/>");
        let d0 = chart_space_of(&x0);
        let m0 = parse_chartex_part(
            d0.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: Some(&style),
                ..Default::default()
            },
        )
        .unwrap();
        assert_eq!(m0.title_font_size_hpt, None);
        assert_eq!(m0.title_font_bold, None);
        assert_eq!(m0.title_font_italic, None);
        assert_eq!(m0.title_font_color, None);
        assert_eq!(m0.title_font_face, None);
        let title_role = &m0.chart_style_roles.as_ref().unwrap()["title"];
        assert_eq!(title_role.font_size_hpt, Some(1400));
        assert_eq!(title_role.font_bold, Some(false));
        assert_eq!(title_role.font_italic, Some(false));
        assert_eq!(title_role.font_color.as_deref(), Some("445566"));
        assert_eq!(title_role.font_face.as_deref(), Some("Calibri Light"));

        // Inline sz on the title wins over the style part.
        let x1 = chart_xml(r#"<a:defRPr sz="2000"/>"#);
        let d1 = chart_space_of(&x1);
        let m1 = parse_chartex_part(
            d1.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: Some(&style),
                ..Default::default()
            },
        )
        .unwrap();
        assert_eq!(m1.title_font_size_hpt, Some(2000));

        // No style part and no inline sz → None (renderer fallback).
        let x2 = chart_xml("<a:defRPr/>");
        let d2 = chart_space_of(&x2);
        let m2 = parse_chartex_part(
            d2.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: None,
                ..Default::default()
            },
        )
        .unwrap();
        assert_eq!(m2.title_font_size_hpt, None);

        // With no character-property node at all, the title remains inheritable.
        let x3 = chart_xml("");
        let d3 = chart_space_of(&x3);
        let m3 = parse_chartex_part(
            d3.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                style_xml: Some(&style),
                ..Default::default()
            },
        )
        .unwrap();
        assert_eq!(m3.title_font_bold, None);
        assert_eq!(m3.title_font_italic, None);
    }
}
