#[cfg(test)]
mod tests {
    use super::super::*;

    #[test]
    fn legend_present_with_pos() {
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:legend><c:legendPos val="t"/></c:legend>
        </c:chart>"#;
        let d = root_of(xml);
        let (show, pos) = extract_legend(d.root_element());
        assert!(show);
        assert_eq!(pos.as_deref(), Some("t"));
    }

    #[test]
    fn legend_absent() {
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#;
        let d = root_of(xml);
        let (show, pos) = extract_legend(d.root_element());
        assert!(!show);
        assert!(pos.is_none());
    }

    #[test]
    fn legend_overlay_and_entry_overrides_preserve_index_delete_and_text_properties() {
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:legend>
              <c:legendEntry><c:idx val="2"/><c:delete/></c:legendEntry>
              <c:legendEntry><c:idx val="0"/><c:txPr><a:p><a:pPr>
                <a:defRPr sz="1400" b="1"><a:solidFill><a:srgbClr val="AABBCC"/></a:solidFill><a:latin typeface="Aptos"/></a:defRPr>
              </a:pPr></a:p></c:txPr></c:legendEntry>
              <c:overlay/>
            </c:legend>
        </c:chart>"#;
        let document = root_of(xml);
        let (overlay, entries) = extract_legend_overrides(document.root_element(), &StubResolver);
        assert_eq!(overlay, Some(true));
        assert_eq!(
            entries,
            Some(vec![
                ChartLegendEntryOverride {
                    idx: 2,
                    deleted: Some(true),
                    font_face: None,
                    font_color: None,
                    font_size_hpt: None,
                    font_bold: None,
                    font_italic: None,
                },
                ChartLegendEntryOverride {
                    idx: 0,
                    deleted: None,
                    font_face: Some("Aptos".to_string()),
                    font_color: Some("AABBCC".to_string()),
                    font_size_hpt: Some(1400),
                    font_bold: Some(true),
                    font_italic: Some(false),
                },
            ])
        );

        let explicit_false = root_of(
            r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:legend><c:overlay val="0"/></c:legend></c:chart>"#,
        );
        assert_eq!(
            extract_legend_overrides(explicit_false.root_element(), &StubResolver).0,
            Some(false)
        );
    }

    #[test]
    fn data_label_position() {
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:plotArea><c:dLbls><c:dLblPos val="ctr"/></c:dLbls></c:plotArea>
        </c:chart>"#;
        let d = root_of(xml);
        assert_eq!(
            extract_data_label_position(d.root_element()).as_deref(),
            Some("ctr")
        );

        let series_only = root_of(
            r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:plotArea><c:scatterChart><c:ser><c:dLbls><c:dLblPos val="l"/></c:dLbls></c:ser></c:scatterChart></c:plotArea>
            </c:chart>"#,
        );
        assert_eq!(
            extract_data_label_position(series_only.root_element()),
            None,
            "a series position must not become the sibling-series fallback",
        );
    }

    #[test]
    fn show_data_labels_over_max_preserves_absent_false_and_bare_true() {
        let absent = root_of(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart/></c:chartSpace>"#,
        );
        assert_eq!(
            extract_show_data_labels_over_max(absent.root_element()),
            None
        );

        let explicit_false = root_of(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:showDLblsOverMax val="0"/></c:chart></c:chartSpace>"#,
        );
        assert_eq!(
            extract_show_data_labels_over_max(explicit_false.root_element()),
            Some(false),
        );

        let bare = root_of(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:showDLblsOverMax/></c:chart></c:chartSpace>"#,
        );
        assert_eq!(
            extract_show_data_labels_over_max(bare.root_element()),
            Some(true),
        );
    }

    #[test]
    fn parse_chart_user_shapes_preserves_relative_text_boxes_and_run_formatting() {
        let doc = roxmltree::Document::parse(
            r#"<c:userShapes xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                 xmlns:cdr="http://schemas.openxmlformats.org/drawingml/2006/chartDrawing"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <cdr:relSizeAnchor>
                <cdr:from><cdr:x>0</cdr:x><cdr:y>0.05</cdr:y></cdr:from>
                <cdr:to><cdr:x>0.8</cdr:x><cdr:y>0.16</cdr:y></cdr:to>
                <cdr:sp>
                  <cdr:nvSpPr><cdr:cNvPr id="1" name="TitleBox"/><cdr:cNvSpPr txBox="1"/></cdr:nvSpPr>
                  <cdr:spPr/>
                  <cdr:txBody>
                    <a:bodyPr anchor="b" wrap="square" lIns="12700" tIns="25400" rIns="38100" bIns="50800"/><a:lstStyle/>
                    <a:p>
                      <a:pPr algn="ctr"><a:defRPr sz="1200"><a:latin typeface="Lato"/></a:defRPr></a:pPr>
                      <a:r><a:rPr sz="2000" b="1"><a:solidFill><a:srgbClr val="1696d2"/></a:solidFill></a:rPr><a:t>Authored </a:t></a:r>
                      <a:r><a:t>title</a:t></a:r>
                    </a:p>
                  </cdr:txBody>
                </cdr:sp>
              </cdr:relSizeAnchor>
            </c:userShapes>"#,
        )
        .unwrap();

        let boxes = parse_chart_user_shapes(doc.root_element(), &StubResolver);
        assert_eq!(boxes.len(), 1);
        assert_eq!(boxes[0].x, 0.0);
        assert_eq!(boxes[0].y, 0.05);
        assert_eq!(boxes[0].w, 0.8);
        assert_eq!(boxes[0].h, 0.11);
        assert_eq!(boxes[0].vertical_anchor.as_deref(), Some("b"));
        assert_eq!(boxes[0].wrap.as_deref(), Some("square"));
        assert_eq!(boxes[0].l_ins, 12700);
        assert_eq!(boxes[0].t_ins, 25400);
        assert_eq!(boxes[0].r_ins, 38100);
        assert_eq!(boxes[0].b_ins, 50800);
        assert_eq!(boxes[0].paragraphs[0].align.as_deref(), Some("ctr"));
        assert_eq!(boxes[0].paragraphs[0].runs[0].text, "Authored ");
        assert_eq!(boxes[0].paragraphs[0].runs[0].font_size_hpt, Some(2000));
        assert_eq!(boxes[0].paragraphs[0].runs[0].bold, Some(true));
        assert_eq!(
            boxes[0].paragraphs[0].runs[0].color.as_deref(),
            Some("1696D2")
        );
        assert_eq!(boxes[0].paragraphs[0].runs[1].text, "title");
        assert_eq!(boxes[0].paragraphs[0].runs[1].font_size_hpt, Some(1200));
        assert_eq!(
            boxes[0].paragraphs[0].runs[1].font_face.as_deref(),
            Some("Lato")
        );
    }

    #[test]
    fn data_label_font_color_resolves_via_resolver() {
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:plotArea>
                <c:dLbls>
                    <c:txPr><a:p><a:r><a:rPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:rPr></a:r></a:p></c:txPr>
                </c:dLbls>
            </c:plotArea>
        </c:chart>"#;
        let d = root_of(xml);
        let got = extract_data_label_font_color(d.root_element(), &StubResolver);
        assert_eq!(got.as_deref(), Some("bg1"));
    }

    #[test]
    fn data_label_font_color_skips_label_background_fill() {
        // `<c:spPr><a:solidFill>` (label background) must not be picked up;
        // only the text fill inside `<c:txPr>` counts.
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:plotArea>
                <c:dLbls>
                    <c:spPr><a:solidFill><a:srgbClr val="aabbcc"/></a:solidFill></c:spPr>
                </c:dLbls>
            </c:plotArea>
        </c:chart>"#;
        let d = root_of(xml);
        let got = extract_data_label_font_color(d.root_element(), &StubResolver);
        assert!(
            got.is_none(),
            "spPr fill must not leak into the font color: got {got:?}"
        );
    }

    #[test]
    fn data_label_font_color_first_dlbls_wins() {
        // Mimics Office writers that put a chart-level dLbls block AND
        // per-series ones — the first txPr resolution wins.
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:plotArea>
                <c:dLbls>
                    <c:txPr><a:p><a:r><a:rPr><a:solidFill><a:srgbClr val="ffffff"/></a:solidFill></a:rPr></a:r></a:p></c:txPr>
                </c:dLbls>
                <c:barChart>
                    <c:ser><c:dLbls>
                        <c:txPr><a:p><a:r><a:rPr><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:rPr></a:r></a:p></c:txPr>
                    </c:dLbls></c:ser>
                </c:barChart>
            </c:plotArea>
        </c:chart>"#;
        let d = root_of(xml);
        let got = extract_data_label_font_color(d.root_element(), &StubResolver);
        assert_eq!(got.as_deref(), Some("FFFFFF"));
    }

    #[test]
    fn chart_title_text_size_bold_srgb() {
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:title><c:tx><c:rich>
                <a:p><a:pPr><a:defRPr sz="1400" b="1"><a:solidFill><a:srgbClr val="1B4332"/></a:solidFill></a:defRPr></a:pPr>
                <a:r><a:t>Carbon &amp; Growth</a:t></a:r></a:p>
            </c:rich></c:tx></c:title>
        </c:chart>"#;
        let d = root_of(xml);
        let root = d.root_element();
        assert_eq!(
            extract_chart_title_text(root).as_deref(),
            Some("Carbon & Growth")
        );
        assert_eq!(extract_chart_title_size(root), Some(1400));
        assert_eq!(extract_chart_title_bold(root), Some(true));
        assert_eq!(extract_chart_title_italic(root), Some(false));
        assert_eq!(extract_chart_title_srgb(root).as_deref(), Some("1B4332"));
    }

    #[test]
    fn chart_title_rich_runs_preserve_newline_and_independent_typography() {
        let xml = format!(
            r#"<c:title xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:tx><c:rich>
              <a:p><a:r><a:rPr sz="1800" b="1"/><a:t>Long heading</a:t></a:r>
                <a:r><a:rPr sz="1400" i="1"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:rPr><a:t>
Subtitle</a:t></a:r></a:p>
            </c:rich></c:tx></c:title>"#
        );
        let document = root_of(&xml);
        let runs = parse_chart_title_rich_runs(document.root_element(), &StubResolver)
            .expect("rich title runs");
        assert_eq!(runs.len(), 2);
        assert_eq!(runs[0].text, "Long heading");
        assert_eq!(runs[0].font_size_hpt, Some(1800));
        assert_eq!(runs[0].bold, Some(true));
        assert_eq!(runs[0].italic, Some(false));
        assert_eq!(runs[1].text, "\nSubtitle");
        assert_eq!(runs[1].font_size_hpt, Some(1400));
        assert_eq!(runs[1].bold, Some(false));
        assert_eq!(runs[1].italic, Some(true));
        assert_eq!(runs[1].color.as_deref(), Some("112233"));
    }

    #[test]
    fn chart_title_character_properties_default_omitted_bold_and_italic_to_false() {
        let xml = format!(
            r#"<c:chart xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:title><c:tx><c:rich>
              <a:p><a:r><a:rPr sz="1400"/><a:t>Regular title</a:t></a:r></a:p>
            </c:rich></c:tx></c:title></c:chart>"#
        );
        let document = root_of(&xml);
        let root = document.root_element();
        assert_eq!(extract_chart_title_bold(root), Some(false));
        assert_eq!(extract_chart_title_italic(root), Some(false));
        let title = child(root, "title").unwrap();
        let runs = parse_chart_title_rich_runs(title, &StubResolver).unwrap();
        assert_eq!(runs[0].bold, Some(false));
        assert_eq!(runs[0].italic, Some(false));
    }

    #[test]
    fn chart_title_text_from_strref_cache() {
        // Title sourced from a strRef cache (`<c:v>`) rather than rich runs.
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:title><c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Sales</c:v></c:pt></c:strCache></c:strRef></c:tx></c:title>
        </c:chart>"#;
        let d = root_of(xml);
        assert_eq!(
            extract_chart_title_text(d.root_element()).as_deref(),
            Some("Sales")
        );
    }

    #[test]
    fn chart_title_text_preserves_an_authored_space() {
        let xml = r#"<c:catAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t> </a:t></a:r></a:p></c:rich></c:tx></c:title>
        </c:catAx>"#;
        let document = root_of(xml);
        assert_eq!(
            extract_chart_title_text(document.root_element()).as_deref(),
            Some(" ")
        );
    }

    #[test]
    fn chart_title_srgb_skips_non_solidfill_srgb() {
        // An `<a:srgbClr>` that is NOT a direct child of `<a:solidFill>` (here a
        // gradient stop) must be ignored.
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:title><c:tx><c:rich><a:p><a:r><a:rPr>
                <a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="ABCDEF"/></a:gs></a:gsLst></a:gradFill>
            </a:rPr><a:t>T</a:t></a:r></a:p></c:rich></c:tx></c:title>
        </c:chart>"#;
        let d = root_of(xml);
        assert!(extract_chart_title_srgb(d.root_element()).is_none());
    }

    #[test]
    fn chart_title_helpers_absent() {
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#;
        let d = root_of(xml);
        let root = d.root_element();
        assert!(extract_chart_title_text(root).is_none());
        assert!(extract_chart_title_size(root).is_none());
        assert!(extract_chart_title_bold(root).is_none());
        assert!(extract_chart_title_italic(root).is_none());
        assert!(extract_chart_title_srgb(root).is_none());
    }

    #[test]
    fn chart_title_color_resolves_scheme_and_srgb() {
        // schemeClr (`tx2`) — resolved via the resolver, unlike the srgb-only
        // `extract_chart_title_srgb` which returns None for a scheme color.
        let scheme = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:title><c:tx><c:rich><a:p><a:pPr>
                <a:defRPr><a:solidFill><a:schemeClr val="tx2"/></a:solidFill></a:defRPr>
            </a:pPr><a:r><a:rPr><a:solidFill><a:schemeClr val="tx2"/></a:solidFill></a:rPr><a:t>T</a:t></a:r></a:p></c:rich></c:tx></c:title>
        </c:chart>"#;
        let d = root_of(scheme);
        assert_eq!(
            extract_chart_title_color(d.root_element(), &StubResolver).as_deref(),
            Some("tx2")
        );
        assert!(extract_chart_title_srgb(d.root_element()).is_none());

        // srgbClr — resolved (uppercased by StubResolver) too.
        let srgb = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:title><c:tx><c:rich><a:p><a:r><a:rPr><a:solidFill><a:srgbClr val="1b4332"/></a:solidFill></a:rPr><a:t>T</a:t></a:r></a:p></c:rich></c:tx></c:title>
        </c:chart>"#;
        let d2 = root_of(srgb);
        assert_eq!(
            extract_chart_title_color(d2.root_element(), &StubResolver).as_deref(),
            Some("1B4332")
        );
    }

    #[test]
    fn chart_title_color_skips_title_frame_sppr_fill() {
        // A `<c:title><c:spPr><a:solidFill>` is the title FRAME fill, not the
        // text color; it must be ignored (only run-property fills count).
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:title>
                <c:tx><c:rich><a:p><a:r><a:t>T</a:t></a:r></a:p></c:rich></c:tx>
                <c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr>
            </c:title>
        </c:chart>"#;
        let d = root_of(xml);
        assert!(extract_chart_title_color(d.root_element(), &StubResolver).is_none());
    }

    #[test]
    fn data_label_face_scoped_to_dlbls() {
        let xml = format!(
            r#"<c:chart xmlns:c="{C_NS}" xmlns:a="{A_NS}">
                 <c:plotArea><c:barChart>
                   <c:dLbls><c:txPr><a:p><a:pPr><a:defRPr><a:latin typeface="Consolas"/></a:defRPr></a:pPr></a:p></c:txPr></c:dLbls>
                 </c:barChart></c:plotArea>
               </c:chart>"#
        );
        assert_eq!(
            extract_data_label_face(root_of(&xml).root_element()).as_deref(),
            Some("Consolas")
        );
    }

    #[test]
    fn legend_text_props_face_size_bold() {
        let xml = format!(
            r#"<c:chart xmlns:c="{C_NS}" xmlns:a="{A_NS}">
                 <c:legend><c:legendPos val="b"/>
                   <c:txPr><a:p><a:pPr><a:defRPr sz="1100" b="1"><a:latin typeface="Calibri"/></a:defRPr></a:pPr></a:p></c:txPr>
                 </c:legend>
               </c:chart>"#
        );
        let (face, size, bold, italic) = extract_legend_text_props(root_of(&xml).root_element());
        assert_eq!(face.as_deref(), Some("Calibri"));
        assert_eq!(size, Some(1100));
        assert_eq!(bold, Some(true));
        assert_eq!(italic, Some(false));
    }

    #[test]
    fn classic_series_data_labels_do_not_enable_unlabelled_sibling_series() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:chart><c:plotArea>
              <c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
                <c:ser><c:idx val="0"/><c:val><c:numLit><c:pt idx="0"><c:v>8</c:v></c:pt></c:numLit></c:val>
                  <c:dLbls><c:txPr><a:p><a:pPr><a:defRPr b="0"/></a:pPr></a:p></c:txPr><c:showVal val="1"/></c:dLbls>
                </c:ser>
                <c:ser><c:idx val="1"/><c:val><c:numLit><c:pt idx="0"><c:v>7</c:v></c:pt></c:numLit></c:val></c:ser>
              </c:barChart></c:plotArea></c:chart></c:chartSpace>"#,
        );
        let model = parse_chart_part(
            chart_space_of(&xml).root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("bar chart parses");

        assert!(
            !model.show_data_labels,
            "series-local dLbls is not chart-wide"
        );
        assert_eq!(
            model.series[0]
                .series_data_labels
                .as_ref()
                .map(|d| d.show_val),
            Some(true)
        );
        assert_eq!(
            model.series[0]
                .series_data_labels
                .as_ref()
                .and_then(|d| d.font_bold),
            Some(false)
        );
        assert!(model.series[1].series_data_labels.is_none());
    }

    #[test]
    fn classic_chart_group_label_style_does_not_inherit_the_first_series_style() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:chart><c:plotArea>
              <c:lineChart><c:grouping val="standard"/>
                <c:ser><c:idx val="0"/><c:val><c:numLit><c:pt idx="0"><c:v>8</c:v></c:pt></c:numLit></c:val>
                  <c:dLbls><c:numFmt formatCode="0.0"/><c:txPr><a:p><a:pPr><a:defRPr sz="800" b="0">
                    <a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:latin typeface="Series Face"/>
                  </a:defRPr></a:pPr></a:p></c:txPr><c:showVal val="1"/></c:dLbls>
                </c:ser>
                <c:dLbls><c:numFmt formatCode="0.00"/><c:txPr><a:p><a:pPr><a:defRPr sz="1200" b="1">
                  <a:solidFill><a:schemeClr val="accent2"/></a:solidFill><a:latin typeface="Group Face"/>
                </a:defRPr></a:pPr></a:p></c:txPr><c:showVal val="1"/></c:dLbls>
              </c:lineChart></c:plotArea></c:chart></c:chartSpace>"#,
        );
        let model = parse_chart_part(
            chart_space_of(&xml).root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("line chart parses");

        assert!(model.show_data_labels);
        assert_eq!(model.data_label_format_code.as_deref(), Some("0.00"));
        assert_eq!(model.data_label_font_size_hpt, Some(1200));
        assert_eq!(model.data_label_font_bold, Some(true));
        assert_eq!(model.data_label_font_color.as_deref(), Some("ED7D31"));
        assert_eq!(model.data_label_font_face.as_deref(), Some("Group Face"));
        let local = model.series[0]
            .series_data_labels
            .as_ref()
            .expect("series-local label style is preserved");
        assert_eq!(local.format_code.as_deref(), Some("0.0"));
        assert_eq!(local.font_size_hpt, Some(800));
        assert_eq!(local.font_bold, Some(false));
        assert_eq!(local.font_color.as_deref(), Some("4472C4"));
        assert_eq!(local.font_face.as_deref(), Some("Series Face"));
    }

    #[test]
    fn classic_chart_group_labels_merge_property_wise_into_owned_series() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:chart><c:plotArea>
              <c:barChart><c:barDir val="col"/>
                <c:ser><c:idx val="0"/><c:val><c:numLit><c:pt idx="0"><c:v>8</c:v></c:pt></c:numLit></c:val>
                  <c:dLbls><c:dLbl><c:idx val="0"/><c:spPr><a:ln><a:solidFill><a:srgbClr val="778899"/></a:solidFill></a:ln></c:spPr><c:showSerName/></c:dLbl>
                    <c:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:ln></c:spPr>
                    <c:txPr><a:p><a:pPr><a:defRPr b="1"><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>
                    <c:showVal val="0"/><c:separator/>
                  </c:dLbls>
                </c:ser>
                <c:ser><c:idx val="1"/><c:val><c:numLit><c:pt idx="0"><c:v>7</c:v></c:pt></c:numLit></c:val></c:ser>
                <c:dLbls><c:dLblPos val="inEnd"/><c:showVal/><c:showCatName/><c:separator>|</c:separator></c:dLbls>
              </c:barChart></c:plotArea></c:chart></c:chartSpace>"#,
        );
        let model = parse_chart_part(
            chart_space_of(&xml).root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("bar chart parses");

        let first = model.series[0]
            .series_data_labels
            .as_ref()
            .expect("merged defaults");
        assert!(!first.show_val && first.show_cat_name);
        assert_eq!(first.font_color.as_deref(), Some("4472C4"));
        assert_eq!(first.font_bold, Some(true));
        assert_eq!(first.position.as_deref(), Some("inEnd"));
        assert_eq!(first.separator.as_deref(), Some(""));
        let first_box = first.label_box.as_ref().expect("merged series box");
        assert_eq!(first_box.fill, None);
        assert_eq!(first_box.border_color.as_deref(), Some("445566"));
        assert_eq!(
            first_box
                .style
                .as_ref()
                .and_then(|style| style.shape_properties_present),
            Some(true)
        );

        let point = &model.series[0].data_label_overrides.as_ref().unwrap()[0];
        assert_eq!(point.text, "");
        assert_eq!(point.show_ser_name, Some(true));
        let point_box = point.label_box.as_ref().expect("merged point box");
        assert_eq!(point_box.fill, None);
        assert_eq!(point_box.border_color.as_deref(), Some("778899"));
        assert_eq!(
            point_box
                .style
                .as_ref()
                .and_then(|style| style.shape_properties_present),
            Some(true)
        );

        let second = model.series[1]
            .series_data_labels
            .as_ref()
            .expect("group defaults");
        assert!(second.show_val);
        assert!(second.show_cat_name);
        assert_eq!(second.position.as_deref(), Some("inEnd"));
        assert_eq!(second.separator.as_deref(), Some("|"));
        assert_eq!(second.font_color, None);
    }

    #[test]
    fn classic_group_labels_do_not_leak_into_an_unlabelled_combo_group() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea>
              <c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:val><c:numLit><c:pt idx="0"><c:v>8</c:v></c:pt></c:numLit></c:val></c:ser><c:dLbls><c:showVal/></c:dLbls></c:barChart>
              <c:lineChart><c:ser><c:idx val="1"/><c:val><c:numLit><c:pt idx="0"><c:v>7</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart>
            </c:plotArea></c:chart></c:chartSpace>"#,
        );
        let model = parse_chart_part(
            chart_space_of(&xml).root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("combo chart parses");

        let bar = model.series[0]
            .series_data_labels
            .as_ref()
            .expect("bar group defaults");
        assert!(bar.show_val);
        let line = model.series[1]
            .series_data_labels
            .as_ref()
            .expect("effective absent group defaults");
        assert!(!line.show_val);
        assert!(!line.show_cat_name);
        assert!(!line.show_ser_name);
        assert!(!line.show_percent);
    }

    #[test]
    fn classic_point_delete_false_overrides_a_deleted_series_label_collection() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea>
              <c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:val><c:numLit><c:pt idx="0"><c:v>8</c:v></c:pt></c:numLit></c:val>
                <c:dLbls><c:dLbl><c:idx val="0"/><c:delete val="0"/></c:dLbl><c:delete/></c:dLbls>
              </c:ser><c:dLbls><c:showVal/></c:dLbls></c:barChart>
            </c:plotArea></c:chart></c:chartSpace>"#,
        );
        let model = parse_chart_part(
            chart_space_of(&xml).root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("bar chart parses");

        let labels = model.series[0]
            .series_data_labels
            .as_ref()
            .expect("deleted collection retained");
        assert_eq!(labels.deleted, Some(true));
        assert!(labels.show_val, "group showVal remains the lower default");
        let point = &model.series[0].data_label_overrides.as_ref().unwrap()[0];
        assert_eq!(point.deleted, Some(false));
    }

    #[test]
    fn classic_legend_frame_fill_and_outline_are_preserved() {
        struct TransformAwareLegendResolver;
        impl ColorResolver for TransformAwareLegendResolver {
            fn resolve_solid_fill(&self, node: Node) -> Option<String> {
                FixtureResolver.resolve_solid_fill(node)
            }

            fn resolve_shape_fill(&self, parent: Node) -> Option<String> {
                match parent.tag_name().name() {
                    "spPr" => Some("DDEEFF".to_string()),
                    "ln" => Some("808080".to_string()),
                    _ => None,
                }
            }
        }
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:chart>
              <c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/>
                <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
              </c:ser></c:barChart></c:plotArea>
              <c:legend><c:legendPos val="r"/><c:spPr>
                <a:solidFill><a:schemeClr val="accent1"><a:lumMod val="20000"/><a:lumOff val="80000"/></a:schemeClr></a:solidFill>
                <a:ln w="3175" cap="sq" cmpd="dbl"><a:solidFill><a:srgbClr val="808080"/></a:solidFill><a:custDash><a:ds d="200000" sp="50000"/></a:custDash><a:round/></a:ln>
              </c:spPr></c:legend>
            </c:chart></c:chartSpace>"#,
        );
        let model = parse_chart_part(
            chart_space_of(&xml).root_element(),
            &ChartParseContext {
                color_resolver: Some(&TransformAwareLegendResolver),
                ..Default::default()
            },
        )
        .expect("legend chart parses");
        let serialized = serde_json::to_value(model).expect("chart serializes");

        assert_eq!(serialized["legendFillColor"], "DDEEFF");
        assert_eq!(serialized["legendFill"]["fillType"], "solid");
        assert_eq!(serialized["legendFill"]["color"], "DDEEFF");
        assert_eq!(serialized["legendFillPaintAuthored"], true);
        assert!(serialized.get("legendFillHidden").is_none());
        assert_eq!(serialized["legendLineColor"], "808080");
        assert_eq!(serialized["legendLineWidthEmu"], 3175);
        assert!(serialized.get("legendLineDash").is_none());
        assert_eq!(serialized["legendLineCustomDash"][0]["dash"], 2.0);
        assert_eq!(serialized["legendLineCustomDash"][0]["space"], 0.5);
        assert_eq!(serialized["legendLineCap"], "sq");
        assert_eq!(serialized["legendLineJoin"], "round");
        assert_eq!(serialized["legendLineCompound"], "dbl");
        assert_eq!(serialized["legendLinePaintAuthored"], true);
        assert!(serialized.get("legendLineHidden").is_none());
    }

    #[test]
    fn classic_legend_frame_no_fill_provenance_is_preserved() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:chart>
              <c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/>
                <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
              </c:ser></c:barChart></c:plotArea>
              <c:legend><c:legendPos val="r"/><c:spPr>
                <a:noFill/><a:ln><a:noFill/></a:ln>
              </c:spPr></c:legend>
            </c:chart></c:chartSpace>"#,
        );
        let model = parse_chart_part(
            chart_space_of(&xml).root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("legend chart parses");
        assert_eq!(model.legend_fill_hidden, Some(true));
        assert_eq!(model.legend_fill_paint_authored, Some(true));
        assert_eq!(model.legend_line_hidden, Some(true));
        assert_eq!(model.legend_line_paint_authored, Some(true));
    }

    #[test]
    fn trendline_rich_run_paint_does_not_become_the_label_default() {
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:trendline>
              <c:trendlineType val="linear"/><c:trendlineLbl><c:tx><c:rich>
                <a:p><a:pPr><a:defRPr><a:noFill/></a:defRPr></a:pPr><a:r><a:t>A</a:t></a:r></a:p>
                <a:p><a:r><a:t>B</a:t></a:r></a:p>
              </c:rich></c:tx></c:trendlineLbl>
            </c:trendline></c:ser>"#
        );
        let trendline = extract_series_trendlines(root_of(&xml).root_element(), &StubResolver)
            .unwrap()
            .remove(0);
        assert_eq!(trendline.label_font_color, None);
        assert_eq!(trendline.label_font_paint_authored, None);
        assert_eq!(trendline.label_font_hidden, None);
        let runs = trendline.label_rich_runs.expect("rich runs");
        assert_eq!(runs[0].color_paint_authored, Some(true));
        assert_eq!(runs[0].color_hidden, Some(true));
        assert_eq!(runs[1].text, "\n");
        assert_eq!(runs[2].color_paint_authored, None);
        assert_eq!(runs[2].color_hidden, None);
    }

    #[test]
    fn dark_classic_text_uses_office_contrast_with_the_observed_title_carrier_gate() {
        let parse = |style: u8, paragraph_default_run: &str| {
            let xml = format!(
                r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
                  <c:style val="{style}"/><c:chart><c:title><c:tx><c:rich>
                    <a:bodyPr/><a:lstStyle/><a:p>{paragraph_default_run}<a:r><a:t>Title</a:t></a:r></a:p>
                  </c:rich></c:tx></c:title><c:plotArea><c:barChart>
                    <c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/>
                      <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                      <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
                    </c:ser>
                  </c:barChart></c:plotArea><c:legend/></c:chart></c:chartSpace>"#,
            );
            let document = chart_space_of(&xml);
            parse_chart_part(
                document.root_element(),
                &ChartParseContext {
                    color_resolver: Some(&WordContrastFixtureResolver),
                    ..Default::default()
                },
            )
            .expect("classic dark chart parses")
        };

        let without_default_run = parse(41, "");
        let roles = without_default_run.classic_chart_style_roles.unwrap();
        assert_eq!(roles["categoryAxis"].font_color.as_deref(), Some("FFFFFF"));
        assert_eq!(roles["legend"].font_color.as_deref(), Some("FFFFFF"));
        assert_eq!(roles["title"].font_color.as_deref(), Some("000000"));

        let with_default_run = parse(41, "<a:pPr><a:defRPr/></a:pPr>");
        let roles = with_default_run.classic_chart_style_roles.unwrap();
        assert_eq!(roles["title"].font_color.as_deref(), Some("FFFFFF"));

        let light_style = parse(40, "<a:pPr><a:defRPr/></a:pPr>");
        let roles = light_style.classic_chart_style_roles.unwrap();
        assert_eq!(roles["title"].font_color.as_deref(), Some("000000"));
        assert_eq!(roles["legend"].font_color.as_deref(), Some("000000"));
    }

    #[test]
    fn dark_classic_title_carrier_gate_uses_only_textual_rich_paragraphs() {
        let parse = |title_body: &str| {
            let xml = format!(
                r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
                  <c:style val="41"/><c:chart><c:title>{title_body}</c:title>
                  <c:plotArea><c:barChart><c:barDir val="col"/>
                    <c:ser><c:idx val="0"/><c:order val="0"/>
                      <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                      <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
                    </c:ser>
                  </c:barChart></c:plotArea></c:chart></c:chartSpace>"#,
            );
            let document = chart_space_of(&xml);
            parse_chart_part(
                document.root_element(),
                &ChartParseContext {
                    color_resolver: Some(&WordContrastFixtureResolver),
                    ..Default::default()
                },
            )
            .expect("classic dark chart parses")
            .classic_chart_style_roles
            .expect("numeric roles")
        };

        let sibling_tx_pr = parse(
            r#"<c:tx><c:rich><a:p><a:r><a:t>Title</a:t></a:r></a:p></c:rich></c:tx>
               <c:txPr><a:p><a:pPr><a:defRPr/></a:pPr></a:p></c:txPr>"#,
        );
        assert_eq!(sibling_tx_pr["title"].font_color.as_deref(), Some("000000"));

        let empty_formatted_paragraph = parse(
            r#"<c:tx><c:rich>
                 <a:p><a:pPr><a:defRPr/></a:pPr></a:p>
                 <a:p><a:r><a:t>Title</a:t></a:r></a:p>
               </c:rich></c:tx>"#,
        );
        assert_eq!(
            empty_formatted_paragraph["title"].font_color.as_deref(),
            Some("000000")
        );

        let run_properties_only = parse(
            r#"<c:tx><c:rich><a:p><a:r><a:rPr lang="en" sz="1200" b="1"/>
                 <a:t>Title</a:t></a:r></a:p></c:rich></c:tx>"#,
        );
        assert_eq!(
            run_properties_only["title"].font_color.as_deref(),
            Some("000000")
        );

        let multiple_paragraphs = parse(
            r#"<c:tx><c:rich>
                 <a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:t>One</a:t></a:r></a:p>
                 <a:p><a:pPr><a:defRPr/></a:pPr><a:fld><a:t>Two</a:t></a:fld></a:p>
               </c:rich></c:tx>"#,
        );
        assert_eq!(
            multiple_paragraphs["title"].font_color.as_deref(),
            Some("000000")
        );

        let mixed_multiple_paragraphs = parse(
            r#"<c:tx><c:rich>
                 <a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:t>One</a:t></a:r></a:p>
                 <a:p><a:r><a:t>Two</a:t></a:r></a:p>
               </c:rich></c:tx>"#,
        );
        assert_eq!(
            mixed_multiple_paragraphs["title"].font_color.as_deref(),
            Some("000000")
        );
    }

    #[test]
    fn ct_boolean_auto_title_deleted_bare_suppresses_fallback_title() {
        // §21.2.2.7 `<c:autoTitleDeleted/>` ⇒ true ⇒ the single-series name is
        // NOT promoted to a fallback chart title. A bare element must read true;
        // its absence (control) leaves the existing empty title frame eligible
        // for the observed series-name fallback.
        let with_bare = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:chart>
                <c:title><c:txPr><a:bodyPr/><a:lstStyle/><a:p/></c:txPr></c:title>
                <c:autoTitleDeleted/>
                <c:plotArea><c:lineChart>
                  <c:ser><c:idx val="0"/><c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>OnlySeries</c:v></c:pt></c:strCache></c:strRef></c:tx>
                    <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>
                </c:lineChart></c:plotArea>
              </c:chart></c:chartSpace>"#
        );
        let d = chart_space_of(&with_bare);
        let m = parse_chart_part(
            d.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("chart");
        assert_eq!(
            m.title, None,
            "bare <c:autoTitleDeleted/> ⇒ no fallback title"
        );

        let without = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:chart>
                <c:title><c:txPr><a:bodyPr/><a:lstStyle/><a:p/></c:txPr></c:title>
                <c:plotArea><c:lineChart>
                  <c:ser><c:idx val="0"/><c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>OnlySeries</c:v></c:pt></c:strCache></c:strRef></c:tx>
                    <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>
                </c:lineChart></c:plotArea>
              </c:chart></c:chartSpace>"#
        );
        let d2 = chart_space_of(&without);
        let m2 = parse_chart_part(
            d2.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("chart");
        assert_eq!(
            m2.title.as_deref(),
            Some("OnlySeries"),
            "no autoTitleDeleted ⇒ series name is the fallback title"
        );
    }

    #[test]
    fn parse_chart_part_preserves_custom_data_label_run_styles() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:chart><c:plotArea>
              <c:lineChart><c:grouping val="standard"/><c:ser>
                <c:idx val="0"/><c:tx><c:v>Employer</c:v></c:tx>
                <c:dLbls><c:dLbl><c:idx val="0"/><c:tx><c:rich><a:p>
                  <a:pPr><a:defRPr sz="800" b="1"/></a:pPr>
                  <a:r><a:rPr sz="1200" b="1"/><a:t>Employer</a:t></a:r>
                  <a:r><a:rPr sz="1100" b="0"/><a:t> 36.0%</a:t></a:r>
                </a:p></c:rich></c:tx><c:dLblPos val="l"/>
                <c:showVal val="1"/><c:showSerName val="1"/>
                </c:dLbl></c:dLbls>
                <c:cat><c:strCache><c:pt idx="0"><c:v>Start</c:v></c:pt></c:strCache></c:cat>
                <c:val><c:numCache><c:pt idx="0"><c:v>36</c:v></c:pt></c:numCache></c:val>
              </c:ser></c:lineChart>
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
        .expect("line chart parses");
        let label = &model.series[0].data_label_overrides.as_ref().unwrap()[0];
        assert_eq!(label.text, "Employer 36.0%");
        assert_eq!(label.position.as_deref(), Some("l"));
        let runs = label.rich_runs.as_ref().expect("rich runs on public model");
        assert_eq!(runs[0].font_size_hpt, Some(1200));
        assert_eq!(runs[0].bold, Some(true));
        assert_eq!(runs[1].font_size_hpt, Some(1100));
        assert_eq!(runs[1].bold, Some(false));
    }

    /// A chart title may be a string reference cache rather than DrawingML
    /// rich text (§21.2.2.210 CT_Title → §21.2.2.214 CT_Tx). The cached `<c:v>`
    /// is the authored title and must win over the single-series auto title.
    #[test]
    fn parse_chart_part_uses_strref_cache_for_title() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:chart>
                <c:title><c:tx><c:strRef><c:f>Sheet1!$A$1</c:f><c:strCache>
                  <c:ptCount val="1"/><c:pt idx="0"><c:v>Authored cached title</c:v></c:pt>
                </c:strCache></c:strRef></c:tx></c:title>
                <c:autoTitleDeleted val="0"/>
                <c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
                  <c:ser><c:idx val="0"/>
                    <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Series fallback</c:v></c:pt></c:strCache></c:strRef></c:tx>
                    <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                    <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
                  </c:ser>
                </c:barChart></c:plotArea>
              </c:chart>
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
        .expect("chart parses");
        assert_eq!(model.title.as_deref(), Some("Authored cached title"));
    }

    #[test]
    fn display_unit_label_preserves_authored_unresolved_text_paint() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:chart><c:plotArea>
              <c:barChart><c:ser><c:idx val="0"/><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart>
              <c:valAx><c:dispUnits><c:builtInUnit val="thousands"/><c:dispUnitsLbl>
                <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr><a:noFill/></a:defRPr></a:pPr></a:p></c:txPr>
              </c:dispUnitsLbl></c:dispUnits></c:valAx>
            </c:plotArea></c:chart></c:chartSpace>"#
        );
        let doc = chart_space_of(&xml);
        let model = parse_chart_part(
            doc.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("chart parses");
        let label = model.val_axis_display_units.unwrap().label.unwrap();
        assert_eq!(label.font_color, None);
        assert_eq!(label.font_paint_authored, Some(true));
        assert_eq!(label.font_hidden, Some(true));
    }

    #[test]
    fn parse_series_data_labels_defaults_and_per_point_override() {
        let cache = std::collections::HashMap::new();
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:dLbls>
                <c:numFmt formatCode="0.0%"/>
                <c:separator>&#10;</c:separator>
                <c:dLbl>
                  <c:idx val="1"/>
                  <c:tx><c:rich><a:p><a:r><a:t>Custom</a:t></a:r></a:p></c:rich></c:tx>
                  <c:dLblPos val="outEnd"/>
                  <c:layout><c:manualLayout>
                    <c:xMode val="edge"/><c:yMode val="edge"/>
                    <c:x val="0.25"/><c:y val="0.4"/>
                    <c:w val="0.2"/><c:h val="0.1"/>
                  </c:manualLayout></c:layout>
                </c:dLbl>
                <c:showVal val="1"/>
                <c:showCatName val="0"/>
                <c:showSerName val="0"/>
                <c:showPercent val="1"/>
                <c:dLblPos val="ctr"/>
                <c:txPr><a:bodyPr rot="1800000" wrap="none" anchor="b" vert="vert" lIns="12700" tIns="25400" rIns="38100" bIns="50800"/><a:p><a:pPr algn="ctr"><a:defRPr i="1" lang="ja-JP" baseline="25000"/></a:pPr></a:p></c:txPr>
              </c:dLbls>
            </c:ser>"#
        );
        let d = root_of(&xml);
        let (defaults, overrides) =
            parse_series_data_labels(d.root_element(), &FixtureResolver, &cache);
        let defaults = defaults.expect("series-level dLbls present");
        assert!(defaults.show_val);
        assert!(!defaults.show_cat_name);
        assert!(!defaults.show_ser_name);
        assert!(defaults.show_percent);
        assert_eq!(defaults.position.as_deref(), Some("ctr"));
        assert_eq!(defaults.format_code.as_deref(), Some("0.0%"));
        assert_eq!(defaults.separator.as_deref(), Some("\n"));
        assert_eq!(defaults.font_italic, Some(true));
        assert_eq!(defaults.font_language.as_deref(), Some("ja-JP"));
        assert_eq!(defaults.font_baseline, Some(0.25));
        assert_eq!(defaults.text_rotation, Some(1_800_000));
        assert_eq!(defaults.text_wrap.as_deref(), Some("none"));
        assert_eq!(defaults.text_vertical_anchor.as_deref(), Some("b"));
        assert_eq!(defaults.text_vertical_mode.as_deref(), Some("vert"));
        assert_eq!(defaults.text_l_ins_emu, Some(12_700));
        assert_eq!(defaults.text_t_ins_emu, Some(25_400));
        assert_eq!(defaults.text_r_ins_emu, Some(38_100));
        assert_eq!(defaults.text_b_ins_emu, Some(50_800));
        assert_eq!(defaults.text_align.as_deref(), Some("ctr"));

        assert_eq!(overrides.len(), 1);
        let o = &overrides[0];
        assert_eq!(o.idx, 1);
        assert_eq!(o.text, "Custom");
        assert_eq!(o.position.as_deref(), Some("outEnd"));
        assert_eq!(
            o.manual_layout,
            Some(ChartManualLayout {
                x_mode: "edge".to_string(),
                y_mode: "edge".to_string(),
                w_mode: "factor".to_string(),
                h_mode: "factor".to_string(),
                layout_target: Some("outer".to_string()),
                x: 0.25,
                y: 0.4,
                w: Some(0.2),
                h: Some(0.1),
            })
        );
    }

    #[test]
    fn parse_series_data_labels_rich_run_style_overrides_txpr_defaults() {
        let cache = std::collections::HashMap::new();
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:dLbls>
                <c:dLbl>
                  <c:idx val="13"/>
                  <c:tx><c:rich><a:bodyPr rot="-2700000" lIns="1pt"/><a:p><a:pPr algn="r"><a:defRPr sz="800" b="1"><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:defRPr></a:pPr>
                    <a:r><a:rPr sz="1200" i="1" lang="en-US" baseline="-12.5%"><a:solidFill><a:srgbClr val="EC008B"/></a:solidFill></a:rPr><a:t>Employer</a:t></a:r>
                    <a:r><a:rPr sz="1100" b="0"/><a:t> 36.0%</a:t></a:r>
                  </a:p></c:rich></c:tx>
                  <c:txPr><a:bodyPr anchor="b" rIns="-12700"/><a:p><a:pPr><a:defRPr sz="900" b="0"><a:solidFill><a:srgbClr val="333333"/></a:solidFill><a:latin typeface="Tx Face"/></a:defRPr></a:pPr></a:p></c:txPr>
                </c:dLbl>
              </c:dLbls>
            </c:ser>"#
        );
        let d = root_of(&xml);
        let (_, overrides) = parse_series_data_labels(d.root_element(), &FixtureResolver, &cache);
        let label = &overrides[0];
        assert_eq!(label.text, "Employer 36.0%");
        assert_eq!(label.font_color.as_deref(), Some("333333"));
        assert_eq!(label.font_size_hpt, Some(900));
        assert_eq!(label.font_bold, Some(false));
        assert_eq!(label.font_italic, Some(false));
        assert_eq!(label.font_language, None);
        assert_eq!(label.font_baseline, None);
        assert_eq!(label.text_rotation, Some(-2_700_000));
        assert_eq!(label.text_l_ins_emu, Some(12_700));
        assert_eq!(label.text_vertical_anchor.as_deref(), Some("b"));
        assert_eq!(label.text_r_ins_emu, Some(-12_700));
        assert_eq!(label.text_body_authored, Some(true));
        assert_eq!(label.text_align, None);
        let runs = label.rich_runs.as_ref().expect("rich runs");
        assert_eq!(runs.len(), 2);
        assert_eq!(runs[0].text, "Employer");
        assert_eq!(runs[0].font_size_hpt, Some(1200));
        assert_eq!(runs[0].bold, Some(true));
        assert_eq!(runs[0].color.as_deref(), Some("EC008B"));
        assert_eq!(runs[0].font_face.as_deref(), Some("Tx Face"));
        assert_eq!(runs[0].italic, Some(true));
        assert_eq!(runs[0].language.as_deref(), Some("en-US"));
        assert_eq!(runs[0].baseline, Some(-0.125));
        assert_eq!(runs[0].paragraph_align.as_deref(), Some("r"));
        assert_eq!(runs[1].text, " 36.0%");
        assert_eq!(runs[1].font_size_hpt, Some(1100));
        assert_eq!(runs[1].bold, Some(false));
        assert_eq!(runs[1].color.as_deref(), Some("445566"));
        assert_eq!(runs[1].font_face.as_deref(), Some("Tx Face"));
        assert_eq!(runs[1].paragraph_align.as_deref(), Some("r"));
    }

    #[test]
    fn data_label_text_fill_provenance_blocks_lower_color_fallbacks() {
        let cache = std::collections::HashMap::new();
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:dLbls>
              <c:txPr><a:p><a:pPr><a:defRPr><a:noFill/></a:defRPr></a:pPr></a:p></c:txPr>
              <c:dLbl><c:idx val="0"/><c:tx><c:rich><a:p><a:r>
                <a:rPr><a:blipFill/></a:rPr><a:t>Unresolved</a:t>
              </a:r></a:p></c:rich></c:tx></c:dLbl>
            </c:dLbls></c:ser>"#
        );
        let document = root_of(&xml);
        let (defaults, overrides) =
            parse_series_data_labels(document.root_element(), &FixtureResolver, &cache);
        let defaults = defaults.expect("series defaults");
        assert_eq!(defaults.font_paint_authored, Some(true));
        assert_eq!(defaults.font_hidden, Some(true));
        assert_eq!(defaults.font_color, None);
        let label = &overrides[0];
        assert_eq!(label.font_paint_authored, None);
        assert_eq!(label.font_hidden, None);
        assert_eq!(label.font_color, None);
        let run = &label.rich_runs.as_ref().expect("rich run")[0];
        assert_eq!(run.color_paint_authored, Some(true));
        assert_eq!(run.color_hidden, None);
        assert_eq!(run.color, None);
    }

    #[test]
    fn data_label_run_paint_does_not_become_the_point_default() {
        let cache = std::collections::HashMap::new();
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:dLbls>
              <c:txPr><a:p><a:pPr><a:defRPr><a:solidFill><a:srgbClr val="008000"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>
              <c:dLbl><c:idx val="0"/><c:tx><c:rich>
                <a:p><a:pPr><a:defRPr><a:noFill/></a:defRPr></a:pPr><a:r><a:t>A</a:t></a:r></a:p>
                <a:p><a:r><a:t>B</a:t></a:r></a:p>
              </c:rich></c:tx></c:dLbl>
            </c:dLbls></c:ser>"#
        );
        let document = root_of(&xml);
        let (defaults, overrides) =
            parse_series_data_labels(document.root_element(), &FixtureResolver, &cache);
        assert_eq!(defaults.unwrap().font_color.as_deref(), Some("008000"));
        let label = &overrides[0];
        assert_eq!(label.font_color, None);
        assert_eq!(label.font_paint_authored, None);
        let runs = label.rich_runs.as_ref().expect("rich runs");
        assert_eq!(runs[0].color_paint_authored, Some(true));
        assert_eq!(runs[0].color_hidden, Some(true));
        assert_eq!(runs[1].text, "\n");
        assert_eq!(runs[2].color_paint_authored, None);
        assert_eq!(runs[2].color_hidden, None);
    }

    #[test]
    fn data_label_rich_break_is_retained_with_paragraph_alignment() {
        let cache = std::collections::HashMap::new();
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:dLbls><c:dLbl>
              <c:idx val="0"/><c:tx><c:rich><a:p><a:pPr algn="r"/>
                <a:r><a:t>A</a:t></a:r><a:br/><a:r><a:t>B</a:t></a:r>
              </a:p></c:rich></c:tx>
            </c:dLbl></c:dLbls></c:ser>"#
        );
        let document = root_of(&xml);
        let (_, overrides) =
            parse_series_data_labels(document.root_element(), &FixtureResolver, &cache);
        let label = &overrides[0];
        assert_eq!(label.text, "A\nB");
        let runs = label.rich_runs.as_ref().expect("rich runs");
        assert_eq!(
            runs.iter().map(|run| run.text.as_str()).collect::<Vec<_>>(),
            vec!["A", "\n", "B"]
        );
        assert!(runs
            .iter()
            .all(|run| run.paragraph_align.as_deref() == Some("r")));
    }

    #[test]
    fn parse_series_data_labels_bounds_rich_runs_by_scalars_and_paragraphs() {
        let cache = std::collections::HashMap::new();
        let paragraph_text = "x".repeat(1100);
        let paragraphs = (0..5)
            .map(|_| format!("<a:p><a:r><a:t>{paragraph_text}</a:t></a:r></a:p>"))
            .collect::<String>();
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:dLbls><c:dLbl>
              <c:idx val="0"/><c:tx><c:rich>{paragraphs}</c:rich></c:tx>
            </c:dLbl></c:dLbls></c:ser>"#
        );
        let document = root_of(&xml);
        let (_, overrides) =
            parse_series_data_labels(document.root_element(), &FixtureResolver, &cache);
        let label = &overrides[0];
        let runs = label.rich_runs.as_ref().expect("bounded rich runs");
        let text = runs.iter().map(|run| run.text.as_str()).collect::<String>();
        assert_eq!(text, label.text);
        assert!(text.chars().count() <= MAX_DATA_LABEL_RICH_SCALARS);
        assert!(text.split('\n').count() <= MAX_DATA_LABEL_RICH_LINES);
    }

    #[test]
    fn parse_data_label_rich_run_size_obeys_st_text_font_size_and_cascades() {
        let cache = std::collections::HashMap::new();
        for (direct, fallback, expected) in [
            ("99", "1200", 1200),
            ("100", "1200", 100),
            ("400000", "1200", 400_000),
            ("400001", "1200", 1200),
        ] {
            let xml = format!(
                r#"<c:ser xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:dLbls><c:dLbl>
                  <c:idx val="0"/><c:tx><c:rich><a:p><a:pPr><a:defRPr sz="{fallback}"/></a:pPr>
                    <a:r><a:rPr sz="{direct}"/><a:t>Label</a:t></a:r>
                  </a:p></c:rich></c:tx>
                </c:dLbl></c:dLbls></c:ser>"#
            );
            let document = root_of(&xml);
            let (_, overrides) =
                parse_series_data_labels(document.root_element(), &FixtureResolver, &cache);
            assert_eq!(
                overrides[0].rich_runs.as_ref().unwrap()[0].font_size_hpt,
                Some(expected),
                "direct size {direct}",
            );
        }
    }

    #[test]
    fn parse_series_data_labels_callout_box_and_leader_lines() {
        // A Word-style pie-callout series `<c:dLbls>` carries a `<c:spPr>` box
        // (white fill + coloured border), a per-point
        // `<c:dLbl>` with its own box, and `<c:showLeaderLines>` +
        // `<c:leaderLines>` style. All must round-trip into the model.
        let cache = std::collections::HashMap::new();
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:dLbls>
                <c:dLbl>
                  <c:idx val="0"/>
                  <c:spPr>
                    <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
                    <a:ln w="12700"><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></a:ln>
                  </c:spPr>
                  <c:showCatName val="1"/>
                  <c:showPercent val="1"/>
                </c:dLbl>
                <c:spPr>
                  <a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FEFEFE"/></a:gs><a:gs pos="100000"><a:srgbClr val="DDEEFF"/></a:gs></a:gsLst><a:lin ang="2700000"/></a:gradFill>
                  <a:ln w="12700"><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill><a:prstDash val="dash"/><a:round/></a:ln>
                </c:spPr>
                <c:showVal val="0"/>
                <c:showCatName val="1"/>
                <c:showPercent val="1"/>
                <c:showLeaderLines val="1"/>
                <c:leaderLines>
                  <c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="A6A6A6"/></a:solidFill></a:ln></c:spPr>
                </c:leaderLines>
              </c:dLbls>
            </c:ser>"#
        );
        let d = root_of(&xml);
        let (defaults, overrides) =
            parse_series_data_labels(d.root_element(), &FixtureResolver, &cache);
        let defaults = defaults.expect("series-level dLbls present");
        let box_ = defaults.label_box.expect("series callout box");
        assert!(box_.fill.is_none());
        assert!(matches!(
            box_.fill_paint,
            Some(ChartStyleFill::Gradient { .. })
        ));
        assert_eq!(box_.border_color.as_deref(), Some("4472C4"));
        assert_eq!(box_.border_width_emu, Some(12700));
        assert_eq!(box_.border_dash.as_deref(), Some("dash"));
        assert_eq!(box_.border_join.as_deref(), Some("round"));
        assert!(defaults.show_leader_lines);
        assert_eq!(defaults.leader_line_color.as_deref(), Some("A6A6A6"));
        assert_eq!(defaults.leader_line_width_emu, Some(9525));

        assert_eq!(overrides.len(), 1);
        let o = &overrides[0];
        assert_eq!(o.idx, 0);
        let obox = o.label_box.as_ref().expect("per-point callout box");
        assert_eq!(obox.fill.as_deref(), Some("FFFFFF"));
        assert_eq!(obox.border_color.as_deref(), Some("4472C4"));
        assert_eq!(obox.border_width_emu, Some(12700));
    }

    #[test]
    fn data_label_shape_paints_exceeding_the_recipe_ceiling_fail_closed_atomically() {
        let stops = (0..=MAX_CHART_LABEL_GRADIENT_STOPS)
            .map(|index| format!(r#"<a:gs pos="{}"><a:srgbClr val="112233"/></a:gs>"#, index))
            .collect::<String>();
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:dLbls>
                <c:spPr><a:solidFill><a:srgbClr val="445566"/></a:solidFill></c:spPr>
                <c:dLbl><c:idx val="0"/><c:spPr><a:gradFill><a:gsLst>{stops}</a:gsLst></a:gradFill></c:spPr></c:dLbl>
                <c:showVal val="1"/>
              </c:dLbls>
            </c:ser>"#
        );
        let document = root_of(&xml);
        let cache = std::collections::HashMap::new();
        let (defaults, overrides) =
            parse_series_data_labels(document.root_element(), &FixtureResolver, &cache);
        let series_box = defaults
            .and_then(|labels| labels.label_box)
            .expect("authored series label shape");
        assert_eq!(series_box.fill_paint_authored, Some(true));
        assert!(series_box.fill.is_none() && series_box.fill_paint.is_none());
        let point_box = overrides[0]
            .label_box
            .as_ref()
            .expect("authored point label shape");
        assert_eq!(point_box.fill_paint_authored, Some(true));
        assert!(point_box.fill.is_none() && point_box.fill_paint.is_none());
    }

    #[test]
    fn data_label_fill_and_outline_apply_the_recipe_ceiling_independently() {
        let stops = (0..2049)
            .map(|index| format!(r#"<a:gs pos="{}"><a:srgbClr val="112233"/></a:gs>"#, index))
            .collect::<String>();
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:dLbls>
              <c:spPr>
                <a:gradFill><a:gsLst>{stops}</a:gsLst></a:gradFill>
                <a:ln><a:gradFill><a:gsLst>{stops}</a:gsLst></a:gradFill></a:ln>
              </c:spPr><c:showVal val="1"/>
            </c:dLbls></c:ser>"#
        );
        let document = root_of(&xml);
        let cache = std::collections::HashMap::new();
        let (defaults, _) =
            parse_series_data_labels(document.root_element(), &FixtureResolver, &cache);
        let box_ = defaults
            .and_then(|labels| labels.label_box)
            .expect("label shape");
        assert!(matches!(
            box_.fill_paint,
            Some(ChartStyleFill::Gradient { .. })
        ));
        assert!(matches!(
            box_.border_fill,
            Some(ChartStyleFill::Gradient { .. })
        ));
    }

    #[test]
    fn parse_series_data_labels_no_box_leaves_callout_fields_unset() {
        // A plain `<c:dLbls>` with no `<c:spPr>` / leader lines must NOT
        // synthesize a callout box (keeps the historical plain-label path).
        let cache = std::collections::HashMap::new();
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}">
              <c:dLbls><c:showPercent val="1"/></c:dLbls>
            </c:ser>"#
        );
        let d = root_of(&xml);
        let (defaults, _) = parse_series_data_labels(d.root_element(), &FixtureResolver, &cache);
        let defaults = defaults.expect("series-level dLbls present");
        assert!(defaults.label_box.is_none());
        assert!(!defaults.show_leader_lines);
        assert!(defaults.leader_line_color.is_none());
    }

    #[test]
    fn parse_series_data_labels_deleted_point_has_empty_text() {
        let cache = std::collections::HashMap::new();
        let xml = format!(
            r#"<c:ser xmlns:c="{C_NS}">
              <c:dLbls>
                <c:dLbl><c:idx val="0"/><c:delete val="1"/></c:dLbl>
              </c:dLbls>
            </c:ser>"#
        );
        let d = root_of(&xml);
        let (_, overrides) = parse_series_data_labels(d.root_element(), &FixtureResolver, &cache);
        assert_eq!(overrides.len(), 1);
        assert_eq!(overrides[0].text, "");
    }

    #[test]
    fn parse_series_data_labels_absent_returns_none_and_empty() {
        let xml = format!(r#"<c:ser xmlns:c="{C_NS}"></c:ser>"#);
        let d = root_of(&xml);
        let cache = std::collections::HashMap::new();
        let (defaults, overrides) =
            parse_series_data_labels(d.root_element(), &FixtureResolver, &cache);
        assert!(defaults.is_none());
        assert!(overrides.is_empty());
    }

    /// Observed Office compatibility: an empty title frame with
    /// `autoTitleDeleted` absent (⇒ not suppressed) and EXACTLY ONE named series
    /// adopts that series' name as the chart title. ECMA-376 §21.2.2.7 defines
    /// only the suppression flag; Word-produced output provides the text rule.
    #[test]
    fn parse_chart_part_auto_title_single_series() {
        let xml = chart_space_auto_title(None, &named_ser(0, "Production in 2017"));
        let d = chart_space_of(&xml);
        let m = parse_chart_part(
            d.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("parses");
        // The series name is promoted VERBATIM (the `cap="all"` uppercase is a
        // rendering-layer transform we do not apply at parse time).
        assert_eq!(m.title.as_deref(), Some("Production in 2017"));

        // An explicit `autoTitleDeleted val="0"` behaves identically (0 ⇒ auto
        // title may be shown).
        let xml0 = chart_space_auto_title(Some("0"), &named_ser(0, "Production in 2017"));
        let d0 = chart_space_of(&xml0);
        let m0 = parse_chart_part(
            d0.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("parses");
        assert_eq!(m0.title.as_deref(), Some("Production in 2017"));
    }

    /// `autoTitleDeleted` controls whether an automatic title is suppressed; it
    /// does not create a title frame when `<c:title>` itself is absent. A named
    /// single-series chart without that element therefore remains untitled.
    #[test]
    fn parse_chart_part_without_title_element_stays_untitled() {
        let series = r#"<c:ser><c:idx val="0"/><c:order val="0"/>
          <c:tx><c:v>A</c:v></c:tx>
          <c:cat><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:strLit></c:cat>
          <c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="1"/>
            <c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
        </c:ser>"#;
        for auto_title_deleted in [None, Some("0")] {
            let atd = auto_title_deleted
                .map(|value| format!(r#"<c:autoTitleDeleted val="{value}"/>"#))
                .unwrap_or_default();
            let xml = format!(
                r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
                  <c:chart>
                    {atd}
                    <c:plotArea><c:pieChart><c:varyColors val="1"/>{series}</c:pieChart></c:plotArea>
                  </c:chart>
                </c:chartSpace>"#
            );
            let document = chart_space_of(&xml);
            let model = parse_chart_part(
                document.root_element(),
                &ChartParseContext {
                    color_resolver: Some(&FixtureResolver),
                    ..Default::default()
                },
            )
            .expect("pie chart parses");
            assert_eq!(
                model.title, None,
                "an absent <c:title> must not synthesize a title when autoTitleDeleted={auto_title_deleted:?}"
            );
        }
    }

    /// §21.2.2.7 `autoTitleDeleted val="1"` (or `"true"`) suppresses the auto
    /// title even for a single named series — Word shows no title.
    #[test]
    fn parse_chart_part_auto_title_deleted_shows_no_title() {
        for v in ["1", "true"] {
            let xml = chart_space_auto_title(Some(v), &named_ser(0, "Production in 2017"));
            let d = chart_space_of(&xml);
            let m = parse_chart_part(
                d.root_element(),
                &ChartParseContext {
                    color_resolver: Some(&FixtureResolver),
                    ..Default::default()
                },
            )
            .expect("parses");
            assert_eq!(m.title, None, "autoTitleDeleted={v} should suppress title");
        }
    }

    /// The observed empty-frame fallback is bounded to single-series charts.
    /// With TWO series, Word shows no synthesized title (a lone series name
    /// would be misleading), so `title` stays `None`.
    #[test]
    fn parse_chart_part_auto_title_multi_series_none() {
        let sers = format!(
            "{}{}",
            named_ser(0, "Series One"),
            named_ser(1, "Series Two")
        );
        let xml = chart_space_auto_title(None, &sers);
        let d = chart_space_of(&xml);
        let m = parse_chart_part(
            d.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("parses");
        assert_eq!(m.series.len(), 2);
        assert_eq!(m.title, None);
    }

    #[test]
    fn parse_chart_part_keeps_bubble_x_source_kind_for_legend_semantics() {
        let series = |x_source: &str| {
            format!(
                r#"<c:ser><c:idx val="0"/><c:order val="0"/>
                  <c:xVal>{x_source}</c:xVal>
                  <c:yVal><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:yVal>
                  <c:bubbleSize><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>3</c:v></c:pt></c:numLit></c:bubbleSize>
                </c:ser>"#
            )
        };
        let text_xml = chart_space_with_group(&format!(
            "<c:bubbleChart>{}</c:bubbleChart>",
            series(
                r#"<c:strRef><c:f>Labels</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Project A</c:v></c:pt></c:strCache></c:strRef>"#
            )
        ));
        let text_doc = chart_space_of(&text_xml);
        let text_chart = parse_chart_part(
            text_doc.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("text-x bubble parses");
        assert_eq!(text_chart.series[0].bubble_x_source_is_string, Some(true));

        let numeric_xml = chart_space_with_group(&format!(
            "<c:bubbleChart>{}</c:bubbleChart>",
            series(r#"<c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit>"#)
        ));
        let numeric_doc = chart_space_of(&numeric_xml);
        let numeric_chart = parse_chart_part(
            numeric_doc.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("numeric-x bubble parses");
        assert_eq!(
            numeric_chart.series[0].bubble_x_source_is_string,
            Some(false)
        );
    }
}
