#[cfg(test)]
mod tests {
    use super::super::*;

    #[test]
    fn legend_manual_layout_preserves_omitted_width_and_height() {
        let document = root_of(
            r#"<c:legend xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:layout><c:manualLayout><c:x val="0.1"/><c:y val="0.2"/></c:manualLayout></c:layout></c:legend>"#,
        );
        let layout = extract_legend_manual_layout(document.root_element()).expect("manual layout");
        assert_eq!(layout.x, 0.1);
        assert_eq!(layout.y, 0.2);
        assert_eq!(layout.w, None);
        assert_eq!(layout.h, None);
    }

    #[test]
    fn axis_delete_truthy_variants() {
        for (val, expect) in [
            ("1", true),
            ("0", false),
            ("true", true),
            ("false", false),
            ("True", true),
        ] {
            let xml = format!(
                r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:delete val="{val}"/>
            </c:valAx>"#
            );
            let d = root_of(&xml);
            assert_eq!(axis_is_deleted(d.root_element()), expect, "val={val}");
        }
    }

    #[test]
    fn axis_delete_default_false() {
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#;
        let d = root_of(xml);
        assert!(!axis_is_deleted(d.root_element()));
    }

    #[test]
    fn axis_min_max() {
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:scaling><c:max val="2500"/><c:min val="0"/></c:scaling>
        </c:valAx>"#;
        let d = root_of(xml);
        assert_eq!(
            extract_axis_min_max(d.root_element()),
            (Some(0.0), Some(2500.0))
        );
    }

    #[test]
    fn axis_format_code_skips_general() {
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:numFmt formatCode="General"/>
        </c:valAx>"#;
        let d = root_of(xml);
        assert!(extract_axis_format_code(d.root_element()).is_none());
    }

    #[test]
    fn axis_format_code_passes_through() {
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:numFmt formatCode="0.0%"/>
        </c:valAx>"#;
        let d = root_of(xml);
        assert_eq!(
            extract_axis_format_code(d.root_element()).as_deref(),
            Some("0.0%")
        );
    }

    #[test]
    fn axis_tick_label_color_from_txpr() {
        let xml = r#"<c:catAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:spPr><a:ln><a:solidFill><a:srgbClr val="d9d9d9"/></a:solidFill></a:ln></c:spPr>
            <c:txPr><a:p><a:pPr><a:defRPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>
        </c:catAx>"#;
        let d = root_of(xml);
        // The txPr text fill (bg1) is returned — the spPr line fill must not leak.
        let got = extract_axis_tick_label_color(d.root_element(), &StubResolver);
        assert_eq!(got.as_deref(), Some("bg1"));
    }

    #[test]
    fn axis_tick_label_color_absent() {
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#;
        let d = root_of(xml);
        assert!(extract_axis_tick_label_color(d.root_element(), &StubResolver).is_none());
    }

    #[test]
    fn axis_line_style_solid_with_width() {
        let xml = r#"<c:catAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:spPr><a:noFill/><a:ln w="9525"><a:solidFill><a:srgbClr val="d9d9d9"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr>
        </c:catAx>"#;
        let d = root_of(xml);
        let (color, width, no_fill) = extract_axis_line_style(d.root_element(), &StubResolver);
        assert_eq!(color.as_deref(), Some("D9D9D9"));
        assert_eq!(width, Some(9525));
        assert!(!no_fill);
        assert_eq!(
            extract_axis_line_dash(d.root_element()).as_deref(),
            Some("dash")
        );
    }

    #[test]
    fn axis_line_style_nofill_line() {
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:spPr><a:ln w="9525"><a:noFill/></a:ln></c:spPr>
        </c:valAx>"#;
        let d = root_of(xml);
        let (color, width, no_fill) = extract_axis_line_style(d.root_element(), &StubResolver);
        assert!(color.is_none());
        assert_eq!(width, Some(9525));
        assert!(no_fill);
    }

    #[test]
    fn axis_line_style_absent() {
        let xml = r#"<c:catAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#;
        let d = root_of(xml);
        assert_eq!(
            extract_axis_line_style(d.root_element(), &StubResolver),
            (None, None, false)
        );
    }

    #[test]
    fn gridline_style_solid_scheme_with_width() {
        // `<c:majorGridlines><c:spPr><a:ln w="3175">
        // <a:solidFill><a:schemeClr val="accent3"/>` → the explicit gridline
        // colour + 0.25 pt width (3175 EMU) the renderer must honor.
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:majorGridlines><c:spPr><a:ln w="3175"><a:solidFill><a:schemeClr val="accent3"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>
        </c:valAx>"#;
        let d = root_of(xml);
        let (color, width, dash) = extract_gridline_style(d.root_element(), &StubResolver);
        assert_eq!(color.as_deref(), Some("accent3"));
        assert_eq!(width, Some(3175));
        assert_eq!(dash, None);
    }

    #[test]
    fn gridline_style_present_without_sppr() {
        // `<c:majorGridlines/>` with no `<c:spPr>` → gridlines are requested
        // (presence-only) but carry no explicit colour/width; the renderer keeps
        // its faint default. `(None, None, None)`.
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:majorGridlines/>
        </c:valAx>"#;
        let d = root_of(xml);
        assert_eq!(
            extract_gridline_style(d.root_element(), &StubResolver),
            (None, None, None)
        );
    }

    #[test]
    fn gridline_style_absent() {
        // No `<c:majorGridlines>` at all → `(None, None, None)`.
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#;
        let d = root_of(xml);
        assert_eq!(
            extract_gridline_style(d.root_element(), &StubResolver),
            (None, None, None)
        );
    }

    #[test]
    fn gridline_style_preserves_drawingml_dash() {
        let xml = format!(
            r#"<c:valAx xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:majorGridlines><c:spPr><a:ln><a:prstDash val="dash"/></a:ln></c:spPr></c:majorGridlines></c:valAx>"#
        );
        let (_, _, dash) = extract_gridline_style(root_of(&xml).root_element(), &StubResolver);
        assert_eq!(dash.as_deref(), Some("dash"));
    }

    #[test]
    fn minor_gridline_style_is_independent_from_major_gridlines() {
        let xml = format!(
            r#"<c:valAx xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:majorGridlines/><c:minorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:prstDash val="dot"/></a:ln></c:spPr></c:minorGridlines></c:valAx>"#
        );
        let (color, width, dash) =
            extract_minor_gridline_style(root_of(&xml).root_element(), &StubResolver);
        assert_eq!(color.as_deref(), Some("112233"));
        assert_eq!(width, Some(6350));
        assert_eq!(dash.as_deref(), Some("dot"));
    }

    #[test]
    fn classic_tick_mark_distinguishes_bare_element_from_omission() {
        let xml =
            format!(r#"<c:valAx xmlns:c="{C_NS}"><c:majorTickMark/><c:minorTickMark/></c:valAx>"#,);
        let document = root_of(&xml);
        let axis = document.root_element();
        assert_eq!(
            extract_axis_tick_mark_or_default(axis, "majorTickMark"),
            "cross"
        );
        assert_eq!(
            extract_axis_tick_mark(axis, "minorTickMark").as_deref(),
            Some("cross")
        );

        let omitted_xml = format!(r#"<c:valAx xmlns:c="{C_NS}"/>"#);
        let omitted = root_of(&omitted_xml);
        let axis = omitted.root_element();
        assert_eq!(
            extract_axis_tick_mark_or_default(axis, "majorTickMark"),
            "out"
        );
        assert_eq!(extract_axis_tick_mark(axis, "minorTickMark"), None);
    }

    #[test]
    fn axis_title_with_props_resolved_scheme_color() {
        // The resolver-based axis-title variant resolves a schemeClr color.
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:axPos val="l"/>
            <c:title><c:tx><c:rich><a:p><a:r><a:rPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:rPr><a:t>Value</a:t></a:r></a:p></c:rich></c:tx></c:title>
        </c:valAx>"#;
        let d = root_of(xml);
        let (text, _sz, _b, color) =
            extract_axis_title_with_props_resolved(d.root_element(), &StubResolver);
        assert_eq!(text.as_deref(), Some("Value"));
        assert_eq!(color.as_deref(), Some("accent1"));
    }

    #[test]
    fn axis_title_inherits_the_axis_text_color_when_its_runs_omit_color() {
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:axPos val="l"/>
            <c:title><c:tx><c:rich><a:p><a:r><a:t>Values</a:t></a:r></a:p></c:rich></c:tx></c:title>
            <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>
        </c:valAx>"#;
        let d = root_of(xml);
        let (_text, _size, _bold, color) =
            extract_axis_title_with_props_resolved(d.root_element(), &StubResolver);
        assert_eq!(color.as_deref(), Some("000000"));
    }

    #[test]
    fn axis_title_with_props_full() {
        let xml = r#"<c:catAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:axPos val="b"/>
            <c:title><c:tx><c:rich>
                <a:p><a:pPr><a:defRPr sz="1000" b="1"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:defRPr></a:pPr>
                <a:r><a:t>Category Axis</a:t></a:r></a:p>
            </c:rich></c:tx></c:title>
        </c:catAx>"#;
        let d = root_of(xml);
        let (text, size, bold, color) = extract_axis_title_with_props(d.root_element());
        assert_eq!(text.as_deref(), Some("Category Axis"));
        assert_eq!(size, Some(1000));
        assert_eq!(bold, Some(true));
        assert_eq!(color.as_deref(), Some("FF0000"));
    }

    #[test]
    fn axis_title_run_properties_override_defaults_property_by_property() {
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:title><c:tx><c:rich><a:bodyPr/><a:p>
              <a:pPr><a:defRPr sz="900" b="0"><a:solidFill><a:srgbClr val="111111"/></a:solidFill><a:latin typeface="Default Face"/></a:defRPr></a:pPr>
              <a:r><a:rPr sz="1200" b="1"><a:solidFill><a:srgbClr val="222222"/></a:solidFill></a:rPr><a:t>Value</a:t></a:r>
            </a:p></c:rich></c:tx></c:title>
        </c:valAx>"#;
        let d = root_of(xml);
        let axis = d.root_element();
        let (_, size, bold, color) = extract_axis_title_with_props(axis);
        assert_eq!(size, Some(1200));
        assert_eq!(bold, Some(true));
        assert_eq!(color.as_deref(), Some("222222"));
        // The run does not author a face, so that property alone cascades to
        // defRPr while the other run-authored properties still win.
        assert_eq!(
            extract_axis_title_face(axis).as_deref(),
            Some("Default Face")
        );
    }

    #[test]
    fn axis_title_size_rejects_values_outside_st_text_font_size() {
        for size in ["NaN", "-1", "0", "99", "400001", "2147483647"] {
            let xml = format!(
                r#"<c:valAx xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:rPr sz="{size}"/><a:t>Value</a:t></a:r></a:p></c:rich></c:tx></c:title></c:valAx>"#
            );
            let d = root_of(&xml);
            assert_eq!(extract_axis_title_size(d.root_element()), None);
        }
        for size in ["100", "400000"] {
            let xml = format!(
                r#"<c:valAx xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:rPr sz="{size}"/><a:t>Value</a:t></a:r></a:p></c:rich></c:tx></c:title></c:valAx>"#
            );
            let d = root_of(&xml);
            assert_eq!(
                extract_axis_title_size(d.root_element()),
                size.parse::<i32>().ok()
            );
        }
    }

    #[test]
    fn axis_title_rotation_and_vertical_mode_are_independent() {
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:title>
              <c:tx><c:rich><a:bodyPr vert="vert270"/><a:p><a:r><a:t>Value</a:t></a:r></a:p></c:rich></c:tx>
              <c:txPr><a:bodyPr rot="1800000"/><a:p/></c:txPr>
            </c:title>
        </c:valAx>"#;
        let d = root_of(xml);
        assert_eq!(
            extract_axis_title_rotation(d.root_element()),
            Some(1_800_000)
        );
        assert_eq!(
            extract_axis_title_vertical_mode(d.root_element()).as_deref(),
            Some("vert270")
        );

        let txpr_only = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:title><c:txPr><a:bodyPr vert="eaVert"/><a:p/></c:txPr></c:title>
        </c:valAx>"#;
        let txpr_doc = root_of(txpr_only);
        assert_eq!(extract_axis_title_rotation(txpr_doc.root_element()), None);
        assert_eq!(
            extract_axis_title_vertical_mode(txpr_doc.root_element()).as_deref(),
            Some("eaVert")
        );
    }

    #[test]
    fn axis_title_vertical_insets_use_text_body_defaults_and_property_cascade() {
        let defaults_xml = r#"<c:catAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t> </a:t></a:r></a:p></c:rich></c:tx></c:title>
        </c:catAx>"#;
        let defaults_doc = root_of(defaults_xml);
        assert_eq!(
            extract_axis_title_vertical_inset(defaults_doc.root_element()),
            Some(91_440)
        );

        let cascade_xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:title>
              <c:tx><c:rich><a:bodyPr tIns="3pt"/><a:p><a:r><a:t>Value</a:t></a:r></a:p></c:rich></c:tx>
              <c:txPr><a:bodyPr tIns="1pt" bIns="2pt"/><a:p/></c:txPr>
            </c:title>
        </c:valAx>"#;
        let cascade_doc = root_of(cascade_xml);
        assert_eq!(
            extract_axis_title_vertical_inset(cascade_doc.root_element()),
            Some(63_500)
        );
    }

    #[test]
    fn axis_title_with_props_text_absent_all_none() {
        // Axis with no `<c:title>` → text None gates the props to None even
        // though run props could in theory be read elsewhere.
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:axPos val="l"/>
        </c:valAx>"#;
        let d = root_of(xml);
        assert_eq!(
            extract_axis_title_with_props(d.root_element()),
            (None, None, None, None)
        );
    }

    #[test]
    fn axis_tick_label_bold_variants() {
        for (b, expect) in [("1", Some(true)), ("0", Some(false)), ("true", Some(true))] {
            let xml = format!(
                r#"<c:catAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                    <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr b="{b}"/></a:pPr><a:endParaRPr/></a:p></c:txPr>
                </c:catAx>"#
            );
            let d = root_of(&xml);
            assert_eq!(
                extract_axis_tick_label_bold(d.root_element()),
                expect,
                "b={b}"
            );
        }
    }

    #[test]
    fn axis_tick_label_bold_absent() {
        let xml = r#"<c:catAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#;
        let d = root_of(xml);
        assert!(extract_axis_tick_label_bold(d.root_element()).is_none());
    }

    #[test]
    fn axis_tick_and_title_faces() {
        // Tick face lives in the axis `<c:txPr>`; the title face in `<c:title>`.
        // Extractors must NOT cross-contaminate.
        let xml = format!(
            r#"<c:valAx xmlns:c="{C_NS}" xmlns:a="{A_NS}">
                 <c:title><a:p><a:r><a:rPr><a:latin typeface="Georgia"/></a:rPr><a:t>Y</a:t></a:r></a:p></c:title>
                 <c:txPr><a:p><a:pPr><a:defRPr><a:latin typeface="Verdana"/></a:defRPr></a:pPr></a:p></c:txPr>
               </c:valAx>"#
        );
        let root = root_of(&xml);
        let ax = root.root_element();
        assert_eq!(extract_axis_tick_label_face(ax).as_deref(), Some("Verdana"));
        assert_eq!(extract_axis_title_face(ax).as_deref(), Some("Georgia"));
    }

    #[test]
    fn axis_gridlines_presence() {
        // Value axis with `<c:majorGridlines>` → true; category axis without → false.
        let val = format!(r#"<c:valAx xmlns:c="{C_NS}"><c:majorGridlines/></c:valAx>"#);
        assert!(axis_has_major_gridlines(root_of(&val).root_element()));
        assert!(!axis_has_minor_gridlines(root_of(&val).root_element()));

        let cat = format!(r#"<c:catAx xmlns:c="{C_NS}"/>"#);
        assert!(!axis_has_major_gridlines(root_of(&cat).root_element()));

        let both = format!(
            r#"<c:valAx xmlns:c="{C_NS}"><c:majorGridlines/><c:minorGridlines/></c:valAx>"#
        );
        assert!(axis_has_major_gridlines(root_of(&both).root_element()));
        assert!(axis_has_minor_gridlines(root_of(&both).root_element()));
    }

    #[test]
    fn minor_gridline_no_fill_suppresses_the_effective_stroke() {
        let visible = format!(
            r#"<c:valAx xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:minorGridlines><c:spPr><a:ln><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:ln></c:spPr></c:minorGridlines></c:valAx>"#
        );
        let hidden = format!(
            r#"<c:valAx xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:minorGridlines><c:spPr><a:ln><a:noFill/></a:ln></c:spPr></c:minorGridlines></c:valAx>"#
        );
        assert!(axis_minor_gridlines_visible(
            root_of(&visible).root_element()
        ));
        assert!(!axis_minor_gridlines_visible(
            root_of(&hidden).root_element()
        ));
        // Presence remains independently queryable for callers that need the
        // authored request rather than its effective paint visibility.
        assert!(axis_has_minor_gridlines(root_of(&hidden).root_element()));
    }

    #[test]
    fn axis_major_minor_unit() {
        let xml = format!(
            r#"<c:valAx xmlns:c="{C_NS}"><c:crossBetween val="between"/><c:majorUnit val="500"/><c:minorUnit val="100"/></c:valAx>"#
        );
        assert_eq!(
            extract_axis_major_unit(root_of(&xml).root_element()),
            Some(500.0)
        );
        assert_eq!(
            extract_axis_minor_unit(root_of(&xml).root_element()),
            Some(100.0)
        );
        // Absent → None (auto step).
        let bare = format!(r#"<c:valAx xmlns:c="{C_NS}"/>"#);
        assert_eq!(extract_axis_major_unit(root_of(&bare).root_element()), None);
        // Non-positive rejected (would wedge the gridline loop).
        let zero = format!(r#"<c:valAx xmlns:c="{C_NS}"><c:majorUnit val="0"/></c:valAx>"#);
        assert_eq!(extract_axis_major_unit(root_of(&zero).root_element()), None);
    }

    #[test]
    fn axis_log_base() {
        let xml = format!(
            r#"<c:valAx xmlns:c="{C_NS}"><c:scaling><c:logBase val="10"/></c:scaling></c:valAx>"#
        );
        assert_eq!(
            extract_axis_log_base(root_of(&xml).root_element()),
            Some(10.0)
        );
        // Base < 2 is invalid per ST_LogBase → rejected.
        let bad = format!(
            r#"<c:valAx xmlns:c="{C_NS}"><c:scaling><c:logBase val="1"/></c:scaling></c:valAx>"#
        );
        assert_eq!(extract_axis_log_base(root_of(&bad).root_element()), None);
        // Absent scaling / logBase → None (linear).
        let bare = format!(r#"<c:valAx xmlns:c="{C_NS}"><c:scaling/></c:valAx>"#);
        assert_eq!(extract_axis_log_base(root_of(&bare).root_element()), None);
    }

    #[test]
    fn axis_orientation() {
        let rev = format!(
            r#"<c:valAx xmlns:c="{C_NS}"><c:scaling><c:orientation val="maxMin"/></c:scaling></c:valAx>"#
        );
        assert_eq!(
            extract_axis_orientation(root_of(&rev).root_element()).as_deref(),
            Some("maxMin")
        );
        let norm = format!(
            r#"<c:valAx xmlns:c="{C_NS}"><c:scaling><c:orientation val="minMax"/></c:scaling></c:valAx>"#
        );
        assert_eq!(
            extract_axis_orientation(root_of(&norm).root_element()).as_deref(),
            Some("minMax")
        );
        // Absent → None (renderer treats as minMax).
        let bare = format!(r#"<c:valAx xmlns:c="{C_NS}"><c:scaling/></c:valAx>"#);
        assert_eq!(
            extract_axis_orientation(root_of(&bare).root_element()),
            None
        );
    }

    #[test]
    fn axis_tick_label_pos_and_rotation() {
        let xml = format!(
            r#"<c:catAx xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:tickLblPos val="low"/><c:txPr><a:bodyPr rot="-2700000"/></c:txPr></c:catAx>"#
        );
        assert_eq!(
            extract_axis_tick_label_pos(root_of(&xml).root_element()).as_deref(),
            Some("low")
        );
        assert_eq!(
            extract_axis_tick_label_rotation(root_of(&xml).root_element()),
            Some(-2_700_000)
        );
        // Absent → None (renderer treats as nextTo / 0°).
        let bare = format!(r#"<c:catAx xmlns:c="{C_NS}"/>"#);
        assert_eq!(
            extract_axis_tick_label_pos(root_of(&bare).root_element()),
            None
        );
        assert_eq!(
            extract_axis_tick_label_rotation(root_of(&bare).root_element()),
            None
        );
    }

    #[test]
    fn parse_chart_part_preserves_primary_axis_display_units_and_labels() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:chart><c:plotArea>
              <c:scatterChart><c:ser><c:idx val="0"/>
                <c:xVal><c:numLit><c:pt idx="0"><c:v>1000</c:v></c:pt></c:numLit></c:xVal>
                <c:yVal><c:numLit><c:pt idx="0"><c:v>20</c:v></c:pt></c:numLit></c:yVal>
              </c:ser></c:scatterChart>
              <c:valAx><c:axPos val="b"/><c:dispUnits><c:builtInUnit val="thousands"/>
                <c:dispUnitsLbl><c:layout><c:manualLayout><c:x val="0.7"/><c:y val="0.8"/></c:manualLayout></c:layout>
                  <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="800" b="1"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr>
                </c:dispUnitsLbl></c:dispUnits></c:valAx>
              <c:valAx><c:axPos val="l"/><c:dispUnits><c:custUnit val="10"/>
                <c:dispUnitsLbl><c:tx><c:rich><a:p><a:r><a:t>per ten</a:t></a:r></a:p></c:rich></c:tx>
                  <c:txPr><a:bodyPr rot="-5400000"/><a:p><a:pPr><a:defRPr sz="900"/></a:pPr></a:p></c:txPr>
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
        .expect("scatter parses");
        let x = model.cat_axis_display_units.expect("horizontal units");
        assert_eq!(x.divisor, 1_000.0);
        assert_eq!(x.built_in_unit.as_deref(), Some("thousands"));
        let x_label = x.label.expect("horizontal unit label");
        assert_eq!(
            x_label.manual_layout.as_ref().map(|layout| layout.x),
            Some(0.7)
        );
        assert_eq!(x_label.font_size_hpt, Some(800));
        assert_eq!(x_label.font_bold, Some(true));
        assert_eq!(x_label.font_color.as_deref(), Some("112233"));
        assert_eq!(x_label.font_paint_authored, Some(true));
        assert_eq!(x_label.font_hidden, None);
        assert_eq!(x_label.font_face.as_deref(), Some("Arial"));

        let y = model.val_axis_display_units.expect("vertical units");
        assert_eq!(y.divisor, 10.0);
        assert_eq!(y.built_in_unit, None);
        let y_label = y.label.expect("vertical unit label");
        assert_eq!(y_label.text.as_deref(), Some("per ten"));
        assert_eq!(y_label.rotation, Some(-5_400_000));
        assert_eq!(y_label.font_size_hpt, Some(900));
    }

    /// A chart without `<c:style>` is valid (`CT_ChartSpace.style` is optional).
    /// The parser retains absent direct axis formatting while the numeric role
    /// table carries the effective built-in style-2 defaults. Keeping those
    /// layers separate is essential: a literal black/0.75pt parser fallback
    /// would incorrectly outrank a non-black theme and its subtle line width.
    #[test]
    fn parse_chart_part_resolves_styleless_legacy_axis_defaults() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:chart>
                <c:plotArea>
                  <c:barChart>
                    <c:barDir val="col"/><c:grouping val="clustered"/>
                    <c:ser><c:idx val="0"/>
                      <c:cat><c:strLit><c:pt idx="0"><c:v>T1</c:v></c:pt></c:strLit></c:cat>
                      <c:val><c:numLit><c:pt idx="0"><c:v>10</c:v></c:pt></c:numLit></c:val>
                    </c:ser>
                  </c:barChart>
                  <c:catAx><c:delete val="0"/><c:majorTickMark val="out"/></c:catAx>
                  <c:valAx><c:delete val="0"/><c:majorGridlines/><c:majorTickMark val="out"/></c:valAx>
                </c:plotArea>
              </c:chart>
              <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="1800"/></a:pPr></a:p></c:txPr>
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

        assert_eq!(m.cat_axis_font_size_hpt, Some(1800));
        assert_eq!(m.val_axis_font_size_hpt, Some(1800));
        assert_eq!(m.cat_axis_font_color, None);
        assert_eq!(m.val_axis_font_color, None);
        assert_eq!(m.cat_axis_line_color, None);
        assert_eq!(m.val_axis_line_color, None);
        assert_eq!(m.cat_axis_line_width_emu, None);
        assert_eq!(m.val_axis_line_width_emu, None);
        assert_eq!(m.cat_axis_line_paint_authored, None);
        assert_eq!(m.val_axis_line_paint_authored, None);
        assert_eq!(m.val_axis_gridline_color, None);
        assert_eq!(m.val_axis_gridline_width_emu, None);

        let roles = m
            .classic_chart_style_roles
            .as_ref()
            .expect("omitted style materializes style-2 roles");
        for role in ["categoryAxis", "valueAxis", "gridlineMajor"] {
            let line = &roles[role];
            // A genuinely absent optional theme uses the Office application
            // default matrix; present-but-broken theme data remains fail-closed.
            assert_eq!(
                line.line_colors.as_deref(),
                Some(&[Some("000000".to_owned())][..])
            );
            assert_eq!(line.line_width_emu, Some(9_525));
            assert_eq!(line.line_paint_authored, Some(true));
        }
        let title = &roles["title"];
        assert_eq!(title.font_size_hpt, None);
        assert_eq!(title.font_bold, Some(false));
        assert_eq!(title.font_paint_authored, None);

        let no_fill_xml = xml.replace(
            "<c:majorGridlines/>",
            "<c:majorGridlines><c:spPr><a:ln><a:noFill/></a:ln></c:spPr></c:majorGridlines>",
        );
        let no_fill_doc = chart_space_of(&no_fill_xml);
        let no_fill_model = parse_chart_part(
            no_fill_doc.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("no-fill chart parses");
        assert_eq!(no_fill_model.val_axis_major_gridlines, Some(false));
        assert_eq!(no_fill_model.val_axis_gridline_color, None);
        assert_eq!(no_fill_model.val_axis_gridline_width_emu, None);

        let no_fill_axis_xml = xml.replace(
            "<c:valAx><c:delete val=\"0\"/><c:majorGridlines/><c:majorTickMark val=\"out\"/></c:valAx>",
            "<c:valAx><c:delete val=\"0\"/><c:majorGridlines/><c:majorTickMark val=\"out\"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr></c:valAx>",
        );
        let no_fill_axis_doc = chart_space_of(&no_fill_axis_xml);
        let no_fill_axis_model = parse_chart_part(
            no_fill_axis_doc.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("no-fill axis chart parses");
        assert!(no_fill_axis_model.val_axis_line_hidden);
        assert_eq!(no_fill_axis_model.val_axis_major_tick_mark, "out");
        assert!(!no_fill_axis_model.val_axis_hidden);
    }

    #[test]
    fn style_two_axis_overlay_inherits_first_theme_line_width() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:style val="2"/>
              <c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser>
                <c:idx val="0"/><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
              </c:ser></c:barChart>
              <c:catAx><c:spPr><a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></c:spPr></c:catAx>
              <c:valAx><c:spPr><a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></c:spPr></c:valAx>
              </c:plotArea></c:chart></c:chartSpace>"#,
        );
        let theme = format!(
            r#"<a:theme xmlns:a="{A_NS}"><a:themeElements><a:fmtScheme name="Office">
              <a:fillStyleLst/><a:lnStyleLst><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
              <a:effectStyleLst/><a:bgFillStyleLst/>
            </a:fmtScheme></a:themeElements></a:theme>"#,
        );
        let resolver = FormatSchemeFixtureResolver {
            format_scheme: crate::theme::ThemeFormatScheme::parse(&theme),
        };
        let doc = chart_space_of(&xml);
        let model = parse_chart_part(
            doc.root_element(),
            &ChartParseContext {
                color_resolver: Some(&resolver),
                ..Default::default()
            },
        )
        .expect("chart parses");
        assert_eq!(model.cat_axis_line_width_emu, Some(12700));
        assert_eq!(model.val_axis_line_width_emu, Some(12700));

        let other_style = xml.replace("<c:style val=\"2\"/>", "<c:style val=\"10\"/>");
        let other_doc = chart_space_of(&other_style);
        let other = parse_chart_part(
            other_doc.root_element(),
            &ChartParseContext {
                color_resolver: Some(&resolver),
                ..Default::default()
            },
        )
        .expect("chart parses");
        assert_eq!(other.cat_axis_line_width_emu, None);
        assert_eq!(other.val_axis_line_width_emu, None);

        let absent_overlay = xml.replace(
            "<c:catAx><c:spPr><a:ln><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill></a:ln></c:spPr></c:catAx>",
            "<c:catAx/>",
        );
        let absent_doc = chart_space_of(&absent_overlay);
        let absent = parse_chart_part(
            absent_doc.root_element(),
            &ChartParseContext {
                color_resolver: Some(&resolver),
                ..Default::default()
            },
        )
        .expect("chart parses");
        assert_eq!(absent.cat_axis_line_width_emu, None);

        let explicit_width = xml.replacen("<a:ln>", "<a:ln w=\"9525\">", 1);
        let explicit_doc = chart_space_of(&explicit_width);
        let explicit = parse_chart_part(
            explicit_doc.root_element(),
            &ChartParseContext {
                color_resolver: Some(&resolver),
                ..Default::default()
            },
        )
        .expect("chart parses");
        assert_eq!(explicit.cat_axis_line_width_emu, Some(9525));
        assert_eq!(explicit.val_axis_line_width_emu, Some(12700));

        let no_fill = xml.replace(
            "<a:ln><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill></a:ln>",
            "<a:ln><a:noFill/></a:ln>",
        );
        let no_fill_doc = chart_space_of(&no_fill);
        let no_fill_model = parse_chart_part(
            no_fill_doc.root_element(),
            &ChartParseContext {
                color_resolver: Some(&resolver),
                ..Default::default()
            },
        )
        .expect("chart parses");
        assert!(no_fill_model.cat_axis_line_hidden);
        assert!(no_fill_model.val_axis_line_hidden);
        assert_eq!(no_fill_model.cat_axis_line_width_emu, None);
        assert_eq!(no_fill_model.val_axis_line_width_emu, None);

        for line in ["<a:ln/>", "<a:ln><a:prstDash val=\"dash\"/></a:ln>"] {
            let unresolved_overlay = xml.replace(
                "<a:ln><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill></a:ln>",
                line,
            );
            let unresolved_doc = chart_space_of(&unresolved_overlay);
            let unresolved = parse_chart_part(
                unresolved_doc.root_element(),
                &ChartParseContext {
                    color_resolver: Some(&resolver),
                    ..Default::default()
                },
            )
            .expect("chart parses");
            assert_eq!(unresolved.cat_axis_line_color, None);
            assert_eq!(unresolved.val_axis_line_color, None);
            assert_eq!(unresolved.cat_axis_line_width_emu, None);
            assert_eq!(unresolved.val_axis_line_width_emu, None);
            assert_eq!(unresolved.cat_axis_line_paint_authored, None);
            assert_eq!(unresolved.val_axis_line_paint_authored, None);
        }
    }

    #[test]
    fn omitted_style_uses_theme_aware_style_two_axis_and_gridline_roles() {
        struct NonBlackStyleTwoResolver {
            format_scheme: crate::theme::ThemeFormatScheme,
        }
        impl ColorResolver for NonBlackStyleTwoResolver {
            fn resolve_solid_fill(&self, node: Node) -> Option<String> {
                let color = node.children().find(|child| {
                    child.is_element() && matches!(child.tag_name().name(), "schemeClr" | "srgbClr")
                })?;
                match color.tag_name().name() {
                    "schemeClr" => self.resolve_scheme_color(color.attribute("val")?),
                    "srgbClr" => color.attribute("val").map(str::to_owned),
                    _ => None,
                }
            }

            fn resolve_scheme_color(&self, name: &str) -> Option<String> {
                Some(
                    match name {
                        "tx1" | "dk1" => "2468AC",
                        "bg1" | "lt1" => "FFFFFF",
                        "accent1" => "4472C4",
                        _ => return None,
                    }
                    .to_owned(),
                )
            }

            fn theme_format_scheme(&self) -> Option<&crate::theme::ThemeFormatScheme> {
                Some(&self.format_scheme)
            }
        }

        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:chart><c:plotArea>
              <c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart>
              <c:catAx/><c:valAx><c:majorGridlines/></c:valAx>
            </c:plotArea></c:chart></c:chartSpace>"#,
        );
        let theme = format!(
            r#"<a:theme xmlns:a="{A_NS}"><a:themeElements><a:fmtScheme name="Office">
              <a:fillStyleLst/><a:lnStyleLst><a:ln w="22222"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
              <a:effectStyleLst/><a:bgFillStyleLst/>
            </a:fmtScheme></a:themeElements></a:theme>"#,
        );
        let resolver = NonBlackStyleTwoResolver {
            format_scheme: crate::theme::ThemeFormatScheme::parse(&theme),
        };
        let model = parse_chart_part(
            chart_space_of(&xml).root_element(),
            &ChartParseContext {
                color_resolver: Some(&resolver),
                ..Default::default()
            },
        )
        .expect("styleless chart parses");
        assert_eq!(model.legacy_chart_style, None);
        assert_eq!(model.cat_axis_line_color, None);
        assert_eq!(model.val_axis_line_color, None);
        assert_eq!(model.val_axis_gridline_color, None);
        let roles = model.classic_chart_style_roles.expect("style-2 roles");
        for role in ["categoryAxis", "valueAxis", "gridlineMajor"] {
            assert_eq!(
                roles[role].line_colors.as_deref(),
                Some(&[Some("2468AC".to_owned())][..])
            );
            assert_eq!(roles[role].line_width_emu, Some(22222));
            assert_eq!(roles[role].line_paint_authored, Some(true));
        }
    }

    /// (b) Combo chart: a bar series on the primary value axis plus a line
    /// series bound to a SECONDARY value axis (`axPos="r"`). Verifies the
    /// series↔axId binding produces `series_type: "line"` and
    /// `use_secondary_axis: true` on the line series only, and that
    /// `secondary_val_axis` is populated from the right-hand `<c:valAx>`.
    #[test]
    fn parse_chart_part_combo_with_secondary_axis() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:chart><c:plotArea>
                <c:barChart>
                  <c:barDir val="col"/>
                  <c:grouping val="clustered"/>
                  <c:ser>
                    <c:idx val="0"/>
                    <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Units</c:v></c:pt></c:strCache></c:strRef></c:tx>
                    <c:cat><c:strCache><c:pt idx="0"><c:v>Jan</c:v></c:pt><c:pt idx="1"><c:v>Feb</c:v></c:pt></c:strCache></c:cat>
                    <c:val><c:numCache><c:pt idx="0"><c:v>5</c:v></c:pt><c:pt idx="1"><c:v>7</c:v></c:pt></c:numCache></c:val>
                  </c:ser>
                  <c:axId val="1"/>
                  <c:axId val="2"/>
                </c:barChart>
                <c:lineChart>
                  <c:grouping val="standard"/>
                  <c:ser>
                    <c:idx val="1"/>
                    <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Margin %</c:v></c:pt></c:strCache></c:strRef></c:tx>
                    <c:cat><c:strCache><c:pt idx="0"><c:v>Jan</c:v></c:pt><c:pt idx="1"><c:v>Feb</c:v></c:pt></c:strCache></c:cat>
                    <c:val><c:numCache><c:pt idx="0"><c:v>0.3</c:v></c:pt><c:pt idx="1"><c:v>0.4</c:v></c:pt></c:numCache></c:val>
                  </c:ser>
                  <c:axId val="1"/>
                  <c:axId val="3"/>
                </c:lineChart>
                <c:catAx><c:axId val="1"/><c:axPos val="b"/></c:catAx>
                <c:valAx>
                  <c:axId val="2"/>
                  <c:axPos val="l"/>
                  <c:crosses val="autoZero"/>
                </c:valAx>
                <c:valAx>
                  <c:axId val="3"/>
                  <c:axPos val="r"/>
                  <c:crosses val="max"/>
                  <c:scaling><c:logBase val="10"/><c:orientation val="maxMin"/><c:min val="0.01"/><c:max val="1"/></c:scaling>
                  <c:tickLblPos val="none"/>
                  <c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill><a:prstDash val="sysDot"/></a:ln></c:spPr>
                  <c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="654321"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr></c:majorGridlines>
                  <c:majorUnit val="0.25"/>
                  <c:minorUnit val="0.05"/>
                  <c:minorTickMark val="cross"/>
                  <c:minorGridlines><c:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:prstDash val="dot"/></a:ln></c:spPr></c:minorGridlines>
                  <c:txPr><a:p><a:pPr><a:defRPr b="1"><a:latin typeface="Tick Face"/></a:defRPr></a:pPr></a:p></c:txPr>
                  <c:title><c:tx><c:rich><a:bodyPr vert="vert"/><a:p><a:r><a:rPr sz="900" b="0"><a:latin typeface="Title Face"/></a:rPr><a:t>Margin</a:t></a:r></a:p></c:rich></c:tx>
                    <c:layout><c:manualLayout><c:x val="0.1"/><c:y val="0.2"/></c:manualLayout></c:layout>
                  </c:title>
                </c:valAx>
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
        .expect("combo chart parses");

        assert_eq!(m.chart_type, "clusteredBar");
        assert_eq!(m.series.len(), 2);

        let bar_series = &m.series[0];
        assert_eq!(bar_series.name, "Units");
        // Every series now carries its chart-group type (the renderer keys line
        // vs. non-line off this; a bar series is `Some("bar")`, treated as
        // non-line, identical in rendering to the old `None`).
        assert_eq!(bar_series.series_type.as_deref(), Some("bar"));
        assert_eq!(bar_series.use_secondary_axis, None);

        let line_series = &m.series[1];
        assert_eq!(line_series.name, "Margin %");
        assert_eq!(line_series.series_type.as_deref(), Some("line"));
        assert_eq!(line_series.use_secondary_axis, Some(true));

        let sec = m.secondary_val_axis.expect("secondary axis populated");
        assert_eq!(sec.min, Some(0.01));
        assert_eq!(sec.max, Some(1.0));
        assert_eq!(sec.title.as_deref(), Some("Margin"));
        assert!(!sec.hidden);
        assert_eq!(sec.font_face.as_deref(), Some("Tick Face"));
        assert_eq!(sec.font_bold, Some(true));
        assert_eq!(sec.log_base, Some(10.0));
        assert_eq!(sec.orientation.as_deref(), Some("maxMin"));
        assert_eq!(sec.tick_label_pos.as_deref(), Some("none"));
        assert_eq!(sec.line_color.as_deref(), Some("ABCDEF"));
        assert_eq!(sec.line_width_emu, Some(19050));
        assert_eq!(sec.line_dash.as_deref(), Some("sysDot"));
        assert_eq!(sec.crosses.as_deref(), Some("max"));
        assert_eq!(sec.crosses_at, None);
        assert!(sec.major_gridlines);
        assert_eq!(sec.major_gridline_color.as_deref(), Some("654321"));
        assert_eq!(sec.major_gridline_width_emu, Some(9525));
        assert_eq!(sec.major_gridline_dash.as_deref(), Some("dash"));
        assert_eq!(sec.minor_tick_mark.as_deref(), Some("cross"));
        assert!(sec.minor_gridlines);
        assert_eq!(sec.minor_gridline_color.as_deref(), Some("123456"));
        assert_eq!(sec.minor_gridline_width_emu, Some(12700));
        assert_eq!(sec.minor_gridline_dash.as_deref(), Some("dot"));
        assert_eq!(sec.minor_unit, Some(0.05));
        assert_eq!(sec.title_font_face.as_deref(), Some("Title Face"));
        assert_eq!(sec.title_font_size_hpt, Some(900));
        assert_eq!(sec.title_font_bold, Some(false));
        assert_eq!(sec.title_rotation, None);
        assert_eq!(sec.title_vertical_mode.as_deref(), Some("vert"));
        assert_eq!(
            sec.title_manual_layout.as_ref().map(|layout| layout.x),
            Some(0.1)
        );
        // #738: an explicit `<c:majorUnit>` on the secondary axis (§21.2.2.103)
        // is threaded into the model (was silently dropped before).
        assert_eq!(sec.major_unit, Some(0.25));
        // The primary value axis declared no majorUnit → stays None.
        assert_eq!(m.val_axis_major_unit, None);
    }

    #[test]
    fn parse_chart_part_two_bar_groups_retains_top_category_axis() {
        let group = |idx: usize, name: &str, cat_axis: u32, val_axis: u32| {
            format!(
                r#"<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
              <c:ser><c:idx val="{idx}"/><c:tx><c:v>{name}</c:v></c:tx>
                <c:cat><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>{name} Cat</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>{}</c:v></c:pt></c:numLit></c:val>
              </c:ser><c:axId val="{cat_axis}"/><c:axId val="{val_axis}"/></c:barChart>"#,
                if idx == 0 { 100 } else { 1 },
            )
        };
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:chart><c:plotArea>
              {}{}
              <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
              <c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
              <c:catAx><c:axId val="3"/><c:axPos val="t"/><c:lblAlgn val="r"/><c:lblOffset val="250%"/>
                <c:tickLblSkip val="2"/><c:tickMarkSkip val="3"/>
                <c:title><c:tx><c:rich><a:p><a:r><a:t>Top Categories</a:t></a:r></a:p></c:rich></c:tx></c:title>
                <c:crossAx val="4"/><c:crosses val="max"/></c:catAx>
              <c:valAx><c:axId val="4"/><c:axPos val="r"/><c:crossAx val="3"/><c:crosses val="max"/>
                <c:scaling><c:min val="0"/><c:max val="1"/></c:scaling></c:valAx>
            </c:plotArea></c:chart></c:chartSpace>"#,
            group(0, "Primary", 1, 2),
            group(1, "Secondary", 3, 4),
        );
        let doc = chart_space_of(&xml);
        let model = parse_chart_part(
            doc.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("dual bar axes parse");

        assert_eq!(model.series[0].use_secondary_axis, None);
        assert_eq!(model.series[1].use_secondary_axis, Some(true));
        let top = model.secondary_cat_axis.expect("top category axis");
        assert_eq!(top.title.as_deref(), Some("Top Categories"));
        assert_eq!(top.label_alignment.as_deref(), Some("r"));
        assert_eq!(top.label_offset_percent, Some(250));
        assert_eq!(top.tick_label_skip, Some(2));
        assert_eq!(top.tick_mark_skip, Some(3));
        let right = model.secondary_val_axis.expect("right value axis");
        assert_eq!(right.max, Some(1.0));
    }

    #[test]
    fn parse_category_axis_label_alignment_and_offset_contracts() {
        let parse_axis = |axis_name: &str, authored: &str| {
            let xml = format!(
                r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:chart><c:plotArea>
                  <c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
                    <c:ser><c:idx val="0"/><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                      <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>
                    <c:axId val="1"/><c:axId val="2"/></c:barChart>
                  <c:{axis_name}><c:axId val="1"/><c:axPos val="b"/>{authored}<c:crossAx val="2"/></c:{axis_name}>
                  <c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
                </c:plotArea></c:chart></c:chartSpace>"#,
            );
            let doc = chart_space_of(&xml);
            parse_chart_part(
                doc.root_element(),
                &ChartParseContext {
                    color_resolver: Some(&FixtureResolver),
                    ..Default::default()
                },
            )
            .expect("axis parses")
        };

        let strict = parse_axis("catAx", r#"<c:lblAlgn val="r"/><c:lblOffset val="250%"/>"#);
        assert_eq!(strict.cat_axis_label_alignment.as_deref(), Some("r"));
        assert_eq!(strict.cat_axis_label_offset_percent, Some(250));

        let transitional = parse_axis("catAx", r#"<c:lblAlgn val="l"/><c:lblOffset val="0"/>"#);
        assert_eq!(transitional.cat_axis_label_alignment.as_deref(), Some("l"));
        assert_eq!(transitional.cat_axis_label_offset_percent, Some(0));

        let bare_date = parse_axis("dateAx", r#"<c:lblOffset/>"#);
        assert_eq!(bare_date.cat_axis_label_alignment, None);
        assert_eq!(bare_date.cat_axis_label_offset_percent, Some(100));

        let invalid = parse_axis(
            "catAx",
            r#"<c:lblAlgn val="justify"/><c:lblOffset val="1001"/>"#,
        );
        assert_eq!(invalid.cat_axis_label_alignment, None);
        assert_eq!(invalid.cat_axis_label_offset_percent, None);

        let omitted = parse_axis("catAx", "");
        assert_eq!(omitted.cat_axis_label_alignment, None);
        assert_eq!(omitted.cat_axis_label_offset_percent, None);
    }

    /// (d) A date-category axis (`<c:dateAx>` instead of `<c:catAx>`) combined
    /// with `<c:date1904/>`. `parse_chart_part` treats `dateAx` identically to
    /// `catAx` for every cat-axis probe (hidden/format-code/etc.) — this pins
    /// that the dateAx path is actually reached (not silently skipped because
    /// the finder only looked for `catAx`).
    #[test]
    fn parse_chart_part_date_axis_and_date1904() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:date1904/>
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:grouping val="standard"/>
                  <c:ser>
                    <c:idx val="0"/>
                    <c:tx><c:v>Temp</c:v></c:tx>
                    <c:cat><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:cat>
                    <c:val><c:numCache><c:pt idx="0"><c:v>21</c:v></c:pt><c:pt idx="1"><c:v>23</c:v></c:pt></c:numCache></c:val>
                  </c:ser>
                </c:lineChart>
                <c:dateAx>
                  <c:axPos val="b"/>
                  <c:numFmt formatCode="m/d/yyyy"/>
                  <c:baseTimeUnit val="months"/>
                  <c:majorUnit val="1.5"/>
                  <c:majorTimeUnit val="months"/>
                  <c:minorUnit val="0.5"/>
                  <c:minorTimeUnit val="months"/>
                </c:dateAx>
                <c:valAx><c:axPos val="l"/></c:valAx>
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
        .expect("dateAx chart parses");

        assert_eq!(m.chart_type, "line");
        assert!(m.date1904);
        assert_eq!(m.cat_axis_format_code.as_deref(), Some("m/d/yyyy"));
        assert_eq!(m.cat_axis_is_date, Some(true));
        assert_eq!(m.cat_axis_base_time_unit.as_deref(), Some("months"));
        assert_eq!(m.cat_axis_major_unit, Some(1.5));
        assert_eq!(m.cat_axis_major_time_unit.as_deref(), Some("months"));
        assert_eq!(m.cat_axis_minor_unit, Some(0.5));
        assert_eq!(m.cat_axis_minor_time_unit.as_deref(), Some("months"));
        assert!(!m.cat_axis_hidden);
    }

    #[test]
    fn ct_boolean_axis_delete_bare_element_is_true() {
        // §21.2.2.40 `<c:delete/>` on an axis ⇒ axis deleted (val default true).
        let bare_xml = format!(r#"<c:catAx xmlns:c="{C_NS}"><c:delete/></c:catAx>"#);
        let bare = root_of(&bare_xml);
        assert!(
            axis_is_deleted(bare.root_element()),
            "bare <c:delete/> ⇒ axis deleted"
        );
        let off_xml = format!(r#"<c:catAx xmlns:c="{C_NS}"><c:delete val="0"/></c:catAx>"#);
        let off = root_of(&off_xml);
        assert!(
            !axis_is_deleted(off.root_element()),
            "val=\"0\" ⇒ not deleted"
        );
        let absent_xml = format!(r#"<c:catAx xmlns:c="{C_NS}"/>"#);
        let absent = root_of(&absent_xml);
        assert!(
            !axis_is_deleted(absent.root_element()),
            "no <c:delete> ⇒ axis shown"
        );
    }

    #[test]
    fn extract_axis_crosses_reads_crosses_and_crosses_at() {
        let xml = format!(r#"<c:valAx xmlns:c="{C_NS}"><c:crosses val="max"/></c:valAx>"#);
        let d = root_of(&xml);
        assert_eq!(
            extract_axis_crosses(d.root_element()),
            (Some("max".to_string()), None)
        );

        let xml2 = format!(r#"<c:valAx xmlns:c="{C_NS}"><c:crossesAt val="3.5"/></c:valAx>"#);
        let d2 = root_of(&xml2);
        assert_eq!(extract_axis_crosses(d2.root_element()), (None, Some(3.5)));

        let xml3 = format!(r#"<c:valAx xmlns:c="{C_NS}"></c:valAx>"#);
        let d3 = root_of(&xml3);
        assert_eq!(extract_axis_crosses(d3.root_element()), (None, None));
    }

    #[test]
    fn extract_manual_layout_full_and_defaults() {
        let xml = format!(
            r#"<c:layout xmlns:c="{C_NS}"><c:manualLayout>
              <c:layoutTarget val="inner"/>
              <c:xMode val="edge"/><c:yMode val="edge"/>
              <c:x val="0.1"/><c:y val="0.2"/><c:w val="0.5"/><c:h val="0.6"/>
            </c:manualLayout></c:layout>"#
        );
        let d = root_of(&xml);
        let layout = extract_manual_layout(d.root_element()).expect("manualLayout present");
        assert_eq!(layout.x_mode, "edge");
        assert_eq!(layout.y_mode, "edge");
        assert_eq!(layout.w_mode, "factor");
        assert_eq!(layout.h_mode, "factor");
        assert_eq!(layout.layout_target.as_deref(), Some("inner"));
        assert_eq!(layout.x, 0.1);
        assert_eq!(layout.y, 0.2);
        assert_eq!(layout.w, Some(0.5));
        assert_eq!(layout.h, Some(0.6));

        let omitted_target_xml = format!(
            r#"<c:layout xmlns:c="{C_NS}"><c:manualLayout>
              <c:xMode val="edge"/><c:yMode val="edge"/>
              <c:x val="0.1"/><c:y val="0.2"/><c:w val="0.5"/><c:h val="0.6"/>
            </c:manualLayout></c:layout>"#
        );
        let omitted_target_doc = root_of(&omitted_target_xml);
        let omitted_target =
            extract_manual_layout(omitted_target_doc.root_element()).expect("manualLayout present");
        assert_eq!(omitted_target.layout_target.as_deref(), Some("outer"));

        let all_modes_xml = format!(
            r#"<c:layout xmlns:c="{C_NS}"><c:manualLayout>
              <c:wMode val="edge"/><c:hMode val="edge"/>
              <c:x val="0.1"/><c:y val="0.2"/><c:w val="0.5"/><c:h val="0.6"/>
            </c:manualLayout></c:layout>"#
        );
        let all_modes_doc = root_of(&all_modes_xml);
        let all_modes =
            extract_manual_layout(all_modes_doc.root_element()).expect("manualLayout present");
        assert_eq!(all_modes.x_mode, "factor");
        assert_eq!(all_modes.y_mode, "factor");
        assert_eq!(all_modes.w_mode, "edge");
        assert_eq!(all_modes.h_mode, "edge");
    }

    #[test]
    fn extract_manual_layout_absent_returns_none() {
        let xml = format!(r#"<c:layout xmlns:c="{C_NS}"></c:layout>"#);
        let d = root_of(&xml);
        assert!(extract_manual_layout(d.root_element()).is_none());
    }

    #[test]
    fn stock_group_series_retain_secondary_value_axis_binding() {
        let series = |index: u32, value: u32| {
            format!(
                r#"<c:ser><c:idx val="{index}"/><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>{value}</c:v></c:pt></c:numLit></c:val></c:ser>"#,
            )
        };
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:chart><c:plotArea>
              <c:stockChart>{}{}{}<c:hiLowLines/><c:axId val="1"/><c:axId val="3"/></c:stockChart>
              <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="3"/></c:catAx>
              <c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
              <c:valAx><c:axId val="3"/><c:axPos val="r"/><c:crossAx val="1"/></c:valAx>
            </c:plotArea></c:chart></c:chartSpace>"#,
            series(0, 55),
            series(1, 11),
            series(2, 32),
        );
        let document = chart_space_of(&xml);
        let model = parse_chart_part(
            document.root_element(),
            &ChartParseContext {
                color_resolver: Some(&FixtureResolver),
                ..Default::default()
            },
        )
        .expect("stock");

        assert_eq!(model.chart_type, "stock");
        assert!(model.secondary_val_axis.is_some());
        assert!(model
            .series
            .iter()
            .all(|stock_series| stock_series.use_secondary_axis == Some(true)));
    }

    #[test]
    fn classic_group_axis_id_count_and_duplicates_fail_closed() {
        let parse_ids = |ids: &[&str]| {
            let axis_ids = ids
                .iter()
                .map(|id| format!(r#"<c:axId val="{id}"/>"#))
                .collect::<String>();
            let xml = format!(
                r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea><c:surfaceChart>{axis_ids}</c:surfaceChart></c:plotArea></c:chart></c:chartSpace>"#,
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
        assert!(parse_ids(&["1", "2", "3"]).is_some());
        assert!(parse_ids(&["1", "2", "3", "4"]).is_none());
        assert!(parse_ids(&["1", "1"]).is_none());
        assert!(parse_ids(&["-2147483648"]).is_some());
        assert!(parse_ids(&["-2147483649"]).is_none());
        assert!(parse_ids(&["4294967296"]).is_none());
        assert!(parse_ids(&["not-an-axis-id"]).is_none());
    }

    #[test]
    fn signed_32_bit_axis_ids_remain_opaque_cross_reference_keys() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}"><c:chart><c:plotArea>
              <c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
                <c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit>
                  <c:ptCount val="1"/><c:pt idx="0"><c:v>2</c:v></c:pt>
                </c:numLit></c:val></c:ser>
                <c:axId val=" -2068027336 "/><c:axId val=" -2113994440 "/>
              </c:barChart>
              <c:catAx><c:axId val=" -2068027336 "/><c:axPos val="b"/>
                <c:crossAx val="-2113994440"/></c:catAx>
              <c:valAx><c:axId val=" -2113994440 "/><c:axPos val="l"/>
                <c:crossAx val="-2068027336"/></c:valAx>
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
        .expect("signed axis identifiers");
        let group = model.plot_groups.expect("group").remove(0);
        assert_eq!(
            group.axis_ids,
            Some(vec!["-2068027336".to_string(), "-2113994440".to_string()]),
        );
        assert_eq!(group.category_axis, "primary");
        assert_eq!(group.value_axis, "primary");
    }
}
