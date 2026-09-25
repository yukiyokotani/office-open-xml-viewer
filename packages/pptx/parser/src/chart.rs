//! Chart element parsing: the legacy DrawingML chart (`c:` namespace) and the
//! newer chartEx (`cx:` namespace) parsers, plus the pptx `ColorResolver` the
//! shared `ooxml_common::chart` helpers use to resolve `<a:solidFill>` colours.
//! Extracted verbatim from `lib.rs`. The general colour grammar
//! (`parse_color_node`) stays in `lib.rs` and is imported here for the
//! `PptxColorResolver`; both chart parsers now delegate their structure walk to
//! `ooxml_common::chart`.

use crate::chart_compatibility::{
    apply_powerpoint_chartex_chart_space_frame, apply_powerpoint_classic_chart_space_frame,
};
use crate::parse_color_node;
use crate::parse_preflighted_pptx_xml;
use crate::theme::PptxRawSchemeResolver;
use crate::types::*;
use ooxml_common::color::ThemeResolver;
use std::collections::HashMap;

/// `ooxml_common::chart::ColorResolver` implementation backed by pptx's
/// `HashMap<String, String>` theme palette and PowerPoint's tint formula.
/// Used by chart helpers in ooxml-common that need to resolve
/// `<a:solidFill>` text colors without owning the theme storage.
pub(crate) struct PptxColorResolver<'a> {
    pub(crate) theme: &'a HashMap<String, String>,
    pub(crate) theme_format_scheme: Option<&'a ooxml_common::theme::ThemeFormatScheme>,
}

impl ooxml_common::chart::ColorResolver for PptxColorResolver<'_> {
    fn resolve_solid_fill(&self, node: roxmltree::Node<'_, '_>) -> Option<String> {
        parse_color_node(node, self.theme)
    }

    fn resolve_scheme_color(&self, name: &str) -> Option<String> {
        PptxRawSchemeResolver { theme: self.theme }.resolve_scheme_color(name)
    }

    fn theme_major_font_latin(&self) -> Option<String> {
        // pptx stores the theme major/minor Latin faces under the `+mj-lt` /
        // `+mn-lt` keys of its color+font map (see lib.rs parse_theme_colors).
        self.theme.get("+mj-lt").cloned()
    }

    fn theme_minor_font_latin(&self) -> Option<String> {
        self.theme.get("+mn-lt").cloned()
    }

    fn resolve_series_accent(&self, idx: usize) -> Option<String> {
        self.theme.get(&format!("accent{}", idx % 6 + 1)).cloned()
    }

    fn theme_format_scheme(&self) -> Option<&ooxml_common::theme::ThemeFormatScheme> {
        self.theme_format_scheme
    }

    fn office_dark_text_contrast_applies(&self, style: u8) -> bool {
        style == 41
    }

    // PowerPoint-produced style-41 controls establish the same carrier
    // boundary as Word for the tested single rich paragraph: a present
    // a:pPr/a:defRPr receives automatic lt1 title text, while run-only lang
    // and size properties retain the numeric dk1 result. Styles 40 and 42..48
    // remain outside this host claim until separately observed.
    fn office_dark_title_contrast_applies(&self, style: u8) -> bool {
        style == 41
    }
}

/// Parse a legacy OOXML chart (`c:` namespace) — barChart / lineChart etc.
///
/// Thin pptx adapter over the shared
/// [`ooxml_common::chart::parse_chart_part`]: it builds a [`PptxColorResolver`]
/// from the theme palette, delegates the entire chart-structure parse, and
/// wraps the resulting [`ChartModel`] in a pptx [`ChartElement`] graphic frame.
/// The frame geometry (`x`/`y`/`width`/`height`) is filled in by the caller
/// from the slide's `<p:graphicFrame><a:xfrm>`; here it defaults to 0.
#[cfg(test)]
pub(crate) fn parse_legacy_chart(
    xml: &str,
    theme: &HashMap<String, String>,
) -> Option<ChartElement> {
    parse_legacy_chart_with_user_shapes(xml, None, theme)
}

#[cfg(test)]
pub(crate) fn parse_legacy_chart_with_user_shapes(
    xml: &str,
    user_shapes_xml: Option<&str>,
    theme: &HashMap<String, String>,
) -> Option<ChartElement> {
    parse_legacy_chart_with_style_parts(xml, None, None, user_shapes_xml, theme, None)
}

#[cfg(test)]
pub(crate) fn parse_legacy_chart_with_style_parts(
    xml: &str,
    style_xml: Option<&str>,
    color_style_xml: Option<&str>,
    user_shapes_xml: Option<&str>,
    theme: &HashMap<String, String>,
    theme_format_scheme: Option<&ooxml_common::theme::ThemeFormatScheme>,
) -> Option<ChartElement> {
    let images = ooxml_common::chart::ChartImageRelationships::default();
    parse_legacy_chart_with_style_parts_and_images(
        xml,
        style_xml,
        color_style_xml,
        user_shapes_xml,
        theme,
        theme_format_scheme,
        &images,
    )
}

pub(crate) fn parse_legacy_chart_with_style_parts_and_images(
    xml: &str,
    style_xml: Option<&str>,
    color_style_xml: Option<&str>,
    user_shapes_xml: Option<&str>,
    theme: &HashMap<String, String>,
    theme_format_scheme: Option<&ooxml_common::theme::ThemeFormatScheme>,
    image_resolver: &dyn ooxml_common::chart::ChartImageResolver,
) -> Option<ChartElement> {
    let doc = parse_preflighted_pptx_xml(xml).ok()?;
    let root = doc.root_element();
    let resolver = PptxColorResolver {
        theme,
        theme_format_scheme,
    };
    let mut chart = ooxml_common::chart::parse_chart_part_with_style_parts_and_images(
        root,
        &resolver,
        style_xml,
        color_style_xml,
        image_resolver,
    )?;
    apply_powerpoint_classic_chart_space_frame(&mut chart);
    if let Some(user_shapes_xml) = user_shapes_xml {
        if let Ok(user_shapes_doc) = parse_preflighted_pptx_xml(user_shapes_xml) {
            let text_boxes = ooxml_common::chart::parse_chart_user_shapes_for_chart(
                root,
                user_shapes_doc.root_element(),
                &resolver,
            );
            if !text_boxes.is_empty() {
                chart.chart_text_boxes = Some(text_boxes);
            }
        }
    }
    Some(ChartElement {
        id: None,
        x: 0,
        y: 0,
        width: 0,
        height: 0,
        rotation: 0.0,
        flip_h: false,
        flip_v: false,
        chart,
    })
}

/// Parse a modern chartEx (cx: namespace) — waterfall, treemap, etc.
///
/// Thin pptx adapter over the shared
/// [`ooxml_common::chart::parse_chartex_part`]: it builds a [`PptxColorResolver`]
/// from the theme palette, delegates the entire chartEx-structure parse, and
/// wraps the resulting [`ChartModel`] in a pptx [`ChartElement`] graphic frame.
/// The frame geometry (`x`/`y`/`width`/`height`) is filled in by the caller
/// from the slide's `<p:graphicFrame><a:xfrm>`; here it defaults to 0.
#[cfg(test)]
pub(crate) fn parse_chartex(
    xml: &str,
    style_xml: Option<&str>,
    color_style_xml: Option<&str>,
    theme: &HashMap<String, String>,
    theme_format_scheme: Option<&ooxml_common::theme::ThemeFormatScheme>,
) -> Option<ChartElement> {
    let images = ooxml_common::chart::ChartImageRelationships::default();
    parse_chartex_with_images(
        xml,
        style_xml,
        color_style_xml,
        theme,
        theme_format_scheme,
        &images,
    )
}

pub(crate) fn parse_chartex_with_images(
    xml: &str,
    style_xml: Option<&str>,
    color_style_xml: Option<&str>,
    theme: &HashMap<String, String>,
    theme_format_scheme: Option<&ooxml_common::theme::ThemeFormatScheme>,
    image_resolver: &dyn ooxml_common::chart::ChartImageResolver,
) -> Option<ChartElement> {
    let doc = parse_preflighted_pptx_xml(xml).ok()?;
    let root = doc.root_element();
    let resolver = PptxColorResolver {
        theme,
        theme_format_scheme,
    };
    // The shared chart grammar reparses the optional style XML. Admit it
    // through the PPTX-local node ceiling first so the second parse only ever
    // sees an already bounded document.
    // chartEx (waterfall/boxWhisker/…) reads its title font size from the
    // associated chartStyle part when the `<cx:title>` itself carries none.
    let mut chart = ooxml_common::chart::parse_chartex_part_with_style_parts_and_images(
        root,
        &resolver,
        style_xml,
        color_style_xml,
        image_resolver,
    )?;
    apply_powerpoint_chartex_chart_space_frame(root, &mut chart);
    Some(ChartElement {
        id: None,
        x: 0,
        y: 0,
        width: 0,
        height: 0,
        rotation: 0.0,
        flip_h: false,
        flip_v: false,
        chart,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    const C_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

    #[test]
    fn powerpoint_chart_host_style_scope_retains_seventh_point_fallback() {
        let theme = HashMap::from([
            ("dk1".to_string(), "111111".to_string()),
            ("lt1".to_string(), "FEFEFE".to_string()),
            ("accent1".to_string(), "808080".to_string()),
        ]);
        let points = (0..7)
            .map(|index| format!(r#"<c:pt idx="{index}"><c:v>1</c:v></c:pt>"#))
            .collect::<String>();
        let parse = |style: u8| {
            let xml = format!(
                r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <c:style val="{style}"/><c:chart><c:plotArea><c:pieChart><c:varyColors val="1"/>
                    <c:ser><c:idx val="0"/><c:order val="0"/>
                      <c:dPt><c:idx val="5"/><c:spPr><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:ln></c:spPr></c:dPt>
                      <c:val><c:numLit><c:ptCount val="7"/>{points}</c:numLit></c:val>
                    </c:ser>
                  </c:pieChart></c:plotArea></c:chart>
                </c:chartSpace>"#,
            );
            parse_legacy_chart(&xml, &theme)
                .expect("PowerPoint chart")
                .chart
        };
        let chart = parse(2);
        let role = &chart
            .classic_varying_point_chart_style_roles
            .as_ref()
            .expect("point-domain numeric roles")["dataPoint"];
        let colors = role
            .fill_colors
            .as_ref()
            .unwrap_or_else(|| panic!("point palette: {role:?}"));
        assert_eq!(colors[0].as_deref(), Some("808080"));
        assert_eq!(colors[6].as_deref(), None);
        assert_eq!(
            role.fill_semantic_fallback_indices.as_deref(),
            Some(&[6][..])
        );
        // Direct point paint stays separate from the automatic palette so
        // the renderer can preserve its precedence at either host boundary.
        assert_eq!(
            chart.series[0]
                .data_point_colors
                .as_ref()
                .expect("direct point color")[5]
                .as_deref(),
            Some("ABCDEF"),
        );
        let point = chart.series[0]
            .data_point_overrides
            .as_ref()
            .expect("point formatting")
            .iter()
            .find(|point| point.idx == 5)
            .expect("formatted point");
        assert_eq!(point.line_color.as_deref(), Some("123456"));
        for (style, expected) in [
            (40, "111111"),
            (41, "FEFEFE"),
            (42, "111111"),
            (48, "111111"),
        ] {
            let chart = parse(style);
            assert_eq!(
                chart
                    .classic_chart_style_roles
                    .as_ref()
                    .expect("numeric roles")["categoryAxis"]
                    .font_color
                    .as_deref(),
                Some(expected),
                "style {style}",
            );
        }
    }

    #[test]
    fn legacy_chart_uses_theme_accents_and_chart_wide_text_defaults() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:chart>
                <c:plotArea>
                  <c:barChart>
                    <c:barDir val="col"/>
                    <c:grouping val="clustered"/>
                    <c:ser>
                      <c:idx val="0"/><c:order val="0"/>
                      <c:tx><c:v>2025</c:v></c:tx>
                      <c:cat><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>T1</c:v></c:pt></c:strLit></c:cat>
                      <c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>12.5</c:v></c:pt></c:numLit></c:val>
                    </c:ser>
                    <c:ser>
                      <c:idx val="1"/><c:order val="1"/>
                      <c:tx><c:v>2026</c:v></c:tx>
                      <c:cat><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>T1</c:v></c:pt></c:strLit></c:cat>
                      <c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>15</c:v></c:pt></c:numLit></c:val>
                    </c:ser>
                    <c:axId val="1"/><c:axId val="2"/>
                  </c:barChart>
                  <c:catAx>
                    <c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/>
                    <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="900"/></a:pPr></a:p></c:txPr>
                  </c:catAx>
                  <c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
                </c:plotArea>
              </c:chart>
              <c:txPr>
                <a:bodyPr/><a:lstStyle/>
                <a:p><a:pPr><a:defRPr sz="1800"/></a:pPr></a:p>
              </c:txPr>
            </c:chartSpace>"#
        );
        let theme = HashMap::from([
            ("accent1".to_string(), "4F81BD".to_string()),
            ("accent2".to_string(), "C0504D".to_string()),
        ]);

        let element = parse_legacy_chart(&xml, &theme).expect("chart should parse");

        assert_eq!(element.chart.series[0].color.as_deref(), Some("4F81BD"));
        assert_eq!(element.chart.series[1].color.as_deref(), Some("C0504D"));
        assert_eq!(element.chart.cat_axis_font_size_hpt, Some(900));
        assert_eq!(element.chart.val_axis_font_size_hpt, Some(1800));
    }

    #[test]
    fn powerpoint_classic_chart_space_frame_tracks_style_boundaries() {
        let chart_xml = |style: Option<u8>, rounded: Option<bool>| {
            let style = style
                .map(|value| format!(r#"<c:style val="{value}"/>"#))
                .unwrap_or_default();
            let rounded = rounded
                .map(|value| format!(r#"<c:roundedCorners val="{}"/>"#, u8::from(value)))
                .unwrap_or_default();
            format!(
                r#"<c:chartSpace xmlns:c="{C_NS}">{style}{rounded}<c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#
            )
        };
        let theme = HashMap::from([
            ("dk1".to_string(), "000000".to_string()),
            ("lt1".to_string(), "FFFFFF".to_string()),
        ]);
        let theme_xml = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements>
          <a:fmtScheme name="Office"><a:fillStyleLst>
            <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
            <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
            <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
          </a:fillStyleLst><a:lnStyleLst>
            <a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
            <a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
            <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
          </a:lnStyleLst><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>
        </a:themeElements></a:theme>"#;
        let format_scheme = ooxml_common::theme::ThemeFormatScheme::parse(theme_xml);
        let parse = |style, rounded| {
            parse_legacy_chart_with_style_parts(
                &chart_xml(style, rounded),
                None,
                None,
                None,
                &theme,
                Some(&format_scheme),
            )
            .expect("classic chart")
            .chart
        };

        for style in [1, 32] {
            let chart = parse(Some(style), None);
            assert_eq!(chart.rounded_corners, Some(true), "style {style}");
            assert!(chart
                .classic_chart_style_roles
                .as_ref()
                .is_none_or(|roles| !roles.contains_key("chartArea")));
        }
        for style in [33, 40] {
            let chart = parse(Some(style), None);
            let frame = &chart.classic_chart_style_roles.as_ref().unwrap()["chartArea"];
            assert_eq!(
                frame.fill_colors.as_deref(),
                Some(&[Some("FFFFFF".to_string())][..]),
                "style {style}"
            );
            assert_eq!(
                frame.line_colors.as_deref(),
                Some(&[Some("898989".to_string())][..]),
                "style {style}"
            );
            assert_eq!(frame.line_width_emu, Some(9_525), "style {style}");
        }
        for style in [41, 48] {
            let chart = parse(Some(style), None);
            let frame = &chart.classic_chart_style_roles.as_ref().unwrap()["chartArea"];
            assert_eq!(
                frame.fill_colors.as_deref(),
                Some(&[Some("000000".to_string())][..]),
                "style {style}"
            );
            assert_eq!(frame.line_hidden, Some(true), "style {style}");
        }

        let omitted = parse(None, None);
        assert_eq!(omitted.rounded_corners, Some(true));
        assert!(omitted
            .classic_chart_style_roles
            .as_ref()
            .is_none_or(|roles| !roles.contains_key("chartArea")));
        assert_eq!(parse(Some(33), Some(false)).rounded_corners, Some(false));

        let linked_style = format!(
            r#"<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="{A_NS}"><cs:chartArea><cs:spPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill><a:ln w="25400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></cs:spPr></cs:chartArea></cs:chartStyle>"#
        );
        let linked = parse_legacy_chart_with_style_parts(
            &chart_xml(Some(33), None),
            Some(&linked_style),
            None,
            None,
            &HashMap::new(),
            None,
        )
        .expect("classic chart with linked style")
        .chart;
        let linked_frame = &linked.chart_style_roles.as_ref().unwrap()["chartArea"];
        assert_eq!(
            linked_frame.fill_colors.as_deref(),
            Some(&[Some("00FF00".to_string())][..])
        );
        assert_eq!(
            linked_frame.line_colors.as_deref(),
            Some(&[Some("FF0000".to_string())][..])
        );
        assert_eq!(linked_frame.line_width_emu, Some(25_400));
    }

    #[test]
    fn powerpoint_style_41_title_contrast_requires_paragraph_default_run() {
        let chart_xml = |style: u8, paragraph_properties: &str, run_properties: &str| {
            format!(
                r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
                  <c:style val="{style}"/><c:chart><c:title><c:tx><c:rich>
                    <a:bodyPr/><a:lstStyle/><a:p>{paragraph_properties}<a:r>{run_properties}<a:t>Title</a:t></a:r></a:p>
                  </c:rich></c:tx></c:title><c:plotArea><c:barChart><c:barDir val="col"/>
                    <c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>
                  </c:barChart></c:plotArea></c:chart></c:chartSpace>"#
            )
        };
        let theme = HashMap::from([
            ("dk1".to_string(), "000000".to_string()),
            ("lt1".to_string(), "FFFFFF".to_string()),
        ]);
        let parse = |style, paragraph_properties: &str, run_properties: &str| {
            parse_legacy_chart(
                &chart_xml(style, paragraph_properties, run_properties),
                &theme,
            )
            .expect("classic chart")
            .chart
            .classic_chart_style_roles
            .expect("numeric roles")["title"]
                .font_color
                .clone()
        };

        assert_eq!(
            parse(41, "<a:pPr><a:defRPr sz=\"1400\"/></a:pPr>", ""),
            Some("FFFFFF".to_string()),
        );
        assert_eq!(
            parse(41, "", "<a:rPr lang=\"en-US\"/>"),
            Some("000000".to_string()),
        );
        assert_eq!(
            parse(40, "<a:pPr><a:defRPr sz=\"1400\"/></a:pPr>", ""),
            Some("000000".to_string()),
        );
        assert_eq!(
            parse(42, "<a:pPr><a:defRPr sz=\"1400\"/></a:pPr>", ""),
            Some("000000".to_string()),
        );
    }

    #[test]
    fn legacy_chart_honors_chart_local_color_map_override() {
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
              </c:barChart></c:plotArea></c:chart>
            </c:chartSpace>"#
        );
        let theme = HashMap::from([
            ("accent1".to_string(), "4472C4".to_string()),
            ("accent2".to_string(), "ED7D31".to_string()),
        ]);

        let element = parse_legacy_chart(&xml, &theme).expect("chart should parse");

        assert_eq!(element.chart.series[0].color.as_deref(), Some("ED7D31"));
    }

    #[test]
    fn legacy_chart_accepts_shared_chart_drawing_text_boxes() {
        let chart_xml = format!(
            r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#
        );
        let user_shapes_xml = format!(
            r#"<c:userShapes xmlns:c="{C_NS}" xmlns:cdr="http://schemas.openxmlformats.org/drawingml/2006/chartDrawing" xmlns:a="{A_NS}"><cdr:relSizeAnchor><cdr:from><cdr:x>0</cdr:x><cdr:y>0</cdr:y></cdr:from><cdr:to><cdr:x>1</cdr:x><cdr:y>0.1</cdr:y></cdr:to><cdr:sp><cdr:nvSpPr><cdr:cNvPr id="1" name="TitleBox"/><cdr:cNvSpPr/></cdr:nvSpPr><cdr:spPr/><cdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"/><a:t>Shared title</a:t></a:r></a:p></cdr:txBody></cdr:sp></cdr:relSizeAnchor></c:userShapes>"#
        );

        let element = parse_legacy_chart_with_user_shapes(
            &chart_xml,
            Some(&user_shapes_xml),
            &HashMap::new(),
        )
        .expect("chart should parse");

        let boxes = element.chart.chart_text_boxes.expect("chart text boxes");
        assert_eq!(boxes[0].paragraphs[0].runs[0].text, "Shared title");
    }
}
