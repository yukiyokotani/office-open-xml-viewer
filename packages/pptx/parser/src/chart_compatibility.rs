//! Isolated PowerPoint deltas over ECMA-376 classic chart-space defaults.

use ooxml_common::chart::ChartModel;
use roxmltree::Node;
use std::collections::BTreeMap;

/// Apply PowerPoint's observed host delta to the normative built-in
/// `chartArea` recipe. ECMA-376 §21.2.3.46 defines the style 1..48 recipes;
/// PowerPoint-produced samples covering every style,
/// the omitted style, and direct-paint controls showed 10pt rounded corners
/// throughout. Styles 1..32 add no chart-area paint; styles 33..40 retain the
/// light theme recipe with a 0.75pt outline; styles 41..48 retain the dark
/// theme recipe without an outline. Default-theme output is white/#898989 and
/// black respectively; custom-theme paint is not replaced by absolute colors.
/// Omitted style behaves as style 2. PowerPoint's PDF export also draws a
/// separate square 0.14pt graphic-frame path for every style; it is not a
/// chart-space style and therefore is intentionally not synthesized here.
/// Direct chart-space paint and a linked `chartArea` role remain authoritative
/// through the shared renderer's direct > linked precedence.
pub(crate) fn apply_powerpoint_classic_chart_space_frame(chart: &mut ChartModel) {
    chart.rounded_corners.get_or_insert(true);

    // Observed PowerPoint compatibility, not an ECMA-376 §21.2.2.145
    // automatic-layout rule. 2-D pie/doughnut controls at omitted/2/32/33/34/
    // 41/48 have no automatic plot fill; explicit gradient paint occupies the
    // circle's bounding square. Keep 3-D, of-pie and mixed groups outside this
    // evidence boundary; the renderer preserves manual layout and authored lines.
    let radial = matches!(chart.chart_type.as_str(), "pie" | "doughnut")
        && chart.three_d.is_none()
        && chart
            .plot_groups
            .as_ref()
            .is_none_or(|groups| groups.iter().all(|group| group.kind == chart.chart_type));
    if radial {
        chart.powerpoint_radial_plot_frame = Some(true);
    }
    let roles = chart
        .classic_chart_style_roles
        .get_or_insert_with(BTreeMap::new);
    // PPTX styles 1..32 also suppress automatic cartesian plot fill, unlike
    // ECMA-376 §21.2.3.46 Tables 1/3. PowerPoint controls and LibreOffice
    // commit 87350de8737ca91ba71b356ca9f33fea99e74591 establish that boundary.
    // Change built-in fill only: direct/linked paint, lines and effects win.
    if radial || matches!(chart.legacy_chart_style.unwrap_or(2), 1..=32) {
        if let Some(plot) = roles.get_mut("plotArea") {
            plot.fill_colors = None;
            plot.fill_paints = None;
            plot.fill_hidden = Some(true);
            plot.fill_paint_authored = Some(true);
        }
    }
    let style = match chart.legacy_chart_style.unwrap_or(2) {
        1..=32 => {
            roles.remove("chartArea");
            return;
        }
        33..=40 => {
            let mut style = roles.remove("chartArea").unwrap_or_default();
            style.line_width_emu = Some(9_525);
            style.line_hidden = None;
            style
        }
        _ => {
            let mut style = roles.remove("chartArea").unwrap_or_default();
            style.line_width_emu = None;
            style.line_hidden = Some(true);
            style.line_paint_authored = Some(true);
            style
        }
    };
    roles.insert("chartArea".to_string(), style);
}

/// Preserve PowerPoint's ChartEx chart-space no-line carrier. In PowerPoint-
/// produced files a present `cx:chartSpace/cx:spPr` that authors another
/// component but omits `a:ln` replaces an `allowNoLineOverride` chart-style
/// outline with no line. MS-ODRAWXML §2.8.4.8 permits the replacement but does
/// not prescribe its carrier; the boundary is established by PowerPoint output
/// with a fill-only chartSpace and its exported PDF. Keep this host delta out of
/// the shared DrawingML cascade: Excel's analogous component omissions inherit,
/// and an explicit local line remains authoritative everywhere.
pub(crate) fn apply_powerpoint_chartex_chart_space_frame(
    chart_space: Node<'_, '_>,
    chart: &mut ChartModel,
) {
    let chart_space_shape = chart_space
        .children()
        .find(|node| node.is_element() && node.tag_name().name() == "spPr");
    let fill_only = chart_space_shape.is_some_and(|shape| {
        let mut has_fill = false;
        let mut has_line = false;
        for node in shape.children().filter(Node::is_element) {
            match node.tag_name().name() {
                "noFill" | "solidFill" | "gradFill" | "pattFill" | "blipFill" | "grpFill" => {
                    has_fill = true;
                }
                "ln" => has_line = true,
                _ => {}
            }
        }
        has_fill && !has_line
    });
    let allows_no_line = chart
        .chart_style_roles
        .as_ref()
        .and_then(|roles| roles.get("chartArea"))
        .and_then(|role| role.allow_no_line_override)
        == Some(true);
    if fill_only && allows_no_line {
        chart.chart_border_hidden = Some(true);
        chart.chart_border_paint_authored = Some(true);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use ooxml_common::chart::ColorResolver;

    struct NoTheme;
    impl ColorResolver for NoTheme {
        fn resolve_solid_fill(&self, _node: roxmltree::Node<'_, '_>) -> Option<String> {
            None
        }
    }

    #[test]
    fn powerpoint_radial_plot_background_is_transparent_for_light_and_dark_styles() {
        for (kind, transparent) in [("pie", true), ("doughnut", true), ("pie3D", false)] {
            for style in [34, 48] {
                let xml = format!(
                    r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:style val="{style}"/><c:chart><c:plotArea><c:{kind}Chart><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:{kind}Chart><c:spPr><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></a:ln></c:spPr></c:plotArea></c:chart></c:chartSpace>"#
                );
                let chart =
                    crate::chart::parse_legacy_chart(&xml, &std::collections::HashMap::new())
                        .unwrap()
                        .chart;
                let role = &chart.classic_chart_style_roles.as_ref().unwrap()["plotArea"];
                assert_eq!(
                    role.fill_hidden == Some(true),
                    transparent,
                    "{kind} style={style}"
                );
                assert_eq!(
                    chart.powerpoint_radial_plot_frame == Some(true),
                    transparent
                );
                assert_eq!(chart.plot_area_bg.as_deref(), Some("123456"));
                assert_eq!(chart.plot_area_line_color.as_deref(), Some("ABCDEF"));
            }
        }
    }

    #[test]
    fn powerpoint_plot_background_defaults_preserve_authored_paint() {
        for style in std::iter::once(None).chain((1..=48).map(Some)) {
            for fill in [
                "",
                "<a:solidFill><a:srgbClr val=\"123456\"/></a:solidFill>",
                "<a:noFill/>",
            ] {
                let style_xml = style
                    .map(|v| format!("<c:style val=\"{v}\"/>"))
                    .unwrap_or_default();
                let shape = if fill.is_empty() {
                    String::new()
                } else {
                    format!("<c:spPr>{fill}</c:spPr>")
                };
                let xml = format!(
                    r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">{style_xml}<c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart>{shape}</c:plotArea></c:chart></c:chartSpace>"#
                );
                let element =
                    crate::chart::parse_legacy_chart(&xml, &std::collections::HashMap::new())
                        .unwrap();
                let chart = element.chart;
                let role = &chart.classic_chart_style_roles.as_ref().unwrap()["plotArea"];
                assert_eq!(
                    role.fill_hidden == Some(true),
                    style.unwrap_or(2) <= 32,
                    "style={style:?}"
                );
                if fill.contains("solidFill") {
                    assert_eq!(chart.plot_area_bg.as_deref(), Some("123456"));
                } else if fill.contains("noFill") {
                    assert_eq!(chart.plot_area_fill_hidden, Some(true));
                }
                if style.unwrap_or(2) > 32 {
                    assert!(role.fill_colors.as_ref().is_some_and(|v| v[0].is_some()));
                }
            }
        }
    }

    #[test]
    fn powerpoint_plot_background_preserves_linked_fill() {
        let xml = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:style val="2"/><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#;
        let style = r#"<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><cs:plotArea><cs:spPr><a:solidFill><a:srgbClr val="123456"/></a:solidFill></cs:spPr></cs:plotArea></cs:chartStyle>"#;
        let document = roxmltree::Document::parse(xml).unwrap();
        let mut chart = ooxml_common::chart::parse_chart_part(
            document.root_element(),
            &ooxml_common::chart::ChartParseContext {
                color_resolver: Some(&NoTheme),
                style_xml: Some(style),
                ..Default::default()
            },
        )
        .unwrap();
        let linked_before = serde_json::to_value(&chart.chart_style_roles).unwrap();
        assert!(linked_before.to_string().contains("123456"));
        apply_powerpoint_classic_chart_space_frame(&mut chart);
        assert_eq!(
            serde_json::to_value(&chart.chart_style_roles).unwrap(),
            linked_before
        );
        assert_eq!(
            chart.classic_chart_style_roles.as_ref().unwrap()["plotArea"].fill_hidden,
            Some(true)
        );
    }

    #[test]
    fn theme_less_package_keeps_default_numeric_role_paint() {
        let xml = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:style val="34"/><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#;
        let document = roxmltree::Document::parse(xml).unwrap();
        let mut chart = ooxml_common::chart::parse_chart_part(
            document.root_element(),
            &ooxml_common::chart::ChartParseContext {
                color_resolver: Some(&NoTheme),
                ..Default::default()
            },
        )
        .expect("classic chart");
        apply_powerpoint_classic_chart_space_frame(&mut chart);
        let roles = chart.classic_chart_style_roles.as_ref().unwrap();
        assert!(roles["chartArea"]
            .fill_colors
            .as_ref()
            .is_some_and(|v| v[0].is_some()));
        assert!(roles["chartArea"]
            .line_colors
            .as_ref()
            .is_some_and(|v| v[0].is_some()));
        assert!(roles["plotArea"]
            .fill_colors
            .as_ref()
            .is_some_and(|v| v[0].is_some()));
        assert!(roles["categoryAxis"]
            .line_colors
            .as_ref()
            .is_some_and(|v| v[0].is_some()));
        assert!(roles["dataPoint"]
            .fill_colors
            .as_ref()
            .is_some_and(|v| v[0].is_some()));
    }

    #[test]
    fn powerpoint_chartex_fill_only_chart_space_suppresses_an_allowed_style_outline() {
        const CX_NS: &str = "http://schemas.microsoft.com/office/drawing/2014/chartex";
        const CS_NS: &str = "http://schemas.microsoft.com/office/drawing/2012/chartStyle";
        const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
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
              <cs:chartArea mods="allowNoLineOverride"><cs:spPr>
                <a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln>
              </cs:spPr></cs:chartArea>
            </cs:chartStyle>"#
        );
        let document = roxmltree::Document::parse(&xml).unwrap();
        let mut chart = ooxml_common::chart::parse_chartex_part(
            document.root_element(),
            &ooxml_common::chart::ChartParseContext {
                color_resolver: Some(&NoTheme),
                style_xml: Some(&style),
                color_style_xml: None,
                ..Default::default()
            },
        )
        .expect("ChartEx chart");

        assert_eq!(chart.chart_border_hidden, None);
        apply_powerpoint_chartex_chart_space_frame(document.root_element(), &mut chart);
        assert_eq!(chart.chart_border_hidden, Some(true));
        assert_eq!(chart.chart_border_paint_authored, Some(true));

        let mut control = ooxml_common::chart::parse_chartex_part(
            document.root_element(),
            &ooxml_common::chart::ChartParseContext {
                color_resolver: Some(&NoTheme),
                style_xml: Some(&style.replace(" mods=\"allowNoLineOverride\"", "")),
                color_style_xml: None,
                ..Default::default()
            },
        )
        .expect("control ChartEx chart");
        apply_powerpoint_chartex_chart_space_frame(document.root_element(), &mut control);
        assert_eq!(control.chart_border_hidden, None);
        assert_eq!(control.chart_border_paint_authored, None);

        let empty_xml = xml.replace(
            "<cx:spPr><a:solidFill><a:schemeClr val=\"bg1\"/></a:solidFill></cx:spPr>",
            "<cx:spPr/>",
        );
        let empty_document = roxmltree::Document::parse(&empty_xml).unwrap();
        let mut empty = ooxml_common::chart::parse_chartex_part(
            empty_document.root_element(),
            &ooxml_common::chart::ChartParseContext {
                color_resolver: Some(&NoTheme),
                style_xml: Some(&style),
                color_style_xml: None,
                ..Default::default()
            },
        )
        .expect("empty chart-space properties");
        apply_powerpoint_chartex_chart_space_frame(empty_document.root_element(), &mut empty);
        assert_eq!(empty.chart_border_hidden, None);
        assert_eq!(empty.chart_border_paint_authored, None);
    }
}
