//! Isolated Excel deltas over ECMA-376 classic chart-space defaults.

#[cfg(test)]
use ooxml_common::chart::ChartExElementStyle;
use ooxml_common::chart::ChartModel;
use std::collections::BTreeMap;

/// Apply Excel's observed host delta to the normative built-in `chartArea`
/// recipe. ECMA-376 §21.2.3.46 defines the style 1..48 recipes;
/// Excel-produced samples covering every style and the omitted style showed
/// 10pt rounded corners throughout.
/// Styles 1..40 use the light theme chart-area recipe with a 1pt outline;
/// styles 41..48 use the dark theme recipe without an outline. Default-theme
/// output is white/#898989 and black respectively, but paint remains resolved
/// from the document theme. Omitted style behaves as style 2.
pub(crate) fn apply_excel_classic_chart_space_frame(chart: &mut ChartModel) {
    chart.rounded_corners.get_or_insert(true);

    let style = chart
        .classic_chart_style_roles
        .get_or_insert_with(BTreeMap::new)
        .entry("chartArea".to_string())
        .or_default();
    if chart.legacy_chart_style.unwrap_or(2) <= 40 {
        style.line_width_emu = Some(12_700);
        style.line_hidden = None;
    } else {
        style.line_width_emu = None;
        style.line_hidden = Some(true);
        style.line_paint_authored = Some(true);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use ooxml_common::chart::ColorResolver;

    struct NoColors;

    impl ColorResolver for NoColors {
        fn resolve_solid_fill(&self, _node: roxmltree::Node<'_, '_>) -> Option<String> {
            None
        }
    }

    fn chart(style: Option<u8>, rounded: Option<bool>) -> ChartModel {
        let style = style
            .map(|value| format!(r#"<c:style val="{value}"/>"#))
            .unwrap_or_default();
        let rounded = rounded
            .map(|value| format!(r#"<c:roundedCorners val="{}"/>"#, u8::from(value)))
            .unwrap_or_default();
        let xml = format!(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">{style}{rounded}<c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#
        );
        let document = roxmltree::Document::parse(&xml).expect("chart XML");
        let mut chart = ooxml_common::chart::parse_chart_part(document.root_element(), &NoColors)
            .expect("classic chart");
        apply_excel_classic_chart_space_frame(&mut chart);
        chart
    }

    #[test]
    fn style_boundary_selects_excel_light_and_dark_frames() {
        let light = chart(Some(40), None);
        let light_frame = &light.classic_chart_style_roles.as_ref().unwrap()["chartArea"];
        assert_eq!(light.rounded_corners, Some(true));
        assert_eq!(
            light_frame.fill_colors.as_deref(),
            Some(&[Some("FFFFFF".to_string())][..])
        );
        assert!(light_frame
            .line_colors
            .as_ref()
            .is_some_and(|colors| colors.iter().any(|color| color.is_some())));
        assert_eq!(light_frame.line_width_emu, Some(12_700));
        let roles = light.classic_chart_style_roles.as_ref().unwrap();
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

        let dark = chart(Some(41), None);
        let dark_frame = &dark.classic_chart_style_roles.as_ref().unwrap()["chartArea"];
        assert!(dark_frame
            .fill_colors
            .as_ref()
            .is_some_and(|colors| colors.iter().any(|color| color.is_some())));
        assert_eq!(dark_frame.line_hidden, Some(true));
    }

    #[test]
    fn omitted_style_and_explicit_square_corners_keep_their_precedence() {
        let omitted = chart(None, None);
        let frame = &omitted.classic_chart_style_roles.as_ref().unwrap()["chartArea"];
        assert_eq!(frame.line_width_emu, Some(12_700));
        assert_eq!(chart(Some(2), Some(false)).rounded_corners, Some(false));

        let mut linked = chart(Some(40), None);
        linked
            .chart_style_roles
            .get_or_insert_with(BTreeMap::new)
            .insert(
                "chartArea".to_string(),
                ChartExElementStyle {
                    line_colors: Some(vec![Some("FF0000".to_string())]),
                    line_width_emu: Some(25_400),
                    ..Default::default()
                },
            );
        apply_excel_classic_chart_space_frame(&mut linked);
        let linked_frame = &linked.chart_style_roles.as_ref().unwrap()["chartArea"];
        assert_eq!(
            linked_frame.line_colors.as_deref(),
            Some(&[Some("FF0000".to_string())][..])
        );
        assert_eq!(linked_frame.line_width_emu, Some(25_400));
    }
}
