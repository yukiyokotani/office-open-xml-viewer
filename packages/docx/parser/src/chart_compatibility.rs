//! Isolated Word deltas over ECMA-376 classic chart-space defaults.

use ooxml_common::chart::{ChartCartesianAutoLayoutProfile, ChartModel};
use std::collections::BTreeMap;

/// Apply Word's observed host delta to the normative built-in `chartArea`
/// recipe. ECMA-376 §21.2.3.46 defines the style 1..48 fill/line table; a
/// complete Word-produced default-theme matrix additionally found 10pt rounded
/// corners and, for omitted style and 1..40, transparent chart areas with a
/// 0.5pt outline. The outline paint remains the theme-resolved ECMA recipe;
/// default-theme output is #898989 without hard-coding that color across
/// custom themes. Styles 41..48 retain Table 3's chart-area fill and suppress
/// the outline. Linked style roles remain a separate, higher layer.
pub(crate) fn apply_word_classic_chart_space_frame(chart: &mut ChartModel) {
    // ECMA-376 leaves automatic plot placement to the consumer. Across a
    // Word-produced matrix covering all built-in numeric styles 1..48, classic
    // 2-D vertical column charts used the same Word-specific title/category
    // reserves. Manual layout, horizontal bars, 3-D charts, combination charts,
    // and non-column families are deliberately excluded: the first three are
    // counterexamples and the remaining classes are outside that evidence.
    let is_single_vertical_bar_group = chart.plot_groups.as_deref().is_some_and(|groups| {
        groups.len() == 1
            && groups[0].kind == "bar"
            && groups[0].bar_direction.as_deref().unwrap_or("col") == "col"
    });
    chart.cartesian_auto_layout_profile = if chart.three_d.is_none()
        && chart.plot_area_manual_layout.is_none()
        && is_single_vertical_bar_group
        && matches!(
            chart.chart_type.as_str(),
            "clusteredBar" | "stackedBar" | "stackedBarPct"
        ) {
        Some(ChartCartesianAutoLayoutProfile::WordClassicColumn)
    } else {
        None
    };

    chart.rounded_corners.get_or_insert(true);

    let style_number = chart.legacy_chart_style.unwrap_or(2);
    let style = chart
        .classic_chart_style_roles
        .get_or_insert_with(BTreeMap::new)
        .entry("chartArea".to_string())
        .or_default();
    if style_number <= 40 {
        style.fill_paints = None;
        style.fill_colors = None;
        style.fill_hidden = Some(true);
        style.fill_paint_authored = Some(true);
        style.fill_no_style = None;
        style.line_width_emu = Some(6_350);
        style.line_hidden = None;
        style.line_no_style = None;
    } else {
        style.line_width_emu = None;
        style.line_hidden = Some(true);
        style.line_paint_authored = Some(true);
        style.line_no_style = None;
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

    fn word_chart(chart_children: &str, plot_children: &str) -> ChartModel {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart>{chart_children}<c:plotArea>{plot_children}</c:plotArea></c:chart></c:chartSpace>"#
        );
        let document = roxmltree::Document::parse(&xml).expect("chart XML");
        let mut chart = ooxml_common::chart::parse_chart_part(document.root_element(), &NoTheme)
            .expect("classic chart");
        apply_word_classic_chart_space_frame(&mut chart);
        chart
    }

    fn bar_group(direction: &str) -> String {
        format!(
            r#"<c:barChart><c:barDir val="{direction}"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart>"#
        )
    }

    #[test]
    fn word_chart_keeps_authored_axis_code_without_worksheet_resolver() {
        let chart = word_chart(
            "",
            &format!(
                r#"{}<c:valAx><c:axId val="100"/><c:axPos val="l"/><c:numFmt formatCode="0.00" sourceLinked="1"/></c:valAx>"#,
                bar_group("col"),
            ),
        );
        assert_eq!(chart.val_axis_format_code.as_deref(), Some("0.00"));
    }

    #[test]
    fn theme_less_package_keeps_default_numeric_role_paint() {
        let xml = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:style val="2"/><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#;
        let document = roxmltree::Document::parse(xml).unwrap();
        let mut chart = ooxml_common::chart::parse_chart_part(document.root_element(), &NoTheme)
            .expect("classic chart");
        apply_word_classic_chart_space_frame(&mut chart);
        let roles = chart.classic_chart_style_roles.as_ref().unwrap();
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
    fn opts_only_automatic_single_group_2d_vertical_columns_into_word_layout() {
        let automatic = word_chart("", &bar_group("col"));
        assert_eq!(
            automatic.cartesian_auto_layout_profile,
            Some(ChartCartesianAutoLayoutProfile::WordClassicColumn)
        );

        let manual = word_chart(
            "",
            &format!(
                r#"<c:layout><c:manualLayout><c:x val="0.1"/><c:y val="0.1"/><c:w val="0.8"/><c:h val="0.8"/></c:manualLayout></c:layout>{}"#,
                bar_group("col")
            ),
        );
        assert_eq!(manual.cartesian_auto_layout_profile, None);

        let horizontal = word_chart("", &bar_group("bar"));
        assert_eq!(horizontal.cartesian_auto_layout_profile, None);

        let three_d = word_chart(
            r#"<c:view3D><c:rotX val="15"/></c:view3D>"#,
            r#"<c:bar3DChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:bar3DChart>"#,
        );
        assert_eq!(three_d.cartesian_auto_layout_profile, None);

        let combination = word_chart(
            "",
            &format!(
                r#"{}<c:lineChart><c:ser><c:idx val="1"/><c:order val="1"/><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart>"#,
                bar_group("col")
            ),
        );
        assert_eq!(combination.cartesian_auto_layout_profile, None);
    }
}
