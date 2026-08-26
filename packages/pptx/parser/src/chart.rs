//! Chart element parsing: the legacy DrawingML chart (`c:` namespace) and the
//! newer chartEx (`cx:` namespace) parsers, plus the pptx `ColorResolver` the
//! shared `ooxml_common::chart` helpers use to resolve `<a:solidFill>` colours.
//! Extracted verbatim from `lib.rs`. The general colour grammar
//! (`parse_color_node`) stays in `lib.rs` and is imported here for the
//! `PptxColorResolver`; both chart parsers now delegate their structure walk to
//! `ooxml_common::chart`.

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
    let style_xml = style_xml.filter(|style| parse_preflighted_pptx_xml(style).is_ok());
    let color_style_xml = color_style_xml.filter(|style| parse_preflighted_pptx_xml(style).is_ok());
    let mut chart = ooxml_common::chart::parse_chart_part_with_style_parts_and_images(
        root,
        &resolver,
        style_xml,
        color_style_xml,
        image_resolver,
    )?;
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
    let style_xml = style_xml.filter(|style| parse_preflighted_pptx_xml(style).is_ok());
    let color_style_xml = color_style_xml.filter(|style| parse_preflighted_pptx_xml(style).is_ok());
    // chartEx (waterfall/boxWhisker/…) reads its title font size from the
    // associated chartStyle part when the `<cx:title>` itself carries none.
    let chart = ooxml_common::chart::parse_chartex_part_with_style_parts_and_images(
        root,
        &resolver,
        style_xml,
        color_style_xml,
        image_resolver,
    )?;
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
