//! Shape / tree / picture / media / table / connector / SmartArt assembly:
//! turns a `<p:spTree>` and its children into `SlideElement`s. Extracted
//! verbatim from `lib.rs`. Pulls fill/stroke/effect/geometry parsing from
//! `fill`, text-body parsing from `text`, chart parsing from `chart`, and the
//! shared XML/zip helpers + `TableStyleDef` from the crate root.

use crate::chart::{parse_chartex_with_images, parse_legacy_chart_with_style_parts_and_images};
use crate::fill::{
    line_properties_to_stroke, parse_blip_alpha, parse_color_node, parse_cust_geom,
    parse_effect_lst, parse_fill, parse_scene3d, parse_sp3d, parse_stroke,
    parse_style_matrix_effects, parse_style_matrix_fill_from_source, parse_table_style_fill,
    parse_xfrm, EffectLst,
};
use crate::master::{InheritedShapeGeometry, LayoutPlaceholders};
use crate::text::{
    empty_level_bullets, parse_text_body, LevelBullets, LevelFontSizes, LevelIndents, ShapeKind,
};
use crate::theme::{PptxRawSchemeResolver, PptxSchemeResolver, PptxThemeSource};
use crate::types::*;
use crate::{
    attr, attr_f64, attr_i64, attr_r, child, read_zip_head, read_zip_str, resolve_path,
    table_style_presets, PptxZip, ResolvedTableCellStyle, TableCellBorderStyle, TableLineStyle,
    TablePartStyle, TableStyleDef, TableStyleFlags, TableTextStyle,
};
use ooxml_common::blip::{mime_from_ext, parse_blip_duotone, parse_src_rect, svg_blip_rid};
use ooxml_common::depth::DepthGuard;
use ooxml_common::line::{
    parse_line_properties, LineDash, LineEnd, LineJoin, LinePaint, LineProperties,
};
use ooxml_common::ns::{is_diagram_uri, is_pml_ole_uri};
use ooxml_common::rels::relationship_part_path;
use ooxml_common::units::EMU_PER_PX_96DPI;
use std::collections::HashMap;

/// ECMA-376 Part 3 §9.3 "given application configuration" for the pptx parser:
/// the extension namespaces whose `<mc:Choice>` content the parser actually
/// renders. chartEx (any dated version — the chart path renders them via the
/// existing `uri.contains("chartex")` gate) and the Microsoft 2010 drawing
/// extensions (a14) that wrap ordinary `<p:sp>`/`<p:pic>`. Deliberately EXCLUDES
/// PowerPoint-2010 (`p14`, ink `<p:contentPart>`) and VML (`v`), which we cannot
/// render, so a Choice requiring them yields the Fallback preview.
pub(crate) fn pptx_understands_ns(ns: &str) -> bool {
    ns.contains("chartex") || ns == "http://schemas.microsoft.com/office/drawing/2010/main"
}

/// Read the chartStyle part (`styleN.xml`) associated with a chart part at
/// `chart_path` (e.g. `ppt/charts/chart1.xml`), following that part's own
/// relationships (`ppt/charts/_rels/chart1.xml.rels`) to the
/// `.../2011/relationships/chartStyle` target. Returns `None` when the chart
/// has no chartStyle relationship or the part cannot be read (the chartEx
/// title then falls back to its inline size, or the renderer's default).
struct ChartRelatedParts {
    style_xml: Option<String>,
    color_style_xml: Option<String>,
    image_relationships: ooxml_common::chart::ChartImageRelationships,
}

fn load_chart_related_parts(zip: &mut PptxZip, chart_path: &str) -> ChartRelatedParts {
    let mut result = ChartRelatedParts {
        style_xml: None,
        color_style_xml: None,
        image_relationships: Default::default(),
    };
    let rels_path = relationship_part_path(chart_path);
    let Ok(rels_xml) = read_zip_str(zip, &rels_path) else {
        return result;
    };
    let relationships = ooxml_common::rels::parse_rels(&rels_xml);
    result.image_relationships.insert_parsed_relationships(
        ooxml_common::chart::ChartImageSource::Chart,
        chart_path,
        &relationships,
    );
    let base_dir = chart_path.rsplit_once('/').map_or("", |(dir, _)| dir);
    let internal_target = |suffix: &str| {
        relationships.values().find(|relationship| {
            relationship.mode == ooxml_common::rels::TargetMode::Internal
                && relationship
                    .relationship_type
                    .as_deref()
                    .is_some_and(|kind| kind.ends_with(suffix))
        })
    };
    if let Some(style_relationship) =
        internal_target(ooxml_common::chart::CHART_STYLE_REL_TYPE_SUFFIX)
    {
        let style_path = resolve_path(base_dir, &style_relationship.target);
        result.style_xml = read_zip_str(zip, &style_path).ok();
        let style_rels_path = relationship_part_path(&style_path);
        if let Ok(style_rels_xml) = read_zip_str(zip, &style_rels_path) {
            let style_relationships = ooxml_common::rels::parse_rels(&style_rels_xml);
            result.image_relationships.insert_parsed_relationships(
                ooxml_common::chart::ChartImageSource::Style,
                &style_path,
                &style_relationships,
            );
        }
    }
    if let Some(color_relationship) =
        internal_target(ooxml_common::chart::CHART_COLOR_STYLE_REL_TYPE_SUFFIX)
    {
        let color_path = resolve_path(base_dir, &color_relationship.target);
        result.color_style_xml = read_zip_str(zip, &color_path).ok();
    }
    result
}

#[cfg(test)]
mod chartex_sidecar_package_tests {
    use std::io::{Cursor, Write};

    #[test]
    fn pptx_package_loads_chartex_style_and_color_sidecars() {
        let content_types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#;
        let root_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#;
        let presentation_xml = r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#;
        let presentation_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSlide" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>"#;
        let slide_xml = r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><p:cSld><p:spTree><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Chart 1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="3000000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart r:id="rIdChart"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>"#;
        let slide_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.microsoft.com/office/2014/relationships/chartEx" Target="../charts/chartEx1.xml"/></Relationships>"#;
        let chart_xml = r#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:lvl ptCount="1"><cx:pt idx="0">1</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="boxWhisker"/></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>"#;
        let rels_xml = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdExternalStyle" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="https://example.invalid/style.xml" TargetMode="External"/><Relationship Id="rIdStyle" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/><Relationship Id="rIdColors" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/></Relationships>"#;
        let style_xml = r#"<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><cs:dataPoint><cs:fillRef idx="1"><cs:styleClr val="auto"/></cs:fillRef><cs:spPr><a:pattFill prst="diagCross"><a:fgClr><a:schemeClr val="phClr"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill></cs:spPr></cs:dataPoint><cs:dataPointMarker><cs:fillRef idx="1"><cs:styleClr val="auto"/></cs:fillRef></cs:dataPointMarker><cs:dataLabelCallout><cs:defRPr><a:noFill/></cs:defRPr><cs:bodyPr/></cs:dataLabelCallout><cs:trendlineLabel><cs:defRPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill></cs:defRPr></cs:trendlineLabel></cs:chartStyle>"#;
        let colors_xml = r#"<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle"><a:srgbClr val="336699"/></cs:colorStyle>"#;
        let theme_xml = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="Theme"><a:themeElements><a:fmtScheme name="Theme"><a:fillStyleLst><a:blipFill><a:blip r:embed="rIdThemeMarker"/><a:stretch/></a:blipFill></a:fillStyleLst><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme></a:themeElements></a:theme>"#;

        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let options = zip::write::SimpleFileOptions::default();
            for (path, xml) in [
                ("[Content_Types].xml", content_types),
                ("_rels/.rels", root_rels),
                ("ppt/presentation.xml", presentation_xml),
                ("ppt/_rels/presentation.xml.rels", presentation_rels),
                ("ppt/slides/slide1.xml", slide_xml),
                ("ppt/slides/_rels/slide1.xml.rels", slide_rels),
                ("ppt/charts/chartEx1.xml", chart_xml),
                ("ppt/charts/_rels/chartEx1.xml.rels", rels_xml),
                ("ppt/charts/style1.xml", style_xml),
                (
                    "ppt/charts/_rels/style1.xml.rels",
                    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdMarker" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/style-marker.png"/></Relationships>"#,
                ),
                ("ppt/charts/colors1.xml", colors_xml),
                ("ppt/theme/theme1.xml", theme_xml),
                (
                    "ppt/theme/_rels/theme1.xml.rels",
                    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdThemeMarker" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/theme-marker.png"/></Relationships>"#,
                ),
                ("ppt/media/theme-marker.png", "png"),
                ("ppt/media/style-marker.png", "png"),
            ] {
                writer.start_file(path, options).unwrap();
                writer.write_all(xml.as_bytes()).unwrap();
            }
            writer.finish().unwrap();
        }

        let json = crate::parse_pptx_native(&bytes).expect("full PPTX package parses");
        let presentation: serde_json::Value =
            serde_json::from_str(&json).expect("presentation JSON");
        let element = &presentation["slides"][0]["elements"][0];
        assert_eq!(element["id"], serde_json::Value::String("2".to_string()));
        let chart = &element["chart"];
        assert_eq!(
            chart["chartexColorPalette"][0],
            serde_json::Value::String("336699".to_string()),
        );
        assert_eq!(
            chart["chartStyleRoles"]["dataLabelCallout"]["fontHidden"],
            serde_json::Value::Bool(true),
        );
        assert_eq!(
            chart["chartStyleRoles"]["dataLabelCallout"]["textBodyAuthored"],
            serde_json::Value::Bool(true),
        );
        assert_eq!(
            chart["chartStyleRoles"]["trendlineLabel"]["fontColor"],
            serde_json::Value::String("112233".to_string()),
        );
        assert_eq!(
            chart["chartexDataPointStyle"]["fillPaints"][0],
            serde_json::json!({
                "fillType": "pattern",
                "fg": "336699",
                "bg": "FFFFFF",
                "preset": "diagCross",
            }),
        );
        assert_eq!(
            chart["chartStyleRoles"]["dataPointMarker"]["fillPaints"][0],
            serde_json::json!({
                "fillType": "image",
                "imagePath": "ppt/media/theme-marker.png",
                "mimeType": "image/png",
                "stretch": true,
            }),
        );
    }

    #[test]
    fn pptx_package_loads_classic_chart_style_roles() {
        let content_types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#;
        let root_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#;
        let presentation_xml = r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#;
        let presentation_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSlide" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>"#;
        let slide_xml = r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><p:cSld><p:spTree><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Chart 1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="3000000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>"#;
        let slide_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#;
        let chart_xml = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:lineChart><c:grouping val="standard"/><c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser><c:dropLines/></c:lineChart></c:plotArea></c:chart></c:chartSpace>"#;
        let rels_xml = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyle" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/><Relationship Id="rIdColors" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/></Relationships>"#;
        let style_xml = r#"<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><cs:dropLine><cs:spPr><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></cs:spPr></cs:dropLine></cs:chartStyle>"#;
        let colors_xml = r#"<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle"><a:srgbClr val="336699"/></cs:colorStyle>"#;

        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let options = zip::write::SimpleFileOptions::default();
            for (path, xml) in [
                ("[Content_Types].xml", content_types),
                ("_rels/.rels", root_rels),
                ("ppt/presentation.xml", presentation_xml),
                ("ppt/_rels/presentation.xml.rels", presentation_rels),
                ("ppt/slides/slide1.xml", slide_xml),
                ("ppt/slides/_rels/slide1.xml.rels", slide_rels),
                ("ppt/charts/chart1.xml", chart_xml),
                ("ppt/charts/_rels/chart1.xml.rels", rels_xml),
                ("ppt/charts/style1.xml", style_xml),
                ("ppt/charts/colors1.xml", colors_xml),
            ] {
                writer.start_file(path, options).unwrap();
                writer.write_all(xml.as_bytes()).unwrap();
            }
            writer.finish().unwrap();
        }

        let json = crate::parse_pptx_native(&bytes).expect("full PPTX package parses");
        let presentation: serde_json::Value =
            serde_json::from_str(&json).expect("presentation JSON");
        let chart = &presentation["slides"][0]["elements"][0]["chart"];
        assert_eq!(
            chart["chartStyleRoles"]["dropLine"]["lineColors"][0],
            serde_json::Value::String("336699".to_string()),
        );
    }

    #[test]
    fn pptx_package_resolves_picture_marker_relationships() {
        let content_types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#;
        let root_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#;
        let presentation_xml = r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#;
        let presentation_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSlide" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>"#;
        let slide_xml = r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><p:cSld><p:spTree><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Chart 1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="3000000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>"#;
        let slide_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#;
        let chart_xml = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart><c:plotArea><c:lineChart><c:ser><c:idx val="0"/><c:order val="0"/><c:marker><c:symbol val="picture"/><c:spPr><a:blipFill><a:blip r:embed="rIdMarker"/><a:stretch/></a:blipFill></c:spPr></c:marker><c:cat><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart></c:plotArea></c:chart></c:chartSpace>"#;
        let chart_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdMarker" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/marker.png"/></Relationships>"#;

        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let options = zip::write::SimpleFileOptions::default();
            for (path, content) in [
                ("[Content_Types].xml", content_types.as_bytes()),
                ("_rels/.rels", root_rels.as_bytes()),
                ("ppt/presentation.xml", presentation_xml.as_bytes()),
                (
                    "ppt/_rels/presentation.xml.rels",
                    presentation_rels.as_bytes(),
                ),
                ("ppt/slides/slide1.xml", slide_xml.as_bytes()),
                ("ppt/slides/_rels/slide1.xml.rels", slide_rels.as_bytes()),
                ("ppt/charts/chart1.xml", chart_xml.as_bytes()),
                ("ppt/charts/_rels/chart1.xml.rels", chart_rels.as_bytes()),
                ("ppt/media/marker.png", b"png".as_slice()),
            ] {
                writer.start_file(path, options).unwrap();
                writer.write_all(content).unwrap();
            }
            writer.finish().unwrap();
        }

        let json = crate::parse_pptx_native(&bytes).expect("full PPTX package parses");
        let presentation: serde_json::Value = serde_json::from_str(&json).unwrap();
        assert_eq!(
            presentation["slides"][0]["elements"][0]["chart"]["series"][0]["markerFillPaint"],
            serde_json::json!({
                "fillType": "image",
                "imagePath": "ppt/media/marker.png",
                "mimeType": "image/png",
                "stretch": true,
            }),
        );
    }
}

fn load_chart_user_shapes_xml(
    zip: &mut PptxZip,
    chart_path: &str,
    chart_xml: &str,
) -> Option<String> {
    let chart_doc = crate::parse_preflighted_pptx_xml(chart_xml).ok()?;
    let rid = chart_doc
        .root_element()
        .descendants()
        .find(|node| node.is_element() && node.tag_name().name() == "userShapes")?
        .attributes()
        .find(|attribute| {
            attribute.name() == "id"
                && attribute
                    .namespace()
                    .is_some_and(|namespace| namespace.ends_with("/relationships"))
        })?
        .value();
    let dir = chart_path.rsplit_once('/').map_or("", |(dir, _)| dir);
    let rels_path = relationship_part_path(chart_path);
    let rels_xml = read_zip_str(zip, &rels_path).ok()?;
    let target = ooxml_common::rels::parse_rels(&rels_xml)
        .get(rid)?
        .target
        .clone();
    let user_shapes_path = resolve_path(dir, &target);
    read_zip_str(zip, &user_shapes_path).ok()
}

/// OOXML spec default positions for common placeholder types.
/// Values are in EMU, assuming a 9144000×6858000 slide (10"×7.5").
pub(crate) fn default_placeholder_transform(ph_type: &str) -> Transform {
    match ph_type {
        "title" | "ctrTitle" => Transform {
            x: 457200,
            y: 274638,
            cx: 8229600,
            cy: 1143000,
            rot: 0.0,
            flip_h: false,
            flip_v: false,
        },
        "subTitle" => Transform {
            x: 457200,
            y: 1600200,
            cx: 8229600,
            cy: 899160,
            rot: 0.0,
            flip_h: false,
            flip_v: false,
        },
        "dt" => Transform {
            x: 0,
            y: 6261600,
            cx: 2286000,
            cy: 596900,
            rot: 0.0,
            flip_h: false,
            flip_v: false,
        },
        "ftr" => Transform {
            x: 2972400,
            y: 6261600,
            cx: 3086100,
            cy: 596900,
            rot: 0.0,
            flip_h: false,
            flip_v: false,
        },
        "sldNum" => Transform {
            x: 6629400,
            y: 6261600,
            cx: 2057400,
            cy: 596900,
            rot: 0.0,
            flip_h: false,
            flip_v: false,
        },
        // "body" and everything else: full-width content area below title
        _ => Transform {
            x: 457200,
            y: 1600200,
            cx: 8229600,
            cy: 4525963,
            rot: 0.0,
            flip_h: false,
            flip_v: false,
        },
    }
}

/// Returns true if the node contains a `p:ph` descendant.
pub(crate) fn is_placeholder(node: roxmltree::Node<'_, '_>) -> bool {
    node.descendants()
        .any(|n| n.is_element() && n.tag_name().name() == "ph")
}

/// Non-visual drawing props pulled from a shape's own `<p:cNvPr>`.
#[derive(Default)]
pub(crate) struct CNvPrInfo {
    pub(crate) id: Option<String>,
    pub(crate) name: Option<String>,
    /// Shape-level hyperlink target resolved via `rels` (URL for external,
    /// internal part name for a slide jump). ECMA-376 §21.1.2.3.5.
    pub(crate) hyperlink: Option<String>,
    /// Raw `<a:hlinkClick @action>` (e.g. "ppaction://hlinksldjump").
    pub(crate) hyperlink_action: Option<String>,
}

/// Pull `id` / `name` / shape-level `hlinkClick` from a node's own
/// `<p:nvSpPr><p:cNvPr>` (or `nvCxnSpPr` etc.). Returns defaults when the
/// wrapper is missing — every field is optional in the JSON output.
///
/// ECMA-376 §21.1.2.3.5: a `<p:cNvPr>` may carry an `<a:hlinkClick @r:id
/// [@action]>`, making the whole shape a click target. `@r:id` resolves via the
/// slide `rels`; `@action` (when a "ppaction://..." verb) marks an internal
/// navigation, mirroring the text-run hyperlink parse in `parse_run`.
pub(crate) fn read_cnv_pr(
    sp_node: roxmltree::Node<'_, '_>,
    rels: &HashMap<String, String>,
) -> CNvPrInfo {
    if let Some(cnv) = own_cnv_pr(sp_node) {
        let hlink_click = child(cnv, "hlinkClick");
        return CNvPrInfo {
            id: attr(&cnv, "id"),
            name: attr(&cnv, "name").filter(|s| !s.is_empty()),
            hyperlink: hlink_click
                .and_then(|h| attr_r(&h, "id"))
                .and_then(|rid| rels.get(&rid).cloned())
                .filter(|s| !s.is_empty()),
            hyperlink_action: hlink_click
                .and_then(|h| attr(&h, "action"))
                .filter(|s| !s.is_empty()),
        };
    }
    CNvPrInfo::default()
}

/// Locate the shape/tree node's OWN `<p:cNvPr>` — the `cNvPr` inside this
/// node's direct non-visual-properties wrapper (`nvSpPr` / `nvCxnSpPr` /
/// `nvPicPr` / `nvGraphicFramePr` / `nvGrpSpPr`). Only a *direct* wrapper child
/// is inspected, never a descendant, so a group's own props are not confused
/// with a nested shape's props.
pub(crate) fn own_cnv_pr<'a, 'input>(
    node: roxmltree::Node<'a, 'input>,
) -> Option<roxmltree::Node<'a, 'input>> {
    for wrapper_name in &[
        "nvSpPr",
        "nvCxnSpPr",
        "nvPicPr",
        "nvGraphicFramePr",
        "nvGrpSpPr",
    ] {
        if let Some(nv) = child(node, wrapper_name) {
            if let Some(cnv) = child(nv, "cNvPr") {
                return Some(cnv);
            }
        }
    }
    None
}

/// True when a tree node's own `<p:cNvPr hidden="1">` marks it hidden
/// (ECMA-376 §20.1.2.2.8 `CT_NonVisualDrawingProps@hidden`, `xsd:boolean`,
/// default `false`). A hidden shape/picture/frame/connector/group is not
/// rendered — skipped at parse time. A hidden `<p:grpSp>` hides its whole
/// subtree (the early return elides all descendants).
pub(crate) fn node_is_hidden(node: roxmltree::Node<'_, '_>) -> bool {
    own_cnv_pr(node)
        .map(ooxml_common::drawing::nv_props_hidden)
        .unwrap_or(false)
}

/// Merge an optional local `<a:ln>` over the inherited line style.
///
/// CT_LineProperties is field-wise: PowerPoint commonly writes a local line
/// containing only head/tail-end defaults while the paint and width remain in
/// `p:style/a:lnRef`. Treating the mere presence of that partial node as a
/// replacement drops the visible theme outline.
fn merge_local_stroke(
    ln_node: Option<roxmltree::Node<'_, '_>>,
    inherited: Option<Stroke>,
    theme: &HashMap<String, String>,
) -> Option<Stroke> {
    let Some(ln) = ln_node else {
        return inherited;
    };
    let properties = parse_line_properties(
        ln,
        &PptxSchemeResolver { theme },
        ooxml_common::color::TintMode::PowerPointLinear,
    );
    if matches!(properties.paint, Some(LinePaint::NoFill)) {
        return None;
    }

    let parsed = line_properties_to_stroke(&properties, None);
    let ln_width = properties.width;
    let has_dash = properties.dash.is_some();
    let ln_dash = preset_dash(&properties);
    let ln_custom_dash = custom_dash(&properties);
    let has_cap = properties.cap.is_some();
    let ln_cap = properties.cap.as_deref().and_then(canvas_line_cap);
    let has_head = properties.head_end.is_some();
    let ln_head = properties.head_end.as_ref().and_then(to_arrow_end);
    let has_tail = properties.tail_end.is_some();
    let ln_tail = properties.tail_end.as_ref().and_then(to_arrow_end);
    let has_cmpd = properties.compound.is_some();
    let ln_cmpd = properties.compound.clone().filter(|value| value != "sng");
    let has_join = properties.join.is_some();
    let (ln_join, ln_miter_limit) = line_join(&properties);
    let has_alignment = properties.alignment.is_some();
    let ln_alignment = properties
        .alignment
        .clone()
        .filter(|value| matches!(value.as_str(), "ctr" | "in"));

    match (parsed, inherited) {
        (Some(authored), base) => Some(Stroke {
            color: authored.color,
            width: ln_width.unwrap_or_else(|| base.as_ref().map(|s| s.width).unwrap_or(9525)),
            fill: authored.fill,
            dash_style: if has_dash {
                ln_dash
            } else {
                authored
                    .dash_style
                    .or_else(|| base.as_ref().and_then(|s| s.dash_style.clone()))
            },
            custom_dash: if has_dash {
                ln_custom_dash
            } else {
                base.as_ref()
                    .map(|stroke| stroke.custom_dash.clone())
                    .unwrap_or_default()
            },
            line_cap: if has_cap {
                ln_cap
            } else {
                authored
                    .line_cap
                    .or_else(|| base.as_ref().and_then(|s| s.line_cap.clone()))
            },
            line_join: if has_join {
                ln_join
            } else {
                base.as_ref().and_then(|stroke| stroke.line_join.clone())
            },
            miter_limit: if has_join {
                ln_miter_limit
            } else {
                base.as_ref().and_then(|stroke| stroke.miter_limit)
            },
            alignment: if has_alignment {
                ln_alignment
            } else {
                base.as_ref().and_then(|stroke| stroke.alignment.clone())
            },
            head_end: if has_head {
                ln_head
            } else {
                authored
                    .head_end
                    .or_else(|| base.as_ref().and_then(|s| s.head_end.clone()))
            },
            tail_end: if has_tail {
                ln_tail
            } else {
                authored
                    .tail_end
                    .or_else(|| base.as_ref().and_then(|s| s.tail_end.clone()))
            },
            cmpd: if has_cmpd {
                ln_cmpd
            } else {
                authored
                    .cmpd
                    .or_else(|| base.as_ref().and_then(|s| s.cmpd.clone()))
            },
        }),
        (None, Some(base)) => Some(Stroke {
            color: base.color,
            width: ln_width.unwrap_or(base.width),
            fill: base.fill,
            dash_style: if has_dash { ln_dash } else { base.dash_style },
            custom_dash: if has_dash {
                ln_custom_dash
            } else {
                base.custom_dash
            },
            line_cap: if has_cap { ln_cap } else { base.line_cap },
            line_join: if has_join { ln_join } else { base.line_join },
            miter_limit: if has_join {
                ln_miter_limit
            } else {
                base.miter_limit
            },
            alignment: if has_alignment {
                ln_alignment
            } else {
                base.alignment
            },
            head_end: if has_head { ln_head } else { base.head_end },
            tail_end: if has_tail { ln_tail } else { base.tail_end },
            cmpd: if has_cmpd { ln_cmpd } else { base.cmpd },
        }),
        (None, None) => None,
    }
}

/// Resolve `p:style/a:lnRef` into the inherited line component. The complete
/// theme line recipe will move to the shared format-scheme model; until then,
/// keeping this fallback in one place prevents shapes, connectors and pictures
/// from disagreeing about the same CT_ShapeStyle component.
fn parse_style_stroke(
    style_node: Option<roxmltree::Node<'_, '_>>,
    theme_source: &(impl PptxThemeSource + ?Sized),
) -> Option<Stroke> {
    let line_ref = style_node.and_then(|style| child(style, "lnRef"))?;
    resolve_style_line_ref(line_ref, theme_source)
}

fn resolve_style_line_ref(
    line_ref: roxmltree::Node<'_, '_>,
    theme_source: &(impl PptxThemeSource + ?Sized),
) -> Option<Stroke> {
    let theme = theme_source.colors();
    let index = attr(&line_ref, "idx")
        .and_then(|value| value.parse::<usize>().ok())
        .unwrap_or(1);
    let reference_color = parse_color_node(line_ref, theme);
    let reference_stroke = |color: String| Stroke {
        color,
        width: 9525,
        fill: None,
        dash_style: None,
        custom_dash: Vec::new(),
        line_cap: None,
        line_join: None,
        miter_limit: None,
        alignment: None,
        head_end: None,
        tail_end: None,
        cmpd: None,
    };
    let Some(format_scheme) = theme_source.format_scheme() else {
        return reference_color.map(reference_stroke);
    };
    let entry = match format_scheme.lookup_line_ref(index) {
        ooxml_common::theme::StyleMatrixLookup::NoStyle => return None,
        ooxml_common::theme::StyleMatrixLookup::Missing => {
            return reference_color.map(reference_stroke);
        }
        ooxml_common::theme::StyleMatrixLookup::Entry(entry) => entry,
    };
    let entry_xml = entry.to_xml();
    let document = roxmltree::Document::parse(&entry_xml).ok()?;
    let line = document.root_element().children().find(|node| {
        node.is_element()
            && node.tag_name().name() == "ln"
            && ooxml_common::ns::is_a_ns(node.tag_name().namespace())
    })?;
    let base_resolver = PptxRawSchemeResolver { theme };
    let resolver = ooxml_common::color::StyleMatrixColorResolver::new(
        &base_resolver,
        reference_color.as_deref(),
    );
    let mut properties = parse_line_properties(
        line,
        &resolver,
        ooxml_common::color::TintMode::PowerPointLinear,
    );
    if properties.paint.is_none() {
        properties.paint = reference_color.map(|color| LinePaint::Solid { color: Some(color) });
    }
    line_properties_to_stroke(&properties, None)
}

fn style_line_ref_is_explicit_no_line(
    line_ref: roxmltree::Node<'_, '_>,
    theme_source: &(impl PptxThemeSource + ?Sized),
) -> bool {
    use ooxml_common::theme::StyleMatrixLookup;

    let index = attr(&line_ref, "idx")
        .and_then(|value| value.parse::<usize>().ok())
        .unwrap_or(0);
    if index == 0 {
        return true;
    }
    let Some(format_scheme) = theme_source.format_scheme() else {
        return false;
    };
    let StyleMatrixLookup::Entry(entry) = format_scheme.lookup_line_ref(index) else {
        return false;
    };
    let entry_xml = entry.to_xml();
    let Ok(document) = roxmltree::Document::parse(&entry_xml) else {
        return false;
    };
    let Some(line) = document.root_element().children().find(|node| {
        node.is_element()
            && node.tag_name().name() == "ln"
            && ooxml_common::ns::is_a_ns(node.tag_name().namespace())
    }) else {
        return false;
    };
    let reference_color = parse_color_node(line_ref, theme_source.colors());
    let raw_resolver = PptxRawSchemeResolver {
        theme: theme_source.colors(),
    };
    let resolver = ooxml_common::color::StyleMatrixColorResolver::new(
        &raw_resolver,
        reference_color.as_deref(),
    );
    matches!(
        parse_line_properties(
            line,
            &resolver,
            ooxml_common::color::TintMode::PowerPointLinear,
        )
        .paint,
        Some(LinePaint::NoFill)
    )
}

fn canvas_line_cap(cap: &str) -> Option<String> {
    match cap {
        "rnd" => Some("round".to_owned()),
        "sq" => Some("square".to_owned()),
        "flat" => Some("butt".to_owned()),
        _ => None,
    }
}

fn preset_dash(line: &LineProperties) -> Option<String> {
    match line.dash.as_ref() {
        Some(LineDash::Preset(Some(value))) if value != "solid" => Some(value.clone()),
        _ => None,
    }
}

fn custom_dash(line: &LineProperties) -> Vec<StrokeDashSegment> {
    match line.dash.as_ref() {
        Some(LineDash::Custom(stops)) => stops
            .iter()
            .map(|stop| StrokeDashSegment {
                dash: stop.dash / 100_000.0,
                space: stop.space / 100_000.0,
            })
            .collect(),
        _ => Vec::new(),
    }
}

fn line_join(line: &LineProperties) -> (Option<String>, Option<f64>) {
    match line.join.as_ref() {
        Some(LineJoin::Round) => (Some("round".to_owned()), None),
        Some(LineJoin::Bevel) => (Some("bevel".to_owned()), None),
        Some(LineJoin::Miter { limit }) => (
            Some("miter".to_owned()),
            limit.map(|value| value as f64 / 100_000.0),
        ),
        None => (None, None),
    }
}

fn to_arrow_end(end: &LineEnd) -> Option<ArrowEnd> {
    let kind = end.kind.clone().unwrap_or_else(|| "none".to_owned());
    (kind != "none").then(|| ArrowEnd {
        kind,
        w: end.width.clone().unwrap_or_else(|| "med".to_owned()),
        len: end.length.clone().unwrap_or_else(|| "med".to_owned()),
    })
}

/// The DrawingML shape-property components that affect a picture's composed
/// bitmap/border silhouette. `p:pic` and a `p:sp` painted by `a:blipFill` both
/// use `CT_ShapeProperties`; keeping their component cascade in one resolver
/// prevents the alternate picture construction paths from silently dropping
/// style-matrix or placeholder-inherited effects.
#[derive(Default, Clone, serde::Serialize)]
pub(crate) struct PictureShapeProperties {
    pub(crate) stroke: Option<Stroke>,
    pub(crate) shadow: Option<Shadow>,
    pub(crate) inner_shadow: Option<Shadow>,
    pub(crate) glow: Option<Glow>,
    pub(crate) soft_edge: Option<SoftEdge>,
    pub(crate) reflection: Option<Reflection>,
    pub(crate) scene3d: Option<Scene3d>,
    pub(crate) sp3d: Option<Sp3d>,
}

impl PictureShapeProperties {
    pub(crate) fn is_empty(&self) -> bool {
        self.stroke.is_none()
            && self.shadow.is_none()
            && self.inner_shadow.is_none()
            && self.glow.is_none()
            && self.soft_edge.is_none()
            && self.reflection.is_none()
            && self.scene3d.is_none()
            && self.sp3d.is_none()
    }
}

/// Resolve the component cascade for a picture-like shape.
///
/// ECMA-376 Part 1, Annex L.3.2.3 defines the slide/layout placeholder
/// inheritance tier. CT_ShapeStyle's `lnRef`/`effectRef` supplies the next
/// tier, and local CT_ShapeProperties components override it. In particular,
/// MS-OI29500 §20.1.2.2.37(b) specifies that an authored local effect component
/// replaces the referenced style component rather than merging its children.
pub(crate) fn resolve_picture_shape_properties(
    sp_pr: Option<roxmltree::Node<'_, '_>>,
    style_node: Option<roxmltree::Node<'_, '_>>,
    inherited: Option<PictureShapeProperties>,
    theme_source: &(impl PptxThemeSource + ?Sized),
) -> PictureShapeProperties {
    let theme = theme_source.colors();
    let inherited = inherited.unwrap_or_default();

    // An explicit lnRef (including idx=0) replaces the placeholder tier.
    // Without an lnRef, the layout placeholder remains the inherited base.
    let style_stroke = if style_node.and_then(|style| child(style, "lnRef")).is_some() {
        parse_style_stroke(style_node, theme_source)
    } else {
        inherited.stroke
    };
    let stroke = merge_local_stroke(
        sp_pr.and_then(|node| child(node, "ln")),
        style_stroke,
        theme,
    );

    let effect_ref = style_node.and_then(|style| child(style, "effectRef"));
    let style_effects = effect_ref
        .map(|node| parse_style_matrix_effects(node, theme))
        .unwrap_or_default();
    let local_effect_node =
        sp_pr.and_then(|node| child(node, "effectLst").or_else(|| child(node, "effectDag")));
    let effects = if local_effect_node.is_some() {
        parse_effect_lst(local_effect_node, theme)
    } else if effect_ref.is_some() {
        style_effects.effects
    } else {
        EffectLst {
            shadow: inherited.shadow,
            inner_shadow: inherited.inner_shadow,
            glow: inherited.glow,
            soft_edge: inherited.soft_edge,
            reflection: inherited.reflection,
        }
    };

    // An explicit effectRef replaces the inherited CT_EffectStyleItem as a
    // whole.  Its absence may inherit layout 3-D; its presence with idx=0 or
    // an item without scene3d/sp3d must clear the older layout components just
    // as it clears/replaces the raster effects above.
    let inherited_scene3d = if effect_ref.is_some() {
        None
    } else {
        inherited.scene3d
    };
    let inherited_sp3d = if effect_ref.is_some() {
        None
    } else {
        inherited.sp3d
    };

    PictureShapeProperties {
        stroke,
        shadow: effects.shadow,
        inner_shadow: effects.inner_shadow,
        glow: effects.glow,
        soft_edge: effects.soft_edge,
        reflection: effects.reflection,
        scene3d: sp_pr
            .and_then(parse_scene3d)
            .or(style_effects.scene3d)
            .or(inherited_scene3d),
        sp3d: sp_pr
            .and_then(|node| parse_sp3d(node, theme))
            .or(style_effects.sp3d)
            .or(inherited_sp3d),
    }
}

pub(crate) fn parse_shape(
    sp_node: roxmltree::Node<'_, '_>,
    lph: &LayoutPlaceholders,
    theme_source: &(impl PptxThemeSource + ?Sized),
    rels: &HashMap<String, String>,
    source_dir: &str,
    group_fill: Option<&Fill>,
    zip: &mut PptxZip,
) -> Option<ShapeElement> {
    let theme = theme_source.colors();
    // --- Placeholder info (for layout fallback) ---
    let ph_node = sp_node
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "ph");
    let ph_type = ph_node
        .as_ref()
        .and_then(|n| attr(n, "type"))
        .unwrap_or_else(|| "body".into());
    let ph_idx: Option<u32> = ph_node
        .as_ref()
        .and_then(|n| attr(n, "idx"))
        .and_then(|v| v.parse().ok())
        // PowerPoint uses the unsigned maximum as an unbound-placeholder
        // sentinel. It is not a real layout slot; treating it as an authored
        // idx suppresses the normal type/bodyStyle fallback.
        .filter(|idx| *idx != u32::MAX);
    // Surface the explicit ph @type / @idx (when present) on the ShapeElement
    // JSON. We keep `ph_type` defaulted to "body" for the internal lookup
    // path, but only emit the `placeholder_type` field when a real `<p:ph>`
    // node was found.
    let placeholder_type_out: Option<String> = ph_node
        .as_ref()
        .map(|n| attr(n, "type").unwrap_or_else(|| "body".into()));
    let cnv = read_cnv_pr(sp_node, rels);

    // --- Transform: slide xfrm OR layout fallback ---
    let sp_pr = child(sp_node, "spPr");
    let slide_xfrm = sp_pr.and_then(|p| child(p, "xfrm"));

    let t: Transform = if let Some(xfrm) = slide_xfrm {
        parse_xfrm(xfrm)
    } else if ph_node.is_some() {
        match lph.lookup(&ph_type, ph_idx) {
            Some(lt) => lt.clone(),
            None => default_placeholder_transform(&ph_type),
        }
    } else {
        return None; // non-placeholder with no xfrm — skip
    };

    // cx=0 AND cy=0 → skip (zero-area invisible shape).
    // cx=0 alone (vertical line/annotation) is permitted when cy > 0.
    // cy=0 means "auto-height": keep 0 when anchor="b" (renderer grows shape upward from off_y),
    // otherwise use a generous fallback so text has room to render.
    if t.cx == 0 && t.cy == 0 {
        return None;
    }
    let inherited_anchor: Option<String> = if ph_node.is_some() {
        lph.lookup_anchor(&ph_type)
    } else {
        None
    };
    let is_bottom_anchor = inherited_anchor.as_deref() == Some("b")
        || child(sp_node, "txBody")
            .and_then(|tb| child(tb, "bodyPr"))
            .and_then(|bp| attr(&bp, "anchor"))
            .map(|a| a == "b")
            .unwrap_or(false);
    // Presentation slides inherit layout information unless they provide a
    // local override (ECMA-376 Part 1, Annex L.3.2.3). Geometry is part of the
    // shape properties: a placeholder with only a local xfrm can still inherit
    // an ellipse/custom geometry from the matching layout placeholder.
    let shape_geometry = sp_pr
        .and_then(|properties| {
            InheritedShapeGeometry::from_sp_pr(properties, t.cx as f64, t.cy as f64)
        })
        .or_else(|| {
            if ph_node.is_some() {
                lph.lookup_geometry(&ph_type, ph_idx)
            } else {
                None
            }
        })
        .unwrap_or_else(|| InheritedShapeGeometry {
            geometry: "rect".to_owned(),
            cust_geom: None,
            adjustments: [None; 8],
        });
    let geometry = shape_geometry.geometry;
    let cust_geom = shape_geometry.cust_geom;
    let [adj, adj2, adj3, adj4, adj5, adj6, adj7, adj8] = shape_geometry.adjustments;

    // cy=0 means "auto-height" for body-text shapes, but connector-type
    // geometries (line, *Connector*) legitimately use cy=0 to represent
    // a perfectly horizontal segment — don't inflate their height.
    let is_connector_geom = matches!(
        geometry.as_str(),
        "line"
            | "straightConnector1"
            | "bentConnector2"
            | "bentConnector3"
            | "bentConnector4"
            | "bentConnector5"
            | "curvedConnector2"
            | "curvedConnector3"
            | "curvedConnector4"
            | "curvedConnector5"
    );
    let cy = if t.cy == 0 && !is_connector_geom {
        if is_bottom_anchor {
            0_i64
        } else {
            2_000_000_i64
        }
    } else {
        t.cy
    };
    // --- Shape style (p:style) provides fill/stroke/text-color fallbacks ---
    let style_node = child(sp_node, "style");

    // fillRef idx=0 → explicit no-fill; idx>0 → resolve the referenced
    // theme fill style, substituting this ref's colour for phClr. Retain the
    // former solid-colour fallback for incomplete/malformed themes.
    let style_fill: Option<Fill> = style_node.and_then(|s| child(s, "fillRef")).and_then(|fr| {
        let idx: u32 = attr(&fr, "idx").and_then(|v| v.parse().ok()).unwrap_or(1);
        if idx == 0 {
            Some(Fill::None)
        } else {
            parse_style_matrix_fill_from_source(fr, theme_source)
                .or_else(|| parse_color_node(fr, theme).map(|c| Fill::Solid { color: c }))
        }
    });

    // lnRef idx=0 → no line; idx>0 → resolve the complete theme fmtScheme >
    // lnStyleLst recipe, substituting the ref color only for `phClr`.
    let style_stroke = parse_style_stroke(style_node, theme_source);

    // fontRef → default text color for this shape.
    // Fall back to layout/master placeholder inherited color (lstStyle > lvl1pPr > defRPr
    // > solidFill) when the shape is a placeholder and has no explicit p:style > fontRef.
    let default_text_color: Option<String> = style_node
        .and_then(|s| child(s, "fontRef"))
        .and_then(|fr| parse_color_node(fr, theme))
        .or_else(|| {
            if ph_node.is_some() {
                lph.lookup_color(&ph_type, ph_idx)
            } else {
                None
            }
        });

    // spPr fill resolution order (ECMA-376 §19.3.1.36 / §20.1.4.2):
    //   1. spPr `<a:grpFill>` → inherit from parent group
    //   2. spPr explicit fill (`<a:solidFill>`, `<a:noFill>`, gradient, pattern, blip)
    //   3. the slide shape's `<p:style><a:fillRef>`
    //   4. layout/master placeholder fill only when this shape did not author
    //      a style component
    // Some(Fill::None) (noFill in spPr) is treated as an explicit choice and
    // must NOT be overridden by step 3 or 4.
    let sp_pr_has_grp_fill = sp_pr.and_then(|p| child(p, "grpFill")).is_some();
    let fill = if sp_pr_has_grp_fill {
        group_fill.cloned()
    } else {
        let own = sp_pr.and_then(|p| parse_fill(p, theme));
        let inherited = if own.is_none() && style_fill.is_none() && ph_node.is_some() {
            lph.lookup_fill(&ph_type, ph_idx)
        } else {
            None
        };
        own.or(style_fill).or(inherited)
    };

    // A local line is a field-wise override over layout/style inheritance;
    // explicit noFill still suppresses the inherited line entirely.
    let has_style_line_ref = style_node.and_then(|style| child(style, "lnRef")).is_some();
    let inherited_stroke = if has_style_line_ref {
        // Presence is authoritative even when idx=0 or the selected theme
        // recipe is noFill. Do not resurrect a layout/master outline.
        style_stroke
    } else if ph_node.is_some() {
        lph.lookup_stroke(&ph_type, ph_idx)
    } else {
        None
    };
    let stroke = merge_local_stroke(sp_pr.and_then(|p| child(p, "ln")), inherited_stroke, theme);

    // Inherited defaults from layout/master for this placeholder type/idx
    let (
        inherited_font_size,
        inherited_font_family,
        inherited_bold,
        inherited_italic,
        inherited_caps,
        inherited_reflection,
        inherited_anchor,
        inherited_text_insets,
        inherited_alignment,
        inherited_ea_ln_brk,
        inherited_space_before,
        inherited_space_after,
        inherited_line_spacing,
    ) = if ph_node.is_some() {
        (
            lph.lookup_font_size(&ph_type, ph_idx),
            lph.lookup_font_family(&ph_type, ph_idx),
            lph.lookup_bold(&ph_type),
            lph.lookup_italic(&ph_type),
            lph.lookup_caps(&ph_type),
            lph.lookup_reflection(&ph_type),
            lph.lookup_anchor(&ph_type),
            lph.lookup_text_insets(&ph_type, ph_idx),
            lph.lookup_alignment(&ph_type, ph_idx),
            lph.lookup_ea_ln_brk(&ph_type),
            lph.lookup_space_before(&ph_type),
            lph.lookup_space_after(&ph_type),
            lph.lookup_line_spacing(&ph_type, ph_idx),
        )
    } else {
        (
            None, None, None, None, None, None, None, None, None, None, None, None, None,
        )
    };
    let inherited_level_font_sizes: LevelFontSizes = if ph_node.is_some() {
        lph.lookup_level_font_sizes(&ph_type, ph_idx)
    } else {
        [None; 9]
    };
    // Per-level paragraph indents (marL/marR/indent) a paragraph inherits when it
    // omits them (ECMA-376 §21.1.2.4.13): the layout/master placeholder cascade.
    let inherited_level_indents: LevelIndents = if ph_node.is_some() {
        lph.lookup_level_indents(&ph_type, ph_idx)
    } else {
        Default::default()
    };

    // ECMA-376 §19.3.1.21 / §20.1.4.2: a slide-level `<p:cNvSpPr txBox="1"/>`
    // marks the shape as a true text box, which means the theme's
    // `<a:txDef>` (rather than `<a:spDef>`) provides the fallback bodyPr.
    let is_text_box = child(sp_node, "nvSpPr")
        .and_then(|n| child(n, "cNvSpPr"))
        .and_then(|n| attr(&n, "txBox"))
        .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
        .unwrap_or(false);
    let shape_kind = if is_text_box {
        ShapeKind::Tx
    } else {
        ShapeKind::Sp
    };
    // Per-level bullets a paragraph inherits when it declares no explicit one
    // (ECMA-376 §19.7.10): the layout/master placeholder cascade for this slot.
    let inherited_level_bullets: LevelBullets = if ph_node.is_some() {
        lph.lookup_level_bullets(&ph_type, ph_idx)
    } else {
        empty_level_bullets()
    };
    let text_body = child(sp_node, "txBody").map(|n| {
        parse_text_body(
            n,
            theme,
            rels,
            source_dir,
            inherited_font_size,
            inherited_font_family,
            inherited_level_font_sizes,
            inherited_level_indents,
            &inherited_level_bullets,
            inherited_bold,
            inherited_italic,
            inherited_caps.clone(),
            inherited_reflection.clone(),
            inherited_anchor,
            inherited_text_insets,
            inherited_alignment,
            inherited_ea_ln_brk,
            inherited_space_before,
            inherited_space_after,
            inherited_line_spacing,
            shape_kind,
            zip,
        )
    });

    // Effects from spPr > effectLst (outerShdw / innerShdw / glow / softEdge /
    // reflection are independent siblings — ECMA-376 §20.1.8.16). Pull each.
    let style_effects = style_node
        .and_then(|style| child(style, "effectRef"))
        .map(|effect_ref| parse_style_matrix_effects(effect_ref, theme))
        .unwrap_or_default();
    let style_scene3d = style_effects.scene3d;
    let style_sp3d = style_effects.sp3d;
    let local_effect_node =
        sp_pr.and_then(|p| child(p, "effectLst").or_else(|| child(p, "effectDag")));
    let EffectLst {
        shadow,
        inner_shadow,
        glow,
        soft_edge,
        reflection,
    } = if local_effect_node.is_some() {
        // MS-OI29500 §20.1.2.2.37(b): a local effect component replaces the
        // style component rather than merging missing fields from effectRef.
        parse_effect_lst(local_effect_node, theme)
    } else {
        style_effects.effects
    };

    Some(ShapeElement {
        x: t.x,
        y: t.y,
        width: t.cx,
        height: cy,
        rotation: t.rot,
        flip_h: t.flip_h,
        flip_v: t.flip_v,
        geometry,
        fill,
        stroke,
        text_body,
        default_text_color,
        cust_geom,
        adj,
        adj2,
        adj3,
        adj4,
        adj5,
        adj6,
        adj7,
        adj8,
        shadow,
        inner_shadow,
        glow,
        soft_edge,
        reflection,
        id: cnv.id,
        name: cnv.name,
        hyperlink: cnv.hyperlink,
        hyperlink_action: cnv.hyperlink_action,
        placeholder_type: placeholder_type_out,
        placeholder_idx: ph_idx,
        text_rect: None,
        scene3d: sp_pr.and_then(parse_scene3d).or(style_scene3d),
        sp3d: sp_pr
            .and_then(|node| parse_sp3d(node, theme))
            .or(style_sp3d),
    })
}

/// Preset clip geometry for a picture body (§20.1.9.18 — the preset geometry of
/// a `<p:pic>` acts as its clip silhouette and the path its border / contour
/// hug). Returns the prstGeom `prst` name plus every `<a:avLst><a:gd>` adjust
/// value in declaration order (1/1000-of-a-percent OOXML units), so any of the
/// 186 presets — roundRect, ellipse, and the rest — can be reconstructed by the
/// shared preset-geometry engine on the TS side. The preset's own declared
/// defaults fill in any omitted guide, so the `adj` Vec may be shorter than the
/// preset's adjust count (the engine substitutes defaults per index).
///
/// `prst="rect"` returns None: a plain rectangle needs no clip path (the bitmap
/// already fills the bbox), matching the previous behaviour.
pub(crate) fn parse_pic_prst_geom(
    sp_pr: roxmltree::Node<'_, '_>,
) -> (Option<String>, Option<Vec<i64>>) {
    let pg = match child(sp_pr, "prstGeom") {
        Some(n) => n,
        None => return (None, None),
    };
    let prst = match attr(&pg, "prst") {
        Some(p) if p != "rect" => p,
        _ => return (None, None),
    };
    let adjust: Vec<i64> = child(pg, "avLst")
        .map(|av| {
            av.children()
                .filter(|n| n.is_element() && n.tag_name().name() == "gd")
                .filter_map(|gd| {
                    attr(&gd, "fmla")
                        .and_then(|f| f.strip_prefix("val ").and_then(|v| v.parse::<i64>().ok()))
                })
                .collect()
        })
        .unwrap_or_default();
    let adjust = if adjust.is_empty() {
        None
    } else {
        Some(adjust)
    };
    (Some(prst), adjust)
}

/// Microsoft 2016 SVG extension on an `<a:blip>` (ECMA-376 §20.1.8.14 carries
/// the core blip fill; the SVG body is a Microsoft extension). When a `<a:blip>`
/// has `<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">` with an
/// `<asvg:svgBlip r:embed="…">`, resolve that rId to the `.svg` part and return
/// it as a `data:image/svg+xml;base64,…` URL. Returns None when there is no
/// svgBlip, the rId is unresolvable, or the part is missing.
///
/// Matching is by namespace-local element name (`svgBlip`), so the `asvg:`
/// prefix (or any other) is irrelevant.
/// Resolve the Microsoft 2016 `asvg:svgBlip` extension on a `<a:blip>` to the
/// embedded **zip path** of the `.svg` part (e.g. "ppt/media/image2.svg"). The
/// renderer fetches the bytes lazily by path. The part's existence is verified
/// so a dangling rId yields None (no SVG twin), matching the prior behaviour.
pub(crate) fn svg_blip_path(
    blip: roxmltree::Node<'_, '_>,
    slide_dir: &str,
    rels: &HashMap<String, String>,
    zip: &mut PptxZip,
) -> Option<String> {
    let svg_rid = svg_blip_rid(blip)?;
    let svg_target = rels.get(&svg_rid)?;
    let svg_path = resolve_path(slide_dir, svg_target);
    // Existence check only — `index_for_name` reads the central directory
    // without inflating, unlike the former `read_zip_bytes` which decompressed
    // the whole SVG part just to discard the bytes.
    zip.index_for_name(&svg_path)?;
    Some(svg_path)
}

/// Resolved drawable source of a `<*:blipFill><a:blip>` — the winning image
/// part plus, when present, the Microsoft-2016 SVG twin surfaced separately so
/// the renderer can prefer the vector original.
pub(crate) struct BlipSource {
    /// Zip path of the raster (preferred) or, when no raster is embedded, of the
    /// SVG part itself, so the element is always drawable.
    pub image_path: String,
    /// MIME of whichever source wins (`image/svg+xml` for the SVG-only case).
    pub mime_type: String,
    /// Intrinsic PNG size for ink-fallback centering (`None` for non-PNG).
    pub intrinsic_width_px: Option<u32>,
    pub intrinsic_height_px: Option<u32>,
    /// The SVG twin's zip path when the blip carries an `asvg:svgBlip`.
    pub svg_image_path: Option<String>,
}

/// Resolve a `<a:blip>`'s drawable source from its parent `blip_fill` node.
///
/// Microsoft 2016 SVG extension: the real vector image rides inside
/// `<a:blip><a:extLst>`, while `<a:blip r:embed>` carries a raster (PNG/JPEG)
/// *fallback* for SVG-incapable clients. Resolve the SVG first so a picture that
/// carries ONLY the svgBlip — with no raster `r:embed`, e.g. an icon inserted as
/// a pure SVG — still parses instead of being dropped. The raster is optional; a
/// pure-SVG picture omits it. A picture needs at least one drawable source: the
/// raster is preferred as `image_path` (the renderer's srcRect / SVG-decode-
/// failure fallback path); when no raster is embedded, fall back to the SVG part
/// itself. Returns `None` when neither source resolves.
///
/// Shared by `parse_picture` and `parse_ole_preview_picture` so the ordinary and
/// OLE-preview picture paths resolve blips identically.
pub(crate) fn resolve_blip_source(
    blip_fill: roxmltree::Node<'_, '_>,
    slide_dir: &str,
    rels: &HashMap<String, String>,
    zip: &mut PptxZip,
) -> Option<BlipSource> {
    let blip = child(blip_fill, "blip")?;

    let svg_image_path = svg_blip_path(blip, slide_dir, rels, zip);

    // Raster blip (`<a:blip r:embed>` → zip path + intrinsic PNG size).
    let raster: Option<(String, Option<(u32, u32)>)> = (|| {
        let r_id = attr_r(&blip, "embed")?;
        let rel_target = rels.get(&r_id)?;
        let path = resolve_path(slide_dir, rel_target);
        // PNG dimensions live in the first 24 bytes. Do not inflate a potentially
        // huge image merely to discover its intrinsic size.
        let image_bytes = read_zip_head(zip, &path, 24).ok()?;
        // Intrinsic PNG size for the ink-fallback centering (None for non-PNG,
        // unchanged from the former png_size_from_data_url semantics).
        let size = png_size_from_bytes(&image_bytes);
        Some((path, size))
    })();

    let (image_path, mime_type, intrinsic_width_px, intrinsic_height_px) =
        match (raster, svg_image_path.as_ref()) {
            (Some((path, size)), _) => {
                let mime = mime_from_ext(&path).to_owned();
                let (w, h) = match size {
                    Some((w, h)) => (Some(w), Some(h)),
                    None => (None, None),
                };
                (path, mime, w, h)
            }
            (None, Some(svg)) => (svg.clone(), "image/svg+xml".to_owned(), None, None),
            (None, None) => return None,
        };

    Some(BlipSource {
        image_path,
        mime_type,
        intrinsic_width_px,
        intrinsic_height_px,
        svg_image_path,
    })
}

pub(crate) fn parse_picture(
    pic_node: roxmltree::Node<'_, '_>,
    slide_dir: &str,
    rels: &HashMap<String, String>,
    theme_source: &(impl PptxThemeSource + ?Sized),
    zip: &mut PptxZip,
) -> Option<PictureElement> {
    let theme = theme_source.colors();
    let sp_pr = child(pic_node, "spPr")?;
    let xfrm_node = child(sp_pr, "xfrm")?;
    let t = parse_xfrm(xfrm_node);

    if t.cx == 0 || t.cy == 0 {
        return None; // pictures always need explicit dimensions
    }

    let blip_fill = child(pic_node, "blipFill")?;
    // Resolve the drawable blip source (raster preferred, SVG twin surfaced).
    let BlipSource {
        image_path,
        mime_type,
        intrinsic_width_px,
        intrinsic_height_px,
        svg_image_path,
    } = resolve_blip_source(blip_fill, slide_dir, rels, zip)?;

    // ECMA-376 §20.1.9.8 — `<p:pic>` may carry `<a:custGeom>` inside `<p:spPr>`,
    // in which case the bitmap is clipped to that custom path (e.g. a laptop
    // silhouette). Re-use the same parser as for shapes so the renderer can
    // build a Path2D and `ctx.clip()` before drawing the image.
    let cust_geom = child(sp_pr, "custGeom")
        .map(|geometry| parse_cust_geom(geometry, t.cx as f64, t.cy as f64));

    let style_node = child(pic_node, "style");
    let PictureShapeProperties {
        stroke,
        shadow,
        inner_shadow,
        glow,
        soft_edge,
        reflection,
        scene3d,
        sp3d,
    } = resolve_picture_shape_properties(Some(sp_pr), style_node, None, theme_source);

    let (prst_geom, prst_adjust) = parse_pic_prst_geom(sp_pr);
    Some(PictureElement {
        id: read_cnv_pr(pic_node, rels).id,
        x: t.x,
        y: t.y,
        width: t.cx,
        height: t.cy,
        rotation: t.rot,
        flip_h: t.flip_h,
        flip_v: t.flip_v,
        image_path,
        mime_type,
        svg_image_path,
        intrinsic_width_px,
        intrinsic_height_px,
        stroke,
        prst_geom,
        prst_adjust,
        src_rect: parse_src_rect(blip_fill),
        alpha: parse_blip_alpha(blip_fill),
        // §20.1.8.23 `<a:duotone>` recolour, resolved through the slide's theme
        // palette with PowerPoint's linear tint. `None` ⇒ no effect.
        duotone: parse_blip_duotone(
            blip_fill,
            &PptxSchemeResolver { theme },
            ooxml_common::color::TintMode::PowerPointLinear,
        ),
        cust_geom,
        shadow,
        inner_shadow,
        glow,
        soft_edge,
        reflection,
        scene3d,
        sp3d,
    })
}

/// Build a `PictureElement` for an OLE preview `<p:pic>` that carries NO
/// `<a:xfrm>` (so `parse_picture` returns None). The enclosing graphicFrame's
/// `Transform` supplies placement and size. Only the blip resolution is
/// borrowed from the ordinary picture path; every drawing decoration
/// (custGeom, effects, srcRect …) that requires an spPr xfrm is irrelevant for
/// an unstyled preview, so we keep them at their neutral defaults.
pub(crate) fn parse_ole_preview_picture(
    pic_node: roxmltree::Node<'_, '_>,
    gf: &Transform,
    slide_dir: &str,
    rels: &HashMap<String, String>,
    theme: &HashMap<String, String>,
    zip: &mut PptxZip,
) -> Option<PictureElement> {
    if gf.cx == 0 || gf.cy == 0 {
        return None; // no drawable box
    }
    let blip_fill = child(pic_node, "blipFill")?;
    // Same blip resolution as the ordinary picture path (raster preferred, SVG
    // twin surfaced), shared via `resolve_blip_source`.
    let BlipSource {
        image_path,
        mime_type,
        intrinsic_width_px,
        intrinsic_height_px,
        svg_image_path,
    } = resolve_blip_source(blip_fill, slide_dir, rels, zip)?;

    Some(PictureElement {
        id: read_cnv_pr(pic_node, rels).id,
        x: gf.x,
        y: gf.y,
        width: gf.cx,
        height: gf.cy,
        rotation: gf.rot,
        flip_h: gf.flip_h,
        flip_v: gf.flip_v,
        image_path,
        mime_type,
        svg_image_path,
        intrinsic_width_px,
        intrinsic_height_px,
        stroke: None,
        prst_geom: None,
        prst_adjust: None,
        src_rect: parse_src_rect(blip_fill),
        alpha: parse_blip_alpha(blip_fill),
        duotone: parse_blip_duotone(
            blip_fill,
            &PptxSchemeResolver { theme },
            ooxml_common::color::TintMode::PowerPointLinear,
        ),
        cust_geom: None,
        shadow: None,
        inner_shadow: None,
        glow: None,
        soft_edge: None,
        reflection: None,
        scene3d: None,
        sp3d: None,
    })
}

/// Decode (width, height) from raw PNG bytes by reading the IHDR chunk. Returns
/// None for non-PNG payloads or malformed data (unchanged semantics from the
/// former `png_size_from_data_url`, only the input is now raw bytes instead of
/// a base64 data URL).
pub(crate) fn png_size_from_bytes(bytes: &[u8]) -> Option<(u32, u32)> {
    if bytes.len() < 24 || &bytes[0..8] != b"\x89PNG\r\n\x1a\n" {
        return None;
    }
    let w = u32::from_be_bytes([bytes[16], bytes[17], bytes[18], bytes[19]]);
    let h = u32::from_be_bytes([bytes[20], bytes[21], bytes[22], bytes[23]]);
    Some((w, h))
}

/// If a `p:pic` declares an `a:audioFile` / `a:videoFile` in its `nvPr`
/// (or the newer `p14:media` extension), emit a `MediaElement` with the
/// poster image and the media bytes. Returns None for regular pictures.
pub(crate) fn parse_media(
    pic_node: roxmltree::Node<'_, '_>,
    slide_dir: &str,
    rels: &HashMap<String, String>,
) -> Option<MediaElement> {
    let nv_pic_pr = child(pic_node, "nvPicPr")?;
    let nv_pr = child(nv_pic_pr, "nvPr")?;

    // A `<p:pic>` can carry several media references at once: the legacy
    // `<a:videoFile>`/`<a:audioFile>` (r:embed or r:link) and the modern
    // `<p14:media r:embed>` extension. We collect every reference and use the
    // first that actually resolves, preferring an *embedded* ref over a
    // *linked* one. Embedded refs always point at internal media we can read;
    // a linked ref may target an External resource we cannot fetch, or its rId
    // may be missing/unresolvable (caught below via `rels.get`). Trying embeds
    // first means a present `<p14:media>` embed is used even when the deck also
    // carries a videoFile link, instead of being shadowed and demoted to a
    // poster-only Picture. (Note: parse_rels stores Id+Target and never reads
    // TargetMode, so a real External link keeps its non-empty URL here.)
    // (ECMA-376 §19.3.1.17/18 a:videoFile/a:audioFile; p14:media is a Microsoft
    // extension.)
    let av = nv_pr
        .children()
        .find(|n| n.is_element() && matches!(n.tag_name().name(), "videoFile" | "audioFile"));
    let av_kind: Option<&str> = av.map(|n| {
        if n.tag_name().name() == "videoFile" {
            "video"
        } else {
            "audio"
        }
    });
    let av_embed = av.and_then(|n| attr_r(&n, "embed"));
    let av_link = av.and_then(|n| attr_r(&n, "link"));

    // `<p14:media>` lives in `nvPr > extLst > ext`.
    let p14 = child(nv_pr, "extLst").and_then(|ext_lst| {
        ext_lst
            .children()
            .filter(|c| c.is_element())
            .find_map(|ext| {
                ext.children()
                    .find(|m| m.is_element() && m.tag_name().name() == "media")
            })
    });
    let p14_embed = p14.and_then(|n| attr_r(&n, "embed"));
    let p14_link = p14.and_then(|n| attr_r(&n, "link"));

    // Preference: embedded refs (always internal) before linked refs (a link may
    // be External / unresolved); within each, the kind-bearing
    // videoFile/audioFile before the kind-less p14:media.
    let candidates = [
        (av_kind, av_embed),
        (None, p14_embed),
        (av_kind, av_link),
        (None, p14_link),
    ];
    let (media_kind, media_path, mime) = candidates.into_iter().find_map(|(kind_hint, rid)| {
        let rid = rid?;
        let target = rels.get(&rid)?;
        if target.is_empty() {
            return None; // malformed rel with an empty Target (External links carry a non-empty URL)
        }
        let path = resolve_path(slide_dir, target);
        let mime = mime_from_ext(&path);
        let kind = match kind_hint {
            Some(k) => k.to_string(),
            None if mime.starts_with("video/") => "video".to_string(),
            None if mime.starts_with("audio/") => "audio".to_string(),
            None => return None, // p14:media with an unknown MIME — try the next ref
        };
        Some((kind, path, mime))
    })?;

    // Geometry
    let sp_pr = child(pic_node, "spPr")?;
    let xfrm_node = child(sp_pr, "xfrm")?;
    let t = parse_xfrm(xfrm_node);
    if t.cx == 0 || t.cy == 0 {
        return None;
    }

    // Poster image (from blipFill). Optional — the renderer falls back to a
    // solid fill when the poster is absent or still loading. We emit just the
    // zip path so the main thread can lazily `getMedia` it; embedding large
    // posters inline ballooned the parse output for video-heavy decks.
    let (poster_path, poster_mime_type) = (|| -> Option<(String, String)> {
        let blip_fill = child(pic_node, "blipFill")?;
        let r_id = child(blip_fill, "blip").and_then(|b| attr_r(&b, "embed"))?;
        let rel_target = rels.get(&r_id)?;
        let image_path = resolve_path(slide_dir, rel_target);
        let img_mime = mime_from_ext(&image_path).to_string();
        Some((image_path, img_mime))
    })()
    .unwrap_or_default();

    Some(MediaElement {
        id: read_cnv_pr(pic_node, rels).id,
        x: t.x,
        y: t.y,
        width: t.cx,
        height: t.cy,
        rotation: t.rot,
        flip_h: t.flip_h,
        flip_v: t.flip_v,
        media_kind,
        poster_path,
        poster_mime_type,
        media_path,
        mime_type: mime.to_string(),
    })
}

/// Parse ppt/tableStyles.xml into a map of styleId → TableStyleDef.
pub(crate) fn parse_table_styles_xml(
    xml: &str,
    theme_source: &(impl PptxThemeSource + ?Sized),
) -> HashMap<String, TableStyleDef> {
    let theme = theme_source.colors();
    let mut map = HashMap::new();
    let Ok(doc) = crate::parse_preflighted_pptx_xml(xml) else {
        return map;
    };
    let root = doc.root_element();
    for style_node in root
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "tblStyle")
    {
        let Some(style_id) = attr(&style_node, "styleId") else {
            continue;
        };
        let style_id = style_id.to_string();
        let mut def = TableStyleDef::default();

        let parse_on_off_default = |value: Option<String>| -> Option<bool> {
            match value.as_deref() {
                Some("on" | "1" | "true") => Some(true),
                Some("off" | "0" | "false") => Some(false),
                // ST_OnOffStyleType `def` inherits; it is not explicit off.
                Some("def") | None => None,
                _ => None,
            }
        };
        let parse_tc_text = |role: roxmltree::Node<'_, '_>| -> TableTextStyle {
            let Some(text) = child(role, "tcTxStyle") else {
                return TableTextStyle::default();
            };
            TableTextStyle {
                color: parse_color_node(text, theme),
                bold: parse_on_off_default(attr(&text, "b")),
                italic: parse_on_off_default(attr(&text, "i")),
            }
        };

        // ECMA-376 §20.1.4.2.4 `CT_TableCellBorderStyle` declares left,
        // right, top, bottom, insideH, and insideV independently. A role's
        // explicit no-fill line is retained as `NoLine`; absence inherits.
        let parse_part = |role: roxmltree::Node<'_, '_>| -> TablePartStyle {
            let Some(tc_style) = child(role, "tcStyle") else {
                return TablePartStyle {
                    text: parse_tc_text(role),
                    ..Default::default()
                };
            };
            // §20.1.4.2.27: the role fill is `fill` or `fillRef`, never a fill
            // child directly on `tcStyle`.
            let fill = child(tc_style, "fill")
                .and_then(|fill| parse_table_style_fill(fill, theme))
                .or_else(|| {
                    child(tc_style, "fillRef").and_then(|fill_ref| {
                        let index = attr(&fill_ref, "idx")
                            .and_then(|value| value.parse::<u32>().ok())
                            .unwrap_or(0);
                        if index == 0 {
                            Some(Fill::None)
                        } else {
                            parse_style_matrix_fill_from_source(fill_ref, theme_source).or_else(
                                || {
                                    parse_color_node(fill_ref, theme)
                                        .map(|color| Fill::Solid { color })
                                },
                            )
                        }
                    })
                });
            let tc_bdr = child(tc_style, "tcBdr");
            let parse_side = |side: &str| -> TableLineStyle {
                let Some(side_node) = tc_bdr.and_then(|borders| child(borders, side)) else {
                    return TableLineStyle::Unspecified;
                };
                if let Some(line) = child(side_node, "ln") {
                    if child(line, "noFill").is_some() {
                        return TableLineStyle::NoLine;
                    }
                    return TableLineStyle::from_stroke(parse_stroke(line, theme));
                }
                if let Some(line_ref) = child(side_node, "lnRef") {
                    let index = attr(&line_ref, "idx")
                        .and_then(|value| value.parse::<u32>().ok())
                        .unwrap_or(0);
                    if index == 0 {
                        return TableLineStyle::NoLine;
                    }
                    if style_line_ref_is_explicit_no_line(line_ref, theme_source) {
                        return TableLineStyle::NoLine;
                    }
                    return TableLineStyle::from_stroke(resolve_style_line_ref(
                        line_ref,
                        theme_source,
                    ));
                }
                TableLineStyle::Unspecified
            };
            TablePartStyle {
                fill,
                text: parse_tc_text(role),
                borders: TableCellBorderStyle {
                    left: parse_side("left"),
                    right: parse_side("right"),
                    top: parse_side("top"),
                    bottom: parse_side("bottom"),
                    inside_h: parse_side("insideH"),
                    inside_v: parse_side("insideV"),
                    diagonal_tl: parse_side("tl2br"),
                    diagonal_tr: parse_side("tr2bl"),
                },
            }
        };

        macro_rules! parse_role {
            ($xml_name:literal, $field:ident) => {
                if let Some(role) = child(style_node, $xml_name) {
                    def.$field = parse_part(role);
                }
            };
        }
        parse_role!("wholeTbl", whole_tbl);
        parse_role!("band1H", band1_h);
        parse_role!("band2H", band2_h);
        parse_role!("band1V", band1_v);
        parse_role!("band2V", band2_v);
        parse_role!("firstRow", first_row);
        parse_role!("lastRow", last_row);
        parse_role!("firstCol", first_col);
        parse_role!("lastCol", last_col);
        parse_role!("nwCell", nw_cell);
        parse_role!("neCell", ne_cell);
        parse_role!("swCell", sw_cell);
        parse_role!("seCell", se_cell);

        map.insert(style_id, def);
    }
    map
}

#[derive(Clone, Copy)]
enum TablePartRegion {
    Whole,
    Row,
    Column,
    Cell,
}

/// Resolve the table-style cascade for one physical cell. ECMA-376 §21.1.3.11
/// defines thirteen `CT_TablePartStyle` components. PowerPoint applies the
/// general table role first, then banding, edge roles, and finally the corner
/// intersection. A role only replaces fields it actually declares.
#[allow(clippy::too_many_arguments)]
pub(crate) fn resolve_table_cell_style(
    style: &TableStyleDef,
    flags: TableStyleFlags,
    row: usize,
    col: usize,
    row_count: usize,
    col_count: usize,
) -> ResolvedTableCellStyle {
    let last_row = row_count.saturating_sub(1);
    let last_col = col_count.saturating_sub(1);
    let mut resolved = ResolvedTableCellStyle::default();

    let overlay =
        |resolved: &mut ResolvedTableCellStyle, role: &TablePartStyle, region: TablePartRegion| {
            if role.fill.is_some() {
                resolved.fill = role.fill.clone();
            }
            resolved.text.overlay(&role.text);

            resolved.diagonal_tl.overlay(&role.borders.diagonal_tl);
            resolved.diagonal_tr.overlay(&role.borders.diagonal_tr);

            let overlay_line = |target: &mut TableLineStyle, source: &TableLineStyle| {
                target.overlay(source);
            };
            match region {
                TablePartRegion::Whole => {
                    overlay_line(
                        &mut resolved.border_l,
                        if col == 0 {
                            &role.borders.left
                        } else {
                            &role.borders.inside_v
                        },
                    );
                    overlay_line(
                        &mut resolved.border_r,
                        if col == last_col {
                            &role.borders.right
                        } else {
                            &role.borders.inside_v
                        },
                    );
                    overlay_line(
                        &mut resolved.border_t,
                        if row == 0 {
                            &role.borders.top
                        } else {
                            &role.borders.inside_h
                        },
                    );
                    overlay_line(
                        &mut resolved.border_b,
                        if row == last_row {
                            &role.borders.bottom
                        } else {
                            &role.borders.inside_h
                        },
                    );
                }
                TablePartRegion::Row => {
                    overlay_line(
                        &mut resolved.border_l,
                        if col == 0 {
                            &role.borders.left
                        } else {
                            &role.borders.inside_v
                        },
                    );
                    overlay_line(
                        &mut resolved.border_r,
                        if col == last_col {
                            &role.borders.right
                        } else {
                            &role.borders.inside_v
                        },
                    );
                    overlay_line(&mut resolved.border_t, &role.borders.top);
                    overlay_line(&mut resolved.border_b, &role.borders.bottom);
                }
                TablePartRegion::Column => {
                    overlay_line(&mut resolved.border_l, &role.borders.left);
                    overlay_line(&mut resolved.border_r, &role.borders.right);
                    overlay_line(
                        &mut resolved.border_t,
                        if row == 0 {
                            &role.borders.top
                        } else {
                            &role.borders.inside_h
                        },
                    );
                    overlay_line(
                        &mut resolved.border_b,
                        if row == last_row {
                            &role.borders.bottom
                        } else {
                            &role.borders.inside_h
                        },
                    );
                }
                TablePartRegion::Cell => {
                    overlay_line(&mut resolved.border_l, &role.borders.left);
                    overlay_line(&mut resolved.border_r, &role.borders.right);
                    overlay_line(&mut resolved.border_t, &role.borders.top);
                    overlay_line(&mut resolved.border_b, &role.borders.bottom);
                }
            }
        };

    overlay(&mut resolved, &style.whole_tbl, TablePartRegion::Whole);
    if flags.band_row {
        let band_row = row.saturating_sub(usize::from(flags.first_row));
        overlay(
            &mut resolved,
            if band_row % 2 == 0 {
                &style.band1_h
            } else {
                &style.band2_h
            },
            TablePartRegion::Row,
        );
    }
    if flags.band_col {
        let band_col = col.saturating_sub(usize::from(flags.first_col));
        overlay(
            &mut resolved,
            if band_col % 2 == 0 {
                &style.band1_v
            } else {
                &style.band2_v
            },
            TablePartRegion::Column,
        );
    }
    if flags.first_row && row == 0 {
        overlay(&mut resolved, &style.first_row, TablePartRegion::Row);
    }
    if flags.last_row && row == last_row {
        overlay(&mut resolved, &style.last_row, TablePartRegion::Row);
    }
    if flags.first_col && col == 0 {
        overlay(&mut resolved, &style.first_col, TablePartRegion::Column);
    }
    if flags.last_col && col == last_col {
        overlay(&mut resolved, &style.last_col, TablePartRegion::Column);
    }
    if flags.first_row && row == 0 && flags.first_col && col == 0 {
        overlay(&mut resolved, &style.nw_cell, TablePartRegion::Cell);
    }
    if flags.first_row && row == 0 && flags.last_col && col == last_col {
        overlay(&mut resolved, &style.ne_cell, TablePartRegion::Cell);
    }
    if flags.last_row && row == last_row && flags.first_col && col == 0 {
        overlay(&mut resolved, &style.sw_cell, TablePartRegion::Cell);
    }
    if flags.last_row && row == last_row && flags.last_col && col == last_col {
        overlay(&mut resolved, &style.se_cell, TablePartRegion::Cell);
    }

    resolved
}

/// Apply the resolved style below direct `tcPr` formatting. This separate step
/// makes the precedence rule testable without constructing a complete package.
pub(crate) fn apply_resolved_table_cell_style(
    cell: &mut TableCell,
    effective: ResolvedTableCellStyle,
) {
    if !cell.has_direct_fill {
        cell.fill = effective.fill;
    }
    if cell.text_color.is_none() {
        cell.text_color = effective.text.color;
    }
    if let (Some(b), Some(tb)) = (effective.text.bold, cell.text_body.as_mut()) {
        if tb.default_bold.is_none() {
            tb.default_bold = Some(b);
        }
    }
    if let (Some(i), Some(tb)) = (effective.text.italic, cell.text_body.as_mut()) {
        if tb.default_italic.is_none() {
            tb.default_italic = Some(i);
        }
    }
    if !cell.has_direct_border_l {
        effective.border_l.apply_to(&mut cell.border_l);
    }
    if !cell.has_direct_border_r {
        effective.border_r.apply_to(&mut cell.border_r);
    }
    if !cell.has_direct_border_t {
        effective.border_t.apply_to(&mut cell.border_t);
    }
    if !cell.has_direct_border_b {
        effective.border_b.apply_to(&mut cell.border_b);
    }
    if !cell.has_direct_diagonal_tl {
        effective.diagonal_tl.apply_to(&mut cell.diagonal_tl);
    }
    if !cell.has_direct_diagonal_tr {
        effective.diagonal_tr.apply_to(&mut cell.diagonal_tr);
    }
}

pub(crate) fn parse_table(
    tbl: roxmltree::Node<'_, '_>,
    t: &Transform,
    theme_source: &(impl PptxThemeSource + ?Sized),
    rels: &HashMap<String, String>,
    source_dir: &str,
    zip: &mut PptxZip,
) -> Option<TableElement> {
    let theme = theme_source.colors();
    // Parse tblPr attributes and look up table style
    let tbl_pr = child(tbl, "tblPr");
    let style_id = tbl_pr
        .and_then(|n| child(n, "tableStyleId"))
        .and_then(|n| n.text())
        .map(|s| s.to_string());
    let flag = |attr_name: &str| -> bool {
        tbl_pr
            .and_then(|n| attr(&n, attr_name))
            .map(|v| v == "1" || v == "true")
            .unwrap_or(false)
    };
    let first_row = flag("firstRow");
    let last_row = flag("lastRow");
    // ECMA-376 §21.1.3.13 (a:tblPr@rtl): right-to-left table layout.
    let rtl = flag("rtl");
    let band_row = flag("bandRow");
    let band_col = flag("bandCol");
    let first_col = flag("firstCol");
    let last_col = flag("lastCol");

    // Load style definitions once
    let table_styles_xml = read_zip_str(zip, "ppt/tableStyles.xml").ok();
    let table_styles = table_styles_xml
        .as_deref()
        .map(|xml| parse_table_styles_xml(xml, theme_source))
        .unwrap_or_default();
    let style_owned: Option<TableStyleDef> = style_id.as_deref().and_then(|id| {
        table_styles
            .get(id)
            .cloned()
            .or_else(|| table_style_presets::lookup_builtin_table_style(id, theme))
    });
    let style = style_owned.as_ref();

    let cols: Vec<i64> = tbl
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "tblGrid")
        .map(|grid| {
            grid.children()
                .filter(|n| n.is_element() && n.tag_name().name() == "gridCol")
                .filter_map(|n| attr_i64(&n, "w"))
                .collect()
        })
        .unwrap_or_default();

    if cols.is_empty() {
        return None;
    }

    let col_count = cols.len();
    let last_col_idx = col_count.saturating_sub(1);

    let mut rows: Vec<TableRow> = tbl
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "tr")
        .map(|tr| parse_table_row(tr, theme, rels, source_dir, zip))
        .collect();

    let row_count = rows.len();
    let last_row_idx = row_count.saturating_sub(1);
    let style_flags = TableStyleFlags {
        first_row,
        last_row,
        first_col,
        last_col,
        band_row,
        band_col,
    };

    // Apply table style fills and borders to each cell
    for (ri, row) in rows.iter_mut().enumerate() {
        for (ci, cell) in row.cells.iter_mut().enumerate() {
            if let Some(s) = style {
                let effective =
                    resolve_table_cell_style(s, style_flags, ri, ci, row_count, col_count);

                // Direct `tcPr` formatting is the final tier. The presence flags
                // preserve an authored noFill/no-line, which is not equivalent to
                // an omitted property inheriting the table style.
                apply_resolved_table_cell_style(cell, effective);
            } else {
                // ── Fallback for built-in styles not defined in tableStyles.xml ──
                // Approximate "Medium Style 2": accent1 header fill + thin outer box + row separators.
                let thin = Stroke {
                    color: "A0A096".to_string(),
                    width: 9525,
                    fill: None,
                    dash_style: None,
                    custom_dash: Vec::new(),
                    line_cap: None,
                    line_join: None,
                    miter_limit: None,
                    alignment: None,
                    head_end: None,
                    tail_end: None,
                    cmpd: None,
                };
                if !cell.has_direct_fill && cell.fill.is_none() && first_row && ri == 0 {
                    if let Some(color) = theme.get("accent1") {
                        cell.fill = Some(Fill::Solid {
                            color: color.clone(),
                        });
                    }
                }
                // Outer top
                if !cell.has_direct_border_t && cell.border_t.is_none() && ri == 0 {
                    cell.border_t = Some(thin.clone());
                }
                // Inner horizontal separators
                if !cell.has_direct_border_t && cell.border_t.is_none() && ri > 0 {
                    cell.border_t = Some(thin.clone());
                }
                // Outer bottom
                if !cell.has_direct_border_b && cell.border_b.is_none() && ri == last_row_idx {
                    cell.border_b = Some(thin.clone());
                }
                // Outer left edge
                if !cell.has_direct_border_l && cell.border_l.is_none() && ci == 0 {
                    cell.border_l = Some(thin.clone());
                }
                // Outer right edge
                if !cell.has_direct_border_r && cell.border_r.is_none() && ci == last_col_idx {
                    cell.border_r = Some(thin.clone());
                }
            }
        }
    }

    Some(TableElement {
        id: None,
        x: t.x,
        y: t.y,
        width: t.cx,
        height: t.cy,
        rotation: t.rot,
        flip_h: t.flip_h,
        flip_v: t.flip_v,
        cols,
        rows,
        rtl,
    })
}

pub(crate) fn parse_table_row(
    tr: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    rels: &HashMap<String, String>,
    source_dir: &str,
    zip: &mut PptxZip,
) -> TableRow {
    let height = attr_i64(&tr, "h").unwrap_or(0);
    let cells: Vec<TableCell> = tr
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "tc")
        .map(|tc| parse_table_cell(tc, theme, rels, source_dir, zip))
        .collect();
    TableRow { height, cells }
}

pub(crate) fn parse_table_cell(
    tc: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    rels: &HashMap<String, String>,
    source_dir: &str,
    zip: &mut PptxZip,
) -> TableCell {
    let tc_pr = child(tc, "tcPr");
    // tcPr > anchor controls vertical text alignment within the cell
    let anchor = tc_pr
        .and_then(|n| attr(&n, "anchor"))
        .map(|a| a.to_string());
    let text_direction = tc_pr.and_then(|n| attr(&n, "vert"));
    let text_insets = tc_pr.map(|n| {
        [
            attr_i64(&n, "marL"),
            attr_i64(&n, "marT"),
            attr_i64(&n, "marR"),
            attr_i64(&n, "marB"),
        ]
    });
    let text_body = child(tc, "txBody").map(|n| {
        let mut body = parse_text_body(
            n,
            theme,
            rels,
            source_dir,
            None,
            None,
            [None; 9],
            Default::default(), // inherited_level_indents
            &empty_level_bullets(),
            None,
            None,
            None,
            None, // inherited_reflection
            anchor,
            text_insets,
            None, // inherited_alignment
            None, // inherited_ea_ln_brk
            None, // inherited_space_before
            None, // inherited_space_after
            None, // inherited_line_spacing
            ShapeKind::Sp,
            zip,
        );
        // Table-cell text direction is authored on tcPr rather than txBody's
        // bodyPr (ECMA-376 §21.1.3.17 CT_TableCellProperties@vert).
        if let Some(direction) = text_direction.as_deref() {
            body.vert = direction.to_owned();
        }
        body
    });

    let fill = tc_pr.and_then(|n| parse_fill(n, theme));

    // ECMA-376 §21.1.3.17 (`CT_TableCellProperties`) makes direct formatting
    // the final cascade tier. Presence must be tracked separately because
    // direct `noFill`/no-line legitimately parse to no painted stroke.
    let has_direct_fill = tc_pr.is_some_and(|n| {
        [
            "noFill",
            "solidFill",
            "gradFill",
            "blipFill",
            "pattFill",
            "grpFill",
        ]
        .iter()
        .any(|name| child(n, name).is_some())
    });
    let has_direct_border_l = tc_pr.is_some_and(|n| child(n, "lnL").is_some());
    let has_direct_border_r = tc_pr.is_some_and(|n| child(n, "lnR").is_some());
    let has_direct_border_t = tc_pr.is_some_and(|n| child(n, "lnT").is_some());
    let has_direct_border_b = tc_pr.is_some_and(|n| child(n, "lnB").is_some());
    let has_direct_diagonal_tl = tc_pr.is_some_and(|n| child(n, "lnTlToBr").is_some());
    let has_direct_diagonal_tr = tc_pr.is_some_and(|n| child(n, "lnBlToTr").is_some());

    let border_l = tc_pr
        .and_then(|n| child(n, "lnL"))
        .and_then(|n| parse_stroke(n, theme));
    let border_r = tc_pr
        .and_then(|n| child(n, "lnR"))
        .and_then(|n| parse_stroke(n, theme));
    let border_t = tc_pr
        .and_then(|n| child(n, "lnT"))
        .and_then(|n| parse_stroke(n, theme));
    let border_b = tc_pr
        .and_then(|n| child(n, "lnB"))
        .and_then(|n| parse_stroke(n, theme));
    // Diagonal borders: lnTlToBr (top-left→bottom-right) and lnBlToTr (bottom-left→top-right)
    // These are CT_LineProperties elements (same structure as lnL/lnR/lnT/lnB)
    let diagonal_tl = tc_pr
        .and_then(|n| child(n, "lnTlToBr"))
        .and_then(|n| parse_stroke(n, theme));
    let diagonal_tr = tc_pr
        .and_then(|n| child(n, "lnBlToTr"))
        .and_then(|n| parse_stroke(n, theme));

    let grid_span: u32 = attr(&tc, "gridSpan")
        .and_then(|v| v.parse().ok())
        .unwrap_or(1);
    let row_span: u32 = attr(&tc, "rowSpan")
        .and_then(|v| v.parse().ok())
        .unwrap_or(1);
    let h_merge = attr(&tc, "hMerge")
        .map(|v| v == "1" || v == "true")
        .unwrap_or(false);
    let v_merge = attr(&tc, "vMerge")
        .map(|v| v == "1" || v == "true")
        .unwrap_or(false);

    TableCell {
        text_body,
        fill,
        has_direct_fill,
        text_color: None,
        border_l,
        border_r,
        border_t,
        border_b,
        has_direct_border_l,
        has_direct_border_r,
        has_direct_border_t,
        has_direct_border_b,
        has_direct_diagonal_tl,
        has_direct_diagonal_tr,
        diagonal_tl,
        diagonal_tr,
        grid_span,
        row_span,
        h_merge,
        v_merge,
    }
}

/// Walk a master/layout `<p:cSld><p:spTree>` and collect its NON-placeholder
/// (decorative) shapes as `SlideElement`s (ECMA-376 §19.3.1.38 showMasterSp /
/// layout decorations). Placeholders are skipped (`skip_placeholders = true`).
/// Extracted verbatim from the two inline spTree walks in `parse_slide` so the
/// master decorations can be pre-computed once per cached master; `theme`,
/// `rels`, `smartart_drawings`, and `part_dir` are the resolution context of the
/// tree being walked. Appends to `out` (callers control ordering/gating).
pub(crate) fn extract_decorative_shapes(
    root: roxmltree::Node<'_, '_>,
    part_dir: &str,
    rels: &HashMap<String, String>,
    smartart_drawings: &HashMap<String, String>,
    theme: &(impl PptxThemeSource + ?Sized),
    zip: &mut PptxZip,
    out: &mut Vec<SlideElement>,
) {
    if let Some(sp_tree) = child(root, "cSld").and_then(|n| child(n, "spTree")) {
        let empty_lph = LayoutPlaceholders::default();
        for node in sp_tree.children().filter(|n| n.is_element()) {
            parse_sp_tree_node(
                node,
                &empty_lph,
                part_dir,
                rels,
                smartart_drawings,
                zip,
                theme,
                out,
                true, // skip placeholder shapes
                None, // no inherited group fill at top level
                DepthGuard::root(),
            );
        }
    }
}

// Recurses the shape tree carrying placeholder/layout context, theme, rels,
// smartart drawings, the zip, the output buffer and inherited group fill.
//
// `depth` bounds the `<p:grpSp>` recursion: a hand-crafted, thousands-deep group
// nest would otherwise blow the fixed WASM stack and trap the whole parse. Past
// the shared limit the group's children are dropped and parsing continues with
// the rest of the slide (graceful degradation). See `ooxml_common::depth`.
#[allow(clippy::too_many_arguments)]
pub(crate) fn parse_sp_tree_node(
    node: roxmltree::Node<'_, '_>,
    lph: &LayoutPlaceholders,
    slide_dir: &str,
    rels: &HashMap<String, String>,
    smartart_drawings: &HashMap<String, String>,
    zip: &mut PptxZip,
    theme_source: &(impl PptxThemeSource + ?Sized),
    out: &mut Vec<SlideElement>,
    skip_placeholders: bool,
    group_fill: Option<&Fill>,
    depth: DepthGuard,
) {
    let theme = theme_source.colors();
    // §20.1.2.2.8 — a shape/pic/graphicFrame/cxnSp/grpSp whose own
    // `<p:cNvPr hidden="1">` marks it hidden is not rendered. Skip at parse
    // time (no viewer-side "show hidden" mode is meaningful for a shape, unlike
    // a whole slide). A hidden grpSp elides its entire subtree here.
    // `mc:AlternateContent` has no cNvPr of its own, so it recurses normally.
    if node_is_hidden(node) {
        return;
    }
    match node.tag_name().name() {
        "sp" => {
            if skip_placeholders && is_placeholder(node) {
                return;
            }
            // Image-filled shape: spPr > blipFill > blip r:embed → render as PictureElement
            let sp_pr_node = child(node, "spPr");
            let blip_fill_node = sp_pr_node.and_then(|p| child(p, "blipFill"));
            let blip_rid = blip_fill_node
                .and_then(|bf| child(bf, "blip"))
                .and_then(|b| attr_r(&b, "embed"));
            if let Some(ref rid) = blip_rid {
                if let Some(xfrm_node) = sp_pr_node.and_then(|p| child(p, "xfrm")) {
                    let t = parse_xfrm(xfrm_node);
                    if t.cx > 0 && t.cy > 0 {
                        if let Some(target) = rels.get(rid) {
                            let image_path = resolve_path(slide_dir, target);
                            if let Ok(bytes) = read_zip_head(zip, &image_path, 24) {
                                let mime_type = mime_from_ext(&image_path).to_owned();
                                let (intrinsic_width_px, intrinsic_height_px) =
                                    match png_size_from_bytes(&bytes) {
                                        Some((w, h)) => (Some(w), Some(h)),
                                        None => (None, None),
                                    };
                                // Microsoft 2016 SVG extension — a blipFill-painted
                                // sp can carry the same svgBlip vector original as a
                                // real p:pic; surface it so the renderer prefers it.
                                let svg_image_path = blip_fill_node
                                    .and_then(|bf| child(bf, "blip"))
                                    .and_then(|b| svg_blip_path(b, slide_dir, rels, zip));
                                // §20.1.9.18 — the sp's prstGeom (any preset, not
                                // just roundRect) is the picture's clip silhouette.
                                let (prst_geom, prst_adjust) =
                                    sp_pr_node.map(parse_pic_prst_geom).unwrap_or((None, None));
                                let cust_geom =
                                    sp_pr_node
                                        .and_then(|p| child(p, "custGeom"))
                                        .map(|geometry| {
                                            parse_cust_geom(geometry, t.cx as f64, t.cy as f64)
                                        });
                                let PictureShapeProperties {
                                    stroke,
                                    shadow,
                                    inner_shadow,
                                    glow,
                                    soft_edge,
                                    reflection,
                                    scene3d,
                                    sp3d,
                                } = resolve_picture_shape_properties(
                                    sp_pr_node,
                                    child(node, "style"),
                                    None,
                                    theme_source,
                                );
                                out.push(SlideElement::Picture(PictureElement {
                                    id: read_cnv_pr(node, rels).id,
                                    x: t.x,
                                    y: t.y,
                                    width: t.cx,
                                    height: t.cy,
                                    rotation: t.rot,
                                    flip_h: t.flip_h,
                                    flip_v: t.flip_v,
                                    image_path,
                                    mime_type,
                                    svg_image_path,
                                    intrinsic_width_px,
                                    intrinsic_height_px,
                                    stroke,
                                    prst_geom,
                                    prst_adjust,
                                    src_rect: blip_fill_node.and_then(parse_src_rect),
                                    alpha: blip_fill_node.and_then(parse_blip_alpha),
                                    duotone: blip_fill_node.and_then(|bf| {
                                        parse_blip_duotone(
                                            bf,
                                            &PptxSchemeResolver { theme },
                                            ooxml_common::color::TintMode::PowerPointLinear,
                                        )
                                    }),
                                    cust_geom,
                                    shadow,
                                    inner_shadow,
                                    glow,
                                    soft_edge,
                                    reflection,
                                    scene3d,
                                    sp3d,
                                }));
                                return;
                            }
                        }
                    }
                }
            }
            // Picture-placeholder inheritance: slide sp has a ph but no own blipFill →
            // look up an inherited blipFill from the layout placeholder. Transform
            // comes from the slide's xfrm when present, otherwise from the layout.
            if blip_rid.is_none() {
                if let Some(ph) = node
                    .descendants()
                    .find(|n| n.is_element() && n.tag_name().name() == "ph")
                {
                    let ph_type = attr(&ph, "type").unwrap_or_else(|| "body".into());
                    let ph_idx: Option<u32> = attr(&ph, "idx")
                        .and_then(|v| v.parse().ok())
                        .filter(|idx| *idx != u32::MAX);
                    if let Some(bf) = lph.lookup_blip_fill(&ph_type, ph_idx) {
                        let slide_xfrm = sp_pr_node.and_then(|p| child(p, "xfrm")).map(parse_xfrm);
                        let t = slide_xfrm.or_else(|| lph.lookup(&ph_type, ph_idx).cloned());
                        if let Some(t) = t {
                            if t.cx > 0 && t.cy > 0 {
                                let PictureShapeProperties {
                                    stroke,
                                    shadow,
                                    inner_shadow,
                                    glow,
                                    soft_edge,
                                    reflection,
                                    scene3d,
                                    sp3d,
                                } = resolve_picture_shape_properties(
                                    sp_pr_node,
                                    child(node, "style"),
                                    lph.lookup_picture_properties(&ph_type, ph_idx),
                                    theme_source,
                                );
                                out.push(SlideElement::Picture(PictureElement {
                                    id: read_cnv_pr(node, rels).id,
                                    x: t.x,
                                    y: t.y,
                                    width: t.cx,
                                    height: t.cy,
                                    rotation: t.rot,
                                    flip_h: t.flip_h,
                                    flip_v: t.flip_v,
                                    image_path: bf.image_path,
                                    mime_type: bf.mime_type,
                                    // TODO: an inherited layout-placeholder blipFill
                                    // (LayoutPlaceholders::lookup_blip_fill) does not
                                    // yet carry the svgBlip extension. Picture
                                    // placeholders pointing at an SVG are rare; thread
                                    // the svg path through BlipFill if a sample needs it.
                                    svg_image_path: None,
                                    // Intrinsic size is only consumed by the ink
                                    // fallback (PNG-IHDR centering); inherited
                                    // placeholder pictures stretch to the box, so
                                    // None matches the prior behaviour.
                                    intrinsic_width_px: None,
                                    intrinsic_height_px: None,
                                    stroke,
                                    prst_geom: None,
                                    prst_adjust: None,
                                    src_rect: bf.src_rect,
                                    alpha: bf.alpha,
                                    // §20.1.8.23 duotone inherited from the layout
                                    // placeholder's blipFill (resolved through the
                                    // theme in InheritedBlipFill); the PictureElement
                                    // render applies it via the shared core cache.
                                    duotone: bf.duotone,
                                    cust_geom: None,
                                    shadow,
                                    inner_shadow,
                                    glow,
                                    soft_edge,
                                    reflection,
                                    scene3d,
                                    sp3d,
                                }));
                                return;
                            }
                        }
                    }
                }
            }
            if let Some(shape) =
                parse_shape(node, lph, theme_source, rels, slide_dir, group_fill, zip)
            {
                out.push(SlideElement::Shape(shape));
            }
        }
        "pic" => {
            if let Some(media) = parse_media(node, slide_dir, rels) {
                out.push(SlideElement::Media(media));
            } else if let Some(pic) = parse_picture(node, slide_dir, rels, theme_source, zip) {
                out.push(SlideElement::Picture(pic));
            } else {
                // Placeholder pic: no xfrm in spPr — position comes from layout by_idx
                let ph_node = node
                    .descendants()
                    .find(|n| n.is_element() && n.tag_name().name() == "ph");
                let ph_type = ph_node
                    .and_then(|ph| attr(&ph, "type"))
                    .unwrap_or_else(|| "body".to_owned());
                let ph_idx = ph_node
                    .and_then(|ph| attr(&ph, "idx"))
                    .and_then(|s| s.parse::<u32>().ok())
                    .filter(|idx| *idx != u32::MAX);
                if let Some(idx) = ph_idx {
                    if let Some(t) = lph.by_idx.get(&idx) {
                        let blip_fill = child(node, "blipFill");
                        let blip = blip_fill.and_then(|bf| child(bf, "blip"));
                        let r_id = blip.and_then(|b| attr_r(&b, "embed"));
                        if let Some(rid) = r_id {
                            if let Some(rel_target) = rels.get(&rid) {
                                let image_path = resolve_path(slide_dir, rel_target);
                                if let Ok(image_bytes) = read_zip_head(zip, &image_path, 24) {
                                    let mime_type = mime_from_ext(&image_path).to_owned();
                                    let (intrinsic_width_px, intrinsic_height_px) =
                                        match png_size_from_bytes(&image_bytes) {
                                            Some((w, h)) => (Some(w), Some(h)),
                                            None => (None, None),
                                        };
                                    // Microsoft 2016 SVG extension on the placeholder
                                    // p:pic's blip — prefer the vector original.
                                    let svg_image_path =
                                        blip.and_then(|b| svg_blip_path(b, slide_dir, rels, zip));
                                    let PictureShapeProperties {
                                        stroke,
                                        shadow,
                                        inner_shadow,
                                        glow,
                                        soft_edge,
                                        reflection,
                                        scene3d,
                                        sp3d,
                                    } = resolve_picture_shape_properties(
                                        child(node, "spPr"),
                                        child(node, "style"),
                                        lph.lookup_picture_properties(&ph_type, Some(idx)),
                                        theme_source,
                                    );
                                    out.push(SlideElement::Picture(PictureElement {
                                        id: read_cnv_pr(node, rels).id,
                                        x: t.x,
                                        y: t.y,
                                        width: t.cx,
                                        height: t.cy,
                                        rotation: t.rot,
                                        flip_h: t.flip_h,
                                        flip_v: t.flip_v,
                                        image_path,
                                        mime_type,
                                        svg_image_path,
                                        intrinsic_width_px,
                                        intrinsic_height_px,
                                        stroke,
                                        prst_geom: None,
                                        prst_adjust: None,
                                        src_rect: blip_fill.and_then(parse_src_rect),
                                        alpha: blip_fill.and_then(parse_blip_alpha),
                                        duotone: blip_fill.and_then(|bf| {
                                            parse_blip_duotone(
                                                bf,
                                                &PptxSchemeResolver { theme },
                                                ooxml_common::color::TintMode::PowerPointLinear,
                                            )
                                        }),
                                        cust_geom: None,
                                        shadow,
                                        inner_shadow,
                                        glow,
                                        soft_edge,
                                        reflection,
                                        scene3d,
                                        sp3d,
                                    }));
                                }
                            }
                        }
                    }
                }
            }
        }
        "AlternateContent" => {
            let before = out.len();
            // ECMA-376 Part 3 §9.3 — select the first Choice whose Requires namespaces
            // are all understood, else the Fallback (replaces the old output-emptiness
            // heuristic that never inspected Requires — issue #787).
            let selected = ooxml_common::mce::select_alternate_content(node, &pptx_understands_ns);
            // Ink/handwriting: a `<p:contentPart>` in any Choice is InkML PowerPoint
            // renders directly and we cannot; its `p14` Choice is not understood, so the
            // Fallback PNG is selected. Detect it so the PNG is drawn at its natural size
            // (a few px for empty/single-tap strokes) rather than stretched to the box.
            let is_ink = node
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "Choice")
                .any(|c| {
                    c.descendants()
                        .any(|n| n.is_element() && n.tag_name().name() == "contentPart")
                });
            let selected_is_fallback = selected.is_some_and(|s| s.tag_name().name() == "Fallback");
            if let Some(sel) = selected {
                for child_node in sel.children().filter(|n| n.is_element()) {
                    // `mc:AlternateContent` is a transparent wrapper, not a group level —
                    // pass `depth` through unchanged.
                    parse_sp_tree_node(
                        child_node,
                        lph,
                        slide_dir,
                        rels,
                        smartart_drawings,
                        zip,
                        theme_source,
                        out,
                        skip_placeholders,
                        group_fill,
                        depth,
                    );
                }
            }
            if is_ink && selected_is_fallback {
                // Render the fallback PNG at its natural pixel size centered
                // inside the original bounding box, so a 6×6 px empty-stroke
                // PNG is drawn as a 6×6 px dot rather than stretched into a
                // blocky cross. Visible strokes whose PNG natural size already
                // matches the box keep their existing extent.
                for el in &mut out[before..] {
                    if let SlideElement::Picture(p) = el {
                        if let (Some(nat_w_px), Some(nat_h_px)) =
                            (p.intrinsic_width_px, p.intrinsic_height_px)
                        {
                            let nat_w = (nat_w_px as i64) * EMU_PER_PX_96DPI;
                            let nat_h = (nat_h_px as i64) * EMU_PER_PX_96DPI;
                            if nat_w < p.width && nat_h < p.height {
                                p.x += (p.width - nat_w) / 2;
                                p.y += (p.height - nat_h) / 2;
                                p.width = nat_w;
                                p.height = nat_h;
                            }
                        }
                    }
                }
            }
        }
        "graphicFrame" => {
            let xfrm_node = child(node, "xfrm");
            let t = xfrm_node.map(parse_xfrm).unwrap_or_default();

            // Table
            let tbl_node = node
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "tbl");
            if let Some(tbl_node) = tbl_node {
                if let Some(mut table) =
                    parse_table(tbl_node, &t, theme_source, rels, slide_dir, zip)
                {
                    table.id = read_cnv_pr(node, rels).id;
                    out.push(SlideElement::Table(table));
                }
                return;
            }

            // SmartArt (ECMA-376 §21.4). Preferred path: replay the
            // PowerPoint-prebaked drawing (`drawingN.xml` / `dsp:spTree`) when
            // Office persisted one — that carries the true layout geometry.
            //
            // Fallback (PP3): when no drawing part exists, recover the diagram's
            // *content* from its data model (`dataN.xml`) — every node's text,
            // indented by its parent/child depth — as a bulleted list filling the
            // frame. Layout-engine reconstruction (hierarchy/cycle geometry) is
            // still not implemented; see `smartart_fallback`.
            if let Some(gd) = node
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "graphicData")
            {
                let uri = attr(&gd, "uri").unwrap_or_default();
                if is_diagram_uri(&uri) {
                    if let Some(rel_ids) = child(gd, "relIds") {
                        if let Some(dm_rid) = attr_r(&rel_ids, "dm") {
                            if let Some(drawing_xml) = smartart_drawings.get(&dm_rid) {
                                parse_smartart_drawing(drawing_xml, &t, theme, out, zip);
                                return;
                            }
                            // No prebaked drawing → data-model fallback. `rels` are
                            // the referencing part's, so `rels[dm_rid]` resolved
                            // against `slide_dir` is the data part (§21.4.2.22
                            // relIds `r:dm`). Emits M (content list) or S
                            // (placeholder); either way this graphicData is a
                            // diagram, so we return regardless of the outcome.
                            crate::smartart_fallback::emit_smartart_fallback(
                                &dm_rid, &t, slide_dir, rels, theme, zip, out,
                            );
                            return;
                        }
                    }
                }
            }

            // Chart
            if let Some(gd) = node
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "graphicData")
            {
                let uri = attr(&gd, "uri").unwrap_or_default();
                // Both <c:chart> and <cx:chart> share the local name "chart"
                let chart_rid = gd
                    .descendants()
                    .find(|n| n.is_element() && n.tag_name().name() == "chart")
                    .and_then(|n| attr_r(&n, "id"));
                if let Some(rid) = chart_rid {
                    if let Some(rel_target) = rels.get(&rid) {
                        let chart_path = resolve_path(slide_dir, rel_target);
                        if let Ok(chart_xml) = read_zip_str(zip, &chart_path) {
                            let related_parts = load_chart_related_parts(zip, &chart_path);
                            let empty_theme_images =
                                ooxml_common::chart::ChartImageRelationships::default();
                            let image_resolver = ooxml_common::chart::ChartImageResolverChain::new(
                                &related_parts.image_relationships,
                                theme_source.chart_images().unwrap_or(&empty_theme_images),
                            );
                            let chart_opt = if uri.contains("chartex") || uri.contains("chartEx") {
                                // chartEx title font size lives in the chart
                                // part's associated chartStyle sidecar
                                // (`styleN.xml`), reached via that part's OWN
                                // rels. Read it best-effort before parsing.
                                parse_chartex_with_images(
                                    &chart_xml,
                                    related_parts.style_xml.as_deref(),
                                    related_parts.color_style_xml.as_deref(),
                                    theme,
                                    theme_source.format_scheme(),
                                    &image_resolver,
                                )
                            } else {
                                let user_shapes_xml =
                                    load_chart_user_shapes_xml(zip, &chart_path, &chart_xml);
                                parse_legacy_chart_with_style_parts_and_images(
                                    &chart_xml,
                                    related_parts.style_xml.as_deref(),
                                    related_parts.color_style_xml.as_deref(),
                                    user_shapes_xml.as_deref(),
                                    theme,
                                    theme_source.format_scheme(),
                                    &image_resolver,
                                )
                            };
                            if let Some(mut chart) = chart_opt {
                                chart.id = read_cnv_pr(node, rels).id;
                                chart.x = t.x;
                                chart.y = t.y;
                                chart.width = t.cx;
                                chart.height = t.cy;
                                chart.rotation = t.rot;
                                chart.flip_h = t.flip_h;
                                chart.flip_v = t.flip_v;
                                out.push(SlideElement::Chart(chart));
                            }
                        }
                    }
                }
            }

            // OLE embedded object (§19.3.2.4 CT_OleObject). We can't run the
            // embedded application, but the OOXML author bakes a preview `<p:pic>`
            // inside `<p:oleObj>` for exactly this case. Route that picture through
            // the ordinary image pipeline so the object is visible instead of a
            // silent hole.
            if let Some(gd) = node
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "graphicData")
            {
                let uri = attr(&gd, "uri").unwrap_or_default();
                if is_pml_ole_uri(&uri) {
                    // `<p:oleObj>` is commonly wrapped in `mc:AlternateContent`.
                    // PowerPoint's canonical output puts a spid-only oleObj in the
                    // `mc:Choice Requires="v"` (VML) branch and the drawable
                    // `<p:pic>`-carrying oleObj in `mc:Fallback` — the `spid`
                    // attribute and a `<p:pic>` child are mutually exclusive on
                    // CT_OleObject (ECMA-376 Part 3 §B.1). Since we do not
                    // understand the `v` (VML) namespace, MCE §9.3 says to take the
                    // representation we can render, i.e. the Fallback's pic. Rather
                    // than depend on document order or Choice/Fallback wrapping,
                    // select the first oleObj that actually carries a `<p:pic>`
                    // (the capability predicate "has something drawable"). This also
                    // makes a bare, unwrapped `<p:oleObj><p:pic/>` work unchanged.
                    let ole_pic = gd
                        .descendants()
                        .filter(|n| n.is_element() && n.tag_name().name() == "oleObj")
                        .find_map(|ole| child(ole, "pic"));
                    if let Some(pic_node) = ole_pic {
                        // The preview pic's own `<a:xfrm>` is 0,0-relative (or
                        // absent), so the graphicFrame's `<p:xfrm>` is authoritative
                        // for placement. Parse the blip, then stamp gf geometry.
                        if let Some(mut pic) =
                            parse_picture(pic_node, slide_dir, rels, theme_source, zip)
                        {
                            pic.x = t.x;
                            pic.y = t.y;
                            pic.width = t.cx;
                            pic.height = t.cy;
                            out.push(SlideElement::Picture(pic));
                        } else if let Some(pic) =
                            parse_ole_preview_picture(pic_node, &t, slide_dir, rels, theme, zip)
                        {
                            // The pic lacked an `<a:xfrm>` (parse_picture requires
                            // one), but still carries a resolvable blip — build the
                            // element directly on the graphicFrame geometry.
                            out.push(SlideElement::Picture(pic));
                        }
                    }
                }
            }
        }
        "grpSp" => {
            let grp_sp_pr = child(node, "grpSpPr");
            let gt: Option<GroupTransform> =
                grp_sp_pr.and_then(|pr| child(pr, "xfrm")).map(|xfrm| {
                    let off = child(xfrm, "off");
                    let ext = child(xfrm, "ext");
                    let ch_off = child(xfrm, "chOff");
                    let ch_ext = child(xfrm, "chExt");
                    GroupTransform {
                        x: off.and_then(|n| attr_i64(&n, "x")).unwrap_or(0),
                        y: off.and_then(|n| attr_i64(&n, "y")).unwrap_or(0),
                        cx: ext.and_then(|n| attr_i64(&n, "cx")).unwrap_or(0),
                        cy: ext.and_then(|n| attr_i64(&n, "cy")).unwrap_or(0),
                        ch_x: ch_off.and_then(|n| attr_i64(&n, "x")).unwrap_or(0),
                        ch_y: ch_off.and_then(|n| attr_i64(&n, "y")).unwrap_or(0),
                        ch_cx: ch_ext.and_then(|n| attr_i64(&n, "cx")).unwrap_or(0),
                        ch_cy: ch_ext.and_then(|n| attr_i64(&n, "cy")).unwrap_or(0),
                        flip_h: attr(&xfrm, "flipH")
                            .map(|v| v == "1" || v == "true")
                            .unwrap_or(false),
                        flip_v: attr(&xfrm, "flipV")
                            .map(|v| v == "1" || v == "true")
                            .unwrap_or(false),
                        rot: attr_f64(&xfrm, "rot").unwrap_or(0.0) / 60000.0,
                    }
                });

            // Determine the fill to propagate to child shapes that use grpFill.
            // - Group has solidFill/noFill → use that as child group fill
            // - Group has grpFill → inherit from parent
            // - Group has no fill → inherit from parent
            let grp_has_grp_fill = grp_sp_pr.and_then(|pr| child(pr, "grpFill")).is_some();
            let grp_explicit_fill = grp_sp_pr.and_then(|pr| parse_fill(pr, theme));
            let child_group_fill: Option<Fill> = if grp_has_grp_fill {
                group_fill.cloned()
            } else if let Some(f) = grp_explicit_fill {
                Some(f)
            } else {
                group_fill.cloned()
            };

            // Bound the group nesting: once the shared depth limit is reached,
            // stop descending into this group's children (drop the subtree) so a
            // pathologically deep `<p:grpSp>` nest cannot overflow the stack.
            let start = out.len();
            if let Some(child_depth) = depth.descend() {
                for child_node in node.children().filter(|n| n.is_element()) {
                    parse_sp_tree_node(
                        child_node,
                        lph,
                        slide_dir,
                        rels,
                        smartart_drawings,
                        zip,
                        theme_source,
                        out,
                        skip_placeholders,
                        child_group_fill.as_ref(),
                        child_depth,
                    );
                }
            }
            if let Some(gt) = gt {
                for el in &mut out[start..] {
                    apply_group_transform_to_element(el, &gt);
                }
            }
        }
        "cxnSp" => {
            // Connector shape: parse as a line/shape element
            if skip_placeholders && is_placeholder(node) {
                return;
            }
            if let Some(shape) = parse_connector(node, theme_source, rels) {
                out.push(SlideElement::Shape(shape));
            }
        }
        _ => {}
    }
}

/// Render a SmartArt diagram's pre-baked fallback drawing (drawing1.xml).
/// The drawing file stores a standard dsp:spTree whose shapes share the same
/// element shape as a regular p:sp / p:cxnSp / p:grpSp (children use the `a:`
/// namespace). Coordinates inside the drawing are local to the enclosing
/// graphicFrame, so we translate every emitted element by the graphicFrame's xfrm.
/// Read a SmartArt shape's explicit text frame from `<dsp:txXfrm>`
/// (`<a:off>` / `<a:ext>`), in the drawing's local EMU coordinates. Returns
/// `None` when the shape has no txXfrm (most ordinary shapes). The optional
/// `rot` attribute is intentionally not consumed here — none of the supported
/// SmartArt layouts rotate the text frame independently of the shape, and the
/// renderer already rotates text by the shape's own rotation.
fn parse_tx_xfrm(sp_node: roxmltree::Node<'_, '_>) -> Option<TextRect> {
    let txx = child(sp_node, "txXfrm")?;
    let off = child(txx, "off")?;
    let ext = child(txx, "ext")?;
    Some(TextRect {
        x: attr_i64(&off, "x")?,
        y: attr_i64(&off, "y")?,
        width: attr_i64(&ext, "cx")?,
        height: attr_i64(&ext, "cy")?,
    })
}

fn parse_smartart_drawing(
    drawing_xml: &str,
    gf_xfrm: &Transform,
    theme: &HashMap<String, String>,
    out: &mut Vec<SlideElement>,
    zip: &mut PptxZip,
) {
    let doc = match crate::parse_preflighted_pptx_xml(drawing_xml) {
        Ok(d) => d,
        Err(_) => return,
    };
    let root = doc.root_element();
    let Some(sp_tree) = child(root, "spTree") else {
        return;
    };

    let empty_lph = LayoutPlaceholders::default();

    let start = out.len();
    for node in sp_tree.children().filter(|n| n.is_element()) {
        // dsp:sp / dsp:cxnSp / dsp:grpSp share local names with their p:* counterparts.
        // Shapes inside drawing1.xml carry no `a:blip r:embed` picture fills, so
        // `rels` is an empty stub. `zip` is still threaded through: a dsp shape's
        // text body may declare a `<a:buBlip>` bullet whose part existence must be
        // verified (it resolves against the empty `rels`, so it always falls
        // through to Bullet::Inherit, but the resolver still needs the archive).
        let empty_rels: HashMap<String, String> = HashMap::new();
        let empty_smartart: HashMap<String, String> = HashMap::new();
        // Dispatch the relevant branches (sp/cxnSp/grpSp) directly rather than via
        // parse_sp_tree_node, since the dsp subtree never contains embedded pictures.
        // §20.1.2.2.8 CT_NonVisualDrawingProps@hidden: this manual dsp:sp/cxnSp/
        // grpSp dispatch bypasses parse_sp_tree_node's shared hidden check, so it
        // needs its own — otherwise a hidden SmartArt fallback shape still renders.
        if node_is_hidden(node) {
            continue;
        }
        match node.tag_name().name() {
            "sp" => {
                if let Some(mut shape) =
                    parse_shape(node, &empty_lph, theme, &empty_rels, "", None, zip)
                {
                    shape.text_rect = parse_tx_xfrm(node);
                    out.push(SlideElement::Shape(shape));
                }
            }
            "cxnSp" => {
                if let Some(shape) = parse_connector(node, theme, &empty_rels) {
                    out.push(SlideElement::Shape(shape));
                }
            }
            "grpSp" => {
                // Recursively render a group by collecting its children into a temp buffer
                // and applying the group's xfrm, mirroring the grpSp branch in parse_sp_tree_node.
                let grp_sp_pr = child(node, "grpSpPr");
                let gt: Option<GroupTransform> =
                    grp_sp_pr.and_then(|pr| child(pr, "xfrm")).map(|xfrm| {
                        let off = child(xfrm, "off");
                        let ext = child(xfrm, "ext");
                        let ch_off = child(xfrm, "chOff");
                        let ch_ext = child(xfrm, "chExt");
                        GroupTransform {
                            x: off.and_then(|n| attr_i64(&n, "x")).unwrap_or(0),
                            y: off.and_then(|n| attr_i64(&n, "y")).unwrap_or(0),
                            cx: ext.and_then(|n| attr_i64(&n, "cx")).unwrap_or(0),
                            cy: ext.and_then(|n| attr_i64(&n, "cy")).unwrap_or(0),
                            ch_x: ch_off.and_then(|n| attr_i64(&n, "x")).unwrap_or(0),
                            ch_y: ch_off.and_then(|n| attr_i64(&n, "y")).unwrap_or(0),
                            ch_cx: ch_ext.and_then(|n| attr_i64(&n, "cx")).unwrap_or(0),
                            ch_cy: ch_ext.and_then(|n| attr_i64(&n, "cy")).unwrap_or(0),
                            flip_h: attr(&xfrm, "flipH")
                                .map(|v| v == "1" || v == "true")
                                .unwrap_or(false),
                            flip_v: attr(&xfrm, "flipV")
                                .map(|v| v == "1" || v == "true")
                                .unwrap_or(false),
                            rot: attr_f64(&xfrm, "rot").unwrap_or(0.0) / 60000.0,
                        }
                    });
                let group_start = out.len();
                let _ = (&empty_rels, &empty_smartart); // silence unused warnings when no nested picture
                for child_node in node.children().filter(|n| n.is_element()) {
                    // Same per-node hidden check as the top-level dispatch above —
                    // a hidden shape/connector nested directly inside a (visible)
                    // SmartArt group must still be elided individually.
                    if node_is_hidden(child_node) {
                        continue;
                    }
                    match child_node.tag_name().name() {
                        "sp" => {
                            if let Some(mut shape) = parse_shape(
                                child_node,
                                &empty_lph,
                                theme,
                                &empty_rels,
                                "",
                                None,
                                zip,
                            ) {
                                shape.text_rect = parse_tx_xfrm(child_node);
                                out.push(SlideElement::Shape(shape));
                            }
                        }
                        "cxnSp" => {
                            if let Some(shape) = parse_connector(child_node, theme, &empty_rels) {
                                out.push(SlideElement::Shape(shape));
                            }
                        }
                        _ => {}
                    }
                }
                if let Some(gt) = gt {
                    for el in &mut out[group_start..] {
                        apply_group_transform_to_element(el, &gt);
                    }
                }
            }
            _ => {}
        }
    }

    // Translate all emitted elements by the graphicFrame's position.
    for el in &mut out[start..] {
        offset_slide_element(el, gf_xfrm.x, gf_xfrm.y);
    }
}

fn offset_slide_element(el: &mut SlideElement, dx: i64, dy: i64) {
    match el {
        SlideElement::Shape(s) => {
            s.x += dx;
            s.y += dy;
            if let Some(tr) = &mut s.text_rect {
                tr.x += dx;
                tr.y += dy;
            }
        }
        SlideElement::Picture(p) => {
            p.x += dx;
            p.y += dy;
        }
        SlideElement::Table(t) => {
            t.x += dx;
            t.y += dy;
        }
        SlideElement::Chart(c) => {
            c.x += dx;
            c.y += dy;
        }
        SlideElement::Media(m) => {
            m.x += dx;
            m.y += dy;
        }
    }
}

/// Parse a connector shape (p:cxnSp) as a ShapeElement with line geometry.
fn parse_connector(
    node: roxmltree::Node<'_, '_>,
    theme_source: &(impl PptxThemeSource + ?Sized),
    rels: &HashMap<String, String>,
) -> Option<ShapeElement> {
    let theme = theme_source.colors();
    let sp_pr = child(node, "spPr")?;
    let xfrm = child(sp_pr, "xfrm")?;
    let t = parse_xfrm(xfrm);
    if t.cx == 0 && t.cy == 0 {
        return None;
    }
    let cnv = read_cnv_pr(node, rels);

    // Style-based stroke fallback. `p:style > a:lnRef idx="N"` references the
    // complete theme fmtScheme lnStyleLst entry N (1-based).
    let style_node = child(node, "style");
    let style_stroke = parse_style_stroke(style_node, theme_source);

    let stroke = merge_local_stroke(child(sp_pr, "ln"), style_stroke, theme);

    let style_effects = style_node
        .and_then(|style| child(style, "effectRef"))
        .map(|effect_ref| parse_style_matrix_effects(effect_ref, theme))
        .unwrap_or_default();
    let style_scene3d = style_effects.scene3d;
    let style_sp3d = style_effects.sp3d;
    let local_effect_node = child(sp_pr, "effectLst").or_else(|| child(sp_pr, "effectDag"));
    let EffectLst {
        shadow,
        inner_shadow,
        glow,
        soft_edge,
        reflection,
    } = if local_effect_node.is_some() {
        parse_effect_lst(local_effect_node, theme)
    } else {
        style_effects.effects
    };

    let cy = if t.cy == 0 { 1 } else { t.cy };

    // Preserve the actual preset geometry (bentConnector3, curvedConnector4, …)
    // rather than collapsing every p:cxnSp to "line" — otherwise the renderer
    // can't distinguish bent/curved paths from a straight segment.
    let prst_geom_node = child(sp_pr, "prstGeom");
    let geometry = prst_geom_node
        .and_then(|n| attr(&n, "prst"))
        .unwrap_or_else(|| "line".to_owned());

    // Pull connector adjust values from avLst (e.g. bentConnector3 adj1 = bend %).
    let parse_gd_val = |gd: roxmltree::Node<'_, '_>| -> Option<f64> {
        attr(&gd, "fmla")
            .and_then(|f| f.strip_prefix("val ").map(|s| s.to_owned()))
            .and_then(|s| s.parse::<f64>().ok())
    };
    let av_node = prst_geom_node.and_then(|n| child(n, "avLst"));
    let gd_nodes: Vec<_> = av_node
        .map(|av| {
            av.children()
                .filter(|n| n.is_element() && n.tag_name().name() == "gd")
                .collect()
        })
        .unwrap_or_default();
    let adj: Option<f64> = gd_nodes
        .iter()
        .find(|n| matches!(attr(n, "name").as_deref(), Some("adj") | Some("adj1")))
        .or_else(|| gd_nodes.first())
        .and_then(|n| parse_gd_val(*n));
    let adj2: Option<f64> = gd_nodes
        .iter()
        .find(|n| attr(n, "name").as_deref() == Some("adj2"))
        .or_else(|| gd_nodes.get(1))
        .and_then(|n| parse_gd_val(*n));
    let adj3: Option<f64> = gd_nodes
        .iter()
        .find(|n| attr(n, "name").as_deref() == Some("adj3"))
        .or_else(|| gd_nodes.get(2))
        .and_then(|n| parse_gd_val(*n));
    let adj4: Option<f64> = gd_nodes
        .iter()
        .find(|n| attr(n, "name").as_deref() == Some("adj4"))
        .or_else(|| gd_nodes.get(3))
        .and_then(|n| parse_gd_val(*n));

    Some(ShapeElement {
        x: t.x,
        y: t.y,
        width: t.cx,
        height: cy,
        rotation: t.rot,
        flip_h: t.flip_h,
        flip_v: t.flip_v,
        geometry,
        fill: None,
        stroke,
        text_body: None,
        default_text_color: None,
        cust_geom: None,
        adj,
        adj2,
        adj3,
        adj4,
        adj5: None,
        adj6: None,
        adj7: None,
        adj8: None,
        shadow,
        inner_shadow,
        glow,
        soft_edge,
        reflection,
        id: cnv.id,
        name: cnv.name,
        hyperlink: cnv.hyperlink,
        hyperlink_action: cnv.hyperlink_action,
        placeholder_type: None,
        placeholder_idx: None,
        text_rect: None,
        scene3d: parse_scene3d(sp_pr).or(style_scene3d),
        sp3d: parse_sp3d(sp_pr, theme).or(style_sp3d),
    })
}

#[cfg(test)]
mod style_ref_tests {
    use super::*;
    use crate::master::LayoutPlaceholders;
    use crate::theme::PptxTheme;
    use std::io::Cursor;

    fn empty_zip() -> PptxZip {
        let mut bytes = Vec::new();
        {
            let writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            writer.finish().unwrap();
        }
        PptxZip::new(Cursor::new(bytes)).unwrap()
    }

    #[test]
    fn partial_shape_line_and_effect_refs_inherit_theme_styles() {
        let theme = PptxTheme::from_xml(
            r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:themeElements>
                <a:clrScheme name="Regression">
                  <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
                  <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
                </a:clrScheme>
                <a:fmtScheme name="Regression">
                  <a:fillStyleLst/>
                  <a:lnStyleLst>
                    <a:ln w="9525"/><a:ln w="25400"/>
                  <a:ln w="38100" cmpd="dbl"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
                    <a:prstDash val="dash"/><a:headEnd type="triangle"/>
                  </a:ln>
                  </a:lnStyleLst>
                  <a:effectStyleLst>
                    <a:effectStyle><a:effectLst>
                      <a:outerShdw blurRad="40000" dist="23000" dir="5400000">
                        <a:schemeClr val="phClr"><a:alpha val="35000"/></a:schemeClr>
                      </a:outerShdw>
                    </a:effectLst>
                      <a:scene3d><a:camera prst="isometricLeftUp"/><a:lightRig rig="threePt" dir="t"/></a:scene3d>
                      <a:sp3d prstMaterial="matte"><a:bevelT w="12700" h="6350"/></a:sp3d>
                    </a:effectStyle>
                  </a:effectStyleLst>
                  <a:bgFillStyleLst/>
                </a:fmtScheme>
              </a:themeElements>
            </a:theme>"#,
        );
        let doc = roxmltree::Document::parse(
            r#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:nvSpPr><p:cNvPr id="1" name="Styled"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
              <p:spPr>
                <a:xfrm><a:off x="0" y="0"/><a:ext cx="100000" cy="100000"/></a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
                <a:ln w="44444" cmpd="sng"><a:prstDash val="solid"/><a:headEnd type="none"/><a:tailEnd type="none"/></a:ln>
                <a:effectLst/>
              </p:spPr>
              <p:style>
                <a:lnRef idx="3"><a:schemeClr val="lt1"/></a:lnRef>
                <a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>
                <a:effectRef idx="1"><a:schemeClr val="accent1"/></a:effectRef>
                <a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
              </p:style>
            </p:sp>"#,
        )
        .unwrap();
        let mut zip = empty_zip();

        let shape = parse_shape(
            doc.root_element(),
            &LayoutPlaceholders::default(),
            &theme,
            &HashMap::new(),
            "ppt/slides",
            None,
            &mut zip,
        )
        .unwrap();

        let stroke = shape.stroke.expect("theme line must survive partial a:ln");
        assert_eq!(stroke.color, "FFFFFF");
        assert_eq!(stroke.width, 44444, "local width overrides styled paint");
        assert_eq!(stroke.dash_style, None, "explicit solid clears theme dash");
        assert!(
            stroke.head_end.is_none(),
            "explicit none clears theme arrow"
        );
        assert_eq!(
            stroke.cmpd, None,
            "explicit single clears theme compound line"
        );
        assert!(
            shape.shadow.is_none(),
            "an authored effectLst replaces the theme effect component"
        );
        assert_eq!(
            shape
                .scene3d
                .as_ref()
                .map(|scene| scene.camera.prst.as_str()),
            Some("isometricLeftUp"),
            "effectRef retains the theme scene3d component"
        );
        assert_eq!(
            shape
                .sp3d
                .as_ref()
                .map(|surface| surface.prst_material.as_str()),
            Some("matte"),
            "effectRef retains the theme sp3d component"
        );
    }

    #[test]
    fn fixed_theme_line_color_does_not_require_a_color_on_ln_ref() {
        let theme = PptxTheme::from_xml(
            r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:themeElements>
                <a:clrScheme name="Test"><a:accent1><a:srgbClr val="123456"/></a:accent1></a:clrScheme>
                <a:fmtScheme name="Test"><a:fillStyleLst/><a:lnStyleLst>
                  <a:ln w="19050" cap="rnd" algn="in"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
                    <a:custDash><a:ds d="125000" sp="75000"/></a:custDash><a:miter lim="800000"/>
                  </a:ln>
                </a:lnStyleLst><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>
              </a:themeElements>
            </a:theme>"#,
        );
        let style = roxmltree::Document::parse(
            r#"<p:style xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <a:lnRef idx="1"/>
                </p:style>"#,
        )
        .unwrap();

        let stroke = parse_style_stroke(Some(style.root_element()), &theme)
            .expect("fixed recipe color should resolve");
        assert_eq!(stroke.color, "123456");
        assert_eq!(stroke.width, 19050);
        assert_eq!(stroke.line_cap.as_deref(), Some("round"));
        assert_eq!(stroke.dash_style, None);
        assert_eq!(stroke.custom_dash.len(), 1);
        assert!((stroke.custom_dash[0].dash - 1.25).abs() < f64::EPSILON);
        assert!((stroke.custom_dash[0].space - 0.75).abs() < f64::EPSILON);
        assert_eq!(stroke.line_join.as_deref(), Some("miter"));
        assert_eq!(stroke.miter_limit, Some(8.0));
        assert_eq!(stroke.alignment.as_deref(), Some("in"));
    }

    #[test]
    fn slide_style_presence_replaces_placeholder_fill_and_line_including_no_style() {
        let theme = PptxTheme::default();
        let mut placeholders = LayoutPlaceholders::default();
        placeholders.by_idx_fill.insert(
            7,
            Fill::Solid {
                color: "AABBCC".to_owned(),
            },
        );
        placeholders.by_idx_stroke.insert(
            7,
            Stroke {
                color: "DDEEFF".to_owned(),
                width: 12700,
                fill: None,
                dash_style: None,
                custom_dash: Vec::new(),
                line_cap: None,
                line_join: None,
                miter_limit: None,
                alignment: None,
                head_end: None,
                tail_end: None,
                cmpd: None,
            },
        );
        let doc = roxmltree::Document::parse(
            r#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:nvSpPr><p:cNvPr id="1" name="Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="7"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100000" cy="100000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:style><a:lnRef idx="0"/><a:fillRef idx="0"/><a:effectRef idx="0"/><a:fontRef idx="minor"/></p:style></p:sp>"#,
        )
        .unwrap();
        let mut zip = empty_zip();
        let shape = parse_shape(
            doc.root_element(),
            &placeholders,
            &theme,
            &HashMap::new(),
            "ppt/slides",
            None,
            &mut zip,
        )
        .expect("shape");

        assert!(shape.stroke.is_none());
        assert!(matches!(shape.fill, Some(Fill::None)));
    }

    #[test]
    fn table_cell_projects_tcpr_vertical_text_and_margins_into_the_text_body() {
        let doc = roxmltree::Document::parse(
            r#"<a:tc xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Vertical label</a:t></a:r></a:p></a:txBody>
              <a:tcPr vert="vert270" anchor="ctr" marL="100" marT="200" marR="300" marB="400"/>
            </a:tc>"#,
        )
        .unwrap();
        let mut zip = empty_zip();
        let cell = parse_table_cell(
            doc.root_element(),
            &HashMap::new(),
            &HashMap::new(),
            "ppt/slides",
            &mut zip,
        );
        let body = cell.text_body.expect("table cell text body");
        assert_eq!(body.vert, "vert270");
        assert_eq!(body.vertical_anchor, "ctr");
        assert_eq!(
            (body.l_ins, body.t_ins, body.r_ins, body.b_ins),
            (100, 200, 300, 400)
        );
    }

    #[test]
    fn table_element_exposes_graphic_frame_id() {
        let doc = roxmltree::Document::parse(
            r#"<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:nvGraphicFramePr><p:cNvPr id="42" name="Table 1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
              <p:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></p:xfrm>
              <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                <a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="100"/></a:tblGrid>
                  <a:tr h="100"><a:tc><a:txBody/></a:tc></a:tr>
                </a:tbl>
              </a:graphicData></a:graphic>
            </p:graphicFrame>"#,
        )
        .unwrap();
        let mut zip = empty_zip();
        let mut out = Vec::new();
        parse_sp_tree_node(
            doc.root_element(),
            &LayoutPlaceholders::default(),
            "ppt/slides",
            &HashMap::new(),
            &HashMap::new(),
            &mut zip,
            &PptxTheme::default(),
            &mut out,
            false,
            None,
            DepthGuard::root(),
        );

        let SlideElement::Table(table) = out.pop().expect("table output") else {
            panic!("expected TableElement")
        };
        assert_eq!(table.id.as_deref(), Some("42"));
    }
}

#[cfg(test)]
mod picture_property_resolution_tests {
    use super::*;
    use crate::master::{InheritedBlipFill, LayoutPlaceholders};
    use std::io::{Cursor, Write};

    fn tiny_png() -> Vec<u8> {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(b"\x89PNG\r\n\x1a\n");
        bytes.extend_from_slice(&[0, 0, 0, 13]);
        bytes.extend_from_slice(b"IHDR");
        bytes.extend_from_slice(&2u32.to_be_bytes());
        bytes.extend_from_slice(&2u32.to_be_bytes());
        bytes.extend_from_slice(&[8, 6, 0, 0, 0]);
        bytes.extend_from_slice(&[0, 0, 0, 0]);
        bytes
    }

    fn image_zip() -> PptxZip {
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            writer
                .start_file(
                    "ppt/media/image1.png",
                    zip::write::SimpleFileOptions::default(),
                )
                .unwrap();
            writer.write_all(&tiny_png()).unwrap();
            writer.finish().unwrap();
        }
        PptxZip::new(Cursor::new(bytes)).unwrap()
    }

    fn theme_with_picture_style() -> crate::theme::PptxTheme {
        crate::theme::PptxTheme::from_xml(
            r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:themeElements>
                <a:clrScheme name="Test"><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1></a:clrScheme>
                <a:fmtScheme name="Test">
                  <a:fillStyleLst/>
                  <a:lnStyleLst><a:ln w="22222"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
                  <a:effectStyleLst><a:effectStyle>
                    <a:effectLst><a:outerShdw blurRad="100" dist="200" dir="0"><a:srgbClr val="112233"/></a:outerShdw></a:effectLst>
                    <a:scene3d><a:camera prst="isometricLeftUp"/><a:lightRig rig="threePt" dir="t"/></a:scene3d>
                    <a:sp3d prstMaterial="matte"/>
                  </a:effectStyle></a:effectStyleLst>
                  <a:bgFillStyleLst/>
                </a:fmtScheme>
              </a:themeElements>
            </a:theme>"#,
        )
    }

    fn run_tree_child(
        xml: &str,
        placeholders: &LayoutPlaceholders,
        theme: &(impl PptxThemeSource + ?Sized),
        zip: &mut PptxZip,
    ) -> PictureElement {
        let doc = roxmltree::Document::parse(xml).unwrap();
        let mut out = Vec::new();
        let rels = HashMap::from([("rIdImg".to_owned(), "../media/image1.png".to_owned())]);
        parse_sp_tree_node(
            doc.root_element(),
            placeholders,
            "ppt/slides",
            &rels,
            &HashMap::new(),
            zip,
            theme,
            &mut out,
            false,
            None,
            DepthGuard::root(),
        );
        let SlideElement::Picture(picture) = out.pop().expect("picture output") else {
            panic!("expected PictureElement")
        };
        picture
    }

    fn assert_theme_properties(picture: &PictureElement) {
        assert_eq!(
            picture.stroke.as_ref().map(|stroke| stroke.width),
            Some(22_222)
        );
        assert_eq!(picture.shadow.as_ref().map(|shadow| shadow.dist), Some(200));
        assert_eq!(
            picture
                .scene3d
                .as_ref()
                .map(|scene| scene.camera.prst.as_str()),
            Some("isometricLeftUp")
        );
        assert_eq!(
            picture
                .sp3d
                .as_ref()
                .map(|surface| surface.prst_material.as_str()),
            Some("matte")
        );
    }

    #[test]
    fn normal_picture_and_blip_filled_shape_share_style_component_resolution() {
        let theme = theme_with_picture_style();
        let mut zip = image_zip();
        let ordinary = run_tree_child(
            r#"<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <p:nvPicPr><p:cNvPr id="1" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
              <p:blipFill><a:blip r:embed="rIdImg"/></p:blipFill>
              <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm></p:spPr>
              <p:style><a:lnRef idx="1"><a:srgbClr val="FFFFFF"/></a:lnRef>
                <a:effectRef idx="1"><a:srgbClr val="FFFFFF"/></a:effectRef></p:style>
            </p:pic>"#,
            &LayoutPlaceholders::default(),
            &theme,
            &mut zip,
        );
        assert_eq!(ordinary.id.as_deref(), Some("1"));
        assert_theme_properties(&ordinary);

        let mut zip = image_zip();
        let blip_shape = run_tree_child(
            r#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                           xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <p:nvSpPr><p:cNvPr id="2" name="Blip shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
              <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm>
                <a:blipFill><a:blip r:embed="rIdImg"/></a:blipFill></p:spPr>
              <p:style><a:lnRef idx="1"><a:srgbClr val="FFFFFF"/></a:lnRef>
                <a:effectRef idx="1"><a:srgbClr val="FFFFFF"/></a:effectRef></p:style>
            </p:sp>"#,
            &LayoutPlaceholders::default(),
            &theme,
            &mut zip,
        );
        assert_eq!(blip_shape.id.as_deref(), Some("2"));
        assert_theme_properties(&blip_shape);
    }

    #[test]
    fn inherited_blip_shape_and_transformless_picture_retain_layout_effects() {
        let property_xml = roxmltree::Document::parse(
            r#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:spPr><a:ln w="33333"><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:ln>
                <a:effectLst><a:outerShdw blurRad="500" dist="700" dir="0"><a:srgbClr val="112233"/></a:outerShdw></a:effectLst>
                <a:scene3d><a:camera prst="perspectiveFront"/><a:lightRig rig="threePt" dir="t"/></a:scene3d>
                <a:sp3d prstMaterial="plastic"/>
              </p:spPr>
            </p:sp>"#,
        )
        .unwrap();
        let properties = resolve_picture_shape_properties(
            child(property_xml.root_element(), "spPr"),
            None,
            None,
            &HashMap::new(),
        );
        let transform = Transform {
            cx: 1000,
            cy: 1000,
            ..Default::default()
        };
        let inherited_blip = InheritedBlipFill {
            image_path: "ppt/media/image1.png".to_owned(),
            mime_type: "image/png".to_owned(),
            src_rect: None,
            alpha: None,
            duotone: None,
        };
        let mut placeholders = LayoutPlaceholders::default();
        placeholders.by_idx.insert(9, transform);
        placeholders.by_idx_blip_fill.insert(9, inherited_blip);
        placeholders.by_idx_picture_properties.insert(9, properties);

        let assert_inherited = |picture: &PictureElement| {
            assert_eq!(
                picture.stroke.as_ref().map(|stroke| stroke.width),
                Some(33_333)
            );
            assert_eq!(picture.shadow.as_ref().map(|shadow| shadow.dist), Some(700));
            assert_eq!(
                picture
                    .scene3d
                    .as_ref()
                    .map(|scene| scene.camera.prst.as_str()),
                Some("perspectiveFront")
            );
            assert_eq!(
                picture
                    .sp3d
                    .as_ref()
                    .map(|surface| surface.prst_material.as_str()),
                Some("plastic")
            );
        };

        let mut zip = image_zip();
        let inherited_shape = run_tree_child(
            r#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                           xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:nvSpPr><p:cNvPr id="3" name="Inherited"/><p:cNvSpPr/><p:nvPr><p:ph type="pic" idx="9"/></p:nvPr></p:nvSpPr>
              <p:spPr/>
            </p:sp>"#,
            &placeholders,
            &HashMap::new(),
            &mut zip,
        );
        assert_inherited(&inherited_shape);

        let mut zip = image_zip();
        let placeholder_picture = run_tree_child(
            r#"<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <p:nvPicPr><p:cNvPr id="4" name="Placeholder"/><p:cNvPicPr/><p:nvPr><p:ph type="pic" idx="9"/></p:nvPr></p:nvPicPr>
              <p:blipFill><a:blip r:embed="rIdImg"/></p:blipFill><p:spPr/>
            </p:pic>"#,
            &placeholders,
            &HashMap::new(),
            &mut zip,
        );
        assert_inherited(&placeholder_picture);
    }

    #[test]
    fn explicit_effect_ref_clears_inherited_3d_components() {
        let inherited_xml = roxmltree::Document::parse(
            r#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:spPr><a:scene3d><a:camera prst="perspectiveFront"/><a:lightRig rig="threePt" dir="t"/></a:scene3d>
                <a:sp3d prstMaterial="plastic"/></p:spPr>
            </p:sp>"#,
        )
        .unwrap();
        let inherited = resolve_picture_shape_properties(
            child(inherited_xml.root_element(), "spPr"),
            None,
            None,
            &HashMap::new(),
        );
        let slide_xml = roxmltree::Document::parse(
            r#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:spPr/><p:style><a:effectRef idx="0"/></p:style>
            </p:sp>"#,
        )
        .unwrap();
        let resolved = resolve_picture_shape_properties(
            child(slide_xml.root_element(), "spPr"),
            child(slide_xml.root_element(), "style"),
            Some(inherited),
            &HashMap::new(),
        );
        assert!(resolved.scene3d.is_none());
        assert!(resolved.sp3d.is_none());
    }
}

#[cfg(test)]
mod ole_tests {
    use super::*;
    use crate::master::LayoutPlaceholders;
    use std::io::{Cursor, Write};

    /// A 2×2 PNG so `png_size_from_bytes` returns Some((2, 2)) and
    /// `read_zip_bytes` succeeds. Only the 8-byte signature + IHDR are read.
    fn tiny_png() -> Vec<u8> {
        let mut b = Vec::new();
        b.extend_from_slice(b"\x89PNG\r\n\x1a\n"); // signature
        b.extend_from_slice(&[0, 0, 0, 13]); // IHDR length
        b.extend_from_slice(b"IHDR");
        b.extend_from_slice(&2u32.to_be_bytes()); // width
        b.extend_from_slice(&2u32.to_be_bytes()); // height
        b.extend_from_slice(&[8, 6, 0, 0, 0]); // bit depth / colour type / etc.
        b.extend_from_slice(&[0, 0, 0, 0]); // fake CRC (unread)
        b
    }

    fn zip_with(parts: &[(&str, &[u8])]) -> PptxZip {
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(Cursor::new(&mut buf));
            let o = zip::write::SimpleFileOptions::default();
            for (path, bytes) in parts {
                w.start_file(*path, o).unwrap();
                w.write_all(bytes).unwrap();
            }
            w.finish().unwrap();
        }
        PptxZip::new(Cursor::new(buf)).unwrap()
    }

    const OLE_URI: &str = "http://schemas.openxmlformats.org/presentationml/2006/ole";

    fn ns() -> &'static str {
        concat!(
            r#"xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" "#,
            r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" "#,
            r#"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006""#
        )
    }

    fn run(xml: &str, zip: &mut PptxZip, rels: &HashMap<String, String>) -> Vec<SlideElement> {
        let doc = roxmltree::Document::parse(xml).unwrap();
        let lph = LayoutPlaceholders::default();
        let theme = HashMap::new();
        let smart = HashMap::new();
        let mut out = Vec::new();
        parse_sp_tree_node(
            doc.root_element(),
            &lph,
            "ppt/slides",
            rels,
            &smart,
            zip,
            &theme,
            &mut out,
            false,
            None,
            DepthGuard::root(),
        );
        out
    }

    // ── MCE AlternateContent selection (ECMA-376 Part 3 §9.3) ───────────────
    //
    // A `<p:pic>` whose `<a:blip r:embed>` names `embed`, sized so
    // `parse_picture` yields a `SlideElement::Picture` with `image_path`
    // `ppt/media/<embed>.png`. Used to tell which MCE branch was selected.
    fn mce_pic(embed: &str) -> String {
        format!(
            r#"<p:pic>
                 <p:nvPicPr><p:cNvPr id="7" name="P"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
                 <p:blipFill><a:blip r:embed="{embed}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
                 <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                   <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
               </p:pic>"#
        )
    }

    /// `<mc:AlternateContent>` whose Choice carries `mce_pic("choice")` behind the
    /// given `requires` prefix, and whose Fallback carries `mce_pic("fallback")`.
    /// The `unk`/`a14`/`p14` prefixes are declared so `Requires` resolves. Returns
    /// the single emitted picture's `image_path` (or a marker string).
    fn mce_selected_pic_path(requires: &str, choice_inner: &str) -> Vec<SlideElement> {
        let mut rels = HashMap::new();
        rels.insert("rIdChoice".to_string(), "../media/choice.png".to_string());
        rels.insert(
            "rIdFallback".to_string(),
            "../media/fallback.png".to_string(),
        );
        let mut zip = zip_with(&[
            ("ppt/media/choice.png", &tiny_png()),
            ("ppt/media/fallback.png", &tiny_png()),
        ]);
        let xml = format!(
            r#"<mc:AlternateContent {ns}
                 xmlns:unk="http://example.com/an/extension/we/do/not/understand"
                 xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
                 xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"
                 xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
                 <mc:Choice Requires="{requires}">{choice_inner}</mc:Choice>
                 <mc:Fallback>{fallback}</mc:Fallback>
               </mc:AlternateContent>"#,
            ns = ns(),
            requires = requires,
            choice_inner = choice_inner,
            fallback = mce_pic("rIdFallback"),
        );
        run(&xml, &mut zip, &rels)
    }

    fn only_pic_path(out: &[SlideElement]) -> String {
        let pics: Vec<&String> = out
            .iter()
            .filter_map(|e| match e {
                SlideElement::Picture(p) => Some(&p.image_path),
                _ => None,
            })
            .collect();
        assert_eq!(pics.len(), 1, "expected exactly one Picture, got {out:?}");
        pics[0].clone()
    }

    /// ECMA-376 Part 3 §9.3(1) — a `<mc:Choice>` is selected only when EVERY
    /// namespace in its `Requires` is understood. `unk` is a namespace the pptx
    /// shape walker has no handler for; although the Choice's `<p:pic>` WOULD
    /// render, the consumer must NOT select an un-understood Choice — it selects
    /// the `<mc:Fallback>` instead. The old output-emptiness heuristic wrongly
    /// kept the Choice because its picture produced output; this pins §9.3.
    #[test]
    fn mce_unhandled_choice_selects_fallback_picture() {
        let out = mce_selected_pic_path("unk", &mce_pic("rIdChoice"));
        assert_eq!(
            only_pic_path(&out),
            "ppt/media/fallback.png",
            "un-understood Choice must yield the Fallback picture, not the Choice's"
        );
    }

    /// §9.3(1) — the drawing-2010 (`a14`) namespace IS understood by the pptx
    /// shape walker (Word/PowerPoint wrap ordinary `<p:sp>`/`<p:pic>` carrying an
    /// a14 extension in `<mc:Choice Requires="a14">`, e.g. sample-8's 115 such
    /// choices). The understood Choice must be selected and its picture kept, the
    /// Fallback dropped — guarding against a regression on those real slides.
    #[test]
    fn mce_understood_a14_choice_selected_over_fallback() {
        let out = mce_selected_pic_path("a14", &mce_pic("rIdChoice"));
        assert_eq!(
            only_pic_path(&out),
            "ppt/media/choice.png",
            "understood a14 Choice must be selected, Fallback dropped"
        );
    }

    /// §9.3 + the "understood = renderable" contract — `p14`
    /// (PowerPoint-2010, ink `<p:contentPart>`) is NOT rendered by the shape
    /// walker, so a `Requires="p14"` Choice must not be selected: the Fallback
    /// preview picture is drawn instead. (Matches sample-3's ink slides.)
    #[test]
    fn mce_ink_p14_contentpart_choice_falls_back_to_picture() {
        let out = mce_selected_pic_path("p14", r#"<p:contentPart r:id="rIdInk"/>"#);
        assert_eq!(
            only_pic_path(&out),
            "ppt/media/fallback.png",
            "un-rendered ink Choice must yield the Fallback picture"
        );
    }

    /// ECMA-376 §21.1.2.3.5 — a `<p:cNvPr><a:hlinkClick @r:id>` on a shape makes
    /// the whole shape an EXTERNAL click target; r:id resolves via slide rels.
    #[test]
    fn shape_cnv_hlink_click_external_resolves_url() {
        let mut rels = HashMap::new();
        rels.insert("rId9".to_string(), "https://example.com/".to_string());
        let mut zip = zip_with(&[]);
        let xml = format!(
            r#"<p:sp {ns}>
                <p:nvSpPr>
                  <p:cNvPr id="3" name="Rect"><a:hlinkClick r:id="rId9"/></p:cNvPr>
                  <p:cNvSpPr/><p:nvPr/>
                </p:nvSpPr>
                <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
               </p:sp>"#,
            ns = ns()
        );
        let out = run(&xml, &mut zip, &rels);
        let shape = match &out[0] {
            SlideElement::Shape(s) => s,
            _ => panic!("expected a shape element"),
        };
        assert_eq!(shape.hyperlink.as_deref(), Some("https://example.com/"));
        assert!(shape.hyperlink_action.is_none());
    }

    /// A `<a:hlinkClick action="ppaction://hlinksldjump" r:id>` on a shape is an
    /// INTERNAL slide jump: `hyperlink` is the resolved internal slide part and
    /// `hyperlink_action` carries the raw ppaction verb. ECMA-376 §21.1.2.3.5.
    #[test]
    fn shape_cnv_hlink_click_internal_slidejump() {
        let mut rels = HashMap::new();
        rels.insert("rId3".to_string(), "../slides/slide3.xml".to_string());
        let mut zip = zip_with(&[]);
        let xml = format!(
            r#"<p:sp {ns}>
                <p:nvSpPr>
                  <p:cNvPr id="4" name="Btn"><a:hlinkClick action="ppaction://hlinksldjump" r:id="rId3"/></p:cNvPr>
                  <p:cNvSpPr/><p:nvPr/>
                </p:nvSpPr>
                <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
               </p:sp>"#,
            ns = ns()
        );
        let out = run(&xml, &mut zip, &rels);
        let shape = match &out[0] {
            SlideElement::Shape(s) => s,
            _ => panic!("expected a shape element"),
        };
        assert_eq!(shape.hyperlink.as_deref(), Some("../slides/slide3.xml"));
        assert_eq!(
            shape.hyperlink_action.as_deref(),
            Some("ppaction://hlinksldjump")
        );
    }

    /// A shape with no cNvPr hlinkClick has both hyperlink fields None.
    #[test]
    fn shape_cnv_without_hlink_click_is_none() {
        let rels = HashMap::new();
        let mut zip = zip_with(&[]);
        let xml = format!(
            r#"<p:sp {ns}>
                <p:nvSpPr><p:cNvPr id="5" name="Plain"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
                <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
               </p:sp>"#,
            ns = ns()
        );
        let out = run(&xml, &mut zip, &rels);
        let shape = match &out[0] {
            SlideElement::Shape(s) => s,
            _ => panic!("expected a shape element"),
        };
        assert!(shape.hyperlink.is_none());
        assert!(shape.hyperlink_action.is_none());
    }

    /// PowerPoint's canonical OLE output: `mc:Choice Requires="v"` holds a
    /// spid-only `<p:oleObj>` (no `<p:pic>` — spid and pic are mutually exclusive,
    /// ECMA-376 Part 3 §B.1) and `mc:Fallback` holds the drawable pic-carrying
    /// oleObj. Because we do not understand the `v` (VML) namespace, MCE §9.3 says
    /// to render the Fallback. Selecting the first oleObj that actually has a
    /// `<p:pic>` must emit exactly ONE preview picture from the Fallback, even
    /// though the Choice's spid-only oleObj comes first in document order.
    #[test]
    fn ole_object_canonical_choice_spid_fallback_pic_emits_single_picture() {
        let mut rels = HashMap::new();
        rels.insert("rIdImg".to_string(), "../media/oleImage1.png".to_string());
        let mut zip = zip_with(&[("ppt/media/oleImage1.png", &tiny_png())]);

        // Choice: spid-only oleObj (VML shape reference, no drawable pic).
        let choice = r#"<p:oleObj spid="_x0000_s1026" r:id="rIdOle" progId="Excel.Sheet.12"><p:embed/></p:oleObj>"#;
        // Fallback: the pic-carrying oleObj we can actually render.
        let fallback = r#"<p:oleObj r:id="rIdOle" progId="Excel.Sheet.12"><p:embed/>
                 <p:pic>
                  <p:nvPicPr><p:cNvPr id="5" name="Object"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
                  <p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
                  <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                   <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
                 </p:pic>
                </p:oleObj>"#;
        let xml = format!(
            r#"<p:graphicFrame {ns}>
              <p:xfrm><a:off x="1000000" y="2000000"/><a:ext cx="1828800" cy="914400"/></p:xfrm>
              <a:graphic><a:graphicData uri="{uri}">
                <mc:AlternateContent>
                  <mc:Choice Requires="v">{choice}</mc:Choice>
                  <mc:Fallback>{fallback}</mc:Fallback>
                </mc:AlternateContent>
              </a:graphicData></a:graphic>
            </p:graphicFrame>"#,
            ns = ns(),
            uri = OLE_URI,
            choice = choice,
            fallback = fallback,
        );

        let out = run(&xml, &mut zip, &rels);
        let pics: Vec<_> = out
            .iter()
            .filter(|e| matches!(e, SlideElement::Picture(_)))
            .collect();
        assert_eq!(
            pics.len(),
            1,
            "canonical Choice(spid)/Fallback(pic) must yield exactly one preview picture, got {}",
            pics.len()
        );
        // Position must be the graphicFrame's (the inner pic xfrm is 0,0-relative).
        if let SlideElement::Picture(p) = pics[0] {
            assert_eq!((p.x, p.y), (1_000_000, 2_000_000));
            assert_eq!((p.width, p.height), (1_828_800, 914_400));
        }
    }

    /// A constructed case where BOTH Choice and Fallback carry a pic-bearing
    /// oleObj (not PowerPoint's real output, but a robustness guard): selecting
    /// the *first* pic-carrying oleObj must still emit exactly ONE picture, never
    /// a double-draw from the second.
    #[test]
    fn ole_object_both_branches_have_pic_emits_single_picture() {
        let mut rels = HashMap::new();
        rels.insert("rIdImg".to_string(), "../media/oleImage1.png".to_string());
        let mut zip = zip_with(&[("ppt/media/oleImage1.png", &tiny_png())]);

        let pic = |embed: &str| {
            format!(
                r#"<p:oleObj r:id="rIdOle" progId="Excel.Sheet.12"><p:embed/>
                 <p:pic>
                  <p:nvPicPr><p:cNvPr id="5" name="Object"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
                  <p:blipFill><a:blip r:embed="{embed}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
                  <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                   <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
                 </p:pic>
                </p:oleObj>"#
            )
        };
        let xml = format!(
            r#"<p:graphicFrame {ns}>
              <p:xfrm><a:off x="1000000" y="2000000"/><a:ext cx="1828800" cy="914400"/></p:xfrm>
              <a:graphic><a:graphicData uri="{uri}">
                <mc:AlternateContent>
                  <mc:Choice Requires="v">{choice}</mc:Choice>
                  <mc:Fallback>{fallback}</mc:Fallback>
                </mc:AlternateContent>
              </a:graphicData></a:graphic>
            </p:graphicFrame>"#,
            ns = ns(),
            uri = OLE_URI,
            choice = pic("rIdImg"),
            fallback = pic("rIdImg"),
        );

        let out = run(&xml, &mut zip, &rels);
        let pics: Vec<_> = out
            .iter()
            .filter(|e| matches!(e, SlideElement::Picture(_)))
            .collect();
        assert_eq!(
            pics.len(),
            1,
            "two pic-carrying oleObjs must still yield exactly one preview picture, got {}",
            pics.len()
        );
    }

    /// A direct (un-wrapped) `<p:oleObj>` whose inner `<p:pic>` has NO xfrm must
    /// still emit a picture, positioned/sized by the graphicFrame xfrm.
    #[test]
    fn ole_object_without_pic_xfrm_uses_graphicframe_xfrm() {
        let mut rels = HashMap::new();
        rels.insert("rIdImg".to_string(), "../media/oleImage2.png".to_string());
        let mut zip = zip_with(&[("ppt/media/oleImage2.png", &tiny_png())]);

        let xml = format!(
            r#"<p:graphicFrame {ns}>
              <p:xfrm><a:off x="500000" y="600000"/><a:ext cx="2743200" cy="1371600"/></p:xfrm>
              <a:graphic><a:graphicData uri="{uri}">
                <p:oleObj r:id="rIdOle" progId="Excel.Sheet.12"><p:embed/>
                 <p:pic>
                  <p:nvPicPr><p:cNvPr id="6" name="Object"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
                  <p:blipFill><a:blip r:embed="rIdImg"/></p:blipFill>
                  <p:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
                 </p:pic>
                </p:oleObj>
              </a:graphicData></a:graphic>
            </p:graphicFrame>"#,
            ns = ns(),
            uri = OLE_URI,
        );

        let out = run(&xml, &mut zip, &rels);
        let SlideElement::Picture(p) = out
            .iter()
            .find(|e| matches!(e, SlideElement::Picture(_)))
            .expect("a picture must be emitted from a pic without its own xfrm")
        else {
            unreachable!()
        };
        assert_eq!((p.x, p.y), (500_000, 600_000));
        assert_eq!((p.width, p.height), (2_743_200, 1_371_600));
    }

    /// A `<p:oleObj>` with NO `<p:pic>` (icon-only / link form) emits nothing —
    /// preserving the prior silent-skip behaviour rather than drawing a blank.
    #[test]
    fn ole_object_without_pic_emits_nothing() {
        let rels = HashMap::new();
        let mut zip = zip_with(&[]);
        let xml = format!(
            r#"<p:graphicFrame {ns}>
              <p:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></p:xfrm>
              <a:graphic><a:graphicData uri="{uri}">
                <p:oleObj r:id="rIdOle" progId="Equation.3" showAsIcon="1"><p:embed/></p:oleObj>
              </a:graphicData></a:graphic>
            </p:graphicFrame>"#,
            ns = ns(),
            uri = OLE_URI,
        );
        let out = run(&xml, &mut zip, &rels);
        assert!(
            out.is_empty(),
            "an oleObj with no preview pic must emit nothing, got {} element(s)",
            out.len()
        );
    }
}

/// §20.1.2.2.8 — `<p:cNvPr hidden="1">` marks a shape/pic/graphicFrame/cxnSp/
/// grpSp as not-to-be-rendered. The parser must skip such nodes (and a hidden
/// group's whole subtree) while leaving `hidden="0"` / absent shapes untouched.
#[cfg(test)]
mod hidden_tests {
    use super::*;
    use crate::master::LayoutPlaceholders;
    use std::collections::HashMap;

    const NS: &str = concat!(
        r#"xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" "#,
        r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
        r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#
    );

    fn run(xml: &str) -> Vec<SlideElement> {
        let doc = roxmltree::Document::parse(xml).unwrap();
        let lph = LayoutPlaceholders::default();
        let theme = HashMap::new();
        let smart = HashMap::new();
        // No zip access needed for plain shapes; build an empty in-memory zip.
        let mut zip = {
            use std::io::Cursor;
            let mut buf = Vec::new();
            {
                let w = zip::ZipWriter::new(Cursor::new(&mut buf));
                w.finish().unwrap();
            }
            PptxZip::new(Cursor::new(buf)).unwrap()
        };
        let rels = HashMap::new();
        let mut out = Vec::new();
        parse_sp_tree_node(
            doc.root_element(),
            &lph,
            "ppt/slides",
            &rels,
            &smart,
            &mut zip,
            &theme,
            &mut out,
            false,
            None,
            DepthGuard::root(),
        );
        out
    }

    fn shape_xml(name: &str, hidden_attr: &str) -> String {
        format!(
            r#"<p:sp {ns}>
              <p:nvSpPr><p:cNvPr id="2" name="{name}"{hidden}/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
              <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
              <p:txBody><a:bodyPr/><a:p><a:r><a:t>{name}</a:t></a:r></a:p></p:txBody>
            </p:sp>"#,
            ns = NS,
            name = name,
            hidden = hidden_attr,
        )
    }

    #[test]
    fn hidden_shape_is_not_emitted() {
        // hidden="1" and hidden="true" both suppress the shape.
        for attr in [r#" hidden="1""#, r#" hidden="true""#] {
            let out = run(&shape_xml("Hidden", attr));
            assert!(
                out.is_empty(),
                "hidden shape emitted (attr={attr}): {out:?}"
            );
        }
    }

    #[test]
    fn visible_shape_is_emitted_unchanged() {
        // Absent, "0", and "false" all keep the shape.
        for attr in ["", r#" hidden="0""#, r#" hidden="false""#] {
            let out = run(&shape_xml("Visible", attr));
            assert_eq!(out.len(), 1, "visible shape dropped (attr={attr})");
            assert!(matches!(out[0], SlideElement::Shape(_)));
        }
    }

    #[test]
    fn hidden_group_elides_whole_subtree() {
        // A hidden grpSp must suppress its visible children too.
        let xml = format!(
            r#"<p:grpSp {ns}>
              <p:nvGrpSpPr><p:cNvPr id="2" name="Grp" hidden="1"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
              <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/>
                <a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/></a:xfrm></p:grpSpPr>
              <p:sp><p:nvSpPr><p:cNvPr id="3" name="Child"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
                <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm>
                  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
                <p:txBody><a:bodyPr/><a:p><a:r><a:t>c</a:t></a:r></a:p></p:txBody></p:sp>
            </p:grpSp>"#,
            ns = NS,
        );
        assert!(run(&xml).is_empty(), "hidden group leaked a child shape");
    }

    #[test]
    fn hidden_child_inside_visible_group_is_skipped() {
        // A visible group with one hidden and one visible child emits only the
        // visible child — proves the check is per-node, not just top-level.
        let xml = format!(
            r#"<p:grpSp {ns}>
              <p:nvGrpSpPr><p:cNvPr id="2" name="Grp"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
              <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/>
                <a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/></a:xfrm></p:grpSpPr>
              <p:sp><p:nvSpPr><p:cNvPr id="3" name="Hidden" hidden="1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
                <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm>
                  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
                <p:txBody><a:bodyPr/><a:p><a:r><a:t>h</a:t></a:r></a:p></p:txBody></p:sp>
              <p:sp><p:nvSpPr><p:cNvPr id="4" name="Shown"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
                <p:spPr><a:xfrm><a:off x="457200" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm>
                  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
                <p:txBody><a:bodyPr/><a:p><a:r><a:t>s</a:t></a:r></a:p></p:txBody></p:sp>
            </p:grpSp>"#,
            ns = NS,
        );
        let out = run(&xml);
        assert_eq!(out.len(), 1, "expected only the visible child: {out:?}");
    }

    // ── Group-shape recursion depth guard (RB2) ────────────────────────────

    /// `<p:grpSp>` nested `levels` deep (levels ≥ 1), with one visible leaf shape
    /// at the very bottom. Each `<p:grpSp>` is a real `parse_sp_tree_node`
    /// recursion. The outermost `<p:grpSp>` carries the namespace decls and is
    /// returned as the root element (the `run` helper parses it as one node).
    fn nested_groups(levels: usize) -> String {
        assert!(levels >= 1);
        let leaf = concat!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="99" name="Leaf"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>"#,
            r#"<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm>"#,
            r#"<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>"#,
            r#"<p:txBody><a:bodyPr/><a:p><a:r><a:t>x</a:t></a:r></a:p></p:txBody></p:sp>"#,
        );
        // Inner groups carry no xmlns (inherited from the root grpSp).
        let open_inner = concat!(
            r#"<p:grpSp><p:nvGrpSpPr><p:cNvPr id="2" name="G"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"#,
            r#"<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/>"#,
            r#"<a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/></a:xfrm></p:grpSpPr>"#,
        );
        let mut inner = leaf.to_string();
        // levels-1 inner groups around the leaf …
        for _ in 0..levels - 1 {
            inner = format!("{open_inner}{inner}</p:grpSp>");
        }
        // … then the outermost group with the namespace declarations, as root.
        format!(
            r#"<p:grpSp xmlns:p="{P}" xmlns:a="{A}"><p:nvGrpSpPr><p:cNvPr id="1" name="Root"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/></a:xfrm></p:grpSpPr>{inner}</p:grpSp>"#,
            P = P_NS,
            A = A_NS,
        )
    }

    fn run_deep(levels: usize) -> Vec<SlideElement> {
        // roxmltree::Document::parse itself recurses on nesting, so parse on a
        // generous stack; this test targets OUR grpSp guard. (The roxmltree layer
        // is bounded separately by the raw-XML depth pre-check in
        // `ooxml_common::depth`.)
        std::thread::Builder::new()
            .stack_size(256 * 1024 * 1024)
            .spawn(move || run(&nested_groups(levels)))
            .unwrap()
            .join()
            .unwrap()
    }

    #[test]
    fn deeply_nested_groups_do_not_trap() {
        // ~2 000 nested groups is ~30× the depth limit; pre-guard this recurses
        // 2 000 frames and traps on the WASM stack. The guard caps the descent so
        // parsing RETURNS instead of aborting. A group deeper than the limit drops
        // its subtree, so the single leaf below it does NOT survive — the assertion
        // is simply that we get here (no trap) with a bounded result.
        //
        // NB: this runs on a 256 MB stack, so it cannot by itself catch a removed
        // DepthGuard (the big stack absorbs the recursion). The LOAD-BEARING
        // coverage for the roxmltree-layer trap is the default-stack neutralization
        // test `theme::tests::deeply_nested_theme_xml_is_rejected_not_trapped`
        // (and the docx/xlsx siblings), which would overflow if the PPTX guarded
        // pre-check were bypassed. This test's job is the truncation contract.
        let out = run_deep(2_000);
        assert!(out.len() <= 1, "guard should bound the emitted shapes");
    }

    #[test]
    fn shallow_nested_groups_keep_their_leaf() {
        // A modest nest (well under the limit) parses in full: the leaf survives.
        let out = run_deep(10);
        assert_eq!(out.len(), 1, "leaf under a shallow group nest must survive");
        assert!(matches!(out[0], SlideElement::Shape(_)));
    }

    // Namespaces for `nested_groups` (kept local to these depth tests).
    const P_NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
}

#[cfg(test)]
mod smartart_hidden_tests {
    // `parse_smartart_drawing` dispatches dsp:sp/cxnSp/grpSp manually (it never
    // routes through `parse_sp_tree_node`), so its own hidden check is a
    // separate code path from `hidden_tests` above and needs its own coverage.
    use super::*;
    use std::collections::HashMap;

    const NS: &str = concat!(
        r#"xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" "#,
        r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main""#
    );

    fn run(sp_tree_children: &str) -> Vec<SlideElement> {
        // parse_smartart_drawing parses the drawing XML itself (it is the
        // dsp:drawing1.xml part, not a subtree passed in already-parsed).
        let xml = format!(
            r#"<dsp:drawing {ns}><dsp:spTree>{children}</dsp:spTree></dsp:drawing>"#,
            ns = NS,
            children = sp_tree_children,
        );
        let theme = HashMap::new();
        let gf_xfrm = Transform::default();
        let mut out = Vec::new();
        let mut zip = {
            use std::io::Cursor;
            let mut buf = Vec::new();
            {
                let w = zip::ZipWriter::new(Cursor::new(&mut buf));
                w.finish().unwrap();
            }
            PptxZip::new(Cursor::new(buf)).unwrap()
        };
        parse_smartart_drawing(&xml, &gf_xfrm, &theme, &mut out, &mut zip);
        out
    }

    fn dsp_sp(name: &str, hidden_attr: &str) -> String {
        format!(
            r#"<dsp:sp>
              <dsp:nvSpPr><dsp:cNvPr id="2" name="{name}"{hidden}/><dsp:cNvSpPr/><dsp:nvPr/></dsp:nvSpPr>
              <dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></dsp:spPr>
              <dsp:txBody><a:bodyPr/><a:p><a:r><a:t>{name}</a:t></a:r></a:p></dsp:txBody>
            </dsp:sp>"#,
            name = name,
            hidden = hidden_attr,
        )
    }

    #[test]
    fn hidden_top_level_smartart_shape_is_not_emitted() {
        let out = run(&dsp_sp("Hidden", r#" hidden="1""#));
        assert!(out.is_empty(), "hidden top-level dsp:sp emitted: {out:?}");
    }

    #[test]
    fn visible_top_level_smartart_shape_is_emitted_unchanged() {
        let out = run(&dsp_sp("Visible", ""));
        assert_eq!(out.len(), 1);
        assert!(matches!(out[0], SlideElement::Shape(_)));
    }

    #[test]
    fn hidden_smartart_group_elides_whole_subtree() {
        let xml = format!(
            r#"<dsp:grpSp>
              <dsp:nvGrpSpPr><dsp:cNvPr id="2" name="Grp" hidden="1"/><dsp:cNvGrpSpPr/><dsp:nvPr/></dsp:nvGrpSpPr>
              <dsp:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/>
                <a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/></a:xfrm></dsp:grpSpPr>
              {child}
            </dsp:grpSp>"#,
            child = dsp_sp("Child", ""),
        );
        assert!(
            run(&xml).is_empty(),
            "hidden SmartArt group leaked a child shape"
        );
    }

    #[test]
    fn hidden_child_inside_visible_smartart_group_is_skipped() {
        let xml = format!(
            r#"<dsp:grpSp>
              <dsp:nvGrpSpPr><dsp:cNvPr id="2" name="Grp"/><dsp:cNvGrpSpPr/><dsp:nvPr/></dsp:nvGrpSpPr>
              <dsp:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/>
                <a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/></a:xfrm></dsp:grpSpPr>
              {hidden}
              {shown}
            </dsp:grpSp>"#,
            hidden = dsp_sp("Hidden", r#" hidden="1""#),
            shown = dsp_sp("Shown", ""),
        );
        let out = run(&xml);
        assert_eq!(out.len(), 1, "expected only the visible child: {out:?}");
    }
}

/// Strict-conformance (ISO/IEC 29500 Strict) fixtures. PresentationML and
/// DrawingML elements are matched by local name (already Strict-safe); the one
/// namespace-pinned site is `attr_r`, which reads `r:embed` / `r:id`. These
/// tests declare the p / a / r namespaces under the `http://purl.oclc.org/ooxml/`
/// Strict URIs and assert a shape's text and a picture's resolved image path come
/// out identical to the Transitional case — proving `attr_r` accepts the Strict
/// relationships namespace. (Real Office-authored Strict `.pptx` end-to-end
/// validation is the QA11 public-conformance-corpus track.)
#[cfg(test)]
mod strict_namespace_tests {
    use super::*;
    use crate::master::LayoutPlaceholders;
    use std::io::{Cursor, Write};

    const STRICT_NS: &str = concat!(
        r#"xmlns:p="http://purl.oclc.org/ooxml/presentationml/main" "#,
        r#"xmlns:a="http://purl.oclc.org/ooxml/drawingml/main" "#,
        r#"xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships""#
    );

    fn tiny_png() -> Vec<u8> {
        let mut b = Vec::new();
        b.extend_from_slice(b"\x89PNG\r\n\x1a\n");
        b.extend_from_slice(&[0, 0, 0, 13]);
        b.extend_from_slice(b"IHDR");
        b.extend_from_slice(&2u32.to_be_bytes());
        b.extend_from_slice(&2u32.to_be_bytes());
        b.extend_from_slice(&[8, 6, 0, 0, 0]);
        b.extend_from_slice(&[0, 0, 0, 0]);
        b
    }

    fn zip_with(parts: &[(&str, &[u8])]) -> PptxZip {
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(Cursor::new(&mut buf));
            let o = zip::write::SimpleFileOptions::default();
            for (path, bytes) in parts {
                w.start_file(*path, o).unwrap();
                w.write_all(bytes).unwrap();
            }
            w.finish().unwrap();
        }
        PptxZip::new(Cursor::new(buf)).unwrap()
    }

    /// Parse each child of the `<p:spTree>` root, mirroring the real slide
    /// driver (which iterates `cSld > spTree`'s element children).
    fn run(xml: &str, zip: &mut PptxZip, rels: &HashMap<String, String>) -> Vec<SlideElement> {
        run_with_smart(xml, zip, rels, &HashMap::new())
    }

    /// Like [`run`], but threads a `dm_rid → drawing_xml` SmartArt cache
    /// through, mirroring [`crate::build_smartart_drawings`]'s output — needed
    /// by tests that exercise the `<a:graphicData uri="…diagram">` fallback
    /// path.
    fn run_with_smart(
        xml: &str,
        zip: &mut PptxZip,
        rels: &HashMap<String, String>,
        smart: &HashMap<String, String>,
    ) -> Vec<SlideElement> {
        let doc = roxmltree::Document::parse(xml).unwrap();
        let lph = LayoutPlaceholders::default();
        let theme = HashMap::new();
        let mut out = Vec::new();
        for node in doc.root_element().children().filter(|n| n.is_element()) {
            parse_sp_tree_node(
                node,
                &lph,
                "ppt/slides",
                rels,
                smart,
                zip,
                &theme,
                &mut out,
                false,
                None,
                DepthGuard::root(),
            );
        }
        out
    }

    // A Strict `<p:sp>` text body yields the run text (local-name matching is
    // namespace-agnostic), and a Strict `<p:pic>` resolves its blip `r:embed`
    // through the relationships map into an `image_path` — the latter exercises
    // `attr_r` against the Strict relationships URI.
    #[test]
    fn strict_shape_text_and_picture_embed_resolve() {
        let mut rels = HashMap::new();
        rels.insert("rIdImg".to_string(), "../media/image1.png".to_string());
        let mut zip = zip_with(&[("ppt/media/image1.png", &tiny_png())]);

        let xml = format!(
            r#"<p:spTree {STRICT_NS}>
                <p:sp>
                  <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
                  <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="457200"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
                  <p:txBody><a:bodyPr/><a:p><a:r><a:t>Strict Slide</a:t></a:r></a:p></p:txBody>
                </p:sp>
                <p:pic>
                  <p:nvPicPr><p:cNvPr id="3" name="Image"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
                  <p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
                  <p:spPr><a:xfrm><a:off x="0" y="500000"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
                </p:pic>
              </p:spTree>"#
        );

        let out = run(&xml, &mut zip, &rels);

        let shape_text = out.iter().find_map(|e| match e {
            SlideElement::Shape(s) => s
                .text_body
                .as_ref()
                .and_then(|tb| tb.paragraphs.first())
                .and_then(|para| para.runs.first())
                .and_then(|r| match r {
                    TextRun::Text(t) => Some(t.text.clone()),
                    _ => None,
                }),
            _ => None,
        });
        assert_eq!(
            shape_text.as_deref(),
            Some("Strict Slide"),
            "Strict shape text must parse"
        );

        let image_path = out.iter().find_map(|e| match e {
            SlideElement::Picture(p) => Some(p.image_path.clone()),
            _ => None,
        });
        assert_eq!(
            image_path.as_deref(),
            Some("ppt/media/image1.png"),
            "Strict r:embed must resolve to the media part via attr_r"
        );
    }

    /// A `<p:graphicFrame>` whose `<a:graphicData uri="…">` is the *Strict*
    /// SmartArt diagram literal (`http://purl.oclc.org/ooxml/drawingml/
    /// diagram`, ECMA-376 Part 1 Strict `dml-diagram.xsd` targetNamespace) must
    /// still be recognized and routed through the cached `drawing1.xml`
    /// fallback — proving the `graphicData@uri` comparison itself (not just
    /// `attr_r`) accepts the Strict value.
    #[test]
    fn strict_graphicframe_diagram_uri_renders_smartart_fallback() {
        let rels = HashMap::new();
        let mut zip = zip_with(&[]);

        let dm_rid = "rIdData1";
        let mut smart = HashMap::new();
        smart.insert(
            dm_rid.to_string(),
            format!(
                r#"<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" {STRICT_NS}>
                  <p:spTree>
                    <p:sp>
                      <p:nvSpPr><p:cNvPr id="2" name="Node1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
                      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>
                        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
                      <p:txBody><a:bodyPr/><a:p><a:r><a:t>SmartArt Node</a:t></a:r></a:p></p:txBody>
                    </p:sp>
                  </p:spTree>
                </dsp:drawing>"#
            ),
        );

        let xml = format!(
            r#"<p:spTree {STRICT_NS}>
                <p:graphicFrame>
                  <p:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></p:xfrm>
                  <a:graphic><a:graphicData uri="http://purl.oclc.org/ooxml/drawingml/diagram">
                    <dgm:relIds xmlns:dgm="http://purl.oclc.org/ooxml/drawingml/diagram" r:dm="{dm_rid}" r:lo="rIdLo1" r:qs="rIdQs1" r:cs="rIdCs1"/>
                  </a:graphicData></a:graphic>
                </p:graphicFrame>
              </p:spTree>"#
        );

        let out = run_with_smart(&xml, &mut zip, &rels, &smart);
        let smartart_text = out.iter().find_map(|e| match e {
            SlideElement::Shape(s) => s
                .text_body
                .as_ref()
                .and_then(|tb| tb.paragraphs.first())
                .and_then(|para| para.runs.first())
                .and_then(|r| match r {
                    TextRun::Text(t) => Some(t.text.clone()),
                    _ => None,
                }),
            _ => None,
        });
        assert_eq!(
            smartart_text.as_deref(),
            Some("SmartArt Node"),
            "Strict graphicData@uri for a diagram must still route to the cached drawing fallback"
        );
    }

    /// A `<p:graphicFrame>` whose `<a:graphicData uri="…">` is the *Strict* OLE
    /// literal (`http://purl.oclc.org/ooxml/presentationml/ole`, confirmed by
    /// ECMA-376 Part 1 Annex L.7.2.5 Strict example markup) must still surface
    /// the embedded object's preview `<p:pic>` as a picture element.
    #[test]
    fn strict_graphicframe_ole_uri_renders_preview_picture() {
        let mut rels = HashMap::new();
        rels.insert("rIdImg".to_string(), "../media/oleImage1.png".to_string());
        let mut zip = zip_with(&[("ppt/media/oleImage1.png", &tiny_png())]);

        let xml = format!(
            r#"<p:spTree {STRICT_NS}>
                <p:graphicFrame>
                  <p:xfrm><a:off x="1000000" y="2000000"/><a:ext cx="1828800" cy="914400"/></p:xfrm>
                  <a:graphic><a:graphicData uri="http://purl.oclc.org/ooxml/presentationml/ole">
                    <p:oleObj r:id="rIdOle" progId="Excel.Sheet.12"><p:embed/>
                      <p:pic>
                        <p:nvPicPr><p:cNvPr id="5" name="Object"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
                        <p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
                        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
                      </p:pic>
                    </p:oleObj>
                  </a:graphicData></a:graphic>
                </p:graphicFrame>
              </p:spTree>"#
        );

        let out = run(&xml, &mut zip, &rels);
        let SlideElement::Picture(p) = out
            .iter()
            .find(|e| matches!(e, SlideElement::Picture(_)))
            .expect("a Strict OLE graphicData@uri must still emit the preview picture")
        else {
            unreachable!()
        };
        assert_eq!((p.x, p.y), (1_000_000, 2_000_000));
        assert_eq!((p.width, p.height), (1_828_800, 914_400));
    }
}
