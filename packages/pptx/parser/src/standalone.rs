//! Parse one standalone PresentationML shape part into the presentation model.
//!
//! A shape part is an OPC part whose root element is a `p:sp` (ECMA-376
//! 19.3.1.43) or `p:cxnSp` (19.3.1.19) outside any slide. Theme references
//! resolve against the supplied theme part (20.1.6.9) and color map
//! (19.3.1.6). There is no layout or master: placeholder inheritance is not
//! resolved, so the result reports whether the shape is a placeholder and the
//! caller decides whether a partially inherited shape is usable.
use super::*;
use crate::theme::{apply_clr_map, parse_clr_map_node, PptxTheme};
use ooxml_common::ns::is_a_ns;
use ooxml_common::theme::StyleMatrixLookup;

/// A parsed standalone shape and the facts a caller needs to judge whether it
/// is complete without a slide context.
pub struct StandaloneShape {
    pub element: ShapeElement,
    /// The shape carries `p:nvPr/p:ph`; layout/master inheritance is absent.
    pub placeholder: bool,
    /// The shape or a selected theme style references a relationship (`r:`
    /// namespace). Theme relationships belong to the theme part, which this
    /// standalone API does not receive; callers must verify them separately.
    pub relationship_references: bool,
}

/// Parse the shape part `part` of the OPC `package`. Returns `Ok(None)` when
/// the root element is not a supported shape. `max_part_bytes` bounds each
/// inflated archive entry and `max_total_bytes` the archive as a whole.
pub fn parse_standalone_shape_part(
    package: &[u8],
    part: &str,
    theme_xml: &str,
    clr_map_xml: Option<&str>,
    max_part_bytes: u64,
    max_total_bytes: u64,
) -> Result<Option<StandaloneShape>, String> {
    let mut zip = open_zip_with_limits(
        package.to_vec(),
        Some(max_part_bytes),
        Some(max_total_bytes),
    )?;
    zip.run_operation("standalone-shape", |zip| {
        let xml = read_zip_str(zip, part).map_err(|e| e.to_string())?;
        let doc = parse_preflighted_pptx_xml(&xml).map_err(|e| e.to_string())?;
        let root = doc.root_element();
        let mut theme = PptxTheme::from_xml(theme_xml);
        if let Some(map_xml) = clr_map_xml {
            let map_doc = parse_preflighted_pptx_xml(map_xml).map_err(|e| e.to_string())?;
            let node = map_doc.root_element();
            // Slide masters carry p:clrMap (ECMA-376 §19.3.1.6), while
            // [MS-PPT] §2.11.9 round-trip color mappings use a:clrMap.
            if node.tag_name().name() != "clrMap"
                || !(is_p_ns(node.tag_name().namespace()) || is_a_ns(node.tag_name().namespace()))
            {
                return Err("color map part has no clrMap root".to_owned());
            }
            apply_clr_map(&mut theme, Some(&parse_clr_map_node(node)));
        }
        let rels_xml = read_zip_str(zip, &relationship_part_path(part)).unwrap_or_default();
        let base = part.rsplit_once('/').map_or("", |(dir, _)| dir);
        let rels = parse_rels(&rels_xml);
        // ECMA-376 §20.1.4.2: style-matrix references select DrawingML
        // fragments in the theme. A blip there owns a relationship in the
        // theme part, even when the shape XML contains no r: attribute. This
        // API has no theme part path or its relationships, so report that
        // dependency rather than presenting a fallback fill as complete.
        let relationship_references = root
            .descendants()
            .any(|n| n.attributes().any(|a| is_r_ns(a.namespace())))
            || theme_style_has_relationship(root, &theme);
        let placeholder = is_placeholder(root);
        let element = match (root.tag_name().name(), is_p_ns(root.tag_name().namespace())) {
            ("sp", true) => parse_shape(
                root,
                &LayoutPlaceholders::default(),
                &theme,
                &rels,
                base,
                None,
                zip,
            ),
            ("cxnSp", true) => parse_connector(root, &theme, &rels),
            _ => None,
        };
        Ok(element.map(|element| StandaloneShape {
            element,
            placeholder,
            relationship_references,
        }))
    })
}

fn theme_style_has_relationship(root: roxmltree::Node<'_, '_>, theme: &PptxTheme) -> bool {
    // A shape may repeat the same style reference in extension markup. Parse
    // each selected entry at most once so work scales with the shape XML plus
    // the selected theme entries, rather than their product.
    let mut inspected = std::collections::HashMap::new();
    root.descendants().any(|node| {
        if !node.is_element() {
            return false;
        }
        let Some(index) = node.attribute("idx").and_then(|value| value.parse().ok()) else {
            return false;
        };
        let (kind, selected) = match node.tag_name().name() {
            "fillRef" => (0, theme.format_scheme.lookup_fill_ref(index)),
            "lnRef" => (1, theme.format_scheme.lookup_line_ref(index)),
            "effectRef" => (2, theme.format_scheme.lookup_effect_ref(index)),
            _ => return false,
        };
        if let Some(&has_relationship) = inspected.get(&(kind, index)) {
            return has_relationship;
        }
        let StyleMatrixLookup::Entry(entry) = selected else {
            return false;
        };
        let has_relationship = roxmltree::Document::parse(&entry.to_xml()).is_ok_and(|doc| {
            doc.descendants()
                .any(|n| n.attributes().any(|a| is_r_ns(a.namespace())))
        });
        inspected.insert((kind, index), has_relationship);
        has_relationship
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{Cursor, Write};
    use zip::write::SimpleFileOptions;

    const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const STRICT_P: &str = "http://purl.oclc.org/ooxml/presentationml/main";
    const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const STRICT_A: &str = "http://purl.oclc.org/ooxml/drawingml/main";
    const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    fn package(xml: &str) -> Vec<u8> {
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        zip.start_file("shape.xml", SimpleFileOptions::default())
            .unwrap();
        zip.write_all(xml.as_bytes()).unwrap();
        zip.finish().unwrap().into_inner()
    }

    fn parse(xml: &str, theme: &str, map: Option<&str>) -> Result<Option<StandaloneShape>, String> {
        parse_standalone_shape_part(&package(xml), "shape.xml", theme, map, 100_000, 1_000_000)
    }

    fn shape(namespace: Option<&str>, kind: &str, extra: &str) -> String {
        let namespace = namespace.map_or(String::new(), |ns| format!(" xmlns:p=\"{ns}\""));
        let prefix = if namespace.is_empty() { "" } else { "p:" };
        format!(
            "<{prefix}{kind}{namespace} xmlns:a=\"{A}\" xmlns:r=\"{R}\"><{prefix}spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"100\" cy=\"100\"/></a:xfrm></{prefix}spPr>{extra}</{prefix}{kind}>"
        )
    }

    #[test]
    fn only_presentationml_shape_roots_are_supported() {
        for namespace in [None, Some("urn:foreign")] {
            assert!(parse(&shape(namespace, "sp", ""), "", None)
                .unwrap()
                .is_none());
        }
        for namespace in [P, STRICT_P] {
            assert!(parse(&shape(Some(namespace), "sp", ""), "", None)
                .unwrap()
                .is_some());
            assert!(parse(&shape(Some(namespace), "cxnSp", ""), "", None)
                .unwrap()
                .is_some());
            assert!(parse(&shape(Some(namespace), "pic", ""), "", None)
                .unwrap()
                .is_none());
        }
    }

    #[test]
    fn color_map_requires_presentationml_or_drawingml_namespace() {
        let xml = shape(Some(P), "sp", "");
        for map in ["<clrMap/>", "<x:clrMap xmlns:x=\"urn:foreign\"/>"] {
            assert!(parse(&xml, "", Some(map)).is_err());
        }
        for namespace in [P, STRICT_P] {
            let map = format!("<p:clrMap xmlns:p=\"{namespace}\"/>");
            assert!(parse(&xml, "", Some(&map)).unwrap().is_some());
        }
        for namespace in [A, STRICT_A] {
            let map = format!("<a:clrMap xmlns:a=\"{namespace}\"/>");
            assert!(parse(&xml, "", Some(&map)).unwrap().is_some());
        }
    }

    #[test]
    fn selected_theme_image_relationship_is_reported() {
        let theme = format!("<a:theme xmlns:a=\"{A}\" xmlns:r=\"{R}\"><a:themeElements><a:fmtScheme name=\"x\"><a:fillStyleLst><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill><a:blipFill><a:blip r:embed=\"rIdImage\"/></a:blipFill></a:fillStyleLst></a:fmtScheme></a:themeElements></a:theme>");
        for (index, expected) in [(1, false), (2, true)] {
            let xml = shape(
                Some(P),
                "sp",
                &format!("<p:style><a:fillRef idx=\"{index}\"/></p:style>"),
            );
            let parsed = parse(&xml, &theme, None).unwrap().unwrap();
            assert_eq!(parsed.relationship_references, expected);
        }
    }
}
