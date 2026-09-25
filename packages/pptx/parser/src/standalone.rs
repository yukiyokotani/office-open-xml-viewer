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

/// A parsed standalone shape and the facts a caller needs to judge whether it
/// is complete without a slide context.
pub struct StandaloneShape {
    pub element: ShapeElement,
    /// The shape carries `p:nvPr/p:ph`; layout/master inheritance is absent.
    pub placeholder: bool,
    /// Some attribute references a relationship (`r:` namespace), which the
    /// part's own relationships may or may not resolve.
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
            if node.tag_name().name() != "clrMap" {
                return Err("color map part has no clrMap root".to_owned());
            }
            apply_clr_map(&mut theme, Some(&parse_clr_map_node(node)));
        }
        let rels_xml = read_zip_str(zip, &relationship_part_path(part)).unwrap_or_default();
        let base = part.rsplit_once('/').map_or("", |(dir, _)| dir);
        let rels = parse_rels(&rels_xml);
        let relationship_references = root
            .descendants()
            .any(|n| n.attributes().any(|a| is_r_ns(a.namespace())));
        let placeholder = is_placeholder(root);
        let element = match root.tag_name().name() {
            "sp" => parse_shape(
                root,
                &LayoutPlaceholders::default(),
                &theme,
                &rels,
                base,
                None,
                zip,
            ),
            "cxnSp" => parse_connector(root, &theme, &rels),
            _ => None,
        };
        Ok(element.map(|element| StandaloneShape {
            element,
            placeholder,
            relationship_references,
        }))
    })
}
