//! OPC `[Content_Types].xml` parsing shared by OOXML format parsers.
//!
//! This module resolves a normalized package part name to either its exact
//! `<Override PartName>` content type or the `<Default Extension>` fallback.
//! It deliberately knows nothing about DOCX/PPTX/XLSX host schemas; callers
//! decide which MIME types are valid for the relationship they are resolving.

use std::collections::HashMap;

const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PackageContentTypes {
    defaults: HashMap<String, String>,
    overrides: HashMap<String, String>,
}

impl PackageContentTypes {
    pub fn parse(xml: &str) -> Option<Self> {
        let nodes_limit =
            u32::try_from(crate::resource::HARD_MAX_XML_DOM_COMPLEXITY).unwrap_or(u32::MAX);
        Self::parse_with_node_limit(xml, nodes_limit)
    }

    /// Parse with a format-calibrated DOM node ceiling. Format parsers that
    /// expose test overrides or tighter policy should pass their existing
    /// ceiling here rather than weakening their XML preflight when sharing this
    /// OPC projection.
    pub fn parse_with_node_limit(xml: &str, nodes_limit: u32) -> Option<Self> {
        let doc = crate::depth::parse_guarded_with_node_limit(xml, nodes_limit).ok()?;
        let root = doc.root_element();
        if root.tag_name().name() != "Types"
            || root.tag_name().namespace() != Some(CONTENT_TYPES_NS)
        {
            return None;
        }
        let mut result = Self::default();
        for node in root.children().filter(|node| {
            node.is_element() && node.tag_name().namespace() == Some(CONTENT_TYPES_NS)
        }) {
            match node.tag_name().name() {
                "Default" => {
                    let (Some(extension), Some(content_type)) =
                        (node.attribute("Extension"), node.attribute("ContentType"))
                    else {
                        continue;
                    };
                    result
                        .defaults
                        .insert(extension.to_ascii_lowercase(), content_type.to_owned());
                }
                "Override" => {
                    let (Some(part_name), Some(content_type)) =
                        (node.attribute("PartName"), node.attribute("ContentType"))
                    else {
                        continue;
                    };
                    result.overrides.insert(
                        part_name.trim_start_matches('/').to_owned(),
                        content_type.to_owned(),
                    );
                }
                _ => {}
            }
        }
        Some(result)
    }

    pub fn for_part(&self, part_path: &str) -> Option<&str> {
        let part_path = part_path.trim_start_matches('/');
        self.overrides
            .get(part_path)
            .or_else(|| {
                let extension = part_path.rsplit_once('.')?.1.to_ascii_lowercase();
                self.defaults.get(&extension)
            })
            .map(String::as_str)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn override_wins_and_default_extensions_are_ascii_case_insensitive() {
        let parsed = PackageContentTypes::parse(
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="FNTDATA" ContentType="application/x-font-ttf"/>
              <Override PartName="/ppt/fonts/font2.fntdata" ContentType="application/x-fontdata"/>
            </Types>"#,
        )
        .unwrap();
        assert_eq!(
            parsed.for_part("ppt/fonts/font1.fntdata"),
            Some("application/x-font-ttf")
        );
        assert_eq!(
            parsed.for_part("ppt/fonts/font2.fntdata"),
            Some("application/x-fontdata")
        );
    }

    #[test]
    fn malformed_xml_is_not_projected() {
        assert!(PackageContentTypes::parse("<Types>").is_none());
    }

    #[test]
    fn foreign_or_missing_content_types_namespaces_are_not_projected() {
        for xml in [
            r#"<Types><Default Extension="fntdata" ContentType="application/x-font-ttf"/></Types>"#,
            r#"<Types xmlns="urn:not-opc"><Default Extension="fntdata" ContentType="application/x-font-ttf"/></Types>"#,
            r#"<Other xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#,
        ] {
            assert!(PackageContentTypes::parse(xml).is_none());
        }
    }

    #[test]
    fn foreign_namespace_children_are_ignored() {
        let parsed = PackageContentTypes::parse(
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"
                xmlns:foreign="urn:not-opc">
              <foreign:Default Extension="fntdata" ContentType="application/x-font-ttf"/>
            </Types>"#,
        )
        .unwrap();
        assert_eq!(parsed.for_part("ppt/fonts/font1.fntdata"), None);
    }

    #[test]
    fn explicit_dom_node_ceiling_is_enforced() {
        let xml = format!(
            r#"<Types xmlns="{CONTENT_TYPES_NS}">{}</Types>"#,
            r#"<Default Extension="a" ContentType="application/a"/>"#.repeat(32),
        );
        assert!(PackageContentTypes::parse_with_node_limit(&xml, 8).is_none());
    }
}
