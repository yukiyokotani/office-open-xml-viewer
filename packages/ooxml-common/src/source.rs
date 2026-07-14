//! Stable source references for text parsed from OOXML XML parts.
//!
//! Renderers are free to split and reorder text for layout, so a source reference
//! maps a UTF-16 range in the parsed text back to a UTF-16 range in one XML text
//! node. The node itself is addressed by a namespace-aware element path.

use roxmltree::Node;
use serde::{Deserialize, Serialize};

/// One namespace-aware step in an OOXML element path.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub struct SourcePathStep {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub namespace_uri: Option<String>,
    pub local_name: String,
    /// Zero-based ordinal among element siblings with the same expanded name.
    pub index: u32,
}

/// Maps a UTF-16 interval in parsed/rendered text to one XML text node.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub struct TextSourceRef {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub part_name: Option<String>,
    pub path: Vec<SourcePathStep>,
    /// UTF-16 interval in the containing parsed text run.
    pub text_start: u32,
    pub text_end: u32,
    /// UTF-16 interval in the source XML text node.
    pub source_start: u32,
    pub source_end: u32,
}

/// Build a reference for the full text of an XML element such as `w:t` or `a:t`.
pub fn text_source_ref(node: Node<'_, '_>, part_name: Option<&str>) -> TextSourceRef {
    let text_len = node.text().unwrap_or("").encode_utf16().count() as u32;
    TextSourceRef {
        part_name: part_name.map(str::to_owned),
        path: element_path(node),
        text_start: 0,
        text_end: text_len,
        source_start: 0,
        source_end: text_len,
    }
}

/// Return the namespace-aware path from the XML root element to `node`.
pub fn element_path(node: Node<'_, '_>) -> Vec<SourcePathStep> {
    node.ancestors()
        .filter(|ancestor| ancestor.is_element())
        .collect::<Vec<_>>()
        .into_iter()
        .rev()
        .map(|element| {
            let tag = element.tag_name();
            let index = element
                .prev_siblings()
                .filter(|sibling| {
                    sibling.is_element()
                        && sibling.tag_name().name() == tag.name()
                        && sibling.tag_name().namespace() == tag.namespace()
                })
                .count() as u32;
            SourcePathStep {
                namespace_uri: tag.namespace().map(str::to_owned),
                local_name: tag.name().to_owned(),
                index,
            }
        })
        .collect()
}

/// Resolve a path produced by [`element_path`] against an OOXML part root.
pub fn resolve_element_path<'a, 'input>(
    root: Node<'a, 'input>,
    path: &[SourcePathStep],
) -> Option<Node<'a, 'input>> {
    let first = path.first()?;
    if !matches_step(root, first) || first.index != 0 {
        return None;
    }

    let mut current = root;
    for step in &path[1..] {
        current = current
            .children()
            .filter(|child| child.is_element() && matches_step(*child, step))
            .nth(step.index as usize)?;
    }
    Some(current)
}

fn matches_step(node: Node<'_, '_>, step: &SourcePathStep) -> bool {
    node.tag_name().name() == step.local_name
        && node.tag_name().namespace() == step.namespace_uri.as_deref()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn source_ref_uses_expanded_name_sibling_ordinals_and_utf16_offsets() {
        let doc = roxmltree::Document::parse(
            r#"<w:document xmlns:w="urn:w"><w:body><w:p/><w:p><w:r><w:t>A😀B</w:t></w:r></w:p></w:body></w:document>"#,
        )
        .unwrap();
        let text = doc
            .descendants()
            .find(|node| node.is_element() && node.tag_name().name() == "t")
            .unwrap();

        let source = text_source_ref(text, Some("word/document.xml"));

        assert_eq!(source.source_end, 4);
        assert_eq!(source.text_end, 4);
        assert_eq!(source.part_name.as_deref(), Some("word/document.xml"));
        assert_eq!(
            source
                .path
                .iter()
                .map(|step| (step.local_name.as_str(), step.index))
                .collect::<Vec<_>>(),
            vec![("document", 0), ("body", 0), ("p", 1), ("r", 0), ("t", 0)]
        );
        assert_eq!(
            resolve_element_path(doc.root_element(), &source.path),
            Some(text)
        );
    }
}
