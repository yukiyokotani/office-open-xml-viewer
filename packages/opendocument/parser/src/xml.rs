//! XML access shared by the readers. Parts are parsed with roxmltree (no DTD,
//! no entity expansion; a bounded node count). Elements and attributes are
//! matched by namespace URI, never by prefix.

pub const NS_MANIFEST: &str = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
pub const NS_OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
pub const NS_STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
pub const NS_TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
pub const NS_FO: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
pub const NS_SVG: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
pub const NS_TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
pub const NS_DRAW: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
pub const NS_XLINK: &str = "http://www.w3.org/1999/xlink";
pub const NS_XML: &str = "http://www.w3.org/XML/1998/namespace";
pub const NS_LOEXT: &str = "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";
pub const NS_OFFICEOOO: &str = "http://openoffice.org/2009/office";
pub const NS_CONFIG: &str = "urn:oasis:names:tc:opendocument:xmlns:config:1.0";

/// Implementation bound on the nodes of one parsed part; the part's inflated
/// size is already bounded by the package governors.
pub const MAX_NODES_PER_PART: u32 = 8_000_000;

pub fn parse(text: &str) -> Result<roxmltree::Document<'_>, String> {
    roxmltree::Document::parse_with_options(
        text,
        roxmltree::ParsingOptions {
            allow_dtd: false,
            nodes_limit: MAX_NODES_PER_PART,
        },
    )
    .map_err(|error| match error {
        roxmltree::Error::NodesLimitReached => {
            "OpenDocument part exceeds the XML node limit".to_string()
        }
        _ => crate::error::invalid("a part is not well-formed XML"),
    })
}

pub fn is(node: roxmltree::Node<'_, '_>, namespace: &str, name: &str) -> bool {
    node.is_element()
        && node.tag_name().namespace() == Some(namespace)
        && node.tag_name().name() == name
}
