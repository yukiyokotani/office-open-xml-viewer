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

/// Implementation bound on the attributes (namespace declarations included)
/// of one element.
pub const MAX_ATTRIBUTES_PER_ELEMENT: usize = 256;
/// Implementation bound on the namespace declarations in scope at an element.
pub const MAX_NAMESPACES_IN_SCOPE: usize = 256;
/// Implementation bound on element nesting.
pub const MAX_XML_DEPTH: usize = 1024;
/// Implementation bound on one part's parsing work, in units of the DOM
/// parser's own costs: every attribute is checked against the element's
/// other attributes (a² per element), every element and attribute name is
/// resolved against the namespaces in scope, and character data merged
/// across adjacent CDATA sections is copied again at each section.
pub const MAX_XML_WORK: usize = 512 * 1024 * 1024;
/// Implementation bound on the DOM parser's up-front reservation: it reserves
/// one node per `<` byte and one attribute per `=` byte of the whole part.
pub const MAX_XML_RESERVATION_BYTES: usize = 256 * 1024 * 1024;
/// Upper estimate of the parser's per-node and per-attribute records.
const XML_RECORD_BYTES: usize = 64;

const XML_LIMIT: &str = "OpenDocument part exceeds the XML parsing limits";

/// Bound the DOM parser's work before it runs: one linear pass over the part
/// counting each start tag's attributes, namespace declarations and depth,
/// and the text merged across CDATA sections. Malformed markup is left to the
/// parser, which rejects it.
fn preflight(text: &str) -> Result<(), String> {
    let bytes = text.as_bytes();
    let records = bytes.iter().filter(|&&b| b == b'<' || b == b'=').count();
    if records.saturating_mul(XML_RECORD_BYTES) > MAX_XML_RESERVATION_BYTES {
        return Err(XML_LIMIT.to_string());
    }
    let mut work = 0usize;
    let mut charge = |units: usize| -> Result<(), String> {
        work = work.saturating_add(units);
        if work > MAX_XML_WORK {
            Err(XML_LIMIT.to_string())
        } else {
            Ok(())
        }
    };
    // Namespace declarations per open element, and their running total.
    let mut scopes: Vec<usize> = Vec::new();
    let mut in_scope = 0usize;
    // Where the current run of character data (text and CDATA) began.
    let mut text_start = 0usize;
    let mut i = 0usize;
    let find = |from: usize, needle: &[u8]| -> usize {
        bytes[from..]
            .windows(needle.len())
            .position(|w| w == needle)
            .map_or(bytes.len(), |at| from + at + needle.len())
    };
    while let Some(offset) = bytes[i..].iter().position(|&b| b == b'<') {
        i += offset + 1;
        match bytes.get(i) {
            Some(b'?') => {
                i = find(i, b"?>");
                text_start = i;
            }
            Some(b'!') if bytes[i..].starts_with(b"!--") => {
                i = find(i, b"-->");
                text_start = i;
            }
            Some(b'!') if bytes[i..].starts_with(b"![CDATA[") => {
                // The section's data is appended to the text merged so far.
                charge(i - text_start)?;
                i = find(i, b"]]>");
            }
            // DOCTYPE and other declarations: the parser rejects them.
            Some(b'!') => return Ok(()),
            Some(b'/') => {
                in_scope -= scopes.pop().unwrap_or(0);
                i = find(i, b">");
                text_start = i;
            }
            Some(_) => {
                let (mut attributes, mut declarations) = (0usize, 0usize);
                let mut quote: Option<u8> = None;
                // The name token being read, and the last completed one
                // (XML allows white space between a name and `=`).
                let mut token: Option<usize> = Some(i);
                let mut last: (usize, usize) = (i, i);
                let mut self_closing = false;
                while let Some(&b) = bytes.get(i) {
                    match quote {
                        Some(q) if b == q => quote = None,
                        Some(_) => {}
                        None => match b {
                            b'"' | b'\'' => quote = Some(b),
                            b'=' => {
                                if let Some(start) = token.take() {
                                    last = (start, i);
                                }
                                attributes += 1;
                                let name = &bytes[last.0..last.1];
                                if name == b"xmlns" || name.starts_with(b"xmlns:") {
                                    declarations += 1;
                                }
                                last = (i, i);
                            }
                            b'>' => break,
                            b'/' => self_closing = true,
                            b if b.is_ascii_whitespace() => {
                                if let Some(start) = token.take() {
                                    last = (start, i);
                                }
                            }
                            _ => {
                                self_closing = false;
                                if token.is_none() {
                                    token = Some(i);
                                }
                            }
                        },
                    }
                    i += 1;
                }
                i += 1;
                text_start = i;
                in_scope += declarations;
                if attributes > MAX_ATTRIBUTES_PER_ELEMENT
                    || in_scope > MAX_NAMESPACES_IN_SCOPE
                    || scopes.len() >= MAX_XML_DEPTH
                {
                    return Err(XML_LIMIT.to_string());
                }
                charge(attributes * attributes + (attributes + 1) * (in_scope + 1))?;
                if self_closing {
                    in_scope -= declarations;
                } else {
                    scopes.push(declarations);
                }
            }
            None => break,
        }
    }
    Ok(())
}

pub fn parse(text: &str) -> Result<roxmltree::Document<'_>, String> {
    preflight(text)?;
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn preflight_bounds_attributes_namespaces_and_depth() {
        let attrs = |n: usize| {
            format!(
                "<r {}/>",
                (0..n)
                    .map(|i| format!(r#"a{i}="1""#))
                    .collect::<Vec<_>>()
                    .join(" ")
            )
        };
        assert!(parse(&attrs(MAX_ATTRIBUTES_PER_ELEMENT)).is_ok());
        assert_eq!(
            parse(&attrs(MAX_ATTRIBUTES_PER_ELEMENT + 1)).unwrap_err(),
            XML_LIMIT
        );
        // Quoted '=', '>' and '/' are attribute data, not markup.
        assert!(preflight(r#"<r a="x=y>/"><!-- <a b=1> --><![CDATA[<x y=1>]]></r>"#).is_ok());
        let namespaces = |n: usize| {
            let open: String = (0..n)
                .map(|i| format!(r#"<e xmlns:p{i}="urn:{i}">"#))
                .collect();
            format!("{open}{}", "</e>".repeat(n))
        };
        assert!(parse(&namespaces(MAX_NAMESPACES_IN_SCOPE)).is_ok());
        assert_eq!(
            parse(&namespaces(MAX_NAMESPACES_IN_SCOPE + 1)).unwrap_err(),
            XML_LIMIT
        );
        // White space around `=` does not hide a declaration.
        for spacing in [" = ", "\t=\n", "\r\n=  "] {
            let spaced = namespaces(MAX_NAMESPACES_IN_SCOPE + 1)
                .replace("=\"urn", &format!("{spacing}\"urn"));
            assert_eq!(parse(&spaced).unwrap_err(), XML_LIMIT, "{spacing:?}");
        }
        // The parser's reservation counts every `<` and `=` byte, text included.
        let equals = format!(
            "<r>{}</r>",
            "=".repeat(MAX_XML_RESERVATION_BYTES / XML_RECORD_BYTES + 1)
        );
        assert_eq!(parse(&equals).unwrap_err(), XML_LIMIT);
        // Text merged across adjacent CDATA sections is charged per section.
        let cdata = format!("<r>{}</r>", "<![CDATA[xxxxxxxx]]>".repeat(40_000));
        assert_eq!(preflight(&cdata).unwrap_err(), XML_LIMIT);
        assert!(preflight(&format!("<r>{}</r>", "<![CDATA[x]]><e/>".repeat(40_000))).is_ok());
        // Siblings release their declarations.
        let siblings: String = (0..1000)
            .map(|i| format!(r#"<e xmlns:p="urn:{i}"/>"#))
            .collect();
        assert!(parse(&format!("<r>{siblings}</r>")).is_ok());
        let deep = format!(
            "{}{}",
            "<e>".repeat(MAX_XML_DEPTH + 1),
            "</e>".repeat(MAX_XML_DEPTH + 1)
        );
        assert_eq!(parse(&deep).unwrap_err(), XML_LIMIT);
        // Aggregate work: many elements at the per-element cap.
        let heavy = format!("<r>{}</r>", attrs(MAX_ATTRIBUTES_PER_ELEMENT).repeat(9000));
        assert_eq!(preflight(&heavy).unwrap_err(), XML_LIMIT);
    }
}
