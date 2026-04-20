use roxmltree::{Document, Node};

pub const W_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
pub const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
pub const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub const WP_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
pub const PIC_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";

/// Find first child with given local name in any namespace.
pub fn child<'a, 'input>(node: Node<'a, 'input>, name: &str) -> Option<Node<'a, 'input>> {
    node.children().find(|n| n.tag_name().name() == name)
}

/// Find first child in w: namespace.
pub fn child_w<'a, 'input>(node: Node<'a, 'input>, name: &str) -> Option<Node<'a, 'input>> {
    node.children()
        .find(|n| n.tag_name().name() == name && n.tag_name().namespace() == Some(W_NS))
}

/// Collect all children in w: namespace with given name.
pub fn children_w<'a, 'input>(node: Node<'a, 'input>, name: &str) -> Vec<Node<'a, 'input>> {
    node.children()
        .filter(|n| n.tag_name().name() == name && n.tag_name().namespace() == Some(W_NS))
        .collect()
}

/// Element children with <w:sdt> wrappers transparently unwrapped. Structured Document
/// Tag (content control) blocks contain their real content inside <w:sdtContent>, and
/// most parsing stages should treat them as inline with the surrounding context.
pub fn element_children_flat<'a, 'input>(node: Node<'a, 'input>) -> Vec<Node<'a, 'input>> {
    let mut out = Vec::new();
    for child in node.children().filter(|n| n.is_element()) {
        let tn = child.tag_name();
        if tn.namespace() == Some(W_NS) && tn.name() == "sdt" {
            if let Some(content) = child_w(child, "sdtContent") {
                out.extend(element_children_flat(content));
            }
        } else {
            out.push(child);
        }
    }
    out
}

/// Like children_w but transparently descends into <w:sdt>/<w:sdtContent> wrappers.
pub fn children_w_flat<'a, 'input>(node: Node<'a, 'input>, name: &str) -> Vec<Node<'a, 'input>> {
    element_children_flat(node)
        .into_iter()
        .filter(|n| n.tag_name().name() == name && n.tag_name().namespace() == Some(W_NS))
        .collect()
}

/// Get attribute in w: namespace, falling back to no-namespace.
pub fn attr_w(node: Node, name: &str) -> Option<String> {
    node.attribute((W_NS, name))
        .or_else(|| node.attribute(name))
        .map(|s| s.to_string())
}

/// Get attribute in any namespace matching local name.
pub fn attr<'a>(node: Node<'a, '_>, ns: &str, name: &str) -> Option<String> {
    node.attribute((ns, name))
        .or_else(|| node.attribute(name))
        .map(|s| s.to_string())
}

/// Parse twips (1/20 pt) string to f64 pt.
pub fn twips_to_pt(s: &str) -> f64 {
    s.parse::<f64>().unwrap_or(0.0) / 20.0
}

/// Parse half-points string to f64 pt.
pub fn half_pt_to_pt(s: &str) -> f64 {
    s.parse::<f64>().unwrap_or(0.0) / 2.0
}

/// Parse EMU string to f64 pt.
pub fn emu_to_pt(s: &str) -> f64 {
    s.parse::<f64>().unwrap_or(0.0) / 12700.0
}

/// Parse a ST_OnOff-style toggle child element. ECMA-376 §17.3.2.22 allows
/// "true"/"false"/"1"/"0"/"on"/"off" (and absent val attribute = true).
/// Returns None if the element itself is absent so the caller can distinguish
/// "explicitly turned off" from "inherited from parent".
pub fn bool_prop(node: Node, tag: &str) -> Option<bool> {
    let child = child_w(node, tag)?;
    let val = attr_w(child, "val");
    Some(match val.as_deref() {
        Some("0") | Some("false") | Some("off") => false,
        _ => true,
    })
}

pub fn find_root_element<'a, 'input>(doc: &'a Document<'input>, tag: &str) -> Option<Node<'a, 'input>> {
    doc.root_element()
        .descendants()
        .find(|n| n.tag_name().name() == tag)
}
