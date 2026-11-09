//! Formula objects (ODF Part 3 §10.4.6 `draw:object` holding a MathML
//! document, ODF 1.3 Part 3 §5.2 of the formula family): MathML 2
//! presentation markup mapped onto the shared OMML math nodes the DOCX
//! renderer draws.

use ooxml_common::math::MathNode;

const NS_MATH: &str = "http://www.w3.org/1998/Math/MathML";
/// Implementation bound on MathML nesting.
const MAX_DEPTH: usize = 64;

/// The nodes of a MathML `math` element.
pub(super) fn convert(math: roxmltree::Node<'_, '_>) -> Vec<MathNode> {
    children(math, 0)
}

fn is_math(node: &roxmltree::Node<'_, '_>) -> bool {
    node.is_element() && node.tag_name().namespace() == Some(NS_MATH)
}

fn children(node: roxmltree::Node<'_, '_>, depth: usize) -> Vec<MathNode> {
    if depth > MAX_DEPTH {
        return Vec::new();
    }
    node.children()
        .filter(is_math)
        .flat_map(|child| element(child, depth + 1))
        .collect()
}

/// The `n`th element child as a node list.
fn argument(node: roxmltree::Node<'_, '_>, n: usize, depth: usize) -> Vec<MathNode> {
    node.children()
        .filter(is_math)
        .nth(n)
        .map(|child| element(child, depth + 1))
        .unwrap_or_default()
}

fn text(node: roxmltree::Node<'_, '_>) -> String {
    node.descendants()
        .filter(|n| n.is_text())
        .filter_map(|n| n.text())
        .collect::<String>()
        .trim()
        .to_string()
}

fn run(text: String, style: &str) -> MathNode {
    MathNode::Run {
        text,
        style: style.to_string(),
        script: None,
    }
}

/// The OMML n-ary operators MathML writes as large operators.
fn nary(op: &str) -> bool {
    matches!(
        op,
        "∑" | "∏"
            | "∐"
            | "∫"
            | "∬"
            | "∭"
            | "∮"
            | "∯"
            | "∰"
            | "⋃"
            | "⋂"
            | "⋁"
            | "⋀"
            | "⨁"
            | "⨂"
    )
}

fn element(node: roxmltree::Node<'_, '_>, depth: usize) -> Vec<MathNode> {
    if depth > MAX_DEPTH {
        return Vec::new();
    }
    let name = node.tag_name().name();
    let variant = node.attribute("mathvariant");
    match name {
        "semantics" => argument(node, 0, depth),
        "annotation" | "annotation-xml" | "none" | "mprescripts" => Vec::new(),
        "mrow" | "mstyle" | "mpadded" | "merror" | "math" | "maction" => {
            vec![MathNode::Group {
                items: children(node, depth),
            }]
        }
        "mphantom" => vec![MathNode::Phant {
            show: false,
            zero_wid: false,
            zero_asc: false,
            zero_desc: false,
            base: children(node, depth),
        }],
        // §3.2.3 mi: a single-character identifier is italic by default.
        "mi" => {
            let content = text(node);
            let style = match variant {
                Some("normal") => "roman",
                Some("bold") => "bold",
                Some("bold-italic") => "boldItalic",
                Some("italic") => "italic",
                _ if content.chars().count() == 1 => "italic",
                _ => "roman",
            };
            vec![run(content, style)]
        }
        "mn" | "mo" | "ms" => {
            let style = match variant {
                Some("bold") => "bold",
                Some("italic") => "italic",
                _ => "roman",
            };
            vec![run(text(node), style)]
        }
        "mtext" => vec![run(text(node), "roman")],
        "mspace" => vec![run("\u{2009}".into(), "roman")],
        "mfrac" => vec![MathNode::Fraction {
            num: argument(node, 0, depth),
            den: argument(node, 1, depth),
            bar: (node.attribute("linethickness") == Some("0")).then_some(false),
        }],
        "msqrt" => vec![MathNode::Radical {
            index: Vec::new(),
            radicand: children(node, depth),
        }],
        "mroot" => vec![MathNode::Radical {
            index: argument(node, 1, depth),
            radicand: argument(node, 0, depth),
        }],
        "msub" | "msup" | "msubsup" | "munder" | "mover" | "munderover" => {
            let base = argument(node, 0, depth);
            let first = argument(node, 1, depth);
            let second = argument(node, 2, depth);
            let base_operator = node
                .children()
                .find(is_math)
                .filter(|b| b.tag_name().name() == "mo")
                .map(text)
                .filter(|op| nary(op));
            if let Some(op) = base_operator {
                let (sub, sup) = match name {
                    "msub" | "munder" => (first, Vec::new()),
                    "msup" | "mover" => (Vec::new(), first),
                    _ => (first, second),
                };
                return vec![MathNode::Nary {
                    op,
                    lim_loc: if name.starts_with("mu") {
                        "undOvr".into()
                    } else {
                        "subSup".into()
                    },
                    sub,
                    sup,
                    body: Vec::new(),
                }];
            }
            match name {
                "msub" => vec![MathNode::Sub { base, sub: first }],
                "msup" => vec![MathNode::Sup { base, sup: first }],
                "msubsup" => vec![MathNode::SubSup {
                    base,
                    sub: first,
                    sup: second,
                }],
                "munder" => vec![MathNode::Limit {
                    base,
                    lower: first,
                    upper: Vec::new(),
                }],
                "mover" => {
                    // A single accent character over the base.
                    let accent = match first.as_slice() {
                        [MathNode::Run { text, .. }] if text.chars().count() == 1 => {
                            Some(text.clone())
                        }
                        _ => None,
                    };
                    match accent {
                        Some(chr) if node.attribute("accent") != Some("false") => {
                            vec![MathNode::Accent { chr, base }]
                        }
                        _ => vec![MathNode::Limit {
                            base,
                            lower: Vec::new(),
                            upper: first,
                        }],
                    }
                }
                _ => vec![MathNode::Limit {
                    base,
                    lower: first,
                    upper: second,
                }],
            }
        }
        "mmultiscripts" => {
            let items: Vec<roxmltree::Node<'_, '_>> = node.children().filter(is_math).collect();
            let base = items
                .first()
                .map(|b| element(*b, depth + 1))
                .unwrap_or_default();
            let pre = items
                .iter()
                .position(|n| n.tag_name().name() == "mprescripts");
            match pre {
                Some(at) => vec![MathNode::SPre {
                    sub: items
                        .get(at + 1)
                        .map(|n| element(*n, depth + 1))
                        .unwrap_or_default(),
                    sup: items
                        .get(at + 2)
                        .map(|n| element(*n, depth + 1))
                        .unwrap_or_default(),
                    base,
                }],
                None => vec![MathNode::SubSup {
                    base,
                    sub: items
                        .get(1)
                        .map(|n| element(*n, depth + 1))
                        .unwrap_or_default(),
                    sup: items
                        .get(2)
                        .map(|n| element(*n, depth + 1))
                        .unwrap_or_default(),
                }],
            }
        }
        "mfenced" => vec![MathNode::Delimiter {
            beg_char: node.attribute("open").unwrap_or("(").to_string(),
            end_char: node.attribute("close").unwrap_or(")").to_string(),
            items: node
                .children()
                .filter(is_math)
                .map(|child| element(child, depth + 1))
                .collect(),
        }],
        "menclose" => vec![MathNode::BorderBox {
            hide_top: false,
            hide_bot: false,
            hide_left: false,
            hide_right: false,
            strike_h: false,
            strike_v: false,
            strike_bltr: false,
            strike_tlbr: false,
            base: children(node, depth),
        }],
        "mtable" => vec![MathNode::Array {
            rows: node
                .children()
                .filter(|row| is_math(row) && matches!(row.tag_name().name(), "mtr" | "mlabeledtr"))
                .map(|row| {
                    row.children()
                        .filter(|cell| is_math(cell) && cell.tag_name().name() == "mtd")
                        .map(|cell| children(cell, depth + 2))
                        .collect()
                })
                .collect(),
            align: "center".into(),
        }],
        _ => children(node, depth),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn converts_presentation_mathml() {
        let xml = r#"<math xmlns="http://www.w3.org/1998/Math/MathML"><semantics><mrow><msup><mi>x</mi><mn>2</mn></msup><mo>+</mo><mfrac><mi>a</mi><mi>b</mi></mfrac><msqrt><mi>y</mi></msqrt></mrow><annotation>x^2</annotation></semantics></math>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let nodes = convert(doc.root_element());
        let MathNode::Group { items } = &nodes[0] else {
            panic!()
        };
        assert!(matches!(items[0], MathNode::Sup { .. }));
        assert!(matches!(items[2], MathNode::Fraction { .. }));
        assert!(matches!(items[3], MathNode::Radical { .. }));
    }
}
