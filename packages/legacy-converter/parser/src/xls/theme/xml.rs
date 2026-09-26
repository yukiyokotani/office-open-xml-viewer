//! Bounded passive metadata XML. No DTDs, PIs, entities or external resolution.
use crate::xls::unsupported;
use quick_xml::{events::Event, name::ResolveResult, NsReader, XmlVersion};
use std::collections::BTreeMap;

pub(super) struct Node {
    pub ns: String,
    pub name: String,
    pub attrs: BTreeMap<String, String>,
    pub children: Vec<Node>,
}

impl Node {
    pub fn only_child(&self, ns: &str, name: &str) -> Result<&Self, String> {
        let mut matches = self
            .children
            .iter()
            .filter(|c| c.ns == ns && c.name == name);
        let child = matches
            .next()
            .ok_or_else(|| unsupported("missing BIFF theme XML child"))?;
        if matches.next().is_some() {
            return Err(unsupported("duplicate BIFF theme XML child"));
        }
        Ok(child)
    }
}

fn xml_character(c: char) -> bool {
    matches!(c as u32, 9 | 10 | 13 | 0x20..=0xd7ff | 0xe000..=0xfffd | 0x10000..=0x10ffff)
}

fn charge(remaining: &mut usize, size: usize) -> Result<(), String> {
    *remaining = remaining
        .checked_sub(size)
        .ok_or_else(|| unsupported("BIFF theme XML retained string budget exceeded"))?;
    Ok(())
}

pub(super) fn parse(bytes: &[u8]) -> Result<Node, String> {
    if bytes.len() > super::MAX_PART_BYTES {
        return Err(unsupported("BIFF theme XML byte budget exceeded"));
    }
    let text =
        std::str::from_utf8(bytes).map_err(|_| unsupported("BIFF theme XML must be UTF-8"))?;
    if !text.chars().all(xml_character) {
        return Err(unsupported("invalid BIFF theme XML character"));
    }
    let mut reader = NsReader::from_reader(bytes);
    reader.config_mut().expand_empty_elements = true;
    reader.config_mut().check_end_names = true;
    let mut stack = Vec::<Node>::new();
    let mut root = None;
    let mut declaration = false;
    // An inherited namespace URI can be copied into many nodes even in a small
    // input. Bound retained strings before copying; event limits bound node and
    // collection overhead separately. This is resource policy, not XML syntax.
    let mut string_budget = 2 * 1024 * 1024;
    // Resource policy: cap parser work and retained metadata tree, not just ZIP size.
    for _ in 0..20_000 {
        let (ns, event) = reader
            .read_resolved_event()
            .map_err(|_| unsupported("invalid BIFF theme XML"))?;
        if matches!(ns, ResolveResult::Unknown(_)) {
            return Err(unsupported("unbound BIFF theme XML prefix"));
        }
        match event {
            Event::Start(element) => {
                if stack.len() >= 32 || (stack.is_empty() && root.is_some()) {
                    return Err(unsupported("BIFF theme XML depth or root violation"));
                }
                let uri = match ns {
                    ResolveResult::Bound(n) => {
                        charge(&mut string_budget, n.as_ref().len())?;
                        String::from_utf8_lossy(n.as_ref()).into_owned()
                    }
                    _ => String::new(),
                };
                let mut attrs = BTreeMap::new();
                for (i, attribute) in element.attributes().enumerate() {
                    if i >= 64 {
                        return Err(unsupported("BIFF theme XML attribute budget exceeded"));
                    }
                    let attribute =
                        attribute.map_err(|_| unsupported("invalid BIFF theme XML attribute"))?;
                    // Only unqualified attributes are meaningful to this projection.
                    let (attribute_ns, local) = reader.resolver().resolve_attribute(attribute.key);
                    if matches!(attribute_ns, ResolveResult::Unknown(_)) {
                        return Err(unsupported("unbound BIFF theme attribute prefix"));
                    }
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                        .map_err(|_| unsupported("invalid BIFF theme XML attribute value"))?;
                    if !value.chars().all(xml_character) {
                        return Err(unsupported("invalid BIFF theme XML character reference"));
                    }
                    if !matches!(attribute_ns, ResolveResult::Unbound) {
                        continue;
                    }
                    charge(&mut string_budget, local.as_ref().len())?;
                    charge(&mut string_budget, value.len())?;
                    attrs.insert(
                        String::from_utf8_lossy(local.as_ref()).into_owned(),
                        value.into_owned(),
                    );
                }
                charge(&mut string_budget, element.local_name().as_ref().len())?;
                stack.push(Node {
                    ns: uri,
                    name: String::from_utf8_lossy(element.local_name().as_ref()).into_owned(),
                    attrs,
                    children: vec![],
                });
            }
            Event::End(_) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| unsupported("unmatched BIFF theme XML end"))?;
                if let Some(parent) = stack.last_mut() {
                    parent.children.push(node);
                } else {
                    root = Some(node);
                }
            }
            Event::Text(t) if t.as_ref().iter().all(u8::is_ascii_whitespace) => {}
            Event::Comment(_) => {}
            Event::Decl(d) if !declaration && stack.is_empty() && root.is_none() => {
                declaration = true;
                if d.version()
                    .map_err(|_| unsupported("invalid BIFF theme XML declaration"))?
                    .as_ref()
                    != b"1.0"
                    || d.encoding()
                        .transpose()
                        .map_err(|_| unsupported("invalid BIFF theme encoding"))?
                        .is_some_and(|e| !e.eq_ignore_ascii_case(b"UTF-8"))
                {
                    return Err(unsupported("unsupported BIFF theme XML declaration"));
                }
            }
            Event::Eof if stack.is_empty() => {
                return root.ok_or_else(|| unsupported("empty BIFF theme XML"))
            }
            _ => {
                return Err(unsupported(
                    "unsupported or incomplete BIFF theme XML content",
                ))
            }
        }
    }
    Err(unsupported("BIFF theme XML event budget exceeded"))
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn rejects_active_entities_malformed_roots_and_namespace_spoofing() {
        for xml in [
            "<!DOCTYPE x [<!ENTITY y SYSTEM 'file:///private'>]><x/>",
            "<?execute test?><x/>",
            "<x>&unknown;</x>",
            "<x/><y/>",
            "<x>",
            "<x a='1' a='2'/>",
            "<x p:a='1'/>",
            "<p:x/>",
            "<x>content</x>",
            "<x a='&#1;'/>",
            "<x>\u{0001}</x>",
        ] {
            assert!(parse(xml.as_bytes()).is_err(), "{xml}");
        }
        let normal =
            parse(b"<p:x xmlns:p='urn:test' a='A&amp;B' xmlns:q='urn:q' q:a='ignored'/>").unwrap();
        assert_eq!(normal.attrs.get("a").unwrap(), "A&B");
    }
    #[test]
    fn bounds_depth_attributes_and_event_fanout() {
        assert!(parse(format!("{}{}", "<x>".repeat(33), "</x>".repeat(33)).as_bytes()).is_err());
        assert!(parse(
            format!(
                "<x {}/>",
                (0..65)
                    .map(|i| format!("a{i}='x'"))
                    .collect::<Vec<_>>()
                    .join(" ")
            )
            .as_bytes()
        )
        .is_err());
        assert!(parse(format!("<x>{}</x>", "<y/>".repeat(10_000)).as_bytes()).is_err());
    }

    #[test]
    fn bounds_expansion_of_inherited_namespace_strings() {
        let xml = format!(
            "<x xmlns='urn:{}'>{}</x>",
            "n".repeat(32_000),
            "<y/>".repeat(100)
        );
        assert!(xml.len() < super::super::MAX_PART_BYTES);
        assert!(parse(xml.as_bytes()).is_err());
    }
}
