//! Typed inspection of one DrawingML alternative shape, MS-ODRAW 2.3.4.41-42.
//!
//! Not yet admitted to conversion: Office can replace alternative text with
//! positional placeholders. Neither equal lengths nor an empty textCheckSum
//! establish correspondence with the binary text. Keep literal evidence intact.
//! The integration test compiles this reader until admission is implemented.
use quick_xml::{events::Event, name::ResolveResult, NsReader, XmlVersion};

const P: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
const A: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";

/// Resource policy, not format limits. One caller retains these counters across
/// shapes. ZIP metadata and inflation must be bounded before calling this reader.
pub struct Budget {
    pub bytes: usize,
    pub events: usize,
    pub paragraphs: usize,
}

#[derive(Debug, Default, PartialEq, Eq)]
pub struct Paragraph {
    pub margin_left: Option<i32>,
    pub margin_right: Option<i32>,
    pub indent: Option<i32>,
    pub default_tab_size: Option<i32>,
    pub level: Option<u8>,
    /// Original XML text, NOT substituted binary text or a correspondence key.
    pub literal: String,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Node {
    Shape,
    Body,
    Paragraph,
    Properties,
    Run,
    Text,
    Other,
}

fn charge(counter: &mut usize, count: usize) -> Result<(), String> {
    *counter = counter
        .checked_sub(count)
        .ok_or("alternative XML resource budget exceeded")?;
    Ok(())
}

/// Read direct scalar paragraph properties of a single p:sp/p:txBody.
/// Unsupported roots, fields and break elements have no projection. Arbitrary
/// XML, relationships, scripts and binary objects never enter the result.
pub fn read(xml: &[u8], budget: &mut Budget) -> Result<Option<Vec<Paragraph>>, String> {
    // Bound a single token as well as total retained literal data. Slice-based
    // quick-xml events borrow input; no unbounded event buffer is allocated.
    if xml.len() > 1024 * 1024 {
        return Err("alternative XML part exceeds byte limit".into());
    }
    charge(&mut budget.bytes, xml.len())?;
    let source = std::str::from_utf8(xml).map_err(|_| "alternative XML must be UTF-8")?;
    if !source.chars().all(xml_character) {
        return Err("invalid XML 1.0 character".into());
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().expand_empty_elements = true;
    reader.config_mut().check_end_names = true;
    let mut stack = Vec::new();
    let mut paragraphs = Vec::<Paragraph>::new();
    let mut root_seen = false;
    let mut declaration_seen = false;
    let mut body_seen = false;
    let mut properties_seen = false;
    loop {
        charge(&mut budget.events, 1)?;
        let (namespace, event) = reader.read_resolved_event().map_err(|e| e.to_string())?;
        if matches!(namespace, ResolveResult::Unknown(_)) {
            return Err("unbound alternative XML prefix".into());
        }
        match event {
            Event::Start(element) => {
                if stack.len() >= 64 {
                    return Err("alternative XML depth limit exceeded".into());
                }
                let uri = match &namespace {
                    ResolveResult::Bound(n) => n.as_ref(),
                    _ => b"",
                };
                let local = element.local_name();
                let name = local.as_ref();
                let parent = stack.last().copied();
                let node = match (parent, uri, name) {
                    (None, P, b"sp") if !root_seen => {
                        root_seen = true;
                        Node::Shape
                    }
                    (None, _, _) if !root_seen => return Ok(None),
                    (None, _, _) => return Err("multiple alternative XML roots".into()),
                    (Some(Node::Shape), P, b"txBody") => {
                        if body_seen {
                            return Err("duplicate alternative text body".into());
                        }
                        body_seen = true;
                        Node::Body
                    }
                    (Some(Node::Body), A, b"p") => {
                        charge(&mut budget.paragraphs, 1)?;
                        paragraphs.push(Paragraph::default());
                        properties_seen = false;
                        Node::Paragraph
                    }
                    (Some(Node::Paragraph), A, b"pPr") => {
                        if properties_seen {
                            return Err("duplicate alternative paragraph properties".into());
                        }
                        properties_seen = true;
                        Node::Properties
                    }
                    (Some(Node::Paragraph), A, b"r") => Node::Run,
                    (Some(Node::Run), A, b"t") => Node::Text,
                    (Some(Node::Paragraph), A, b"fld" | b"br") => return Ok(None),
                    _ => Node::Other,
                };
                // Check duplicate attributes even when they are not projected.
                for (index, attribute) in element.attributes().enumerate() {
                    if index >= 64 {
                        return Err("alternative XML attribute limit exceeded".into());
                    }
                    charge(&mut budget.events, 1)?;
                    let attribute = attribute.map_err(|e| e.to_string())?;
                    if node != Node::Properties {
                        continue;
                    }
                    let key = attribute.key.as_ref();
                    if !matches!(key, b"marL" | b"marR" | b"indent" | b"defTabSz" | b"lvl") {
                        continue;
                    }
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                        .map_err(|e| e.to_string())?;
                    let value = value
                        .trim()
                        .parse::<i32>()
                        .map_err(|_| "invalid alternative paragraph integer")?;
                    let paragraph = paragraphs
                        .last_mut()
                        .ok_or("missing alternative paragraph")?;
                    // ECMA-376 ST_TextMargin, ST_TextIndent, ST_Coordinate32,
                    // ST_TextIndentLevelType. No clamping or unit conversion.
                    match key {
                        b"marL" | b"marR" if !(0..=51_206_400).contains(&value) => {
                            return Err("alternative margin outside schema range".into())
                        }
                        b"indent" if !(-51_206_400..=51_206_400).contains(&value) => {
                            return Err("alternative indent outside schema range".into())
                        }
                        b"lvl" if !(0..=8).contains(&value) => {
                            return Err("alternative level outside schema range".into())
                        }
                        _ => {}
                    }
                    match key {
                        b"marL" => paragraph.margin_left = Some(value),
                        b"marR" => paragraph.margin_right = Some(value),
                        b"indent" => paragraph.indent = Some(value),
                        b"defTabSz" => paragraph.default_tab_size = Some(value),
                        b"lvl" => paragraph.level = Some(value as u8),
                        _ => unreachable!(),
                    }
                }
                stack.push(node);
            }
            Event::End(_) => {
                stack.pop().ok_or("unmatched alternative XML end")?;
            }
            Event::Text(text) => {
                let text = text.xml10_content().map_err(|e| e.to_string())?;
                if stack.last() == Some(&Node::Text) {
                    paragraphs
                        .last_mut()
                        .ok_or("text outside paragraph")?
                        .literal
                        .push_str(&text);
                } else if stack.is_empty() && !text.trim().is_empty() {
                    return Err("text outside alternative XML root".into());
                }
            }
            Event::CData(text) => {
                if stack.last() != Some(&Node::Text) {
                    return Err("unexpected alternative CDATA".into());
                }
                paragraphs
                    .last_mut()
                    .ok_or("text outside paragraph")?
                    .literal
                    .push_str(&text.decode().map_err(|e| e.to_string())?);
            }
            Event::GeneralRef(reference) => {
                let character = match reference.resolve_char_ref().map_err(|e| e.to_string())? {
                    Some(c) => c,
                    None => match reference.as_ref() {
                        b"amp" => '&',
                        b"lt" => '<',
                        b"gt" => '>',
                        b"apos" => '\'',
                        b"quot" => '"',
                        _ => return Err("unsupported alternative XML entity".into()),
                    },
                };
                if !xml_character(character) {
                    return Err("invalid XML 1.0 character reference".into());
                }
                if stack.last() != Some(&Node::Text) {
                    return Err("entity outside alternative text".into());
                }
                paragraphs
                    .last_mut()
                    .ok_or("text outside paragraph")?
                    .literal
                    .push(character);
            }
            Event::DocType(_) | Event::PI(_) => {
                return Err("active or DTD alternative XML is unsupported".into())
            }
            Event::Decl(declaration) => {
                if root_seen || declaration_seen {
                    return Err("misplaced alternative XML declaration".into());
                }
                declaration_seen = true;
                if declaration.version().map_err(|e| e.to_string())?.as_ref() != b"1.0" {
                    return Err("alternative XML must be XML 1.0".into());
                }
                if let Some(encoding) = declaration.encoding() {
                    if !encoding
                        .map_err(|e| e.to_string())?
                        .eq_ignore_ascii_case(b"UTF-8")
                    {
                        return Err("alternative XML must declare UTF-8".into());
                    }
                }
            }
            Event::Eof => {
                if !stack.is_empty() || !root_seen {
                    return Err("incomplete alternative XML".into());
                }
                return Ok(body_seen.then_some(paragraphs));
            }
            Event::Comment(_) => {}
            Event::Empty(_) => unreachable!("empty elements expanded"),
        }
    }
}

fn xml_character(c: char) -> bool {
    matches!(c, '\u{9}' | '\u{a}' | '\u{d}' | '\u{20}'..='\u{d7ff}' | '\u{e000}'..='\u{fffd}' | '\u{10000}'..='\u{10ffff}')
}
