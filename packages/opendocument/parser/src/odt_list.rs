//! Lists (ODF Part 3 §5.3) and list styles (§16.30-§16.34), measured against
//! LibreOffice 26.8 PDF export (design addendum rev 9).
//!
//! A label is emitted as the paragraph's leading runs, flagged `__listLabel`,
//! which the DOCX layout places with LibreOffice's label geometry: label lines
//! keep the first-line indent through the body's first line, the separator
//! tab breaks before a stop at the line end and has no width at a line start,
//! a label-only line fits the next word as an empty line would, label spaces
//! take no justification, and end/centre labels end or centre at M + I. The
//! DOCX `NumberingInfo` marker cannot carry this: it keeps the body at or
//! after the paragraph indent, while LibreOffice starts it at the list tab
//! stop T even when T is before M.

use super::*;
use docx_model::TabStop;

/// Levels a list style defines (§19.834 `text:level`).
const MAX_LEVELS: u32 = 10;
/// Authored start values LibreOffice keeps (an item start of 32768 renders
/// "1.", measured); counting itself was measured to 32768.
const MAX_AUTHORED_START: u32 = 32767;
const MAX_COUNTED: u32 = 32768;
/// Generated tab stops per paragraph, and the smallest grid interval.
const MAX_TAB_STOPS: usize = 1024;
const MIN_TAB_INTERVAL_PT: f64 = 0.1 * 72.0 / 2.54;
/// Affix length bound (bytes), checked before the label is built.
const MAX_AFFIX_BYTES: usize = 256;

#[derive(Clone, Copy, PartialEq, Debug)]
pub(super) enum Follow {
    Tab,
    Space,
    Nothing,
}

#[derive(Clone, Copy, PartialEq, Debug)]
pub(super) enum Align {
    Start,
    End,
    Center,
}

impl Align {
    fn wire(self) -> &'static str {
        match self {
            Align::Start => "start",
            Align::End => "end",
            Align::Center => "center",
        }
    }
}

#[derive(Clone, Copy, PartialEq, Debug)]
enum Format {
    Decimal,
    LowerLetter,
    UpperLetter,
    LowerRoman,
    UpperRoman,
    /// "①, ②, ③, ..." (measured past 20: ⑲ ⑳ ㉑ ㉒; up to ㊿).
    Circled,
    /// "ア, イ, ウ, ..." and its half-width form: the 46 kana of the
    /// gojūon order, repeating (measured: ヲ ン ア).
    Katakana,
    HalfwidthKatakana,
    /// An empty `style:num-format`: no counter component.
    Empty,
}

const KATAKANA: &str =
    "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン";
const HALFWIDTH_KATAKANA: &str = "ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜｦﾝ";

#[derive(Clone, Debug)]
enum Kind {
    Number {
        format: Format,
        prefix: String,
        suffix: String,
        display: u32,
        letter_sync: bool,
    },
    Bullet(char),
}

/// One level of a list style, in label alignment mode.
#[derive(Clone)]
struct Level<'a> {
    kind: Kind,
    start: u32,
    follow: Follow,
    align: Align,
    /// `fo:margin-left` (M), `fo:text-indent` (I) and
    /// `text:list-tab-stop-position` (T), in pt; T from the text area.
    margin: f64,
    indent: f64,
    tab: Option<f64>,
    /// The label's own formatting: the level's text properties (a bullet
    /// face) before its character style's chain.
    styles: Vec<StyleRef<'a>>,
}

/// A list paragraph's label and geometry, handed to `Reader::paragraph`.
pub(super) struct ListParagraph<'a> {
    pub style_name: String,
    pub margin: f64,
    pub indent: f64,
    pub tab: Option<f64>,
    pub label: Option<Label<'a>>,
}

pub(super) struct Label<'a> {
    pub text: String,
    pub follow: Follow,
    pub align: Align,
    pub styles: Vec<StyleRef<'a>>,
}

/// One LibreOffice list: a list and the lists that continue it share it.
#[derive(Default)]
struct Chain {
    counters: [Option<u32>; MAX_LEVELS as usize],
    counted: bool,
}

#[derive(Default)]
pub(super) struct Lists {
    chains: Vec<Chain>,
    /// `xml:id` of top-level lists, to their chains.
    ids: HashMap<String, usize>,
    /// `xml:id` of nested lists (continuing one rejects).
    nested_ids: HashSet<String>,
    /// The immediately preceding top-level list: its style and chain.
    previous: Option<(String, usize)>,
}

/// The list a paragraph sits in while it is read.
#[derive(Clone)]
pub(super) struct Frame {
    chain: usize,
    style: String,
    level: u32,
}

impl<'a> Reader<'a> {
    /// A `text:list` (§5.3.1). Every item's first paragraph is a numbered
    /// node at its nesting depth; wrapper items (whose first child is a list)
    /// and headers are not. A continuation chain is one LibreOffice list.
    pub(super) fn list(
        &mut self,
        node: roxmltree::Node<'_, '_>,
        parent: Option<&Frame>,
        depth: usize,
    ) -> Result<(), String> {
        if depth > MAX_SECTION_DEPTH + MAX_LEVELS as usize {
            return Err(unsupported("list-depth"));
        }
        let mut continue_numbering = false;
        let mut continue_list = None;
        let mut style_name = None;
        let mut id = None;
        for attribute in node.attributes() {
            match (attribute.namespace(), attribute.name()) {
                (Some(NS_TEXT), "style-name") => style_name = Some(attribute.value()),
                (Some(NS_TEXT), "continue-numbering") => {
                    continue_numbering = attribute.value() == "true"
                }
                (Some(NS_TEXT), "continue-list") => continue_list = Some(attribute.value()),
                (Some(NS_XML), "id") => id = Some(attribute.value()),
                _ => return Err(unsupported("list-attribute")),
            }
        }
        let frame = match parent {
            Some(parent) => {
                // A nested list continues its enclosing list; another style
                // or a continuation is not measured.
                if style_name.is_some_and(|name| name != parent.style)
                    || continue_numbering
                    || continue_list.is_some()
                {
                    return Err(unsupported("list-nested-style"));
                }
                if parent.level >= MAX_LEVELS {
                    return Err(unsupported("list-level"));
                }
                if let Some(id) = id {
                    self.charge_construction(id.len())?;
                    self.lists.nested_ids.insert(id.to_string());
                }
                Frame {
                    chain: parent.chain,
                    style: parent.style.clone(),
                    level: parent.level + 1,
                }
            }
            None => {
                let style = style_name.ok_or_else(|| unsupported("list-style-missing"))?;
                if continue_numbering && continue_list.is_some() {
                    return Err(unsupported("list-continuation"));
                }
                let continued = if let Some(target) = continue_list {
                    if self.lists.nested_ids.contains(target) {
                        return Err(unsupported("list-continuation"));
                    }
                    Some(
                        *self
                            .lists
                            .ids
                            .get(target)
                            .ok_or_else(|| unsupported("list-continuation"))?,
                    )
                } else {
                    // `continue-numbering` without a preceding list of the
                    // same style restarts (§19.783).
                    match &self.lists.previous {
                        Some((previous, chain)) if continue_numbering && previous == style => {
                            Some(*chain)
                        }
                        _ => None,
                    }
                };
                let chain = match continued {
                    // LibreOffice initialises a list without a counted item
                    // differently for an empty and a header-only list
                    // (measured by the reviewer); continuing one rejects.
                    Some(chain) if !self.lists.chains[chain].counted => {
                        return Err(unsupported("list-continuation"));
                    }
                    Some(chain) => chain,
                    None => {
                        self.charge_construction(std::mem::size_of::<Chain>())?;
                        self.lists.chains.push(Chain::default());
                        self.lists.chains.len() - 1
                    }
                };
                self.charge_construction(style.len())?;
                self.lists.previous = Some((style.to_string(), chain));
                if let Some(id) = id {
                    self.charge_construction(id.len())?;
                    self.lists.ids.insert(id.to_string(), chain);
                }
                Frame {
                    chain,
                    style: style.to_string(),
                    level: 1,
                }
            }
        };
        let level = self.list_level(&frame.style, frame.level)?;
        for child in node.children() {
            if child.is_text() {
                if child.text().is_some_and(|t| !t.trim().is_empty()) {
                    return Err(invalid("text in a list"));
                }
                continue;
            }
            if !child.is_element() {
                continue;
            }
            match (child.tag_name().namespace(), child.tag_name().name()) {
                (Some(NS_TEXT), "list-item") => self.list_item(child, &frame, &level, depth)?,
                (Some(NS_TEXT), "list-header") => self.list_header(child, &frame, &level, depth)?,
                _ => return Err(unsupported("list-child")),
            }
        }
        Ok(())
    }

    fn list_item(
        &mut self,
        node: roxmltree::Node<'_, '_>,
        frame: &Frame,
        level: &Level<'a>,
        depth: usize,
    ) -> Result<(), String> {
        let mut start = None;
        for attribute in node.attributes() {
            match (attribute.namespace(), attribute.name()) {
                (Some(NS_TEXT), "start-value") => {
                    let value = attribute
                        .value()
                        .parse::<u32>()
                        .map_err(|_| invalid("bad list item start value"))?;
                    if value > MAX_AUTHORED_START {
                        return Err(unsupported("list-start-value"));
                    }
                    start = Some(value);
                }
                (Some(NS_XML), "id") => {}
                _ => return Err(unsupported("list-attribute")),
            }
        }
        let mut sublists = 0;
        for (index, child) in block_children(node)?.into_iter().enumerate() {
            match (child.tag_name().namespace(), child.tag_name().name()) {
                // The item's first paragraph is its numbered node.
                (Some(NS_TEXT), "p") if index == 0 => {
                    let values = self.count(frame, start)?;
                    let label = self.label(frame, level, &values)?;
                    self.list_paragraph = Some(ListParagraph {
                        style_name: frame.style.clone(),
                        margin: level.margin,
                        indent: level.indent,
                        tab: level.tab,
                        label: Some(label),
                    });
                    self.paragraph(child, None)?;
                }
                (Some(NS_TEXT), "p") => self.unlabelled(child, frame, level)?,
                (Some(NS_TEXT), "list") => {
                    // One sublist per item: a second restarts its level in
                    // LibreOffice, outside the counter model (real corpus: 0).
                    sublists += 1;
                    if sublists > 1 {
                        return Err(unsupported("list-sublists"));
                    }
                    self.list(child, Some(frame), depth + 1)?;
                }
                // Headings in lists are not measured yet.
                (Some(NS_TEXT), "h") => return Err(unsupported("list-heading")),
                _ => return Err(unsupported("list-child")),
            }
        }
        Ok(())
    }

    /// A `text:list-header`: unnumbered paragraphs at M.
    fn list_header(
        &mut self,
        node: roxmltree::Node<'_, '_>,
        frame: &Frame,
        level: &Level<'a>,
        depth: usize,
    ) -> Result<(), String> {
        for attribute in node.attributes() {
            if (attribute.namespace(), attribute.name()) != (Some(NS_XML), "id") {
                return Err(unsupported("list-attribute"));
            }
        }
        let mut sublists = 0;
        for child in block_children(node)? {
            match (child.tag_name().namespace(), child.tag_name().name()) {
                (Some(NS_TEXT), "p") => self.unlabelled(child, frame, level)?,
                (Some(NS_TEXT), "list") => {
                    sublists += 1;
                    if sublists > 1 {
                        return Err(unsupported("list-sublists"));
                    }
                    self.list(child, Some(frame), depth + 1)?;
                }
                (Some(NS_TEXT), "h") => return Err(unsupported("list-heading")),
                _ => return Err(unsupported("list-child")),
            }
        }
        Ok(())
    }

    fn unlabelled(
        &mut self,
        node: roxmltree::Node<'_, '_>,
        frame: &Frame,
        level: &Level<'a>,
    ) -> Result<(), String> {
        self.list_paragraph = Some(ListParagraph {
            style_name: frame.style.clone(),
            margin: level.margin,
            indent: 0.0,
            tab: level.tab,
            label: None,
        });
        self.paragraph(node, None)
    }

    /// LibreOffice's numbering of a node at `frame.level` (measured): every
    /// unset ancestor level takes its level's start (a phantom parent, whose
    /// item attributes are ignored), the node's level takes the item's start,
    /// else counts on, else takes its start, and deeper levels are unset.
    fn count(&mut self, frame: &Frame, start: Option<u32>) -> Result<Vec<u32>, String> {
        let mut starts = Vec::with_capacity(frame.level as usize);
        for level in 1..=frame.level {
            starts.push(self.list_level(&frame.style, level)?.start);
        }
        let chain = &mut self.lists.chains[frame.chain];
        let index = (frame.level - 1) as usize;
        for (ancestor, ancestor_start) in starts.iter().enumerate().take(index) {
            chain.counters[ancestor].get_or_insert(*ancestor_start);
        }
        let value = match (start, chain.counters[index]) {
            (Some(start), _) => start,
            (None, Some(previous)) => previous + 1,
            (None, None) => starts[index],
        };
        if value > MAX_COUNTED {
            return Err(unsupported("list-counter"));
        }
        chain.counters[index] = Some(value);
        for deeper in chain.counters.iter_mut().skip(index + 1) {
            *deeper = None;
        }
        chain.counted = true;
        Ok(chain.counters[..=index]
            .iter()
            .map(|v| v.expect("ancestors are set"))
            .collect())
    }

    /// The label text (measured composition): a bullet's character; for a
    /// number, an empty own format shows only the affixes, otherwise the
    /// selected levels joined by "." (an empty-format ancestor is left out
    /// with its separator, a bullet ancestor contributes an empty component)
    /// between the prefix and the suffix.
    fn label(
        &mut self,
        frame: &Frame,
        level: &Level<'a>,
        values: &[u32],
    ) -> Result<Label<'a>, String> {
        let text = match &level.kind {
            Kind::Bullet(c) => c.to_string(),
            Kind::Number {
                format,
                prefix,
                suffix,
                display,
                letter_sync,
            } => {
                let mut components = Vec::new();
                if *format != Format::Empty {
                    let first = frame.level.saturating_sub(display - 1).max(1);
                    for l in first..=frame.level {
                        let component = if l == frame.level {
                            (*format, *letter_sync)
                        } else {
                            match self.list_level(&frame.style, l)?.kind {
                                Kind::Bullet(_) => {
                                    components.push(String::new());
                                    continue;
                                }
                                Kind::Number {
                                    format: Format::Empty,
                                    ..
                                } => continue,
                                Kind::Number {
                                    format,
                                    letter_sync,
                                    ..
                                } => (format, letter_sync),
                            }
                        };
                        components.push(self.format_counter(
                            values[(l - 1) as usize],
                            component.0,
                            component.1,
                        )?);
                    }
                }
                format!("{prefix}{}{suffix}", components.join("."))
            }
        };
        self.charge_construction(text.len())?;
        Ok(Label {
            text,
            follow: level.follow,
            align: level.align,
            styles: level.styles.clone(),
        })
    }

    fn format_counter(&self, value: u32, format: Format, sync: bool) -> Result<String, String> {
        // 0 renders "0" in every format (measured by the reviewer).
        if value == 0 && format != Format::Empty {
            return Ok("0".to_string());
        }
        Ok(match format {
            Format::Decimal => value.to_string(),
            Format::Empty => String::new(),
            Format::LowerRoman => roman(value).to_ascii_lowercase(),
            Format::Circled => match value {
                1..=20 => char::from_u32(0x2460 + value - 1),
                21..=35 => char::from_u32(0x3251 + value - 21),
                36..=50 => char::from_u32(0x32B1 + value - 36),
                _ => None,
            }
            .ok_or_else(|| unsupported("list-format-range"))?
            .to_string(),
            Format::Katakana | Format::HalfwidthKatakana => {
                let table = if format == Format::Katakana {
                    KATAKANA
                } else {
                    HALFWIDTH_KATAKANA
                };
                table
                    .chars()
                    .nth(((value - 1) % 46) as usize)
                    .expect("46 kana")
                    .to_string()
            }
            Format::UpperRoman => roman(value),
            Format::LowerLetter | Format::UpperLetter => {
                // ODF 1.4 §19.505 matches LibreOffice beyond "z"; ODF 1.3
                // states the opposite, so only 1.4 documents admit it.
                if value > 26 && !self.odf_1_4 {
                    return Err(unsupported("list-letter-sync"));
                }
                let letters = letters(value, sync);
                if format == Format::UpperLetter {
                    letters.to_ascii_uppercase()
                } else {
                    letters
                }
            }
        })
    }

    /// A list style's level definition (§16.31-§16.34), label alignment mode
    /// only.
    fn list_level(&mut self, style: &str, level: u32) -> Result<Level<'a>, String> {
        let list = self
            .sheet
            .find_list(Scope::Content, style)
            .ok_or_else(|| invalid("a list names a missing list style"))?;
        self.charge_work(16 + style.len())?;
        for (key, value) in &list.element.attributes {
            match (key.0.as_str(), key.1.as_str()) {
                (NS_STYLE, "name" | "display-name") => {}
                (NS_TEXT, "consecutive-numbering") if value == "false" => {}
                _ => return Err(unsupported("list-style-attribute")),
            }
        }
        self.charge_work(list.element.children.len())?;
        let definition = list
            .element
            .children
            .iter()
            .find(|child| {
                child
                    .attributes
                    .get(&(NS_TEXT.to_string(), "level".to_string()))
                    .and_then(|value| value.parse::<u32>().ok())
                    == Some(level)
            })
            .ok_or_else(|| unsupported("list-level-missing"))?;
        let attribute = |name: (&str, &str)| {
            definition
                .attributes
                .get(&(name.0.to_string(), name.1.to_string()))
                .map(String::as_str)
        };
        let number = match (definition.namespace.as_str(), definition.name.as_str()) {
            (NS_TEXT, "list-level-style-number") => true,
            (NS_TEXT, "list-level-style-bullet") => false,
            (NS_TEXT, "list-level-style-image") => return Err(unsupported("list-image")),
            _ => return Err(unsupported("list-style-child")),
        };
        for (key, value) in &definition.attributes {
            let admitted = match (key.0.as_str(), key.1.as_str()) {
                (NS_TEXT, "level" | "style-name") => true,
                (NS_STYLE, "num-format" | "num-prefix" | "num-suffix" | "num-letter-sync") => true,
                (NS_TEXT, "display-levels" | "start-value") => number,
                (NS_TEXT, "bullet-char") => !number,
                (NS_TEXT, "bullet-relative-size") => !number && value == "100%",
                // LibreOffice's label template, checked against the
                // composition below; a bullet draws its character only.
                (NS_LOEXT, "num-list-format") => true,
                (NS_LOEXT, "is-legal") => value == "false",
                _ => false,
            };
            if !admitted {
                return Err(unsupported("list-level-attribute"));
            }
        }
        let kind = if number {
            let format = match attribute((NS_STYLE, "num-format")).unwrap_or("") {
                "1" => Format::Decimal,
                "a" => Format::LowerLetter,
                "A" => Format::UpperLetter,
                "i" => Format::LowerRoman,
                "I" => Format::UpperRoman,
                "①, ②, ③, ..." => Format::Circled,
                "ア, イ, ウ, ..." => Format::Katakana,
                "ｱ, ｲ, ｳ, ..." => Format::HalfwidthKatakana,
                "" => Format::Empty,
                _ => return Err(unsupported("list-format")),
            };
            let affix = |name: &str| -> Result<String, String> {
                let value = attribute((NS_STYLE, name)).unwrap_or("");
                // LibreOffice breaks a label at a line feed, draws U+2011 as
                // a gap, never breaks at U+00AD and interleaves body text into
                // a long label with U+00A0 (measured by the reviewer).
                if value.len() > MAX_AFFIX_BYTES || value.chars().any(special_label_character) {
                    return Err(unsupported("list-affix"));
                }
                Ok(value.to_string())
            };
            let authored_display = attribute((NS_TEXT, "display-levels"))
                .map(|value| {
                    value
                        .parse::<u32>()
                        .ok()
                        .filter(|v| (1..=level).contains(v))
                        .ok_or_else(|| unsupported("list-display-levels"))
                })
                .transpose()?;
            let prefix = affix("num-prefix")?;
            let suffix = affix("num-suffix")?;
            // LibreOffice composes the label from its template when one is
            // written (it may drop `text:display-levels`, measured). The
            // template must be the prefix, the displayed levels' placeholders
            // joined by "." and ending at this level, and the suffix.
            let display = match attribute((NS_LOEXT, "num-list-format")) {
                None => authored_display.unwrap_or(1),
                Some(template) => {
                    let display = (1..=level)
                        .find(|display| {
                            let levels: Vec<String> = (level + 1 - display..=level)
                                .map(|l| format!("%{l}%"))
                                .collect();
                            template == format!("{prefix}{}{suffix}", levels.join("."))
                        })
                        .ok_or_else(|| unsupported("list-format-template"))?;
                    if authored_display.is_some_and(|authored| authored != display) {
                        return Err(unsupported("list-format-template"));
                    }
                    display
                }
            };
            Kind::Number {
                format,
                prefix,
                suffix,
                display,
                letter_sync: attribute((NS_STYLE, "num-letter-sync")) == Some("true"),
            }
        } else {
            // LibreOffice draws only the bullet; its affixes are ignored
            // (measured).
            let mut chars = attribute((NS_TEXT, "bullet-char")).unwrap_or("").chars();
            match (chars.next(), chars.next()) {
                (Some(c), None) if !special_label_character(c) => Kind::Bullet(c),
                _ => return Err(unsupported("list-bullet")),
            }
        };
        let start = match attribute((NS_TEXT, "start-value")) {
            None => 1,
            Some(value) => value
                .parse::<u32>()
                .ok()
                .filter(|v| *v <= MAX_AUTHORED_START)
                .ok_or_else(|| unsupported("list-start-value"))?,
        };
        let mut follow = Follow::Tab;
        let mut align = Align::Start;
        let (mut margin, mut indent, mut tab) = (0.0, 0.0, None);
        let mut label_alignment = false;
        for child in &definition.children {
            match (child.namespace.as_str(), child.name.as_str()) {
                (NS_STYLE, "list-level-properties") => {
                    for (key, value) in &child.attributes {
                        match (key.0.as_str(), key.1.as_str()) {
                            (NS_TEXT, "list-level-position-and-space-mode") => {
                                label_alignment = value == "label-alignment"
                            }
                            (NS_FO, "text-align") => {
                                align = match value.as_str() {
                                    "start" | "left" => Align::Start,
                                    "end" | "right" => Align::End,
                                    "center" => Align::Center,
                                    _ => return Err(unsupported("list-label-align")),
                                }
                            }
                            // The legacy box's parameters, unused in label
                            // alignment mode.
                            (
                                NS_TEXT,
                                "space-before" | "min-label-width" | "min-label-distance",
                            ) => {}
                            _ => return Err(unsupported("list-level-properties")),
                        }
                    }
                    for alignment in &child.children {
                        if (alignment.namespace.as_str(), alignment.name.as_str())
                            != (NS_STYLE, "list-level-label-alignment")
                            || !alignment.children.is_empty()
                        {
                            return Err(unsupported("list-level-properties"));
                        }
                        for (key, value) in &alignment.attributes {
                            match (key.0.as_str(), key.1.as_str()) {
                                (NS_TEXT, "label-followed-by") => {
                                    follow = match value.as_str() {
                                        "listtab" => Follow::Tab,
                                        "space" => Follow::Space,
                                        "nothing" => Follow::Nothing,
                                        _ => return Err(unsupported("list-label-followed-by")),
                                    }
                                }
                                // A stop before the text area is not modelled
                                // (measured by the reviewer).
                                (NS_TEXT, "list-tab-stop-position") => {
                                    let position = length_pt(value)?;
                                    if position < 0.0 {
                                        return Err(unsupported("list-tab-position"));
                                    }
                                    tab = Some(position);
                                }
                                (NS_FO, "margin-left") => margin = length_pt(value)?,
                                (NS_FO, "text-indent") => indent = length_pt(value)?,
                                _ => return Err(unsupported("list-level-properties")),
                            }
                        }
                    }
                }
                (NS_STYLE, "text-properties") => {}
                _ => return Err(unsupported("list-style-child")),
            }
        }
        // The legacy label box (the default mode) needs label widths at
        // admission.
        if !label_alignment {
            return Err(unsupported("list-legacy-mode"));
        }
        let mut styles = Vec::new();
        if let Some(text) = list.level_text.get(&level.to_string()) {
            // Numbered levels' own text properties are ignored by LibreOffice
            // (measured); only a bullet's face is honoured.
            for key in text.groups[TEXT].attributes.keys() {
                let face = matches!(
                    (key.0.as_str(), key.1.as_str()),
                    (
                        NS_STYLE,
                        "font-name"
                            | "font-pitch"
                            | "font-charset"
                            | "font-family-generic"
                            | "font-style-name"
                    ) | (NS_FO, "font-family")
                );
                if number || !face {
                    return Err(unsupported("list-level-text-properties"));
                }
            }
            styles.push(StyleRef::new(text));
        }
        if let Some(name) = attribute((NS_TEXT, "style-name")) {
            let style = self
                .sheet
                .find(Scope::Content, Family::Text, name)
                .ok_or_else(|| invalid("a list level names a missing style"))?;
            styles.extend(self.text_chain(style)?.iter().copied());
        }
        Ok(Level {
            kind,
            start,
            follow,
            align,
            margin,
            indent,
            tab,
            styles,
        })
    }

    /// A list paragraph's indents and tab stops (measured). An automatic
    /// paragraph style, or a named one linked to the list's style
    /// (`style:list-style-name`), replaces the level's M and I; a named style
    /// that is not linked is ignored by LibreOffice. Later paragraphs of an
    /// item sit at M with no first-line offset.
    pub(super) fn list_geometry(
        &mut self,
        list: &ListParagraph<'a>,
        chain: &[StyleRef<'a>],
        properties: &Cascade<'a>,
        paragraph: &mut DocParagraph,
    ) -> Result<(), String> {
        let mut margin = list.margin;
        let mut indent = list.indent;
        for (name, value) in [("margin-left", &mut margin), ("text-indent", &mut indent)] {
            let key = (NS_FO.to_string(), name.to_string());
            let Some(index) = chain.iter().position(|style| {
                style
                    .style()
                    .groups
                    .get(PARAGRAPH)
                    .is_some_and(|group| group.attributes.contains_key(&key))
            }) else {
                continue;
            };
            let setter = chain[index].style();
            // The setter's link is inherited from its parents like any
            // paragraph property (measured by the reviewer).
            let link = (NS_STYLE.to_string(), "list-style-name".to_string());
            let linked = chain[index..]
                .iter()
                .find_map(|style| style.style().attributes.get(&link))
                == Some(&list.style_name);
            if setter.automatic || linked {
                // Not measured on an item's later paragraphs.
                if list.label.is_none() {
                    return Err(unsupported("list-paragraph-indent"));
                }
                *value = length_pt(&setter.groups[PARAGRAPH].attributes[&key])?;
            }
        }
        paragraph.indent_left = margin;
        paragraph.indent_first = if list.label.is_some() { indent } else { 0.0 };
        paragraph.no_hanging_indent_tab = true;
        let Some(label) = &list.label else {
            return Ok(());
        };
        // LibreOffice leaves a label-only line at the start position in a
        // centred or right-aligned paragraph (measured by the reviewer).
        if matches!(paragraph.alignment.as_str(), "center" | "right") {
            return Err(unsupported("list-paragraph-alignment"));
        }
        // LibreOffice clamps a label beyond the text area (measured by the
        // reviewer); the DOCX layout does not. A label before it is drawn in
        // the margin by both.
        let width = self.text_width - paragraph.indent_right;
        if !(margin + indent < width && margin < width) {
            return Err(unsupported("list-geometry"));
        }
        if label.follow == Follow::Tab {
            let interval = properties
                .get(NS_STYLE, "tab-stop-distance")
                .map(length_pt)
                .transpose()?
                .ok_or_else(|| unsupported("list-tab-interval"))?;
            paragraph.tab_stops = self.list_tab_stops(list, margin, interval, width)?;
        }
        Ok(())
    }

    /// The label's leading runs. An empty label adds none (the body starts at
    /// M + I, measured).
    pub(super) fn list_label(
        &mut self,
        label: Label<'a>,
        contexts: &mut Vec<TextContext<'a>>,
        paragraph: &mut DocParagraph,
    ) -> Result<(), String> {
        // An empty labelled paragraph is taller in LibreOffice (an 18 pt line
        // where the label and text give 13.8 pt, measured): out of model.
        let body = paragraph.runs.iter().any(|run| match run {
            DocRun::Text(text) => !text.text.is_empty(),
            _ => true,
        });
        if !body {
            return Err(unsupported("list-empty-label"));
        }
        if label.text.is_empty() {
            return Ok(());
        }
        // A label's space under BalanceSpacesAndIdeographicSpaces is not
        // measured.
        if label.follow == Follow::Space && self.layout_flags.balance {
            return Err(unsupported("list-label-balance"));
        }
        for c in label.text.chars() {
            match table_22(c) {
                Some(ScriptClass::Latin) => {}
                None if is_weak_unmapped(c) => {}
                _ => return Err(unsupported("list-label-script")),
            }
        }
        let context = if label.styles.is_empty() {
            0
        } else {
            if contexts.len() >= MAX_PARAGRAPH_CONTEXTS {
                return Err(unsupported("span-count"));
            }
            contexts.push(TextContext {
                styles: label.styles.into(),
                parent: Some(0),
            });
            (contexts.len() - 1) as u32
        };
        let properties = self.run_properties(contexts, context, ScriptClass::Latin)?;
        // An end- or centre-aligned label that would start before the page
        // edge, or a centred one reaching T, is laid out irregularly by
        // LibreOffice (measured: the label is cut at the page edge and the
        // body moves right; the tab collapses). The label width is bounded by
        // two ems per ASCII character, so only labels that provably stay
        // clear of both are admitted.
        if label.align != Align::Start {
            if !label.text.is_ascii() {
                return Err(unsupported("list-label-extent"));
            }
            let bound = label.text.len() as f64 * 2.0 * properties.size;
            let start = paragraph.indent_left + paragraph.indent_first;
            let before = if label.align == Align::End {
                bound
            } else {
                bound / 2.0
            };
            let reaches_tab = label.align == Align::Center
                && label.follow == Follow::Tab
                && paragraph
                    .tab_stops
                    .first()
                    .is_none_or(|stop| start + bound / 2.0 >= stop.pos);
            // The label is fitted at its full width from M + I before it is
            // shifted, so it must also fit there: a wrapped aligned label
            // breaks differently in LibreOffice (measured by the reviewer).
            let width = self.text_width - paragraph.indent_right;
            if start - before < -self.page_left_margin || reaches_tab || start + bound > width {
                return Err(unsupported("list-label-extent"));
            }
        }
        // LibreOffice extends a label's decoration over its tab gap, which a
        // DOCX tab does not paint.
        if properties.underline.is_some()
            || properties.strike.is_some()
            || properties.background.is_some()
        {
            return Err(unsupported("list-label-decoration"));
        }
        let separator = match label.follow {
            Follow::Tab => Some("\t"),
            Follow::Space => Some(" "),
            Follow::Nothing => None,
        };
        let mut runs = Vec::with_capacity(2);
        for text in std::iter::once(label.text.as_str()).chain(separator) {
            self.charge_construction(properties.retained_estimate() + text.len())?;
            let mut run = properties.text_run();
            run.text = text.to_string();
            run.list_label = Some(label.align.wire());
            runs.push(DocRun::Text(Box::new(run)));
        }
        paragraph.runs.splice(0..0, runs);
        Ok(())
    }

    /// The tab stops a list paragraph needs (§19.836 and LibreOffice's grid,
    /// measured): T, then the grid stops beyond it, from M when
    /// TabsRelativeToIndent is true and from the text area otherwise, up to
    /// one interval past `width`.
    pub(super) fn list_tab_stops(
        &mut self,
        list: &ListParagraph<'a>,
        margin: f64,
        interval: f64,
        width: f64,
    ) -> Result<Vec<TabStop>, String> {
        let relative = self
            .layout_flags
            .tabs_relative
            .ok_or_else(|| unsupported("list-tabs-relative-unset"))?;
        if !interval.is_finite() || interval < MIN_TAB_INTERVAL_PT {
            return Err(unsupported("list-tab-interval"));
        }
        let origin = if relative { margin } else { 0.0 };
        let mut stops = Vec::new();
        let mut push = |pos: f64, reader: &mut Self| -> Result<(), String> {
            if stops.len() >= MAX_TAB_STOPS {
                return Err(unsupported("list-tab-stops"));
            }
            reader.charge_construction(std::mem::size_of::<TabStop>() + 12)?;
            stops.push(TabStop {
                pos,
                alignment: "left".to_string(),
                leader: "none".to_string(),
            });
            Ok(())
        };
        let first = list.tab.unwrap_or(f64::NEG_INFINITY);
        if let Some(tab) = list.tab {
            push(tab, self)?;
        }
        let mut k = ((first.max(origin) - origin) / interval).floor().max(0.0);
        loop {
            let pos = origin + k * interval;
            if pos > width + interval {
                break;
            }
            if pos > first {
                push(pos, self)?;
            }
            k += 1.0;
        }
        Ok(stops)
    }
}

/// Label characters LibreOffice does not draw as ordinary text: a line feed
/// breaks the label, U+2011 is drawn as a gap, U+00AD never breaks it, and
/// U+00A0 lets body text interleave into a long label (measured by the
/// reviewer).
fn special_label_character(c: char) -> bool {
    c.is_control() || matches!(c, '\u{a0}' | '\u{2011}' | '\u{ad}')
}

/// The element children of an item or header, without soft page breaks.
fn block_children<'n, 'i>(
    node: roxmltree::Node<'n, 'i>,
) -> Result<Vec<roxmltree::Node<'n, 'i>>, String> {
    let mut children = Vec::new();
    for child in node.children() {
        if child.is_text() {
            if child.text().is_some_and(|t| !t.trim().is_empty()) {
                return Err(invalid("text in a list item"));
            }
            continue;
        }
        if !child.is_element() || xml::is(child, NS_TEXT, "soft-page-break") {
            continue;
        }
        children.push(child);
    }
    Ok(children)
}

/// Roman numerals, thousands as repeated "M" (LibreOffice: 32767 is 32 "M"
/// and "DCCLXVII", measured by the reviewer).
fn roman(mut value: u32) -> String {
    const TABLE: [(u32, &str); 13] = [
        (1000, "M"),
        (900, "CM"),
        (500, "D"),
        (400, "CD"),
        (100, "C"),
        (90, "XC"),
        (50, "L"),
        (40, "XL"),
        (10, "X"),
        (9, "IX"),
        (5, "V"),
        (4, "IV"),
        (1, "I"),
    ];
    let mut out = String::new();
    for (unit, numeral) in TABLE {
        while value >= unit {
            out.push_str(numeral);
            value -= unit;
        }
    }
    out
}

/// Lower-case letters (ODF 1.4 §19.505): with letter sync every place repeats
/// one letter ("aa, bb"), without it the places count like digits ("aa, ab").
fn letters(value: u32, sync: bool) -> String {
    if sync {
        let letter = (b'a' + ((value - 1) % 26) as u8) as char;
        return std::iter::repeat_n(letter, ((value - 1) / 26 + 1) as usize).collect();
    }
    let mut n = value;
    let mut out = Vec::new();
    while n > 0 {
        n -= 1;
        out.push(b'a' + (n % 26) as u8);
        n /= 26;
    }
    out.reverse();
    String::from_utf8(out).expect("ASCII letters")
}
