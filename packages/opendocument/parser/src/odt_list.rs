//! Lists (ODF Part 3 §5.3), list styles (§16.30-§16.34) and heading
//! numbering (§16.36), measured against LibreOffice 26.8 PDF export.
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
//!
//! The older label-width-and-position mode (§20.428 et al.) places the label
//! at the paragraph indent plus `text:space-before` and the text at least
//! `text:min-label-width` further: it is mapped to the same model with
//! M = indent + space-before + min-label-width, I = -min-label-width and a
//! tab to M.
//!
//! A picture label is the label font's zero-advance metric followed by the
//! picture, both flagged: LibreOffice gives the picture a box 0.5 pt larger
//! on every side, whose top it places at the label font's ascent while it
//! sizes the line and which it then centres in the line (measured,
//! `style:vertical-pos="middle"` relative to the line).

use super::*;
use crate::style::ListStyle;
use docx_model::{AnchorHostMetrics, TabStop};

/// Levels a list style defines (§19.840 `text:level`).
const MAX_LEVELS: u32 = 10;
/// Counting bound (LibreOffice counts to 32768, measured).
const MAX_COUNTED: u32 = 1_000_000;
/// LibreOffice's margin around a picture label, on every side (measured:
/// the box is 1 pt wider and taller than `fo:width` by `fo:height`).
const IMAGE_LABEL_INSET_PT: f64 = 0.5;
/// Generated tab stops per paragraph, and the smallest grid interval.
const MAX_TAB_STOPS: usize = 1024;
const MIN_TAB_INTERVAL_PT: f64 = 0.1 * 72.0 / 2.54;
/// The reserved list style name standing for the outline style.
const OUTLINE: &str = "\u{1}outline";

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

/// A `style:num-format` (§19.500, with LibreOffice's extended formats).
#[derive(Clone, Copy, PartialEq, Debug)]
enum Format {
    Decimal,
    LowerLetter,
    UpperLetter,
    LowerRoman,
    UpperRoman,
    /// Digits of another numeral system (zero digit code point).
    Digits(u32),
    /// Zero-padded decimal ("01, 02, ...").
    Padded(usize),
    /// A cycling alphabet (the letters, one per number; `sync` repeats).
    Alphabet(&'static str),
    /// Circled numbers "①, ②, ③, ...".
    Circled,
    /// Chinese/Japanese counting ("一, 二, 三, ...").
    Ideographic(&'static [&'static str; 10]),
    /// English ordinals and words.
    Ordinal,
    CardinalText,
    OrdinalText,
    /// An empty `style:num-format`: no counter component.
    Empty,
}

const KATAKANA: &str =
    "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン";
const HALFWIDTH_KATAKANA: &str = "ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜｦﾝ";
const IROHA: &str = "イロハニホヘトチリヌルヲワカヨタレソツネナラムウヰノオクヤマケフコエテアサキユメミシヱヒモセス";
const HIRAGANA: &str =
    "あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわをん";
const HIRAGANA_IROHA: &str = "いろはにほへとちりぬるをわかよたれそつねならむうゐのおくやまけふこえてあさきゆめみしゑひもせす";
const HANGUL_SYLLABLE: &str = "가나다라마바사아자차카타파하";
const HANGUL_JAMO: &str = "ㄱㄴㄷㄹㅁㅂㅅㅇㅈㅊㅋㅌㅍㅎ";
const HEAVENLY_STEMS: &str = "甲乙丙丁戊己庚辛壬癸";
const EARTHLY_BRANCHES: &str = "子丑寅卯辰巳午未申酉戌亥";
const GREEK_LOWER: &str = "αβγδεζηθικλμνξοπρστυφχψω";
const GREEK_UPPER: &str = "ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ";
const CYRILLIC_LOWER: &str = "абвгдежзиклмнопрстуфхцчшщэюя";
const CYRILLIC_UPPER: &str = "АБВГДЕЖЗИКЛМНОПРСТУФХЦЧШЩЭЮЯ";
const HEBREW: &str = "אבגדהוזחטיכלמנסעפצקרשת";
const ARABIC_ALPHABET: &str = "أبتثجحخدذرزسشصضطظعغفقكلمنهوي";
const THAI_ALPHABET: &str = "กขฃคฅฆงจฉชซฌญฎฏฐฑฒณดตถทธนบปผฝพฟภมยรลวศษสหฬอฮ";
const CJK_LOWER: [&str; 10] = ["〇", "一", "二", "三", "四", "五", "六", "七", "八", "九"];
const CJK_LEGAL_SIMPLIFIED: [&str; 10] =
    ["零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖"];
const CJK_LEGAL_TRADITIONAL: [&str; 10] =
    ["零", "壹", "貳", "參", "肆", "伍", "陸", "柒", "捌", "玖"];
const KOREAN_HANJA: [&str; 10] = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"];
const KOREAN_HANGUL_DIGITS: [&str; 10] =
    ["영", "일", "이", "삼", "사", "오", "육", "칠", "팔", "구"];

/// `style:num-format` values (ODF's five, LibreOffice's extended formats by
/// their UI names). Unknown formats count in decimal.
fn parse_format(value: &str) -> Format {
    match value {
        "" => Format::Empty,
        "1" => Format::Decimal,
        "a" => Format::LowerLetter,
        "A" => Format::UpperLetter,
        "i" => Format::LowerRoman,
        "I" => Format::UpperRoman,
        "①, ②, ③, ..." => Format::Circled,
        "ア, イ, ウ, ..." => Format::Alphabet(KATAKANA),
        "ｱ, ｲ, ｳ, ..." => Format::Alphabet(HALFWIDTH_KATAKANA),
        "イ, ロ, ハ, ..." => Format::Alphabet(IROHA),
        "あ, い, う, ..." => Format::Alphabet(HIRAGANA),
        "い, ろ, は, ..." => Format::Alphabet(HIRAGANA_IROHA),
        "가, 나, 다, ..." => Format::Alphabet(HANGUL_SYLLABLE),
        "ㄱ, ㄴ, ㄷ, ..." => Format::Alphabet(HANGUL_JAMO),
        "甲, 乙, 丙, ..." => Format::Alphabet(HEAVENLY_STEMS),
        "子, 丑, 寅, ..." => Format::Alphabet(EARTHLY_BRANCHES),
        "α, β, γ, ..." => Format::Alphabet(GREEK_LOWER),
        "Α, Β, Γ, ..." => Format::Alphabet(GREEK_UPPER),
        "а, б, .., аа, аб, ... (Russian)" | "а, б, .., аа, бб, ... (Russian)" => {
            Format::Alphabet(CYRILLIC_LOWER)
        }
        "А, Б, .., Аа, Аб, ... (Russian)" | "А, Б, .., Аа, Бб, ... (Russian)" => {
            Format::Alphabet(CYRILLIC_UPPER)
        }
        "א, ב, ג, ..." => Format::Alphabet(HEBREW),
        "أ, ب, ت, ..." => Format::Alphabet(ARABIC_ALPHABET),
        "ก, ข, ฃ, ..." => Format::Alphabet(THAI_ALPHABET),
        "一, 二, 三, ..." | "一, 二, 三, ... (zh-CN)" | "一, 二, 三, ... (ja)" => {
            Format::Ideographic(&CJK_LOWER)
        }
        "壹, 贰, 叁, ..." => Format::Ideographic(&CJK_LEGAL_SIMPLIFIED),
        "壹, 貳, 參, ..." => Format::Ideographic(&CJK_LEGAL_TRADITIONAL),
        "일, 이, 삼, ..." => Format::Ideographic(&KOREAN_HANGUL_DIGITS),
        "一, 二, 三, ... (ko)" => Format::Ideographic(&KOREAN_HANJA),
        "١, ٢, ٣, ٤, ..." => Format::Digits(0x0660),
        "۱, ۲, ۳, ۴, ..." => Format::Digits(0x06F0),
        "१, २, ३, ..." => Format::Digits(0x0966),
        "১, ২, ৩, ..." => Format::Digits(0x09E6),
        "๑, ๒, ๓, ..." => Format::Digits(0x0E50),
        "１, ２, ３, ..." => Format::Digits(0xFF10),
        "01, 02, 03, ..." => Format::Padded(2),
        "001, 002, 003, ..." => Format::Padded(3),
        "0001, 0002, 0003, ..." => Format::Padded(4),
        "00001, 00002, 00003, ..." => Format::Padded(5),
        "1st" => Format::Ordinal,
        "One" => Format::CardinalText,
        "First" => Format::OrdinalText,
        _ => Format::Decimal,
    }
}

/// A note number format (§16.29.2 `style:num-format`) as ST_NumberFormat.
pub(super) fn note_number_format(value: &str) -> Option<&'static str> {
    match value {
        "1" => Some("decimal"),
        "a" => Some("lowerLetter"),
        "A" => Some("upperLetter"),
        "i" => Some("lowerRoman"),
        "I" => Some("upperRoman"),
        "①, ②, ③, ..." => Some("decimalEnclosedCircle"),
        _ => None,
    }
}

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
    /// A picture (§16.33), drawn at `fo:width` by `fo:height` (pt).
    Image {
        /// Shared by every copy of the cached level.
        href: Rc<str>,
        width: Option<f64>,
        height: Option<f64>,
    },
}

/// One level of a list style.
#[derive(Clone)]
pub(super) struct Level<'a> {
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
    /// Bullet size relative to the text (§19.762 text:bullet-relative-size).
    relative_size: Option<f64>,
}

/// A list paragraph's label and geometry, handed to `Reader::paragraph`.
pub(super) struct ListParagraph<'a> {
    pub margin: f64,
    pub indent: f64,
    pub tab: Option<f64>,
    pub label: Option<Label<'a>>,
}

pub(super) struct Label<'a> {
    pub text: String,
    /// A picture label's link and size (pt); `text` is then empty.
    pub image: Option<(String, Option<f64>, Option<f64>)>,
    pub follow: Follow,
    pub align: Align,
    pub styles: Vec<StyleRef<'a>>,
    pub relative_size: Option<f64>,
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
    /// `xml:id` of lists, to their chains.
    ids: HashMap<String, usize>,
    /// The immediately preceding top-level list: its style and chain.
    previous: Option<(String, usize)>,
    /// Chains by list style for `text:numbered-paragraph` (§5.3.6, by
    /// `text:list-id`).
    numbered: HashMap<String, usize>,
    outline: Option<usize>,
}

/// The list a paragraph sits in while it is read.
#[derive(Clone)]
pub(super) struct Frame {
    chain: usize,
    style: String,
    level: u32,
}

/// A number a list counter starts at (§19.868 `text:start-value`): beyond
/// the counting bound it is refused, never clipped.
fn list_number(value: Option<&str>) -> Result<Option<u32>, String> {
    Ok(
        xml::bounded_count(value, u64::from(MAX_COUNTED), "list number")?
            .map(|number| number as u32),
    )
}

impl<'a> Reader<'a> {
    fn list_style(&self, name: &str) -> Option<&'a ListStyle> {
        let sheet: &'a StyleSheet = self.sheet;
        if name == OUTLINE {
            return sheet.outline();
        }
        sheet
            .find_list(self.scope, name)
            .or_else(|| sheet.find_list(Scope::Content, name))
            .or_else(|| sheet.find_list(Scope::Styles, name))
    }

    /// A `text:list` (§5.3.1). Every item's first paragraph is a numbered
    /// node at its nesting depth; wrapper items (whose first child is a list)
    /// and headers are not. A continuation chain is one LibreOffice list.
    pub(super) fn list(
        &mut self,
        node: roxmltree::Node<'a, 'a>,
        parent: Option<&Frame>,
    ) -> Result<(), String> {
        self.enter()?;
        let style_name = node.attribute((NS_TEXT, "style-name"));
        let continue_numbering = node.attribute((NS_TEXT, "continue-numbering")) == Some("true");
        let continue_list = node.attribute((NS_TEXT, "continue-list"));
        let id = node.attribute((xml::NS_XML, "id"));
        let frame = match parent {
            Some(parent) => Frame {
                chain: parent.chain,
                // A nested list may name its own style (§19.880): its levels
                // come from that style.
                style: style_name
                    .filter(|name| self.list_style(name).is_some())
                    .unwrap_or(&parent.style)
                    .to_string(),
                level: (parent.level + 1).min(MAX_LEVELS),
            },
            None => {
                let style = style_name.unwrap_or("").to_string();
                let continued = if let Some(target) = continue_list {
                    self.lists.ids.get(target).copied()
                } else {
                    match &self.lists.previous {
                        Some((previous, chain)) if continue_numbering && *previous == style => {
                            Some(*chain)
                        }
                        _ => None,
                    }
                };
                let chain = match continued {
                    Some(chain) => chain,
                    None => {
                        budget::tick(1)?;
                        self.lists.chains.push(Chain::default());
                        self.lists.chains.len() - 1
                    }
                };
                self.lists.previous = Some((style.clone(), chain));
                Frame {
                    chain,
                    style,
                    level: 1,
                }
            }
        };
        if let Some(id) = id {
            self.lists.ids.insert(id.to_string(), frame.chain);
        }
        let level = self.list_level(&frame.style, frame.level)?;
        for child in node.kids().filter(|n| n.is_element()) {
            budget::tick(1)?;
            match (child.tag_name().namespace(), child.tag_name().name()) {
                (Some(NS_TEXT), "list-item") => self.list_item(child, &frame, &level, false)?,
                (Some(NS_TEXT), "list-header") => self.list_item(child, &frame, &level, true)?,
                _ => {}
            }
        }
        self.leave();
        Ok(())
    }

    fn list_item(
        &mut self,
        node: roxmltree::Node<'a, 'a>,
        frame: &Frame,
        level: &Level<'a>,
        header: bool,
    ) -> Result<(), String> {
        let start = list_number(node.attribute((NS_TEXT, "start-value")))?;
        let mut first = !header;
        for child in node.kids().filter(|n| n.is_element()) {
            budget::tick(1)?;
            match (child.tag_name().namespace(), child.tag_name().name()) {
                (Some(NS_TEXT), "p" | "h") if first => {
                    first = false;
                    let values = self.count(frame, start)?;
                    let label = self.label(frame, level, &values)?;
                    self.list_paragraph = Some(ListParagraph {
                        margin: level.margin,
                        indent: level.indent,
                        tab: level.tab,
                        label: Some(label),
                    });
                    self.list_child(child)?;
                }
                (Some(NS_TEXT), "p" | "h") => {
                    self.list_paragraph = Some(ListParagraph {
                        margin: level.margin,
                        indent: 0.0,
                        tab: level.tab,
                        label: None,
                    });
                    self.list_child(child)?;
                }
                (Some(NS_TEXT), "list") => {
                    first = false;
                    self.list(child, Some(frame))?;
                }
                (Some(NS_TEXT), "soft-page-break") => {}
                _ => {
                    first = false;
                    self.block(child)?;
                }
            }
        }
        Ok(())
    }

    fn list_child(&mut self, node: roxmltree::Node<'a, 'a>) -> Result<(), String> {
        if xml::is(node, NS_TEXT, "h") {
            self.paragraph(node, Some(outline_level(node)))
        } else {
            self.paragraph(node, None)
        }
    }

    /// `text:numbered-paragraph` (§5.3.6): numbered in the chain of its
    /// `text:list-id` at its `text:level`.
    pub(super) fn numbered(
        &mut self,
        node: roxmltree::Node<'a, 'a>,
        paragraph: roxmltree::Node<'a, 'a>,
    ) -> Result<(), String> {
        let style = node
            .attribute((NS_TEXT, "style-name"))
            .unwrap_or("")
            .to_string();
        let list_id = node
            .attribute((NS_TEXT, "list-id"))
            .unwrap_or(&style)
            .to_string();
        let level = node
            .attribute((NS_TEXT, "level"))
            .and_then(unsigned)
            .map_or(1, |level| level.clamp(1, u64::from(MAX_LEVELS)) as u32);
        let chain = match self.lists.numbered.get(&list_id) {
            Some(chain) => *chain,
            None => {
                self.lists.chains.push(Chain::default());
                let chain = self.lists.chains.len() - 1;
                self.lists.numbered.insert(list_id, chain);
                chain
            }
        };
        let frame = Frame {
            chain,
            style,
            level,
        };
        let definition = self.list_level(&frame.style, level)?;
        let start = list_number(node.attribute((NS_TEXT, "start-value")))?;
        let label = if node.attribute((NS_TEXT, "is-list-header")) == Some("true") {
            None
        } else {
            let values = self.count(&frame, start)?;
            Some(self.label(&frame, &definition, &values)?)
        };
        self.list_paragraph = Some(ListParagraph {
            margin: definition.margin,
            indent: if label.is_some() {
                definition.indent
            } else {
                0.0
            },
            tab: definition.tab,
            label,
        });
        self.list_child(paragraph)
    }

    /// A heading numbered by the list style its paragraph style names, at
    /// its outline level, in one list per list style.
    pub(super) fn styled_list_paragraph(
        &mut self,
        node: roxmltree::Node<'a, 'a>,
        style: &str,
        level: u32,
    ) -> Result<Option<ListParagraph<'a>>, String> {
        let level = level.clamp(1, MAX_LEVELS);
        let key = format!("\u{2}{style}");
        let chain = match self.lists.numbered.get(&key) {
            Some(chain) => *chain,
            None => {
                self.lists.chains.push(Chain::default());
                let chain = self.lists.chains.len() - 1;
                self.lists.numbered.insert(key, chain);
                chain
            }
        };
        let frame = Frame {
            chain,
            style: style.to_string(),
            level,
        };
        let definition = self.list_level(style, level)?;
        let start = list_number(node.attribute((NS_TEXT, "start-value")))?;
        let values = self.count(&frame, start)?;
        let label = self.label(&frame, &definition, &values)?;
        Ok(Some(ListParagraph {
            margin: definition.margin,
            indent: definition.indent,
            tab: definition.tab,
            label: Some(label),
        }))
    }

    /// Heading numbering (§16.36): the outline style's level counted over
    /// the document's headings. A level without a label leaves the heading
    /// unnumbered and its geometry to its paragraph style.
    pub(super) fn outline_paragraph(
        &mut self,
        node: roxmltree::Node<'a, 'a>,
        level: u32,
    ) -> Result<Option<ListParagraph<'a>>, String> {
        if self.sheet.outline().is_none() {
            return Ok(None);
        }
        let level = level.clamp(1, MAX_LEVELS);
        let definition = self.list_level(OUTLINE, level)?;
        let numbered = match &definition.kind {
            Kind::Number {
                format,
                prefix,
                suffix,
                ..
            } => *format != Format::Empty || !prefix.is_empty() || !suffix.is_empty(),
            Kind::Bullet(_) | Kind::Image { .. } => true,
        };
        if !numbered {
            return Ok(None);
        }
        let chain = match self.lists.outline {
            Some(chain) => chain,
            None => {
                self.lists.chains.push(Chain::default());
                let chain = self.lists.chains.len() - 1;
                self.lists.outline = Some(chain);
                chain
            }
        };
        let frame = Frame {
            chain,
            style: OUTLINE.to_string(),
            level,
        };
        if node.attribute((NS_TEXT, "restart-numbering")) == Some("true") {
            self.lists.chains[chain].counters = Default::default();
        }
        let start = list_number(node.attribute((NS_TEXT, "start-value")))?;
        let values = self.count(&frame, start)?;
        let label = self.label(&frame, &definition, &values)?;
        Ok(Some(ListParagraph {
            margin: definition.margin,
            indent: definition.indent,
            tab: definition.tab,
            label: Some(label),
        }))
    }

    /// LibreOffice's numbering of a node at `frame.level` (measured): every
    /// unset ancestor level takes its level's start (a phantom parent), the
    /// node's level takes the item's start, else counts on, else takes its
    /// start, and deeper levels are unset.
    fn count(&mut self, frame: &Frame, start: Option<u32>) -> Result<Vec<u32>, String> {
        let mut starts = Vec::with_capacity(frame.level as usize);
        for level in 1..=frame.level {
            starts.push(self.list_level(&frame.style, level)?.start);
        }
        let chain = &mut self.lists.chains[frame.chain];
        let index = (frame.level - 1) as usize;
        // Every value entering a counter is within MAX_COUNTED (start values
        // are read with that bound), so no label component outgrows it; a
        // counter that counts past it is beyond the reader's limits.
        for (ancestor, ancestor_start) in starts.iter().enumerate().take(index) {
            chain.counters[ancestor].get_or_insert(*ancestor_start);
        }
        let value = match (start, chain.counters[index]) {
            (Some(start), _) => start,
            (None, Some(previous)) => previous + 1,
            (None, None) => starts[index],
        };
        if value > MAX_COUNTED {
            return Err(resource("list number"));
        }
        chain.counters[index] = Some(value);
        for deeper in chain.counters.iter_mut().skip(index + 1) {
            *deeper = None;
        }
        chain.counted = true;
        Ok(chain.counters[..=index]
            .iter()
            .map(|v| v.unwrap_or(1))
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
            Kind::Image { .. } => String::new(),
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
                                Kind::Bullet(_) | Kind::Image { .. } => {
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
        let image = match &level.kind {
            Kind::Image {
                href,
                width,
                height,
            } => Some((href.to_string(), *width, *height)),
            _ => None,
        };
        Ok(Label {
            text,
            image,
            follow: level.follow,
            align: level.align,
            styles: level.styles.clone(),
            relative_size: level.relative_size,
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
            Format::LowerRoman => roman(value)?.to_ascii_lowercase(),
            Format::UpperRoman => roman(value)?,
            Format::Circled => match value {
                1..=20 => char::from_u32(0x2460 + value - 1),
                21..=35 => char::from_u32(0x3251 + value - 21),
                36..=50 => char::from_u32(0x32B1 + value - 36),
                _ => None,
            }
            .map_or_else(|| value.to_string(), |c| c.to_string()),
            Format::Alphabet(table) => {
                let letters: Vec<char> = table.chars().collect();
                let n = letters.len() as u32;
                let letter = letters[((value - 1) % n) as usize];
                let repeat = if sync { (value - 1) / n + 1 } else { 1 };
                repeated(letter, repeat as usize)?
            }
            Format::Digits(zero) => value
                .to_string()
                .chars()
                .map(|d| char::from_u32(zero + d.to_digit(10).unwrap_or(0)).unwrap_or(d))
                .collect(),
            Format::Padded(width) => format!("{value:0width$}"),
            Format::Ideographic(digits) => ideographic(value, digits),
            Format::Ordinal => {
                let suffix = match (value % 10, value % 100) {
                    (_, 11..=13) => "th",
                    (1, _) => "st",
                    (2, _) => "nd",
                    (3, _) => "rd",
                    _ => "th",
                };
                format!("{value}{suffix}")
            }
            Format::CardinalText => english(value, false),
            Format::OrdinalText => english(value, true),
            Format::LowerLetter | Format::UpperLetter => {
                // Beyond "z" the places count like digits ("aa, ab"), as
                // ODF 1.4 §19.505 and LibreOffice do.
                let letters = letters(value, sync)?;
                if format == Format::UpperLetter {
                    letters.to_ascii_uppercase()
                } else {
                    letters
                }
            }
        })
    }

    /// A list style's level definition (§16.31-§16.34), resolved once per
    /// scope, style and level; later uses copy the bounded definition.
    fn list_level(&mut self, style: &str, level: u32) -> Result<Level<'a>, String> {
        budget::tick(1)?;
        let key = (self.scope, style.to_string(), level);
        if let Some(cached) = self.level_cache.get(&key) {
            return Ok(cached.clone());
        }
        let resolved = self.resolve_list_level(style, level)?;
        self.level_cache.insert(key, resolved.clone());
        Ok(resolved)
    }

    fn resolve_list_level(&mut self, style: &str, level: u32) -> Result<Level<'a>, String> {
        let fallback = Level {
            kind: Kind::Number {
                format: Format::Empty,
                prefix: String::new(),
                suffix: String::new(),
                display: 1,
                letter_sync: false,
            },
            start: 1,
            follow: Follow::Nothing,
            align: Align::Start,
            margin: 0.0,
            indent: 0.0,
            tab: None,
            styles: Vec::new(),
            relative_size: None,
        };
        let Some(list) = self.list_style(style) else {
            return Ok(fallback);
        };
        // The level search and its format templates read the whole
        // captured definition.
        budget::tick(list.element.nodes())?;
        // The definition of this level, else the nearest lower one (a list
        // nested deeper than its style defines).
        let level_of = |child: &crate::style::Element| {
            child
                .attributes
                .get(&(NS_TEXT.to_string(), "level".to_string()))
                .and_then(|value| value.parse::<u32>().ok())
        };
        let definition = list
            .element
            .children
            .iter()
            .rfind(|child| level_of(child) == Some(level))
            .or_else(|| {
                list.element
                    .children
                    .iter()
                    .filter(|child| level_of(child).is_some_and(|l| l <= level))
                    .max_by_key(|child| level_of(child))
            });
        let Some(definition) = definition else {
            return Ok(fallback);
        };
        let attribute = |name: (&str, &str)| {
            definition
                .attributes
                .get(&(name.0.to_string(), name.1.to_string()))
                .map(String::as_str)
        };
        let (number, image) = match definition.name.as_str() {
            "list-level-style-number" | "outline-level-style" => (true, false),
            "list-level-style-bullet" => (false, false),
            "list-level-style-image" => (false, true),
            _ => (true, false),
        };
        // (A list level's values are bounded by `style::check_values`.)
        let affix =
            |name: &str| -> String { attribute((NS_STYLE, name)).unwrap_or("").to_string() };
        let kind = if number {
            let format = parse_format(attribute((NS_STYLE, "num-format")).unwrap_or(""));
            let prefix = affix("num-prefix");
            let suffix = affix("num-suffix");
            let authored_display = attribute((NS_TEXT, "display-levels"))
                .and_then(unsigned)
                .map(|v| v.clamp(1, u64::from(level)) as u32);
            // LibreOffice composes the label from its template when one is
            // written (it may drop `text:display-levels`, measured): the
            // number of level placeholders it holds.
            let display = match attribute((NS_LOEXT, "num-list-format")) {
                Some(template) => {
                    let placeholders = (1..=level)
                        .filter(|l| template.contains(&format!("%{l}%")))
                        .count() as u32;
                    authored_display.unwrap_or(placeholders.max(1))
                }
                None => authored_display.unwrap_or(1),
            };
            Kind::Number {
                format,
                prefix,
                suffix,
                display,
                letter_sync: attribute((NS_STYLE, "num-letter-sync")) == Some("true"),
            }
        } else if image {
            match attribute((NS_XLINK, "href")) {
                Some(href) => Kind::Image {
                    href: Rc::from(href),
                    width: None,
                    height: None,
                },
                None => Kind::Bullet('\u{2022}'),
            }
        } else {
            // LibreOffice draws only the bullet; its affixes are ignored
            // (measured).
            Kind::Bullet(
                attribute((NS_TEXT, "bullet-char"))
                    .and_then(|v| v.chars().next())
                    .unwrap_or('\u{2022}'),
            )
        };
        let start = list_number(attribute((NS_TEXT, "start-value")))?.unwrap_or(1);
        let relative_size = attribute((NS_TEXT, "bullet-relative-size"))
            .and_then(|v| percent(v).ok())
            .filter(|v| *v > 0.0 && (*v - 1.0).abs() > 1e-9);
        let mut follow = Follow::Tab;
        let mut align = Align::Start;
        let (mut margin, mut indent, mut tab) = (0.0, 0.0, None);
        let mut label_alignment = false;
        let (mut space_before, mut min_width, mut min_distance) = (0.0, 0.0, 0.0);
        let (mut width, mut height) = (None, None);
        for child in &definition.children {
            if child.name != "list-level-properties" {
                continue;
            }
            for (key, value) in &child.attributes {
                match (key.0.as_str(), key.1.as_str()) {
                    (NS_FO, "width") => width = length_pt(value).ok(),
                    (NS_FO, "height") => height = length_pt(value).ok(),
                    (NS_TEXT, "list-level-position-and-space-mode") => {
                        label_alignment = value == "label-alignment"
                    }
                    (NS_FO, "text-align") => {
                        align = match value.as_str() {
                            "end" | "right" => Align::End,
                            "center" => Align::Center,
                            _ => Align::Start,
                        }
                    }
                    (NS_TEXT, "space-before") => space_before = length_pt(value).unwrap_or(0.0),
                    (NS_TEXT, "min-label-width") => min_width = length_pt(value).unwrap_or(0.0),
                    (NS_TEXT, "min-label-distance") => {
                        min_distance = length_pt(value).unwrap_or(0.0)
                    }
                    _ => {}
                }
            }
            for alignment in &child.children {
                if alignment.name != "list-level-label-alignment" {
                    continue;
                }
                for (key, value) in &alignment.attributes {
                    match (key.0.as_str(), key.1.as_str()) {
                        (NS_TEXT, "label-followed-by") => {
                            follow = match value.as_str() {
                                "space" => Follow::Space,
                                "nothing" => Follow::Nothing,
                                _ => Follow::Tab,
                            }
                        }
                        (NS_TEXT, "list-tab-stop-position") => tab = length_pt(value).ok(),
                        (NS_FO, "margin-left") => margin = length_pt(value).unwrap_or(0.0),
                        (NS_FO, "text-indent") => indent = length_pt(value).unwrap_or(0.0),
                        _ => {}
                    }
                }
            }
        }
        if !label_alignment {
            // The label-width-and-position mode (see the module notes).
            margin = space_before + min_width;
            indent = -min_width;
            tab = Some(margin);
            follow = if min_width > 0.0 || min_distance > 0.0 {
                Follow::Tab
            } else {
                Follow::Nothing
            };
        }
        let kind = match kind {
            Kind::Image { href, .. } => Kind::Image {
                href,
                width: width.filter(|w| *w > 0.0),
                height: height.filter(|h| *h > 0.0),
            },
            kind => kind,
        };
        let mut styles = Vec::new();
        if let Some(text) = list.level_text.get(&level.to_string()) {
            styles.push(StyleRef::new(text));
        }
        if let Some(name) = attribute((NS_TEXT, "style-name")) {
            let sheet: &'a StyleSheet = self.sheet;
            if let Some(style) = sheet
                .find(self.scope, Family::Text, name)
                .or_else(|| sheet.find_common(Family::Text, name))
            {
                styles.extend(self.style_chain(style)?.iter().copied());
            }
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
            relative_size,
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
        let automatic_indent = properties.get(NS_STYLE, "auto-text-indent") == Some("true");
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
            let link = (NS_STYLE.to_string(), "list-style-name".to_string());
            let nearest = chain
                .iter()
                .position(|style| style.style().attributes.contains_key(&link));
            if let Some(nearest) = nearest {
                if nearest < index {
                    continue;
                }
            }
            let linked = nearest.is_some();
            if setter.automatic || linked {
                if list.label.is_none() && name == "text-indent" {
                    continue;
                }
                // (§20.256: an automatic indent leaves fo:text-indent
                // unused.)
                if name == "text-indent" && automatic_indent {
                    continue;
                }
                let raw = &setter.groups[PARAGRAPH].attributes[&key];
                if raw.ends_with('%') {
                    // A percentage of the parent style's value (§20.207,
                    // §20.225).
                    *value = relative_length(&chain[index..], PARAGRAPH, NS_FO, name, raw);
                } else if let Ok(length) = length_pt(raw) {
                    *value = length;
                }
            }
        }
        paragraph.indent_left = margin + self.section_indent.0;
        // (An automatic indent, already resolved from the font size,
        // stays.)
        if !automatic_indent {
            paragraph.indent_first = if list.label.is_some() { indent } else { 0.0 };
        }
        paragraph.no_hanging_indent_tab = true;
        let Some(label) = &list.label else {
            return Ok(());
        };
        if label.follow == Follow::Tab {
            // (The cascade ends in the paragraph default style.)
            let interval = properties
                .get(NS_STYLE, "tab-stop-distance")
                .and_then(|v| length_pt(v).ok())
                .unwrap_or(36.0);
            let width = self.text_width - paragraph.indent_right;
            let mut stops = self.list_tab_stops(list, margin, interval, width)?;
            // The paragraph's own custom stops after the list's.
            let custom = self.tab_stops(properties, paragraph.indent_left)?;
            stops.extend(custom);
            stops.sort_by(|a, b| a.pos.total_cmp(&b.pos));
            paragraph.tab_stops = stops;
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
        if label.image.is_some() {
            return self.list_label_image(label, contexts, paragraph);
        }
        if label.text.is_empty() {
            return Ok(());
        }
        let context = if label.styles.is_empty() || contexts.len() >= MAX_PARAGRAPH_CONTEXTS {
            0
        } else {
            contexts.push(TextContext {
                styles: label.styles.into(),
                parent: Some(0),
                link: None,
            });
            (contexts.len() - 1) as u32
        };
        // The label's script class: that of its first classified character.
        let class = label
            .text
            .chars()
            .find_map(crate::script::table_22)
            .unwrap_or(ScriptClass::Latin);
        let mut properties = self.run_properties(contexts, context, class)?;
        if let Some(relative) = label.relative_size {
            properties.size *= relative;
        }
        // LibreOffice paints a label without the paragraph text's
        // decorations (they belong to the text, not the label).
        let separator = match label.follow {
            Follow::Tab => Some("\t"),
            Follow::Space => Some(" "),
            Follow::Nothing => None,
        };
        let mut runs = Vec::with_capacity(2);
        for text in std::iter::once(label.text.as_str()).chain(separator) {
            budget::tick(1)?;
            let mut run = properties.text_run();
            run.text = text.to_string();
            run.list_label = Some(label.align.wire());
            runs.push(DocRun::Text(Box::new(run)));
        }
        paragraph.runs.splice(0..0, runs);
        Ok(())
    }

    /// A picture label's leading runs: the label font's zero-advance metric,
    /// the picture and the separator, all flagged. The picture's box (0.5 pt
    /// more on every side) advances the pen.
    fn list_label_image(
        &mut self,
        label: Label<'a>,
        contexts: &mut [TextContext<'a>],
        paragraph: &mut DocParagraph,
    ) -> Result<(), String> {
        let (href, width, height) = label.image.expect("a picture label");
        // The picture's line geometry follows the formatting of the
        // paragraph's first character (measured).
        // Only the first run's properties are read: its text is set aside
        // while the properties are copied (a charged copy, no text).
        let first = match paragraph.runs.iter_mut().find_map(|run| match run {
            DocRun::Text(text) if !text.text.is_empty() => Some(text),
            _ => None,
        }) {
            Some(text) => {
                let kept = std::mem::take(&mut text.text);
                let first = text.as_ref().clone();
                text.text = kept;
                first
            }
            None => self
                .run_properties(contexts, 0, ScriptClass::Latin)?
                .text_run(),
        };
        let Some((path, mime_type)) = self.package_image(&href)? else {
            return Ok(());
        };
        let width = width.unwrap_or(first.font_size * 0.8);
        let height = height.unwrap_or(first.font_size * 0.8);
        let align = label.align.wire();
        let family = first.font_family.clone();
        let _ = IMAGE_LABEL_INSET_PT;
        let mut runs = vec![
            DocRun::AnchorHost(AnchorHostMetrics {
                font_size: first.font_size,
                font_family: family,
                font_family_east_asia: None,
                bold: first.bold,
                italic: first.italic,
                anchor_occurrence_id: None,
                list_label: Some(align),
            }),
            DocRun::Image(Box::new(draw::inline_image(
                path,
                mime_type,
                width,
                height,
                Some(align),
            ))),
        ];
        let separator = match label.follow {
            Follow::Tab => Some("\t"),
            Follow::Space => Some(" "),
            Follow::Nothing => None,
        };
        if let Some(separator) = separator {
            let mut run = first;
            run.text = separator.to_string();
            run.list_label = Some(align);
            runs.push(DocRun::Text(Box::new(run)));
        }
        paragraph.runs.splice(0..0, runs);
        Ok(())
    }

    /// The tab stops a list paragraph needs (§19.836 and LibreOffice's grid,
    /// measured): T, then the grid stops beyond it, from M when
    /// TabsRelativeToIndent is true and from the text area otherwise, up to
    /// one interval past `width`.
    fn list_tab_stops(
        &mut self,
        list: &ListParagraph<'a>,
        margin: f64,
        interval: f64,
        width: f64,
    ) -> Result<Vec<TabStop>, String> {
        let relative = self.settings.tabs_relative.unwrap_or(true);
        let interval = if interval.is_finite() && interval >= MIN_TAB_INTERVAL_PT {
            interval
        } else {
            36.0
        };
        let origin = if relative { margin } else { 0.0 };
        let mut stops = Vec::new();
        let first = list.tab.unwrap_or(f64::NEG_INFINITY);
        if let Some(tab) = list.tab {
            stops.push(TabStop {
                pos: tab,
                alignment: "left".to_string(),
                leader: "none".to_string(),
            });
        }
        let mut k = ((first.max(origin) - origin) / interval).floor().max(0.0);
        loop {
            budget::tick(1)?;
            let pos = origin + k * interval;
            if pos.is_nan() || pos > width + interval || stops.len() >= MAX_TAB_STOPS {
                break;
            }
            if pos > first {
                stops.push(TabStop {
                    pos,
                    alignment: "left".to_string(),
                    leader: "none".to_string(),
                });
            }
            k += 1.0;
        }
        Ok(stops)
    }
}

/// Roman numerals, thousands as repeated "M" (LibreOffice: 32767 is 32 "M"
/// and "DCCLXVII", measured by the reviewer).
fn roman(mut value: u32) -> Result<String, String> {
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
    // The thousands are a run of the count's length; the rest is at most
    // fifteen numerals.
    let mut out = repeated('M', (value / 1000) as usize)?;
    value %= 1000;
    for (unit, numeral) in TABLE {
        while value >= unit {
            out.push_str(numeral);
            value -= unit;
        }
    }
    Ok(out)
}

/// `count` times `c`, written in pieces with a step each: a run as long as
/// a number the file chose is never allocated at once.
fn repeated(c: char, count: usize) -> Result<String, String> {
    const PIECE: usize = 4096;
    let mut out = String::new();
    let mut left = count;
    while left > 0 {
        let piece = left.min(PIECE);
        budget::tick(1)?;
        out.extend(std::iter::repeat_n(c, piece));
        left -= piece;
    }
    Ok(out)
}

/// Lower-case letters (ODF 1.4 §19.505): with letter sync every place repeats
/// one letter ("aa, bb"), without it the places count like digits ("aa, ab").
fn letters(value: u32, sync: bool) -> Result<String, String> {
    if sync {
        let letter = (b'a' + ((value - 1) % 26) as u8) as char;
        return repeated(letter, ((value - 1) / 26 + 1) as usize);
    }
    let mut n = value;
    let mut out = Vec::new();
    while n > 0 {
        n -= 1;
        out.push(b'a' + (n % 26) as u8);
        n /= 26;
    }
    out.reverse();
    Ok(String::from_utf8(out).expect("ASCII letters"))
}

/// Chinese-style counting with 十, 百 and 千 (the 〇 digit set shows zero
/// places as in "一百〇五").
fn ideographic(value: u32, digits: &[&str; 10]) -> String {
    if value >= 10_000 {
        return value
            .to_string()
            .chars()
            .map(|d| digits[d.to_digit(10).unwrap_or(0) as usize])
            .collect();
    }
    let units = ["", "十", "百", "千"];
    let places: Vec<u32> = value
        .to_string()
        .chars()
        .map(|d| d.to_digit(10).unwrap_or(0))
        .collect();
    let count = places.len();
    let mut out = String::new();
    let mut zero_pending = false;
    for (i, digit) in places.iter().enumerate() {
        let place = count - 1 - i;
        if *digit == 0 {
            zero_pending = !out.is_empty();
            continue;
        }
        if zero_pending {
            out.push_str(digits[0]);
            zero_pending = false;
        }
        // "十" alone for 10-19 at the head.
        if !(*digit == 1 && place == 1 && i == 0) {
            out.push_str(digits[*digit as usize]);
        }
        out.push_str(units[place]);
    }
    out
}

/// English cardinal or ordinal words (LibreOffice "One", "First").
fn english(value: u32, ordinal: bool) -> String {
    const ONES: [&str; 20] = [
        "zero",
        "one",
        "two",
        "three",
        "four",
        "five",
        "six",
        "seven",
        "eight",
        "nine",
        "ten",
        "eleven",
        "twelve",
        "thirteen",
        "fourteen",
        "fifteen",
        "sixteen",
        "seventeen",
        "eighteen",
        "nineteen",
    ];
    const TENS: [&str; 10] = [
        "", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety",
    ];
    fn words(n: u32) -> String {
        match n {
            0..=19 => ONES[n as usize].to_string(),
            20..=99 => {
                let tens = TENS[(n / 10) as usize];
                if n.is_multiple_of(10) {
                    tens.to_string()
                } else {
                    format!("{tens}-{}", ONES[(n % 10) as usize])
                }
            }
            100..=999 => {
                let rest = n % 100;
                let head = format!("{} hundred", ONES[(n / 100) as usize]);
                if rest == 0 {
                    head
                } else {
                    format!("{head} {}", words(rest))
                }
            }
            _ => {
                let rest = n % 1000;
                let head = format!("{} thousand", words(n / 1000));
                if rest == 0 {
                    head
                } else {
                    format!("{head} {}", words(rest))
                }
            }
        }
    }
    let mut text = words(value);
    if ordinal {
        const IRREGULAR: [(&str, &str); 8] = [
            ("one", "first"),
            ("two", "second"),
            ("three", "third"),
            ("five", "fifth"),
            ("eight", "eighth"),
            ("nine", "ninth"),
            ("twelve", "twelfth"),
            ("y", "ieth"),
        ];
        let mut done = false;
        for (from, to) in IRREGULAR {
            if text.ends_with(from) {
                text.truncate(text.len() - from.len());
                text.push_str(to);
                done = true;
                break;
            }
        }
        if !done {
            text.push_str("th");
        }
    }
    let mut chars = text.chars();
    match chars.next() {
        Some(first) => first.to_uppercase().chain(chars).collect(),
        None => text,
    }
}
