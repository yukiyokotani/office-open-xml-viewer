//! Display of a Word EQ field (ECMA-376 Part 1 17.16.5.19; the switches are
//! those of Word's EQ field: `\a` array, `\b` bracket, `\d` displace, `\f`
//! fraction, `\i` integral, `\l` list, `\o` overstrike, `\r` radical, `\s`
//! superscript or subscript and `\x` box).
//!
//! Word draws an EQ field itself: the field stores no result. The DOCX model
//! has two constructs that hold what Word draws:
//!
//! * a phonetic guide (`w:ruby`), for the overstrike of a raised guide over
//!   its base that Word writes for every phonetic guide;
//! * an inline equation (`m:oMath`), for every other instruction. Each EQ
//!   element maps onto the OMML object of the same meaning.
//!
//! The reading never fails: text that is not a well-formed element stays
//! literal text, as Word leaves it.

use ooxml_common::math::MathNode;

/// Resource policy for nested EQ elements, not a Word limit.
const MAX_DEPTH: usize = 32;

/// A Word phonetic guide stored as an EQ field (an `\o` overstrike of a
/// `\s\up` raised guide over its base).
/// Evidence: the local DOC files Word saved from DOCX files store every DOCX
/// `w:ruby` as `EQ \* jc2 \* "Font:..." \* hpsN \o\ad(\s\up K(guide),base)`
/// with `w:rubyAlign="distributeSpace"`, `w:hps` = N and `w:hpsRaise` = 2K; a
/// Word control with the centered alignment stores `\* jc0` and `\o(`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(in crate::doc) struct RubyForm {
    /// Guide text and its CPs inside the (hidden) instruction.
    pub(in crate::doc) guide: String,
    pub(in crate::doc) guide_cp: usize,
    /// Base text and its CPs inside the instruction.
    pub(in crate::doc) base: String,
    pub(in crate::doc) base_cp: usize,
    /// `w:hps`: guide text size in half-points, when the instruction states
    /// it; the guide characters' own size otherwise.
    pub(in crate::doc) guide_half_points: Option<u16>,
    /// `w:hpsRaise` / 2: raise of the guide in points.
    pub(in crate::doc) raise_pt: u16,
    /// ECMA-376 17.18.75 ST_RubyAlign.
    pub(in crate::doc) align: &'static str,
}

#[derive(Debug, Clone)]
pub(in crate::doc) enum Equation {
    Ruby(RubyForm),
    /// The equation and the CP of its first displayed character, whose
    /// properties size the equation.
    Math(Vec<MathNode>, Option<usize>),
}

struct Reader<'a> {
    chars: &'a [(char, usize)],
    at: usize,
    first_text_cp: Option<usize>,
}

/// General-formatting switches Word writes before a phonetic guide.
#[derive(Default)]
struct Preamble {
    jc: Option<u16>,
    hps: Option<u16>,
}

/// `chars` is the instruction with its CPs, starting at the `EQ` keyword.
pub(super) fn parse(chars: &[(char, usize)]) -> Equation {
    let mut reader = Reader {
        chars,
        at: 0,
        first_text_cp: None,
    };
    reader.skip_space();
    // The keyword.
    while reader.peek().is_some_and(|c| c.is_ascii_alphabetic()) {
        reader.at += 1;
    }
    let preamble = reader.preamble();
    let body = reader.at;
    if let Some(ruby) = reader.ruby(&preamble) {
        return Equation::Ruby(ruby);
    }
    reader.at = body;
    reader.first_text_cp = None;
    let mut nodes = reader.sequence(0, false);
    // Text after an unbalanced `)` is literal.
    while reader.at < reader.chars.len() {
        let (character, cp) = reader.chars[reader.at];
        reader.at += 1;
        reader.first_text_cp.get_or_insert(cp);
        push_text(&mut nodes, character);
        let rest = reader.sequence(0, false);
        append(&mut nodes, rest);
    }
    trim_trailing_space(&mut nodes);
    Equation::Math(nodes, reader.first_text_cp)
}

fn run(text: String) -> MathNode {
    MathNode::Run {
        text,
        // Word draws EQ text upright in the field's own font.
        style: "roman".into(),
        script: None,
    }
}

fn push_text(nodes: &mut Vec<MathNode>, character: char) {
    if let Some(MathNode::Run { text, .. }) = nodes.last_mut() {
        text.push(character);
    } else {
        nodes.push(run(character.to_string()));
    }
}

/// Append `more`, joining text that meets text.
fn append(nodes: &mut Vec<MathNode>, more: Vec<MathNode>) {
    for node in more {
        match (nodes.last_mut(), node) {
            (Some(MathNode::Run { text, .. }), MathNode::Run { text: next, .. }) => {
                text.push_str(&next);
            }
            (_, node) => nodes.push(node),
        }
    }
}

fn trim_trailing_space(nodes: &mut Vec<MathNode>) {
    if let Some(MathNode::Run { text, .. }) = nodes.last_mut() {
        let kept = text.trim_end().len();
        text.truncate(kept);
        if text.is_empty() {
            nodes.pop();
        }
    }
}

fn is_separator(character: char) -> bool {
    // Word separates arguments with the list separator of the system
    // locale: a comma, or a semicolon where the comma is the decimal symbol.
    matches!(character, ',' | ';')
}

impl Reader<'_> {
    fn peek(&self) -> Option<char> {
        self.chars.get(self.at).map(|(c, _)| *c)
    }

    fn peek_at(&self, offset: usize) -> Option<char> {
        self.chars.get(self.at + offset).map(|(c, _)| *c)
    }

    fn skip_space(&mut self) {
        while self.peek().is_some_and(char::is_whitespace) {
            self.at += 1;
        }
    }

    fn literal(&mut self, text: &str) -> bool {
        let matched = text.chars().enumerate().all(|(offset, expected)| {
            self.peek_at(offset)
                .is_some_and(|c| c.eq_ignore_ascii_case(&expected))
        });
        if matched {
            self.at += text.chars().count();
        }
        matched
    }

    fn number(&mut self) -> Option<u16> {
        let start = self.at;
        while self.peek().is_some_and(|c| c.is_ascii_digit()) {
            self.at += 1;
        }
        self.chars[start..self.at]
            .iter()
            .map(|(c, _)| *c)
            .collect::<String>()
            .parse()
            .ok()
    }

    /// `\* argument` general-formatting switches (ECMA-376 17.16.4.3); an
    /// argument may be quoted (Word also writes typographic quotes).
    fn preamble(&mut self) -> Preamble {
        let mut preamble = Preamble::default();
        loop {
            self.skip_space();
            if !(self.peek() == Some('\\') && self.peek_at(1) == Some('*')) {
                return preamble;
            }
            self.at += 2;
            self.skip_space();
            let close = match self.peek() {
                Some('"') => Some('"'),
                Some('\u{201c}') => Some('\u{201d}'),
                _ => None,
            };
            if let Some(close) = close {
                self.at += 1;
                while self.peek().is_some_and(|c| c != close) {
                    self.at += 1;
                }
                self.at = (self.at + 1).min(self.chars.len());
                continue;
            }
            if self.literal("jc") {
                preamble.jc = self.number();
            } else if self.literal("hps") {
                preamble.hps = self.number().filter(|value| *value > 0);
            }
            while self.peek().is_some_and(|c| !c.is_whitespace()) {
                self.at += 1;
            }
        }
    }

    /// `\o[\al|\ac|\ar|\ad](\s\up K(guide),base)` and nothing else.
    fn ruby(&mut self, preamble: &Preamble) -> Option<RubyForm> {
        self.skip_space();
        if !self.literal("\\o") {
            return None;
        }
        self.skip_space();
        let mut option = None;
        for (name, align) in [
            ("\\ad", "distributeSpace"),
            ("\\al", "left"),
            ("\\ac", "center"),
            ("\\ar", "right"),
        ] {
            if self.literal(name) {
                option = Some(align);
                break;
            }
        }
        self.skip_space();
        if !self.literal("(") {
            return None;
        }
        self.skip_space();
        if !self.literal("\\s") {
            return None;
        }
        self.skip_space();
        if !self.literal("\\up") {
            return None;
        }
        self.skip_space();
        let raise_pt = self.number()?;
        self.skip_space();
        if !self.literal("(") {
            return None;
        }
        let (guide, guide_cp) = self.plain(|c| c == ')')?;
        if !self.literal(")") {
            return None;
        }
        self.skip_space();
        if !self.peek().is_some_and(is_separator) {
            return None;
        }
        self.at += 1;
        let (base, base_cp) = self.plain(|c| c == ')')?;
        self.at += 1;
        self.skip_space();
        if self.at != self.chars.len() {
            return None;
        }
        // Word's WdPhoneticGuideAlignmentType, which the `\* jcN` switch
        // stores: 0 center, 1 0-1-0 (distributeLetter), 2 1-2-1
        // (distributeSpace), 3 left, 4 right, 5 right vertical.
        let align = match preamble.jc {
            Some(0) => "center",
            Some(1) => "distributeLetter",
            Some(2) => "distributeSpace",
            Some(3) => "left",
            Some(4) => "right",
            Some(5) => "rightVertical",
            _ => option.unwrap_or("center"),
        };
        Some(RubyForm {
            guide,
            guide_cp,
            base,
            base_cp,
            guide_half_points: preamble.hps,
            raise_pt,
            align,
        })
    }

    /// Non-empty text without element syntax, up to (not including) `end`.
    fn plain(&mut self, end: impl Fn(char) -> bool) -> Option<(String, usize)> {
        let start = self.at;
        while let Some(character) = self.peek() {
            if end(character) {
                break;
            }
            if matches!(character, '(' | ')' | '\\') || is_separator(character) {
                return None;
            }
            self.at += 1;
        }
        self.peek()?;
        (self.at > start).then(|| {
            (
                self.chars[start..self.at].iter().map(|(c, _)| *c).collect(),
                self.chars[start].1,
            )
        })
    }

    /// Text and elements up to an argument separator or closing parenthesis
    /// (neither consumed) when `in_arguments`, else up to a stray `)`.
    fn sequence(&mut self, depth: usize, in_arguments: bool) -> Vec<MathNode> {
        let mut nodes = Vec::new();
        while let Some(character) = self.peek() {
            let cp = self.chars[self.at].1;
            match character {
                ')' => break,
                c if in_arguments && is_separator(c) => break,
                '\\' => {
                    let next = self.peek_at(1);
                    match next {
                        // An escaped delimiter or backslash.
                        Some(escaped)
                            if matches!(escaped, '(' | ')' | '\\') || is_separator(escaped) =>
                        {
                            self.at += 2;
                            self.first_text_cp.get_or_insert(cp);
                            push_text(&mut nodes, escaped);
                        }
                        Some(name) if depth < MAX_DEPTH && is_element(name) => {
                            self.at += 2;
                            self.element(name.to_ascii_lowercase(), depth, &mut nodes);
                        }
                        _ => {
                            self.at += 1;
                            self.first_text_cp.get_or_insert(cp);
                            push_text(&mut nodes, '\\');
                        }
                    }
                }
                _ => {
                    self.at += 1;
                    // Leading blanks of an argument are not drawn.
                    if character.is_whitespace() && nodes.is_empty() {
                        continue;
                    }
                    self.first_text_cp.get_or_insert(cp);
                    push_text(&mut nodes, character);
                }
            }
        }
        nodes
    }

    /// The options and arguments of the element `name`, appended as its OMML
    /// equivalent.
    fn element(&mut self, name: char, depth: usize, nodes: &mut Vec<MathNode>) {
        let mut options: Vec<(String, Option<u16>, Option<char>)> = Vec::new();
        loop {
            self.skip_space();
            if self.peek() != Some('\\') {
                break;
            }
            let (Some(first), Some(second)) = (self.peek_at(1), self.peek_at(2)) else {
                break;
            };
            if !first.is_ascii_alphabetic() || !second.is_ascii_alphabetic() {
                break;
            }
            let option: String = [first, second]
                .iter()
                .collect::<String>()
                .to_ascii_lowercase();
            self.at += 3;
            // `\lc\c`, `\rc\c`, `\bc\c`, `\fc\c` and `\vc\c` name a character.
            let character = if matches!(option.as_str(), "lc" | "rc" | "bc" | "fc" | "vc")
                && self.peek() == Some('\\')
            {
                let character = self.peek_at(1);
                self.at = (self.at + 2).min(self.chars.len());
                character
            } else {
                None
            };
            self.skip_space();
            let value = self.number();
            options.push((option, value, character));
        }
        let has = |name: &str| options.iter().any(|(option, ..)| option == name);
        let character = |name: &str| {
            options
                .iter()
                .find(|(option, ..)| option == name)
                .and_then(|(_, _, character)| *character)
        };
        let mut arguments: Vec<Vec<MathNode>> = Vec::new();
        if self.peek() == Some('(') {
            self.at += 1;
            loop {
                let mut argument = self.sequence(depth + 1, true);
                trim_trailing_space(&mut argument);
                arguments.push(argument);
                match self.peek() {
                    Some(c) if is_separator(c) => self.at += 1,
                    Some(')') => {
                        self.at += 1;
                        break;
                    }
                    _ => break,
                }
            }
        }
        let mut arguments = arguments.into_iter();
        let mut next = || arguments.next().unwrap_or_default();
        match name {
            'f' => nodes.push(MathNode::Fraction {
                num: next(),
                den: next(),
                bar: None,
            }),
            'r' => {
                let first = next();
                let second = arguments.next();
                nodes.push(match second {
                    Some(radicand) => MathNode::Radical {
                        index: first,
                        radicand,
                    },
                    None => MathNode::Radical {
                        index: Vec::new(),
                        radicand: first,
                    },
                });
            }
            'i' => {
                let op = character("fc")
                    .or_else(|| character("vc"))
                    .map(String::from)
                    .unwrap_or_else(|| {
                        if has("su") {
                            "\u{2211}"
                        } else if has("pr") {
                            "\u{220f}"
                        } else {
                            "\u{222b}"
                        }
                        .into()
                    });
                nodes.push(MathNode::Nary {
                    op,
                    lim_loc: if has("in") { "subSup" } else { "undOvr" }.into(),
                    sub: next(),
                    sup: next(),
                    body: next(),
                });
            }
            'b' => {
                let both = character("bc");
                let closing = |open: char| match open {
                    '(' => ')',
                    '[' => ']',
                    '{' => '}',
                    '<' => '>',
                    other => other,
                };
                let (begin, end) = match (both, character("lc"), character("rc")) {
                    (Some(both), left, right) => (
                        Some(left.unwrap_or(both)),
                        Some(right.unwrap_or(closing(both))),
                    ),
                    (None, None, None) => (Some('('), Some(')')),
                    (None, left, right) => (left, right),
                };
                nodes.push(MathNode::Delimiter {
                    beg_char: begin.map(String::from).unwrap_or_default(),
                    end_char: end.map(String::from).unwrap_or_default(),
                    items: vec![next()],
                });
            }
            'a' => {
                let columns = options
                    .iter()
                    .find(|(option, ..)| option == "co")
                    .and_then(|(_, value, _)| *value)
                    .filter(|value| *value > 0)
                    .map_or(1, usize::from);
                let cells: Vec<Vec<MathNode>> = arguments.collect();
                let rows = cells.chunks(columns).map(<[_]>::to_vec).collect();
                nodes.push(MathNode::Array {
                    rows,
                    align: if has("al") { "left" } else { "center" }.into(),
                });
            }
            'l' => {
                let mut items = Vec::new();
                for (index, argument) in arguments.enumerate() {
                    if index != 0 {
                        push_text(&mut items, ',');
                    }
                    append(&mut items, argument);
                }
                nodes.push(MathNode::Group { items });
            }
            's' => {
                let cells: Vec<Vec<MathNode>> = arguments.collect();
                if has("up") {
                    nodes.push(MathNode::Sup {
                        base: Vec::new(),
                        sup: cells.into_iter().flatten().collect(),
                    });
                } else if has("do") {
                    nodes.push(MathNode::Sub {
                        base: Vec::new(),
                        sub: cells.into_iter().flatten().collect(),
                    });
                } else {
                    // Without a placement the elements are stacked.
                    nodes.push(MathNode::Array {
                        rows: cells.into_iter().map(|cell| vec![cell]).collect(),
                        align: "center".into(),
                    });
                }
            }
            'x' => {
                let any = ["to", "bo", "le", "ri"].iter().any(|side| has(side));
                nodes.push(MathNode::BorderBox {
                    hide_top: any && !has("to"),
                    hide_bot: any && !has("bo"),
                    hide_left: any && !has("le"),
                    hide_right: any && !has("ri"),
                    strike_h: false,
                    strike_v: false,
                    strike_bltr: false,
                    strike_tlbr: false,
                    base: next(),
                });
            }
            'o' => {
                // OMML has no overstrike: the elements follow each other.
                crate::rejection::approximated("Word EQ overstrike drawn in sequence", ());
                for argument in arguments {
                    append(nodes, argument);
                }
            }
            // `\d` moves the next character; it draws nothing itself.
            _ => crate::rejection::skipped("Word EQ displacement"),
        }
    }
}

fn is_element(name: char) -> bool {
    matches!(
        name.to_ascii_lowercase(),
        'a' | 'b' | 'd' | 'f' | 'i' | 'l' | 'o' | 'r' | 's' | 'x'
    )
}

/// Retained heap bytes of an equation, for the model budget.
pub(super) fn payload(nodes: &[MathNode]) -> usize {
    fn node(value: &MathNode) -> usize {
        let children: usize = match value {
            MathNode::Run { text, style, .. } => text.capacity() + style.capacity(),
            MathNode::Fraction { num, den, .. } => payload(num) + payload(den),
            MathNode::Sup { base, sup } => payload(base) + payload(sup),
            MathNode::Sub { base, sub } => payload(base) + payload(sub),
            MathNode::Nary {
                op,
                lim_loc,
                sub,
                sup,
                body,
            } => op.capacity() + lim_loc.capacity() + payload(sub) + payload(sup) + payload(body),
            MathNode::Delimiter {
                beg_char,
                end_char,
                items,
            } => {
                beg_char.capacity()
                    + end_char.capacity()
                    + items.capacity() * std::mem::size_of::<Vec<MathNode>>()
                    + items.iter().map(|item| payload(item)).sum::<usize>()
            }
            MathNode::Radical { index, radicand } => payload(index) + payload(radicand),
            MathNode::Array { rows, align } => {
                align.capacity()
                    + rows.capacity() * std::mem::size_of::<Vec<Vec<MathNode>>>()
                    + rows
                        .iter()
                        .map(|row| {
                            row.capacity() * std::mem::size_of::<Vec<MathNode>>()
                                + row.iter().map(|cell| payload(cell)).sum::<usize>()
                        })
                        .sum::<usize>()
            }
            MathNode::Group { items } => payload(items),
            MathNode::BorderBox { base, .. } => payload(base),
            // The reader above builds no other object.
            _ => 0,
        };
        children
    }
    std::mem::size_of_val(nodes) + nodes.iter().map(node).sum::<usize>()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn chars(text: &str) -> Vec<(char, usize)> {
        text.chars().enumerate().map(|(i, c)| (c, i + 10)).collect()
    }

    fn math(text: &str) -> serde_json::Value {
        match parse(&chars(text)) {
            Equation::Math(nodes, _) => serde_json::to_value(nodes).unwrap(),
            Equation::Ruby(_) => panic!("phonetic guide"),
        }
    }

    fn text(value: &str) -> serde_json::Value {
        serde_json::json!([{ "kind": "run", "text": value, "style": "roman" }])
    }

    #[test]
    fn fractions_and_radicals_map_to_their_omml_objects() {
        assert_eq!(
            math(" EQ \\f(1,2) "),
            serde_json::json!([{ "kind": "fraction", "num": text("1"), "den": text("2") }])
        );
        assert_eq!(
            math("EQ \\r(3,x)"),
            serde_json::json!([{ "kind": "radical", "index": text("3"), "radicand": text("x") }])
        );
        assert_eq!(
            math("EQ \\r(x)"),
            serde_json::json!([{ "kind": "radical", "radicand": text("x") }])
        );
        // Nested elements, a semicolon separator and escaped delimiters.
        assert_eq!(
            math("EQ a\\f(\\r(x);\\(y\\))"),
            serde_json::json!([
                { "kind": "run", "text": "a", "style": "roman" },
                { "kind": "fraction",
                  "num": [{ "kind": "radical", "radicand": text("x") }],
                  "den": text("(y)") }
            ])
        );
    }

    #[test]
    fn the_first_displayed_character_sizes_the_equation() {
        let Equation::Math(_, cp) = parse(&chars("EQ \\f(1,2)")) else {
            panic!()
        };
        assert_eq!(cp, Some(16));
        let Equation::Math(nodes, cp) = parse(&chars("EQ ")) else {
            panic!()
        };
        assert!(nodes.is_empty());
        assert_eq!(cp, None);
    }

    #[test]
    fn integrals_brackets_arrays_lists_scripts_and_boxes() {
        assert_eq!(
            math("EQ \\i\\su(i=1,n,x)"),
            serde_json::json!([{ "kind": "nary", "op": "\u{2211}", "limLoc": "undOvr",
                "sub": text("i=1"), "sup": text("n"), "body": text("x") }])
        );
        assert_eq!(
            math("EQ \\i \\in(0,1,x)")[0]["limLoc"],
            serde_json::json!("subSup")
        );
        assert_eq!(
            math("EQ \\b\\bc\\[(x)"),
            serde_json::json!([{ "kind": "delimiter", "begChar": "[", "endChar": "]",
                "items": [text("x")] }])
        );
        assert_eq!(
            math("EQ \\b\\lc\\{(x)"),
            serde_json::json!([{ "kind": "delimiter", "begChar": "{", "endChar": "",
                "items": [text("x")] }])
        );
        assert_eq!(math("EQ \\b(x)")[0]["endChar"], serde_json::json!(")"));
        assert_eq!(
            math("EQ \\a\\al\\co2\\vs3(a,b,c,d)"),
            serde_json::json!([{ "kind": "array", "align": "left",
                "rows": [[text("a"), text("b")], [text("c"), text("d")]] }])
        );
        assert_eq!(
            math("EQ \\l(a,b)"),
            serde_json::json!([{ "kind": "group", "items": text("a,b") }])
        );
        assert_eq!(
            math("EQ \\s\\up8(a)\\s\\do 8(b)"),
            serde_json::json!([
                { "kind": "sup", "base": [], "sup": text("a") },
                { "kind": "sub", "base": [], "sub": text("b") }
            ])
        );
        assert_eq!(
            math("EQ \\x\\to\\bo(a)"),
            serde_json::json!([{ "kind": "borderBox", "hideLeft": true, "hideRight": true,
                "base": text("a") }])
        );
    }

    #[test]
    fn overstrike_and_displacement_record_their_approximation() {
        crate::rejection::take_best_effort();
        assert_eq!(math("EQ \\o(a,b)\\d\\fo10()c"), text("abc"));
        assert_eq!(
            crate::rejection::take_best_effort(),
            [
                (
                    crate::rejection::BestEffort::Approximated,
                    "Word EQ overstrike drawn in sequence"
                ),
                (
                    crate::rejection::BestEffort::Skipped,
                    "Word EQ displacement"
                ),
            ]
        );
    }

    #[test]
    fn malformed_instructions_stay_literal_text() {
        assert_eq!(math("EQ a)b\\q"), text("a)b\\q"));
        assert_eq!(
            math("EQ \\f(1"),
            serde_json::json!([{ "kind": "fraction", "num": text("1"), "den": [] }])
        );
        // Nesting beyond the depth budget is kept as text.
        let deep = format!("EQ {}", "\\r(".repeat(MAX_DEPTH + 2));
        assert!(serde_json::to_string(&math(&deep))
            .unwrap()
            .contains("\\\\r("));
    }

    #[test]
    fn phonetic_guides_keep_their_alignment_size_and_raise() {
        let Equation::Ruby(form) = parse(&chars(
            "EQ \\* jc2 \\* \"Font:MS Mincho\" \\* hps10 \\o\\ad(\\s\\up 9(かん),漢)",
        )) else {
            panic!()
        };
        assert_eq!(
            (form.guide.as_str(), form.base.as_str(), form.align),
            ("かん", "漢", "distributeSpace")
        );
        assert_eq!((form.guide_half_points, form.raise_pt), (Some(10), 9));
        let Equation::Ruby(form) = parse(&chars(
            "EQ \\* jc0 \\* \u{201c}Font:游明朝\u{201d} \\* hps20 \\o(\\s\\up 9(かん),漢)",
        )) else {
            panic!()
        };
        assert_eq!((form.align, form.guide_half_points), ("center", Some(20)));
        // Without the switches the overstrike option aligns the guide and
        // the guide characters keep their own size.
        let Equation::Ruby(form) = parse(&chars("EQ \\o\\al(\\s\\up 5(a),b)")) else {
            panic!()
        };
        assert_eq!((form.align, form.guide_half_points), ("left", None));
        // Anything else is an equation.
        assert!(matches!(
            parse(&chars("EQ \\o(\\s\\up 5(a),b)c")),
            Equation::Math(..)
        ));
    }
}
